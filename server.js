const express = require("express");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const nodemailer = require("nodemailer");
const admin = require("firebase-admin");

function loadLocalEnv(baseDir) {
  const envPath = path.join(baseDir, ".env.local");
  if (!fs.existsSync(envPath)) return;

  const lines = fs.readFileSync(envPath, "utf8").split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const eqIndex = line.indexOf("=");
    if (eqIndex <= 0) continue;

    const key = line.slice(0, eqIndex).trim();
    if (!key || Object.prototype.hasOwnProperty.call(process.env, key)) continue;

    let value = line.slice(eqIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"'))
      || (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    process.env[key] = value;
  }
}

loadLocalEnv(__dirname);

const { createStore } = require("./datastore");

const app = express();
const PORT = Number(process.env.PORT || 5501);
const IS_VERCEL = String(process.env.VERCEL || "").toLowerCase() === "1";
const JWT_SECRET = process.env.JWT_SECRET || "routes_payroll_dev_secret_change_me";
const FIREBASE_WEB_API_KEY = process.env.FIREBASE_WEB_API_KEY || "AIzaSyCqK3ZVR-9qN9WmsXycGzYkar5hnZBEpW0";
const SUPABASE_URL = String(process.env.SUPABASE_URL || "").replace(/\/+$/, "");
const SUPABASE_ANON_KEY = String(process.env.SUPABASE_ANON_KEY || "").trim();
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
const SUPABASE_REQUEST_TIMEOUT_MS = Math.max(1000, Number(process.env.SUPABASE_REQUEST_TIMEOUT_MS || 8000));
const store = createStore({ baseDir: __dirname });
const RATE_BUCKETS = new Map();
const AUTH_CACHE = new Map();
let startupDegradedError = null;
let startupReady = false;
let startupState = "pending";
let startupInitPromise = null;

app.use(express.json({ limit: "2mb" }));
app.use(express.static(__dirname, { index: "index.html" }));

function isValidMonth(month) {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(month);
}

function toNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function sanitizeText(value, maxLength = 200) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

function sanitizeEmail(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .slice(0, 120);
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ""));
}

function isStrongPassword(value) {
  const raw = String(value || "");
  if (raw.length < 8 || raw.length > 80) return false;
  if (!/[a-z]/.test(raw)) return false;
  if (!/[A-Z]/.test(raw)) return false;
  if (!/[0-9]/.test(raw)) return false;
  return true;
}

function isUserEmailVerified(user) {
  return user?.email_verified !== false && user?.email_verified !== 0;
}

function cacheUser(user) {
  if (!user) return;
  const normalizedUser = {
    id: String(user.id || ""),
    username: String(user.username || ""),
    email: String(user.email || ""),
    email_verified: isUserEmailVerified(user),
    password_hash: String(user.password_hash || ""),
  };
  if (normalizedUser.id) AUTH_CACHE.set(`id:${normalizedUser.id}`, normalizedUser);
  if (normalizedUser.username) AUTH_CACHE.set(`username:${normalizedUser.username.toLowerCase()}`, normalizedUser);
  if (normalizedUser.email) AUTH_CACHE.set(`email:${normalizedUser.email.toLowerCase()}`, normalizedUser);
}

function getCachedUserByIdentifier(identifier) {
  const normalized = String(identifier || "").trim().toLowerCase();
  return AUTH_CACHE.get(`username:${normalized}`) || AUTH_CACHE.get(`email:${normalized}`) || null;
}

function isQuotaExceededError(error) {
  const message = String(error?.message || "").toLowerCase();
  return error?.code === 8 || message.includes("quota exceeded");
}

function isStoreUnavailableError(error) {
  const message = String(error?.message || "").toLowerCase();
  const causeMessage = String(error?.cause?.message || "").toLowerCase();
  return message.includes("fetch failed")
    || message.includes("request timed out")
    || message.includes("network")
    || message.includes("socket")
    || causeMessage.includes("enotfound")
    || causeMessage.includes("eai_again")
    || causeMessage.includes("timed out")
    || causeMessage.includes("network");
}

function startupStatusPayload() {
  if (startupState === "pending") {
    return { ok: true, pending: true, degraded: false, provider: store.provider };
  }
  if (!startupDegradedError) {
    return { ok: true, pending: false, degraded: false, provider: store.provider };
  }
  return {
    ok: true,
    pending: false,
    degraded: true,
    provider: store.provider,
    reason: isQuotaExceededError(startupDegradedError) ? "cloud_quota_exceeded" : "startup_failed",
  };
}

function fallbackCompanies() {
  return [
    {
      id: 1,
      name: "Red Lantern Restaurant",
      logo_data_url: "",
    },
  ];
}

const MONTHLY_BACKUP_DIR = IS_VERCEL
  ? path.join("/tmp", "routes-payroll", "monthly-backups")
  : path.join(__dirname, "data", "monthly-backups");

function ensureMonthlyBackupDir() {
  fs.mkdirSync(MONTHLY_BACKUP_DIR, { recursive: true });
}

function monthIsoFromDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function isLastCalendarDay(date) {
  const lastDay = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  return date.getDate() === lastDay;
}

function nextScheduledBackupDate(date = new Date()) {
  const candidate = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  if (date.getDate() <= candidate.getDate() && !isLastCalendarDay(date)) {
    return candidate;
  }
  return new Date(date.getFullYear(), date.getMonth() + 2, 0);
}

function parseBackupFilename(fileName) {
  const match = String(fileName || "").match(/^monthly-backup-(\d{4}-\d{2})-(.+)\.json$/);
  if (!match) return null;
  return {
    month: match[1],
    stamp: match[2],
  };
}

function listMonthlyBackupFiles() {
  ensureMonthlyBackupDir();
  return fs.readdirSync(MONTHLY_BACKUP_DIR)
    .filter((fileName) => fileName.endsWith(".json"))
    .map((fileName) => {
      const parsed = parseBackupFilename(fileName);
      if (!parsed) return null;
      const fullPath = path.join(MONTHLY_BACKUP_DIR, fileName);
      const stats = fs.statSync(fullPath);
      return {
        fileName,
        fullPath,
        month: parsed.month,
        createdAt: stats.mtime.toISOString(),
        sizeBytes: Number(stats.size || 0),
      };
    })
    .filter(Boolean)
    .sort((a, b) => Date.parse(b.createdAt) - Date.parse(a.createdAt));
}

async function getSystemBackupPayload() {
  const companies = await store.listCompaniesById();
  const rows = await store.listPayrollEntriesAll();
  const employees = [];
  const designations = [];
  for (const company of companies) {
    // eslint-disable-next-line no-await-in-loop
    const companyEmployees = await store.listEmployeesByCompany(company.id);
    employees.push(...companyEmployees.map(dbRowToEmployee));
    // eslint-disable-next-line no-await-in-loop
    const companyDesignations = await store.listDesignationPresets(company.id);
    designations.push(...companyDesignations.map(dbRowToDesignation));
  }

  return {
    companies: companies.map((company) => ({
      id: company.id,
      name: company.name,
      logoDataUrl: company.logo_data_url || "",
    })),
    employees,
    designations,
    entries: rows.map(dbRowToRecord),
  };
}

async function ensureMonthlyEmergencyBackup({ force = false } = {}) {
  const now = new Date();
  const currentMonth = monthIsoFromDate(now);
  const recentFiles = listMonthlyBackupFiles();
  const currentMonthBackup = recentFiles.find((item) => item.month === currentMonth) || null;
  const dueToday = isLastCalendarDay(now);

  if (!force && !dueToday) {
    return {
      createdThisCheck: false,
      hasCurrentMonthBackup: Boolean(currentMonthBackup),
      latestBackup: recentFiles[0] || null,
      recentBackups: recentFiles.slice(0, 12),
      isLastDayOfMonth: false,
      nextScheduledDate: nextScheduledBackupDate(now).toISOString().slice(0, 10),
    };
  }

  if (!force && currentMonthBackup) {
    return {
      createdThisCheck: false,
      hasCurrentMonthBackup: true,
      latestBackup: recentFiles[0] || null,
      recentBackups: recentFiles.slice(0, 12),
      isLastDayOfMonth: dueToday,
      nextScheduledDate: nextScheduledBackupDate(now).toISOString().slice(0, 10),
    };
  }

  const payload = await getSystemBackupPayload();
  ensureMonthlyBackupDir();
  const stamp = now.toISOString().replace(/[:.]/g, "-");
  const fileName = `monthly-backup-${currentMonth}-${stamp}.json`;
  const fullPath = path.join(MONTHLY_BACKUP_DIR, fileName);
  const backupEnvelope = {
    kind: "routes-payroll-monthly-emergency-backup",
    createdAt: now.toISOString(),
    backupMonth: currentMonth,
    provider: store.provider,
    payload,
  };
  fs.writeFileSync(fullPath, JSON.stringify(backupEnvelope, null, 2), "utf8");
  const updatedFiles = listMonthlyBackupFiles();
  return {
    createdThisCheck: true,
    hasCurrentMonthBackup: true,
    latestBackup: updatedFiles[0] || null,
    recentBackups: updatedFiles.slice(0, 12),
    isLastDayOfMonth: dueToday,
    nextScheduledDate: nextScheduledBackupDate(now).toISOString().slice(0, 10),
  };
}

function authProvider() {
  return store.provider === "supabase" ? "supabase" : "firebase";
}

function firebaseAuth() {
  return admin.auth();
}

function hasSupabaseAuthConfig() {
  return Boolean(SUPABASE_URL && SUPABASE_ANON_KEY && SUPABASE_SERVICE_ROLE_KEY);
}

async function fetchWithTimeout(url, options = {}, timeoutMs = SUPABASE_REQUEST_TIMEOUT_MS) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error?.name === "AbortError") {
      throw new Error(`Request timed out after ${timeoutMs}ms.`);
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

async function supabaseAuthAdmin(pathname, { method = "GET", body } = {}) {
  if (!hasSupabaseAuthConfig()) {
    throw new Error("Supabase Auth is not configured. Set SUPABASE_URL, SUPABASE_ANON_KEY, and SUPABASE_SERVICE_ROLE_KEY.");
  }

  const response = await fetchWithTimeout(`${SUPABASE_URL}${pathname}`, {
    method,
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      "Content-Type": "application/json",
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(payload?.msg || payload?.error_description || payload?.message || "Supabase auth request failed.");
    error.code = payload?.code || response.status;
    throw error;
  }
  return payload;
}

async function supabaseAuthPasswordSignIn(email, password) {
  if (!hasSupabaseAuthConfig()) {
    throw new Error("Supabase Auth is not configured. Set SUPABASE_URL, SUPABASE_ANON_KEY, and SUPABASE_SERVICE_ROLE_KEY.");
  }

  const response = await fetchWithTimeout(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ email, password }),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = String(payload?.error_code || payload?.error || payload?.msg || "").toLowerCase();
    if (message.includes("invalid") || message.includes("credentials")) {
      const error = new Error("Invalid username or password.");
      error.code = "auth/invalid-login";
      throw error;
    }
    throw new Error(payload?.msg || payload?.error_description || payload?.message || "Supabase sign-in failed.");
  }

  return payload;
}

async function listAllFirebaseAuthUsers() {
  const users = [];
  let nextPageToken;
  do {
    // eslint-disable-next-line no-await-in-loop
    const page = await firebaseAuth().listUsers(1000, nextPageToken);
    users.push(...(page.users || []));
    nextPageToken = page.pageToken;
  } while (nextPageToken);
  return users;
}

async function listAllSupabaseAuthUsers() {
  const users = [];
  let page = 1;

  while (true) {
    // eslint-disable-next-line no-await-in-loop
    const payload = await supabaseAuthAdmin(`/auth/v1/admin/users?page=${page}&per_page=200`);
    const batch = Array.isArray(payload?.users) ? payload.users : [];
    users.push(...batch);
    if (batch.length < 200) break;
    page += 1;
  }

  return users;
}

async function listAllAuthUsers() {
  if (authProvider() === "supabase") {
    return listAllSupabaseAuthUsers();
  }
  return listAllFirebaseAuthUsers();
}

async function countAuthUsers() {
  return store.countUsers();
}

function authUserEmailVerified(authUser) {
  if (!authUser) return false;
  if (typeof authUser.email_confirmed_at === "string" && authUser.email_confirmed_at) return true;
  if (authUser.email_confirmed_at instanceof Date) return true;
  return Boolean(authUser.user_metadata?.email_verified);
}

function authUserUsername(authUser) {
  const metadataUsername = sanitizeText(authUser?.user_metadata?.username, 40);
  if (metadataUsername) return metadataUsername;
  const email = sanitizeEmail(authUser?.email);
  if (!email) return "";
  return sanitizeText(email.split("@")[0], 40);
}

async function reconcileSupabaseUserFromAuth(identifier) {
  if (store.provider !== "supabase") return null;

  const normalized = sanitizeEmail(identifier) || sanitizeText(identifier, 120).toLowerCase();
  if (!normalized) return null;

  const authUsers = await listAllSupabaseAuthUsers();
  const authUser = authUsers.find((candidate) => {
    const email = sanitizeEmail(candidate?.email);
    const username = sanitizeText(candidate?.user_metadata?.username, 40).toLowerCase();
    return email === normalized || username === normalized;
  });

  if (!authUser?.id || !authUser?.email) return null;

  let storeUser = await store.getUserById(String(authUser.id));
  if (storeUser) return storeUser;

  storeUser = await store.findUserByEmail(String(authUser.email));
  if (storeUser) return storeUser;

  const username = authUserUsername(authUser);
  if (!username) return null;

  await store.createUser(username, String(authUser.email), "", {
    id: String(authUser.id),
    emailVerified: authUserEmailVerified(authUser),
  });

  return store.getUserById(String(authUser.id));
}

async function signInWithFirebasePassword(email, password) {
  const response = await fetchWithTimeout(
    `https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${encodeURIComponent(FIREBASE_WEB_API_KEY)}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email,
        password,
        returnSecureToken: true,
      }),
    }
  );

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = String(payload?.error?.message || "").toUpperCase();
    if (message.includes("INVALID_LOGIN_CREDENTIALS") || message.includes("INVALID_PASSWORD") || message.includes("EMAIL_NOT_FOUND")) {
      const error = new Error("Invalid username or password.");
      error.code = "auth/invalid-login";
      throw error;
    }
    throw new Error(payload?.error?.message || "Firebase sign-in failed.");
  }

  return payload;
}

async function createAuthUser({ email, password, username, emailVerified }) {
  if (authProvider() === "supabase") {
    return {
      id: crypto.randomUUID(),
      email,
    };
  }

  const firebaseUser = await firebaseAuth().createUser({
    email,
    password,
    displayName: username,
    emailVerified: Boolean(emailVerified),
  });
  return {
    id: String(firebaseUser.uid || ""),
    email: String(firebaseUser.email || email || ""),
  };
}

async function signInWithAuthPassword(email, password, passwordHash = "") {
  if (authProvider() === "supabase") {
    const normalizedHash = String(passwordHash || "");
    if (!normalizedHash) {
      const error = new Error("Invalid username or password.");
      error.code = "auth/invalid-login";
      throw error;
    }
    const valid = await bcrypt.compare(password, normalizedHash);
    if (!valid) {
      const error = new Error("Invalid username or password.");
      error.code = "auth/invalid-login";
      throw error;
    }
    return { ok: true };
  }
  return signInWithFirebasePassword(email, password);
}

async function updateAuthUserPassword(userId, newPassword) {
  if (authProvider() === "supabase") {
    return;
  }
  await firebaseAuth().updateUser(String(userId || ""), { password: newPassword });
}

function checkRateLimit(key, limit, windowMs) {
  const now = Date.now();
  const slot = RATE_BUCKETS.get(key) || { count: 0, resetAt: now + windowMs };
  if (now > slot.resetAt) {
    slot.count = 0;
    slot.resetAt = now + windowMs;
  }
  slot.count += 1;
  RATE_BUCKETS.set(key, slot);
  return slot.count <= limit;
}

function clientIp(req) {
  const xf = String(req.headers["x-forwarded-for"] || "");
  if (xf) return xf.split(",")[0].trim();
  return req.ip || "unknown";
}

function hashResetCode(rawCode) {
  return crypto.createHash("sha256").update(String(rawCode)).digest("hex");
}

function createMailer() {
  const host = process.env.SMTP_HOST || "";
  const port = Number(process.env.SMTP_PORT || 587);
  const user = process.env.SMTP_USER || "";
  const pass = process.env.SMTP_PASS || "";
  const from = process.env.SMTP_FROM || user;
  const secure = String(process.env.SMTP_SECURE || "").toLowerCase() === "true";

  if (!host || !port || !user || !pass || !from) {
    return null;
  }

  const transport = nodemailer.createTransport({
    host,
    port,
    secure,
    auth: { user, pass },
  });

  return { transport, from };
}

async function sendUsernameRecoveryEmail({ to, username }) {
  const mailer = createMailer();
  if (!mailer) {
    throw new Error("Email service is not configured. Set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM.");
  }

  await mailer.transport.sendMail({
    from: mailer.from,
    to,
    subject: "Routes Payroll Username Recovery",
    text: [
      "You requested username recovery for Routes Payroll.",
      `Username: ${username}`,
      "If this was not you, please ignore this email.",
    ].join("\n"),
  });
}

async function sendPasswordResetLinkEmail({ to, resetLink }) {
  const mailer = createMailer();
  if (!mailer) {
    throw new Error("Email service is not configured. Set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM.");
  }

  await mailer.transport.sendMail({
    from: mailer.from,
    to,
    subject: "Routes Payroll Password Reset Link",
    text: [
      "You requested a password reset for Routes Payroll.",
      "Open this link to set a new password:",
      resetLink,
      "This link expires in 15 minutes.",
      "If this was not you, please ignore this email.",
    ].join("\n"),
  });
}

async function sendEmailVerificationLinkEmail({ to, verifyLink }) {
  const mailer = createMailer();
  if (!mailer) {
    throw new Error("Email service is not configured. Set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM.");
  }

  await mailer.transport.sendMail({
    from: mailer.from,
    to,
    subject: "Verify Your Routes Payroll Email",
    text: [
      "Welcome to Routes Payroll.",
      "Please verify your registered email by opening this link:",
      verifyLink,
      "After verification, you can go back to your system or log in to your system.",
      "If this was not you, please ignore this email.",
    ].join("\n"),
  });
}

function appBaseUrl(req) {
  const forwardedProto = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
  const protocol = forwardedProto || req.protocol || "http";
  const host = String(req.get("host") || `127.0.0.1:${PORT}`);
  return `${protocol}://${host}`;
}

function sanitizeLogoDataUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (!/^data:image\/[a-zA-Z0-9.+-]+;base64,/.test(raw)) return "";
  return raw.slice(0, 1_200_000);
}

function parseCompanyId(value, fallback = 1) {
  if (value === undefined || value === null || value === "") return fallback;
  const id = Number(value);
  if (!Number.isInteger(id) || id <= 0) return null;
  return id;
}

function normalizeEmploymentStatus(value) {
  return String(value || "").trim().toLowerCase();
}

async function resolveCompanyId(req, res, source = "query") {
  const rawValue = source === "body" ? req.body?.companyId : req.query?.companyId;
  const companyId = parseCompanyId(rawValue, 1);
  if (!companyId) {
    res.status(400).json({ error: "Invalid company id." });
    return null;
  }

  try {
    if (startupDegradedError && companyId === 1) {
      return 1;
    }
    const exists = await store.companyExists(companyId);
    if (!exists) {
      res.status(404).json({ error: "Company not found." });
      return null;
    }
    return companyId;
  } catch (error) {
    if (companyId === 1 && (startupDegradedError || isStoreUnavailableError(error))) {
      return 1;
    }
    if (isStoreUnavailableError(error)) {
      res.status(503).json({ error: "Payroll data service is temporarily unreachable. Please retry in a moment." });
      return null;
    }
    res.status(500).json({ error: "Failed to resolve company." });
    return null;
  }
}

function normalizeRecord(record, index) {
  return {
    id: Number.isFinite(Number(record.id)) ? Number(record.id) : null,
    employeeId: sanitizeText(record.employeeId, 60),
    employeeName: sanitizeText(record.employeeName, 120),
    designation: sanitizeText(record.designation, 120),
    presentSalary: Math.max(0, toNumber(record.presentSalary)),
    increment: Math.max(0, toNumber(record.increment)),
    oldAdvanceTaken: Math.max(0, toNumber(record.oldAdvanceTaken)),
    extraAdvanceAdded: Math.max(0, toNumber(record.extraAdvanceAdded)),
    deductionEntered: Math.max(0, toNumber(record.deductionEntered)),
    daysAbsent: clamp(toNumber(record.daysAbsent), 0, 31),
    comment: sanitizeText(record.comment, 350),
    positionIndex: Number.isFinite(Number(record.positionIndex)) ? Number(record.positionIndex) : index,
  };
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
}

function normalizeEmployee(employee, index) {
  const rawStatus = String(employee?.status || "working").toLowerCase();
  const status = rawStatus === "leave" || rawStatus === "resumed" || rawStatus === "terminated" ? rawStatus : "working";
  const leaveFrom = isIsoDate(employee?.leaveFrom) ? String(employee.leaveFrom) : "";
  const leaveTo = isIsoDate(employee?.leaveTo) ? String(employee.leaveTo) : "";
  const terminatedOn = isIsoDate(employee?.terminatedOn) ? String(employee.terminatedOn) : "";

  return {
    id: Number.isFinite(Number(employee?.id)) ? Number(employee.id) : null,
    employeeId: sanitizeText(employee?.employeeId, 60),
    employeeName: sanitizeText(employee?.employeeName, 120),
    joiningDate: isIsoDate(employee?.joiningDate) ? String(employee.joiningDate) : "",
    birthDate: isIsoDate(employee?.birthDate) ? String(employee.birthDate) : "",
    baseSalary: Math.max(0, toNumber(employee?.baseSalary)),
    openingAdvance: Math.max(0, toNumber(employee?.openingAdvance)),
    designation: sanitizeText(employee?.designation, 120),
    mobileNumber: sanitizeText(employee?.mobileNumber, 30),
    status,
    leaveFrom,
    leaveTo,
    terminatedOn,
    notes: sanitizeText(employee?.notes, 350),
    positionIndex: Number.isFinite(Number(employee?.positionIndex)) ? Number(employee.positionIndex) : index,
  };
}

function dbRowToEmployee(row) {
  return {
    id: Number(row.id),
    companyId: Number(row.company_id),
    employeeId: row.employee_id,
    employeeName: row.employee_name,
    joiningDate: row.joining_date || "",
    birthDate: row.birth_date || "",
    baseSalary: Number(row.base_salary || 0),
    openingAdvance: Number(row.opening_advance || 0),
    designation: row.designation || "",
    mobileNumber: row.mobile_number || "",
    status: row.status || "working",
    leaveFrom: row.leave_from || "",
    leaveTo: row.leave_to || "",
    terminatedOn: row.terminated_on || "",
    notes: row.notes || "",
    positionIndex: Number(row.position_index || 0),
  };
}

function previousMonthIso(month) {
  if (!/^\d{4}-\d{2}$/.test(String(month || ""))) return "";
  const [year, mon] = String(month).split("-").map((item) => Number(item));
  const date = new Date(year, mon - 1, 1);
  date.setMonth(date.getMonth() - 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function advanceRemainedFromPayrollRow(row) {
  const oldAdvance = Math.max(0, toNumber(row?.old_advance_taken));
  const extraAdvance = Math.max(0, toNumber(row?.extra_advance_added));
  const totalAdvance = oldAdvance + extraAdvance;
  const deductionEntered = Math.max(0, toNumber(row?.deduction_entered));
  const deductionApplied = Math.min(deductionEntered, totalAdvance);
  return totalAdvance - deductionApplied;
}

function dbRowToDesignation(row) {
  return {
    id: Number(row.id),
    companyId: Number(row.company_id),
    name: String(row.name || ""),
    positionIndex: Number(row.position_index || 0),
  };
}

function dbRowToPayrollReport(row) {
  return {
    id: Number(row.id),
    companyId: Number(row.company_id),
    month: String(row.month || ""),
    checkedAt: String(row.checked_at || ""),
    generatedAt: String(row.generated_at || ""),
    employeeCount: Number(row.employee_count || 0),
  };
}

function parseSnapshotJson(value) {
  try {
    const parsed = JSON.parse(String(value || "{}"));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

async function buildPayrollReportSnapshot(month, companyId) {
  const [company, records] = await Promise.all([
    store.getCompanyById(companyId),
    listMergedPayrollMonthRecords(month, companyId),
  ]);

  return {
    company: {
      id: Number(company?.id || companyId),
      name: String(company?.name || "Company"),
      logoDataUrl: String(company?.logo_data_url || ""),
    },
    month,
    records,
  };
}

async function listMergedPayrollMonthRecords(month, companyId) {
  const [rows, employees] = await Promise.all([
    store.listPayrollByMonthCompany(month, companyId),
    store.listEmployeesByCompany(companyId),
  ]);
  const employeeByEmployeeId = new Map(
    employees.map((employee) => [String(employee.employee_id || ""), employee])
  );
  const prevMonth = previousMonthIso(month);
  const previousRows = prevMonth ? await store.listPayrollByMonthCompany(prevMonth, companyId) : [];
  const previousByEmployeeId = new Map(previousRows.map((row) => [String(row.employee_id || ""), row]));
  const byEmployeeId = new Map(rows.map((row) => [String(row.employee_id || ""), row]));
  const records = [];

  employees.forEach((employee, index) => {
    const linked = byEmployeeId.get(String(employee.employee_id || "")) || null;
    const previous = previousByEmployeeId.get(String(employee.employee_id || "")) || null;
    const carriedAdvance = previous ? advanceRemainedFromPayrollRow(previous) : null;
    records.push({
      id: linked?.id || null,
      companyId,
      month,
      employeeId: employee.employee_id,
      employeeName: employee.employee_name,
      joiningDate: employee.joining_date || "",
      designation: linked?.designation || employee.designation || "",
      presentSalary: linked?.present_salary ?? Number(employee.base_salary || 0),
      increment: linked?.increment ?? 0,
      oldAdvanceTaken: linked?.old_advance_taken ?? (carriedAdvance ?? Number(employee.opening_advance || 0)),
      extraAdvanceAdded: linked?.extra_advance_added ?? 0,
      deductionEntered: linked?.deduction_entered ?? 0,
      daysAbsent: linked?.days_absent ?? 0,
      comment: linked?.comment || "",
      positionIndex: Number(employee.position_index ?? index),
      employeeStatus: employee.status || "working",
      leaveFrom: employee.leave_from || "",
      leaveTo: employee.leave_to || "",
      terminatedOn: employee.terminated_on || "",
    });
    byEmployeeId.delete(String(employee.employee_id || ""));
  });

  for (const extra of byEmployeeId.values()) {
    const linkedEmployee = employeeByEmployeeId.get(String(extra.employee_id || ""));
    if (!linkedEmployee) {
      continue;
    }
    records.push({
      ...dbRowToRecord(extra),
      joiningDate: linkedEmployee?.joining_date || "",
      employeeStatus: linkedEmployee?.status || "working",
      leaveFrom: linkedEmployee?.leave_from || "",
      leaveTo: linkedEmployee?.leave_to || "",
      terminatedOn: linkedEmployee?.terminated_on || "",
    });
  }

  return records;
}

async function buildPayrollReportRecord(month, companyId, employeeId) {
  const snapshot = await buildPayrollReportSnapshot(month, companyId);
  const record = (snapshot.records || []).find((item) => String(item.employeeId || "") === String(employeeId || ""));
  if (!record) return null;
  return {
    company: snapshot.company,
    month: snapshot.month,
    record,
  };
}

async function ensurePayrollCheckedForMonth(month, companyId) {
  const snapshot = await buildPayrollReportSnapshot(month, companyId);
  if (!Array.isArray(snapshot.records) || snapshot.records.length === 0) {
    const error = new Error("No payroll rows found for this month.");
    error.statusCode = 400;
    throw error;
  }

  const existing = await store.getPayrollReportByMonthCompany(month, companyId);
  if (existing?.checked_at) {
    return {
      checkedAt: String(existing.checked_at || ""),
      reportId: existing.id,
      snapshot,
    };
  }

  const checkedAt = new Date().toISOString();
  const result = await store.markPayrollChecked(month, companyId, checkedAt);
  return {
    checkedAt,
    reportId: result.id,
    snapshot,
  };
}

function createToken(user) {
  return jwt.sign(
    {
      userId: user.id,
      username: user.username,
      email: user.email || "",
      emailVerified: isUserEmailVerified(user),
    },
    JWT_SECRET,
    { expiresIn: "30d" }
  );
}

function authMiddleware(req, res, next) {
  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";

  if (!token) {
    res.status(401).json({ error: "Missing authentication token." });
    return;
  }

  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.user = payload;
    next();
  } catch {
    res.status(401).json({ error: "Invalid or expired token." });
  }
}

app.get("/api/health", (_req, res) => {
  res.json({ ...startupStatusPayload(), time: new Date().toISOString() });
});

app.get("/api/auth/bootstrap", async (_req, res) => {
  if (!startupReady && !startupDegradedError) {
    res.json({
      needsSetup: false,
      pending: true,
      degraded: false,
      message: "Connecting to cloud payroll service. Please wait...",
    });
    return;
  }

  if (startupDegradedError && store.provider === "supabase") {
    res.json({
      needsSetup: false,
      pending: false,
      degraded: true,
      message: startupDegradedError.message || "Cloud startup failed. Check internet access and Supabase settings.",
    });
    return;
  }

  try {
    const count = await countAuthUsers();
    res.json({
      needsSetup: count === 0,
      pending: false,
      degraded: Boolean(startupDegradedError && isQuotaExceededError(startupDegradedError) && store.provider !== "supabase"),
      message: startupDegradedError && isQuotaExceededError(startupDegradedError) && store.provider !== "supabase"
        ? "Cloud payroll data is temporarily busy, but sign-in remains available."
        : "",
    });
  } catch (error) {
    res.status(500).json({ error: "Failed to check bootstrap state." });
  }
});

app.post("/api/auth/register", async (req, res) => {
  const companyName = sanitizeText(req.body?.companyName, 120);
  const email = sanitizeEmail(req.body?.email);
  const username = sanitizeText(req.body?.username, 40);
  const password = String(req.body?.password || "");

  if (companyName.length < 2 || !isValidEmail(email) || username.length < 3 || !isStrongPassword(password)) {
    res.status(400).json({ error: "Company name, email, username, or password is invalid." });
    return;
  }

  try {
    const userCount = await countAuthUsers();
    if (userCount > 0) {
      res.status(409).json({ error: "Admin account already exists." });
      return;
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const authUser = await createAuthUser({
      email,
      password,
      username,
      emailVerified: false,
    });

    const verifyToken = crypto.randomBytes(24).toString("hex");
    const verifyTokenHash = hashResetCode(verifyToken);
    const expiresAtIso = new Date(Date.now() + (24 * 60 * 60 * 1000)).toISOString();
    await store.createUser(username, email, passwordHash, { id: authUser.id, emailVerified: false });
    await store.setUserEmailVerificationById(authUser.id, verifyTokenHash, expiresAtIso);

    if (!startupDegradedError || store.provider === "supabase") {
      await store.updateCompanyName(1, companyName, "");
    }

    const verifyLink = `${appBaseUrl(req)}/verify-email.html?token=${encodeURIComponent(verifyToken)}`;
    let verificationEmailSent = true;
    let message = "Account created. Verify your email from the link we sent, or resend it from Settings.";
    try {
      await sendEmailVerificationLinkEmail({ to: email, verifyLink });
    } catch (error) {
      verificationEmailSent = false;
      message = "Account created, but verification email could not be sent. Use Verify Email in Settings after SMTP is configured.";
    }

    const createdUser = { id: authUser.id, username, email, email_verified: false, password_hash: passwordHash };
    cacheUser(createdUser);
    const token = createToken(createdUser);
    res.status(201).json({
      token,
      message,
      verificationEmailSent,
      user: { id: authUser.id, username, email, emailVerified: false },
    });
  } catch (error) {
    if (
      error?.code === "auth/email-already-exists"
      || error?.code === "auth/uid-already-exists"
      || error?.code === "email_exists"
      || error?.code === "user_already_exists"
      || error?.code === "unique"
      || String(error?.message || "").toLowerCase().includes("already")
    ) {
      res.status(409).json({ error: "Admin account or email already exists." });
      return;
    }
    res.status(500).json({ error: "Failed to register admin account." });
  }
});

app.post("/api/auth/login", async (req, res) => {
  const identifier = sanitizeText(req.body?.username, 120);
  const password = String(req.body?.password || "");
  const ip = clientIp(req);

  if (!identifier || !password) {
    res.status(400).json({ error: "Username/email and password are required." });
    return;
  }
  if (!checkRateLimit(`login:${ip}`, 10, 60_000)) {
    res.status(429).json({ error: "Too many login attempts. Please wait and retry." });
    return;
  }

  try {
    const user = await store.findUserByIdentifier(identifier);
    if (!user) {
      res.status(401).json({ error: "Invalid username or password." });
      return;
    }

    await signInWithAuthPassword(user.email, password, user.password_hash);
    if (!isUserEmailVerified(user)) {
      res.status(403).json({ error: "Email not verified. Open the verification email or resend it from Settings on your registered system." });
      return;
    }

    cacheUser(user);
    const token = createToken(user);
    res.json({
      token,
      user: {
        id: user.id,
        username: user.username,
        email: user.email || "",
        emailVerified: isUserEmailVerified(user),
      },
    });
  } catch (error) {
    if (error?.code === "auth/invalid-login") {
      res.status(401).json({ error: "Invalid username or password." });
      return;
    }
    if (isQuotaExceededError(error)) {
      res.status(503).json({ error: "Cloud login service is temporarily busy. Please keep the app open and retry shortly." });
      return;
    }
    res.status(500).json({ error: "Login failed." });
  }
});

app.post("/api/auth/send-email-verification", authMiddleware, async (req, res) => {
  const ip = clientIp(req);
  if (!checkRateLimit(`verify-email:${ip}`, 5, 60_000)) {
    res.status(429).json({ error: "Too many verification requests. Please wait and retry." });
    return;
  }

  try {
    const user = await store.getUserById(String(req.user.userId || ""));
    if (!user) {
      res.status(404).json({ error: "User not found." });
      return;
    }
    if (isUserEmailVerified(user)) {
      res.json({ ok: true, message: "Email is already verified." });
      return;
    }
    if (!isValidEmail(user.email || "")) {
      res.status(400).json({ error: "Registered email is invalid." });
      return;
    }

    const verifyToken = crypto.randomBytes(24).toString("hex");
    const verifyTokenHash = hashResetCode(verifyToken);
    const expiresAtIso = new Date(Date.now() + (24 * 60 * 60 * 1000)).toISOString();
    await store.setUserEmailVerificationById(user.id, verifyTokenHash, expiresAtIso);
    const verifyLink = `${appBaseUrl(req)}/verify-email.html?token=${encodeURIComponent(verifyToken)}`;
    await sendEmailVerificationLinkEmail({ to: user.email, verifyLink });
    res.json({ ok: true, message: "Verification email sent to your registered email address." });
  } catch (error) {
    res.status(500).json({ error: String(error?.message || "Failed to send verification email.") });
  }
});

app.post("/api/auth/recover-email", async (req, res) => {
  const companyName = sanitizeText(req.body?.companyName, 120);
  const email = sanitizeEmail(req.body?.email);
  const ip = clientIp(req);

  if (companyName.length < 2 || !isValidEmail(email)) {
    res.status(400).json({ error: "Company name or email is invalid." });
    return;
  }
  if (!checkRateLimit(`recover-email:${ip}`, 5, 60_000)) {
    res.status(429).json({ error: "Too many recovery attempts. Please wait and retry." });
    return;
  }

  try {
    const company = startupDegradedError ? { name: "Red Lantern Restaurant" } : await store.getCompanyById(1);
    const expectedCompanyName = sanitizeText(company?.name || "", 120).toLowerCase();
    const user = await store.findUserByEmail(email);
    const validMatch = expectedCompanyName
      && expectedCompanyName === companyName.toLowerCase()
      && user
      && isValidEmail(user.email || "")
      && String(user.email || "").toLowerCase() === email;

    if (validMatch) {
      await sendUsernameRecoveryEmail({
        to: user.email,
        username: user.username,
      });
    }

    res.json({ ok: true, message: "If account exists, recovery details have been sent." });
  } catch {
    res.status(500).json({ error: "Recovery request failed." });
  }
});

app.post("/api/auth/request-password-reset", async (req, res) => {
  const identifier = sanitizeText(req.body?.identifier, 120);
  const ip = clientIp(req);
  if (!identifier) {
    res.status(400).json({ error: "Username or email is required." });
    return;
  }
  if (!checkRateLimit(`request-reset:${ip}`, 5, 60_000)) {
    res.status(429).json({ error: "Too many reset requests. Please wait and retry." });
    return;
  }

  try {
    const user = await store.findUserByIdentifier(identifier);
    if (!user) {
      res.json({ ok: true, message: "If account exists, password reset link has been sent." });
      return;
    }
    if (!isValidEmail(user.email || "")) {
      res.json({ ok: true, message: "If account exists, password reset link has been sent." });
      return;
    }

    const resetToken = crypto.randomBytes(24).toString("hex");
    const resetTokenHash = hashResetCode(resetToken);
    const expiresAtIso = new Date(Date.now() + (15 * 60 * 1000)).toISOString();
    await store.setUserResetCodeById(user.id, resetTokenHash, expiresAtIso);
    const resetLink = `${appBaseUrl(req)}/reset-password.html?token=${encodeURIComponent(resetToken)}`;
    await sendPasswordResetLinkEmail({
      to: user.email,
      resetLink,
    });

    res.json({ ok: true, message: "If account exists, password reset link has been sent." });
  } catch (error) {
    res.status(500).json({ error: String(error?.message || "Reset request failed.") });
  }
});

app.post("/api/auth/reset-password-with-token", async (req, res) => {
  const token = sanitizeText(req.body?.token, 200);
  const newPassword = String(req.body?.newPassword || "");
  const ip = clientIp(req);

  if (!token || !isStrongPassword(newPassword)) {
    res.status(400).json({ error: "Token or new password is invalid." });
    return;
  }
  if (!checkRateLimit(`confirm-reset:${ip}`, 10, 60_000)) {
    res.status(429).json({ error: "Too many attempts. Please wait and retry." });
    return;
  }

  try {
    const tokenHash = hashResetCode(token);
    const user = await store.findUserByResetCodeHash(tokenHash);
    if (!user) {
      res.status(401).json({ error: "Invalid or expired reset link." });
      return;
    }

    const expectedHash = String(user.reset_code_hash || "");
    const expiresAt = String(user.reset_code_expires_at || "");
    const isExpired = !expiresAt || Date.parse(expiresAt) < Date.now();
    if (!expectedHash || isExpired || expectedHash !== tokenHash) {
      res.status(401).json({ error: "Invalid or expired reset link." });
      return;
    }

    const passwordHash = await bcrypt.hash(newPassword, 12);
    await updateAuthUserPassword(user.id, newPassword);
    await store.updateUserPasswordById(user.id, passwordHash);
    await store.clearUserResetCodeById(user.id);
    res.json({ ok: true, message: "Password reset successful." });
  } catch (error) {
    res.status(500).json({ error: String(error?.message || "Password reset failed.") });
  }
});

app.post("/api/auth/verify-email-token", async (req, res) => {
  const token = sanitizeText(req.body?.token, 200);
  const ip = clientIp(req);

  if (!token) {
    res.status(400).json({ error: "Verification token is required." });
    return;
  }
  if (!checkRateLimit(`verify-email-token:${ip}`, 10, 60_000)) {
    res.status(429).json({ error: "Too many attempts. Please wait and retry." });
    return;
  }

  try {
    const tokenHash = hashResetCode(token);
    const user = await store.findUserByEmailVerificationHash(tokenHash);
    if (!user) {
      res.status(401).json({ error: "Invalid or expired verification link." });
      return;
    }

    const expectedHash = String(user.email_verification_hash || "");
    const expiresAt = String(user.email_verification_expires_at || "");
    const isExpired = !expiresAt || Date.parse(expiresAt) < Date.now();
    if (!expectedHash || isExpired || expectedHash !== tokenHash) {
      res.status(401).json({ error: "Invalid or expired verification link." });
      return;
    }

    await store.markUserEmailVerifiedById(user.id);
    res.json({ ok: true, message: "Email verified successfully. You can go back to your system or log in to your system." });
  } catch (error) {
    res.status(500).json({ error: String(error?.message || "Email verification failed.") });
  }
});

app.get("/api/auth/me", authMiddleware, async (req, res) => {
  try {
    res.json({
      user: {
        id: req.user.userId,
        username: req.user.username,
        email: req.user.email || "",
        emailVerified: req.user.emailVerified !== false,
      },
    });
  } catch {
    res.status(500).json({ error: "Failed to load user profile." });
  }
});

app.post("/api/auth/change-password", authMiddleware, async (req, res) => {
  const newPassword = String(req.body?.newPassword || "");
  const confirmPassword = String(req.body?.confirmPassword || "");

  if (!isStrongPassword(newPassword)) {
    res.status(400).json({ error: "Password must be 8+ chars with uppercase, lowercase, and number." });
    return;
  }
  if (newPassword !== confirmPassword) {
    res.status(400).json({ error: "New password and confirm password do not match." });
    return;
  }

  try {
    const user = await store.getUserById(String(req.user.userId || ""));
    if (!user) {
      res.status(404).json({ error: "User not found." });
      return;
    }
    await updateAuthUserPassword(String(req.user.userId || ""), newPassword);
    const passwordHash = await bcrypt.hash(newPassword, 12);
    await store.updateUserPasswordById(String(req.user.userId || ""), passwordHash);
    cacheUser(user);
    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ error: String(error?.message || "Failed to update password.") });
  }
});

app.get("/api/companies", authMiddleware, async (_req, res) => {
  try {
    const rows = await store.listCompanies();
    res.json({
      companies: rows.map((row) => ({
        id: row.id,
        name: row.name,
        logoDataUrl: row.logo_data_url || "",
      })),
    });
  } catch (error) {
    if (isQuotaExceededError(error) || startupDegradedError) {
      const rows = fallbackCompanies();
      res.json({
        companies: rows.map((row) => ({
          id: row.id,
          name: row.name,
          logoDataUrl: row.logo_data_url || "",
        })),
        degraded: true,
      });
      return;
    }
    if (isStoreUnavailableError(error)) {
      const rows = fallbackCompanies();
      res.json({
        companies: rows.map((row) => ({
          id: row.id,
          name: row.name,
          logoDataUrl: row.logo_data_url || "",
        })),
        degraded: true,
      });
      return;
    }
    res.status(500).json({ error: "Failed to fetch companies." });
  }
});

app.get("/api/employees", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  try {
    const rows = await store.listEmployeesByCompany(companyId);
    res.json({ employees: rows.map(dbRowToEmployee) });
  } catch {
    res.status(500).json({ error: "Failed to fetch employees." });
  }
});

app.post("/api/employees", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;

  const existing = await store.listEmployeesByCompany(companyId).catch(() => []);
  const normalized = normalizeEmployee(req.body, existing.length);
  if (
    normalized.employeeName.length < 2
    || normalized.employeeId.length < 2
    || normalized.joiningDate.length !== 10
    || normalized.designation.length < 2
  ) {
    res.status(400).json({ error: "Employee name, ID, joining date, and designation are required." });
    return;
  }

  try {
    const result = await store.createEmployee(companyId, normalized);
    const saved = await store.getEmployeeByIdCompany(result.id, companyId);
    res.status(201).json({ employee: dbRowToEmployee(saved) });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Employee ID already exists for this company." });
      return;
    }
    res.status(500).json({ error: "Failed to create employee." });
  }
});

app.put("/api/employees/:id", authMiddleware, async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) {
    res.status(400).json({ error: "Invalid employee id." });
    return;
  }
  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;

  const existing = await store.getEmployeeByIdCompany(id, companyId);
  if (!existing) {
    res.status(404).json({ error: "Employee not found." });
    return;
  }

  const normalized = normalizeEmployee(req.body, Number(existing.position_index || 0));
  if (
    normalized.employeeName.length < 2
    || normalized.employeeId.length < 2
    || normalized.joiningDate.length !== 10
    || normalized.designation.length < 2
  ) {
    res.status(400).json({ error: "Employee name, ID, joining date, and designation are required." });
    return;
  }

  try {
    await store.updateEmployee(companyId, id, normalized);
    const saved = await store.getEmployeeByIdCompany(id, companyId);
    res.json({ employee: dbRowToEmployee(saved) });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Employee ID already exists for this company." });
      return;
    }
    res.status(500).json({ error: "Failed to update employee." });
  }
});

app.delete("/api/employees/:id", authMiddleware, async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) {
    res.status(400).json({ error: "Invalid employee id." });
    return;
  }
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  try {
    await store.deleteEmployeeByIdCompany(id, companyId);
    res.json({ ok: true });
  } catch {
    res.status(500).json({ error: "Failed to delete employee." });
  }
});

app.post("/api/companies", authMiddleware, async (req, res) => {
  const name = sanitizeText(req.body?.name, 120);
  const logoDataUrl = sanitizeLogoDataUrl(req.body?.logoDataUrl);
  if (name.length < 2) {
    res.status(400).json({ error: "Company name is too short." });
    return;
  }

  try {
    const result = await store.createCompany(name, logoDataUrl);
    res.status(201).json({
      company: {
        id: result.id,
        name,
        logoDataUrl,
      },
    });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Company name already exists." });
      return;
    }
    res.status(500).json({ error: "Failed to create company." });
  }
});

app.put("/api/companies/:id", authMiddleware, async (req, res) => {
  const id = parseCompanyId(req.params.id, null);
  if (!id) {
    res.status(400).json({ error: "Invalid company id." });
    return;
  }

  const name = sanitizeText(req.body?.name, 120);
  if (name.length < 2) {
    res.status(400).json({ error: "Company name is too short." });
    return;
  }

  try {
    const existing = await store.getCompanyById(id);
    if (!existing) {
      res.status(404).json({ error: "Company not found." });
      return;
    }

    await store.updateCompanyName(id, name, String(existing.logo_data_url || ""));
    res.json({
      ok: true,
      company: {
        id,
        name,
        logoDataUrl: String(existing.logo_data_url || ""),
      },
    });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Company name already exists." });
      return;
    }
    res.status(500).json({ error: "Failed to update company name." });
  }
});

app.put("/api/companies/:id/logo", authMiddleware, async (req, res) => {
  const id = parseCompanyId(req.params.id, null);
  if (!id) {
    res.status(400).json({ error: "Invalid company id." });
    return;
  }

  const logoDataUrl = sanitizeLogoDataUrl(req.body?.logoDataUrl);
  try {
    const exists = await store.companyExists(id);
    if (!exists) {
      res.status(404).json({ error: "Company not found." });
      return;
    }
    await store.updateCompanyLogo(id, logoDataUrl);
    res.json({ ok: true, companyId: id, logoDataUrl });
  } catch {
    res.status(500).json({ error: "Failed to update company logo." });
  }
});

app.get("/api/settings/designations", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;
  try {
    const rows = await store.listDesignationPresets(companyId);
    res.json({ designations: rows.map(dbRowToDesignation) });
  } catch (error) {
    if (isStoreUnavailableError(error)) {
      res.json({ designations: [], degraded: true });
      return;
    }
    res.status(500).json({ error: "Failed to fetch designation presets." });
  }
});

app.post("/api/settings/designations", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;
  const name = sanitizeText(req.body?.name, 120);
  if (name.length < 2) {
    res.status(400).json({ error: "Designation name is too short." });
    return;
  }

  try {
    const existing = await store.listDesignationPresets(companyId);
    const result = await store.addDesignationPreset(companyId, name, existing.length);
    const rows = await store.listDesignationPresets(companyId);
    const created = rows.find((row) => Number(row.id) === Number(result.id));
    res.status(201).json({ designation: dbRowToDesignation(created) });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Designation already exists." });
      return;
    }
    res.status(500).json({ error: "Failed to add designation." });
  }
});

app.delete("/api/settings/designations/:id", authMiddleware, async (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) {
    res.status(400).json({ error: "Invalid designation id." });
    return;
  }
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;
  try {
    await store.deleteDesignationPreset(companyId, id);
    res.json({ ok: true });
  } catch (error) {
    if (isStoreUnavailableError(error)) {
      res.status(503).json({ error: "Payroll data service is temporarily unreachable. Designation was not deleted." });
      return;
    }
    res.status(500).json({ error: "Failed to delete designation." });
  }
});

app.get("/api/payroll/all", authMiddleware, async (_req, res) => {
  try {
    res.json(await getSystemBackupPayload());
  } catch (error) {
    if (isStoreUnavailableError(error) || startupDegradedError) {
      res.json({
        companies: fallbackCompanies().map((row) => ({
          id: row.id,
          name: row.name,
          logoDataUrl: row.logo_data_url || "",
        })),
        employees: [],
        designations: [],
        entries: [],
        degraded: true,
      });
      return;
    }
    res.status(500).json({ error: "Failed to fetch payroll backup data." });
  }
});

app.get("/api/system-backups/status", authMiddleware, async (_req, res) => {
  try {
    const status = await ensureMonthlyEmergencyBackup();
    res.json(status);
  } catch {
    res.status(500).json({ error: "Failed to fetch monthly backup status." });
  }
});

app.get("/api/system-backups/latest", authMiddleware, async (_req, res) => {
  try {
    const recent = listMonthlyBackupFiles();
    const latest = recent[0];
    if (!latest) {
      res.status(404).json({ error: "No monthly emergency backup found yet." });
      return;
    }
    res.setHeader("Content-Type", "application/json");
    res.setHeader("Content-Disposition", `attachment; filename="${latest.fileName}"`);
    res.sendFile(latest.fullPath);
  } catch {
    res.status(500).json({ error: "Failed to download monthly backup." });
  }
});

app.get("/api/payroll-reports", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  try {
    const rows = await store.listPayrollReportsByCompany(companyId);
    res.json({ reports: rows.map(dbRowToPayrollReport) });
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error("Failed to fetch payroll reports:", error);
    res.status(500).json({ error: "Failed to fetch payroll reports." });
  }
});

app.get("/api/payroll-reports/:id", authMiddleware, async (req, res) => {
  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  const id = Number(req.params.id);
  if (!Number.isInteger(id) || id <= 0) {
    res.status(400).json({ error: "Invalid payroll report id." });
    return;
  }

  try {
    const row = await store.getPayrollReportByIdCompany(id, companyId);
    if (!row) {
      res.status(404).json({ error: "Payroll report not found." });
      return;
    }
    res.json({
      report: dbRowToPayrollReport(row),
      snapshot: parseSnapshotJson(row.snapshot_json),
    });
  } catch {
    res.status(500).json({ error: "Failed to fetch payroll report." });
  }
});

app.post("/api/payroll-reports/check", authMiddleware, async (req, res) => {
  const month = String(req.body?.month || "");
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }

  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;

  try {
    const rows = await store.listPayrollByMonthCompany(month, companyId);
    if (!rows.length) {
      res.status(400).json({ error: "No payroll rows found for this month." });
      return;
    }

    const checkedAt = new Date().toISOString();
    const result = await store.markPayrollChecked(month, companyId, checkedAt);
    const report = await store.getPayrollReportByIdCompany(result.id, companyId);
    res.json({ report: report ? dbRowToPayrollReport(report) : null });
  } catch {
    res.status(500).json({ error: "Failed to mark payroll as checked." });
  }
});

app.post("/api/payroll-reports/generate", authMiddleware, async (req, res) => {
  const month = String(req.body?.month || "");
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }

  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;

  try {
    const ensured = await ensurePayrollCheckedForMonth(month, companyId);
    const snapshot = ensured.snapshot;

    const generatedAt = new Date().toISOString();
    const saved = await store.savePayrollReportSnapshot(month, companyId, {
      checkedAt: ensured.checkedAt,
      generatedAt,
      employeeCount: snapshot.records.length,
      snapshot,
    });
    const report = await store.getPayrollReportByIdCompany(saved.id, companyId);
    res.json({
      report: report ? dbRowToPayrollReport(report) : null,
      snapshot,
    });
  } catch {
    res.status(500).json({ error: "Failed to generate payroll report." });
  }
});

app.post("/api/payroll-reports/generate-entry", authMiddleware, async (req, res) => {
  const month = String(req.body?.month || "");
  const employeeId = sanitizeText(req.body?.employeeId, 60);
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }
  if (employeeId.length < 2) {
    res.status(400).json({ error: "Employee id is required." });
    return;
  }

  const companyId = await resolveCompanyId(req, res, "body");
  if (!companyId) return;

  try {
    const ensured = await ensurePayrollCheckedForMonth(month, companyId);
    const payload = await buildPayrollReportRecord(month, companyId, employeeId);
    if (!payload?.record) {
      res.status(404).json({ error: "Payroll record not found for this employee." });
      return;
    }

    const existing = await store.getPayrollReportByMonthCompany(month, companyId);
    const existingSnapshot = parseSnapshotJson(existing?.snapshot_json);
    const existingRecords = Array.isArray(existingSnapshot.records) ? existingSnapshot.records : [];
    const mergedByEmployee = new Map(existingRecords.map((record) => [String(record.employeeId || ""), record]));
    mergedByEmployee.set(String(payload.record.employeeId || ""), payload.record);
    const records = Array.from(mergedByEmployee.values()).sort((a, b) => Number(a.positionIndex || 0) - Number(b.positionIndex || 0));

    const snapshot = {
      company: payload.company,
      month: payload.month,
      records,
    };
    const generatedAt = new Date().toISOString();
    const saved = await store.savePayrollReportSnapshot(month, companyId, {
      checkedAt: ensured.checkedAt,
      generatedAt,
      employeeCount: records.length,
      snapshot,
    });
    const report = await store.getPayrollReportByIdCompany(saved.id, companyId);
    res.json({
      report: report ? dbRowToPayrollReport(report) : null,
      snapshot,
      generatedEmployeeId: payload.record.employeeId,
    });
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error("Failed to generate payroll report entry:", error);
    if (error?.statusCode) {
      res.status(error.statusCode).json({ error: error.message });
      return;
    }
    if (isStoreUnavailableError(error)) {
      res.status(503).json({ error: "Payroll data service is temporarily unreachable. Payslip generation can continue once the connection is back." });
      return;
    }
    res.status(500).json({ error: "Failed to generate payslip for this employee." });
  }
});

app.post("/api/payroll/restore", authMiddleware, async (req, res) => {
  const companiesPayload = req.body?.companies;
  const employeesPayload = req.body?.employees;
  const designationsPayload = req.body?.designations;
  const entriesPayload = req.body?.entries;
  const months = req.body?.months;
  const isLegacyPayload = months && typeof months === "object" && !Array.isArray(months);
  const isNewPayload = Array.isArray(companiesPayload) && Array.isArray(entriesPayload);

  if (!isLegacyPayload && !isNewPayload) {
    res.status(400).json({ error: "Invalid restore payload." });
    return;
  }

  try {
    await store.clearPayrollEntries();
    if (isNewPayload) {
      await store.clearCompanies();
      await store.clearEmployees();
      await store.clearDesignationPresets();
      for (const rawCompany of companiesPayload) {
        const id = parseCompanyId(rawCompany?.id, null);
        const name = sanitizeText(rawCompany?.name, 120);
        const logoDataUrl = sanitizeLogoDataUrl(rawCompany?.logoDataUrl);
        if (!id || name.length < 2) continue;
        await store.insertCompanyWithId(id, name, logoDataUrl);
      }
      await store.ensureDefaultCompany();

      if (Array.isArray(designationsPayload)) {
        for (const rawDesignation of designationsPayload) {
          const companyId = parseCompanyId(rawDesignation?.companyId, null);
          const name = sanitizeText(rawDesignation?.name, 120);
          if (!companyId || name.length < 2 || !(await store.companyExists(companyId))) continue;
          try {
            // eslint-disable-next-line no-await-in-loop
            await store.addDesignationPreset(companyId, name, Number(rawDesignation?.positionIndex) || 0);
          } catch {
            // Ignore duplicates during restore.
          }
        }
      }

      if (Array.isArray(employeesPayload)) {
        for (const rawEmployee of employeesPayload) {
          const companyId = parseCompanyId(rawEmployee?.companyId, null);
          if (!companyId || !(await store.companyExists(companyId))) continue;
          const employee = normalizeEmployee(rawEmployee, Number(rawEmployee?.positionIndex) || 0);
          if (employee.employeeId.length < 2 || employee.employeeName.length < 2) continue;
          try {
            // eslint-disable-next-line no-await-in-loop
            await store.createEmployee(companyId, employee);
          } catch {
            // Skip invalid/duplicate rows during bulk restore.
          }
        }
      }

      for (const rawEntry of entriesPayload) {
        const month = String(rawEntry?.month || "");
        const companyId = parseCompanyId(rawEntry?.companyId, null);
        if (!isValidMonth(month) || !companyId) {
          continue;
        }
        if (!(await store.companyExists(companyId))) continue;
        const record = normalizeRecord(rawEntry, Number(rawEntry?.positionIndex) || 0);
        await store.insertPayrollRecord(month, companyId, record);
      }
    } else {
      for (const [month, records] of Object.entries(months)) {
        if (!isValidMonth(month) || !Array.isArray(records)) {
          continue;
        }
        let idx = 0;
        for (const rawRecord of records) {
          const record = normalizeRecord(rawRecord, idx);
          await store.insertPayrollRecord(month, 1, record);
          idx += 1;
        }
      }
    }

    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ error: "Failed to restore backup data." });
  }
});

app.get("/api/payroll/:month", authMiddleware, async (req, res) => {
  const { month } = req.params;
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }

  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  try {
    const records = await listMergedPayrollMonthRecords(month, companyId);
    res.json({ records });
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch payroll month data." });
  }
});

app.put("/api/payroll/:month", authMiddleware, async (req, res) => {
  const { month } = req.params;
  const records = req.body?.records;
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }
  if (!Array.isArray(records)) {
    res.status(400).json({ error: "records must be an array." });
    return;
  }

  const companyId = await resolveCompanyId(req, res, "query");
  if (!companyId) return;

  const normalized = records.map((record, index) => normalizeRecord(record, index));

  try {
    await store.deletePayrollByMonthCompany(month, companyId);
    for (const record of normalized) {
      await store.insertPayrollRecord(month, companyId, record);
    }
    res.json({ ok: true, count: normalized.length });
  } catch (error) {
    console.error(`Failed to save payroll month data for ${month} company ${companyId}:`, error);
    res.status(500).json({ error: "Failed to save payroll month data." });
  }
});

app.use("/api", (_req, res) => {
  res.status(404).json({ error: "Not found" });
});

app.get("/{*any}", (_req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

function dbRowToRecord(row) {
  return {
    id: row.id,
    companyId: row.company_id,
    month: row.month,
    employeeId: row.employee_id,
    employeeName: row.employee_name,
    designation: row.designation,
    presentSalary: row.present_salary,
    increment: row.increment,
    oldAdvanceTaken: row.old_advance_taken,
    extraAdvanceAdded: row.extra_advance_added,
    deductionEntered: row.deduction_entered,
    daysAbsent: row.days_absent,
    comment: row.comment,
    positionIndex: row.position_index,
    advanceRemained: advanceRemainedFromPayrollRow(row),
  };
}

function ensureStartupInit() {
  if (!startupInitPromise) {
    startupInitPromise = store.init()
      .then(() => {
        startupReady = true;
        startupState = "ready";
        // eslint-disable-next-line no-console
        console.log(`Routes Payroll cloud startup completed (db: ${store.provider})`);
      })
      .catch((error) => {
        startupDegradedError = error;
        startupState = "degraded";
        // eslint-disable-next-line no-console
        console.error("Database initialized in degraded mode:", error);
      });
  }

  return startupInitPromise;
}

ensureStartupInit();

if (require.main === module) {
  app.listen(PORT, "127.0.0.1", () => {
    // eslint-disable-next-line no-console
    console.log(`Routes Payroll API running at http://127.0.0.1:${PORT} (db: ${store.provider}, startup: ${startupState})`);
  });
}

module.exports = app;
