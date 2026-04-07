const express = require("express");
const path = require("path");
const crypto = require("crypto");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const nodemailer = require("nodemailer");
const { createStore } = require("./datastore");

const app = express();
const PORT = Number(process.env.PORT || 5501);
const JWT_SECRET = process.env.JWT_SECRET || "routes_payroll_dev_secret_change_me";
const store = createStore({ baseDir: __dirname });
const RATE_BUCKETS = new Map();

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

async function resolveCompanyId(req, res, source = "query") {
  const rawValue = source === "body" ? req.body?.companyId : req.query?.companyId;
  const companyId = parseCompanyId(rawValue, 1);
  if (!companyId) {
    res.status(400).json({ error: "Invalid company id." });
    return null;
  }

  try {
    const exists = await store.companyExists(companyId);
    if (!exists) {
      res.status(404).json({ error: "Company not found." });
      return null;
    }
    return companyId;
  } catch {
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
    daysAbsent: clamp(toNumber(record.daysAbsent), 0, 30),
    comment: sanitizeText(record.comment, 350),
    positionIndex: Number.isFinite(Number(record.positionIndex)) ? Number(record.positionIndex) : index,
  };
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
}

function normalizeEmployee(employee, index) {
  const rawStatus = String(employee?.status || "working").toLowerCase();
  const status = rawStatus === "leave" || rawStatus === "terminated" ? rawStatus : "working";
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

function createToken(user) {
  return jwt.sign(
    { userId: user.id, username: user.username },
    JWT_SECRET,
    { expiresIn: "12h" }
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
  res.json({ ok: true, time: new Date().toISOString() });
});

app.get("/api/auth/bootstrap", async (_req, res) => {
  try {
    const count = await store.countUsers();
    res.json({ needsSetup: count === 0 });
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
    const userCount = await store.countUsers();
    if (userCount > 0) {
      res.status(409).json({ error: "Admin account already exists." });
      return;
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const result = await store.createUser(username, email, passwordHash);
    await store.updateCompanyName(1, companyName, "");

    const token = createToken({ id: result.id, username });
    res.status(201).json({ token, user: { id: result.id, username, email } });
  } catch (error) {
    if (error?.code === "unique") {
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

    const isValid = await bcrypt.compare(password, user.password_hash);
    if (!isValid) {
      res.status(401).json({ error: "Invalid username or password." });
      return;
    }

    const token = createToken(user);
    res.json({ token, user: { id: user.id, username: user.username } });
  } catch (error) {
    res.status(500).json({ error: "Login failed." });
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
    const company = await store.getCompanyById(1);
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

    const resetToken = crypto.randomBytes(32).toString("hex");
    const expiresAt = new Date(Date.now() + 15 * 60 * 1000).toISOString();
    await store.setUserResetCodeById(user.id, hashResetCode(resetToken), expiresAt);
    const appBaseUrl = String(process.env.APP_BASE_URL || `http://127.0.0.1:${PORT}`).replace(/\/+$/, "");
    const resetLink = `${appBaseUrl}/reset-password.html?token=${encodeURIComponent(resetToken)}`;
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
    await store.updateUserPasswordById(user.id, passwordHash);
    await store.clearUserResetCodeById(user.id);
    res.json({ ok: true, message: "Password reset successful." });
  } catch {
    res.status(500).json({ error: "Password reset failed." });
  }
});

app.get("/api/auth/me", authMiddleware, (req, res) => {
  res.json({
    user: {
      id: req.user.userId,
      username: req.user.username,
    },
  });
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
    const user = await store.findUserByIdentifier(req.user.username);
    if (!user) {
      res.status(404).json({ error: "User not found." });
      return;
    }
    const sameAsCurrent = await bcrypt.compare(newPassword, user.password_hash || "");
    if (sameAsCurrent) {
      res.status(400).json({ error: "New password must be different from current password." });
      return;
    }

    const passwordHash = await bcrypt.hash(newPassword, 12);
    await store.updateUserPasswordById(req.user.userId, passwordHash);
    res.json({ ok: true });
  } catch {
    res.status(500).json({ error: "Failed to update password." });
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
  } catch {
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
  } catch {
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
  } catch {
    res.status(500).json({ error: "Failed to delete designation." });
  }
});

app.get("/api/payroll/all", authMiddleware, async (_req, res) => {
  try {
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

    res.json({
      companies: companies.map((company) => ({
        id: company.id,
        name: company.name,
        logoDataUrl: company.logo_data_url || "",
      })),
      employees,
      designations,
      entries: rows.map(dbRowToRecord),
    });
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch payroll backup data." });
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
    const rows = await store.listPayrollByMonthCompany(month, companyId);
    const employees = await store.listEmployeesByCompany(companyId);
    const prevMonth = previousMonthIso(month);
    const previousRows = prevMonth ? await store.listPayrollByMonthCompany(prevMonth, companyId) : [];
    const previousByEmployeeId = new Map(previousRows.map((row) => [String(row.employee_id || ""), row]));
    const byEmployeeId = new Map(rows.map((row) => [String(row.employee_id || ""), row]));
    const records = [];

    employees.forEach((employee, index) => {
      if (String(employee.status || "").toLowerCase() === "terminated") {
        return;
      }
      const linked = byEmployeeId.get(String(employee.employee_id || "")) || null;
      const previous = previousByEmployeeId.get(String(employee.employee_id || "")) || null;
      const carriedAdvance = previous ? advanceRemainedFromPayrollRow(previous) : null;
      records.push({
        id: linked?.id || null,
        companyId,
        month,
        employeeId: employee.employee_id,
        employeeName: employee.employee_name,
        designation: employee.designation || "",
        presentSalary: Number(employee.base_salary || 0),
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
      if (String(extra.employee_status || "").toLowerCase() === "terminated") {
        continue;
      }
      records.push({
        ...dbRowToRecord(extra),
        employeeStatus: "working",
        leaveFrom: "",
        leaveTo: "",
        terminatedOn: "",
      });
    }

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

store.init()
  .then(() => {
    app.listen(PORT, "127.0.0.1", () => {
      // eslint-disable-next-line no-console
      console.log(`Routes Payroll API running at http://127.0.0.1:${PORT} (db: ${store.provider})`);
    });
  })
  .catch((error) => {
    // eslint-disable-next-line no-console
    console.error("Failed to initialize database:", error);
    process.exit(1);
  });
