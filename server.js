const express = require("express");
const path = require("path");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const { createStore } = require("./datastore");

const app = express();
const PORT = Number(process.env.PORT || 5501);
const JWT_SECRET = process.env.JWT_SECRET || "routes_payroll_dev_secret_change_me";
const store = createStore({ baseDir: __dirname });

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
  const username = sanitizeText(req.body?.username, 40);
  const password = String(req.body?.password || "");

  if (companyName.length < 2 || username.length < 3 || password.length < 6) {
    res.status(400).json({ error: "Company name, username, or password is invalid." });
    return;
  }

  try {
    const userCount = await store.countUsers();
    if (userCount > 0) {
      res.status(409).json({ error: "Admin account already exists." });
      return;
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const result = await store.createUser(username, passwordHash);
    await store.updateCompanyName(1, companyName, "");

    const token = createToken({ id: result.lastID, username });
    res.status(201).json({ token, user: { id: result.id, username } });
  } catch (error) {
    if (error?.code === "unique") {
      res.status(409).json({ error: "Admin account already exists." });
      return;
    }
    res.status(500).json({ error: "Failed to register admin account." });
  }
});

app.post("/api/auth/login", async (req, res) => {
  const username = sanitizeText(req.body?.username, 40);
  const password = String(req.body?.password || "");

  if (!username || !password) {
    res.status(400).json({ error: "Username and password are required." });
    return;
  }

  try {
    const user = await store.findUserByUsername(username);
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

app.get("/api/auth/me", authMiddleware, (req, res) => {
  res.json({
    user: {
      id: req.user.userId,
      username: req.user.username,
    },
  });
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

app.get("/api/payroll/all", authMiddleware, async (_req, res) => {
  try {
    const companies = await store.listCompaniesById();
    const rows = await store.listPayrollEntriesAll();

    res.json({
      companies: companies.map((company) => ({
        id: company.id,
        name: company.name,
        logoDataUrl: company.logo_data_url || "",
      })),
      entries: rows.map(dbRowToRecord),
    });
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch payroll backup data." });
  }
});

app.post("/api/payroll/restore", authMiddleware, async (req, res) => {
  const companiesPayload = req.body?.companies;
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
      for (const rawCompany of companiesPayload) {
        const id = parseCompanyId(rawCompany?.id, null);
        const name = sanitizeText(rawCompany?.name, 120);
        const logoDataUrl = sanitizeLogoDataUrl(rawCompany?.logoDataUrl);
        if (!id || name.length < 2) continue;
        await store.insertCompanyWithId(id, name, logoDataUrl);
      }
      await store.ensureDefaultCompany();

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

    res.json({ records: rows.map(dbRowToRecord) });
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
