const express = require("express");
const path = require("path");
const fs = require("fs");
const sqlite3 = require("sqlite3").verbose();
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");

const app = express();
const PORT = Number(process.env.PORT || 5501);
const JWT_SECRET = process.env.JWT_SECRET || "routes_payroll_dev_secret_change_me";
const DB_DIR = path.join(__dirname, "data");
const DB_PATH = path.join(DB_DIR, "payroll.db");

if (!fs.existsSync(DB_DIR)) {
  fs.mkdirSync(DB_DIR, { recursive: true });
}

const db = new sqlite3.Database(DB_PATH);

function run(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function onRun(error) {
      if (error) {
        reject(error);
        return;
      }
      resolve({ lastID: this.lastID, changes: this.changes });
    });
  });
}

function get(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (error, row) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(row || null);
    });
  });
}

function all(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (error, rows) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(rows || []);
    });
  });
}

async function initDb() {
  await run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await run(`
    CREATE TABLE IF NOT EXISTS payroll_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      month TEXT NOT NULL,
      employee_id TEXT NOT NULL,
      employee_name TEXT NOT NULL,
      designation TEXT NOT NULL,
      present_salary REAL NOT NULL DEFAULT 0,
      increment REAL NOT NULL DEFAULT 0,
      old_advance_taken REAL NOT NULL DEFAULT 0,
      extra_advance_added REAL NOT NULL DEFAULT 0,
      deduction_entered REAL NOT NULL DEFAULT 0,
      days_absent REAL NOT NULL DEFAULT 0,
      comment TEXT NOT NULL DEFAULT '',
      position_index INTEGER NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await run(`CREATE INDEX IF NOT EXISTS idx_payroll_month ON payroll_entries(month)`);
  await run(
    `CREATE INDEX IF NOT EXISTS idx_payroll_month_position ON payroll_entries(month, position_index)`
  );
}

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
    const row = await get("SELECT COUNT(*) AS count FROM users");
    res.json({ needsSetup: (row?.count || 0) === 0 });
  } catch (error) {
    res.status(500).json({ error: "Failed to check bootstrap state." });
  }
});

app.post("/api/auth/register", async (req, res) => {
  const username = sanitizeText(req.body?.username, 40);
  const password = String(req.body?.password || "");

  if (username.length < 3 || password.length < 6) {
    res.status(400).json({ error: "Username or password length is invalid." });
    return;
  }

  try {
    const countRow = await get("SELECT COUNT(*) AS count FROM users");
    if ((countRow?.count || 0) > 0) {
      res.status(409).json({ error: "Admin account already exists." });
      return;
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const result = await run(
      "INSERT INTO users (username, password_hash) VALUES (?, ?)",
      [username, passwordHash]
    );

    const token = createToken({ id: result.lastID, username });
    res.status(201).json({ token, user: { id: result.lastID, username } });
  } catch (error) {
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
    const user = await get("SELECT id, username, password_hash FROM users WHERE username = ?", [username]);
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

app.get("/api/payroll/all", authMiddleware, async (_req, res) => {
  try {
    const rows = await all(
      `SELECT id, month, employee_id, employee_name, designation,
              present_salary, increment, old_advance_taken, extra_advance_added,
              deduction_entered, days_absent, comment, position_index
       FROM payroll_entries
       ORDER BY month, position_index, id`
    );

    const months = {};
    for (const row of rows) {
      if (!months[row.month]) {
        months[row.month] = [];
      }
      months[row.month].push(dbRowToRecord(row));
    }

    res.json({ months });
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch payroll backup data." });
  }
});

app.post("/api/payroll/restore", authMiddleware, async (req, res) => {
  const months = req.body?.months;
  if (!months || typeof months !== "object" || Array.isArray(months)) {
    res.status(400).json({ error: "Invalid restore payload." });
    return;
  }

  try {
    await run("BEGIN IMMEDIATE TRANSACTION");
    await run("DELETE FROM payroll_entries");

    for (const [month, records] of Object.entries(months)) {
      if (!isValidMonth(month) || !Array.isArray(records)) {
        continue;
      }
      let idx = 0;
      for (const rawRecord of records) {
        const record = normalizeRecord(rawRecord, idx);
        await insertRecord(month, record);
        idx += 1;
      }
    }

    await run("COMMIT");
    res.json({ ok: true });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
    res.status(500).json({ error: "Failed to restore backup data." });
  }
});

app.get("/api/payroll/:month", authMiddleware, async (req, res) => {
  const { month } = req.params;
  if (!isValidMonth(month)) {
    res.status(400).json({ error: "Invalid month format. Use YYYY-MM." });
    return;
  }

  try {
    const rows = await all(
      `SELECT id, month, employee_id, employee_name, designation,
              present_salary, increment, old_advance_taken, extra_advance_added,
              deduction_entered, days_absent, comment, position_index
       FROM payroll_entries
       WHERE month = ?
       ORDER BY position_index, id`,
      [month]
    );

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

  const normalized = records.map((record, index) => normalizeRecord(record, index));

  try {
    await run("BEGIN IMMEDIATE TRANSACTION");
    await run("DELETE FROM payroll_entries WHERE month = ?", [month]);
    for (const record of normalized) {
      await insertRecord(month, record);
    }
    await run("COMMIT");
    res.json({ ok: true, count: normalized.length });
  } catch (error) {
    await run("ROLLBACK").catch(() => {});
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

function insertRecord(month, record) {
  return run(
    `INSERT INTO payroll_entries (
      month, employee_id, employee_name, designation,
      present_salary, increment, old_advance_taken, extra_advance_added,
      deduction_entered, days_absent, comment, position_index, updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
    [
      month,
      record.employeeId || "",
      record.employeeName || "",
      record.designation || "",
      record.presentSalary,
      record.increment,
      record.oldAdvanceTaken,
      record.extraAdvanceAdded,
      record.deductionEntered,
      record.daysAbsent,
      record.comment || "",
      record.positionIndex,
    ]
  );
}

initDb()
  .then(() => {
    app.listen(PORT, "127.0.0.1", () => {
      // eslint-disable-next-line no-console
      console.log(`Routes Payroll API running at http://127.0.0.1:${PORT}`);
    });
  })
  .catch((error) => {
    // eslint-disable-next-line no-console
    console.error("Failed to initialize database:", error);
    process.exit(1);
  });
