const fs = require("fs");
const path = require("path");
const sqlite3 = require("sqlite3").verbose();
const admin = require("firebase-admin");

function createStore({ baseDir }) {
  const provider = String(process.env.DB_PROVIDER || "sqlite").toLowerCase();
  if (provider === "firebase") {
    return createFirebaseStore();
  }
  return createSqliteStore(baseDir);
}

function createSqliteStore(baseDir) {
  const dbDir = path.join(baseDir, "data");
  const dbPath = path.join(dbDir, "payroll.db");

  if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
  }

  const db = new sqlite3.Database(dbPath);

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

  return {
    provider: "sqlite",
    async init() {
      await run(`
        CREATE TABLE IF NOT EXISTS users (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          username TEXT NOT NULL UNIQUE,
          username_lc TEXT NOT NULL DEFAULT '',
          email TEXT NOT NULL DEFAULT '',
          email_lc TEXT NOT NULL DEFAULT '',
          password_hash TEXT NOT NULL,
          reset_code_hash TEXT NOT NULL DEFAULT '',
          reset_code_expires_at TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
      `);

      await run(`
        CREATE TABLE IF NOT EXISTS companies (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL UNIQUE,
          logo_data_url TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
      `);

      await run(`
        CREATE TABLE IF NOT EXISTS payroll_entries (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER NOT NULL DEFAULT 1,
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

      await run(
        "INSERT OR IGNORE INTO companies (id, name, logo_data_url) VALUES (1, 'Routes Payroll', '')"
      );

      try {
        await run("ALTER TABLE payroll_entries ADD COLUMN company_id INTEGER NOT NULL DEFAULT 1");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN username_lc TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN email TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN email_lc TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN reset_code_hash TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN reset_code_expires_at TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      await run("UPDATE users SET username_lc = lower(username) WHERE username_lc = ''");
      await run("UPDATE users SET email_lc = lower(email) WHERE email_lc = ''");
      await run("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_lc_unique ON users(username_lc)");
      await run(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email_lc_unique ON users(email_lc) WHERE email_lc <> ''"
      );

      await run("CREATE INDEX IF NOT EXISTS idx_payroll_month ON payroll_entries(month)");
      await run(
        "CREATE INDEX IF NOT EXISTS idx_payroll_month_position ON payroll_entries(month, position_index)"
      );
      await run(
        "CREATE INDEX IF NOT EXISTS idx_payroll_company_month_position ON payroll_entries(company_id, month, position_index)"
      );
    },

    async countUsers() {
      const row = await get("SELECT COUNT(*) AS count FROM users");
      return Number(row?.count || 0);
    },

    async findUserByIdentifier(identifier) {
      const normalized = String(identifier || "").trim().toLowerCase();
      return get(
        `SELECT id, username, email, password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE username_lc = ? OR email_lc = ?
         LIMIT 1`,
        [normalized, normalized]
      );
    },

    async findUserByEmail(email) {
      const normalized = String(email || "").trim().toLowerCase();
      return get(
        `SELECT id, username, email, password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE email_lc = ?
         LIMIT 1`,
        [normalized]
      );
    },

    async createUser(username, email, passwordHash) {
      const result = await run(
        "INSERT INTO users (username, username_lc, email, email_lc, password_hash) VALUES (?, ?, ?, ?, ?)",
        [username, String(username).toLowerCase(), email, String(email).toLowerCase(), passwordHash]
      );
      return { id: result.lastID };
    },

    async updateUserPasswordByEmail(email, passwordHash) {
      await run("UPDATE users SET password_hash = ? WHERE email_lc = ?", [
        passwordHash,
        String(email || "").trim().toLowerCase(),
      ]);
    },

    async updateUserPasswordById(userId, passwordHash) {
      await run("UPDATE users SET password_hash = ? WHERE id = ?", [passwordHash, userId]);
    },

    async setUserResetCodeById(userId, resetCodeHash, expiresAtIso) {
      await run(
        "UPDATE users SET reset_code_hash = ?, reset_code_expires_at = ? WHERE id = ?",
        [resetCodeHash, expiresAtIso, userId]
      );
    },

    async clearUserResetCodeById(userId) {
      await run(
        "UPDATE users SET reset_code_hash = '', reset_code_expires_at = '' WHERE id = ?",
        [userId]
      );
    },

    async findUserByResetCodeHash(resetCodeHash) {
      return get(
        `SELECT id, username, email, password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE reset_code_hash = ?
         LIMIT 1`,
        [String(resetCodeHash || "")]
      );
    },

    async updateCompanyName(id, name, logoDataUrl = "") {
      await run("UPDATE companies SET name = ?, logo_data_url = ? WHERE id = ?", [name, logoDataUrl, id]);
    },

    async companyExists(id) {
      const row = await get("SELECT id FROM companies WHERE id = ?", [id]);
      return Boolean(row);
    },

    async listCompanies() {
      return all(
        `SELECT id, name, logo_data_url
         FROM companies
         ORDER BY name, id`
      );
    },

    async listCompaniesById() {
      return all(
        `SELECT id, name, logo_data_url
         FROM companies
         ORDER BY id`
      );
    },

    async getCompanyById(id) {
      return get(
        `SELECT id, name, logo_data_url
         FROM companies
         WHERE id = ?`,
        [id]
      );
    },

    async createCompany(name, logoDataUrl) {
      try {
        const result = await run("INSERT INTO companies (name, logo_data_url) VALUES (?, ?)", [
          name,
          logoDataUrl,
        ]);
        return { id: result.lastID };
      } catch (error) {
        if (String(error?.message || "").toLowerCase().includes("unique")) {
          const conflict = new Error("Company name already exists.");
          conflict.code = "unique";
          throw conflict;
        }
        throw error;
      }
    },

    async clearPayrollEntries() {
      await run("DELETE FROM payroll_entries");
    },

    async clearCompanies() {
      await run("DELETE FROM companies");
    },

    async insertCompanyWithId(id, name, logoDataUrl) {
      await run("INSERT INTO companies (id, name, logo_data_url) VALUES (?, ?, ?)", [id, name, logoDataUrl]);
    },

    async ensureDefaultCompany() {
      await run(
        "INSERT OR IGNORE INTO companies (id, name, logo_data_url) VALUES (1, 'Routes Payroll', '')"
      );
    },

    async listPayrollEntriesAll() {
      return all(
        `SELECT id, company_id, month, employee_id, employee_name, designation,
                present_salary, increment, old_advance_taken, extra_advance_added,
                deduction_entered, days_absent, comment, position_index
         FROM payroll_entries
         ORDER BY company_id, month, position_index, id`
      );
    },

    async listPayrollByMonthCompany(month, companyId) {
      return all(
        `SELECT id, company_id, month, employee_id, employee_name, designation,
                present_salary, increment, old_advance_taken, extra_advance_added,
                deduction_entered, days_absent, comment, position_index
         FROM payroll_entries
         WHERE month = ? AND company_id = ?
         ORDER BY position_index, id`,
        [month, companyId]
      );
    },

    async deletePayrollByMonthCompany(month, companyId) {
      await run("DELETE FROM payroll_entries WHERE month = ? AND company_id = ?", [month, companyId]);
    },

    async insertPayrollRecord(month, companyId, record) {
      await run(
        `INSERT INTO payroll_entries (
          company_id, month, employee_id, employee_name, designation,
          present_salary, increment, old_advance_taken, extra_advance_added,
          deduction_entered, days_absent, comment, position_index, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
        [
          companyId,
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
    },
  };
}

function createFirebaseStore() {
  const serviceAccountJson = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || "";
  const firebaseProjectId = process.env.FIREBASE_PROJECT_ID || "";

  if (!admin.apps.length) {
    if (serviceAccountJson) {
      const credentials = JSON.parse(serviceAccountJson);
      admin.initializeApp({
        credential: admin.credential.cert(credentials),
        projectId: firebaseProjectId || credentials.project_id,
      });
    } else {
      admin.initializeApp(
        firebaseProjectId
          ? { credential: admin.credential.applicationDefault(), projectId: firebaseProjectId }
          : { credential: admin.credential.applicationDefault() }
      );
    }
  }

  const firestore = admin.firestore();
  const nowIso = () => new Date().toISOString();

  function col(name) {
    return firestore.collection(name);
  }

  async function getCounter(kind) {
    const ref = firestore.collection("metadata").doc("counters");
    return firestore.runTransaction(async (tx) => {
      const snap = await tx.get(ref);
      const data = snap.exists ? snap.data() || {} : {};
      const current = Number(data[kind] || 0);
      const next = current + 1;
      tx.set(ref, { ...data, [kind]: next }, { merge: true });
      return next;
    });
  }

  async function batchDeleteByCollection(name) {
    const snapshot = await col(name).get();
    let batch = firestore.batch();
    let count = 0;

    for (const docSnap of snapshot.docs) {
      batch.delete(docSnap.ref);
      count += 1;
      if (count % 450 === 0) {
        await batch.commit();
        batch = firestore.batch();
      }
    }

    if (count % 450 !== 0) {
      await batch.commit();
    }
  }

  async function findByField(name, field, value) {
    const snapshot = await col(name).where(field, "==", value).limit(1).get();
    if (snapshot.empty) return null;
    return snapshot.docs[0].data() || null;
  }

  return {
    provider: "firebase",

    async init() {
      const defaultCompanyRef = col("companies").doc("1");
      const snapshot = await defaultCompanyRef.get();
      if (!snapshot.exists) {
        await defaultCompanyRef.set({
          id: 1,
          name: "Routes Payroll",
          name_lc: "routes payroll",
          logo_data_url: "",
          created_at: nowIso(),
        });
      }

      await firestore.collection("metadata").doc("counters").set(
        {
          companies: admin.firestore.FieldValue.increment(0),
          users: admin.firestore.FieldValue.increment(0),
          payroll_entries: admin.firestore.FieldValue.increment(0),
        },
        { merge: true }
      );
    },

    async countUsers() {
      const snapshot = await col("users").count().get();
      return Number(snapshot.data()?.count || 0);
    },

    async findUserByIdentifier(identifier) {
      const normalized = String(identifier || "").trim().toLowerCase();
      const user = await findByField("users", "username_lc", normalized);
      if (user) return user;
      return findByField("users", "email_lc", normalized);
    },

    async findUserByEmail(email) {
      const user = await findByField("users", "email_lc", String(email || "").toLowerCase());
      return user || null;
    },

    async createUser(username, email, passwordHash) {
      const usernameLc = String(username || "").toLowerCase();
      const emailLc = String(email || "").toLowerCase();
      const existing = await findByField("users", "username_lc", usernameLc);
      if (existing) {
        const conflict = new Error("Username already exists.");
        conflict.code = "unique";
        throw conflict;
      }
      const existingEmail = await findByField("users", "email_lc", emailLc);
      if (existingEmail) {
        const conflict = new Error("Email already exists.");
        conflict.code = "unique";
        throw conflict;
      }

      const id = await getCounter("users");
      await col("users").doc(String(id)).set({
        id,
        username,
        username_lc: usernameLc,
        email,
        email_lc: emailLc,
        password_hash: passwordHash,
        created_at: nowIso(),
      });
      return { id };
    },

    async updateUserPasswordByEmail(email, passwordHash) {
      const user = await findByField("users", "email_lc", String(email || "").toLowerCase());
      if (!user || !user.id) return;
      await col("users").doc(String(user.id)).set(
        {
          password_hash: passwordHash,
        },
        { merge: true }
      );
    },

    async updateUserPasswordById(userId, passwordHash) {
      await col("users").doc(String(userId)).set(
        {
          password_hash: passwordHash,
        },
        { merge: true }
      );
    },

    async setUserResetCodeById(userId, resetCodeHash, expiresAtIso) {
      await col("users").doc(String(userId)).set(
        {
          reset_code_hash: resetCodeHash,
          reset_code_expires_at: expiresAtIso,
        },
        { merge: true }
      );
    },

    async clearUserResetCodeById(userId) {
      await col("users").doc(String(userId)).set(
        {
          reset_code_hash: "",
          reset_code_expires_at: "",
        },
        { merge: true }
      );
    },

    async findUserByResetCodeHash(resetCodeHash) {
      const user = await findByField("users", "reset_code_hash", String(resetCodeHash || ""));
      return user || null;
    },

    async updateCompanyName(id, name, logoDataUrl = "") {
      await col("companies").doc(String(id)).set(
        {
          id,
          name,
          name_lc: String(name || "").toLowerCase(),
          logo_data_url: logoDataUrl,
        },
        { merge: true }
      );
    },

    async companyExists(id) {
      const snapshot = await col("companies").doc(String(id)).get();
      return snapshot.exists;
    },

    async listCompanies() {
      const snapshot = await col("companies").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .sort((a, b) => {
          const nameCmp = String(a.name || "").localeCompare(String(b.name || ""));
          if (nameCmp !== 0) return nameCmp;
          return Number(a.id || 0) - Number(b.id || 0);
        });
    },

    async listCompaniesById() {
      const snapshot = await col("companies").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .sort((a, b) => Number(a.id || 0) - Number(b.id || 0));
    },

    async getCompanyById(id) {
      const snapshot = await col("companies").doc(String(id)).get();
      return snapshot.exists ? (snapshot.data() || null) : null;
    },

    async createCompany(name, logoDataUrl) {
      const nameLc = String(name || "").toLowerCase();
      const existing = await findByField("companies", "name_lc", nameLc);
      if (existing) {
        const conflict = new Error("Company name already exists.");
        conflict.code = "unique";
        throw conflict;
      }

      const id = await getCounter("companies");
      await col("companies").doc(String(id)).set({
        id,
        name,
        name_lc: nameLc,
        logo_data_url: logoDataUrl,
        created_at: nowIso(),
      });
      return { id };
    },

    async clearPayrollEntries() {
      await batchDeleteByCollection("payroll_entries");
    },

    async clearCompanies() {
      await batchDeleteByCollection("companies");
    },

    async insertCompanyWithId(id, name, logoDataUrl) {
      await col("companies").doc(String(id)).set({
        id,
        name,
        name_lc: String(name || "").toLowerCase(),
        logo_data_url: logoDataUrl,
        created_at: nowIso(),
      });

      await firestore.collection("metadata").doc("counters").set(
        { companies: admin.firestore.FieldValue.increment(0) },
        { merge: true }
      );
    },

    async ensureDefaultCompany() {
      const defaultRef = col("companies").doc("1");
      const snapshot = await defaultRef.get();
      if (!snapshot.exists) {
        await defaultRef.set({
          id: 1,
          name: "Routes Payroll",
          name_lc: "routes payroll",
          logo_data_url: "",
          created_at: nowIso(),
        });
      }
    },

    async listPayrollEntriesAll() {
      const snapshot = await col("payroll_entries").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .sort((a, b) => {
          const aCompany = Number(a.company_id || 0);
          const bCompany = Number(b.company_id || 0);
          if (aCompany !== bCompany) return aCompany - bCompany;
          const monthCmp = String(a.month || "").localeCompare(String(b.month || ""));
          if (monthCmp !== 0) return monthCmp;
          const posCmp = Number(a.position_index || 0) - Number(b.position_index || 0);
          if (posCmp !== 0) return posCmp;
          return Number(a.id || 0) - Number(b.id || 0);
        });
    },

    async listPayrollByMonthCompany(month, companyId) {
      const snapshot = await col("payroll_entries").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .filter((row) => row.month === month && Number(row.company_id) === Number(companyId))
        .sort((a, b) => {
          const posCmp = Number(a.position_index || 0) - Number(b.position_index || 0);
          if (posCmp !== 0) return posCmp;
          return Number(a.id || 0) - Number(b.id || 0);
        });
    },

    async deletePayrollByMonthCompany(month, companyId) {
      const rows = await this.listPayrollByMonthCompany(month, companyId);
      if (rows.length === 0) return;

      let batch = firestore.batch();
      let count = 0;
      for (const row of rows) {
        batch.delete(col("payroll_entries").doc(String(row.id)));
        count += 1;
        if (count % 450 === 0) {
          await batch.commit();
          batch = firestore.batch();
        }
      }
      if (count % 450 !== 0) {
        await batch.commit();
      }
    },

    async insertPayrollRecord(month, companyId, record) {
      const id = await getCounter("payroll_entries");
      await col("payroll_entries").doc(String(id)).set({
        id,
        company_id: companyId,
        month,
        employee_id: record.employeeId || "",
        employee_name: record.employeeName || "",
        designation: record.designation || "",
        present_salary: Number(record.presentSalary || 0),
        increment: Number(record.increment || 0),
        old_advance_taken: Number(record.oldAdvanceTaken || 0),
        extra_advance_added: Number(record.extraAdvanceAdded || 0),
        deduction_entered: Number(record.deductionEntered || 0),
        days_absent: Number(record.daysAbsent || 0),
        comment: record.comment || "",
        position_index: Number(record.positionIndex || 0),
        updated_at: nowIso(),
      });
    },
  };
}

module.exports = {
  createStore,
};
