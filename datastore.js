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

      await run(`
        CREATE TABLE IF NOT EXISTS designation_presets (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER NOT NULL,
          name TEXT NOT NULL,
          name_lc TEXT NOT NULL,
          position_index INTEGER NOT NULL DEFAULT 0,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
      `);

      await run(`
        CREATE TABLE IF NOT EXISTS employees (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER NOT NULL DEFAULT 1,
          employee_id TEXT NOT NULL,
          employee_name TEXT NOT NULL,
          joining_date TEXT NOT NULL DEFAULT '',
          birth_date TEXT NOT NULL DEFAULT '',
          base_salary REAL NOT NULL DEFAULT 0,
          opening_advance REAL NOT NULL DEFAULT 0,
          designation TEXT NOT NULL DEFAULT '',
          mobile_number TEXT NOT NULL DEFAULT '',
          status TEXT NOT NULL DEFAULT 'working',
          leave_from TEXT NOT NULL DEFAULT '',
          leave_to TEXT NOT NULL DEFAULT '',
          terminated_on TEXT NOT NULL DEFAULT '',
          notes TEXT NOT NULL DEFAULT '',
          position_index INTEGER NOT NULL DEFAULT 0,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
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
        await run("ALTER TABLE employees ADD COLUMN opening_advance REAL NOT NULL DEFAULT 0");
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
      await run(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_designation_company_name_lc_unique ON designation_presets(company_id, name_lc)"
      );
      await run(
        "CREATE INDEX IF NOT EXISTS idx_designation_company_position ON designation_presets(company_id, position_index)"
      );
      await run(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_company_employee_id_unique ON employees(company_id, employee_id)"
      );
      await run(
        "CREATE INDEX IF NOT EXISTS idx_employees_company_position ON employees(company_id, position_index)"
      );

      // Backfill employee master from payroll rows for older installs.
      await run(`
        INSERT OR IGNORE INTO employees (
          company_id, employee_id, employee_name, joining_date, base_salary, designation, status, position_index
        )
        SELECT
          p.company_id,
          p.employee_id,
          p.employee_name,
          date('now'),
          p.present_salary,
          p.designation,
          'working',
          MIN(p.position_index)
        FROM payroll_entries p
        WHERE trim(p.employee_id) <> ''
        GROUP BY p.company_id, p.employee_id
      `);

      await this.ensureDefaultDesignationPresets(1);
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

    async updateCompanyLogo(id, logoDataUrl = "") {
      await run("UPDATE companies SET logo_data_url = ? WHERE id = ?", [logoDataUrl, id]);
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

    async listDesignationPresets(companyId) {
      return all(
        `SELECT id, company_id, name, name_lc, position_index
         FROM designation_presets
         WHERE company_id = ?
         ORDER BY position_index, id`,
        [companyId]
      );
    },

    async addDesignationPreset(companyId, name, positionIndex) {
      const nameLc = String(name || "").trim().toLowerCase();
      try {
        const result = await run(
          `INSERT INTO designation_presets (company_id, name, name_lc, position_index)
           VALUES (?, ?, ?, ?)`,
          [companyId, name, nameLc, Number(positionIndex || 0)]
        );
        return { id: result.lastID };
      } catch (error) {
        if (String(error?.message || "").toLowerCase().includes("unique")) {
          const conflict = new Error("Designation already exists.");
          conflict.code = "unique";
          throw conflict;
        }
        throw error;
      }
    },

    async deleteDesignationPreset(companyId, id) {
      await run("DELETE FROM designation_presets WHERE id = ? AND company_id = ?", [id, companyId]);
    },

    async ensureDefaultDesignationPresets(companyId) {
      const defaults = ["Manager", "Chef", "Accountant", "Supervisor", "Staff"];
      let idx = 0;
      // eslint-disable-next-line no-restricted-syntax
      for (const name of defaults) {
        try {
          // eslint-disable-next-line no-await-in-loop
          await this.addDesignationPreset(companyId, name, idx);
        } catch (error) {
          if (error?.code !== "unique") throw error;
        }
        idx += 1;
      }
    },

    async listEmployeesByCompany(companyId) {
      return all(
        `SELECT id, company_id, employee_id, employee_name, joining_date, birth_date,
                base_salary, opening_advance, designation, mobile_number, status, leave_from, leave_to,
                terminated_on, notes, position_index
         FROM employees
         WHERE company_id = ?
         ORDER BY position_index, id`,
        [companyId]
      );
    },

    async getEmployeeByIdCompany(id, companyId) {
      return get(
        `SELECT id, company_id, employee_id, employee_name, joining_date, birth_date,
                base_salary, opening_advance, designation, mobile_number, status, leave_from, leave_to,
                terminated_on, notes, position_index
         FROM employees
         WHERE id = ? AND company_id = ?`,
        [id, companyId]
      );
    },

    async createEmployee(companyId, employee) {
      try {
        const result = await run(
          `INSERT INTO employees (
            company_id, employee_id, employee_name, joining_date, birth_date, base_salary,
            opening_advance, designation, mobile_number, status, leave_from, leave_to, terminated_on, notes,
            position_index, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
          [
            companyId,
            employee.employeeId,
            employee.employeeName,
            employee.joiningDate,
            employee.birthDate,
            employee.baseSalary,
            employee.openingAdvance,
            employee.designation,
            employee.mobileNumber,
            employee.status,
            employee.leaveFrom,
            employee.leaveTo,
            employee.terminatedOn,
            employee.notes,
            employee.positionIndex,
          ]
        );
        return { id: result.lastID };
      } catch (error) {
        if (String(error?.message || "").toLowerCase().includes("unique")) {
          const conflict = new Error("Employee ID already exists for this company.");
          conflict.code = "unique";
          throw conflict;
        }
        throw error;
      }
    },

    async updateEmployee(companyId, id, employee) {
      try {
        await run(
          `UPDATE employees
           SET employee_id = ?, employee_name = ?, joining_date = ?, birth_date = ?, base_salary = ?,
               opening_advance = ?, designation = ?, mobile_number = ?, status = ?, leave_from = ?, leave_to = ?,
               terminated_on = ?, notes = ?, position_index = ?, updated_at = CURRENT_TIMESTAMP
           WHERE id = ? AND company_id = ?`,
          [
            employee.employeeId,
            employee.employeeName,
            employee.joiningDate,
            employee.birthDate,
            employee.baseSalary,
            employee.openingAdvance,
            employee.designation,
            employee.mobileNumber,
            employee.status,
            employee.leaveFrom,
            employee.leaveTo,
            employee.terminatedOn,
            employee.notes,
            employee.positionIndex,
            id,
            companyId,
          ]
        );
      } catch (error) {
        if (String(error?.message || "").toLowerCase().includes("unique")) {
          const conflict = new Error("Employee ID already exists for this company.");
          conflict.code = "unique";
          throw conflict;
        }
        throw error;
      }
    },

    async deleteEmployeeByIdCompany(id, companyId) {
      await run("DELETE FROM employees WHERE id = ? AND company_id = ?", [id, companyId]);
    },

    async clearPayrollEntries() {
      await run("DELETE FROM payroll_entries");
    },

    async clearEmployees() {
      await run("DELETE FROM employees");
    },

    async clearDesignationPresets() {
      await run("DELETE FROM designation_presets");
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
          employees: admin.firestore.FieldValue.increment(0),
          designation_presets: admin.firestore.FieldValue.increment(0),
        },
        { merge: true }
      );

      await this.ensureDefaultDesignationPresets(1);
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

    async updateCompanyLogo(id, logoDataUrl = "") {
      await col("companies").doc(String(id)).set(
        {
          id,
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
      await this.ensureDefaultDesignationPresets(id);
      return { id };
    },

    async listDesignationPresets(companyId) {
      const snapshot = await col("designation_presets").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .filter((row) => Number(row.company_id) === Number(companyId))
        .sort((a, b) => {
          const posCmp = Number(a.position_index || 0) - Number(b.position_index || 0);
          if (posCmp !== 0) return posCmp;
          return Number(a.id || 0) - Number(b.id || 0);
        });
    },

    async addDesignationPreset(companyId, name, positionIndex) {
      const nameLc = String(name || "").trim().toLowerCase();
      const existing = await this.listDesignationPresets(companyId);
      if (existing.some((row) => String(row.name_lc || "") === nameLc)) {
        const conflict = new Error("Designation already exists.");
        conflict.code = "unique";
        throw conflict;
      }

      const id = await getCounter("designation_presets");
      await col("designation_presets").doc(String(id)).set({
        id,
        company_id: companyId,
        name,
        name_lc: nameLc,
        position_index: Number(positionIndex || 0),
        created_at: nowIso(),
      });
      return { id };
    },

    async deleteDesignationPreset(companyId, id) {
      const snapshot = await col("designation_presets").doc(String(id)).get();
      if (!snapshot.exists) return;
      const data = snapshot.data() || {};
      if (Number(data.company_id) !== Number(companyId)) return;
      await col("designation_presets").doc(String(id)).delete();
    },

    async ensureDefaultDesignationPresets(companyId) {
      const defaults = ["Manager", "Chef", "Accountant", "Supervisor", "Staff"];
      let idx = 0;
      // eslint-disable-next-line no-restricted-syntax
      for (const name of defaults) {
        try {
          // eslint-disable-next-line no-await-in-loop
          await this.addDesignationPreset(companyId, name, idx);
        } catch (error) {
          if (error?.code !== "unique") throw error;
        }
        idx += 1;
      }
    },

    async listEmployeesByCompany(companyId) {
      const snapshot = await col("employees").get();
      return snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .filter((row) => Number(row.company_id) === Number(companyId))
        .sort((a, b) => {
          const posCmp = Number(a.position_index || 0) - Number(b.position_index || 0);
          if (posCmp !== 0) return posCmp;
          return Number(a.id || 0) - Number(b.id || 0);
        });
    },

    async getEmployeeByIdCompany(id, companyId) {
      const snap = await col("employees").doc(String(id)).get();
      if (!snap.exists) return null;
      const row = snap.data() || {};
      if (Number(row.company_id) !== Number(companyId)) return null;
      return row;
    },

    async createEmployee(companyId, employee) {
      const existing = await this.listEmployeesByCompany(companyId);
      if (existing.some((row) => String(row.employee_id || "") === String(employee.employeeId || ""))) {
        const conflict = new Error("Employee ID already exists for this company.");
        conflict.code = "unique";
        throw conflict;
      }

      const id = await getCounter("employees");
      await col("employees").doc(String(id)).set({
        id,
        company_id: companyId,
        employee_id: employee.employeeId,
        employee_name: employee.employeeName,
        joining_date: employee.joiningDate,
        birth_date: employee.birthDate,
        base_salary: Number(employee.baseSalary || 0),
        opening_advance: Number(employee.openingAdvance || 0),
        designation: employee.designation,
        mobile_number: employee.mobileNumber,
        status: employee.status,
        leave_from: employee.leaveFrom,
        leave_to: employee.leaveTo,
        terminated_on: employee.terminatedOn,
        notes: employee.notes,
        position_index: Number(employee.positionIndex || 0),
        created_at: nowIso(),
        updated_at: nowIso(),
      });
      return { id };
    },

    async updateEmployee(companyId, id, employee) {
      const existing = await this.listEmployeesByCompany(companyId);
      if (
        existing.some((row) => Number(row.id) !== Number(id) && String(row.employee_id || "") === String(employee.employeeId || ""))
      ) {
        const conflict = new Error("Employee ID already exists for this company.");
        conflict.code = "unique";
        throw conflict;
      }

      await col("employees").doc(String(id)).set(
        {
          id: Number(id),
          company_id: companyId,
          employee_id: employee.employeeId,
          employee_name: employee.employeeName,
          joining_date: employee.joiningDate,
          birth_date: employee.birthDate,
          base_salary: Number(employee.baseSalary || 0),
          opening_advance: Number(employee.openingAdvance || 0),
          designation: employee.designation,
          mobile_number: employee.mobileNumber,
          status: employee.status,
          leave_from: employee.leaveFrom,
          leave_to: employee.leaveTo,
          terminated_on: employee.terminatedOn,
          notes: employee.notes,
          position_index: Number(employee.positionIndex || 0),
          updated_at: nowIso(),
        },
        { merge: true }
      );
    },

    async deleteEmployeeByIdCompany(id, companyId) {
      const row = await this.getEmployeeByIdCompany(id, companyId);
      if (!row) return;
      await col("employees").doc(String(id)).delete();
    },

    async clearPayrollEntries() {
      await batchDeleteByCollection("payroll_entries");
    },

    async clearEmployees() {
      await batchDeleteByCollection("employees");
    },

    async clearDesignationPresets() {
      await batchDeleteByCollection("designation_presets");
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
