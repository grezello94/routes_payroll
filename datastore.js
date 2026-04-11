const fs = require("fs");
const path = require("path");
const admin = require("firebase-admin");
const PAYROLL_REPORT_ROW_ID = "__PAYROLL_REPORT__";

function resolveFirebaseCredentialFile(baseDir) {
  const explicitPath = String(process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "").trim();
  const candidatePaths = [
    explicitPath,
    path.join(baseDir, "firebase-service-account.json"),
    path.join(baseDir, "routespayroll-firebase-adminsdk-fbsvc-6a7da5deea.json"),
  ].filter(Boolean);

  for (const candidate of candidatePaths) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return "";
}

function createStore({ baseDir }) {
  const requestedProvider = String(process.env.DB_PROVIDER || "").toLowerCase();
  const hasSupabase = Boolean(process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY);
  const provider = requestedProvider || (hasSupabase ? "supabase" : "firebase");
  if (provider === "supabase") {
    return createSupabaseStore();
  }
  if (provider === "firebase") {
    return createFirebaseStore(baseDir);
  }
  return createSqliteStore(baseDir);
}

function createSqliteStore(baseDir) {
  const sqlite3 = require("sqlite3").verbose();
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
          email_verified INTEGER NOT NULL DEFAULT 1,
          email_verification_hash TEXT NOT NULL DEFAULT '',
          email_verification_expires_at TEXT NOT NULL DEFAULT '',
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
        CREATE TABLE IF NOT EXISTS payroll_reports (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          company_id INTEGER NOT NULL DEFAULT 1,
          month TEXT NOT NULL,
          checked_at TEXT NOT NULL DEFAULT '',
          generated_at TEXT NOT NULL DEFAULT '',
          employee_count INTEGER NOT NULL DEFAULT 0,
          snapshot_json TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
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
        await run("ALTER TABLE users ADD COLUMN email_verified INTEGER NOT NULL DEFAULT 1");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN email_verification_hash TEXT NOT NULL DEFAULT ''");
      } catch (error) {
        if (!String(error?.message || "").toLowerCase().includes("duplicate column name")) {
          throw error;
        }
      }

      try {
        await run("ALTER TABLE users ADD COLUMN email_verification_expires_at TEXT NOT NULL DEFAULT ''");
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
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_payroll_reports_company_month_unique ON payroll_reports(company_id, month)"
      );
      await run(
        "CREATE INDEX IF NOT EXISTS idx_payroll_reports_company_month ON payroll_reports(company_id, month)"
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
        `SELECT id, username, email, email_verified, email_verification_hash, email_verification_expires_at,
                password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE username_lc = ? OR email_lc = ?
         LIMIT 1`,
        [normalized, normalized]
      );
    },

    async findUserByEmail(email) {
      const normalized = String(email || "").trim().toLowerCase();
      return get(
        `SELECT id, username, email, email_verified, email_verification_hash, email_verification_expires_at,
                password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE email_lc = ?
         LIMIT 1`,
        [normalized]
      );
    },

    async getUserById(userId) {
      return get(
        `SELECT id, username, email, email_verified, email_verification_hash, email_verification_expires_at,
                password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE id = ?
         LIMIT 1`,
        [userId]
      );
    },

    async createUser(username, email, passwordHash, options = {}) {
      const emailVerified = options.emailVerified === true ? 1 : 0;
      const result = await run(
        `INSERT INTO users
          (username, username_lc, email, email_lc, email_verified, password_hash)
         VALUES (?, ?, ?, ?, ?, ?)`,
        [username, String(username).toLowerCase(), email, String(email).toLowerCase(), emailVerified, passwordHash]
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

    async setUserEmailVerificationById(userId, verificationHash, expiresAtIso) {
      await run(
        "UPDATE users SET email_verified = 0, email_verification_hash = ?, email_verification_expires_at = ? WHERE id = ?",
        [verificationHash, expiresAtIso, userId]
      );
    },

    async markUserEmailVerifiedById(userId) {
      await run(
        `UPDATE users
         SET email_verified = 1, email_verification_hash = '', email_verification_expires_at = ''
         WHERE id = ?`,
        [userId]
      );
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
        `SELECT id, username, email, email_verified, email_verification_hash, email_verification_expires_at,
                password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE reset_code_hash = ?
         LIMIT 1`,
        [String(resetCodeHash || "")]
      );
    },

    async findUserByEmailVerificationHash(verificationHash) {
      return get(
        `SELECT id, username, email, email_verified, email_verification_hash, email_verification_expires_at,
                password_hash, reset_code_hash, reset_code_expires_at
         FROM users
         WHERE email_verification_hash = ?
         LIMIT 1`,
        [String(verificationHash || "")]
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
         WHERE employee_id <> ?
         ORDER BY company_id, month, position_index, id`
      , [PAYROLL_REPORT_ROW_ID]);
    },

    async listPayrollByMonthCompany(month, companyId) {
      return all(
        `SELECT id, company_id, month, employee_id, employee_name, designation,
                present_salary, increment, old_advance_taken, extra_advance_added,
                deduction_entered, days_absent, comment, position_index
         FROM payroll_entries
         WHERE month = ? AND company_id = ? AND employee_id <> ?
         ORDER BY position_index, id`,
        [month, companyId, PAYROLL_REPORT_ROW_ID]
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

    async listPayrollReportsByCompany(companyId) {
      const rows = await all(
        `SELECT id, company_id, month, checked_at, generated_at, employee_count, snapshot_json
         FROM payroll_reports
         WHERE company_id = ?
         ORDER BY month DESC, id DESC`,
        [companyId]
      );
      if (rows.length) {
        return rows.map((row) => ({
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.checked_at || "",
          generated_at: row.generated_at || "",
          employee_count: Number(row.employee_count || 0),
          snapshot_json: row.snapshot_json || "",
          legacy: false,
        }));
      }
      return all(
        `SELECT id, company_id, month, employee_name, designation, present_salary, comment
         FROM payroll_entries
         WHERE company_id = ? AND employee_id = ?
         ORDER BY month DESC, id DESC`,
        [companyId, PAYROLL_REPORT_ROW_ID]
      ).then((legacyRows) => legacyRows.map((row) => ({
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      })));
    },

    async getPayrollReportByIdCompany(id, companyId) {
      const row = await get(
        `SELECT id, company_id, month, checked_at, generated_at, employee_count, snapshot_json
         FROM payroll_reports
         WHERE id = ? AND company_id = ?
         LIMIT 1`,
        [id, companyId]
      );
      if (row) {
        return {
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.checked_at || "",
          generated_at: row.generated_at || "",
          employee_count: Number(row.employee_count || 0),
          snapshot_json: row.snapshot_json || "",
          legacy: false,
        };
      }
      return get(
        `SELECT id, company_id, month, employee_name, designation, present_salary, comment
         FROM payroll_entries
         WHERE id = ? AND company_id = ? AND employee_id = ?
         LIMIT 1`,
        [id, companyId, PAYROLL_REPORT_ROW_ID]
      ).then((row) => (!row ? null : {
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      }));
    },

    async getPayrollReportByMonthCompany(month, companyId) {
      const row = await get(
        `SELECT id, company_id, month, checked_at, generated_at, employee_count, snapshot_json
         FROM payroll_reports
         WHERE month = ? AND company_id = ?
         LIMIT 1`,
        [month, companyId]
      );
      if (row) {
        return {
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.checked_at || "",
          generated_at: row.generated_at || "",
          employee_count: Number(row.employee_count || 0),
          snapshot_json: row.snapshot_json || "",
          legacy: false,
        };
      }
      return get(
        `SELECT id, company_id, month, employee_name, designation, present_salary, comment
         FROM payroll_entries
         WHERE month = ? AND company_id = ? AND employee_id = ?
         LIMIT 1`,
        [month, companyId, PAYROLL_REPORT_ROW_ID]
      ).then((row) => (!row ? null : {
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      }));
    },

    async markPayrollChecked(month, companyId, checkedAt) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      if (existing && !existing.legacy) {
        await run(
          `UPDATE payroll_reports
           SET checked_at = ?, updated_at = CURRENT_TIMESTAMP
           WHERE id = ?`,
          [checkedAt, existing.id]
        );
        return { id: existing.id };
      }

      const result = await run(
        `INSERT INTO payroll_reports (
          company_id, month, checked_at, generated_at, employee_count, snapshot_json, updated_at
         ) VALUES (?, ?, ?, '', 0, '', CURRENT_TIMESTAMP)`,
        [companyId, month, checkedAt]
      );
      return { id: result.lastID };
    },

    async savePayrollReportSnapshot(month, companyId, payload) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      const snapshotJson = JSON.stringify(payload?.snapshot || {});
      if (existing && !existing.legacy) {
        await run(
          `UPDATE payroll_reports
           SET checked_at = ?, generated_at = ?, employee_count = ?, snapshot_json = ?, updated_at = CURRENT_TIMESTAMP
           WHERE id = ?`,
          [
            payload?.checkedAt || existing.checked_at || "",
            payload?.generatedAt || "",
            Number(payload?.employeeCount || 0),
            snapshotJson,
            existing.id,
          ]
        );
        return { id: existing.id };
      }

      const result = await run(
        `INSERT INTO payroll_reports (
          company_id, month, checked_at, generated_at, employee_count, snapshot_json, updated_at
         ) VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
        [
          companyId,
          month,
          payload?.checkedAt || "",
          payload?.generatedAt || "",
          Number(payload?.employeeCount || 0),
          snapshotJson,
        ]
      );
      return { id: result.lastID };
    },
  };
}

function createFirebaseStore(baseDir) {
  const serviceAccountJson = process.env.FIREBASE_SERVICE_ACCOUNT_JSON || "";
  const serviceAccountPath = resolveFirebaseCredentialFile(baseDir);
  const firebaseProjectId = process.env.FIREBASE_PROJECT_ID || "";

  if (!admin.apps.length) {
    if (serviceAccountJson) {
      const credentials = JSON.parse(serviceAccountJson);
      admin.initializeApp({
        credential: admin.credential.cert(credentials),
        projectId: firebaseProjectId || credentials.project_id,
      });
    } else if (serviceAccountPath) {
      const credentials = JSON.parse(fs.readFileSync(serviceAccountPath, "utf8"));
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
          payroll_reports: admin.firestore.FieldValue.increment(0),
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

    async getUserById(userId) {
      const snapshot = await col("users").doc(String(userId)).get();
      return snapshot.exists ? (snapshot.data() || null) : null;
    },

    async createUser(username, email, passwordHash, options = {}) {
      const usernameLc = String(username || "").toLowerCase();
      const emailLc = String(email || "").toLowerCase();
      const emailVerified = options.emailVerified === true;
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
        email_verified: emailVerified,
        email_verification_hash: "",
        email_verification_expires_at: "",
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

    async setUserEmailVerificationById(userId, verificationHash, expiresAtIso) {
      await col("users").doc(String(userId)).set(
        {
          email_verified: false,
          email_verification_hash: verificationHash,
          email_verification_expires_at: expiresAtIso,
        },
        { merge: true }
      );
    },

    async markUserEmailVerifiedById(userId) {
      await col("users").doc(String(userId)).set(
        {
          email_verified: true,
          email_verification_hash: "",
          email_verification_expires_at: "",
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

    async findUserByEmailVerificationHash(verificationHash) {
      const user = await findByField("users", "email_verification_hash", String(verificationHash || ""));
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
        .filter((row) => String(row.employee_id || "") !== PAYROLL_REPORT_ROW_ID)
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
        .filter((row) => row.month === month && Number(row.company_id) === Number(companyId) && String(row.employee_id || "") !== PAYROLL_REPORT_ROW_ID)
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

    async listPayrollReportsByCompany(companyId) {
      const snapshot = await col("payroll_reports").get();
      const reports = snapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .filter((row) => Number(row.company_id) === Number(companyId))
        .map((row) => ({
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.checked_at || "",
          generated_at: row.generated_at || "",
          employee_count: Number(row.employee_count || 0),
          snapshot_json: row.snapshot_json || "",
          created_at: row.created_at || "",
          legacy: false,
        }))
        .sort((a, b) => String(b.month || "").localeCompare(String(a.month || "")) || (Number(b.id || 0) - Number(a.id || 0)));
      if (reports.length) return reports;

      const legacySnapshot = await col("payroll_entries").get();
      return legacySnapshot.docs
        .map((docSnap) => docSnap.data() || {})
        .filter((row) => Number(row.company_id) === Number(companyId) && String(row.employee_id || "") === PAYROLL_REPORT_ROW_ID)
        .map((row) => ({
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.employee_name || "",
          generated_at: row.designation || "",
          employee_count: Number(row.present_salary || 0),
          snapshot_json: row.comment || "",
          created_at: row.created_at || "",
          legacy: true,
        }))
        .sort((a, b) => String(b.month || "").localeCompare(String(a.month || "")) || (Number(b.id || 0) - Number(a.id || 0)));
    },

    async getPayrollReportByIdCompany(id, companyId) {
      const snap = await col("payroll_reports").doc(String(id)).get();
      if (snap.exists) {
        const row = snap.data() || {};
        if (Number(row.company_id) !== Number(companyId)) return null;
        return {
          id: row.id,
          company_id: row.company_id,
          month: row.month,
          checked_at: row.checked_at || "",
          generated_at: row.generated_at || "",
          employee_count: Number(row.employee_count || 0),
          snapshot_json: row.snapshot_json || "",
          created_at: row.created_at || "",
          legacy: false,
        };
      }
      const legacySnap = await col("payroll_entries").doc(String(id)).get();
      if (!legacySnap.exists) return null;
      const row = legacySnap.data() || {};
      if (Number(row.company_id) !== Number(companyId) || String(row.employee_id || "") !== PAYROLL_REPORT_ROW_ID) return null;
      return {
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        created_at: row.created_at || "",
        legacy: true,
      };
    },

    async getPayrollReportByMonthCompany(month, companyId) {
      const rows = await this.listPayrollReportsByCompany(companyId);
      return rows.find((row) => String(row.month || "") === String(month)) || null;
    },

    async markPayrollChecked(month, companyId, checkedAt) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      if (existing?.id && !existing.legacy) {
        await col("payroll_reports").doc(String(existing.id)).set({
          checked_at: checkedAt,
          updated_at: nowIso(),
        }, { merge: true });
        return { id: existing.id };
      }

      const id = await getCounter("payroll_reports");
      await col("payroll_reports").doc(String(id)).set({
        id,
        company_id: Number(companyId),
        month,
        checked_at: checkedAt,
        generated_at: "",
        employee_count: 0,
        snapshot_json: "",
        created_at: nowIso(),
        updated_at: nowIso(),
      });
      return { id };
    },

    async savePayrollReportSnapshot(month, companyId, payload) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      const reportId = existing?.id && !existing.legacy ? existing.id : await getCounter("payroll_reports");
      await col("payroll_reports").doc(String(reportId)).set({
        id: Number(reportId),
        company_id: Number(companyId),
        month,
        checked_at: payload?.checkedAt || existing?.checked_at || "",
        generated_at: payload?.generatedAt || "",
        employee_count: Number(payload?.employeeCount || 0),
        snapshot_json: JSON.stringify(payload?.snapshot || {}),
        created_at: existing && !existing.legacy ? existing.created_at : nowIso(),
        updated_at: nowIso(),
      }, { merge: true });
      return { id: reportId };
    },
  };
}

function createSupabaseStore() {
  const baseUrl = String(process.env.SUPABASE_URL || "").replace(/\/+$/, "");
  const serviceRoleKey = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
  const requestTimeoutMs = Math.max(1000, Number(process.env.SUPABASE_REQUEST_TIMEOUT_MS || 8000));
  const nowIso = () => new Date().toISOString();
  const supabaseCapabilities = {
    requiredTablesChecked: false,
    payrollReportsTable: null,
  };

  if (!baseUrl || !serviceRoleKey) {
    throw new Error("Supabase is not configured. Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY.");
  }

  async function fetchWithTimeout(url, options = {}) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), requestTimeoutMs);

    try {
      return await fetch(url, { ...options, signal: controller.signal });
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(`Supabase request timed out after ${requestTimeoutMs}ms.`);
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  async function supabaseRest(table, {
    method = "GET",
    query,
    body,
    headers = {},
    allowEmpty = false,
  } = {}) {
    const url = new URL(`/rest/v1/${table}`, `${baseUrl}/`);
    if (query) {
      Object.entries(query).forEach(([key, value]) => {
        if (value !== undefined && value !== null) {
          url.searchParams.set(key, value);
        }
      });
    }

    const response = await fetchWithTimeout(url, {
      method,
      headers: {
        apikey: serviceRoleKey,
        Authorization: `Bearer ${serviceRoleKey}`,
        "Content-Type": "application/json",
        ...headers,
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });

    if (!response.ok) {
      let payload = null;
      try {
        payload = await response.json();
      } catch {
        payload = null;
      }
      const error = new Error(payload?.message || payload?.error_description || payload?.hint || `Supabase request failed for ${table}.`);
      error.code = payload?.code || response.status;
      error.details = payload;
      throw error;
    }

    if (response.status === 204) return allowEmpty ? [] : null;
    const text = await response.text();
    if (!text) return allowEmpty ? [] : null;
    return JSON.parse(text);
  }

  async function maybeSingle(table, query) {
    const rows = await supabaseRest(table, { query: { ...query, limit: "1" }, allowEmpty: true });
    return Array.isArray(rows) && rows.length > 0 ? rows[0] : null;
  }

  async function nextNumericId(table) {
    const rows = await supabaseRest(table, {
      query: { select: "id", order: "id.desc", limit: "1" },
      allowEmpty: true,
    });
    const current = Array.isArray(rows) && rows.length > 0 ? Number(rows[0].id || 0) : 0;
    return current + 1;
  }

  async function insertSupabaseRowWithRetry(table, buildBody) {
    let attemptedIds = [];
    for (let attempt = 0; attempt < 3; attempt += 1) {
      const generatedId = attempt === 0
        ? await nextNumericId(table)
        : Math.max(Date.now(), ...attemptedIds) + attempt;
      attemptedIds.push(generatedId);
      try {
        await supabaseRest(table, {
          method: "POST",
          body: buildBody(generatedId),
        });
        return generatedId;
      } catch (error) {
        if (String(error?.code || "") !== "23505") throw error;
      }
    }
    throw new Error(`Failed to allocate a unique id for Supabase table "${table}".`);
  }

  function mapConflict(error, message) {
    if (String(error?.code || "") === "23505") {
      const conflict = new Error(message);
      conflict.code = "unique";
      throw conflict;
    }
    throw error;
  }

  function isMissingSupabaseTable(error, tableName) {
    const code = String(error?.code || "");
    const details = String(error?.details?.message || error?.details?.details || error?.message || "").toLowerCase();
    const table = String(tableName || "").toLowerCase();
    return code === "42P01"
      || code === "PGRST205"
      || details.includes(`relation "${table}" does not exist`)
      || details.includes(`relation 'public.${table}' does not exist`)
      || details.includes(`could not find the table '${table}'`)
      || details.includes(`could not find the table 'public.${table}'`)
      || details.includes(`could not find the relation '${table}'`);
  }

  async function probeTableExists(table) {
    try {
      await supabaseRest(table, {
        query: { select: "id", limit: "1" },
        allowEmpty: true,
      });
      return true;
    } catch (error) {
      if (isMissingSupabaseTable(error, table)) return false;
      throw error;
    }
  }

  async function ensureSupabaseSchemaReady() {
    if (supabaseCapabilities.requiredTablesChecked) return;

    const requiredTables = ["users", "companies", "designation_presets", "employees", "payroll_entries"];
    for (const table of requiredTables) {
      // eslint-disable-next-line no-await-in-loop
      const exists = await probeTableExists(table);
      if (!exists) {
        throw new Error(`Supabase table "${table}" is missing. Run supabase-setup.sql in the Supabase SQL editor.`);
      }
    }

    supabaseCapabilities.payrollReportsTable = await probeTableExists("payroll_reports");
    supabaseCapabilities.requiredTablesChecked = true;
  }

  async function canUsePayrollReportsTable() {
    await ensureSupabaseSchemaReady();
    return supabaseCapabilities.payrollReportsTable === true;
  }

  return {
    provider: "supabase",

    async init() {
      await ensureSupabaseSchemaReady();
      await this.ensureDefaultCompany();
      await this.ensureDefaultDesignationPresets(1);
    },

    async countUsers() {
      const rows = await supabaseRest("users", {
        query: { select: "id", limit: "1000" },
        allowEmpty: true,
      });
      return Array.isArray(rows) ? rows.length : 0;
    },

    async findUserByIdentifier(identifier) {
      const normalized = String(identifier || "").trim().toLowerCase();
      if (!normalized) return null;
      const byUsername = await maybeSingle("users", {
        select: "*",
        username_lc: `eq.${normalized}`,
      });
      if (byUsername) return byUsername;
      return maybeSingle("users", {
        select: "*",
        email_lc: `eq.${normalized}`,
      });
    },

    async findUserByEmail(email) {
      return maybeSingle("users", {
        select: "*",
        email_lc: `eq.${String(email || "").trim().toLowerCase()}`,
      });
    },

    async getUserById(userId) {
      return maybeSingle("users", {
        select: "*",
        id: `eq.${String(userId || "").trim()}`,
      });
    },

    async createUser(username, email, passwordHash, options = {}) {
      try {
        const id = String(options.id || "").trim();
        if (!id) {
          throw new Error("Supabase user id is required.");
        }
        const rows = await supabaseRest("users", {
          method: "POST",
          headers: { Prefer: "return=representation" },
          body: [{
            id,
            username,
            username_lc: String(username || "").toLowerCase(),
            email,
            email_lc: String(email || "").toLowerCase(),
            email_verified: options.emailVerified === true,
            email_verification_hash: "",
            email_verification_expires_at: "",
            password_hash: passwordHash || "",
            reset_code_hash: "",
            reset_code_expires_at: "",
            created_at: nowIso(),
          }],
        });
        return { id: rows?.[0]?.id || id };
      } catch (error) {
        mapConflict(error, "Username or email already exists.");
      }
    },

    async updateUserPasswordByEmail(email, passwordHash) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { email_lc: `eq.${String(email || "").trim().toLowerCase()}` },
        body: { password_hash: passwordHash || "" },
      });
    },

    async updateUserPasswordById(userId, passwordHash) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { id: `eq.${String(userId || "").trim()}` },
        body: { password_hash: passwordHash || "" },
      });
    },

    async setUserEmailVerificationById(userId, verificationHash, expiresAtIso) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { id: `eq.${String(userId || "").trim()}` },
        body: {
          email_verified: false,
          email_verification_hash: verificationHash,
          email_verification_expires_at: expiresAtIso,
        },
      });
    },

    async markUserEmailVerifiedById(userId) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { id: `eq.${String(userId || "").trim()}` },
        body: {
          email_verified: true,
          email_verification_hash: "",
          email_verification_expires_at: "",
        },
      });
    },

    async setUserResetCodeById(userId, resetCodeHash, expiresAtIso) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { id: `eq.${String(userId || "").trim()}` },
        body: {
          reset_code_hash: resetCodeHash,
          reset_code_expires_at: expiresAtIso,
        },
      });
    },

    async clearUserResetCodeById(userId) {
      await supabaseRest("users", {
        method: "PATCH",
        query: { id: `eq.${String(userId || "").trim()}` },
        body: {
          reset_code_hash: "",
          reset_code_expires_at: "",
        },
      });
    },

    async findUserByResetCodeHash(resetCodeHash) {
      return maybeSingle("users", {
        select: "*",
        reset_code_hash: `eq.${String(resetCodeHash || "")}`,
      });
    },

    async findUserByEmailVerificationHash(verificationHash) {
      return maybeSingle("users", {
        select: "*",
        email_verification_hash: `eq.${String(verificationHash || "")}`,
      });
    },

    async updateCompanyName(id, name, logoDataUrl = "") {
      await supabaseRest("companies", {
        method: "PATCH",
        query: { id: `eq.${Number(id)}` },
        body: {
          name,
          name_lc: String(name || "").toLowerCase(),
          logo_data_url: logoDataUrl,
        },
      });
    },

    async updateCompanyLogo(id, logoDataUrl = "") {
      await supabaseRest("companies", {
        method: "PATCH",
        query: { id: `eq.${Number(id)}` },
        body: { logo_data_url: logoDataUrl },
      });
    },

    async companyExists(id) {
      const row = await maybeSingle("companies", {
        select: "id",
        id: `eq.${Number(id)}`,
      });
      return Boolean(row);
    },

    async listCompanies() {
      return supabaseRest("companies", {
        query: { select: "*", order: "name.asc,id.asc" },
        allowEmpty: true,
      });
    },

    async listCompaniesById() {
      return supabaseRest("companies", {
        query: { select: "*", order: "id.asc" },
        allowEmpty: true,
      });
    },

    async getCompanyById(id) {
      return maybeSingle("companies", {
        select: "*",
        id: `eq.${Number(id)}`,
      });
    },

    async createCompany(name, logoDataUrl) {
      try {
        const id = await nextNumericId("companies");
        const rows = await supabaseRest("companies", {
          method: "POST",
          headers: { Prefer: "return=representation" },
          body: [{
            id,
            name,
            name_lc: String(name || "").toLowerCase(),
            logo_data_url: logoDataUrl,
            created_at: nowIso(),
          }],
        });
        await this.ensureDefaultDesignationPresets(id);
        return { id: rows?.[0]?.id || id };
      } catch (error) {
        mapConflict(error, "Company name already exists.");
      }
    },

    async listDesignationPresets(companyId) {
      return supabaseRest("designation_presets", {
        query: {
          select: "*",
          company_id: `eq.${Number(companyId)}`,
          order: "position_index.asc,id.asc",
        },
        allowEmpty: true,
      });
    },

    async addDesignationPreset(companyId, name, positionIndex) {
      try {
        const id = await nextNumericId("designation_presets");
        const rows = await supabaseRest("designation_presets", {
          method: "POST",
          headers: { Prefer: "return=representation" },
          body: [{
            id,
            company_id: Number(companyId),
            name,
            name_lc: String(name || "").trim().toLowerCase(),
            position_index: Number(positionIndex || 0),
            created_at: nowIso(),
          }],
        });
        return { id: rows?.[0]?.id || id };
      } catch (error) {
        mapConflict(error, "Designation already exists.");
      }
    },

    async deleteDesignationPreset(companyId, id) {
      await supabaseRest("designation_presets", {
        method: "DELETE",
        query: {
          id: `eq.${Number(id)}`,
          company_id: `eq.${Number(companyId)}`,
        },
      });
    },

    async ensureDefaultDesignationPresets(companyId) {
      const defaults = ["Manager", "Chef", "Accountant", "Supervisor", "Staff"];
      let idx = 0;
      for (const name of defaults) {
        try {
          await this.addDesignationPreset(companyId, name, idx);
        } catch (error) {
          if (error?.code !== "unique") throw error;
        }
        idx += 1;
      }
    },

    async listEmployeesByCompany(companyId) {
      return supabaseRest("employees", {
        query: {
          select: "*",
          company_id: `eq.${Number(companyId)}`,
          order: "position_index.asc,id.asc",
        },
        allowEmpty: true,
      });
    },

    async getEmployeeByIdCompany(id, companyId) {
      return maybeSingle("employees", {
        select: "*",
        id: `eq.${Number(id)}`,
        company_id: `eq.${Number(companyId)}`,
      });
    },

    async createEmployee(companyId, employee) {
      try {
        const id = await nextNumericId("employees");
        const rows = await supabaseRest("employees", {
          method: "POST",
          headers: { Prefer: "return=representation" },
          body: [{
            id,
            company_id: Number(companyId),
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
          }],
        });
        return { id: rows?.[0]?.id || id };
      } catch (error) {
        mapConflict(error, "Employee ID already exists for this company.");
      }
    },

    async updateEmployee(companyId, id, employee) {
      try {
        await supabaseRest("employees", {
          method: "PATCH",
          query: {
            id: `eq.${Number(id)}`,
            company_id: `eq.${Number(companyId)}`,
          },
          body: {
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
        });
      } catch (error) {
        mapConflict(error, "Employee ID already exists for this company.");
      }
    },

    async deleteEmployeeByIdCompany(id, companyId) {
      await supabaseRest("employees", {
        method: "DELETE",
        query: {
          id: `eq.${Number(id)}`,
          company_id: `eq.${Number(companyId)}`,
        },
      });
    },

    async clearPayrollEntries() {
      await supabaseRest("payroll_entries", {
        method: "DELETE",
        query: { id: "gt.0" },
      });
    },

    async clearEmployees() {
      await supabaseRest("employees", {
        method: "DELETE",
        query: { id: "gt.0" },
      });
    },

    async clearDesignationPresets() {
      await supabaseRest("designation_presets", {
        method: "DELETE",
        query: { id: "gt.0" },
      });
    },

    async clearCompanies() {
      await supabaseRest("companies", {
        method: "DELETE",
        query: { id: "gt.0" },
      });
    },

    async insertCompanyWithId(id, name, logoDataUrl) {
      await supabaseRest("companies", {
        method: "POST",
        body: [{
          id: Number(id),
          name,
          name_lc: String(name || "").toLowerCase(),
          logo_data_url: logoDataUrl,
          created_at: nowIso(),
        }],
      });
    },

    async ensureDefaultCompany() {
      const exists = await this.companyExists(1);
      if (!exists) {
        await this.insertCompanyWithId(1, "Routes Payroll", "");
      }
    },

    async listPayrollEntriesAll() {
      return supabaseRest("payroll_entries", {
        query: {
          select: "*",
          employee_id: `neq.${PAYROLL_REPORT_ROW_ID}`,
          order: "company_id.asc,month.asc,position_index.asc,id.asc",
        },
        allowEmpty: true,
      });
    },

    async listPayrollByMonthCompany(month, companyId) {
      return supabaseRest("payroll_entries", {
        query: {
          select: "*",
          month: `eq.${month}`,
          company_id: `eq.${Number(companyId)}`,
          employee_id: `neq.${PAYROLL_REPORT_ROW_ID}`,
          order: "position_index.asc,id.asc",
        },
        allowEmpty: true,
      });
    },

    async deletePayrollByMonthCompany(month, companyId) {
      await supabaseRest("payroll_entries", {
        method: "DELETE",
        query: {
          month: `eq.${month}`,
          company_id: `eq.${Number(companyId)}`,
        },
      });
    },

    async insertPayrollRecord(month, companyId, record) {
      const id = await nextNumericId("payroll_entries");
      await supabaseRest("payroll_entries", {
        method: "POST",
        body: [{
          id,
          company_id: Number(companyId),
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
        }],
      });
    },

    async listPayrollReportsByCompany(companyId) {
      if (await canUsePayrollReportsTable()) {
        const rows = await supabaseRest("payroll_reports", {
          query: {
            select: "*",
            company_id: `eq.${Number(companyId)}`,
            order: "month.desc,id.desc",
          },
          allowEmpty: true,
        });
        if (rows.length) {
          return rows.map((row) => ({
            id: row.id,
            company_id: row.company_id,
            month: row.month,
            checked_at: row.checked_at || "",
            generated_at: row.generated_at || "",
            employee_count: Number(row.employee_count || 0),
            snapshot_json: row.snapshot_json || "",
            legacy: false,
          }));
        }
      }
      return supabaseRest("payroll_entries", {
        query: {
          select: "*",
          company_id: `eq.${Number(companyId)}`,
          employee_id: `eq.${PAYROLL_REPORT_ROW_ID}`,
          order: "month.desc,id.desc",
        },
        allowEmpty: true,
      }).then((legacyRows) => legacyRows.map((row) => ({
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      })));
    },

    async getPayrollReportByIdCompany(id, companyId) {
      if (await canUsePayrollReportsTable()) {
        const row = await maybeSingle("payroll_reports", {
          select: "*",
          id: `eq.${Number(id)}`,
          company_id: `eq.${Number(companyId)}`,
        });
        if (row) {
          return {
            id: row.id,
            company_id: row.company_id,
            month: row.month,
            checked_at: row.checked_at || "",
            generated_at: row.generated_at || "",
            employee_count: Number(row.employee_count || 0),
            snapshot_json: row.snapshot_json || "",
            legacy: false,
          };
        }
      }
      return maybeSingle("payroll_entries", {
        select: "*",
        id: `eq.${Number(id)}`,
        company_id: `eq.${Number(companyId)}`,
        employee_id: `eq.${PAYROLL_REPORT_ROW_ID}`,
      }).then((row) => (!row ? null : {
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      }));
    },

    async getPayrollReportByMonthCompany(month, companyId) {
      if (await canUsePayrollReportsTable()) {
        const row = await maybeSingle("payroll_reports", {
          select: "*",
          month: `eq.${month}`,
          company_id: `eq.${Number(companyId)}`,
        });
        if (row) {
          return {
            id: row.id,
            company_id: row.company_id,
            month: row.month,
            checked_at: row.checked_at || "",
            generated_at: row.generated_at || "",
            employee_count: Number(row.employee_count || 0),
            snapshot_json: row.snapshot_json || "",
            legacy: false,
          };
        }
      }
      return maybeSingle("payroll_entries", {
        select: "*",
        month: `eq.${month}`,
        company_id: `eq.${Number(companyId)}`,
        employee_id: `eq.${PAYROLL_REPORT_ROW_ID}`,
      }).then((row) => (!row ? null : {
        id: row.id,
        company_id: row.company_id,
        month: row.month,
        checked_at: row.employee_name || "",
        generated_at: row.designation || "",
        employee_count: Number(row.present_salary || 0),
        snapshot_json: row.comment || "",
        legacy: true,
      }));
    },

    async markPayrollChecked(month, companyId, checkedAt) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      if ((await canUsePayrollReportsTable()) && existing?.id && !existing.legacy) {
        await supabaseRest("payroll_reports", {
          method: "PATCH",
          query: { id: `eq.${Number(existing.id)}` },
          body: {
            checked_at: checkedAt,
            updated_at: nowIso(),
          },
        });
        return { id: existing.id };
      }

      if (await canUsePayrollReportsTable()) {
        const id = await nextNumericId("payroll_reports");
        await supabaseRest("payroll_reports", {
          method: "POST",
          body: [{
            id,
            company_id: Number(companyId),
            month,
            checked_at: checkedAt,
            generated_at: "",
            employee_count: 0,
            snapshot_json: "",
            created_at: nowIso(),
            updated_at: nowIso(),
          }],
        });
        return { id };
      }

      if (existing?.id && existing.legacy) {
        await supabaseRest("payroll_entries", {
          method: "PATCH",
          query: { id: `eq.${Number(existing.id)}` },
          body: {
            employee_name: checkedAt,
            updated_at: nowIso(),
          },
        });
        return { id: existing.id };
      }

      const legacyId = await insertSupabaseRowWithRetry("payroll_entries", (generatedId) => ([{
        id: generatedId,
        company_id: Number(companyId),
        month,
        employee_id: PAYROLL_REPORT_ROW_ID,
        employee_name: checkedAt,
        designation: "",
        present_salary: 0,
        increment: 0,
        old_advance_taken: 0,
        extra_advance_added: 0,
        deduction_entered: 0,
        days_absent: 0,
        comment: "",
        position_index: 999999,
        updated_at: nowIso(),
      }]));
      return { id: legacyId };
    },

    async savePayrollReportSnapshot(month, companyId, payload) {
      const existing = await this.getPayrollReportByMonthCompany(month, companyId);
      const body = {
        checked_at: payload?.checkedAt || existing?.checked_at || "",
        generated_at: payload?.generatedAt || "",
        employee_count: Number(payload?.employeeCount || 0),
        snapshot_json: JSON.stringify(payload?.snapshot || {}),
        updated_at: nowIso(),
      };

      if ((await canUsePayrollReportsTable()) && existing?.id && !existing.legacy) {
        await supabaseRest("payroll_reports", {
          method: "PATCH",
          query: { id: `eq.${Number(existing.id)}` },
          body,
        });
        return { id: existing.id };
      }

      if (await canUsePayrollReportsTable()) {
        const id = await nextNumericId("payroll_reports");
        await supabaseRest("payroll_reports", {
          method: "POST",
          body: [{
            id,
            company_id: Number(companyId),
            month,
            checked_at: payload?.checkedAt || "",
            generated_at: payload?.generatedAt || "",
            employee_count: Number(payload?.employeeCount || 0),
            snapshot_json: JSON.stringify(payload?.snapshot || {}),
            created_at: nowIso(),
            updated_at: nowIso(),
          }],
        });
        return { id };
      }

      const legacyBody = {
        employee_name: payload?.checkedAt || existing?.checked_at || "",
        designation: payload?.generatedAt || "",
        present_salary: Number(payload?.employeeCount || 0),
        comment: JSON.stringify(payload?.snapshot || {}),
        position_index: 999999,
        updated_at: nowIso(),
      };
      if (existing?.id && existing.legacy) {
        await supabaseRest("payroll_entries", {
          method: "PATCH",
          query: { id: `eq.${Number(existing.id)}` },
          body: legacyBody,
        });
        return { id: existing.id };
      }

      const legacyId = await insertSupabaseRowWithRetry("payroll_entries", (generatedId) => ([{
        id: generatedId,
        company_id: Number(companyId),
        month,
        employee_id: PAYROLL_REPORT_ROW_ID,
        employee_name: legacyBody.employee_name,
        designation: legacyBody.designation,
        present_salary: legacyBody.present_salary,
        increment: 0,
        old_advance_taken: 0,
        extra_advance_added: 0,
        deduction_entered: 0,
        days_absent: 0,
        comment: legacyBody.comment,
        position_index: 999999,
        updated_at: nowIso(),
      }]));
      return { id: legacyId };
    },
  };
}

module.exports = {
  createStore,
};
