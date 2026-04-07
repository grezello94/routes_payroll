const admin = require("firebase-admin");
const fs = require("fs");
const path = require("path");
const sqlite3 = require("sqlite3").verbose();

const baseDir = __dirname;
const sqlitePath = path.join(baseDir, "data", "payroll.db");
const serviceAccountPath = [
  process.env.FIREBASE_SERVICE_ACCOUNT_PATH || "",
  path.join(baseDir, "firebase-service-account.json"),
  path.join(baseDir, "routespayroll-firebase-adminsdk-fbsvc-6a7da5deea.json"),
].find((candidate) => candidate && fs.existsSync(candidate));

if (!serviceAccountPath) {
  console.error("No Firebase service account file found.");
  process.exit(1);
}

const serviceAccount = JSON.parse(fs.readFileSync(serviceAccountPath, "utf8"));

if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
    projectId: process.env.FIREBASE_PROJECT_ID || serviceAccount.project_id,
  });
}

const firestore = admin.firestore();

function getSqliteRow(sql, params = []) {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(sqlitePath);
    db.get(sql, params, (error, row) => {
      db.close();
      if (error) {
        reject(error);
        return;
      }
      resolve(row || null);
    });
  });
}

async function main() {
  const user = await getSqliteRow(
    `SELECT id, username, email, password_hash, reset_code_hash, reset_code_expires_at, created_at
     FROM users
     ORDER BY id
     LIMIT 1`
  );

  if (!user) {
    console.error("No local SQLite admin user found to sync.");
    process.exit(1);
  }

  const company = await getSqliteRow(
    `SELECT id, name, logo_data_url, created_at
     FROM companies
     WHERE id = 1
     LIMIT 1`
  );

  await firestore.collection("users").doc(String(user.id)).set({
    id: Number(user.id),
    username: String(user.username || ""),
    username_lc: String(user.username || "").toLowerCase(),
    email: String(user.email || "").toLowerCase(),
    email_lc: String(user.email || "").toLowerCase(),
    email_verified: true,
    email_verification_hash: "",
    email_verification_expires_at: "",
    password_hash: String(user.password_hash || ""),
    reset_code_hash: String(user.reset_code_hash || ""),
    reset_code_expires_at: String(user.reset_code_expires_at || ""),
    created_at: user.created_at ? new Date(user.created_at).toISOString() : new Date().toISOString(),
  }, { merge: true });

  if (company) {
    await firestore.collection("companies").doc(String(company.id)).set({
      id: Number(company.id),
      name: String(company.name || ""),
      name_lc: String(company.name || "").toLowerCase(),
      logo_data_url: String(company.logo_data_url || ""),
      created_at: company.created_at ? new Date(company.created_at).toISOString() : new Date().toISOString(),
    }, { merge: true });
  }

  console.log(JSON.stringify({
    ok: true,
    syncedUser: {
      id: Number(user.id),
      username: user.username,
      email: String(user.email || "").toLowerCase(),
    },
    syncedCompany: company ? {
      id: Number(company.id),
      name: company.name,
    } : null,
    projectId: process.env.FIREBASE_PROJECT_ID || serviceAccount.project_id,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
