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

async function deleteIfExistsByEmail(email) {
  try {
    const user = await admin.auth().getUserByEmail(email);
    await admin.auth().deleteUser(user.uid);
  } catch (error) {
    if (error?.code !== "auth/user-not-found") {
      throw error;
    }
  }
}

async function main() {
  const user = await getSqliteRow(
    `SELECT id, username, email, password_hash, created_at
     FROM users
     ORDER BY id
     LIMIT 1`
  );

  if (!user) {
    console.error("No local SQLite admin user found to sync.");
    process.exit(1);
  }

  const uid = String(user.id);
  await deleteIfExistsByEmail(String(user.email || "").toLowerCase());
  try {
    await admin.auth().deleteUser(uid);
  } catch (error) {
    if (error?.code !== "auth/user-not-found") {
      throw error;
    }
  }

  const result = await admin.auth().importUsers(
    [
      {
        uid,
        email: String(user.email || "").toLowerCase(),
        displayName: String(user.username || ""),
        emailVerified: true,
        passwordHash: Buffer.from(String(user.password_hash || ""), "utf8"),
      },
    ],
    {
      hash: {
        algorithm: "BCRYPT",
      },
    }
  );

  if (result.errors?.length) {
    console.error(JSON.stringify(result.errors, null, 2));
    process.exit(1);
  }

  console.log(JSON.stringify({
    ok: true,
    importedUser: {
      uid,
      username: user.username,
      email: String(user.email || "").toLowerCase(),
    },
    projectId: process.env.FIREBASE_PROJECT_ID || serviceAccount.project_id,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
