// Restore missing collections from payroll-backup.json to Firestore
// Usage: node restore-firestore.js

const admin = require('firebase-admin');
const fs = require('fs');

const serviceAccount = require('./firebase-service-account.json');
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});
const db = admin.firestore();

const backup = JSON.parse(fs.readFileSync('payroll-backup.json', 'utf8'));

async function restoreCollection(collectionName, data) {
  if (!Array.isArray(data)) {
    console.log(`No array data for ${collectionName}`);
    return;
  }
  console.log(`Restoring ${data.length} documents to ${collectionName}...`);
  let count = 0;
  for (const doc of data) {
    const id = doc.id ? String(doc.id) : undefined;
    try {
      if (id) {
        await db.collection(collectionName).doc(id).set(doc);
      } else {
        await db.collection(collectionName).add(doc);
      }
      count++;
      if (count % 100 === 0) console.log(`  ${count}...`);
    } catch (e) {
      console.error(`Error writing to ${collectionName}:`, e.message);
    }
  }
  console.log(`Done: ${collectionName}`);
}

async function main() {
  await restoreCollection('employees', backup.employees);
  await restoreCollection('payroll_entries', backup.entries);
  await restoreCollection('designation_presets', backup.designations);
  console.log('All done!');
  process.exit(0);
}

main();
