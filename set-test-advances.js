// Set test advanceRemained values for first 5 employees in payroll_entries
// Usage: node set-test-advances.js

const admin = require('firebase-admin');
const serviceAccount = require('./firebase-service-account.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

async function setTestAdvances() {
  const snapshot = await db.collection('payroll_entries').orderBy('employeeId').limit(5).get();
  const testValues = [5000, 3000, 2000, 1000, 700];
  let i = 0;
  for (const doc of snapshot.docs) {
    await doc.ref.update({ advanceRemained: testValues[i] });
    console.log(`Set advanceRemained for ${doc.id} to ${testValues[i]}`);
    i++;
  }
  console.log('Test advances set.');
}

setTestAdvances().then(() => process.exit(0));
