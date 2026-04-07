// Fix advanceRemained in payroll_entries based on oldAdvanceTaken, extraAdvanceAdded, deductionEntered
// Usage: node fix-advanceRemained.js

const admin = require('firebase-admin');
const serviceAccount = require('./firebase-service-account.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

async function fixPayrollAdvances() {
  const snapshot = await db.collection('payroll_entries').get();
  let updated = 0;
  for (const doc of snapshot.docs) {
    const data = doc.data();
    const oldAdvance = Number(data.oldAdvanceTaken ?? data.old_advance_taken ?? 0);
    const extraAdvance = Number(data.extraAdvanceAdded ?? data.extra_advance_added ?? 0);
    const deduction = Number(data.deductionEntered ?? data.deduction_entered ?? 0);
    const totalAdvance = oldAdvance + extraAdvance;
    const deductionApplied = Math.min(deduction, totalAdvance);
    const advanceRemained = totalAdvance - deductionApplied;
    if (data.advanceRemained !== advanceRemained) {
      await doc.ref.update({ advanceRemained });
      updated++;
    }
  }
  console.log(`Updated advanceRemained for ${updated} payroll_entries.`);
}

fixPayrollAdvances().then(() => process.exit(0));
