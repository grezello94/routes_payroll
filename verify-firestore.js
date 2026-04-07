// Firestore verification script
// Usage: node verify-firestore.js

const admin = require('firebase-admin');
const fs = require('fs');

// Path to your service account key
const serviceAccount = require('./firebase-service-account.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

async function verifyCollection(collectionName, sample = 3) {
  const snapshot = await db.collection(collectionName).limit(sample).get();
  const countSnap = await db.collection(collectionName).count().get();
  console.log(`Collection: ${collectionName}`);
  console.log(`  Total documents: ${countSnap.data().count}`);
  snapshot.forEach(doc => {
    console.log(`  - ${doc.id}:`, doc.data());
  });
  console.log('');
}

async function main() {
  const collections = [
    'companies',
    'employees',
    'payroll_entries',
    'designation_presets'
  ];
  for (const col of collections) {
    try {
      await verifyCollection(col);
    } catch (e) {
      console.error(`Error reading collection ${col}:`, e.message);
    }
  }
  process.exit(0);
}

main();
