const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createDriveBackupUploader, backupDriveName, DEFAULT_FOLDER_ID } = require('./google-drive-backup');
const { createBackupScheduler } = require('./backup-scheduler');
const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'payroll-drive-test-'));
const file = { fullPath: path.join(directory, 'test.json'), fileName: 'test.json', createdAt: '2026-10-10T06:26:56Z' };
fs.writeFileSync(file.fullPath, '{"employees":[]}');
assert.equal(backupDriveName(file.createdAt), 'Routes-Payroll-Backup-2026-10-October-2026-10-10_11-56-56-IST.json');
assert.ok(backupDriveName('2026-10-31T20:00:00Z').includes('2026-11-November-2026-11-01'));
let uploads = 0;
let tokens = 0;
let duplicate = false;
const uploader = createDriveBackupUploader({
  env: { GOOGLE_DRIVE_CLIENT_ID: 'test', GOOGLE_DRIVE_CLIENT_SECRET: 'test', GOOGLE_DRIVE_REFRESH_TOKEN: 'test' },
  fetchImpl: async (url, options) => {
    assert.ok(options.signal);
    if (url.includes('oauth2')) { tokens++; return { ok: true, json: async () => ({ access_token: 'test', expires_in: 3600 }) }; }
    if (url.includes('/upload/')) {
      uploads++;
      const body = options.body.toString();
      assert.ok(body.includes(DEFAULT_FOLDER_ID));
      assert.ok(body.includes(backupDriveName(file.createdAt)));
      assert.ok(body.includes('{"employees":[]}'));
      return { ok: true, json: async () => ({ id: 'drive-id', name: backupDriveName(file.createdAt) }) };
    }
    return { ok: true, json: async () => ({ files: duplicate ? [{ id: 'drive-id' }] : [] }) };
  },
});
(async () => {
  assert.equal((await uploader.upload(file)).id, 'drive-id');
  duplicate = true;
  assert.equal((await uploader.upload(file)).id, 'drive-id');
  assert.equal(uploads, 1, 'Remote matching prevents duplicate uploads');
  assert.equal(tokens, 1, 'Access token is reused');
  const disabled = createDriveBackupUploader({ env: {}, fetchImpl: () => assert.fail('Unauthenticated uploader made a request') });
  assert.equal(await disabled.upload(file), null);
  let attempts = 0;
  const scheduler = createBackupScheduler({ directory, getProvider: () => 'test', getPayload: async () => ({ entries: [] }), driveUploader: {
    status: () => ({ configured: true }),
    upload: async () => { attempts++; if (attempts === 1) throw Error('Mock outage'); return { id: 'drive-id' }; },
  } });
  await scheduler.check();
  assert.equal(scheduler.list().length, 1, 'Upload outage cannot prevent the local backup');
  await scheduler.check();
  assert.equal(attempts, 2, 'Pending upload retries without waiting seven days');
  await scheduler.check();
  assert.equal(attempts, 2, 'Successful upload receipt prevents repeat uploads');
  console.log('Drive backup tests passed: month names in IST, folder targeting, OAuth reuse, deduplication, disabled uploads, and retry after outage.');
})().catch((error) => { console.error(error); process.exitCode = 1; }).finally(() => fs.rmSync(directory, { recursive: true, force: true }));
