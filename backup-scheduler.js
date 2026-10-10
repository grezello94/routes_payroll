const fs = require('node:fs');
const path = require('node:path');
const BACKUP_INTERVAL_MS = 7 * 24 * 60 * 60 * 1000;

function createBackupScheduler({ directory, getPayload, getProvider, driveUploader = null, now = () => new Date() }) {
  let inFlight = null;
  let timer = null;
  let lastError = null;

  function list() {
    fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
    return fs.readdirSync(directory)
      .filter((name) => /^(weekly|monthly)-backup-\d{4}-\d{2}-.+\.json$/.test(name))
      .map((fileName) => {
        const fullPath = path.join(directory, fileName);
        const stat = fs.statSync(fullPath);
        return { fileName, fullPath, month: fileName.match(/^(?:weekly|monthly)-backup-(\d{4}-\d{2})-/)[1], createdAt: stat.mtime.toISOString(), sizeBytes: stat.size };
      })
      .filter((file) => file.sizeBytes > 0)
      .sort((a, b) => Date.parse(b.createdAt) - Date.parse(a.createdAt));
  }

  function status(createdThisCheck = false) {
    const files = list();
    const next = files.length ? new Date(Date.parse(files[0].createdAt) + BACKUP_INTERVAL_MS) : now();
    return {
      intervalDays: 7,
      createdThisCheck,
      latestBackup: files[0] || null,
      recentBackups: files.slice(0, 12),
      nextScheduledAt: next.toISOString(),
      nextScheduledDate: next.toISOString().slice(0, 10),
      overdue: next.getTime() <= now().getTime(),
      lastError,
      googleDrive: driveUploader?.status() || { configured: false },
    };
  }

  async function run() {
    if (!status().overdue) return status();
    const payload = await getPayload();
    const createdAt = now();
    const stamp = createdAt.toISOString().replace(/[:.]/g, '-');
    const fileName = `weekly-backup-${createdAt.toISOString().slice(0, 7)}-${stamp}.json`;
    const fullPath = path.join(directory, fileName);
    const temporaryPath = `${fullPath}.tmp`;
    try {
      fs.writeFileSync(temporaryPath, JSON.stringify({
        kind: 'routes-payroll-weekly-emergency-backup',
        createdAt: createdAt.toISOString(),
        backupMonth: createdAt.toISOString().slice(0, 7),
        provider: getProvider(),
        payload,
      }, null, 2), { encoding: 'utf8', mode: 0o600 });
      fs.renameSync(temporaryPath, fullPath);
      fs.utimesSync(fullPath, createdAt, createdAt);
    } finally {
      if (fs.existsSync(temporaryPath)) fs.unlinkSync(temporaryPath);
    }
    lastError = null;
    return status(true);
  }

  async function uploadPending() {
    if (!driveUploader?.status().configured) return;
    // Keep upload receipts next to the backups so restarts do not resend them.
    const pending = list().reverse().find((file) => !fs.existsSync(`${file.fullPath}.drive-uploaded`));
    if (!pending) return;
    const receipt = await driveUploader.upload(pending);
    if (receipt?.id) {
      fs.writeFileSync(`${pending.fullPath}.drive-uploaded`, JSON.stringify(receipt), { mode: 0o600 });
    }
  }

  function check() {
    if (inFlight) return inFlight;
    inFlight = run().then(async (result) => {
      try { await uploadPending(); } catch (error) { console.error("Google Drive backup upload failed:", error.message); }
      return { ...result, googleDrive: driveUploader?.status() || { configured: false } };
    }).catch((error) => {
      lastError = String(error.message || error);
      throw error;
    }).finally(() => { inFlight = null; });
    return inFlight;
  }

  function start() {
    if (timer) return;
    const tick = () => check().catch((error) => console.error('Automatic payroll backup failed:', error.message));
    tick();
    timer = setInterval(tick, 60 * 1000);
    timer.unref();
  }

  function stop() {
    if (timer) clearInterval(timer);
    timer = null;
  }

  return { check, list, status, start, stop };
}

module.exports = { createBackupScheduler, BACKUP_INTERVAL_MS };
