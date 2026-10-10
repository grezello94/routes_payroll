const fs = require('node:fs');
const crypto = require('node:crypto');
const DEFAULT_FOLDER_ID = '1LNxVNKGYa3pQWNzcgrW9GyIDX_Y7XVyW';

function backupDriveName(createdAt) {
  const date = new Date(createdAt);
  if (Number.isNaN(date.getTime())) throw Error('Invalid backup date');
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(date).map((part) => [part.type, part.value]));
  const month = new Intl.DateTimeFormat('en-US', { timeZone: 'Asia/Kolkata', month: 'long' }).format(date);
  return `Routes-Payroll-Backup-${parts.year}-${parts.month}-${month}-${parts.year}-${parts.month}-${parts.day}_${parts.hour}-${parts.minute}-${parts.second}-IST.json`;
}

function createDriveBackupUploader({ env = process.env, fetchImpl = fetch } = {}) {
  const folderId = env.GOOGLE_DRIVE_BACKUP_FOLDER_ID || DEFAULT_FOLDER_ID;
  const configured = Boolean(env.GOOGLE_DRIVE_CLIENT_ID && env.GOOGLE_DRIVE_CLIENT_SECRET && env.GOOGLE_DRIVE_REFRESH_TOKEN);
  let token = null;
  let expiresAt = 0;
  let lastError = null;
  let lastUpload = null;
  async function request(url, options) {
    const response = await fetchImpl(url, { ...options, signal: AbortSignal.timeout(60000) });
    if (!response.ok) throw Error(`Google Drive request failed (HTTP ${response.status}). Check authorization, folder access, and storage quota.`);
    return response;
  }
  async function accessToken() {
    if (token && Date.now() < expiresAt) return token;
    const response = await request('https://oauth2.googleapis.com/token', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: env.GOOGLE_DRIVE_CLIENT_ID, client_secret: env.GOOGLE_DRIVE_CLIENT_SECRET, refresh_token: env.GOOGLE_DRIVE_REFRESH_TOKEN, grant_type: 'refresh_token' }),
    });
    const data = await response.json();
    if (!data.access_token) throw Error('Google authorization did not return an access token.');
    token = data.access_token;
    expiresAt = Date.now() + Math.max(0, Number(data.expires_in || 3600) - 60) * 1000;
    return token;
  }
  async function upload(file) {
    if (!configured) return null;
    try {
      const content = fs.readFileSync(file.fullPath);
      const key = crypto.createHash('sha256').update(file.fileName).update(content).digest('hex');
      const name = backupDriveName(file.createdAt);
      const headers = { Authorization: `Bearer ${await accessToken()}` };
      const query = new URLSearchParams({ q: `'${folderId.replace(/'/g, "\\'")}' in parents and trashed = false and appProperties has { key='payrollBackupKey' and value='${key}' }`, fields: 'files(id,name)', pageSize: '1' });
      const existing = await (await request(`https://www.googleapis.com/drive/v3/files?${query}`, { headers })).json();
      let remote = existing.files?.[0];
      if (!remote) {
        const boundary = `payroll_${crypto.randomBytes(16).toString('hex')}`;
        const metadata = { name, parents: [folderId], mimeType: 'application/json', appProperties: { payrollBackupKey: key } };
        const body = Buffer.concat([
          Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: application/json\r\n\r\n`),
          content, Buffer.from(`\r\n--${boundary}--\r\n`),
        ]);
        remote = await (await request('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
          method: 'POST', headers: { ...headers, 'Content-Type': `multipart/related; boundary=${boundary}` }, body,
        })).json();
      }
      if (!remote.id) throw Error('Google Drive did not confirm the uploaded backup.');
      lastUpload = { id: remote.id, name: remote.name || name, uploadedAt: new Date().toISOString() };
      lastError = null;
      return lastUpload;
    } catch (error) {
      lastError = error.message;
      throw error;
    }
  }
  return { upload, status: () => ({ configured, folderId, lastUpload, lastError }) };
}
module.exports = { createDriveBackupUploader, backupDriveName, DEFAULT_FOLDER_ID };
