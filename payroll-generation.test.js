const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('server.js', 'utf8');
const handlers = {};
const records = Array.from({ length: 27 }, (_, index) => ({ employeeId: `E${index + 1}`, positionIndex: index, presentSalary: 15000 + index }));
let saves = 0;
let builds = 0;
let payload;
const previousRecord = { employeeId: 'E27', positionIndex: 26, presentSalary: 9000 };
const context = vm.createContext({
  app: { post: (route, ...args) => { handlers[route] = args.at(-1); } },
  authMiddleware() {},
  isValidMonth: (month) => /^\d{4}-(0[1-9]|1[0-2])$/.test(month),
  sanitizeText: (value, max) => String(value ?? '').trim().slice(0, max),
  resolveCompanyId: async () => 1,
  ensurePayrollCheckedForMonth: async () => { builds++; return { checkedAt: '2026-10-10', snapshot: { company: { id: 1 }, month: '2026-09', records } }; },
  parseSnapshotJson: (value) => JSON.parse(value || '{}'),
  dbRowToPayrollReport: (row) => row,
  isStoreUnavailableError: () => false,
  console,
  store: {
    getPayrollReportByMonthCompany: async () => ({ snapshot_json: JSON.stringify({ records: [previousRecord] }) }),
    savePayrollReportSnapshot: async (month, companyId, data) => { saves++; payload = data; return { id: 7 }; },
    getPayrollReportByIdCompany: async () => ({ id: 7 }),
  },
});
const start = source.indexOf('app.post("/api/payroll-reports/generate",');
const end = source.indexOf('app.post("/api/payroll/restore",', start);
vm.runInContext(source.slice(start, end), context);
async function call(route, body) {
  let status = 200;
  let result;
  await handlers[route]({ body }, { status(code) { status = code; return this; }, json(data) { result = data; } });
  return { status, result };
}
(async () => {
  let response = await call('/api/payroll-reports/generate', { month: '2026-09', employeeIds: records.map((record) => record.employeeId) });
  assert.equal(response.status, 200);
  assert.equal(response.result.generatedCount, 27);
  assert.equal(saves, 1, '27 employees must use one report save');
  assert.equal(builds, 1, '27 employees must use one payroll build');
  assert.equal(payload.employeeCount, 27);
  response = await call('/api/payroll-reports/generate', { month: '2026-09', employeeIds: ['E1', 'E1', 'E2'] });
  assert.equal(response.result.generatedCount, 2);
  assert.equal(payload.snapshot.records.length, 3);
  assert.equal(payload.snapshot.records.find((record) => record.employeeId === 'E27').presentSalary, 9000, 'Filtered generation preserves unselected payslips');
  const before = saves;
  assert.equal((await call('/api/payroll-reports/generate', { month: '2026-09', employeeIds: ['missing'] })).status, 404);
  assert.equal((await call('/api/payroll-reports/generate', { month: '2026-09', employeeIds: [] })).status, 400);
  assert.equal(saves, before, 'Invalid selections must not save reports');
  response = await call('/api/payroll-reports/generate', { month: '2026-09' });
  assert.equal(response.result.generatedCount, 27, 'Existing whole-month API remains supported');
  const beforeBuild = builds;
  response = await call('/api/payroll-reports/generate-entry', { month: '2026-09', employeeId: 'E1' });
  assert.equal(response.status, 200);
  assert.equal(builds - beforeBuild, 1, 'Single payslip must not rebuild payroll twice');
  console.log('Payroll generation tests passed: one build/save for 27 employees, filtered selections, deduplication, invalid IDs, and single generation.');
})().catch((error) => { console.error(error); process.exitCode = 1; });
