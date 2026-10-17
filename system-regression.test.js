const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const appSource = fs.readFileSync('app.js', 'utf8');
const serverSource = fs.readFileSync('server.js', 'utf8');
function extract(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `Missing function ${name}`);
  const end = source.indexOf('\nfunction ', start + 1);
  return source.slice(start, end < 0 ? source.length : end);
}
const browser = vm.createContext({ SELECTORS: { monthPicker: {} }, getPayrollMonthMode: () => 'previous' });
for (const name of ['setDefaultMonth', 'parseIsoDate', 'isoFromDate', 'monthBounds', 'shiftDateByDays', 'payrollWorkWindow', 'isPayrollBlocked', 'payrollMonthDayCount', 'systemWorkedDaysInMonth', 'computePayroll', 'toMoney', 'clamp']) {
  vm.runInContext(extract(appSource, name), browser);
}
for (const [today, expected] of [['2026-03-31', '2026-02'], ['2026-05-31', '2026-04'], ['2026-01-31', '2025-12']]) {
  browser.Date = class extends Date { constructor(...args) { super(...(args.length ? args : [`${today}T12:00:00`])); } };
  browser.setDefaultMonth();
  assert.equal(browser.SELECTORS.monthPicker.value, expected);
}
delete browser.Date;
const backend = vm.createContext({});
vm.runInContext(extract(serverSource, 'isIsoDate'), backend);
for (const invalid of ['2026-02-30', '2025-02-29', '2026-04-31', '2026-13-01', 'garbage']) {
  assert.equal(browser.parseIsoDate(invalid), null);
  assert.equal(backend.isIsoDate(invalid), false);
}
for (const valid of ['2024-02-29', '2026-10-10']) {
  assert.ok(browser.parseIsoDate(valid));
  assert.equal(backend.isIsoDate(valid), true);
}
const base = { presentSalary: 31000, increment: 0, employeeStatus: 'working' };
assert.equal(browser.computePayroll(base, '2026-10').netSalary, 31000);
assert.equal(browser.computePayroll({ ...base, joiningDate: '2026-10-16' }, '2026-10').workedDays, 16);
assert.equal(browser.computePayroll({ ...base, joiningDate: '2026-10-16', daysAbsent: 20 }, '2026-10').netSalary, 0);
assert.equal(browser.computePayroll({ ...base, employeeStatus: 'leave', leaveFrom: '2026-10-01' }, '2026-10').netSalary, 0);
assert.equal(browser.computePayroll({ ...base, oldAdvanceTaken: 5000, deductionEntered: 10000 }, '2026-10').advanceRemained, 0);
assert.equal(browser.computePayroll({ ...base, employeeStatus: 'resumed', leaveTo: '2026-10-20' }, '2026-10').workedDays, 12);
const middleware = [];
let staticCalls = 0;
const assets = vm.createContext({ app: { use: (handler) => middleware.push(handler) }, express: { static: () => () => { staticCalls++; } }, __dirname: '.' });
vm.runInContext(serverSource.slice(serverSource.indexOf('const PUBLIC_ASSETS'), serverSource.indexOf('\nfunction isValidMonth')), assets);
for (const path of ['/server.js', '/data/payroll.db', '/data/monthly-backups/test.json', '/package.json', '/.env.local', '/node_modules/test.js']) {
  let status;
  middleware[0]({ path }, { status: (code) => { status = code; return { end() {} }; } }, () => assert.fail(`Private file allowed: ${path}`));
  assert.equal(status, 404);
}
for (const path of ['/', '/app.js', '/styles.css', '/employee-workspace.css', '/payroll-workspace.css', '/dashboard-workspace.css', '/reports-workspace.css', '/workspace-states.css', '/icons/icon-192.svg']) {
  let passed = false;
  middleware[0]({ path }, {}, () => { passed = true; });
  assert.equal(passed, true);
}
middleware[1]({ path: '/server%2ejs' }, {}, () => {});
assert.equal(staticCalls, 0, 'Encoded private paths cannot reach static serving');
let header;
middleware[0]({ path: '/api/payroll/all' }, { set: (key, value) => { header = [key, value]; } }, () => {});
assert.deepEqual(header, ['Cache-Control', 'no-store']);
const listeners = {};
const worker = vm.createContext({ URL, self: { location: { origin: 'https://payroll.test' }, addEventListener: (name, handler) => { listeners[name] = handler; } } });
vm.runInContext(fs.readFileSync('sw.js', 'utf8'), worker);
listeners.fetch({ request: { method: 'GET', url: 'https://payroll.test/api/payroll/all' }, respondWith: () => assert.fail('API response was intercepted for caching') });
console.log('System regression tests passed: month rollover, date validation, payroll edge cases, private assets, and API cache isolation.');
