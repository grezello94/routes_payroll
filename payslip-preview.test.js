const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('app.js', 'utf8');
let releaseSave;
let saveBarrier = new Promise((resolve) => { releaseSave = resolve; });
let requests = 0;
let navigations = 0;
let updates = 0;
let month = '2026-09';
let shouldFail = false;
const messages = [];
const context = vm.createContext({
  currentRecords: [{ employeeId: 'E001', employeeName: 'Test Employee', presentSalary: 15000 }],
  getSelectedMonth: () => month,
  getSelectedCompanyId: () => 1,
  getActiveCompany: () => ({ id: 1, name: 'Test Company' }),
  openPayslip: () => assert.fail('Generation must not open the preview'),
  setWorkspace: () => { navigations++; assert.fail('Generation must stay in Payroll'); },
  setReportPanel: () => assert.fail('Generation must not switch report tabs'),
  SELECTORS: { railButtons: [{ dataset: { railAction: 'reports' }, classList: { toggle: (name, active) => assert.equal(active, true) } }], reportsSection: { scrollIntoView() {} } },
  showAppMessage: (message) => messages.push(message),
  flushPendingSave: () => saveBarrier,
  apiRequest: async (url, options) => {
    requests++;
    assert.equal(options.body.month, '2026-09');
    if (shouldFail) throw Error('Connection unavailable');
    return { report: { id: 7 }, snapshot: {} };
  },
  applyGeneratedPayrollResponse: () => { updates++; },
  renderPayrollWorkflow() {},
});
const start = source.indexOf('const pendingPayslipGenerations');
const end = source.indexOf('\nasync function generatePayslipsForEntries', start);
vm.runInContext(source.slice(start, end), context);
(async () => {
  const first = context.generatePayslipForCurrentRecord(0);
  assert.equal(navigations, 0, 'Wait for successful generation before opening reports');
  assert.equal(requests, 0);
  await context.generatePayslipForCurrentRecord(0);
  assert.equal(navigations, 0, 'Repeated clicks do not duplicate pending generation');
  month = '2026-10';
  releaseSave();
  await first;
  assert.equal(requests, 1);
  assert.equal(updates, 0, 'Previous-month response cannot replace current workspace');
  month = '2026-09';
  saveBarrier = Promise.resolve();
  shouldFail = true;
  await context.generatePayslipForCurrentRecord(0);
  assert.ok(messages.some((message) => message.includes('generation failed')));
  assert.equal(navigations, 0, 'Failed generation stays in payroll');
  shouldFail = false;
  await context.generatePayslipForCurrentRecord(0);
  assert.equal(updates, 1, 'Generation can be retried after failure');
  assert.equal(requests, 3);
  assert.equal(navigations, 0, "Successful generation stays in Payroll");
  assert.ok(messages.some((message) => message === "Payslip generated for Test Employee."));
  console.log('Payslip navigation tests passed: no automatic preview, stays in payroll after saving, duplicate prevention, workspace isolation, failure, and retry.');
})().catch((error) => { console.error(error); process.exitCode = 1; });
