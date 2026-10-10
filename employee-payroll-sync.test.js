const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { effectiveSalary } = require('./salary-progression');
const server = fs.readFileSync('server.js', 'utf8');
const client = fs.readFileSync('app.js', 'utf8');
let employee = { employee_id: 'E001', employee_name: 'Old Name', designation: 'Chef', base_salary: 15000, joining_date: '2026-01-01', opening_advance: 1000 };
const stored = { id: 1, employee_id: 'E001', employee_name: 'Old Name', designation: 'Chef', increment: 500, extra_advance_added: 100, deduction_entered: 200, days_absent: 2, comment: 'Keep note' };
const history = [{ employee_id: 'E001', company_id: 1, month: '2026-08', present_salary: 15000, increment: 1000 }];
const context = vm.createContext({
  store: {
    listPayrollByMonthCompany: async () => [stored],
    listEmployeesByCompany: async () => [employee],
    listPayrollEntriesAll: async () => history,
  },
  effectiveSalary,
  buildAdvanceLedger: () => ({ latestAdvanceByEmployeeId: new Map([['E001', 800]]) }),
  resolveEmployeeStatusForMonth: (item) => ({ employeeStatus: item.status || 'working', leaveFrom: item.leave_from || '', leaveTo: '', terminatedOn: '' }),
});
const start = server.indexOf('async function listMergedPayrollMonthRecords');
const end = server.indexOf('\nasync function buildPayrollReportRecord', start);
vm.runInContext(server.slice(start, end), context);
(async () => {
  const first = (await context.listMergedPayrollMonthRecords('2026-09', 1))[0];
  assert.equal(first.presentSalary, 16000);
  employee = { ...employee, employee_name: 'Updated Name', designation: 'Head Chef', base_salary: 18000, joining_date: '2026-02-01', nationality: 'Indian', status: 'leave', leave_from: '2026-09-20' };
  const updated = (await context.listMergedPayrollMonthRecords('2026-09', 1))[0];
  assert.equal(updated.employeeName, 'Updated Name');
  assert.equal(updated.designation, 'Head Chef', 'Old monthly designation cannot override employee edits');
  assert.equal(updated.presentSalary, 18000);
  assert.equal(updated.joiningDate, '2026-02-01');
  assert.equal(updated.nationality, 'Indian');
  assert.equal(updated.employeeStatus, 'leave');
  assert.equal(updated.leaveFrom, '2026-09-20');
  assert.equal(updated.increment, 500);
  assert.equal(updated.extraAdvanceAdded, 100);
  assert.equal(updated.deductionEntered, 200);
  assert.equal(updated.daysAbsent, 2);
  assert.equal(updated.comment, 'Keep note');
  assert.equal(updated.oldAdvanceTaken, 800);
  const browser = vm.createContext({
    getSelectedMonth: () => '2026-09', payrollMonthUrl: () => '/test',
    apiRequest: async () => ({ records: [{ ...updated, id: null, presentSalary: 19000, employeeName: 'Stale Name', designation: 'Stale Role' }] }),
    employeeMaster: [{ employeeId: 'E001', employeeName: 'Updated Name', designation: 'Head Chef', baseSalary: 18000 }],
    resolveEmployeeStatusForMonth: () => ({ employeeStatus: 'working' }),
    isPayrollVisibleForMonth: () => true,
    activePayrollEmployeeId: '', renderPayrollTable() {}, setSaveStatus() {}, renderPayrollWorkflow() {}, renderDashboardInsights() {}, showAppMessage(message) { throw Error(message); },
  });
  const clientStart = client.indexOf('async function loadMonthRecords()');
  const clientEnd = client.indexOf('\nfunction payrollReportsUrl', clientStart);
  vm.runInContext(client.slice(clientStart, clientEnd), browser);
  await browser.loadMonthRecords();
  assert.equal(browser.currentRecords[0].presentSalary, 19000, 'Carried-forward salary survives loading an unsaved month');
  assert.equal(browser.currentRecords[0].employeeName, 'Updated Name');
  assert.equal(browser.currentRecords[0].designation, 'Head Chef');
  console.log('Employee payroll sync tests passed: current employee details, salary carry-forward, and preserved monthly payroll inputs.');
})().catch((error) => { console.error(error); process.exitCode = 1; });
