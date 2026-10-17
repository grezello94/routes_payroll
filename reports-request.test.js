const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('app.js', 'utf8');

function extract(name) {
  const match = new RegExp(`^(?:async )?function ${name}\\(`, 'm').exec(source);
  assert.ok(match, `Missing Reports helper ${name}`);
  const remainder = source.slice(match.index + match[0].length);
  const next = /\n(?:async )?function \w+\(/.exec(remainder);
  return source.slice(match.index, next ? match.index + match[0].length + next.index : source.length);
}

function harness() {
  const nodes = Object.fromEntries([
    'generatedPayrollBody', 'generatedPayrollEmployeesBody', 'generatedPayrollDetailTitle',
    'reportsMonthPicker', 'reportsPayslipSearchInput', 'reportsPayslipCount',
    'reportsPayslipTotalCount', 'reportsPayslipNetTotal', 'reportsPayslipAdvanceTotal',
  ].map((name) => [name, { value: '', disabled: false }]));
  const context = vm.createContext({
    SELECTORS: nodes,
    companyId: 1,
    payrollReportsRequestVersion: 0,
    pendingPayrollReportRequest: null,
    activePayrollReportId: 1,
    activePayrollReportSnapshot: { month: '2026-08', company: { id: 1 }, records: [] },
    payrollReports: [{ id: 1, month: '2026-08', employeeCount: 0 }],
    requests: [],
    messages: [],
    getSelectedCompanyId: () => context.companyId,
    getSelectedMonth: () => '2026-09',
    apiRequest: (url) => new Promise((resolve, reject) => { context.requests.push({ url, resolve, reject }); }),
    showAppMessage: (message) => { context.messages.push(message); },
    renderReportsSummary() {},
    renderDashboardInsights() {},
    renderPayrollWorkflow() {},
    normalizeGeneratedPayslipSelection() {},
    syncGeneratedPayslipSelectionControls() {},
    getGeneratedPayslipRecordKey: (record, index) => record.employeeId || String(index),
    isGeneratedPayslipSelected: () => false,
    computePayroll: () => ({ netSalary: 1000, advanceRemained: 0 }),
    formatCurrency: (value) => String(value),
    formatDateTime: () => '',
    formatMonth: (month) => month,
    escapeHtml: (value) => String(value),
  });
  for (const name of ['isPayrollReportRequestPending', 'renderGeneratedPayrollReports', 'loadPayrollReports', 'openPayrollReport']) {
    vm.runInContext(extract(name), context);
  }
  return context;
}

function reportPayload(id, companyId = 1) {
  return { report: { id }, snapshot: { month: '2026-09', company: { id: companyId }, records: [{ employeeId: `E${id}`, employeeName: `Employee ${id}` }] } };
}

async function testOpenRequests() {
  {
    const context = harness();
    const pending = context.openPayrollReport(10);
    assert.equal(context.requests[0].url, '/api/payroll-reports/10?companyId=1');
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, true);
    context.SELECTORS.reportsPayslipSearchInput.value = 'Employee';
    context.renderGeneratedPayrollReports();
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, true, 'Search renders cannot unlock a pending report picker');
    assert.match(context.SELECTORS.generatedPayrollBody.innerHTML, /data-report-action="open" disabled/, 'Summary opens remain disabled during loading');
    context.requests[0].resolve(reportPayload(10));
    assert.equal(await pending, true);
    assert.equal(context.activePayrollReportId, 10);
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, false);
    assert.equal(context.pendingPayrollReportRequest, null);
  }
  for (const rejectStale of [false, true]) {
    const context = harness();
    const older = context.openPayrollReport(10);
    const newer = context.openPayrollReport(11);
    context.requests[1].resolve(reportPayload(11));
    assert.equal(await newer, true);
    if (rejectStale) context.requests[0].reject(new Error('Old request failed'));
    else context.requests[0].resolve(reportPayload(10));
    assert.equal(await older, false);
    assert.equal(context.activePayrollReportId, 11, 'An older response cannot replace the most recently requested report');
    assert.equal(context.activePayrollReportSnapshot.records[0].employeeId, 'E11');
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, false);
    assert.equal(context.messages.length, 0);
  }
  {
    const context = harness();
    const pending = context.openPayrollReport(10);
    context.companyId = 2;
    const currentSnapshot = { company: { id: 2 }, month: '2026-09', records: [] };
    context.activePayrollReportId = 20;
    context.activePayrollReportSnapshot = currentSnapshot;
    context.payrollReports = [{ id: 20, month: '2026-09' }];
    context.requests[0].resolve(reportPayload(10, 1));
    assert.equal(await pending, false);
    assert.equal(context.activePayrollReportId, 20);
    assert.equal(context.activePayrollReportSnapshot, currentSnapshot, 'A response from the previous company cannot replace the visible payslips');
    assert.equal(context.isPayrollReportRequestPending(), false);
  }
  {
    const context = harness();
    const previousSnapshot = context.activePayrollReportSnapshot;
    const pending = context.openPayrollReport(10);
    context.requests[0].reject(new Error('Report unavailable'));
    await assert.rejects(pending, /Report unavailable/);
    assert.equal(context.activePayrollReportSnapshot, previousSnapshot);
    assert.equal(context.pendingPayrollReportRequest, null);
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, false, 'A failed request restores the picker');
  }
}

async function testListRequests() {
  for (const rejectStale of [false, true]) {
    const context = harness();
    const older = context.loadPayrollReports(null, { openPreferredReport: false });
    const newer = context.loadPayrollReports(null, { openPreferredReport: false });
    context.requests[1].resolve({ reports: [{ id: 11, month: '2026-09' }] });
    assert.equal(await newer, true);
    if (rejectStale) context.requests[0].reject(new Error('Stale list failure'));
    else context.requests[0].resolve({ reports: [{ id: 10, month: '2026-08' }] });
    assert.equal(await older, false);
    assert.equal(context.payrollReports[0].id, 11, 'An older list cannot replace a newer list');
    assert.equal(context.activePayrollReportId, 11);
    assert.equal(context.activePayrollReportSnapshot, null, 'A list-only selection cannot expose a snapshot from another report');
    assert.equal(context.messages.length, 0);
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, false);
  }
  for (const rejectStale of [false, true]) {
    const context = harness();
    const pending = context.loadPayrollReports();
    context.companyId = 2;
    const companyReports = [{ id: 20, month: '2026-09' }];
    context.payrollReports = companyReports;
    context.activePayrollReportId = 20;
    if (rejectStale) context.requests[0].reject(new Error('Previous company failed'));
    else context.requests[0].resolve({ reports: [{ id: 10, month: '2026-08' }] });
    assert.equal(await pending, false);
    assert.equal(context.payrollReports, companyReports, 'A previous company list or error cannot clear the current reports');
    assert.equal(context.activePayrollReportId, 20);
    assert.equal(context.requests.length, 1, 'A stale list cannot fetch a report under the newly selected company');
    assert.equal(context.messages.length, 0);
  }
  {
    const context = harness();
    const pending = context.loadPayrollReports();
    context.requests[0].resolve({ reports: [{ id: 10, month: '2026-09' }] });
    await new Promise(setImmediate);
    assert.equal(context.requests[1].url, '/api/payroll-reports/10?companyId=1');
    context.renderGeneratedPayrollReports();
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, true, 'The picker remains disabled between list and snapshot responses');
    context.requests[1].resolve(reportPayload(10));
    assert.equal(await pending, true);
    assert.equal(context.activePayrollReportId, 10);
    assert.equal(context.pendingPayrollReportRequest, null);
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, false);
  }
  {
    const context = harness();
    const pending = context.loadPayrollReports();
    context.requests[0].resolve({ reports: [{ id: 10, month: '2026-09' }] });
    await new Promise(setImmediate);
    context.requests[1].reject(new Error('Saved report failed'));
    assert.equal(await pending, false);
    assert.equal(context.payrollReports.length, 0);
    assert.equal(context.activePayrollReportSnapshot, null);
    assert.equal(context.pendingPayrollReportRequest, null);
    assert.equal(context.messages[0], 'Saved report failed');
  }
}

async function testOpenAndListOrdering() {
  {
    const context = harness();
    const oldList = context.loadPayrollReports();
    const newOpen = context.openPayrollReport(11);
    context.requests[0].resolve({ reports: [{ id: 10, month: '2026-09' }] });
    assert.equal(await oldList, false);
    assert.equal(context.requests.length, 2, 'An older list cannot launch an unwanted snapshot after a newer explicit open');
    assert.equal(context.SELECTORS.reportsMonthPicker.disabled, true, 'An older request completion cannot unlock a newer pending open');
    context.requests[1].resolve(reportPayload(11));
    await newOpen;
    assert.equal(context.activePayrollReportId, 11);
  }
  {
    const context = harness();
    const oldOpen = context.openPayrollReport(10);
    const newList = context.loadPayrollReports();
    context.requests[1].resolve({ reports: [{ id: 11, month: '2026-09' }] });
    await new Promise(setImmediate);
    context.requests[2].resolve(reportPayload(11));
    assert.equal(await newList, true);
    context.requests[0].resolve(reportPayload(10));
    assert.equal(await oldOpen, false);
    assert.equal(context.activePayrollReportId, 11, 'A new list and its snapshot supersede previous open requests');
  }
}

(async () => {
  await testOpenRequests();
  await testListRequests();
  await testOpenAndListOrdering();
  console.log('Reports request tests passed: company isolation, ordered list/snapshot responses, persistent loading controls, and failure recovery.');
})().catch((error) => { console.error(error); process.exitCode = 1; });
