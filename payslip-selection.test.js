const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('app.js', 'utf8');
const nodes = { generatedPayslipsSelectionCount: {}, generatedPayslipsNetTotal: {}, generatedPayslipsCheckAll: {}, printSelectedPayslipsBtn: {} };
const context = vm.createContext({
  SELECTORS: nodes,
  activePayrollReportId: 1,
  activePayrollReportSnapshot: { month: '2026-09', records: [{ employeeId: 'E1', net: 27500 }, { employeeId: 'E2', net: 10700 }, { employeeId: 'E3', net: 23160 }, { employeeId: 'E4', net: 0 }, { employeeId: 'E5', net: 466.666666 }] },
  selectedGeneratedPayslipsByReport: new Map(),
  getSelectedCompanyId: () => 1,
  computePayroll: (record, month) => { assert.equal(month, '2026-09'); return { netSalary: record.net }; },
  formatCurrency: (amount) => new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR' }).format(amount),
  renderGeneratedPayrollReports: () => context.syncGeneratedPayslipSelectionControls(),
});
const start = source.indexOf('function getGeneratedPayslipSelectionKey(');
const end = source.indexOf('\nfunction getSelectedGeneratedPayslipData', start);
vm.runInContext(source.slice(start, end), context);
context.syncGeneratedPayslipSelectionControls();
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹0.00');
context.toggleGeneratedPayslipSelection('E1', true);
context.toggleGeneratedPayslipSelection('E2', true);
context.toggleGeneratedPayslipSelection('E3', true);
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹61,360.00');
context.toggleGeneratedPayslipSelection('E2', false);
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹50,660.00');
context.setAllGeneratedPayslipsSelected(true);
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹61,826.67');
assert.equal(nodes.generatedPayslipsCheckAll.checked, true);
context.setAllGeneratedPayslipsSelected(false);
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹0.00');
assert.equal(nodes.printSelectedPayslipsBtn.disabled, true);
context.activePayrollReportId = 2;
context.syncGeneratedPayslipSelectionControls();
assert.equal(nodes.generatedPayslipsNetTotal.textContent, '₹0.00', 'Selection is isolated by report');
console.log('Payslip selection tests passed: live sum, unchecking, check all, rounded currency, zero salaries, and report isolation.');
