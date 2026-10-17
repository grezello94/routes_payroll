const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('app.js', 'utf8');

function extract(name) {
  const declaration = new RegExp(`^(?:async )?function ${name}\\(`, 'm').exec(source);
  assert.ok(declaration, `Missing deletion helper: ${name}`);
  const remainder = source.slice(declaration.index + declaration[0].length);
  const next = /\n(?:async )?function \w+\(/.exec(remainder);
  return source.slice(declaration.index, next ? declaration.index + declaration[0].length + next.index : source.length);
}

class MockNode {
  constructor() {
    this.listeners = new Map();
    this.value = '';
    this.textContent = '';
    this.disabled = false;
    this.open = false;
    this.returnValue = '';
    this.dataset = {};
    this.classList = { add() {}, remove() {}, toggle() {} };
  }
  addEventListener(name, listener) {
    const listeners = this.listeners.get(name) || [];
    listeners.push(listener);
    this.listeners.set(name, listeners);
  }
  emit(name, overrides = {}) {
    const event = { target: this, defaultPrevented: false, preventDefault() { this.defaultPrevented = true; }, ...overrides };
    for (const listener of this.listeners.get(name) || []) listener(event);
    return event;
  }
  showModal() { this.open = true; }
  close(value = '') {
    this.open = false;
    this.returnValue = value;
    queueMicrotask(() => this.emit('close'));
  }
  focus() { this.focused = true; }
  reset() { this.value = ''; }
  setAttribute(name, value) { this[name] = value; }
  removeAttribute(name) { delete this[name]; }
}

function dialogContext() {
  const selectors = {};
  for (const name of ['Dialog', 'Form', 'Input', 'SubmitBtn', 'CancelBtn', 'CloseBtn', 'Title', 'Description', 'EmployeeName', 'EmployeeId']) {
    selectors[`deleteConfirmation${name}`] = new MockNode();
  }
  selectors.deleteConfirmationForm.reset = () => { selectors.deleteConfirmationInput.value = ''; };
  const context = vm.createContext({
    SELECTORS: selectors,
    pendingDeleteConfirmation: null,
    pendingEmployeeDeletions: new Set(),
    queueMicrotask,
    requestAnimationFrame: (callback) => callback(),
    setTimeout: (callback) => callback(),
    showAppMessage() {},
  });
  vm.runInContext(extract('wireDeleteConfirmation'), context);
  vm.runInContext(extract('confirmEmployeeDeletion'), context);
  context.wireDeleteConfirmation();
  return context;
}

function typeToken(context, token) {
  context.SELECTORS.deleteConfirmationInput.value = token;
  context.SELECTORS.deleteConfirmationInput.emit('input');
}

function submit(context) {
  return context.SELECTORS.deleteConfirmationForm.emit('submit');
}

async function testDialog() {
  const context = dialogContext();
  const nodes = context.SELECTORS;
  const first = context.confirmEmployeeDeletion({ employeeName: 'Deepak', employeeId: 'E001' });
  assert.equal(nodes.deleteConfirmationDialog.open, true);
  assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, true);
  assert.equal(await context.confirmEmployeeDeletion({ employeeName: 'Other', employeeId: 'E002' }), false, 'A second prompt cannot replace an active request');
  let settled = false;
  first.then(() => { settled = true; });
  for (const invalid of ['', 'CONFIRM', 'Confirm', 'confirm ', ' confirm', 'delete']) {
    typeToken(context, invalid);
    assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, true, `Token ${JSON.stringify(invalid)} must stay blocked`);
    submit(context);
    await Promise.resolve();
    await Promise.resolve();
    assert.equal(settled, false, 'Synthetic form submission cannot bypass the token');
    assert.equal(nodes.deleteConfirmationDialog.open, true);
  }
  typeToken(context, 'confirm');
  assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, false);
  nodes.deleteConfirmationSubmitBtn.disabled = true;
  submit(context);
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(settled, false, 'A synthetic submit cannot bypass a disabled confirmation button');
  typeToken(context, 'confirm');
  submit(context);
  assert.equal(await first, true);
  assert.equal(nodes.deleteConfirmationDialog.open, false);
  assert.equal(nodes.deleteConfirmationInput.value, '', 'Closing clears a previously valid token');
  assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, true);

  for (const action of ['cancel', 'close', 'escape', 'native-close']) {
    const pending = context.confirmEmployeeDeletion({ employeeName: 'Vinod', employeeId: 'E002' });
    assert.equal(nodes.deleteConfirmationInput.value, '');
    assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, true);
    typeToken(context, 'confirm');
    if (action === 'cancel') nodes.deleteConfirmationCancelBtn.emit('click');
    if (action === 'close') nodes.deleteConfirmationCloseBtn.emit('click');
    if (action === 'escape') {
      const event = nodes.deleteConfirmationDialog.emit('cancel');
      if (!event.defaultPrevented) nodes.deleteConfirmationDialog.close();
    }
    if (action === 'native-close') nodes.deleteConfirmationDialog.close();
    assert.equal(await pending, false, `${action} must leave the record intact`);
    assert.equal(nodes.deleteConfirmationInput.value, '');
    assert.equal(nodes.deleteConfirmationSubmitBtn.disabled, true);
  }

  const failed = dialogContext();
  failed.SELECTORS.deleteConfirmationDialog.showModal = () => { throw new Error('Dialog unavailable'); };
  assert.equal(await failed.confirmEmployeeDeletion({ employeeName: 'Deepak', employeeId: 'E001' }), false);
  assert.equal(failed.pendingDeleteConfirmation, null, 'A dialog-open failure cannot authorize a deletion or leave a pending request');
}

function actionContext() {
  const context = dialogContext();
  Object.assign(context, {
    companyId: 1,
    month: '2026-09',
    employeeMaster: [],
    employeeVerifications: [],
    currentRecords: [],
    activePayrollEmployeeId: '',
    requests: [],
    messages: [],
    saves: [],
    refreshes: [],
    getSelectedCompanyId: () => context.companyId,
    getSelectedMonth: () => context.month,
    formatMonth: (month) => month,
    apiRequest: async (url, options) => { context.requests.push({ url, ...options }); return {}; },
    loadEmployees: async () => { context.refreshes.push('employees'); },
    loadMonthRecords: async () => { context.refreshes.push('payroll'); },
    loadEmployeeVerifications: async () => { context.refreshes.push('verifications'); },
    renderPayrollTable: () => { context.refreshes.push('payroll-table'); },
    scheduleSave: (immediate) => { context.saves.push(immediate); },
    showAppMessage: (message) => { context.messages.push(message); },
    setEmployeeMessage: (message) => { context.messages.push(message); },
    employeeApiErrorMessage: (error) => error.message,
    setVerificationMessage: (message) => { context.messages.push(message); },
  });
  for (const name of ['deleteEmployeeWithConfirmation', 'removePayrollRecordWithConfirmation', 'deleteVerificationWithConfirmation']) {
    vm.runInContext(extract(name), context);
  }
  return context;
}

const actions = [
  { name: 'employee', helper: 'deleteEmployeeWithConfirmation', collection: 'employeeMaster', record: () => ({ id: 17, employeeId: 'E017', employeeName: 'Deepak' }), url: '/api/employees/17?companyId=1' },
  { name: 'payroll', helper: 'removePayrollRecordWithConfirmation', collection: 'currentRecords', record: () => ({ employeeId: 'E017', employeeName: 'Deepak' }) },
  { name: 'verification', helper: 'deleteVerificationWithConfirmation', collection: 'employeeVerifications', record: () => ({ id: 17, primaryBioData: { fullName: 'Deepak' } }), url: '/api/employee-verifications/17?companyId=1' },
];

function startAction(action) {
  const context = actionContext();
  const record = action.record();
  context[action.collection] = [record];
  context.activePayrollEmployeeId = record.employeeId;
  const pending = context[action.helper](record);
  return { context, record, pending };
}

function confirm(context) {
  typeToken(context, 'confirm');
  submit(context);
}

async function testDeletionActions() {
  for (const action of actions) {
    {
      const { context, record, pending } = startAction(action);
      assert.equal(context.requests.length, 0, `${action.name}: no request before confirmation`);
      assert.equal(context.saves.length, 0, `${action.name}: no payroll write before confirmation`);
      await context[action.helper](record);
      assert.equal(context.pendingEmployeeDeletions.size, 1, `${action.name}: duplicate action shares one pending operation`);
      assert.equal(context.SELECTORS.deleteConfirmationDialog.open, true);
      context.SELECTORS.deleteConfirmationCancelBtn.emit('click');
      await pending;
      assert.equal(context.requests.length, 0, `${action.name}: cancellation cannot send DELETE`);
      assert.equal(context[action.collection][0], record);
      assert.equal(context.saves.length, 0);
      assert.equal(context.pendingEmployeeDeletions.size, 0, `${action.name}: cancellation releases pending key`);
    }
    {
      const { context, record, pending } = startAction(action);
      confirm(context);
      await pending;
      if (action.url) {
        assert.equal(context.requests.length, 1, `${action.name}: one confirmed DELETE`);
        assert.equal(context.requests[0].url, action.url);
        assert.equal(context.requests[0].method, 'DELETE');
      } else {
        assert.equal(context.currentRecords.length, 0, 'Confirmed payroll deletion removes the target');
        assert.equal(context.activePayrollEmployeeId, '');
        assert.deepEqual(context.saves, [true]);
      }
      assert.equal(context.pendingEmployeeDeletions.size, 0, `${action.name}: success releases pending key`);
    }
    {
      const { context, record, pending } = startAction(action);
      context.companyId = 2;
      confirm(context);
      await pending;
      assert.equal(context.requests.length, 0, `${action.name}: switching company invalidates confirmation`);
      assert.equal(context.saves.length, 0);
      assert.equal(context[action.collection][0], record);
      assert.equal(context.pendingEmployeeDeletions.size, 0);
    }
    {
      const { context, record, pending } = startAction(action);
      const replacement = { ...record };
      context[action.collection] = [replacement];
      confirm(context);
      await pending;
      assert.equal(context.requests.length, 0, `${action.name}: replacement objects are not deleted by stale prompts`);
      assert.equal(context.saves.length, 0);
      assert.equal(context[action.collection][0], replacement);
      assert.equal(context.pendingEmployeeDeletions.size, 0);
    }
  }

  const payroll = actions.find((action) => action.name === 'payroll');
  {
    const { context, record, pending } = startAction(payroll);
    context.month = '2026-10';
    confirm(context);
    await pending;
    assert.equal(context.currentRecords[0], record, 'Switching payroll month invalidates confirmation');
    assert.equal(context.saves.length, 0);
  }
  {
    const { context, record, pending } = startAction(payroll);
    const other = { employeeId: 'E018', employeeName: 'Vinod' };
    context.currentRecords = [other, record];
    confirm(context);
    await pending;
    assert.equal(context.currentRecords.length, 1);
    assert.equal(context.currentRecords[0], other, 'Reordering while the prompt is open still removes only the target object');
    assert.deepEqual(context.saves, [true]);
  }

  for (const action of actions.filter((action) => action.url)) {
    {
      const { context, record, pending } = startAction(action);
      let finishRequest;
      context.apiRequest = (url, options) => {
        context.requests.push({ url, ...options });
        return new Promise((resolve) => { finishRequest = resolve; });
      };
      confirm(context);
      await Promise.resolve();
      await Promise.resolve();
      assert.equal(context.requests.length, 1);
      await context[action.helper](record);
      assert.equal(context.SELECTORS.deleteConfirmationDialog.open, false, `${action.name}: an in-flight duplicate cannot open a new prompt`);
      assert.equal(context.requests.length, 1);
      context.companyId = 2;
      finishRequest({});
      await pending;
      assert.equal(context.refreshes.length, 0, `${action.name}: stale responses cannot refresh the newly selected company`);
      assert.equal(context.messages.length, 0);
      assert.equal(context.pendingEmployeeDeletions.size, 0);
    }
    {
      const { context, pending } = startAction(action);
      context.apiRequest = async () => { throw new Error('Deletion failed'); };
      confirm(context);
      await pending;
      assert.equal(context.messages.at(-1), 'Deletion failed');
      assert.equal(context.pendingEmployeeDeletions.size, 0, `${action.name}: failed requests release the pending key for retry`);
    }
  }
}

async function run() {
  await testDialog();
  await testDeletionActions();
  console.log('Deletion confirmation tests passed: exact token, disabled submit, cancel/escape, reset, duplicate operations, target identity, company/month changes, and request failures.');
}
run().catch((error) => { console.error(error); process.exitCode = 1; });
