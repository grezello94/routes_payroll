const STORAGE_KEYS = {
  authToken: "routes_payroll_auth_token_v1",
};

const SELECTORS = {
  authView: document.getElementById("authView"),
  appView: document.getElementById("appView"),
  setupForm: document.getElementById("setupForm"),
  loginForm: document.getElementById("loginForm"),
  authMessage: document.getElementById("authMessage"),
  appMessage: document.getElementById("appMessage"),
  saveStatus: document.getElementById("saveStatus"),
  companyPicker: document.getElementById("companyPicker"),
  addCompanyBtn: document.getElementById("addCompanyBtn"),
  companyDialog: document.getElementById("companyDialog"),
  closeCompanyBtn: document.getElementById("closeCompanyBtn"),
  companyForm: document.getElementById("companyForm"),
  companyNameInput: document.getElementById("companyNameInput"),
  companyLogoInput: document.getElementById("companyLogoInput"),
  companyMessage: document.getElementById("companyMessage"),
  monthPicker: document.getElementById("monthPicker"),
  payrollBody: document.getElementById("payrollBody"),
  addEmployeeBtn: document.getElementById("addEmployeeBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  backupBtn: document.getElementById("backupBtn"),
  restoreInput: document.getElementById("restoreInput"),
  logoutBtn: document.getElementById("logoutBtn"),
  metricEmployees: document.getElementById("metricEmployees"),
  metricGross: document.getElementById("metricGross"),
  metricNet: document.getElementById("metricNet"),
  metricAdvance: document.getElementById("metricAdvance"),
  installBtn: document.getElementById("installBtn"),
  actionMenu: document.getElementById("actionMenu"),
  railButtons: Array.from(document.querySelectorAll("[data-rail-action]")),
  dashboardSection: document.getElementById("dashboardSection"),
  metricsSection: document.getElementById("metricsSection"),
  payrollSection: document.getElementById("payrollSection"),
  payslipDialog: document.getElementById("payslipDialog"),
  payslipPreview: document.getElementById("payslipPreview"),
  closePayslipBtn: document.getElementById("closePayslipBtn"),
  printPayslipBtn: document.getElementById("printPayslipBtn"),
  downloadPayslipBtn: document.getElementById("downloadPayslipBtn"),
  shareWebBtn: document.getElementById("shareWebBtn"),
  shareWhatsappBtn: document.getElementById("shareWhatsappBtn"),
  shareMessengerBtn: document.getElementById("shareMessengerBtn"),
  copyPayslipBtn: document.getElementById("copyPayslipBtn"),
};

let pendingInstallPrompt = null;
let authToken = localStorage.getItem(STORAGE_KEYS.authToken) || "";
let activePayslip = null;
let currentRecords = [];
let companies = [];
let activeCompanyId = null;
let pendingSaveTimer = null;
let saveInFlight = false;
let saveQueued = false;

init();

async function init() {
  wireAuth();
  wireAppActions();
  setDefaultMonth();
  registerServiceWorker();
  wireInstallPrompt();
  await updateAuthView();
}

function wireAuth() {
  SELECTORS.setupForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const companyName = document.getElementById("setupCompanyName").value.trim();
    const username = document.getElementById("setupUsername").value.trim();
    const password = document.getElementById("setupPassword").value;
    const confirm = document.getElementById("setupConfirm").value;

    if (companyName.length < 2) {
      showAuthMessage("Company name is too short.");
      return;
    }

    if (password !== confirm) {
      showAuthMessage("Passwords do not match.");
      return;
    }

    try {
      const response = await apiRequest("/api/auth/register", {
        method: "POST",
        body: { companyName, username, password },
      }, false);

      setToken(response.token);
      SELECTORS.setupForm.reset();
      showAuthMessage("");
      await switchToApp();
    } catch (error) {
      showAuthMessage(error.message);
    }
  });

  SELECTORS.loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const username = document.getElementById("loginUsername").value.trim();
    const password = document.getElementById("loginPassword").value;

    try {
      const response = await apiRequest("/api/auth/login", {
        method: "POST",
        body: { username, password },
      }, false);

      setToken(response.token);
      SELECTORS.loginForm.reset();
      showAuthMessage("");
      await switchToApp();
    } catch (error) {
      showAuthMessage(error.message);
    }
  });
}

function wireAppActions() {
  wireRailActions();
  wireCompanyActions();

  SELECTORS.monthPicker.addEventListener("change", async () => {
    await flushPendingSave();
    await loadMonthRecords();
  });

  SELECTORS.addEmployeeBtn.addEventListener("click", async () => {
    currentRecords.push(createEmptyEmployee(currentRecords));
    renderPayrollTable();
    scheduleSave(true);
    showAppMessage("Employee row added.");
  });

  SELECTORS.exportCsvBtn.addEventListener("click", () => {
    const month = getSelectedMonth();
    const csv = makeCsv(currentRecords, month);
    downloadBlob(
      new Blob([csv], { type: "text/csv;charset=utf-8" }),
      `payroll-${month}.csv`
    );
    showAppMessage("CSV exported.");
    closeActionMenu();
  });

  SELECTORS.exportExcelBtn.addEventListener("click", () => {
    const month = getSelectedMonth();
    const excelXml = makeExcelSpreadsheet(currentRecords, month);
    downloadBlob(
      new Blob([excelXml], { type: "application/vnd.ms-excel;charset=utf-8" }),
      `payroll-${month}.xls`
    );
    showAppMessage("Excel exported.");
    closeActionMenu();
  });

  SELECTORS.backupBtn.addEventListener("click", async () => {
    try {
      const response = await apiRequest("/api/payroll/all");
      const backup = JSON.stringify(response, null, 2);
      downloadBlob(
        new Blob([backup], { type: "application/json" }),
        `payroll-backup-${new Date().toISOString().slice(0, 10)}.json`
      );
      showAppMessage("Database backup downloaded.");
      closeActionMenu();
    } catch (error) {
      showAppMessage(error.message);
    }
  });

  SELECTORS.restoreInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      await apiRequest("/api/payroll/restore", {
        method: "POST",
        body: parsed,
      });
      await loadCompanies();
      await loadMonthRecords();
      showAppMessage("Database backup restored.");
      closeActionMenu();
    } catch (error) {
      showAppMessage(`Restore failed: ${error.message}`);
    } finally {
      event.target.value = "";
    }
  });

  SELECTORS.logoutBtn.addEventListener("click", async () => {
    setToken("");
    currentRecords = [];
    renderPayrollTable();
    await updateAuthView();
    showAuthMessage("Logged out.");
  });

  SELECTORS.payrollBody.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) {
      return;
    }

    const row = target.closest("tr");
    if (!row) return;
    const index = Number(row.dataset.index);
    const field = target.dataset.field;
    if (Number.isNaN(index) || !field || !currentRecords[index]) return;

    currentRecords[index][field] = target.value;
    scheduleSave(false);
  });

  SELECTORS.payrollBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) {
      return;
    }
    const row = target.closest("tr");
    if (!row) return;
    const index = Number(row.dataset.index);
    if (Number.isNaN(index) || !currentRecords[index]) return;
    renderPayrollTable();
    scheduleSave(true);
  });

  SELECTORS.payrollBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button");
    if (!button) return;
    const row = button.closest("tr");
    if (!row) return;
    const index = Number(row.dataset.index);
    if (Number.isNaN(index) || !currentRecords[index]) return;

    const action = button.dataset.action;
    if (action === "delete") {
      currentRecords.splice(index, 1);
      renderPayrollTable();
      scheduleSave(true);
      showAppMessage("Employee removed.");
      return;
    }

    if (action === "payslip") {
      openPayslip(currentRecords[index], getSelectedMonth());
    }
  });

  SELECTORS.closePayslipBtn.addEventListener("click", () => SELECTORS.payslipDialog.close());
  SELECTORS.printPayslipBtn.addEventListener("click", printCurrentPayslip);
  SELECTORS.downloadPayslipBtn.addEventListener("click", downloadCurrentPayslipText);
  SELECTORS.shareWhatsappBtn.addEventListener("click", shareCurrentPayslipWhatsapp);
  SELECTORS.shareMessengerBtn.addEventListener("click", shareCurrentPayslipMessenger);
  SELECTORS.shareWebBtn.addEventListener("click", shareCurrentPayslipWeb);
  SELECTORS.copyPayslipBtn.addEventListener("click", copyCurrentPayslip);
}

function wireCompanyActions() {
  SELECTORS.companyPicker.addEventListener("change", async () => {
    await flushPendingSave();
    activeCompanyId = Number(SELECTORS.companyPicker.value) || null;
    await loadMonthRecords();
  });

  SELECTORS.addCompanyBtn.addEventListener("click", () => {
    SELECTORS.companyForm.reset();
    SELECTORS.companyMessage.textContent = "";
    SELECTORS.companyDialog.showModal();
  });

  SELECTORS.closeCompanyBtn.addEventListener("click", () => {
    SELECTORS.companyDialog.close();
  });

  SELECTORS.companyForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const name = SELECTORS.companyNameInput.value.trim();
    const file = SELECTORS.companyLogoInput.files?.[0];
    let logoDataUrl = "";

    if (file) {
      if (file.size > 700 * 1024) {
        SELECTORS.companyMessage.textContent = "Logo too large. Keep it under 700KB.";
        return;
      }
      logoDataUrl = await fileToDataUrl(file);
    }

    try {
      const response = await apiRequest("/api/companies", {
        method: "POST",
        body: { name, logoDataUrl },
      });
      activeCompanyId = response.company?.id || activeCompanyId;
      await loadCompanies();
      await loadMonthRecords();
      SELECTORS.companyDialog.close();
      showAppMessage("Company created.");
    } catch (error) {
      SELECTORS.companyMessage.textContent = error.message;
    }
  });
}

function wireRailActions() {
  SELECTORS.railButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.railAction;
      SELECTORS.railButtons.forEach((item) => item.classList.remove("active"));
      button.classList.add("active");

      if (action === "dashboard") {
        SELECTORS.dashboardSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }

      if (action === "employees") {
        if (currentRecords.length === 0) {
          SELECTORS.addEmployeeBtn.click();
        }
        SELECTORS.payrollSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        setTimeout(() => {
          const firstNameField = SELECTORS.payrollBody.querySelector('input[data-field="employeeName"]');
          firstNameField?.focus();
        }, 260);
        return;
      }

      if (action === "payroll") {
        SELECTORS.payrollSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        setTimeout(() => {
          const salaryField = SELECTORS.payrollBody.querySelector('input[data-field="presentSalary"]');
          salaryField?.focus();
        }, 240);
        return;
      }

      if (action === "reports") {
        SELECTORS.actionMenu?.setAttribute("open", "");
        SELECTORS.actionMenu?.scrollIntoView({ behavior: "smooth", block: "center" });
        return;
      }

      if (action === "settings") {
        SELECTORS.monthPicker?.focus();
        SELECTORS.metricsSection?.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  });
}

function closeActionMenu() {
  SELECTORS.actionMenu?.removeAttribute("open");
}

async function updateAuthView() {
  if (authToken) {
    const valid = await validateToken();
    if (valid) {
      await switchToApp();
      return;
    }
  }

  SELECTORS.authView.classList.remove("hidden");
  SELECTORS.appView.classList.add("hidden");

  try {
    const bootstrap = await apiRequest("/api/auth/bootstrap", {}, false);
    if (bootstrap.needsSetup) {
      SELECTORS.setupForm.classList.remove("hidden");
      SELECTORS.loginForm.classList.add("hidden");
    } else {
      SELECTORS.setupForm.classList.add("hidden");
      SELECTORS.loginForm.classList.remove("hidden");
    }
  } catch {
    SELECTORS.setupForm.classList.add("hidden");
    SELECTORS.loginForm.classList.add("hidden");
    showAuthMessage("Server is not reachable. Start backend with: node server.js");
  }
}

async function validateToken() {
  if (!authToken) return false;
  try {
    await apiRequest("/api/auth/me");
    return true;
  } catch (error) {
    if (String(error.message).toLowerCase().includes("expired") || String(error.message).toLowerCase().includes("authentication")) {
      setToken("");
    }
    return false;
  }
}

async function switchToApp() {
  SELECTORS.authView.classList.add("hidden");
  SELECTORS.appView.classList.remove("hidden");
  await loadCompanies();
  await loadMonthRecords();
}

function setDefaultMonth() {
  const now = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  SELECTORS.monthPicker.value = month;
}

function getSelectedMonth() {
  return SELECTORS.monthPicker.value || new Date().toISOString().slice(0, 7);
}

function getSelectedCompanyId() {
  if (activeCompanyId) return activeCompanyId;
  const parsed = Number(SELECTORS.companyPicker.value);
  if (Number.isInteger(parsed) && parsed > 0) return parsed;
  return companies[0]?.id || 1;
}

function displayCompanyName(company) {
  const raw = String(company?.name ?? "").trim();
  if (raw.length >= 2) return raw;
  const id = Number(company?.id) || 0;
  return id > 0 ? `Company ${id}` : "Company";
}

function getActiveCompany() {
  const companyId = getSelectedCompanyId();
  return companies.find((company) => company.id === companyId) || {
    id: 1,
    name: "Routes Payroll",
    logoDataUrl: "",
  };
}

function payrollMonthUrl(month) {
  const companyId = getSelectedCompanyId();
  return `/api/payroll/${month}?companyId=${companyId}`;
}

async function loadCompanies() {
  const previousCompanyId = activeCompanyId;
  const response = await apiRequest("/api/companies");
  const incomingCompanies = Array.isArray(response.companies) ? response.companies : [];
  companies = incomingCompanies
    .map((company) => {
      const id = Number(company?.id);
      if (!Number.isInteger(id) || id <= 0) return null;
      return {
        id,
        name: String(company?.name ?? ""),
        logoDataUrl: String(company?.logoDataUrl || ""),
      };
    })
    .filter(Boolean);

  if (companies.length === 0) {
    companies = [{ id: 1, name: "Routes Payroll", logoDataUrl: "" }];
  }

  if (previousCompanyId && companies.some((company) => company.id === previousCompanyId)) {
    activeCompanyId = previousCompanyId;
  } else {
    activeCompanyId = companies[0].id;
  }

  SELECTORS.companyPicker.innerHTML = companies
    .map((company) => `<option value="${company.id}">${escapeHtml(displayCompanyName(company))}</option>`)
    .join("");
  SELECTORS.companyPicker.value = String(activeCompanyId);
}

async function loadMonthRecords() {
  const month = getSelectedMonth();
  try {
    const response = await apiRequest(payrollMonthUrl(month));
    currentRecords = Array.isArray(response.records) ? response.records : [];
    renderPayrollTable();
    setSaveStatus("All changes saved.");
  } catch (error) {
    currentRecords = [];
    renderPayrollTable();
    showAppMessage(error.message);
  }
}

function createEmptyEmployee(records) {
  const nextId = findNextEmployeeId(records);
  return {
    employeeId: nextId,
    employeeName: "",
    designation: "",
    presentSalary: "0",
    increment: "0",
    oldAdvanceTaken: "0",
    extraAdvanceAdded: "0",
    deductionEntered: "0",
    daysAbsent: "0",
    comment: "",
  };
}

function findNextEmployeeId(records) {
  let maxId = 0;
  for (const item of records) {
    const match = String(item.employeeId || "").match(/(\d+)/);
    if (match) {
      maxId = Math.max(maxId, Number(match[1]));
    }
  }
  return `EMP${String(maxId + 1).padStart(3, "0")}`;
}

function computePayroll(record) {
  const presentSalary = toMoney(record.presentSalary);
  const increment = toMoney(record.increment);
  const grossSalary = presentSalary + increment;

  const oldAdvanceTaken = toMoney(record.oldAdvanceTaken);
  const extraAdvanceAdded = toMoney(record.extraAdvanceAdded);
  const totalAdvance = oldAdvanceTaken + extraAdvanceAdded;

  const deductionEntered = Math.max(0, toMoney(record.deductionEntered));
  const daysAbsentRaw = toMoney(record.daysAbsent);
  const daysAbsent = clamp(daysAbsentRaw, 0, 30);
  const proratedAbsenceDeduction = (daysAbsent / 30) * grossSalary;
  const deductionApplied = Math.min(deductionEntered, totalAdvance);
  const advanceRemained = totalAdvance - deductionApplied;
  const netSalary = grossSalary - deductionApplied - proratedAbsenceDeduction;

  return {
    presentSalary,
    increment,
    grossSalary,
    oldAdvanceTaken,
    extraAdvanceAdded,
    totalAdvance,
    deductionEntered,
    daysAbsent,
    proratedAbsenceDeduction,
    deductionApplied,
    advanceRemained,
    netSalary,
  };
}

function renderPayrollTable() {
  const month = getSelectedMonth();

  if (currentRecords.length === 0) {
    SELECTORS.payrollBody.innerHTML = `
      <tr>
        <td colspan="17" class="empty">
          No employees for ${formatMonth(month)}. Click "Add Employee" to begin.
        </td>
      </tr>
    `;
    updateMetrics([]);
    return;
  }

  SELECTORS.payrollBody.innerHTML = currentRecords
    .map((record, index) => {
      const calc = computePayroll(record);
      return `
        <tr data-index="${index}">
          <td><input data-field="employeeId" value="${escapeHtml(record.employeeId || "")}" /></td>
          <td><input data-field="employeeName" value="${escapeHtml(record.employeeName || "")}" /></td>
          <td><input data-field="designation" value="${escapeHtml(record.designation || "")}" /></td>
          <td><input class="field-salary" data-field="presentSalary" type="number" min="0" step="0.01" value="${toRaw(record.presentSalary)}" /></td>
          <td><input class="field-increment" data-field="increment" type="number" min="0" step="0.01" value="${toRaw(record.increment)}" /></td>
          <td class="read money-gross">${formatCurrency(calc.grossSalary)}</td>
          <td><input class="field-advance" data-field="oldAdvanceTaken" type="number" min="0" step="0.01" value="${toRaw(record.oldAdvanceTaken)}" /></td>
          <td><input class="field-advance" data-field="extraAdvanceAdded" type="number" min="0" step="0.01" value="${toRaw(record.extraAdvanceAdded)}" /></td>
          <td class="read money-advance">${formatCurrency(calc.totalAdvance)}</td>
          <td><input class="field-deduction" data-field="deductionEntered" type="number" min="0" step="0.01" value="${toRaw(record.deductionEntered)}" /></td>
          <td><input class="field-absent" data-field="daysAbsent" type="number" min="0" max="30" step="0.5" value="${toRaw(record.daysAbsent)}" /></td>
          <td class="read money-deduction">${formatCurrency(calc.proratedAbsenceDeduction)}</td>
          <td class="read money-deduction">${formatCurrency(calc.deductionApplied)}</td>
          <td class="read money-advance">${formatCurrency(calc.advanceRemained)}</td>
          <td class="read money-net">${formatCurrency(calc.netSalary)}</td>
          <td><textarea data-field="comment" rows="2">${escapeHtml(record.comment || "")}</textarea></td>
          <td>
            <div class="row-actions">
              <button data-action="payslip" class="mini">Payslip</button>
              <button data-action="delete" class="mini danger">Delete</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");

  updateMetrics(currentRecords.map((record) => computePayroll(record)));
}

function updateMetrics(computedRecords) {
  const totalEmployees = computedRecords.length;
  const gross = computedRecords.reduce((sum, item) => sum + item.grossSalary, 0);
  const net = computedRecords.reduce((sum, item) => sum + item.netSalary, 0);
  const advance = computedRecords.reduce((sum, item) => sum + item.advanceRemained, 0);

  SELECTORS.metricEmployees.textContent = String(totalEmployees);
  SELECTORS.metricGross.textContent = formatCurrency(gross);
  SELECTORS.metricNet.textContent = formatCurrency(net);
  SELECTORS.metricAdvance.textContent = formatCurrency(advance);
}

function scheduleSave(immediate) {
  setSaveStatus("Saving changes...");

  if (pendingSaveTimer) {
    clearTimeout(pendingSaveTimer);
    pendingSaveTimer = null;
  }

  if (immediate) {
    persistRecords();
    return;
  }

  pendingSaveTimer = setTimeout(() => {
    pendingSaveTimer = null;
    persistRecords();
  }, 600);
}

async function flushPendingSave() {
  if (pendingSaveTimer) {
    clearTimeout(pendingSaveTimer);
    pendingSaveTimer = null;
    await persistRecords();
  }
}

async function persistRecords() {
  if (saveInFlight) {
    saveQueued = true;
    return;
  }

  saveInFlight = true;
  const month = getSelectedMonth();

  try {
    await apiRequest(payrollMonthUrl(month), {
      method: "PUT",
      body: { records: currentRecords },
    });
    setSaveStatus("All changes saved.");
  } catch (error) {
    setSaveStatus("Save failed.");
    showAppMessage(error.message);
  } finally {
    saveInFlight = false;
    if (saveQueued) {
      saveQueued = false;
      persistRecords();
    }
  }
}

function openPayslip(record, month) {
  const calc = computePayroll(record);
  const company = getActiveCompany();
  activePayslip = { record, calc, month, company };
  SELECTORS.payslipPreview.innerHTML = renderPayslipCard(record, calc, month, company);
  SELECTORS.payslipDialog.showModal();
}

function renderPayslipCard(record, calc, month, company) {
  return `
    <article class="payslip-card">
      <header class="payslip-brand">
        <div class="payslip-brand-row">
          ${company.logoDataUrl ? `<img src="${company.logoDataUrl}" alt="${escapeHtml(company.name)} logo" class="payslip-logo" />` : ""}
          <div>
            <p class="eyebrow">Routes Payroll Payslip</p>
            <h4>${escapeHtml(company.name || "Company")}</h4>
          </div>
        </div>
        <h3>${escapeHtml(record.employeeName || "Employee Name")}</h3>
        <p>${escapeHtml(record.designation || "Designation")}</p>
      </header>
      <dl>
        <div><dt>Company</dt><dd>${escapeHtml(company.name || "-")}</dd></div>
        <div><dt>Year-Month</dt><dd>${escapeHtml(formatMonth(month))}</dd></div>
        <div><dt>Employee ID</dt><dd>${escapeHtml(record.employeeId || "-")}</dd></div>
        <div><dt>Present Salary</dt><dd>${formatCurrency(calc.presentSalary)}</dd></div>
        <div><dt>Increment</dt><dd>${formatCurrency(calc.increment)}</dd></div>
        <div><dt>Gross Salary</dt><dd>${formatCurrency(calc.grossSalary)}</dd></div>
        <div><dt>Old Advance Taken</dt><dd>${formatCurrency(calc.oldAdvanceTaken)}</dd></div>
        <div><dt>Extra Advance Added</dt><dd>${formatCurrency(calc.extraAdvanceAdded)}</dd></div>
        <div><dt>Total Advance</dt><dd>${formatCurrency(calc.totalAdvance)}</dd></div>
        <div><dt>Deduction Entered</dt><dd>${formatCurrency(calc.deductionEntered)}</dd></div>
        <div><dt>Days Absent</dt><dd>${calc.daysAbsent} / 30</dd></div>
        <div><dt>Prorated Absence Deduction</dt><dd>${formatCurrency(calc.proratedAbsenceDeduction)}</dd></div>
        <div><dt>Deduction Applied</dt><dd>${formatCurrency(calc.deductionApplied)}</dd></div>
        <div><dt>Advance Remained</dt><dd>${formatCurrency(calc.advanceRemained)}</dd></div>
        <div><dt>Salary In Hand</dt><dd class="money-net">${formatCurrency(calc.netSalary)}</dd></div>
        <div><dt>Comment</dt><dd>${escapeHtml(record.comment || "-")}</dd></div>
      </dl>
    </article>
  `;
}

function ensureActivePayslip() {
  if (!activePayslip) {
    showAppMessage("No payslip selected.");
    return false;
  }
  return true;
}

function getPayslipText(data) {
  const { record, calc, month, company } = data;
  return [
    "Routes Payroll Payslip",
    `Company: ${company?.name || "-"}`,
    `Year-Month: ${formatMonth(month)}`,
    `Employee ID: ${record.employeeId || "-"}`,
    `Employee Name: ${record.employeeName || "-"}`,
    `Designation: ${record.designation || "-"}`,
    `Present Salary: ${formatCurrency(calc.presentSalary)}`,
    `Increment: ${formatCurrency(calc.increment)}`,
    `Gross Salary: ${formatCurrency(calc.grossSalary)}`,
    `Old Advance Taken: ${formatCurrency(calc.oldAdvanceTaken)}`,
    `Extra Advance Added: ${formatCurrency(calc.extraAdvanceAdded)}`,
    `Total Advance: ${formatCurrency(calc.totalAdvance)}`,
    `Deduction Entered: ${formatCurrency(calc.deductionEntered)}`,
    `Days Absent: ${calc.daysAbsent} / 30`,
    `Prorated Absence Deduction: ${formatCurrency(calc.proratedAbsenceDeduction)}`,
    `Deduction Applied: ${formatCurrency(calc.deductionApplied)}`,
    `Advance Remained: ${formatCurrency(calc.advanceRemained)}`,
    `Salary In Hand: ${formatCurrency(calc.netSalary)}`,
    `Comment: ${record.comment || "-"}`,
  ].join("\n");
}

function printCurrentPayslip() {
  if (!ensureActivePayslip()) return;
  const html = `
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Payslip - ${escapeHtml(activePayslip.record.employeeName || "")}</title>
        <style>
          body { font-family: "Segoe UI", Tahoma, sans-serif; margin: 24px; color: #111; }
          h1 { margin: 0 0 4px; color: #111; }
          p { margin: 0 0 18px; color: #444; }
          table { width: 100%; border-collapse: collapse; }
          td { border-bottom: 1px solid #ddd; padding: 8px 0; }
          td:last-child { font-weight: 600; text-align: right; }
          .net { color: #0d8a53; font-size: 1.1rem; }
        </style>
      </head>
      <body>
        <h1>Routes Payroll Payslip</h1>
        <p>${escapeHtml(activePayslip.company?.name || "Company")}</p>
        <p>${escapeHtml(formatMonth(activePayslip.month))}</p>
        <table>
          ${buildPrintRows(activePayslip)}
        </table>
      </body>
    </html>
  `;

  const printWindow = window.open("", "_blank", "width=900,height=800");
  if (!printWindow) {
    showAppMessage("Pop-up blocked. Allow pop-ups to print.");
    return;
  }
  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

function buildPrintRows(data) {
  const { record, calc, company } = data;
  const rows = [
    ["Company", escapeHtml(company?.name || "-")],
    ["Employee ID", escapeHtml(record.employeeId || "-")],
    ["Employee Name", escapeHtml(record.employeeName || "-")],
    ["Designation", escapeHtml(record.designation || "-")],
    ["Present Salary", formatCurrency(calc.presentSalary)],
    ["Increment", formatCurrency(calc.increment)],
    ["Gross Salary", formatCurrency(calc.grossSalary)],
    ["Old Advance Taken", formatCurrency(calc.oldAdvanceTaken)],
    ["Extra Advance Added", formatCurrency(calc.extraAdvanceAdded)],
    ["Total Advance", formatCurrency(calc.totalAdvance)],
    ["Deduction Entered", formatCurrency(calc.deductionEntered)],
    ["Days Absent", `${calc.daysAbsent} / 30`],
    ["Prorated Absence Deduction", formatCurrency(calc.proratedAbsenceDeduction)],
    ["Deduction Applied", formatCurrency(calc.deductionApplied)],
    ["Advance Remained", formatCurrency(calc.advanceRemained)],
    ["Salary In Hand", `<span class="net">${formatCurrency(calc.netSalary)}</span>`],
    ["Comment", escapeHtml(record.comment || "-")],
  ];

  return rows.map(([left, right]) => `<tr><td>${left}</td><td>${right}</td></tr>`).join("");
}

function downloadCurrentPayslipText() {
  if (!ensureActivePayslip()) return;
  const text = getPayslipText(activePayslip);
  const name = slugify(activePayslip.record.employeeName || activePayslip.record.employeeId || "employee");
  downloadBlob(new Blob([text], { type: "text/plain;charset=utf-8" }), `payslip-${name}-${activePayslip.month}.txt`);
}

async function shareCurrentPayslipWeb() {
  if (!ensureActivePayslip()) return;
  const text = getPayslipText(activePayslip);
  if (!navigator.share) {
    showAppMessage("Web Share is not supported on this device.");
    return;
  }
  try {
    await navigator.share({ title: `Payslip ${activePayslip.record.employeeName || ""}`, text });
  } catch {
    // Ignore if user cancels sharing.
  }
}

function shareCurrentPayslipWhatsapp() {
  if (!ensureActivePayslip()) return;
  const text = getPayslipText(activePayslip);
  const url = `https://wa.me/?text=${encodeURIComponent(text)}`;
  window.open(url, "_blank", "noopener");
}

function shareCurrentPayslipMessenger() {
  if (!ensureActivePayslip()) return;
  const text = getPayslipText(activePayslip);
  const appUrl = window.location.href.split("#")[0];
  const messengerDeepLink = `fb-messenger://share?link=${encodeURIComponent(appUrl)}`;
  const shareUrl = `https://www.facebook.com/sharer/sharer.php?u=${encodeURIComponent(appUrl)}&quote=${encodeURIComponent(text)}`;
  window.location.href = messengerDeepLink;
  setTimeout(() => {
    window.open(shareUrl, "_blank", "noopener");
  }, 450);
}

async function copyCurrentPayslip() {
  if (!ensureActivePayslip()) return;
  try {
    await navigator.clipboard.writeText(getPayslipText(activePayslip));
    showAppMessage("Payslip copied to clipboard.");
  } catch {
    showAppMessage("Clipboard access failed.");
  }
}

function makeCsv(records, month) {
  const company = getActiveCompany();
  const headers = [
    "Company",
    "Year-Month",
    "Employee ID",
    "Employee Names",
    "Designation",
    "Present Salary (₹)",
    "Increment (₹)",
    "Gross Salary (₹)",
    "Old Advance Taken (₹)",
    "Extra Advance Added (₹)",
    "Total Advance (₹)",
    "Deduction Entered (₹)",
    "Days Absent (out of 30)",
    "Prorated Absence Deduction (₹)",
    "Deduction Applied (Advance only)",
    "Advance Remained (₹)",
    "Salary In Hand (Net Salary ₹)",
    "Comment",
  ];

  const rows = records.map((record) => {
    const calc = computePayroll(record);
    return [
      company.name || "",
      month,
      record.employeeId || "",
      record.employeeName || "",
      record.designation || "",
      calc.presentSalary,
      calc.increment,
      calc.grossSalary,
      calc.oldAdvanceTaken,
      calc.extraAdvanceAdded,
      calc.totalAdvance,
      calc.deductionEntered,
      calc.daysAbsent,
      calc.proratedAbsenceDeduction,
      calc.deductionApplied,
      calc.advanceRemained,
      calc.netSalary,
      record.comment || "",
    ];
  });

  return [headers, ...rows].map((row) => row.map(escapeCsv).join(",")).join("\n");
}

function makeExcelSpreadsheet(records, month) {
  const company = getActiveCompany();
  const computed = records.map((record) => ({ record, calc: computePayroll(record) }));
  const totals = computed.reduce((acc, item) => {
    acc.gross += item.calc.grossSalary;
    acc.net += item.calc.netSalary;
    acc.advance += item.calc.advanceRemained;
    return acc;
  }, { gross: 0, net: 0, advance: 0 });

  const title = `${company.name || "Company"} - Monthly Payroll Report`;
  const monthLabel = formatMonth(month);
  const generatedAt = new Date().toLocaleString("en-IN");

  const columns = [
    90, 150, 130, 110, 95, 110, 120, 120, 110, 120, 100, 140, 140, 120, 140, 180, 120,
  ];

  const headerCells = [
    ["Employee ID", "sHeader"],
    ["Employee Name", "sHeader"],
    ["Designation", "sHeader"],
    ["Present Salary (INR)", "sHeaderSalary"],
    ["Increment (INR)", "sHeaderIncrement"],
    ["Gross Salary (INR)", "sHeaderGross"],
    ["Old Advance Taken (INR)", "sHeaderAdvance"],
    ["Extra Advance Added (INR)", "sHeaderAdvance"],
    ["Total Advance (INR)", "sHeaderAdvance"],
    ["Deduction Entered (INR)", "sHeaderDeduction"],
    ["Days Absent", "sHeaderAbsent"],
    ["Prorated Absence Deduction (INR)", "sHeaderDeduction"],
    ["Deduction Applied (INR)", "sHeaderDeduction"],
    ["Advance Remained (INR)", "sHeaderAdvance"],
    ["Salary In Hand (INR)", "sHeaderNet"],
    ["Comment", "sHeader"],
    ["Updated Month", "sHeader"],
  ];

  const headerRow = `<Row ss:Height="34">${headerCells.map(([text, style]) => xmlCellString(text, style)).join("")}</Row>`;

  const dataRows = computed.map((item, idx) => {
    const rowStyle = idx % 2 === 0 ? "sCell" : "sCellAlt";
    return `<Row>
      ${xmlCellString(item.record.employeeId || "", rowStyle)}
      ${xmlCellString(item.record.employeeName || "", rowStyle)}
      ${xmlCellString(item.record.designation || "", rowStyle)}
      ${xmlCellNumber(item.calc.presentSalary, "sMoneySalary")}
      ${xmlCellNumber(item.calc.increment, "sMoneyIncrement")}
      ${xmlCellNumber(item.calc.grossSalary, "sMoneyGross")}
      ${xmlCellNumber(item.calc.oldAdvanceTaken, "sMoneyAdvance")}
      ${xmlCellNumber(item.calc.extraAdvanceAdded, "sMoneyAdvance")}
      ${xmlCellNumber(item.calc.totalAdvance, "sMoneyAdvanceStrong")}
      ${xmlCellNumber(item.calc.deductionEntered, "sMoneyDeduction")}
      ${xmlCellNumber(item.calc.daysAbsent, "sNum")}
      ${xmlCellNumber(item.calc.proratedAbsenceDeduction, "sMoneyDeduction")}
      ${xmlCellNumber(item.calc.deductionApplied, "sMoneyDeductionStrong")}
      ${xmlCellNumber(item.calc.advanceRemained, "sMoneyAdvanceStrong")}
      ${xmlCellNumber(item.calc.netSalary, "sMoneyNet")}
      ${xmlCellString(item.record.comment || "", rowStyle)}
      ${xmlCellString(month, rowStyle)}
    </Row>`;
  }).join("");

  const worksheetRows = `
    <Row ss:Height="30">
      <Cell ss:StyleID="sTitle" ss:MergeAcross="16"><Data ss:Type="String">${escapeXml(title)}</Data></Cell>
    </Row>
    <Row>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Month</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="String">${escapeXml(monthLabel)}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Company</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="String">${escapeXml(company.name || "")}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Generated</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="String">${escapeXml(generatedAt)}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Employees</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="Number">${computed.length}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Gross</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="Number">${toFixedNumber(totals.gross)}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Net</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="Number">${toFixedNumber(totals.net)}</Data></Cell>
      <Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">Advance Remained</Data></Cell>
      <Cell ss:StyleID="sMetaValue"><Data ss:Type="Number">${toFixedNumber(totals.advance)}</Data></Cell>
    </Row>
    <Row ss:Height="8"></Row>
    ${headerRow}
    ${dataRows || `<Row><Cell ss:StyleID="sEmpty" ss:MergeAcross="16"><Data ss:Type="String">No records for ${escapeXml(monthLabel)}</Data></Cell></Row>`}
  `;

  return `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Author>Routes Payroll</Author>
    <Created>${new Date().toISOString()}</Created>
  </DocumentProperties>
  <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
    <WindowHeight>12000</WindowHeight>
    <WindowWidth>24000</WindowWidth>
    <ProtectStructure>False</ProtectStructure>
    <ProtectWindows>False</ProtectWindows>
  </ExcelWorkbook>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Alignment ss:Vertical="Center"/>
      <Borders/>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#0F1728"/>
      <Interior/>
      <NumberFormat/>
      <Protection/>
    </Style>
    <Style ss:ID="sTitle">
      <Font ss:FontName="Calibri" ss:Size="16" ss:Bold="1" ss:Color="#0F294D"/>
      <Interior ss:Color="#E9F2FF" ss:Pattern="Solid"/>
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
    </Style>
    <Style ss:ID="sMetaLabel">
      <Font ss:FontName="Calibri" ss:Size="10" ss:Bold="1" ss:Color="#2C3E57"/>
      <Interior ss:Color="#F3F7FC" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D5E0ED"/>
      </Borders>
    </Style>
    <Style ss:ID="sMetaValue">
      <Font ss:FontName="Calibri" ss:Size="10" ss:Color="#10243D"/>
      <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
      <NumberFormat ss:Format="#,##0.00"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E1E8F1"/>
      </Borders>
    </Style>
    <Style ss:ID="sHeader">
      <Font ss:Bold="1" ss:Color="#10243D"/>
      <Interior ss:Color="#EEF4FB" ss:Pattern="Solid"/>
      <Alignment ss:WrapText="1"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/>
      </Borders>
    </Style>
    <Style ss:ID="sHeaderSalary"><Font ss:Bold="1" ss:Color="#0F67C6"/><Interior ss:Color="#EAF3FF" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderIncrement"><Font ss:Bold="1" ss:Color="#6246C7"/><Interior ss:Color="#EFEAFF" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderGross"><Font ss:Bold="1" ss:Color="#157347"/><Interior ss:Color="#E9F8F0" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderAdvance"><Font ss:Bold="1" ss:Color="#8B6B10"/><Interior ss:Color="#FBF5E8" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderDeduction"><Font ss:Bold="1" ss:Color="#B91C2D"/><Interior ss:Color="#FDEEEF" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderAbsent"><Font ss:Bold="1" ss:Color="#9A4F07"/><Interior ss:Color="#FFF4E8" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sHeaderNet"><Font ss:Bold="1" ss:Color="#0F7A55"/><Interior ss:Color="#EAF8F2" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C9D7E8"/></Borders></Style>
    <Style ss:ID="sCell"><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sCellAlt"><Interior ss:Color="#FAFCFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sNum"><NumberFormat ss:Format="0.00"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneySalary"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#0F67C6"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyIncrement"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#6246C7"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyGross"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#157347"/><Interior ss:Color="#F1FBF5" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyAdvance"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#8B6B10"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyAdvanceStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8B6B10"/><Interior ss:Color="#FFF7E8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyDeduction"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#B91C2D"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyDeductionStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#B91C2D"/><Interior ss:Color="#FFF1F3" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sMoneyNet"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F7A55"/><Interior ss:Color="#ECFAF3" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E4EAF2"/></Borders></Style>
    <Style ss:ID="sEmpty">
      <Font ss:Bold="1" ss:Color="#526680"/>
      <Interior ss:Color="#F7FAFF" ss:Pattern="Solid"/>
      <Alignment ss:Horizontal="Center"/>
    </Style>
  </Styles>
  <Worksheet ss:Name="${escapeXml(`Payroll ${month}`)}">
    <Table>
      ${columns.map((width) => `<Column ss:AutoFitWidth="0" ss:Width="${width}"/>`).join("")}
      ${worksheetRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>4</SplitHorizontal>
      <TopRowBottomPane>4</TopRowBottomPane>
      <ActivePane>2</ActivePane>
      <Panes>
        <Pane>
          <Number>3</Number>
        </Pane>
      </Panes>
      <ProtectObjects>False</ProtectObjects>
      <ProtectScenarios>False</ProtectScenarios>
    </WorksheetOptions>
  </Worksheet>
</Workbook>`;
}

function xmlCellString(value, styleId) {
  const style = styleId ? ` ss:StyleID="${styleId}"` : "";
  return `<Cell${style}><Data ss:Type="String">${escapeXml(value)}</Data></Cell>`;
}

function xmlCellNumber(value, styleId) {
  const style = styleId ? ` ss:StyleID="${styleId}"` : "";
  return `<Cell${style}><Data ss:Type="Number">${toFixedNumber(value)}</Data></Cell>`;
}

function toFixedNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num.toFixed(2) : "0.00";
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Failed to read logo file."));
    reader.readAsDataURL(file);
  });
}

function escapeCsv(value) {
  const text = String(value ?? "");
  if (text.includes(",") || text.includes('"') || text.includes("\n")) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

async function apiRequest(url, options = {}, withAuth = true) {
  const headers = { ...(options.headers || {}) };
  if (options.body !== undefined && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }
  if (withAuth && authToken) {
    headers.Authorization = `Bearer ${authToken}`;
  }

  const response = await fetch(url, {
    method: options.method || "GET",
    headers,
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
  });

  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json")
    ? await response.json()
    : null;

  if (!response.ok) {
    const message = payload?.error || `Request failed (${response.status}).`;
    if (response.status === 401 && withAuth) {
      setToken("");
      updateAuthView();
    }
    throw new Error(message);
  }

  return payload || {};
}

function setToken(token) {
  authToken = token || "";
  if (authToken) {
    localStorage.setItem(STORAGE_KEYS.authToken, authToken);
  } else {
    localStorage.removeItem(STORAGE_KEYS.authToken);
  }
}

function setSaveStatus(message) {
  applyStatusTone(SELECTORS.saveStatus, message);
  SELECTORS.saveStatus.textContent = message;
}

function toMoney(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function toRaw(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatCurrency(number) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2,
  }).format(number || 0);
}

function formatMonth(isoMonth) {
  if (!isoMonth || !isoMonth.includes("-")) return isoMonth || "";
  const [year, month] = isoMonth.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return date.toLocaleDateString("en-IN", { month: "long", year: "numeric" });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function slugify(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function showAuthMessage(message) {
  applyStatusTone(SELECTORS.authMessage, message);
  SELECTORS.authMessage.textContent = message;
}

function showAppMessage(message) {
  applyStatusTone(SELECTORS.appMessage, message);
  SELECTORS.appMessage.textContent = message;
  if (message) {
    setTimeout(() => {
      if (SELECTORS.appMessage.textContent === message) {
        SELECTORS.appMessage.textContent = "";
      }
    }, 3200);
  }
}

function applyStatusTone(element, message) {
  element.classList.remove("status-info", "status-success", "status-error");
  if (!message) return;

  const text = String(message).toLowerCase();
  if (text.includes("failed") || text.includes("error") || text.includes("invalid")) {
    element.classList.add("status-error");
    return;
  }
  if (text.includes("saved") || text.includes("restored") || text.includes("exported") || text.includes("downloaded") || text.includes("added") || text.includes("removed") || text.includes("copied")) {
    element.classList.add("status-success");
    return;
  }
  element.classList.add("status-info");
}

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js").catch(() => {
      // Non-critical if registration fails.
    });
  }
}

function wireInstallPrompt() {
  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    pendingInstallPrompt = event;
    SELECTORS.installBtn.classList.remove("hidden");
  });

  SELECTORS.installBtn.addEventListener("click", async () => {
    if (!pendingInstallPrompt) return;
    pendingInstallPrompt.prompt();
    await pendingInstallPrompt.userChoice;
    pendingInstallPrompt = null;
    SELECTORS.installBtn.classList.add("hidden");
  });
}
