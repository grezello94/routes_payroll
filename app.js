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
  monthPicker: document.getElementById("monthPicker"),
  payrollBody: document.getElementById("payrollBody"),
  addEmployeeBtn: document.getElementById("addEmployeeBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  backupBtn: document.getElementById("backupBtn"),
  restoreInput: document.getElementById("restoreInput"),
  logoutBtn: document.getElementById("logoutBtn"),
  metricEmployees: document.getElementById("metricEmployees"),
  metricGross: document.getElementById("metricGross"),
  metricNet: document.getElementById("metricNet"),
  metricAdvance: document.getElementById("metricAdvance"),
  installBtn: document.getElementById("installBtn"),
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
    const username = document.getElementById("setupUsername").value.trim();
    const password = document.getElementById("setupPassword").value;
    const confirm = document.getElementById("setupConfirm").value;

    if (password !== confirm) {
      showAuthMessage("Passwords do not match.");
      return;
    }

    try {
      const response = await apiRequest("/api/auth/register", {
        method: "POST",
        body: { username, password },
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
      await loadMonthRecords();
      showAppMessage("Database backup restored.");
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

async function loadMonthRecords() {
  const month = getSelectedMonth();
  try {
    const response = await apiRequest(`/api/payroll/${month}`);
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
    await apiRequest(`/api/payroll/${month}`, {
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
  activePayslip = { record, calc, month };
  SELECTORS.payslipPreview.innerHTML = renderPayslipCard(record, calc, month);
  SELECTORS.payslipDialog.showModal();
}

function renderPayslipCard(record, calc, month) {
  return `
    <article class="payslip-card">
      <header>
        <p class="eyebrow">Routes Payroll Payslip</p>
        <h3>${escapeHtml(record.employeeName || "Employee Name")}</h3>
        <p>${escapeHtml(record.designation || "Designation")}</p>
      </header>
      <dl>
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
  const { record, calc, month } = data;
  return [
    "Routes Payroll Payslip",
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
  const { record, calc } = data;
  const rows = [
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
  const headers = [
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

function showAuthMessage(message) {
  SELECTORS.authMessage.textContent = message;
}

function showAppMessage(message) {
  SELECTORS.appMessage.textContent = message;
  if (message) {
    setTimeout(() => {
      if (SELECTORS.appMessage.textContent === message) {
        SELECTORS.appMessage.textContent = "";
      }
    }, 3200);
  }
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
