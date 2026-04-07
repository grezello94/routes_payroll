const STORAGE_KEYS = {
  authToken: "routes_payroll_auth_token_v1",
  payrollMonthMode: "routes_payroll_month_mode_v1",
};

const SELECTORS = {
  authView: document.getElementById("authView"),
  appView: document.getElementById("appView"),
  setupForm: document.getElementById("setupForm"),
  loginForm: document.getElementById("loginForm"),
  showRegisterBtn: document.getElementById("showRegisterBtn"),
  showLoginBtn: document.getElementById("showLoginBtn"),
  setupEmail: document.getElementById("setupEmail"),
  authMessage: document.getElementById("authMessage"),
  recoverForm: document.getElementById("recoverForm"),
  recoverUsernameCompanyName: document.getElementById("recoverUsernameCompanyName"),
  recoverUsernameEmail: document.getElementById("recoverUsernameEmail"),
  recoverIdentifier: document.getElementById("recoverIdentifier"),
  recoverUsernameBtn: document.getElementById("recoverUsernameBtn"),
  recoverBtn: document.getElementById("recoverBtn"),
  appMessage: document.getElementById("appMessage"),
  saveStatus: document.getElementById("saveStatus"),
  employeesSection: document.getElementById("employeesSection"),
  employeeForm: document.getElementById("employeeForm"),
  employeeFormId: document.getElementById("employeeFormId"),
  employeeNameInput: document.getElementById("employeeNameInput"),
  employeeJoiningDateInput: document.getElementById("employeeJoiningDateInput"),
  employeeBirthDateInput: document.getElementById("employeeBirthDateInput"),
  employeeSalaryInput: document.getElementById("employeeSalaryInput"),
  employeeOpeningAdvanceInput: document.getElementById("employeeOpeningAdvanceInput"),
  employeeDesignationInput: document.getElementById("employeeDesignationInput"),
  employeeMobileInput: document.getElementById("employeeMobileInput"),
  employeeStatusFields: document.getElementById("employeeStatusFields"),
  employeeStatusInput: document.getElementById("employeeStatusInput"),
  employeeLeaveFromInput: document.getElementById("employeeLeaveFromInput"),
  employeeResumeOnInput: document.getElementById("employeeResumeOnInput"),
  employeeTerminatedOnInput: document.getElementById("employeeTerminatedOnInput"),
  employeeSaveBtn: document.getElementById("employeeSaveBtn"),
  employeeCancelBtn: document.getElementById("employeeCancelBtn"),
  employeeMessage: document.getElementById("employeeMessage"),
  employeesBody: document.getElementById("employeesBody"),
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
  settingsSection: document.getElementById("settingsSection"),
  reportsSection: document.getElementById("reportsSection"),
  settingsUsernameValue: document.getElementById("settingsUsernameValue"),
  settingsEmailValue: document.getElementById("settingsEmailValue"),
  settingsEmailStatusValue: document.getElementById("settingsEmailStatusValue"),
  verifyEmailBtn: document.getElementById("verifyEmailBtn"),
  settingsCompanyNameInput: document.getElementById("settingsCompanyNameInput"),
  saveCompanyNameBtn: document.getElementById("saveCompanyNameBtn"),
  settingsPayrollMonthMode: document.getElementById("settingsPayrollMonthMode"),
  designationPresetInput: document.getElementById("designationPresetInput"),
  addDesignationBtn: document.getElementById("addDesignationBtn"),
  designationList: document.getElementById("designationList"),
  settingsLogoInput: document.getElementById("settingsLogoInput"),
  settingsLogoPreview: document.getElementById("settingsLogoPreview"),
  saveLogoBtn: document.getElementById("saveLogoBtn"),
  legacyImportInput: document.getElementById("legacyImportInput"),
  importLegacyBtn: document.getElementById("importLegacyBtn"),
  importLegacySummary: document.getElementById("importLegacySummary"),
  changePasswordForm: document.getElementById("changePasswordForm"),
  newPasswordInput: document.getElementById("newPasswordInput"),
  confirmPasswordInput: document.getElementById("confirmPasswordInput"),
  changePasswordBtn: document.getElementById("changePasswordBtn"),
  settingsMessage: document.getElementById("settingsMessage"),
  leaveResumeReportBody: document.getElementById("leaveResumeReportBody"),
  reportsMessage: document.getElementById("reportsMessage"),
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
let employeeMaster = [];
let activePayrollEmployeeId = "";
let companies = [];
let activeCompanyId = null;
let pendingSaveTimer = null;
let saveInFlight = false;
let saveQueued = false;
let needsSetupFlow = false;
let designationPresets = [];
let pendingSettingsLogoDataUrl = "";
let currentUser = null;
let serverReconnectTimer = null;
let serverReconnectInFlight = false;

init();

async function init() {
  wireAuth();
  wirePasswordToggles();
  wireEmployeeManagement();
  wireAppActions();
  wireSettingsActions();
  setDefaultMonth();
  registerServiceWorker();
  wireInstallPrompt();
  await updateAuthView();
}

function wireAuth() {
  SELECTORS.showRegisterBtn?.addEventListener("click", () => {
    setAuthMode("register");
  });

  SELECTORS.showLoginBtn?.addEventListener("click", () => {
    setAuthMode("login");
  });

  SELECTORS.setupForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitBtn = SELECTORS.setupForm.querySelector('button[type="submit"]');
    const companyName = document.getElementById("setupCompanyName").value.trim();
    const email = SELECTORS.setupEmail.value.trim().toLowerCase();
    const username = document.getElementById("setupUsername").value.trim();
    const password = document.getElementById("setupPassword").value;
    const confirm = document.getElementById("setupConfirm").value;

    if (companyName.length < 2) {
      showAuthMessage("Company name is too short.");
      return;
    }
    if (!isValidEmail(email)) {
      showAuthMessage("Enter a valid recovery email.");
      return;
    }

    if (password !== confirm) {
      showAuthMessage("Passwords do not match.");
      return;
    }

    try {
      if (submitBtn) submitBtn.disabled = true;
      showAuthMessage("Creating account...");
      const response = await apiRequest("/api/auth/register", {
        method: "POST",
        body: { companyName, email, username, password },
      }, false);

      setToken(response.token);
      currentUser = response.user || null;
      hydrateSettingsAccount();
      needsSetupFlow = false;
      SELECTORS.setupForm.reset();
      showAuthMessage("");
      await switchToApp();
      showAppMessage(response.message || "Account created.");
    } catch (error) {
      showAuthMessage(error.message);
    } finally {
      if (submitBtn) submitBtn.disabled = false;
    }
  });

  SELECTORS.loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitBtn = SELECTORS.loginForm.querySelector('button[type="submit"]');
    if (needsSetupFlow) {
      setAuthMode("register");
      showAuthMessage("Please register company and admin first, then use Sign In.");
      return;
    }
    const username = document.getElementById("loginUsername").value.trim();
    const password = document.getElementById("loginPassword").value;

    try {
      if (submitBtn) submitBtn.disabled = true;
      showAuthMessage("Signing in...");
      const response = await apiRequest("/api/auth/login", {
        method: "POST",
        body: { username, password },
      }, false);

      setToken(response.token);
      currentUser = response.user || null;
      hydrateSettingsAccount();
      SELECTORS.loginForm.reset();
      showAuthMessage("");
      await switchToApp();
    } catch (error) {
      showAuthMessage(error.message);
    } finally {
      if (submitBtn) submitBtn.disabled = false;
    }
  });

  SELECTORS.recoverUsernameBtn?.addEventListener("click", async () => {
    const companyName = SELECTORS.recoverUsernameCompanyName?.value.trim() || "";
    const email = SELECTORS.recoverUsernameEmail?.value.trim().toLowerCase() || "";

    if (companyName.length < 2) {
      showAuthMessage("Enter your registered company name.");
      return;
    }
    if (!isValidEmail(email)) {
      showAuthMessage("Enter your registered email.");
      return;
    }

    try {
      await apiRequest("/api/auth/recover-email", {
        method: "POST",
        body: { companyName, email },
      }, false);
      showAuthMessage("If account exists, username reminder has been sent.");
    } catch (error) {
      showAuthMessage(error.message);
    }
  });

  SELECTORS.recoverBtn?.addEventListener("click", async () => {
    const identifier = SELECTORS.recoverIdentifier?.value.trim() || "";
    if (!identifier) {
      showAuthMessage("Enter username or email.");
      return;
    }

    try {
      await apiRequest("/api/auth/request-password-reset", {
        method: "POST",
        body: { identifier },
      }, false);
      showAuthMessage("If account exists, password reset link has been sent to registered email.");
    } catch (error) {
      showAuthMessage(error.message);
    }
  });
}

function wirePasswordToggles() {
  const toggleButtons = Array.from(document.querySelectorAll(".password-toggle"));
  for (const button of toggleButtons) {
    button.addEventListener("click", () => {
      const targetId = button.getAttribute("data-password-target") || "";
      const input = document.getElementById(targetId);
      if (!input) return;

      const showing = input.type === "text";
      input.type = showing ? "password" : "text";
      button.setAttribute("aria-pressed", showing ? "false" : "true");
      button.setAttribute("aria-label", showing ? "Show password" : "Hide password");
      button.textContent = "👁";
    });
  }
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || ""));
}

function wireEmployeeManagement() {
  SELECTORS.employeeStatusInput?.addEventListener("change", () => {
    syncEmployeeStatusFields();
  });

  SELECTORS.employeeForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitBtn = SELECTORS.employeeSaveBtn;
    const editingId = Number(SELECTORS.employeeForm.dataset.editingId || 0);
    const isEditMode = editingId > 0;
    const statusValue = isEditMode ? (SELECTORS.employeeStatusInput.value || "working") : "working";
    const payload = {
      companyId: getSelectedCompanyId(),
      employeeId: SELECTORS.employeeFormId.value.trim() || undefined,
      employeeName: SELECTORS.employeeNameInput.value.trim(),
      joiningDate: SELECTORS.employeeJoiningDateInput.value || "",
      birthDate: SELECTORS.employeeBirthDateInput.value || "",
      baseSalary: Number(SELECTORS.employeeSalaryInput.value || 0),
      openingAdvance: Math.max(0, Number(SELECTORS.employeeOpeningAdvanceInput.value || 0)),
      designation: SELECTORS.employeeDesignationInput.value.trim(),
      mobileNumber: SELECTORS.employeeMobileInput.value.trim(),
      status: statusValue,
      leaveFrom: isEditMode ? (SELECTORS.employeeLeaveFromInput.value || "") : "",
      leaveTo: isEditMode ? (SELECTORS.employeeResumeOnInput.value || "") : "",
      terminatedOn: isEditMode ? (SELECTORS.employeeTerminatedOnInput.value || "") : "",
      notes: "",
    };

    if (payload.employeeName.length < 2) {
      setEmployeeMessage("Employee name is too short.");
      return;
    }
    if (!payload.joiningDate) {
      setEmployeeMessage("Joining date is required.");
      return;
    }
    if (payload.openingAdvance < 0) {
      setEmployeeMessage("Old advance cannot be negative.");
      return;
    }
    if (payload.designation.length < 2) {
      setEmployeeMessage("Select designation from Settings presets.");
      return;
    }
    if (isEditMode && payload.status === "leave" && !payload.leaveFrom) {
      setEmployeeMessage("For leave status, set the leave from date.");
      return;
    }
    if (isEditMode && payload.status === "resumed" && !payload.leaveTo) {
      setEmployeeMessage("For resumed work status, select the resumed on date.");
      return;
    }
    if (isEditMode && payload.status === "terminated" && !payload.terminatedOn) {
      setEmployeeMessage("For terminated status, terminated date is required.");
      return;
    }

    try {
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = isEditMode ? "Updating..." : "Saving...";
      }
      if (editingId > 0) {
        await apiRequest(`/api/employees/${editingId}`, {
          method: "PUT",
          body: payload,
        });
        showAppMessage("Employee updated.");
      } else {
        payload.employeeId = nextEmployeeCode(employeeMaster);
        await apiRequest("/api/employees", {
          method: "POST",
          body: payload,
        });
        showAppMessage("Employee added.");
      }

      await loadEmployees();
      await loadMonthRecords();
      resetEmployeeForm();
      setEmployeeMessage("");
    } catch (error) {
      setEmployeeMessage(employeeApiErrorMessage(error));
    } finally {
      if (submitBtn) {
        submitBtn.disabled = false;
        const stillEditing = Number(SELECTORS.employeeForm?.dataset.editingId || 0) > 0;
        submitBtn.textContent = stillEditing ? "Update Employee" : "Save Employee";
      }
    }
  });

  SELECTORS.employeeCancelBtn?.addEventListener("click", () => {
    resetEmployeeForm();
    setEmployeeMessage("");
  });

  SELECTORS.employeesBody?.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button");
    if (!button) return;
    const row = button.closest("tr");
    if (!row) return;
    const id = Number(row.dataset.id);
    if (!id) return;

    if (button.dataset.action === "edit") {
      const employee = employeeMaster.find((item) => Number(item.id) === id);
      if (!employee) return;
      fillEmployeeForm(employee);
      SELECTORS.employeesSection?.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }

    if (button.dataset.action === "delete") {
      if (!window.confirm("Delete this employee from management list?")) return;
      try {
        await apiRequest(`/api/employees/${id}?companyId=${getSelectedCompanyId()}`, {
          method: "DELETE",
        });
        await loadEmployees();
        await loadMonthRecords();
        showAppMessage("Employee deleted.");
      } catch (error) {
        setEmployeeMessage(employeeApiErrorMessage(error));
      }
    }
  });

  resetEmployeeForm();
}

function syncEmployeeStatusFields() {
  const status = SELECTORS.employeeStatusInput?.value || "working";
  const leaveMode = status === "leave";
  const resumedMode = status === "resumed";
  const terminatedMode = status === "terminated";

  SELECTORS.employeeLeaveFromInput.disabled = !leaveMode;
  SELECTORS.employeeResumeOnInput.disabled = !(leaveMode || resumedMode);
  SELECTORS.employeeTerminatedOnInput.disabled = !terminatedMode;

  if (!leaveMode) {
    SELECTORS.employeeLeaveFromInput.value = "";
  }
  if (!(leaveMode || resumedMode)) {
    SELECTORS.employeeResumeOnInput.value = "";
  }
  if (!terminatedMode) {
    SELECTORS.employeeTerminatedOnInput.value = "";
  }
}

function resetEmployeeForm() {
  SELECTORS.employeeForm?.reset();
  SELECTORS.employeeForm.dataset.editingId = "";
  SELECTORS.employeeFormId.value = "";
  SELECTORS.employeeSaveBtn.textContent = "Save Employee";
  SELECTORS.employeeCancelBtn.classList.add("hidden");
  SELECTORS.employeeStatusFields?.classList.add("hidden");
  SELECTORS.employeeStatusInput.value = "working";
  if (SELECTORS.employeeDesignationInput) {
    SELECTORS.employeeDesignationInput.value = "";
  }
  syncEmployeeStatusFields();
}

function fillEmployeeForm(employee) {
  SELECTORS.employeeForm.dataset.editingId = String(employee.id);
  SELECTORS.employeeFormId.value = employee.employeeId || "";
  SELECTORS.employeeNameInput.value = employee.employeeName || "";
  SELECTORS.employeeJoiningDateInput.value = employee.joiningDate || "";
  SELECTORS.employeeBirthDateInput.value = employee.birthDate || "";
  SELECTORS.employeeSalaryInput.value = Number(employee.baseSalary || 0);
  SELECTORS.employeeOpeningAdvanceInput.value = Number(employee.openingAdvance || 0);
  const designationValue = String(employee.designation || "");
  if (designationValue) {
    const exists = Array.from(SELECTORS.employeeDesignationInput.options || [])
      .some((option) => option.value === designationValue);
    if (!exists) {
      const option = document.createElement("option");
      option.value = designationValue;
      option.textContent = designationValue;
      SELECTORS.employeeDesignationInput.append(option);
    }
    SELECTORS.employeeDesignationInput.disabled = false;
    SELECTORS.employeeDesignationInput.value = designationValue;
  } else {
    SELECTORS.employeeDesignationInput.value = "";
  }
  SELECTORS.employeeMobileInput.value = employee.mobileNumber || "";
  SELECTORS.employeeStatusInput.value = employee.status || "working";
  SELECTORS.employeeLeaveFromInput.value = employee.leaveFrom || "";
  SELECTORS.employeeResumeOnInput.value = employee.leaveTo || "";
  SELECTORS.employeeTerminatedOnInput.value = employee.terminatedOn || "";
  SELECTORS.employeeSaveBtn.textContent = "Update Employee";
  SELECTORS.employeeCancelBtn.classList.remove("hidden");
  SELECTORS.employeeStatusFields?.classList.remove("hidden");
  syncEmployeeStatusFields();
}

function nextEmployeeCode(employees) {
  let maxId = 0;
  for (const item of employees || []) {
    const match = String(item.employeeId || "").match(/(\d+)/);
    if (match) maxId = Math.max(maxId, Number(match[1]));
  }
  return `EMP${String(maxId + 1).padStart(3, "0")}`;
}

async function loadEmployees() {
  try {
    const companyId = getSelectedCompanyId();
    const response = await apiRequest(`/api/employees?companyId=${companyId}`);
    employeeMaster = Array.isArray(response.employees) ? response.employees : [];
    // Fetch all payroll records for pending advances
    try {
      const payrollResp = await apiRequest('/api/payroll/all');
      window.allPayrollRecords = Array.isArray(payrollResp.entries) ? payrollResp.entries : [];
    } catch (e) {
      window.allPayrollRecords = [];
    }
    renderEmployeeTable();
    renderLeaveResumeReport();
    renderDesignationSuggestions();
  } catch (error) {
    employeeMaster = [];
    window.allPayrollRecords = [];
    renderEmployeeTable();
    renderLeaveResumeReport();
    renderDesignationSuggestions();
    setEmployeeMessage(employeeApiErrorMessage(error));
  }
}

function renderEmployeeTable() {
  if (!SELECTORS.employeesBody) return;
  if (employeeMaster.length === 0) {
    SELECTORS.employeesBody.innerHTML = `
      <tr><td colspan="8" class="empty">No employees yet. Add employee details above.</td></tr>
    `;
    return;
  }

  // Advance Remained is the carried balance left after the latest month's deduction is applied.
  const payrollByEmployee = {};
  const companyId = getSelectedCompanyId && getSelectedCompanyId();
  if (window.allPayrollRecords) {
    for (const rec of window.allPayrollRecords) {
      if (!rec.employeeId) continue;
      if (companyId && Number(rec.companyId || rec.company_id || 1) !== Number(companyId)) continue;
      if (!payrollByEmployee[rec.employeeId] || (rec.month > payrollByEmployee[rec.employeeId].month)) {
        payrollByEmployee[rec.employeeId] = rec;
      }
    }
  }

  SELECTORS.employeesBody.innerHTML = employeeMaster
    .map((employee) => {
      const statusLabel = employee.status === "leave"
        ? "On Leave"
        : employee.status === "terminated"
          ? "Terminated"
          : "Working";
      const statusClass = employee.status === "leave"
        ? "status-pill leave"
        : employee.status === "terminated"
          ? "status-pill terminated"
          : "status-pill working";
      let advanceRemainedDisplay = Number(employee.openingAdvance || 0);
      const payroll = payrollByEmployee[employee.employeeId];
      if (payroll) {
        if (payroll.advanceRemained != null && !isNaN(Number(payroll.advanceRemained))) {
          advanceRemainedDisplay = Number(payroll.advanceRemained);
        }
        if (advanceRemainedDisplay <= 0) {
          advanceRemainedDisplay = 0;
        } else {
          advanceRemainedDisplay = Math.max(0, advanceRemainedDisplay);
        }
      }

      return `
        <tr data-id="${employee.id}">
          <td>${escapeHtml(employee.employeeId || "-")}</td>
          <td>
            <div class="emp-list-name">${escapeHtml(employee.employeeName || "-")}</div>
            <div class="emp-list-sub">${escapeHtml(employee.designation || "-")}</div>
          </td>
          <td>${formatCurrency(Number(employee.baseSalary || 0))}</td>
          <td>${formatCurrency(advanceRemainedDisplay)}</td>
          <td><span class="${statusClass}">${statusLabel}</span></td>
          <td>${escapeHtml(employee.leaveFrom || "-")}</td>
          <td>${escapeHtml(employee.leaveTo || "-")}</td>
          <td>
            <div class="row-actions">
              <button type="button" class="mini ghost" data-action="edit">Edit</button>
              <button type="button" class="mini danger" data-action="delete">Delete</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");
}

function renderLeaveResumeReport() {
  if (!SELECTORS.leaveResumeReportBody) return;

  const rows = employeeMaster
    .filter((employee) => employee.leaveFrom || employee.leaveTo || employee.terminatedOn)
    .sort((a, b) => String(a.employeeId || "").localeCompare(String(b.employeeId || "")));

  if (rows.length === 0) {
    SELECTORS.leaveResumeReportBody.innerHTML = `
      <tr><td colspan="6" class="empty">No leave, resume, or termination details recorded yet.</td></tr>
    `;
    if (SELECTORS.reportsMessage) {
      SELECTORS.reportsMessage.textContent = "";
    }
    return;
  }

  SELECTORS.leaveResumeReportBody.innerHTML = rows
    .map((employee) => {
      const status = String(employee.status || "working").toLowerCase();
      const statusLabel = status === "leave"
        ? "On Leave"
        : status === "terminated"
          ? "Terminated"
          : "Working";

      return `
        <tr>
          <td>${escapeHtml(employee.employeeId || "-")}</td>
          <td>
            <div class="emp-list-name">${escapeHtml(employee.employeeName || "-")}</div>
            <div class="emp-list-sub">${escapeHtml(employee.designation || "-")}</div>
          </td>
          <td>${escapeHtml(statusLabel)}</td>
          <td>${escapeHtml(employee.leaveFrom || "-")}</td>
          <td>${escapeHtml(employee.leaveTo || "-")}</td>
          <td>${escapeHtml(employee.terminatedOn || "-")}</td>
        </tr>
      `;
    })
    .join("");

  if (SELECTORS.reportsMessage) {
    SELECTORS.reportsMessage.textContent = `${rows.length} employee record(s) with leave, resume, or termination details.`;
  }
}

function setEmployeeMessage(message) {
  applyStatusTone(SELECTORS.employeeMessage, message);
  SELECTORS.employeeMessage.textContent = message || "";
}

function renderDesignationSuggestions() {
  if (!SELECTORS.employeeDesignationInput) return;

  const values = new Map();
  for (const preset of designationPresets) {
    const name = String(preset?.name || "").trim();
    const key = name.toLowerCase();
    if (name && !values.has(key)) {
      values.set(key, name);
    }
  }

  const currentValue = String(SELECTORS.employeeDesignationInput.value || "").trim();
  const sortedValues = Array.from(values.values()).sort((a, b) => a.localeCompare(b));
  const hasPresets = sortedValues.length > 0;

  SELECTORS.employeeDesignationInput.innerHTML = [
    `<option value="">${hasPresets ? "Select designation" : "Add designation in Settings first"}</option>`,
    ...sortedValues.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`),
  ].join("");
  SELECTORS.employeeDesignationInput.disabled = !hasPresets;

  if (hasPresets && currentValue && sortedValues.includes(currentValue)) {
    SELECTORS.employeeDesignationInput.value = currentValue;
  } else {
    SELECTORS.employeeDesignationInput.value = "";
  }
}

function employeeApiErrorMessage(error) {
  const raw = String(error?.message || "");
  if (raw.toLowerCase() === "not found") {
    return "Employees API not available on current backend process (old server on port 5501). Stop it and run latest backend: cd /Users/grezello/Desktop/routes_payroll && node server.js";
  }
  return raw || "Employee request failed.";
}

function settingsApiErrorMessage(error) {
  const raw = String(error?.message || "");
  if (raw.toLowerCase() === "not found") {
    return "Settings API not available on current backend process (old server on port 5501). Stop it and run latest backend: cd /Users/grezello/Desktop/routes_payroll && node server.js";
  }
  return raw || "Settings request failed.";
}

function wireAppActions() {
  wireRailActions();
  wireCompanyActions();

  SELECTORS.monthPicker.addEventListener("change", async () => {
    await flushPendingSave();
    await loadMonthRecords();
  });

  SELECTORS.addEmployeeBtn.addEventListener("click", async () => {
    setWorkspace("employees");
    SELECTORS.railButtons.forEach((item) => item.classList.toggle("active", item.dataset.railAction === "employees"));
    SELECTORS.employeesSection?.scrollIntoView({ behavior: "smooth", block: "start" });
    setTimeout(() => SELECTORS.employeeNameInput?.focus(), 240);
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
      await loadEmployees();
      await loadMonthRecords();
      await loadDesignationPresets();
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
    employeeMaster = [];
    designationPresets = [];
    pendingSettingsLogoDataUrl = "";
    renderEmployeeTable();
    renderDesignationPresets();
    renderDesignationSuggestions();
    hydrateSettingsLogoPreview("");
    setSettingsMessage("");
    renderPayrollTable();
    await updateAuthView();
    showAuthMessage("Logged out.");
  });

  SELECTORS.payrollBody.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) {
      return;
    }

    const row = target.closest("[data-index]");
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
      if (!(target instanceof HTMLSelectElement)) return;
    }
    const row = target.closest("[data-index]");
    if (!row) return;
    const index = Number(row.dataset.index);
    if (Number.isNaN(index) || !currentRecords[index]) return;

    if (target instanceof HTMLSelectElement && target.dataset.action === "pick-record") {
      activePayrollEmployeeId = target.value || "";
      renderPayrollTable();
      return;
    }

    renderPayrollTable();
    scheduleSave(true);
  });

  SELECTORS.payrollBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button");
    if (!button) return;
    const row = button.closest("[data-index]");
    if (!row) return;
    const index = Number(row.dataset.index);
    if (Number.isNaN(index) || !currentRecords[index]) return;

    const action = button.dataset.action;
    if (action === "delete") {
      const removedId = currentRecords[index]?.employeeId || "";
      currentRecords.splice(index, 1);
      if (String(activePayrollEmployeeId) === String(removedId)) {
        activePayrollEmployeeId = "";
      }
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

function wireSettingsActions() {
  SELECTORS.verifyEmailBtn?.addEventListener("click", async () => {
    const button = SELECTORS.verifyEmailBtn;
    try {
      if (button) {
        button.disabled = true;
        button.textContent = "Sending...";
      }
      const response = await apiRequest("/api/auth/send-email-verification", {
        method: "POST",
      });
      const message = response.message || "Verification email sent.";
      setSettingsMessage(message);
      showAppMessage(message);
    } catch (error) {
      const message = String(error?.message || "Failed to send verification email.");
      setSettingsMessage(message);
      showAppMessage(message);
    } finally {
      if (button) {
        button.disabled = false;
        button.textContent = "Verify Email";
      }
    }
  });

  SELECTORS.settingsPayrollMonthMode?.addEventListener("change", async () => {
    const mode = SELECTORS.settingsPayrollMonthMode.value === "current" ? "current" : "previous";
    localStorage.setItem(STORAGE_KEYS.payrollMonthMode, mode);
    setDefaultMonth();
    try {
      await loadMonthRecords();
      setSettingsMessage(
        mode === "previous"
          ? "Payroll view now defaults to previous month. Example: in April, payroll opens March."
          : "Payroll view now defaults to current month."
      );
      showAppMessage(
        mode === "previous"
          ? "Payroll will open on the previous month by default."
          : "Payroll will open on the current month by default."
      );
    } catch (error) {
      const message = settingsApiErrorMessage(error);
      setSettingsMessage(message);
      showAppMessage(message);
    }
  });

  SELECTORS.designationPresetInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    SELECTORS.addDesignationBtn?.click();
  });

  SELECTORS.addDesignationBtn?.addEventListener("click", async () => {
    const name = SELECTORS.designationPresetInput?.value.trim() || "";
    if (name.length < 2) {
      setSettingsMessage("Designation name is too short.");
      showAppMessage("Designation name is too short.");
      return;
    }
    const exists = designationPresets.some((item) => String(item.name || "").toLowerCase() === name.toLowerCase());
    if (exists) {
      setSettingsMessage("Designation already exists.");
      showAppMessage("Designation already exists.");
      return;
    }

    const addBtn = SELECTORS.addDesignationBtn;
    try {
      if (addBtn) {
        addBtn.disabled = true;
        addBtn.textContent = "Adding...";
      }
      await apiRequest("/api/settings/designations", {
        method: "POST",
        body: {
          companyId: getSelectedCompanyId(),
          name,
        },
      });
      if (SELECTORS.designationPresetInput) {
        SELECTORS.designationPresetInput.value = "";
      }
      await loadDesignationPresets();
      setSettingsMessage("Designation added and saved to database.");
      showAppMessage("Designation added and saved to database.");
    } catch (error) {
      const message = settingsApiErrorMessage(error);
      setSettingsMessage(message);
      showAppMessage(message);
    } finally {
      if (addBtn) {
        addBtn.disabled = false;
        addBtn.textContent = "Add";
      }
    }
  });

  SELECTORS.saveCompanyNameBtn?.addEventListener("click", async () => {
    const companyId = getSelectedCompanyId();
    if (!companyId) {
      setSettingsMessage("Select a company first.");
      showAppMessage("Select a company first.");
      return;
    }

    const name = SELECTORS.settingsCompanyNameInput?.value.trim() || "";
    if (name.length < 2) {
      setSettingsMessage("Company name is too short.");
      showAppMessage("Company name is too short.");
      return;
    }

    const saveBtn = SELECTORS.saveCompanyNameBtn;
    try {
      if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.textContent = "Saving...";
      }
      const response = await apiRequest(`/api/companies/${companyId}`, {
        method: "PUT",
        body: { name },
      });
      const company = companies.find((item) => Number(item.id) === Number(companyId));
      if (company) {
        company.name = String(response?.company?.name || name);
      }
      await loadCompanies();
      hydrateSettingsCompany();
      renderPayrollTable();
      setSettingsMessage("Company name saved to database.");
      showAppMessage("Company name updated.");
    } catch (error) {
      const message = settingsApiErrorMessage(error);
      setSettingsMessage(message);
      showAppMessage(message);
    } finally {
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = "Save Company Name";
      }
    }
  });

  SELECTORS.designationList?.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button[data-action='delete-designation']");
    if (!button) return;

    const id = Number(button.dataset.id);
    if (!id) return;

    try {
      await apiRequest(`/api/settings/designations/${id}?companyId=${getSelectedCompanyId()}`, {
        method: "DELETE",
      });
      await loadDesignationPresets();
      setSettingsMessage("Designation removed from database.");
      showAppMessage("Designation removed from database.");
    } catch (error) {
      const message = settingsApiErrorMessage(error);
      setSettingsMessage(message);
      showAppMessage(message);
    }
  });

  SELECTORS.settingsLogoInput?.addEventListener("change", async (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    if (!file) return;
    if (file.size > 700 * 1024) {
      setSettingsMessage("Logo too large. Keep it under 700KB.");
      input.value = "";
      return;
    }

    try {
      pendingSettingsLogoDataUrl = await fileToDataUrl(file);
      hydrateSettingsLogoPreview(pendingSettingsLogoDataUrl);
      setSettingsMessage("Logo selected. Click Save Logo.");
    } catch (error) {
      setSettingsMessage(error.message || "Failed to read logo file.");
    }
  });

  SELECTORS.saveLogoBtn?.addEventListener("click", async () => {
    const companyId = getSelectedCompanyId();
    if (!companyId) {
      setSettingsMessage("Select a company first.");
      showAppMessage("Select a company first.");
      return;
    }

    const currentLogo = String(getActiveCompany()?.logoDataUrl || "");
    const logoDataUrl = pendingSettingsLogoDataUrl || currentLogo;
    if (!logoDataUrl) {
      setSettingsMessage("Choose a logo first.");
      showAppMessage("Choose a logo first.");
      return;
    }

    const saveBtn = SELECTORS.saveLogoBtn;
    try {
      if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.textContent = "Saving...";
      }
      await apiRequest(`/api/companies/${companyId}/logo`, {
        method: "PUT",
        body: { logoDataUrl },
      });
      const company = companies.find((item) => Number(item.id) === Number(companyId));
      if (company) company.logoDataUrl = logoDataUrl;
      pendingSettingsLogoDataUrl = "";
      if (SELECTORS.settingsLogoInput) SELECTORS.settingsLogoInput.value = "";
      hydrateSettingsLogoPreview(logoDataUrl);
      setSettingsMessage("Company logo saved to database.");
      showAppMessage("Company logo updated for payslips.");
    } catch (error) {
      const message = settingsApiErrorMessage(error);
      setSettingsMessage(message);
      showAppMessage(message);
    } finally {
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = "Save Logo";
      }
    }
  });

  SELECTORS.changePasswordForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const newPassword = SELECTORS.newPasswordInput?.value || "";
    const confirmPassword = SELECTORS.confirmPasswordInput?.value || "";
    const submitBtn = SELECTORS.changePasswordBtn;

    if (newPassword.length < 8) {
      setSettingsMessage("New password must be at least 8 characters.");
      showAppMessage("New password must be at least 8 characters.");
      return;
    }
    if (!/[a-z]/.test(newPassword) || !/[A-Z]/.test(newPassword) || !/[0-9]/.test(newPassword)) {
      setSettingsMessage("Password must include uppercase, lowercase, and number.");
      showAppMessage("Password must include uppercase, lowercase, and number.");
      return;
    }
    if (newPassword !== confirmPassword) {
      setSettingsMessage("New password and confirm password do not match.");
      showAppMessage("New password and confirm password do not match.");
      return;
    }

    try {
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = "Updating...";
      }
      await apiRequest("/api/auth/change-password", {
        method: "POST",
        body: { newPassword, confirmPassword },
      });
      SELECTORS.changePasswordForm.reset();
      setSettingsMessage("Password updated successfully.");
      showAppMessage("Password updated successfully.");
    } catch (error) {
      const message = String(error?.message || "Failed to update password.");
      setSettingsMessage(message);
      showAppMessage(message);
    } finally {
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = "Update Password";
      }
    }
  });

  SELECTORS.importLegacyBtn?.addEventListener("click", async () => {
    const file = SELECTORS.legacyImportInput?.files?.[0];
    if (!file) {
      setSettingsMessage("Choose an Excel/CSV file first.");
      showAppMessage("Choose an Excel/CSV file first.");
      return;
    }

    const btn = SELECTORS.importLegacyBtn;
    try {
      if (btn) {
        btn.disabled = true;
        btn.textContent = "Analyzing...";
      }
      setImportSummary("Analyzing file...");
      const rows = await parseLegacyPayrollFile(file);
      if (rows.length === 0) {
        setImportSummary("Import failed: no usable rows found in file.");
        setSettingsMessage("Import failed: no usable rows found in file.");
        return;
      }

      const analyzed = analyzeLegacyRows(rows);
      const inferenceText = analyzed.diagnostics?.inferenceNotes?.length
        ? ` Mapped fields: ${analyzed.diagnostics.inferenceNotes.slice(0, 4).join("; ")}.`
        : "";
      if (analyzed.payrollRecords.length === 0) {
        const reasonText = analyzed.diagnostics?.reasons?.length
          ? analyzed.diagnostics.reasons.join(" ")
          : "No payroll rows matched expected columns.";
        setImportSummary(`Import failed: ${reasonText}${inferenceText}`);
        setSettingsMessage(`Import failed: ${reasonText}`);
        return;
      }

      if (btn) btn.textContent = "Importing... 1%";
      setImportSummary(`Rows analyzed: ${analyzed.totalRows}. Importing... 1%`);
      await importLegacyAnalysis(analyzed, ({ percent, label }) => {
        const safePercent = Math.max(1, Math.min(100, Number(percent) || 1));
        if (btn) btn.textContent = `Importing... ${safePercent}%`;
        setImportSummary(`Rows analyzed: ${analyzed.totalRows}. Importing... ${safePercent}%${label ? ` (${label})` : ""}`);
      });

      await Promise.all([
        loadDesignationPresets(),
        loadEmployees(),
        loadMonthRecords(),
      ]);
      if (btn) btn.textContent = "Importing... 100%";
      const warningText = analyzed.diagnostics?.warnings?.length
        ? ` Warnings: ${analyzed.diagnostics.warnings.join(" ")}`
        : "";
      setImportSummary(`Import successful! 100%. Imported ${analyzed.payrollRecords.length} payroll rows across ${analyzed.months.length} month(s).${warningText}${inferenceText}`);
      setSettingsMessage(`Import successful!${warningText}`);
      showAppMessage("Import successful!");
      if (SELECTORS.legacyImportInput) SELECTORS.legacyImportInput.value = "";
    } catch (error) {
      const message = String(error?.message || "Legacy import failed.");
      setSettingsMessage(message);
      showAppMessage(message);
      setImportSummary(`Import failed: ${message}`);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = "Analyze & Import";
      }
    }
  });
}

async function loadDesignationPresets() {
  try {
    const companyId = getSelectedCompanyId();
    const response = await apiRequest(`/api/settings/designations?companyId=${companyId}`);
    designationPresets = Array.isArray(response.designations) ? response.designations : [];
    renderDesignationPresets();
    renderDesignationSuggestions();
    hydrateSettingsLogoPreview(String(getActiveCompany()?.logoDataUrl || ""));
    setSettingsMessage("");
  } catch (error) {
    designationPresets = [];
    renderDesignationPresets();
    renderDesignationSuggestions();
    setSettingsMessage(settingsApiErrorMessage(error));
  }
}

function renderDesignationPresets() {
  if (!SELECTORS.designationList) return;
  if (designationPresets.length === 0) {
    SELECTORS.designationList.innerHTML = '<p class="empty">No designation presets yet.</p>';
    return;
  }

  SELECTORS.designationList.innerHTML = designationPresets
    .map((item) => `
      <div class="designation-item" data-id="${Number(item.id) || 0}">
        <span>${escapeHtml(String(item.name || ""))}</span>
        <button type="button" class="mini danger ghost" data-action="delete-designation" data-id="${Number(item.id) || 0}">Delete</button>
      </div>
    `)
    .join("");
}

function hydrateSettingsLogoPreview(logoDataUrl) {
  if (!SELECTORS.settingsLogoPreview) return;
  const src = String(logoDataUrl || "");
  if (!src) {
    SELECTORS.settingsLogoPreview.removeAttribute("src");
    SELECTORS.settingsLogoPreview.classList.add("hidden");
    return;
  }
  SELECTORS.settingsLogoPreview.src = src;
  SELECTORS.settingsLogoPreview.classList.remove("hidden");
}

function setSettingsMessage(message) {
  if (!SELECTORS.settingsMessage) return;
  applyStatusTone(SELECTORS.settingsMessage, message);
  SELECTORS.settingsMessage.textContent = message || "";
}

function applyEmployeeSelection(index, employeeId) {
  const picked = employeeMaster.find((item) => String(item.employeeId) === String(employeeId));
  if (!picked || !currentRecords[index]) return;

  currentRecords[index].employeeId = picked.employeeId || "";
  currentRecords[index].employeeName = picked.employeeName || "";
  currentRecords[index].designation = picked.designation || "";
  currentRecords[index].presentSalary = Number(picked.baseSalary || 0);
  currentRecords[index].employeeStatus = picked.status || "working";
  currentRecords[index].leaveFrom = picked.leaveFrom || "";
  currentRecords[index].leaveTo = picked.leaveTo || "";
  currentRecords[index].terminatedOn = picked.terminatedOn || "";
}

function revealPayrollSectionsForRow(row) {
  if (!row) return;
  const sections = row.querySelectorAll(".payroll-reveal-section");
  sections.forEach((section) => {
    section.setAttribute("open", "");
  });
}

function wireCompanyActions() {
  SELECTORS.companyPicker.addEventListener("change", async () => {
    await flushPendingSave();
    activeCompanyId = Number(SELECTORS.companyPicker.value) || null;
    await loadEmployees();
    await loadMonthRecords();
    await loadDesignationPresets();
    resetEmployeeForm();
    setSettingsMessage("");
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
      await loadEmployees();
      await loadMonthRecords();
      await loadDesignationPresets();
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
        setWorkspace("dashboard");
        SELECTORS.dashboardSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }

      if (action === "employees") {
        setWorkspace("employees");
        SELECTORS.employeesSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        setTimeout(() => {
          SELECTORS.employeeNameInput?.focus();
        }, 260);
        return;
      }

      if (action === "payroll") {
        setWorkspace("payroll");
        SELECTORS.payrollSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        setTimeout(() => {
          const incrementField = SELECTORS.payrollBody.querySelector('input[data-field="increment"]:not([disabled])');
          incrementField?.focus();
        }, 240);
        return;
      }

      if (action === "reports") {
        setWorkspace("reports");
        SELECTORS.reportsSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }

      if (action === "settings") {
        setWorkspace("settings");
        SELECTORS.settingsSection?.scrollIntoView({ behavior: "smooth", block: "start" });
        setTimeout(() => {
          SELECTORS.designationPresetInput?.focus();
        }, 220);
      }
    });
  });
}

function setWorkspace(view) {
  const dashboardView = view === "dashboard";
  const employeesView = view === "employees";
  const payrollView = view === "payroll";
  const reportsView = view === "reports";
  const settingsView = view === "settings";

  SELECTORS.metricsSection?.classList.toggle("hidden", !dashboardView);
  SELECTORS.employeesSection?.classList.toggle("hidden", !employeesView);
  SELECTORS.payrollSection?.classList.toggle("hidden", !payrollView);
  SELECTORS.reportsSection?.classList.toggle("hidden", !reportsView);
  SELECTORS.settingsSection?.classList.toggle("hidden", !settingsView);
  SELECTORS.saveStatus?.classList.toggle("hidden", !payrollView);
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
    stopServerReconnectPolling();
    needsSetupFlow = Boolean(bootstrap.needsSetup);
    SELECTORS.showRegisterBtn?.classList.remove("hidden");
    SELECTORS.showLoginBtn?.classList.remove("hidden");

    if (bootstrap.degraded) {
      setAuthMode("login");
      showAuthMessage(bootstrap.message || "Cloud service is temporarily busy. Please retry shortly.");
      return;
    }

    if (bootstrap.needsSetup) {
      setAuthMode("register");
      showAuthMessage("First-time setup: register company + admin email + username + password.");
    } else {
      setAuthMode("login");
      showAuthMessage("Company is already registered. Sign in with username or email.");
    }
  } catch {
    SELECTORS.setupForm.classList.add("hidden");
    SELECTORS.loginForm.classList.add("hidden");
    startServerReconnectPolling();
    showAuthMessage("Server is not reachable. Start backend with `npm start` or double-click `start-routes-payroll.command`. We will reconnect automatically.");
  }
}

function setAuthMode(mode) {
  const showRegister = mode === "register";
  SELECTORS.setupForm.classList.toggle("hidden", !showRegister);
  SELECTORS.loginForm.classList.toggle("hidden", showRegister);
  SELECTORS.showRegisterBtn?.classList.toggle("active", showRegister);
  SELECTORS.showLoginBtn?.classList.toggle("active", !showRegister);
}

async function validateToken() {
  if (!authToken) return false;
  try {
    const response = await apiRequest("/api/auth/me");
    currentUser = response.user || null;
    hydrateSettingsAccount();
    return true;
  } catch (error) {
    if (String(error.message).toLowerCase().includes("expired") || String(error.message).toLowerCase().includes("authentication")) {
      setToken("");
    }
    currentUser = null;
    hydrateSettingsAccount();
    return false;
  }
}

async function switchToApp() {
  stopServerReconnectPolling();
  SELECTORS.authView.classList.add("hidden");
  SELECTORS.appView.classList.remove("hidden");
  hydrateSettingsAccount();
  await loadCompanies();
  await Promise.all([
    loadEmployees(),
    loadMonthRecords(),
    loadDesignationPresets(),
  ]);
  setWorkspace("dashboard");
  SELECTORS.railButtons.forEach((item) => item.classList.toggle("active", item.dataset.railAction === "dashboard"));
}

async function isServerReachable() {
  try {
    const response = await fetch("/api/health", { cache: "no-store" });
    return response.ok;
  } catch {
    return false;
  }
}

function stopServerReconnectPolling() {
  if (serverReconnectTimer) {
    window.clearInterval(serverReconnectTimer);
    serverReconnectTimer = null;
  }
  serverReconnectInFlight = false;
}

function startServerReconnectPolling() {
  if (serverReconnectTimer) return;
  serverReconnectTimer = window.setInterval(async () => {
    if (serverReconnectInFlight) return;
    serverReconnectInFlight = true;
    try {
      const reachable = await isServerReachable();
      if (reachable) {
        stopServerReconnectPolling();
        await updateAuthView();
      }
    } finally {
      serverReconnectInFlight = false;
    }
  }, 3000);
}

function hydrateSettingsAccount() {
  if (SELECTORS.settingsUsernameValue) {
    SELECTORS.settingsUsernameValue.textContent = currentUser?.username || "-";
  }
  if (SELECTORS.settingsEmailValue) {
    SELECTORS.settingsEmailValue.textContent = currentUser?.email || "-";
  }
  if (SELECTORS.settingsEmailStatusValue) {
    SELECTORS.settingsEmailStatusValue.textContent = currentUser?.emailVerified ? "Verified" : "Not Verified";
  }
  if (SELECTORS.verifyEmailBtn) {
    SELECTORS.verifyEmailBtn.classList.toggle("hidden", Boolean(currentUser?.emailVerified) || !currentUser?.email);
  }
  hydrateSettingsCompany();
  hydratePayrollMonthModeSetting();
}

function hydrateSettingsCompany() {
  if (!SELECTORS.settingsCompanyNameInput) return;
  SELECTORS.settingsCompanyNameInput.value = displayCompanyName(getActiveCompany());
}

function getPayrollMonthMode() {
  const stored = localStorage.getItem(STORAGE_KEYS.payrollMonthMode) || "previous";
  return stored === "current" ? "current" : "previous";
}

function hydratePayrollMonthModeSetting() {
  if (!SELECTORS.settingsPayrollMonthMode) return;
  SELECTORS.settingsPayrollMonthMode.value = getPayrollMonthMode();
}

function setDefaultMonth() {
  const now = new Date();
  if (getPayrollMonthMode() === "previous") {
    now.setMonth(now.getMonth() - 1);
  }
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  SELECTORS.monthPicker.value = month;
}

function getSelectedMonth() {
  return SELECTORS.monthPicker.value || new Date().toISOString().slice(0, 7);
}

function isPayrollActiveStatus(status) {
  const normalized = String(status || "working").toLowerCase();
  return normalized !== "terminated" && normalized !== "leave";
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
  hydrateSettingsCompany();
}

async function loadMonthRecords() {
  const month = getSelectedMonth();
  try {
    const response = await apiRequest(payrollMonthUrl(month));
    currentRecords = Array.isArray(response.records) ? response.records : [];
    const visible = currentRecords.filter((record) => isPayrollActiveStatus(record.employeeStatus));
    if (!visible.some((record) => String(record.employeeId) === String(activePayrollEmployeeId))) {
      activePayrollEmployeeId = visible[0]?.employeeId || "";
    }
    renderPayrollTable();
    setSaveStatus("All changes saved.");
  } catch (error) {
    currentRecords = [];
    activePayrollEmployeeId = "";
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

function monthBounds(isoMonth) {
  if (!isoMonth || !isoMonth.includes("-")) {
    const now = new Date();
    return {
      start: new Date(now.getFullYear(), now.getMonth(), 1),
      end: new Date(now.getFullYear(), now.getMonth() + 1, 0),
    };
  }
  const [year, month] = isoMonth.split("-").map((item) => Number(item));
  return {
    start: new Date(year, month - 1, 1),
    end: new Date(year, month, 0),
  };
}

function parseIsoDate(value) {
  if (!value || !/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function isPayrollBlocked(record, month) {
  const status = String(record.employeeStatus || "working").toLowerCase();
  if (status === "terminated") {
    const terminatedOn = parseIsoDate(record.terminatedOn);
    if (!terminatedOn) return true;
    const { end } = monthBounds(month);
    return terminatedOn <= end;
  }
  if (status === "leave") {
    const { start, end } = monthBounds(month);
    const leaveFrom = parseIsoDate(record.leaveFrom);
    const leaveTo = parseIsoDate(record.leaveTo);
    if (!leaveFrom && !leaveTo) return true;
    if (leaveFrom && leaveFrom > end) return false;
    if (leaveTo && leaveTo < start) return false;
    return true;
  }
  return false;
}

function daysBeforeResumeInMonth(record, month) {
  const status = String(record.employeeStatus || "working").toLowerCase();
  if (status !== "resumed") return 0;
  const resumedOn = parseIsoDate(record.leaveTo);
  if (!resumedOn) return 0;
  const { start, end } = monthBounds(month);
  if (resumedOn <= start) return 0;
  if (resumedOn > end) return 30;
  return clamp(resumedOn.getDate() - 1, 0, 30);
}

function computePayroll(record, month = getSelectedMonth()) {
  const blocked = isPayrollBlocked(record, month);
  const presentSalary = blocked ? 0 : toMoney(record.presentSalary);
  const increment = blocked ? 0 : toMoney(record.increment);
  const grossSalary = presentSalary + increment;

  const oldAdvanceTaken = toMoney(record.oldAdvanceTaken);
  const extraAdvanceAdded = toMoney(record.extraAdvanceAdded);
  const totalAdvance = oldAdvanceTaken + extraAdvanceAdded;

  const deductionEntered = blocked ? 0 : Math.max(0, toMoney(record.deductionEntered));
  const manualDaysAbsent = blocked ? 0 : toMoney(record.daysAbsent);
  const resumedDaysAbsent = blocked ? 0 : daysBeforeResumeInMonth(record, month);
  const daysAbsent = clamp(manualDaysAbsent + resumedDaysAbsent, 0, 30);
  const proratedAbsenceDeduction = (daysAbsent / 30) * grossSalary;
  const deductionApplied = Math.min(deductionEntered, totalAdvance);
  const advanceRemained = totalAdvance - deductionApplied;
  const netSalary = grossSalary - deductionApplied - proratedAbsenceDeduction;

  return {
    blocked,
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
  const now = new Date();
  const currentMonthIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const isCurrentMonth = month === currentMonthIso;
  const isFirstMonth = (() => {
    if (!currentRecords.length) return true;
    const months = currentRecords.map(r => r.month || month).filter(Boolean);
    const sorted = months.slice().sort();
    return month === sorted[0];
  })();
  const allowOverride = isFirstMonth || isCurrentMonth;
  const visibleRecords = currentRecords
    .map((record, index) => ({ record, index }))
    .filter((item) => isPayrollActiveStatus(item.record.employeeStatus));
  const visibleCount = visibleRecords.length;

  if (visibleCount === 0) {
    SELECTORS.payrollBody.innerHTML = `
      <div class="empty">
        No employees for ${formatMonth(month)}. Add employees from Employee Management.
      </div>
    `;
    updateMetrics([]);
    return;
  }

  const activeEntry = visibleRecords.find((item) => String(item.record.employeeId) === String(activePayrollEmployeeId)) || visibleRecords[0];
  const record = activeEntry.record;
  const index = activeEntry.index;
  activePayrollEmployeeId = record.employeeId || "";
  const calc = computePayroll(record, month);
  const status = String(record.employeeStatus || "working").toLowerCase();
  const statusLabel = status === "leave" ? "On Leave" : "Working";
  const badgeClass = status === "leave" ? "status-pill leave" : "status-pill working";
  const statusBoxClass = status === "leave" ? "status-box leave" : "status-box active";
  const statusDetail = status === "leave"
    ? `Leave: ${record.leaveFrom || "-"} to ${record.leaveTo || "-"}`
    : status === "resumed"
      ? `Resumed on: ${record.leaveTo || "-"}`
    : "";
  const employeeOptions = visibleRecords
    .map((item) => {
      const selected = String(item.record.employeeId) === String(record.employeeId) ? "selected" : "";
      return `<option value="${escapeHtml(item.record.employeeId)}" ${selected}>${escapeHtml(item.record.employeeId)} - ${escapeHtml(item.record.employeeName)}</option>`;
    })
    .join("");

  const selectClass = status === "leave" ? "emp-picker leave" : "emp-picker working";
  SELECTORS.payrollBody.innerHTML = `
    <article class="payroll-entry-card payroll-modern" data-index="${index}">
      <div class="payroll-modern-grid">
        <div class="payroll-modern-main">
          <section class="payroll-stack-section">
            <h3>Employee Profile</h3>
            <div class="payroll-fields">
              <label>Employee ID
                <select data-action="pick-record" class="${selectClass}">
                  ${employeeOptions}
                </select>
              </label>
              <label>Employee Name
                <input data-field="employeeName" value="${escapeHtml(record.employeeName || "")}" readonly />
              </label>
              <label>Designation
                <input data-field="designation" value="${escapeHtml(record.designation || "")}" readonly />
              </label>
              <label>Status
                <div class="${statusBoxClass}">${status === "leave" ? "ON LEAVE" : "ACTIVE ✓"}</div>
              </label>
            </div>
            ${statusDetail ? `<p class="status-note">${escapeHtml(statusDetail)}</p>` : ""}
          </section>

          <section class="payroll-stack-section">
            <h3>Core Salary & Increment</h3>
            <div class="payroll-fields">
              <label>Present Salary (₹)
                <input class="field-salary locked-field" data-field="presentSalary" type="number" min="0" step="0.01" value="${toRaw(record.presentSalary)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
              <label>Increment (₹)
                <span class="field-tag">NEW</span>
                <input class="field-increment" data-field="increment" type="number" min="0" step="0.01" value="${toRaw(record.increment)}" ${calc.blocked || !allowOverride ? "disabled" : ""} />
                ${!allowOverride ? '<span class="field-note">Locked (past month)</span>' : ''}
              </label>
              <label>Gross Salary (₹)
                <input type="text" value="${formatCurrency(calc.grossSalary)}" readonly />
              </label>
            </div>
          </section>

          <section class="payroll-stack-section">
            <h3>Advance Tracking</h3>
            <div class="payroll-fields">
              <label>Old Advance Taken (₹)
                <input class="field-advance locked-field" data-field="oldAdvanceTaken" type="number" min="0" step="0.01" value="${toRaw(record.oldAdvanceTaken)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
              <label>Extra Advance Added (₹)
                <input class="field-advance" data-field="extraAdvanceAdded" type="number" min="0" step="0.01" value="${toRaw(record.extraAdvanceAdded)}" ${allowOverride ? "" : "readonly"} />
                ${!allowOverride ? '<span class="field-note">Locked (past month)</span>' : ''}
              </label>
              <label>Total Advance (₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.totalAdvance)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
            </div>
          </section>

          <section class="payroll-stack-section">
            <h3>Deductions & Absence</h3>
            <div class="payroll-fields">
              <label>Deduction Entered (₹)
                <input class="field-deduction" data-field="deductionEntered" type="number" min="0" step="0.01" value="${toRaw(record.deductionEntered)}" ${allowOverride ? "" : "readonly"} />
                ${!allowOverride ? '<span class="field-note">Locked (past month)</span>' : ''}
              </label>
              <label>Days Absent (out of 30)
                <input class="field-absent" data-field="daysAbsent" type="number" min="0" max="30" step="1" value="${toRaw(record.daysAbsent)}" ${allowOverride ? "" : "readonly"} />
                ${!allowOverride ? '<span class="field-note">Locked (past month)</span>' : ''}
              </label>
              <label>Prorated Absence Deduction (₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.proratedAbsenceDeduction)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
              <label>Deduction Applied (Advance only)
                <input type="text" class="locked-field" value="${formatCurrency(calc.deductionApplied)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
            </div>
          </section>

          <section class="payroll-stack-section">
            <h3>Summary</h3>
            <div class="payroll-fields">
              <label>Advance Remained (₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.advanceRemained)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
              <label>Salary In Hand (Net Salary ₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.netSalary)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
            </div>
          </section>
        </div>

        <aside class="payroll-modern-side">
          <div class="kpi-card net">
            <p>Net Salary In Hand</p>
            <strong>${formatCurrency(calc.netSalary)}</strong>
            <div class="net-sub">Calculated: Gross - Deductions</div>
          </div>

          <div class="kpi-card soft">
            <h4>Deductions</h4>
            <div class="kpi-line"><span>Deduction Entered</span><b>${formatCurrency(calc.deductionEntered)}</b></div>
            <div class="kpi-line"><span>Absence Deduction</span><b>${formatCurrency(calc.proratedAbsenceDeduction)}</b></div>
            <div class="kpi-line total"><span>Total Applied</span><b>${formatCurrency(calc.deductionApplied)}</b></div>
          </div>

          <div class="kpi-card amber">
            <p>Advance Remained</p>
            <strong>${formatCurrency(calc.advanceRemained)}</strong>
          </div>
        </aside>
      </div>

      <section class="payroll-footer">
        <label>Notes / Comments
          <textarea data-field="comment" rows="3" placeholder="Add professional notes here...">${escapeHtml(record.comment || "")}</textarea>
        </label>
        <div class="row-actions">
          <button data-action="payslip" class="mini" type="button">Generate Payslip</button>
          <button data-action="delete" class="mini danger" type="button">Delete Record</button>
        </div>
      </section>
    </article>
  `;

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
        <div><dt>Status</dt><dd>${escapeHtml(String(record.employeeStatus || "working"))}</dd></div>
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
    `Status: ${record.employeeStatus || "working"}`,
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
    ["Status", escapeHtml(record.employeeStatus || "working")],
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
    "Status",
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
      record.employeeStatus || "working",
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

function setImportSummary(message) {
  if (!SELECTORS.importLegacySummary) return;
  SELECTORS.importLegacySummary.textContent = String(message || "");
}

async function parseLegacyPayrollFile(file) {
  const name = String(file?.name || "").toLowerCase();
  if (name.endsWith(".csv")) {
    const text = await file.text();
    return parseCsvRows(text);
  }

  if (!window.XLSX) {
    throw new Error("Excel parser not loaded. Reload app and try again.");
  }
  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) return [];

  // Helper to convert sheet name like 'Mar_2026' or 'Apr_2025' to '2026-03' or '2025-04'
  function parseMonthFromSheetName(sheetName) {
    const match = String(sheetName).match(/([A-Za-z]{3})[_-](\d{4})/);
    if (match) {
      const monthMap = {
        jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
        jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12',
      };
      const m = monthMap[match[1].toLowerCase()];
      if (m) return `${match[2]}-${m}`;
    }
    return '';
  }

  let allRows = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    const monthFromSheet = parseMonthFromSheetName(sheetName);
    if (Array.isArray(rows) && rows.length > 0) {
      // If no explicit month column, add it from the sheet name
      for (const row of rows) {
        const hasMonth = Object.keys(row).some(k => /month|period|year/i.test(k));
        if (!hasMonth && monthFromSheet) {
          row["Month"] = monthFromSheet;
        }
      }
      allRows = allRows.concat(rows);
    }
  }
  return allRows;
}

function parseCsvRows(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0);
  if (lines.length < 2) return [];
  const headers = parseCsvLine(lines[0]).map((item) => String(item || "").trim());
  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const values = parseCsvLine(lines[i]);
    const row = {};
    headers.forEach((key, idx) => {
      row[key] = values[idx] ?? "";
    });
    rows.push(row);
  }
  return rows;
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === "," && !inQuotes) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out.map((item) => item.trim());
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function headerMatchesAlias(normalizedHeader, alias) {
  const normalizedAlias = normalizeHeader(alias);
  if (!normalizedHeader || !normalizedAlias) return false;
  if (normalizedHeader === normalizedAlias) return true;
  if (normalizedHeader.includes(normalizedAlias) || normalizedAlias.includes(normalizedHeader)) return true;

  const headerTokens = normalizedHeader.split(" ").filter(Boolean);
  const aliasTokens = normalizedAlias.split(" ").filter(Boolean);
  if (aliasTokens.length > 0 && aliasTokens.every((token) => headerTokens.includes(token))) return true;
  if (headerTokens.length > 0 && headerTokens.every((token) => aliasTokens.includes(token))) return true;
  return false;
}

function collectImportHeaders(rows) {
  const headerMap = new Map();
  for (const row of rows || []) {
    for (const key of Object.keys(row || {})) {
      const normalized = normalizeHeader(key);
      if (normalized && !headerMap.has(normalized)) {
        headerMap.set(normalized, key);
      }
    }
  }
  return headerMap;
}

function detectImportHeader(headerMap, aliases) {
  for (const [normalized, original] of headerMap.entries()) {
    if (aliases.some((alias) => headerMatchesAlias(normalized, alias))) {
      return original;
    }
  }
  return "";
}

function parseMoney(value) {
  const cleaned = String(value ?? "")
    .replace(/[₹,\s]/g, "")
    .replace(/[^0-9.-]/g, "");
  const num = Number(cleaned);
  return Number.isFinite(num) ? Math.max(0, num) : 0;
}

function parseMonthFromImport(value) {
  const raw = String(value || "").trim();
  if (!raw) return getSelectedMonth();
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}\/\d{2}$/.test(raw)) return raw.replace("/", "-");
  if (/^\d{2}[-/]\d{4}$/.test(raw)) {
    const [m, y] = raw.split(/[-/]/);
    return `${y}-${m}`;
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }
  return getSelectedMonth();
}

function parseIsoFromImport(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function isLikelyEmployeeId(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (raw.length > 24) return false;
  return /[a-z]/i.test(raw) && /\d/.test(raw);
}

function isLikelyPersonName(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (raw.length < 2 || raw.length > 80) return false;
  if (/\d{3,}/.test(raw)) return false;
  return /^[a-z .'-]+$/i.test(raw);
}

function isLikelyMonthValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  return /^\d{4}[-/]\d{2}$/.test(raw)
    || /^\d{2}[-/]\d{4}$/.test(raw)
    || /^[a-z]{3,9}[\s_-]?\d{4}$/i.test(raw)
    || !Number.isNaN(new Date(raw).getTime());
}

function isLikelyMoneyValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  return Number.isFinite(Number(String(raw).replace(/[₹,\s]/g, "").replace(/[^0-9.-]/g, "")));
}

function isLikelyDaysAbsentValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  const num = Number(raw);
  return Number.isFinite(num) && num >= 0 && num <= 31;
}

function buildImportSchema(rows) {
  const headerMap = collectImportHeaders(rows);
  const headers = Array.from(headerMap.entries()).map(([normalized, original]) => ({ normalized, original }));
  const sampleRows = Array.isArray(rows) ? rows.slice(0, 25) : [];
  const fieldDefinitions = {
    employeeId: {
      aliases: ["employee id", "emp id", "employeeid", "staff id", "staff code", "id code", "code"],
      scoreValue: isLikelyEmployeeId,
    },
    employeeName: {
      aliases: ["employee name", "employee names", "name", "emp name", "staff name", "full name", "worker name"],
      scoreValue: isLikelyPersonName,
    },
    designation: {
      aliases: ["designation", "role", "position", "job title", "title", "department role"],
      scoreValue: (value) => {
        const raw = String(value || "").trim();
        return raw.length >= 2 && raw.length <= 60 && !/\d{4,}/.test(raw);
      },
    },
    month: {
      aliases: ["year month", "updated month", "month", "period", "salary month", "pay month", "for month"],
      scoreValue: isLikelyMonthValue,
    },
    presentSalary: {
      aliases: ["present salary", "salary", "base salary", "net salary", "gross salary", "monthly salary", "salary amount"],
      scoreValue: isLikelyMoneyValue,
    },
    increment: {
      aliases: ["increment", "salary increment", "raise", "hike"],
      scoreValue: isLikelyMoneyValue,
    },
    oldAdvanceTaken: {
      aliases: ["old advance taken", "old advance", "advance taken", "previous advance", "opening advance"],
      scoreValue: isLikelyMoneyValue,
    },
    extraAdvanceAdded: {
      aliases: ["extra advance added", "extra advance", "new advance", "advance added", "fresh advance"],
      scoreValue: isLikelyMoneyValue,
    },
    deductionEntered: {
      aliases: ["deduction entered", "deduction", "advance deduction", "recovery", "recovered amount"],
      scoreValue: isLikelyMoneyValue,
    },
    daysAbsent: {
      aliases: ["days absent", "absent", "attendance absent", "absent days", "leave days"],
      scoreValue: isLikelyDaysAbsentValue,
    },
    comment: {
      aliases: ["comment", "notes", "remark", "remarks", "description"],
      scoreValue: (value) => String(value || "").trim().length > 0,
    },
    joiningDate: {
      aliases: ["joining date", "date of joining", "doj", "joined on"],
      scoreValue: (value) => Boolean(parseIsoFromImport(value)),
    },
    birthDate: {
      aliases: ["birth date", "date of birth", "dob"],
      scoreValue: (value) => Boolean(parseIsoFromImport(value)),
    },
    mobileNumber: {
      aliases: ["mobile", "phone", "contact", "mobile number", "phone number"],
      scoreValue: (value) => /\d{7,}/.test(String(value || "").replace(/\D/g, "")),
    },
  };

  const schema = {};
  const inferenceNotes = [];

  for (const [field, definition] of Object.entries(fieldDefinitions)) {
    let best = { original: "", score: -1, matchedBy: "" };
    for (const header of headers) {
      let score = 0;
      let matchedBy = "content";
      for (const alias of definition.aliases) {
        if (headerMatchesAlias(header.normalized, alias)) {
          score += 12;
          matchedBy = "header";
        }
      }

      let sampleMatches = 0;
      let sampleChecks = 0;
      for (const row of sampleRows) {
        const rawValue = row?.[header.original];
        if (String(rawValue || "").trim() === "") continue;
        sampleChecks += 1;
        if (definition.scoreValue(rawValue)) sampleMatches += 1;
      }
      if (sampleChecks > 0) {
        score += Math.round((sampleMatches / sampleChecks) * 8);
      }

      if (score > best.score) {
        best = { original: header.original, score, matchedBy };
      }
    }

    if (best.original && best.score >= 6) {
      schema[field] = best.original;
      inferenceNotes.push(`${field} -> ${best.original} (${best.matchedBy} inference)`);
    } else {
      schema[field] = "";
    }
  }

  return { schema, inferenceNotes };
}

function pickCellFromSchema(row, schema, field) {
  const header = schema?.[field];
  if (!header) return "";
  const value = row?.[header];
  return String(value || "").trim() === "" ? "" : value;
}

function analyzeLegacyRows(rows) {
  const payrollRecords = [];
  const designationSet = new Set();
  const employeeMap = new Map();
  const monthSet = new Set();
  const { schema: detectedHeaders, inferenceNotes } = buildImportSchema(rows);
  let skippedMissingEmployeeId = 0;
  let skippedMissingEmployeeName = 0;
  let defaultedMonthCount = 0;

  for (const row of rows) {
    const employeeId = String(pickCellFromSchema(row, detectedHeaders, "employeeId") || "").trim();
    const employeeName = String(pickCellFromSchema(row, detectedHeaders, "employeeName") || "").trim();
    if (!employeeId) {
      skippedMissingEmployeeId += 1;
      continue;
    }
    if (!employeeName) {
      skippedMissingEmployeeName += 1;
      continue;
    }

    const designation = String(pickCellFromSchema(row, detectedHeaders, "designation") || "").trim();
    const rawMonthValue = pickCellFromSchema(row, detectedHeaders, "month");
    const month = parseMonthFromImport(rawMonthValue);
    const presentSalary = parseMoney(pickCellFromSchema(row, detectedHeaders, "presentSalary"));
    const increment = parseMoney(pickCellFromSchema(row, detectedHeaders, "increment"));
    const oldAdvanceTaken = parseMoney(pickCellFromSchema(row, detectedHeaders, "oldAdvanceTaken"));
    const extraAdvanceAdded = parseMoney(pickCellFromSchema(row, detectedHeaders, "extraAdvanceAdded"));
    const deductionEntered = parseMoney(pickCellFromSchema(row, detectedHeaders, "deductionEntered"));
    const daysAbsentRaw = Number(pickCellFromSchema(row, detectedHeaders, "daysAbsent"));
    const daysAbsent = Number.isFinite(daysAbsentRaw) ? Math.min(30, Math.max(0, daysAbsentRaw)) : 0;
    const comment = String(pickCellFromSchema(row, detectedHeaders, "comment") || "").trim();
    if (!String(rawMonthValue || "").trim()) {
      defaultedMonthCount += 1;
    }

    if (designation) designationSet.add(designation);
    monthSet.add(month);

    payrollRecords.push({
      month,
      employeeId,
      employeeName,
      designation,
      presentSalary,
      increment,
      oldAdvanceTaken,
      extraAdvanceAdded,
      deductionEntered,
      daysAbsent,
      comment,
    });

    if (!employeeMap.has(employeeId)) {
      employeeMap.set(employeeId, {
        employeeId,
        employeeName,
        joiningDate: parseIsoFromImport(pickCellFromSchema(row, detectedHeaders, "joiningDate")) || new Date().toISOString().slice(0, 10),
        birthDate: parseIsoFromImport(pickCellFromSchema(row, detectedHeaders, "birthDate")),
        baseSalary: presentSalary,
        openingAdvance: oldAdvanceTaken,
        designation: designation || "Staff",
        mobileNumber: String(pickCellFromSchema(row, detectedHeaders, "mobileNumber") || "").trim(),
        status: "working",
      });
    }
  }

  const reasons = [];
  if (!detectedHeaders.employeeId) {
    reasons.push("Could not detect an Employee ID column.");
  }
  if (!detectedHeaders.employeeName) {
    reasons.push("Could not detect an Employee Name column.");
  }
  if (rows.length > 0 && payrollRecords.length === 0 && skippedMissingEmployeeId > 0) {
    reasons.push(`Skipped ${skippedMissingEmployeeId} row(s) because Employee ID was empty or not recognized.`);
  }
  if (rows.length > 0 && payrollRecords.length === 0 && skippedMissingEmployeeName > 0) {
    reasons.push(`Skipped ${skippedMissingEmployeeName} row(s) because Employee Name was empty or not recognized.`);
  }
  if (!detectedHeaders.month) {
    reasons.push("No Month column was detected, so the importer would fall back to the currently selected month.");
  }

  const warnings = [];
  if (defaultedMonthCount > 0) {
    warnings.push(`${defaultedMonthCount} row(s) did not include a month, so the selected app month was used.`);
  }
  if (skippedMissingEmployeeId > 0) {
    warnings.push(`${skippedMissingEmployeeId} row(s) were skipped due to missing Employee ID.`);
  }
  if (skippedMissingEmployeeName > 0) {
    warnings.push(`${skippedMissingEmployeeName} row(s) were skipped due to missing Employee Name.`);
  }

  return {
    totalRows: rows.length,
    payrollRecords,
    designations: Array.from(designationSet),
    employees: Array.from(employeeMap.values()),
    months: Array.from(monthSet).sort(),
    diagnostics: {
      detectedHeaders,
      reasons,
      warnings,
      skippedMissingEmployeeId,
      skippedMissingEmployeeName,
      defaultedMonthCount,
      inferenceNotes,
    },
  };
}

async function importLegacyAnalysis(analyzed, onProgress = () => {}) {
  const companyId = getSelectedCompanyId();
  const totalSteps = Math.max(1, analyzed.designations.length + analyzed.employees.length + analyzed.months.length);
  let completedSteps = 0;

  function reportProgress(label) {
    completedSteps += 1;
    const percent = Math.min(100, Math.max(1, Math.round((completedSteps / totalSteps) * 100)));
    onProgress({ percent, label });
  }

  for (const name of analyzed.designations) {
    if (!name) continue;
    try {
      // eslint-disable-next-line no-await-in-loop
      await apiRequest("/api/settings/designations", {
        method: "POST",
        body: { companyId, name },
      });
    } catch (error) {
      const msg = String(error?.message || "").toLowerCase();
      if (!msg.includes("already exists")) throw error;
    }
    reportProgress(`designation: ${name}`);
  }

  const existingEmployeesResp = await apiRequest(`/api/employees?companyId=${companyId}`);
  const existingEmployees = Array.isArray(existingEmployeesResp.employees) ? existingEmployeesResp.employees : [];
  const byEmployeeId = new Map(existingEmployees.map((emp) => [String(emp.employeeId || ""), emp]));

  for (const incoming of analyzed.employees) {
    const existing = byEmployeeId.get(String(incoming.employeeId || ""));
    const payload = {
      companyId,
      employeeId: incoming.employeeId,
      employeeName: incoming.employeeName,
      joiningDate: incoming.joiningDate,
      birthDate: incoming.birthDate,
      baseSalary: incoming.baseSalary,
      openingAdvance: incoming.openingAdvance,
      designation: incoming.designation,
      mobileNumber: incoming.mobileNumber,
      status: incoming.status,
      leaveFrom: "",
      leaveTo: "",
      terminatedOn: "",
      notes: "",
    };

    if (existing) {
      // eslint-disable-next-line no-await-in-loop
      await apiRequest(`/api/employees/${existing.id}`, { method: "PUT", body: payload });
    } else {
      // eslint-disable-next-line no-await-in-loop
      await apiRequest("/api/employees", { method: "POST", body: payload });
    }
    reportProgress(`employee: ${incoming.employeeName}`);
  }

  // Sort months chronologically
  const sortedMonths = Array.from(new Set(analyzed.payrollRecords.map(r => r.month))).sort();
  // Map: employeeId -> last advance remained
  const lastAdvanceByEmployee = new Map();
  // Group records by month and employee
  const recordsByMonth = new Map();
  for (const row of analyzed.payrollRecords) {
    const month = row.month || getSelectedMonth();
    if (!recordsByMonth.has(month)) recordsByMonth.set(month, new Map());
    recordsByMonth.get(month).set(row.employeeId, row);
  }

  for (const month of sortedMonths) {
    const incomingMap = recordsByMonth.get(month) || new Map();
    // eslint-disable-next-line no-await-in-loop
    const existingResp = await apiRequest(`/api/payroll/${month}?companyId=${companyId}`);
    const existingRecords = Array.isArray(existingResp.records) ? existingResp.records : [];
    const map = new Map(existingRecords.map((item) => [String(item.employeeId || ""), item]));

    // For each employee in this month
    for (const [employeeId, row] of incomingMap.entries()) {
      const existing = map.get(String(employeeId));
      // Always use the imported value for this month if present
      let oldAdvanceTaken = row.oldAdvanceTaken;
      // If missing or blank, carry forward from previous month
      if ((oldAdvanceTaken === undefined || oldAdvanceTaken === null || oldAdvanceTaken === "") && lastAdvanceByEmployee.has(employeeId)) {
        oldAdvanceTaken = lastAdvanceByEmployee.get(employeeId);
      }
      const merged = {
        ...(existing || {}),
        employeeId: row.employeeId,
        employeeName: row.employeeName,
        designation: row.designation || existing?.designation || "",
        presentSalary: row.presentSalary,
        increment: row.increment,
        oldAdvanceTaken,
        extraAdvanceAdded: row.extraAdvanceAdded,
        deductionEntered: row.deductionEntered,
        daysAbsent: row.daysAbsent,
        comment: row.comment || existing?.comment || "",
      };
      // Calculate advance remained for next month
      const totalAdvance = Number(oldAdvanceTaken) + Number(row.extraAdvanceAdded);
      const deductionApplied = Math.min(Number(row.deductionEntered), totalAdvance);
      const advanceRemained = totalAdvance - deductionApplied;
      lastAdvanceByEmployee.set(employeeId, advanceRemained);
      map.set(String(employeeId), merged);
    }

    const finalRecords = Array.from(map.values()).map((item, index) => ({
      ...item,
      positionIndex: Number(item.positionIndex ?? index),
    }));
    // eslint-disable-next-line no-await-in-loop
    await apiRequest(`/api/payroll/${month}?companyId=${companyId}`, {
      method: "PUT",
      body: { records: finalRecords },
    });
    reportProgress(`month: ${month}`);
  }
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

  let response;
  try {
    response = await fetch(url, {
      method: options.method || "GET",
      headers,
      body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
    });
  } catch {
    startServerReconnectPolling();
    throw new Error("Server is not reachable. Start backend with `npm start` or double-click `start-routes-payroll.command`.");
  }

  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json")
    ? await response.json()
    : null;

  stopServerReconnectPolling();

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
    currentUser = null;
    hydrateSettingsAccount();
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
  if (
    text.includes("saved")
    || text.includes("restored")
    || text.includes("exported")
    || text.includes("downloaded")
    || text.includes("added")
    || text.includes("removed")
    || text.includes("copied")
    || text.includes("updated")
    || text.includes("changed")
  ) {
    element.classList.add("status-success");
    return;
  }
  element.classList.add("status-info");
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  navigator.serviceWorker.getRegistrations()
    .then((registrations) => Promise.all(registrations.map((registration) => registration.unregister())))
    .catch(() => {
      // Non-critical if unregister fails.
    });
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
