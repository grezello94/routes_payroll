const STORAGE_KEYS = {
  authToken: "routes_payroll_auth_token_v1",
  payrollMonthMode: "routes_payroll_month_mode_v1",
  errorLogs: "routes_payroll_error_logs_v1",
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
  dashboardSummaryCards: document.getElementById("dashboardSummaryCards"),
  dashboardHighlightsBody: document.getElementById("dashboardHighlightsBody"),
  dashboardActivityBody: document.getElementById("dashboardActivityBody"),
  metricEmployees: document.getElementById("metricEmployees"),
  metricGross: document.getElementById("metricGross"),
  metricNet: document.getElementById("metricNet"),
  metricAdvance: document.getElementById("metricAdvance"),
  dashboardInsightsSection: document.getElementById("dashboardInsightsSection"),
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
  backupStatusSummary: document.getElementById("backupStatusSummary"),
  refreshBackupStatusBtn: document.getElementById("refreshBackupStatusBtn"),
  downloadLatestBackupBtn: document.getElementById("downloadLatestBackupBtn"),
  backupRecentList: document.getElementById("backupRecentList"),
  changePasswordForm: document.getElementById("changePasswordForm"),
  newPasswordInput: document.getElementById("newPasswordInput"),
  confirmPasswordInput: document.getElementById("confirmPasswordInput"),
  changePasswordBtn: document.getElementById("changePasswordBtn"),
  settingsMessage: document.getElementById("settingsMessage"),
  leaveResumeReportBody: document.getElementById("leaveResumeReportBody"),
  reportsMessage: document.getElementById("reportsMessage"),
  reportsSummaryCards: document.getElementById("reportsSummaryCards"),
  errorLogBody: document.getElementById("errorLogBody"),
  errorLogCount: document.getElementById("errorLogCount"),
  clearErrorLogsBtn: document.getElementById("clearErrorLogsBtn"),
  generatedPayrollBody: document.getElementById("generatedPayrollBody"),
  generatedPayrollDetailTitle: document.getElementById("generatedPayrollDetailTitle"),
  generatedPayrollEmployeesBody: document.getElementById("generatedPayrollEmployeesBody"),
  payslipDialog: document.getElementById("payslipDialog"),
  payslipPreview: document.getElementById("payslipPreview"),
  closePayslipBtn: document.getElementById("closePayslipBtn"),
  printPayslipBtn: document.getElementById("printPayslipBtn"),
  downloadPayslipBtn: document.getElementById("downloadPayslipBtn"),
  shareWebBtn: document.getElementById("shareWebBtn"),
  shareWhatsappBtn: document.getElementById("shareWhatsappBtn"),
  shareMessengerBtn: document.getElementById("shareMessengerBtn"),
  copyPayslipBtn: document.getElementById("copyPayslipBtn"),
  payrollWorkflowMessage: document.getElementById("payrollWorkflowMessage"),
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
let payrollReports = [];
let activePayrollReportId = null;
let activePayrollReportSnapshot = null;
let errorLogs = readStoredErrorLogs();
let backupStatus = null;

const MOTIVATING_MESSAGES = [
  "Great job this month! Keep up the excellent work.",
  "Your hard work and dedication are truly appreciated.",
  "Thank you for your continuous effort and positive energy.",
  "We value your commitment to the team. Keep it up!",
  "Outstanding performance! Proud to have you with us.",
  "Your contributions make a big difference. Thank you!",
  "Keep shining! Your dedication does not go unnoticed.",
  "Thank you for your outstanding teamwork and reliability.",
  "Your positive attitude is contagious. Thank you!",
  "We appreciate all the extra effort you put in this month."
];

init();

async function init() {
  wireAuth();
  wirePasswordToggles();
  wireEmployeeManagement();
  wireAppActions();
  wireSettingsActions();
  wireGlobalErrorLogging();
  setDefaultMonth();
  renderErrorLogs();
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

      if (response.token) {
        setToken(response.token);
        currentUser = response.user || null;
        hydrateSettingsAccount();
        needsSetupFlow = false;
        SELECTORS.setupForm.reset();
        showAuthMessage("");
        await switchToApp();
        showAppMessage(response.message || "Account created.");
      } else {
        needsSetupFlow = false;
        SELECTORS.setupForm.reset();
        setAuthMode("login");
        showAuthMessage(response.message || "Account created. Please check your email to verify before logging in.");
      }
    } catch (error) {
      showAuthMessage(error.message);
    } finally {
      if (submitBtn) submitBtn.disabled = false;
    }
  });

  SELECTORS.loginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const submitBtn = SELECTORS.loginForm.querySelector('button[type="submit"]');
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
    renderDashboardInsights();
  } catch (error) {
    employeeMaster = [];
    window.allPayrollRecords = [];
    renderEmployeeTable();
    renderLeaveResumeReport();
    renderDesignationSuggestions();
    renderDashboardInsights();
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
  const activeMonth = getSelectedMonth();
  if (window.allPayrollRecords) {
    for (const rec of window.allPayrollRecords) {
      if (!rec.employeeId) continue;
      if (companyId && Number(rec.companyId || rec.company_id || 1) !== Number(companyId)) continue;
      // Strictly ignore any future months beyond the currently selected month dropdown
      if (String(rec.month || "") > String(activeMonth)) continue;

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
      let salaryDisplay = Number(employee.baseSalary || 0);
      const payroll = payrollByEmployee[employee.employeeId];
      if (payroll) {
        const latestPresentSalary = Number(
          payroll.presentSalary ?? payroll.present_salary ?? payroll.baseSalary ?? payroll.base_salary
        );
        if (Number.isFinite(latestPresentSalary) && latestPresentSalary > 0) {
          salaryDisplay = latestPresentSalary;
        }
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
          <td>${formatCurrency(salaryDisplay)}</td>
          <td><span class="${advanceRemainedDisplay > 0 ? "advance-remained-alert" : ""}">${formatCurrency(advanceRemainedDisplay)}</span></td>
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
      const badgeClass = status === "leave"
        ? "status-pill leave"
        : status === "terminated"
          ? "status-pill terminated"
          : "status-pill working";

      return `
        <tr>
          <td>${escapeHtml(employee.employeeId || "-")}</td>
          <td>
            <div class="emp-list-name">${escapeHtml(employee.employeeName || "-")}</div>
            <div class="emp-list-sub">${escapeHtml(employee.designation || "-")}</div>
          </td>
          <td><span class="${badgeClass}">${escapeHtml(statusLabel)}</span></td>
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

function setPayrollWorkflowMessage(message) {
  applyStatusTone(SELECTORS.payrollWorkflowMessage, message);
  if (SELECTORS.payrollWorkflowMessage) {
    SELECTORS.payrollWorkflowMessage.textContent = message || "";
  }
}

function readStoredErrorLogs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.errorLogs);
    const parsed = JSON.parse(String(raw || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistErrorLogs() {
  try {
    localStorage.setItem(STORAGE_KEYS.errorLogs, JSON.stringify(errorLogs.slice(0, 200)));
  } catch {
    // Ignore storage quota issues for logging.
  }
}

function formatLogDetails(context) {
  if (!context || typeof context !== "object") return "";
  const fields = [];
  if (context.status) fields.push(`Status ${context.status}`);
  if (context.url) fields.push(String(context.url));
  if (context.action) fields.push(String(context.action));
  if (context.detail) fields.push(String(context.detail));
  return fields.join(" • ");
}

function recordAppError(message, source = "app", context = {}) {
  const text = String(message || "").trim();
  if (!text) return;

  const now = new Date().toISOString();
  const fingerprint = text;
  const latest = errorLogs[0];
  if (latest && latest.fingerprint === fingerprint) {
    const delta = Date.parse(now) - Date.parse(latest.createdAt || 0);
    if (Number.isFinite(delta) && delta >= 0 && delta < 8000) {
      latest.createdAt = now;
      latest.count = Number(latest.count || 1) + 1;
      latest.source = source || latest.source || "app";
      latest.context = { ...(latest.context || {}), ...(context || {}) };
      persistErrorLogs();
      renderErrorLogs();
      return;
    }
  }

  errorLogs.unshift({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    fingerprint,
    createdAt: now,
    source,
    message: text,
    context,
    count: 1,
  });
  errorLogs = errorLogs.slice(0, 200);
  persistErrorLogs();
  renderErrorLogs();
}

function renderErrorLogs() {
  if (!SELECTORS.errorLogBody || !SELECTORS.errorLogCount) return;
  const entries = Array.isArray(errorLogs) ? errorLogs : [];
  SELECTORS.errorLogCount.textContent = `${entries.length} ${entries.length === 1 ? "entry" : "entries"}`;

  if (!entries.length) {
    SELECTORS.errorLogBody.innerHTML = `
      <tr><td colspan="4" class="empty">No errors logged yet. App and API issues will appear here automatically.</td></tr>
    `;
    return;
  }

  SELECTORS.errorLogBody.innerHTML = entries
    .map((entry) => `
      <tr>
        <td>
          <div class="error-log-time">${escapeHtml(formatDateTime(entry.createdAt || ""))}</div>
          ${Number(entry.count || 1) > 1 ? `<div class="error-log-repeat">${escapeHtml(`${entry.count} times`)}</div>` : ""}
        </td>
        <td><span class="error-log-source">${escapeHtml(String(entry.source || "app"))}</span></td>
        <td><div class="error-log-message">${escapeHtml(String(entry.message || ""))}</div></td>
        <td><div class="error-log-detail">${escapeHtml(formatLogDetails(entry.context)) || "-"}</div></td>
      </tr>
    `)
    .join("");
}

function formatBackupScheduleDate(value) {
  const raw = String(value || "").trim();
  if (!raw) return "-";
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  return date.toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
}

function renderBackupStatus() {
  if (!SELECTORS.backupStatusSummary || !SELECTORS.backupRecentList) return;
  const status = backupStatus;
  if (!status) {
    SELECTORS.backupStatusSummary.textContent = "Monthly backup status is not loaded yet.";
    SELECTORS.backupRecentList.innerHTML = '<p class="empty">No backup status loaded.</p>';
    return;
  }

  const latest = status.latestBackup || null;
  const latestText = latest
    ? `Latest backup: ${formatDateTime(latest.createdAt)} (${latest.month}).`
    : "No monthly backup file has been created yet.";
  const scheduleText = status.isLastDayOfMonth
    ? (status.createdThisCheck ? "Auto backup ran during this check." : status.hasCurrentMonthBackup ? "This month's backup is already complete." : "Backup is due today and will run when refresh is triggered.")
    : `Next scheduled backup date: ${formatBackupScheduleDate(status.nextScheduledDate)}.`;
  SELECTORS.backupStatusSummary.textContent = `${latestText} ${scheduleText}`;

  const recent = Array.isArray(status.recentBackups) ? status.recentBackups : [];
  if (!recent.length) {
    SELECTORS.backupRecentList.innerHTML = '<p class="empty">No monthly backup files yet.</p>';
    return;
  }

  SELECTORS.backupRecentList.innerHTML = recent
    .map((item) => `
      <div class="designation-item">
        <span>${escapeHtml(`${item.month} • ${formatDateTime(item.createdAt)}`)}</span>
        <small>${escapeHtml(String(item.fileName || ""))}</small>
      </div>
    `)
    .join("");
}

async function loadBackupStatus() {
  try {
    const response = await apiRequest("/api/system-backups/status");
    backupStatus = response || null;
    renderBackupStatus();
  } catch (error) {
    backupStatus = null;
    renderBackupStatus();
    recordAppError(error.message, "backup");
  }
}

async function downloadLatestEmergencyBackup() {
  try {
    const response = await fetch("/api/system-backups/latest", {
      headers: authToken ? { Authorization: `Bearer ${authToken}` } : {},
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload?.error || `Request failed (${response.status}).`);
    }
    const disposition = response.headers.get("content-disposition") || "";
    const match = disposition.match(/filename=\"?([^"]+)\"?/i);
    const filename = match?.[1] || `routes-payroll-backup-${new Date().toISOString().slice(0, 10)}.json`;
    const blob = await response.blob();
    downloadBlob(blob, filename);
    showAppMessage("Monthly emergency backup downloaded.");
  } catch (error) {
    showAppMessage(error.message || "Failed to download latest backup.");
  }
}

function clearErrorLogs() {
  errorLogs = [];
  try {
    localStorage.removeItem(STORAGE_KEYS.errorLogs);
  } catch {
    // Ignore storage cleanup failures.
  }
  renderErrorLogs();
}

function isErrorLikeMessage(message) {
  const text = String(message || "").toLowerCase();
  if (!text) return false;
  return text.includes("failed")
    || text.includes("error")
    || text.includes("invalid")
    || text.includes("not found")
    || text.includes("not reachable")
    || text.includes("blocked")
    || text.includes("timed out")
    || text.includes("quota")
    || text.includes("denied");
}

function wireGlobalErrorLogging() {
  window.addEventListener("error", (event) => {
    recordAppError(event?.message || "Unexpected browser error.", "browser", {
      detail: event?.filename ? `${event.filename}:${event.lineno || 0}` : "",
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    const reason = event?.reason;
    const message = reason?.message || String(reason || "Unhandled promise rejection.");
    recordAppError(message, "browser", {
      detail: "Unhandled promise rejection",
    });
  });
}

function renderDashboardInsights() {
  if (SELECTORS.dashboardSummaryCards) {
    const currentMonth = getSelectedMonth();
    const currentMonthLabel = formatMonth(currentMonth);
        const activeHeadcount = employeeMaster.filter((employee) => isPayrollActiveStatus(employee.status)).length;
    const onLeave = employeeMaster.filter((employee) => String(employee.status || "").toLowerCase() === "leave").length;
    const terminated = employeeMaster.filter((employee) => String(employee.status || "").toLowerCase() === "terminated").length;
    const generatedCurrentMonth = payrollReports.find((report) => String(report.month) === String(currentMonth)) || null;
    const latestReport = payrollReports[0] || null;

    const summaryCards = [
      {
        label: "Current Payroll Month",
        value: currentMonthLabel || "-",
        note: generatedCurrentMonth?.generatedAt
          ? `Generated ${formatDateTime(generatedCurrentMonth.generatedAt)}`
          : "No saved payroll report yet",
        tone: "blue",
      },
      {
        label: "Saved Payroll Months",
        value: String(payrollReports.length),
        note: latestReport ? `Latest: ${formatMonth(latestReport.month)}` : "No reports generated",
        tone: "violet",
      },
      {
        label: "Employees On Leave",
        value: String(onLeave),
        note: `${terminated} terminated record(s) on file`,
        tone: "amber",
      },
      {
            label: "Active In Company",
            value: String(activeHeadcount),
        note: `${employeeMaster.length} total employee profile(s)`,
        tone: "green",
      },
    ];

    SELECTORS.dashboardSummaryCards.innerHTML = summaryCards.map((card) => `
      <article class="dashboard-stat-card tone-${card.tone}">
        <p>${escapeHtml(card.label)}</p>
        <strong>${escapeHtml(card.value)}</strong>
        <span>${escapeHtml(card.note)}</span>
      </article>
    `).join("");
  }

  if (SELECTORS.dashboardHighlightsBody) {
    const currentMonth = getSelectedMonth();
    const report = payrollReports.find((item) => String(item.month) === String(currentMonth)) || null;
    const computed = currentRecords.map((record) => computePayroll(record, currentMonth));
    const net = computed.reduce((sum, item) => sum + item.netSalary, 0);
    const gross = computed.reduce((sum, item) => sum + item.grossSalary, 0);
    const advance = computed.reduce((sum, item) => sum + item.advanceRemained, 0);
    const blockedCount = currentRecords.filter((record) => isPayrollBlocked(record, currentMonth)).length;

    const highlights = [
      {
        title: "Payroll Status",
        detail: report?.generatedAt
          ? `${report.employeeCount || 0} payslip(s) saved for ${formatMonth(currentMonth)}`
          : `No generated payroll saved yet for ${formatMonth(currentMonth)}`,
      },
      {
        title: "Gross vs Net",
        detail: `${formatCurrency(gross)} gross payroll and ${formatCurrency(net)} net payout`,
      },
      {
        title: "Advance Exposure",
        detail: `${formatCurrency(advance)} remains pending as employee advances`,
      },
      {
        title: "Locked or Inactive Records",
        detail: `${blockedCount} record(s) are currently blocked from payroll payout`,
      },
    ];

    SELECTORS.dashboardHighlightsBody.innerHTML = highlights.map((item) => `
      <div class="dashboard-list-item">
        <strong>${escapeHtml(item.title)}</strong>
        <span>${escapeHtml(item.detail)}</span>
      </div>
    `).join("");
  }

  if (SELECTORS.dashboardActivityBody) {
    const items = payrollReports.slice(0, 4);
    SELECTORS.dashboardActivityBody.innerHTML = items.length
      ? items.map((report, index) => `
        <div class="dashboard-list-item">
          <strong>${escapeHtml(`${index === 0 ? "Latest report" : "Saved report"} • ${formatMonth(report.month)}`)}</strong>
          <span>${escapeHtml(`${report.employeeCount || 0} employee payslip(s) • checked ${formatDateTime(report.checkedAt) || "-"} • generated ${formatDateTime(report.generatedAt) || "-"}`)}</span>
        </div>
      `).join("")
      : `
        <div class="dashboard-list-item empty-dashboard-item">
          <strong>No payroll reports generated yet</strong>
          <span>Generate employee payslips from Payroll to build a month-wise report timeline here.</span>
        </div>
      `;
  }
}

function renderReportsSummary() {
  if (!SELECTORS.reportsSummaryCards) return;
  const openReport = payrollReports.find((report) => Number(report.id) === Number(activePayrollReportId)) || null;
  const totalPayslips = payrollReports.reduce((sum, report) => sum + Number(report.employeeCount || 0), 0);
  const latestReport = payrollReports[0] || null;
  const summaryCards = [
    {
      label: "Generated Months",
      value: String(payrollReports.length),
      note: latestReport ? `Latest month: ${formatMonth(latestReport.month)}` : "No month generated yet",
      tone: "blue",
    },
    {
      label: "Saved Payslips",
      value: String(totalPayslips),
      note: "Employee payslips captured across generated reports",
      tone: "green",
    },
    {
      label: "Current Open Report",
      value: openReport ? formatMonth(openReport.month) : "None",
      note: openReport ? `${openReport.employeeCount || 0} employee(s) in active view` : "Open a generated payroll month below",
      tone: "violet",
    },
    {
      label: "Last Generated",
      value: latestReport?.generatedAt ? formatDateTime(latestReport.generatedAt) : "-",
      note: latestReport ? "Most recent payroll snapshot" : "Waiting for first generated report",
      tone: "amber",
    },
  ];

  SELECTORS.reportsSummaryCards.innerHTML = summaryCards.map((card) => `
    <article class="dashboard-stat-card compact tone-${card.tone}">
      <p>${escapeHtml(card.label)}</p>
      <strong>${escapeHtml(card.value)}</strong>
      <span>${escapeHtml(card.note)}</span>
    </article>
  `).join("");
}

function getCurrentMonthPayrollReport() {
  const month = getSelectedMonth();
  return payrollReports.find((report) => String(report.month) === String(month)) || null;
}

function renderPayrollWorkflow() {
  const report = getCurrentMonthPayrollReport();
  const monthLabel = formatMonth(getSelectedMonth());
  const parts = [];
  if (report?.generatedAt) {
    parts.push(`${report.employeeCount || 0} payslip(s) generated for ${monthLabel}`);
    parts.push(`Last generated on ${formatDateTime(report.generatedAt)}`);
  }
  if (!parts.length) {
    parts.push(`Generate payslips one by one for ${monthLabel} from the employee payroll register.`);
  }

  setPayrollWorkflowMessage(parts.join(" • "));
}

async function generatePayslipForCurrentRecord(index) {
  const record = currentRecords[index];
  if (!record) return;

  try {
    await flushPendingSave();
    const month = getSelectedMonth();
    const response = await apiRequest("/api/payroll-reports/generate-entry", {
      method: "POST",
      body: {
        companyId: getSelectedCompanyId(),
        month,
        employeeId: record.employeeId,
      },
    });
    activePayrollReportId = Number(response.report?.id || activePayrollReportId || 0) || null;
    activePayrollReportSnapshot = response.snapshot || null;
    await loadPayrollReports(activePayrollReportId, { openPreferredReport: false });
    renderPayrollWorkflow();
    showAppMessage(`Payslip generated for ${record.employeeName || record.employeeId}.`);
  } catch (error) {
    showAppMessage(error.message);
  }
}

function renderGeneratedPayrollReports() {
  if (!SELECTORS.generatedPayrollBody || !SELECTORS.generatedPayrollEmployeesBody || !SELECTORS.generatedPayrollDetailTitle) return;

  if (!payrollReports.length) {
    SELECTORS.generatedPayrollBody.innerHTML = `
      <tr><td colspan="5" class="empty">No payroll has been generated yet for this company.</td></tr>
    `;
    SELECTORS.generatedPayrollEmployeesBody.innerHTML = `
      <tr><td colspan="5" class="empty">Open a generated payroll month to see employee payslips.</td></tr>
    `;
    SELECTORS.generatedPayrollDetailTitle.textContent = "Generated Payslips";
    renderPayrollWorkflow();
    return;
  }

  SELECTORS.generatedPayrollBody.innerHTML = payrollReports
    .map((report) => {
      const activeClass = Number(report.id) === Number(activePayrollReportId) ? "active-report-row" : "";
      const isActive = Number(report.id) === Number(activePayrollReportId);
      return `
        <tr class="${activeClass}">
          <td>
            <div class="report-month-title">Payroll Generated for the month of ${escapeHtml(formatMonth(report.month))}</div>
            <div class="report-month-sub">Saved monthly payroll snapshot</div>
          </td>
          <td>${escapeHtml(formatDateTime(report.checkedAt) || "-")}</td>
          <td>${escapeHtml(formatDateTime(report.generatedAt) || "-")}</td>
          <td><span class="report-count-pill">${escapeHtml(String(report.employeeCount || 0))}</span></td>
          <td>
            <div class="row-actions">
              <button type="button" class="mini ghost" data-report-id="${report.id}" data-report-action="open" aria-pressed="${isActive ? "true" : "false"}">${isActive ? "Opened" : "Open"}</button>
              <button type="button" class="mini ghost" data-report-id="${report.id}" data-report-action="excel">Excel</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");

  const snapshot = activePayrollReportSnapshot;
  if (!snapshot || !Array.isArray(snapshot.records) || snapshot.records.length === 0) {
    SELECTORS.generatedPayrollEmployeesBody.innerHTML = `
      <tr><td colspan="5" class="empty">Open a generated payroll month to see employee payslips.</td></tr>
    `;
    SELECTORS.generatedPayrollDetailTitle.textContent = "Generated Payslips";
    renderPayrollWorkflow();
    return;
  }

  SELECTORS.generatedPayrollDetailTitle.textContent = `Generated Payslips for ${formatMonth(snapshot.month)}`;
  SELECTORS.generatedPayrollEmployeesBody.innerHTML = snapshot.records
    .map((record, index) => {
      const calc = computePayroll(record, snapshot.month);
      return `
        <tr>
          <td>${escapeHtml(record.employeeId || "-")}</td>
          <td><div class="emp-list-name">${escapeHtml(record.employeeName || "-")}</div></td>
          <td>${escapeHtml(record.designation || "-")}</td>
          <td><span class="report-net-salary">${escapeHtml(formatCurrency(calc.netSalary))}</span></td>
          <td>
            <div class="row-actions">
              <button type="button" class="mini ghost icon-mini" title="View Payslip" aria-label="View Payslip" data-report-index="${index}" data-action="preview">&#128065;</button>
              <button type="button" class="mini ghost icon-mini" title="Save Payslip PDF" aria-label="Save Payslip PDF" data-report-index="${index}" data-action="download">&#8681;</button>
              <button type="button" class="mini icon-mini" title="Print or Save PDF" aria-label="Print or Save PDF" data-report-index="${index}" data-action="print">&#128424;</button>
            </div>
          </td>
        </tr>
      `;
    })
    .join("");

  renderPayrollWorkflow();
}

function setEmployeeMessage(message) {
  if (!SELECTORS.employeeMessage) return;
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

  window.addEventListener("beforeunload", (event) => {
    if (saveInFlight || pendingSaveTimer || saveQueued) {
      event.preventDefault();
      event.returnValue = "Changes are still saving to the cloud. Please wait a moment before leaving.";
    }
  });

  SELECTORS.clearErrorLogsBtn?.addEventListener("click", () => {
    clearErrorLogs();
    showAppMessage("Error logs cleared.");
  });

  SELECTORS.monthPicker.addEventListener("change", async () => {
    await flushPendingSave();
    activePayrollReportId = null;
    activePayrollReportSnapshot = null;
    await loadMonthRecords();
    await loadPayrollReports();
    renderEmployeeTable();
  });

  SELECTORS.addEmployeeBtn.addEventListener("click", async () => {
    setWorkspace("employees");
    syncCurrentRecordsToAllPayroll();
    renderEmployeeTable();
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
    payrollReports = [];
    activePayrollReportId = null;
    activePayrollReportSnapshot = null;
    pendingSettingsLogoDataUrl = "";
    renderEmployeeTable();
    renderDesignationPresets();
    renderDesignationSuggestions();
    renderGeneratedPayrollReports();
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
      generatePayslipForCurrentRecord(index);
    }
  });

  SELECTORS.closePayslipBtn.addEventListener("click", () => SELECTORS.payslipDialog.close());
  SELECTORS.printPayslipBtn.addEventListener("click", printCurrentPayslip);
  SELECTORS.downloadPayslipBtn.addEventListener("click", downloadCurrentPayslipPdf);
  SELECTORS.shareWhatsappBtn.addEventListener("click", shareCurrentPayslipWhatsapp);
  SELECTORS.shareMessengerBtn.addEventListener("click", shareCurrentPayslipMessenger);
  SELECTORS.shareWebBtn.addEventListener("click", shareCurrentPayslipWeb);
  SELECTORS.copyPayslipBtn.addEventListener("click", copyCurrentPayslip);

  SELECTORS.generatedPayrollBody?.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button[data-report-id]");
    if (!button) return;
    const reportId = Number(button.dataset.reportId);
    if (!Number.isInteger(reportId) || reportId <= 0) return;
    try {
      if (button.dataset.reportAction === "excel") {
        await downloadPayrollReportExcel(reportId);
        return;
      }
      await openPayrollReport(reportId);
      SELECTORS.generatedPayrollDetailTitle?.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      showAppMessage(error.message);
    }
  });

  SELECTORS.generatedPayrollEmployeesBody?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const button = target.closest("button[data-report-index]");
    if (!button || !activePayrollReportSnapshot) return;
    const index = Number(button.dataset.reportIndex);
    if (Number.isNaN(index)) return;
    const record = activePayrollReportSnapshot.records?.[index];
    if (!record) return;
    openPayslip(record, activePayrollReportSnapshot.month, activePayrollReportSnapshot.company || null);
    if (button.dataset.action === "print") {
      printCurrentPayslip();
      return;
    }
    if (button.dataset.action === "download") {
      downloadCurrentPayslipPdf();
    }
  });
}

function wireSettingsActions() {
  SELECTORS.refreshBackupStatusBtn?.addEventListener("click", async () => {
    await loadBackupStatus();
    showAppMessage("Monthly backup status refreshed.");
  });

  SELECTORS.downloadLatestBackupBtn?.addEventListener("click", async () => {
    await downloadLatestEmergencyBackup();
  });

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
      const monthPreview = buildLegacyImportPreview(analyzed);
      if (analyzed.payrollRecords.length === 0) {
        const reasonText = analyzed.diagnostics?.reasons?.length
          ? analyzed.diagnostics.reasons.join(" ")
          : "No payroll rows matched expected columns.";
        setImportSummary(`Import failed: ${reasonText}${inferenceText}${monthPreview ? `\n${monthPreview}` : ""}`);
        setSettingsMessage(`Import failed: ${reasonText}`);
        return;
      }

      if (btn) btn.textContent = "Importing... 1%";
      setImportSummary(`Rows analyzed: ${analyzed.totalRows}. Importing... 1%${monthPreview ? `\n${monthPreview}` : ""}`);
      await importLegacyAnalysis(analyzed, ({ percent, label }) => {
        const safePercent = Math.max(1, Math.min(100, Number(percent) || 1));
        if (btn) btn.textContent = `Importing... ${safePercent}%`;
        setImportSummary(`Rows analyzed: ${analyzed.totalRows}. Importing... ${safePercent}%${label ? ` (${label})` : ""}${monthPreview ? `\n${monthPreview}` : ""}`);
      });

      const latestImportedMonth = analyzed.months?.length ? analyzed.months[analyzed.months.length - 1] : "";
      if (latestImportedMonth && SELECTORS.monthPicker) {
        SELECTORS.monthPicker.value = latestImportedMonth;
      }

      await Promise.all([
        loadDesignationPresets(),
        loadEmployees(),
        loadMonthRecords(),
      ]);
      if (btn) btn.textContent = "Importing... 100%";
      const warningText = analyzed.diagnostics?.warnings?.length
        ? ` Warnings: ${analyzed.diagnostics.warnings.join(" ")}`
        : "";
      setImportSummary(`Import successful! 100%. Imported ${analyzed.payrollRecords.length} payroll rows across ${analyzed.months.length} month(s).${warningText}${inferenceText}${monthPreview ? `\n${monthPreview}` : ""}`);
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
    activePayrollReportId = null;
    activePayrollReportSnapshot = null;
    await loadEmployees();
    await loadMonthRecords();
    await loadDesignationPresets();
    await loadPayrollReports();
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
      await loadPayrollReports();
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
        syncCurrentRecordsToAllPayroll();
        renderEmployeeTable();
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
        loadBackupStatus();
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
  SELECTORS.dashboardInsightsSection?.classList.toggle("hidden", !dashboardView);
  SELECTORS.employeesSection?.classList.toggle("hidden", !employeesView);
  SELECTORS.payrollSection?.classList.toggle("hidden", !payrollView);
  SELECTORS.reportsSection?.classList.toggle("hidden", !reportsView);
  SELECTORS.settingsSection?.classList.toggle("hidden", !settingsView);
  SELECTORS.saveStatus?.classList.toggle("hidden", !payrollView);
  
  localStorage.setItem("activeWorkspace", view);
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
    needsSetupFlow = Boolean(bootstrap.needsSetup);
    SELECTORS.showRegisterBtn?.classList.remove("hidden");
    SELECTORS.showLoginBtn?.classList.remove("hidden");

    if (bootstrap.pending) {
      SELECTORS.setupForm.classList.add("hidden");
      SELECTORS.loginForm.classList.add("hidden");
      startServerReconnectPolling();
      showAuthMessage(bootstrap.message || "Connecting...");
      return;
    }

    if (bootstrap.degraded) {
      SELECTORS.setupForm.classList.add("hidden");
      SELECTORS.loginForm.classList.add("hidden");
      startServerReconnectPolling();
      setAuthMode("login");
      showAuthMessage(bootstrap.message || "Cloud service is temporarily busy. Please retry shortly.");
      return;
    }

    stopServerReconnectPolling();
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
    const msg = String(error.message).toLowerCase();
    if (msg.includes("expired") || msg.includes("authentication") || msg.includes("invalid") || !authToken) {
      setToken("");
      currentUser = null;
      hydrateSettingsAccount();
      return false;
    }
    return true;
  }
}

async function switchToApp() {
  stopServerReconnectPolling();
  SELECTORS.authView.classList.add("hidden");
  SELECTORS.appView.classList.remove("hidden");
  hydrateSettingsAccount();
  try {
    await loadCompanies();
    await Promise.all([
      loadEmployees(),
      loadMonthRecords(),
      loadDesignationPresets(),
      loadPayrollReports(),
      loadBackupStatus(),
    ]);
  } catch (error) {
    showAppMessage(error.message || "Failed to load some application data.");
  }
  
  const savedWorkspace = localStorage.getItem("activeWorkspace") || "dashboard";
  setWorkspace(savedWorkspace);
  SELECTORS.railButtons.forEach((item) => item.classList.toggle("active", item.dataset.railAction === savedWorkspace));
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
      if (reachable) await updateAuthView();
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

function isPayrollVisibleForMonth(record, month = getSelectedMonth()) {
  const status = String(record?.employeeStatus || "working").toLowerCase();
  if (status === "leave" || status === "terminated") return false;
  if (status !== "resumed") return true;

  const resumedOn = parseIsoDate(record?.leaveTo);
  if (!resumedOn) return false;
  const { end } = monthBounds(month);
  return resumedOn <= end;
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
  let incomingCompanies = [];
  try {
    const response = await apiRequest("/api/companies");
    incomingCompanies = Array.isArray(response.companies) ? response.companies : [];
  } catch (error) {
    showAppMessage(error.message || "Failed to load companies.");
  }
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
    currentRecords = (Array.isArray(response.records) ? response.records : []).map((record) => {
      const employee = employeeMaster.find((item) => String(item.employeeId) === String(record.employeeId));
      if (!employee) return record;

      const masterSalary = Number(employee.baseSalary || 0);
      const monthSalary = Number(record.presentSalary || 0);
      return {
        ...record,
        employeeName: employee.employeeName || record.employeeName || "",
        designation: employee.designation || record.designation || "",
        presentSalary: masterSalary > 0 ? masterSalary : monthSalary,
        employeeStatus: employee.status || record.employeeStatus || "working",
        leaveFrom: employee.leaveFrom || record.leaveFrom || "",
        leaveTo: employee.leaveTo || record.leaveTo || "",
        terminatedOn: employee.terminatedOn || record.terminatedOn || "",
      };
    });
    const visible = currentRecords.filter((record) => isPayrollVisibleForMonth(record, month));
    if (!visible.some((record) => String(record.employeeId) === String(activePayrollEmployeeId))) {
      activePayrollEmployeeId = visible[0]?.employeeId || "";
    }
    renderPayrollTable();
    setSaveStatus("All changes saved.");
    renderPayrollWorkflow();
    renderDashboardInsights();
  } catch (error) {
    currentRecords = [];
    activePayrollEmployeeId = "";
    renderPayrollTable();
    renderPayrollWorkflow();
    renderDashboardInsights();
    showAppMessage(error.message);
  }
}

function payrollReportsUrl() {
  return `/api/payroll-reports?companyId=${getSelectedCompanyId()}`;
}

async function loadPayrollReports(preferredReportId = activePayrollReportId, options = {}) {
  const { openPreferredReport = true } = options;
  try {
    const response = await apiRequest(payrollReportsUrl());
    payrollReports = Array.isArray(response.reports) ? response.reports : [];
    const selectedReport = payrollReports.find((report) => Number(report.id) === Number(preferredReportId))
      || payrollReports.find((report) => String(report.month) === String(getSelectedMonth()))
      || payrollReports[0]
      || null;
    activePayrollReportId = selectedReport ? Number(selectedReport.id) : null;
    if (activePayrollReportId && openPreferredReport) {
      await openPayrollReport(activePayrollReportId, true);
    } else {
      renderGeneratedPayrollReports();
      renderReportsSummary();
      renderDashboardInsights();
    }
  } catch (error) {
    payrollReports = [];
    activePayrollReportId = null;
    activePayrollReportSnapshot = null;
    renderGeneratedPayrollReports();
    renderReportsSummary();
    renderDashboardInsights();
    showAppMessage(error.message);
  }
}

async function openPayrollReport(reportId, skipListRender = false) {
  const response = await apiRequest(`/api/payroll-reports/${reportId}?companyId=${getSelectedCompanyId()}`);
  activePayrollReportId = Number(response.report?.id || reportId);
  activePayrollReportSnapshot = response.snapshot || null;
  renderGeneratedPayrollReports();
  renderReportsSummary();
  renderDashboardInsights();
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
  const monthDays = end.getDate();
  if (resumedOn > end) return monthDays;
  return clamp(resumedOn.getDate() - 1, 0, monthDays);
}

function payrollMonthDayCount(month) {
  return monthBounds(month).end.getDate();
}

function daysWorkedFromDateInMonth(date, month) {
  if (!date) return 0;
  const { start, end } = monthBounds(month);
  if (date > end) return 0;
  const effectiveStart = date < start ? start : date;
  return clamp(end.getDate() - effectiveStart.getDate() + 1, 0, end.getDate());
}

function systemWorkedDaysInMonth(record, month) {
  const status = String(record.employeeStatus || "working").toLowerCase();
  if (status === "leave" || status === "terminated") return 0;

  const monthDays = payrollMonthDayCount(month);
  const joiningDate = parseIsoDate(record.joiningDate);
  const resumedOn = status === "resumed" ? parseIsoDate(record.leaveTo) : null;
  const candidateDates = [];

  if (joiningDate) candidateDates.push(joiningDate);
  if (resumedOn) candidateDates.push(resumedOn);

  if (candidateDates.length === 0) return monthDays;

  let workedDays = monthDays;
  for (const candidate of candidateDates) {
    workedDays = Math.min(workedDays, daysWorkedFromDateInMonth(candidate, month));
  }
  return clamp(workedDays, 0, monthDays);
}

function computePayroll(record, month = getSelectedMonth()) {
  const blocked = isPayrollBlocked(record, month);
  const monthDays = payrollMonthDayCount(month);
  const presentSalary = blocked ? 0 : toMoney(record.presentSalary);
  const increment = blocked ? 0 : toMoney(record.increment);
  const grossSalary = presentSalary + increment;

  const oldAdvanceTaken = toMoney(record.oldAdvanceTaken);
  const extraAdvanceAdded = toMoney(record.extraAdvanceAdded);
  const totalAdvance = oldAdvanceTaken + extraAdvanceAdded;

  const deductionEntered = blocked ? 0 : Math.max(0, toMoney(record.deductionEntered));
  const systemWorkedDays = blocked ? 0 : systemWorkedDaysInMonth(record, month);
  const autoDaysAbsent = blocked ? monthDays : Math.max(0, monthDays - systemWorkedDays);
  const manualDaysAbsent = blocked ? 0 : clamp(toMoney(record.daysAbsent), 0, systemWorkedDays);
  const workedDays = blocked ? 0 : clamp(systemWorkedDays - manualDaysAbsent, 0, monthDays);
  const daysAbsent = clamp(monthDays - workedDays, 0, monthDays);
  const payableSalary = monthDays > 0 ? (workedDays / monthDays) * grossSalary : 0;
  const proratedAbsenceDeduction = Math.max(0, grossSalary - payableSalary);
  const deductionApplied = Math.min(deductionEntered, totalAdvance);
  const advanceRemained = totalAdvance - deductionApplied;
  const netSalary = Math.max(0, payableSalary - deductionApplied);

  return {
    blocked,
    monthDays,
    presentSalary,
    increment,
    grossSalary,
    oldAdvanceTaken,
    extraAdvanceAdded,
    totalAdvance,
    deductionEntered,
    systemWorkedDays,
    autoDaysAbsent,
    manualDaysAbsent,
    workedDays,
    daysAbsent,
    payableSalary,
    proratedAbsenceDeduction,
    deductionApplied,
    advanceRemained,
    netSalary,
  };
}

function payrollZeroSalaryReason(record, calc, month) {
  if (calc.netSalary > 0) return "";
  if (calc.blocked) {
    return `Payroll is locked for ${formatMonth(month)}.`;
  }
  if (calc.workedDays <= 0) {
    const joiningDate = parseIsoDate(record.joiningDate);
    const resumedOn = String(record.employeeStatus || "").toLowerCase() === "resumed"
      ? parseIsoDate(record.leaveTo)
      : null;
    const { end } = monthBounds(month);

    if (joiningDate && joiningDate > end) {
      return `No salary is payable for ${formatMonth(month)} because joining date ${record.joiningDate} falls after this payroll month.`;
    }
    if (resumedOn && resumedOn > end) {
      return `No salary is payable for ${formatMonth(month)} because resume date ${record.leaveTo} falls after this payroll month.`;
    }
    if (calc.manualDaysAbsent >= calc.systemWorkedDays && calc.systemWorkedDays > 0) {
      return `No salary is payable because manual absent days removed all ${formatNumberValue(calc.systemWorkedDays)} payable day(s).`;
    }
    return `No salary is payable for ${formatMonth(month)} because the system counted 0 worked day(s). Check payroll month, joining date, resume date, or absences.`;
  }
  if (calc.proratedAbsenceDeduction >= calc.grossSalary) {
    return `No salary remains after full-month absence proration for ${formatMonth(month)}.`;
  }
  if (calc.deductionApplied >= calc.payableSalary && calc.deductionApplied > 0) {
    return `Advance deduction consumed the full payable salary for ${formatMonth(month)}.`;
  }
  return "";
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
    .filter((item) => isPayrollVisibleForMonth(item.record, month));
  const visibleCount = visibleRecords.length;

  if (visibleCount === 0) {
    SELECTORS.payrollBody.innerHTML = `
      <div class="empty">
        No active payroll employees for ${formatMonth(month)}. Employees on leave or terminated are hidden until they resume.
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
  const zeroSalaryReason = payrollZeroSalaryReason(record, calc, month);
  const status = String(record.employeeStatus || "working").toLowerCase();
  const statusLabel = status === "leave"
    ? "On Leave"
    : status === "terminated"
      ? "Terminated"
      : status === "resumed"
        ? "Resumed"
        : "Working";
  const badgeClass = status === "leave"
    ? "status-pill leave"
    : status === "terminated"
      ? "status-pill terminated"
      : "status-pill working";
  const statusBoxClass = status === "leave"
    ? "status-box leave"
    : status === "terminated"
      ? "status-box terminated"
      : "status-box active";
  const statusDetail = status === "leave"
    ? `Leave: ${record.leaveFrom || "-"} to ${record.leaveTo || "-"}`
    : status === "terminated"
      ? `Terminated on: ${record.terminatedOn || "-"}`
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
                <div class="${statusBoxClass}">${escapeHtml(statusLabel.toUpperCase())}</div>
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
              <label>Worked Days (out of ${calc.monthDays})
                <input type="text" class="locked-field" value="${formatNumberValue(calc.workedDays)}" readonly />
                <span class="field-note">System calculated from joining/resume date and manual absences.</span>
              </label>
              <label>Manual Days Absent (out of ${calc.systemWorkedDays})
                <input class="field-absent" data-field="daysAbsent" type="number" min="0" max="${calc.systemWorkedDays}" step="1" value="${toRaw(record.daysAbsent)}" ${allowOverride ? "" : "readonly"} />
                ${!allowOverride ? '<span class="field-note">Locked (past month)</span>' : ''}
              </label>
              <label>Salary For Worked Days (₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.payableSalary)}" readonly />
                <span class="field-note">Based on ${formatNumberValue(calc.workedDays)} of ${calc.monthDays} calendar day(s).</span>
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
            ${(calc.autoDaysAbsent > 0 || status === "resumed")
              ? `<p class="status-note">${escapeHtml(`System counted ${formatNumberValue(calc.autoDaysAbsent)} non-working day(s) for this month before joining/resume. Manual absence is added separately.`)}</p>`
              : ""}
          </section>

          <section class="payroll-stack-section">
            <h3>Summary</h3>
            <div class="payroll-fields">
              <label>Advance Remained (₹)
                <input type="text" class="locked-field ${calc.advanceRemained > 0 ? "advance-remained-input" : ""}" value="${formatCurrency(calc.advanceRemained)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
              <label>Salary In Hand (Net Salary ₹)
                <input type="text" class="locked-field" value="${formatCurrency(calc.netSalary)}" readonly />
                <span class="field-note">Locked (system calculated)</span>
              </label>
            </div>
            ${zeroSalaryReason ? `<p class="status-note">${escapeHtml(zeroSalaryReason)}</p>` : ""}
          </section>
        </div>

        <aside class="payroll-modern-side">
          <div class="kpi-card net">
            <p>Net Salary In Hand</p>
            <strong>${formatCurrency(calc.netSalary)}</strong>
            <div class="net-sub">Calculated: Worked Days Pay - Advance Deduction</div>
          </div>

          <div class="kpi-card soft">
            <h4>Deductions</h4>
            <div class="kpi-line"><span>Deduction Entered</span><b>${formatCurrency(calc.deductionEntered)}</b></div>
            <div class="kpi-line"><span>Absence Deduction</span><b>${formatCurrency(calc.proratedAbsenceDeduction)}</b></div>
            <div class="kpi-line total"><span>Advance Deduction Applied</span><b>${formatCurrency(calc.deductionApplied)}</b></div>
          </div>

          <div class="kpi-card amber">
            <p>Advance Remained</p>
            <strong class="${calc.advanceRemained > 0 ? "advance-remained-alert" : ""}">${formatCurrency(calc.advanceRemained)}</strong>
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

  const visibleRecordsForMetrics = currentRecords.filter((record) => isPayrollVisibleForMonth(record, month));
  updateMetrics(visibleRecordsForMetrics.map((record) => computePayroll(record, month)));
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
  renderDashboardInsights();
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

function syncCurrentRecordsToAllPayroll() {
  if (!window.allPayrollRecords) window.allPayrollRecords = [];
  const month = getSelectedMonth();
  const companyId = getSelectedCompanyId();
  
  for (const rec of currentRecords) {
    if (!rec.employeeId) continue;
    const calc = computePayroll(rec, month);
    const existingIdx = window.allPayrollRecords.findIndex(
      (r) => String(r.month) === String(month) && 
             String(r.employeeId) === String(rec.employeeId) && 
             Number(r.companyId || r.company_id || 1) === Number(companyId)
    );
    
    const updatedRec = {
      ...rec,
      companyId,
      month,
      advanceRemained: calc.advanceRemained,
      presentSalary: calc.presentSalary
    };
    
    if (existingIdx >= 0) {
      window.allPayrollRecords[existingIdx] = updatedRec;
    } else {
      window.allPayrollRecords.push(updatedRec);
    }
  }
}

async function persistRecords() {
  if (saveInFlight) {
    saveQueued = true;
    return;
  }

  saveInFlight = true;
  const month = getSelectedMonth();
  syncCurrentRecordsToAllPayroll();

  try {
    await apiRequest(payrollMonthUrl(month), {
      method: "PUT",
      body: { records: currentRecords },
    });
    setSaveStatus("All changes saved.");
    saveInFlight = false;
    if (saveQueued) {
      saveQueued = false;
      persistRecords();
    }
  } catch (error) {
    setSaveStatus("Network busy. Auto-retrying save...");
    saveInFlight = false;
    saveQueued = true;
    if (!pendingSaveTimer) {
      pendingSaveTimer = setTimeout(() => {
        pendingSaveTimer = null;
        persistRecords();
      }, 4000);
    }
  }
}

function getRandomMotivatingMessage() {
  return MOTIVATING_MESSAGES[Math.floor(Math.random() * MOTIVATING_MESSAGES.length)];
}

function openPayslip(record, month, companyOverride = null) {
  const calc = computePayroll(record, month);
  const company = companyOverride || getActiveCompany();
  const displayComment = String(record.comment || "").trim() || getRandomMotivatingMessage();
  activePayslip = { record, calc, month, company, displayComment };
  SELECTORS.payslipPreview.innerHTML = renderPayslipCard(record, calc, month, company, displayComment);
  SELECTORS.payslipDialog.showModal();
}

function renderPayslipCard(record, calc, month, company, displayComment = null) {
  const monthSheet = formatPayslipMonth(month);
  const comment = displayComment || String(record.comment || "").trim() || getRandomMotivatingMessage();
  const status = formatStatusLabel(record.employeeStatus || "working");
  const absenceSummary = calc.daysAbsent > 0
    ? `${formatNumberValue(calc.workedDays)} day(s) worked out of ${calc.monthDays} • Absence deduction ${formatCurrency(calc.proratedAbsenceDeduction)}`
    : "No absence deduction applied";

  return `
    <article class="payslip-sheet">
      <header class="payslip-sheet-header">
        <div class="payslip-sheet-brand">
          <div>
            <h2>${escapeHtml(company.name || "Company")}</h2>
            <p>Official Payroll Statement</p>
          </div>
          ${company.logoDataUrl ? `<img src="${company.logoDataUrl}" alt="${escapeHtml(company.name)} logo" class="payslip-sheet-logo" />` : ""}
        </div>
        <div class="payslip-sheet-period">
          <span class="payslip-label">Month Sheet</span>
          <strong>${escapeHtml(monthSheet)}</strong>
        </div>
      </header>
      <section class="payslip-identity">
        <div>
          <span class="payslip-label">Employee Name</span>
          <strong>${escapeHtml(record.employeeName || "Employee Name")}</strong>
        </div>
        <div>
          <span class="payslip-label">Designation</span>
          <strong>${escapeHtml(record.designation || "Designation")}</strong>
        </div>
        <div>
          <span class="payslip-label">Employee ID</span>
          <strong>${escapeHtml(record.employeeId || "-")}</strong>
        </div>
      </section>
      <section class="payslip-financials">
        <div class="payslip-panel payslip-panel-addition">
          <h3>ADDITION</h3>
          <div class="payslip-lines">
            <div class="payslip-line"><span>Present Salary (₹)</span><strong>${formatPayslipAmount(calc.presentSalary)}</strong></div>
            <div class="payslip-line"><span>Increment (₹)</span><strong class="amount-green">${formatPayslipAmount(calc.increment)}</strong></div>
            <div class="payslip-line payslip-line-total"><span>Gross Salary (₹)</span><strong class="amount-underline">${formatPayslipAmount(calc.grossSalary)}</strong></div>
          </div>
        </div>
        <div class="payslip-panel payslip-panel-deduction">
          <h3>DEDUCTIONS</h3>
          <div class="payslip-lines">
            <div class="payslip-line"><span>Old Advance Taken (₹)</span><strong>${formatPayslipAmount(calc.oldAdvanceTaken)}</strong></div>
            <div class="payslip-line"><span>Extra Advance Added (₹)</span><strong>${formatPayslipAmount(calc.extraAdvanceAdded)}</strong></div>
            <div class="payslip-line"><span>Total Advance (₹)</span><strong>${formatPayslipAmount(calc.totalAdvance)}</strong></div>
            <div class="payslip-line payslip-line-highlight"><span>Deduction Applied (₹)</span><strong class="amount-red">${formatPayslipAmount(calc.deductionApplied)}</strong></div>
            <div class="payslip-line payslip-line-total"><span>Advance Remained (₹)</span><strong class="amount-red-large">${formatPayslipAmount(calc.advanceRemained)}</strong></div>
          </div>
        </div>
      </section>
      <section class="payslip-note-strip">
        <div>
          <span class="payslip-label">Payroll Notes</span>
          <p>${escapeHtml(`Status: ${status} • Deduction Entered: ${formatCurrency(calc.deductionEntered)} • ${absenceSummary}`)}</p>
        </div>
      </section>
      <footer class="payslip-footer">
        <div class="payslip-comment">
          <span class="payslip-label">Comment</span>
          <p>${escapeHtml(comment)}</p>
        </div>
        <div class="payslip-net">
          <span class="payslip-label">Salary In Hand (Net Salary ₹)</span>
          <div class="payslip-net-box">${formatPayslipAmount(calc.netSalary)}</div>
        </div>
      </footer>
      <div class="payslip-signoff">
        <span>Prepared By: System</span>
        <div class="payslip-signature">Employer Signature</div>
      </div>
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
  const { record, calc, month, company, displayComment } = data;
  const comment = displayComment || String(record.comment || "").trim() || getRandomMotivatingMessage();
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
    `Worked Days: ${calc.workedDays} / ${calc.monthDays}`,
    `Total Days Absent: ${calc.daysAbsent} / ${calc.monthDays}`,
    `Salary For Worked Days: ${formatCurrency(calc.payableSalary)}`,
    `Prorated Absence Deduction: ${formatCurrency(calc.proratedAbsenceDeduction)}`,
    `Deduction Applied: ${formatCurrency(calc.deductionApplied)}`,
    `Advance Remained: ${formatCurrency(calc.advanceRemained)}`,
    `Salary In Hand: ${formatCurrency(calc.netSalary)}`,
    `Comment: ${comment}`,
  ].join("\n");
}

function printCurrentPayslip() {
  if (!ensureActivePayslip()) return;
  const printWindow = openPayslipPrintWindow();
  if (!printWindow) {
    showAppMessage("Pop-up blocked. Allow pop-ups to print.");
    return;
  }
  triggerPayslipPrint(printWindow);
}

function downloadCurrentPayslipPdf() {
  if (!ensureActivePayslip()) return;
  const printWindow = openPayslipPrintWindow();
  if (!printWindow) {
    showAppMessage("Pop-up blocked. Allow pop-ups to save PDF.");
    return;
  }
  showAppMessage("Choose Save as PDF in the print dialog.");
  triggerPayslipPrint(printWindow);
}

function openPayslipPrintWindow() {
  const html = buildPayslipPrintDocument();
  const printWindow = window.open("", "_blank", "width=900,height=800");
  if (!printWindow) {
    return null;
  }
  printWindow.document.write(html);
  printWindow.document.close();
  return printWindow;
}

function triggerPayslipPrint(printWindow) {
  printWindow.focus();
  window.setTimeout(() => {
    printWindow.print();
  }, 180);
}

function buildPayslipPrintDocument() {
  if (!activePayslip) return "";
  return `
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Payslip - ${escapeHtml(activePayslip.record.employeeName || "")}</title>
        <style>
          ${getPayslipPrintStyles()}
        </style>
      </head>
      <body>
        ${renderPayslipCard(
          activePayslip.record,
          activePayslip.calc,
          activePayslip.month,
          activePayslip.company,
          activePayslip.displayComment
        )}
      </body>
    </html>
  `;
}

function downloadCurrentPayslipText() {
  if (!ensureActivePayslip()) return;
  const text = getPayslipText(activePayslip);
  const name = slugify(activePayslip.record.employeeName || activePayslip.record.employeeId || "employee");
  downloadBlob(new Blob([text], { type: "text/plain;charset=utf-8" }), `payslip-${name}-${activePayslip.month}.txt`);
}

function formatPayslipMonth(month) {
  const raw = String(month || "");
  if (!raw.includes("-")) return raw;
  const [year, monthPart] = raw.split("-");
  const date = new Date(Number(year), Number(monthPart) - 1, 1);
  if (Number.isNaN(date.getTime())) return raw;
  return `${date.toLocaleDateString("en-IN", { month: "short" })}_${year}`;
}

function formatNumberValue(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "0";
  const hasFraction = Math.abs(num % 1) > 0.0001;
  return new Intl.NumberFormat("en-IN", {
    minimumFractionDigits: hasFraction ? 2 : 0,
    maximumFractionDigits: hasFraction ? 2 : 0,
  }).format(num);
}

function formatPayslipAmount(value) {
  return formatNumberValue(value);
}

function formatStatusLabel(status) {
  const raw = String(status || "working").trim().toLowerCase();
  if (!raw) return "Working";
  return raw.replace(/\b\w/g, (char) => char.toUpperCase());
}

function getPayslipPrintStyles() {
  return `
    @page {
      size: A5 portrait;
      margin: 8mm;
    }

    * {
      box-sizing: border-box;
    }

    html, body {
      margin: 0;
      padding: 0;
      background: #ffffff;
      color: #111827;
      font-family: Inter, "Segoe UI", Arial, sans-serif;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    body {
      padding: 0;
    }

    .payslip-sheet {
      width: 100%;
      min-height: calc(210mm - 16mm);
      padding: 9mm;
      border: 1px solid #d8e0ea;
      border-radius: 3mm;
      background: #ffffff;
      display: flex;
      flex-direction: column;
      gap: 0;
    }

    .payslip-sheet-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      gap: 8mm;
      padding-bottom: 6mm;
      border-bottom: 1px solid #e7edf3;
      margin-bottom: 7mm;
    }

    .payslip-sheet-brand {
      display: flex;
      align-items: center;
      gap: 4mm;
      min-width: 0;
      flex: 1;
    }

    .payslip-sheet-logo {
      width: 13mm;
      height: 13mm;
      object-fit: cover;
      border-radius: 2mm;
      border: 1px solid #d8e0ea;
      flex-shrink: 0;
    }

    .payslip-sheet-brand h2 {
      margin: 0;
      font-size: 8.8mm;
      line-height: 1.02;
      font-weight: 800;
      letter-spacing: -0.04em;
      color: #17223b;
      text-transform: uppercase;
    }

    .payslip-sheet-brand p {
      margin: 1mm 0 0;
      font-size: 3.4mm;
      font-weight: 600;
      color: #64748b;
    }

    .payslip-sheet-period {
      text-align: right;
      flex-shrink: 0;
    }

    .payslip-sheet-period strong {
      display: block;
      margin-top: 1mm;
      font-size: 5.8mm;
      font-weight: 800;
      color: #111827;
    }

    .payslip-label {
      display: block;
      margin-bottom: 1.2mm;
      font-size: 2.8mm;
      line-height: 1.2;
      font-weight: 700;
      color: #6b7280;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }

    .payslip-identity {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 6mm;
      margin-bottom: 8mm;
    }

    .payslip-identity strong {
      display: block;
      font-size: 4.2mm;
      line-height: 1.25;
      font-weight: 700;
      color: #111827;
    }

    .payslip-financials {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8mm;
      margin-bottom: 7mm;
    }

    .payslip-panel h3 {
      margin: 0 0 3mm;
      padding-bottom: 1.2mm;
      font-size: 4mm;
      line-height: 1.2;
      font-weight: 800;
      border-bottom: 0.45mm solid #111827;
    }

    .payslip-panel-addition h3 {
      color: #047857;
      border-color: #0f9f78;
    }

    .payslip-panel-deduction h3 {
      color: #b91c1c;
      border-color: #ef4444;
    }

    .payslip-lines {
      display: flex;
      flex-direction: column;
      gap: 0.8mm;
    }

    .payslip-line {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 4mm;
      padding: 2.4mm 0;
      border-bottom: 1px solid #eef2f7;
    }

    .payslip-line span {
      font-size: 3.6mm;
      line-height: 1.3;
      color: #64748b;
    }

    .payslip-line strong {
      font-size: 4mm;
      line-height: 1.2;
      font-weight: 700;
      color: #111827;
      text-align: right;
    }

    .payslip-line-total {
      border-bottom: none;
      padding-top: 3.5mm;
    }

    .payslip-line-total span {
      font-weight: 700;
      color: #1f2937;
    }

    .payslip-line-total strong {
      font-size: 5.4mm;
      font-weight: 800;
    }

    .payslip-line-highlight {
      background: #f8fafc;
      padding-left: 2mm;
      padding-right: 2mm;
    }

    .payslip-line-highlight span,
    .amount-red {
      color: #ef2b24;
      font-weight: 700;
    }

    .amount-green {
      color: #06a168;
    }

    .amount-underline {
      padding-bottom: 0.6mm;
      border-bottom: 0.5mm solid #111827;
    }

    .amount-red-large {
      color: #c81e1e;
      font-size: 5.8mm;
      font-weight: 800;
    }

    .payslip-note-strip {
      margin-top: auto;
      padding-top: 4mm;
      border-top: 1.2mm double #dce5ef;
    }

    .payslip-note-strip p {
      margin: 0;
      font-size: 3.1mm;
      line-height: 1.45;
      color: #5b708c;
    }

    .payslip-footer {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 8mm;
      margin-top: 7mm;
    }

    .payslip-comment {
      flex: 1;
      min-width: 0;
    }

    .payslip-comment p {
      margin: 0;
      padding-left: 3mm;
      border-left: 0.5mm solid #d8e0ea;
      font-size: 3.4mm;
      line-height: 1.45;
      color: #64748b;
      font-style: italic;
    }

    .payslip-net {
      width: 43mm;
      text-align: right;
      flex-shrink: 0;
    }

    .payslip-net-box {
      display: inline-block;
      margin-top: 1.5mm;
      min-width: 36mm;
      padding: 2.8mm 4.5mm;
      border: 0.6mm solid #17223b;
      border-radius: 1.6mm;
      font-size: 8.5mm;
      line-height: 1;
      font-weight: 800;
      letter-spacing: -0.05em;
      color: #202226;
      text-align: center;
      background: #ffffff;
    }

    .payslip-signoff {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      margin-top: 11mm;
      font-size: 2.9mm;
      line-height: 1.2;
      font-weight: 700;
      color: #8ca0bf;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }

    .payslip-signature {
      width: 38mm;
      padding-top: 1.6mm;
      border-top: 1px solid #cbd5e1;
      text-align: center;
    }
  `;
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
    "Worked Days",
    "Payroll Month Days",
    "Manual Days Absent",
    "Total Days Absent",
    "Salary For Worked Days (₹)",
    "Prorated Absence Deduction (₹)",
    "Deduction Applied (Advance only)",
    "Advance Remained (₹)",
    "Salary In Hand (Net Salary ₹)",
    "Comment",
  ];

  const rows = records.map((record) => {
    const calc = computePayroll(record, month);
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
      calc.workedDays,
      calc.monthDays,
      calc.manualDaysAbsent,
      calc.daysAbsent,
      calc.payableSalary,
      calc.proratedAbsenceDeduction,
      calc.deductionApplied,
      calc.advanceRemained,
      calc.netSalary,
      record.comment || "",
    ];
  });

  return [headers, ...rows].map((row) => row.map(escapeCsv).join(",")).join("\n");
}

async function fetchPayrollReportSnapshot(reportId) {
  const report = payrollReports.find((item) => Number(item.id) === Number(reportId)) || null;
  if (Number(activePayrollReportId) === Number(reportId) && activePayrollReportSnapshot) {
    return {
      report,
      snapshot: activePayrollReportSnapshot,
    };
  }

  const response = await apiRequest(`/api/payroll-reports/${reportId}?companyId=${getSelectedCompanyId()}`);
  return {
    report: response.report || report,
    snapshot: response.snapshot || null,
  };
}

async function downloadPayrollReportExcel(reportId) {
  const payload = await fetchPayrollReportSnapshot(reportId);
  const report = payload.report || {};
  const fallbackMonth = String(report.month || "");
  const snapshot = payload.snapshot && typeof payload.snapshot === "object"
    ? payload.snapshot
    : { company: getActiveCompany(), month: fallbackMonth, records: [] };

  const workbook = makeDetailedPayrollReportSpreadsheet(snapshot, report);
  const companyName = snapshot.company?.name || getActiveCompany().name || "company";
  const month = String(snapshot.month || fallbackMonth || "payroll-report");
  downloadBlob(
    new Blob([workbook], { type: "application/vnd.ms-excel;charset=utf-8" }),
    `generated-payroll-${slugify(companyName)}-${month}.xls`
  );
  showAppMessage(`Detailed Excel exported for ${formatMonth(month)}.`);
}

function summarizePayrollSnapshot(snapshot, report = {}) {
  const month = String(snapshot?.month || report?.month || getSelectedMonth());
  const company = snapshot?.company || getActiveCompany();
  const records = Array.isArray(snapshot?.records) ? snapshot.records : [];
  const computed = records.map((record) => ({
    record,
    calc: computePayroll(record, month),
  }));

  const totals = computed.reduce((acc, item) => {
    acc.presentSalary += item.calc.presentSalary;
    acc.increment += item.calc.increment;
    acc.grossSalary += item.calc.grossSalary;
    acc.oldAdvanceTaken += item.calc.oldAdvanceTaken;
    acc.extraAdvanceAdded += item.calc.extraAdvanceAdded;
    acc.totalAdvance += item.calc.totalAdvance;
    acc.deductionEntered += item.calc.deductionEntered;
    acc.proratedAbsenceDeduction += item.calc.proratedAbsenceDeduction;
    acc.deductionApplied += item.calc.deductionApplied;
    acc.advanceRemained += item.calc.advanceRemained;
    acc.netSalary += item.calc.netSalary;
    acc.daysAbsent += item.calc.daysAbsent;
    return acc;
  }, {
    presentSalary: 0,
    increment: 0,
    grossSalary: 0,
    oldAdvanceTaken: 0,
    extraAdvanceAdded: 0,
    totalAdvance: 0,
    deductionEntered: 0,
    proratedAbsenceDeduction: 0,
    deductionApplied: 0,
    advanceRemained: 0,
    netSalary: 0,
    daysAbsent: 0,
  });

  const statusCounts = new Map();
  const designationCounts = new Map();
  for (const item of computed) {
    const statusKey = String(item.record.employeeStatus || "working").trim().toLowerCase() || "working";
    statusCounts.set(statusKey, (statusCounts.get(statusKey) || 0) + 1);

    const designationKey = String(item.record.designation || "Unassigned").trim() || "Unassigned";
    designationCounts.set(designationKey, (designationCounts.get(designationKey) || 0) + 1);
  }

  const statusRows = Array.from(statusCounts.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([status, count]) => ({
      status: status.replace(/\b\w/g, (char) => char.toUpperCase()),
      count,
    }));

  const designationRows = Array.from(designationCounts.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([designation, count]) => ({ designation, count }));

  const topNetRows = computed
    .slice()
    .sort((a, b) => b.calc.netSalary - a.calc.netSalary || String(a.record.employeeName || "").localeCompare(String(b.record.employeeName || "")))
    .slice(0, 10);

  return {
    month,
    company,
    records,
    computed,
    totals,
    statusRows,
    designationRows,
    topNetRows,
    checkedAt: report?.checkedAt || "",
    generatedAt: report?.generatedAt || "",
  };
}

function makeDetailedPayrollReportSpreadsheet(snapshot, report = {}) {
  const summary = summarizePayrollSnapshot(snapshot, report);
  const monthLabel = formatMonth(summary.month);
  const generatedAt = formatDateTime(summary.generatedAt) || formatDateTime(new Date().toISOString());
  const checkedAt = formatDateTime(summary.checkedAt) || "-";
  const companyName = summary.company?.name || "Company";
  const records = summary.records;
  const computed = summary.computed;

  const summaryRows = [
    `<Row ss:Height="30"><Cell ss:StyleID="rTitle"><Data ss:Type="String">${escapeXml(`${companyName} - Generated Payroll Report`)}</Data></Cell><Cell ss:StyleID="rTitleFill"/><Cell ss:StyleID="rTitleFill"/><Cell ss:StyleID="rTitleFill"/><Cell ss:StyleID="rTitleFill"/><Cell ss:StyleID="rTitleFill"/></Row>`,
    `<Row ss:Height="22"><Cell ss:StyleID="rSubtitle" ss:MergeAcross="5"><Data ss:Type="String">${escapeXml(`Detailed monthly payroll workbook for ${monthLabel}`)}</Data></Cell></Row>`,
    `<Row ss:Height="8"></Row>`,
    `<Row>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Company</Data></Cell>
      <Cell ss:StyleID="rMetaValue"><Data ss:Type="String">${escapeXml(companyName)}</Data></Cell>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Payroll Month</Data></Cell>
      <Cell ss:StyleID="rMetaValue"><Data ss:Type="String">${escapeXml(monthLabel)}</Data></Cell>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Employees</Data></Cell>
      <Cell ss:StyleID="rMetaValueNumber"><Data ss:Type="Number">${records.length}</Data></Cell>
    </Row>`,
    `<Row>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Checked At</Data></Cell>
      <Cell ss:StyleID="rMetaValue"><Data ss:Type="String">${escapeXml(checkedAt)}</Data></Cell>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Generated At</Data></Cell>
      <Cell ss:StyleID="rMetaValue"><Data ss:Type="String">${escapeXml(generatedAt)}</Data></Cell>
      <Cell ss:StyleID="rMetaLabel"><Data ss:Type="String">Workbook Created</Data></Cell>
      <Cell ss:StyleID="rMetaValue"><Data ss:Type="String">${escapeXml(new Date().toLocaleString("en-IN"))}</Data></Cell>
    </Row>`,
    `<Row ss:Height="10"></Row>`,
    `<Row><Cell ss:StyleID="rSection" ss:MergeAcross="5"><Data ss:Type="String">Monthly Totals</Data></Cell></Row>`,
    `<Row>
      <Cell ss:StyleID="rKpiBlueLabel"><Data ss:Type="String">Gross Salary</Data></Cell>
      <Cell ss:StyleID="rKpiBlueValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.grossSalary)}</Data></Cell>
      <Cell ss:StyleID="rKpiGreenLabel"><Data ss:Type="String">Net Salary</Data></Cell>
      <Cell ss:StyleID="rKpiGreenValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.netSalary)}</Data></Cell>
      <Cell ss:StyleID="rKpiGoldLabel"><Data ss:Type="String">Advance Remaining</Data></Cell>
      <Cell ss:StyleID="rKpiGoldValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.advanceRemained)}</Data></Cell>
    </Row>`,
    `<Row>
      <Cell ss:StyleID="rKpiPurpleLabel"><Data ss:Type="String">Present Salary</Data></Cell>
      <Cell ss:StyleID="rKpiPurpleValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.presentSalary)}</Data></Cell>
      <Cell ss:StyleID="rKpiSkyLabel"><Data ss:Type="String">Increment</Data></Cell>
      <Cell ss:StyleID="rKpiSkyValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.increment)}</Data></Cell>
      <Cell ss:StyleID="rKpiRedLabel"><Data ss:Type="String">Total Deductions</Data></Cell>
      <Cell ss:StyleID="rKpiRedValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.deductionApplied + summary.totals.proratedAbsenceDeduction)}</Data></Cell>
    </Row>`,
    `<Row>
      <Cell ss:StyleID="rKpiSlateLabel"><Data ss:Type="String">Total Advance</Data></Cell>
      <Cell ss:StyleID="rKpiSlateValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.totalAdvance)}</Data></Cell>
      <Cell ss:StyleID="rKpiSlateLabel"><Data ss:Type="String">Absence Deduction</Data></Cell>
      <Cell ss:StyleID="rKpiSlateValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.proratedAbsenceDeduction)}</Data></Cell>
      <Cell ss:StyleID="rKpiSlateLabel"><Data ss:Type="String">Total Days Absent</Data></Cell>
      <Cell ss:StyleID="rKpiSlateValue"><Data ss:Type="Number">${toFixedNumber(summary.totals.daysAbsent)}</Data></Cell>
    </Row>`,
    `<Row ss:Height="10"></Row>`,
    `<Row><Cell ss:StyleID="rSection" ss:MergeAcross="2"><Data ss:Type="String">Employee Status Mix</Data></Cell><Cell/><Cell ss:StyleID="rSection" ss:MergeAcross="2"><Data ss:Type="String">Designation Breakdown</Data></Cell></Row>`,
    ...buildPairedSummaryRows(
      summary.statusRows.map((item) => [xmlCellString(item.status, "rSummaryLabel"), xmlCellNumber(item.count, "rSummaryCount"), xmlCellBlank("rSummarySpacer")]),
      summary.designationRows.slice(0, 12).map((item) => [xmlCellString(item.designation, "rSummaryLabel"), xmlCellNumber(item.count, "rSummaryCount"), xmlCellBlank("rSummarySpacer")])
    ),
    `<Row ss:Height="10"></Row>`,
    `<Row><Cell ss:StyleID="rSection" ss:MergeAcross="5"><Data ss:Type="String">Top Net Salary Earners</Data></Cell></Row>`,
    `<Row>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Employee ID</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Employee Name</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Designation</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Status</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Net Salary</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Advance Remaining</Data></Cell>
    </Row>`,
    ...(summary.topNetRows.length
      ? summary.topNetRows.map((item, index) => {
        const rowStyle = index % 2 === 0 ? "rData" : "rDataAlt";
        return `<Row>
          ${xmlCellString(item.record.employeeId || "", rowStyle)}
          ${xmlCellString(item.record.employeeName || "", rowStyle)}
          ${xmlCellString(item.record.designation || "", rowStyle)}
          ${xmlCellString(item.record.employeeStatus || "working", rowStyle)}
          ${xmlCellNumber(item.calc.netSalary, "rMoneyNet")}
          ${xmlCellNumber(item.calc.advanceRemained, "rMoneyGold")}
        </Row>`;
      }).join("")
      : `<Row><Cell ss:StyleID="rEmpty" ss:MergeAcross="5"><Data ss:Type="String">No generated payroll records available for this month.</Data></Cell></Row>`),
  ].join("");

  const detailRows = computed.map((item, index) => {
    const rowStyle = index % 2 === 0 ? "rData" : "rDataAlt";
    return `<Row>
      ${xmlCellString(item.record.employeeId || "", rowStyle)}
      ${xmlCellString(item.record.employeeName || "", rowStyle)}
      ${xmlCellString(item.record.designation || "", rowStyle)}
      ${xmlCellString(item.record.employeeStatus || "working", rowStyle)}
      ${xmlCellString(item.record.leaveFrom || "-", rowStyle)}
      ${xmlCellString(item.record.leaveTo || "-", rowStyle)}
      ${xmlCellString(item.record.terminatedOn || "-", rowStyle)}
      ${xmlCellNumber(item.calc.presentSalary, "rMoneyBlue")}
      ${xmlCellNumber(item.calc.increment, "rMoneyPurple")}
      ${xmlCellNumber(item.calc.grossSalary, "rMoneyGreen")}
      ${xmlCellNumber(item.calc.oldAdvanceTaken, "rMoneyGold")}
      ${xmlCellNumber(item.calc.extraAdvanceAdded, "rMoneyGold")}
      ${xmlCellNumber(item.calc.totalAdvance, "rMoneyGoldStrong")}
      ${xmlCellNumber(item.calc.deductionEntered, "rMoneyRed")}
      ${xmlCellNumber(item.calc.daysAbsent, "rNumber")}
      ${xmlCellNumber(item.calc.proratedAbsenceDeduction, "rMoneyRed")}
      ${xmlCellNumber(item.calc.deductionApplied, "rMoneyRedStrong")}
      ${xmlCellNumber(item.calc.advanceRemained, "rMoneyGoldStrong")}
      ${xmlCellNumber(item.calc.netSalary, "rMoneyNet")}
      ${xmlCellString(item.record.comment || "", rowStyle)}
    </Row>`;
  }).join("");

  const detailSheetRows = `
    <Row ss:Height="28"><Cell ss:StyleID="rTitle" ss:MergeAcross="19"><Data ss:Type="String">${escapeXml(`${companyName} - Employee Payroll Details`)}</Data></Cell></Row>
    <Row ss:Height="22"><Cell ss:StyleID="rSubtitle" ss:MergeAcross="19"><Data ss:Type="String">${escapeXml(`Detailed employee-level payroll for ${monthLabel}`)}</Data></Cell></Row>
    <Row ss:Height="8"></Row>
    <Row ss:Height="34">
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Employee ID</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Employee Name</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Designation</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Status</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Leave From</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Resumed On</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Terminated On</Data></Cell>
      <Cell ss:StyleID="rTableHeaderBlue"><Data ss:Type="String">Present Salary</Data></Cell>
      <Cell ss:StyleID="rTableHeaderPurple"><Data ss:Type="String">Increment</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGreen"><Data ss:Type="String">Gross Salary</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGold"><Data ss:Type="String">Old Advance</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGold"><Data ss:Type="String">Extra Advance</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGold"><Data ss:Type="String">Total Advance</Data></Cell>
      <Cell ss:StyleID="rTableHeaderRed"><Data ss:Type="String">Deduction Entered</Data></Cell>
      <Cell ss:StyleID="rTableHeaderAmber"><Data ss:Type="String">Days Absent</Data></Cell>
      <Cell ss:StyleID="rTableHeaderRed"><Data ss:Type="String">Absence Deduction</Data></Cell>
      <Cell ss:StyleID="rTableHeaderRed"><Data ss:Type="String">Deduction Applied</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGold"><Data ss:Type="String">Advance Remaining</Data></Cell>
      <Cell ss:StyleID="rTableHeaderGreen"><Data ss:Type="String">Net Salary</Data></Cell>
      <Cell ss:StyleID="rTableHeader"><Data ss:Type="String">Comment</Data></Cell>
    </Row>
    ${detailRows || `<Row><Cell ss:StyleID="rEmpty" ss:MergeAcross="19"><Data ss:Type="String">No generated payroll rows found for this month.</Data></Cell></Row>`}
    <Row ss:Height="4"></Row>
    <Row>
      <Cell ss:StyleID="rTotalsLabel" ss:MergeAcross="6"><Data ss:Type="String">Monthly Totals</Data></Cell>
      ${xmlCellNumber(summary.totals.presentSalary, "rTotalsBlue")}
      ${xmlCellNumber(summary.totals.increment, "rTotalsPurple")}
      ${xmlCellNumber(summary.totals.grossSalary, "rTotalsGreen")}
      ${xmlCellNumber(summary.totals.oldAdvanceTaken, "rTotalsGold")}
      ${xmlCellNumber(summary.totals.extraAdvanceAdded, "rTotalsGold")}
      ${xmlCellNumber(summary.totals.totalAdvance, "rTotalsGoldStrong")}
      ${xmlCellNumber(summary.totals.deductionEntered, "rTotalsRed")}
      ${xmlCellNumber(summary.totals.daysAbsent, "rTotalsNumber")}
      ${xmlCellNumber(summary.totals.proratedAbsenceDeduction, "rTotalsRed")}
      ${xmlCellNumber(summary.totals.deductionApplied, "rTotalsRedStrong")}
      ${xmlCellNumber(summary.totals.advanceRemained, "rTotalsGoldStrong")}
      ${xmlCellNumber(summary.totals.netSalary, "rTotalsGreen")}
      ${xmlCellBlank("rTotalsComment")}
    </Row>
  `;

  const financeRows = `
    <Row ss:Height="28"><Cell ss:StyleID="rTitle" ss:MergeAcross="5"><Data ss:Type="String">${escapeXml(`${companyName} - Finance Summary`)}</Data></Cell></Row>
    <Row ss:Height="22"><Cell ss:StyleID="rSubtitle" ss:MergeAcross="5"><Data ss:Type="String">${escapeXml(`Payroll reconciliation for ${monthLabel}`)}</Data></Cell></Row>
    <Row ss:Height="8"></Row>
    <Row><Cell ss:StyleID="rSection" ss:MergeAcross="5"><Data ss:Type="String">Totals Reconciliation</Data></Cell></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Present Salary</Data></Cell><Cell ss:StyleID="rSummaryMoneyBlue"><Data ss:Type="Number">${toFixedNumber(summary.totals.presentSalary)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Deduction Entered</Data></Cell><Cell ss:StyleID="rSummaryMoneyRed"><Data ss:Type="Number">${toFixedNumber(summary.totals.deductionEntered)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Increment</Data></Cell><Cell ss:StyleID="rSummaryMoneyPurple"><Data ss:Type="Number">${toFixedNumber(summary.totals.increment)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Absence Deduction</Data></Cell><Cell ss:StyleID="rSummaryMoneyRed"><Data ss:Type="Number">${toFixedNumber(summary.totals.proratedAbsenceDeduction)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Gross Salary</Data></Cell><Cell ss:StyleID="rSummaryMoneyGreen"><Data ss:Type="Number">${toFixedNumber(summary.totals.grossSalary)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Deduction Applied</Data></Cell><Cell ss:StyleID="rSummaryMoneyRedStrong"><Data ss:Type="Number">${toFixedNumber(summary.totals.deductionApplied)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Old Advance</Data></Cell><Cell ss:StyleID="rSummaryMoneyGold"><Data ss:Type="Number">${toFixedNumber(summary.totals.oldAdvanceTaken)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Advance Remaining</Data></Cell><Cell ss:StyleID="rSummaryMoneyGold"><Data ss:Type="Number">${toFixedNumber(summary.totals.advanceRemained)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Extra Advance</Data></Cell><Cell ss:StyleID="rSummaryMoneyGold"><Data ss:Type="Number">${toFixedNumber(summary.totals.extraAdvanceAdded)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Net Salary</Data></Cell><Cell ss:StyleID="rSummaryMoneyGreenStrong"><Data ss:Type="Number">${toFixedNumber(summary.totals.netSalary)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
    <Row><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Total Advance</Data></Cell><Cell ss:StyleID="rSummaryMoneyGoldStrong"><Data ss:Type="Number">${toFixedNumber(summary.totals.totalAdvance)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/><Cell ss:StyleID="rSummaryLabel"><Data ss:Type="String">Average Net Salary</Data></Cell><Cell ss:StyleID="rSummaryMoneyGreen"><Data ss:Type="Number">${toFixedNumber(records.length ? summary.totals.netSalary / records.length : 0)}</Data></Cell><Cell ss:StyleID="rSummarySpacer"/></Row>
  `;

  const summaryColumns = [150, 170, 140, 170, 150, 160];
  const detailColumns = [90, 150, 130, 95, 90, 90, 95, 100, 90, 100, 100, 100, 105, 110, 80, 115, 110, 115, 105, 180];
  const financeColumns = [180, 140, 18, 180, 140, 18];

  return `<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Author>Routes Payroll</Author>
    <Created>${new Date().toISOString()}</Created>
    <Company>${escapeXml(companyName)}</Company>
  </DocumentProperties>
  <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
    <WindowHeight>12800</WindowHeight>
    <WindowWidth>26000</WindowWidth>
    <ProtectStructure>False</ProtectStructure>
    <ProtectWindows>False</ProtectWindows>
  </ExcelWorkbook>
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Borders/><Font ss:FontName="Calibri" ss:Size="11" ss:Color="#0F172A"/><Interior/><NumberFormat/><Protection/></Style>
    <Style ss:ID="rTitle"><Font ss:FontName="Calibri" ss:Size="16" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#123A63" ss:Pattern="Solid"/><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0C2743"/></Borders></Style>
    <Style ss:ID="rTitleFill"><Interior ss:Color="#123A63" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0C2743"/></Borders></Style>
    <Style ss:ID="rSubtitle"><Font ss:FontName="Calibri" ss:Size="11" ss:Italic="1" ss:Color="#334155"/><Interior ss:Color="#EEF5FB" ss:Pattern="Solid"/><Alignment ss:Horizontal="Left" ss:Vertical="Center"/></Style>
    <Style ss:ID="rSection"><Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#123A63"/><Interior ss:Color="#E7F0F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C3D4E5"/></Borders></Style>
    <Style ss:ID="rMetaLabel"><Font ss:Bold="1" ss:Color="#334155"/><Interior ss:Color="#F5F8FC" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#DDE7F0"/></Borders></Style>
    <Style ss:ID="rMetaValue"><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E5EDF5"/></Borders></Style>
    <Style ss:ID="rMetaValueNumber"><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><NumberFormat ss:Format="#,##0"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E5EDF5"/></Borders></Style>
    <Style ss:ID="rKpiBlueLabel"><Font ss:Bold="1" ss:Color="#0F4C81"/><Interior ss:Color="#EAF3FF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CFDEEF"/></Borders></Style>
    <Style ss:ID="rKpiBlueValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F4C81"/><Interior ss:Color="#F5FAFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CFDEEF"/></Borders></Style>
    <Style ss:ID="rKpiGreenLabel"><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#EAF8F1" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CCE9DB"/></Borders></Style>
    <Style ss:ID="rKpiGreenValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#F4FCF7" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CCE9DB"/></Borders></Style>
    <Style ss:ID="rKpiGoldLabel"><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF5DB" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F2DDAB"/></Borders></Style>
    <Style ss:ID="rKpiGoldValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF9EB" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F2DDAB"/></Borders></Style>
    <Style ss:ID="rKpiPurpleLabel"><Font ss:Bold="1" ss:Color="#5B3FA3"/><Interior ss:Color="#F1EBFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#DCD0FB"/></Borders></Style>
    <Style ss:ID="rKpiPurpleValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#5B3FA3"/><Interior ss:Color="#F8F5FF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#DCD0FB"/></Borders></Style>
    <Style ss:ID="rKpiSkyLabel"><Font ss:Bold="1" ss:Color="#0A6B8E"/><Interior ss:Color="#E7F7FD" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C7E5F2"/></Borders></Style>
    <Style ss:ID="rKpiSkyValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0A6B8E"/><Interior ss:Color="#F3FBFE" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C7E5F2"/></Borders></Style>
    <Style ss:ID="rKpiRedLabel"><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FDEFF1" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F1CBD1"/></Borders></Style>
    <Style ss:ID="rKpiRedValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FFF7F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F1CBD1"/></Borders></Style>
    <Style ss:ID="rKpiSlateLabel"><Font ss:Bold="1" ss:Color="#334155"/><Interior ss:Color="#F6F8FB" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2EC"/></Borders></Style>
    <Style ss:ID="rKpiSlateValue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#334155"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2EC"/></Borders></Style>
    <Style ss:ID="rSummaryLabel"><Font ss:Bold="1" ss:Color="#334155"/><Interior ss:Color="#F8FAFC" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryCount"><NumberFormat ss:Format="#,##0"/><Font ss:Bold="1" ss:Color="#123A63"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummarySpacer"><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/></Style>
    <Style ss:ID="rSummaryMoneyBlue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F4C81"/><Interior ss:Color="#F6FAFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyPurple"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#5B3FA3"/><Interior ss:Color="#FBF9FF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyGreen"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#F5FCF8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyGreenStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#EAF8F1" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CCE9DB"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyGold"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF9EF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyGoldStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF4D6" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F2DDAB"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyRed"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FFF7F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rSummaryMoneyRedStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FDEFF1" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#F1CBD1"/></Borders></Style>
    <Style ss:ID="rTableHeader"><Font ss:Bold="1" ss:Color="#123A63"/><Interior ss:Color="#EAF1F8" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderBlue"><Font ss:Bold="1" ss:Color="#0F4C81"/><Interior ss:Color="#EAF3FF" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderPurple"><Font ss:Bold="1" ss:Color="#5B3FA3"/><Interior ss:Color="#F1EBFF" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderGreen"><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#EAF8F1" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderGold"><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF5DB" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderRed"><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FDEFF1" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rTableHeaderAmber"><Font ss:Bold="1" ss:Color="#9A5B13"/><Interior ss:Color="#FFF4E5" ss:Pattern="Solid"/><Alignment ss:WrapText="1"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D6E5"/></Borders></Style>
    <Style ss:ID="rData"><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rDataAlt"><Interior ss:Color="#FBFDFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rNumber"><NumberFormat ss:Format="#,##0.00"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyBlue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#0F4C81"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyPurple"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#5B3FA3"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyGreen"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#F5FCF8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyGold"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#8A5A00"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyGoldStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF8E8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyRed"><NumberFormat ss:Format="#,##0.00"/><Font ss:Color="#A61B29"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyRedStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FFF4F6" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
    <Style ss:ID="rMoneyNet"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#EAF8F1" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CCE9DB"/></Borders></Style>
    <Style ss:ID="rTotalsLabel"><Font ss:Bold="1" ss:Color="#123A63"/><Interior ss:Color="#EAF1F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsBlue"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F4C81"/><Interior ss:Color="#F5FAFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsPurple"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#5B3FA3"/><Interior ss:Color="#FBF9FF" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsGreen"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#0F6B46"/><Interior ss:Color="#EDF9F1" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsGold"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFFAEE" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsGoldStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#8A5A00"/><Interior ss:Color="#FFF2CE" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsRed"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FFF7F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsRedStrong"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#A61B29"/><Interior ss:Color="#FDEFF1" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsNumber"><NumberFormat ss:Format="#,##0.00"/><Font ss:Bold="1" ss:Color="#334155"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rTotalsComment"><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#BFD0E0"/></Borders></Style>
    <Style ss:ID="rEmpty"><Font ss:Bold="1" ss:Color="#64748B"/><Interior ss:Color="#F8FAFC" ss:Pattern="Solid"/><Alignment ss:Horizontal="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  </Styles>
  <Worksheet ss:Name="Summary">
    <Table>
      ${summaryColumns.map((width) => `<Column ss:AutoFitWidth="0" ss:Width="${width}"/>`).join("")}
      ${summaryRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane></Panes></WorksheetOptions>
  </Worksheet>
  <Worksheet ss:Name="Employee Details">
    <Table>
      ${detailColumns.map((width) => `<Column ss:AutoFitWidth="0" ss:Width="${width}"/>`).join("")}
      ${detailSheetRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane></Panes></WorksheetOptions>
  </Worksheet>
  <Worksheet ss:Name="Finance Summary">
    <Table>
      ${financeColumns.map((width) => `<Column ss:AutoFitWidth="0" ss:Width="${width}"/>`).join("")}
      ${financeRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane></Panes></WorksheetOptions>
  </Worksheet>
</Workbook>`;
}

function buildPairedSummaryRows(leftRows, rightRows) {
  const total = Math.max(leftRows.length, rightRows.length);
  const rows = [];
  for (let index = 0; index < total; index += 1) {
    const left = leftRows[index] || [xmlCellBlank("rSummaryLabel"), xmlCellBlank("rSummaryCount"), xmlCellBlank("rSummarySpacer")];
    const right = rightRows[index] || [xmlCellBlank("rSummaryLabel"), xmlCellBlank("rSummaryCount"), xmlCellBlank("rSummarySpacer")];
    rows.push(`<Row>${left.join("")}${right.join("")}</Row>`);
  }
  if (!rows.length) {
    rows.push(`<Row><Cell ss:StyleID="rEmpty" ss:MergeAcross="5"><Data ss:Type="String">No summary data available.</Data></Cell></Row>`);
  }
  return rows;
}

function makeExcelSpreadsheet(records, month) {
  const company = getActiveCompany();
  const computed = records.map((record) => ({ record, calc: computePayroll(record, month) }));
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
    ["Worked Days", "sHeaderAbsent"],
    ["Month Days", "sHeaderAbsent"],
    ["Manual Days Absent", "sHeaderAbsent"],
    ["Total Days Absent", "sHeaderAbsent"],
    ["Salary For Worked Days (INR)", "sHeaderGross"],
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
      ${xmlCellNumber(item.calc.workedDays, "sNum")}
      ${xmlCellNumber(item.calc.monthDays, "sNum")}
      ${xmlCellNumber(item.calc.manualDaysAbsent, "sNum")}
      ${xmlCellNumber(item.calc.daysAbsent, "sNum")}
      ${xmlCellNumber(item.calc.payableSalary, "sMoneyGross")}
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
      <Cell ss:StyleID="sTitle" ss:MergeAcross="20"><Data ss:Type="String">${escapeXml(title)}</Data></Cell>
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

function xmlCellBlank(styleId) {
  const style = styleId ? ` ss:StyleID="${styleId}"` : "";
  return `<Cell${style}/>`;
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

function buildLegacyImportPreview(analyzed) {
  const monthSummaries = Array.isArray(analyzed?.monthSummaries) ? analyzed.monthSummaries : [];
  if (!monthSummaries.length) return "";

  const latestMonth = monthSummaries[monthSummaries.length - 1];
  const monthLines = monthSummaries.map((item) => (
    `${item.month}: ${item.rowCount} row(s), ${item.employeeCount} employee(s), `
    + `salary ${formatCurrency(item.presentSalary)}, increment ${formatCurrency(item.increment)}, `
    + `advance ${formatCurrency(item.totalAdvance)}, deduction ${formatCurrency(item.deductionEntered)}`
  ));

  return [
    `Detected ${monthSummaries.length} month(s). Latest month: ${latestMonth.month}.`,
    ...monthLines,
  ].join("\n");
}

function excelSerialToDate(value) {
  const raw = Number(value);
  if (!Number.isFinite(raw) || raw <= 0) return null;
  const utcDays = Math.floor(raw - 25569);
  const utcValue = utcDays * 86400;
  const date = new Date(utcValue * 1000);
  return Number.isNaN(date.getTime()) ? null : date;
}

async function parseLegacyPayrollFile(file) {
  const name = String(file?.name || "").toLowerCase();
  if (name.endsWith(".csv")) {
    const text = await file.text();
    const inferredMonth = inferMonthValue(file?.name || "");
    return parseCsvRows(text).map((row) => ({
      ...row,
      __sourceSheetName: String(file?.name || "CSV"),
      __sourceSheetMonth: inferredMonth,
    }));
  }

  if (!window.XLSX) {
    throw new Error("Excel parser not loaded. Reload app and try again.");
  }
  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) return [];

  function parseMonthFromSheetName(sheetName) {
    return inferMonthValue(String(sheetName || ""));
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
        row.__sourceSheetName = sheetName;
        row.__sourceSheetMonth = monthFromSheet;
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

function parseOptionalMoney(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  return parseMoney(raw);
}

function firstFiniteNumber(...values) {
  for (const value of values) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }
  }
  return null;
}

function inferMonthValue(value) {
  if (typeof value === "number") {
    const serialDate = excelSerialToDate(value);
    if (serialDate) {
      return `${serialDate.getUTCFullYear()}-${String(serialDate.getUTCMonth() + 1).padStart(2, "0")}`;
    }
  }

  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}\/\d{2}$/.test(raw)) return raw.replace("/", "-");
  if (/^\d{2}[-/]\d{4}$/.test(raw)) {
    const [m, y] = raw.split(/[-/]/);
    return `${y}-${m}`;
  }
  if (/^\d{4}[-/]\d{1,2}$/.test(raw)) {
    const [y, m] = raw.split(/[-/]/);
    return `${y}-${String(Number(m)).padStart(2, "0")}`;
  }
  if (/^\d{1,2}[-/]\d{4}$/.test(raw)) {
    const [m, y] = raw.split(/[-/]/);
    return `${y}-${String(Number(m)).padStart(2, "0")}`;
  }
  const namedMonthMatch = raw.match(/^([a-z]{3,9})[\s,_/-]+(\d{4})$/i);
  if (namedMonthMatch) {
    const monthMap = {
      january: "01", jan: "01",
      february: "02", feb: "02",
      march: "03", mar: "03",
      april: "04", apr: "04",
      may: "05",
      june: "06", jun: "06",
      july: "07", jul: "07",
      august: "08", aug: "08",
      september: "09", sept: "09", sep: "09",
      october: "10", oct: "10",
      november: "11", nov: "11",
      december: "12", dec: "12",
    };
    const month = monthMap[namedMonthMatch[1].toLowerCase()];
    if (month) return `${namedMonthMatch[2]}-${month}`;
  }
  const reversedNamedMonthMatch = raw.match(/^(\d{4})[\s,_/-]+([a-z]{3,9})$/i);
  if (reversedNamedMonthMatch) {
    const monthMap = {
      january: "01", jan: "01",
      february: "02", feb: "02",
      march: "03", mar: "03",
      april: "04", apr: "04",
      may: "05",
      june: "06", jun: "06",
      july: "07", jul: "07",
      august: "08", aug: "08",
      september: "09", sept: "09", sep: "09",
      october: "10", oct: "10",
      november: "11", nov: "11",
      december: "12", dec: "12",
    };
    const month = monthMap[reversedNamedMonthMatch[2].toLowerCase()];
    if (month) return `${reversedNamedMonthMatch[1]}-${month}`;
  }
  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }
  return "";
}

function parseMonthFromImport(value) {
  const inferred = inferMonthValue(value);
  if (inferred) return inferred;
  return getSelectedMonth();
}

function parseIsoFromImport(value) {
  if (typeof value === "number") {
    const serialDate = excelSerialToDate(value);
    if (serialDate) {
      return `${serialDate.getUTCFullYear()}-${String(serialDate.getUTCMonth() + 1).padStart(2, "0")}-${String(serialDate.getUTCDate()).padStart(2, "0")}`;
    }
  }

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
  if (/^\d{1,12}$/.test(raw)) return true;
  if (/^[a-z0-9][a-z0-9/_-]{1,23}$/i.test(raw) && /\d/.test(raw)) return true;
  return false;
}

function isLikelyPersonName(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (raw.length < 2 || raw.length > 80) return false;
  if (/\d{3,}/.test(raw)) return false;
  return /^[a-z .'-]+$/i.test(raw);
}

function isSpreadsheetErrorValue(value) {
  const raw = String(value || "").trim().toUpperCase();
  return raw === "#REF!" || raw === "#VALUE!" || raw === "#NAME?" || raw === "#N/A" || raw === "#NULL!";
}

function rowHasSpreadsheetError(row, schema, fields) {
  return fields.some((field) => {
    const value = pickCellFromSchema(row, schema, field);
    return isSpreadsheetErrorValue(value);
  });
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

function makeImportHeaderStats(rows, headers) {
  const sampleRows = Array.isArray(rows) ? rows.slice(0, 50) : [];
  const stats = new Map();
  for (const header of headers) {
    let nonEmpty = 0;
    let moneyLike = 0;
    let monthLike = 0;
    let nameLike = 0;
    let idLike = 0;
    let absentLike = 0;
    for (const row of sampleRows) {
      const rawValue = row?.[header.original];
      if (String(rawValue ?? "").trim() === "") continue;
      nonEmpty += 1;
      if (isLikelyMoneyValue(rawValue)) moneyLike += 1;
      if (isLikelyMonthValue(rawValue)) monthLike += 1;
      if (isLikelyPersonName(rawValue)) nameLike += 1;
      if (isLikelyEmployeeId(rawValue)) idLike += 1;
      if (isLikelyDaysAbsentValue(rawValue)) absentLike += 1;
    }
    stats.set(header.original, { nonEmpty, moneyLike, monthLike, nameLike, idLike, absentLike });
  }
  return stats;
}

function buildImportSchema(rows) {
  const headerMap = collectImportHeaders(rows);
  const headers = Array.from(headerMap.entries()).map(([normalized, original]) => ({ normalized, original }));
  const sampleRows = Array.isArray(rows) ? rows.slice(0, 25) : [];
  const headerStats = makeImportHeaderStats(rows, headers);
  const fieldDefinitions = {
    employeeId: {
      aliases: ["employee id", "emp id", "employeeid", "staff id", "staff code", "id code", "code"],
      scoreValue: isLikelyEmployeeId,
      negativeAliases: ["phone", "mobile", "month", "salary", "advance", "deduction"],
    },
    employeeName: {
      aliases: ["employee name", "employee names", "name", "emp name", "staff name", "full name", "worker name"],
      scoreValue: isLikelyPersonName,
      negativeAliases: ["company name", "month", "designation", "role", "comment", "remark"],
    },
    designation: {
      aliases: ["designation", "role", "position", "job title", "title", "department role"],
      scoreValue: (value) => {
        const raw = String(value || "").trim();
        return raw.length >= 2 && raw.length <= 60 && !/\d{4,}/.test(raw);
      },
      negativeAliases: ["employee", "name", "month", "salary", "advance", "deduction"],
    },
    month: {
      aliases: ["year month", "updated month", "month", "period", "salary month", "pay month", "for month"],
      scoreValue: isLikelyMonthValue,
      negativeAliases: ["joining", "birth", "salary", "advance", "deduction"],
    },
    presentSalary: {
      aliases: ["present salary", "salary", "base salary", "monthly salary", "salary amount", "basic salary"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["increment", "advance", "deduction", "days absent", "old advance", "extra advance"],
    },
    grossSalary: {
      aliases: ["gross salary", "gross", "gross pay", "total salary before deduction"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["net salary", "present salary", "increment", "deduction", "advance remained"],
    },
    increment: {
      aliases: ["increment", "salary increment", "raise", "hike", "inc", "incr"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["salary", "advance", "deduction", "absent"],
    },
    oldAdvanceTaken: {
      aliases: ["old advance taken", "old advance", "advance taken", "previous advance", "opening advance"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["extra advance", "deduction", "salary", "increment"],
    },
    extraAdvanceAdded: {
      aliases: ["extra advance added", "extra advance", "new advance", "advance added", "fresh advance", "advance add"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["old advance", "deduction", "salary", "increment"],
    },
    deductionEntered: {
      aliases: ["deduction entered", "deduction", "advance deduction", "recovery", "recovered amount", "deducted", "recovery amount"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["days absent", "salary", "increment", "advance remained"],
    },
    totalAdvance: {
      aliases: ["total advance", "advance total", "total advance taken"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["old advance", "extra advance", "deduction", "salary"],
    },
    deductionApplied: {
      aliases: ["deduction applied", "advance deduction applied", "deduction applied advance only", "recovery applied"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["days absent", "salary", "increment", "advance remained"],
    },
    advanceRemained: {
      aliases: ["advance remained", "advance remaining", "remaining advance", "advance balance"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["deduction", "salary", "increment", "total advance"],
    },
    netSalary: {
      aliases: ["salary in hand", "net salary", "salary in hand net salary", "take home salary", "net pay"],
      scoreValue: isLikelyMoneyValue,
      negativeAliases: ["gross salary", "advance remained", "deduction", "increment"],
    },
    daysAbsent: {
      aliases: ["days absent", "absent", "attendance absent", "absent days", "leave days"],
      scoreValue: isLikelyDaysAbsentValue,
      negativeAliases: ["deduction", "salary", "advance", "month"],
    },
    comment: {
      aliases: ["comment", "notes", "remark", "remarks", "description"],
      scoreValue: (value) => String(value || "").trim().length > 0,
      negativeAliases: ["month", "salary", "advance", "deduction", "employee id"],
    },
    joiningDate: {
      aliases: ["joining date", "date of joining", "doj", "joined on"],
      scoreValue: (value) => Boolean(parseIsoFromImport(value)),
      negativeAliases: ["birth", "month", "salary", "advance"],
    },
    birthDate: {
      aliases: ["birth date", "date of birth", "dob"],
      scoreValue: (value) => Boolean(parseIsoFromImport(value)),
      negativeAliases: ["joining", "month", "salary", "advance"],
    },
    mobileNumber: {
      aliases: ["mobile", "phone", "contact", "mobile number", "phone number"],
      scoreValue: (value) => /\d{7,}/.test(String(value || "").replace(/\D/g, "")),
      negativeAliases: ["employee id", "month", "salary", "advance", "deduction"],
    },
  };

  const schema = {};
  const inferenceNotes = [];
  const candidates = [];

  for (const [field, definition] of Object.entries(fieldDefinitions)) {
    for (const header of headers) {
      let score = 0;
      let matchedBy = "content";
      for (const alias of definition.aliases) {
        const normalizedAlias = normalizeHeader(alias);
        if (header.normalized === normalizedAlias) {
          score += 18;
          matchedBy = "header";
        } else if (headerMatchesAlias(header.normalized, alias)) {
          score += 10;
          matchedBy = "header";
        }
      }
      for (const negativeAlias of definition.negativeAliases || []) {
        if (headerMatchesAlias(header.normalized, negativeAlias)) {
          score -= 8;
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
      const stats = headerStats.get(header.original);
      if (stats) {
        if (field === "month") score += Math.round(((stats.monthLike || 0) / Math.max(1, stats.nonEmpty || 1)) * 10);
        if (field === "employeeName") score += Math.round(((stats.nameLike || 0) / Math.max(1, stats.nonEmpty || 1)) * 6);
        if (field === "employeeId") score += Math.round(((stats.idLike || 0) / Math.max(1, stats.nonEmpty || 1)) * 6);
        if ([
          "presentSalary",
          "grossSalary",
          "increment",
          "oldAdvanceTaken",
          "extraAdvanceAdded",
          "totalAdvance",
          "deductionEntered",
          "deductionApplied",
          "advanceRemained",
          "netSalary",
        ].includes(field)) {
          score += Math.round(((stats.moneyLike || 0) / Math.max(1, stats.nonEmpty || 1)) * 4);
        }
        if (field === "daysAbsent") {
          score += Math.round(((stats.absentLike || 0) / Math.max(1, stats.nonEmpty || 1)) * 8);
        }
      }
      candidates.push({ field, original: header.original, score, matchedBy });
    }
  }

  const assignedHeaders = new Set();
  const sortedCandidates = candidates
    .filter((item) => item.original && item.score >= 6)
    .sort((a, b) => b.score - a.score || a.field.localeCompare(b.field));

  for (const candidate of sortedCandidates) {
    if (schema[candidate.field]) continue;
    if (assignedHeaders.has(candidate.original)) continue;
    schema[candidate.field] = candidate.original;
    assignedHeaders.add(candidate.original);
    inferenceNotes.push(`${candidate.field} -> ${candidate.original} (${candidate.matchedBy} inference, score ${candidate.score})`);
  }

  for (const field of Object.keys(fieldDefinitions)) {
    if (!schema[field]) {
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

function makeFallbackEmployeeId(employeeName, usedIds, rowIndex) {
  const base = slugify(employeeName || `employee-${rowIndex + 1}`) || `employee-${rowIndex + 1}`;
  let candidate = `imp-${base}`;
  let suffix = 2;
  while (usedIds.has(candidate)) {
    candidate = `imp-${base}-${suffix}`;
    suffix += 1;
  }
  usedIds.add(candidate);
  return candidate;
}

function analyzeLegacyRows(rows) {
  const payrollRecords = [];
  const designationSet = new Set();
  const employeeMap = new Map();
  const earliestMonthByEmployee = new Map();
  const latestMonthByEmployee = new Map();
  const monthSet = new Set();
  const { schema: detectedHeaders, inferenceNotes } = buildImportSchema(rows);
  const usedEmployeeIds = new Set();
  let skippedMissingEmployeeName = 0;
  let skippedSpreadsheetErrorRows = 0;
  let defaultedMonthCount = 0;
  let sheetInferredMonthCount = 0;
  let generatedEmployeeIdCount = 0;
  let duplicateEmployeeMonthCount = 0;
  const seenEmployeeMonthKeys = new Set();

  rows.forEach((row, rowIndex) => {
    let employeeId = String(pickCellFromSchema(row, detectedHeaders, "employeeId") || "").trim();
    const employeeName = String(pickCellFromSchema(row, detectedHeaders, "employeeName") || "").trim();
    if (!employeeName || isSpreadsheetErrorValue(employeeName) || isSpreadsheetErrorValue(employeeId)) {
      skippedMissingEmployeeName += 1;
      return;
    }
    if (rowHasSpreadsheetError(row, detectedHeaders, [
      "employeeId",
      "employeeName",
      "designation",
      "month",
      "presentSalary",
      "grossSalary",
      "increment",
      "oldAdvanceTaken",
      "extraAdvanceAdded",
      "totalAdvance",
      "deductionEntered",
      "deductionApplied",
      "advanceRemained",
      "daysAbsent",
      "comment",
    ])) {
      skippedSpreadsheetErrorRows += 1;
      return;
    }
    if (!employeeId) {
      employeeId = makeFallbackEmployeeId(employeeName, usedEmployeeIds, rowIndex);
      generatedEmployeeIdCount += 1;
    } else {
      usedEmployeeIds.add(employeeId);
    }

    const designation = String(pickCellFromSchema(row, detectedHeaders, "designation") || "").trim();
    const rawMonthValue = pickCellFromSchema(row, detectedHeaders, "month");
    const explicitMonth = inferMonthValue(rawMonthValue);
    const sheetMonth = inferMonthValue(row?.__sourceSheetMonth || row?.__sourceSheetName || "");
    const month = explicitMonth || sheetMonth || getSelectedMonth();
    const presentSalary = parseMoney(pickCellFromSchema(row, detectedHeaders, "presentSalary"));
    const grossSalary = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "grossSalary"));
    const explicitIncrement = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "increment"));
    const explicitOldAdvanceTaken = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "oldAdvanceTaken"));
    const explicitExtraAdvanceAdded = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "extraAdvanceAdded"));
    const totalAdvance = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "totalAdvance"));
    const explicitDeductionEntered = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "deductionEntered"));
    const deductionApplied = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "deductionApplied"));
    const advanceRemained = parseOptionalMoney(pickCellFromSchema(row, detectedHeaders, "advanceRemained"));
    const increment = firstFiniteNumber(
      explicitIncrement,
      grossSalary !== null ? Math.max(0, grossSalary - presentSalary) : null,
      0
    );
    const extraAdvanceAdded = firstFiniteNumber(
      explicitExtraAdvanceAdded,
      totalAdvance !== null && explicitOldAdvanceTaken !== null
        ? Math.max(0, totalAdvance - explicitOldAdvanceTaken)
        : null,
      0
    );
    const oldAdvanceTaken = firstFiniteNumber(
      explicitOldAdvanceTaken,
      totalAdvance !== null && explicitExtraAdvanceAdded !== null
        ? Math.max(0, totalAdvance - explicitExtraAdvanceAdded)
        : null
    );
    const deductionEntered = firstFiniteNumber(
      explicitDeductionEntered,
      deductionApplied,
      totalAdvance !== null && advanceRemained !== null
        ? Math.max(0, totalAdvance - advanceRemained)
        : null,
      0
    );
    const daysAbsentRaw = Number(pickCellFromSchema(row, detectedHeaders, "daysAbsent"));
    const daysAbsent = Number.isFinite(daysAbsentRaw) ? Math.min(30, Math.max(0, daysAbsentRaw)) : 0;
    const comment = String(pickCellFromSchema(row, detectedHeaders, "comment") || "").trim();
    if (!explicitMonth && sheetMonth) {
      sheetInferredMonthCount += 1;
    } else if (!explicitMonth) {
      defaultedMonthCount += 1;
    }
    const employeeMonthKey = `${employeeId}__${month}`;
    if (seenEmployeeMonthKeys.has(employeeMonthKey)) {
      duplicateEmployeeMonthCount += 1;
    }
    seenEmployeeMonthKeys.add(employeeMonthKey);

    if (designation) designationSet.add(designation);
    monthSet.add(month);
    const previousEarliestMonth = earliestMonthByEmployee.get(employeeId);
    if (!previousEarliestMonth || String(month).localeCompare(previousEarliestMonth) < 0) {
      earliestMonthByEmployee.set(employeeId, month);
    }
    const previousLatestMonth = latestMonthByEmployee.get(employeeId);
    if (!previousLatestMonth || String(month).localeCompare(previousLatestMonth) > 0) {
      latestMonthByEmployee.set(employeeId, month);
    }

    payrollRecords.push({
      month,
      employeeId,
        employeeName,
        designation,
        presentSalary,
        increment,
        oldAdvanceTaken: oldAdvanceTaken ?? "",
        extraAdvanceAdded,
        deductionEntered,
        daysAbsent,
        comment,
    });

    if (!employeeMap.has(employeeId)) {
      const importedJoiningDate = parseIsoFromImport(pickCellFromSchema(row, detectedHeaders, "joiningDate"));
      const fallbackJoiningDate = `${String(month || getSelectedMonth())}-01`;
      employeeMap.set(employeeId, {
        employeeId,
        employeeName,
        joiningDate: importedJoiningDate || fallbackJoiningDate,
        birthDate: parseIsoFromImport(pickCellFromSchema(row, detectedHeaders, "birthDate")),
        baseSalary: presentSalary,
        openingAdvance: oldAdvanceTaken,
        designation: designation || "Staff",
        mobileNumber: String(pickCellFromSchema(row, detectedHeaders, "mobileNumber") || "").trim(),
        status: "working",
      });
    } else {
      const existingEmployee = employeeMap.get(employeeId);
      const importedJoiningDate = parseIsoFromImport(pickCellFromSchema(row, detectedHeaders, "joiningDate"));
      const earliestMonth = earliestMonthByEmployee.get(employeeId) || month || getSelectedMonth();
      const latestMonth = latestMonthByEmployee.get(employeeId) || month || getSelectedMonth();
      if (existingEmployee && !existingEmployee.joiningDate) {
        existingEmployee.joiningDate = importedJoiningDate || `${String(earliestMonth)}-01`;
      }
      if (existingEmployee) {
        const existingLatestMonth = String(existingEmployee.__latestImportedMonth || "");
        if (!existingLatestMonth || String(latestMonth).localeCompare(existingLatestMonth) >= 0) {
          existingEmployee.baseSalary = presentSalary;
          existingEmployee.openingAdvance = oldAdvanceTaken ?? existingEmployee.openingAdvance ?? 0;
          existingEmployee.designation = designation || existingEmployee.designation || "Staff";
          existingEmployee.__latestImportedMonth = latestMonth;
        }
      }
    }
  });

  const reasons = [];
  if (!detectedHeaders.employeeId && generatedEmployeeIdCount === 0) {
    reasons.push("Could not detect an Employee ID column.");
  }
  if (!detectedHeaders.employeeName) {
    reasons.push("Could not detect an Employee Name column.");
  }
  if (rows.length > 0 && payrollRecords.length === 0 && skippedMissingEmployeeName > 0) {
    reasons.push(`Skipped ${skippedMissingEmployeeName} row(s) because Employee Name was empty or not recognized.`);
  }
  if (skippedSpreadsheetErrorRows > 0) {
    reasons.push(`Skipped ${skippedSpreadsheetErrorRows} row(s) because they contained spreadsheet error values like #REF!.`);
  }
  if (!detectedHeaders.month) {
    reasons.push("No Month column was detected, so the importer would fall back to the currently selected month.");
  }

  const warnings = [];
  if (defaultedMonthCount > 0) {
    warnings.push(`${defaultedMonthCount} row(s) did not include a month, so the selected app month was used.`);
  }
  if (sheetInferredMonthCount > 0) {
    warnings.push(`${sheetInferredMonthCount} row(s) used the workbook sheet name to infer the payroll month.`);
  }
  if (generatedEmployeeIdCount > 0) {
    warnings.push(`${generatedEmployeeIdCount} employee ID(s) were generated because the legacy sheet did not include usable IDs.`);
  }
  if (skippedMissingEmployeeName > 0) {
    warnings.push(`${skippedMissingEmployeeName} row(s) were skipped due to missing Employee Name.`);
  }
  if (skippedSpreadsheetErrorRows > 0) {
    warnings.push(`${skippedSpreadsheetErrorRows} row(s) were skipped because they contained spreadsheet errors like #REF!.`);
  }
  if (duplicateEmployeeMonthCount > 0) {
    warnings.push(`${duplicateEmployeeMonthCount} duplicate employee-month row(s) were found. The last row for the same employee and month is kept during import.`);
  }

  const monthSummaryMap = new Map();
  for (const row of payrollRecords) {
    const month = String(row.month || "");
    if (!monthSummaryMap.has(month)) {
      monthSummaryMap.set(month, {
        month,
        rowCount: 0,
        employeeIds: new Set(),
        presentSalary: 0,
        increment: 0,
        totalAdvance: 0,
        deductionEntered: 0,
      });
    }
    const summary = monthSummaryMap.get(month);
    summary.rowCount += 1;
    summary.employeeIds.add(String(row.employeeId || ""));
    summary.presentSalary += Number(row.presentSalary || 0);
    summary.increment += Number(row.increment || 0);
    summary.totalAdvance += Number(row.oldAdvanceTaken || 0) + Number(row.extraAdvanceAdded || 0);
    summary.deductionEntered += Number(row.deductionEntered || 0);
  }
  const monthSummaries = Array.from(monthSummaryMap.values())
    .sort((a, b) => String(a.month).localeCompare(String(b.month)))
    .map((item) => ({
      month: item.month,
      rowCount: item.rowCount,
      employeeCount: item.employeeIds.size,
      presentSalary: item.presentSalary,
      increment: item.increment,
      totalAdvance: item.totalAdvance,
      deductionEntered: item.deductionEntered,
    }));

  return {
    totalRows: rows.length,
    payrollRecords,
    designations: Array.from(designationSet),
    employees: Array.from(employeeMap.values()).map((employee) => {
      const copy = { ...employee };
      delete copy.__latestImportedMonth;
      return copy;
    }),
    months: Array.from(monthSet).sort(),
    monthSummaries,
    diagnostics: {
      detectedHeaders,
      reasons,
      warnings,
      skippedMissingEmployeeName,
      skippedSpreadsheetErrorRows,
      defaultedMonthCount,
      sheetInferredMonthCount,
      generatedEmployeeIdCount,
      duplicateEmployeeMonthCount,
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
  const byEmployeeName = new Map();
  for (const emp of existingEmployees) {
    const key = String(emp.employeeName || "").trim().toLowerCase();
    if (!key) continue;
    const bucket = byEmployeeName.get(key) || [];
    bucket.push(emp);
    byEmployeeName.set(key, bucket);
  }
  const employeeIdRemap = new Map();

  for (const incoming of analyzed.employees) {
    const incomingEmployeeId = String(incoming.employeeId || "");
    let existing = byEmployeeId.get(incomingEmployeeId);
    if (!existing) {
      const nameKey = String(incoming.employeeName || "").trim().toLowerCase();
      const nameMatches = byEmployeeName.get(nameKey) || [];
      if (nameMatches.length === 1) {
        existing = nameMatches[0];
        employeeIdRemap.set(incomingEmployeeId, String(existing.employeeId || incomingEmployeeId));
      }
    }

    const resolvedEmployeeId = String(existing?.employeeId || incomingEmployeeId);
    if (resolvedEmployeeId && resolvedEmployeeId !== incomingEmployeeId) {
      employeeIdRemap.set(incomingEmployeeId, resolvedEmployeeId);
    }

    const payload = {
      companyId,
      employeeId: resolvedEmployeeId,
      employeeName: incoming.employeeName || existing?.employeeName || "",
      joiningDate: incoming.joiningDate || existing?.joiningDate || "",
      birthDate: incoming.birthDate || existing?.birthDate || "",
      baseSalary: Number(incoming.baseSalary || 0) > 0 ? incoming.baseSalary : Number(existing?.baseSalary || 0),
      openingAdvance: Number(incoming.openingAdvance || 0) > 0 ? incoming.openingAdvance : Number(existing?.openingAdvance || 0),
      designation: incoming.designation || existing?.designation || "Staff",
      mobileNumber: incoming.mobileNumber || existing?.mobileNumber || "",
      status: incoming.status || existing?.status || "working",
      leaveFrom: "",
      leaveTo: "",
      terminatedOn: "",
      notes: existing?.notes || "",
    };

    if (existing) {
      // eslint-disable-next-line no-await-in-loop
      await apiRequest(`/api/employees/${existing.id}`, { method: "PUT", body: payload });
      byEmployeeId.set(String(payload.employeeId || ""), {
        ...existing,
        ...payload,
      });
    } else {
      // eslint-disable-next-line no-await-in-loop
      const createdResp = await apiRequest("/api/employees", { method: "POST", body: payload });
      const createdEmployee = createdResp?.employee || payload;
      byEmployeeId.set(String(createdEmployee.employeeId || payload.employeeId || ""), createdEmployee);
      const nameKey = String(createdEmployee.employeeName || payload.employeeName || "").trim().toLowerCase();
      if (nameKey) {
        const bucket = byEmployeeName.get(nameKey) || [];
        bucket.push(createdEmployee);
        byEmployeeName.set(nameKey, bucket);
      }
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
      const resolvedEmployeeId = String(employeeIdRemap.get(String(employeeId)) || employeeId);
      const existing = map.get(resolvedEmployeeId);
      // Always use the imported value for this month if present
      let oldAdvanceTaken = row.oldAdvanceTaken;
      // If missing or blank, carry forward from previous month
      if ((oldAdvanceTaken === undefined || oldAdvanceTaken === null || oldAdvanceTaken === "") && lastAdvanceByEmployee.has(resolvedEmployeeId)) {
        oldAdvanceTaken = lastAdvanceByEmployee.get(resolvedEmployeeId);
      }
      const merged = {
        ...(existing || {}),
        employeeId: resolvedEmployeeId,
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
      lastAdvanceByEmployee.set(resolvedEmployeeId, advanceRemained);
      map.set(resolvedEmployeeId, merged);
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
  headers["Cache-Control"] = "no-cache, no-store, must-revalidate";

  let response;
  try {
    response = await fetch(url, {
      method: options.method || "GET",
      headers,
      body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
    });
  } catch {
    startServerReconnectPolling();
    const message = "Server is not reachable. Start backend with `npm start` or double-click `start-routes-payroll.command`.";
    recordAppError(message, "api", { url, action: options.method || "GET" });
    throw new Error(message);
  }

  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json")
    ? await response.json()
    : null;

  stopServerReconnectPolling();

  if (!response.ok) {
    const message = payload?.error || `Request failed (${response.status}).`;
    recordAppError(message, "api", { url, status: response.status, action: options.method || "GET" });
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
  if (!SELECTORS.saveStatus) return;
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

function formatDateTime(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw;
  return date.toLocaleString("en-IN");
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
  if (!SELECTORS.authMessage) return;
  applyStatusTone(SELECTORS.authMessage, message);
  SELECTORS.authMessage.textContent = message;
}

function showAppMessage(message) {
  if (!SELECTORS.appMessage) return;
  if (isErrorLikeMessage(message)) {
    recordAppError(message, "ui");
  }
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
  if (!element?.classList) return;
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
