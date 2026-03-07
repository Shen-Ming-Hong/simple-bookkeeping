const STORAGE_KEY = "accounting_forecast_v1";
const SCHEMA_VERSION = 2;
const MAX_MONTHS = 120;
const DEFAULT_CURRENCY = "TWD";
const DEFAULT_LOCALE = "zh-TW";
const CADENCE_MONTHLY = "monthly";
const CADENCE_YEARLY = "yearly";
const APP_ID = "accounting_forecast";
const EXPORT_VERSION = 1;
const BACKUP_FILENAME_PREFIX = "accounting-backup";
const REQUIRED_BACKUP_DATA_KEYS = [
  "schemaVersion",
  "initialBalance",
  "horizonMonths",
  "recurringIncomes",
  "recurringExpenses",
  "installments",
  "currency",
  "locale",
  "updatedAt",
];

let state = createDefaultState();
let forecastChart = null;

const refs = {};
const currencyFormatter = new Intl.NumberFormat(DEFAULT_LOCALE, {
  style: "currency",
  currency: DEFAULT_CURRENCY,
  maximumFractionDigits: 0,
});

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  state = loadState();
  bindEvents();
  renderAll();
});

function createDefaultState() {
  return {
    schemaVersion: SCHEMA_VERSION,
    initialBalance: 0,
    horizonMonths: 12,
    recurringIncomes: [],
    recurringExpenses: [],
    installments: [],
    currency: DEFAULT_CURRENCY,
    locale: DEFAULT_LOCALE,
    updatedAt: new Date().toISOString(),
  };
}

function cacheDom() {
  refs.settingsForm = document.getElementById("settings-form");
  refs.initialBalance = document.getElementById("initial-balance");
  refs.horizonMonths = document.getElementById("horizon-months");
  refs.exportBackupBtn = document.getElementById("export-backup-btn");
  refs.importBackupBtn = document.getElementById("import-backup-btn");
  refs.importBackupInput = document.getElementById("import-backup-input");
  refs.backupStatus = document.getElementById("backup-status");
  refs.settingsError = document.getElementById("settings-error");

  refs.summaryMonthlyNet = document.getElementById("summary-monthly-net");
  refs.summaryEndingBalance = document.getElementById("summary-ending-balance");
  refs.summaryNegativeCount = document.getElementById("summary-negative-count");

  refs.incomeForm = document.getElementById("income-form");
  refs.incomeEditId = document.getElementById("income-edit-id");
  refs.incomeName = document.getElementById("income-name");
  refs.incomeCategory = document.getElementById("income-category");
  refs.incomeAmount = document.getElementById("income-amount");
  refs.incomeCadence = document.getElementById("income-cadence");
  refs.incomeMonthWrap = document.getElementById("income-month-wrap");
  refs.incomeMonthOfYear = document.getElementById("income-month-of-year");
  refs.incomeSubmitBtn = document.getElementById("income-submit-btn");
  refs.incomeCancelBtn = document.getElementById("income-cancel-btn");
  refs.incomeError = document.getElementById("income-error");
  refs.incomeTbody = document.getElementById("income-tbody");

  refs.expenseForm = document.getElementById("expense-form");
  refs.expenseEditId = document.getElementById("expense-edit-id");
  refs.expenseName = document.getElementById("expense-name");
  refs.expenseCategory = document.getElementById("expense-category");
  refs.expenseAmount = document.getElementById("expense-amount");
  refs.expenseCadence = document.getElementById("expense-cadence");
  refs.expenseMonthWrap = document.getElementById("expense-month-wrap");
  refs.expenseMonthOfYear = document.getElementById("expense-month-of-year");
  refs.expenseSubmitBtn = document.getElementById("expense-submit-btn");
  refs.expenseCancelBtn = document.getElementById("expense-cancel-btn");
  refs.expenseError = document.getElementById("expense-error");
  refs.expenseTbody = document.getElementById("expense-tbody");

  refs.installmentForm = document.getElementById("installment-form");
  refs.installmentEditId = document.getElementById("installment-edit-id");
  refs.installmentName = document.getElementById("installment-name");
  refs.installmentAmount = document.getElementById("installment-amount");
  refs.installmentMonths = document.getElementById("installment-months");
  refs.installmentSubmitBtn = document.getElementById("installment-submit-btn");
  refs.installmentCancelBtn = document.getElementById("installment-cancel-btn");
  refs.installmentError = document.getElementById("installment-error");
  refs.installmentTbody = document.getElementById("installment-tbody");

  refs.forecastTbody = document.getElementById("forecast-tbody");
  refs.chartCanvas = document.getElementById("forecast-chart");
  refs.chartError = document.getElementById("chart-error");
}

function bindEvents() {
  refs.settingsForm.addEventListener("submit", onSettingsSubmit);
  refs.exportBackupBtn.addEventListener("click", downloadJsonBackup);
  refs.importBackupBtn.addEventListener("click", triggerImportBackup);
  refs.importBackupInput.addEventListener("change", handleImportBackup);

  refs.incomeForm.addEventListener("submit", onIncomeSubmit);
  refs.incomeCancelBtn.addEventListener("click", resetIncomeForm);
  refs.incomeCadence.addEventListener("change", () => syncCadenceField("income"));

  refs.expenseForm.addEventListener("submit", onExpenseSubmit);
  refs.expenseCancelBtn.addEventListener("click", resetExpenseForm);
  refs.expenseCadence.addEventListener("change", () => syncCadenceField("expense"));

  refs.installmentForm.addEventListener("submit", onInstallmentSubmit);
  refs.installmentCancelBtn.addEventListener("click", resetInstallmentForm);

  syncCadenceField("income");
  syncCadenceField("expense");
}

function onSettingsSubmit(event) {
  event.preventDefault();
  clearError(refs.settingsError);

  const initialBalance = parseIntegerInput(refs.initialBalance.value);
  if (initialBalance === null || initialBalance < 0) {
    setError(refs.settingsError, "目前存款需為 0 以上整數。");
    return;
  }

  const horizonMonths = parseIntegerInput(refs.horizonMonths.value);
  if (horizonMonths === null || horizonMonths < 1 || horizonMonths > MAX_MONTHS) {
    setError(refs.settingsError, `預測月數需為 1 到 ${MAX_MONTHS} 的整數。`);
    return;
  }

  state.initialBalance = initialBalance;
  state.horizonMonths = horizonMonths;
  persistAndRender();
}

function downloadJsonBackup() {
  clearStatusMessage(refs.backupStatus);

  try {
    const payload = buildExportPayload();
    downloadBackupPayload(payload, BACKUP_FILENAME_PREFIX);
  } catch (_error) {
    setStatusMessage(refs.backupStatus, "匯出失敗：瀏覽器目前無法下載備份檔。", "error");
  }
}

function triggerImportBackup() {
  clearStatusMessage(refs.backupStatus);
  refs.importBackupInput.value = "";
  refs.importBackupInput.click();
}

async function handleImportBackup(event) {
  const input = event.target;
  const [file] = input.files || [];

  if (!file) {
    return;
  }

  clearStatusMessage(refs.backupStatus);

  try {
    const payload = await parseBackupFile(file);
    validateImportPayload(payload);

    const confirmed = window.confirm(
      `即將匯入 ${formatImportLabel(payload.exportedAt)} 的備份，這會覆蓋目前所有資料。若要保留目前資料，請先點擊「匯出備份」。是否繼續？`,
    );

    if (!confirmed) {
      setStatusMessage(refs.backupStatus, "已取消匯入，目前資料未變更。", "info");
      return;
    }

    applyImportedState(payload.data);
    setStatusMessage(refs.backupStatus, "匯入完成，目前資料已由備份檔還原。", "success");
  } catch (error) {
    setStatusMessage(refs.backupStatus, getImportErrorMessage(error), "error");
  } finally {
    input.value = "";
  }
}

function buildExportPayload() {
  const normalizedState = sanitizeState(state);
  normalizedState.schemaVersion = SCHEMA_VERSION;

  return {
    appId: APP_ID,
    exportVersion: EXPORT_VERSION,
    exportedAt: new Date().toISOString(),
    data: {
      schemaVersion: normalizedState.schemaVersion,
      initialBalance: normalizedState.initialBalance,
      horizonMonths: normalizedState.horizonMonths,
      recurringIncomes: normalizedState.recurringIncomes,
      recurringExpenses: normalizedState.recurringExpenses,
      installments: normalizedState.installments,
      currency: normalizedState.currency,
      locale: normalizedState.locale,
      updatedAt: normalizedState.updatedAt,
    },
  };
}

function downloadBackupPayload(payload, filenamePrefix) {
  if (typeof Blob === "undefined" || typeof URL === "undefined" || typeof URL.createObjectURL !== "function") {
    throw new Error("browser_unsupported");
  }

  const json = JSON.stringify(payload, null, 2);
  const blob = new Blob([json], { type: "application/json;charset=utf-8" });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = objectUrl;
  link.download = buildBackupFilename(payload.exportedAt, filenamePrefix);
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

async function parseBackupFile(file) {
  const rawText = await readFileAsText(file);

  try {
    const payload = JSON.parse(rawText);
    if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
      throw new Error("invalid_payload");
    }
    return payload;
  } catch (error) {
    if (error instanceof SyntaxError) {
      throw new Error("invalid_json");
    }
    throw error;
  }
}

function readFileAsText(file) {
  if (typeof file.text === "function") {
    return file.text();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(new Error("read_failed"));
    reader.readAsText(file, "utf-8");
  });
}

function validateImportPayload(payload) {
  if (payload.appId !== APP_ID) {
    throw new Error("invalid_app");
  }

  if (payload.exportVersion !== EXPORT_VERSION) {
    throw new Error("unsupported_export_version");
  }

  if (typeof payload.exportedAt !== "string" || !payload.exportedAt.trim()) {
    throw new Error("invalid_payload");
  }

  const { data } = payload;
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("invalid_payload");
  }

  for (const key of REQUIRED_BACKUP_DATA_KEYS) {
    if (!(key in data)) {
      throw new Error("invalid_payload");
    }
  }

  if (!Number.isInteger(data.schemaVersion)) {
    throw new Error("invalid_payload");
  }

  if (data.schemaVersion > SCHEMA_VERSION) {
    throw new Error("unsupported_schema_version");
  }

  if (
    typeof data.initialBalance !== "number" ||
    typeof data.horizonMonths !== "number" ||
    !Array.isArray(data.recurringIncomes) ||
    !Array.isArray(data.recurringExpenses) ||
    !Array.isArray(data.installments) ||
    typeof data.currency !== "string" ||
    typeof data.locale !== "string" ||
    typeof data.updatedAt !== "string"
  ) {
    throw new Error("invalid_payload");
  }
}

function applyImportedState(importedData) {
  state = sanitizeState(importedData);
  persistState();
  resetIncomeForm();
  resetExpenseForm();
  resetInstallmentForm();
  clearError(refs.settingsError);
  renderAll();
}

function getImportErrorMessage(error) {
  const code = error instanceof Error ? error.message : "";

  switch (code) {
    case "invalid_json":
      return "匯入失敗：檔案不是有效的 JSON 備份。";
    case "invalid_app":
      return "匯入失敗：這不是此網站產生的正式備份檔。";
    case "unsupported_export_version":
      return "匯入失敗：備份格式版本不支援。";
    case "unsupported_schema_version":
      return "匯入失敗：備份資料版本比目前網站新，請使用較新的版本開啟。";
    case "read_failed":
      return "匯入失敗：瀏覽器無法讀取該備份檔。";
    case "invalid_payload":
    default:
      return "匯入失敗：備份檔格式不完整或內容不正確。";
  }
}

function formatImportLabel(isoTimestamp) {
  const date = new Date(isoTimestamp);
  if (Number.isNaN(date.getTime())) {
    return "未標記時間";
  }

  const year = date.getFullYear();
  const month = padDatePart(date.getMonth() + 1);
  const day = padDatePart(date.getDate());
  const hours = padDatePart(date.getHours());
  const minutes = padDatePart(date.getMinutes());
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function buildBackupFilename(isoTimestamp, filenamePrefix = BACKUP_FILENAME_PREFIX) {
  const date = new Date(isoTimestamp);
  const year = date.getFullYear();
  const month = padDatePart(date.getMonth() + 1);
  const day = padDatePart(date.getDate());
  const hours = padDatePart(date.getHours());
  const minutes = padDatePart(date.getMinutes());
  const seconds = padDatePart(date.getSeconds());
  return `${filenamePrefix}-${year}-${month}-${day}-${hours}-${minutes}-${seconds}.json`;
}

function padDatePart(value) {
  return String(value).padStart(2, "0");
}

function onIncomeSubmit(event) {
  event.preventDefault();
  clearError(refs.incomeError);

  const name = normalizeRequiredText(refs.incomeName.value);
  if (!name) {
    setError(refs.incomeError, "收入名稱為必填。");
    return;
  }

  const amount = parseIntegerInput(refs.incomeAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.incomeError, "金額需為 0 以上整數。");
    return;
  }

  const cadence = parseCadence(refs.incomeCadence.value);
  if (!cadence) {
    setError(refs.incomeError, "請選擇有效週期。");
    return;
  }

  let monthOfYear;
  if (cadence === CADENCE_YEARLY) {
    monthOfYear = parseIntegerInput(refs.incomeMonthOfYear.value);
    if (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12) {
      setError(refs.incomeError, "年度週期需選擇 1 到 12 月。");
      return;
    }
  }

  const category = normalizeCategory(refs.incomeCategory.value);
  const editId = refs.incomeEditId.value.trim();
  const payload = {
    id: editId || makeId(),
    name,
    category,
    amount,
    cadence,
    ...(cadence === CADENCE_YEARLY ? { monthOfYear } : {}),
  };

  if (editId) {
    const found = updateById(state.recurringIncomes, payload);
    if (!found) {
      setError(refs.incomeError, "找不到要更新的收入項目，請重新操作。");
      resetIncomeForm();
      return;
    }
  } else {
    state.recurringIncomes.push(payload);
  }

  resetIncomeForm();
  persistAndRender();
}

function onExpenseSubmit(event) {
  event.preventDefault();
  clearError(refs.expenseError);

  const name = normalizeRequiredText(refs.expenseName.value);
  if (!name) {
    setError(refs.expenseError, "支出名稱為必填。");
    return;
  }

  const amount = parseIntegerInput(refs.expenseAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.expenseError, "金額需為 0 以上整數。");
    return;
  }

  const cadence = parseCadence(refs.expenseCadence.value);
  if (!cadence) {
    setError(refs.expenseError, "請選擇有效週期。");
    return;
  }

  let monthOfYear;
  if (cadence === CADENCE_YEARLY) {
    monthOfYear = parseIntegerInput(refs.expenseMonthOfYear.value);
    if (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12) {
      setError(refs.expenseError, "年度週期需選擇 1 到 12 月。");
      return;
    }
  }

  const category = normalizeCategory(refs.expenseCategory.value);
  const editId = refs.expenseEditId.value.trim();
  const payload = {
    id: editId || makeId(),
    name,
    category,
    amount,
    cadence,
    ...(cadence === CADENCE_YEARLY ? { monthOfYear } : {}),
  };

  if (editId) {
    const found = updateById(state.recurringExpenses, payload);
    if (!found) {
      setError(refs.expenseError, "找不到要更新的支出項目，請重新操作。");
      resetExpenseForm();
      return;
    }
  } else {
    state.recurringExpenses.push(payload);
  }

  resetExpenseForm();
  persistAndRender();
}

function onInstallmentSubmit(event) {
  event.preventDefault();
  clearError(refs.installmentError);

  const name = normalizeRequiredText(refs.installmentName.value);
  if (!name) {
    setError(refs.installmentError, "分期名稱為必填。");
    return;
  }

  const amount = parseIntegerInput(refs.installmentAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.installmentError, "每月金額需為 0 以上整數。");
    return;
  }

  const remainingMonths = parseIntegerInput(refs.installmentMonths.value);
  if (remainingMonths === null || remainingMonths < 1 || remainingMonths > MAX_MONTHS) {
    setError(refs.installmentError, `剩餘月數需為 1 到 ${MAX_MONTHS} 的整數。`);
    return;
  }

  const editId = refs.installmentEditId.value.trim();
  const payload = { id: editId || makeId(), name, amount, remainingMonths };

  if (editId) {
    const found = updateById(state.installments, payload);
    if (!found) {
      setError(refs.installmentError, "找不到要更新的分期項目，請重新操作。");
      resetInstallmentForm();
      return;
    }
  } else {
    state.installments.push(payload);
  }

  resetInstallmentForm();
  persistAndRender();
}

function updateById(list, nextItem) {
  const index = list.findIndex((item) => item.id === nextItem.id);
  if (index < 0) {
    return false;
  }
  list[index] = nextItem;
  return true;
}

function resetIncomeForm() {
  refs.incomeForm.reset();
  refs.incomeEditId.value = "";
  refs.incomeCadence.value = CADENCE_MONTHLY;
  refs.incomeMonthOfYear.value = "";
  syncCadenceField("income");
  refs.incomeSubmitBtn.textContent = "新增收入";
  refs.incomeCancelBtn.classList.add("hidden");
  clearError(refs.incomeError);
}

function resetExpenseForm() {
  refs.expenseForm.reset();
  refs.expenseEditId.value = "";
  refs.expenseCadence.value = CADENCE_MONTHLY;
  refs.expenseMonthOfYear.value = "";
  syncCadenceField("expense");
  refs.expenseSubmitBtn.textContent = "新增支出";
  refs.expenseCancelBtn.classList.add("hidden");
  clearError(refs.expenseError);
}

function resetInstallmentForm() {
  refs.installmentForm.reset();
  refs.installmentEditId.value = "";
  refs.installmentSubmitBtn.textContent = "新增分期";
  refs.installmentCancelBtn.classList.add("hidden");
  clearError(refs.installmentError);
}

function enterIncomeEditMode(item) {
  refs.incomeEditId.value = item.id;
  refs.incomeName.value = item.name;
  refs.incomeCategory.value = item.category;
  refs.incomeAmount.value = item.amount.toString();
  refs.incomeCadence.value = item.cadence;
  refs.incomeMonthOfYear.value = item.cadence === CADENCE_YEARLY ? String(item.monthOfYear) : "";
  syncCadenceField("income");
  refs.incomeSubmitBtn.textContent = "更新收入";
  refs.incomeCancelBtn.classList.remove("hidden");
  clearError(refs.incomeError);
}

function enterExpenseEditMode(item) {
  refs.expenseEditId.value = item.id;
  refs.expenseName.value = item.name;
  refs.expenseCategory.value = item.category;
  refs.expenseAmount.value = item.amount.toString();
  refs.expenseCadence.value = item.cadence;
  refs.expenseMonthOfYear.value = item.cadence === CADENCE_YEARLY ? String(item.monthOfYear) : "";
  syncCadenceField("expense");
  refs.expenseSubmitBtn.textContent = "更新支出";
  refs.expenseCancelBtn.classList.remove("hidden");
  clearError(refs.expenseError);
}

function enterInstallmentEditMode(item) {
  refs.installmentEditId.value = item.id;
  refs.installmentName.value = item.name;
  refs.installmentAmount.value = item.amount.toString();
  refs.installmentMonths.value = item.remainingMonths.toString();
  refs.installmentSubmitBtn.textContent = "更新分期";
  refs.installmentCancelBtn.classList.remove("hidden");
  clearError(refs.installmentError);
}

function syncCadenceField(type) {
  const isIncome = type === "income";
  const cadence = isIncome ? refs.incomeCadence.value : refs.expenseCadence.value;
  const monthWrap = isIncome ? refs.incomeMonthWrap : refs.expenseMonthWrap;
  const monthInput = isIncome ? refs.incomeMonthOfYear : refs.expenseMonthOfYear;
  const isYearly = cadence === CADENCE_YEARLY;

  monthWrap.classList.toggle("hidden", !isYearly);
  monthInput.required = isYearly;

  if (!isYearly) {
    monthInput.value = "";
  }
}

function renderAll() {
  renderSettings();
  renderIncomeTable();
  renderExpenseTable();
  renderInstallmentTable();

  const rows = calculateForecastRows();
  renderSummary(rows);
  renderForecastTable(rows);
  renderChart(rows);
}

function renderSettings() {
  refs.initialBalance.value = state.initialBalance.toString();
  refs.horizonMonths.value = state.horizonMonths.toString();
}

function renderIncomeTable() {
  refs.incomeTbody.textContent = "";
  if (state.recurringIncomes.length === 0) {
    appendEmptyRow(refs.incomeTbody, 5, "尚未新增定期收入");
    return;
  }

  state.recurringIncomes.forEach((item) => {
    const tr = document.createElement("tr");
    appendCell(tr, item.name);
    appendCell(tr, item.category);
    appendCell(tr, formatCadence(item));
    appendCell(tr, formatCurrency(item.amount));

    const actionTd = document.createElement("td");
    const actionWrap = document.createElement("div");
    actionWrap.className = "table-actions";

    const editBtn = buildActionButton("編輯", "btn-secondary", () => {
      enterIncomeEditMode(item);
      refs.incomeName.focus();
    });

    const deleteBtn = buildActionButton("刪除", "btn-danger", () => {
      if (!window.confirm(`確定刪除收入項目「${item.name}」嗎？`)) {
        return;
      }
      state.recurringIncomes = state.recurringIncomes.filter((entry) => entry.id !== item.id);
      if (refs.incomeEditId.value === item.id) {
        resetIncomeForm();
      }
      persistAndRender();
    });

    actionWrap.append(editBtn, deleteBtn);
    actionTd.appendChild(actionWrap);
    tr.appendChild(actionTd);
    refs.incomeTbody.appendChild(tr);
  });
}

function renderExpenseTable() {
  refs.expenseTbody.textContent = "";
  if (state.recurringExpenses.length === 0) {
    appendEmptyRow(refs.expenseTbody, 5, "尚未新增定期支出");
    return;
  }

  state.recurringExpenses.forEach((item) => {
    const tr = document.createElement("tr");
    appendCell(tr, item.name);
    appendCell(tr, item.category);
    appendCell(tr, formatCadence(item));
    appendCell(tr, formatCurrency(item.amount));

    const actionTd = document.createElement("td");
    const actionWrap = document.createElement("div");
    actionWrap.className = "table-actions";

    const editBtn = buildActionButton("編輯", "btn-secondary", () => {
      enterExpenseEditMode(item);
      refs.expenseName.focus();
    });

    const deleteBtn = buildActionButton("刪除", "btn-danger", () => {
      if (!window.confirm(`確定刪除支出項目「${item.name}」嗎？`)) {
        return;
      }
      state.recurringExpenses = state.recurringExpenses.filter((entry) => entry.id !== item.id);
      if (refs.expenseEditId.value === item.id) {
        resetExpenseForm();
      }
      persistAndRender();
    });

    actionWrap.append(editBtn, deleteBtn);
    actionTd.appendChild(actionWrap);
    tr.appendChild(actionTd);
    refs.expenseTbody.appendChild(tr);
  });
}

function renderInstallmentTable() {
  refs.installmentTbody.textContent = "";
  if (state.installments.length === 0) {
    appendEmptyRow(refs.installmentTbody, 4, "尚未新增分期付款");
    return;
  }

  state.installments.forEach((item) => {
    const tr = document.createElement("tr");
    appendCell(tr, item.name);
    appendCell(tr, formatCurrency(item.amount));
    appendCell(tr, `${item.remainingMonths} 個月`);

    const actionTd = document.createElement("td");
    const actionWrap = document.createElement("div");
    actionWrap.className = "table-actions";

    const editBtn = buildActionButton("編輯", "btn-secondary", () => {
      enterInstallmentEditMode(item);
      refs.installmentName.focus();
    });

    const deleteBtn = buildActionButton("刪除", "btn-danger", () => {
      if (!window.confirm(`確定刪除分期項目「${item.name}」嗎？`)) {
        return;
      }
      state.installments = state.installments.filter((entry) => entry.id !== item.id);
      if (refs.installmentEditId.value === item.id) {
        resetInstallmentForm();
      }
      persistAndRender();
    });

    actionWrap.append(editBtn, deleteBtn);
    actionTd.appendChild(actionWrap);
    tr.appendChild(actionTd);
    refs.installmentTbody.appendChild(tr);
  });
}

function renderSummary(rows) {
  const firstMonthNet = rows.length > 0 ? rows[0].net : 0;
  const endingBalance = rows.length > 0 ? rows[rows.length - 1].endingBalance : state.initialBalance;
  const negativeCount = rows.filter((row) => row.endingBalance < 0).length;

  refs.summaryMonthlyNet.textContent = formatCurrency(firstMonthNet);
  refs.summaryEndingBalance.textContent = formatCurrency(endingBalance);
  refs.summaryNegativeCount.textContent = `${negativeCount} 個月`;
}

function renderForecastTable(rows) {
  refs.forecastTbody.textContent = "";
  if (rows.length === 0) {
    appendEmptyRow(refs.forecastTbody, 8, "無預測資料");
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.endingBalance < 0) {
      tr.classList.add("negative-row");
    }

    appendCell(tr, row.monthLabel);
    appendCell(tr, formatCurrency(row.startingBalance));
    appendCell(tr, formatCurrency(row.incomeTotal));
    appendCell(tr, formatCurrency(row.expenseTotal));
    appendCell(tr, formatCurrency(row.installmentTotal));
    appendCell(tr, formatCurrency(row.net));
    appendCell(tr, formatCurrency(row.endingBalance));

    const statusTd = document.createElement("td");
    const chip = document.createElement("span");
    chip.className = `status-chip ${row.endingBalance < 0 ? "warn" : "ok"}`;
    chip.textContent = row.endingBalance < 0 ? "警示：資金不足" : "正常";
    statusTd.appendChild(chip);
    tr.appendChild(statusTd);

    refs.forecastTbody.appendChild(tr);
  });
}

function renderChart(rows) {
  clearError(refs.chartError);

  if (typeof window.Chart === "undefined") {
    setError(refs.chartError, "圖表載入失敗，請檢查網路後重整頁面。");
    if (forecastChart) {
      forecastChart.destroy();
      forecastChart = null;
    }
    return;
  }

  const labels = rows.map((row) => row.monthLabel);
  const values = rows.map((row) => row.endingBalance);
  const pointColors = rows.map((row) => (row.endingBalance < 0 ? "#b42318" : "#0f766e"));
  const borderColors = rows.map((row) => (row.endingBalance < 0 ? "#f59e0b" : "#0f766e"));

  if (forecastChart) {
    forecastChart.data.labels = labels;
    forecastChart.data.datasets[0].data = values;
    forecastChart.data.datasets[0].pointBackgroundColor = pointColors;
    forecastChart.data.datasets[0].pointBorderColor = borderColors;
    forecastChart.update();
    return;
  }

  forecastChart = new window.Chart(refs.chartCanvas.getContext("2d"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "月末餘額",
          data: values,
          borderColor: "#0f766e",
          pointBackgroundColor: pointColors,
          pointBorderColor: borderColors,
          pointRadius: 4,
          pointHoverRadius: 6,
          borderWidth: 2.5,
          fill: true,
          tension: 0.25,
          backgroundColor: "rgba(15, 118, 110, 0.12)",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `月末餘額：${formatCurrency(context.raw)}`;
            },
          },
        },
      },
      scales: {
        y: {
          ticks: {
            callback(value) {
              return formatCurrency(value);
            },
          },
        },
      },
    },
  });
}

function calculateForecastRows() {
  const rows = [];
  const today = new Date();
  let prevEndingBalance = state.initialBalance;

  for (let i = 0; i < state.horizonMonths; i += 1) {
    const date = new Date(today.getFullYear(), today.getMonth() + i, 1);
    const incomeTotal = sumActiveAmounts(state.recurringIncomes, date);
    const expenseTotal = sumActiveAmounts(state.recurringExpenses, date);
    const installmentTotal = state.installments.reduce((sum, installment) => {
      return sum + (i < installment.remainingMonths ? installment.amount : 0);
    }, 0);

    const net = incomeTotal - expenseTotal - installmentTotal;
    const startingBalance = i === 0 ? state.initialBalance : prevEndingBalance;
    const endingBalance = startingBalance + net;

    rows.push({
      monthLabel: formatMonth(date),
      startingBalance,
      incomeTotal,
      expenseTotal,
      installmentTotal,
      net,
      endingBalance,
    });

    prevEndingBalance = endingBalance;
  }

  return rows;
}

function sumActiveAmounts(items, date) {
  return items.reduce((sum, item) => {
    return sum + (isItemActiveInMonth(item, date) ? item.amount : 0);
  }, 0);
}

function isItemActiveInMonth(item, date) {
  if (item.cadence === CADENCE_YEARLY) {
    return item.monthOfYear === date.getMonth() + 1;
  }
  return true;
}

function formatMonth(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}年${month}月`;
}

function formatCurrency(value) {
  return currencyFormatter.format(value);
}

function formatCadence(item) {
  if (item.cadence === CADENCE_YEARLY) {
    return `每年 ${item.monthOfYear} 月`;
  }
  return "每月";
}

function appendCell(row, value) {
  const td = document.createElement("td");
  td.textContent = value;
  row.appendChild(td);
  return td;
}

function appendEmptyRow(tbody, colSpan, message) {
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.colSpan = colSpan;
  td.textContent = message;
  tr.appendChild(td);
  tbody.appendChild(tr);
}

function buildActionButton(label, styleClass, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `btn ${styleClass}`;
  button.textContent = label;
  button.addEventListener("click", onClick);
  return button;
}

function parseIntegerInput(rawValue) {
  const value = String(rawValue).trim();
  if (!/^\d+$/.test(value)) {
    return null;
  }

  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) {
    return null;
  }

  return parsed;
}

function parseCadence(rawValue) {
  if (rawValue === CADENCE_MONTHLY || rawValue === CADENCE_YEARLY) {
    return rawValue;
  }
  return null;
}

function normalizeRequiredText(rawValue) {
  const value = String(rawValue ?? "").trim();
  return value.length > 0 ? value : "";
}

function normalizeCategory(rawValue) {
  const value = String(rawValue ?? "").trim();
  return value.length > 0 ? value : "未分類";
}

function makeId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function setError(element, message) {
  element.textContent = message;
}

function clearError(element) {
  element.textContent = "";
}

function setStatusMessage(element, message, type) {
  element.textContent = message;
  element.classList.remove("is-error", "is-success", "is-info");

  if (message) {
    element.classList.add(`is-${type}`);
  }
}

function clearStatusMessage(element) {
  element.textContent = "";
  element.classList.remove("is-error", "is-success", "is-info");
}

function persistAndRender() {
  persistState();
  renderAll();
}

function persistState() {
  state.schemaVersion = SCHEMA_VERSION;
  state.updatedAt = new Date().toISOString();
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (_error) {
    setError(refs.settingsError, "儲存失敗：瀏覽器無法寫入本地資料。");
  }
}

function persistMigratedState(nextState) {
  const migrated = {
    ...nextState,
    schemaVersion: SCHEMA_VERSION,
    updatedAt: new Date().toISOString(),
  };

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(migrated));
  } catch (_error) {
    return nextState;
  }

  return migrated;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return createDefaultState();
    }

    const parsed = JSON.parse(raw);
    const isV2 = parsed && typeof parsed === "object" && parsed.schemaVersion === SCHEMA_VERSION;
    const source = isV2 ? parsed : migrateStateToV2(parsed);
    const sanitized = sanitizeState(source);

    if (!isV2) {
      return persistMigratedState(sanitized);
    }

    return sanitized;
  } catch (_error) {
    return createDefaultState();
  }
}

function migrateStateToV2(legacyState) {
  const migrated = createDefaultState();
  if (!legacyState || typeof legacyState !== "object") {
    return migrated;
  }

  migrated.initialBalance = legacyState.initialBalance;
  migrated.horizonMonths = legacyState.horizonMonths;
  migrated.recurringIncomes = migrateLegacyItems(legacyState.recurringIncomes);
  migrated.recurringExpenses = migrateLegacyItems(legacyState.recurringExpenses);
  migrated.installments = Array.isArray(legacyState.installments) ? legacyState.installments : [];
  migrated.updatedAt = typeof legacyState.updatedAt === "string"
    ? legacyState.updatedAt
    : migrated.updatedAt;

  return migrated;
}

function migrateLegacyItems(input) {
  if (!Array.isArray(input)) {
    return [];
  }

  return input.map((item) => {
    if (!item || typeof item !== "object") {
      return item;
    }

    return {
      ...item,
      cadence: CADENCE_MONTHLY,
      monthOfYear: undefined,
    };
  });
}

function sanitizeState(input) {
  const safeState = createDefaultState();
  if (!input || typeof input !== "object") {
    return safeState;
  }

  safeState.schemaVersion = SCHEMA_VERSION;
  safeState.initialBalance = sanitizeBoundedInteger(input.initialBalance, 0, Number.MAX_SAFE_INTEGER, 0);
  safeState.horizonMonths = sanitizeBoundedInteger(input.horizonMonths, 1, MAX_MONTHS, 12);
  safeState.recurringIncomes = sanitizeItems(input.recurringIncomes);
  safeState.recurringExpenses = sanitizeItems(input.recurringExpenses);
  safeState.installments = sanitizeInstallments(input.installments);
  safeState.currency = DEFAULT_CURRENCY;
  safeState.locale = DEFAULT_LOCALE;
  safeState.updatedAt = typeof input.updatedAt === "string" ? input.updatedAt : new Date().toISOString();

  return safeState;
}

function sanitizeItems(input) {
  if (!Array.isArray(input)) {
    return [];
  }

  return input
    .map((rawItem) => {
      if (!rawItem || typeof rawItem !== "object") {
        return null;
      }

      const name = normalizeRequiredText(rawItem.name);
      if (!name) {
        return null;
      }

      const amount = sanitizeBoundedInteger(rawItem.amount, 0, Number.MAX_SAFE_INTEGER, null);
      if (amount === null) {
        return null;
      }

      const cadence = parseCadence(rawItem.cadence) || CADENCE_MONTHLY;
      const monthOfYear = cadence === CADENCE_YEARLY
        ? sanitizeBoundedInteger(rawItem.monthOfYear, 1, 12, null)
        : null;

      if (cadence === CADENCE_YEARLY && monthOfYear === null) {
        return null;
      }

      const item = {
        id: typeof rawItem.id === "string" && rawItem.id.trim() ? rawItem.id : makeId(),
        name,
        category: normalizeCategory(rawItem.category),
        amount,
        cadence,
      };

      if (monthOfYear !== null) {
        item.monthOfYear = monthOfYear;
      }

      return item;
    })
    .filter(Boolean);
}

function sanitizeInstallments(input) {
  if (!Array.isArray(input)) {
    return [];
  }

  return input
    .map((rawItem) => {
      if (!rawItem || typeof rawItem !== "object") {
        return null;
      }

      const name = normalizeRequiredText(rawItem.name);
      if (!name) {
        return null;
      }

      const amount = sanitizeBoundedInteger(rawItem.amount, 0, Number.MAX_SAFE_INTEGER, null);
      const remainingMonths = sanitizeBoundedInteger(rawItem.remainingMonths, 1, MAX_MONTHS, null);
      if (amount === null || remainingMonths === null) {
        return null;
      }

      return {
        id: typeof rawItem.id === "string" && rawItem.id.trim() ? rawItem.id : makeId(),
        name,
        amount,
        remainingMonths,
      };
    })
    .filter(Boolean);
}

function sanitizeBoundedInteger(value, min, max, fallback) {
  const parsed = parseIntegerInput(String(value ?? ""));
  if (parsed === null || parsed < min || parsed > max) {
    return fallback;
  }
  return parsed;
}
