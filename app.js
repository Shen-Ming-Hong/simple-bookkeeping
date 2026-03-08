const STORAGE_KEY = "accounting_forecast_v1";
const SCHEMA_VERSION = 3;
const EXPORT_VERSION = 2;
const LEGACY_EXPORT_VERSION = 1;
const MAX_MONTHS = 120;
const DEFAULT_CURRENCY = "TWD";
const DEFAULT_LOCALE = "zh-TW";
const CADENCE_MONTHLY = "monthly";
const CADENCE_YEARLY = "yearly";
const CADENCE_ONE_TIME = "one-time";
const APP_ID = "accounting_forecast";
const BACKUP_FILENAME_PREFIX = "accounting-backup";
const MAIN_ACCOUNT_NAME = "主帳戶";
const BATCH_TYPES = ["income", "expense", "installment"];
const LEGACY_BACKUP_DATA_KEYS = [
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
const CURRENT_BACKUP_DATA_KEYS = [
  "schemaVersion",
  "accounts",
  "horizonMonths",
  "recurringIncomes",
  "recurringExpenses",
  "installments",
  "monthlyTransfers",
  "currency",
  "locale",
  "updatedAt",
];

let state = createDefaultState();
let forecastChart = null;

const refs = {};
const batchState = createBatchState();

if (typeof document !== "undefined") {
  document.addEventListener("DOMContentLoaded", () => {
    cacheDom();
    state = loadState();
    bindEvents();
    refreshStaticAccountSelects();
    resetAllForms();
    renderAll();
  });
}

function createBatchState() {
  return {
    income: createEmptyBatchBucket(),
    expense: createEmptyBatchBucket(),
    installment: createEmptyBatchBucket(),
  };
}

function createEmptyBatchBucket() {
  return {
    selectedIds: new Set(),
    isOpen: false,
    pendingAccounts: {},
  };
}

function createDefaultState() {
  return {
    schemaVersion: SCHEMA_VERSION,
    accounts: [createAccount(MAIN_ACCOUNT_NAME, 0)],
    horizonMonths: 12,
    recurringIncomes: [],
    recurringExpenses: [],
    installments: [],
    monthlyTransfers: [],
    currency: DEFAULT_CURRENCY,
    locale: DEFAULT_LOCALE,
    updatedAt: new Date().toISOString(),
  };
}

function createAccount(name, initialBalance) {
  return {
    id: makeId(),
    name,
    initialBalance,
  };
}

function cacheDom() {
  refs.settingsForm = document.getElementById("settings-form");
  refs.horizonMonths = document.getElementById("horizon-months");
  refs.exportBackupBtn = document.getElementById("export-backup-btn");
  refs.importBackupBtn = document.getElementById("import-backup-btn");
  refs.importBackupInput = document.getElementById("import-backup-input");
  refs.backupStatus = document.getElementById("backup-status");
  refs.settingsError = document.getElementById("settings-error");

  refs.summaryMonthlyNet = document.getElementById("summary-monthly-net");
  refs.summaryEndingBalance = document.getElementById("summary-ending-balance");
  refs.summaryNegativeTotalCount = document.getElementById("summary-negative-total-count");
  refs.summaryNegativeAccountCount = document.getElementById("summary-negative-account-count");

  refs.accountForm = document.getElementById("account-form");
  refs.accountEditId = document.getElementById("account-edit-id");
  refs.accountName = document.getElementById("account-name");
  refs.accountBalance = document.getElementById("account-balance");
  refs.accountSubmitBtn = document.getElementById("account-submit-btn");
  refs.accountCancelBtn = document.getElementById("account-cancel-btn");
  refs.accountError = document.getElementById("account-error");
  refs.accountTbody = document.getElementById("account-tbody");

  refs.incomeForm = document.getElementById("income-form");
  refs.incomeEditId = document.getElementById("income-edit-id");
  refs.incomeName = document.getElementById("income-name");
  refs.incomeAmount = document.getElementById("income-amount");
  refs.incomeAccountId = document.getElementById("income-account-id");
  refs.incomeCadence = document.getElementById("income-cadence");
  refs.incomeMonthWrap = document.getElementById("income-month-wrap");
  refs.incomeMonthOfYear = document.getElementById("income-month-of-year");
  refs.incomeSubmitBtn = document.getElementById("income-submit-btn");
  refs.incomeCancelBtn = document.getElementById("income-cancel-btn");
  refs.incomeError = document.getElementById("income-error");
  refs.incomeSelectAll = document.getElementById("income-select-all");
  refs.incomeBatchBar = document.getElementById("income-batch-bar");
  refs.incomeSelectedCount = document.getElementById("income-selected-count");
  refs.incomeBatchBtn = document.getElementById("income-batch-btn");
  refs.incomeClearBtn = document.getElementById("income-clear-btn");
  refs.incomeBatchPanel = document.getElementById("income-batch-panel");
  refs.incomeBatchList = document.getElementById("income-batch-list");
  refs.incomeBatchApplyBtn = document.getElementById("income-batch-apply-btn");
  refs.incomeBatchCancelBtn = document.getElementById("income-batch-cancel-btn");
  refs.incomeTbody = document.getElementById("income-tbody");

  refs.expenseForm = document.getElementById("expense-form");
  refs.expenseEditId = document.getElementById("expense-edit-id");
  refs.expenseName = document.getElementById("expense-name");
  refs.expenseAmount = document.getElementById("expense-amount");
  refs.expenseAccountId = document.getElementById("expense-account-id");
  refs.expenseCadence = document.getElementById("expense-cadence");
  refs.expenseMonthWrap = document.getElementById("expense-month-wrap");
  refs.expenseMonthOfYear = document.getElementById("expense-month-of-year");
  refs.expenseOneTimeWrap = document.getElementById("expense-one-time-wrap");
  refs.expenseOneTimeMonth = document.getElementById("expense-one-time-month");
  refs.expenseSubmitBtn = document.getElementById("expense-submit-btn");
  refs.expenseCancelBtn = document.getElementById("expense-cancel-btn");
  refs.expenseError = document.getElementById("expense-error");
  refs.expenseSelectAll = document.getElementById("expense-select-all");
  refs.expenseBatchBar = document.getElementById("expense-batch-bar");
  refs.expenseSelectedCount = document.getElementById("expense-selected-count");
  refs.expenseBatchBtn = document.getElementById("expense-batch-btn");
  refs.expenseClearBtn = document.getElementById("expense-clear-btn");
  refs.expenseBatchPanel = document.getElementById("expense-batch-panel");
  refs.expenseBatchList = document.getElementById("expense-batch-list");
  refs.expenseBatchApplyBtn = document.getElementById("expense-batch-apply-btn");
  refs.expenseBatchCancelBtn = document.getElementById("expense-batch-cancel-btn");
  refs.expenseTbody = document.getElementById("expense-tbody");

  refs.installmentForm = document.getElementById("installment-form");
  refs.installmentEditId = document.getElementById("installment-edit-id");
  refs.installmentName = document.getElementById("installment-name");
  refs.installmentAmount = document.getElementById("installment-amount");
  refs.installmentMonths = document.getElementById("installment-months");
  refs.installmentAccountId = document.getElementById("installment-account-id");
  refs.installmentSubmitBtn = document.getElementById("installment-submit-btn");
  refs.installmentCancelBtn = document.getElementById("installment-cancel-btn");
  refs.installmentError = document.getElementById("installment-error");
  refs.installmentSelectAll = document.getElementById("installment-select-all");
  refs.installmentBatchBar = document.getElementById("installment-batch-bar");
  refs.installmentSelectedCount = document.getElementById("installment-selected-count");
  refs.installmentBatchBtn = document.getElementById("installment-batch-btn");
  refs.installmentClearBtn = document.getElementById("installment-clear-btn");
  refs.installmentBatchPanel = document.getElementById("installment-batch-panel");
  refs.installmentBatchList = document.getElementById("installment-batch-list");
  refs.installmentBatchApplyBtn = document.getElementById("installment-batch-apply-btn");
  refs.installmentBatchCancelBtn = document.getElementById("installment-batch-cancel-btn");
  refs.installmentTbody = document.getElementById("installment-tbody");

  refs.transferForm = document.getElementById("transfer-form");
  refs.transferEditId = document.getElementById("transfer-edit-id");
  refs.transferName = document.getElementById("transfer-name");
  refs.transferAmount = document.getElementById("transfer-amount");
  refs.transferSourceAccountId = document.getElementById("transfer-source-account-id");
  refs.transferTargetAccountId = document.getElementById("transfer-target-account-id");
  refs.transferSubmitBtn = document.getElementById("transfer-submit-btn");
  refs.transferCancelBtn = document.getElementById("transfer-cancel-btn");
  refs.transferError = document.getElementById("transfer-error");
  refs.transferTbody = document.getElementById("transfer-tbody");

  refs.forecastTbody = document.getElementById("forecast-tbody");
  refs.accountForecastTbody = document.getElementById("account-forecast-tbody");
  refs.chartCanvas = document.getElementById("forecast-chart");
  refs.chartError = document.getElementById("chart-error");
}

function bindEvents() {
  refs.settingsForm.addEventListener("submit", onSettingsSubmit);
  refs.exportBackupBtn.addEventListener("click", downloadJsonBackup);
  refs.importBackupBtn.addEventListener("click", triggerImportBackup);
  refs.importBackupInput.addEventListener("change", handleImportBackup);

  refs.accountForm.addEventListener("submit", onAccountSubmit);
  refs.accountCancelBtn.addEventListener("click", resetAccountForm);

  refs.incomeForm.addEventListener("submit", onIncomeSubmit);
  refs.incomeCancelBtn.addEventListener("click", resetIncomeForm);
  refs.incomeCadence.addEventListener("change", syncIncomeCadenceField);

  refs.expenseForm.addEventListener("submit", onExpenseSubmit);
  refs.expenseCancelBtn.addEventListener("click", resetExpenseForm);
  refs.expenseCadence.addEventListener("change", syncExpenseCadenceField);

  refs.installmentForm.addEventListener("submit", onInstallmentSubmit);
  refs.installmentCancelBtn.addEventListener("click", resetInstallmentForm);

  refs.transferForm.addEventListener("submit", onTransferSubmit);
  refs.transferCancelBtn.addEventListener("click", resetTransferForm);
  refs.transferSourceAccountId.addEventListener("change", () => {
    syncTransferTargetSelection();
  });

  bindBatchEvents("income");
  bindBatchEvents("expense");
  bindBatchEvents("installment");
}

function bindBatchEvents(type) {
  const batchRefs = getBatchRefs(type);
  batchRefs.selectAll.addEventListener("change", () => {
    toggleSelectAll(type, batchRefs.selectAll.checked);
  });
  batchRefs.batchBtn.addEventListener("click", () => openBatchPanel(type));
  batchRefs.clearBtn.addEventListener("click", () => clearBatchSelection(type));
  batchRefs.applyBtn.addEventListener("click", () => applyBatchAssignments(type));
  batchRefs.cancelBtn.addEventListener("click", () => {
    closeBatchPanel(type);
    renderAll();
  });
}

function resetAllForms() {
  resetAccountForm();
  resetIncomeForm();
  resetExpenseForm();
  resetInstallmentForm();
  resetTransferForm();
}

function refreshStaticAccountSelects() {
  const defaultAccountId = getDefaultAccountId();
  setAccountSelectOptions(
    refs.incomeAccountId,
    getValidAccountId(refs.incomeAccountId.value) || defaultAccountId
  );
  setAccountSelectOptions(
    refs.expenseAccountId,
    getValidAccountId(refs.expenseAccountId.value) || defaultAccountId
  );
  setAccountSelectOptions(
    refs.installmentAccountId,
    getValidAccountId(refs.installmentAccountId.value) || defaultAccountId
  );
  setAccountSelectOptions(
    refs.transferSourceAccountId,
    getValidAccountId(refs.transferSourceAccountId.value) || defaultAccountId
  );
  syncTransferTargetSelection(refs.transferTargetAccountId.value);
}

function setAccountSelectOptions(selectElement, preferredValue) {
  if (!selectElement) {
    return;
  }

  const fragment = document.createDocumentFragment();
  state.accounts.forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.name;
    fragment.appendChild(option);
  });

  selectElement.replaceChildren(fragment);
  selectElement.disabled = state.accounts.length === 0;
  selectElement.value = getValidAccountId(preferredValue) || getDefaultAccountId();
}

function syncTransferTargetSelection(preferredTargetId) {
  if (!refs.transferSourceAccountId || !refs.transferTargetAccountId) {
    return;
  }

  const sourceAccountId = getValidAccountId(refs.transferSourceAccountId.value) || getDefaultAccountId();
  refs.transferSourceAccountId.value = sourceAccountId;

  const candidates = state.accounts.filter((account) => account.id !== sourceAccountId);
  refs.transferTargetAccountId.replaceChildren();

  if (!candidates.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "請先新增第二個帳戶";
    refs.transferTargetAccountId.appendChild(option);
    refs.transferTargetAccountId.value = "";
    refs.transferTargetAccountId.disabled = true;
    return;
  }

  refs.transferTargetAccountId.disabled = false;
  candidates.forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.name;
    refs.transferTargetAccountId.appendChild(option);
  });

  const nextTargetId = candidates.some((account) => account.id === preferredTargetId)
    ? preferredTargetId
    : candidates[0].id;
  refs.transferTargetAccountId.value = nextTargetId;
}

function onSettingsSubmit(event) {
  event.preventDefault();
  clearError(refs.settingsError);
  clearStatusMessage(refs.backupStatus);

  const horizonMonths = parseIntegerInput(refs.horizonMonths.value);
  if (horizonMonths === null || horizonMonths < 1 || horizonMonths > MAX_MONTHS) {
    setError(refs.settingsError, `預測月數需為 1 到 ${MAX_MONTHS} 的整數。`);
    return;
  }

  state.horizonMonths = horizonMonths;
  persistAndRender();
  setStatusMessage(refs.backupStatus, "已更新預測月數。", "success");
}

function onAccountSubmit(event) {
  event.preventDefault();
  clearError(refs.accountError);

  const name = normalizeRequiredText(refs.accountName.value);
  if (!name) {
    setError(refs.accountError, "請輸入帳戶名稱。");
    return;
  }

  const initialBalance = parseIntegerInput(refs.accountBalance.value);
  if (initialBalance === null || initialBalance < 0) {
    setError(refs.accountError, "期初餘額需為 0 以上的整數。");
    return;
  }

  const editId = refs.accountEditId.value.trim();
  if (isDuplicateAccountName(name, editId)) {
    setError(refs.accountError, "帳戶名稱不可重複。");
    return;
  }

  if (editId) {
    const account = getAccountById(editId);
    if (!account) {
      setError(refs.accountError, "找不到要更新的帳戶，請重新操作。");
      resetAccountForm();
      return;
    }
    account.name = name;
    account.initialBalance = initialBalance;
  } else {
    state.accounts.push({
      id: makeId(),
      name,
      initialBalance,
    });
  }

  resetAccountForm();
  persistAndRender();
}

function onIncomeSubmit(event) {
  event.preventDefault();
  clearError(refs.incomeError);

  const name = normalizeRequiredText(refs.incomeName.value);
  if (!name) {
    setError(refs.incomeError, "請輸入收入名稱。");
    return;
  }

  const amount = parseIntegerInput(refs.incomeAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.incomeError, "收入金額需為 0 以上的整數。");
    return;
  }

  const accountId = getValidAccountId(refs.incomeAccountId.value);
  if (!accountId) {
    setError(refs.incomeError, "請選擇帳戶。");
    return;
  }

  const cadence = parseIncomeCadence(refs.incomeCadence.value);
  if (!cadence) {
    setError(refs.incomeError, "請選擇正確的收入週期。");
    return;
  }

  const monthOfYear =
    cadence === CADENCE_YEARLY
      ? parseIntegerInput(refs.incomeMonthOfYear.value)
      : null;
  if (cadence === CADENCE_YEARLY && (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12)) {
    setError(refs.incomeError, "年度收入需要指定 1 到 12 月。");
    return;
  }

  const payload = {
    id: refs.incomeEditId.value.trim() || makeId(),
    name,
    amount,
    accountId,
    cadence,
    monthOfYear: cadence === CADENCE_YEARLY ? monthOfYear : undefined,
  };

  if (refs.incomeEditId.value.trim()) {
    if (!updateById(state.recurringIncomes, payload)) {
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
    setError(refs.expenseError, "請輸入支出名稱。");
    return;
  }

  const amount = parseIntegerInput(refs.expenseAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.expenseError, "支出金額需為 0 以上的整數。");
    return;
  }

  const accountId = getValidAccountId(refs.expenseAccountId.value);
  if (!accountId) {
    setError(refs.expenseError, "請選擇帳戶。");
    return;
  }

  const cadence = parseExpenseCadence(refs.expenseCadence.value);
  if (!cadence) {
    setError(refs.expenseError, "請選擇正確的支出週期。");
    return;
  }

  let monthOfYear;
  let year;

  if (cadence === CADENCE_YEARLY) {
    monthOfYear = parseIntegerInput(refs.expenseMonthOfYear.value);
    if (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12) {
      setError(refs.expenseError, "年度支出需要指定 1 到 12 月。");
      return;
    }
  }

  if (cadence === CADENCE_ONE_TIME) {
    const parsed = parseOneTimeMonth(refs.expenseOneTimeMonth.value);
    if (!parsed) {
      setError(refs.expenseError, "一次性支出需要指定有效的年月。");
      return;
    }
    year = parsed.year;
    monthOfYear = parsed.monthOfYear;
  }

  const payload = {
    id: refs.expenseEditId.value.trim() || makeId(),
    name,
    amount,
    accountId,
    cadence,
    monthOfYear,
    year,
  };

  if (refs.expenseEditId.value.trim()) {
    if (!updateById(state.recurringExpenses, payload)) {
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
    setError(refs.installmentError, "請輸入分期名稱。");
    return;
  }

  const amount = parseIntegerInput(refs.installmentAmount.value);
  if (amount === null || amount < 0) {
    setError(refs.installmentError, "分期金額需為 0 以上的整數。");
    return;
  }

  const remainingMonths = parseIntegerInput(refs.installmentMonths.value);
  if (remainingMonths === null || remainingMonths < 1 || remainingMonths > MAX_MONTHS) {
    setError(refs.installmentError, `剩餘月數需為 1 到 ${MAX_MONTHS} 的整數。`);
    return;
  }

  const accountId = getValidAccountId(refs.installmentAccountId.value);
  if (!accountId) {
    setError(refs.installmentError, "請選擇帳戶。");
    return;
  }

  const payload = {
    id: refs.installmentEditId.value.trim() || makeId(),
    name,
    amount,
    remainingMonths,
    accountId,
  };

  if (refs.installmentEditId.value.trim()) {
    if (!updateById(state.installments, payload)) {
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

function onTransferSubmit(event) {
  event.preventDefault();
  clearError(refs.transferError);

  const name = normalizeRequiredText(refs.transferName.value);
  if (!name) {
    setError(refs.transferError, "請輸入轉帳名稱。");
    return;
  }

  const amount = parseIntegerInput(refs.transferAmount.value);
  if (amount === null || amount <= 0) {
    setError(refs.transferError, "每月轉帳金額需為大於 0 的整數。");
    return;
  }

  const sourceAccountId = getValidAccountId(refs.transferSourceAccountId.value);
  const targetAccountId = getValidAccountId(refs.transferTargetAccountId.value);
  if (!sourceAccountId || !targetAccountId) {
    setError(refs.transferError, "轉帳需要同時指定來源與目標帳戶。");
    return;
  }
  if (sourceAccountId === targetAccountId) {
    setError(refs.transferError, "來源與目標帳戶不可相同。");
    return;
  }

  const payload = {
    id: refs.transferEditId.value.trim() || makeId(),
    name,
    amount,
    sourceAccountId,
    targetAccountId,
  };

  if (refs.transferEditId.value.trim()) {
    if (!updateById(state.monthlyTransfers, payload)) {
      setError(refs.transferError, "找不到要更新的轉帳項目，請重新操作。");
      resetTransferForm();
      return;
    }
  } else {
    state.monthlyTransfers.push(payload);
  }

  resetTransferForm();
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

function resetAccountForm() {
  refs.accountForm.reset();
  refs.accountEditId.value = "";
  refs.accountSubmitBtn.textContent = "新增帳戶";
  refs.accountCancelBtn.classList.add("hidden");
  clearError(refs.accountError);
}

function resetIncomeForm() {
  refs.incomeForm.reset();
  refs.incomeEditId.value = "";
  refs.incomeCadence.value = CADENCE_MONTHLY;
  refs.incomeMonthOfYear.value = "";
  refs.incomeAccountId.value = getDefaultAccountId();
  syncIncomeCadenceField();
  refs.incomeSubmitBtn.textContent = "新增收入";
  refs.incomeCancelBtn.classList.add("hidden");
  clearError(refs.incomeError);
}

function resetExpenseForm() {
  refs.expenseForm.reset();
  refs.expenseEditId.value = "";
  refs.expenseCadence.value = CADENCE_MONTHLY;
  refs.expenseMonthOfYear.value = "";
  refs.expenseOneTimeMonth.value = "";
  refs.expenseAccountId.value = getDefaultAccountId();
  syncExpenseCadenceField();
  refs.expenseSubmitBtn.textContent = "新增支出";
  refs.expenseCancelBtn.classList.add("hidden");
  clearError(refs.expenseError);
}

function resetInstallmentForm() {
  refs.installmentForm.reset();
  refs.installmentEditId.value = "";
  refs.installmentAccountId.value = getDefaultAccountId();
  refs.installmentSubmitBtn.textContent = "新增分期";
  refs.installmentCancelBtn.classList.add("hidden");
  clearError(refs.installmentError);
}

function resetTransferForm() {
  refs.transferForm.reset();
  refs.transferEditId.value = "";
  refs.transferSourceAccountId.value = getDefaultAccountId();
  syncTransferTargetSelection(getDefaultTransferTargetId(refs.transferSourceAccountId.value));
  refs.transferSubmitBtn.textContent = "新增轉帳";
  refs.transferCancelBtn.classList.add("hidden");
  clearError(refs.transferError);
}

function enterAccountEditMode(account) {
  refs.accountEditId.value = account.id;
  refs.accountName.value = account.name;
  refs.accountBalance.value = account.initialBalance;
  refs.accountSubmitBtn.textContent = "儲存帳戶";
  refs.accountCancelBtn.classList.remove("hidden");
  clearError(refs.accountError);
}

function enterIncomeEditMode(item) {
  refs.incomeEditId.value = item.id;
  refs.incomeName.value = item.name;
  refs.incomeAmount.value = item.amount;
  refs.incomeAccountId.value = item.accountId;
  refs.incomeCadence.value = item.cadence;
  refs.incomeMonthOfYear.value = item.cadence === CADENCE_YEARLY ? item.monthOfYear : "";
  syncIncomeCadenceField();
  refs.incomeSubmitBtn.textContent = "儲存收入";
  refs.incomeCancelBtn.classList.remove("hidden");
  clearError(refs.incomeError);
}

function enterExpenseEditMode(item) {
  refs.expenseEditId.value = item.id;
  refs.expenseName.value = item.name;
  refs.expenseAmount.value = item.amount;
  refs.expenseAccountId.value = item.accountId;
  refs.expenseCadence.value = item.cadence;
  refs.expenseMonthOfYear.value =
    item.cadence === CADENCE_YEARLY ? item.monthOfYear : "";
  refs.expenseOneTimeMonth.value =
    item.cadence === CADENCE_ONE_TIME
      ? `${item.year}-${padDatePart(item.monthOfYear)}`
      : "";
  syncExpenseCadenceField();
  refs.expenseSubmitBtn.textContent = "儲存支出";
  refs.expenseCancelBtn.classList.remove("hidden");
  clearError(refs.expenseError);
}

function enterInstallmentEditMode(item) {
  refs.installmentEditId.value = item.id;
  refs.installmentName.value = item.name;
  refs.installmentAmount.value = item.amount;
  refs.installmentMonths.value = item.remainingMonths;
  refs.installmentAccountId.value = item.accountId;
  refs.installmentSubmitBtn.textContent = "儲存分期";
  refs.installmentCancelBtn.classList.remove("hidden");
  clearError(refs.installmentError);
}

function enterTransferEditMode(item) {
  refs.transferEditId.value = item.id;
  refs.transferName.value = item.name;
  refs.transferAmount.value = item.amount;
  refs.transferSourceAccountId.value = item.sourceAccountId;
  syncTransferTargetSelection(item.targetAccountId);
  refs.transferSubmitBtn.textContent = "儲存轉帳";
  refs.transferCancelBtn.classList.remove("hidden");
  clearError(refs.transferError);
}

function syncIncomeCadenceField() {
  const showMonth = refs.incomeCadence.value === CADENCE_YEARLY;
  refs.incomeMonthWrap.classList.toggle("hidden", !showMonth);
  refs.incomeMonthOfYear.required = showMonth;
}

function syncExpenseCadenceField() {
  const cadence = refs.expenseCadence.value;
  const showYearlyMonth = cadence === CADENCE_YEARLY;
  const showOneTimeMonth = cadence === CADENCE_ONE_TIME;
  refs.expenseMonthWrap.classList.toggle("hidden", !showYearlyMonth);
  refs.expenseOneTimeWrap.classList.toggle("hidden", !showOneTimeMonth);
  refs.expenseMonthOfYear.required = showYearlyMonth;
  refs.expenseOneTimeMonth.required = showOneTimeMonth;
}

function renderAll() {
  state = sanitizeState(state);
  refreshStaticAccountSelects();
  renderSettings();
  renderAccountTable();
  renderTransferTable();
  renderIncomeTable();
  renderExpenseTable();
  renderInstallmentTable();

  const forecast = calculateForecastData(state);
  renderSummary(forecast);
  renderForecastTable(forecast.totalRows);
  renderAccountForecastTable(forecast.accountRows);
  renderChart(forecast.totalRows);
}

function renderSettings() {
  refs.horizonMonths.value = state.horizonMonths;
}

function renderAccountTable() {
  refs.accountTbody.replaceChildren();

  state.accounts.forEach((account) => {
    const row = document.createElement("tr");
    appendCell(row, account.name);
    appendCell(row, formatCurrency(account.initialBalance));
    appendCell(row, `${getAccountReferenceCount(account.id)} 筆`);

    const actions = document.createElement("td");
    actions.className = "table-actions";
    actions.appendChild(
      buildActionButton("編輯", "btn-secondary", () => {
        enterAccountEditMode(account);
      })
    );
    actions.appendChild(
      buildActionButton("刪除", "btn-danger", () => {
        const reason = getAccountDeletionBlockReason(account.id);
        if (reason) {
          setError(refs.accountError, reason);
          return;
        }
        const confirmed =
          typeof window === "undefined" || typeof window.confirm !== "function"
            ? true
            : window.confirm(`確定刪除帳戶「${account.name}」嗎？`);
        if (!confirmed) {
          return;
        }
        state.accounts = state.accounts.filter((item) => item.id !== account.id);
        clearEditFormIfNeeded("account", account.id);
        persistAndRender();
      })
    );
    row.appendChild(actions);
    refs.accountTbody.appendChild(row);
  });
}

function renderIncomeTable() {
  refs.incomeTbody.replaceChildren();

  if (!state.recurringIncomes.length) {
    appendEmptyRow(refs.incomeTbody, 6, "尚未設定收入項目。");
    renderBatchUi("income");
    return;
  }

  state.recurringIncomes.forEach((item) => {
    const row = document.createElement("tr");
    row.appendChild(buildSelectionCell("income", item.id));
    appendCell(row, item.name);
    appendCell(row, getAccountName(item.accountId));
    appendCell(row, formatCadence(item));
    appendCell(row, formatCurrency(item.amount));
    row.appendChild(buildItemActionCell("income", item.id));
    refs.incomeTbody.appendChild(row);
  });

  renderBatchUi("income");
}

function renderExpenseTable() {
  refs.expenseTbody.replaceChildren();

  if (!state.recurringExpenses.length) {
    appendEmptyRow(refs.expenseTbody, 6, "尚未設定支出項目。");
    renderBatchUi("expense");
    return;
  }

  state.recurringExpenses.forEach((item) => {
    const row = document.createElement("tr");
    row.appendChild(buildSelectionCell("expense", item.id));
    appendCell(row, item.name);
    appendCell(row, getAccountName(item.accountId));
    appendCell(row, formatCadence(item));
    appendCell(row, formatCurrency(item.amount));
    row.appendChild(buildItemActionCell("expense", item.id));
    refs.expenseTbody.appendChild(row);
  });

  renderBatchUi("expense");
}

function renderInstallmentTable() {
  refs.installmentTbody.replaceChildren();

  if (!state.installments.length) {
    appendEmptyRow(refs.installmentTbody, 6, "尚未設定分期項目。");
    renderBatchUi("installment");
    return;
  }

  state.installments.forEach((item) => {
    const row = document.createElement("tr");
    row.appendChild(buildSelectionCell("installment", item.id));
    appendCell(row, item.name);
    appendCell(row, getAccountName(item.accountId));
    appendCell(row, formatCurrency(item.amount));
    appendCell(row, `${item.remainingMonths} 個月`);
    row.appendChild(buildItemActionCell("installment", item.id));
    refs.installmentTbody.appendChild(row);
  });

  renderBatchUi("installment");
}

function renderTransferTable() {
  refs.transferTbody.replaceChildren();

  if (!state.monthlyTransfers.length) {
    appendEmptyRow(refs.transferTbody, 5, "尚未設定每月存款轉帳。");
    return;
  }

  state.monthlyTransfers.forEach((item) => {
    const row = document.createElement("tr");
    appendCell(row, item.name);
    appendCell(row, getAccountName(item.sourceAccountId));
    appendCell(row, getAccountName(item.targetAccountId));
    appendCell(row, formatCurrency(item.amount));

    const actions = document.createElement("td");
    actions.className = "table-actions";
    actions.appendChild(
      buildActionButton("編輯", "btn-secondary", () => {
        enterTransferEditMode(item);
      })
    );
    actions.appendChild(
      buildActionButton("刪除", "btn-danger", () => {
        const confirmed =
          typeof window === "undefined" || typeof window.confirm !== "function"
            ? true
            : window.confirm(`確定刪除轉帳「${item.name}」嗎？`);
        if (!confirmed) {
          return;
        }
        state.monthlyTransfers = state.monthlyTransfers.filter((entry) => entry.id !== item.id);
        clearEditFormIfNeeded("transfer", item.id);
        persistAndRender();
      })
    );
    row.appendChild(actions);
    refs.transferTbody.appendChild(row);
  });
}

function buildSelectionCell(type, itemId) {
  const bucket = batchState[type];
  const cell = document.createElement("td");
  cell.className = "selection-cell";

  const input = document.createElement("input");
  input.type = "checkbox";
  input.className = "row-selector";
  input.checked = bucket.selectedIds.has(itemId);
  input.addEventListener("change", () => {
    toggleBatchItem(type, itemId, input.checked);
  });

  cell.appendChild(input);
  return cell;
}

function buildItemActionCell(type, itemId) {
  const cell = document.createElement("td");
  cell.className = "table-actions";
  const collection = getCollectionByType(type);
  const item = collection.find((entry) => entry.id === itemId);

  cell.appendChild(
    buildActionButton("編輯", "btn-secondary", () => {
      if (item) {
        enterItemEditMode(type, item);
      }
    })
  );
  cell.appendChild(
    buildActionButton("刪除", "btn-danger", () => {
      if (!item) {
        return;
      }
      const confirmed =
        typeof window === "undefined" || typeof window.confirm !== "function"
          ? true
          : window.confirm(`確定刪除「${item.name}」嗎？`);
      if (!confirmed) {
        return;
      }
      removeItemFromCollection(type, itemId);
    })
  );

  return cell;
}

function enterItemEditMode(type, item) {
  if (type === "income") {
    enterIncomeEditMode(item);
    return;
  }
  if (type === "expense") {
    enterExpenseEditMode(item);
    return;
  }
  if (type === "installment") {
    enterInstallmentEditMode(item);
  }
}

function removeItemFromCollection(type, itemId) {
  const collection = getCollectionByType(type);
  const index = collection.findIndex((item) => item.id === itemId);
  if (index < 0) {
    return;
  }

  collection.splice(index, 1);
  clearEditFormIfNeeded(type, itemId);
  clearBatchItemSelection(type, itemId);
  persistAndRender();
}

function clearEditFormIfNeeded(type, itemId) {
  if (type === "account" && refs.accountEditId.value === itemId) {
    resetAccountForm();
    return;
  }
  if (type === "income" && refs.incomeEditId.value === itemId) {
    resetIncomeForm();
    return;
  }
  if (type === "expense" && refs.expenseEditId.value === itemId) {
    resetExpenseForm();
    return;
  }
  if (type === "installment" && refs.installmentEditId.value === itemId) {
    resetInstallmentForm();
    return;
  }
  if (type === "transfer" && refs.transferEditId.value === itemId) {
    resetTransferForm();
  }
}

function renderBatchUi(type) {
  const bucket = batchState[type];
  const items = getCollectionByType(type);
  const batchRefs = getBatchRefs(type);
  pruneBatchBucket(type);

  const selectedCount = bucket.selectedIds.size;
  const hasItems = items.length > 0;
  const isAllSelected = hasItems && selectedCount === items.length;

  batchRefs.selectAll.checked = isAllSelected;
  batchRefs.selectAll.indeterminate = selectedCount > 0 && selectedCount < items.length;
  batchRefs.batchBar.classList.toggle("hidden", selectedCount === 0);
  batchRefs.selectedCount.textContent = `${selectedCount} 筆已選`;

  if (!bucket.isOpen || selectedCount === 0) {
    batchRefs.batchPanel.classList.add("hidden");
    batchRefs.batchList.replaceChildren();
    return;
  }

  batchRefs.batchPanel.classList.remove("hidden");
  batchRefs.batchList.replaceChildren();

  items
    .filter((item) => bucket.selectedIds.has(item.id))
    .forEach((item) => {
      const row = document.createElement("div");
      row.className = "batch-item-row";

      const meta = document.createElement("div");
      meta.className = "batch-item-meta";
      const title = document.createElement("strong");
      title.textContent = item.name;
      const subtitle = document.createElement("span");
      subtitle.textContent = `目前帳戶：${getAccountName(item.accountId)}`;
      meta.appendChild(title);
      meta.appendChild(subtitle);

      const select = document.createElement("select");
      select.className = "batch-item-select";
      state.accounts.forEach((account) => {
        const option = document.createElement("option");
        option.value = account.id;
        option.textContent = account.name;
        select.appendChild(option);
      });

      const selectedAccountId =
        getValidAccountId(bucket.pendingAccounts[item.id]) || item.accountId;
      select.value = selectedAccountId;
      select.addEventListener("change", () => {
        bucket.pendingAccounts[item.id] = getValidAccountId(select.value) || item.accountId;
      });

      row.appendChild(meta);
      row.appendChild(select);
      batchRefs.batchList.appendChild(row);
    });
}

function pruneBatchBucket(type) {
  const bucket = batchState[type];
  const validIds = new Set(getCollectionByType(type).map((item) => item.id));

  Array.from(bucket.selectedIds).forEach((itemId) => {
    if (!validIds.has(itemId)) {
      bucket.selectedIds.delete(itemId);
      delete bucket.pendingAccounts[itemId];
    }
  });

  Object.keys(bucket.pendingAccounts).forEach((itemId) => {
    if (!validIds.has(itemId)) {
      delete bucket.pendingAccounts[itemId];
    }
  });

  if (bucket.selectedIds.size === 0) {
    bucket.isOpen = false;
  }
}

function toggleSelectAll(type, isSelected) {
  const bucket = batchState[type];
  const items = getCollectionByType(type);

  bucket.selectedIds.clear();
  bucket.pendingAccounts = {};

  if (isSelected) {
    items.forEach((item) => {
      bucket.selectedIds.add(item.id);
      bucket.pendingAccounts[item.id] = item.accountId;
    });
  } else {
    bucket.isOpen = false;
  }

  renderAll();
}

function toggleBatchItem(type, itemId, isSelected) {
  const bucket = batchState[type];
  const collection = getCollectionByType(type);
  const item = collection.find((entry) => entry.id === itemId);
  if (!item) {
    return;
  }

  if (isSelected) {
    bucket.selectedIds.add(itemId);
    bucket.pendingAccounts[itemId] = item.accountId;
  } else {
    bucket.selectedIds.delete(itemId);
    delete bucket.pendingAccounts[itemId];
  }

  if (bucket.selectedIds.size === 0) {
    bucket.isOpen = false;
  }

  renderAll();
}

function clearBatchSelection(type) {
  const bucket = batchState[type];
  bucket.selectedIds.clear();
  bucket.pendingAccounts = {};
  bucket.isOpen = false;
  renderAll();
}

function clearBatchItemSelection(type, itemId) {
  const bucket = batchState[type];
  bucket.selectedIds.delete(itemId);
  delete bucket.pendingAccounts[itemId];
  if (bucket.selectedIds.size === 0) {
    bucket.isOpen = false;
  }
}

function openBatchPanel(type) {
  const bucket = batchState[type];
  if (bucket.selectedIds.size === 0) {
    return;
  }

  getCollectionByType(type)
    .filter((item) => bucket.selectedIds.has(item.id))
    .forEach((item) => {
      bucket.pendingAccounts[item.id] = getValidAccountId(bucket.pendingAccounts[item.id]) || item.accountId;
    });

  bucket.isOpen = true;
  renderAll();
}

function closeBatchPanel(type) {
  const bucket = batchState[type];
  bucket.isOpen = false;
  bucket.pendingAccounts = {};
}

function applyBatchAssignments(type) {
  const bucket = batchState[type];
  const collection = getCollectionByType(type);
  let hasChanges = false;

  collection.forEach((item) => {
    if (!bucket.selectedIds.has(item.id)) {
      return;
    }
    const nextAccountId = getValidAccountId(bucket.pendingAccounts[item.id]) || item.accountId;
    if (nextAccountId !== item.accountId) {
      item.accountId = nextAccountId;
      syncEditFormAccountIfNeeded(type, item.id, nextAccountId);
      hasChanges = true;
    }
  });

  bucket.selectedIds.clear();
  bucket.pendingAccounts = {};
  bucket.isOpen = false;

  if (hasChanges) {
    persistAndRender();
  } else {
    renderAll();
  }
}

function syncEditFormAccountIfNeeded(type, itemId, accountId) {
  if (type === "income" && refs.incomeEditId.value === itemId) {
    refs.incomeAccountId.value = accountId;
    return;
  }
  if (type === "expense" && refs.expenseEditId.value === itemId) {
    refs.expenseAccountId.value = accountId;
    return;
  }
  if (type === "installment" && refs.installmentEditId.value === itemId) {
    refs.installmentAccountId.value = accountId;
  }
}

function renderSummary(forecast) {
  const firstRow = forecast.totalRows[0];
  const lastRow = forecast.totalRows[forecast.totalRows.length - 1];

  refs.summaryMonthlyNet.textContent = firstRow ? formatCurrency(firstRow.net) : formatCurrency(0);
  refs.summaryEndingBalance.textContent = lastRow ? formatCurrency(lastRow.endingBalance) : formatCurrency(getTotalInitialBalance());
  refs.summaryNegativeTotalCount.textContent = `${forecast.negativeTotalCount} 個月`;
  refs.summaryNegativeAccountCount.textContent = `${forecast.negativeAccountCount} 個月`;
}

function renderForecastTable(totalRows) {
  refs.forecastTbody.replaceChildren();

  if (!totalRows.length) {
    appendEmptyRow(refs.forecastTbody, 8, "尚無預測資料。");
    return;
  }

  totalRows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.status !== "ok") {
      tr.classList.add("warning-row");
    }

    appendCell(tr, row.monthLabel);
    appendCell(tr, formatCurrency(row.startingBalance));
    appendCell(tr, formatCurrency(row.income));
    appendCell(tr, formatCurrency(row.expense));
    appendCell(tr, formatCurrency(row.installment));
    appendCell(tr, formatCurrency(row.net));
    appendCell(tr, formatCurrency(row.endingBalance));
    tr.appendChild(buildStatusCell(row.status));
    refs.forecastTbody.appendChild(tr);
  });
}

function renderAccountForecastTable(accountRows) {
  refs.accountForecastTbody.replaceChildren();

  if (!accountRows.length) {
    appendEmptyRow(refs.accountForecastTbody, 10, "尚無帳戶明細資料。");
    return;
  }

  accountRows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.status !== "ok") {
      tr.classList.add("warning-row");
    }

    appendCell(tr, row.monthLabel);
    appendCell(tr, row.accountName);
    appendCell(tr, formatCurrency(row.startingBalance));
    appendCell(tr, formatCurrency(row.income));
    appendCell(tr, formatCurrency(row.expense));
    appendCell(tr, formatCurrency(row.installment));
    appendCell(tr, formatCurrency(row.transferIn));
    appendCell(tr, formatCurrency(row.transferOut));
    appendCell(tr, formatCurrency(row.endingBalance));
    tr.appendChild(buildAccountStatusCell(row));
    refs.accountForecastTbody.appendChild(tr);
  });
}

function buildStatusCell(status) {
  const cell = document.createElement("td");
  const chip = document.createElement("span");
  chip.className = "status-chip ok";

  if (status === "total-negative") {
    chip.className = "status-chip warn";
    chip.textContent = "總額不足";
  } else if (status === "account-negative") {
    chip.className = "status-chip caution";
    chip.textContent = "部分帳戶不足";
  } else {
    chip.textContent = "正常";
  }

  cell.appendChild(chip);
  return cell;
}

function buildAccountStatusCell(row) {
  const cell = document.createElement("td");
  const chip = document.createElement("span");
  chip.className = row.status === "negative" ? "status-chip warn" : "status-chip ok";
  chip.textContent = row.status === "negative" ? "帳戶不足" : "正常";
  cell.appendChild(chip);
  return cell;
}

function renderChart(totalRows) {
  clearSectionError(refs.chartError);

  if (!refs.chartCanvas || typeof Chart === "undefined") {
    return;
  }

  const context = typeof refs.chartCanvas.getContext === "function"
    ? refs.chartCanvas.getContext("2d")
    : null;

  if (!context) {
    setSectionError(refs.chartError, "目前無法繪製圖表。");
    return;
  }

  if (forecastChart && typeof forecastChart.destroy === "function") {
    forecastChart.destroy();
  }

  forecastChart = new Chart(context, {
    type: "line",
    data: {
      labels: totalRows.map((row) => row.monthLabel),
      datasets: [
        {
          label: "月末總餘額",
          data: totalRows.map((row) => row.endingBalance),
          borderColor: "#0f766e",
          backgroundColor: "rgba(15, 118, 110, 0.16)",
          fill: true,
          tension: 0.3,
          pointRadius: 3,
          pointHoverRadius: 5,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          ticks: {
            callback(value) {
              return formatCurrency(value);
            },
          },
        },
      },
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `月末總餘額：${formatCurrency(context.parsed.y)}`;
            },
          },
        },
      },
    },
  });
}

function calculateForecastData(currentState) {
  const sourceState = currentState || state;
  const forecastStart = new Date();
  const startMonth = new Date(forecastStart.getFullYear(), forecastStart.getMonth(), 1);
  const balances = {};
  const totalRows = [];
  const accountRows = [];

  sourceState.accounts.forEach((account) => {
    balances[account.id] = account.initialBalance;
  });

  for (let monthIndex = 0; monthIndex < sourceState.horizonMonths; monthIndex += 1) {
    const monthDate = new Date(startMonth.getFullYear(), startMonth.getMonth() + monthIndex, 1);
    const monthLabel = formatMonth(monthDate);
    const monthAccountRows = [];
    let anyNegativeAccount = false;

    sourceState.accounts.forEach((account) => {
      const startingBalance = balances[account.id] || 0;

      const income = sourceState.recurringIncomes.reduce((sum, item) => {
        if (item.accountId !== account.id || !isIncomeActiveInMonth(item, monthDate)) {
          return sum;
        }
        return sum + item.amount;
      }, 0);

      const expense = sourceState.recurringExpenses.reduce((sum, item) => {
        if (item.accountId !== account.id || !isExpenseActiveInMonth(item, monthDate)) {
          return sum;
        }
        return sum + item.amount;
      }, 0);

      const installment = sourceState.installments.reduce((sum, item) => {
        if (item.accountId !== account.id || monthIndex >= item.remainingMonths) {
          return sum;
        }
        return sum + item.amount;
      }, 0);

      const transferIn = sourceState.monthlyTransfers.reduce((sum, item) => {
        return item.targetAccountId === account.id ? sum + item.amount : sum;
      }, 0);

      const transferOut = sourceState.monthlyTransfers.reduce((sum, item) => {
        return item.sourceAccountId === account.id ? sum + item.amount : sum;
      }, 0);

      const endingBalance =
        startingBalance +
        income +
        transferIn -
        expense -
        installment -
        transferOut;

      const row = {
        monthLabel,
        monthDate,
        accountId: account.id,
        accountName: account.name,
        startingBalance,
        income,
        expense,
        installment,
        transferIn,
        transferOut,
        endingBalance,
        status: endingBalance < 0 ? "negative" : "ok",
      };

      if (row.status === "negative") {
        anyNegativeAccount = true;
      }

      balances[account.id] = endingBalance;
      monthAccountRows.push(row);
      accountRows.push(row);
    });

    totalRows.push(buildTotalForecastRow(monthLabel, monthDate, monthAccountRows, anyNegativeAccount));
  }

  return {
    totalRows,
    accountRows,
    negativeTotalCount: totalRows.filter((row) => row.status === "total-negative").length,
    negativeAccountCount: totalRows.filter((row) => row.status !== "ok").length,
  };
}

function buildTotalForecastRow(monthLabel, monthDate, accountRows, anyNegativeAccount) {
  const startingBalance = accountRows.reduce((sum, row) => sum + row.startingBalance, 0);
  const income = accountRows.reduce((sum, row) => sum + row.income, 0);
  const expense = accountRows.reduce((sum, row) => sum + row.expense, 0);
  const installment = accountRows.reduce((sum, row) => sum + row.installment, 0);
  const endingBalance = accountRows.reduce((sum, row) => sum + row.endingBalance, 0);
  const net = income - expense - installment;

  let status = "ok";
  if (endingBalance < 0) {
    status = "total-negative";
  } else if (anyNegativeAccount) {
    status = "account-negative";
  }

  return {
    monthLabel,
    monthDate,
    startingBalance,
    income,
    expense,
    installment,
    net,
    endingBalance,
    status,
  };
}

function isIncomeActiveInMonth(item, monthDate) {
  if (item.cadence === CADENCE_MONTHLY) {
    return true;
  }
  if (item.cadence === CADENCE_YEARLY) {
    return monthDate.getMonth() + 1 === item.monthOfYear;
  }
  return false;
}

function isExpenseActiveInMonth(item, monthDate) {
  if (item.cadence === CADENCE_MONTHLY) {
    return true;
  }
  if (item.cadence === CADENCE_YEARLY) {
    return monthDate.getMonth() + 1 === item.monthOfYear;
  }
  if (item.cadence === CADENCE_ONE_TIME) {
    return monthDate.getFullYear() === item.year && monthDate.getMonth() + 1 === item.monthOfYear;
  }
  return false;
}

function formatMonth(date) {
  return `${date.getFullYear()}/${padDatePart(date.getMonth() + 1)}`;
}

function formatCurrency(amount) {
  const formatter = new Intl.NumberFormat(state.locale || DEFAULT_LOCALE, {
    style: "currency",
    currency: state.currency || DEFAULT_CURRENCY,
    maximumFractionDigits: 0,
  });
  return formatter.format(Number(amount) || 0);
}

function formatCadence(item) {
  if (item.cadence === CADENCE_MONTHLY) {
    return "每月";
  }
  if (item.cadence === CADENCE_YEARLY) {
    return `每年 ${item.monthOfYear} 月`;
  }
  if (item.cadence === CADENCE_ONE_TIME) {
    return `${item.year}/${padDatePart(item.monthOfYear)} 一次`;
  }
  return "未知";
}

function appendCell(row, content) {
  const cell = document.createElement("td");
  cell.textContent = content;
  row.appendChild(cell);
  return cell;
}

function appendEmptyRow(tbody, colspan, message) {
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = colspan;
  cell.textContent = message;
  row.appendChild(cell);
  tbody.appendChild(row);
}

function buildActionButton(label, styleClass, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `btn ${styleClass}`;
  button.textContent = label;
  button.addEventListener("click", onClick);
  return button;
}

function parseIntegerInput(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed) || !Number.isInteger(parsed)) {
    return null;
  }
  return parsed;
}

function parseIncomeCadence(value) {
  if (value === CADENCE_MONTHLY || value === CADENCE_YEARLY) {
    return value;
  }
  return null;
}

function parseExpenseCadence(value) {
  if (
    value === CADENCE_MONTHLY ||
    value === CADENCE_YEARLY ||
    value === CADENCE_ONE_TIME
  ) {
    return value;
  }
  return null;
}

function parseOneTimeMonth(value) {
  if (typeof value !== "string") {
    return null;
  }

  const match = value.match(/^(\d{4})-(\d{2})$/);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const monthOfYear = Number(match[2]);
  if (!Number.isInteger(year) || !Number.isInteger(monthOfYear) || monthOfYear < 1 || monthOfYear > 12) {
    return null;
  }

  return { year, monthOfYear };
}

function normalizeRequiredText(value) {
  return typeof value === "string" ? value.trim() : "";
}

function makeId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function setError(node, message) {
  if (!node) {
    return;
  }
  node.textContent = message || "";
}

function clearError(node) {
  setError(node, "");
}

function setStatusMessage(node, message, status) {
  if (!node) {
    return;
  }
  node.textContent = message || "";
  if (node.classList) {
    node.classList.toggle("is-error", status === "error");
    node.classList.toggle("is-success", status === "success");
  }
}

function clearStatusMessage(node) {
  setStatusMessage(node, "", "");
}

function clearSectionError(node) {
  clearError(node);
}

function setSectionError(node, message) {
  setError(node, message);
}

function downloadJsonBackup() {
  clearStatusMessage(refs.backupStatus);
  clearError(refs.settingsError);

  const payload = {
    appId: APP_ID,
    exportVersion: EXPORT_VERSION,
    exportedAt: new Date().toISOString(),
    data: buildExportPayload(),
  };

  downloadBackupPayload(payload, buildBackupFilename());
  setStatusMessage(refs.backupStatus, "已匯出目前資料。", "success");
}

function triggerImportBackup() {
  clearStatusMessage(refs.backupStatus);
  clearError(refs.settingsError);
  refs.importBackupInput.value = "";
  refs.importBackupInput.click();
}

async function handleImportBackup(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) {
    return;
  }

  clearStatusMessage(refs.backupStatus);
  clearError(refs.settingsError);

  try {
    const importedState = await parseBackupFile(file);
    applyImportedState(importedState);
    setStatusMessage(
      refs.backupStatus,
      `已匯入 ${formatImportLabel(file.name)}。`,
      "success"
    );
  } catch (error) {
    setStatusMessage(refs.backupStatus, getImportErrorMessage(error), "error");
  } finally {
    refs.importBackupInput.value = "";
  }
}

function buildExportPayload() {
  return {
    schemaVersion: SCHEMA_VERSION,
    accounts: state.accounts.map((account) => ({
      id: account.id,
      name: account.name,
      initialBalance: account.initialBalance,
    })),
    horizonMonths: state.horizonMonths,
    recurringIncomes: state.recurringIncomes.map((item) => ({
      id: item.id,
      name: item.name,
      amount: item.amount,
      accountId: item.accountId,
      cadence: item.cadence,
      monthOfYear: item.monthOfYear,
    })),
    recurringExpenses: state.recurringExpenses.map((item) => ({
      id: item.id,
      name: item.name,
      amount: item.amount,
      accountId: item.accountId,
      cadence: item.cadence,
      monthOfYear: item.monthOfYear,
      year: item.year,
    })),
    installments: state.installments.map((item) => ({ ...item })),
    monthlyTransfers: state.monthlyTransfers.map((item) => ({ ...item })),
    currency: state.currency,
    locale: state.locale,
    updatedAt: state.updatedAt,
  };
}

function downloadBackupPayload(payload, filename) {
  if (typeof document === "undefined" || typeof Blob === "undefined") {
    throw new Error("目前環境不支援匯出備份。");
  }

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function parseBackupFile(file) {
  const rawText = await readFileAsText(file);
  let payload;

  try {
    payload = JSON.parse(rawText);
  } catch (error) {
    throw new Error("備份檔不是有效的 JSON。");
  }

  return validateImportPayload(payload);
}

function readFileAsText(file) {
  if (file && typeof file.text === "function") {
    return file.text();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("讀取備份檔失敗。"));
    reader.readAsText(file);
  });
}

function validateImportPayload(payload) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("備份格式不正確。");
  }

  if (Object.prototype.hasOwnProperty.call(payload, "exportVersion")) {
    const exportVersion = Number(payload.exportVersion);
    if (!Object.prototype.hasOwnProperty.call(payload, "data")) {
      throw new Error("備份缺少 data 欄位。");
    }

    if (exportVersion === LEGACY_EXPORT_VERSION) {
      return validateLegacyImportData(payload.data);
    }
    if (exportVersion === EXPORT_VERSION) {
      return validateCurrentImportData(payload.data);
    }
    throw new Error("不支援的備份版本。");
  }

  if (Object.prototype.hasOwnProperty.call(payload, "accounts")) {
    return validateCurrentImportData(payload);
  }

  return validateLegacyImportData(payload);
}

function validateLegacyImportData(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("舊版備份內容格式不正確。");
  }

  const hasKnownKeys = LEGACY_BACKUP_DATA_KEYS.some((key) => Object.prototype.hasOwnProperty.call(data, key));
  if (!hasKnownKeys) {
    throw new Error("找不到可匯入的舊版資料欄位。");
  }

  return sanitizeState(migrateStateToV3(normalizeStateShape(data)));
}

function validateCurrentImportData(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("新版備份內容格式不正確。");
  }

  const hasKnownKeys = CURRENT_BACKUP_DATA_KEYS.some((key) => Object.prototype.hasOwnProperty.call(data, key));
  if (!hasKnownKeys) {
    throw new Error("找不到可匯入的資料欄位。");
  }

  const normalized = normalizeStateShape(data);
  if (Number(normalized.schemaVersion) <= 2 || !Array.isArray(normalized.accounts)) {
    return sanitizeState(migrateStateToV3(normalized));
  }

  if (Number(normalized.schemaVersion) !== SCHEMA_VERSION) {
    throw new Error(`不支援的 schemaVersion：${normalized.schemaVersion}`);
  }

  return sanitizeState(normalized);
}

function applyImportedState(importedState) {
  state = sanitizeState(importedState);
  clearTransientUiState();
  resetAllForms();
  persistAndRender();
}

function getImportErrorMessage(error) {
  return error instanceof Error ? error.message : "匯入備份時發生未知錯誤。";
}

function formatImportLabel(fileName) {
  return fileName ? `備份檔 ${fileName}` : "備份檔";
}

function buildBackupFilename() {
  const now = new Date();
  return `${BACKUP_FILENAME_PREFIX}-${now.getFullYear()}${padDatePart(now.getMonth() + 1)}${padDatePart(now.getDate())}-${padDatePart(now.getHours())}${padDatePart(now.getMinutes())}${padDatePart(now.getSeconds())}.json`;
}

function padDatePart(value) {
  return String(value).padStart(2, "0");
}

function persistAndRender() {
  state = sanitizeState(state);
  state.updatedAt = new Date().toISOString();
  persistState();
  renderAll();
}

function persistState() {
  if (typeof localStorage === "undefined") {
    return;
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(buildExportPayload()));
}

function persistMigratedState(nextState) {
  if (typeof localStorage === "undefined") {
    return;
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(nextState));
}

function loadState() {
  if (typeof localStorage === "undefined") {
    return createDefaultState();
  }

  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return createDefaultState();
  }

  try {
    const parsed = JSON.parse(raw);
    const normalized = normalizeStateShape(parsed);
    const needsMigration =
      Number(normalized.schemaVersion) <= 2 ||
      !Array.isArray(normalized.accounts) ||
      Object.prototype.hasOwnProperty.call(normalized, "initialBalance");

    const nextState = sanitizeState(
      needsMigration ? migrateStateToV3(normalized) : normalized
    );

    persistMigratedState(nextState);
    return nextState;
  } catch (error) {
    return createDefaultState();
  }
}

function normalizeStateShape(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    return {};
  }

  return {
    schemaVersion: input.schemaVersion,
    initialBalance: input.initialBalance,
    accounts: input.accounts,
    horizonMonths: input.horizonMonths,
    recurringIncomes: input.recurringIncomes,
    recurringExpenses: input.recurringExpenses,
    installments: input.installments,
    monthlyTransfers: input.monthlyTransfers,
    currency: input.currency,
    locale: input.locale,
    updatedAt: input.updatedAt,
  };
}

function migrateStateToV3(legacyState) {
  if (Array.isArray(legacyState.accounts) && Number(legacyState.schemaVersion) >= SCHEMA_VERSION) {
    return {
      schemaVersion: SCHEMA_VERSION,
      accounts: legacyState.accounts,
      horizonMonths: legacyState.horizonMonths,
      recurringIncomes: legacyState.recurringIncomes,
      recurringExpenses: legacyState.recurringExpenses,
      installments: legacyState.installments,
      monthlyTransfers: legacyState.monthlyTransfers,
      currency: legacyState.currency,
      locale: legacyState.locale,
      updatedAt: legacyState.updatedAt,
    };
  }

  const mainAccount = createAccount(
    MAIN_ACCOUNT_NAME,
    sanitizeBoundedInteger(legacyState.initialBalance, 0, 0)
  );

  return {
    schemaVersion: SCHEMA_VERSION,
    accounts: [mainAccount],
    horizonMonths: legacyState.horizonMonths,
    recurringIncomes: migrateLegacyRecurringItems(
      legacyState.recurringIncomes,
      mainAccount.id,
      "income"
    ),
    recurringExpenses: migrateLegacyRecurringItems(
      legacyState.recurringExpenses,
      mainAccount.id,
      "expense"
    ),
    installments: migrateLegacyInstallments(legacyState.installments, mainAccount.id),
    monthlyTransfers: [],
    currency: legacyState.currency,
    locale: legacyState.locale,
    updatedAt: legacyState.updatedAt,
  };
}

function migrateLegacyRecurringItems(items, accountId, type) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const cadence =
        type === "income"
          ? parseIncomeCadence(item.cadence) || CADENCE_MONTHLY
          : parseExpenseCadence(item.cadence) || CADENCE_MONTHLY;

      const migrated = {
        id: typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId(),
        name: normalizeRequiredText(item.name),
        amount: sanitizeBoundedInteger(item.amount, 0, 0),
        accountId,
        cadence,
      };

      if (!migrated.name) {
        return null;
      }

      if (cadence === CADENCE_YEARLY) {
        const monthOfYear = sanitizeBoundedInteger(item.monthOfYear, null, 1, 12);
        if (monthOfYear === null) {
          return null;
        }
        migrated.monthOfYear = monthOfYear;
      }

      if (cadence === CADENCE_ONE_TIME) {
        const year = sanitizeBoundedInteger(item.year, null, 2000, 9999);
        const monthOfYear = sanitizeBoundedInteger(item.monthOfYear, null, 1, 12);
        if (year === null || monthOfYear === null) {
          return null;
        }
        migrated.year = year;
        migrated.monthOfYear = monthOfYear;
      }

      return migrated;
    })
    .filter(Boolean);
}

function migrateLegacyInstallments(items, accountId) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const remainingMonths =
        sanitizeBoundedInteger(item.remainingMonths, null, 1, MAX_MONTHS) ??
        sanitizeBoundedInteger(item.monthsRemaining, null, 1, MAX_MONTHS) ??
        sanitizeBoundedInteger(item.months, null, 1, MAX_MONTHS);

      if (!normalizeRequiredText(item.name) || remainingMonths === null) {
        return null;
      }

      return {
        id: typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId(),
        name: normalizeRequiredText(item.name),
        amount: sanitizeBoundedInteger(item.amount, 0, 0),
        remainingMonths,
        accountId,
      };
    })
    .filter(Boolean);
}

function sanitizeState(inputState) {
  const source = inputState || {};
  const accounts = sanitizeAccounts(source.accounts);
  return {
    schemaVersion: SCHEMA_VERSION,
    accounts,
    horizonMonths: sanitizeBoundedInteger(source.horizonMonths, 12, 1, MAX_MONTHS),
    recurringIncomes: sanitizeRecurringIncomes(source.recurringIncomes, accounts),
    recurringExpenses: sanitizeRecurringExpenses(source.recurringExpenses, accounts),
    installments: sanitizeInstallments(source.installments, accounts),
    monthlyTransfers: sanitizeMonthlyTransfers(source.monthlyTransfers, accounts),
    currency: typeof source.currency === "string" && source.currency.trim()
      ? source.currency.trim()
      : DEFAULT_CURRENCY,
    locale: typeof source.locale === "string" && source.locale.trim()
      ? source.locale.trim()
      : DEFAULT_LOCALE,
    updatedAt:
      typeof source.updatedAt === "string" && source.updatedAt.trim()
        ? source.updatedAt
        : new Date().toISOString(),
  };
}

function sanitizeAccounts(accounts) {
  if (!Array.isArray(accounts) || accounts.length === 0) {
    return [createAccount(MAIN_ACCOUNT_NAME, 0)];
  }

  const usedNames = [];
  const usedIds = new Set();

  const sanitized = accounts
    .map((account, index) => {
      if (!account || typeof account !== "object") {
        return null;
      }

      const baseName = normalizeRequiredText(account.name) || (index === 0 ? MAIN_ACCOUNT_NAME : `帳戶 ${index + 1}`);
      const name = makeUniqueAccountName(baseName, usedNames);
      usedNames.push(name);

      let id = typeof account.id === "string" && account.id.trim() ? account.id.trim() : makeId();
      while (usedIds.has(id)) {
        id = makeId();
      }
      usedIds.add(id);

      return {
        id,
        name,
        initialBalance: sanitizeBoundedInteger(account.initialBalance, 0, 0),
      };
    })
    .filter(Boolean);

  return sanitized.length ? sanitized : [createAccount(MAIN_ACCOUNT_NAME, 0)];
}

function makeUniqueAccountName(baseName, usedNames) {
  const normalizedName = normalizeRequiredText(baseName) || MAIN_ACCOUNT_NAME;
  let candidate = normalizedName;
  let suffix = 2;

  while (usedNames.includes(candidate)) {
    candidate = `${normalizedName} (${suffix})`;
    suffix += 1;
  }

  return candidate;
}

function sanitizeRecurringIncomes(items, accounts) {
  if (!Array.isArray(items)) {
    return [];
  }

  const accountIds = new Set(accounts.map((account) => account.id));
  const fallbackAccountId = accounts[0].id;
  const usedIds = new Set();

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const name = normalizeRequiredText(item.name);
      if (!name) {
        return null;
      }

      let id = typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId();
      while (usedIds.has(id)) {
        id = makeId();
      }
      usedIds.add(id);

      const cadence = parseIncomeCadence(item.cadence) || CADENCE_MONTHLY;
      const monthOfYear =
        cadence === CADENCE_YEARLY
          ? sanitizeBoundedInteger(item.monthOfYear, null, 1, 12)
          : undefined;

      if (cadence === CADENCE_YEARLY && monthOfYear === null) {
        return null;
      }

      return {
        id,
        name,
        amount: sanitizeBoundedInteger(item.amount, 0, 0),
        accountId: accountIds.has(item.accountId) ? item.accountId : fallbackAccountId,
        cadence,
        monthOfYear,
      };
    })
    .filter(Boolean);
}

function sanitizeRecurringExpenses(items, accounts) {
  if (!Array.isArray(items)) {
    return [];
  }

  const accountIds = new Set(accounts.map((account) => account.id));
  const fallbackAccountId = accounts[0].id;
  const usedIds = new Set();

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const name = normalizeRequiredText(item.name);
      if (!name) {
        return null;
      }

      let id = typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId();
      while (usedIds.has(id)) {
        id = makeId();
      }
      usedIds.add(id);

      const cadence = parseExpenseCadence(item.cadence) || CADENCE_MONTHLY;
      const base = {
        id,
        name,
        amount: sanitizeBoundedInteger(item.amount, 0, 0),
        accountId: accountIds.has(item.accountId) ? item.accountId : fallbackAccountId,
        cadence,
      };

      if (cadence === CADENCE_YEARLY) {
        const monthOfYear = sanitizeBoundedInteger(item.monthOfYear, null, 1, 12);
        if (monthOfYear === null) {
          return null;
        }
        base.monthOfYear = monthOfYear;
      }

      if (cadence === CADENCE_ONE_TIME) {
        const year = sanitizeBoundedInteger(item.year, null, 2000, 9999);
        const monthOfYear = sanitizeBoundedInteger(item.monthOfYear, null, 1, 12);
        if (year === null || monthOfYear === null) {
          return null;
        }
        base.year = year;
        base.monthOfYear = monthOfYear;
      }

      return base;
    })
    .filter(Boolean);
}

function sanitizeInstallments(items, accounts) {
  if (!Array.isArray(items)) {
    return [];
  }

  const accountIds = new Set(accounts.map((account) => account.id));
  const fallbackAccountId = accounts[0].id;
  const usedIds = new Set();

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const name = normalizeRequiredText(item.name);
      const remainingMonths =
        sanitizeBoundedInteger(item.remainingMonths, null, 1, MAX_MONTHS) ??
        sanitizeBoundedInteger(item.monthsRemaining, null, 1, MAX_MONTHS) ??
        sanitizeBoundedInteger(item.months, null, 1, MAX_MONTHS);

      if (!name || remainingMonths === null) {
        return null;
      }

      let id = typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId();
      while (usedIds.has(id)) {
        id = makeId();
      }
      usedIds.add(id);

      return {
        id,
        name,
        amount: sanitizeBoundedInteger(item.amount, 0, 0),
        remainingMonths,
        accountId: accountIds.has(item.accountId) ? item.accountId : fallbackAccountId,
      };
    })
    .filter(Boolean);
}

function sanitizeMonthlyTransfers(items, accounts) {
  if (!Array.isArray(items)) {
    return [];
  }

  const accountIds = new Set(accounts.map((account) => account.id));
  const usedIds = new Set();

  return items
    .map((item) => {
      if (!item || typeof item !== "object") {
        return null;
      }

      const name = normalizeRequiredText(item.name);
      if (!name) {
        return null;
      }

      const amount = sanitizeBoundedInteger(item.amount, null, 1);
      if (amount === null) {
        return null;
      }

      if (!accountIds.has(item.sourceAccountId) || !accountIds.has(item.targetAccountId)) {
        return null;
      }
      if (item.sourceAccountId === item.targetAccountId) {
        return null;
      }

      let id = typeof item.id === "string" && item.id.trim() ? item.id.trim() : makeId();
      while (usedIds.has(id)) {
        id = makeId();
      }
      usedIds.add(id);

      return {
        id,
        name,
        amount,
        sourceAccountId: item.sourceAccountId,
        targetAccountId: item.targetAccountId,
      };
    })
    .filter(Boolean);
}

function sanitizeBoundedInteger(value, fallback, min, max) {
  const parsed = parseIntegerInput(value);
  if (parsed === null) {
    return fallback;
  }
  if (typeof min === "number" && parsed < min) {
    return fallback;
  }
  if (typeof max === "number" && parsed > max) {
    return fallback;
  }
  return parsed;
}

function getCollectionByType(type) {
  if (type === "income") {
    return state.recurringIncomes;
  }
  if (type === "expense") {
    return state.recurringExpenses;
  }
  if (type === "installment") {
    return state.installments;
  }
  return [];
}

function getBatchRefs(type) {
  if (type === "income") {
    return {
      selectAll: refs.incomeSelectAll,
      batchBar: refs.incomeBatchBar,
      selectedCount: refs.incomeSelectedCount,
      batchBtn: refs.incomeBatchBtn,
      clearBtn: refs.incomeClearBtn,
      batchPanel: refs.incomeBatchPanel,
      batchList: refs.incomeBatchList,
      applyBtn: refs.incomeBatchApplyBtn,
      cancelBtn: refs.incomeBatchCancelBtn,
    };
  }
  if (type === "expense") {
    return {
      selectAll: refs.expenseSelectAll,
      batchBar: refs.expenseBatchBar,
      selectedCount: refs.expenseSelectedCount,
      batchBtn: refs.expenseBatchBtn,
      clearBtn: refs.expenseClearBtn,
      batchPanel: refs.expenseBatchPanel,
      batchList: refs.expenseBatchList,
      applyBtn: refs.expenseBatchApplyBtn,
      cancelBtn: refs.expenseBatchCancelBtn,
    };
  }
  return {
    selectAll: refs.installmentSelectAll,
    batchBar: refs.installmentBatchBar,
    selectedCount: refs.installmentSelectedCount,
    batchBtn: refs.installmentBatchBtn,
    clearBtn: refs.installmentClearBtn,
    batchPanel: refs.installmentBatchPanel,
    batchList: refs.installmentBatchList,
    applyBtn: refs.installmentBatchApplyBtn,
    cancelBtn: refs.installmentBatchCancelBtn,
  };
}

function getAccountById(accountId) {
  return state.accounts.find((account) => account.id === accountId) || null;
}

function getAccountName(accountId) {
  const account = getAccountById(accountId);
  return account ? account.name : "未知帳戶";
}

function getDefaultAccountId() {
  return state.accounts[0] ? state.accounts[0].id : "";
}

function getDefaultTransferTargetId(sourceAccountId) {
  const alternative = state.accounts.find((account) => account.id !== sourceAccountId);
  return alternative ? alternative.id : "";
}

function getValidAccountId(accountId) {
  return state.accounts.some((account) => account.id === accountId) ? accountId : "";
}

function getTotalInitialBalance() {
  return state.accounts.reduce((sum, account) => sum + account.initialBalance, 0);
}

function getAccountReferenceCount(accountId) {
  const incomeCount = state.recurringIncomes.filter((item) => item.accountId === accountId).length;
  const expenseCount = state.recurringExpenses.filter((item) => item.accountId === accountId).length;
  const installmentCount = state.installments.filter((item) => item.accountId === accountId).length;
  const transferCount = state.monthlyTransfers.filter(
    (item) => item.sourceAccountId === accountId || item.targetAccountId === accountId
  ).length;
  return incomeCount + expenseCount + installmentCount + transferCount;
}

function getAccountDeletionBlockReason(accountId) {
  const account = getAccountById(accountId);
  if (!account) {
    return "找不到要刪除的帳戶。";
  }
  if (state.accounts.length <= 1) {
    return "至少需要保留一個帳戶。";
  }
  if (account.initialBalance !== 0) {
    return "帳戶期初餘額不為 0，無法刪除。";
  }
  if (getAccountReferenceCount(accountId) > 0) {
    return "帳戶仍被收入、支出、分期或轉帳引用，請先改掛其他帳戶。";
  }
  if (isAccountReferencedInPendingBatches(accountId)) {
    return "帳戶仍在批次歸戶面板中被指定，請先取消或套用。";
  }
  return "";
}

function isAccountReferencedInPendingBatches(accountId) {
  return BATCH_TYPES.some((type) => {
    const bucket = batchState[type];
    if (!bucket.isOpen) {
      return false;
    }
    return Object.values(bucket.pendingAccounts).some((value) => value === accountId);
  });
}

function clearTransientUiState() {
  BATCH_TYPES.forEach((type) => {
    batchState[type].selectedIds.clear();
    batchState[type].isOpen = false;
    batchState[type].pendingAccounts = {};
  });
}

function isDuplicateAccountName(name, excludeId) {
  return state.accounts.some((account) => account.id !== excludeId && account.name === name);
}
