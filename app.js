const STORAGE_KEY = "accounting_forecast_v1";
const MAX_MONTHS = 120;
const DEFAULT_CURRENCY = "TWD";
const DEFAULT_LOCALE = "zh-TW";

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
  refs.settingsError = document.getElementById("settings-error");

  refs.summaryMonthlyNet = document.getElementById("summary-monthly-net");
  refs.summaryEndingBalance = document.getElementById("summary-ending-balance");
  refs.summaryNegativeCount = document.getElementById("summary-negative-count");

  refs.incomeForm = document.getElementById("income-form");
  refs.incomeEditId = document.getElementById("income-edit-id");
  refs.incomeName = document.getElementById("income-name");
  refs.incomeCategory = document.getElementById("income-category");
  refs.incomeAmount = document.getElementById("income-amount");
  refs.incomeSubmitBtn = document.getElementById("income-submit-btn");
  refs.incomeCancelBtn = document.getElementById("income-cancel-btn");
  refs.incomeError = document.getElementById("income-error");
  refs.incomeTbody = document.getElementById("income-tbody");

  refs.expenseForm = document.getElementById("expense-form");
  refs.expenseEditId = document.getElementById("expense-edit-id");
  refs.expenseName = document.getElementById("expense-name");
  refs.expenseCategory = document.getElementById("expense-category");
  refs.expenseAmount = document.getElementById("expense-amount");
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

  refs.incomeForm.addEventListener("submit", onIncomeSubmit);
  refs.incomeCancelBtn.addEventListener("click", resetIncomeForm);

  refs.expenseForm.addEventListener("submit", onExpenseSubmit);
  refs.expenseCancelBtn.addEventListener("click", resetExpenseForm);

  refs.installmentForm.addEventListener("submit", onInstallmentSubmit);
  refs.installmentCancelBtn.addEventListener("click", resetInstallmentForm);
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
    setError(refs.incomeError, "每月金額需為 0 以上整數。");
    return;
  }

  const category = normalizeCategory(refs.incomeCategory.value);
  const editId = refs.incomeEditId.value.trim();
  const payload = { id: editId || makeId(), name, category, amount };

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
    setError(refs.expenseError, "每月金額需為 0 以上整數。");
    return;
  }

  const category = normalizeCategory(refs.expenseCategory.value);
  const editId = refs.expenseEditId.value.trim();
  const payload = { id: editId || makeId(), name, category, amount };

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
  refs.incomeSubmitBtn.textContent = "新增收入";
  refs.incomeCancelBtn.classList.add("hidden");
  clearError(refs.incomeError);
}

function resetExpenseForm() {
  refs.expenseForm.reset();
  refs.expenseEditId.value = "";
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
  refs.incomeSubmitBtn.textContent = "更新收入";
  refs.incomeCancelBtn.classList.remove("hidden");
  clearError(refs.incomeError);
}

function enterExpenseEditMode(item) {
  refs.expenseEditId.value = item.id;
  refs.expenseName.value = item.name;
  refs.expenseCategory.value = item.category;
  refs.expenseAmount.value = item.amount.toString();
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
    appendEmptyRow(refs.incomeTbody, 4, "尚未新增定期收入");
    return;
  }

  state.recurringIncomes.forEach((item) => {
    const tr = document.createElement("tr");
    appendCell(tr, item.name);
    appendCell(tr, item.category);
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
    appendEmptyRow(refs.expenseTbody, 4, "尚未新增定期支出");
    return;
  }

  state.recurringExpenses.forEach((item) => {
    const tr = document.createElement("tr");
    appendCell(tr, item.name);
    appendCell(tr, item.category);
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
  const monthlyIncome = sumAmounts(state.recurringIncomes);
  const monthlyExpense = sumAmounts(state.recurringExpenses);
  let prevEndingBalance = state.initialBalance;

  for (let i = 0; i < state.horizonMonths; i += 1) {
    const date = new Date(today.getFullYear(), today.getMonth() + i, 1);
    const installmentTotal = state.installments.reduce((sum, installment) => {
      return sum + (i < installment.remainingMonths ? installment.amount : 0);
    }, 0);

    const net = monthlyIncome - monthlyExpense - installmentTotal;
    const startingBalance = i === 0 ? state.initialBalance : prevEndingBalance;
    const endingBalance = startingBalance + net;

    rows.push({
      monthLabel: formatMonth(date),
      startingBalance,
      incomeTotal: monthlyIncome,
      expenseTotal: monthlyExpense,
      installmentTotal,
      net,
      endingBalance,
    });

    prevEndingBalance = endingBalance;
  }

  return rows;
}

function sumAmounts(items) {
  return items.reduce((sum, item) => sum + item.amount, 0);
}

function formatMonth(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}年${month}月`;
}

function formatCurrency(value) {
  return currencyFormatter.format(value);
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

function persistAndRender() {
  persistState();
  renderAll();
}

function persistState() {
  state.updatedAt = new Date().toISOString();
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (_error) {
    setError(refs.settingsError, "儲存失敗：瀏覽器無法寫入本地資料。");
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return createDefaultState();
    }

    const parsed = JSON.parse(raw);
    return sanitizeState(parsed);
  } catch (_error) {
    return createDefaultState();
  }
}

function sanitizeState(input) {
  const safeState = createDefaultState();
  if (!input || typeof input !== "object") {
    return safeState;
  }

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

      return {
        id: typeof rawItem.id === "string" && rawItem.id.trim() ? rawItem.id : makeId(),
        name,
        category: normalizeCategory(rawItem.category),
        amount,
      };
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
