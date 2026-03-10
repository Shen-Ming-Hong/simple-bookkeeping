const STORAGE_KEY = 'accounting_forecast_v1';
const SCHEMA_VERSION = 3;
const EXPORT_VERSION = 2;
const LEGACY_EXPORT_VERSION = 1;
const MAX_MONTHS = 120;
const DEFAULT_CURRENCY = 'TWD';
const DEFAULT_LOCALE = 'zh-TW';
const CADENCE_MONTHLY = 'monthly';
const CADENCE_YEARLY = 'yearly';
const CADENCE_ONE_TIME = 'one-time';
const APP_ID = 'accounting_forecast';
const BACKUP_FILENAME_PREFIX = 'accounting-backup';
const MAIN_ACCOUNT_NAME = '主帳戶';
const BATCH_TYPES = ['income', 'expense', 'installment'];
const SYNC_META_STORAGE_KEY = 'accounting_forecast_sync_meta_v1';
const DRIVE_SESSION_STORAGE_KEY = 'accounting_forecast_drive_session_v1';
const THEME_STORAGE_KEY = 'accounting_forecast_theme_v1';
const THEME_LIGHT = 'light';
const THEME_DARK = 'dark';
const THEME_COLOR_META_NAME = 'theme-color';
const LIGHT_THEME_COLOR = '#f2f6f8';
const DARK_THEME_COLOR = '#09131f';
const GOOGLE_CLIENT_ID_META_NAME = 'google-client-id';
const GOOGLE_DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.appdata';
const GOOGLE_DRIVE_FILE_NAME = 'accounting_forecast.json';
const GOOGLE_DRIVE_API_URL = 'https://www.googleapis.com/drive/v3/files';
const GOOGLE_DRIVE_UPLOAD_URL = 'https://www.googleapis.com/upload/drive/v3/files';
const SYNC_UPLOAD_DEBOUNCE_MS = 1200;
const ACCESS_TOKEN_EXPIRY_BUFFER_MS = 60 * 1000;
const SERVICE_WORKER_SCRIPT_URL = './service-worker.js';
const LEGACY_BACKUP_DATA_KEYS = [
	'schemaVersion',
	'initialBalance',
	'horizonMonths',
	'recurringIncomes',
	'recurringExpenses',
	'installments',
	'currency',
	'locale',
	'updatedAt',
];
const CURRENT_BACKUP_DATA_KEYS = [
	'schemaVersion',
	'accounts',
	'horizonMonths',
	'recurringIncomes',
	'recurringExpenses',
	'installments',
	'monthlyTransfers',
	'currency',
	'locale',
	'updatedAt',
];

let state = createDefaultState();
let forecastChart = null;
let syncUploadTimer = null;

const refs = {};
const batchState = createBatchState();
const syncState = createSyncState();
const uiState = createUiState();

if (typeof document !== 'undefined') {
	document.addEventListener('DOMContentLoaded', () => {
		cacheDom();
		applyTheme(loadThemePreference(), { persist: false, rerenderChart: false });
		state = loadState();
		bindEvents();
		refreshStaticAccountSelects();
		resetAllForms();
		renderAll();
		registerServiceWorker();
		void initializeCloudSync();
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

function createDefaultSyncMeta() {
	return {
		lastSyncedAt: '',
		lastKnownCloudUpdatedAt: '',
		lastKnownCloudFileId: '',
		syncStatus: 'local-only',
		pendingUpload: false,
		lastAuthAttemptAt: '',
	};
}

function createDefaultDriveSession() {
	return {
		accessToken: '',
		accessTokenExpiresAt: 0,
	};
}

function createSyncState() {
	return {
		meta: createDefaultSyncMeta(),
		clientId: '',
		isConfigured: false,
		isGoogleReady: false,
		hasGoogleScriptListeners: false,
		tokenClient: null,
		accessToken: '',
		accessTokenExpiresAt: 0,
		isConnected: false,
		isAuthorizing: false,
		isSyncing: false,
		isNetworkOnline: isBrowserOnline(),
		hasSessionChanges: false,
		hadLocalDataAtStartup: false,
		statusCode: 'local-only',
		statusMessage: '目前使用本機資料。',
		statusTone: 'info',
	};
}

function createUiState() {
	return {
		theme: getDocumentTheme(),
		isSyncPopoverOpen: false,
		isTotalForecastExpanded: false,
		expandedAccountForecastAccounts: new Set(),
		hasTouchedAccountForecastExpansion: false,
	};
}

function sanitizeTheme(value) {
	return value === THEME_DARK ? THEME_DARK : THEME_LIGHT;
}

function isBrowserOnline() {
	if (typeof navigator === 'undefined' || typeof navigator.onLine !== 'boolean') {
		return true;
	}

	return navigator.onLine;
}

function getDocumentTheme() {
	if (typeof document === 'undefined') {
		return THEME_LIGHT;
	}

	return sanitizeTheme(document.documentElement.dataset.theme);
}

function loadThemePreference() {
	if (typeof localStorage === 'undefined') {
		return getDocumentTheme();
	}

	try {
		return sanitizeTheme(localStorage.getItem(THEME_STORAGE_KEY));
	} catch (error) {
		return getDocumentTheme();
	}
}

function persistThemePreference(theme) {
	if (typeof localStorage === 'undefined') {
		return;
	}

	try {
		localStorage.setItem(THEME_STORAGE_KEY, sanitizeTheme(theme));
	} catch (error) {
		// Ignore storage write failures and keep the active in-memory theme.
	}
}

function buildThemeToggleLabel(theme = uiState.theme) {
	return sanitizeTheme(theme) === THEME_DARK ? '切換為亮色主題' : '切換為暗色主題';
}

function getThemeColor(theme = uiState.theme) {
	return sanitizeTheme(theme) === THEME_DARK ? DARK_THEME_COLOR : LIGHT_THEME_COLOR;
}

function updateThemeColorMeta(theme = uiState.theme) {
	if (typeof document === 'undefined') {
		return;
	}

	const themeColorMeta = document.querySelector(`meta[name="${THEME_COLOR_META_NAME}"]`);
	if (themeColorMeta) {
		themeColorMeta.setAttribute('content', getThemeColor(theme));
	}
}

function renderThemeToggle() {
	if (!refs.themeToggleBtn) {
		return;
	}

	const label = buildThemeToggleLabel();
	refs.themeToggleBtn.dataset.theme = uiState.theme;
	refs.themeToggleBtn.setAttribute('aria-label', label);
	refs.themeToggleBtn.setAttribute('title', label);
	refs.themeToggleBtn.setAttribute('aria-pressed', String(uiState.theme === THEME_DARK));
}

function applyTheme(nextTheme, options = {}) {
	const theme = sanitizeTheme(nextTheme);
	const persist = options.persist !== false;
	const rerenderChart = options.rerenderChart !== false;

	uiState.theme = theme;
	if (typeof document !== 'undefined') {
		document.documentElement.dataset.theme = theme;
	}
	updateThemeColorMeta(theme);
	if (persist) {
		persistThemePreference(theme);
	}

	renderThemeToggle();
	if (rerenderChart) {
		renderChart(calculateForecastData(state).totalRows);
	}
}

function toggleTheme() {
	const nextTheme = uiState.theme === THEME_DARK ? THEME_LIGHT : THEME_DARK;
	applyTheme(nextTheme);
}

function setSyncPopoverOpen(isOpen, options = {}) {
	const nextOpen = Boolean(isOpen);
	if (uiState.isSyncPopoverOpen === nextOpen) {
		return;
	}

	uiState.isSyncPopoverOpen = nextOpen;
	renderSyncPopoverState();

	if (nextOpen) {
		if (options.focus !== false && typeof window !== 'undefined') {
			window.requestAnimationFrame(() => {
				focusSyncPopoverAction();
			});
		}
		return;
	}

	if (options.restoreFocus !== false && refs.syncToggleBtn && typeof refs.syncToggleBtn.focus === 'function') {
		refs.syncToggleBtn.focus();
	}
}

function toggleSyncPopover() {
	setSyncPopoverOpen(!uiState.isSyncPopoverOpen);
}

function renderSyncPopoverState() {
	if (!refs.syncToggleBtn || !refs.syncPopover) {
		return;
	}

	refs.syncToggleBtn.setAttribute('aria-expanded', String(uiState.isSyncPopoverOpen));
	refs.syncPopover.classList.toggle('hidden', !uiState.isSyncPopoverOpen);
	refs.syncToggleBtn.setAttribute('aria-label', buildSyncToggleLabel());
	refs.syncToggleBtn.setAttribute('title', buildSyncToggleLabel());
}

function focusSyncPopoverAction() {
	const focusTarget = [refs.syncConnectBtn, refs.syncUploadBtn, refs.syncDownloadBtn, refs.syncDisconnectBtn].find(
		button => button && !button.disabled
	);

	if (focusTarget && typeof focusTarget.focus === 'function') {
		focusTarget.focus();
		return;
	}

	if (refs.syncPopover && typeof refs.syncPopover.focus === 'function') {
		refs.syncPopover.focus();
	}
}

function handleDocumentPointerDown(event) {
	if (!uiState.isSyncPopoverOpen || !refs.syncUtility) {
		return;
	}

	if (!(event.target instanceof Node)) {
		return;
	}

	if (refs.syncUtility.contains(event.target)) {
		return;
	}

	setSyncPopoverOpen(false, { restoreFocus: false });
}

function handleGlobalKeydown(event) {
	if (event.key !== 'Escape' || !uiState.isSyncPopoverOpen) {
		return;
	}

	event.preventDefault();
	setSyncPopoverOpen(false);
}

function buildSyncToggleLabel() {
	const action = uiState.isSyncPopoverOpen ? '關閉' : '開啟';
	return `${action}雲端同步面板。目前狀態：${syncState.statusMessage}`;
}

function getSyncIndicatorTone() {
	const errorStatuses = new Set(['error', 'auth-required', 'sync-error', 'gis-error']);
	const successStatuses = new Set(['synced', 'connected', 'cloud-applied']);
	const progressStatuses = new Set(['syncing', 'pending-upload', 'auth-check', 'gis-loading', 'local-newer']);

	if (syncState.statusTone === 'error' || errorStatuses.has(syncState.statusCode)) {
		return 'error';
	}
	if (syncState.statusTone === 'success' || successStatuses.has(syncState.statusCode)) {
		return 'success';
	}
	if (progressStatuses.has(syncState.statusCode)) {
		return 'progress';
	}
	return 'neutral';
}

function shouldPulseSyncIndicator() {
	const pulsingStatuses = new Set(['syncing', 'pending-upload', 'auth-check', 'gis-loading']);
	return pulsingStatuses.has(syncState.statusCode);
}

function renderSyncIndicator() {
	if (!refs.syncToggleBtn || !refs.syncStatusDot) {
		return;
	}

	const tone = getSyncIndicatorTone();
	refs.syncToggleBtn.dataset.syncTone = tone;
	refs.syncStatusDot.dataset.syncTone = tone;
	refs.syncStatusDot.classList.toggle('is-pulsing', shouldPulseSyncIndicator());
	refs.syncToggleBtn.setAttribute('aria-label', buildSyncToggleLabel());
	refs.syncToggleBtn.setAttribute('title', buildSyncToggleLabel());
}

function getCssVariableValue(name, fallback) {
	if (typeof window === 'undefined' || typeof document === 'undefined') {
		return fallback;
	}

	const value = window.getComputedStyle(document.documentElement).getPropertyValue(name).trim();
	return value || fallback;
}

function getChartPalette() {
	return {
		line: getCssVariableValue('--chart-line', '#0f766e'),
		fill: getCssVariableValue('--chart-fill', 'rgba(15, 118, 110, 0.16)'),
		grid: getCssVariableValue('--chart-grid', 'rgba(15, 36, 48, 0.1)'),
		text: getCssVariableValue('--chart-text', '#5d7280'),
		tooltipBackground: getCssVariableValue('--chart-tooltip-bg', 'rgba(255, 255, 255, 0.96)'),
		tooltipBorder: getCssVariableValue('--chart-tooltip-border', 'rgba(15, 36, 48, 0.12)'),
		tooltipText: getCssVariableValue('--chart-tooltip-text', '#0f2430'),
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
	refs.googleIdentityScript = document.getElementById('google-identity-script');
	refs.themeToggleBtn = document.getElementById('theme-toggle-btn');
	refs.syncUtility = document.getElementById('sync-utility');
	refs.syncToggleBtn = document.getElementById('sync-toggle-btn');
	refs.syncPopover = document.getElementById('sync-popover');
	refs.syncStatusDot = document.getElementById('sync-status-dot');
	refs.settingsForm = document.getElementById('settings-form');
	refs.horizonMonths = document.getElementById('horizon-months');
	refs.exportBackupBtn = document.getElementById('export-backup-btn');
	refs.importBackupBtn = document.getElementById('import-backup-btn');
	refs.importBackupInput = document.getElementById('import-backup-input');
	refs.backupStatus = document.getElementById('backup-status');
	refs.settingsError = document.getElementById('settings-error');
	refs.syncPanelTitle = document.getElementById('sync-panel-title');
	refs.syncConnectionStatus = document.getElementById('sync-connection-status');
	refs.syncStatusMessage = document.getElementById('sync-status-message');
	refs.syncLocalUpdatedAt = document.getElementById('sync-local-updated-at');
	refs.syncCloudUpdatedAt = document.getElementById('sync-cloud-updated-at');
	refs.syncLastSyncedAt = document.getElementById('sync-last-synced-at');
	refs.syncConnectBtn = document.getElementById('sync-connect-btn');
	refs.syncUploadBtn = document.getElementById('sync-upload-btn');
	refs.syncDownloadBtn = document.getElementById('sync-download-btn');
	refs.syncDisconnectBtn = document.getElementById('sync-disconnect-btn');

	refs.summaryMonthlyNet = document.getElementById('summary-monthly-net');
	refs.summaryEndingBalance = document.getElementById('summary-ending-balance');
	refs.summaryNegativeTotalCount = document.getElementById('summary-negative-total-count');
	refs.summaryNegativeAccountCount = document.getElementById('summary-negative-account-count');

	refs.accountForm = document.getElementById('account-form');
	refs.accountEditId = document.getElementById('account-edit-id');
	refs.accountName = document.getElementById('account-name');
	refs.accountBalance = document.getElementById('account-balance');
	refs.accountSubmitBtn = document.getElementById('account-submit-btn');
	refs.accountCancelBtn = document.getElementById('account-cancel-btn');
	refs.accountError = document.getElementById('account-error');
	refs.accountTbody = document.getElementById('account-tbody');

	refs.incomeForm = document.getElementById('income-form');
	refs.incomeEditId = document.getElementById('income-edit-id');
	refs.incomeName = document.getElementById('income-name');
	refs.incomeAmount = document.getElementById('income-amount');
	refs.incomeAccountId = document.getElementById('income-account-id');
	refs.incomeCadence = document.getElementById('income-cadence');
	refs.incomeMonthWrap = document.getElementById('income-month-wrap');
	refs.incomeMonthOfYear = document.getElementById('income-month-of-year');
	refs.incomeSubmitBtn = document.getElementById('income-submit-btn');
	refs.incomeCancelBtn = document.getElementById('income-cancel-btn');
	refs.incomeError = document.getElementById('income-error');
	refs.incomeSelectAll = document.getElementById('income-select-all');
	refs.incomeBatchBar = document.getElementById('income-batch-bar');
	refs.incomeSelectedCount = document.getElementById('income-selected-count');
	refs.incomeBatchBtn = document.getElementById('income-batch-btn');
	refs.incomeClearBtn = document.getElementById('income-clear-btn');
	refs.incomeBatchPanel = document.getElementById('income-batch-panel');
	refs.incomeBatchList = document.getElementById('income-batch-list');
	refs.incomeBatchApplyBtn = document.getElementById('income-batch-apply-btn');
	refs.incomeBatchCancelBtn = document.getElementById('income-batch-cancel-btn');
	refs.incomeTbody = document.getElementById('income-tbody');

	refs.expenseForm = document.getElementById('expense-form');
	refs.expenseEditId = document.getElementById('expense-edit-id');
	refs.expenseName = document.getElementById('expense-name');
	refs.expenseAmount = document.getElementById('expense-amount');
	refs.expenseAccountId = document.getElementById('expense-account-id');
	refs.expenseCadence = document.getElementById('expense-cadence');
	refs.expenseMonthWrap = document.getElementById('expense-month-wrap');
	refs.expenseMonthOfYear = document.getElementById('expense-month-of-year');
	refs.expenseOneTimeWrap = document.getElementById('expense-one-time-wrap');
	refs.expenseOneTimeMonth = document.getElementById('expense-one-time-month');
	refs.expenseSubmitBtn = document.getElementById('expense-submit-btn');
	refs.expenseCancelBtn = document.getElementById('expense-cancel-btn');
	refs.expenseError = document.getElementById('expense-error');
	refs.expenseSelectAll = document.getElementById('expense-select-all');
	refs.expenseBatchBar = document.getElementById('expense-batch-bar');
	refs.expenseSelectedCount = document.getElementById('expense-selected-count');
	refs.expenseBatchBtn = document.getElementById('expense-batch-btn');
	refs.expenseClearBtn = document.getElementById('expense-clear-btn');
	refs.expenseBatchPanel = document.getElementById('expense-batch-panel');
	refs.expenseBatchList = document.getElementById('expense-batch-list');
	refs.expenseBatchApplyBtn = document.getElementById('expense-batch-apply-btn');
	refs.expenseBatchCancelBtn = document.getElementById('expense-batch-cancel-btn');
	refs.expenseTbody = document.getElementById('expense-tbody');

	refs.installmentForm = document.getElementById('installment-form');
	refs.installmentEditId = document.getElementById('installment-edit-id');
	refs.installmentName = document.getElementById('installment-name');
	refs.installmentAmount = document.getElementById('installment-amount');
	refs.installmentMonths = document.getElementById('installment-months');
	refs.installmentAccountId = document.getElementById('installment-account-id');
	refs.installmentSubmitBtn = document.getElementById('installment-submit-btn');
	refs.installmentCancelBtn = document.getElementById('installment-cancel-btn');
	refs.installmentError = document.getElementById('installment-error');
	refs.installmentSelectAll = document.getElementById('installment-select-all');
	refs.installmentBatchBar = document.getElementById('installment-batch-bar');
	refs.installmentSelectedCount = document.getElementById('installment-selected-count');
	refs.installmentBatchBtn = document.getElementById('installment-batch-btn');
	refs.installmentClearBtn = document.getElementById('installment-clear-btn');
	refs.installmentBatchPanel = document.getElementById('installment-batch-panel');
	refs.installmentBatchList = document.getElementById('installment-batch-list');
	refs.installmentBatchApplyBtn = document.getElementById('installment-batch-apply-btn');
	refs.installmentBatchCancelBtn = document.getElementById('installment-batch-cancel-btn');
	refs.installmentTbody = document.getElementById('installment-tbody');

	refs.transferForm = document.getElementById('transfer-form');
	refs.transferEditId = document.getElementById('transfer-edit-id');
	refs.transferName = document.getElementById('transfer-name');
	refs.transferAmount = document.getElementById('transfer-amount');
	refs.transferSourceAccountId = document.getElementById('transfer-source-account-id');
	refs.transferTargetAccountId = document.getElementById('transfer-target-account-id');
	refs.transferSubmitBtn = document.getElementById('transfer-submit-btn');
	refs.transferCancelBtn = document.getElementById('transfer-cancel-btn');
	refs.transferError = document.getElementById('transfer-error');
	refs.transferTbody = document.getElementById('transfer-tbody');

	refs.forecastTbody = document.getElementById('forecast-tbody');
	refs.accountForecastTbody = document.getElementById('account-forecast-tbody');
	refs.forecastMobile = document.getElementById('forecast-mobile');
	refs.accountForecastMobile = document.getElementById('account-forecast-mobile');
	refs.chartCanvas = document.getElementById('forecast-chart');
	refs.chartError = document.getElementById('chart-error');
	refs.actionGroups = Array.from(document.querySelectorAll('.form-actions, .sync-actions, .batch-actions'));
}

function bindEvents() {
	refs.themeToggleBtn.addEventListener('click', toggleTheme);
	refs.syncToggleBtn.addEventListener('click', toggleSyncPopover);
	refs.settingsForm.addEventListener('submit', onSettingsSubmit);
	refs.exportBackupBtn.addEventListener('click', downloadJsonBackup);
	refs.importBackupBtn.addEventListener('click', triggerImportBackup);
	refs.importBackupInput.addEventListener('change', handleImportBackup);
	refs.syncConnectBtn.addEventListener('click', () => {
		void connectGoogleDrive();
	});
	refs.syncUploadBtn.addEventListener('click', () => {
		void uploadToGoogleDrive();
	});
	refs.syncDownloadBtn.addEventListener('click', () => {
		void syncFromGoogleDrive();
	});
	refs.syncDisconnectBtn.addEventListener('click', () => {
		void disconnectGoogleDrive();
	});
	document.addEventListener('pointerdown', handleDocumentPointerDown);
	document.addEventListener('keydown', handleGlobalKeydown);
	if (typeof window !== 'undefined') {
		window.addEventListener('online', handleBrowserOnline);
		window.addEventListener('offline', handleBrowserOffline);
	}

	refs.accountForm.addEventListener('submit', onAccountSubmit);
	refs.accountCancelBtn.addEventListener('click', resetAccountForm);

	refs.incomeForm.addEventListener('submit', onIncomeSubmit);
	refs.incomeCancelBtn.addEventListener('click', resetIncomeForm);
	refs.incomeCadence.addEventListener('change', syncIncomeCadenceField);

	refs.expenseForm.addEventListener('submit', onExpenseSubmit);
	refs.expenseCancelBtn.addEventListener('click', resetExpenseForm);
	refs.expenseCadence.addEventListener('change', syncExpenseCadenceField);

	refs.installmentForm.addEventListener('submit', onInstallmentSubmit);
	refs.installmentCancelBtn.addEventListener('click', resetInstallmentForm);

	refs.transferForm.addEventListener('submit', onTransferSubmit);
	refs.transferCancelBtn.addEventListener('click', resetTransferForm);
	refs.transferSourceAccountId.addEventListener('change', () => {
		syncTransferTargetSelection();
	});

	bindBatchEvents('income');
	bindBatchEvents('expense');
	bindBatchEvents('installment');
	setupActionGroupLayoutTracking();
}

function setupActionGroupLayoutTracking() {
	if (!Array.isArray(refs.actionGroups) || !refs.actionGroups.length) {
		return;
	}

	refs.actionGroups.forEach(group => {
		syncActionGroupLayout(group);
		if (typeof MutationObserver === 'undefined' || group.dataset.actionLayoutObserved === 'true') {
			return;
		}

		const observer = new MutationObserver(() => {
			syncActionGroupLayout(group);
		});
		observer.observe(group, {
			attributes: true,
			subtree: true,
			attributeFilter: ['class'],
		});
		group.dataset.actionLayoutObserved = 'true';
	});
}

function syncActionGroupLayout(group) {
	if (!(group instanceof Element)) {
		return;
	}

	const visibleCount = Array.from(group.querySelectorAll('.btn')).filter(button => !button.classList.contains('hidden'))
		.length;
	group.dataset.visibleCount = String(visibleCount);
}

function registerServiceWorker() {
	if (typeof window === 'undefined' || typeof navigator === 'undefined' || !('serviceWorker' in navigator)) {
		return;
	}

	window.addEventListener(
		'load',
		() => {
			void navigator.serviceWorker.register(SERVICE_WORKER_SCRIPT_URL).catch(() => {});
		},
		{ once: true }
	);
}

function buildOfflineSyncMessage() {
	return syncState.meta.pendingUpload
		? '目前離線，本機資料會在恢復連線後再同步到 Google Drive。'
		: '目前離線，仍可使用本機資料。';
}

function handleBrowserOnline() {
	syncState.isNetworkOnline = true;

	if (!syncState.isConfigured) {
		renderSyncPanel();
		return;
	}

	if (!syncState.isGoogleReady) {
		void initializeCloudSync();
		return;
	}

	if (syncState.meta.pendingUpload) {
		setSyncFeedback('pending-upload', '已恢復連線，本機資料將自動同步到 Google Drive。', 'info', {
			persistStatus: false,
		});
		maybeContinuePendingCloudUpload();
		return;
	}

	if (syncState.isConnected || hasValidDriveAccessToken()) {
		setSyncFeedback('connected', '已恢復連線，可繼續與 Google Drive 同步。', 'success', {
			persistStatus: false,
		});
		return;
	}

	setSyncFeedback('local-only', '已恢復連線，可隨時連接 Google Drive。', 'info', {
		persistStatus: false,
	});
}

function handleBrowserOffline() {
	syncState.isNetworkOnline = false;

	if (!syncState.isConfigured) {
		renderSyncPanel();
		return;
	}

	setSyncFeedback('offline', buildOfflineSyncMessage(), 'info', {
		persistStatus: false,
	});
}

function bindBatchEvents(type) {
	const batchRefs = getBatchRefs(type);
	batchRefs.selectAll.addEventListener('change', () => {
		toggleSelectAll(type, batchRefs.selectAll.checked);
	});
	batchRefs.batchBtn.addEventListener('click', () => openBatchPanel(type));
	batchRefs.clearBtn.addEventListener('click', () => clearBatchSelection(type));
	batchRefs.applyBtn.addEventListener('click', () => applyBatchAssignments(type));
	batchRefs.cancelBtn.addEventListener('click', () => {
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
	setAccountSelectOptions(refs.incomeAccountId, getValidAccountId(refs.incomeAccountId.value) || defaultAccountId);
	setAccountSelectOptions(refs.expenseAccountId, getValidAccountId(refs.expenseAccountId.value) || defaultAccountId);
	setAccountSelectOptions(refs.installmentAccountId, getValidAccountId(refs.installmentAccountId.value) || defaultAccountId);
	setAccountSelectOptions(refs.transferSourceAccountId, getValidAccountId(refs.transferSourceAccountId.value) || defaultAccountId);
	syncTransferTargetSelection(refs.transferTargetAccountId.value);
}

function setAccountSelectOptions(selectElement, preferredValue) {
	if (!selectElement) {
		return;
	}

	const fragment = document.createDocumentFragment();
	state.accounts.forEach(account => {
		const option = document.createElement('option');
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

	const candidates = state.accounts.filter(account => account.id !== sourceAccountId);
	refs.transferTargetAccountId.replaceChildren();

	if (!candidates.length) {
		const option = document.createElement('option');
		option.value = '';
		option.textContent = '請先新增第二個帳戶';
		refs.transferTargetAccountId.appendChild(option);
		refs.transferTargetAccountId.value = '';
		refs.transferTargetAccountId.disabled = true;
		return;
	}

	refs.transferTargetAccountId.disabled = false;
	candidates.forEach(account => {
		const option = document.createElement('option');
		option.value = account.id;
		option.textContent = account.name;
		refs.transferTargetAccountId.appendChild(option);
	});

	const nextTargetId = candidates.some(account => account.id === preferredTargetId) ? preferredTargetId : candidates[0].id;
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
	setStatusMessage(refs.backupStatus, '已更新預測月數。', 'success');
}

function onAccountSubmit(event) {
	event.preventDefault();
	clearError(refs.accountError);

	const name = normalizeRequiredText(refs.accountName.value);
	if (!name) {
		setError(refs.accountError, '請輸入帳戶名稱。');
		return;
	}

	const initialBalance = parseIntegerInput(refs.accountBalance.value);
	if (initialBalance === null || initialBalance < 0) {
		setError(refs.accountError, '期初餘額需為 0 以上的整數。');
		return;
	}

	const editId = refs.accountEditId.value.trim();
	if (isDuplicateAccountName(name, editId)) {
		setError(refs.accountError, '帳戶名稱不可重複。');
		return;
	}

	if (editId) {
		const account = getAccountById(editId);
		if (!account) {
			setError(refs.accountError, '找不到要更新的帳戶，請重新操作。');
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
		setError(refs.incomeError, '請輸入收入名稱。');
		return;
	}

	const amount = parseIntegerInput(refs.incomeAmount.value);
	if (amount === null || amount < 0) {
		setError(refs.incomeError, '收入金額需為 0 以上的整數。');
		return;
	}

	const accountId = getValidAccountId(refs.incomeAccountId.value);
	if (!accountId) {
		setError(refs.incomeError, '請選擇帳戶。');
		return;
	}

	const cadence = parseIncomeCadence(refs.incomeCadence.value);
	if (!cadence) {
		setError(refs.incomeError, '請選擇正確的收入週期。');
		return;
	}

	const monthOfYear = cadence === CADENCE_YEARLY ? parseIntegerInput(refs.incomeMonthOfYear.value) : null;
	if (cadence === CADENCE_YEARLY && (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12)) {
		setError(refs.incomeError, '年度收入需要指定 1 到 12 月。');
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
			setError(refs.incomeError, '找不到要更新的收入項目，請重新操作。');
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
		setError(refs.expenseError, '請輸入支出名稱。');
		return;
	}

	const amount = parseIntegerInput(refs.expenseAmount.value);
	if (amount === null || amount < 0) {
		setError(refs.expenseError, '支出金額需為 0 以上的整數。');
		return;
	}

	const accountId = getValidAccountId(refs.expenseAccountId.value);
	if (!accountId) {
		setError(refs.expenseError, '請選擇帳戶。');
		return;
	}

	const cadence = parseExpenseCadence(refs.expenseCadence.value);
	if (!cadence) {
		setError(refs.expenseError, '請選擇正確的支出週期。');
		return;
	}

	let monthOfYear;
	let year;

	if (cadence === CADENCE_YEARLY) {
		monthOfYear = parseIntegerInput(refs.expenseMonthOfYear.value);
		if (monthOfYear === null || monthOfYear < 1 || monthOfYear > 12) {
			setError(refs.expenseError, '年度支出需要指定 1 到 12 月。');
			return;
		}
	}

	if (cadence === CADENCE_ONE_TIME) {
		const parsed = parseOneTimeMonth(refs.expenseOneTimeMonth.value);
		if (!parsed) {
			setError(refs.expenseError, '一次性支出需要指定有效的年月。');
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
			setError(refs.expenseError, '找不到要更新的支出項目，請重新操作。');
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
		setError(refs.installmentError, '請輸入分期名稱。');
		return;
	}

	const amount = parseIntegerInput(refs.installmentAmount.value);
	if (amount === null || amount < 0) {
		setError(refs.installmentError, '分期金額需為 0 以上的整數。');
		return;
	}

	const remainingMonths = parseIntegerInput(refs.installmentMonths.value);
	if (remainingMonths === null || remainingMonths < 1 || remainingMonths > MAX_MONTHS) {
		setError(refs.installmentError, `剩餘月數需為 1 到 ${MAX_MONTHS} 的整數。`);
		return;
	}

	const accountId = getValidAccountId(refs.installmentAccountId.value);
	if (!accountId) {
		setError(refs.installmentError, '請選擇帳戶。');
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
			setError(refs.installmentError, '找不到要更新的分期項目，請重新操作。');
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
		setError(refs.transferError, '請輸入轉帳名稱。');
		return;
	}

	const amount = parseIntegerInput(refs.transferAmount.value);
	if (amount === null || amount <= 0) {
		setError(refs.transferError, '每月轉帳金額需為大於 0 的整數。');
		return;
	}

	const sourceAccountId = getValidAccountId(refs.transferSourceAccountId.value);
	const targetAccountId = getValidAccountId(refs.transferTargetAccountId.value);
	if (!sourceAccountId || !targetAccountId) {
		setError(refs.transferError, '轉帳需要同時指定來源與目標帳戶。');
		return;
	}
	if (sourceAccountId === targetAccountId) {
		setError(refs.transferError, '來源與目標帳戶不可相同。');
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
			setError(refs.transferError, '找不到要更新的轉帳項目，請重新操作。');
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
	const index = list.findIndex(item => item.id === nextItem.id);
	if (index < 0) {
		return false;
	}
	list[index] = nextItem;
	return true;
}

function resetAccountForm() {
	refs.accountForm.reset();
	refs.accountEditId.value = '';
	refs.accountSubmitBtn.textContent = '新增帳戶';
	refs.accountCancelBtn.classList.add('hidden');
	clearError(refs.accountError);
}

function resetIncomeForm() {
	refs.incomeForm.reset();
	refs.incomeEditId.value = '';
	refs.incomeCadence.value = CADENCE_MONTHLY;
	refs.incomeMonthOfYear.value = '';
	refs.incomeAccountId.value = getDefaultAccountId();
	syncIncomeCadenceField();
	refs.incomeSubmitBtn.textContent = '新增收入';
	refs.incomeCancelBtn.classList.add('hidden');
	clearError(refs.incomeError);
}

function resetExpenseForm() {
	refs.expenseForm.reset();
	refs.expenseEditId.value = '';
	refs.expenseCadence.value = CADENCE_MONTHLY;
	refs.expenseMonthOfYear.value = '';
	refs.expenseOneTimeMonth.value = '';
	refs.expenseAccountId.value = getDefaultAccountId();
	syncExpenseCadenceField();
	refs.expenseSubmitBtn.textContent = '新增支出';
	refs.expenseCancelBtn.classList.add('hidden');
	clearError(refs.expenseError);
}

function resetInstallmentForm() {
	refs.installmentForm.reset();
	refs.installmentEditId.value = '';
	refs.installmentAccountId.value = getDefaultAccountId();
	refs.installmentSubmitBtn.textContent = '新增分期';
	refs.installmentCancelBtn.classList.add('hidden');
	clearError(refs.installmentError);
}

function resetTransferForm() {
	refs.transferForm.reset();
	refs.transferEditId.value = '';
	refs.transferSourceAccountId.value = getDefaultAccountId();
	syncTransferTargetSelection(getDefaultTransferTargetId(refs.transferSourceAccountId.value));
	refs.transferSubmitBtn.textContent = '新增轉帳';
	refs.transferCancelBtn.classList.add('hidden');
	clearError(refs.transferError);
}

function enterAccountEditMode(account) {
	refs.accountEditId.value = account.id;
	refs.accountName.value = account.name;
	refs.accountBalance.value = account.initialBalance;
	refs.accountSubmitBtn.textContent = '儲存帳戶';
	refs.accountCancelBtn.classList.remove('hidden');
	clearError(refs.accountError);
}

function enterIncomeEditMode(item) {
	refs.incomeEditId.value = item.id;
	refs.incomeName.value = item.name;
	refs.incomeAmount.value = item.amount;
	refs.incomeAccountId.value = item.accountId;
	refs.incomeCadence.value = item.cadence;
	refs.incomeMonthOfYear.value = item.cadence === CADENCE_YEARLY ? item.monthOfYear : '';
	syncIncomeCadenceField();
	refs.incomeSubmitBtn.textContent = '儲存收入';
	refs.incomeCancelBtn.classList.remove('hidden');
	clearError(refs.incomeError);
}

function enterExpenseEditMode(item) {
	refs.expenseEditId.value = item.id;
	refs.expenseName.value = item.name;
	refs.expenseAmount.value = item.amount;
	refs.expenseAccountId.value = item.accountId;
	refs.expenseCadence.value = item.cadence;
	refs.expenseMonthOfYear.value = item.cadence === CADENCE_YEARLY ? item.monthOfYear : '';
	refs.expenseOneTimeMonth.value = item.cadence === CADENCE_ONE_TIME ? `${item.year}-${padDatePart(item.monthOfYear)}` : '';
	syncExpenseCadenceField();
	refs.expenseSubmitBtn.textContent = '儲存支出';
	refs.expenseCancelBtn.classList.remove('hidden');
	clearError(refs.expenseError);
}

function enterInstallmentEditMode(item) {
	refs.installmentEditId.value = item.id;
	refs.installmentName.value = item.name;
	refs.installmentAmount.value = item.amount;
	refs.installmentMonths.value = item.remainingMonths;
	refs.installmentAccountId.value = item.accountId;
	refs.installmentSubmitBtn.textContent = '儲存分期';
	refs.installmentCancelBtn.classList.remove('hidden');
	clearError(refs.installmentError);
}

function enterTransferEditMode(item) {
	refs.transferEditId.value = item.id;
	refs.transferName.value = item.name;
	refs.transferAmount.value = item.amount;
	refs.transferSourceAccountId.value = item.sourceAccountId;
	syncTransferTargetSelection(item.targetAccountId);
	refs.transferSubmitBtn.textContent = '儲存轉帳';
	refs.transferCancelBtn.classList.remove('hidden');
	clearError(refs.transferError);
}

function syncIncomeCadenceField() {
	const showMonth = refs.incomeCadence.value === CADENCE_YEARLY;
	refs.incomeMonthWrap.classList.toggle('hidden', !showMonth);
	refs.incomeMonthOfYear.required = showMonth;
}

function syncExpenseCadenceField() {
	const cadence = refs.expenseCadence.value;
	const showYearlyMonth = cadence === CADENCE_YEARLY;
	const showOneTimeMonth = cadence === CADENCE_ONE_TIME;
	refs.expenseMonthWrap.classList.toggle('hidden', !showYearlyMonth);
	refs.expenseOneTimeWrap.classList.toggle('hidden', !showOneTimeMonth);
	refs.expenseMonthOfYear.required = showYearlyMonth;
	refs.expenseOneTimeMonth.required = showOneTimeMonth;
}

function renderAll() {
	state = sanitizeState(state);
	refreshStaticAccountSelects();
	renderSettings();
	renderSyncPanel();
	renderAccountTable();
	renderTransferTable();
	renderIncomeTable();
	renderExpenseTable();
	renderInstallmentTable();

	const forecast = calculateForecastData(state);
	renderSummary(forecast);
	renderForecastTable(forecast.totalRows);
	renderForecastMobile(forecast.totalRows);
	renderAccountForecastTable(forecast.accountRows);
	renderAccountForecastMobile(forecast.accountRows);
	renderChart(forecast.totalRows);
}

function renderSettings() {
	refs.horizonMonths.value = state.horizonMonths;
	renderThemeToggle();
	renderSyncPopoverState();
}

function renderSyncPanel() {
	if (!refs.syncConnectionStatus) {
		return;
	}

	refs.syncPanelTitle.textContent = syncState.isConfigured ? 'Google Drive 同步' : 'Google Drive 同步（未設定）';
	refs.syncConnectionStatus.textContent = getSyncConnectionLabel();
	refs.syncLocalUpdatedAt.textContent = formatSyncTimestamp(state.updatedAt);
	refs.syncCloudUpdatedAt.textContent = formatSyncTimestamp(syncState.meta.lastKnownCloudUpdatedAt);
	refs.syncLastSyncedAt.textContent = formatSyncTimestamp(syncState.meta.lastSyncedAt);
	setStatusMessage(refs.syncStatusMessage, syncState.statusMessage, syncState.statusTone);

	const allowManualSync =
		syncState.isNetworkOnline && syncState.isConfigured && syncState.isGoogleReady && !syncState.isAuthorizing && !syncState.isSyncing;

	refs.syncConnectBtn.disabled =
		!syncState.isNetworkOnline || !syncState.isConfigured || !syncState.isGoogleReady || syncState.isAuthorizing || syncState.isSyncing;
	refs.syncUploadBtn.disabled = !allowManualSync;
	refs.syncDownloadBtn.disabled = !allowManualSync;
	refs.syncDisconnectBtn.disabled =
		!syncState.isNetworkOnline ||
		!syncState.isConfigured ||
		syncState.isAuthorizing ||
		syncState.isSyncing ||
		(!syncState.isConnected && !syncState.accessToken);

	refs.syncConnectBtn.textContent = syncState.isConnected ? '重新授權' : '連接 Google Drive';
	renderSyncIndicator();
	renderSyncPopoverState();
}

function renderAccountTable() {
	refs.accountTbody.replaceChildren();

	state.accounts.forEach(account => {
		const row = document.createElement('tr');
		appendCell(row, account.name);
		appendCell(row, formatCurrency(account.initialBalance));
		appendCell(row, `${getAccountReferenceCount(account.id)} 筆`);

		const actions = document.createElement('td');
		actions.className = 'table-actions';
		actions.appendChild(
			buildActionButton('編輯', 'btn-secondary', () => {
				enterAccountEditMode(account);
			})
		);
		actions.appendChild(
			buildActionButton('刪除', 'btn-danger', () => {
				const reason = getAccountDeletionBlockReason(account.id);
				if (reason) {
					setError(refs.accountError, reason);
					return;
				}
				const confirmed =
					typeof window === 'undefined' || typeof window.confirm !== 'function'
						? true
						: window.confirm(`確定刪除帳戶「${account.name}」嗎？`);
				if (!confirmed) {
					return;
				}
				state.accounts = state.accounts.filter(item => item.id !== account.id);
				clearEditFormIfNeeded('account', account.id);
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
		appendEmptyRow(refs.incomeTbody, 6, '尚未設定收入項目。');
		renderBatchUi('income');
		return;
	}

	state.recurringIncomes.forEach(item => {
		const row = document.createElement('tr');
		row.appendChild(buildSelectionCell('income', item.id));
		appendCell(row, item.name);
		appendCell(row, getAccountName(item.accountId));
		appendCell(row, formatCadence(item));
		appendCell(row, formatCurrency(item.amount));
		row.appendChild(buildItemActionCell('income', item.id));
		refs.incomeTbody.appendChild(row);
	});

	renderBatchUi('income');
}

function renderExpenseTable() {
	refs.expenseTbody.replaceChildren();

	if (!state.recurringExpenses.length) {
		appendEmptyRow(refs.expenseTbody, 6, '尚未設定支出項目。');
		renderBatchUi('expense');
		return;
	}

	state.recurringExpenses.forEach(item => {
		const row = document.createElement('tr');
		row.appendChild(buildSelectionCell('expense', item.id));
		appendCell(row, item.name);
		appendCell(row, getAccountName(item.accountId));
		appendCell(row, formatCadence(item));
		appendCell(row, formatCurrency(item.amount));
		row.appendChild(buildItemActionCell('expense', item.id));
		refs.expenseTbody.appendChild(row);
	});

	renderBatchUi('expense');
}

function renderInstallmentTable() {
	refs.installmentTbody.replaceChildren();

	if (!state.installments.length) {
		appendEmptyRow(refs.installmentTbody, 6, '尚未設定分期項目。');
		renderBatchUi('installment');
		return;
	}

	state.installments.forEach(item => {
		const row = document.createElement('tr');
		row.appendChild(buildSelectionCell('installment', item.id));
		appendCell(row, item.name);
		appendCell(row, getAccountName(item.accountId));
		appendCell(row, formatCurrency(item.amount));
		appendCell(row, `${item.remainingMonths} 個月`);
		row.appendChild(buildItemActionCell('installment', item.id));
		refs.installmentTbody.appendChild(row);
	});

	renderBatchUi('installment');
}

function renderTransferTable() {
	refs.transferTbody.replaceChildren();

	if (!state.monthlyTransfers.length) {
		appendEmptyRow(refs.transferTbody, 5, '尚未設定每月存款轉帳。');
		return;
	}

	state.monthlyTransfers.forEach(item => {
		const row = document.createElement('tr');
		appendCell(row, item.name);
		appendCell(row, getAccountName(item.sourceAccountId));
		appendCell(row, getAccountName(item.targetAccountId));
		appendCell(row, formatCurrency(item.amount));

		const actions = document.createElement('td');
		actions.className = 'table-actions';
		actions.appendChild(
			buildActionButton('編輯', 'btn-secondary', () => {
				enterTransferEditMode(item);
			})
		);
		actions.appendChild(
			buildActionButton('刪除', 'btn-danger', () => {
				const confirmed =
					typeof window === 'undefined' || typeof window.confirm !== 'function'
						? true
						: window.confirm(`確定刪除轉帳「${item.name}」嗎？`);
				if (!confirmed) {
					return;
				}
				state.monthlyTransfers = state.monthlyTransfers.filter(entry => entry.id !== item.id);
				clearEditFormIfNeeded('transfer', item.id);
				persistAndRender();
			})
		);
		row.appendChild(actions);
		refs.transferTbody.appendChild(row);
	});
}

function buildSelectionCell(type, itemId) {
	const bucket = batchState[type];
	const cell = document.createElement('td');
	cell.className = 'selection-cell';

	const input = document.createElement('input');
	input.type = 'checkbox';
	input.className = 'row-selector';
	input.checked = bucket.selectedIds.has(itemId);
	input.addEventListener('change', () => {
		toggleBatchItem(type, itemId, input.checked);
	});

	cell.appendChild(input);
	return cell;
}

function buildItemActionCell(type, itemId) {
	const cell = document.createElement('td');
	cell.className = 'table-actions';
	const collection = getCollectionByType(type);
	const item = collection.find(entry => entry.id === itemId);

	cell.appendChild(
		buildActionButton('編輯', 'btn-secondary', () => {
			if (item) {
				enterItemEditMode(type, item);
			}
		})
	);
	cell.appendChild(
		buildActionButton('刪除', 'btn-danger', () => {
			if (!item) {
				return;
			}
			const confirmed =
				typeof window === 'undefined' || typeof window.confirm !== 'function'
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
	if (type === 'income') {
		enterIncomeEditMode(item);
		return;
	}
	if (type === 'expense') {
		enterExpenseEditMode(item);
		return;
	}
	if (type === 'installment') {
		enterInstallmentEditMode(item);
	}
}

function removeItemFromCollection(type, itemId) {
	const collection = getCollectionByType(type);
	const index = collection.findIndex(item => item.id === itemId);
	if (index < 0) {
		return;
	}

	collection.splice(index, 1);
	clearEditFormIfNeeded(type, itemId);
	clearBatchItemSelection(type, itemId);
	persistAndRender();
}

function clearEditFormIfNeeded(type, itemId) {
	if (type === 'account' && refs.accountEditId.value === itemId) {
		resetAccountForm();
		return;
	}
	if (type === 'income' && refs.incomeEditId.value === itemId) {
		resetIncomeForm();
		return;
	}
	if (type === 'expense' && refs.expenseEditId.value === itemId) {
		resetExpenseForm();
		return;
	}
	if (type === 'installment' && refs.installmentEditId.value === itemId) {
		resetInstallmentForm();
		return;
	}
	if (type === 'transfer' && refs.transferEditId.value === itemId) {
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
	batchRefs.batchBar.classList.toggle('hidden', selectedCount === 0);
	batchRefs.selectedCount.textContent = `${selectedCount} 筆已選`;

	if (!bucket.isOpen || selectedCount === 0) {
		batchRefs.batchPanel.classList.add('hidden');
		batchRefs.batchList.replaceChildren();
		return;
	}

	batchRefs.batchPanel.classList.remove('hidden');
	batchRefs.batchList.replaceChildren();

	items
		.filter(item => bucket.selectedIds.has(item.id))
		.forEach(item => {
			const row = document.createElement('div');
			row.className = 'batch-item-row';

			const meta = document.createElement('div');
			meta.className = 'batch-item-meta';
			const title = document.createElement('strong');
			title.textContent = item.name;
			const subtitle = document.createElement('span');
			subtitle.textContent = `目前帳戶：${getAccountName(item.accountId)}`;
			meta.appendChild(title);
			meta.appendChild(subtitle);

			const select = document.createElement('select');
			select.className = 'batch-item-select';
			state.accounts.forEach(account => {
				const option = document.createElement('option');
				option.value = account.id;
				option.textContent = account.name;
				select.appendChild(option);
			});

			const selectedAccountId = getValidAccountId(bucket.pendingAccounts[item.id]) || item.accountId;
			select.value = selectedAccountId;
			select.addEventListener('change', () => {
				bucket.pendingAccounts[item.id] = getValidAccountId(select.value) || item.accountId;
			});

			row.appendChild(meta);
			row.appendChild(select);
			batchRefs.batchList.appendChild(row);
		});
}

function pruneBatchBucket(type) {
	const bucket = batchState[type];
	const validIds = new Set(getCollectionByType(type).map(item => item.id));

	Array.from(bucket.selectedIds).forEach(itemId => {
		if (!validIds.has(itemId)) {
			bucket.selectedIds.delete(itemId);
			delete bucket.pendingAccounts[itemId];
		}
	});

	Object.keys(bucket.pendingAccounts).forEach(itemId => {
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
		items.forEach(item => {
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
	const item = collection.find(entry => entry.id === itemId);
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
		.filter(item => bucket.selectedIds.has(item.id))
		.forEach(item => {
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

	collection.forEach(item => {
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
	if (type === 'income' && refs.incomeEditId.value === itemId) {
		refs.incomeAccountId.value = accountId;
		return;
	}
	if (type === 'expense' && refs.expenseEditId.value === itemId) {
		refs.expenseAccountId.value = accountId;
		return;
	}
	if (type === 'installment' && refs.installmentEditId.value === itemId) {
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
		appendEmptyRow(refs.forecastTbody, 8, '尚無預測資料。');
		return;
	}

	totalRows.forEach(row => {
		const tr = document.createElement('tr');
		tr.className = 'forecast-row';
		if (row.status !== 'ok') {
			tr.classList.add('warning-row');
		}

		appendCell(tr, row.monthLabel, { label: '月份', className: 'forecast-cell forecast-cell-month' });
		appendCell(tr, formatCurrency(row.startingBalance), { label: '起始餘額', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.income), { label: '收入', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.expense), { label: '支出', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.installment), { label: '分期', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.net), { label: '淨額', className: 'forecast-cell forecast-cell-highlight' });
		appendCell(tr, formatCurrency(row.endingBalance), { label: '月末總餘額', className: 'forecast-cell forecast-cell-highlight' });
		tr.appendChild(buildStatusCell(row.status, { label: '狀態', className: 'forecast-cell forecast-cell-status' }));
		refs.forecastTbody.appendChild(tr);
	});
}

function renderForecastMobile(totalRows) {
	if (!refs.forecastMobile) {
		return;
	}

	refs.forecastMobile.replaceChildren();

	if (!totalRows.length) {
		const emptyState = document.createElement('p');
		emptyState.className = 'mobile-forecast-empty';
		emptyState.textContent = '尚無預測資料。';
		refs.forecastMobile.appendChild(emptyState);
		return;
	}

	const lastRow = totalRows[totalRows.length - 1];
	const details = document.createElement('details');
	details.className = 'forecast-mobile-accordion-item';
	details.open = uiState.isTotalForecastExpanded;

	const summary = document.createElement('summary');
	summary.className = 'forecast-mobile-summary';

	const titleWrap = document.createElement('div');
	titleWrap.className = 'forecast-mobile-summary-main';

	const title = document.createElement('strong');
	title.textContent = '總表';
	const subtitle = document.createElement('span');
	subtitle.textContent = `${totalRows.length} 個月 · 末餘額 ${formatCurrency(lastRow.endingBalance)}`;
	titleWrap.appendChild(title);
	titleWrap.appendChild(subtitle);

	const indicator = document.createElement('div');
	indicator.className = 'forecast-mobile-summary-side';
	const totalStatus = getTotalForecastStatusMeta(totalRows);
	indicator.appendChild(createStatusChip(totalStatus.tone, totalStatus.text));

	summary.appendChild(titleWrap);
	summary.appendChild(indicator);
	details.appendChild(summary);

	const cards = document.createElement('div');
	cards.className = 'forecast-mobile-cards';

	totalRows.forEach(row => {
		cards.appendChild(buildMobileForecastCard(row));
	});

	details.appendChild(cards);
	details.addEventListener('toggle', () => {
		uiState.isTotalForecastExpanded = details.open;
	});

	refs.forecastMobile.appendChild(details);
}

function renderAccountForecastTable(accountRows) {
	refs.accountForecastTbody.replaceChildren();

	if (!accountRows.length) {
		appendEmptyRow(refs.accountForecastTbody, 10, '尚無帳戶明細資料。');
		return;
	}

	accountRows.forEach(row => {
		const tr = document.createElement('tr');
		tr.className = 'account-forecast-row';
		if (row.status !== 'ok') {
			tr.classList.add('warning-row');
		}

		appendCell(tr, row.monthLabel, { label: '月份', className: 'forecast-cell forecast-cell-month' });
		appendCell(tr, row.accountName, { label: '帳戶', className: 'forecast-cell forecast-cell-highlight' });
		appendCell(tr, formatCurrency(row.startingBalance), { label: '起始餘額', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.income), { label: '收入', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.expense), { label: '支出', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.installment), { label: '分期', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.transferIn), { label: '轉入', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.transferOut), { label: '轉出', className: 'forecast-cell' });
		appendCell(tr, formatCurrency(row.endingBalance), { label: '月末餘額', className: 'forecast-cell forecast-cell-highlight' });
		tr.appendChild(buildAccountStatusCell(row, { label: '狀態', className: 'forecast-cell forecast-cell-status' }));
		refs.accountForecastTbody.appendChild(tr);
	});
}

function renderAccountForecastMobile(accountRows) {
	if (!refs.accountForecastMobile) {
		return;
	}

	refs.accountForecastMobile.replaceChildren();

	if (!accountRows.length) {
		const emptyState = document.createElement('p');
		emptyState.className = 'mobile-forecast-empty';
		emptyState.textContent = '尚無帳戶明細資料。';
		refs.accountForecastMobile.appendChild(emptyState);
		return;
	}

	const groupedRows = groupAccountForecastRowsByAccount(accountRows);
	const expandedAccounts = uiState.expandedAccountForecastAccounts;
	const groupIds = new Set(groupedRows.map(group => group.accountId));
	const firstNegativeGroup = groupedRows.find(group => group.hasNegative);

	Array.from(expandedAccounts).forEach(accountId => {
		if (!groupIds.has(accountId)) {
			expandedAccounts.delete(accountId);
		}
	});

	groupedRows.forEach(group => {
		const details = document.createElement('details');
		details.className = 'forecast-mobile-accordion-item';
		details.open = uiState.hasTouchedAccountForecastExpansion ? expandedAccounts.has(group.accountId) : firstNegativeGroup?.accountId === group.accountId;

		const summary = document.createElement('summary');
		summary.className = 'forecast-mobile-summary';

		const titleWrap = document.createElement('div');
		titleWrap.className = 'forecast-mobile-summary-main';

		const title = document.createElement('strong');
		title.textContent = group.accountName;
		const subtitle = document.createElement('span');
		subtitle.textContent = `${group.rows.length} 個月 · 末餘額 ${formatCurrency(group.latestEndingBalance)}`;
		titleWrap.appendChild(title);
		titleWrap.appendChild(subtitle);

		const indicator = document.createElement('div');
		indicator.className = 'forecast-mobile-summary-side';
		indicator.appendChild(createStatusChip(group.hasNegative ? 'warn' : 'ok', group.hasNegative ? '有月份不足' : '正常'));

		summary.appendChild(titleWrap);
		summary.appendChild(indicator);
		details.appendChild(summary);

		const cards = document.createElement('div');
		cards.className = 'forecast-mobile-cards';

		group.rows.forEach(row => {
			cards.appendChild(buildMobileAccountForecastCard(row));
		});

		details.appendChild(cards);
		details.addEventListener('toggle', () => {
			uiState.hasTouchedAccountForecastExpansion = true;
			if (details.open) {
				expandedAccounts.add(group.accountId);
			} else {
				expandedAccounts.delete(group.accountId);
			}
		});

		refs.accountForecastMobile.appendChild(details);
	});
}

function groupAccountForecastRowsByAccount(accountRows) {
	const groups = new Map();

	state.accounts.forEach(account => {
		groups.set(account.id, {
			accountId: account.id,
			accountName: account.name,
			rows: [],
			latestEndingBalance: account.initialBalance,
			hasNegative: false,
		});
	});

	accountRows.forEach(row => {
		if (!groups.has(row.accountId)) {
			groups.set(row.accountId, {
				accountId: row.accountId,
				accountName: row.accountName,
				rows: [],
				latestEndingBalance: row.endingBalance,
				hasNegative: false,
			});
		}

		const group = groups.get(row.accountId);
		group.accountName = row.accountName;
		group.rows.push(row);
		group.latestEndingBalance = row.endingBalance;
		group.hasNegative = group.hasNegative || row.status === 'negative';
	});

	return Array.from(groups.values()).filter(group => group.rows.length);
}

function buildMobileForecastCard(row) {
	const card = document.createElement('article');
	card.className = 'forecast-mobile-card';
	if (row.status !== 'ok') {
		card.classList.add('is-warning');
	}

	const header = document.createElement('div');
	header.className = 'forecast-mobile-card-header';

	const title = document.createElement('strong');
	title.textContent = row.monthLabel;
	header.appendChild(title);
	header.appendChild(buildTotalForecastStatusChip(row.status));

	const grid = document.createElement('div');
	grid.className = 'forecast-mobile-grid';
	const fields = [
		['起始餘額', formatCurrency(row.startingBalance)],
		['收入', formatCurrency(row.income)],
		['支出', formatCurrency(row.expense)],
		['分期', formatCurrency(row.installment)],
		['淨額', formatCurrency(row.net)],
		['月末總餘額', formatCurrency(row.endingBalance)],
	];

	fields.forEach(([label, value]) => {
		grid.appendChild(buildMobileForecastPair(label, value));
	});

	card.appendChild(header);
	card.appendChild(grid);
	return card;
}

function buildMobileAccountForecastCard(row) {
	const card = document.createElement('article');
	card.className = 'forecast-mobile-card';
	if (row.status === 'negative') {
		card.classList.add('is-warning');
	}

	const header = document.createElement('div');
	header.className = 'forecast-mobile-card-header';

	const title = document.createElement('strong');
	title.textContent = row.monthLabel;
	header.appendChild(title);
	header.appendChild(buildAccountForecastStatusChip(row));

	const grid = document.createElement('div');
	grid.className = 'forecast-mobile-grid';
	const fields = [
		['起始餘額', formatCurrency(row.startingBalance)],
		['收入', formatCurrency(row.income)],
		['支出', formatCurrency(row.expense)],
		['分期', formatCurrency(row.installment)],
		['轉入', formatCurrency(row.transferIn)],
		['轉出', formatCurrency(row.transferOut)],
		['月末餘額', formatCurrency(row.endingBalance)],
	];

	fields.forEach(([label, value]) => {
		grid.appendChild(buildMobileForecastPair(label, value));
	});

	card.appendChild(header);
	card.appendChild(grid);
	return card;
}

function buildMobileForecastPair(label, value) {
	const pair = document.createElement('div');
	pair.className = 'forecast-mobile-pair';

	const pairLabel = document.createElement('span');
	pairLabel.className = 'forecast-mobile-pair-label';
	pairLabel.textContent = label;

	const pairValue = document.createElement('strong');
	pairValue.className = 'forecast-mobile-pair-value';
	pairValue.textContent = value;

	pair.appendChild(pairLabel);
	pair.appendChild(pairValue);
	return pair;
}

function buildStatusCell(status, options = {}) {
	const cell = document.createElement('td');
	applyCellMeta(cell, options);
	const chip = buildTotalForecastStatusChip(status);
	cell.appendChild(chip);
	return cell;
}

function buildAccountStatusCell(row, options = {}) {
	const cell = document.createElement('td');
	applyCellMeta(cell, options);
	const chip = buildAccountForecastStatusChip(row);
	cell.appendChild(chip);
	return cell;
}

function buildTotalForecastStatusChip(status) {
	if (status === 'total-negative') {
		return createStatusChip('warn', '總額不足');
	}
	if (status === 'account-negative') {
		return createStatusChip('caution', '部分帳戶不足');
	}
	return createStatusChip('ok', '正常');
}

function buildAccountForecastStatusChip(row) {
	return createStatusChip(row.status === 'negative' ? 'warn' : 'ok', row.status === 'negative' ? '帳戶不足' : '正常');
}

function getTotalForecastStatusMeta(totalRows) {
	if (totalRows.some(row => row.status === 'total-negative')) {
		return { tone: 'warn', text: '總額不足' };
	}
	if (totalRows.some(row => row.status === 'account-negative')) {
		return { tone: 'caution', text: '部分帳戶不足' };
	}
	return { tone: 'ok', text: '正常' };
}

function createStatusChip(tone, text) {
	const chip = document.createElement('span');
	chip.className = `status-chip ${tone}`;
	chip.textContent = text;
	return chip;
}

function renderChart(totalRows) {
	clearSectionError(refs.chartError);

	if (!refs.chartCanvas || typeof Chart === 'undefined') {
		return;
	}

	const context = typeof refs.chartCanvas.getContext === 'function' ? refs.chartCanvas.getContext('2d') : null;

	if (!context) {
		setSectionError(refs.chartError, '目前無法繪製圖表。');
		return;
	}

	if (forecastChart && typeof forecastChart.destroy === 'function') {
		forecastChart.destroy();
	}

	const palette = getChartPalette();
	forecastChart = new Chart(context, {
		type: 'line',
		data: {
			labels: totalRows.map(row => row.monthLabel),
			datasets: [
				{
					label: '月末總餘額',
					data: totalRows.map(row => row.endingBalance),
					borderColor: palette.line,
					backgroundColor: palette.fill,
					pointBackgroundColor: palette.line,
					pointBorderColor: palette.line,
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
				x: {
					ticks: {
						color: palette.text,
					},
					grid: {
						color: palette.grid,
					},
					border: {
						color: palette.grid,
					},
				},
				y: {
					ticks: {
						color: palette.text,
						callback(value) {
							return formatCurrency(value);
						},
					},
					grid: {
						color: palette.grid,
					},
					border: {
						color: palette.grid,
					},
				},
			},
			plugins: {
				legend: {
					display: false,
				},
				tooltip: {
					displayColors: false,
					backgroundColor: palette.tooltipBackground,
					borderColor: palette.tooltipBorder,
					borderWidth: 1,
					titleColor: palette.tooltipText,
					bodyColor: palette.tooltipText,
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

	sourceState.accounts.forEach(account => {
		balances[account.id] = account.initialBalance;
	});

	for (let monthIndex = 0; monthIndex < sourceState.horizonMonths; monthIndex += 1) {
		const monthDate = new Date(startMonth.getFullYear(), startMonth.getMonth() + monthIndex, 1);
		const monthLabel = formatMonth(monthDate);
		const monthAccountRows = [];
		let anyNegativeAccount = false;

		sourceState.accounts.forEach(account => {
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

			const endingBalance = startingBalance + income + transferIn - expense - installment - transferOut;

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
				status: endingBalance < 0 ? 'negative' : 'ok',
			};

			if (row.status === 'negative') {
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
		negativeTotalCount: totalRows.filter(row => row.status === 'total-negative').length,
		negativeAccountCount: totalRows.filter(row => row.status !== 'ok').length,
	};
}

function buildTotalForecastRow(monthLabel, monthDate, accountRows, anyNegativeAccount) {
	const startingBalance = accountRows.reduce((sum, row) => sum + row.startingBalance, 0);
	const income = accountRows.reduce((sum, row) => sum + row.income, 0);
	const expense = accountRows.reduce((sum, row) => sum + row.expense, 0);
	const installment = accountRows.reduce((sum, row) => sum + row.installment, 0);
	const endingBalance = accountRows.reduce((sum, row) => sum + row.endingBalance, 0);
	const net = income - expense - installment;

	let status = 'ok';
	if (endingBalance < 0) {
		status = 'total-negative';
	} else if (anyNegativeAccount) {
		status = 'account-negative';
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

function formatSyncTimestamp(value) {
	const timeValue = parseTimestamp(value);
	if (!timeValue) {
		return '尚無資料';
	}

	return new Intl.DateTimeFormat(state.locale || DEFAULT_LOCALE, {
		dateStyle: 'medium',
		timeStyle: 'short',
	}).format(new Date(timeValue));
}

function formatCurrency(amount) {
	const formatter = new Intl.NumberFormat(state.locale || DEFAULT_LOCALE, {
		style: 'currency',
		currency: state.currency || DEFAULT_CURRENCY,
		maximumFractionDigits: 0,
	});
	return formatter.format(Number(amount) || 0);
}

function formatCadence(item) {
	if (item.cadence === CADENCE_MONTHLY) {
		return '每月';
	}
	if (item.cadence === CADENCE_YEARLY) {
		return `每年 ${item.monthOfYear} 月`;
	}
	if (item.cadence === CADENCE_ONE_TIME) {
		return `${item.year}/${padDatePart(item.monthOfYear)} 一次`;
	}
	return '未知';
}

function appendCell(row, content, options = {}) {
	const cell = document.createElement('td');
	applyCellMeta(cell, options);
	cell.textContent = content;
	row.appendChild(cell);
	return cell;
}

function applyCellMeta(cell, options = {}) {
	if (!cell || !options) {
		return;
	}

	if (options.label) {
		cell.dataset.label = options.label;
	}

	if (options.className) {
		cell.className = options.className;
	}
}

function appendEmptyRow(tbody, colspan, message) {
	const row = document.createElement('tr');
	const cell = document.createElement('td');
	cell.colSpan = colspan;
	cell.textContent = message;
	row.appendChild(cell);
	tbody.appendChild(row);
}

function buildActionButton(label, styleClass, onClick) {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = `btn ${styleClass}`;
	button.textContent = label;
	button.addEventListener('click', onClick);
	return button;
}

function parseIntegerInput(value) {
	if (value === null || value === undefined || String(value).trim() === '') {
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
	if (value === CADENCE_MONTHLY || value === CADENCE_YEARLY || value === CADENCE_ONE_TIME) {
		return value;
	}
	return null;
}

function parseOneTimeMonth(value) {
	if (typeof value !== 'string') {
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
	return typeof value === 'string' ? value.trim() : '';
}

function parseTimestamp(value) {
	if (typeof value !== 'string' || !value.trim()) {
		return 0;
	}

	const parsed = Date.parse(value);
	return Number.isFinite(parsed) ? parsed : 0;
}

function compareUpdatedAt(left, right) {
	return parseTimestamp(left) - parseTimestamp(right);
}

function makeId() {
	if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
		return crypto.randomUUID();
	}
	return `id_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function setError(node, message) {
	if (!node) {
		return;
	}
	node.textContent = message || '';
}

function clearError(node) {
	setError(node, '');
}

function setStatusMessage(node, message, status) {
	if (!node) {
		return;
	}
	node.textContent = message || '';
	if (node.classList) {
		node.classList.toggle('is-error', status === 'error');
		node.classList.toggle('is-success', status === 'success');
		node.classList.toggle('is-info', status === 'info');
	}
}

function clearStatusMessage(node) {
	setStatusMessage(node, '', '');
}

function getSyncConnectionLabel() {
	if (!syncState.isConfigured) {
		return '尚未設定 Client ID';
	}
	if (!syncState.isNetworkOnline) {
		return '離線中';
	}
	if (syncState.isAuthorizing) {
		return '驗證中';
	}
	if (syncState.isSyncing) {
		return '同步中';
	}
	if (syncState.isConnected) {
		return '已連接 Google Drive';
	}
	return '本機模式';
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
	setStatusMessage(refs.backupStatus, '已匯出目前資料。', 'success');
}

function triggerImportBackup() {
	clearStatusMessage(refs.backupStatus);
	clearError(refs.settingsError);
	refs.importBackupInput.value = '';
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
		setStatusMessage(refs.backupStatus, `已匯入 ${formatImportLabel(file.name)}。`, 'success');
	} catch (error) {
		setStatusMessage(refs.backupStatus, getImportErrorMessage(error), 'error');
	} finally {
		refs.importBackupInput.value = '';
	}
}

function buildExportPayload() {
	return {
		schemaVersion: SCHEMA_VERSION,
		accounts: state.accounts.map(account => ({
			id: account.id,
			name: account.name,
			initialBalance: account.initialBalance,
		})),
		horizonMonths: state.horizonMonths,
		recurringIncomes: state.recurringIncomes.map(item => ({
			id: item.id,
			name: item.name,
			amount: item.amount,
			accountId: item.accountId,
			cadence: item.cadence,
			monthOfYear: item.monthOfYear,
		})),
		recurringExpenses: state.recurringExpenses.map(item => ({
			id: item.id,
			name: item.name,
			amount: item.amount,
			accountId: item.accountId,
			cadence: item.cadence,
			monthOfYear: item.monthOfYear,
			year: item.year,
		})),
		installments: state.installments.map(item => ({ ...item })),
		monthlyTransfers: state.monthlyTransfers.map(item => ({ ...item })),
		currency: state.currency,
		locale: state.locale,
		updatedAt: state.updatedAt,
	};
}

function downloadBackupPayload(payload, filename) {
	if (typeof document === 'undefined' || typeof Blob === 'undefined') {
		throw new Error('目前環境不支援匯出備份。');
	}

	const blob = new Blob([JSON.stringify(payload, null, 2)], {
		type: 'application/json',
	});
	const url = URL.createObjectURL(blob);
	const link = document.createElement('a');
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
		throw new Error('備份檔不是有效的 JSON。');
	}

	return validateImportPayload(payload);
}

function readFileAsText(file) {
	if (file && typeof file.text === 'function') {
		return file.text();
	}

	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = () => resolve(String(reader.result || ''));
		reader.onerror = () => reject(new Error('讀取備份檔失敗。'));
		reader.readAsText(file);
	});
}

function validateImportPayload(payload) {
	if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
		throw new Error('備份格式不正確。');
	}

	if (Object.prototype.hasOwnProperty.call(payload, 'exportVersion')) {
		const exportVersion = Number(payload.exportVersion);
		if (!Object.prototype.hasOwnProperty.call(payload, 'data')) {
			throw new Error('備份缺少 data 欄位。');
		}

		if (exportVersion === LEGACY_EXPORT_VERSION) {
			return validateLegacyImportData(payload.data);
		}
		if (exportVersion === EXPORT_VERSION) {
			return validateCurrentImportData(payload.data);
		}
		throw new Error('不支援的備份版本。');
	}

	if (Object.prototype.hasOwnProperty.call(payload, 'accounts')) {
		return validateCurrentImportData(payload);
	}

	return validateLegacyImportData(payload);
}

function validateLegacyImportData(data) {
	if (!data || typeof data !== 'object' || Array.isArray(data)) {
		throw new Error('舊版備份內容格式不正確。');
	}

	const hasKnownKeys = LEGACY_BACKUP_DATA_KEYS.some(key => Object.prototype.hasOwnProperty.call(data, key));
	if (!hasKnownKeys) {
		throw new Error('找不到可匯入的舊版資料欄位。');
	}

	return sanitizeState(migrateStateToV3(normalizeStateShape(data)));
}

function validateCurrentImportData(data) {
	if (!data || typeof data !== 'object' || Array.isArray(data)) {
		throw new Error('新版備份內容格式不正確。');
	}

	const hasKnownKeys = CURRENT_BACKUP_DATA_KEYS.some(key => Object.prototype.hasOwnProperty.call(data, key));
	if (!hasKnownKeys) {
		throw new Error('找不到可匯入的資料欄位。');
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

function applySyncedState(nextState) {
	state = sanitizeState(nextState);
	clearTransientUiState();
	resetAllForms();
	persistState();
	renderAll();
}

function getImportErrorMessage(error) {
	return error instanceof Error ? error.message : '匯入備份時發生未知錯誤。';
}

function formatImportLabel(fileName) {
	return fileName ? `備份檔 ${fileName}` : '備份檔';
}

function buildBackupFilename() {
	const now = new Date();
	return `${BACKUP_FILENAME_PREFIX}-${now.getFullYear()}${padDatePart(now.getMonth() + 1)}${padDatePart(now.getDate())}-${padDatePart(now.getHours())}${padDatePart(now.getMinutes())}${padDatePart(now.getSeconds())}.json`;
}

function padDatePart(value) {
	return String(value).padStart(2, '0');
}

function persistAndRender() {
	state = sanitizeState(state);
	state.updatedAt = new Date().toISOString();
	persistState();
	handleLocalStateMutation();
	renderAll();
}

function persistState() {
	if (typeof localStorage === 'undefined') {
		return;
	}
	localStorage.setItem(STORAGE_KEY, JSON.stringify(buildExportPayload()));
}

function persistMigratedState(nextState) {
	if (typeof localStorage === 'undefined') {
		return;
	}
	localStorage.setItem(STORAGE_KEY, JSON.stringify(nextState));
}

function loadState() {
	if (typeof localStorage === 'undefined') {
		syncState.hadLocalDataAtStartup = false;
		return createDefaultState();
	}

	const raw = localStorage.getItem(STORAGE_KEY);
	syncState.hadLocalDataAtStartup = Boolean(raw);
	if (!raw) {
		return createDefaultState();
	}

	try {
		const parsed = JSON.parse(raw);
		const normalized = normalizeStateShape(parsed);
		const needsMigration =
			Number(normalized.schemaVersion) <= 2 ||
			!Array.isArray(normalized.accounts) ||
			Object.prototype.hasOwnProperty.call(normalized, 'initialBalance');

		const nextState = sanitizeState(needsMigration ? migrateStateToV3(normalized) : normalized);

		persistMigratedState(nextState);
		return nextState;
	} catch (error) {
		return createDefaultState();
	}
}

function normalizeStateShape(input) {
	if (!input || typeof input !== 'object' || Array.isArray(input)) {
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

	const mainAccount = createAccount(MAIN_ACCOUNT_NAME, sanitizeBoundedInteger(legacyState.initialBalance, 0, 0));

	return {
		schemaVersion: SCHEMA_VERSION,
		accounts: [mainAccount],
		horizonMonths: legacyState.horizonMonths,
		recurringIncomes: migrateLegacyRecurringItems(legacyState.recurringIncomes, mainAccount.id, 'income'),
		recurringExpenses: migrateLegacyRecurringItems(legacyState.recurringExpenses, mainAccount.id, 'expense'),
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
		.map(item => {
			if (!item || typeof item !== 'object') {
				return null;
			}

			const cadence =
				type === 'income'
					? parseIncomeCadence(item.cadence) || CADENCE_MONTHLY
					: parseExpenseCadence(item.cadence) || CADENCE_MONTHLY;

			const migrated = {
				id: typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId(),
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
		.map(item => {
			if (!item || typeof item !== 'object') {
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
				id: typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId(),
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
		currency: typeof source.currency === 'string' && source.currency.trim() ? source.currency.trim() : DEFAULT_CURRENCY,
		locale: typeof source.locale === 'string' && source.locale.trim() ? source.locale.trim() : DEFAULT_LOCALE,
		updatedAt: typeof source.updatedAt === 'string' && source.updatedAt.trim() ? source.updatedAt : new Date().toISOString(),
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
			if (!account || typeof account !== 'object') {
				return null;
			}

			const baseName = normalizeRequiredText(account.name) || (index === 0 ? MAIN_ACCOUNT_NAME : `帳戶 ${index + 1}`);
			const name = makeUniqueAccountName(baseName, usedNames);
			usedNames.push(name);

			let id = typeof account.id === 'string' && account.id.trim() ? account.id.trim() : makeId();
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

	const accountIds = new Set(accounts.map(account => account.id));
	const fallbackAccountId = accounts[0].id;
	const usedIds = new Set();

	return items
		.map(item => {
			if (!item || typeof item !== 'object') {
				return null;
			}

			const name = normalizeRequiredText(item.name);
			if (!name) {
				return null;
			}

			let id = typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId();
			while (usedIds.has(id)) {
				id = makeId();
			}
			usedIds.add(id);

			const cadence = parseIncomeCadence(item.cadence) || CADENCE_MONTHLY;
			const monthOfYear = cadence === CADENCE_YEARLY ? sanitizeBoundedInteger(item.monthOfYear, null, 1, 12) : undefined;

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

	const accountIds = new Set(accounts.map(account => account.id));
	const fallbackAccountId = accounts[0].id;
	const usedIds = new Set();

	return items
		.map(item => {
			if (!item || typeof item !== 'object') {
				return null;
			}

			const name = normalizeRequiredText(item.name);
			if (!name) {
				return null;
			}

			let id = typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId();
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

	const accountIds = new Set(accounts.map(account => account.id));
	const fallbackAccountId = accounts[0].id;
	const usedIds = new Set();

	return items
		.map(item => {
			if (!item || typeof item !== 'object') {
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

			let id = typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId();
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

	const accountIds = new Set(accounts.map(account => account.id));
	const usedIds = new Set();

	return items
		.map(item => {
			if (!item || typeof item !== 'object') {
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

			let id = typeof item.id === 'string' && item.id.trim() ? item.id.trim() : makeId();
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
	if (typeof min === 'number' && parsed < min) {
		return fallback;
	}
	if (typeof max === 'number' && parsed > max) {
		return fallback;
	}
	return parsed;
}

function getCollectionByType(type) {
	if (type === 'income') {
		return state.recurringIncomes;
	}
	if (type === 'expense') {
		return state.recurringExpenses;
	}
	if (type === 'installment') {
		return state.installments;
	}
	return [];
}

function getBatchRefs(type) {
	if (type === 'income') {
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
	if (type === 'expense') {
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
	return state.accounts.find(account => account.id === accountId) || null;
}

function getAccountName(accountId) {
	const account = getAccountById(accountId);
	return account ? account.name : '未知帳戶';
}

function getDefaultAccountId() {
	return state.accounts[0] ? state.accounts[0].id : '';
}

function getDefaultTransferTargetId(sourceAccountId) {
	const alternative = state.accounts.find(account => account.id !== sourceAccountId);
	return alternative ? alternative.id : '';
}

function getValidAccountId(accountId) {
	return state.accounts.some(account => account.id === accountId) ? accountId : '';
}

function getTotalInitialBalance() {
	return state.accounts.reduce((sum, account) => sum + account.initialBalance, 0);
}

function getAccountReferenceCount(accountId) {
	const incomeCount = state.recurringIncomes.filter(item => item.accountId === accountId).length;
	const expenseCount = state.recurringExpenses.filter(item => item.accountId === accountId).length;
	const installmentCount = state.installments.filter(item => item.accountId === accountId).length;
	const transferCount = state.monthlyTransfers.filter(
		item => item.sourceAccountId === accountId || item.targetAccountId === accountId
	).length;
	return incomeCount + expenseCount + installmentCount + transferCount;
}

function getAccountDeletionBlockReason(accountId) {
	const account = getAccountById(accountId);
	if (!account) {
		return '找不到要刪除的帳戶。';
	}
	if (state.accounts.length <= 1) {
		return '至少需要保留一個帳戶。';
	}
	if (account.initialBalance !== 0) {
		return '帳戶期初餘額不為 0，無法刪除。';
	}
	if (getAccountReferenceCount(accountId) > 0) {
		return '帳戶仍被收入、支出、分期或轉帳引用，請先改掛其他帳戶。';
	}
	if (isAccountReferencedInPendingBatches(accountId)) {
		return '帳戶仍在批次歸戶面板中被指定，請先取消或套用。';
	}
	return '';
}

function isAccountReferencedInPendingBatches(accountId) {
	return BATCH_TYPES.some(type => {
		const bucket = batchState[type];
		if (!bucket.isOpen) {
			return false;
		}
		return Object.values(bucket.pendingAccounts).some(value => value === accountId);
	});
}

function clearTransientUiState() {
	BATCH_TYPES.forEach(type => {
		batchState[type].selectedIds.clear();
		batchState[type].isOpen = false;
		batchState[type].pendingAccounts = {};
	});
}

function isDuplicateAccountName(name, excludeId) {
	return state.accounts.some(account => account.id !== excludeId && account.name === name);
}

function loadSyncMeta() {
	if (typeof localStorage === 'undefined') {
		return createDefaultSyncMeta();
	}

	const raw = localStorage.getItem(SYNC_META_STORAGE_KEY);
	if (!raw) {
		return createDefaultSyncMeta();
	}

	try {
		return sanitizeSyncMeta(JSON.parse(raw));
	} catch (error) {
		return createDefaultSyncMeta();
	}
}

function sanitizeDriveSession(input) {
	const source = input || {};
	const accessToken = typeof source.accessToken === 'string' && source.accessToken.trim() ? source.accessToken.trim() : '';
	const accessTokenExpiresAt = Number.isFinite(Number(source.accessTokenExpiresAt))
		? Math.max(0, Number(source.accessTokenExpiresAt))
		: 0;

	return {
		accessToken,
		accessTokenExpiresAt,
	};
}

function loadDriveSession() {
	if (typeof sessionStorage === 'undefined') {
		return createDefaultDriveSession();
	}

	let raw = '';
	try {
		raw = sessionStorage.getItem(DRIVE_SESSION_STORAGE_KEY);
	} catch (error) {
		return createDefaultDriveSession();
	}

	if (!raw) {
		return createDefaultDriveSession();
	}

	try {
		return sanitizeDriveSession(JSON.parse(raw));
	} catch (error) {
		return createDefaultDriveSession();
	}
}

function persistDriveSession() {
	if (typeof sessionStorage === 'undefined') {
		return;
	}

	const nextSession = sanitizeDriveSession({
		accessToken: syncState.accessToken,
		accessTokenExpiresAt: syncState.accessTokenExpiresAt,
	});

	try {
		if (!nextSession.accessToken || !nextSession.accessTokenExpiresAt) {
			sessionStorage.removeItem(DRIVE_SESSION_STORAGE_KEY);
			return;
		}

		sessionStorage.setItem(DRIVE_SESSION_STORAGE_KEY, JSON.stringify(nextSession));
	} catch (error) {
		// Ignore storage failures and keep the in-memory token only.
	}
}

function restoreDriveAccessTokenFromSession() {
	const driveSession = loadDriveSession();
	syncState.accessToken = driveSession.accessToken;
	syncState.accessTokenExpiresAt = driveSession.accessTokenExpiresAt;

	if (!hasValidDriveAccessToken()) {
		clearDriveAccessToken();
		return false;
	}

	syncState.isConnected = true;
	return true;
}

function sanitizeSyncMeta(input) {
	const source = input || {};
	return {
		lastSyncedAt: typeof source.lastSyncedAt === 'string' && source.lastSyncedAt.trim() ? source.lastSyncedAt : '',
		lastKnownCloudUpdatedAt:
			typeof source.lastKnownCloudUpdatedAt === 'string' && source.lastKnownCloudUpdatedAt.trim()
				? source.lastKnownCloudUpdatedAt
				: '',
		lastKnownCloudFileId:
			typeof source.lastKnownCloudFileId === 'string' && source.lastKnownCloudFileId.trim() ? source.lastKnownCloudFileId : '',
		syncStatus: typeof source.syncStatus === 'string' && source.syncStatus.trim() ? source.syncStatus : 'local-only',
		pendingUpload: Boolean(source.pendingUpload),
		lastAuthAttemptAt: typeof source.lastAuthAttemptAt === 'string' && source.lastAuthAttemptAt.trim() ? source.lastAuthAttemptAt : '',
	};
}

function persistSyncMeta() {
	if (typeof localStorage === 'undefined') {
		return;
	}

	localStorage.setItem(SYNC_META_STORAGE_KEY, JSON.stringify(syncState.meta));
}

function updateSyncMeta(patch, options = {}) {
	syncState.meta = sanitizeSyncMeta({
		...syncState.meta,
		...patch,
	});
	persistSyncMeta();
	if (options.render !== false) {
		renderSyncPanel();
	}
}

async function initializeCloudSync() {
	syncState.meta = loadSyncMeta();
	syncState.clientId = resolveGoogleClientId();
	syncState.isConfigured = Boolean(syncState.clientId);
	syncState.isNetworkOnline = isBrowserOnline();

	if (!syncState.isConfigured) {
		setSyncFeedback('not-configured', '尚未設定 Google Client ID，雲端同步目前停用。', 'info');
		return;
	}

	if (!syncState.isNetworkOnline) {
		setSyncFeedback('offline', buildOfflineSyncMessage(), 'info', {
			persistStatus: false,
		});
		return;
	}

	if (isGoogleIdentityReady()) {
		await handleGoogleIdentityScriptReady();
		return;
	}

	if (!refs.googleIdentityScript) {
		setSyncFeedback('gis-missing', '找不到 Google Identity Services 腳本，無法啟用雲端同步。', 'error');
		return;
	}

	if (syncState.hasGoogleScriptListeners) {
		setSyncFeedback('gis-loading', '正在初始化 Google Drive 同步...', 'info', {
			persistStatus: false,
		});
		return;
	}

	syncState.hasGoogleScriptListeners = true;
	setSyncFeedback('gis-loading', '正在初始化 Google Drive 同步...', 'info');
	refs.googleIdentityScript.addEventListener(
		'load',
		() => {
			syncState.hasGoogleScriptListeners = false;
			void handleGoogleIdentityScriptReady();
		},
		{ once: true }
	);
	refs.googleIdentityScript.addEventListener(
		'error',
		() => {
			syncState.hasGoogleScriptListeners = false;
			setSyncFeedback('gis-error', 'Google Identity Services 載入失敗，請稍後再試。', 'error');
		},
		{ once: true }
	);
}

async function handleGoogleIdentityScriptReady() {
	try {
		setupGoogleTokenClient();
		setSyncFeedback('auth-check', '正在檢查 Google Drive 連線狀態...', 'info');

		if (restoreDriveAccessTokenFromSession()) {
			setSyncFeedback('connected', '已恢復 Google Drive 連線。', 'success');
			await synchronizeCloudState({ reason: 'startup' });
			return;
		}

		const authorized = await ensureDriveAuthorization({ interactive: false });
		if (!authorized) {
			return;
		}
		await synchronizeCloudState({ reason: 'startup' });
	} catch (error) {
		setSyncFeedback('gis-error', getSyncFriendlyErrorMessage(error), 'error');
	}
}

function resolveGoogleClientId() {
	if (typeof document === 'undefined') {
		return '';
	}

	const meta = document.querySelector(`meta[name="${GOOGLE_CLIENT_ID_META_NAME}"]`);
	return meta && typeof meta.content === 'string' ? meta.content.trim() : '';
}

function isGoogleIdentityReady() {
	return Boolean(window.google?.accounts?.oauth2);
}

function setupGoogleTokenClient() {
	if (!isGoogleIdentityReady()) {
		throw new Error('Google Identity Services 尚未就緒。');
	}

	syncState.isGoogleReady = true;
	syncState.tokenClient = window.google.accounts.oauth2.initTokenClient({
		client_id: syncState.clientId,
		scope: GOOGLE_DRIVE_SCOPE,
		callback: () => {},
		error_callback: () => {},
	});
}

async function connectGoogleDrive() {
	if (!(await ensureDriveAuthorization({ interactive: true }))) {
		return;
	}

	await synchronizeCloudState({ reason: 'manual-connect' });
}

async function uploadToGoogleDrive() {
	if (!(await ensureDriveAuthorization({ interactive: true }))) {
		return;
	}

	await synchronizeCloudState({ forceUpload: true, reason: 'manual-upload' });
}

async function syncFromGoogleDrive() {
	if (!(await ensureDriveAuthorization({ interactive: true }))) {
		return;
	}

	await synchronizeCloudState({ forceDownload: true, reason: 'manual-download' });
}

async function disconnectGoogleDrive() {
	if (!syncState.isConfigured) {
		return;
	}

	const confirmed =
		typeof window === 'undefined' || typeof window.confirm !== 'function'
			? true
			: window.confirm('確定中斷 Google Drive 連線嗎？本機資料會保留。');
	if (!confirmed) {
		return;
	}

	if (syncUploadTimer) {
		window.clearTimeout(syncUploadTimer);
		syncUploadTimer = null;
	}

	syncState.isAuthorizing = true;
	renderSyncPanel();

	try {
		if (syncState.accessToken && window.google?.accounts?.oauth2?.revoke) {
			await new Promise(resolve => {
				window.google.accounts.oauth2.revoke(syncState.accessToken, () => {
					resolve();
				});
			});
		}
	} finally {
		clearDriveAccessToken();
		syncState.isConnected = false;
		syncState.isAuthorizing = false;
		syncState.hasSessionChanges = false;
		syncState.meta = createDefaultSyncMeta();
		persistSyncMeta();
		setSyncFeedback('local-only', '已中斷 Google Drive 連線，本機資料仍可繼續使用。', 'info', { persistStatus: false });
	}
}

async function ensureDriveAuthorization(options = {}) {
	const interactive = Boolean(options.interactive);
	if (!syncState.isConfigured) {
		setSyncFeedback('not-configured', '尚未設定 Google Client ID，雲端同步目前停用。', 'info');
		return false;
	}
	syncState.isNetworkOnline = isBrowserOnline();
	if (!syncState.isNetworkOnline) {
		setSyncFeedback('offline', buildOfflineSyncMessage(), 'info', {
			persistStatus: false,
		});
		return false;
	}
	if (!syncState.isGoogleReady || !syncState.tokenClient) {
		setSyncFeedback('gis-loading', 'Google Drive 同步仍在初始化，請稍後再試。', 'info');
		return false;
	}
	if (hasValidDriveAccessToken()) {
		syncState.isConnected = true;
		renderSyncPanel();
		return true;
	}

	const prompt = interactive ? 'select_account' : 'none';
	updateSyncMeta(
		{
			lastAuthAttemptAt: new Date().toISOString(),
		},
		{ render: false }
	);

	try {
		const tokenResponse = await requestGoogleAccessToken(prompt);
		syncState.accessToken = tokenResponse.access_token;
		syncState.accessTokenExpiresAt = Date.now() + Math.max(0, Number(tokenResponse.expires_in) || 0) * 1000;
		persistDriveSession();
		syncState.isConnected = true;
		setSyncFeedback('connected', '已連接 Google Drive。', 'success');
		return true;
	} catch (error) {
		clearDriveAccessToken();
		syncState.isConnected = false;
		if (interactive) {
			setSyncFeedback('auth-required', getSyncFriendlyErrorMessage(error), 'error');
		} else if (syncState.meta.pendingUpload) {
			setSyncFeedback('pending-auth', '本機有尚未同步的資料，請連接 Google Drive 後續傳。', 'info');
		} else {
			setSyncFeedback('local-only', '目前使用本機資料，可隨時連接 Google Drive。', 'info');
		}
		return false;
	}
}

function hasValidDriveAccessToken() {
	return Boolean(
		syncState.accessToken &&
		syncState.accessTokenExpiresAt &&
		Date.now() < syncState.accessTokenExpiresAt - ACCESS_TOKEN_EXPIRY_BUFFER_MS
	);
}

function clearDriveAccessToken() {
	syncState.accessToken = '';
	syncState.accessTokenExpiresAt = 0;
	persistDriveSession();
}

function requestGoogleAccessToken(promptValue) {
	return new Promise((resolve, reject) => {
		if (!syncState.tokenClient) {
			reject(new Error('Google Token Client 尚未初始化。'));
			return;
		}

		syncState.isAuthorizing = true;
		renderSyncPanel();

		syncState.tokenClient.callback = response => {
			syncState.isAuthorizing = false;
			renderSyncPanel();
			if (response && !response.error && response.access_token) {
				resolve(response);
				return;
			}

			const error = new Error(response?.error_description || response?.error || 'Google 授權失敗。');
			error.code = response?.error || '';
			reject(error);
		};

		syncState.tokenClient.error_callback = error => {
			syncState.isAuthorizing = false;
			renderSyncPanel();
			const nextError = error instanceof Error ? error : new Error(error?.type || 'Google 授權失敗。');
			nextError.code = error?.type || '';
			reject(nextError);
		};

		syncState.tokenClient.requestAccessToken({
			prompt: promptValue,
		});
	});
}

function setSyncFeedback(statusCode, message, tone, options = {}) {
	syncState.statusCode = statusCode;
	syncState.statusMessage = message;
	syncState.statusTone = tone;
	if (options.persistStatus !== false) {
		updateSyncMeta(
			{
				syncStatus: statusCode,
			},
			{ render: false }
		);
	}
	if (options.render !== false) {
		renderSyncPanel();
	}
}

async function synchronizeCloudState(options = {}) {
	const forceDownload = Boolean(options.forceDownload);
	const forceUpload = Boolean(options.forceUpload);

	syncState.isNetworkOnline = isBrowserOnline();
	if (syncState.isSyncing) {
		return false;
	}
	if (!syncState.isNetworkOnline) {
		setSyncFeedback('offline', buildOfflineSyncMessage(), 'info', {
			persistStatus: false,
		});
		return false;
	}

	syncState.isSyncing = true;
	setSyncFeedback('syncing', '正在與 Google Drive 同步...', 'info', {
		render: false,
	});
	renderSyncPanel();

	try {
		const cloudFile = await getCloudFileMetadata();

		if (!cloudFile) {
			updateSyncMeta(
				{
					lastKnownCloudFileId: '',
					lastKnownCloudUpdatedAt: '',
				},
				{ render: false }
			);

			if (forceDownload) {
				setSyncFeedback('cloud-empty', 'Google Drive 尚無可下載的資料。', 'info');
				return false;
			}

			if (forceUpload || shouldCreateCloudFileOnFirstSync() || syncState.hasSessionChanges || syncState.meta.pendingUpload) {
				await uploadLocalStateToCloud({ reason: options.reason || 'sync', createIfMissing: true });
				return true;
			}

			setSyncFeedback('cloud-empty', 'Google Drive 尚無資料，本機仍可正常使用。', 'info');
			return false;
		}

		updateSyncMeta(
			{
				lastKnownCloudFileId: cloudFile.id,
				lastKnownCloudUpdatedAt: cloudFile.updatedAt,
			},
			{ render: false }
		);

		if (!syncState.hadLocalDataAtStartup && !syncState.hasSessionChanges && !syncState.meta.pendingUpload && !forceUpload) {
			await applyCloudStateFromDrive(cloudFile);
			return true;
		}

		if (
			forceUpload &&
			compareUpdatedAt(cloudFile.updatedAt, state.updatedAt) > 0 &&
			typeof window !== 'undefined' &&
			typeof window.confirm === 'function'
		) {
			const confirmed = window.confirm('雲端資料時間較新，仍要用本機資料覆蓋雲端嗎？');
			if (!confirmed) {
				setSyncFeedback('upload-cancelled', '已取消上傳到 Google Drive。', 'info');
				return false;
			}
		}

		if (
			forceDownload &&
			compareUpdatedAt(state.updatedAt, cloudFile.updatedAt) > 0 &&
			typeof window !== 'undefined' &&
			typeof window.confirm === 'function'
		) {
			const confirmed = window.confirm('本機資料時間較新，仍要用雲端資料覆蓋本機嗎？');
			if (!confirmed) {
				setSyncFeedback('download-cancelled', '已取消從 Google Drive 覆蓋本機資料。', 'info');
				return false;
			}
		}

		if (forceUpload) {
			await uploadLocalStateToCloud({
				fileId: cloudFile.id,
				reason: options.reason || 'manual-upload',
			});
			return true;
		}

		if (forceDownload) {
			await applyCloudStateFromDrive(cloudFile);
			return true;
		}

		const comparison = compareUpdatedAt(cloudFile.updatedAt, state.updatedAt);
		if (comparison > 0) {
			await applyCloudStateFromDrive(cloudFile);
			return true;
		}

		if (comparison < 0) {
			if (syncState.hasSessionChanges || syncState.meta.pendingUpload) {
				await uploadLocalStateToCloud({
					fileId: cloudFile.id,
					reason: options.reason || 'pending-upload',
				});
				return true;
			}

			setSyncFeedback('local-newer', '本機資料較新，尚未推送到雲端。', 'info');
			return false;
		}

		syncState.hasSessionChanges = false;
		updateSyncMeta(
			{
				lastSyncedAt: new Date().toISOString(),
				pendingUpload: false,
				syncStatus: 'synced',
			},
			{ render: false }
		);
		setSyncFeedback('synced', '本機與 Google Drive 已同步。', 'success');
		return true;
	} catch (error) {
		setSyncFeedback('sync-error', getSyncFriendlyErrorMessage(error), 'error');
		return false;
	} finally {
		syncState.isSyncing = false;
		renderSyncPanel();
		maybeContinuePendingCloudUpload();
	}
}

async function getCloudFileMetadata() {
	const params = new URLSearchParams({
		spaces: 'appDataFolder',
		pageSize: '1',
		orderBy: 'modifiedTime desc',
		fields: 'files(id,name,modifiedTime,appProperties)',
		q: `name='${GOOGLE_DRIVE_FILE_NAME}' and 'appDataFolder' in parents and trashed=false`,
	});
	const payload = await driveRequestJson(`${GOOGLE_DRIVE_API_URL}?${params.toString()}`);
	const firstFile = Array.isArray(payload.files) ? payload.files[0] : null;
	return firstFile ? normalizeCloudFileMetadata(firstFile) : null;
}

function normalizeCloudFileMetadata(file) {
	const appUpdatedAt = typeof file?.appProperties?.updatedAt === 'string' ? file.appProperties.updatedAt : '';
	return {
		id: typeof file?.id === 'string' ? file.id : '',
		updatedAt: appUpdatedAt || file?.modifiedTime || '',
		modifiedTime: typeof file?.modifiedTime === 'string' ? file.modifiedTime : '',
	};
}

async function fetchCloudStatePayload(fileId) {
	const response = await driveRequest(`${GOOGLE_DRIVE_API_URL}/${encodeURIComponent(fileId)}?alt=media`);
	return response.json();
}

async function applyCloudStateFromDrive(file) {
	const payload = await fetchCloudStatePayload(file.id);
	const nextState = validateImportPayload(payload);
	applySyncedState(nextState);
	syncState.hasSessionChanges = false;
	updateSyncMeta(
		{
			lastKnownCloudFileId: file.id,
			lastKnownCloudUpdatedAt: nextState.updatedAt,
			lastSyncedAt: new Date().toISOString(),
			pendingUpload: false,
			syncStatus: 'cloud-applied',
		},
		{ render: false }
	);
	setSyncFeedback('cloud-applied', '已套用 Google Drive 上較新的資料。', 'success');
}

async function uploadLocalStateToCloud(options = {}) {
	const payload = buildExportPayload();
	const fileId = options.fileId || syncState.meta.lastKnownCloudFileId || '';
	const uploadedUpdatedAt = payload.updatedAt;
	const metadata = buildDriveFileMetadata(uploadedUpdatedAt, !fileId);
	const boundary = `batch_${makeId()}`;
	const url = fileId
		? `${GOOGLE_DRIVE_UPLOAD_URL}/${encodeURIComponent(fileId)}?uploadType=multipart&fields=id,modifiedTime,appProperties`
		: `${GOOGLE_DRIVE_UPLOAD_URL}?uploadType=multipart&fields=id,modifiedTime,appProperties`;

	try {
		const response = await driveRequestJson(url, {
			method: fileId ? 'PATCH' : 'POST',
			headers: {
				'Content-Type': `multipart/related; boundary=${boundary}`,
			},
			body: buildDriveMultipartBody(boundary, metadata, payload),
		});
		const cloudFile = normalizeCloudFileMetadata(response);
		const unchangedDuringUpload = state.updatedAt === uploadedUpdatedAt;
		syncState.hasSessionChanges = !unchangedDuringUpload;
		updateSyncMeta(
			{
				lastKnownCloudFileId: cloudFile.id,
				lastKnownCloudUpdatedAt: uploadedUpdatedAt,
				lastSyncedAt: new Date().toISOString(),
				pendingUpload: !unchangedDuringUpload,
				syncStatus: unchangedDuringUpload ? 'synced' : 'pending-upload',
			},
			{ render: false }
		);
		setSyncFeedback(
			unchangedDuringUpload ? 'synced' : 'pending-upload',
			unchangedDuringUpload ? '已將本機資料同步到 Google Drive。' : '本機仍有新變更，已排入下一次自動上傳。',
			unchangedDuringUpload ? 'success' : 'info'
		);

		if (!unchangedDuringUpload) {
			scheduleCloudUpload();
		}

		return cloudFile;
	} catch (error) {
		if (fileId && error?.statusCode === 404) {
			updateSyncMeta(
				{
					lastKnownCloudFileId: '',
					lastKnownCloudUpdatedAt: '',
				},
				{ render: false }
			);
			return uploadLocalStateToCloud({
				...options,
				fileId: '',
			});
		}
		throw error;
	}
}

function shouldCreateCloudFileOnFirstSync() {
	return !syncState.hadLocalDataAtStartup && !syncState.hasSessionChanges && !syncState.meta.pendingUpload;
}

function buildDriveFileMetadata(updatedAt, includeParent) {
	const metadata = {
		name: GOOGLE_DRIVE_FILE_NAME,
		mimeType: 'application/json',
		appProperties: {
			appId: APP_ID,
			updatedAt,
		},
	};

	if (includeParent) {
		metadata.parents = ['appDataFolder'];
	}

	return metadata;
}

function buildDriveMultipartBody(boundary, metadata, payload) {
	return [
		`--${boundary}\r\n`,
		'Content-Type: application/json; charset=UTF-8\r\n\r\n',
		JSON.stringify(metadata),
		'\r\n',
		`--${boundary}\r\n`,
		'Content-Type: application/json; charset=UTF-8\r\n\r\n',
		JSON.stringify(payload),
		'\r\n',
		`--${boundary}--`,
	].join('');
}

async function driveRequestJson(url, options = {}) {
	const response = await driveRequest(url, options);
	if (response.status === 204) {
		return {};
	}
	return response.json();
}

async function driveRequest(url, options = {}) {
	if (!syncState.accessToken) {
		throw new Error('尚未取得 Google 授權。');
	}
	syncState.isNetworkOnline = isBrowserOnline();
	if (!syncState.isNetworkOnline) {
		throw createOfflineSyncError();
	}

	const headers = new Headers(options.headers || {});
	headers.set('Authorization', `Bearer ${syncState.accessToken}`);

	const response = await fetch(url, {
		...options,
		headers,
	});

	if (response.status === 401 || response.status === 403) {
		handleDriveAuthFailure();
		const error = new Error('Google 授權已失效，請重新連接。');
		error.statusCode = response.status;
		throw error;
	}

	if (!response.ok) {
		throw await parseDriveError(response);
	}

	return response;
}

async function parseDriveError(response) {
	let message = `Google Drive 同步失敗（${response.status}）。`;

	try {
		const payload = await response.json();
		if (typeof payload?.error?.message === 'string' && payload.error.message.trim()) {
			message = payload.error.message.trim();
		}
	} catch (error) {
		// Keep the fallback message when the response body is not JSON.
	}

	const nextError = new Error(message);
	nextError.statusCode = response.status;
	return nextError;
}

function handleDriveAuthFailure() {
	clearDriveAccessToken();
	syncState.isConnected = false;
	updateSyncMeta(
		{
			pendingUpload: syncState.meta.pendingUpload || syncState.hasSessionChanges,
			syncStatus: 'auth-required',
		},
		{ render: false }
	);
	setSyncFeedback('auth-required', 'Google 授權已失效，請重新連接。', 'error');
}

function handleLocalStateMutation() {
	if (!syncState.isConfigured) {
		return;
	}
	syncState.isNetworkOnline = isBrowserOnline();

	syncState.hasSessionChanges = true;
	updateSyncMeta(
		{
			pendingUpload: true,
			syncStatus: 'pending-upload',
		},
		{ render: false }
	);

	if (!syncState.isNetworkOnline) {
		setSyncFeedback('offline', buildOfflineSyncMessage(), 'info', {
			persistStatus: false,
		});
		return;
	}

	if (syncState.isGoogleReady) {
		setSyncFeedback('pending-upload', '本機資料已更新，將自動同步到 Google Drive。', 'info');
		scheduleCloudUpload();
		return;
	}

	setSyncFeedback('pending-upload', '本機資料已更新，待 Google Drive 同步初始化後續傳。', 'info');
}

function scheduleCloudUpload() {
	syncState.isNetworkOnline = isBrowserOnline();
	if (!syncState.isConfigured || !syncState.isNetworkOnline) {
		return;
	}

	if (syncUploadTimer) {
		window.clearTimeout(syncUploadTimer);
	}

	syncUploadTimer = window.setTimeout(() => {
		syncUploadTimer = null;
		void runScheduledCloudUpload();
	}, SYNC_UPLOAD_DEBOUNCE_MS);
}

async function runScheduledCloudUpload() {
	syncState.isNetworkOnline = isBrowserOnline();
	if (!syncState.isConfigured || !syncState.isNetworkOnline) {
		return;
	}

	if (syncState.isSyncing || syncState.isAuthorizing) {
		scheduleCloudUpload();
		return;
	}

	if (!(await ensureDriveAuthorization({ interactive: false }))) {
		updateSyncMeta(
			{
				pendingUpload: true,
			},
			{ render: false }
		);
		return;
	}

	try {
		await uploadLocalStateToCloud({
			reason: 'auto-upload',
		});
	} catch (error) {
		setSyncFeedback('sync-error', getSyncFriendlyErrorMessage(error), 'error');
	}
}

function maybeContinuePendingCloudUpload() {
	syncState.isNetworkOnline = isBrowserOnline();
	if (syncState.meta.pendingUpload && syncState.isConfigured && syncState.isNetworkOnline && !syncState.isSyncing && !syncState.isAuthorizing) {
		scheduleCloudUpload();
	}
}

function getSyncFriendlyErrorMessage(error) {
	const message = error instanceof Error ? error.message : '';
	if (!syncState.isNetworkOnline || message.includes('Failed to fetch') || message.includes('NetworkError')) {
		return '目前離線，恢復連線後可再同步 Google Drive。';
	}
	if (!message) {
		return 'Google Drive 同步失敗。';
	}
	if (error?.statusCode === 404) {
		return 'Google Drive 上找不到同步檔案，請重新上傳。';
	}
	if (message.includes('popup_closed')) {
		return '你已取消 Google 授權。';
	}
	if (message.includes('popup_failed_to_open')) {
		return 'Google 授權視窗無法開啟，請確認瀏覽器未封鎖彈窗。';
	}
	if (message.includes('immediate_failed')) {
		return '目前無法自動恢復 Google 授權，可手動連接後再同步。';
	}
	return message;
}

function createOfflineSyncError() {
	const error = new Error('目前離線，恢復連線後可再同步 Google Drive。');
	error.code = 'offline';
	return error;
}
