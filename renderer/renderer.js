const DEFAULT_CUSTOMIZATION = Object.freeze({
  includeSummary: true,
  includeDetail: false,
  includeCharts: true,
  includeMovements: true,
  includeObservations: false,
  includeUniverse: true,
  csv: {
    includePoResumen: true,
    includeRemisiones: true,
    includeFacturas: true,
    includeTotales: true,
    includeUniverseInfo: true
  }
});

function createDefaultCustomization() {
  return {
    ...DEFAULT_CUSTOMIZATION,
    csv: { ...DEFAULT_CUSTOMIZATION.csv }
  };
}

const state = {
  empresas: [],
  pos: [],
  selectedEmpresa: '',
  selectedPoIds: [],
  selectedPoDetails: new Map(),
  manualExtensions: new Map(),
  charts: new Map(),
  summary: null,
  users: [],
  reportSettings: null,
  reportEngines: [],
  selectedEngine: '',
  reportFormatsCatalog: [],
  selectedFormat: 'pdf',
  customization: createDefaultCustomization(),
  customReportName: readSession('porpt-custom-report-name', ''),
  editingUserId: null,
  universeFilterMode: 'global',
  universeStartDate: '',
  universeEndDate: '',
  universeFormat: readSession('porpt-universe-format', 'pdf'),
  universeCustomName: readSession('porpt-universe-report-name', ''),
  isGeneratingReport: false,
  userCompanySelection: new Set(),
  overview: {
    items: [],
    filteredItems: [],
    selection: new Map(),
    itemIndex: new Map(),
    filters: {
      mode: 'global',
      startDate: '',
      endDate: '',
      search: '',
      alertsOnly: false
    },
    loading: false
  },
  dashboardPanel: {
    isOpen: false,
    userCollapsed: false,
    isExpanded: false
  }
};

const searchableSelectStates = new WeakMap();

const ALERT_TYPE_CLASS = {
  success: 'success',
  danger: 'danger',
  warning: 'warning',
  alerta: 'warning',
  info: 'info'
};

const ALERT_SETTINGS = {
  cooldownMs: 15000,
  maxVisible: 4,
  autoHideMs: 7000
};

const alertRegistry = new Map();
const ALERT_UI_ENABLED = false;

function mapAlertType(type) {
  return ALERT_TYPE_CLASS[type] || 'info';
}

function pruneAlertHistory(message, timestamp) {
  setTimeout(() => {
    const stored = alertRegistry.get(message);
    if (stored && stored === timestamp) {
      alertRegistry.delete(message);
    }
  }, ALERT_SETTINGS.cooldownMs);
}

function showAlert(message, type = 'info', delay = ALERT_SETTINGS.autoHideMs) {
  const normalizedMessage = typeof message === 'string' ? message.trim() : String(message ?? '');
  if (!ALERT_UI_ENABLED) {
    const logger = type === 'danger' ? console.error : type === 'warning' ? console.warn : console.log;
    logger(normalizedMessage);
    return;
  }

  const container = document.getElementById('alertBanner');
  if (!container) {
    const logger = type === 'danger' ? console.error : type === 'warning' ? console.warn : console.log;
    logger(normalizedMessage);
    return;
  }

  const now = Date.now();
  const lastShown = alertRegistry.get(normalizedMessage);
  if (lastShown && now - lastShown < ALERT_SETTINGS.cooldownMs) {
    return;
  }

  while (container.children.length >= ALERT_SETTINGS.maxVisible) {
    container.removeChild(container.firstElementChild);
  }

  const wrapper = document.createElement('div');
  wrapper.className = `alert alert-${mapAlertType(type)} alert-dismissible fade show shadow`;
  wrapper.setAttribute('role', 'alert');
  wrapper.dataset.message = normalizedMessage;
  const safeMessage = escapeHtml(normalizedMessage).replace(/\n/g, '<br>');
  wrapper.innerHTML = `
    <div class="d-flex align-items-start gap-2">
      <div class="fw-medium flex-grow-1">${safeMessage}</div>
      <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Cerrar"></button>
    </div>
  `;
  container.appendChild(wrapper);
  window.bootstrap?.Alert?.getOrCreateInstance(wrapper);

  alertRegistry.set(normalizedMessage, now);
  pruneAlertHistory(normalizedMessage, now);

  wrapper.addEventListener('closed.bs.alert', () => {
    wrapper.remove();
  });

  const autoHide = typeof delay === 'number' ? delay : ALERT_SETTINGS.autoHideMs;
  if (autoHide > 0) {
    setTimeout(() => {
      if (!wrapper.isConnected) return;
      const instance = window.bootstrap?.Alert?.getOrCreateInstance(wrapper);
      if (instance) {
        instance.close();
      } else {
        wrapper.remove();
      }
    }, autoHide);
  }
}

function saveSession(key, value) {
  sessionStorage.setItem(key, JSON.stringify(value));
}

function readSession(key, defaultValue) {
  try {
    const stored = sessionStorage.getItem(key);
    return stored ? JSON.parse(stored) : defaultValue;
  } catch (err) {
    return defaultValue;
  }
}

function isAdmin() {
  return !!readSession('porpt-is-admin', false);
}

function adminFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  if (isAdmin()) {
    headers.set('X-PORPT-ADMIN', 'true');
  }
  return fetch(url, { ...options, headers });
}

function getInputValue(id) {
  const element = document.getElementById(id);
  return element ? element.value.trim() : '';
}

function escapeHtml(value) {
  if (typeof value !== 'string') {
    return value ?? '';
  }
  return value
    .replace(/&/gu, '&amp;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
    .replace(/"/gu, '&quot;')
    .replace(/'/gu, '&#39;');
}

function sanitizeFileName(value, fallback) {
  const base = typeof value === 'string' ? value.trim() : '';
  if (!base) {
    return fallback;
  }
  const sanitized = base.replace(/[^0-9a-zA-Z-_]+/gu, '_');
  return sanitized || fallback;
}

function resetChartCanvas(canvas) {
  if (!canvas) return;
  const context = canvas.getContext?.('2d');
  if (context) {
    context.clearRect(0, 0, canvas.width || 0, canvas.height || 0);
  }
  canvas.removeAttribute('height');
  canvas.style.removeProperty('height');
  canvas.style.removeProperty('width');
  const wrapper = canvas.closest?.('.chart-wrapper');
  if (wrapper) {
    wrapper.style.removeProperty('minHeight');
    wrapper.style.removeProperty('maxHeight');
    wrapper.style.removeProperty('height');
  }
}

function destroyCharts() {
  state.charts.forEach(chart => {
    if (chart && typeof chart.destroy === 'function') {
      const canvas = chart.canvas || chart.ctx?.canvas;
      chart.destroy();
      resetChartCanvas(canvas);
    }
  });
  state.charts.clear();
}

function syncDashboardToggleState() {
  const toggleBtn = document.getElementById('dashboardToggleBtn');
  const panel = document.getElementById('dashboardContent');
  if (!toggleBtn || !panel) {
    return;
  }
  const available = !panel.classList.contains('d-none');
  const open = available && panel.classList.contains('dashboard-panel-open');
  toggleBtn.disabled = !available;
  toggleBtn.setAttribute('aria-disabled', (!available).toString());
  toggleBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
  const openLabel = toggleBtn.getAttribute('data-open-label') || 'Abrir dashboard';
  const closeLabel = toggleBtn.getAttribute('data-close-label') || 'Ocultar dashboard';
  toggleBtn.textContent = open ? closeLabel : openLabel;
  syncDashboardExpandButton();
}

function resetDashboardPanelScroll() {
  const panelBody = document.getElementById('dashboardPanelContent');
  if (!panelBody) {
    return;
  }
  panelBody.scrollTop = 0;
}

function syncDashboardExpandButton() {
  const expandBtn = document.getElementById('dashboardExpandBtn');
  const panel = document.getElementById('dashboardContent');
  if (!expandBtn || !panel) {
    return;
  }
  const available = !panel.classList.contains('d-none');
  const expanded = available && panel.classList.contains('dashboard-panel-expanded');
  expandBtn.disabled = !available;
  expandBtn.setAttribute('aria-disabled', (!available).toString());
  expandBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
  expandBtn.setAttribute('aria-pressed', expanded ? 'true' : 'false');
  const expandLabel = expandBtn.getAttribute('data-expand-label') || 'Expandir';
  const collapseLabel = expandBtn.getAttribute('data-collapse-label') || 'Minimizar';
  expandBtn.textContent = expanded ? collapseLabel : expandLabel;
  expandBtn.title = expanded ? 'Minimizar panel' : 'Expandir panel';
}

function setDashboardPanelOpen(open, options = {}) {
  const panel = document.getElementById('dashboardContent');
  if (!panel || panel.classList.contains('d-none')) {
    state.dashboardPanel.isOpen = false;
    state.dashboardPanel.isExpanded = false;
    syncDashboardToggleState();
    return;
  }
  const { userAction = false } = options;
  panel.classList.toggle('dashboard-panel-open', open);
  panel.setAttribute('aria-hidden', open ? 'false' : 'true');
  document.body.classList.toggle('dashboard-panel-open', open);
  state.dashboardPanel.isOpen = open;
  if (userAction) {
    state.dashboardPanel.userCollapsed = !open;
  } else if (open) {
    state.dashboardPanel.userCollapsed = false;
  }
  if (!open) {
    panel.classList.remove('dashboard-panel-expanded');
    document.body.classList.remove('dashboard-panel-expanded');
    state.dashboardPanel.isExpanded = false;
  } else if (state.dashboardPanel.isExpanded) {
    panel.classList.add('dashboard-panel-expanded');
    document.body.classList.add('dashboard-panel-expanded');
  }
  syncDashboardToggleState();
}

function setDashboardPanelExpanded(expanded) {
  const panel = document.getElementById('dashboardContent');
  if (!panel || panel.classList.contains('d-none')) {
    state.dashboardPanel.isExpanded = false;
    document.body.classList.remove('dashboard-panel-expanded');
    syncDashboardToggleState();
    return;
  }
  const shouldExpand = Boolean(expanded);
  panel.classList.toggle('dashboard-panel-expanded', shouldExpand);
  document.body.classList.toggle('dashboard-panel-expanded', shouldExpand);
  state.dashboardPanel.isExpanded = shouldExpand;
  syncDashboardToggleState();
  if (shouldExpand) {
    state.charts.forEach(chart => {
      if (chart && typeof chart.resize === 'function') {
        try {
          chart.resize();
        } catch (error) {
          console.warn('No se pudo reajustar la gráfica al expandir el dashboard:', error);
        }
      }
    });
  }
}

function showDashboardPanel(options = {}) {
  const panel = document.getElementById('dashboardContent');
  if (!panel) return;
  panel.classList.remove('d-none');
  resetDashboardPanelScroll();
  const { autoOpen = false } = options;
  const shouldOpen = autoOpen
    ? !state.dashboardPanel.userCollapsed
    : state.dashboardPanel.isOpen && !state.dashboardPanel.userCollapsed;
  requestAnimationFrame(() => {
    setDashboardPanelOpen(shouldOpen, { userAction: false });
  });
}

function hideDashboardPanel() {
  const panel = document.getElementById('dashboardContent');
  if (!panel) return;
  panel.classList.add('d-none');
  panel.classList.remove('dashboard-panel-open');
  panel.classList.remove('dashboard-panel-expanded');
  panel.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('dashboard-panel-open');
  document.body.classList.remove('dashboard-panel-expanded');
  state.dashboardPanel.isOpen = false;
  state.dashboardPanel.userCollapsed = false;
  state.dashboardPanel.isExpanded = false;
  syncDashboardToggleState();
}

function initializeDashboardPanelControls() {
  const toggleBtn = document.getElementById('dashboardToggleBtn');
  if (toggleBtn) {
    toggleBtn.addEventListener('click', () => {
      const panel = document.getElementById('dashboardContent');
      if (!panel || panel.classList.contains('d-none')) {
        return;
      }
      const willOpen = !panel.classList.contains('dashboard-panel-open');
      setDashboardPanelOpen(willOpen, { userAction: true });
    });
  }
  const expandBtn = document.getElementById('dashboardExpandBtn');
  if (expandBtn) {
    expandBtn.addEventListener('click', () => {
      const panel = document.getElementById('dashboardContent');
      if (!panel || panel.classList.contains('d-none')) {
        return;
      }
      if (!panel.classList.contains('dashboard-panel-open')) {
        setDashboardPanelOpen(true, { userAction: true });
      }
      const willExpand = !panel.classList.contains('dashboard-panel-expanded');
      setDashboardPanelExpanded(willExpand);
    });
  }
  const closeBtn = document.getElementById('dashboardCloseBtn');
  closeBtn?.addEventListener('click', () => {
    setDashboardPanelOpen(false, { userAction: true });
  });
  const tabs = document.getElementById('dashboardPanelTabs');
  tabs?.addEventListener('shown.bs.tab', () => {
    resetDashboardPanelScroll();
    state.charts.forEach(chart => {
      if (chart && typeof chart.resize === 'function') {
        try {
          chart.resize();
        } catch (error) {
          console.warn('No se pudo reajustar la gráfica:', error);
        }
      }
    });
  });
  syncDashboardToggleState();
}

function normalizeNumber(value) {
  const number = Number(value ?? 0);
  return Number.isFinite(number) ? number : 0;
}

function roundTo(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round(normalizeNumber(value) * factor) / factor;
}

function clampPercentage(value, decimals = 2) {
  const normalized = roundTo(value, decimals);
  return Math.min(100, Math.max(0, normalized));
}

function computeResponsiveChartHeight(itemCount) {
  const count = Number(itemCount) || 0;
  if (count <= 1) {
    return 220;
  }
  const base = 220;
  const perItem = 26;
  const computed = base + Math.max(0, count - 4) * perItem;
  return Math.min(420, Math.max(200, computed));
}

function formatCurrency(value) {
  const amount = normalizeNumber(value);
  return amount.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getSearchableSelectState(select) {
  if (!select) return null;
  let state = searchableSelectStates.get(select);
  if (!state) {
    state = {
      text: '',
      placeholder: select.getAttribute('data-placeholder') || 'Selecciona una opción',
      filterInput: null
    };
    searchableSelectStates.set(select, state);
  }
  return state;
}

function getSearchablePlaceholder(select) {
  if (!select) return '';
  const state = getSearchableSelectState(select);
  const hasItems = select.dataset.searchableHasItems !== 'false';
  const emptyLabel = select.dataset.searchableEmptyLabel;
  if (!hasItems && emptyLabel) {
    return emptyLabel;
  }
  return state?.placeholder || select.getAttribute('placeholder') || '';
}

function updateTomSelectPlaceholder(control, placeholder) {
  if (!control) return;
  if (typeof placeholder === 'string') {
    control.settings.placeholder = placeholder;
    if (control.control_input) {
      control.control_input.placeholder = placeholder;
    }
    if (control.dropdown_input) {
      control.dropdown_input.placeholder = placeholder;
    }
  }
}

function refreshTomSelectOptions(select) {
  if (!select?.tomselect) return;
  const control = select.tomselect;
  const options = Array.from(select.options || [])
    .filter(option => option.dataset?.placeholderOption !== 'true')
    .map(option => ({
      value: option.value,
      text: option.textContent,
      disabled: option.disabled
    }));
  const placeholder = getSearchablePlaceholder(select);
  try {
    if (typeof control.clearOptions === 'function') {
      control.clearOptions();
    }
    if (options.length) {
      if (typeof control.addOptions === 'function') {
        control.addOptions(options);
      } else if (typeof control.addOption === 'function') {
        options.forEach(option => control.addOption(option));
      }
    }
    updateTomSelectPlaceholder(control, placeholder);
    const currentValue = select.value;
    if (currentValue) {
      control.setValue(currentValue, true);
    } else {
      control.clear(true);
    }
    if (typeof control.refreshOptions === 'function') {
      control.refreshOptions(false);
    }
  } catch (error) {
    console.error('Error sincronizando TomSelect', error);
  }
}

function ensureSearchableFilterInput(select) {
  if (!select) return null;
  if (select.tomselect) return null;
  const state = getSearchableSelectState(select);
  if (!state) return null;
  if (state.filterInput && state.filterInput.isConnected) {
    state.filterInput.value = state.text || '';
    return state.filterInput;
  }
  const filterInput = document.createElement('input');
  filterInput.type = 'search';
  filterInput.className = 'form-control form-control-sm mb-2 searchable-select-filter';
  filterInput.placeholder = 'Escribe para filtrar';
  filterInput.setAttribute('aria-label', 'Filtrar opciones');
  filterInput.autocomplete = 'off';
  filterInput.spellcheck = false;
  filterInput.addEventListener('input', () => {
    applySearchableSelectFilter(select, filterInput.value);
  });
  filterInput.addEventListener('keydown', event => {
    if (event.key === 'Escape') {
      event.preventDefault();
      filterInput.value = '';
      applySearchableSelectFilter(select, '');
      return;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      select.focus();
    }
  });
  const parent = select.parentElement;
  if (parent?.classList.contains('input-group')) {
    parent.parentElement?.insertBefore(filterInput, parent);
  } else {
    parent?.insertBefore(filterInput, select);
  }
  state.filterInput = filterInput;
  filterInput.value = state.text || '';
  return filterInput;
}

function updateSearchableSelectPlaceholder(select, matches = 0) {
  if (!select) return;
  if (select.tomselect) {
    updateTomSelectPlaceholder(select.tomselect, getSearchablePlaceholder(select));
    return;
  }
  const state = getSearchableSelectState(select);
  const placeholderOption = select.querySelector('option[data-placeholder-option="true"]');
  if (!placeholderOption || !state) return;
  if (state.text) {
    placeholderOption.textContent = matches
      ? `Filtrando: "${state.text}" (${matches} coincidencia${matches === 1 ? '' : 's'})`
      : `Sin coincidencias para "${state.text}"`;
    placeholderOption.disabled = true;
  } else {
    const hasItems = select.dataset.searchableHasItems !== 'false';
    const emptyLabel = select.dataset.searchableEmptyLabel;
    placeholderOption.textContent = !hasItems && emptyLabel
      ? emptyLabel
      : state.placeholder;
    placeholderOption.disabled = true;
  }
}

function applySearchableSelectFilter(select, text = '') {
  if (!select) return;
  if (select.tomselect) {
    updateTomSelectPlaceholder(select.tomselect, getSearchablePlaceholder(select));
    return;
  }
  const state = getSearchableSelectState(select);
  if (!state) return;
  const nextText = typeof text === 'string' ? text : '';
  state.text = nextText;
  if (state.filterInput && state.filterInput.value !== nextText) {
    state.filterInput.value = nextText;
  }
  const normalized = nextText.trim().toLowerCase();
  let matches = 0;
  Array.from(select.options).forEach(option => {
    if (option.dataset.placeholderOption === 'true') {
      option.hidden = false;
      option.style.display = '';
      return;
    }
    const label = (option.textContent || '').toLowerCase();
    const value = (option.value || '').toLowerCase();
    const match = !normalized || label.includes(normalized) || value.includes(normalized);
    option.hidden = !match;
    option.style.display = match ? '' : 'none';
    if (match) {
      matches += 1;
    }
  });
  updateSearchableSelectPlaceholder(select, matches);
  const currentValue = select.value;
  if (currentValue) {
    const currentOption = Array.from(select.options).find(option => option.value === currentValue && !option.hidden);
    if (!currentOption) {
      const placeholderOption = select.querySelector('option[data-placeholder-option="true"]');
      if (placeholderOption) {
        placeholderOption.selected = true;
        select.value = placeholderOption.value;
      }
    }
  } else {
    const placeholderOption = select.querySelector('option[data-placeholder-option="true"]');
    if (placeholderOption) {
      placeholderOption.selected = true;
    }
  }
}

function clearSearchableSelectFilter(select) {
  if (!select) return;
  if (select.tomselect) {
    const control = select.tomselect;
    if (typeof control.setTextboxValue === 'function') {
      control.setTextboxValue('');
    }
    updateTomSelectPlaceholder(control, getSearchablePlaceholder(select));
    return;
  }
  const state = getSearchableSelectState(select);
  if (!state) return;
  state.text = '';
  if (state.filterInput) {
    state.filterInput.value = '';
  }
  applySearchableSelectFilter(select, '');
}

function resetSearchableSelect(select) {
  if (!select) return;
  if (select.tomselect) {
    const control = select.tomselect;
    control.clear(true);
    updateTomSelectPlaceholder(control, getSearchablePlaceholder(select));
    if (typeof control.close === 'function') {
      control.close();
    }
    return;
  }
  select.value = '';
  clearSearchableSelectFilter(select);
  const placeholderOption = select.querySelector('option[data-placeholder-option="true"]');
  if (placeholderOption) {
    placeholderOption.selected = true;
  }
}

function setSearchableSelectValue(select, value) {
  if (!select) return;
  if (select.tomselect) {
    const control = select.tomselect;
    if (value) {
      control.setValue(value, true);
    } else {
      control.clear(true);
    }
    return;
  }
  const option = Array.from(select.options).find(item => item.value === value);
  if (option) {
    select.value = value;
    option.selected = true;
  } else {
    select.value = '';
  }
  clearSearchableSelectFilter(select);
}

function handleSearchableSelectKeydown(event) {
  const select = event.target;
  if (!select || select.tagName !== 'SELECT' || !select.dataset.searchableSelect) {
    return;
  }
  if (select.tomselect) {
    return;
  }
  const state = getSearchableSelectState(select);
  if (!state) return;
  const { key } = event;
  if (key === 'Escape') {
    event.preventDefault();
    resetSearchableSelect(select);
    return;
  }
  if (key === 'Backspace') {
    event.preventDefault();
    state.text = state.text.slice(0, -1);
    applySearchableSelectFilter(select, state.text);
    return;
  }
  if (key === 'Delete') {
    event.preventDefault();
    state.text = '';
    applySearchableSelectFilter(select, '');
    return;
  }
  if (key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault();
    const nextText = `${state.text}${key}`;
    applySearchableSelectFilter(select, nextText);
    return;
  }
  if (key === 'Enter') {
    const value = select.value;
    if (!value) {
      return;
    }
    event.preventDefault();
    select.dispatchEvent(new Event('change', { bubbles: true }));
  }
}

function initializeSearchableSelect(select, { placeholder } = {}) {
  if (!select) return;
  const state = getSearchableSelectState(select);
  if (placeholder) {
    state.placeholder = placeholder;
  }
  if (window.TomSelect) {
    if (!select.tomselect) {
      const control = new TomSelect(select, {
        valueField: 'value',
        labelField: 'text',
        searchField: ['text', 'value'],
        allowEmptyOption: true,
        create: false,
        persist: false,
        maxOptions: 500,
        plugins: ['clear_button'],
        placeholder: getSearchablePlaceholder(select)
      });
      updateTomSelectPlaceholder(control, getSearchablePlaceholder(select));
      select.dataset.searchableInitialized = 'tomselect';
    } else if (placeholder) {
      updateTomSelectPlaceholder(select.tomselect, getSearchablePlaceholder(select));
    }
    refreshTomSelectOptions(select);
    return;
  }
  if (select.dataset.searchableInitialized === 'true') return;
  ensureSearchableFilterInput(select);
  const hasInitialOptions = Array.from(select.options).some(option => option.value && !option.disabled);
  select.dataset.searchableHasItems = hasInitialOptions ? 'true' : 'false';
  if (!hasInitialOptions) {
    const emptyPlaceholder = select.getAttribute('data-empty-placeholder');
    if (emptyPlaceholder) {
      select.dataset.searchableEmptyLabel = emptyPlaceholder;
    }
  }
  select.dataset.searchableInitialized = 'true';
  select.addEventListener('keydown', handleSearchableSelectKeydown);
  select.addEventListener('focus', () => {
    applySearchableSelectFilter(select, state.text);
  });
  select.addEventListener('blur', () => {
    if (!state.text) {
      applySearchableSelectFilter(select, '');
    }
  });
}

function formatPercentageValue(value, decimals = 2) {
  return roundTo(value, decimals).toFixed(decimals);
}

function formatPercentageLabel(value, decimals = 2) {
  return `${formatPercentageValue(value, decimals)}%`;
}

function computePercentagesFromTotals(totals = {}) {
  const total = normalizeNumber(totals.total);
  if (total <= 0) {
    return { rem: 0, fac: 0, rest: 0 };
  }
  const totalRem = normalizeNumber(totals.totalRem);
  const totalFac = normalizeNumber(totals.totalFac);
  const consumo = totals.totalConsumo != null
    ? normalizeNumber(totals.totalConsumo)
    : roundTo(totalRem + totalFac);
  const restante = totals.restante != null
    ? Math.max(normalizeNumber(totals.restante), 0)
    : Math.max(roundTo(total - consumo), 0);
  const rem = roundTo((totalRem / total) * 100);
  const fac = roundTo((totalFac / total) * 100);
  const rest = roundTo((restante / total) * 100);
  return { rem, fac, rest };
}

function computeGlobalPercentages(totals = {}, baseAmount = 0) {
  const base = normalizeNumber(baseAmount);
  const fallbackTotal = normalizeNumber(totals.total);
  const scale = base > 0 ? base : fallbackTotal > 0 ? fallbackTotal : 1;
  const totalRem = Math.max(0, normalizeNumber(totals.totalRem));
  const totalFac = Math.max(0, normalizeNumber(totals.totalFac));
  const consumo = totals.totalConsumo != null
    ? Math.max(0, normalizeNumber(totals.totalConsumo))
    : roundTo(totalRem + totalFac);
  const restanteCalculado = totals.restante != null
    ? normalizeNumber(totals.restante)
    : roundTo((totals.total || 0) - consumo);
  const restante = Math.max(restanteCalculado, 0);
  const rem = roundTo((totalRem / scale) * 100);
  const fac = roundTo((totalFac / scale) * 100);
  const consumoPerc = roundTo((consumo / scale) * 100);
  const rest = roundTo((restante / scale) * 100);
  return { rem, fac, consumo: consumoPerc, rest };
}

function getTotalsWithDefaults(totals) {
  return {
    total: 0,
    totalRem: 0,
    totalFac: 0,
    totalConsumo: 0,
    restante: 0,
    porcRem: 0,
    porcFac: 0,
    porcRest: 0,
    ...(totals || {})
  };
}

function normalizeDocumentId(value) {
  return typeof value === 'string' ? value.trim().toUpperCase() : '';
}

function buildPoGroups(items = [], selectionDetails = []) {
  const groups = new Map();
  const manualMap = new Map();
  const normalizePoKey = value => (typeof value === 'string' ? value.trim().toUpperCase() : '');

  if (Array.isArray(selectionDetails)) {
    selectionDetails.forEach(detail => {
      const baseValue = typeof detail?.baseId === 'string' ? detail.baseId.trim() : '';
      const normalizedBase = normalizePoKey(baseValue);
      if (!normalizedBase) return;
      manualMap.set(normalizedBase, baseValue);
      const variants = new Set();
      variants.add(baseValue);
      (Array.isArray(detail?.variants) ? detail.variants : []).forEach(id => {
        if (typeof id === 'string' && id.trim()) {
          variants.add(id.trim());
        }
      });
      variants.forEach(id => {
        const normalizedVariant = normalizePoKey(id);
        if (!normalizedVariant) return;
        manualMap.set(normalizedVariant, baseValue);
      });
    });
  }

  items.forEach(item => {
    const itemId = typeof item?.id === 'string' ? item.id.trim() : '';
    if (!itemId) {
      return;
    }

    const manualBase = manualMap.get(normalizePoKey(itemId));
    const baseCandidate = manualBase || item?.baseId || itemId;
    const key = typeof baseCandidate === 'string' ? baseCandidate.trim() : '';
    if (!key) {
      return;
    }

    if (!groups.has(key)) {
      groups.set(key, {
        baseId: key,
        ids: new Set([key]),
        total: 0,
        remisiones: new Map(),
        facturas: new Map(),
        remFallback: 0,
        facFallback: 0,
        items: new Map()
      });
    }

    const group = groups.get(key);
    group.ids.add(itemId);

    const totals = getTotalsWithDefaults(item?.totals);
    const authorized = normalizeNumber(item?.total ?? totals.total);
    const baseTotal = authorized > 0 ? authorized : normalizeNumber(totals.total);
    const totalRem = normalizeNumber(totals.totalRem);
    const totalFac = normalizeNumber(totals.totalFac);
    const restante = totals.restante != null
      ? normalizeNumber(totals.restante)
      : Math.max(baseTotal - (totalRem + totalFac), 0);
    const normalizedTotals = {
      total: roundTo(baseTotal),
      totalRem: roundTo(totalRem),
      totalFac: roundTo(totalFac),
      totalConsumo: roundTo(totalRem + totalFac),
      restante: roundTo(restante)
    };
    const percentages = computePercentagesFromTotals(normalizedTotals);

    group.total += baseTotal;

    const remisiones = Array.isArray(item?.remisiones) ? item.remisiones : [];
    if (remisiones.length) {
      remisiones.forEach(rem => {
        const docKey = normalizeDocumentId(rem?.id);
        if (!docKey) return;
        const amount = normalizeNumber(rem?.monto);
        if (!group.remisiones.has(docKey)) {
          group.remisiones.set(docKey, {
            id: typeof rem?.id === 'string' ? rem.id.trim() : docKey,
            monto: amount
          });
        } else {
          const entry = group.remisiones.get(docKey);
          entry.monto = Math.max(entry.monto, amount);
          if (!entry.id && rem?.id) {
            entry.id = rem.id.trim();
          }
        }
      });
    } else if (totalRem > 0) {
      group.remFallback += totalRem;
    }

    const facturas = Array.isArray(item?.facturas) ? item.facturas : [];
    if (facturas.length) {
      facturas.forEach(fac => {
        const docKey = normalizeDocumentId(fac?.id);
        if (!docKey) return;
        const amount = normalizeNumber(fac?.monto);
        if (!group.facturas.has(docKey)) {
          group.facturas.set(docKey, {
            id: typeof fac?.id === 'string' ? fac.id.trim() : docKey,
            monto: amount
          });
        } else {
          const entry = group.facturas.get(docKey);
          entry.monto = Math.max(entry.monto, amount);
          if (!entry.id && fac?.id) {
            entry.id = fac.id.trim();
          }
        }
      });
    } else if (totalFac > 0) {
      group.facFallback += totalFac;
    }

    const isBaseVariant = normalizePoKey(itemId) === normalizePoKey(group.baseId);
    group.items.set(itemId, {
      id: itemId,
      label: itemId,
      totals: normalizedTotals,
      percentages,
      isBase: isBaseVariant,
      fecha: item?.fecha || ''
    });
  });

  return Array.from(groups.values())
    .map(group => {
      const variantIds = Array.from(group.ids);
      const orderedVariants = variantIds.includes(group.baseId)
        ? [group.baseId, ...variantIds.filter(id => id !== group.baseId).sort((a, b) => a.localeCompare(b))]
        : variantIds.sort((a, b) => a.localeCompare(b));
      const extensionIds = orderedVariants.filter(id => id !== group.baseId);
      const totalRem = Array.from(group.remisiones.values()).reduce((sum, entry) => sum + entry.monto, 0) + group.remFallback;
      const totalFac = Array.from(group.facturas.values()).reduce((sum, entry) => sum + entry.monto, 0) + group.facFallback;
      const totalConsumo = totalRem + totalFac;
      const restante = Math.max(group.total - totalConsumo, 0);
      const totals = {
        total: roundTo(group.total),
        totalRem: roundTo(totalRem),
        totalFac: roundTo(totalFac),
        totalConsumo: roundTo(totalConsumo),
        restante: roundTo(restante)
      };
      const baseKey = normalizePoKey(group.baseId);
      const items = Array.from(group.items.values())
        .sort((a, b) => {
          const aKey = normalizePoKey(a.id);
          const bKey = normalizePoKey(b.id);
          if (aKey === baseKey) return -1;
          if (bKey === baseKey) return 1;
          return aKey.localeCompare(bKey);
        });

      return {
        baseId: group.baseId,
        ids: orderedVariants,
        extensionIds,
        totals,
        percentages: computePercentagesFromTotals(totals),
        items
      };
    })
    .sort((a, b) => a.baseId.localeCompare(b.baseId));
}

function buildGroupBadgeList(baseId, extensionIds = []) {
  const normalizedBase = typeof baseId === 'string' ? baseId.trim() : '';
  const normalizedExtensions = Array.from(
    new Set(
      (Array.isArray(extensionIds) ? extensionIds : [])
        .map(id => (typeof id === 'string' ? id.trim() : ''))
        .filter(Boolean)
    )
  );

  const badges = [];
  if (normalizedBase) {
    badges.push(
      `<span class="badge extension-badge-base">Base ${escapeHtml(normalizedBase)}</span>`
    );
  }
  normalizedExtensions.forEach(id => {
    badges.push(
      `<span class="badge extension-badge-extension">Ext ${escapeHtml(id)}</span>`
    );
  });
  return badges.join('');
}

function buildGroupCompositionNote(extensionIds = []) {
  const hasExtensions = Array.isArray(extensionIds) && extensionIds.length > 0;
  return hasExtensions
    ? 'Gráficas y montos consolidados del subgrupo base + extensiones seleccionadas.'
    : 'Gráficas y montos calculados solo con la PO base.';
}

const AXIS_CURRENCY_FORMATTER = new Intl.NumberFormat('es-MX', {
  style: 'currency',
  currency: 'MXN',
  maximumFractionDigits: 0
});

const CHART_COLORS = {
  rem: '#2563eb',
  fac: '#f97316',
  rest: '#16a34a'
};

const FORMAT_LABEL_MAP = {
  pdf: 'PDF (predeterminado)',
  csv: 'CSV (hoja de cálculo)',
  json: 'JSON (datos crudos)'
};

const FORMAT_DESCRIPTION_MAP = {
  csv: 'Ideal para abrir en Excel u hojas de cálculo.',
  json: 'Incluye toda la información del resumen en formato estructurado.'
};

function getFormatLabel(format) {
  if (!format) return 'Formato';
  return FORMAT_LABEL_MAP[format] || format.toUpperCase();
}

function getFormatDescription(format) {
  return FORMAT_DESCRIPTION_MAP[format] || '';
}

async function login(event) {
  event?.preventDefault();
  const form = document.getElementById('loginForm');
  const usernameInput = document.getElementById('username');
  const passwordInput = document.getElementById('password');
  const username = usernameInput?.value.trim();
  const password = passwordInput?.value || '';

  let isValid = true;
  if (!username) {
    usernameInput?.classList.add('is-invalid');
    isValid = false;
  } else {
    usernameInput?.classList.remove('is-invalid');
  }

  if (!password) {
    passwordInput?.classList.add('is-invalid');
    isValid = false;
  } else {
    passwordInput?.classList.remove('is-invalid');
  }

  if (!isValid) {
    form?.classList.add('was-validated');
    return;
  }

  try {
    const response = await fetch('/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const data = await response.json();
    if (!data.success) {
      showAlert(data.message || 'Credenciales inválidas', 'danger');
      return;
    }
    saveSession('porpt-empresas', data.empresas || []);
    saveSession('porpt-is-admin', !!data.isAdmin);
    showAlert('¡Bienvenido! Selecciona empresa y PO para ver el dashboard.', 'success');
    window.location.href = 'dashboard.html';
  } catch (error) {
    console.error('Error en login:', error);
    showAlert('No fue posible iniciar sesión. Verifica la conexión.', 'danger');
  }
}

function renderEmpresaOptions() {
  const select = document.getElementById('empresaSearch');
  if (!select) return;
  const selectState = getSearchableSelectState(select);
  const previousValue = state.selectedEmpresa || select.value || '';
  select.innerHTML = '';
  select.dataset.searchableHasItems = state.empresas.length ? 'true' : 'false';
  delete select.dataset.searchableEmptyLabel;
  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.dataset.placeholderOption = 'true';
  placeholderOption.textContent = selectState?.placeholder || 'Selecciona una empresa';
  placeholderOption.disabled = true;
  select.appendChild(placeholderOption);
  state.empresas.forEach(empresa => {
    const option = document.createElement('option');
    option.value = empresa;
    option.textContent = empresa;
    select.appendChild(option);
  });
  if (previousValue && state.empresas.includes(previousValue)) {
    select.value = previousValue;
  } else {
    select.value = '';
  }
  applySearchableSelectFilter(select, selectState?.text || '');
  refreshTomSelectOptions(select);
  renderOverviewEmpresaOptions();
}

function renderUniverseEmpresaSelect() {
  const select = document.getElementById('universeEmpresaSelect');
  if (!select) return;
  const currentValue = state.selectedEmpresa || '';
  select.innerHTML = '<option value="">Selecciona una empresa</option>';
  state.empresas.forEach(empresa => {
    const option = document.createElement('option');
    option.value = empresa;
    option.textContent = empresa;
    select.appendChild(option);
  });
  select.value = currentValue;
}

function renderPoOptions() {
  const select = document.getElementById('poSearch');
  if (!select) return;
  const selectState = getSearchableSelectState(select);
  const previousValue = select.value || '';
  select.innerHTML = '';
  const emptyPlaceholder = select.getAttribute('data-empty-placeholder');
  const selectedIds = getAllSelectedPoIds();
  const availableOptions = state.pos.filter(po => !selectedIds.has(po.id));
  const hasItems = availableOptions.length > 0;
  select.dataset.searchableHasItems = hasItems ? 'true' : 'false';
  if (!hasItems && emptyPlaceholder) {
    select.dataset.searchableEmptyLabel = emptyPlaceholder;
  } else {
    delete select.dataset.searchableEmptyLabel;
  }
  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.dataset.placeholderOption = 'true';
  placeholderOption.textContent = hasItems
    ? (selectState?.placeholder || 'Selecciona una PO base o extensión')
    : (emptyPlaceholder || selectState?.placeholder || 'Selecciona una empresa primero');
  placeholderOption.disabled = true;
  select.appendChild(placeholderOption);
  availableOptions.forEach(po => {
    const option = document.createElement('option');
    option.value = po.id;
    option.textContent = po.display;
    select.appendChild(option);
  });
  if (previousValue && Array.from(select.options).some(option => option.value === previousValue)) {
    select.value = previousValue;
  } else {
    select.value = '';
  }
  applySearchableSelectFilter(select, selectState?.text || '');
  updatePoFoundList();
  refreshTomSelectOptions(select);
}

function updatePoFoundList() {
  const container = document.getElementById('poFoundList');
  if (!container) return;
  if (!state.selectedEmpresa) {
    container.textContent = '';
    container.title = '';
    return;
  }
  const total = Array.isArray(state.pos) ? state.pos.length : 0;
  if (!total) {
    container.textContent = 'Sin POs disponibles';
    container.title = '';
    return;
  }
  const summary = `${total} PO${total !== 1 ? 's' : ''} encontradas`;
  container.textContent = summary;
  container.title = summary;
}

async function ensureEmpresasCatalog() {
  if (Array.isArray(state.empresas) && state.empresas.length > 0) {
    return state.empresas;
  }
  const stored = readSession('porpt-empresas', []);
  if (Array.isArray(stored) && stored.length > 0) {
    state.empresas = stored;
    return state.empresas;
  }
  try {
    const response = await fetch('/empresas');
    const data = await response.json();
    if (data.success) {
      state.empresas = data.empresas || [];
      saveSession('porpt-empresas', state.empresas);
    }
  } catch (error) {
    console.error('Error obteniendo empresas:', error);
  }
  return state.empresas;
}

function getPoMetadataByBase(baseId) {
  if (!baseId) return null;
  return state.pos.find(po => po.id === baseId) || state.pos.find(po => po.baseId === baseId) || null;
}

function sanitizeVariantKey(value) {
  return typeof value === 'string' ? value.replace(/[^0-9a-zA-Z-_]+/gu, '_') : '';
}

function ensureSelectedPoDetail(baseId, initialId) {
  if (!baseId) return null;
  const normalizedBase = baseId.trim();
  if (!state.selectedPoDetails.has(normalizedBase)) {
    const variants = new Set();
    if (normalizedBase) {
      variants.add(normalizedBase);
    }
    if (initialId && initialId !== normalizedBase) {
      variants.add(initialId.trim());
    }
    state.selectedPoDetails.set(normalizedBase, { baseId: normalizedBase, variants });
  } else if (initialId && initialId !== normalizedBase) {
    const detail = state.selectedPoDetails.get(normalizedBase);
    detail.variants.add(initialId.trim());
  }
  return state.selectedPoDetails.get(normalizedBase) || null;
}

function removeSelectedPoDetail(baseId) {
  if (!baseId) return;
  state.selectedPoDetails.delete(baseId.trim());
}

function getSelectedVariantIds(baseId) {
  const detail = baseId ? state.selectedPoDetails.get(baseId.trim()) : null;
  if (detail && detail.variants.size > 0) {
    return Array.from(detail.variants);
  }
  return baseId ? [baseId.trim()] : [];
}

function getSuggestedVariants(baseId) {
  if (!baseId) return [];
  const normalized = baseId.trim();
  const variants = state.pos.filter(po => (po.baseId || po.id) === normalized);
  if (variants.length > 0) {
    return variants.sort((a, b) => a.id.localeCompare(b.id));
  }
  const fallback = {
    id: normalized,
    baseId: normalized,
    display: normalized,
    fecha: '',
    total: 0,
    isExtension: false,
    isFallback: true
  };
  return [fallback];
}

function getManualVariants(baseId) {
  if (!baseId) return [];
  const store = state.manualExtensions.get(baseId.trim());
  if (!store) return [];
  return Array.from(store.values());
}

function getAllSelectedPoIds() {
  const selected = new Set();
  state.selectedPoIds.forEach(baseId => {
    if (!baseId) return;
    const normalized = baseId.trim();
    if (!normalized) return;
    selected.add(normalized);
    const detail = state.selectedPoDetails.get(normalized);
    if (detail?.variants) {
      detail.variants.forEach(id => {
        if (typeof id === 'string' && id.trim()) {
          selected.add(id.trim());
        }
      });
    }
  });
  return selected;
}

function getAvailableManualExtensionOptions(baseId) {
  if (!baseId) return [];
  const normalizedBase = baseId.trim();
  if (!normalizedBase) return [];
  const selectedVariants = new Set(getSelectedVariantIds(normalizedBase));
  selectedVariants.add(normalizedBase);
  const manualVariants = getManualVariants(normalizedBase);
  manualVariants.forEach(variant => {
    if (variant?.id) {
      selectedVariants.add(variant.id);
    }
  });
  const globalSelected = getAllSelectedPoIds();

  const candidateMap = new Map();
  state.pos.forEach(po => {
    const id = po?.id;
    if (!id || !id.trim()) return;
    if (id === normalizedBase) return;
    if (selectedVariants.has(id)) return;
    if (globalSelected.has(id)) return;
    const matchesBase = po.baseId
      ? po.baseId === normalizedBase
      : id.startsWith(normalizedBase);
    if (matchesBase) {
      candidateMap.set(id, po);
    }
  });

  if (!candidateMap.size) {
    state.pos.forEach(po => {
      const id = po?.id;
      if (!id || !id.trim()) return;
      if (id === normalizedBase) return;
      if (selectedVariants.has(id)) return;
      if (globalSelected.has(id)) return;
      if (!candidateMap.has(id)) {
        candidateMap.set(id, po);
      }
    });
  }

  return Array.from(candidateMap.values()).sort((a, b) => a.id.localeCompare(b.id));
}

function renderManualExtensionSelectOptions(select, baseId) {
  if (!select) return;
  const selectState = getSearchableSelectState(select);
  const previousValue = select.value || '';
  const options = getAvailableManualExtensionOptions(baseId);
  select.innerHTML = '';
  if (!options.length) {
    select.dataset.searchableHasItems = 'false';
    select.dataset.searchableEmptyLabel = 'Sin extensiones disponibles';
  } else {
    select.dataset.searchableHasItems = 'true';
    delete select.dataset.searchableEmptyLabel;
  }
  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.dataset.placeholderOption = 'true';
  placeholderOption.textContent = selectState?.placeholder || 'Selecciona una extensión disponible';
  placeholderOption.disabled = true;
  select.appendChild(placeholderOption);
  options.forEach(po => {
    const option = document.createElement('option');
    option.value = po.id;
    option.textContent = po.display || po.id;
    select.appendChild(option);
  });
  if (previousValue && options.some(option => option.id === previousValue)) {
    select.value = previousValue;
  } else {
    select.value = '';
  }
  applySearchableSelectFilter(select, selectState?.text || '');
  refreshTomSelectOptions(select);
}

function storeManualVariant(baseId, variant) {
  if (!baseId || !variant?.id) return;
  const normalizedBase = baseId.trim();
  const normalizedId = variant.id.trim();
  if (!normalizedBase || !normalizedId) return;
  if (!state.manualExtensions.has(normalizedBase)) {
    state.manualExtensions.set(normalizedBase, new Map());
  }
  const baseStore = state.manualExtensions.get(normalizedBase);
  baseStore.set(normalizedId, {
    id: normalizedId,
    baseId: normalizedBase,
    display: variant.display || normalizedId,
    fecha: variant.fecha || '',
    total: normalizeNumber(variant.total) || 0,
    isExtension: false,
    isManual: true
  });
}

function addManualExtensionToPo(baseId, rawId) {
  if (!baseId) return;
  const manualId = typeof rawId === 'string' ? rawId.trim() : '';
  if (!manualId) {
    showAlert('Selecciona la extensión que deseas agregar.', 'warning');
    return;
  }
  const normalizedBase = baseId.trim();
  const normalizedId = manualId;
  if (normalizedId === normalizedBase) {
    showAlert('La extensión coincide con la PO base; ya está incluida automáticamente.', 'info');
    return;
  }

  const detail = ensureSelectedPoDetail(normalizedBase);
  if (!detail) {
    showAlert('Selecciona primero la PO base antes de agregar una extensión manual.', 'warning');
    return;
  }
  if (detail.variants.has(normalizedId)) {
    showAlert(`La extensión ${normalizedId} ya está seleccionada para la PO ${normalizedBase}.`, 'info');
    return;
  }

  detail.variants.add(normalizedId);
  detail.variants.add(normalizedBase);

  const existsInCatalog = state.pos.some(po => po.id === normalizedId && (po.baseId || po.id) === normalizedBase);
  if (!existsInCatalog) {
    storeManualVariant(normalizedBase, { id: normalizedId });
  }

  renderSelectedPoChips();
  updateSummary({ silent: true });
  showAlert(`Extensión ${normalizedId} agregada manualmente a la PO ${normalizedBase}.`, 'success');
}

function renderExtensionSelection() {
  const container = document.getElementById('extensionSelectionContainer');
  if (!container) return;
  container.innerHTML = '';

  state.selectedPoIds.forEach(baseId => {
    const suggestedVariants = getSuggestedVariants(baseId);
    const manualVariants = getManualVariants(baseId);
    const variantMap = new Map();
    suggestedVariants.forEach(variant => {
      variantMap.set(variant.id, variant);
    });
    manualVariants.forEach(variant => {
      variantMap.set(variant.id, variant);
    });
    const detail = ensureSelectedPoDetail(baseId);
    const selectedVariants = detail ? Array.from(detail.variants) : [baseId];
    selectedVariants.forEach(id => {
      if (!variantMap.has(id)) {
        variantMap.set(id, {
          id,
          baseId,
          display: id,
          fecha: '',
          total: 0,
          isExtension: false,
          isManual: true
        });
      }
    });
    const variants = Array.from(variantMap.values()).sort((a, b) => a.id.localeCompare(b.id));
    const totalVariants = Math.max(variants.length, 1);
    const selectedCount = Math.min(selectedVariants.length, totalVariants);
    const badgeLabel = `${selectedCount}/${totalVariants} seleccionadas`;
    const poMeta = getPoMetadataByBase(baseId);

    const card = document.createElement('section');
    card.className = 'extension-selection-card';
    card.dataset.baseId = baseId;

    const header = document.createElement('div');
    header.className = 'extension-selection-header';
    const authorizedText = poMeta
      ? `Autorizado: $${formatCurrency(poMeta.total || 0)}`
      : 'Marca solo las extensiones que suman consumo.';
    header.innerHTML = `
      <div>
        <h5>PO ${escapeHtml(baseId)}</h5>
        <p>${escapeHtml(authorizedText)}</p>
      </div>
      <span class="extension-selection-badge">${escapeHtml(badgeLabel)}</span>
    `;
    card.appendChild(header);

    const actions = document.createElement('div');
    actions.className = 'extension-selection-actions';
    actions.innerHTML = `
      <button type="button" class="btn btn-outline-primary" data-action="select-all" data-base-id="${escapeHtml(baseId)}">Seleccionar todas</button>
      <button type="button" class="btn btn-outline-secondary" data-action="select-base" data-base-id="${escapeHtml(baseId)}">Solo base</button>
    `;
    card.appendChild(actions);

    const optionsGrid = document.createElement('div');
    optionsGrid.className = 'extension-selection-options';

    variants.forEach(variant => {
      const option = document.createElement('div');
      option.className = 'extension-option form-check form-switch';
      const input = document.createElement('input');
      input.className = 'form-check-input';
      input.type = 'checkbox';
      input.value = variant.id;
      input.dataset.baseId = baseId;
      input.id = `variant-${sanitizeVariantKey(baseId)}-${sanitizeVariantKey(variant.id)}`;
      if (variant.id === baseId) {
        input.checked = true;
        input.disabled = true;
      } else {
        input.checked = selectedVariants.includes(variant.id);
      }

      const label = document.createElement('label');
      label.className = 'form-check-label extension-option-label';
      label.setAttribute('for', input.id);
      const title = escapeHtml(variant.display || variant.id);
      const metaParts = [];
      if (variant.fecha) {
        metaParts.push(`Fecha: ${escapeHtml(variant.fecha)}`);
      }
      const totalValue = normalizeNumber(variant.total);
      if (totalValue > 0) {
        metaParts.push(`Autorizado: $${formatCurrency(totalValue)}`);
      }
      const badges = [];
      if (variant.id === baseId) {
        badges.push('<span class="badge text-bg-secondary">Base</span>');
      } else if (variant.isExtension === false) {
        badges.push('<span class="badge text-bg-warning-subtle text-warning-emphasis">Manual</span>');
      } else {
        badges.push('<span class="badge text-bg-info-subtle text-info-emphasis">Extensión</span>');
      }
      label.innerHTML = `
        <span class="extension-option-title">${title}</span>
        ${metaParts.length ? `<span class="extension-option-meta">${metaParts.join(' · ')}</span>` : ''}
        ${badges.length ? `<div class="extension-option-badges">${badges.join('')}</div>` : ''}
      `;

      option.appendChild(input);
      option.appendChild(label);
      optionsGrid.appendChild(option);
    });

    if (variants.length <= 1) {
      const empty = document.createElement('p');
      empty.className = 'extension-selection-empty';
      optionsGrid.appendChild(empty);
    }

    card.appendChild(optionsGrid);
    const manualEntry = document.createElement('div');
    manualEntry.className = 'extension-manual-entry';
    manualEntry.innerHTML = `
      <label class="form-label">Agregar extensión manual</label>
      <div class="input-group">
        <span class="input-group-text">Extensión</span>
        <select
          class="form-select"
          data-searchable-select
          data-placeholder="Selecciona una extensión disponible"
          data-empty-placeholder="Sin extensiones disponibles"
          data-manual-extension-select
          data-base-id="${escapeHtml(baseId)}"
        ></select>
        <button type="button" class="btn btn-outline-success" data-action="add-manual-extension" data-base-id="${escapeHtml(baseId)}">Agregar</button>
      </div>
    `;
    card.appendChild(manualEntry);
    const manualSelect = manualEntry.querySelector('select[data-manual-extension-select]');
    renderManualExtensionSelectOptions(manualSelect, baseId);
    initializeSearchableSelect(manualSelect, { placeholder: 'Selecciona una extensión disponible' });
    container.appendChild(card);
  });
}

function handleExtensionSelectionChange(event) {
  const input = event.target;
  if (!input || input.type !== 'checkbox') return;
  const baseId = input.dataset.baseId;
  const variantId = input.value;
  if (!baseId || !variantId) return;
  const detail = ensureSelectedPoDetail(baseId, variantId);
  if (!detail) return;
  if (variantId === baseId) {
    input.checked = true;
    return;
  }
  if (input.checked) {
    detail.variants.add(variantId);
  } else {
    detail.variants.delete(variantId);
    if (detail.variants.size === 0) {
      detail.variants.add(baseId);
    }
  }
  renderSelectedPoChips();
  updateSummary({ silent: true });
}

function handleExtensionSelectionAction(event) {
  const button = event.target.closest('button[data-action][data-base-id]');
  if (!button) return;
  const baseId = button.dataset.baseId;
  const action = button.dataset.action;
  if (!baseId || !action) return;
  const card = button.closest('.extension-selection-card');
  if (!card) return;
  if (action === 'add-manual-extension') {
    const select = card.querySelector(`select[data-manual-extension-select][data-base-id="${baseId}"]`);
    const value = select ? select.value : '';
    addManualExtensionToPo(baseId, value);
    if (select) {
      resetSearchableSelect(select);
    }
    return;
  }
  const detail = ensureSelectedPoDetail(baseId);
  if (!detail) return;
  const checkboxes = Array.from(
    card?.querySelectorAll('input[type="checkbox"][data-base-id]') || []
  );
  if (!checkboxes.length) return;

  const mutableSet = detail.variants instanceof Set ? detail.variants : new Set(detail.variants || []);

  if (action === 'select-all') {
    checkboxes.forEach(checkbox => {
      checkbox.checked = true;
      if (!checkbox.disabled) {
        mutableSet.add(checkbox.value);
      }
    });
  } else if (action === 'select-base') {
    mutableSet.clear();
    mutableSet.add(baseId);
    checkboxes.forEach(checkbox => {
      checkbox.checked = checkbox.value === baseId;
    });
  } else {
    return;
  }

  mutableSet.add(baseId);
  detail.variants = mutableSet;
  renderSelectedPoChips();
  updateSummary({ silent: true });
}

function handleManualExtensionSelectChange(event) {
  const select = event.target;
  if (!select || !select.matches('select[data-manual-extension-select]')) return;
  const baseId = select.dataset.baseId;
  if (!baseId) return;
  const value = select.value;
  if (!value) return;
  addManualExtensionToPo(baseId, value);
  resetSearchableSelect(select);
}

function syncSelectedVariantsFromSummary(summary) {
  if (!summary || !Array.isArray(summary.selectionDetails)) {
    return false;
  }
  let changed = false;
  summary.selectionDetails.forEach(detail => {
    if (!detail?.baseId) return;
    const variants = Array.isArray(detail.variants) && detail.variants.length
      ? detail.variants
      : [detail.baseId];
    const normalized = variants.map(id => (typeof id === 'string' ? id.trim() : '')).filter(Boolean);
    const stateDetail = ensureSelectedPoDetail(detail.baseId);
    if (!stateDetail) {
      return;
    }
    const current = stateDetail.variants instanceof Set
      ? stateDetail.variants
      : new Set(stateDetail.variants || []);
    let updated = false;
    normalized.forEach(id => {
      if (!current.has(id)) {
        current.add(id);
        updated = true;
      }
    });
    if (!current.has(detail.baseId)) {
      current.add(detail.baseId);
      updated = true;
    }
    if (!(stateDetail.variants instanceof Set)) {
      updated = true;
    }
    if (updated) {
      stateDetail.variants = current;
      changed = true;
    }
  });
  return changed;
}

function renderSelectedPoChips() {
  const container = document.getElementById('selectedPoContainer');
  if (!container) return;
  container.innerHTML = '';
  const help = document.getElementById('selectedPoHelp');
  if (!state.selectedPoIds.length) {
    const empty = document.createElement('p');
    empty.className = 'text-muted small mb-0';
    empty.textContent = 'No hay POs seleccionadas.';
    container.appendChild(empty);
    hideDashboardPanel();
    renderReportSelectionOverview();
    if (help) {
      help.textContent = 'Agrega las extensiones de un PO manualmente o elige la sugerencia';
    }
    renderExtensionSelection();
    return;
  }
  state.selectedPoIds.forEach(baseId => {
    const chip = document.createElement('span');
    chip.className = 'selected-po-chip';
    const meta = getPoMetadataByBase(baseId);
    const variantCount = getSelectedVariantIds(baseId).length;
    const extensionCount = Math.max(variantCount - 1, 0);
    const suffix = extensionCount > 0 ? ` (+${extensionCount} ext.)` : '';
    const label = meta
      ? `${baseId}${suffix} • $${formatCurrency(meta.total || 0)}`
      : `${baseId}${suffix}`;
    chip.innerHTML = `
      <span>${label}</span>
      <button type="button" aria-label="Quitar" data-action="remove-po" data-po="${baseId}">&times;</button>
    `;
    container.appendChild(chip);
  });
  renderReportSelectionOverview();
  renderExtensionSelection();
}

function addPoToSelection(poId) {
  if (!poId) return;
  const found = state.pos.find(po => po.id === poId);
  if (!found) {
    showAlert('Selecciona una PO válida de la lista desplegable.', 'warning');
    return;
  }
  const baseId = found.baseId || found.id;
  if (state.selectedPoIds.includes(baseId)) {
    const variants = getSelectedVariantIds(baseId);
    const alreadyIncluded = variants.includes(found.id);
    ensureSelectedPoDetail(baseId, found.id);
    if (found.id !== baseId && !alreadyIncluded) {
      showAlert(`Extensión ${found.id} agregada a la PO ${baseId}.`, 'success');
      renderSelectedPoChips();
      updateSummary({ silent: true });
    } else {
      showAlert(`La PO ${baseId} ya está en la selección.`, 'info');
    }
    return;
  }
  ensureSelectedPoDetail(baseId, found.id);
  state.selectedPoIds.push(baseId);
  showAlert(`PO ${baseId} agregada a la selección.`, 'success');
  renderSelectedPoChips();
  updateSummary();
}

function removePoFromSelection(baseId) {
  const index = state.selectedPoIds.indexOf(baseId);
  if (index === -1) return;
  state.selectedPoIds.splice(index, 1);
  removeSelectedPoDetail(baseId);
  state.manualExtensions.delete(baseId);
  renderSelectedPoChips();
  if (state.selectedPoIds.length === 0) {
    state.summary = null;
    destroyCharts();
    document.getElementById('summaryCards')?.replaceChildren();
    document.getElementById('poTable')?.replaceChildren();
    document.getElementById('extensionsContainer')?.replaceChildren();
    hideDashboardPanel();
  } else {
    updateSummary();
  }
}

function clearSelectedPos() {
  if (!state.selectedPoIds.length) return;
  state.selectedPoIds = [];
  state.selectedPoDetails.clear();
  state.manualExtensions.clear();
  renderSelectedPoChips();
  state.summary = null;
  destroyCharts();
  document.getElementById('summaryCards')?.replaceChildren();
  document.getElementById('poTable')?.replaceChildren();
  document.getElementById('extensionsContainer')?.replaceChildren();
  hideDashboardPanel();
  resetSearchableSelect(document.getElementById('poSearch'));
  renderReportSelectionOverview();
  showAlert('Selección de POs limpiada.', 'info');
}

async function selectEmpresa(value) {
  if (!value || !state.empresas.includes(value)) {
    showAlert('Selecciona una empresa válida de la lista.', 'warning');
    return;
  }
  if (state.selectedEmpresa === value) return;
  state.selectedEmpresa = value;
  state.selectedPoIds = [];
  state.selectedPoDetails.clear();
  state.manualExtensions.clear();
  const poSelect = document.getElementById('poSearch');
  resetSearchableSelect(poSelect);
  setSearchableSelectValue(document.getElementById('empresaSearch'), value);
  setOverviewEmpresaSelectValue(value);
  state.overview.selection = new Map();
  state.overview.itemIndex = new Map();
  state.overview.filters.search = '';
  state.overview.filters.alertsOnly = false;
  updateOverviewSearchInput('');
  const alertsToggle = document.getElementById('overviewAlertsOnly');
  if (alertsToggle) {
    alertsToggle.checked = false;
  }
  renderUniverseEmpresaSelect();
  hideDashboardPanel();
  destroyCharts();
  renderSelectedPoChips();
  updateUniverseControls();
  await loadPOs();
  await loadPoOverview();
}

async function loadPOs() {
  if (!state.selectedEmpresa) return;
  state.pos = [];
  state.summary = null;
  renderPoOptions();
  try {
    const response = await fetch(`/pos/${state.selectedEmpresa}`);
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener las POs');
    }
    state.pos = (data.pos || []).map(po => ({
      ...po,
      display: `${po.id} • ${po.fecha || '-'} • $${formatCurrency(po.total || 0)}${po.isExtension ? ' (Extensión)' : ''}`
    }));
    if (state.pos.length === 0) {
      showAlert('La empresa seleccionada no tiene POs activas.', 'warning');
    }
    renderPoOptions();
    showAlert('POs disponibles listas. Selecciona una o varias para visualizar el consumo.', 'success');
  } catch (error) {
    console.error('Error cargando POs:', error);
    showAlert(error.message || 'Error consultando POs', 'danger');
    renderPoOptions();
  }
}

function renderOverviewEmpresaOptions() {
  const select = document.getElementById('overviewEmpresaSelect');
  if (!select) return;
  const currentValue = state.selectedEmpresa || select.value || '';
  select.innerHTML = '<option value="">Selecciona una empresa</option>';
  state.empresas.forEach(empresa => {
    const option = document.createElement('option');
    option.value = empresa;
    option.textContent = empresa;
    select.appendChild(option);
  });
  if (currentValue && state.empresas.includes(currentValue)) {
    select.value = currentValue;
  }
}

function setOverviewEmpresaSelectValue(value) {
  const select = document.getElementById('overviewEmpresaSelect');
  if (!select) return;
  if (value && state.empresas.includes(value)) {
    select.value = value;
  } else {
    select.value = '';
  }
}

function updateOverviewSearchInput(value) {
  const input = document.getElementById('overviewSearch');
  if (!input) return;
  if (input.value !== value) {
    input.value = value;
  }
}

function updateOverviewModeVisibility() {
  const mode = state.overview.filters.mode || 'global';
  const startGroup = document.getElementById('overviewStartDateGroup');
  const endGroup = document.getElementById('overviewEndDateGroup');
  const startLabel = document.getElementById('overviewStartDateLabel');
  if (startGroup) {
    startGroup.classList.toggle('d-none', mode === 'global');
  }
  if (endGroup) {
    endGroup.classList.toggle('d-none', mode !== 'range');
  }
  if (startLabel) {
    startLabel.textContent = mode === 'single' ? 'Fecha específica' : 'Fecha inicial';
  }
}

function updateOverviewModeControls() {
  const modeSelect = document.getElementById('overviewFilterMode');
  if (modeSelect) {
    modeSelect.value = state.overview.filters.mode || 'global';
  }
  const startInput = document.getElementById('overviewStartDate');
  if (startInput) {
    startInput.value = state.overview.filters.startDate || '';
  }
  const endInput = document.getElementById('overviewEndDate');
  if (endInput) {
    endInput.value = state.overview.filters.endDate || '';
  }
  updateOverviewModeVisibility();
}

function resetOverviewState(options = {}) {
  const keepFilters = options.keepFilters === true;
  state.overview.items = [];
  state.overview.filteredItems = [];
  state.overview.selection = new Map();
  state.overview.itemIndex = new Map();
  state.overview.loading = false;
  if (!keepFilters) {
    state.overview.filters = { mode: 'global', startDate: '', endDate: '', search: '', alertsOnly: false };
    updateOverviewSearchInput('');
    updateOverviewModeControls();
    const alertsToggle = document.getElementById('overviewAlertsOnly');
    if (alertsToggle) {
      alertsToggle.checked = false;
    }
  }
  renderOverviewTable();
  renderOverviewSummary();
  updateOverviewSelectAllState();
}

async function loadPoOverview(options = {}) {
  if (!state.selectedEmpresa) {
    resetOverviewState({ keepFilters: true });
    return;
  }
  const overview = state.overview;
  const mode = overview.filters.mode || 'global';
  if (mode === 'range') {
    if (!overview.filters.startDate || !overview.filters.endDate) {
      showAlert('Selecciona las fechas de inicio y fin para aplicar el rango.', 'warning');
      return;
    }
    if (overview.filters.startDate > overview.filters.endDate) {
      showAlert('La fecha inicial no puede ser posterior a la fecha final.', 'warning');
      return;
    }
  }
  if (mode === 'single' && !overview.filters.startDate) {
    showAlert('Selecciona la fecha para aplicar el filtro unitario.', 'warning');
    return;
  }

  overview.loading = true;
  renderOverviewTable();
  renderOverviewSummary();

  try {
    const params = new URLSearchParams();
    params.set('mode', mode);
    if (mode === 'range') {
      params.set('startDate', overview.filters.startDate);
      params.set('endDate', overview.filters.endDate);
    } else if (mode === 'single') {
      params.set('date', overview.filters.startDate);
    }
    const query = params.toString() ? `?${params.toString()}` : '';
    const response = await fetch(`/po-overview/${state.selectedEmpresa}${query}`);
    const data = await response.json();
    if (!response.ok || !data.success) {
      throw new Error(data.message || 'No fue posible obtener el listado de POs');
    }
    const items = Array.isArray(data.summary?.items) ? data.summary.items : [];
    const normalizedItems = items.map(item => {
      const id = typeof item.id === 'string' ? item.id.trim() : '';
      const baseIdRaw = typeof item.baseId === 'string' ? item.baseId.trim() : '';
      const baseId = baseIdRaw || id;
      const totals = getTotalsWithDefaults(item.totals);
      return {
        ...item,
        id,
        baseId,
        fecha: item.fecha || '',
        remisiones: Array.isArray(item.remisiones) ? item.remisiones : [],
        facturas: Array.isArray(item.facturas) ? item.facturas : [],
        remisionesTexto: item.remisionesTexto || '',
        facturasTexto: item.facturasTexto || '',
        totals
      };
    });
    overview.items = normalizedItems;
    overview.itemIndex = new Map(normalizedItems.map(item => [item.id, item]));
    if (options.preserveSelection) {
      const preserved = new Map();
      normalizedItems.forEach(item => {
        const baseId = item.baseId;
        const variantId = item.id;
        if (!baseId || !variantId) return;
        const existing = state.overview.selection.get(baseId);
        if (existing && existing.has(variantId)) {
          if (!preserved.has(baseId)) {
            preserved.set(baseId, new Set());
          }
          preserved.get(baseId).add(variantId);
        }
      });
      overview.selection = preserved;
    } else {
      const defaultSelection = new Map();
      normalizedItems.forEach(item => {
        const baseId = item.baseId;
        const variantId = item.id;
        if (!baseId || !variantId) return;
        if (!defaultSelection.has(baseId)) {
          defaultSelection.set(baseId, new Set());
        }
        defaultSelection.get(baseId).add(variantId);
      });
      overview.selection = defaultSelection;
    }
    overview.loading = false;
    applyOverviewFilters();
    const overviewPanel = document.getElementById('panel-overview');
    const panelIsActive = overviewPanel?.classList.contains('active');
    const shouldAutoApply = !options.skipAutoApply && (options.forceAutoApply || panelIsActive);
    if (shouldAutoApply) {
      await handleOverviewApplySelection();
    }
  } catch (error) {
    overview.loading = false;
    console.error('Error cargando resumen general de POs:', error);
    showAlert(error.message || 'Error consultando la tabla principal de POs', 'danger');
    overview.items = [];
    overview.filteredItems = [];
    overview.selection = new Map();
    overview.itemIndex = new Map();
    renderOverviewTable();
    renderOverviewSummary();
  }
}

function isOverviewItemSelected(item) {
  if (!item) return false;
  const baseId = typeof item.baseId === 'string' ? item.baseId.trim() : '';
  const variantId = typeof item.id === 'string' ? item.id.trim() : '';
  if (!baseId || !variantId) return false;
  const variants = state.overview.selection.get(baseId);
  return variants ? variants.has(variantId) : false;
}

function setOverviewItemSelected(item, selected) {
  if (!item) return false;
  const baseId = typeof item.baseId === 'string' ? item.baseId.trim() : '';
  const variantId = typeof item.id === 'string' ? item.id.trim() : '';
  if (!baseId || !variantId) return false;
  let variants = state.overview.selection.get(baseId);
  if (!variants) {
    if (!selected) {
      return false;
    }
    variants = new Set();
    state.overview.selection.set(baseId, variants);
  }
  const hasVariant = variants.has(variantId);
  if (selected) {
    if (hasVariant) {
      return false;
    }
    variants.add(variantId);
    return true;
  }
  if (hasVariant) {
    variants.delete(variantId);
    if (variants.size === 0) {
      state.overview.selection.delete(baseId);
    }
    return true;
  }
  return false;
}

function applyOverviewFilters() {
  const items = Array.isArray(state.overview.items) ? state.overview.items : [];
  const term = (state.overview.filters.search || '').trim().toLowerCase();
  const alertsOnly = state.overview.filters.alertsOnly === true;
  const filtered = items.filter(item => {
    const fields = [item.id, item.baseId, item.remisionesTexto, item.facturasTexto];
    const matchesSearch = !term
      || fields.some(value => typeof value === 'string' && value.toLowerCase().includes(term));
    if (!matchesSearch) {
      return false;
    }
    if (!alertsOnly) {
      return true;
    }
    const totals = getTotalsWithDefaults(item.totals);
    const total = normalizeNumber(totals.total || item.total);
    const totalRem = normalizeNumber(totals.totalRem);
    const totalFac = normalizeNumber(totals.totalFac);
    const consumido = totals.totalConsumo != null
      ? normalizeNumber(totals.totalConsumo)
      : roundTo(totalRem + totalFac);
    const alertLevel = getOverviewAlertInfo(consumido, total).level;
    return alertLevel === 'warning' || alertLevel === 'critical';
  });
  state.overview.filteredItems = filtered;
  renderOverviewTable();
  renderOverviewSummary();
}

function getOverviewAlertInfo(consumido, total) {
  const totalAmount = normalizeNumber(total);
  const consumedAmount = normalizeNumber(consumido);
  if (totalAmount <= 0) {
    return {
      level: 'unknown',
      icon: '—',
      label: 'Sin dato',
      description: 'No hay monto autorizado para calcular el consumo de la PO.',
      subtext: 'Sin presupuesto registrado',
      badgeClass: 'po-alert-chip-muted',
      rowClass: '',
      ariaLabel: 'Sin datos de consumo disponibles.'
    };
  }
  const rawPercentage = (consumedAmount / totalAmount) * 100;
  const limitedPercentage = Math.max(0, Math.min(rawPercentage, 999.99));
  const formattedPercentage = formatPercentageLabel(limitedPercentage);
  const consumptionPhrase = `Consumo ${formattedPercentage}`;
  if (rawPercentage >= 100) {
    return {
      level: 'critical',
      icon: '⛔',
      label: 'Crítica',
      description: `${consumptionPhrase} del presupuesto autorizado.`,
      subtext: consumptionPhrase,
      badgeClass: 'po-alert-chip-critical',
      rowClass: 'overview-alert-critical',
      ariaLabel: `Alerta crítica: ${consumptionPhrase} del presupuesto autorizado.`
    };
  }
  if (rawPercentage >= 90) {
    return {
      level: 'warning',
      icon: '⚠️',
      label: 'Atención',
      description: `${consumptionPhrase} del presupuesto autorizado.`,
      subtext: consumptionPhrase,
      badgeClass: 'po-alert-chip-warning',
      rowClass: 'overview-alert-warning',
      ariaLabel: `Alerta preventiva: ${consumptionPhrase} del presupuesto autorizado.`
    };
  }
  return {
    level: 'safe',
    icon: '✓',
    label: 'En rango',
    description: `${consumptionPhrase} del presupuesto autorizado.`,
    subtext: consumptionPhrase,
    badgeClass: 'po-alert-chip-safe',
    rowClass: '',
    ariaLabel: `Consumo dentro de rango: ${consumptionPhrase} del presupuesto autorizado.`
  };
}

function renderOverviewTable() {
  const tbody = document.getElementById('overviewTableBody');
  if (!tbody) return;
  const columns = 9;
  if (!state.selectedEmpresa) {
    tbody.innerHTML = `<tr><td colspan="${columns}" class="text-center py-4 text-muted">Selecciona una empresa para consultar sus POs.</td></tr>`;
    updateOverviewSelectAllState();
    return;
  }
  if (state.overview.loading) {
    tbody.innerHTML = `<tr><td colspan="${columns}" class="text-center py-4 text-muted">Cargando POs...</td></tr>`;
    updateOverviewSelectAllState();
    return;
  }
  const filtered = Array.isArray(state.overview.filteredItems) ? state.overview.filteredItems : [];
  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="${columns}" class="text-center py-4 text-muted">No se encontraron POs con los filtros aplicados.</td></tr>`;
    updateOverviewSelectAllState();
    return;
  }
  const rows = filtered
    .map(item => {
      const totals = getTotalsWithDefaults(item.totals);
      const total = normalizeNumber(totals.total || item.total);
      const totalRem = normalizeNumber(totals.totalRem);
      const totalFac = normalizeNumber(totals.totalFac);
      const consumido = totals.totalConsumo != null ? normalizeNumber(totals.totalConsumo) : roundTo(totalRem + totalFac);
      const restante = totals.restante != null ? normalizeNumber(totals.restante) : Math.max(total - consumido, 0);
      const remCount = Array.isArray(item.remisiones) ? item.remisiones.length : 0;
      const facCount = Array.isArray(item.facturas) ? item.facturas.length : 0;
      const selected = isOverviewItemSelected(item);
      const isExtension = item.baseId && item.baseId !== item.id;
      const baseLabel = isExtension ? `Base ${escapeHtml(item.baseId)}` : 'PO base';
      const remTooltip = escapeHtml(item.remisionesTexto || '');
      const facTooltip = escapeHtml(item.facturasTexto || '');
      const alertInfo = getOverviewAlertInfo(consumido, total);
      const rowClasses = [selected ? '' : 'overview-row-unchecked', alertInfo.rowClass || '']
        .filter(Boolean)
        .join(' ');
      const alertChip = `
        <span
          class="po-alert-chip ${alertInfo.badgeClass}"
          role="img"
          aria-label="${escapeHtml(alertInfo.ariaLabel)}"
          title="${escapeHtml(alertInfo.description)}"
        >
          <span class="po-alert-icon" aria-hidden="true">${escapeHtml(alertInfo.icon)}</span>
          <span class="po-alert-chip-label">${escapeHtml(alertInfo.label)}</span>
        </span>
      `;
      const alertSubtext = alertInfo.subtext
        ? `<div class="po-alert-chip-subtext text-muted">${escapeHtml(alertInfo.subtext)}</div>`
        : '';
      return `
        <tr data-po-id="${escapeHtml(item.id)}" data-base-id="${escapeHtml(item.baseId)}" class="${rowClasses}">
          <td class="text-center">
            <input class="form-check-input" type="checkbox" data-po-checkbox aria-label="Seleccionar ${escapeHtml(item.id)}" ${selected ? 'checked' : ''}>
          </td>
          <td>
            <div class="fw-semibold">${escapeHtml(item.id)}</div>
            <div class="small text-muted">${baseLabel}</div>
          </td>
          <td class="text-center">
            ${alertChip}
            ${alertSubtext}
          </td>
          <td>${escapeHtml(item.fecha || '-')}</td>
          <td class="text-end">$${formatCurrency(total)}</td>
          <td class="text-end" title="${remTooltip}">
            <div class="fw-semibold">$${formatCurrency(totalRem)}</div>
            <div class="small text-muted">${remCount || 0} ${remCount === 1 ? 'doc.' : 'docs.'}</div>
          </td>
          <td class="text-end" title="${facTooltip}">
            <div class="fw-semibold">$${formatCurrency(totalFac)}</div>
            <div class="small text-muted">${facCount || 0} ${facCount === 1 ? 'doc.' : 'docs.'}</div>
          </td>
          <td class="text-end">$${formatCurrency(consumido)}</td>
          <td class="text-end">$${formatCurrency(restante)}</td>
        </tr>
      `;
    })
    .join('');
  tbody.innerHTML = rows;
  updateOverviewSelectAllState();
}

function renderOverviewSummary() {
  const label = document.getElementById('overviewSummaryLabel');
  const badges = document.getElementById('overviewSummaryBadges');
  if (!label || !badges) return;
  if (!state.selectedEmpresa) {
    label.textContent = 'Selecciona una empresa para ver sus POs.';
    badges.innerHTML = '';
    return;
  }
  if (state.overview.loading) {
    label.textContent = 'Actualizando tabla...';
    badges.innerHTML = '';
    return;
  }
  const filtered = Array.isArray(state.overview.filteredItems) ? state.overview.filteredItems : [];
  if (!filtered.length) {
    label.textContent = 'No se encontraron POs con los filtros aplicados.';
    badges.innerHTML = '';
    return;
  }
  let selectedCount = 0;
  let total = 0;
  let totalRem = 0;
  let totalFac = 0;
  filtered.forEach(item => {
    if (!isOverviewItemSelected(item)) {
      return;
    }
    selectedCount += 1;
    const totals = getTotalsWithDefaults(item.totals);
    const itemTotal = normalizeNumber(totals.total || item.total);
    const itemRem = normalizeNumber(totals.totalRem);
    const itemFac = normalizeNumber(totals.totalFac);
    total += itemTotal;
    totalRem += itemRem;
    totalFac += itemFac;
  });
  label.textContent = `${selectedCount} de ${filtered.length} POs visibles seleccionadas`;
  if (selectedCount === 0) {
    badges.innerHTML = '';
    return;
  }
  const consumido = roundTo(totalRem + totalFac);
  const disponible = Math.max(roundTo(total - consumido), 0);
  badges.innerHTML = `
    <span class="badge text-bg-primary-subtle text-primary-emphasis">Total $${formatCurrency(total)}</span>
    <span class="badge text-bg-warning-subtle text-warning-emphasis">Consumido $${formatCurrency(consumido)}</span>
    <span class="badge text-bg-success-subtle text-success-emphasis">Disponible $${formatCurrency(disponible)}</span>
  `;
}

function updateOverviewSelectAllState() {
  const selectAll = document.getElementById('overviewSelectAll');
  if (!selectAll) return;
  const filtered = Array.isArray(state.overview.filteredItems) ? state.overview.filteredItems : [];
  if (!state.selectedEmpresa || state.overview.loading || filtered.length === 0) {
    selectAll.checked = false;
    selectAll.indeterminate = false;
    selectAll.disabled = true;
    return;
  }
  let selectedRows = 0;
  filtered.forEach(item => {
    if (isOverviewItemSelected(item)) {
      selectedRows += 1;
    }
  });
  selectAll.disabled = false;
  if (selectedRows === 0) {
    selectAll.checked = false;
    selectAll.indeterminate = false;
  } else if (selectedRows === filtered.length) {
    selectAll.checked = true;
    selectAll.indeterminate = false;
  } else {
    selectAll.checked = false;
    selectAll.indeterminate = true;
  }
}

function getOverviewSelectionEntries() {
  const filtered = Array.isArray(state.overview.filteredItems) ? state.overview.filteredItems : [];
  const entries = new Map();
  filtered.forEach(item => {
    if (!isOverviewItemSelected(item)) {
      return;
    }
    const baseId = typeof item.baseId === 'string' ? item.baseId.trim() : '';
    const variantId = typeof item.id === 'string' ? item.id.trim() : '';
    if (!baseId || !variantId) {
      return;
    }
    if (!entries.has(baseId)) {
      entries.set(baseId, new Set());
    }
    entries.get(baseId).add(variantId);
  });
  return Array.from(entries.entries())
    .map(([baseId, variants]) => ({ baseId, variants: Array.from(variants) }))
    .sort((a, b) => a.baseId.localeCompare(b.baseId));
}

function syncOverviewSelectionToState() {
  if (!state.selectedEmpresa) {
    showAlert('Selecciona una empresa válida antes de continuar.', 'warning');
    return false;
  }
  const entries = getOverviewSelectionEntries();
  if (entries.length === 0) {
    showAlert('Selecciona al menos una PO visible en la tabla principal.', 'warning');
    return false;
  }
  state.selectedPoIds = entries.map(entry => entry.baseId);
  state.selectedPoDetails = new Map();
  state.manualExtensions.clear();
  entries.forEach(entry => {
    state.selectedPoDetails.set(entry.baseId, { baseId: entry.baseId, variants: new Set(entry.variants) });
  });
  renderSelectedPoChips();
  renderPoOptions();
  renderReportSelectionOverview();
  updatePoFoundList();
  return true;
}

function handleOverviewSelectAllChange(event) {
  const checkbox = event.target;
  if (!checkbox) return;
  const filtered = Array.isArray(state.overview.filteredItems) ? state.overview.filteredItems : [];
  filtered.forEach(item => {
    setOverviewItemSelected(item, checkbox.checked);
  });
  const tbody = document.getElementById('overviewTableBody');
  if (tbody) {
    tbody.querySelectorAll('tr[data-po-id]').forEach(row => {
      const poId = row.getAttribute('data-po-id');
      const item = state.overview.itemIndex.get(poId || '');
      const selected = item ? isOverviewItemSelected(item) : false;
      const input = row.querySelector('input[data-po-checkbox]');
      if (input) {
        input.checked = selected;
      }
      row.classList.toggle('overview-row-unchecked', !selected);
    });
  }
  renderOverviewSummary();
  updateOverviewSelectAllState();
}

function handleOverviewTableChange(event) {
  const checkbox = event.target.closest('input[data-po-checkbox]');
  if (!checkbox) return;
  const row = checkbox.closest('tr[data-po-id]');
  if (!row) return;
  const poId = row.getAttribute('data-po-id');
  const item = state.overview.itemIndex.get(poId || '');
  if (!item) {
    checkbox.checked = false;
    return;
  }
  setOverviewItemSelected(item, checkbox.checked);
  row.classList.toggle('overview-row-unchecked', !checkbox.checked);
  renderOverviewSummary();
  updateOverviewSelectAllState();
}

function handleOverviewModeChange(mode) {
  const normalized = typeof mode === 'string' ? mode.trim().toLowerCase() : 'global';
  state.overview.filters.mode = ['range', 'single'].includes(normalized) ? normalized : 'global';
  if (state.overview.filters.mode === 'global') {
    state.overview.filters.startDate = '';
    state.overview.filters.endDate = '';
  }
  if (state.overview.filters.mode !== 'range') {
    state.overview.filters.endDate = '';
  }
  updateOverviewModeControls();
  if (state.overview.filters.mode === 'global') {
    if (state.selectedEmpresa) {
      loadPoOverview();
    }
    return;
  }
  handleOverviewDateChange();
}

function handleOverviewDateChange() {
  if (!state.selectedEmpresa) return;
  const mode = state.overview.filters.mode || 'global';
  if (mode === 'range') {
    if (!state.overview.filters.startDate || !state.overview.filters.endDate) {
      return;
    }
    if (state.overview.filters.startDate > state.overview.filters.endDate) {
      showAlert('La fecha inicial no puede ser posterior a la fecha final.', 'warning');
      return;
    }
    loadPoOverview({ preserveSelection: false });
    return;
  }
  if (mode === 'single') {
    if (!state.overview.filters.startDate) {
      return;
    }
    loadPoOverview({ preserveSelection: false });
    return;
  }
  loadPoOverview({ preserveSelection: false });
}

async function handleOverviewApplySelection() {
  if (!syncOverviewSelectionToState()) {
    return;
  }
  await updateSummary();
}

async function handleOverviewGenerateReport() {
  if (!syncOverviewSelectionToState()) {
    return;
  }
  await openReportPreview();
}

function initializeOverviewTab() {
  renderOverviewEmpresaOptions();
  updateOverviewModeControls();
  updateOverviewSearchInput(state.overview.filters.search || '');
  renderOverviewTable();
  renderOverviewSummary();
  const empresaSelect = document.getElementById('overviewEmpresaSelect');
  empresaSelect?.addEventListener('change', event => {
    const value = event.target.value;
    if (!value) {
      setSearchableSelectValue(document.getElementById('empresaSearch'), '');
      state.selectedEmpresa = '';
      state.pos = [];
      state.selectedPoIds = [];
      state.selectedPoDetails.clear();
      state.manualExtensions.clear();
      resetOverviewState({ keepFilters: true });
      renderSelectedPoChips();
      renderPoOptions();
      renderReportSelectionOverview();
      updateUniverseControls();
      destroyCharts();
      hideDashboardPanel();
      resetSearchableSelect(document.getElementById('poSearch'));
      return;
    }
    selectEmpresa(value);
  });
  document.getElementById('overviewFilterMode')?.addEventListener('change', event => handleOverviewModeChange(event.target.value));
  document.getElementById('overviewStartDate')?.addEventListener('change', event => {
    state.overview.filters.startDate = event.target.value;
    handleOverviewDateChange();
  });
  document.getElementById('overviewEndDate')?.addEventListener('change', event => {
    state.overview.filters.endDate = event.target.value;
    handleOverviewDateChange();
  });
  document.getElementById('overviewSearch')?.addEventListener('input', event => {
    state.overview.filters.search = event.target.value || '';
    applyOverviewFilters();
  });
  const alertsToggle = document.getElementById('overviewAlertsOnly');
  if (alertsToggle) {
    alertsToggle.checked = state.overview.filters.alertsOnly === true;
    alertsToggle.addEventListener('change', event => {
      state.overview.filters.alertsOnly = event.target.checked;
      applyOverviewFilters();
    });
  }
  document.getElementById('overviewRefreshBtn')?.addEventListener('click', () => loadPoOverview({ preserveSelection: false }));
  document.getElementById('overviewApplySelectionBtn')?.addEventListener('click', handleOverviewApplySelection);
  document.getElementById('overviewReportBtn')?.addEventListener('click', handleOverviewGenerateReport);
  document.getElementById('overviewSelectAll')?.addEventListener('change', handleOverviewSelectAllChange);
  document.getElementById('overviewTableBody')?.addEventListener('change', handleOverviewTableChange);
  updateOverviewSelectAllState();
}

async function selectPo(poId) {
  if (!poId) return;
  addPoToSelection(poId);
  resetSearchableSelect(document.getElementById('poSearch'));
}

function renderSummaryCards(summary) {
  const container = document.getElementById('summaryCards');
  if (!container) return;
  const rawTotals = getTotalsWithDefaults(summary.totals);
  const total = normalizeNumber(rawTotals.total);
  const totalRem = normalizeNumber(rawTotals.totalRem);
  const totalFac = normalizeNumber(rawTotals.totalFac);
  const restante = rawTotals.restante != null
    ? normalizeNumber(rawTotals.restante)
    : Math.max(total - (totalRem + totalFac), 0);
  const totalConsumo = rawTotals.totalConsumo != null
    ? normalizeNumber(rawTotals.totalConsumo)
    : roundTo(totalRem + totalFac);
  const totals = {
    ...rawTotals,
    total,
    totalRem,
    totalFac,
    restante,
    totalConsumo
  };
  const percentages = computePercentagesFromTotals(totals);
  const selectionCount = (summary.selectedIds || []).length || 0;
  const selectionLabel = selectionCount === 1
    ? '1 PO base (extensiones según tu selección)'
    : `${selectionCount} POs base combinadas`;
  const companyLabel = escapeHtml(summary.empresaLabel || summary.empresa || 'Empresa');
  container.innerHTML = `
    <article class="summary-card summary-card-total">
      <span class="summary-card-title">${selectionCount > 1 ? 'Total combinado' : 'Total PO (Grupo)'}</span>
      <p class="summary-card-value">$${formatCurrency(totals.total)}</p>
      <p class="summary-card-subtext">${escapeHtml(selectionLabel)}.</p>
      <span class="badge rounded-pill text-bg-light text-dark fw-semibold mt-3">${companyLabel}</span>
    </article>
    <article class="summary-card summary-card-consumo">
      <span class="summary-card-title">Consumo acumulado</span>
      <p class="summary-card-value">$${formatCurrency(totals.totalConsumo)}</p>
      <p class="summary-card-subtext">${formatPercentageLabel(percentages.rem)} Remisiones · ${formatPercentageLabel(percentages.fac)} Facturas</p>
    </article>
    <article class="summary-card summary-card-disponible">
      <span class="summary-card-title">Disponible</span>
      <p class="summary-card-value">$${formatCurrency(totals.restante)}</p>
      <p class="summary-card-subtext">${formatPercentageLabel(percentages.rest)} restante del presupuesto autorizado.</p>
    </article>
  `;
}

function renderTable(summary) {
  const table = document.getElementById('poTable');
  if (!table) return;
  table.setAttribute('aria-busy', 'true');
  try {
    const items = Array.isArray(summary?.items) ? summary.items : [];
    const rows = items
      .map(item => {
        const totalAmount = Number(item.total || 0);
        const totalOriginalAmount = Number(item.totalOriginal ?? item.total ?? 0);
        const subtotalAmount = Number(item.subtotal || 0);
        const showSubtotal = subtotalAmount > 0 && Math.abs(subtotalAmount - totalAmount) > 0.009;
        const adjustment = item.ajusteDocSig;
        const appliedDiff = Number(adjustment?.diferenciaAplicada ?? 0);
        const hasAdjustment = appliedDiff > 0.009;
        const totalMetaLines = [];
        if (hasAdjustment) {
          totalMetaLines.push(`<span class="table-meta text-muted">Original: $${formatCurrency(totalOriginalAmount)}</span>`);
          const tipoLabel = adjustment?.tipo === 'F' ? 'Factura' : 'Remisión';
          const docLabel = escapeHtml(adjustment?.docSig || '-');
          totalMetaLines.push(
            `<span class="table-meta text-warning">Ajuste por ${tipoLabel} ${docLabel}: -$${formatCurrency(appliedDiff)}</span>`
          );
        }
        if (showSubtotal) {
          totalMetaLines.push(`<span class="table-meta text-muted">Subtotal sin IVA: $${formatCurrency(subtotalAmount)}</span>`);
        }
        const totals = getTotalsWithDefaults(item.totals);
        const totalBase = normalizeNumber(totals.total || totalAmount);
        const totalRem = normalizeNumber(totals.totalRem);
        const totalFac = normalizeNumber(totals.totalFac);
        const restante = totals.restante != null
          ? normalizeNumber(totals.restante)
          : Math.max(totalBase - (totalRem + totalFac), 0);
        const normalizedTotals = {
          ...totals,
          total: totalBase,
          totalRem,
          totalFac,
          restante
        };
        const percentages = computePercentagesFromTotals(normalizedTotals);
        const consumptionPercentage = clampPercentage(percentages.rem + percentages.fac);
        let alertClass = 'po-alert-safe';
        let rowAlertClass = 'po-row-safe';
        if (consumptionPercentage >= 100) {
          alertClass = 'po-alert-critical';
          rowAlertClass = 'po-row-critical';
        } else if (consumptionPercentage >= 90) {
          alertClass = 'po-alert-warning';
          rowAlertClass = 'po-row-warning';
        }
        const rawAlertsText = typeof item.alertasTexto === 'string' ? item.alertasTexto : '';
        const alertLines = rawAlertsText
          ? rawAlertsText.split(/\n+/u).map(line => line.trim()).filter(Boolean)
          : [];
        const meaningfulAlerts = alertLines.filter(line => !/^sin alertas/i.test(line));
        const primaryAlertLine = meaningfulAlerts[0] || '';
        const normalizedPrimaryAlert = primaryAlertLine.replace(/^\[[^\]]*\]\s*/u, '').trim();
        const alertMessage = normalizedPrimaryAlert
          || (alertClass === 'po-alert-critical'
            ? 'Consumo al 100%'
            : alertClass === 'po-alert-warning'
              ? `Consumo ${formatPercentageLabel(consumptionPercentage)}`
              : 'Sin alertas relevantes');
        const detailParts = [`Consumo ${formatPercentageLabel(consumptionPercentage)}`];
        if (meaningfulAlerts.length > 1) {
          detailParts.unshift(`${meaningfulAlerts.length} alertas registradas`);
        }
        const alertDetail = detailParts.join(' · ');
        return `
        <tr data-po="${escapeHtml(item.id || '')}" class="po-alert-row ${rowAlertClass}">
          <td class="fw-semibold">${escapeHtml(item.id || '')}</td>
          <td>${escapeHtml(item.fecha || '-')}</td>
          <td>
            <span class="table-amount">$${formatCurrency(totalAmount)}</span>
            ${totalMetaLines.join('')}
          </td>
          <td>
            <span class="table-amount">$${formatCurrency(normalizedTotals.totalRem)}</span>
            <span class="table-meta">${formatPercentageLabel(percentages.rem)}</span>
          </td>
          <td>
            <span class="table-amount">$${formatCurrency(normalizedTotals.totalFac)}</span>
            <span class="table-meta">${formatPercentageLabel(percentages.fac)}</span>
          </td>
          <td>
            <span class="table-amount">$${formatCurrency(normalizedTotals.restante)}</span>
            <span class="table-meta">${formatPercentageLabel(percentages.rest)}</span>
          </td>
          <td class="po-alert-cell ${alertClass}">
            <span>${escapeHtml(alertMessage)}</span>
            <span class="po-alert-detail">${escapeHtml(alertDetail)}</span>
          </td>
          <td class="text-center">
            <button class="btn btn-outline-primary btn-sm" data-po="${escapeHtml(item.id || '')}" data-action="show-modal">Detalle</button>
          </td>
        </tr>
        `;
      })
      .join('');
    const emptyState = `
        <tr>
          <td class="text-center text-muted py-4" colspan="8">No hay POs seleccionadas.</td>
        </tr>
      `;
    table.innerHTML = `
      <thead>
        <tr>
          <th scope="col">PO</th>
          <th scope="col">Fecha</th>
          <th scope="col">
            <span class="table-heading-title">Total autorizado</span>
            <span class="table-heading-sub">Monto aprobado</span>
          </th>
          <th scope="col">
            <span class="table-heading-title">Remisiones</span>
            <span class="table-heading-sub">Monto · % consumo</span>
          </th>
          <th scope="col">
            <span class="table-heading-title">Facturas</span>
            <span class="table-heading-sub">Monto · % consumo</span>
          </th>
          <th scope="col">
            <span class="table-heading-title">Disponible</span>
            <span class="table-heading-sub">Saldo restante</span>
          </th>
          <th scope="col">
            <span class="table-heading-title">Alerta</span>
            <span class="table-heading-sub">Estado de consumo</span>
          </th>
          <th scope="col" class="text-center">Acción</th>
        </tr>
      </thead>
      <tbody>${rows || emptyState}</tbody>
    `;
  } finally {
    table.setAttribute('aria-busy', 'false');
  }
}

function attachTableHandlers(summary) {
  const table = document.getElementById('poTable');
  const modalElement = document.getElementById('modalExt');
  if (!table || !modalElement) return;
  const modalBody = modalElement.querySelector('.modal-body');
  const modalTitle = modalElement.querySelector('.modal-title');
  table.onclick = event => {
    const button = event.target.closest('button[data-action="show-modal"]');
    if (!button) return;
    const poId = button.getAttribute('data-po');
    const item = summary.items.find(it => it.id === poId);
    if (!item) return;
    modalTitle.textContent = `Detalle PO ${item.id}`;
    const totalAutorizado = formatCurrency(item.total || 0);
    const subtotalAmount = Number(item.subtotal || 0);
    const subtotalLinea =
      subtotalAmount > 0 ? `Subtotal sin IVA: $${formatCurrency(subtotalAmount)}` : '';
    const totalesResumen = [
      `Total de importe: $${totalAutorizado}`,
      subtotalLinea,
      `Remisiones acumuladas: $${formatCurrency(item.totals?.totalRem ?? 0)}`,
      `Facturas acumuladas: $${formatCurrency(item.totals?.totalFac ?? 0)}`,
      `Disponible: $${formatCurrency(item.totals?.restante ?? 0)}`
    ]
      .filter(Boolean)
      .join('\n');
    const resumenHtml = escapeHtml(totalesResumen);
    const remisionesHtml = escapeHtml(item.remisionesTexto || 'Sin remisiones registradas');
    const facturasHtml = escapeHtml(item.facturasTexto || 'Sin facturas registradas');
    const alertasHtml = escapeHtml(item.alertasTexto || 'Sin alertas');
    modalBody.innerHTML = `
      <div class="mb-3">
        <h6 class="text-secondary">Totales del pedido</h6>
        <pre class="bg-light rounded p-3 small">${resumenHtml}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-primary">Remisiones</h6>
        <pre class="bg-light rounded p-3 small">${remisionesHtml}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-danger">Facturas</h6>
        <pre class="bg-light rounded p-3 small">${facturasHtml}</pre>
      </div>
      <div class="mb-0">
        <h6 class="text-danger">Alertas</h6>
        <pre class="bg-light border border-danger-subtle rounded p-3 small text-danger fw-semibold">${alertasHtml}</pre>
      </div>
    `;
    const modal = new bootstrap.Modal(modalElement);
    modal.show();
  };
}

function renderCharts(summary) {
  destroyCharts();
  ['chartRem', 'chartFac', 'chartJunto'].forEach(id => {
    const canvas = document.getElementById(id);
    if (canvas) {
      resetChartCanvas(canvas);
    }
  });
  const items = Array.isArray(summary.items) ? summary.items : [];
  const groups = buildPoGroups(items, summary.selectionDetails);
  const extensionContainer = document.getElementById('extensionsContainer');
  if (!groups.length) {
    if (extensionContainer) {
      extensionContainer.innerHTML = '<p class="text-muted">Sin información suficiente para graficar esta selección.</p>';
    }
    return;
  }

  const totals = getTotalsWithDefaults(summary.totals);
  const total = normalizeNumber(totals.total);
  const totalRem = normalizeNumber(totals.totalRem);
  const totalFac = normalizeNumber(totals.totalFac);
  const restante = totals.restante != null
    ? normalizeNumber(totals.restante)
    : Math.max(total - (totalRem + totalFac), 0);
  const aggregatedTotals = {
    total,
    totalRem,
    totalFac,
    restante,
    totalConsumo: roundTo(totalRem + totalFac)
  };
  const aggregatedPercentages = computePercentagesFromTotals(aggregatedTotals);
  const combinedScaleBase = aggregatedTotals.total > 0
    ? aggregatedTotals.total
    : Math.max(roundTo(aggregatedTotals.totalRem + aggregatedTotals.totalFac), 0);
  const percentageBase = combinedScaleBase > 0 ? combinedScaleBase : 1;

  const chartEntries = groups.map(group => {
    const variantCount = group.ids.length;
    const extensionIds = Array.isArray(group.extensionIds) ? group.extensionIds : [];
    const extensionCount = extensionIds.length;
    const suffix = extensionCount > 0 ? ` (base + ${extensionCount} ext.)` : '';
    return {
      id: group.baseId,
      label: `PO ${group.baseId}${suffix}`,
      shortLabel: extensionCount > 0 ? `${group.baseId} (+${extensionCount})` : group.baseId,
      variantCount,
      variantIds: group.ids,
      extensionIds,
      totals: group.totals,
      percentages: group.percentages,
      globalPercentages: computeGlobalPercentages(group.totals, percentageBase),
      items: Array.isArray(group.items) ? group.items : []
    };
  });

  const topRemEntry = chartEntries.reduce((acc, entry) => (
    !acc || entry.totals.totalRem > acc.totals.totalRem ? entry : acc
  ), null);
  const topFacEntry = chartEntries.reduce((acc, entry) => (
    !acc || entry.totals.totalFac > acc.totals.totalFac ? entry : acc
  ), null);

  const mainMetrics = [
    {
      containerId: 'chartRemMetrics',
      rows: [
        {
          label: 'Remisiones acumuladas',
          className: 'metric-rem',
          amount: aggregatedTotals.totalRem,
          percentage: aggregatedPercentages.rem
        },
        topRemEntry
          ? {
              label: `Mayor consumo: ${escapeHtml(topRemEntry.label)}`,
              className: 'text-muted',
              amount: topRemEntry.totals.totalRem,
              percentage: aggregatedTotals.totalRem > 0
                ? clampPercentage((topRemEntry.totals.totalRem / aggregatedTotals.totalRem) * 100)
                : 0
            }
          : null
      ].filter(Boolean)
    },
    {
      containerId: 'chartFacMetrics',
      rows: [
        {
          label: 'Facturas acumuladas',
          className: 'metric-fac',
          amount: aggregatedTotals.totalFac,
          percentage: aggregatedPercentages.fac
        },
        topFacEntry
          ? {
              label: `Mayor consumo: ${escapeHtml(topFacEntry.label)}`,
              className: 'text-muted',
              amount: topFacEntry.totals.totalFac,
              percentage: aggregatedTotals.totalFac > 0
                ? clampPercentage((topFacEntry.totals.totalFac / aggregatedTotals.totalFac) * 100)
                : 0
            }
          : null
      ].filter(Boolean)
    },
    {
      containerId: 'chartStackMetrics',
      rows: [
        {
          label: 'Consumido (Rem + Fac)',
          className: 'metric-rem',
          amount: aggregatedTotals.totalRem + aggregatedTotals.totalFac,
          percentage: roundTo(aggregatedPercentages.rem + aggregatedPercentages.fac)
        },
        {
          label: 'Disponible',
          className: 'metric-rest',
          amount: aggregatedTotals.restante,
          percentage: aggregatedPercentages.rest
        }
      ]
    }
  ];

  mainMetrics.forEach(definition => {
    const container = document.getElementById(definition.containerId);
    if (!container) return;
    container.innerHTML = definition.rows
      .map(row => `
        <div class="metric-row">
          <span class="metric-label ${row.className}">${row.label}</span>
          <span>$${formatCurrency(row.amount)} · ${formatPercentageLabel(row.percentage)}</span>
        </div>
      `)
      .join('');
  });

  const labels = chartEntries.map(entry => entry.label);
  const ctxRem = document.getElementById('chartRem');
  const ctxFac = document.getElementById('chartFac');
  const ctxStack = document.getElementById('chartJunto');
  const dynamicHeight = computeResponsiveChartHeight(chartEntries.length);
  const stackMaxPercentage = chartEntries.reduce((max, entry) => {
    const global = entry.globalPercentages || { rem: 0, fac: 0, rest: 0 };
    const totalPercentage = global.rem + global.fac + global.rest;
    return Math.max(max, totalPercentage);
  }, 0);
  const stackAxisMax = Math.max(100, Math.ceil(stackMaxPercentage / 10) * 10);

  const barChartsMeta = [
    { canvas: ctxRem, key: 'rem', amountKey: 'totalRem', percKey: 'rem', label: 'Remisiones', color: CHART_COLORS.rem },
    { canvas: ctxFac, key: 'fac', amountKey: 'totalFac', percKey: 'fac', label: 'Facturas', color: CHART_COLORS.fac }
  ];

  barChartsMeta.forEach(meta => {
    if (!meta.canvas) return;
    const wrapper = meta.canvas.closest('.chart-wrapper');
    if (wrapper) {
      wrapper.style.minHeight = `${dynamicHeight}px`;
      wrapper.style.maxHeight = `${Math.max(dynamicHeight, 240)}px`;
    }
    meta.canvas.height = dynamicHeight;
    meta.canvas.style.height = `${dynamicHeight}px`;
    const values = chartEntries.map(entry => entry.globalPercentages?.[meta.percKey] ?? 0);
    const maxValue = values.length ? Math.max(...values) : 0;
    const axisMax = Math.max(100, Math.ceil(maxValue / 10) * 10 || 100);
    const dataset = {
      label: `${meta.label} (%)`,
      data: values,
      backgroundColor: meta.color,
      borderRadius: 10,
      maxBarThickness: 32
    };
    const chart = new Chart(meta.canvas.getContext('2d'), {
      type: 'bar',
      data: { labels, datasets: [dataset] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        animation: { duration: 400 },
        layout: { padding: { top: 12, right: 16, bottom: 12, left: 8 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title(context) {
                return chartEntries[context[0].dataIndex].label;
              },
              label(context) {
                const entry = chartEntries[context.dataIndex];
                const percentage = entry.globalPercentages?.[meta.percKey] ?? 0;
                return `${meta.label}: $${formatCurrency(entry.totals[meta.amountKey])} (${formatPercentageLabel(percentage)})`;
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            max: axisMax,
            grid: { color: 'rgba(148, 163, 184, 0.25)', borderDash: [4, 4] },
            ticks: {
              color: '#475569',
              callback: value => `${formatPercentageValue(value)}%`
            }
          },
          y: {
            grid: { display: false },
            ticks: { color: '#475569', autoSkip: false }
          }
        }
      }
    });
    state.charts.set(meta.key, chart);
  });

  if (ctxStack) {
    const wrapper = ctxStack.closest('.chart-wrapper');
    if (wrapper) {
      wrapper.style.minHeight = `${dynamicHeight}px`;
      wrapper.style.maxHeight = `${Math.max(dynamicHeight, 240)}px`;
    }
    ctxStack.height = dynamicHeight;
    ctxStack.style.height = `${dynamicHeight}px`;
    const stackMeta = [
      { label: 'Remisiones', percKey: 'rem', amountKey: 'totalRem', color: CHART_COLORS.rem },
      { label: 'Facturas', percKey: 'fac', amountKey: 'totalFac', color: CHART_COLORS.fac },
      { label: 'Disponible', percKey: 'rest', amountKey: 'restante', color: CHART_COLORS.rest }
    ];
    const chart = new Chart(ctxStack.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: stackMeta.map(meta => ({
          label: `${meta.label} %`,
          data: chartEntries.map(entry => entry.globalPercentages?.[meta.percKey] ?? 0),
          backgroundColor: meta.color,
          borderRadius: 8,
          stack: 'total'
        }))
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        layout: { padding: { top: 12, right: 16, bottom: 12, left: 8 } },
        animation: { duration: 400 },
        plugins: {
          legend: {
            position: 'bottom',
            labels: { usePointStyle: true, padding: 16 }
          },
          tooltip: {
            callbacks: {
              title(context) {
                return chartEntries[context[0].dataIndex].label;
              },
              label(context) {
                const meta = stackMeta[context.datasetIndex];
                const entry = chartEntries[context.dataIndex];
                const percentage = entry.globalPercentages?.[meta.percKey] ?? 0;
                return `${meta.label}: $${formatCurrency(entry.totals[meta.amountKey])} (${formatPercentageLabel(percentage)})`;
              }
            }
          }
        },
        scales: {
          x: {
            stacked: true,
            beginAtZero: true,
            max: stackAxisMax,
            grid: { color: 'rgba(148, 163, 184, 0.25)', borderDash: [4, 4] },
            ticks: {
              color: '#475569',
              callback: value => `${formatPercentageValue(value)}%`
            }
          },
          y: {
            stacked: true,
            grid: { display: false },
            ticks: { color: '#475569', autoSkip: false }
          }
        }
      }
    });
    state.charts.set('stack', chart);
  }

  if (extensionContainer) {
    const donutMeta = [
      { label: 'Remisiones', amountKey: 'totalRem', percKey: 'rem', color: CHART_COLORS.rem, labelClass: 'metric-rem' },
      { label: 'Facturas', amountKey: 'totalFac', percKey: 'fac', color: CHART_COLORS.fac, labelClass: 'metric-fac' },
      { label: 'Disponible', amountKey: 'restante', percKey: 'rest', color: CHART_COLORS.rest, labelClass: 'metric-rest' }
    ];
    extensionContainer.innerHTML = '';
    chartEntries.forEach((entry, index) => {
      const card = document.createElement('article');
      card.className = 'card h-100 shadow-sm chart-card';
      const extensionCount = entry.extensionIds.length;
      const variantBadge = extensionCount === 0
        ? 'Solo base'
        : `Conjunto (${entry.variantCount} variantes)`;
      const donutHeight = Math.min(320, Math.max(220, 180 + Math.max(entry.variantCount - 1, 0) * 22));
      const badgeListMarkup = buildGroupBadgeList(entry.id, entry.extensionIds);
      const compositionNote = buildGroupCompositionNote(entry.extensionIds);
      const breakdownRows = entry.items
        .map(item => {
          const badge = item.isBase
            ? '<span class="badge text-bg-secondary ms-2">Base</span>'
            : '<span class="badge text-bg-info ms-2">Extensión</span>';
          return `
            <tr>
              <td>
                <span class="fw-semibold">${escapeHtml(item.id)}</span>
                ${badge}
              </td>
              <td class="text-end">
                <span class="table-amount">$${formatCurrency(item.totals.total)}</span>
              </td>
              <td class="text-end">
                <span class="table-amount">$${formatCurrency(item.totals.totalRem)}</span>
                <span class="table-meta">${formatPercentageLabel(item.percentages.rem)}</span>
              </td>
              <td class="text-end">
                <span class="table-amount">$${formatCurrency(item.totals.totalFac)}</span>
                <span class="table-meta">${formatPercentageLabel(item.percentages.fac)}</span>
              </td>
              <td class="text-end">
                <span class="table-amount">$${formatCurrency(item.totals.restante)}</span>
                <span class="table-meta">${formatPercentageLabel(item.percentages.rest)}</span>
              </td>
            </tr>
          `;
        })
        .join('');
      const breakdownTable = breakdownRows
        ? `
          <div class="group-breakdown mt-3">
            <h6 class="text-secondary">Desglose por PO</h6>
            <div class="table-responsive">
              <table class="table table-sm align-middle mb-0 group-breakdown-table table-modern table-modern-compact">
                <thead>
                  <tr>
                    <th scope="col">PO</th>
                    <th scope="col">
                      <span class="table-heading-title">Total autorizado</span>
                      <span class="table-heading-sub">Monto aprobado</span>
                    </th>
                    <th scope="col">
                      <span class="table-heading-title">Remisiones</span>
                      <span class="table-heading-sub">Monto · % consumo</span>
                    </th>
                    <th scope="col">
                      <span class="table-heading-title">Facturas</span>
                      <span class="table-heading-sub">Monto · % consumo</span>
                    </th>
                    <th scope="col">
                      <span class="table-heading-title">Disponible</span>
                      <span class="table-heading-sub">Saldo restante</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  ${breakdownRows}
                </tbody>
              </table>
            </div>
          </div>
        `
        : '';
      card.innerHTML = `
        <div class="card-body">
          <div class="d-flex justify-content-between align-items-start gap-3">
            <div>
              <h5 class="card-title mb-1">${escapeHtml(entry.label)}</h5>
              <p class="small text-muted mb-0">Autorizado: $${formatCurrency(entry.totals.total)}</p>
            </div>
            <span class="badge text-bg-light text-dark">${escapeHtml(variantBadge)}</span>
          </div>
          <div class="extension-group-summary mt-2">${badgeListMarkup}</div>
          <p class="extension-group-note mb-0">${escapeHtml(compositionNote)}</p>
          <div class="chart-wrapper mt-2" style="min-height: ${donutHeight}px; height: ${donutHeight}px; max-height: ${donutHeight}px">
            <canvas id="extChart-${index}"></canvas>
          </div>
          <div class="chart-metrics mt-2">
            ${donutMeta
              .map(meta => `
                <div class="metric-row">
                  <span class="metric-label ${meta.labelClass}">${meta.label}</span>
                  <span>$${formatCurrency(entry.totals[meta.amountKey])} · ${formatPercentageLabel(entry.percentages[meta.percKey])}</span>
                </div>
              `)
              .join('')}
          </div>
          ${breakdownTable}
        </div>
      `;
      extensionContainer.appendChild(card);
      const canvas = card.querySelector('canvas');
      const wrapper = canvas?.closest('.chart-wrapper');
      if (wrapper) {
        wrapper.style.minHeight = `${donutHeight}px`;
        wrapper.style.height = `${donutHeight}px`;
        wrapper.style.maxHeight = `${donutHeight}px`;
      }
      canvas.height = donutHeight;
      canvas.style.height = `${donutHeight}px`;
      canvas.style.maxHeight = `${donutHeight}px`;
      const dataset = {
        data: donutMeta.map(meta => entry.totals[meta.amountKey]),
        backgroundColor: donutMeta.map(meta => meta.color),
        hoverOffset: 8,
        metaInfo: donutMeta.map(meta => ({
          label: meta.label,
          amount: entry.totals[meta.amountKey],
          percentage: entry.percentages[meta.percKey]
        }))
      };
      const chart = new Chart(canvas.getContext('2d'), {
        type: 'doughnut',
        data: {
          labels: donutMeta.map(meta => meta.label),
          datasets: [dataset]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '60%',
          layout: { padding: 6 },
          plugins: {
            legend: {
              position: 'bottom',
              labels: { usePointStyle: true, padding: 16 }
            },
            tooltip: {
              callbacks: {
                title() {
                  return entry.label;
                },
                label(context) {
                  const info = dataset.metaInfo?.[context.dataIndex];
                  if (info) {
                    return `${info.label}: $${formatCurrency(info.amount)} (${formatPercentageLabel(info.percentage)})`;
                  }
                  const totalValue = dataset.data.reduce((sum, value) => sum + value, 0);
                  const percentage = totalValue > 0 ? (context.parsed / totalValue) * 100 : 0;
                  return `${context.label}: $${formatCurrency(context.parsed)} (${formatPercentageLabel(percentage)})`;
                }
              }
            }
          }
        }
      });
      state.charts.set(`ext-${index}`, chart);
    });
  }
}

function renderAlerts(alerts) {
  alerts.forEach(alert => showAlert(alert.message, alert.type || 'info', 8000));
}

function resolveCustomizationFlag(value, defaultValue) {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'false') {
      return false;
    }
    if (normalized === 'true') {
      return true;
    }
  }
  return Boolean(value);
}

function normalizeCustomizationState(customization = {}) {
  const csv = customization.csv || {};
  const defaults = DEFAULT_CUSTOMIZATION;
  return {
    includeSummary: resolveCustomizationFlag(customization.includeSummary, defaults.includeSummary),
    includeDetail: resolveCustomizationFlag(customization.includeDetail, defaults.includeDetail),
    includeCharts: resolveCustomizationFlag(customization.includeCharts, defaults.includeCharts),
    includeMovements: resolveCustomizationFlag(customization.includeMovements, defaults.includeMovements),
    includeObservations: resolveCustomizationFlag(customization.includeObservations, defaults.includeObservations),
    includeUniverse: resolveCustomizationFlag(customization.includeUniverse, defaults.includeUniverse),
    csv: {
      includePoResumen: resolveCustomizationFlag(csv.includePoResumen, defaults.csv.includePoResumen),
      includeRemisiones: resolveCustomizationFlag(csv.includeRemisiones, defaults.csv.includeRemisiones),
      includeFacturas: resolveCustomizationFlag(csv.includeFacturas, defaults.csv.includeFacturas),
      includeTotales: resolveCustomizationFlag(csv.includeTotales, defaults.csv.includeTotales),
      includeUniverseInfo: resolveCustomizationFlag(csv.includeUniverseInfo, defaults.csv.includeUniverseInfo)
    }
  };
}

function mergeCustomizationSettings(baseCustomization = {}, overrides) {
  const normalizedBase = normalizeCustomizationState(baseCustomization);
  if (!overrides) {
    return normalizedBase;
  }
  const merged = {
    ...normalizedBase,
    ...overrides,
    csv: {
      ...normalizedBase.csv,
      ...(overrides.csv || {})
    }
  };
  return normalizeCustomizationState(merged);
}

function buildCustomizationBadgesMarkup() {
  const config = normalizeCustomizationState(state.customization);
  const badges = [];
  badges.push({ label: config.includeSummary ? 'Resumen incluido' : 'Sin resumen', active: config.includeSummary });
  badges.push({ label: config.includeDetail ? 'Detalle incluido' : 'Sin detalle', active: config.includeDetail });
  badges.push({ label: config.includeCharts ? 'Gráficas activas' : 'Sin gráficas', active: config.includeCharts });
  badges.push({ label: config.includeMovements ? 'Movimientos incluidos' : 'Ocultar movimientos', active: config.includeMovements });
  badges.push({ label: config.includeObservations ? 'Observaciones visibles' : 'Sin observaciones', active: config.includeObservations });
  badges.push({ label: config.includeUniverse ? 'Bloque universo' : 'Ocultar universo', active: config.includeUniverse });

  const csvConfig = config.csv || {};
  badges.push({ label: csvConfig.includePoResumen ? 'CSV con resumen de PO' : 'CSV sin resumen de PO', active: csvConfig.includePoResumen });
  badges.push({ label: csvConfig.includeRemisiones ? 'CSV con remisiones' : 'CSV sin remisiones', active: csvConfig.includeRemisiones });
  badges.push({ label: csvConfig.includeFacturas ? 'CSV con facturas' : 'CSV sin facturas', active: csvConfig.includeFacturas });
  badges.push({ label: csvConfig.includeTotales ? 'CSV con totales' : 'CSV sin totales', active: csvConfig.includeTotales });
  badges.push({ label: csvConfig.includeUniverseInfo ? 'CSV con datos de universo' : 'CSV sin datos de universo', active: csvConfig.includeUniverseInfo });

  return badges
    .map(badge => `
      <span class="customization-badge ${badge.active ? 'on' : 'off'}">${escapeHtml(badge.label)}</span>
    `)
    .join('');
}

function renderCustomizationSummary() {
  const container = document.getElementById('customizationSummary');
  if (!container) return;
  container.innerHTML = buildCustomizationBadgesMarkup();
}

function applySavedCustomization() {
  const saved = readSession('porpt-customization', null);
  if (!saved) return;
  state.customization = mergeCustomizationSettings(state.customization, saved);
}

function setCustomizationControlStates() {
  const config = normalizeCustomizationState(state.customization);
  document.querySelectorAll('[data-customization-select]').forEach(select => {
    const key = select.getAttribute('data-customization-select');
    if (!key) return;
    const group = select.getAttribute('data-customization-group');
    const value = group === 'csv' ? config.csv?.[key] !== false : config[key] !== false;
    select.value = value ? 'true' : 'false';
  });
  document.querySelectorAll('input[data-customization-toggle]').forEach(input => {
    const key = input.getAttribute('data-customization-toggle');
    if (!key) return;
    const group = input.getAttribute('data-customization-group');
    const enabled = group === 'csv' ? config.csv?.[key] !== false : config[key] !== false;
    input.checked = enabled;
  });
}

function handleCustomizationSelectChange(event) {
  const select = event.target;
  if (!select || select.tagName !== 'SELECT') return;
  const key = select.getAttribute('data-customization-select');
  if (!key) return;
  const group = select.getAttribute('data-customization-group');
  const enabled = select.value !== 'false';
  if (group === 'csv') {
    state.customization = {
      ...state.customization,
      csv: {
        ...state.customization.csv,
        [key]: enabled
      }
    };
  } else {
    state.customization = {
      ...state.customization,
      [key]: enabled
    };
  }
  state.customization = normalizeCustomizationState(state.customization);
  saveSession('porpt-customization', state.customization);
  renderCustomizationSummary();
  updateReportOverviewMeta();
}

function setupCustomizationControls() {
  setCustomizationControlStates();
  document.querySelectorAll('[data-customization-select]').forEach(select => {
    select.removeEventListener('change', handleCustomizationSelectChange);
    select.addEventListener('change', handleCustomizationSelectChange);
  });
  document.querySelectorAll('input[data-customization-toggle]').forEach(input => {
    if (input.dataset.customizationBound === 'true') return;
    input.dataset.customizationBound = 'true';
    input.addEventListener('change', event => {
      const target = event.target;
      if (!target || target.tagName !== 'INPUT') return;
      const key = target.getAttribute('data-customization-toggle');
      if (!key) return;
      const group = target.getAttribute('data-customization-group');
      const enabled = target.checked;
      if (group === 'csv') {
        state.customization = {
          ...state.customization,
          csv: {
            ...state.customization.csv,
            [key]: enabled
          }
        };
      } else {
        state.customization = {
          ...state.customization,
          [key]: enabled
        };
      }
      state.customization = normalizeCustomizationState(state.customization);
      saveSession('porpt-customization', state.customization);
      setCustomizationControlStates();
      renderCustomizationSummary();
      updateReportOverviewMeta();
    });
  });
  setCustomizationControlStates();
}

function handleReportFileNameChange(event) {
  state.customReportName = event.target.value.trim();
  saveSession('porpt-custom-report-name', state.customReportName);
  updateReportOverviewMeta();
}

function handleUniverseFileNameChange(event) {
  state.universeCustomName = event.target.value.trim();
  saveSession('porpt-universe-report-name', state.universeCustomName);
}

function renderReportSelectionOverview() {
  const container = document.getElementById('reportSelectionOverview');
  if (!container) return;
  const empresaBadge = document.getElementById('overviewEmpresaBadge');
  if (empresaBadge) {
    empresaBadge.textContent = state.selectedEmpresa || 'Sin empresa';
  }
  const countBadge = document.getElementById('overviewPoCount');
  if (countBadge) {
    const count = state.selectedPoIds.length;
    countBadge.textContent = count === 1 ? '1 PO base' : `${count} POs base`;
  }
  const help = document.getElementById('reportDownloadHelp');
  if (state.selectedPoIds.length === 0 || !state.selectedEmpresa) {
    container.innerHTML = '<p class="text-muted mb-0">Selecciona una empresa y al menos una PO en la pestaña "Selección".</p>';
    
  }
  if (!state.summary) {
    container.innerHTML = '<p class="text-muted mb-0">Cuando el dashboard termine de cargar, verás aquí el resumen del reporte.</p>';
    return;
  }
  const items = Array.isArray(state.summary.items) ? state.summary.items : [];
  const groups = buildPoGroups(items, state.summary?.selectionDetails);
  if (!groups.length) {
    container.innerHTML = '<p class="text-muted mb-0">No se encontraron POs activas con la selección actual.</p>';
    if (help) {
      help.textContent = 'Verifica los filtros o selecciona otra PO.';
    }
    return;
  }
  const rows = groups
    .map(group => {
      const total = formatCurrency(group.totals.total);
      const restante = formatCurrency(group.totals.restante);
      const extensionCount = Array.isArray(group.extensionIds) ? group.extensionIds.length : 0;
      const variantLabel = extensionCount === 0
        ? 'Solo base'
        : `Conjunto (${group.ids.length} variantes)`;
      const badgeListMarkup = buildGroupBadgeList(group.baseId, group.extensionIds);
      const compositionNote = buildGroupCompositionNote(group.extensionIds);
      return `
        <li class="list-group-item d-flex flex-wrap justify-content-between gap-2">
          <div class="flex-grow-1">
            <div class="d-flex flex-wrap align-items-center gap-2">
              <span class="fw-semibold">PO ${escapeHtml(group.baseId)}</span>
              <span class="badge text-bg-light text-dark">${escapeHtml(variantLabel)}</span>
            </div>
            <div class="extension-group-summary mt-1">${badgeListMarkup}</div>
            <p class="extension-group-note mb-0">${escapeHtml(compositionNote)}</p>
          </div>
          <div class="text-end">
            <span class="badge text-bg-primary-subtle text-primary-emphasis me-2">Total $${total}</span>
            <span class="badge text-bg-success-subtle text-success-emphasis">Restante $${restante}</span>
          </div>
        </li>
      `;
    })
    .join('');
  container.innerHTML = `
    <ul class="list-group list-group-flush rounded-4 border">
      ${rows}
    </ul>
  `;
}

function updateReportOverviewMeta() {
  const engineLabel = document.getElementById('overviewEngineLabel');
  const formatLabel = document.getElementById('overviewFormatLabel');
  if (engineLabel) {
    const engine = state.reportEngines.find(item => item.id === state.selectedEngine);
    engineLabel.textContent = engine ? engine.label : 'Sin seleccionar';
  }
  if (formatLabel) {
    formatLabel.textContent = getFormatLabel(state.selectedFormat);
  }
  const filenameLabel = document.getElementById('overviewFilename');
  if (filenameLabel) {
    const custom = state.customReportName?.trim();
    filenameLabel.textContent = custom ? `${custom}.${state.selectedFormat}` : 'Se generará automáticamente';
  }
}

function showWizardTab(target) {
  if (!target) return;
  const normalized = target.startsWith('#') ? target : `#${target}`;
  const trigger = document.querySelector(`[data-bs-target="${normalized}"]`);
  if (!trigger) return;
  const tabInstance = bootstrap.Tab.getOrCreateInstance(trigger);
  tabInstance.show();
}

function setupWizardNavigation() {
  document.querySelectorAll('[data-wizard-action]').forEach(button => {
    if (!button || button.dataset.wizardBound === 'true') {
      return;
    }
    button.dataset.wizardBound = 'true';
    button.addEventListener('click', event => {
      event.preventDefault();
      const explicitTarget = button.getAttribute('data-wizard-target');
      if (explicitTarget) {
        showWizardTab(explicitTarget);
        return;
      }
      const action = button.getAttribute('data-wizard-action');
      if (action === 'next') {
        showWizardTab('#panel-personal');
      } else if (action === 'prev') {
        showWizardTab('#panel-global');
      }
    });
  });
}

function isSummarySynced() {
  if (!state.summary) return false;
  const selected = [...state.selectedPoIds].sort();
  const details = Array.isArray(state.summary.selectionDetails)
    ? state.summary.selectionDetails.reduce((map, entry) => {
      if (entry?.baseId) {
        map.set(entry.baseId, new Set((entry.variants || []).map(id => id.trim())));
      }
      return map;
    }, new Map())
    : null;
  const summaryBases = details
    ? Array.from(details.keys()).sort()
    : Array.isArray(state.summary.selectedIds)
      ? [...state.summary.selectedIds].sort()
      : [];
  if (summaryBases.length !== selected.length) return false;
  if (!selected.every((value, index) => value === summaryBases[index])) {
    return false;
  }
  if (!details) {
    return true;
  }
  return selected.every(baseId => {
    const summaryVariants = details.get(baseId) || new Set();
    const selectedVariants = new Set(getSelectedVariantIds(baseId).map(id => id.trim()));
    if (summaryVariants.size !== selectedVariants.size) {
      return false;
    }
    for (const id of selectedVariants) {
      if (!summaryVariants.has(id)) {
        return false;
      }
    }
    return true;
  });
}

async function updateSummary(options = {}) {
  const silent = options.silent === true;
  if (!state.selectedEmpresa || state.selectedPoIds.length === 0) return;
  try {
    const poTargets = state.selectedPoIds.map(baseId => ({
      baseId,
      ids: getSelectedVariantIds(baseId)
    }));
    const response = await fetch('/po-summary', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        empresa: state.selectedEmpresa,
        poIds: state.selectedPoIds,
        poTargets
      })
    });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener el resumen');
    }
    state.summary = data.summary;
    const selectionAdjusted = syncSelectedVariantsFromSummary(state.summary);
    renderSummaryCards(state.summary);
    renderTable(state.summary);
    attachTableHandlers(state.summary);
    renderCharts(state.summary);
    renderAlerts(state.summary.alerts || []);
    if (selectionAdjusted) {
      renderSelectedPoChips();
    }
    if (!silent) {
      showAlert('Dashboard actualizado con la selección de POs.', 'success');
    }
    showDashboardPanel({ autoOpen: true });
    renderReportSelectionOverview();
  } catch (error) {
    console.error('Error actualizando resumen:', error);
    showAlert(error.message || 'No se pudo actualizar el dashboard', 'danger');
  }
}

async function generateReport() {
  if (state.isGeneratingReport) {
    return false;
  }
  if (!state.selectedEmpresa || state.selectedPoIds.length === 0) {
    showAlert('Selecciona primero una empresa y al menos una PO base.', 'warning');
    return false;
  }
  const formatSelect = document.getElementById('reportFormatSelect');
  if (formatSelect) {
    state.selectedFormat = formatSelect.value;
  }
  const filenameInput = document.getElementById('reportFileNameInput');
  if (filenameInput) {
    state.customReportName = filenameInput.value.trim();
    saveSession('porpt-custom-report-name', state.customReportName);
  }
  const engine = state.selectedEngine || state.reportSettings?.defaultEngine || 'simple-pdf';
  const format = state.selectedFormat || state.reportSettings?.export?.defaultFormat || 'pdf';
  const customization = normalizeCustomizationState(state.customization);
  const poTargets = state.selectedPoIds.map(baseId => ({
    baseId,
    ids: getSelectedVariantIds(baseId)
  }));
  state.isGeneratingReport = true;
  try {
    const response = await fetch('/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        empresa: state.selectedEmpresa,
        poIds: state.selectedPoIds,
        poTargets,
        engine,
        format,
        customization
      })
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: 'Error generando reporte' }));
      throw new Error(error.message || 'Error generando reporte');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const defaultFilename = state.selectedPoIds.length === 1
      ? `POrpt_${state.selectedPoIds[0]}`
      : `POrpt_${state.selectedPoIds.length}_POs`;
    const contentType = response.headers.get('Content-Type') || '';
    let extension = 'pdf';
    if (contentType.includes('csv')) {
      extension = 'csv';
    } else if (contentType.includes('json')) {
      extension = 'json';
    } else if (contentType.includes('pdf')) {
      extension = 'pdf';
    } else if (format) {
      extension = format;
    }
    const filenameBase = sanitizeFileName(state.customReportName, defaultFilename);
    link.download = `${filenameBase}.${extension}`;
    link.click();
    URL.revokeObjectURL(url);
    const selectedEngine = state.reportEngines.find(item => item.id === engine);
    const engineLabel = selectedEngine ? selectedEngine.label : 'Reporte';
    const formatLabel = getFormatLabel(format);
    showAlert(`${engineLabel} (${formatLabel}) generado correctamente.`, 'success');
    return true;
  } catch (error) {
    console.error('Error generando reporte:', error);
    showAlert(error.message || 'No se pudo generar el reporte.', 'danger');
    return false;
  } finally {
    state.isGeneratingReport = false;
  }
}

function populateReportPreview() {
  const summary = state.summary;
  if (!summary) return;
  const empresaElement = document.getElementById('previewEmpresa');
  if (empresaElement) {
    empresaElement.textContent = summary.companyName || summary.empresaLabel || state.selectedEmpresa || '-';
  }
  const engineElement = document.getElementById('previewEngine');
  if (engineElement) {
    const engine = state.reportEngines.find(item => item.id === state.selectedEngine);
    engineElement.textContent = engine ? engine.label : 'Sin seleccionar';
  }
  const formatElement = document.getElementById('previewFormat');
  if (formatElement) {
    formatElement.textContent = getFormatLabel(state.selectedFormat);
  }
  const filenameElement = document.getElementById('previewFileName');
  if (filenameElement) {
    const defaultBase = summary.selectedIds && summary.selectedIds.length
      ? summary.selectedIds.join('_')
      : summary.baseId || 'reporte';
    const filenameBase = sanitizeFileName(state.customReportName, defaultBase);
    filenameElement.textContent = `${filenameBase}.${state.selectedFormat}`;
  }
  const poList = document.getElementById('previewPoList');
  if (poList) {
    const items = Array.isArray(summary.items) ? summary.items : [];
    if (items.length === 0) {
      poList.innerHTML = '<li class="list-group-item text-muted">Sin POs detalladas para esta selección.</li>';
    } else {
      poList.innerHTML = items
        .map(item => {
          const total = formatCurrency(item.total || 0);
          const restante = formatCurrency(item.totals?.restante ?? 0);
          return `
            <li class="list-group-item d-flex justify-content-between gap-2">
              <div>
                <span class="fw-semibold">${escapeHtml(item.id || '')}</span>
                <span class="text-muted">• ${escapeHtml(item.fecha || '')}</span>
              </div>
              <div class="text-end">
                <span class="badge text-bg-primary-subtle text-primary-emphasis me-2">$${total}</span>
                <span class="badge text-bg-success-subtle text-success-emphasis">Restante $${restante}</span>
              </div>
            </li>
          `;
        })
        .join('');
    }
  }
  const totalsList = document.getElementById('previewTotals');
  if (totalsList) {
    const totals = summary.totals || {};
    totalsList.innerHTML = `
      <li class="list-group-item d-flex justify-content-between">
        <span>Total autorizado</span>
        <span class="fw-semibold">$${formatCurrency(totals.total || 0)}</span>
      </li>
      <li class="list-group-item d-flex justify-content-between">
        <span>Remisiones</span>
        <span class="fw-semibold">$${formatCurrency(totals.totalRem || 0)}</span>
      </li>
      <li class="list-group-item d-flex justify-content-between">
        <span>Facturas</span>
        <span class="fw-semibold">$${formatCurrency(totals.totalFac || 0)}</span>
      </li>
      <li class="list-group-item d-flex justify-content-between">
        <span>Disponible</span>
        <span class="fw-semibold">$${formatCurrency(totals.restante || 0)}</span>
      </li>
    `;
  }
  const customizationContainer = document.getElementById('previewCustomization');
  if (customizationContainer) {
    customizationContainer.innerHTML = buildCustomizationBadgesMarkup();
  }
  const alertsList = document.getElementById('previewAlerts');
  if (alertsList) {
    const alerts = Array.isArray(summary.alerts) ? summary.alerts : [];
    if (alerts.length === 0) {
      alertsList.innerHTML = '<li class="list-group-item text-muted">Sin alertas registradas.</li>';
    } else {
      const topAlerts = alerts.slice(0, 3);
      alertsList.innerHTML = topAlerts
        .map(alert => `
          <li class="list-group-item d-flex justify-content-between align-items-center gap-2">
            <span>${escapeHtml(alert.message)}</span>
            <span class="badge text-bg-${mapAlertType(alert.type || 'info')}">${escapeHtml((alert.type || 'info').toUpperCase())}</span>
          </li>
        `)
        .join('');
      if (alerts.length > 3) {
        alertsList.innerHTML += `
          <li class="list-group-item text-muted">${alerts.length - 3} alerta(s) adicional(es)...</li>
        `;
      }
    }
  }
}

async function openReportPreview(event) {
  event?.preventDefault();
  if (!state.selectedEmpresa || state.selectedPoIds.length === 0) {
    showAlert('Selecciona primero una empresa y al menos una PO base.', 'warning');
    return;
  }
  const formatSelect = document.getElementById('reportFormatSelect');
  if (formatSelect) {
    state.selectedFormat = formatSelect.value;
  }
  const filenameInput = document.getElementById('reportFileNameInput');
  if (filenameInput) {
    state.customReportName = filenameInput.value.trim();
    saveSession('porpt-custom-report-name', state.customReportName);
  }
  if (!isSummarySynced()) {
    await updateSummary({ silent: true });
  }
  if (!state.summary) {
    showAlert('No fue posible preparar el resumen para la descarga.', 'danger');
    return;
  }
  updateReportOverviewMeta();
  renderCustomizationSummary();
  renderReportSelectionOverview();
  populateReportPreview();
  const modalElement = document.getElementById('reportPreviewModal');
  if (!modalElement) {
    await generateReport();
    return;
  }
  const modal = bootstrap.Modal.getOrCreateInstance(modalElement);
  modal.show();
}

async function confirmReportDownload(event) {
  event?.preventDefault();
  const button = document.getElementById('confirmReportDownloadBtn');
  const modalElement = document.getElementById('reportPreviewModal');
  const modal = modalElement ? bootstrap.Modal.getInstance(modalElement) : null;
  const originalContent = button ? button.innerHTML : '';
  if (button) {
    button.disabled = true;
    button.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Generando...';
  }
  const success = await generateReport();
  if (button) {
    button.disabled = false;
    button.innerHTML = originalContent || 'Descargar ahora';
  }
  if (success && modal) {
    modal.hide();
  }
}

function getUniverseFilterLabel(mode, start, end) {
  if (mode === 'range') {
    return `Del ${start} al ${end}`;
  }
  if (mode === 'single') {
    return `Fecha ${start}`;
  }
  return 'Global (todas las fechas)';
}

function updateUniverseFilterVisibility() {
  const mode = state.universeFilterMode;
  const startGroup = document.getElementById('universeStartDateGroup');
  const endGroup = document.getElementById('universeEndDateGroup');
  const startLabel = document.getElementById('universeStartDateLabel');
  if (startGroup) {
    startGroup.classList.toggle('d-none', mode === 'global');
  }
  if (endGroup) {
    endGroup.classList.toggle('d-none', mode !== 'range');
  }
  if (startLabel) {
    startLabel.textContent = mode === 'single' ? 'Fecha específica' : 'Fecha inicial';
  }
  if (mode === 'global') {
    state.universeStartDate = '';
    state.universeEndDate = '';
    const startInput = document.getElementById('universeStartDate');
    const endInput = document.getElementById('universeEndDate');
    if (startInput) startInput.value = '';
    if (endInput) endInput.value = '';
  }
}

function updateUniverseControls() {
  const button = document.getElementById('generateUniverseReportBtn');
  if (button) {
    button.disabled = !state.selectedEmpresa;
  }
  const empresaLabel = document.getElementById('universeSelectedEmpresa');
  if (empresaLabel) {
    empresaLabel.textContent = state.selectedEmpresa || 'Selecciona una empresa';
  }
  const companyLabel = document.getElementById('universeCompanyLabel');
  if (companyLabel) {
    companyLabel.textContent = state.reportSettings?.branding?.companyName || 'SSITEL';
  }
  const formatSelect = document.getElementById('universeFormat');
  if (formatSelect) {
    formatSelect.value = state.universeFormat || 'pdf';
  }
  const filenameInput = document.getElementById('universeFileName');
  if (filenameInput) {
    filenameInput.value = state.universeCustomName || '';
  }
}

function handleUniverseModeChange(mode) {
  state.universeFilterMode = mode || 'global';
  if (state.universeFilterMode !== 'range') {
    state.universeEndDate = '';
    const endInput = document.getElementById('universeEndDate');
    if (endInput && state.universeFilterMode !== 'range') {
      endInput.value = '';
    }
  }
  updateUniverseFilterVisibility();
}

async function generateUniverseReport(event) {
  event?.preventDefault();
  if (!state.selectedEmpresa) {
    showAlert('Selecciona primero una empresa para generar el reporte del universo.', 'warning');
    return;
  }
  const mode = state.universeFilterMode;
  const startInput = document.getElementById('universeStartDate');
  const endInput = document.getElementById('universeEndDate');
  const startDate = startInput?.value || state.universeStartDate || '';
  const endDate = endInput?.value || state.universeEndDate || '';
  state.universeStartDate = startDate;
  state.universeEndDate = endDate;
  const formatSelect = document.getElementById('universeFormat');
  if (formatSelect) {
    state.universeFormat = formatSelect.value;
    saveSession('porpt-universe-format', state.universeFormat);
  }
  const filenameInput = document.getElementById('universeFileName');
  if (filenameInput) {
    state.universeCustomName = filenameInput.value.trim();
    saveSession('porpt-universe-report-name', state.universeCustomName);
  }

  const payload = {
    empresa: state.selectedEmpresa,
    filter: { mode },
    format: state.universeFormat || 'pdf',
    customization: state.customization
  };
  if (mode === 'range') {
    if (!startDate || !endDate) {
      showAlert('Selecciona las fechas de inicio y fin para el rango.', 'warning');
      return;
    }
    if (startDate > endDate) {
      showAlert('La fecha inicial no puede ser posterior a la fecha final.', 'warning');
      return;
    }
    payload.filter.startDate = startDate;
    payload.filter.endDate = endDate;
  } else if (mode === 'single') {
    if (!startDate) {
      showAlert('Selecciona la fecha específica para el reporte unitario.', 'warning');
      return;
    }
    payload.filter.startDate = startDate;
  }

  const button = document.getElementById('generateUniverseReportBtn');
  if (button) {
    button.disabled = true;
  }
  try {
    const response = await fetch('/report-universe', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: 'Error generando reporte' }));
      throw new Error(error.message || 'Error generando reporte');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const suffix = mode === 'range'
      ? `${startDate}_a_${endDate}`
      : mode === 'single'
        ? startDate
        : 'global';
    const contentType = response.headers.get('Content-Type') || '';
    let extension = 'pdf';
    if (contentType.includes('csv')) {
      extension = 'csv';
    } else if (contentType.includes('json')) {
      extension = 'json';
    } else if (payload.format && payload.format !== 'pdf') {
      extension = payload.format;
    }
    const baseName = sanitizeFileName(state.universeCustomName, `POrpt_universo_${suffix || 'global'}`);
    link.download = `${baseName}.${extension}`;
    link.click();
    URL.revokeObjectURL(url);
    const label = getUniverseFilterLabel(mode, startDate || '-', endDate || '-');
    const formatLabel = getFormatLabel(payload.format);
    showAlert(`Reporte universo (${label}) generado en ${formatLabel}.`, 'success');
  } catch (error) {
    console.error('Error generando reporte universo:', error);
    showAlert(error.message || 'No se pudo generar el reporte del universo.', 'danger');
  } finally {
    if (button) {
      button.disabled = !state.selectedEmpresa;
    }
  }
}

function updateReportEngineStatus() {
  const badge = document.getElementById('reportEngineStatus');
  if (!badge) return;
  const engineId = state.selectedEngine || state.reportSettings?.defaultEngine;
  const engine = state.reportEngines.find(item => item.id === engineId);
  if (!engine) {
    badge.classList.add('d-none');
    return;
  }
  let text = '';
  let badgeClass = 'text-bg-secondary';
  if (engine.available) {
    text = 'PDF directo listo';
    badgeClass = 'text-bg-success';
  } else {
    text = 'Instala pdfkit para habilitar el PDF';
    badgeClass = 'text-bg-warning text-dark';
  }
  badge.textContent = text;
  badge.className = `badge ${badgeClass}`;
  badge.classList.remove('d-none');
}

function renderReportEngineSelector() {
  const menu = document.getElementById('reportEngineMenu');
  const label = document.getElementById('reportEngineLabel');
  if (!menu) return;

  menu.innerHTML = '';
  if (!Array.isArray(state.reportEngines) || state.reportEngines.length === 0) {
    if (label) label.textContent = '-';
    updateReportEngineStatus();
    return;
  }

  if (!state.selectedEngine || !state.reportEngines.some(engine => engine.id === state.selectedEngine && engine.available)) {
    const availableEngine = state.reportEngines.find(engine => engine.available);
    state.selectedEngine = availableEngine?.id || state.reportSettings?.defaultEngine || state.reportEngines[0].id;
  }

  state.reportEngines.forEach(engine => {
    const item = document.createElement('li');
    item.innerHTML = `
      <button type="button" class="dropdown-item ${state.selectedEngine === engine.id ? 'active' : ''}" data-engine="${engine.id}" ${engine.available ? '' : 'disabled'}>
        ${engine.label}
        <small>${engine.description}</small>
      </button>
    `;
    menu.appendChild(item);
  });

  const activeEngine = state.reportEngines.find(engine => engine.id === state.selectedEngine);
  if (label && activeEngine) {
    label.textContent = activeEngine.label;
  }
  updateReportEngineStatus();
  updateReportOverviewMeta();
}

function renderReportFormatSelect() {
  const select = document.getElementById('reportFormatSelect');
  if (!select) return;
  const availableFormats = (Array.isArray(state.reportSettings?.export?.availableFormats)
    ? state.reportSettings.export.availableFormats
    : ['pdf']).filter(format => {
    if (!state.reportFormatsCatalog || state.reportFormatsCatalog.length === 0) return true;
    return state.reportFormatsCatalog.includes(format);
  });
  if (availableFormats.length === 0) {
    availableFormats.push('pdf');
  }
  select.innerHTML = availableFormats
    .map(format => `<option value="${format}">${getFormatLabel(format)}</option>`)
    .join('');
  const defaultFormat = state.reportSettings?.export?.defaultFormat;
  const selected = availableFormats.includes(state.selectedFormat)
    ? state.selectedFormat
    : availableFormats.includes(defaultFormat)
      ? defaultFormat
      : availableFormats[0];
  select.value = selected;
  state.selectedFormat = selected;
  updateReportOverviewMeta();
}

function setCompaniesFieldDisabled(disabled) {
  const picker = document.getElementById('userCompanyPicker');
  if (picker) {
    picker.disabled = disabled;
    if (disabled) {
      picker.classList.remove('is-invalid');
    }
  }
  const checklist = document.getElementById('userCompanyChecklist');
  if (checklist) {
    checklist.querySelectorAll('input[type="checkbox"]').forEach(input => {
      input.disabled = disabled;
    });
    if (disabled) {
      checklist.classList.remove('is-invalid');
    }
  }
}

function setUserCompaniesValidity(isValid) {
  const picker = document.getElementById('userCompanyPicker');
  const checklist = document.getElementById('userCompanyChecklist');
  if (picker) {
    picker.classList.toggle('is-invalid', !isValid);
  }
  if (checklist) {
    checklist.classList.toggle('is-invalid', !isValid);
  }
}

function updateUserCompaniesValidity() {
  const allCompanies = document.getElementById('userAllCompanies')?.checked;
  const selectedCount = state.userCompanySelection instanceof Set ? state.userCompanySelection.size : 0;
  const isValid = allCompanies || selectedCount > 0;
  setUserCompaniesValidity(isValid);
}

function updateUserCompanyPickerOptions() {
  const picker = document.getElementById('userCompanyPicker');
  if (!picker) return;
  const selected = state.userCompanySelection instanceof Set ? state.userCompanySelection : new Set();
  const currentValue = picker.value;
  picker.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Selecciona una empresa';
  picker.appendChild(placeholder);
  state.empresas
    .filter(empresa => !selected.has(empresa))
    .forEach(empresa => {
      const option = document.createElement('option');
      option.value = empresa;
      option.textContent = empresa;
      picker.appendChild(option);
    });
  const hasCurrent = Array.from(picker.options).some(option => option.value === currentValue);
  picker.value = hasCurrent ? currentValue : '';
}

function renderUserCompanyChecklist() {
  const container = document.getElementById('userCompanyChecklist');
  if (!container) return;
  container.innerHTML = '';
  const selected = state.userCompanySelection instanceof Set
    ? Array.from(state.userCompanySelection)
    : [];
  if (!selected.length) {
    const empty = document.createElement('p');
    empty.className = 'text-muted small mb-0';
    empty.textContent = 'No hay empresas seleccionadas.';
    container.appendChild(empty);
    return;
  }
  selected
    .slice()
    .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }))
    .forEach(company => {
      const wrapper = document.createElement('div');
      wrapper.className = 'form-check';
      const id = `userCompany-${sanitizeVariantKey(company)}`;
      const input = document.createElement('input');
      input.className = 'form-check-input';
      input.type = 'checkbox';
      input.id = id;
      input.checked = true;
      input.dataset.company = company;
      const label = document.createElement('label');
      label.className = 'form-check-label';
      label.setAttribute('for', id);
      label.textContent = company;
      wrapper.appendChild(input);
      wrapper.appendChild(label);
      container.appendChild(wrapper);
    });
}

function resetUserCompanySelection(initialCompanies = [], { skipValidityUpdate = false } = {}) {
  const entries = Array.isArray(initialCompanies) ? initialCompanies.filter(Boolean) : [];
  state.userCompanySelection = new Set(entries);
  renderUserCompanyChecklist();
  updateUserCompanyPickerOptions();
  if (!skipValidityUpdate) {
    updateUserCompaniesValidity();
  }
}

function handleUserCompanyPickerChange(event) {
  const value = event.target.value;
  if (!value) return;
  if (!(state.userCompanySelection instanceof Set)) {
    state.userCompanySelection = new Set();
  }
  state.userCompanySelection.add(value);
  event.target.value = '';
  renderUserCompanyChecklist();
  updateUserCompanyPickerOptions();
  updateUserCompaniesValidity();
}

function handleUserCompanyChecklistChange(event) {
  const input = event.target;
  if (!input || input.type !== 'checkbox') return;
  const company = input.dataset.company;
  if (!company) return;
  if (!(state.userCompanySelection instanceof Set)) {
    state.userCompanySelection = new Set();
  }
  if (!input.checked) {
    state.userCompanySelection.delete(company);
    renderUserCompanyChecklist();
    updateUserCompanyPickerOptions();
    updateUserCompaniesValidity();
  }
}

function toggleBrandingLetterheadFields(disabled) {
  ['brandingLetterheadTop', 'brandingLetterheadBottom'].forEach(id => {
    const input = document.getElementById(id);
    if (input) {
      input.disabled = disabled;
      if (disabled) {
        input.classList.remove('is-invalid');
      }
    }
  });
}

function parseDialogFilters(filtersAttr) {
  if (!filtersAttr) return undefined;
  try {
    const parsed = JSON.parse(filtersAttr);
    return Array.isArray(parsed) ? parsed : undefined;
  } catch (error) {
    console.warn('No se pudieron interpretar los filtros del selector de archivos:', error);
    return undefined;
  }
}

function setupPathPickers(root = document) {
  if (!root) return;
  const buttons = root.querySelectorAll('[data-select-path]');
  buttons.forEach(button => {
    button.addEventListener('click', async () => {
      if (!window.electronAPI || typeof window.electronAPI.selectPath !== 'function') {
        showAlert('El selector nativo no está disponible en esta versión. Ingresa la ruta manualmente.', 'info');
        return;
      }
      const targetId = button.getAttribute('data-select-target');
      if (!targetId) return;
      const input = document.getElementById(targetId);
      const type = button.getAttribute('data-select-type') || 'file';
      const filters = parseDialogFilters(button.getAttribute('data-select-filters'));
      try {
        const selectedPath = await window.electronAPI.selectPath({
          type,
          defaultPath: input?.value || undefined,
          filters
        });
        if (selectedPath && input) {
          input.value = selectedPath;
          input.dispatchEvent(new Event('change', { bubbles: true }));
        }
      } catch (error) {
        console.error('Error al seleccionar ruta:', error);
        showAlert('No se pudo abrir el selector de archivos.', 'danger');
      }
    });
  });
}

function renderReportSettingsForm() {
  const form = document.getElementById('reportSettingsForm');
  if (!form || !state.reportSettings) return;
  const exportCfg = state.reportSettings.export || {};
  const formatCatalog = Array.isArray(state.reportFormatsCatalog) && state.reportFormatsCatalog.length
    ? state.reportFormatsCatalog
    : Array.from(new Set(['pdf', ...(Array.isArray(exportCfg.availableFormats) ? exportCfg.availableFormats : [])]));
  const availableFormats = (Array.isArray(exportCfg.availableFormats) && exportCfg.availableFormats.length
    ? exportCfg.availableFormats
    : ['pdf']).filter(format => formatCatalog.includes(format));
  const defaultFormat = formatCatalog.includes(exportCfg.defaultFormat)
    ? exportCfg.defaultFormat
    : availableFormats[0] || formatCatalog[0] || 'pdf';
  const defaultFormatSelect = document.getElementById('exportDefaultFormat');
  if (defaultFormatSelect) {
    defaultFormatSelect.innerHTML = formatCatalog
      .map(format => `<option value="${format}">${getFormatLabel(format)}</option>`)
      .join('');
    defaultFormatSelect.value = defaultFormat;
  }
  const formatsContainer = document.getElementById('exportFormatsContainer');
  if (formatsContainer) {
    formatsContainer.innerHTML = formatCatalog
      .map(format => {
        const id = `exportFormat-${format}`;
        const checked = format === 'pdf' ? true : availableFormats.includes(format);
        const disabledAttr = format === 'pdf' ? ' disabled' : '';
        const description = getFormatDescription(format);
        return `
          <div class="form-check">
            <input class="form-check-input" type="checkbox" id="${id}" value="${format}" data-format-option ${checked ? 'checked' : ''}${disabledAttr}>
            <label class="form-check-label" for="${id}">${getFormatLabel(format)}${description ? `<span class="d-block small text-muted">${description}</span>` : ''}</label>
          </div>
        `;
      })
      .join('');
  }
  if (!availableFormats.includes(state.selectedFormat)) {
    state.selectedFormat = defaultFormat;
  }
  state.reportFormatsCatalog = formatCatalog;
  const customization = state.reportSettings.customization || {};
  state.customization = mergeCustomizationSettings(customization);
  const customizationMappings = [
    ['customIncludeCharts', state.customization.includeCharts !== false],
    ['customIncludeMovements', state.customization.includeMovements !== false],
    ['customIncludeObservations', state.customization.includeObservations !== false],
    ['customIncludeUniverse', state.customization.includeUniverse !== false]
  ];
  customizationMappings.forEach(([id, value]) => {
    const input = document.getElementById(id);
    if (input) {
      input.value = value ? 'true' : 'false';
    }
  });
  const csvCustomization = state.customization.csv || {};
  const csvMappings = [
    ['customCsvIncludePoResumen', csvCustomization.includePoResumen !== false],
    ['customCsvIncludeRemisiones', csvCustomization.includeRemisiones !== false],
    ['customCsvIncludeFacturas', csvCustomization.includeFacturas !== false],
    ['customCsvIncludeTotales', csvCustomization.includeTotales !== false],
    ['customCsvIncludeUniverseInfo', csvCustomization.includeUniverseInfo !== false]
  ];
  csvMappings.forEach(([id, value]) => {
    const input = document.getElementById(id);
    if (input) {
      input.value = value ? 'true' : 'false';
    }
  });
  const branding = state.reportSettings.branding || {};
  const brandingMappings = [
    ['brandingHeaderTitle', branding.headerTitle || ''],
    ['brandingHeaderSubtitle', branding.headerSubtitle || ''],
    ['brandingCompanyName', branding.companyName || ''],
    ['brandingFooterText', branding.footerText || ''],
    ['brandingLetterheadTop', branding.letterheadTop || ''],
    ['brandingLetterheadBottom', branding.letterheadBottom || ''],
    ['brandingRemColor', branding.remColor || '#2563eb'],
    ['brandingFacColor', branding.facColor || '#dc2626'],
    ['brandingRestanteColor', branding.restanteColor || '#16a34a'],
    ['brandingAccentColor', branding.accentColor || '#1f2937']
  ];
  brandingMappings.forEach(([id, value]) => {
    const input = document.getElementById(id);
    if (input) {
      input.value = value;
    }
  });
  const letterheadToggle = document.getElementById('brandingLetterheadEnabled');
  if (letterheadToggle) {
    letterheadToggle.checked = branding.letterheadEnabled === true;
  }
  toggleBrandingLetterheadFields(!(branding.letterheadEnabled === true));
  renderReportFormatSelect();
}

async function loadReportSettings(options = {}) {
  const requireAdmin = options.requireAdmin || false;
  try {
    const response = await (requireAdmin ? adminFetch('/report-settings') : fetch('/report-settings'));
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener la configuración');
    }
    state.reportSettings = data.settings;
    state.reportEngines = Array.isArray(data.engines) ? data.engines : [];
    state.reportFormatsCatalog = Array.isArray(data.formats) ? data.formats : state.reportFormatsCatalog;
    if (state.reportSettings?.export?.defaultFormat) {
      state.selectedFormat = state.reportSettings.export.defaultFormat;
    }
    if (!state.selectedEngine || !state.reportEngines.some(engine => engine.id === state.selectedEngine)) {
      state.selectedEngine = state.reportSettings?.defaultEngine || state.reportEngines[0]?.id;
    }
    renderReportEngineSelector();
    renderReportFormatSelect();
    updateUniverseControls();
    if (requireAdmin) {
      renderReportSettingsForm();
    }
  } catch (error) {
    console.error('Error cargando configuración de reportes:', error);
    showAlert(error.message || 'No se pudo cargar la configuración de reportes.', 'danger');
  }
}

async function loadUsers() {
  try {
    const response = await adminFetch('/users');
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener usuarios');
    }
    state.users = Array.isArray(data.users) ? data.users : [];
    renderUsers(state.users);
    showAlert('Usuarios sincronizados.', 'success');
  } catch (error) {
    console.error('Error cargando usuarios:', error);
    showAlert(error.message || 'No se pudieron cargar los usuarios.', 'danger');
  }
}

function getInitials(text) {
  if (!text) return '?';
  const parts = text.trim().split(/\s+/u).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
}

function formatCompaniesForDisplay(user) {
  if (user.empresas === '*') {
    return '<span class="badge text-bg-primary">Todas</span>';
  }
  const companies = Array.isArray(user.empresas) ? user.empresas : [];
  if (companies.length === 0) {
    return '<span class="text-muted small">Sin asignar</span>';
  }
  return companies
    .map(empresa => `<span class="badge text-bg-light text-dark me-1 mb-1">${escapeHtml(empresa)}</span>`)
    .join('');
}

function renderUsers(users) {
  const tbody = document.querySelector('#usersTable tbody');
  const emptyState = document.getElementById('usersEmptyState');
  if (!tbody) return;

  if (!Array.isArray(users) || users.length === 0) {
    tbody.innerHTML = '';
    emptyState?.classList.remove('d-none');
    return;
  }

  emptyState?.classList.add('d-none');
  tbody.innerHTML = users
    .map(user => {
      const roleBadge = user.empresas === '*'
        ? '<span class="badge text-bg-primary">Administrador</span>'
        : '<span class="badge text-bg-info text-dark">Colaborador</span>';
      return `
        <tr data-id="${user.id}">
          <td>
            <div class="d-flex align-items-center gap-2">
              <div class="avatar-initial">${getInitials(user.nombre || user.usuario)}</div>
              <div>
                <div class="fw-semibold">@${escapeHtml(user.usuario)}</div>
                <div class="text-muted small">ID ${user.id}</div>
              </div>
            </div>
          </td>
          <td>${user.nombre ? escapeHtml(user.nombre) : '<span class="text-muted">Sin nombre</span>'}</td>
          <td><div class="d-flex flex-wrap gap-1">${formatCompaniesForDisplay(user)}</div></td>
          <td>${roleBadge}</td>
          <td class="text-end">
            <button class="btn btn-sm btn-outline-secondary me-1" data-action="edit" data-id="${user.id}">Editar</button>
            ${user.usuario === 'admin'
              ? ''
              : `<button class="btn btn-sm btn-outline-danger" data-action="delete" data-id="${user.id}">Eliminar</button>`}
          </td>
        </tr>
      `;
    })
    .join('');
}

function resetUserModal() {
  const form = document.getElementById('userForm');
  form?.reset();
  ['userUsername', 'userPassword', 'userPasswordConfirm', 'userCompanyPicker'].forEach(id => {
    const input = document.getElementById(id);
    input?.classList.remove('is-invalid');
  });
  const usernameInput = document.getElementById('userUsername');
  if (usernameInput) {
    usernameInput.disabled = false;
  }
  state.editingUserId = null;
  resetUserCompanySelection([], { skipValidityUpdate: true });
  setCompaniesFieldDisabled(false);
}

function openUserModal(user = null) {
  resetUserModal();
  const modalTitle = document.getElementById('userModalTitle');
  const usernameInput = document.getElementById('userUsername');
  const nameInput = document.getElementById('userName');
  const allCompaniesInput = document.getElementById('userAllCompanies');
  const submitButton = document.getElementById('userModalSubmit');

  if (user) {
    state.editingUserId = user.id;
    modalTitle.textContent = 'Editar usuario';
    usernameInput.value = user.usuario;
    usernameInput.disabled = user.usuario === 'admin';
    nameInput.value = user.nombre || '';
    allCompaniesInput.checked = user.empresas === '*';
    if (user.empresas !== '*') {
      const selectedCompanies = Array.isArray(user.empresas) ? user.empresas : [];
      resetUserCompanySelection(selectedCompanies);
    }
    submitButton.textContent = 'Actualizar';
  } else {
    modalTitle.textContent = 'Nuevo usuario';
    usernameInput.disabled = false;
    submitButton.textContent = 'Crear usuario';
  }

  if (!user || user.empresas === '*') {
    resetUserCompanySelection([], { skipValidityUpdate: true });
  }

  setCompaniesFieldDisabled(allCompaniesInput.checked);
  if (allCompaniesInput.checked) {
    setUserCompaniesValidity(true);
  } else if (user && user.empresas !== '*') {
    updateUserCompaniesValidity();
  } else {
    setUserCompaniesValidity(true);
  }
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('userModal'));
  modal.show();
}

function handleUsersTableClick(event) {
  const button = event.target.closest('button[data-action][data-id]');
  if (!button) return;
  const action = button.getAttribute('data-action');
  const id = button.getAttribute('data-id');
  if (action === 'edit') {
    editUser(id);
  }
  if (action === 'delete') {
    deleteUser(id);
  }
}

function editUser(id) {
  const user = state.users.find(item => item.id === id);
  if (!user) {
    showAlert('No se encontró el usuario seleccionado.', 'warning');
    return;
  }
  openUserModal(user);
}

async function deleteUser(id) {
  const user = state.users.find(item => item.id === id);
  if (!user) {
    showAlert('No se encontró el usuario seleccionado.', 'warning');
    return;
  }
  if (!confirm(`¿Eliminar al usuario @${user.usuario}?`)) return;
  try {
    const response = await adminFetch(`/users/${id}`, { method: 'DELETE' });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No se pudo eliminar');
    }
    showAlert('Usuario eliminado.', 'success');
    await loadUsers();
  } catch (error) {
    console.error('Error eliminando usuario:', error);
    showAlert(error.message || 'Error eliminando usuario.', 'danger');
  }
}

async function handleUserFormSubmit(event) {
  event.preventDefault();
  const usernameInput = document.getElementById('userUsername');
  const nameInput = document.getElementById('userName');
  const passwordInput = document.getElementById('userPassword');
  const confirmInput = document.getElementById('userPasswordConfirm');
  const allCompaniesInput = document.getElementById('userAllCompanies');

  const username = usernameInput.value.trim();
  const nombre = nameInput.value.trim();
  const password = passwordInput.value;
  const confirmPassword = confirmInput.value;
  const allCompanies = allCompaniesInput.checked;
  const companies = state.userCompanySelection instanceof Set
    ? Array.from(state.userCompanySelection)
    : [];

  [usernameInput, passwordInput, confirmInput].forEach(input => input.classList.remove('is-invalid'));
  setUserCompaniesValidity(true);

  let hasError = false;
  if (!username) {
    usernameInput.classList.add('is-invalid');
    hasError = true;
  }

  const isEdit = !!state.editingUserId;
  const passwordProvided = password.trim().length > 0;
  if ((passwordProvided || !isEdit) && password.length < 6) {
    passwordInput.classList.add('is-invalid');
    hasError = true;
  }

  if (passwordProvided || confirmPassword.trim().length > 0) {
    if (password !== confirmPassword) {
      confirmInput.classList.add('is-invalid');
      hasError = true;
    }
  }

  if (!allCompanies && companies.length === 0) {
    setUserCompaniesValidity(false);
    hasError = true;
  }

  if (hasError) {
    return;
  }

  const payload = {
    usuario: username,
    nombre: nombre || null,
    empresas: allCompanies ? '*' : companies
  };
  if (passwordProvided) {
    payload.password = password;
  }

  try {
    const response = await adminFetch(state.editingUserId ? `/users/${state.editingUserId}` : '/users', {
      method: state.editingUserId ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No se pudo guardar el usuario');
    }
    const modal = bootstrap.Modal.getInstance(document.getElementById('userModal'));
    modal?.hide();
    state.editingUserId = null;
    showAlert(isEdit ? 'Usuario actualizado.' : 'Usuario creado.', 'success');
    await loadUsers();
  } catch (error) {
    console.error('Error guardando usuario:', error);
    showAlert(error.message || 'No fue posible guardar el usuario.', 'danger');
  }
}

async function handleReportSettingsSubmit(event) {
  event.preventDefault();
  if (!isAdmin()) {
    showAlert('Solo un administrador puede actualizar la configuración.', 'danger');
    return;
  }
  const payload = {
    export: {
      defaultFormat: document.getElementById('exportDefaultFormat')?.value || 'pdf',
      availableFormats: Array.from(document.querySelectorAll('#exportFormatsContainer input[data-format-option]:checked'))
        .map(input => input.value)
    },
    customization: {
      includeCharts: (document.getElementById('customIncludeCharts')?.value ?? 'true') !== 'false',
      includeMovements: (document.getElementById('customIncludeMovements')?.value ?? 'true') !== 'false',
      includeObservations: (document.getElementById('customIncludeObservations')?.value ?? 'true') !== 'false',
      includeUniverse: (document.getElementById('customIncludeUniverse')?.value ?? 'true') !== 'false',
      csv: {
        includePoResumen: (document.getElementById('customCsvIncludePoResumen')?.value ?? 'true') !== 'false',
        includeRemisiones: (document.getElementById('customCsvIncludeRemisiones')?.value ?? 'true') !== 'false',
        includeFacturas: (document.getElementById('customCsvIncludeFacturas')?.value ?? 'true') !== 'false',
        includeTotales: (document.getElementById('customCsvIncludeTotales')?.value ?? 'true') !== 'false',
        includeUniverseInfo: (document.getElementById('customCsvIncludeUniverseInfo')?.value ?? 'true') !== 'false'
      }
    },
    branding: {
      headerTitle: getInputValue('brandingHeaderTitle'),
      headerSubtitle: getInputValue('brandingHeaderSubtitle'),
      companyName: getInputValue('brandingCompanyName'),
      footerText: getInputValue('brandingFooterText'),
      letterheadEnabled: document.getElementById('brandingLetterheadEnabled')?.checked ?? false,
      letterheadTop: getInputValue('brandingLetterheadTop'),
      letterheadBottom: getInputValue('brandingLetterheadBottom'),
      remColor: document.getElementById('brandingRemColor')?.value,
      facColor: document.getElementById('brandingFacColor')?.value,
      restanteColor: document.getElementById('brandingRestanteColor')?.value,
      accentColor: document.getElementById('brandingAccentColor')?.value
    }
  };
  if (!payload.export.availableFormats.includes('pdf')) {
    payload.export.availableFormats.unshift('pdf');
  }
  try {
    const response = await adminFetch('/report-settings', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No se pudo actualizar la configuración');
    }
    state.reportSettings = data.settings;
    state.reportEngines = Array.isArray(data.engines) ? data.engines : [];
    state.reportFormatsCatalog = Array.isArray(data.formats) ? data.formats : state.reportFormatsCatalog;
    if (state.reportSettings?.export?.defaultFormat) {
      state.selectedFormat = state.reportSettings.export.defaultFormat;
    }
    state.selectedEngine = state.reportSettings.defaultEngine;
    renderReportSettingsForm();
    renderReportEngineSelector();
    renderReportFormatSelect();
    showAlert('Configuración de reportes actualizada.', 'success');
  } catch (error) {
    console.error('Error guardando configuración de reportes:', error);
    showAlert(error.message || 'No fue posible guardar la configuración de reportes.', 'danger');
  }
}

function logout() {
  sessionStorage.removeItem('porpt-empresas');
  sessionStorage.removeItem('porpt-is-admin');
  window.location.href = 'index.html';
}

function setupLoginPage() {
  const form = document.getElementById('loginForm');
  form?.addEventListener('submit', login);
  document.getElementById('loginButton')?.addEventListener('click', login);
  const usernameInput = document.getElementById('username');
  const passwordInput = document.getElementById('password');
  usernameInput?.addEventListener('input', () => usernameInput.classList.remove('is-invalid'));
  passwordInput?.addEventListener('input', () => passwordInput.classList.remove('is-invalid'));
}

async function setupDashboard() {
  await ensureEmpresasCatalog();
  initializeDashboardPanelControls();
  const empresaSelect = document.getElementById('empresaSearch');
  initializeSearchableSelect(empresaSelect, { placeholder: 'Selecciona una empresa' });
  const poSelect = document.getElementById('poSearch');
  initializeSearchableSelect(poSelect, { placeholder: 'Selecciona una PO base o extensión' });
  renderEmpresaOptions();
  renderPoOptions();
  renderUniverseEmpresaSelect();
  renderSelectedPoChips();
  initializeOverviewTab();
  empresaSelect?.addEventListener('change', event => {
    const value = event.target.value;
    if (!value) {
      return;
    }
    selectEmpresa(value);
  });
  poSelect?.addEventListener('change', event => {
    const value = event.target.value;
    if (!value) {
      return;
    }
    selectPo(value);
  });
  document.getElementById('generateReportBtn')?.addEventListener('click', openReportPreview);
  document.getElementById('clearSelectionBtn')?.addEventListener('click', clearSelectedPos);
  document.getElementById('selectedPoContainer')?.addEventListener('click', event => {
    const button = event.target.closest('button[data-action="remove-po"]');
    if (!button) return;
    const baseId = button.getAttribute('data-po');
    removePoFromSelection(baseId);
  });
  const extensionContainer = document.getElementById('extensionSelectionContainer');
  extensionContainer?.addEventListener('change', handleExtensionSelectionChange);
  extensionContainer?.addEventListener('click', handleExtensionSelectionAction);
  extensionContainer?.addEventListener('change', handleManualExtensionSelectChange);
  document.getElementById('universeEmpresaSelect')?.addEventListener('change', event => {
    const value = event.target.value;
    if (!value) {
      state.selectedEmpresa = '';
      state.pos = [];
      state.selectedPoIds = [];
      state.selectedPoDetails.clear();
      state.manualExtensions.clear();
      destroyCharts();
      renderSelectedPoChips();
      updateUniverseControls();
      resetSearchableSelect(document.getElementById('empresaSearch'));
      resetSearchableSelect(document.getElementById('poSearch'));
      renderPoOptions();
      renderUniverseEmpresaSelect();
      setOverviewEmpresaSelectValue('');
      state.overview.filters.search = '';
      updateOverviewSearchInput('');
      resetOverviewState({ keepFilters: true });
      hideDashboardPanel();
      return;
    }
    selectEmpresa(value);
  });
  const adminLink = document.getElementById('adminLink');
  if (adminLink && !isAdmin()) {
    adminLink.classList.add('d-none');
  }
  const reportSettingsLink = document.getElementById('reportSettingsLink');
  if (reportSettingsLink && !isAdmin()) {
    reportSettingsLink.classList.add('d-none');
  }
  const universeMode = document.getElementById('universeFilterMode');
  universeMode?.addEventListener('change', event => handleUniverseModeChange(event.target.value));
  document.getElementById('universeStartDate')?.addEventListener('change', event => {
    state.universeStartDate = event.target.value;
  });
  document.getElementById('universeEndDate')?.addEventListener('change', event => {
    state.universeEndDate = event.target.value;
  });
  const universeFormatSelect = document.getElementById('universeFormat');
  if (universeFormatSelect) {
    universeFormatSelect.value = state.universeFormat || 'pdf';
    universeFormatSelect.addEventListener('change', event => {
      state.universeFormat = event.target.value || 'pdf';
      saveSession('porpt-universe-format', state.universeFormat);
    });
  }
  const universeFileInput = document.getElementById('universeFileName');
  if (universeFileInput) {
    universeFileInput.value = state.universeCustomName || '';
    universeFileInput.addEventListener('input', handleUniverseFileNameChange);
  }
  document.getElementById('generateUniverseReportBtn')?.addEventListener('click', generateUniverseReport);
  handleUniverseModeChange(state.universeFilterMode);
  updateUniverseControls();
  const engineMenu = document.getElementById('reportEngineMenu');
  engineMenu?.addEventListener('click', event => {
    const button = event.target.closest('button[data-engine]');
    if (!button) return;
    state.selectedEngine = button.getAttribute('data-engine');
    renderReportEngineSelector();
  });
  const formatSelect = document.getElementById('reportFormatSelect');
  formatSelect?.addEventListener('change', event => {
    state.selectedFormat = event.target.value;
    updateReportOverviewMeta();
  });
  const reportNameInput = document.getElementById('reportFileNameInput');
  if (reportNameInput) {
    reportNameInput.value = state.customReportName || '';
    reportNameInput.addEventListener('input', handleReportFileNameChange);
  }
  document.getElementById('confirmReportDownloadBtn')?.addEventListener('click', confirmReportDownload);
  await loadReportSettings();
  applySavedCustomization();
  setupCustomizationControls();
  renderCustomizationSummary();
  setupWizardNavigation();
  updateReportOverviewMeta();
  renderReportSelectionOverview();
  showAlert('Selecciona empresa y PO para visualizar el tablero.', 'info', 9000);
}

async function setupReportSettings() {
  if (!isAdmin()) {
    showAlert('Solo los administradores pueden acceder a la configuración de reportes.', 'danger');
    window.location.href = 'dashboard.html';
    return;
  }
  await loadReportSettings({ requireAdmin: true });
  document.getElementById('reportSettingsForm')?.addEventListener('submit', handleReportSettingsSubmit);
  document.getElementById('brandingLetterheadEnabled')?.addEventListener('change', event => {
    toggleBrandingLetterheadFields(!event.target.checked);
  });
  setupPathPickers();
  showAlert('Configura los formatos de exportación y la identidad visual desde este panel.', 'info', 7000);
}

async function setupAdmin() {
  if (!isAdmin()) {
    showAlert('Solo los administradores pueden acceder a esta sección.', 'danger');
    window.location.href = 'dashboard.html';
    return;
  }
  await ensureEmpresasCatalog();
  resetUserCompanySelection([], { skipValidityUpdate: true });
  await loadUsers();
  document.getElementById('loadUsersBtn')?.addEventListener('click', loadUsers);
  document.getElementById('newUserBtn')?.addEventListener('click', () => openUserModal());
  const table = document.getElementById('usersTable');
  table?.addEventListener('click', handleUsersTableClick);
  const userForm = document.getElementById('userForm');
  userForm?.addEventListener('submit', handleUserFormSubmit);
  document.getElementById('userAllCompanies')?.addEventListener('change', event => {
    setCompaniesFieldDisabled(event.target.checked);
    updateUserCompaniesValidity();
  });
  document.getElementById('userCompanyPicker')?.addEventListener('change', handleUserCompanyPickerChange);
  document.getElementById('userCompanyChecklist')?.addEventListener('change', handleUserCompanyChecklistChange);
  ['userUsername', 'userPassword', 'userPasswordConfirm'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', event => event.target.classList.remove('is-invalid'));
  });
  showAlert('Desde aquí administra usuarios. Usa el botón “Configuración de reportes” para ajustar los formatos.', 'info', 7000);
}

function init() {
  const path = window.location.pathname;
  if (path.endsWith('index.html') || path === '/' || path === '') {
    setupLoginPage();
  } else if (path.endsWith('dashboard.html')) {
    setupDashboard();
  } else if (path.endsWith('admin.html')) {
    setupAdmin();
  } else if (path.endsWith('report-settings.html')) {
    setupReportSettings();
  }
}

document.addEventListener('DOMContentLoaded', init);

window.login = login;
window.loadUsers = loadUsers;
window.logout = logout;
window.generateReport = generateReport;
window.editUser = editUser;
window.deleteUser = deleteUser;
