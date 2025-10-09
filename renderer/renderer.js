const state = {
  empresas: [],
  pos: [],
  selectedEmpresa: '',
  selectedPoIds: [],
  selectedPoDetails: new Map(),
  charts: new Map(),
  summary: null,
  users: [],
  reportSettings: null,
  reportEngines: [],
  selectedEngine: '',
  reportFormatsCatalog: [],
  selectedFormat: 'pdf',
  customization: {
    includeCharts: true,
    includeMovements: true,
    includeObservations: true,
    includeUniverse: true,
    csv: {
      includePoResumen: true,
      includeRemisiones: true,
      includeFacturas: true,
      includeTotales: true,
      includeUniverseInfo: true
    }
  },
  customReportName: readSession('porpt-custom-report-name', ''),
  editingUserId: null,
  universeFilterMode: 'global',
  universeStartDate: '',
  universeEndDate: '',
  universeFormat: readSession('porpt-universe-format', 'pdf'),
  universeCustomName: readSession('porpt-universe-report-name', ''),
  isGeneratingReport: false
};

const ALERT_TYPE_CLASS = {
  success: 'success',
  danger: 'danger',
  warning: 'warning',
  info: 'info'
};

function mapAlertType(type) {
  return ALERT_TYPE_CLASS[type] || 'info';
}

function showAlert(message, type = 'info', delay = 6000) {
  const container = document.getElementById('alerts');
  if (!container) return;
  const wrapper = document.createElement('div');
  wrapper.className = `toast align-items-center text-bg-${mapAlertType(type)} border-0 shadow`;
  wrapper.setAttribute('role', 'alert');
  wrapper.setAttribute('aria-live', 'assertive');
  wrapper.setAttribute('aria-atomic', 'true');
  wrapper.innerHTML = `
    <div class="d-flex">
      <div class="toast-body fw-medium">${message}</div>
      <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Cerrar"></button>
    </div>
  `;
  container.appendChild(wrapper);
  const toast = new bootstrap.Toast(wrapper, { delay });
  toast.show();
  wrapper.addEventListener('hidden.bs.toast', () => wrapper.remove());
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

function destroyCharts() {
  state.charts.forEach(chart => {
    if (chart && typeof chart.destroy === 'function') {
      chart.destroy();
    }
  });
  state.charts.clear();
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

function formatCurrency(value) {
  const amount = normalizeNumber(value);
  return amount.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
  const restanteAmount = normalizeNumber(
    totals.restante != null ? totals.restante : total - (totalRem + totalFac)
  );
  const rem = clampPercentage((totalRem / total) * 100);
  const fac = clampPercentage((totalFac / total) * 100);
  let rest = clampPercentage((restanteAmount / total) * 100);
  if (roundTo(rem + fac + rest, 2) !== 100) {
    rest = clampPercentage(100 - (rem + fac));
  }
  return { rem, fac, rest };
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

function buildPoGroups(items = []) {
  const groups = new Map();
  items.forEach(item => {
    const baseId = item?.baseId || item?.id;
    if (!baseId) {
      return;
    }
    const key = baseId.trim();
    if (!key) {
      return;
    }
    if (!groups.has(key)) {
      groups.set(key, {
        baseId: key,
        ids: [],
        total: 0,
        totalRem: 0,
        totalFac: 0
      });
    }
    const group = groups.get(key);
    const totals = getTotalsWithDefaults(item?.totals);
    const authorized = normalizeNumber(item?.total ?? totals.total);
    group.ids.push(item.id);
    group.total += authorized;
    group.totalRem += normalizeNumber(totals.totalRem);
    group.totalFac += normalizeNumber(totals.totalFac);
  });

  return Array.from(groups.values())
    .map(group => {
      const totalConsumo = group.totalRem + group.totalFac;
      const restante = Math.max(group.total - totalConsumo, 0);
      const totals = {
        total: roundTo(group.total),
        totalRem: roundTo(group.totalRem),
        totalFac: roundTo(group.totalFac),
        totalConsumo: roundTo(totalConsumo),
        restante: roundTo(restante)
      };
      return {
        baseId: group.baseId,
        ids: group.ids,
        totals,
        percentages: computePercentagesFromTotals(totals)
      };
    })
    .sort((a, b) => a.baseId.localeCompare(b.baseId));
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

function renderEmpresaOptions(filterText = '') {
  const datalist = document.getElementById('empresaOptions');
  if (!datalist) return;
  const normalized = filterText.toLowerCase();
  datalist.innerHTML = '';
  state.empresas
    .filter(empresa => empresa.toLowerCase().includes(normalized))
    .slice(0, 50)
    .forEach(empresa => {
      const option = document.createElement('option');
      option.value = empresa;
      option.textContent = empresa;
      datalist.appendChild(option);
    });
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

function renderPoOptions(filterText = '') {
  const datalist = document.getElementById('poOptions');
  if (!datalist) return;
  const normalized = filterText.toLowerCase();
  datalist.innerHTML = '';
  state.pos
    .filter(po => po.display.toLowerCase().includes(normalized))
    .slice(0, 100)
    .forEach(po => {
      const option = document.createElement('option');
      option.value = po.id;
      option.textContent = po.display;
      datalist.appendChild(option);
    });
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

function renderExtensionSelection() {
  const container = document.getElementById('extensionSelectionContainer');
  if (!container) return;
  container.innerHTML = '';
  if (!state.selectedPoIds.length) {
    container.innerHTML = '<p class="text-muted small mb-0">Selecciona una PO base para ver sus extensiones sugeridas.</p>';
    return;
  }

  state.selectedPoIds.forEach(baseId => {
    const variants = getSuggestedVariants(baseId);
    const detail = ensureSelectedPoDetail(baseId);
    const selectedVariants = detail ? Array.from(detail.variants) : [baseId];
    const selectedCount = selectedVariants.length;
    const badgeLabel = selectedCount === 1
      ? 'Solo base seleccionada'
      : `${selectedCount} variantes incluidas`;

    const card = document.createElement('div');
    card.className = 'card shadow-sm border border-primary-subtle mb-3';
    const rows = document.createElement('div');
    rows.className = 'row g-2';

    variants.forEach(variant => {
      const col = document.createElement('div');
      col.className = 'col-12 col-md-6';
      const formCheck = document.createElement('div');
      formCheck.className = 'form-check form-switch';
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
      label.className = 'form-check-label small';
      label.setAttribute('for', input.id);
      const suggestionBadge = variant.id === baseId
        ? '<span class="badge text-bg-secondary ms-2">Base</span>'
        : variant.isExtension !== false
          ? '<span class="badge text-bg-info-subtle text-info-emphasis ms-2">Sugerencia</span>'
          : '';
      label.innerHTML = `${escapeHtml(variant.display || variant.id)}${suggestionBadge}`;
      formCheck.appendChild(input);
      formCheck.appendChild(label);
      col.appendChild(formCheck);
      rows.appendChild(col);
    });

    if (variants.length <= 1) {
      const col = document.createElement('div');
      col.className = 'col-12';
      col.innerHTML = '<p class="text-muted small mb-0">Sin extensiones adicionales registradas para esta PO.</p>';
      rows.appendChild(col);
    }

    card.innerHTML = `
      <div class="card-body">
        <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
          <div>
            <h5 class="card-title h6 mb-0">PO ${escapeHtml(baseId)}</h5>
            <p class="text-muted small mb-0">Marca únicamente las extensiones que deban sumarse al consumo.</p>
          </div>
          <span class="badge text-bg-light text-dark">${escapeHtml(badgeLabel)}</span>
        </div>
      </div>
    `;
    card.querySelector('.card-body').appendChild(rows);
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
    const current = stateDetail.variants || new Set();
    const normalizedSet = new Set(normalized);
    const sameSize = current.size === normalizedSet.size;
    const isSame = sameSize && Array.from(normalizedSet).every(value => current.has(value));
    if (!isSame) {
      stateDetail.variants = new Set(normalizedSet);
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
    empty.textContent = 'No hay POs seleccionadas. Utiliza el buscador para agregarlas.';
    container.appendChild(empty);
    document.getElementById('dashboardContent')?.classList.add('d-none');
    renderReportSelectionOverview();
    if (help) {
      help.textContent = 'Agrega una PO base y luego elige manualmente qué extensiones deseas incluir.';
    }
    renderExtensionSelection();
    return;
  }
  if (help) {
    const label = state.selectedPoIds.length === 1
      ? '1 PEO base seleccionada'
      : `${state.selectedPoIds.length} PEOs base seleccionadas`;
    help.textContent = `${label}. Ajusta las extensiones en la tarjeta inferior o quita cualquiera dando clic en la “x”.`;
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
  renderSelectedPoChips();
  if (state.selectedPoIds.length === 0) {
    state.summary = null;
    destroyCharts();
    document.getElementById('summaryCards')?.replaceChildren();
    document.getElementById('poTable')?.replaceChildren();
    document.getElementById('extensionsContainer')?.replaceChildren();
    document.getElementById('dashboardContent')?.classList.add('d-none');
  } else {
    updateSummary();
  }
}

function clearSelectedPos() {
  if (!state.selectedPoIds.length) return;
  state.selectedPoIds = [];
  state.selectedPoDetails.clear();
  renderSelectedPoChips();
  state.summary = null;
  destroyCharts();
  document.getElementById('summaryCards')?.replaceChildren();
  document.getElementById('poTable')?.replaceChildren();
  document.getElementById('extensionsContainer')?.replaceChildren();
  document.getElementById('dashboardContent')?.classList.add('d-none');
  const input = document.getElementById('poSearch');
  if (input) {
    input.value = '';
  }
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
  const poInput = document.getElementById('poSearch');
  if (poInput) {
    poInput.value = '';
  }
  const empresaInput = document.getElementById('empresaSearch');
  if (empresaInput) {
    empresaInput.value = value;
  }
  renderUniverseEmpresaSelect();
  document.getElementById('dashboardContent').classList.add('d-none');
  destroyCharts();
  renderSelectedPoChips();
  updateUniverseControls();
  await loadPOs();
}

async function loadPOs() {
  if (!state.selectedEmpresa) return;
  state.pos = [];
  state.summary = null;
  try {
    showAlert(`Cargando POs de ${state.selectedEmpresa}...`, 'info', 3500);
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
  }
}

async function selectPo(poId) {
  if (!poId) return;
  addPoToSelection(poId);
  const input = document.getElementById('poSearch');
  if (input) {
    input.value = '';
  }
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
    ? '1 PEO base (extensiones según tu selección)'
    : `${selectionCount} PEOs base combinadas`;
  container.innerHTML = `
    <div class="col-md-4">
      <div class="card shadow-sm h-100 border-primary border-2">
        <div class="card-body">
          <h5 class="card-title text-primary">${selectionCount > 1 ? 'Total combinado' : 'Total PO (Grupo)'}</h5>
          <p class="display-6 fw-bold">$${formatCurrency(totals.total)}</p>
          <p class="text-muted mb-0">${selectionLabel}.</p>
          <span class="badge text-bg-light text-dark mt-2">${summary.empresaLabel || summary.empresa || 'Empresa'}</span>
        </div>
      </div>
    </div>
    <div class="col-md-4">
      <div class="card shadow-sm h-100 border-info border-2">
        <div class="card-body">
          <h5 class="card-title text-info">Consumo acumulado</h5>
          <p class="display-6 fw-bold">$${formatCurrency(totals.totalConsumo)}</p>
          <p class="text-muted mb-0">${formatPercentageLabel(percentages.rem)} Remisiones · ${formatPercentageLabel(percentages.fac)} Facturas</p>
        </div>
      </div>
    </div>
    <div class="col-md-4">
      <div class="card shadow-sm h-100 border-success border-2">
        <div class="card-body">
          <h5 class="card-title text-success">Disponible</h5>
          <p class="display-6 fw-bold">$${formatCurrency(totals.restante)}</p>
          <p class="text-muted mb-0">${formatPercentageLabel(percentages.rest)} restante del presupuesto autorizado.</p>
        </div>
      </div>
    </div>
  `;
}

function renderTable(summary) {
  const table = document.getElementById('poTable');
  if (!table) return;
  const rows = summary.items
    .map(item => {
      const totalAmount = Number(item.total || 0);
      const totalOriginalAmount = Number(item.totalOriginal ?? item.total ?? 0);
      const subtotalAmount = Number(item.subtotal || 0);
      const showSubtotal = subtotalAmount > 0 && Math.abs(subtotalAmount - totalAmount) > 0.009;
      const subtotalHtml = showSubtotal
        ? `<div class="small text-muted">Subtotal: $${formatCurrency(subtotalAmount)}</div>`
        : '';
      const adjustment = item.ajusteDocSig;
      const appliedDiff = Number(adjustment?.diferenciaAplicada ?? 0);
      const hasAdjustment = appliedDiff > 0.009;
      const adjustmentDetails = [];
      if (hasAdjustment) {
        adjustmentDetails.push(`<div class="small text-muted">Original: $${formatCurrency(totalOriginalAmount)}</div>`);
        const tipoLabel = adjustment?.tipo === 'F' ? 'Factura' : 'Remisión';
        const docLabel = escapeHtml(adjustment?.docSig || '-');
        adjustmentDetails.push(
          `<div class="small text-warning">Ajuste por ${tipoLabel} ${docLabel}: -$${formatCurrency(appliedDiff)}</div>`
        );
      }
      if (subtotalHtml) {
        adjustmentDetails.push(subtotalHtml);
      }
      const totalCellDetails = adjustmentDetails.join('');
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
      return `
      <tr data-po="${item.id}">
        <td class="fw-semibold">${item.id}</td>
        <td>${item.fecha || '-'}</td>
        <td>$${formatCurrency(totalAmount)}${totalCellDetails}</td>
        <td>$${formatCurrency(normalizedTotals.totalRem)} (${formatPercentageLabel(percentages.rem)})</td>
        <td>$${formatCurrency(normalizedTotals.totalFac)} (${formatPercentageLabel(percentages.fac)})</td>
        <td>$${formatCurrency(normalizedTotals.restante)} (${formatPercentageLabel(percentages.rest)})</td>
        <td>
          <button class="btn btn-outline-primary btn-sm" data-po="${item.id}" data-action="show-modal">Detalle</button>
        </td>
      </tr>
      `;
    })
    .join('');
  table.innerHTML = `
    <thead class="table-light">
      <tr>
        <th>PO</th>
        <th>Fecha</th>
        <th>Total</th>
        <th>Remisiones</th>
        <th>Facturas</th>
        <th>Disponible</th>
        <th>Acción</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  `;
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
    const totalOriginal = Number(item.totalOriginal ?? item.total ?? 0);
    const adjustment = item.ajusteDocSig;
    const appliedDiff = Number(adjustment?.diferenciaAplicada ?? 0);
    const hasAdjustment = appliedDiff > 0.009;
    const showSubtotal = subtotalAmount > 0 && Math.abs(subtotalAmount - Number(item.total || 0)) > 0.009;
    const subtotalLinea = showSubtotal ? `Subtotal (CAN_TOT): $${formatCurrency(subtotalAmount)}` : '';
    const ajusteLinea = hasAdjustment
      ? `Ajuste aplicado (${adjustment?.tipo === 'F' ? 'Factura' : 'Remisión'} ${adjustment?.docSig || '-' }): -$${formatCurrency(appliedDiff)}`
      : '';
    const originalLinea = hasAdjustment ? `Total original FACTP: $${formatCurrency(totalOriginal)}` : '';
    const totalLinea = hasAdjustment
      ? `Total autorizado ajustado: $${totalAutorizado}`
      : `Total autorizado (IMPORTE): $${totalAutorizado}`;
    const docSigLinea = item.docSig ? `DOC_SIG relacionado: ${item.docSig}` : '';
    const totalesResumen = [
      totalLinea,
      originalLinea,
      ajusteLinea,
      subtotalLinea,
      docSigLinea,
      `Remisiones acumuladas: $${formatCurrency(item.totals?.totalRem ?? 0)}`,
      `Facturas acumuladas: $${formatCurrency(item.totals?.totalFac ?? 0)}`,
      `Disponible: $${formatCurrency(item.totals?.restante ?? 0)}`
    ]
      .filter(Boolean)
      .join('\n');
    modalBody.innerHTML = `
      <div class="mb-3">
        <h6 class="text-secondary">Totales del pedido</h6>
        <pre class="bg-light rounded p-3 small">${totalesResumen}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-primary">Remisiones</h6>
        <pre class="bg-light rounded p-3 small">${item.remisionesTexto}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-danger">Facturas</h6>
        <pre class="bg-light rounded p-3 small">${item.facturasTexto}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-success">Notas de venta vinculadas</h6>
        <pre class="bg-light rounded p-3 small">${item.notasVentaTexto || 'Sin notas de venta vinculadas'}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-info">Cotizaciones relacionadas</h6>
        <pre class="bg-light rounded p-3 small">${item.cotizacionesTexto || 'Sin cotizaciones relacionadas'}</pre>
      </div>
      <div>
        <h6 class="text-warning">Alertas</h6>
        <pre class="bg-light rounded p-3 small">${item.alertasTexto}</pre>
      </div>
    `;
    const modal = new bootstrap.Modal(modalElement);
    modal.show();
  };
}

function renderCharts(summary) {
  destroyCharts();
  const items = Array.isArray(summary.items) ? summary.items : [];
  const groups = buildPoGroups(items);
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

  const chartEntries = groups.map(group => {
    const variantCount = group.ids.length;
    const suffix = variantCount > 1 ? ` (+${variantCount - 1} ext.)` : '';
    return {
      id: group.baseId,
      label: `PO ${group.baseId}${suffix}`,
      shortLabel: group.baseId,
      variantCount,
      totals: group.totals,
      percentages: group.percentages
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
          percentage: clampPercentage(aggregatedPercentages.rem + aggregatedPercentages.fac)
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
  const dynamicHeight = Math.max(190, chartEntries.length * 38);

  const barChartsMeta = [
    { canvas: ctxRem, key: 'rem', amountKey: 'totalRem', percKey: 'rem', label: 'Remisiones', color: CHART_COLORS.rem },
    { canvas: ctxFac, key: 'fac', amountKey: 'totalFac', percKey: 'fac', label: 'Facturas', color: CHART_COLORS.fac }
  ];

  barChartsMeta.forEach(meta => {
    if (!meta.canvas) return;
    const wrapper = meta.canvas.closest('.chart-wrapper');
    if (wrapper) {
      wrapper.style.minHeight = `${dynamicHeight}px`;
    }
    const dataset = {
      label: `${meta.label} ($)`,
      data: chartEntries.map(entry => entry.totals[meta.amountKey]),
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
                return `${meta.label}: $${formatCurrency(entry.totals[meta.amountKey])} (${formatPercentageLabel(entry.percentages[meta.percKey])})`;
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            grid: { color: 'rgba(148, 163, 184, 0.25)', borderDash: [4, 4] },
            ticks: {
              color: '#475569',
              callback: value => AXIS_CURRENCY_FORMATTER.format(value)
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
    }
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
          data: chartEntries.map(entry => entry.percentages[meta.percKey]),
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
                return `${meta.label}: $${formatCurrency(entry.totals[meta.amountKey])} (${formatPercentageLabel(entry.percentages[meta.percKey])})`;
              }
            }
          }
        },
        scales: {
          x: {
            stacked: true,
            beginAtZero: true,
            max: 100,
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
      const card = document.createElement('div');
      card.className = 'col-md-6 col-xl-4';
      const variantBadge = entry.variantCount === 1 ? 'Solo base' : `${entry.variantCount} variantes`;
      card.innerHTML = `
        <div class="card h-100 shadow-sm chart-card">
          <div class="card-body">
            <div class="d-flex justify-content-between align-items-start gap-3">
              <div>
                <h5 class="card-title mb-1">${escapeHtml(entry.label)}</h5>
                <p class="small text-muted mb-0">Autorizado: $${formatCurrency(entry.totals.total)}</p>
              </div>
              <span class="badge text-bg-light text-dark">${escapeHtml(variantBadge)}</span>
            </div>
            <div class="chart-wrapper mt-2">
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
          </div>
        </div>
      `;
      extensionContainer.appendChild(card);
      const canvas = card.querySelector('canvas');
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

function normalizeCustomizationState(customization = {}) {
  const csv = customization.csv || {};
  return {
    includeCharts: customization.includeCharts !== false,
    includeMovements: customization.includeMovements !== false,
    includeObservations: customization.includeObservations !== false,
    includeUniverse: customization.includeUniverse !== false,
    csv: {
      includePoResumen: csv.includePoResumen !== false,
      includeRemisiones: csv.includeRemisiones !== false,
      includeFacturas: csv.includeFacturas !== false,
      includeTotales: csv.includeTotales !== false,
      includeUniverseInfo: csv.includeUniverseInfo !== false
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
  document.querySelectorAll('[data-customization-toggle]').forEach(input => {
    const key = input.getAttribute('data-customization-toggle');
    if (!key) return;
    const group = input.getAttribute('data-customization-group');
    const value = group === 'csv' ? config.csv?.[key] !== false : config[key] !== false;
    input.checked = !!value;
  });
}

function handleCustomizationToggleChange(event) {
  const input = event.target;
  if (!input || input.type !== 'checkbox') return;
  const key = input.getAttribute('data-customization-toggle');
  if (!key) return;
  const group = input.getAttribute('data-customization-group');
  if (group === 'csv') {
    state.customization = {
      ...state.customization,
      csv: {
        ...state.customization.csv,
        [key]: input.checked
      }
    };
  } else {
    state.customization = {
      ...state.customization,
      [key]: input.checked
    };
  }
  state.customization = normalizeCustomizationState(state.customization);
  saveSession('porpt-customization', state.customization);
  renderCustomizationSummary();
  updateReportOverviewMeta();
}

function setupCustomizationControls() {
  setCustomizationControlStates();
  document.querySelectorAll('[data-customization-toggle]').forEach(input => {
    input.removeEventListener('change', handleCustomizationToggleChange);
    input.addEventListener('change', handleCustomizationToggleChange);
  });
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
    if (help) {
      help.textContent = 'Aún no hay suficientes datos para preparar el reporte.';
    }
    return;
  }
  if (!state.summary) {
    container.innerHTML = '<p class="text-muted mb-0">Cuando el dashboard termine de cargar, verás aquí el resumen del reporte.</p>';
    if (help) {
      help.textContent = 'Actualiza el dashboard para habilitar la descarga.';
    }
    return;
  }
  const items = Array.isArray(state.summary.items) ? state.summary.items : [];
  const groups = buildPoGroups(items);
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
      const variantLabel = group.ids.length === 1 ? 'Solo base' : `${group.ids.length} variantes`;
      return `
        <li class="list-group-item d-flex flex-wrap justify-content-between gap-2">
          <div>
            <span class="fw-semibold">PO ${escapeHtml(group.baseId)}</span>
            <span class="badge text-bg-light text-dark ms-2">${escapeHtml(variantLabel)}</span>
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
  if (help) {
    help.textContent = 'Todo listo. Revisa el resumen y después descarga el reporte.';
  }
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
    document.getElementById('dashboardContent')?.classList.remove('d-none');
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
  const engine = state.selectedEngine || state.reportSettings?.defaultEngine || 'jasper';
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
    companyLabel.textContent = state.reportSettings?.branding?.companyName || 'SITTEL';
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
  if (engine.id === 'jasper') {
    if (state.reportSettings?.jasper?.enabled === false) {
      text = 'Jasper deshabilitado';
      badgeClass = 'text-bg-secondary';
    } else if (state.reportSettings?.jasper?.available) {
      text = 'Jasper listo';
      badgeClass = 'text-bg-success';
    } else {
      text = 'Configura Jasper';
      badgeClass = 'text-bg-warning text-dark';
    }
  } else {
    text = 'PDF directo disponible';
    badgeClass = 'text-bg-info text-dark';
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
  const input = document.getElementById('userCompanies');
  if (input) {
    input.disabled = disabled;
    if (disabled) {
      input.classList.remove('is-invalid');
    }
  }
}

function updateJasperStatusBadge() {
  const badge = document.getElementById('jasperStatusBadge');
  if (!badge) return;
  if (!state.reportSettings) {
    badge.textContent = 'Sin datos';
    badge.className = 'badge text-bg-secondary';
    return;
  }
  const jasper = state.reportSettings.jasper || {};
  if (jasper.enabled === false) {
    badge.textContent = 'Jasper deshabilitado';
    badge.className = 'badge text-bg-secondary';
    return;
  }
  if (jasper.available) {
    badge.textContent = 'Jasper listo';
    badge.className = 'badge text-bg-success';
  } else {
    badge.textContent = 'Jasper requiere configuración';
    badge.className = 'badge text-bg-warning text-dark';
  }
}

function toggleJasperFields(disabled) {
  ['jasperCompiledDir', 'jasperTemplatesDir', 'jasperFontsDir', 'jasperReportName', 'jasperDataSource', 'jasperJsonQuery']
    .forEach(id => {
      const input = document.getElementById(id);
      if (input) {
        input.disabled = disabled;
        if (disabled) {
          input.classList.remove('is-invalid');
        }
      }
    });
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
  const select = document.getElementById('defaultReportEngine');
  if (select) {
    select.innerHTML = (state.reportEngines || [])
      .map(engine => `<option value="${engine.id}" ${engine.available ? '' : 'disabled'}>${engine.label}</option>`)
      .join('');
    select.value = state.reportSettings.defaultEngine || select.value;
  }
  const jasper = state.reportSettings.jasper || {};
  const jasperEnabledInput = document.getElementById('jasperEnabled');
  if (jasperEnabledInput) {
    jasperEnabledInput.checked = jasper.enabled !== false;
  }
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
      input.checked = value;
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
      input.checked = value;
    }
  });
  const mappings = [
    ['jasperCompiledDir', jasper.compiledDir || ''],
    ['jasperTemplatesDir', jasper.templatesDir || ''],
    ['jasperFontsDir', jasper.fontsDir || ''],
    ['jasperReportName', jasper.defaultReport || ''],
    ['jasperDataSource', jasper.dataSourceName || ''],
    ['jasperJsonQuery', jasper.jsonQuery || '']
  ];
  mappings.forEach(([id, value]) => {
    const input = document.getElementById(id);
    if (input) {
      input.value = value;
    }
  });
  toggleJasperFields(jasper.enabled === false);
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
  updateJasperStatusBadge();
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

function parseCompaniesInput(value) {
  return value
    .split(',')
    .map(item => item.trim())
    .filter(Boolean);
}

function resetUserModal() {
  const form = document.getElementById('userForm');
  form?.reset();
  ['userUsername', 'userPassword', 'userPasswordConfirm', 'userCompanies'].forEach(id => {
    const input = document.getElementById(id);
    input?.classList.remove('is-invalid');
  });
  const usernameInput = document.getElementById('userUsername');
  if (usernameInput) {
    usernameInput.disabled = false;
  }
  state.editingUserId = null;
  setCompaniesFieldDisabled(false);
}

function openUserModal(user = null) {
  resetUserModal();
  const modalTitle = document.getElementById('userModalTitle');
  const usernameInput = document.getElementById('userUsername');
  const nameInput = document.getElementById('userName');
  const allCompaniesInput = document.getElementById('userAllCompanies');
  const companiesInput = document.getElementById('userCompanies');
  const submitButton = document.getElementById('userModalSubmit');

  if (user) {
    state.editingUserId = user.id;
    modalTitle.textContent = 'Editar usuario';
    usernameInput.value = user.usuario;
    usernameInput.disabled = user.usuario === 'admin';
    nameInput.value = user.nombre || '';
    allCompaniesInput.checked = user.empresas === '*';
    if (user.empresas !== '*') {
      companiesInput.value = Array.isArray(user.empresas) ? user.empresas.join(', ') : '';
    }
    submitButton.textContent = 'Actualizar';
  } else {
    modalTitle.textContent = 'Nuevo usuario';
    usernameInput.disabled = false;
    submitButton.textContent = 'Crear usuario';
  }

  setCompaniesFieldDisabled(allCompaniesInput.checked);
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
  const companiesInput = document.getElementById('userCompanies');
  const allCompaniesInput = document.getElementById('userAllCompanies');

  const username = usernameInput.value.trim();
  const nombre = nameInput.value.trim();
  const password = passwordInput.value;
  const confirmPassword = confirmInput.value;
  const allCompanies = allCompaniesInput.checked;
  const companies = parseCompaniesInput(companiesInput.value);

  [usernameInput, passwordInput, confirmInput, companiesInput].forEach(input => input.classList.remove('is-invalid'));

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
    companiesInput.classList.add('is-invalid');
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
  const select = document.getElementById('defaultReportEngine');
  const jasperEnabledInput = document.getElementById('jasperEnabled');
  const payload = {
    defaultEngine: select?.value,
    jasper: {
      enabled: jasperEnabledInput?.checked ?? true,
      compiledDir: getInputValue('jasperCompiledDir'),
      templatesDir: getInputValue('jasperTemplatesDir'),
      fontsDir: getInputValue('jasperFontsDir'),
      defaultReport: getInputValue('jasperReportName'),
      dataSourceName: getInputValue('jasperDataSource'),
      jsonQuery: getInputValue('jasperJsonQuery')
    },
    export: {
      defaultFormat: document.getElementById('exportDefaultFormat')?.value || 'pdf',
      availableFormats: Array.from(document.querySelectorAll('#exportFormatsContainer input[data-format-option]:checked'))
        .map(input => input.value)
    },
    customization: {
      includeCharts: document.getElementById('customIncludeCharts')?.checked ?? true,
      includeMovements: document.getElementById('customIncludeMovements')?.checked ?? true,
      includeObservations: document.getElementById('customIncludeObservations')?.checked ?? true,
      includeUniverse: document.getElementById('customIncludeUniverse')?.checked ?? true,
      csv: {
        includePoResumen: document.getElementById('customCsvIncludePoResumen')?.checked ?? true,
        includeRemisiones: document.getElementById('customCsvIncludeRemisiones')?.checked ?? true,
        includeFacturas: document.getElementById('customCsvIncludeFacturas')?.checked ?? true,
        includeTotales: document.getElementById('customCsvIncludeTotales')?.checked ?? true,
        includeUniverseInfo: document.getElementById('customCsvIncludeUniverseInfo')?.checked ?? true
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
  state.empresas = readSession('porpt-empresas', []);
  if (!Array.isArray(state.empresas) || state.empresas.length === 0) {
    try {
      const response = await fetch('/empresas');
      const data = await response.json();
      if (data.success) {
        state.empresas = data.empresas || [];
      }
    } catch (error) {
      console.error('Error obteniendo empresas:', error);
    }
  }
  renderEmpresaOptions();
  renderUniverseEmpresaSelect();
  renderSelectedPoChips();
  const empresaInput = document.getElementById('empresaSearch');
  empresaInput?.addEventListener('input', event => renderEmpresaOptions(event.target.value));
  empresaInput?.addEventListener('change', event => selectEmpresa(event.target.value));
  const poInput = document.getElementById('poSearch');
  poInput?.addEventListener('input', event => renderPoOptions(event.target.value));
  poInput?.addEventListener('change', event => selectPo(event.target.value));
  poInput?.addEventListener('keydown', event => {
    if (event.key === 'Enter') {
      event.preventDefault();
      selectPo(event.target.value);
    }
  });
  document.getElementById('generateReportBtn')?.addEventListener('click', openReportPreview);
  document.getElementById('clearSelectionBtn')?.addEventListener('click', clearSelectedPos);
  document.getElementById('selectedPoContainer')?.addEventListener('click', event => {
    const button = event.target.closest('button[data-action="remove-po"]');
    if (!button) return;
    const baseId = button.getAttribute('data-po');
    removePoFromSelection(baseId);
  });
  document.getElementById('extensionSelectionContainer')?.addEventListener('change', handleExtensionSelectionChange);
  document.getElementById('universeEmpresaSelect')?.addEventListener('change', event => {
    const value = event.target.value;
    if (!value) {
      state.selectedEmpresa = '';
      state.selectedPoIds = [];
      state.selectedPoDetails.clear();
      destroyCharts();
      renderSelectedPoChips();
      updateUniverseControls();
      const empresaInput = document.getElementById('empresaSearch');
      if (empresaInput) {
        empresaInput.value = '';
      }
      renderUniverseEmpresaSelect();
      document.getElementById('dashboardContent')?.classList.add('d-none');
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
  document.getElementById('jasperEnabled')?.addEventListener('change', event => {
    const enabled = event.target.checked;
    if (state.reportSettings && state.reportSettings.jasper) {
      state.reportSettings.jasper.enabled = enabled;
    }
    toggleJasperFields(!enabled);
    updateJasperStatusBadge();
    updateReportEngineStatus();
  });
  document.getElementById('brandingLetterheadEnabled')?.addEventListener('change', event => {
    toggleBrandingLetterheadFields(!event.target.checked);
  });
  setupPathPickers();
  showAlert('Actualiza los motores de reporte y la identidad visual desde este panel.', 'info', 7000);
}

async function setupAdmin() {
  if (!isAdmin()) {
    showAlert('Solo los administradores pueden acceder a esta sección.', 'danger');
    window.location.href = 'dashboard.html';
    return;
  }
  await loadUsers();
  document.getElementById('loadUsersBtn')?.addEventListener('click', loadUsers);
  document.getElementById('newUserBtn')?.addEventListener('click', () => openUserModal());
  const table = document.getElementById('usersTable');
  table?.addEventListener('click', handleUsersTableClick);
  const userForm = document.getElementById('userForm');
  userForm?.addEventListener('submit', handleUserFormSubmit);
  document.getElementById('userAllCompanies')?.addEventListener('change', event => {
    setCompaniesFieldDisabled(event.target.checked);
  });
  ['userUsername', 'userPassword', 'userPasswordConfirm', 'userCompanies'].forEach(id => {
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
