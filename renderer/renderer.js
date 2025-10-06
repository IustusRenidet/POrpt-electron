const state = {
  empresas: [],
  pos: [],
  selectedEmpresa: '',
  selectedPoIds: [],
  charts: new Map(),
  summary: null,
  users: [],
  reportSettings: null,
  reportEngines: [],
  selectedEngine: '',
  editingUserId: null,
  universeFilterMode: 'global',
  universeStartDate: '',
  universeEndDate: ''
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

function destroyCharts() {
  state.charts.forEach(chart => {
    if (chart && typeof chart.destroy === 'function') {
      chart.destroy();
    }
  });
  state.charts.clear();
}

function formatCurrency(value) {
  return value.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
    if (help) {
      help.textContent = 'Puedes agregar varias POs base (el sistema combinará sus extensiones automáticamente).';
    }
    return;
  }
  if (help) {
    const label = state.selectedPoIds.length === 1
      ? '1 PEO base seleccionada'
      : `${state.selectedPoIds.length} PEOs base seleccionadas`;
    help.textContent = `${label}. Puedes quitar cualquiera dando clic en la “x”.`;
  }
  state.selectedPoIds.forEach(baseId => {
    const chip = document.createElement('span');
    chip.className = 'selected-po-chip';
    const meta = getPoMetadataByBase(baseId);
    const label = meta
      ? `${baseId} • $${formatCurrency(meta.total || 0)}`
      : baseId;
    chip.innerHTML = `
      <span>${label}</span>
      <button type="button" aria-label="Quitar" data-action="remove-po" data-po="${baseId}">&times;</button>
    `;
    container.appendChild(chip);
  });
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
    showAlert(`La PO ${baseId} ya está en la selección.`, 'info');
    return;
  }
  state.selectedPoIds.push(baseId);
  showAlert(`PO ${baseId} agregada a la selección.`, 'success');
  renderSelectedPoChips();
  updateSummary();
}

function removePoFromSelection(baseId) {
  const index = state.selectedPoIds.indexOf(baseId);
  if (index === -1) return;
  state.selectedPoIds.splice(index, 1);
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
  document.getElementById('poSearch').value = '';
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
  const totals = summary.totals;
  const selectionCount = (summary.selectedIds || []).length || 0;
  const selectionLabel = selectionCount === 1
    ? '1 PEO base (extensiones incluidas)'
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
          <p class="text-muted mb-0">${totals.porcRem.toFixed(2)}% Remisiones · ${totals.porcFac.toFixed(2)}% Facturas</p>
        </div>
      </div>
    </div>
    <div class="col-md-4">
      <div class="card shadow-sm h-100 border-success border-2">
        <div class="card-body">
          <h5 class="card-title text-success">Disponible</h5>
          <p class="display-6 fw-bold">$${formatCurrency(totals.restante)}</p>
          <p class="text-muted mb-0">${totals.porcRest.toFixed(2)}% restante del presupuesto autorizado.</p>
        </div>
      </div>
    </div>
  `;
}

function renderTable(summary) {
  const table = document.getElementById('poTable');
  if (!table) return;
  const rows = summary.items
    .map(item => `
      <tr data-po="${item.id}">
        <td class="fw-semibold">${item.id}</td>
        <td>${item.fecha || '-'}</td>
        <td>$${formatCurrency(item.total || 0)}</td>
        <td>$${formatCurrency(item.totals.totalRem)} (${item.totals.porcRem.toFixed(2)}%)</td>
        <td>$${formatCurrency(item.totals.totalFac)} (${item.totals.porcFac.toFixed(2)}%)</td>
        <td>$${formatCurrency(item.totals.restante)} (${item.totals.porcRest.toFixed(2)}%)</td>
        <td>
          <button class="btn btn-outline-primary btn-sm" data-po="${item.id}" data-action="show-modal">Detalle</button>
        </td>
      </tr>
    `)
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
    modalBody.innerHTML = `
      <div class="mb-3">
        <h6 class="text-primary">Remisiones</h6>
        <pre class="bg-light rounded p-3 small">${item.remisionesTexto}</pre>
      </div>
      <div class="mb-3">
        <h6 class="text-danger">Facturas</h6>
        <pre class="bg-light rounded p-3 small">${item.facturasTexto}</pre>
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
  const items = summary.items;
  if (!items.length) return;
  const remLabels = items.map(item => item.id);
  const remData = items.map(item => Number(item.totals.totalRem.toFixed(2)));
  const facData = items.map(item => Number(item.totals.totalFac.toFixed(2)));
  const ctxRem = document.getElementById('chartRem');
  const ctxFac = document.getElementById('chartFac');
  const ctxStack = document.getElementById('chartJunto');
  if (ctxRem) {
    state.charts.set('rem', new Chart(ctxRem.getContext('2d'), {
      type: 'bar',
      data: {
        labels: remLabels,
        datasets: [{
          label: 'Remisiones ($)',
          data: remData,
          backgroundColor: '#2563eb'
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        scales: { y: { beginAtZero: true } }
      }
    }));
  }
  if (ctxFac) {
    state.charts.set('fac', new Chart(ctxFac.getContext('2d'), {
      type: 'bar',
      data: {
        labels: remLabels,
        datasets: [{
          label: 'Facturas ($)',
          data: facData,
          backgroundColor: '#f97316'
        }]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        scales: { y: { beginAtZero: true } }
      }
    }));
  }
  if (ctxStack) {
    state.charts.set('stack', new Chart(ctxStack.getContext('2d'), {
      type: 'bar',
      data: {
        labels: remLabels,
        datasets: [
          {
            label: 'Remisiones %',
            data: items.map(item => Number(item.totals.porcRem.toFixed(2))),
            backgroundColor: '#2563eb'
          },
          {
            label: 'Facturas %',
            data: items.map(item => Number(item.totals.porcFac.toFixed(2))),
            backgroundColor: '#f97316'
          },
          {
            label: 'Disponible %',
            data: items.map(item => Number(item.totals.porcRest.toFixed(2))),
            backgroundColor: '#16a34a'
          }
        ]
      },
      options: {
        responsive: true,
        scales: {
          x: { stacked: true },
          y: { stacked: true, beginAtZero: true, max: 100 }
        }
      }
    }));
  }
  const extensionContainer = document.getElementById('extensionsContainer');
  if (extensionContainer) {
    extensionContainer.innerHTML = '';
    items.forEach((item, index) => {
      const restante = Math.max(item.total - (item.totals.totalRem + item.totals.totalFac), 0);
      const card = document.createElement('div');
      card.className = 'col-md-6 col-lg-4';
      card.innerHTML = `
        <div class="card h-100 shadow-sm">
          <div class="card-body">
            <h5 class="card-title">${item.id}</h5>
            <p class="small text-muted mb-2">$${formatCurrency(item.total)} totales</p>
            <canvas id="extChart-${index}" height="220"></canvas>
          </div>
        </div>
      `;
      extensionContainer.appendChild(card);
      const canvas = card.querySelector('canvas');
      state.charts.set(`ext-${index}`, new Chart(canvas.getContext('2d'), {
        type: 'doughnut',
        data: {
          labels: ['Remisiones', 'Facturas', 'Disponible'],
          datasets: [{
            data: [item.totals.totalRem, item.totals.totalFac, restante],
            backgroundColor: ['#2563eb', '#f97316', '#16a34a']
          }]
        },
        options: {
          plugins: {
            legend: { position: 'bottom' }
          }
        }
      }));
    });
  }
}

function renderAlerts(alerts) {
  alerts.forEach(alert => showAlert(alert.message, alert.type || 'info', 8000));
}

async function updateSummary() {
  if (!state.selectedEmpresa || state.selectedPoIds.length === 0) return;
  try {
    const response = await fetch('/po-summary', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ empresa: state.selectedEmpresa, poIds: state.selectedPoIds })
    });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener el resumen');
    }
    state.summary = data.summary;
    renderSummaryCards(state.summary);
    renderTable(state.summary);
    attachTableHandlers(state.summary);
    renderCharts(state.summary);
    renderAlerts(state.summary.alerts || []);
    showAlert('Dashboard actualizado con la selección de POs.', 'success');
    document.getElementById('dashboardContent')?.classList.remove('d-none');
  } catch (error) {
    console.error('Error actualizando resumen:', error);
    showAlert(error.message || 'No se pudo actualizar el dashboard', 'danger');
  }
}

async function generateReport() {
  if (!state.selectedEmpresa || state.selectedPoIds.length === 0) {
    showAlert('Selecciona primero una empresa y al menos una PO base.', 'warning');
    return;
  }
  try {
    const engine = state.selectedEngine || state.reportSettings?.defaultEngine || 'jasper';
    const response = await fetch('/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ empresa: state.selectedEmpresa, poIds: state.selectedPoIds, engine })
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: 'Error generando reporte' }));
      throw new Error(error.message || 'Error generando reporte');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const filename = state.selectedPoIds.length === 1
      ? `POrpt_${state.selectedPoIds[0]}`
      : `POrpt_${state.selectedPoIds.length}_POs`;
    link.download = `${filename}.pdf`;
    link.click();
    URL.revokeObjectURL(url);
    const selectedEngine = state.reportEngines.find(item => item.id === engine);
    const engineLabel = selectedEngine ? selectedEngine.label : 'Reporte';
    showAlert(`${engineLabel} generado correctamente.`, 'success');
  } catch (error) {
    console.error('Error generando reporte:', error);
    showAlert(error.message || 'No se pudo generar el reporte.', 'danger');
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

  const payload = { empresa: state.selectedEmpresa, filter: { mode } };
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
    link.download = `POrpt_universo_${suffix || 'global'}.pdf`;
    link.click();
    URL.revokeObjectURL(url);
    const label = getUniverseFilterLabel(mode, startDate || '-', endDate || '-');
    showAlert(`Reporte universo (${label}) generado correctamente.`, 'success');
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
    if (!state.selectedEngine || !state.reportEngines.some(engine => engine.id === state.selectedEngine)) {
      state.selectedEngine = state.reportSettings?.defaultEngine || state.reportEngines[0]?.id;
    }
    renderReportEngineSelector();
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
    state.selectedEngine = state.reportSettings.defaultEngine;
    renderReportSettingsForm();
    renderReportEngineSelector();
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
  document.getElementById('generateReportBtn')?.addEventListener('click', generateReport);
  document.getElementById('clearSelectionBtn')?.addEventListener('click', clearSelectedPos);
  document.getElementById('selectedPoContainer')?.addEventListener('click', event => {
    const button = event.target.closest('button[data-action="remove-po"]');
    if (!button) return;
    const baseId = button.getAttribute('data-po');
    removePoFromSelection(baseId);
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
  await loadReportSettings();
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
