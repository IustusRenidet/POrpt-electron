const state = {
  empresas: [],
  pos: [],
  selectedEmpresa: '',
  selectedPoId: '',
  charts: new Map(),
  summary: null
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
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value;
  if (!username || !password) {
    showAlert('Ingresa usuario y contraseña.', 'warning');
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

async function selectEmpresa(value) {
  if (!value || !state.empresas.includes(value)) {
    showAlert('Selecciona una empresa válida de la lista.', 'warning');
    return;
  }
  if (state.selectedEmpresa === value) return;
  state.selectedEmpresa = value;
  state.selectedPoId = '';
  document.getElementById('poSearch').value = '';
  document.getElementById('dashboardContent').classList.add('d-none');
  destroyCharts();
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
    document.getElementById('dashboardContent').classList.remove('d-none');
    showAlert('POs disponibles listas. Busca por número, fecha o monto.', 'success');
  } catch (error) {
    console.error('Error cargando POs:', error);
    showAlert(error.message || 'Error consultando POs', 'danger');
  }
}

async function selectPo(poId) {
  if (!poId) return;
  const found = state.pos.find(po => po.id === poId);
  if (!found) {
    showAlert('Selecciona una PO válida de la lista desplegable.', 'warning');
    return;
  }
  state.selectedPoId = poId;
  await updateSummary();
}

function renderSummaryCards(summary) {
  const container = document.getElementById('summaryCards');
  if (!container) return;
  const totals = summary.totals;
  container.innerHTML = `
    <div class="col-md-4">
      <div class="card shadow-sm h-100 border-primary border-2">
        <div class="card-body">
          <h5 class="card-title text-primary">Total PO (Grupo)</h5>
          <p class="display-6 fw-bold">$${formatCurrency(totals.total)}</p>
          <p class="text-muted mb-0">${summary.items.length} documento(s) incluyendo extensiones.</p>
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
          <p class="text-muted mb-0">${totals.porcRest.toFixed(2)}% restante del presupuesto.</p>
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
  if (!state.selectedEmpresa || !state.selectedPoId) return;
  try {
    const response = await fetch(`/po-summary/${state.selectedEmpresa}/${state.selectedPoId}`);
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
    showAlert('Dashboard actualizado con la PO seleccionada.', 'success');
  } catch (error) {
    console.error('Error actualizando resumen:', error);
    showAlert(error.message || 'No se pudo actualizar el dashboard', 'danger');
  }
}

async function generateReport() {
  if (!state.selectedEmpresa || !state.selectedPoId) {
    showAlert('Selecciona primero una empresa y PO.', 'warning');
    return;
  }
  try {
    const response = await fetch('/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ empresa: state.selectedEmpresa, poId: state.selectedPoId })
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: 'Error generando reporte' }));
      throw new Error(error.message || 'Error generando reporte');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `POrpt_${state.selectedPoId}.pdf`;
    link.click();
    URL.revokeObjectURL(url);
    showAlert('Reporte Jasper generado correctamente.', 'success');
  } catch (error) {
    console.error('Error generando reporte:', error);
    showAlert(error.message || 'No se pudo generar el reporte.', 'danger');
  }
}

async function loadUsers() {
  try {
    const response = await fetch('/users');
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No fue posible obtener usuarios');
    }
    const tbody = document.querySelector('#usersTable tbody');
    if (!tbody) return;
    tbody.innerHTML = data.users
      .map(user => `
        <tr>
          <td>${user.id}</td>
          <td>${user.usuario}</td>
          <td>${user.nombre || ''}</td>
          <td>${Array.isArray(user.empresas) ? user.empresas.join(', ') : user.empresas}</td>
          <td>
            <button class="btn btn-sm btn-outline-secondary" data-id="${user.id}" data-action="edit">Editar</button>
            <button class="btn btn-sm btn-outline-danger" data-id="${user.id}" data-action="delete">Eliminar</button>
          </td>
        </tr>
      `)
      .join('');
    showAlert('Usuarios cargados.', 'success');
  } catch (error) {
    console.error('Error cargando usuarios:', error);
    showAlert(error.message || 'No se pudieron cargar los usuarios.', 'danger');
  }
}

async function editUser(id) {
  const nuevoNombre = prompt('Nombre (deja vacío para mantener actual):');
  const nuevasEmpresas = prompt('Empresas asignadas separadas por coma (usa * para todas):');
  const nuevaPassword = prompt('Nueva contraseña (deja vacío para mantener actual):');
  const payload = {};
  if (nuevoNombre !== null && nuevoNombre.trim() !== '') payload.nombre = nuevoNombre.trim();
  if (nuevasEmpresas !== null && nuevasEmpresas.trim() !== '') {
    payload.empresas = nuevasEmpresas.trim() === '*'
      ? '*'
      : nuevasEmpresas.split(',').map(e => e.trim()).filter(Boolean);
  }
  if (nuevaPassword !== null && nuevaPassword.trim() !== '') payload.password = nuevaPassword;
  if (Object.keys(payload).length === 0) {
    showAlert('No se modificó ningún dato.', 'info');
    return;
  }
  try {
    const response = await fetch(`/users/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!data.success) {
      throw new Error(data.message || 'No se pudo actualizar el usuario');
    }
    showAlert('Usuario actualizado.', 'success');
    await loadUsers();
  } catch (error) {
    console.error('Error editando usuario:', error);
    showAlert(error.message || 'Error editando usuario.', 'danger');
  }
}

async function deleteUser(id) {
  if (!confirm('¿Eliminar usuario seleccionado?')) return;
  try {
    const response = await fetch(`/users/${id}`, { method: 'DELETE' });
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

function logout() {
  sessionStorage.removeItem('porpt-empresas');
  sessionStorage.removeItem('porpt-is-admin');
  window.location.href = 'index.html';
}

function setupLoginPage() {
  const form = document.getElementById('loginForm');
  form?.addEventListener('submit', login);
  document.getElementById('loginButton')?.addEventListener('click', login);
  showAlert("Instrucción: Usuario admin, contraseña 569OpEGvwh'8", 'info', 10000);
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
  const empresaInput = document.getElementById('empresaSearch');
  empresaInput?.addEventListener('input', event => renderEmpresaOptions(event.target.value));
  empresaInput?.addEventListener('change', event => selectEmpresa(event.target.value));
  const poInput = document.getElementById('poSearch');
  poInput?.addEventListener('input', event => renderPoOptions(event.target.value));
  poInput?.addEventListener('change', event => selectPo(event.target.value));
  document.getElementById('generateReportBtn')?.addEventListener('click', generateReport);
  showAlert('Selecciona empresa y PO para visualizar el tablero.', 'info', 9000);
}

function setupAdmin() {
  document.getElementById('loadUsersBtn')?.addEventListener('click', loadUsers);
  const table = document.getElementById('usersTable');
  table?.addEventListener('click', event => {
    const action = event.target.getAttribute('data-action');
    const id = event.target.getAttribute('data-id');
    if (!action || !id) return;
    if (action === 'edit') {
      editUser(id);
    }
    if (action === 'delete') {
      deleteUser(id);
    }
  });
  showAlert('Desde aquí administra usuarios y contraseñas.', 'info', 7000);
}

function init() {
  const path = window.location.pathname;
  if (path.endsWith('index.html') || path === '/' || path === '') {
    setupLoginPage();
  } else if (path.endsWith('dashboard.html')) {
    setupDashboard();
  } else if (path.endsWith('admin.html')) {
    setupAdmin();
  }
}

document.addEventListener('DOMContentLoaded', init);

window.login = login;
window.loadUsers = loadUsers;
window.logout = logout;
window.generateReport = generateReport;
window.editUser = editUser;
window.deleteUser = deleteUser;
