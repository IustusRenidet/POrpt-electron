let empresas = [];
let pos = [];
let selectedEmpresa = '';
let selectedPoId = '';
let charts = {};  // Para dispose

// Alertas
function showAlert(msg, type = 'info') {
  const toastHtml = `
    <div class="toast align-items-center text-white bg-${type}" role="alert" aria-live="assertive" aria-atomic="true">
      <div class="d-flex">
        <div class="toast-body">${msg}</div>
        <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
      </div>
    </div>`;
  document.getElementById('alerts').innerHTML += toastHtml;
  new bootstrap.Toast(document.querySelector('#alerts .toast:last-child')).show();
}

// Login
async function login() {
  const username = document.getElementById('username').value;
  const password = document.getElementById('password').value;
  try {
    const res = await fetch('/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const data = await res.json();
    if (data.success) {
      empresas = data.empresas;
      loadEmpresaSelect();
      showAlert('Login exitoso', 'success');
    } else {
      showAlert(data.message, 'danger');
    }
  } catch (err) {
    showAlert('Error de conexión', 'danger');
  }
  showAlert('Instrucción: Usa admin / 569OpEGvwh\'8', 'info');
}

// Cargar empresas en select (search dinámico con datalist)
function loadEmpresaSelect() {
  const select = document.getElementById('empresaSelect');
  select.innerHTML = '<option value="">-- Selecciona --</option>';
  empresas.forEach(emp => {
    const option = document.createElement('option');
    option.value = emp;
    select.appendChild(option);
  });
  document.getElementById('poSection').style.display = 'none';
}

// Cargar POs
async function loadPOs() {
  selectedEmpresa = document.getElementById('empresaSelect').value.split(' - ')[0];  // e.g., "Empresa26"
  if (!selectedEmpresa) return;
  try {
    const res = await fetch(`/pos/${selectedEmpresa}`);
    const data = await res.json();
    if (data.success) {
      pos = data.pos;
      const select = document.getElementById('poSelect');
      select.innerHTML = '<option value="">-- Selecciona PO --</option>';
      pos.forEach(po => {
        const option = document.createElement('option');
        option.value = po.id;
        option.textContent = `${po.id} - ${po.fecha} - $${po.total.toLocaleString()}`;
        select.appendChild(option);
      });
      document.getElementById('poSection').style.display = 'block';
      showAlert('POs cargados', 'success');
    }
  } catch (err) {
    showAlert('Error cargando POs', 'danger');
  }
}

// Actualizar gráficos y tabla
async function updateCharts() {
  selectedPoId = document.getElementById('poSelect').value;
  if (!selectedPoId) return;
  const po = pos.find(p => p.id === selectedPoId);
  if (!po) return;
  try {
    const [resRem, resFac] = await Promise.all([
      fetch(`/rems/${selectedEmpresa}/${selectedPoId}`),
      fetch(`/facts/${selectedEmpresa}/${selectedPoId}`)
    ]);
    const { rems: remData } = await resRem.json();
    const { facts: facData } = await resFac.json();
    if (!remData.success || !facData.success) return showAlert('Error cargando datos', 'danger');

    const totalPO = po.total;
    const totalRem = remData.rems.reduce((s, r) => s + r.monto, 0);
    const totalFac = facData.facts.reduce((s, f) => s + f.monto, 0);
    const totalConsumo = totalRem + totalFac;
    const porcRem = (totalRem / totalPO * 100);
    const porcFac = (totalFac / totalPO * 100);
    const porcRest = 100 - porcRem - porcFac;

    if (totalConsumo > totalPO * 0.1) {
      showAlert(`¡Alerta! Consumo del PO ${selectedPoId} supera el 10% (${porcRest.toFixed(1)}% restante).`, 'warning');
    }

    // Gráficos separados
    charts.chartRem = new Chart(document.getElementById('chartRem').getContext('2d'), {
      type: 'bar',
      data: { labels: remData.rems.map(r => r.id), datasets: [{ label: 'Remisiones', data: remData.rems.map(r => (r.monto / totalPO * 100).toFixed(1)), backgroundColor: 'blue' }] },
      options: { scales: { y: { beginAtZero: true } } }
    });
    charts.chartFac = new Chart(document.getElementById('chartFac').getContext('2d'), {
      type: 'bar',
      data: { labels: facData.facts.map(f => f.id), datasets: [{ label: 'Facturas', data: facData.facts.map(f => (f.monto / totalPO * 100).toFixed(1)), backgroundColor: 'red' }] },
      options: { scales: { y: { beginAtZero: true } } }
    });

    // Gráfico conjunto apilado 100%
    charts.chartJunto = new Chart(document.getElementById('chartJunto').getContext('2d'), {
      type: 'bar',
      data: {
        labels: ['Consumo PO'],
        datasets: [
          { label: 'Remisiones', data: [porcRem], backgroundColor: 'blue' },
          { label: 'Facturas', data: [porcFac], backgroundColor: 'red' },
          { label: 'Restante', data: [porcRest], backgroundColor: 'green' }
        ]
      },
      options: {
        scales: {
          x: { stacked: true },
          y: { stacked: true, max: 100, beginAtZero: true }
        }
      }
    });

    // Tabla PO
    const table = document.getElementById('poTable');
    table.innerHTML = `
      <thead><tr><th>ID</th><th>Fecha</th><th>Total</th><th>#Rem</th><th>#Fac</th><th>%Consumo</th></tr></thead>
      <tbody><tr><td>${po.id}</td><td>${po.fecha}</td><td>$${po.total.toLocaleString()}</td><td>${remData.rems.length}</td><td>${facData.facts.length}</td><td>${((totalConsumo / totalPO * 100).toFixed(1)}%</td></tr></tbody>`;
    table.style.display = 'table';

    // Extensión: Si selectedPoId contiene '-', abre modal con gráfico adicional
    if (selectedPoId.includes('-')) {
      // Llama a /pos con LIKE para base, genera chartExt similar
      const modal = new bootstrap.Modal(document.getElementById('modalExt'));
      modal.show();
      // Lógica para chartExt (placeholder: usa mismo chartJunto)
    }

    showAlert('Gráficos actualizados', 'success');
  } catch (err) {
    showAlert('Error actualizando gráficos', 'danger');
  }
}

// Generar reporte
async function generateReport() {
  if (!selectedPoId) return showAlert('Selecciona PO primero', 'warning');
  const po = pos.find(p => p.id === selectedPoId);
  const ssitel = selectedEmpresa.split(' - ')[1] || 'SSITEL';  // De CLIE si necesitas query extra
  try {
    const [resRem, resFac] = await Promise.all([fetch(`/rems/${selectedEmpresa}/${selectedPoId}`), fetch(`/facts/${selectedEmpresa}/${selectedPoId}`)]);
    const { rems } = await resRem.json();
    const { facts } = await resFac.json();
    const resPdf = await fetch('/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ssitel, po, rems: rems.rems, facts: facts.facts })
    });
    const blob = await resPdf.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `reporte_${selectedPoId}.pdf`;
    a.click();
    showAlert('PDF generado', 'success');
  } catch (err) {
    showAlert('Error generando PDF', 'danger');
  }
}

// Admin: Cargar users
async function loadUsers() {
  try {
    const res = await fetch('/users');
    const data = await res.json();
    if (data.success) {
      const tbody = document.querySelector('#usersTable tbody');
      tbody.innerHTML = data.users.map(u => `
        <tr>
          <td>${u.id}</td><td>${u.usuario}</td><td>${u.nombre || ''}</td><td>${Array.isArray(u.empresas) ? u.empresas.join(', ') : u.empresas}</td>
          <td><button onclick="editUser(${u.id})" class="btn btn-sm btn-warning">Editar</button> <button onclick="deleteUser(${u.id})" class="btn btn-sm btn-danger">Eliminar</button></td>
        </tr>`).join('');
      showAlert('Usuarios cargados', 'success');
    }
  } catch (err) {
    showAlert('Error cargando usuarios', 'danger');
  }
}

// CRUD admin (editUser, deleteUser: fetch PUT/DELETE con form modal, similar a login)

function logout() {
  window.location.href = 'index.html';
}

// Onload dashboard
if (window.location.pathname.includes('dashboard.html')) {
  fetch('/empresas').then(r => r.json()).then(d => {
    if (d.success) {
      empresas = d.empresas;
      loadEmpresaSelect();
    }
  });
}