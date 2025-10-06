require('dotenv').config();
const express = require('express');
const Firebird = require('node-firebird');
const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const bcrypt = require('bcryptjs');
const jasperManager = require('./reports/jasper');
const simplePdf = require('./reports/simple-pdf');
const reportSettingsStore = require('./reports/settings-store');
const { ALLOWED_ENGINES } = reportSettingsStore;

const ADMIN_HEADER = 'x-porpt-admin';
let sqliteDb;
let serverInstance;
let initPromise;

const runtimeBaseDir = __dirname;
const app = express();
const defaultPort = parseInt(process.env.PORT || '3000', 10);
const defaultHost = process.env.HOST || '127.0.0.1';

function getEnvVar(name, defaultValue) {
  const value = process.env[name];
  if (value === undefined || value === '') {
    if (defaultValue !== undefined) {
      console.warn(`Variable de entorno ${name} no definida. Usando default: "${defaultValue}".`);
      return defaultValue;
    }
    throw new Error(`Variable de entorno obligatoria faltante: ${name}`);
  }
  return value;
}

function padTo2(value) {
  return value.toString().padStart(2, '0');
}

function formatFirebirdDate(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  const year = date.getFullYear();
  const month = padTo2(date.getMonth() + 1);
  const day = padTo2(date.getDate());
  return `${year}-${month}-${day}`;
}

function isAdminRequest(req) {
  const headerValue = req.headers?.[ADMIN_HEADER];
  return typeof headerValue === 'string' && headerValue.toLowerCase() === 'true';
}

app.use(express.json());
app.use(express.static(path.join(runtimeBaseDir, 'renderer')));

const baseDir = getEnvVar('BASE_DIR', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');
const baseDirFb = getEnvVar('BASE_DIR_FB', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');
const sqlitePath = getEnvVar('SQLITE_DB', path.join(runtimeBaseDir, 'PERFILES.DB'));

const baseOptions = {
  host: getEnvVar('FB_HOST', 'localhost'),
  port: parseInt(getEnvVar('FB_PORT', '3050'), 10),
  user: getEnvVar('FB_USER', 'SYSDBA'),
  password: getEnvVar('FB_PASSWORD', 'masterkey'),
  lowercase_keys: getEnvVar('FB_LOWERCASE_KEYS', 'false').toLowerCase() === 'true',
  role: getEnvVar('FB_ROLE', ''),
  pageSize: parseInt(getEnvVar('FB_PAGE_SIZE', '4096'), 10),
  charset: getEnvVar('FB_CHARSET', 'ISO8859_1')
};

function queryWithTimeout(target, sql, params = [], timeoutMs = 5000) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error('Timeout en consulta')), timeoutMs);
    target.query(sql, params, (err, result) => {
      clearTimeout(timeout);
      if (err) return reject(err);
      resolve(result);
    });
  });
}

function runAsync(db, sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) return reject(err);
      resolve(this);
    });
  });
}

function getAsync(db, sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) return reject(err);
      resolve(row);
    });
  });
}

function allAsync(db, sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => {
      if (err) return reject(err);
      resolve(rows);
    });
  });
}

app.get('/favicon.ico', (req, res) => res.status(204).end());

app.get('/', (req, res) => {
  res.sendFile(path.join(runtimeBaseDir, 'renderer', 'index.html'));
});

app.post('/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ success: false, message: 'Usuario y contraseña son obligatorios' });
  }
  try {
    const user = await getAsync(sqliteDb, 'SELECT * FROM usuarios WHERE usuario = ?', [username]);
    if (!user) {
      return res.status(401).json({ success: false, message: 'Credenciales inválidas' });
    }
    const match = await bcrypt.compare(password, user.password);
    if (!match) {
      return res.status(401).json({ success: false, message: 'Credenciales inválidas' });
    }
    let empresas = [];
    const isAdmin = user.usuario === 'admin' || user.empresas === '*';
    if (isAdmin) {
      try {
        const dirs = await fs.promises.readdir(baseDir);
        empresas = dirs
          .filter(dir => dir.match(/^Empresa\d+$/))
          .sort((a, b) => parseInt(a.replace('Empresa', '')) - parseInt(b.replace('Empresa', '')));
      } catch (e) {
        return res.status(500).json({ success: false, message: 'Error listando empresas: ' + e.message });
      }
    } else {
      try {
        empresas = user.empresas ? JSON.parse(user.empresas) : [];
        if (Array.isArray(empresas) && empresas.includes('*')) {
          const dirs = await fs.promises.readdir(baseDir);
          empresas = dirs
            .filter(dir => dir.match(/^Empresa\d+$/))
            .sort((a, b) => parseInt(a.replace('Empresa', '')) - parseInt(b.replace('Empresa', '')));
        }
      } catch {
        empresas = [];
      }
    }
    res.json({ success: true, message: 'Login exitoso', empresas, isAdmin });
  } catch (err) {
    console.error('Error consultando usuario:', err);
    res.status(500).json({ success: false, message: 'Error consultando usuario' });
  }
});

app.get('/users', async (req, res) => {
  if (!isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: 'Acceso restringido a administradores' });
  }
  try {
    const users = await allAsync(sqliteDb, 'SELECT id, usuario, nombre, empresas FROM usuarios');
    const mapped = users.map(u => ({
      id: u.id.toString(),
      usuario: u.usuario,
      nombre: u.nombre,
      empresas: u.empresas === '*' ? '*' : JSON.parse(u.empresas || '[]')
    }));
    res.json({ success: true, users: mapped });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error listando usuarios' });
  }
});

app.post('/users', async (req, res) => {
  if (!isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: 'Acceso restringido a administradores' });
  }
  const { usuario, password, nombre, empresas } = req.body;
  const empresasPayload = empresas === '*'
    ? '*'
    : Array.isArray(empresas)
      ? JSON.stringify(empresas)
      : null;
  if (!usuario || !password || empresasPayload === null) {
    return res.status(400).json({ success: false, message: 'Datos inválidos' });
  }
  try {
    const hash = await bcrypt.hash(password, 10);
    const result = await runAsync(sqliteDb, 'INSERT INTO usuarios (usuario, password, nombre, empresas) VALUES (?,?,?,?)', [
      usuario,
      hash,
      nombre || null,
      empresasPayload
    ]);
    res.json({ success: true, id: result.lastID });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error creando usuario: ' + err.message });
  }
});

app.put('/users/:id', async (req, res) => {
  if (!isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: 'Acceso restringido a administradores' });
  }
  const { id } = req.params;
  const { usuario, password, nombre, empresas } = req.body;
  try {
    const existing = await getAsync(sqliteDb, 'SELECT * FROM usuarios WHERE id = ?', [id]);
    if (!existing) return res.status(404).json({ success: false, message: 'Usuario no encontrado' });
    const newUsuario = usuario || existing.usuario;
    const newNombre = nombre !== undefined ? nombre : existing.nombre;
    const newEmpresas = empresas === '*'
      ? '*'
      : Array.isArray(empresas)
        ? JSON.stringify(empresas)
        : existing.empresas;
    const newPassword = password ? await bcrypt.hash(password, 10) : existing.password;
    await runAsync(sqliteDb, 'UPDATE usuarios SET usuario = ?, password = ?, nombre = ?, empresas = ? WHERE id = ?', [
      newUsuario,
      newPassword,
      newNombre,
      newEmpresas,
      id
    ]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error actualizando usuario: ' + err.message });
  }
});

app.delete('/users/:id', async (req, res) => {
  if (!isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: 'Acceso restringido a administradores' });
  }
  const { id } = req.params;
  try {
    const result = await runAsync(sqliteDb, 'DELETE FROM usuarios WHERE id = ?', [id]);
    if (result.changes === 0) return res.status(404).json({ success: false, message: 'Usuario no encontrado' });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error eliminando usuario: ' + err.message });
  }
});

app.get('/empresas', async (req, res) => {
  try {
    const dirs = await fs.promises.readdir(baseDir);
    const empresas = dirs
      .filter(dir => dir.match(/^Empresa\d+$/))
      .sort((a, b) => parseInt(a.replace('Empresa', '')) - parseInt(b.replace('Empresa', '')));
    if (empresas.length === 0) {
      return res.status(404).json({ success: false, message: 'No se encontraron empresas' });
    }
    res.json({ success: true, empresas });
  } catch (err) {
    console.error('Error listando empresas:', err);
    res.status(500).json({ success: false, message: 'Error listando empresas: ' + err.message });
  }
});

function getDatabasePaths(empresa) {
  const num = empresa.replace('Empresa', '').padStart(2, '0');
  const fsPath = path.join(baseDir, empresa, 'Datos', `SAE90EMPRE${num}.FDB`);
  const fbPath = path.win32.join(baseDirFb, empresa, 'Datos', `SAE90EMPRE${num}.FDB`);
  return { fsPath, fbPath };
}

function getCompanyTables(empresa) {
  const num = empresa.replace('Empresa', '').padStart(2, '0');
  return {
    num,
    FACTP: `FACTP${num}`,
    PAR_FACTP: `PAR_FACTP${num}`,
    FACTR: `FACTR${num}`,
    PAR_FACTR: `PAR_FACTR${num}`,
    FACTF: `FACTF${num}`,
    PAR_FACTF: `PAR_FACTF${num}`,
    CLIE: `CLIE${num}`
  };
}

function basePoId(poId) {
  if (!poId) return '';
  return poId.replace(/-\d+$/u, '');
}

async function withFirebirdConnection(empresa, handler) {
  const { fsPath, fbPath } = getDatabasePaths(empresa);
  if (!fs.existsSync(fsPath)) {
    throw new Error(`Base de datos no encontrada para ${empresa}`);
  }
  const options = { ...baseOptions, database: fbPath };
  return await new Promise((resolve, reject) => {
    Firebird.attach(options, async (err, db) => {
      if (err) {
        return reject(err);
      }
      try {
        const tables = getCompanyTables(empresa);
        const result = await handler(db, tables);
        resolve(result);
      } catch (error) {
        reject(error);
      } finally {
        try {
          db.detach();
        } catch (detachError) {
          console.warn('Error al cerrar conexión Firebird:', detachError.message);
        }
      }
    });
  });
}

function buildAlert(message, type = 'info') {
  return { message, type };
}

function calculateTotals(total, totalRem, totalFac) {
  const consumo = totalRem + totalFac;
  const restante = Math.max(total - consumo, 0);
  const porcRem = total > 0 ? (totalRem / total) * 100 : 0;
  const porcFac = total > 0 ? (totalFac / total) * 100 : 0;
  const porcRest = Math.max(100 - (porcRem + porcFac), 0);
  return {
    total,
    totalRem,
    totalFac,
    totalConsumo: consumo,
    restante,
    porcRem,
    porcFac,
    porcRest
  };
}

function extractEmpresaNumero(empresa) {
  if (typeof empresa !== 'string') return '';
  const match = empresa.match(/(\d+)/u);
  return match ? match[1] : '';
}

function buildEmpresaLabel(empresa) {
  const numero = extractEmpresaNumero(empresa);
  return numero ? `Empresa ${numero}` : empresa;
}

function normalizePoId(value) {
  if (typeof value !== 'string') return '';
  return value.trim();
}

function uniqueBasePoIds(ids = []) {
  const seen = new Set();
  ids.forEach(id => {
    const normalized = normalizePoId(id);
    if (!normalized) return;
    const base = basePoId(normalized);
    if (base) {
      seen.add(base);
    }
  });
  return Array.from(seen);
}

async function getPoSummary(empresa, poId) {
  const targetPo = (poId || '').trim();
  if (!targetPo) {
    throw new Error('PO inválida');
  }
  return await withFirebirdConnection(empresa, async (db, tables) => {
    const baseId = basePoId(targetPo);
    const poRows = await queryWithTimeout(
      db,
      `SELECT TRIM(f.CVE_DOC) AS id, f.FECHA_DOC AS fecha, COALESCE(SUM(p.TOT_PARTIDA), 0) AS total
       FROM ${tables.FACTP} f
       LEFT JOIN ${tables.PAR_FACTP} p ON f.CVE_DOC = p.CVE_DOC
       WHERE f.STATUS <> 'C' AND (TRIM(f.CVE_DOC) = ? OR TRIM(f.CVE_DOC) LIKE ?)
       GROUP BY f.CVE_DOC, f.FECHA_DOC
       ORDER BY TRIM(f.CVE_DOC)`,
      [targetPo, `${baseId}-%`]
    );
    if (poRows.length === 0) {
      throw new Error(`No se encontró información para el PO ${targetPo}`);
    }
    const poIds = [...new Set(poRows.map(row => (row.ID || row.id || '').trim()))];
    const placeholders = poIds.map(() => '?').join(',');
    let remRows = [];
    let factRows = [];
    if (poIds.length > 0) {
      remRows = await queryWithTimeout(
        db,
        `SELECT TRIM(r.CVE_DOC) AS id, TRIM(r.CVE_PEDI) AS po, r.FECHA_DOC AS fecha, COALESCE(SUM(pr.TOT_PARTIDA), 0) AS monto
         FROM ${tables.FACTR} r
         LEFT JOIN ${tables.PAR_FACTR} pr ON r.CVE_DOC = pr.CVE_DOC
         WHERE r.STATUS <> 'C' AND TRIM(r.CVE_PEDI) IN (${placeholders})
         GROUP BY r.CVE_DOC, r.CVE_PEDI, r.FECHA_DOC
         ORDER BY TRIM(r.CVE_DOC)`,
        poIds
      );
      factRows = await queryWithTimeout(
        db,
        `SELECT TRIM(f.CVE_DOC) AS id, TRIM(f.CVE_PEDI) AS po, f.FECHA_DOC AS fecha, COALESCE(SUM(pf.TOT_PARTIDA), 0) AS monto
         FROM ${tables.FACTF} f
         LEFT JOIN ${tables.PAR_FACTF} pf ON f.CVE_DOC = pf.CVE_DOC
         WHERE f.STATUS <> 'C' AND TRIM(f.CVE_PEDI) IN (${placeholders})
         GROUP BY f.CVE_DOC, f.CVE_PEDI, f.FECHA_DOC
         ORDER BY TRIM(f.CVE_DOC)`,
        poIds
      );
    }

    const remisionesMap = remRows.reduce((acc, row) => {
      const key = (row.PO || row.po || '').trim();
      if (!acc[key]) acc[key] = [];
      acc[key].push({
        id: (row.ID || row.id || '').trim(),
        fecha: formatFirebirdDate(row.FECHA || row.fecha),
        monto: Number(row.MONTO ?? row.monto ?? 0)
      });
      return acc;
    }, {});

    const facturasMap = factRows.reduce((acc, row) => {
      const key = (row.PO || row.po || '').trim();
      if (!acc[key]) acc[key] = [];
      acc[key].push({
        id: (row.ID || row.id || '').trim(),
        fecha: formatFirebirdDate(row.FECHA || row.fecha),
        monto: Number(row.MONTO ?? row.monto ?? 0)
      });
      return acc;
    }, {});

    const items = poRows.map(row => {
      const id = (row.ID || row.id || '').trim();
      const fecha = formatFirebirdDate(row.FECHA || row.fecha);
      const total = Number(row.TOTAL ?? row.total ?? 0);
      const remisiones = (remisionesMap[id] || []).map(rem => ({
        ...rem,
        porcentaje: total > 0 ? (rem.monto / total) * 100 : 0
      }));
      const facturas = (facturasMap[id] || []).map(fac => ({
        ...fac,
        porcentaje: total > 0 ? (fac.monto / total) * 100 : 0
      }));
      const totalRem = remisiones.reduce((sum, item) => sum + item.monto, 0);
      const totalFac = facturas.reduce((sum, item) => sum + item.monto, 0);
      const totals = calculateTotals(total, totalRem, totalFac);
      const alerts = [];
      if (totals.totalConsumo >= total * 0.1 && total > 0) {
        alerts.push(buildAlert(`El consumo del PO ${id} ha alcanzado el ${(totals.totalConsumo / total * 100).toFixed(2)}%`, 'warning'));
      }
      const remisionesTexto = remisiones.length
        ? remisiones
            .map(rem => `${rem.id} • ${formatFirebirdDate(rem.fecha)} • $${rem.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${rem.porcentaje.toFixed(2)}%)`)
            .join('\n')
        : 'Sin remisiones registradas';
      const facturasTexto = facturas.length
        ? facturas
            .map(fac => `${fac.id} • ${formatFirebirdDate(fac.fecha)} • $${fac.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${fac.porcentaje.toFixed(2)}%)`)
            .join('\n')
        : 'Sin facturas registradas';
      const alertasTexto = alerts.length
        ? alerts.map(alerta => `[${alerta.type.toUpperCase()}] ${alerta.message}`).join('\n')
        : 'Sin alertas';
      return {
        id,
        baseId: basePoId(id),
        fecha,
        total,
        remisiones,
        facturas,
        remisionesTexto,
        facturasTexto,
        alertasTexto,
        totals,
        alerts
      };
    });

    const aggregatedTotals = items.reduce(
      (acc, item) => {
        acc.total += item.total;
        acc.totalRem += item.totals.totalRem;
        acc.totalFac += item.totals.totalFac;
        return acc;
      },
      { total: 0, totalRem: 0, totalFac: 0 }
    );
    const totals = calculateTotals(aggregatedTotals.total, aggregatedTotals.totalRem, aggregatedTotals.totalFac);
    const totalsTexto = `Total grupo: $${totals.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}\n` +
      `Remisiones: $${totals.totalRem.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRem.toFixed(2)}%)\n` +
      `Facturas: $${totals.totalFac.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcFac.toFixed(2)}%)\n` +
      `Restante: $${totals.restante.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRest.toFixed(2)}%)`;
    const alerts = items.flatMap(item => item.alerts);
    if (totals.totalConsumo >= totals.total * 0.1 && totals.total > 0) {
      alerts.push(buildAlert(`El consumo total del grupo ${baseId} supera el 10% (${(totals.totalConsumo / totals.total * 100).toFixed(2)}%)`, 'warning'));
    }
    const alertasTexto = alerts.length
      ? alerts.map(alerta => `[${alerta.type.toUpperCase()}] ${alerta.message}`).join('\n')
      : 'Sin alertas generales';
    const empresaLabel = buildEmpresaLabel(empresa);
    return {
      empresa,
      empresaNumero: extractEmpresaNumero(empresa),
      empresaLabel,
      companyName: empresaLabel,
      baseId,
      selectedId: targetPo,
      selectedIds: [baseId],
      totals,
      totalsTexto,
      items,
      alerts,
      alertasTexto
    };
  });
}

async function getPoSummaryGroup(empresa, poIds) {
  const normalizedIds = uniqueBasePoIds(poIds);
  if (normalizedIds.length === 0) {
    throw new Error('Selecciona al menos una PO válida.');
  }
  const summaries = [];
  for (const id of normalizedIds) {
    const summary = await getPoSummary(empresa, id);
    summaries.push(summary);
  }
  if (summaries.length === 1) {
    return summaries[0];
  }

  const firstSummary = summaries[0];
  const itemsMap = new Map();
  summaries.forEach(summary => {
    summary.items.forEach(item => {
      itemsMap.set(item.id, item);
    });
  });
  const items = Array.from(itemsMap.values()).sort((a, b) => a.id.localeCompare(b.id));
  const aggregated = items.reduce(
    (acc, item) => {
      acc.total += item.total;
      acc.totalRem += item.totals.totalRem;
      acc.totalFac += item.totals.totalFac;
      return acc;
    },
    { total: 0, totalRem: 0, totalFac: 0 }
  );
  const totals = calculateTotals(aggregated.total, aggregated.totalRem, aggregated.totalFac);
  const totalsTexto =
    `Total combinado: $${totals.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}\n` +
    `Remisiones: $${totals.totalRem.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRem.toFixed(2)}%)\n` +
    `Facturas: $${totals.totalFac.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcFac.toFixed(2)}%)\n` +
    `Restante: $${totals.restante.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRest.toFixed(2)}%)`;

  const alerts = summaries.flatMap(summary => summary.alerts || []);
  if (totals.total > 0 && totals.totalConsumo >= totals.total * 0.1) {
    alerts.push(
      buildAlert(
        `El consumo total combinado supera el 10% (${((totals.totalConsumo / totals.total) * 100).toFixed(2)}%)`,
        'warning'
      )
    );
  }
  const alertasTexto = alerts.length
    ? alerts.map(alerta => `[${(alerta.type || 'info').toUpperCase()}] ${alerta.message}`).join('\n')
    : 'Sin alertas generales';

  const selectedIds = Array.from(new Set(summaries.flatMap(summary => summary.selectedIds || [summary.baseId]))).filter(Boolean);
  const selectedTargets = Array.from(new Set(summaries.map(summary => summary.selectedId).filter(Boolean)));

  return {
    empresa: firstSummary.empresa,
    empresaNumero: firstSummary.empresaNumero,
    empresaLabel: firstSummary.empresaLabel,
    companyName: firstSummary.companyName,
    baseId: null,
    selectedId: selectedIds.join(', '),
    selectedIds,
    selectedTargets,
    totals,
    totalsTexto,
    items,
    alerts,
    alertasTexto
  };
}

async function tableExists(db, tableName) {
  try {
    const rows = await queryWithTimeout(
      db,
      "SELECT 1 AS CNT FROM RDB$RELATIONS WHERE TRIM(UPPER(RDB$RELATION_NAME)) = UPPER(?)",
      [tableName]
    );
    return rows.length > 0;
  } catch (err) {
    console.warn(`Error verificando existencia de tabla ${tableName}: ${err.message}`);
    return false;
  }
}

app.get('/pos/:empresa', async (req, res) => {
  const empresa = req.params.empresa;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const rows = await withFirebirdConnection(empresa, async (db, tables) => {
      const result = await queryWithTimeout(db, `
        SELECT TRIM(f.CVE_DOC) AS id, f.FECHA_DOC AS fecha, COALESCE(SUM(p.TOT_PARTIDA), 0) AS total
        FROM ${tables.FACTP} f
        LEFT JOIN ${tables.PAR_FACTP} p ON f.CVE_DOC = p.CVE_DOC
        WHERE f.STATUS <> 'C'
        GROUP BY f.CVE_DOC, f.FECHA_DOC
        ORDER BY f.FECHA_DOC DESC
      `);
      return result.map(row => ({
        id: (row.ID || row.id || '').trim(),
        fecha: formatFirebirdDate(row.FECHA || row.fecha),
        total: Number(row.TOTAL ?? row.total ?? 0),
        baseId: basePoId((row.ID || row.id || '').trim()),
        isExtension: basePoId((row.ID || row.id || '').trim()) !== (row.ID || row.id || '').trim()
      }));
    });
    res.json({ success: true, pos: rows });
  } catch (err) {
    const status = err.message && err.message.includes('Base de datos no encontrada') ? 404 : 500;
    res.status(status).json({ success: false, message: 'Error consultando POs: ' + err.message });
  }
});

app.get('/rems/:empresa/:poId', async (req, res) => {
  const { empresa, poId } = req.params;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const summary = await getPoSummary(empresa, poId);
    const rems = summary.items.flatMap(item => item.remisiones.map(rem => ({
      ...rem,
      poId: item.id
    })));
    res.json({ success: true, rems });
  } catch (err) {
    const status = err.message && err.message.includes('No se encontró') ? 404 : 500;
    res.status(status).json({ success: false, message: 'Error consultando remisiones: ' + err.message });
  }
});

app.get('/facts/:empresa/:poId', async (req, res) => {
  const { empresa, poId } = req.params;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const summary = await getPoSummary(empresa, poId);
    const facts = summary.items.flatMap(item => item.facturas.map(fac => ({
      ...fac,
      poId: item.id
    })));
    res.json({ success: true, facts });
  } catch (err) {
    const status = err.message && err.message.includes('No se encontró') ? 404 : 500;
    res.status(status).json({ success: false, message: 'Error consultando facturas: ' + err.message });
  }
});

app.post('/po-summary', async (req, res) => {
  const { empresa, poIds } = req.body || {};
  if (!empresa || !empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const summary = await getPoSummaryGroup(empresa, Array.isArray(poIds) ? poIds : []);
    res.json({ success: true, summary });
  } catch (err) {
    const status = err.message && err.message.includes('PO') ? 404 : 500;
    res.status(status).json({ success: false, message: err.message });
  }
});

app.get('/po-summary/:empresa/:poId', async (req, res) => {
  const { empresa, poId } = req.params;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const summary = await getPoSummaryGroup(empresa, [poId]);
    res.json({ success: true, summary });
  } catch (err) {
    const status = err.message && err.message.includes('No se encontró') ? 404 : 500;
    res.status(status).json({ success: false, message: err.message });
  }
});

app.get('/report-settings', async (req, res) => {
  try {
    const settings = reportSettingsStore.getSettings();
    const jasperEnabled = settings.jasper.enabled !== false;
    const jasperAvailable = jasperEnabled && jasperManager.isAvailable();
    const engines = [
      {
        id: 'jasper',
        label: 'JasperReports',
        description: 'Usa plantillas JRXML y node-jasper para generar reportes.',
        available: jasperAvailable
      },
      {
        id: 'simple-pdf',
        label: 'PDF directo',
        description: 'Genera un PDF resumido sin depender de Jasper.',
        available: simplePdf.isAvailable()
      }
    ];
    res.json({
      success: true,
      settings: {
        ...settings,
        jasper: {
          ...settings.jasper,
          available: jasperAvailable
        }
      },
      engines
    });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error leyendo configuración de reportes: ' + err.message });
  }
});

app.put('/report-settings', async (req, res) => {
  if (!isAdminRequest(req)) {
    return res.status(403).json({ success: false, message: 'Acceso restringido a administradores' });
  }
  try {
    const payload = req.body || {};
    const updates = {};

    if (payload.defaultEngine && ALLOWED_ENGINES.includes(payload.defaultEngine)) {
      updates.defaultEngine = payload.defaultEngine;
    }

    if (payload.jasper) {
      const jasper = {};
      if (typeof payload.jasper.enabled === 'boolean') jasper.enabled = payload.jasper.enabled;
      const keys = ['compiledDir', 'templatesDir', 'fontsDir', 'defaultReport', 'dataSourceName', 'jsonQuery'];
      keys.forEach(key => {
        if (payload.jasper[key] !== undefined) {
          jasper[key] = payload.jasper[key];
        }
      });
      updates.jasper = jasper;
    }

    if (payload.branding) {
      const branding = {};
      const stringKeys = ['headerTitle', 'headerSubtitle', 'footerText', 'letterheadTop', 'letterheadBottom', 'remColor', 'facColor', 'restanteColor', 'accentColor'];
      stringKeys.forEach(key => {
        if (payload.branding[key] !== undefined) {
          branding[key] = payload.branding[key];
        }
      });
      if (payload.branding.letterheadEnabled !== undefined) {
        branding.letterheadEnabled = !!payload.branding.letterheadEnabled;
      }
      updates.branding = branding;
    }

    const updatedSettings = reportSettingsStore.updateSettings(updates);
    if (updates.jasper) {
      await jasperManager.reload();
    }

    const jasperEnabled = updatedSettings.jasper.enabled !== false;
    const jasperAvailable = jasperEnabled && jasperManager.isAvailable();
    const engines = [
      {
        id: 'jasper',
        label: 'JasperReports',
        description: 'Usa plantillas JRXML y node-jasper para generar reportes.',
        available: jasperAvailable
      },
      {
        id: 'simple-pdf',
        label: 'PDF directo',
        description: 'Genera un PDF resumido sin depender de Jasper.',
        available: simplePdf.isAvailable()
      }
    ];

    res.json({ success: true, settings: { ...updatedSettings, jasper: { ...updatedSettings.jasper, available: jasperAvailable } }, engines });
  } catch (err) {
    console.error('Error actualizando configuración de reportes:', err);
    res.status(500).json({ success: false, message: 'No se pudo actualizar la configuración de reportes: ' + err.message });
  }
});

app.post('/report', async (req, res) => {
  const { empresa, poId, poIds, engine: requestedEngine } = req.body || {};
  const ids = Array.isArray(poIds) && poIds.length ? poIds : poId ? [poId] : [];
  if (!empresa || ids.length === 0) {
    return res.status(400).json({ success: false, message: 'Empresa y PO son obligatorios para el reporte' });
  }
  try {
    const summary = await getPoSummaryGroup(empresa, ids);
    const settings = reportSettingsStore.getSettings();
    const engine = ALLOWED_ENGINES.includes(requestedEngine) ? requestedEngine : settings.defaultEngine;

    if (engine === 'simple-pdf') {
      if (!simplePdf.isAvailable()) {
        return res.status(503).json({
          success: false,
          message: simplePdf.getUnavailableMessage()
        });
      }
      const buffer = await simplePdf.generate(summary, settings.branding || {});
      res.set('Content-Type', 'application/pdf');
      const filenameBase = summary.selectedIds && summary.selectedIds.length
        ? summary.selectedIds.join('_')
        : summary.baseId || 'reporte';
      res.set('Content-Disposition', `attachment; filename=${filenameBase}.pdf`);
      return res.send(buffer);
    }

    if (settings.jasper.enabled === false) {
      return res.status(503).json({
        success: false,
        message: 'JasperReports está deshabilitado en la configuración.'
      });
    }

    const jasper = jasperManager.getInstance();
    if (!jasper) {
      return res.status(503).json({
        success: false,
        message: 'JasperReports no se inicializó. Verifica la instalación y la configuración.'
      });
    }

    const buffer = await jasperManager.generatePoSummary(jasper, summary);
    res.set('Content-Type', 'application/pdf');
    const filenameBase = summary.selectedIds && summary.selectedIds.length
      ? summary.selectedIds.join('_')
      : summary.baseId || 'reporte';
    res.set('Content-Disposition', `attachment; filename=${filenameBase}.pdf`);
    res.send(buffer);
  } catch (err) {
    console.error('Error generando reporte:', err);
    const status = err.message && err.message.includes('PO') ? 404 : 500;
    res.status(status).json({ success: false, message: 'Error generando reporte: ' + err.message });
  }
});

app.use((err, req, res, next) => {
  console.error('Error en el servidor:', err);
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({ success: false, message: err.message || 'Error interno del servidor' });
});

async function init(options = {}) {
  if (initPromise) return initPromise;
  const listenPort = options.port ?? defaultPort;
  const listenHost = options.host || defaultHost;
  initPromise = (async () => {
    sqliteDb = new sqlite3.Database(sqlitePath);
    try {
      await runAsync(sqliteDb, `CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        nombre TEXT,
        empresas TEXT NOT NULL
      )`);
      const adminUser = 'admin';
      const adminPass = "569OpEGvwh'8";
      const admin = await getAsync(sqliteDb, 'SELECT id FROM usuarios WHERE usuario = ?', [adminUser]);
      if (!admin) {
        const hash = await bcrypt.hash(adminPass, 10);
        await runAsync(sqliteDb, 'INSERT INTO usuarios (usuario, password, nombre, empresas) VALUES (?,?,?,?)', [
          adminUser, hash, null, '*'
        ]);
      }
      reportSettingsStore.loadSettings();
      await jasperManager.init();
      serverInstance = await new Promise((resolve, reject) => {
        const server = app.listen(listenPort, listenHost, () => {
          console.log(`Servidor corriendo en puerto ${listenPort}`);
          resolve(server);
        });
        server.on('error', reject);
      });
      return serverInstance;
    } catch (err) {
      if (sqliteDb) sqliteDb.close();
      throw err;
    }
  })();
  return await initPromise;
}

async function shutdown() {
  if (serverInstance) {
    await new Promise((resolve, reject) => serverInstance.close(error => (error ? reject(error) : resolve())));
    serverInstance = null;
  }
  if (sqliteDb) {
    await new Promise((resolve, reject) => sqliteDb.close(error => (error ? reject(error) : resolve())));
    sqliteDb = null;
  }
  initPromise = null;
}

if (require.main === module) {
  init().catch(err => {
    console.error('Error inicializando la aplicación:', err);
    process.exit(1);
  });
}

module.exports = { init, shutdown, app };