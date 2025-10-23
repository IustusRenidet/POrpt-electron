require('dotenv').config();
const express = require('express');
const Firebird = require('node-firebird');
const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const bcrypt = require('bcryptjs');
const simplePdf = require('./reports/simple-pdf');
const exporters = require('./reports/exporters');
const reportSettingsStore = require('./reports/settings-store');
const { ALLOWED_ENGINES, ALLOWED_FORMATS } = reportSettingsStore;

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
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }
    if (/^\d{4}-\d{2}-\d{2}$/u.test(trimmed)) {
      return trimmed;
    }
    if (/^\d{8}$/u.test(trimmed)) {
      const year = trimmed.slice(0, 4);
      const month = trimmed.slice(4, 6);
      const day = trimmed.slice(6, 8);
      return `${year}-${month}-${day}`;
    }
    const isoPrefix = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/u);
    if (isoPrefix) {
      const [, year, month, day] = isoPrefix;
      return `${year}-${month}-${day}`;
    }
    const normalized = trimmed.replace(/ /u, 'T');
    const parsedFromNormalized = new Date(normalized);
    if (!Number.isNaN(parsedFromNormalized.getTime())) {
      const year = parsedFromNormalized.getFullYear();
      const month = padTo2(parsedFromNormalized.getMonth() + 1);
      const day = padTo2(parsedFromNormalized.getDate());
      return `${year}-${month}-${day}`;
    }
  }
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  const year = date.getFullYear();
  const month = padTo2(date.getMonth() + 1);
  const day = padTo2(date.getDate());
  return `${year}-${month}-${day}`;
}

function pickEarliestDate(values = []) {
  if (!Array.isArray(values) || values.length === 0) {
    return '';
  }
  const normalized = values
    .map(value => formatFirebirdDate(value))
    .filter(Boolean)
    .sort();
  return normalized[0] || '';
}

function getRowValue(row, ...keys) {
  if (!row || typeof row !== 'object') {
    return undefined;
  }
  for (const key of keys) {
    if (!key) continue;
    const variants = [key, key.toUpperCase(), key.toLowerCase()];
    for (const variant of variants) {
      if (Object.prototype.hasOwnProperty.call(row, variant)) {
        const value = row[variant];
        if (value !== undefined) {
          return value;
        }
      }
    }
  }
  return undefined;
}

function parseToggle(value, defaultValue) {
  if (value === undefined || value === null) {
    return defaultValue;
  }
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['false', '0', 'no', 'off'].includes(normalized)) {
      return false;
    }
    if (['true', '1', 'si', 'sí', 'on'].includes(normalized)) {
      return true;
    }
  }
  return Boolean(value);
}

function normalizeNumber(value) {
  const number = Number(value ?? 0);
  return Number.isFinite(number) ? number : 0;
}

function roundTo(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round(normalizeNumber(value) * factor) / factor;
}

function mergeCustomization(baseCustomization = {}, requestCustomization = {}) {
  const baseCsv = baseCustomization.csv || {};
  const merged = {
    includeSummary: baseCustomization.includeSummary !== false,
    includeDetail: baseCustomization.includeDetail !== false,
    includeCharts: baseCustomization.includeCharts !== false,
    includeMovements: baseCustomization.includeMovements !== false,
    includeObservations: baseCustomization.includeObservations !== false,
    includeUniverse: baseCustomization.includeUniverse !== false,
    csv: {
      includePoResumen: baseCsv.includePoResumen !== false,
      includeRemisiones: baseCsv.includeRemisiones !== false,
      includeFacturas: baseCsv.includeFacturas !== false,
      includeTotales: baseCsv.includeTotales !== false,
      includeUniverseInfo: baseCsv.includeUniverseInfo !== false
    }
  };

  const overrides = requestCustomization && typeof requestCustomization === 'object'
    ? requestCustomization
    : {};
  ['includeSummary', 'includeDetail', 'includeCharts', 'includeMovements', 'includeObservations', 'includeUniverse'].forEach(key => {
    if (Object.prototype.hasOwnProperty.call(overrides, key)) {
      merged[key] = parseToggle(overrides[key], merged[key]);
    }
  });

  const csvOverrides = overrides.csv && typeof overrides.csv === 'object' ? overrides.csv : {};
  ['includePoResumen', 'includeRemisiones', 'includeFacturas', 'includeTotales', 'includeUniverseInfo'].forEach(key => {
    if (Object.prototype.hasOwnProperty.call(csvOverrides, key)) {
      merged.csv[key] = parseToggle(csvOverrides[key], merged.csv[key]);
    }
  });

  return merged;
}

function isAdminRequest(req) {
  const headerValue = req.headers?.[ADMIN_HEADER];
  return typeof headerValue === 'string' && headerValue.toLowerCase() === 'true';
}

app.use(express.json());
app.use(express.static(path.join(runtimeBaseDir, 'renderer')));

const baseDir = getEnvVar('BASE_DIR', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');
const baseDirFb = getEnvVar('BASE_DIR_FB', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');
function resolveSqlitePath() {
  return getEnvVar('SQLITE_DB', path.join(runtimeBaseDir, 'PERFILES.DB'));
}

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
    FACTR: `FACTR${num}`,
    FACTF: `FACTF${num}`
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
  const normalizedType = typeof type === 'string' ? type.trim().toLowerCase() : 'info';
  const finalType = normalizedType === 'warning' ? 'alerta' : normalizedType;
  return { message, type: finalType };
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

const DATE_REGEX = /^\d{4}-\d{2}-\d{2}$/u;

function normalizeUniverseFilter(filter = {}) {
  const rawMode = (filter.mode || '').toString().toLowerCase();
  const parseDate = value => {
    if (typeof value !== 'string') return null;
    const trimmed = value.trim();
    return DATE_REGEX.test(trimmed) ? trimmed : null;
  };

  let mode;
  if (['range', 'rango', 'intervalo'].includes(rawMode)) {
    mode = 'range';
  } else if (['single', 'unit', 'unitario', 'día', 'dia', 'unico', 'único'].includes(rawMode)) {
    mode = 'single';
  } else {
    mode = 'global';
  }

  if (mode === 'range') {
    const start = parseDate(filter.startDate ?? filter.start ?? filter.from);
    const end = parseDate(filter.endDate ?? filter.end ?? filter.to);
    if (!start || !end) {
      throw new Error('Proporciona fechas de inicio y fin válidas (AAAA-MM-DD) para el filtro por rango.');
    }
    if (start > end) {
      throw new Error('La fecha inicial no puede ser posterior a la fecha final.');
    }
    return {
      mode,
      startDate: start,
      endDate: end,
      label: `Del ${start} al ${end}`,
      shortLabel: `${start}_a_${end}`,
      description: 'Periodo específico del universo de POs.',
      title: 'Reporte del universo por rango',
      isUniverse: true
    };
  }

  if (mode === 'single') {
    const date = parseDate(filter.date ?? filter.startDate ?? filter.endDate);
    if (!date) {
      throw new Error('Proporciona una fecha válida (AAAA-MM-DD) para el filtro unitario.');
    }
    return {
      mode,
      startDate: date,
      endDate: date,
      label: `Fecha ${date}`,
      shortLabel: date,
      description: 'Concentrado del universo para un solo día.',
      title: 'Reporte del universo (1 día)',
      isUniverse: true
    };
  }

  return {
    mode: 'global',
    startDate: null,
    endDate: null,
    label: 'Global (todas las fechas)',
    shortLabel: 'global',
    description: 'Incluye todo el historial disponible en la empresa seleccionada.',
    title: 'Reporte del universo global',
    isUniverse: true
  };
}

function buildDateFilterClause(field, filter) {
  if (!filter || filter.mode === 'global') {
    return { clause: '', params: [] };
  }
  if (filter.mode === 'range') {
    return { clause: ` AND ${field} BETWEEN ? AND ?`, params: [filter.startDate, filter.endDate] };
  }
  if (filter.mode === 'single') {
    return { clause: ` AND ${field} = ?`, params: [filter.startDate] };
  }
  return { clause: '', params: [] };
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

function normalizePoSelection(poTargets = [], fallbackIds = []) {
  const candidates = [];
  if (Array.isArray(poTargets)) {
    poTargets.forEach(target => {
      if (!target) return;
      if (typeof target === 'string') {
        const normalized = normalizePoId(target);
        const baseId = basePoId(normalized);
        if (baseId) {
          candidates.push({ baseId, includeAll: true, ids: null });
        }
        return;
      }
      if (typeof target === 'object') {
        const baseValue = normalizePoId(target.baseId || target.id || target.poId || '');
        const baseId = basePoId(baseValue);
        if (!baseId) {
          return;
        }
        const rawIds = []
          .concat(Array.isArray(target.ids) ? target.ids : [])
          .concat(Array.isArray(target.variants) ? target.variants : [])
          .concat(Array.isArray(target.selectedIds) ? target.selectedIds : [])
          .concat(baseValue ? [baseValue] : []);
        const normalizedIds = uniqueNormalizedValues(rawIds);
        if (normalizedIds.length > 0) {
          candidates.push({ baseId, includeAll: false, ids: normalizedIds });
        } else {
          candidates.push({ baseId, includeAll: true, ids: null });
        }
      }
    });
  }
  if (candidates.length === 0) {
    const bases = uniqueBasePoIds(fallbackIds);
    return bases.map(baseId => ({ baseId, includeAll: true, ids: null }));
  }
  const merged = new Map();
  candidates.forEach(entry => {
    if (!entry.baseId) return;
    if (!merged.has(entry.baseId)) {
      merged.set(entry.baseId, {
        baseId: entry.baseId,
        includeAll: entry.includeAll,
        ids: entry.includeAll ? null : [...(entry.ids || [])]
      });
      return;
    }
    const current = merged.get(entry.baseId);
    if (entry.includeAll) {
      current.includeAll = true;
      current.ids = null;
    } else if (!current.includeAll) {
      current.ids = uniqueNormalizedValues([...(current.ids || []), ...(entry.ids || [])]);
    }
  });
  return Array.from(merged.values());
}

function normalizeDocKey(value) {
  if (typeof value !== 'string') return '';
  return value.trim().toUpperCase();
}

function uniqueNormalizedValues(values = []) {
  const set = new Set();
  values.forEach(value => {
    const trimmed = typeof value === 'string' ? value.trim() : '';
    if (trimmed) {
      set.add(trimmed);
    }
  });
  return Array.from(set);
}

async function fetchRemisiones(db, tableName, { docAntValues = [], docIds = [] } = {}) {
  if (!tableName) return [];
  const alias = 'r';
  const conditions = [];
  const params = [];
  const normalizedDocAnt = uniqueNormalizedValues(docAntValues);
  const normalizedDocIds = uniqueNormalizedValues(docIds);

  if (normalizedDocAnt.length > 0) {
    const placeholders = normalizedDocAnt.map(() => '?').join(',');
    conditions.push(`TRIM(${alias}.DOC_ANT) IN (${placeholders})`);
    params.push(...normalizedDocAnt);
  }

  if (normalizedDocIds.length > 0) {
    const placeholders = normalizedDocIds.map(() => '?').join(',');
    conditions.push(`TRIM(${alias}.CVE_DOC) IN (${placeholders})`);
    params.push(...normalizedDocIds);
  }

  if (conditions.length === 0) {
    return [];
  }

  const whereClause = conditions.map(condition => `(${condition})`).join(' OR ');
    const query = `
      SELECT
        TRIM(${alias}.CVE_DOC) AS id,
        ${alias}.FECHA_DOC AS fecha,
        COALESCE(SUM(${alias}.IMPORTE), 0) AS importe,
        MAX(${alias}.TIP_DOC_SIG) AS tip_doc_sig,
        MAX(${alias}.DOC_SIG) AS doc_sig,
        MAX(${alias}.TIP_DOC_ANT) AS tip_doc_ant,
        MAX(${alias}.DOC_ANT) AS doc_ant,
        MAX(${alias}.STATUS) AS status
      FROM ${tableName} ${alias}
      WHERE ${alias}.STATUS <> 'C' AND (${whereClause})
      GROUP BY ${alias}.CVE_DOC, ${alias}.FECHA_DOC
      ORDER BY TRIM(${alias}.CVE_DOC)
    `;
  return await queryWithTimeout(db, query, params);
}

async function fetchFacturas(
  db,
  tableName,
  { docAntPoValues = [], docAntRemValues = [], docIds = [] } = {}
) {
  if (!tableName) return [];
  const alias = 'f';
  const conditions = [];
  const params = [];
  const normalizedDocAntPo = uniqueNormalizedValues(docAntPoValues);
  const normalizedDocAntRem = uniqueNormalizedValues(docAntRemValues);
  const normalizedDocIds = uniqueNormalizedValues(docIds);

  if (normalizedDocAntPo.length > 0) {
    const placeholders = normalizedDocAntPo.map(() => '?').join(',');
    conditions.push(`(TRIM(${alias}.DOC_ANT) IN (${placeholders}) AND UPPER(TRIM(${alias}.TIP_DOC_ANT)) = 'P')`);
    params.push(...normalizedDocAntPo);
  }

  if (normalizedDocAntRem.length > 0) {
    const placeholders = normalizedDocAntRem.map(() => '?').join(',');
    conditions.push(`(TRIM(${alias}.DOC_ANT) IN (${placeholders}) AND UPPER(TRIM(${alias}.TIP_DOC_ANT)) = 'R')`);
    params.push(...normalizedDocAntRem);
  }

  if (normalizedDocIds.length > 0) {
    const placeholders = normalizedDocIds.map(() => '?').join(',');
    conditions.push(`TRIM(${alias}.CVE_DOC) IN (${placeholders})`);
    params.push(...normalizedDocIds);
  }

  if (conditions.length === 0) {
    return [];
  }

  const whereClause = conditions.map(condition => `(${condition})`).join(' OR ');
  const query = `
    SELECT
      TRIM(${alias}.CVE_DOC) AS id,
      ${alias}.FECHA_DOC AS fecha,
      COALESCE(SUM(${alias}.IMPORTE), 0) AS importe,
      MAX(${alias}.TIP_DOC_SIG) AS tip_doc_sig,
      MAX(${alias}.DOC_SIG) AS doc_sig,
      MAX(${alias}.TIP_DOC_ANT) AS tip_doc_ant,
      MAX(${alias}.DOC_ANT) AS doc_ant,
      MAX(${alias}.STATUS) AS status
    FROM ${tableName} ${alias}
    WHERE ${alias}.STATUS <> 'C' AND (${whereClause})
    GROUP BY ${alias}.CVE_DOC, ${alias}.FECHA_DOC
    ORDER BY TRIM(${alias}.CVE_DOC)
  `;
  return await queryWithTimeout(db, query, params);
}

async function fetchFacturaStatuses(db, tableName, docIds = []) {
  if (!tableName) return new Map();
  const alias = 'fs';
  const normalizedDocIds = uniqueNormalizedValues(docIds);
  if (normalizedDocIds.length === 0) {
    return new Map();
  }
  const placeholders = normalizedDocIds.map(() => '?').join(',');
  const query = `
    SELECT
      TRIM(${alias}.CVE_DOC) AS id,
      MAX(${alias}.STATUS) AS status
    FROM ${tableName} ${alias}
    WHERE TRIM(${alias}.CVE_DOC) IN (${placeholders})
    GROUP BY TRIM(${alias}.CVE_DOC)
  `;
  const rows = await queryWithTimeout(db, query, normalizedDocIds);
  const map = new Map();
  rows.forEach(row => {
    const id = (row.ID || row.id || '').trim();
    if (!id) return;
    const statusRaw = row.STATUS ?? row.status ?? '';
    const status = typeof statusRaw === 'string' ? statusRaw.trim().toUpperCase() : '';
    map.set(normalizeDocKey(id), status);
  });
  return map;
}

async function getPoSummary(empresa, poId, options = {}) {
  const targetPo = (poId || '').trim();
  if (!targetPo) {
    throw new Error('PO inválida');
  }
  const allowedIds = Array.isArray(options.allowedIds) ? uniqueNormalizedValues(options.allowedIds) : [];
  const allowedSet = allowedIds.length > 0 ? new Set(allowedIds) : null;
  return await withFirebirdConnection(empresa, async (db, tables) => {
    const baseId = basePoId(targetPo);
    const extensionPrefix = baseId ? `${baseId}-` : null;
    const manualIds = allowedSet
      ? Array.from(
          new Set(
            allowedIds
              .map(id => normalizePoId(id))
              .filter(id =>
                id &&
                id !== targetPo &&
                id !== baseId &&
                !(extensionPrefix && id.startsWith(extensionPrefix))
              )
          )
        )
      : [];

    const whereParts = ['TRIM(f.CVE_DOC) = ?'];
    const params = [targetPo];

    if (baseId) {
      whereParts.push('TRIM(f.CVE_DOC) LIKE ?');
      params.push(`${baseId}-%`);
    }

    if (manualIds.length > 0) {
      const placeholders = manualIds.map(() => '?').join(',');
      whereParts.push(`TRIM(f.CVE_DOC) IN (${placeholders})`);
      params.push(...manualIds);
    }

    const poRows = await queryWithTimeout(
      db,
      `SELECT
         TRIM(f.CVE_DOC) AS id,
         f.FECHA_DOC AS fecha,
         COALESCE(f.IMPORTE, 0) AS importe,
         COALESCE(f.CAN_TOT, 0) AS subtotal,
         TRIM(f.TIP_DOC_SIG) AS tip_doc_sig,
         TRIM(f.DOC_SIG) AS doc_sig
       FROM ${tables.FACTP} f
       WHERE f.STATUS <> 'C' AND (${whereParts.join(' OR ')})
       ORDER BY TRIM(f.CVE_DOC)`,
      params
    );
    let targetRows = allowedSet
      ? poRows.filter(row => allowedSet.has(normalizePoId(row.ID || row.id || '')))
      : poRows;
    if (targetRows.length === 0) {
      throw new Error(`No se encontró información para el PO ${targetPo}`);
    }
    if (targetRows.length === 0) {
      throw new Error(`No se encontró información para el PO ${targetPo}`);
    }
    const remDocIds = [];
    const remDocKeys = new Set();
    const factDocIdsFromPo = [];
    const factDocIdKeysFromPo = new Set();
    const factDocAntFromPo = new Set();
    const poDocKeys = new Set();
    targetRows.forEach(row => {
      const id = (row.ID || row.id || '').trim();
      if (id) {
        factDocAntFromPo.add(id);
        poDocKeys.add(normalizeDocKey(id));
      }
      const tipDocRaw = row.TIP_DOC_SIG ?? row.tip_doc_sig;
      const docSigRaw = row.DOC_SIG ?? row.doc_sig;
      const tipDoc = typeof tipDocRaw === 'string' ? tipDocRaw.trim().toUpperCase() : '';
      const docSig = typeof docSigRaw === 'string' ? docSigRaw.trim() : '';
      if (!docSig) {
        return;
      }
      if (tipDoc === 'R') {
        remDocIds.push(docSig);
        remDocKeys.add(normalizeDocKey(docSig));
      } else if (tipDoc === 'F') {
        factDocIdsFromPo.push(docSig);
        factDocIdKeysFromPo.add(normalizeDocKey(docSig));
      }
    });

    const remRows = await fetchRemisiones(db, tables.FACTR, {
      docAntValues: Array.from(factDocAntFromPo),
      docIds: remDocIds
    });
    const filteredRemRows = remRows.filter(row => {
      const remKey = normalizeDocKey(row.ID ?? row.id ?? '');
      const docAntKey = normalizeDocKey(row.DOC_ANT ?? row.doc_ant ?? '');
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      const linkedById = remKey && remDocKeys.has(remKey);
      const linkedByPo = docAntKey && poDocKeys.has(docAntKey) && (tipDocAnt === 'P' || !tipDocAnt);
      return linkedById || linkedByPo;
    });
    const remDataById = new Map();
    const remIdsByPoKey = new Map();
    const factDocIdsFromRem = new Set();
    const factDocAntFromRem = new Set();

    filteredRemRows.forEach(row => {
      const remId = (row.ID || row.id || '').trim();
      if (!remId) return;
      const remKey = normalizeDocKey(remId);
      if (!remKey) return;
      const monto = Number(row.IMPORTE ?? row.importe ?? 0);
      const fecha = row.FECHA ?? row.fecha ?? row.FECHA_DOC ?? null;
      const tipDocSigRaw = row.TIP_DOC_SIG ?? row.tip_doc_sig;
      const tipDocSig = typeof tipDocSigRaw === 'string' ? tipDocSigRaw.trim().toUpperCase() : '';
      const docSigRaw = row.DOC_SIG ?? row.doc_sig ?? '';
      const docSig = typeof docSigRaw === 'string' ? docSigRaw.trim() : '';
      const docAntRaw = row.DOC_ANT ?? row.doc_ant ?? '';
      const docAnt = typeof docAntRaw === 'string' ? docAntRaw.trim() : '';
      const docAntKey = normalizeDocKey(docAnt);
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant;
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      factDocAntFromRem.add(remId);
      if (tipDocSig === 'F' && docSig) {
        factDocIdsFromRem.add(docSig);
      }
      if (docAntKey) {
        if (!remIdsByPoKey.has(docAntKey)) {
          remIdsByPoKey.set(docAntKey, new Set());
        }
        remIdsByPoKey.get(docAntKey).add(remKey);
      }
      remDataById.set(remKey, {
        id: remId,
        fecha,
        monto,
        tipDocSig,
        docSig,
        docAnt,
        docAntKey,
        tipDocAnt
      });
    });

    const factRows = await fetchFacturas(db, tables.FACTF, {
      docAntPoValues: Array.from(factDocAntFromPo),
      docAntRemValues: Array.from(factDocAntFromRem),
      docIds: [...factDocIdsFromPo, ...factDocIdsFromRem]
    });

    const remKeysFromData = new Set(filteredRemRows.map(row => normalizeDocKey(row.ID ?? row.id ?? '')));
    const factDocIdKeysFromRem = new Set(Array.from(factDocIdsFromRem, normalizeDocKey));
    const filteredFactRows = factRows.filter(row => {
      const factKey = normalizeDocKey(row.ID ?? row.id ?? '');
      const docAntKey = normalizeDocKey(row.DOC_ANT ?? row.doc_ant ?? '');
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      const linkedByPo = docAntKey && poDocKeys.has(docAntKey) && (tipDocAnt === 'P' || !tipDocAnt);
      const linkedByRem = docAntKey && remKeysFromData.has(docAntKey) && (tipDocAnt === 'R' || !tipDocAnt);
      const linkedById =
        factKey && (factDocIdKeysFromPo.has(factKey) || factDocIdKeysFromRem.has(factKey));
      return linkedByPo || linkedByRem || linkedById;
    });

    const factIdsByPoKey = new Map();
    const factIdsByRemKey = new Map();
    const factDataById = new Map();
    const factStatusByKey = new Map();

    filteredFactRows.forEach(row => {
      const facId = (row.ID || row.id || '').trim();
      if (!facId) return;
      const facKey = normalizeDocKey(facId);
      if (!facKey) return;
      const monto = Number(row.IMPORTE ?? row.importe ?? 0);
      const fecha = row.FECHA ?? row.fecha ?? row.FECHA_DOC ?? null;
      const docAntRaw = row.DOC_ANT ?? row.doc_ant ?? '';
      const docAnt = typeof docAntRaw === 'string' ? docAntRaw.trim() : '';
      const docAntKey = normalizeDocKey(docAnt);
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      const statusRaw = row.STATUS ?? row.status ?? '';
      const status = typeof statusRaw === 'string' ? statusRaw.trim().toUpperCase() : '';
      if (docAntKey) {
        if (tipDocAnt === 'R') {
          if (!factIdsByRemKey.has(docAntKey)) {
            factIdsByRemKey.set(docAntKey, new Set());
          }
          factIdsByRemKey.get(docAntKey).add(facKey);
        } else if (tipDocAnt === 'P' || !tipDocAnt) {
          if (!factIdsByPoKey.has(docAntKey)) {
            factIdsByPoKey.set(docAntKey, new Set());
          }
          factIdsByPoKey.get(docAntKey).add(facKey);
        }
      }
      const linkedByDocSig = factDocIdKeysFromPo.has(facKey) || factDocIdKeysFromRem.has(facKey);
      const linkedByPo = docAntKey && (tipDocAnt === 'P' || !tipDocAnt) && poDocKeys.has(docAntKey);
      const linkedByRem = docAntKey && tipDocAnt === 'R' && remKeysFromData.has(docAntKey);
      const remSources = tipDocAnt === 'R' && docAntKey ? [docAntKey] : [];
      factDataById.set(facKey, {
        id: facId,
        fecha,
        monto,
        tipDocAnt,
        docAnt,
        docAntKey,
        linkedByDocSig,
        linkedByPo,
        linkedByRem,
        remSources,
        status,
        cancelada: status === 'C'
      });
      if (status) {
        factStatusByKey.set(facKey, status);
      }
    });

    const statusLookup = await fetchFacturaStatuses(
      db,
      tables.FACTF,
      [...factDocIdsFromPo, ...Array.from(factDocIdsFromRem)]
    );
    statusLookup.forEach((status, key) => {
      if (status && !factStatusByKey.has(key)) {
        factStatusByKey.set(key, status);
      }
    });

    const linkingAvailable = true;

    const items = targetRows.map(row => {
      const id = (row.ID || row.id || '').trim();
      const fecha = formatFirebirdDate(row.FECHA || row.fecha);
      const totalOriginal = Number(row.IMPORTE ?? row.importe ?? row.TOTAL ?? row.total ?? 0);
      const subtotal = Number(row.SUBTOTAL ?? row.subtotal ?? row.CAN_TOT ?? row.can_tot ?? 0);
      const tipDocRaw = row.TIP_DOC_SIG ?? row.tip_doc_sig;
      const docSigRaw = row.DOC_SIG ?? row.doc_sig;
      const tipDoc = typeof tipDocRaw === 'string' ? tipDocRaw.trim().toUpperCase() : '';
      const docSig = typeof docSigRaw === 'string' ? docSigRaw.trim() : '';
      const docSigKey = normalizeDocKey(docSig);
      const poKey = normalizeDocKey(id);
      const remisiones = [];
      const facturas = [];
      const remKeysForItem = new Set();
      const factKeysForItem = new Set();

      if (tipDoc === 'R' && docSigKey) {
        remKeysForItem.add(docSigKey);
      }
      if (poKey && remIdsByPoKey.has(poKey)) {
        for (const remKey of remIdsByPoKey.get(poKey)) {
          remKeysForItem.add(remKey);
        }
      }

      const orderedRemKeys = Array.from(remKeysForItem).sort();

      orderedRemKeys.forEach(remKey => {
        const remEntry = remDataById.get(remKey);
        if (remEntry) {
          const monto = Number(remEntry.monto || 0);
          const linkedByDocSig = tipDoc === 'R' && docSigKey === remKey;
          const linkedByDocAnt = remEntry.docAntKey && remEntry.docAntKey === poKey;
          remisiones.push({
            id: remEntry.id,
            fecha: formatFirebirdDate(remEntry.fecha),
            monto,
            porcentaje: totalOriginal > 0 ? (monto / totalOriginal) * 100 : 0,
            vinculado: linkedByDocSig || linkedByDocAnt
          });
          if (remEntry.tipDocSig === 'F') {
            const facKeyFromRem = normalizeDocKey(remEntry.docSig);
            if (facKeyFromRem) {
              factKeysForItem.add(facKeyFromRem);
            }
          }
        } else {
          remisiones.push({
            id: remKey,
            fecha: null,
            monto: 0,
            porcentaje: 0,
            vinculado: false
          });
        }
      });

      if (tipDoc === 'F' && docSigKey) {
        factKeysForItem.add(docSigKey);
      }
      if (poKey && factIdsByPoKey.has(poKey)) {
        for (const factKey of factIdsByPoKey.get(poKey)) {
          factKeysForItem.add(factKey);
        }
      }
      for (const remKey of remKeysForItem) {
        const factsFromRem = factIdsByRemKey.get(remKey);
        if (factsFromRem) {
          for (const factKey of factsFromRem) {
            factKeysForItem.add(factKey);
          }
        }
      }

      const orderedFactKeys = Array.from(factKeysForItem).sort();

      orderedFactKeys.forEach(factKey => {
        const factEntry = factDataById.get(factKey);
        const status = factStatusByKey.get(factKey) || (factEntry && factEntry.status) || '';
        const cancelada = status === 'C';
        if (factEntry) {
          const monto = Number(factEntry.monto || 0);
          const linkedByDocSig = tipDoc === 'F' && docSigKey === factKey;
          const linkedByPo =
            factEntry.docAntKey &&
            factEntry.docAntKey === poKey &&
            (factEntry.tipDocAnt === 'P' || !factEntry.tipDocAnt);
          const remSources = Array.isArray(factEntry.remSources) ? factEntry.remSources : [];
          const linkedByRem = remSources.some(source => remKeysForItem.has(source));
          facturas.push({
            id: factEntry.id,
            fecha: formatFirebirdDate(factEntry.fecha),
            monto,
            porcentaje: totalOriginal > 0 ? (monto / totalOriginal) * 100 : 0,
            vinculado: linkedByDocSig || linkedByPo || linkedByRem,
            status,
            cancelada
          });
        } else {
          facturas.push({
            id: factKey,
            fecha: null,
            monto: 0,
            porcentaje: 0,
            vinculado: false,
            status,
            cancelada
          });
        }
      });

      const notasVenta = [];
      const cotizacionesOrigen = [];
      const cotizacionesPosteriores = [];
      const totalRem = remisiones.reduce((sum, item) => sum + item.monto, 0);
      const totalFac = facturas.reduce((sum, item) => sum + item.monto, 0);
      const totals = calculateTotals(totalOriginal, totalRem, totalFac);
      const alerts = [];
      if (totalOriginal > 0) {
        const ratio = totals.totalConsumo / totalOriginal;
        const fullyConsumed = totals.restante <= 0.01;
        if (fullyConsumed) {
          alerts.push(buildAlert(`El PO ${id} está consumido al 100%`, 'alerta'));
        } else if (ratio >= 0.9) {
          alerts.push(
            buildAlert(`El consumo del PO ${id} ha alcanzado el ${(ratio * 100).toFixed(2)}%`, 'alerta')
          );
        }
      }
      const remSinVinculo = remisiones.filter(rem => !rem.vinculado);
      if (remSinVinculo.length > 0) {
        alerts.push(
          buildAlert(
            `Remisiones sin documento encontrado por DOC_SIG: ${remSinVinculo.map(rem => rem.id).join(', ')}`,
            'warning'
          )
        );
      }
      const facSinVinculo = facturas.filter(fac => !fac.vinculado);
      const facSinVinculoActivas = facSinVinculo.filter(fac => !fac.cancelada);
      if (facSinVinculoActivas.length > 0) {
        alerts.push(
          buildAlert(
            `Facturas sin documento encontrado por DOC_SIG: ${facSinVinculoActivas
              .map(fac => fac.id)
              .join(', ')}`,
            'error'
          )
        );
      }
      const remisionesTexto = remisiones.length
        ? remisiones
            .map(rem => {
              const baseLinea = `${rem.id} • ${formatFirebirdDate(rem.fecha)} • $${rem.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${rem.porcentaje.toFixed(2)}%)`;
              return `${baseLinea} • ${rem.vinculado ? 'DOC_SIG encontrado' : 'Sin coincidencia DOC_SIG'}`;
            })
            .join('\n')
        : 'Sin remisiones registradas';
      const facturasTexto = facturas.length
        ? facturas
            .map(fac => {
              const baseLinea = `${fac.id} • ${formatFirebirdDate(fac.fecha)} • $${fac.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${fac.porcentaje.toFixed(2)}%)`;
              return `${baseLinea} • ${fac.vinculado ? 'DOC_SIG encontrado' : 'Sin coincidencia DOC_SIG'}`;
            })
            .join('\n')
        : 'Sin facturas registradas';
      const notasVentaTexto = 'Seguimiento de notas de venta no disponible.';
      const cotizacionesTexto = 'Seguimiento de cotizaciones no disponible.';
      const alertasTexto = alerts.length
        ? alerts.map(alerta => `[${alerta.type.toUpperCase()}] ${alerta.message}`).join('\n')
        : 'Sin alertas';
      return {
        id,
        baseId: basePoId(id),
        fecha,
        total: totalOriginal,
        subtotal,
        remisiones,
        facturas,
        remisionesTexto,
        facturasTexto,
        notasVenta,
        notasVentaTexto,
        cotizacionesOrigen,
        cotizacionesPosteriores,
        cotizacionesTexto,
        linking: {
          disponible: linkingAvailable,
          remisionesSinVinculo: remisiones.filter(rem => !rem.vinculado).map(rem => rem.id),
          facturasSinVinculo: facturas
            .filter(fac => !fac.vinculado && !fac.cancelada)
            .map(fac => fac.id)
        },
        alertasTexto,
        tipDoc,
        docSig,
        totalOriginal,
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
    if (totals.total > 0) {
      const ratio = totals.totalConsumo / totals.total;
      const fullyConsumed = totals.restante <= 0.01;
      if (fullyConsumed) {
        alerts.push(buildAlert(`Los recursos del grupo ${baseId} están consumidos al 100%`, 'alerta'));
      } else if (ratio >= 0.9) {
        alerts.push(
          buildAlert(
            `El consumo total del grupo ${baseId} supera el 90% (${(ratio * 100).toFixed(2)}%)`,
            'alerta'
          )
        );
      }
    }
    const alertasTexto = alerts.length
      ? alerts.map(alerta => `[${alerta.type.toUpperCase()}] ${alerta.message}`).join('\n')
      : 'Sin alertas generales';
    const empresaLabel = buildEmpresaLabel(empresa);
    const selectedVariants = items.map(item => item.id);
    const selection = {
      baseId,
      variants: selectedVariants,
      includeAll: !allowedSet,
      count: selectedVariants.length
    };
    return {
      empresa,
      empresaNumero: extractEmpresaNumero(empresa),
      empresaLabel,
      companyName: empresaLabel,
      baseId,
      selectedId: targetPo,
      selectedIds: [baseId],
      selectedTargets: selectedVariants,
      totals,
      totalsTexto,
      items,
      alerts,
      alertasTexto,
      selection,
      selectionDetails: [selection]
    };
  });
}

async function getPoSummaryGroup(empresa, selectionEntries) {
  const entries = Array.isArray(selectionEntries) ? selectionEntries.filter(entry => entry && entry.baseId) : [];
  if (entries.length === 0) {
    throw new Error('Selecciona al menos una PO válida.');
  }
  const summaries = [];
  for (const entry of entries) {
    const summary = await getPoSummary(empresa, entry.baseId, {
      allowedIds: entry.includeAll ? null : entry.ids
    });
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
  if (totals.total > 0) {
    const ratio = totals.totalConsumo / totals.total;
    const fullyConsumed = totals.restante <= 0.01;
    if (fullyConsumed) {
      alerts.push(buildAlert('Los recursos combinados están consumidos al 100%', 'alerta'));
    } else if (ratio >= 0.9) {
      alerts.push(
        buildAlert(
          `El consumo total combinado supera el 90% (${(ratio * 100).toFixed(2)}%)`,
          'alerta'
        )
      );
    }
  }
  const alertasTexto = alerts.length
    ? alerts.map(alerta => `[${(alerta.type || 'info').toUpperCase()}] ${alerta.message}`).join('\n')
    : 'Sin alertas generales';

  const selectedIds = Array.from(new Set(entries.map(entry => entry.baseId))).filter(Boolean);
  const selectionDetails = summaries.flatMap(summary => summary.selectionDetails || []);
  const selectedTargets = Array.from(
    new Set(selectionDetails.flatMap(detail => detail.variants || []))
  );

  const selection = {
    baseId: null,
    variants: selectedTargets,
    includeAll: entries.some(entry => entry.includeAll),
    count: selectedTargets.length
  };

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
    alertasTexto,
    selection,
    selectionDetails
  };
}

async function getUniverseSummary(empresa, rawFilter) {
  const filter = normalizeUniverseFilter(rawFilter);
  return await withFirebirdConnection(empresa, async (db, tables) => {
    const isDateWithinFilter = value => {
      if (!filter || filter.mode === 'global') {
        return true;
      }
      const formatted = formatFirebirdDate(value);
      if (!formatted) {
        return false;
      }
      if (filter.mode === 'range') {
        return formatted >= filter.startDate && formatted <= filter.endDate;
      }
      if (filter.mode === 'single') {
        return formatted === filter.startDate;
      }
      return true;
    };

    const { clause: poClause, params: poParams } = buildDateFilterClause('f.FECHA_DOC', filter);
    const poRows = await queryWithTimeout(
      db,
      `SELECT
         TRIM(f.CVE_DOC) AS id,
         f.FECHA_DOC AS fecha,
         COALESCE(f.IMPORTE, 0) AS importe,
         COALESCE(f.CAN_TOT, 0) AS subtotal,
         TRIM(f.TIP_DOC_SIG) AS tip_doc_sig,
         TRIM(f.DOC_SIG) AS doc_sig
       FROM ${tables.FACTP} f
       WHERE f.STATUS <> 'C'${poClause}`,
      poParams
    );

    const poIds = new Set();
    const poDocKeys = new Set();
    const remDocIds = new Set();
    const remDocKeys = new Set();
    const factDocIdsFromPo = new Set();
    const factDocIdKeysFromPo = new Set();
    const poMap = new Map();
    const ensureUniversePo = poId => {
      const normalized = normalizePoId(poId);
      if (!normalized) return null;
      if (!poMap.has(normalized)) {
        const baseId = basePoId(normalized);
        poMap.set(normalized, {
          id: normalized,
          baseId: baseId || normalized,
          fecha: '',
          total: 0,
          subtotal: 0,
          docSig: '',
          tipDoc: '',
          remisiones: [],
          facturas: [],
          totals: { total: 0, totalRem: 0, totalFac: 0, totalConsumo: 0, restante: 0 },
          alerts: []
        });
      }
      return poMap.get(normalized);
    };

    poRows.forEach(row => {
      const id = normalizePoId(row.ID ?? row.id ?? '');
      if (id) {
        poIds.add(id);
        const poKey = normalizeDocKey(id);
        if (poKey) {
          poDocKeys.add(poKey);
        }
        const entry = ensureUniversePo(id);
        const fecha = formatFirebirdDate(getRowValue(row, 'FECHA_DOC', 'FECHA', 'fecha'));
        const totalImporte = Number(row.IMPORTE ?? row.importe ?? 0);
        const subtotalImporte = Number(row.SUBTOTAL ?? row.subtotal ?? row.CAN_TOT ?? row.can_tot ?? 0);
        if (fecha) {
          entry.fecha = fecha;
        }
        entry.total = roundTo(totalImporte);
        entry.subtotal = roundTo(subtotalImporte);
        entry.totals.total = roundTo(totalImporte);
        entry.docSig = typeof row.DOC_SIG === 'string' ? row.DOC_SIG.trim() : typeof row.doc_sig === 'string' ? row.doc_sig.trim() : '';
        entry.tipDoc = typeof row.TIP_DOC_SIG === 'string'
          ? row.TIP_DOC_SIG.trim().toUpperCase()
          : typeof row.tip_doc_sig === 'string'
            ? row.tip_doc_sig.trim().toUpperCase()
            : '';
      }
      const tipDoc = typeof row.TIP_DOC_SIG === 'string' ? row.TIP_DOC_SIG.trim().toUpperCase() :
        typeof row.tip_doc_sig === 'string' ? row.tip_doc_sig.trim().toUpperCase() : '';
      const docSig = typeof row.DOC_SIG === 'string' ? row.DOC_SIG.trim() :
        typeof row.doc_sig === 'string' ? row.doc_sig.trim() : '';
      if (!docSig) {
        return;
      }
      if (tipDoc === 'R') {
        remDocIds.add(docSig);
        remDocKeys.add(normalizeDocKey(docSig));
      } else if (tipDoc === 'F') {
        factDocIdsFromPo.add(docSig);
        factDocIdKeysFromPo.add(normalizeDocKey(docSig));
      }
    });

    const remRows = await fetchRemisiones(db, tables.FACTR, {
      docAntValues: Array.from(poIds),
      docIds: Array.from(remDocIds)
    });

    const filteredRemRows = remRows.filter(row => {
      const remKey = normalizeDocKey(row.ID ?? row.id ?? '');
      const docAntKey = normalizeDocKey(row.DOC_ANT ?? row.doc_ant ?? '');
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      const linkedById = remKey && remDocKeys.has(remKey);
      const linkedByPo = docAntKey && poDocKeys.has(docAntKey) && (tipDocAnt === 'P' || !tipDocAnt);
      const fechaOk = isDateWithinFilter(getRowValue(row, 'FECHA_DOC', 'FECHA', 'fecha'));
      return fechaOk && (linkedById || linkedByPo);
    });

    const remKeysFromData = new Set(
      filteredRemRows
        .map(row => normalizeDocKey(row.ID ?? row.id ?? ''))
        .filter(Boolean)
    );

    const remToPo = new Map();
    filteredRemRows.forEach(row => {
      const remId = normalizePoId(row.ID ?? row.id ?? '');
      const docAnt = normalizePoId(row.DOC_ANT ?? row.doc_ant ?? '');
      if (!remId) return;
      let targetPo = '';
      if (docAnt) {
        if (poMap.has(docAnt)) {
          targetPo = docAnt;
        } else {
          const baseCandidate = basePoId(docAnt);
          if (baseCandidate && poMap.has(baseCandidate)) {
            targetPo = baseCandidate;
          } else {
            targetPo = docAnt;
          }
        }
      }
      if (!targetPo) return;
      const entry = ensureUniversePo(targetPo);
      if (!entry) return;
      const monto = Number(row.IMPORTE ?? row.importe ?? 0);
      const fecha = formatFirebirdDate(getRowValue(row, 'FECHA_DOC', 'FECHA', 'fecha'));
      entry.remisiones.push({
        id: remId,
        fecha,
        monto,
        porcentaje: entry.total > 0 ? roundTo((monto / entry.total) * 100) : 0,
        vinculado: true
      });
      entry.totals.totalRem += monto;
      remToPo.set(remId, entry.id);
    });

    const factDocIdsFromRem = new Set();
    const factDocIdKeysFromRem = new Set();
    const factDocAntFromRem = new Set();

    filteredRemRows.forEach(row => {
      const remId = (row.ID || row.id || '').trim();
      if (remId) {
        factDocAntFromRem.add(remId);
      }
      const tipDocSigRaw = row.TIP_DOC_SIG ?? row.tip_doc_sig ?? '';
      const tipDocSig = typeof tipDocSigRaw === 'string' ? tipDocSigRaw.trim().toUpperCase() : '';
      const docSigRaw = row.DOC_SIG ?? row.doc_sig ?? '';
      const docSig = typeof docSigRaw === 'string' ? docSigRaw.trim() : '';
      if (tipDocSig === 'F' && docSig) {
        factDocIdsFromRem.add(docSig);
        factDocIdKeysFromRem.add(normalizeDocKey(docSig));
      }
    });

    const factRows = await fetchFacturas(db, tables.FACTF, {
      docAntPoValues: Array.from(poIds),
      docAntRemValues: Array.from(factDocAntFromRem),
      docIds: Array.from(new Set([...factDocIdsFromPo, ...factDocIdsFromRem]))
    });

    const filteredFactRows = factRows.filter(row => {
      const factKey = normalizeDocKey(row.ID ?? row.id ?? '');
      const docAntKey = normalizeDocKey(row.DOC_ANT ?? row.doc_ant ?? '');
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      const linkedByPo = docAntKey && poDocKeys.has(docAntKey) && (tipDocAnt === 'P' || !tipDocAnt);
      const linkedByRem = docAntKey && remKeysFromData.has(docAntKey) && (tipDocAnt === 'R' || !tipDocAnt);
      const linkedById =
        factKey && (factDocIdKeysFromPo.has(factKey) || factDocIdKeysFromRem.has(factKey));
      const fechaOk = isDateWithinFilter(getRowValue(row, 'FECHA_DOC', 'FECHA', 'fecha'));
      return fechaOk && (linkedByPo || linkedByRem || linkedById);
    });

    filteredFactRows.forEach(row => {
      const factId = normalizePoId(row.ID ?? row.id ?? '');
      const docAnt = normalizePoId(row.DOC_ANT ?? row.doc_ant ?? '');
      const tipDocAntRaw = row.TIP_DOC_ANT ?? row.tip_doc_ant ?? '';
      const tipDocAnt = typeof tipDocAntRaw === 'string' ? tipDocAntRaw.trim().toUpperCase() : '';
      let targetPo = '';
      if (tipDocAnt === 'P' && docAnt) {
        if (poMap.has(docAnt)) {
          targetPo = docAnt;
        } else {
          const baseCandidate = basePoId(docAnt);
          if (baseCandidate && poMap.has(baseCandidate)) {
            targetPo = baseCandidate;
          } else {
            targetPo = docAnt;
          }
        }
      } else if (tipDocAnt === 'R' && docAnt) {
        targetPo = remToPo.get(docAnt) || '';
      }
      if (!targetPo && docAnt) {
        const baseCandidate = basePoId(docAnt);
        if (baseCandidate && poMap.has(baseCandidate)) {
          targetPo = baseCandidate;
        }
      }
      if (!targetPo && factId) {
        const baseCandidate = basePoId(factId);
        if (baseCandidate && poMap.has(baseCandidate)) {
          targetPo = baseCandidate;
        }
      }
      if (!targetPo) return;
      const entry = ensureUniversePo(targetPo);
      if (!entry) return;
      const monto = Number(row.IMPORTE ?? row.importe ?? 0);
      const fecha = formatFirebirdDate(getRowValue(row, 'FECHA_DOC', 'FECHA', 'fecha'));
      entry.facturas.push({
        id: factId,
        fecha,
        monto,
        porcentaje: entry.total > 0 ? roundTo((monto / entry.total) * 100) : 0,
        vinculado: true
      });
      entry.totals.totalFac += monto;
    });

    const items = Array.from(poMap.values())
      .map(entry => {
        const totals = calculateTotals(entry.total, entry.totals.totalRem, entry.totals.totalFac);
        const alerts = [];
        if (totals.total > 0) {
          const ratio = totals.totalConsumo / totals.total;
          const fullyConsumed = totals.restante <= 0.01;
          if (fullyConsumed) {
            alerts.push(buildAlert(`El PO ${entry.id} está consumido al 100%`, 'alerta'));
          } else if (ratio >= 0.9) {
            alerts.push(
              buildAlert(
                `El consumo del PO ${entry.id} ha alcanzado el ${(ratio * 100).toFixed(2)}%`,
                'alerta'
              )
            );
          }
        }
        const alertasTexto = alerts.length
          ? alerts.map(alerta => `[${(alerta.type || 'info').toUpperCase()}] ${alerta.message}`).join('\n')
          : 'Sin alertas';
        const remisionesTexto = entry.remisiones.length
          ? entry.remisiones
              .map(rem => {
                const fechaLabel = rem.fecha || '-';
                return `${rem.id} • ${fechaLabel} • $${rem.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${rem.porcentaje.toFixed(2)}%)`;
              })
              .join('\n')
          : 'Sin remisiones registradas';
        const facturasTexto = entry.facturas.length
          ? entry.facturas
              .map(fac => {
                const fechaLabel = fac.fecha || '-';
                return `${fac.id} • ${fechaLabel} • $${fac.monto.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${fac.porcentaje.toFixed(2)}%)`;
              })
              .join('\n')
          : 'Sin facturas registradas';
        const resolvedDate = pickEarliestDate([
          entry.fecha,
          ...entry.remisiones.map(rem => rem.fecha),
          ...entry.facturas.map(fac => fac.fecha)
        ]);
        const finalDate = resolvedDate || formatFirebirdDate(entry.fecha);
        entry.fecha = finalDate;
        return {
          id: entry.id,
          baseId: entry.baseId,
          fecha: finalDate,
          total: entry.total,
          subtotal: entry.subtotal,
          remisiones: entry.remisiones,
          facturas: entry.facturas,
          remisionesTexto,
          facturasTexto,
          totals,
          alerts,
          alertasTexto,
          docSig: entry.docSig,
          tipDoc: entry.tipDoc
        };
      })
      .sort((a, b) => a.id.localeCompare(b.id));

    const aggregated = items.reduce(
      (acc, item) => {
        acc.total += Number(item.total || 0);
        acc.totalRem += Number(item.totals?.totalRem || 0);
        acc.totalFac += Number(item.totals?.totalFac || 0);
        return acc;
      },
      { total: 0, totalRem: 0, totalFac: 0 }
    );

    const totals = calculateTotals(aggregated.total, aggregated.totalRem, aggregated.totalFac);
    const totalsTexto =
      `Total universo: $${totals.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}\n` +
      `Remisiones: $${totals.totalRem.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRem.toFixed(2)}%)\n` +
      `Facturas: $${totals.totalFac.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcFac.toFixed(2)}%)\n` +
      `Restante: $${totals.restante.toLocaleString('es-MX', { minimumFractionDigits: 2 })} (${totals.porcRest.toFixed(2)}%)`;

    const alerts = [];
    if (totals.total === 0) {
      alerts.push(buildAlert('No se encontraron POs activas con el filtro seleccionado.', 'info'));
    } else {
      const ratio = totals.totalConsumo / totals.total;
      const fullyConsumed = totals.restante <= 0.01;
      if (fullyConsumed) {
        alerts.push(buildAlert('El universo está consumido al 100%', 'alerta'));
      } else if (ratio >= 0.9) {
        alerts.push(
          buildAlert(
            `El consumo del universo supera el 90% (${(ratio * 100).toFixed(2)}%)`,
            'alerta'
          )
        );
      }
    }

    const alertasTexto = alerts.length
      ? alerts.map(alerta => `[${(alerta.type || 'info').toUpperCase()}] ${alerta.message}`).join('\n')
      : 'Sin alertas generales';

    const empresaLabel = buildEmpresaLabel(empresa);

    return {
      empresa,
      empresaNumero: extractEmpresaNumero(empresa),
      empresaLabel,
      companyName: empresaLabel,
      baseId: null,
      selectedId: `Universo - ${filter.label}`,
      selectedIds: ['UNIVERSO'],
      selectedTargets: [],
      totals,
      totalsTexto,
      items,
      alerts,
      alertasTexto,
      universe: filter
    };
  });
}

app.get('/po-overview/:empresa', async (req, res) => {
  const empresa = req.params.empresa;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  const { mode, startDate, endDate, date } = req.query;
  const filter = { mode, startDate, endDate, date };
  try {
    const summary = await getUniverseSummary(empresa, filter);
    res.json({ success: true, summary });
  } catch (err) {
    const status = err.message && err.message.includes('Proporciona') ? 400 : 500;
    res.status(status).json({ success: false, message: 'Error consultando el resumen de POs: ' + err.message });
  }
});

app.get('/pos/:empresa', async (req, res) => {
  const empresa = req.params.empresa;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const rows = await withFirebirdConnection(empresa, async (db, tables) => {
      const result = await queryWithTimeout(db, `
        SELECT
          TRIM(f.CVE_DOC) AS id,
          f.FECHA_DOC AS fecha,
          COALESCE(f.IMPORTE, 0) AS total,
          COALESCE(f.CAN_TOT, 0) AS subtotal
        FROM ${tables.FACTP} f
        WHERE f.STATUS <> 'C'
        ORDER BY f.FECHA_DOC DESC
      `);
      return result.map(row => ({
        id: (row.ID || row.id || '').trim(),
        fecha: formatFirebirdDate(row.FECHA || row.fecha),
        total: Number(row.TOTAL ?? row.total ?? row.IMPORTE ?? row.importe ?? 0),
        subtotal: Number(row.SUBTOTAL ?? row.subtotal ?? row.CAN_TOT ?? row.can_tot ?? 0),
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
  const { empresa, poIds, poTargets } = req.body || {};
  if (!empresa || !empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const selection = normalizePoSelection(Array.isArray(poTargets) ? poTargets : [], Array.isArray(poIds) ? poIds : []);
    const summary = await getPoSummaryGroup(empresa, selection);
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
    const selection = normalizePoSelection([], [poId]);
    const summary = await getPoSummaryGroup(empresa, selection);
    res.json({ success: true, summary });
  } catch (err) {
    const status = err.message && err.message.includes('No se encontró') ? 404 : 500;
    res.status(status).json({ success: false, message: err.message });
  }
});

app.get('/report-settings', async (req, res) => {
  try {
    const settings = reportSettingsStore.getSettings();
    const engines = [
      {
        id: 'simple-pdf',
        label: 'PDF directo',
        description: 'Genera un PDF resumido usando pdfkit.',
        available: simplePdf.isAvailable()
      }
    ];
    res.json({
      success: true,
      settings,
      engines,
      formats: ALLOWED_FORMATS
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

    if (payload.export) {
      const exportCfg = {};
      if (payload.export.defaultFormat && ALLOWED_FORMATS.includes(payload.export.defaultFormat)) {
        exportCfg.defaultFormat = payload.export.defaultFormat;
      }
      if (Array.isArray(payload.export.availableFormats)) {
        const filtered = payload.export.availableFormats.filter(format => ALLOWED_FORMATS.includes(format));
        if (filtered.length) {
          exportCfg.availableFormats = Array.from(new Set(['pdf', ...filtered]));
        }
      }
      if (Object.keys(exportCfg).length) {
        updates.export = exportCfg;
      }
    }

    if (payload.customization) {
      const customization = {};
      ['includeCharts', 'includeMovements', 'includeObservations', 'includeUniverse'].forEach(key => {
        if (payload.customization[key] !== undefined) {
          customization[key] = !!payload.customization[key];
        }
      });
      if (Object.keys(customization).length) {
        updates.customization = customization;
      }
    }

    if (payload.branding) {
      const branding = {};
      const stringKeys = ['headerTitle', 'headerSubtitle', 'companyName', 'footerText', 'letterheadTop', 'letterheadBottom', 'remColor', 'facColor', 'restanteColor', 'accentColor'];
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
    const engines = [
      {
        id: 'simple-pdf',
        label: 'PDF directo',
        description: 'Genera un PDF resumido usando pdfkit.',
        available: simplePdf.isAvailable()
      }
    ];

    res.json({
      success: true,
      settings: updatedSettings,
      engines,
      formats: ALLOWED_FORMATS
    });
  } catch (err) {
    console.error('Error actualizando configuración de reportes:', err);
    res.status(500).json({ success: false, message: 'No se pudo actualizar la configuración de reportes: ' + err.message });
  }
});

app.post('/report-universe', async (req, res) => {
  const { empresa, filter, format: requestedFormat, customization: requestedCustomization } = req.body || {};
  if (!empresa || !empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  try {
    const summary = await getUniverseSummary(empresa, filter || {});
    const settings = reportSettingsStore.getSettings();
    const customization = mergeCustomization(settings.customization || {}, requestedCustomization || {});
    summary.customization = customization;
    if (settings.branding?.companyName) {
      summary.companyName = settings.branding.companyName;
    }
    const normalizedRequestFormat = typeof requestedFormat === 'string' ? requestedFormat.toLowerCase() : '';
    const allowedFormats = ['pdf', 'csv', 'json'];
    const format = allowedFormats.includes(normalizedRequestFormat) ? normalizedRequestFormat : 'pdf';
    const branding = settings.branding || {};
    const rawLabel = summary.universe?.shortLabel || 'global';
    const sanitized = rawLabel.replace(/[^0-9a-zA-Z_-]+/gu, '-');
    const filenameBase = `universo_${sanitized}`;

    if (format === 'csv') {
      const buffer = exporters.createCsv(summary, { customization });
      res.set('Content-Type', 'text/csv; charset=utf-8');
      res.set('Content-Disposition', `attachment; filename=${filenameBase}.csv`);
      return res.send(buffer);
    }

    if (format === 'json') {
      const buffer = exporters.createJson(summary, {
        engine: 'simple-pdf',
        format,
        empresa,
        selectedIds: summary.selectedIds,
        customization
      });
      res.set('Content-Type', 'application/json');
      res.set('Content-Disposition', `attachment; filename=${filenameBase}.json`);
      return res.send(buffer);
    }

    if (!simplePdf.isAvailable()) {
      return res.status(503).json({ success: false, message: simplePdf.getUnavailableMessage() });
    }
    const buffer = await simplePdf.generate(summary, branding, customization);
    res.set('Content-Type', 'application/pdf');
    res.set('Content-Disposition', `attachment; filename=${filenameBase}.pdf`);
    res.send(buffer);
  } catch (err) {
    const status = err.message && err.message.includes('Proporciona') ? 400 : 500;
    res.status(status).json({ success: false, message: 'Error generando reporte universo: ' + err.message });
  }
});

app.post('/report', async (req, res) => {
  const {
    empresa,
    poId,
    poIds,
    poTargets,
    engine: requestedEngine,
    format: requestedFormat,
    customization: requestedCustomization
  } = req.body || {};
  const ids = Array.isArray(poIds) && poIds.length ? poIds : poId ? [poId] : [];
  const selection = normalizePoSelection(Array.isArray(poTargets) ? poTargets : [], ids);
  if (!empresa || selection.length === 0) {
    return res.status(400).json({ success: false, message: 'Empresa y selección de PO son obligatorias para el reporte' });
  }
  try {
    const summary = await getPoSummaryGroup(empresa, selection);
    const settings = reportSettingsStore.getSettings();
    const customization = mergeCustomization(settings.customization || {}, requestedCustomization || {});
    summary.customization = customization;
    if (settings.branding?.companyName) {
      summary.companyName = settings.branding.companyName;
    }
    const engine = ALLOWED_ENGINES.includes(requestedEngine)
      ? requestedEngine
      : (ALLOWED_ENGINES.includes(settings.defaultEngine) ? settings.defaultEngine : 'simple-pdf');
    const availableFormats = Array.isArray(settings.export?.availableFormats)
      ? settings.export.availableFormats.filter(item => ALLOWED_FORMATS.includes(item))
      : ['pdf'];
    if (!availableFormats.includes('pdf')) {
      availableFormats.unshift('pdf');
    }
    const normalizedRequestFormat = typeof requestedFormat === 'string' ? requestedFormat.toLowerCase() : '';
    const defaultFormat = availableFormats.includes(settings.export?.defaultFormat)
      ? settings.export.defaultFormat
      : 'pdf';
    const format = availableFormats.includes(normalizedRequestFormat) ? normalizedRequestFormat : defaultFormat;
    const filenameBase = summary.selectedIds && summary.selectedIds.length
      ? summary.selectedIds.join('_')
      : summary.baseId || 'reporte';
    if (format === 'csv') {
      const buffer = exporters.createCsv(summary, { customization });
      res.set('Content-Type', 'text/csv; charset=utf-8');
      res.set('Content-Disposition', `attachment; filename=${filenameBase}.csv`);
      return res.send(buffer);
    }

    if (format === 'json') {
      const buffer = exporters.createJson(summary, {
        engine,
        format,
        empresa,
        selectedIds: summary.selectedIds,
        customization
      });
      res.set('Content-Type', 'application/json');
      res.set('Content-Disposition', `attachment; filename=${filenameBase}.json`);
      return res.send(buffer);
    }

    if (!simplePdf.isAvailable()) {
      return res.status(503).json({
        success: false,
        message: simplePdf.getUnavailableMessage()
      });
    }

    const buffer = await simplePdf.generate(summary, settings.branding || {}, customization);
    res.set('Content-Type', 'application/pdf');
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
    const sqlitePath = resolveSqlitePath();
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