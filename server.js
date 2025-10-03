require('dotenv').config();
const express = require('express');
const Firebird = require('node-firebird');
const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const bcrypt = require('bcryptjs');
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

// Middleware
app.use(express.json());
app.use(express.static(path.join(runtimeBaseDir, 'renderer')));  // Sirve HTML/CSS/JS

// Directorio base
const baseDir = getEnvVar('BASE_DIR', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');
const baseDirFb = getEnvVar('BASE_DIR_FB', 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\');

// SQLite
const sqlitePath = getEnvVar('SQLITE_DB', path.join(runtimeBaseDir, 'PERFILES.DB'));

// Firebird base
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

// Helpers (queryWithTimeout, runAsync, getAsync, allAsync) - copia de tu código

// Ruta principal (redirige a login)
app.get('/', (req, res) => res.sendFile(path.join(runtimeBaseDir, 'renderer', 'index.html')));

// Login (adaptado: admin "admin" / "569OpEGvwh'8")
app.post('/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ success: false, message: 'Usuario y contraseña obligatorios' });
  try {
    const user = await getAsync(sqliteDb, 'SELECT * FROM usuarios WHERE usuario = ?', [username]);
    if (!user) return res.status(401).json({ success: false, message: 'Credenciales inválidas' });
    const match = await bcrypt.compare(password, user.password);
    if (!match) return res.status(401).json({ success: false, message: 'Credenciales inválidas' });
    let empresas = [];
    const isAdmin = user.usuario === 'admin' || user.empresas === '*';
    if (isAdmin) {
      const dirs = await fs.promises.readdir(baseDir);
      empresas = dirs.filter(dir => dir.match(/^Empresa\d+$/)).sort((a, b) => parseInt(a.replace('Empresa', '')) - parseInt(b.replace('Empresa', '')));
    } else {
      empresas = user.empresas ? JSON.parse(user.empresas) : [];
    }
    res.json({ success: true, message: 'Login exitoso', empresas, isAdmin });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error en login' });
  }
});

// CRUD users (copia de tu código, adaptado para "admin")

// /empresas (copia de tu código)

// Nuevas rutas para POrpt
function getDatabasePaths(empresa) {  // Copia de tu código
  const num = empresa.replace('Empresa', '').padStart(2, '0');
  const fsPath = path.join(baseDir, empresa, 'Datos', `SAE90EMPRE${num}.FDB`);
  const fbPath = path.win32.join(baseDirFb, empresa, 'Datos', `SAE90EMPRE${num}.FDB`);
  return { fsPath, fbPath };
}

function getCompanyTables(empresa) {  // Adaptado para FACT*
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

async function tableExists(db, tableName) {  // Copia

}

// /pos/:empresa (POs)
app.get('/pos/:empresa', async (req, res) => {
  const empresa = req.params.empresa;
  if (!empresa.match(/^Empresa\d+$/)) return res.status(400).json({ success: false, message: 'Empresa inválida' });
  const tables = getCompanyTables(empresa);
  const { fsPath, fbPath } = getDatabasePaths(empresa);
  if (!fs.existsSync(fsPath)) return res.status(404).json({ success: false, message: `DB no encontrada` });
  const options = { ...baseOptions, database: fbPath };
  Firebird.attach(options, async (err, db) => {
    if (err) return res.status(500).json({ success: false, message: err.message });
    try {
      const rows = await queryWithTimeout(db, `
        SELECT f.CVE_DOC as id, f.FECHA_DOC as fecha, COALESCE(SUM(p.TOT_PARTIDA), 0) as total
        FROM ${tables.FACTP} f LEFT JOIN ${tables.PAR_FACTP} p ON f.CVE_DOC = p.CVE_DOC
        WHERE f.STATUS <> 'C' GROUP BY f.CVE_DOC, f.FECHA_DOC ORDER BY f.FECHA_DOC DESC
      `);
      res.json({ success: true, pos: rows });
    } catch (e) {
      res.status(500).json({ success: false, message: e.message });
    } finally {
      db.detach();
    }
  });
});

// /rems/:empresa/:poId (remisiones enlazadas)
app.get('/rems/:empresa/:poId', async (req, res) => {
  // Similar: Query FACTR + PAR_FACTR, WHERE CVE_PEDI LIKE poId + '%'
  // Retorna [{id, monto, porcentaje:0}]
});

// /facts/:empresa/:poId (similar para facturas)

// /report (POST data: {ssitel, po, rems, facts} -> retorna buffer PDF)
app.post('/report', (req, res) => {
  const pdfmake = require('pdfmake/build/pdfmake');
  const pdfFonts = require('pdfmake/build/vfs_fonts');
  pdfmake.vfs = pdfFonts.pdfMake.vfs;
  // Usa generatePOReport de paso anterior, pero recibe data en body
  // Calcula % , genera docDefinition, pdfMake.createPdf(doc).getBuffer((buf) => res.send(buf))
});

// Diagnóstico /diagnostico/pos/:empresa (similar a tu /diagnostico/recepciones)

// Middleware errores (copia)

// init/shutdown (adaptado: Crea admin "admin" / "569OpEGvwh'8" hashed)
async function init(options = {}) {
  // ... copia, pero en CREATE TABLE y admin insert: usuario='admin', password=bcrypt.hashSync('569OpEGvwh\'8', 10), empresas='*'
}

module.exports = { init, shutdown, app };  // Exporta app para main.js