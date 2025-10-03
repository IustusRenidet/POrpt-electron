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
  const { usuario, password, nombre, empresas } = req.body;
  if (!usuario || !password || !Array.isArray(empresas)) {
    return res.status(400).json({ success: false, message: 'Datos inválidos' });
  }
  try {
    const hash = await bcrypt.hash(password, 10);
    const result = await runAsync(sqliteDb, 'INSERT INTO usuarios (usuario, password, nombre, empresas) VALUES (?,?,?,?)', [
      usuario,
      hash,
      nombre || null,
      JSON.stringify(empresas)
    ]);
    res.json({ success: true, id: result.lastID });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Error creando usuario: ' + err.message });
  }
});

app.put('/users/:id', async (req, res) => {
  const { id } = req.params;
  const { usuario, password, nombre, empresas } = req.body;
  try {
    const existing = await getAsync(sqliteDb, 'SELECT * FROM usuarios WHERE id = ?', [id]);
    if (!existing) return res.status(404).json({ success: false, message: 'Usuario no encontrado' });
    const newUsuario = usuario || existing.usuario;
    const newNombre = nombre !== undefined ? nombre : existing.nombre;
    const newEmpresas = Array.isArray(empresas) ? JSON.stringify(empresas) : existing.empresas;
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
  const tables = getCompanyTables(empresa);
  const { fsPath, fbPath } = getDatabasePaths(empresa);
  if (!fs.existsSync(fsPath)) {
    return res.status(404).json({ success: false, message: `Base de datos no encontrada para ${empresa}` });
  }
  const options = { ...baseOptions, database: fbPath };
  Firebird.attach(options, async (err, db) => {
    if (err) {
      console.error('Error conectando a la BD:', err);
      return res.status(500).json({ success: false, message: 'Error conectando a la BD: ' + err.message });
    }
    try {
      const rows = await queryWithTimeout(db, `
        SELECT TRIM(f.CVE_DOC) AS id, f.FECHA_DOC as fecha, COALESCE(SUM(p.TOT_PARTIDA), 0) as total
        FROM ${tables.FACTP} f
        LEFT JOIN ${tables.PAR_FACTP} p ON f.CVE_DOC = p.CVE_DOC
        WHERE f.STATUS <> 'C'
        GROUP BY f.CVE_DOC, f.FECHA_DOC
        ORDER BY f.FECHA_DOC DESC
      `);
      res.json({ success: true, pos: rows });
    } catch (e) {
      res.status(500).json({ success: false, message: 'Error consultando POs: ' + e.message });
    } finally {
      db.detach();
    }
  });
});

app.get('/rems/:empresa/:poId', async (req, res) => {
  const { empresa, poId } = req.params;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  const tables = getCompanyTables(empresa);
  const { fsPath, fbPath } = getDatabasePaths(empresa);
  if (!fs.existsSync(fsPath)) {
    return res.status(404).json({ success: false, message: `Base de datos no encontrada para ${empresa}` });
  }
  const options = { ...baseOptions, database: fbPath };
  Firebird.attach(options, async (err, db) => {
    if (err) {
      console.error('Error conectando a la BD:', err);
      return res.status(500).json({ success: false, message: 'Error conectando a la BD: ' + err.message });
    }
    try {
      const rows = await queryWithTimeout(db, `
        SELECT TRIM(r.CVE_DOC) AS id, COALESCE(SUM(pr.TOT_PARTIDA), 0) as monto
        FROM ${tables.FACTR} r
        LEFT JOIN ${tables.PAR_FACTR} pr ON r.CVE_DOC = pr.CVE_DOC
        WHERE r.CVE_PEDI LIKE ?
        AND r.STATUS <> 'C'
        GROUP BY r.CVE_DOC
      `, [poId + '%']);
      res.json({ success: true, rems: rows });
    } catch (e) {
      res.status(500).json({ success: false, message: 'Error consultando remisiones: ' + e.message });
    } finally {
      db.detach();
    }
  });
});

app.get('/facts/:empresa/:poId', async (req, res) => {
  const { empresa, poId } = req.params;
  if (!empresa.match(/^Empresa\d+$/)) {
    return res.status(400).json({ success: false, message: 'Nombre de empresa inválido' });
  }
  const tables = getCompanyTables(empresa);
  const { fsPath, fbPath } = getDatabasePaths(empresa);
  if (!fs.existsSync(fsPath)) {
    return res.status(404).json({ success: false, message: `Base de datos no encontrada para ${empresa}` });
  }
  const options = { ...baseOptions, database: fbPath };
  Firebird.attach(options, async (err, db) => {
    if (err) {
      console.error('Error conectando a la BD:', err);
      return res.status(500).json({ success: false, message: 'Error conectando a la BD: ' + err.message });
    }
    try {
      const rows = await queryWithTimeout(db, `
        SELECT TRIM(f.CVE_DOC) AS id, COALESCE(SUM(pf.TOT_PARTIDA), 0) as monto
        FROM ${tables.FACTF} f
        LEFT JOIN ${tables.PAR_FACTF} pf ON f.CVE_DOC = pf.CVE_DOC
        WHERE f.CVE_PEDI LIKE ?
        AND f.STATUS <> 'C'
        GROUP BY f.CVE_DOC
      `, [poId + '%']);
      res.json({ success: true, facts: rows });
    } catch (e) {
      res.status(500).json({ success: false, message: 'Error consultando facturas: ' + e.message });
    } finally {
      db.detach();
    }
  });
});

app.post('/report', (req, res) => {
  const { ssitel, po, rems, facts } = req.body;
  const pdfMake = require('pdfmake/build/pdfmake');
  const pdfFonts = require('pdfmake/build/vfs_fonts');
  pdfMake.vfs = pdfFonts.pdfMake.vfs;

  const totalPO = po.total;
  const totalRem = rems.reduce((sum, r) => sum + r.monto, 0);
  const totalFac = facts.reduce((sum, f) => sum + f.monto, 0);
  const totalConsumo = totalRem + totalFac;
  const porcRem = (totalRem / totalPO * 100).toFixed(1);
  const porcFac = (totalFac / totalPO * 100).toFixed(1);
  const porcRest = (100 - parseFloat(porcRem) - parseFloat(porcFac)).toFixed(1);

  rems.forEach(r => r.porcentaje = (r.monto / totalPO * 100).toFixed(1));
  facts.forEach(f => f.porcentaje = (f.monto / totalPO * 100).toFixed(1));

  const docDefinition = {
    content: [
      { text: `Reporte de POs - Consumo (${ssitel})`, style: 'header' },
      {
        table: { body: [['PO ID', 'Fecha', 'Total'], [po.id, po.fecha, `$${totalPO.toLocaleString()}`]] },
        layout: 'lightHorizontalLines'
      },
      { text: 'Remisiones', style: 'subheader', color: 'blue' },
      {
        table: { body: [['ID', 'Monto (%)'], ...rems.map(r => [r.id, `$${r.monto.toLocaleString()} (${r.porcentaje}%)`])] },
        layout: { fillColor: rowIndex => (rowIndex === 0 ? '#e3f2fd' : null) }
      },
      { text: 'Facturas', style: 'subheader', color: 'red' },
      {
        table: { body: [['ID', 'Monto (%)'], ...facts.map(f => [f.id, `$${f.monto.toLocaleString()} (${f.porcentaje}%)`])] },
        layout: { fillColor: rowIndex => (rowIndex === 0 ? '#ffebee' : null) }
      },
      {
        layout: 'noBorders',
        table: {
  body: [[
    {
      text: `Consumido: $${totalConsumo.toLocaleString()} (${((totalConsumo / totalPO) * 100).toFixed(1)}%)`,
      colSpan: 2
    },
    // si usas colSpan en pdfmake, agrega una celda vacía como “placeholder”
    {},
    {
      text: `Restante: $${(totalPO - totalConsumo).toLocaleString()} (${(100 - ((totalConsumo / totalPO) * 100)).toFixed(1)}%)`,
      colSpan: 2
    },
    {}
  ]]
}

      },
      // Barra apilada (canvas)
      { canvas: [{ type: 'rect', x: 40, y: 50, w: 520, h: 36, r: 0, lineWidth: 1, lineColor: '#ddd' }] },
      { canvas: [{ type: 'rect', x: 40, y: 50, w: (porcRem / 100 * 520), h: 36, color: 'blue' }] },
      { canvas: [{ type: 'rect', x: 40 + (porcRem / 100 * 520), y: 50, w: (porcFac / 100 * 520), h: 36, color: 'red' }] },
      { canvas: [{ type: 'rect', x: 40 + ((parseFloat(porcRem) + parseFloat(porcFac)) / 100 * 520), y: 50, w: (porcRest / 100 * 520), h: 36, color: 'green' }] },
      { text: `Rem: ${porcRem}% ($${totalRem.toLocaleString()})`, color: 'blue' },
      { text: `Fac: ${porcFac}% ($${totalFac.toLocaleString()})`, color: 'red' },
      { text: `Restante: ${porcRest}% ($${(totalPO - totalConsumo).toLocaleString()})`, color: 'green' },
      { text: 'Observaciones: *Sin extensiones. Para "-2", ver gráfico adicional.', italics: true },
      { text: 'POrpt • Aspel SAE 9', style: 'footer' }
    ],
    styles: { header: { fontSize: 20, bold: true }, subheader: { fontSize: 14, bold: true }, footer: { fontSize: 10, alignment: 'center' } }
  };

  pdfMake.createPdf(docDefinition).getBuffer((buffer) => {
    res.set('Content-Type', 'application/pdf');
    res.set('Content-Disposition', 'attachment; filename=reporte.pdf');
    res.send(buffer);
  });
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