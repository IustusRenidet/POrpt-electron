const Firebird = require('node-firebird');
const path = require('path');
const fs = require('fs');

const BASE_PATH = 'C:\\Program Files (x86)\\Common Files\\Aspel\\Sistemas Aspel\\SAE9.00\\';
const OPTIONS = { host: 'localhost', port: 3050, user: 'SYSDBA', password: 'masterkey', lowercase_keys: false, role: null, pageSize: 4096 };

function getDbPath(empresaFolder) {
  const xx = empresaFolder.replace('Empresa', '');
  return path.join(BASE_PATH, empresaFolder, 'Datos', `SAE90EMPRE${xx}.FDB`);
}

async function getEmpresas() {
  const empresas = [];
  const folders = fs.readdirSync(BASE_PATH).filter(f => f.match(/^Empresa\d+$/));
  for (const folder of folders) {
    const dbPath = getDbPath(folder);
    if (fs.existsSync(dbPath)) {
      const ssitel = await getEmpresaName(dbPath);
      empresas.push(`${folder} - ${ssitel}`);
    }
  }
  return empresas;
}

async function getEmpresaName(dbPath) {
  return new Promise((resolve) => {
    Firebird.attach({ filename: dbPath, ...OPTIONS }, (err, db) => {
      if (err) return resolve('Empresa Desconocida');
      db.query("SELECT NOMBRE FROM CLIE26 WHERE CLAVE = '0000'", (err, rows) => {
        db.detach();
        if (err || !rows.length) return resolve('Empresa Desconocida');
        resolve(rows[0].NOMBRE);
      });
    });
  });
}

function getConnection(empresaFolder) {
  const dbPath = getDbPath(empresaFolder);
  return new Promise((resolve, reject) => {
    Firebird.attach({ filename: dbPath, ...OPTIONS }, (err, db) => {
      if (err) reject(err);
      else resolve(db);
    });
  });
}

async function getPOs(empresaFolder) {
  const db = await getConnection(empresaFolder);
  return new Promise((resolve) => {
    db.query(`
      SELECT
        f.CVE_DOC AS id,
        f.FECHA_DOC AS fecha,
        COALESCE(f.IMPORTE, 0) AS total,
        COALESCE(f.CAN_TOT, 0) AS subtotal
      FROM FACTP26 f
      WHERE f.STATUS <> 'C'
      ORDER BY f.FECHA_DOC DESC
    `, (err, rows) => {
      if (err) return resolve([]);
      const pos = [];
      rows.forEach(row => {
        pos.push({
          id: row.ID,
          fecha: row.FECHA,
          total: parseFloat(row.TOTAL),
          subtotal: parseFloat(row.SUBTOTAL)
        });
      });
      db.detach();
      resolve(pos);
    });
  });
}

async function getRemisionesForPO(db, poIdBase) {
  return new Promise((resolve) => {
    db.query(`
      SELECT
        r.CVE_DOC AS id,
        r.FECHA_DOC AS fecha,
        COALESCE(SUM(r.IMPORTE), 0) AS monto
      FROM FACTR26 r
      WHERE r.CVE_PEDI LIKE '${poIdBase}%'
        AND r.STATUS <> 'C'
      GROUP BY r.CVE_DOC, r.FECHA_DOC
    `, (err, rows) => {
      const rems = rows.map(row => ({
        id: row.ID,
        fecha: row.FECHA,
        monto: parseFloat(row.MONTO),
        porcentaje: 0
      }));
      resolve(rems);
    });
  });
}

async function getFacturasForPO(db, poIdBase) {
  // Similar a remisiones, pero FACTF26 / PAR_FACTF26
  return new Promise((resolve) => {
    db.query(`
      SELECT
        f.CVE_DOC AS id,
        f.FECHA_DOC AS fecha,
        COALESCE(SUM(f.IMPORTE), 0) AS monto
      FROM FACTF26 f
      WHERE f.CVE_PEDI LIKE '${poIdBase}%'
        AND f.STATUS <> 'C'
      GROUP BY f.CVE_DOC, f.FECHA_DOC
    `, (err, rows) => {
      const facts = rows.map(row => ({
        id: row.ID,
        fecha: row.FECHA,
        monto: parseFloat(row.MONTO),
        porcentaje: 0
      }));
      resolve(facts);
    });
  });
}

// Para extensiones: getPOWithExtensions(empresa, baseId) - agrupa LIKE '%', suma totals, sub para "-2"
async function getPOWithExtensions(empresaFolder, baseId) {
  const db = await getConnection(empresaFolder);
  // Query LIKE baseId + '%', GROUP BY SUBSTRING(id, 1, LENGTH(baseId))
  // Retorna { base: {...}, ext2: {...}, total: {...} }
  db.detach();
  return {};  // Expande con lógica similar
}

module.exports = { getEmpresas, getPOs, getRemisionesForPO, getFacturasForPO, getPOWithExtensions, getConnection };