const fs = require('fs');
let jasperFactory = null;
try {
  jasperFactory = require('node-jasper');
} catch (err) {
  console.warn('No fue posible cargar node-jasper. JasperReports quedará deshabilitado hasta instalar la dependencia.', err.message);
}
const config = require('./jasper.config');

let instance = null;
let initPromise = null;
let compiled = false;

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function compileReport(target, reportCfg) {
  if (!reportCfg || !reportCfg.jrxml || !reportCfg.jasper) return;
  if (!fs.existsSync(reportCfg.jrxml)) {
    throw new Error(`No se encontró la plantilla JRXML en ${reportCfg.jrxml}`);
  }
  const jasperExists = fs.existsSync(reportCfg.jasper);
  const jrxmlStat = fs.statSync(reportCfg.jrxml);
  const jasperStat = jasperExists ? fs.statSync(reportCfg.jasper) : null;
  if (jasperExists && jasperStat && jasperStat.mtimeMs >= jrxmlStat.mtimeMs) {
    return;
  }
  if (typeof target.compile === 'function') {
    target.compile(reportCfg.jrxml, reportCfg.jasper);
  }
}

async function init() {
  if (instance) return instance;
  if (initPromise) return initPromise;
  initPromise = (async () => {
    try {
      if (!jasperFactory) {
        return null;
      }
      ensureDir(config.jasperPath);
      const options = {
        path: config.jasperPath,
        reports: config.reports,
        dataSources: config.dataSources || {}
      };
      if (config.drivers && config.drivers.jaybird) {
        const driver = config.drivers.jaybird;
        if (fs.existsSync(driver.jar)) {
          options.drivers = { jaybird: { path: driver.jar, className: driver.className } };
        } else {
          console.warn(`No se encontró el driver Jaybird en ${driver.jar}. Añádelo para habilitar JDBC.`);
        }
      }
      instance = jasperFactory(options);
      if (instance && !compiled) {
        Object.values(config.reports).forEach(reportCfg => compileReport(instance, reportCfg));
        compiled = true;
      }
      return instance;
    } catch (err) {
      console.warn('No se pudo inicializar JasperReports:', err.message);
      instance = null;
      return null;
    }
  })();
  return await initPromise;
}

function getInstance() {
  return instance;
}

async function generatePoSummary(jasperInstance, summary) {
  if (!jasperInstance) {
    throw new Error('Instancia de JasperReports no disponible');
  }
  const payload = { summary };
  if (jasperInstance.dataSources && config.dataSourceName) {
    jasperInstance.dataSources[config.dataSourceName] = {
      driver: 'json',
      data: payload,
      jsonQuery: config.dataSources?.[config.dataSourceName]?.jsonQuery || 'summary.items'
    };
  }
  return await new Promise((resolve, reject) => {
    jasperInstance.pdf(
      {
        report: config.defaultReport,
        dataSource: {
          type: 'json',
          name: 'SummaryData',
          jsonQuery: 'summary.items',
          data: JSON.stringify(payload)
        },
        parameters: {
          REPORT_TITLE: `POrpt • ${summary.companyName}`,
          COMPANY_NAME: summary.companyName,
          BASE_ID: summary.baseId,
          TOTALS_TEXT: summary.totalsTexto,
          GENERAL_ALERTS: summary.alertasTexto
        }
      },
      (err, buffer) => {
        if (err) return reject(err);
        resolve(buffer);
      }
    );
  });
}

module.exports = {
  init,
  getInstance,
  generatePoSummary
};
