const fs = require('fs');
const path = require('path');
const settingsStore = require('./settings-store');
const configModule = require('./jasper.config');

let jasperFactory = null;
let jasperLoadError = null;

try {
  jasperFactory = require('node-jasper');
  console.log('✓ node-jasper cargado correctamente');
} catch (err) {
  jasperLoadError = err;
  console.error('✗ Error cargando node-jasper:', err.message);
  console.log('\nPara habilitar JasperReports:');
  console.log('1. Instala la dependencia: npm install node-jasper');
  console.log('2. Asegúrate de tener Java JDK instalado (java -version)');
  console.log('3. Reinicia el servidor\n');
}

let config = configModule.get();
let instance = null;
let initPromise = null;
let compiled = false;

function refreshConfig() {
  config = configModule.load(settingsStore.getSettings());
  compiled = false;
}

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`✓ Directorio creado: ${dirPath}`);
  }
}

function compileReport(target, reportCfg) {
  if (!reportCfg || !reportCfg.jrxml || !reportCfg.jasper) {
    console.warn('Configuración de reporte incompleta:', reportCfg);
    return;
  }

  if (!fs.existsSync(reportCfg.jrxml)) {
    throw new Error(`No se encontró la plantilla JRXML en ${reportCfg.jrxml}`);
  }

  const jasperExists = fs.existsSync(reportCfg.jasper);
  const jrxmlStat = fs.statSync(reportCfg.jrxml);
  const jasperStat = jasperExists ? fs.statSync(reportCfg.jasper) : null;

  if (jasperExists && jasperStat && jasperStat.mtimeMs >= jrxmlStat.mtimeMs) {
    console.log(`✓ Reporte ya compilado: ${path.basename(reportCfg.jasper)}`);
    return;
  }

  console.log(`Compilando reporte: ${path.basename(reportCfg.jrxml)}...`);

  if (typeof target.compile === 'function') {
    try {
      target.compile(reportCfg.jrxml, reportCfg.jasper);
      console.log(`✓ Reporte compilado exitosamente: ${path.basename(reportCfg.jasper)}`);
    } catch (err) {
      console.error(`✗ Error compilando reporte ${path.basename(reportCfg.jrxml)}:`, err.message);
      throw err;
    }
  } else {
    console.warn('Método compile() no disponible en la instancia de jasper');
  }
}

async function init() {
  if (instance) {
    return instance;
  }

  if (initPromise) {
    return initPromise;
  }

  const { jasper } = settingsStore.getSettings();
  if (jasper && jasper.enabled === false) {
    console.warn('JasperReports deshabilitado por configuración.');
    instance = null;
    compiled = false;
    return null;
  }

  initPromise = (async () => {
    try {
      refreshConfig();

      if (!jasperFactory) {
        console.error('JasperReports no disponible. Error de carga:', jasperLoadError?.message);
        return null;
      }

      ensureDir(config.jasperPath);
      ensureDir(config.templatesPath);

      console.log('• Directorio de compilados Jasper:', config.jasperPath);
      console.log('• Directorio de plantillas Jasper:', config.templatesPath);
      if (config.fontsPath) {
        console.log('• Directorio de fuentes Jasper:', config.fontsPath);
      }

      console.log('Inicializando JasperReports...');

      const options = {
        path: config.jasperPath,
        reports: config.reports,
        dataSources: config.dataSources || {}
      };

      if (config.drivers && config.drivers.jaybird) {
        const driver = config.drivers.jaybird;
        if (fs.existsSync(driver.jar)) {
          options.drivers = {
            jaybird: {
              path: driver.jar,
              className: driver.className
            }
          };
          console.log('✓ Driver Jaybird configurado');
        } else {
          console.warn(`⚠ Driver Jaybird no encontrado en ${driver.jar}`);
        }
      }

      instance = jasperFactory(options);

      if (!instance) {
        throw new Error('jasperFactory retornó null o undefined');
      }

      console.log('✓ Instancia de JasperReports creada');

      if (!compiled) {
        console.log('Verificando compilación de reportes...');
        const reportNames = Object.keys(config.reports);

        for (const reportName of reportNames) {
          const reportCfg = config.reports[reportName];
          try {
            compileReport(instance, reportCfg);
          } catch (err) {
            console.error(`Error compilando reporte ${reportName}:`, err.message);
          }
        }

        compiled = true;
        console.log('✓ Proceso de compilación completado');
      }

      console.log('✓ JasperReports inicializado correctamente\n');
      return instance;
    } catch (err) {
      console.error('✗ Error inicializando JasperReports:', err.message);
      console.error('Stack:', err.stack);
      instance = null;
      return null;
    } finally {
      initPromise = null;
    }
  })();

  return await initPromise;
}

function getInstance() {
  return instance;
}

function isAvailable() {
  return instance !== null;
}

async function reload() {
  instance = null;
  initPromise = null;
  compiled = false;
  return await init();
}

async function generatePoSummary(jasperInstance, summary) {
  if (!jasperInstance) {
    throw new Error('JasperReports no está inicializado. Verifica la instalación de node-jasper y Java JDK.');
  }

  console.log('Generando reporte para PO:', summary.baseId);

  const payload = { summary };

  if (jasperInstance.dataSources && config.dataSourceName) {
    jasperInstance.dataSources[config.dataSourceName] = {
      driver: 'json',
      data: payload,
      jsonQuery: config.dataSources[config.dataSourceName]?.jsonQuery || 'summary.items'
    };
  }

  const reportName = config.defaultReport;
  const reportConfig = config.reports[reportName];

  if (!reportConfig) {
    throw new Error(`No se encontró la configuración del reporte ${reportName}`);
  }

  return await new Promise((resolve, reject) => {
    try {
      jasperInstance.pdf(reportConfig, (err, buffer) => {
        if (err) {
          return reject(err);
        }
        resolve(buffer);
      });
    } catch (err) {
      reject(err);
    }
  });
}

module.exports = {
  init,
  reload,
  getInstance,
  isAvailable,
  generatePoSummary
};
