const fs = require('fs');
const path = require('path');

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

const config = require('./jasper.config');

let instance = null;
let initPromise = null;
let compiled = false;

/**
 * Crea directorio si no existe
 */
function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`✓ Directorio creado: ${dirPath}`);
  }
}

/**
 * Compila reporte JRXML a .jasper si es necesario
 */
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

  // Solo compilar si el .jasper no existe o si el JRXML es más nuevo
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

/**
 * Inicializa JasperReports
 */
async function init() {
  if (instance) {
    return instance;
  }
  
  if (initPromise) {
    return initPromise;
  }

  initPromise = (async () => {
    try {
      // Verificar que node-jasper esté disponible
      if (!jasperFactory) {
        console.error('JasperReports no disponible. Error de carga:', jasperLoadError?.message);
        return null;
      }

      // Crear directorios necesarios
      ensureDir(config.jasperPath);
      ensureDir(config.templatesPath);

      console.log('Inicializando JasperReports...');

      // Configurar opciones
      const options = {
        path: config.jasperPath,
        reports: config.reports,
        dataSources: config.dataSources || {}
      };

      // Agregar driver Jaybird si existe
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

      // Crear instancia
      instance = jasperFactory(options);

      if (!instance) {
        throw new Error('jasperFactory retornó null o undefined');
      }

      console.log('✓ Instancia de JasperReports creada');

      // Compilar reportes si es necesario
      if (!compiled) {
        console.log('Verificando compilación de reportes...');
        const reportNames = Object.keys(config.reports);
        
        for (const reportName of reportNames) {
          const reportCfg = config.reports[reportName];
          try {
            compileReport(instance, reportCfg);
          } catch (err) {
            console.error(`Error compilando reporte ${reportName}:`, err.message);
            // Continuar con otros reportes
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
    }
  })();

  return await initPromise;
}

/**
 * Obtiene la instancia actual (puede ser null)
 */
function getInstance() {
  return instance;
}

/**
 * Verifica si JasperReports está disponible
 */
function isAvailable() {
  return instance !== null;
}

/**
 * Genera reporte PDF de resumen de PO
 */
async function generatePoSummary(jasperInstance, summary) {
  if (!jasperInstance) {
    throw new Error('JasperReports no está inicializado. Verifica la instalación de node-jasper y Java JDK.');
  }

  console.log('Generando reporte para PO:', summary.baseId);

  const payload = { summary };

  // Configurar data source
  if (jasperInstance.dataSources && config.dataSourceName) {
    jasperInstance.dataSources[config.dataSourceName] = {
      driver: 'json',
      data: payload,
      jsonQuery: config.dataSources?.[config.dataSourceName]?.jsonQuery || 'summary.items'
    };
  }

  return await new Promise((resolve, reject) => {
    const reportConfig = {
      report: config.defaultReport,
      dataSource: {
        type: 'json',
        name: 'SummaryData',
        jsonQuery: 'summary.items',
        data: JSON.stringify(payload)
      },
      parameters: {
        REPORT_TITLE: `POrpt • ${summary.companyName || 'Sin empresa'}`,
        COMPANY_NAME: summary.companyName || 'SSITEL',
        BASE_ID: summary.baseId || '-',
        TOTALS_TEXT: summary.totalsTexto || '',
        GENERAL_ALERTS: summary.alertasTexto || ''
      }
    };

    jasperInstance.pdf(reportConfig, (err, buffer) => {
      if (err) {
        console.error('Error generando PDF:', err.message);
        return reject(new Error(`Error generando reporte: ${err.message}`));
      }
      
      console.log('✓ Reporte PDF generado correctamente');
      resolve(buffer);
    });
  });
}

/**
 * Reinicia la instancia (útil para desarrollo)
 */
async function reset() {
  console.log('Reiniciando JasperReports...');
  instance = null;
  initPromise = null;
  compiled = false;
  return await init();
}

module.exports = {
  init,
  getInstance,
  isAvailable,
  generatePoSummary,
  reset
};