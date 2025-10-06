const fs = require('fs');
const os = require('os');
const path = require('path');
const settingsStore = require('./settings-store');

const envOrDefault = (envName, fallback) => {
  const value = process.env[envName];
  if (!value) {
    return fallback;
  }
  return path.isAbsolute(value) ? value : path.resolve(value);
};

function buildConfig(overrides = {}) {
  const jasperOverrides = overrides.jasper || settingsStore.getSettings().jasper || {};

  const basePath = envOrDefault('JASPER_BASE_PATH', __dirname);
  const isPackaged = basePath.includes('.asar');
  const fallbackWritableBase = envOrDefault(
    'JASPER_WRITABLE_BASE',
    path.join(os.tmpdir(), 'porpt-electron', 'jasper')
  );

  const compiledPath = jasperOverrides.compiledDir || envOrDefault(
    'JASPER_COMPILED_DIR',
    isPackaged ? path.join(fallbackWritableBase, 'compiled') : path.join(basePath, 'compiled')
  );

  const templatesPath = jasperOverrides.templatesDir || envOrDefault(
    'JASPER_TEMPLATES_DIR',
    path.join(basePath, 'templates')
  );

  const fontsPath = jasperOverrides.fontsDir || envOrDefault(
    'JASPER_FONTS_DIR',
    path.join(basePath, 'fonts')
  );

  const summaryReportBase = process.env.JASPER_SUMMARY_REPORT_BASE || 'po_summary';
  const summaryReportName = jasperOverrides.defaultReport || process.env.JASPER_SUMMARY_REPORT || 'poSummary';
  const dataSourceName = jasperOverrides.dataSourceName || process.env.JASPER_DATASOURCE_NAME || 'poSummaryJson';
  const jsonQuery = jasperOverrides.jsonQuery || process.env.JASPER_JSON_QUERY || 'summary.items';

  const jaybirdJar = envOrDefault(
    'JAYBIRD_JAR_PATH',
    path.join(basePath, 'lib', 'jaybird-full-5.0.1.jar')
  );

  try {
    const compiledParent = path.dirname(compiledPath);
    if (!fs.existsSync(compiledParent)) {
      fs.mkdirSync(compiledParent, { recursive: true });
    }
  } catch (err) {
    console.warn('⚠ No se pudo preparar el directorio para los compilados de Jasper:', err.message);
  }

  return {
    jasperPath: compiledPath,
    templatesPath,
    fontsPath,
    defaultReport: summaryReportName,
    reports: {
      [summaryReportName]: {
        jasper: path.join(compiledPath, `${summaryReportBase}.jasper`),
        jrxml: path.join(templatesPath, `${summaryReportBase}.jrxml`),
        conn: 'poSummaryJson'
      }
    },
    dataSourceName,
    dataSources: {
      [dataSourceName]: {
        driver: 'json',
        jsonQuery,
        data: []
      }
    },
    drivers: {
      jaybird: {
        jar: jaybirdJar,
        className: process.env.JAYBIRD_CLASS_NAME || 'org.firebirdsql.jdbc.FBDriver'
      }
    }
  };
}

let cachedConfig = null;

function load(overrides = {}) {
  cachedConfig = buildConfig(overrides);
  return cachedConfig;
}

function get() {
  if (!cachedConfig) {
    cachedConfig = buildConfig(settingsStore.getSettings());
  }
  return cachedConfig;
}

module.exports = {
  load,
  get
};
