const fs = require('fs');
const path = require('path');

const SETTINGS_FILE = path.join(__dirname, 'report-settings.json');
const ALLOWED_ENGINES = ['simple-pdf'];
const ALLOWED_FORMATS = ['pdf', 'xlsx', 'csv', 'json'];

const defaultSettings = {
  defaultEngine: 'simple-pdf',
  export: {
    defaultFormat: 'pdf',
    availableFormats: ['pdf', 'xlsx', 'csv', 'json']
  },
  customization: {
    includeCharts: true,
    includeMovements: true,
    includeObservations: true,
    includeUniverse: true,
    csv: {
      includePoResumen: true,
      includeRemisiones: true,
      includeFacturas: true,
      includeTotales: true,
      includeUniverseInfo: true
    }
  },
  branding: {
    headerTitle: 'Reporte de POs - Consumo',
    headerSubtitle: '',
    footerText: 'POrpt • Aspel SAE 9',
    letterheadEnabled: false,
    letterheadTop: '',
    letterheadBottom: '',
    remColor: '#2563eb',
    facColor: '#dc2626',
    restanteColor: '#16a34a',
    accentColor: '#1f2937',
    companyName: 'SSITEL'
  }
};

let settingsCache = null;

function ensureFileExists() {
  try {
    if (!fs.existsSync(SETTINGS_FILE)) {
      fs.writeFileSync(SETTINGS_FILE, JSON.stringify(defaultSettings, null, 2), 'utf8');
    }
  } catch (err) {
    console.warn('No se pudo garantizar la creación de report-settings.json:', err.message);
  }
}

function sanitizeBrandingConfig(branding = {}) {
  const sanitizeString = value => (typeof value === 'string' ? value.trim() : '');
  const sanitizeColor = (value, fallback) => {
    if (typeof value !== 'string') return fallback;
    const hex = value.trim();
    return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(hex) ? hex : fallback;
  };
  return {
    ...defaultSettings.branding,
    ...branding,
    headerTitle: sanitizeString(branding.headerTitle) || defaultSettings.branding.headerTitle,
    headerSubtitle: sanitizeString(branding.headerSubtitle),
    footerText: sanitizeString(branding.footerText) || defaultSettings.branding.footerText,
    letterheadEnabled: branding.letterheadEnabled === true,
    letterheadTop: sanitizeString(branding.letterheadTop),
    letterheadBottom: sanitizeString(branding.letterheadBottom),
    remColor: sanitizeColor(branding.remColor, defaultSettings.branding.remColor),
    facColor: sanitizeColor(branding.facColor, defaultSettings.branding.facColor),
    restanteColor: sanitizeColor(branding.restanteColor, defaultSettings.branding.restanteColor),
    accentColor: sanitizeColor(branding.accentColor, defaultSettings.branding.accentColor),
    companyName: sanitizeString(branding.companyName) || defaultSettings.branding.companyName
  };
}

function sanitizeExportConfig(exportConfig = {}) {
  const userFormats = Array.isArray(exportConfig.availableFormats) ? exportConfig.availableFormats : [];
  const mergedFormats = [...userFormats, ...defaultSettings.export.availableFormats];
  const catalog = Array.from(new Set(mergedFormats)).filter(format => ALLOWED_FORMATS.includes(format));
  const unique = catalog.length ? catalog : ['pdf'];
  const defaultFormat = ALLOWED_FORMATS.includes(exportConfig.defaultFormat) ? exportConfig.defaultFormat : unique[0];
  return {
    defaultFormat,
    availableFormats: unique
  };
}

function sanitizeCsvCustomization(csv = {}) {
  return {
    includePoResumen: csv.includePoResumen !== false,
    includeRemisiones: csv.includeRemisiones !== false,
    includeFacturas: csv.includeFacturas !== false,
    includeTotales: csv.includeTotales !== false,
    includeUniverseInfo: csv.includeUniverseInfo !== false
  };
}

function sanitizeCustomizationConfig(customization = {}) {
  return {
    includeCharts: customization.includeCharts !== false,
    includeMovements: customization.includeMovements !== false,
    includeObservations: customization.includeObservations !== false,
    includeUniverse: customization.includeUniverse !== false,
    csv: sanitizeCsvCustomization(customization.csv || {})
  };
}

function normalizeSettings(rawSettings = {}) {
  const merged = {
    ...defaultSettings,
    ...rawSettings,
    export: sanitizeExportConfig(rawSettings.export),
    customization: sanitizeCustomizationConfig(rawSettings.customization),
    branding: sanitizeBrandingConfig(rawSettings.branding)
  };

  if (!ALLOWED_ENGINES.includes(merged.defaultEngine)) {
    merged.defaultEngine = defaultSettings.defaultEngine;
  }

  return merged;
}

function loadSettings() {
  if (settingsCache) {
    return settingsCache;
  }

  ensureFileExists();

  try {
    const fileContent = fs.readFileSync(SETTINGS_FILE, 'utf8');
    const parsed = JSON.parse(fileContent || '{}');
    settingsCache = normalizeSettings(parsed);
  } catch (err) {
    console.warn('No se pudo leer report-settings.json, usando valores predeterminados:', err.message);
    settingsCache = normalizeSettings(defaultSettings);
  }

  try {
    fs.writeFileSync(SETTINGS_FILE, JSON.stringify(settingsCache, null, 2), 'utf8');
  } catch (err) {
    console.warn('No se pudo escribir configuración normalizada de reportes:', err.message);
  }

  return settingsCache;
}

function getSettings() {
  return settingsCache || loadSettings();
}

function saveSettings(newSettings) {
  settingsCache = normalizeSettings(newSettings);
  try {
    fs.writeFileSync(SETTINGS_FILE, JSON.stringify(settingsCache, null, 2), 'utf8');
  } catch (err) {
    console.warn('No se pudo guardar la configuración de reportes:', err.message);
  }
  return settingsCache;
}

function updateSettings(partial = {}) {
  const current = getSettings();
  const updated = {
    ...current,
    ...partial,
    export: sanitizeExportConfig({
      ...current.export,
      ...(partial.export || {})
    }),
    customization: sanitizeCustomizationConfig({
      ...current.customization,
      ...(partial.customization || {})
    }),
    branding: sanitizeBrandingConfig({
      ...current.branding,
      ...(partial.branding || {})
    })
  };
  return saveSettings(updated);
}

module.exports = {
  ALLOWED_ENGINES,
  ALLOWED_FORMATS,
  SETTINGS_FILE,
  defaultSettings,
  loadSettings,
  getSettings,
  saveSettings,
  updateSettings
};
