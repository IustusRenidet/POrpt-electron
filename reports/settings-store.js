const fs = require('fs');
const path = require('path');

const SETTINGS_FILE = path.join(__dirname, 'report-settings.json');
const ALLOWED_ENGINES = ['jasper', 'simple-pdf'];

const defaultSettings = {
  defaultEngine: 'jasper',
  jasper: {
    enabled: true,
    compiledDir: '',
    templatesDir: '',
    fontsDir: '',
    defaultReport: '',
    dataSourceName: '',
    jsonQuery: ''
  },
  branding: {
    headerTitle: 'Reporte de PEOs - Consumo',
    headerSubtitle: '',
    footerText: 'PEOrpt • Aspel SAE 9',
    letterheadEnabled: false,
    letterheadTop: '',
    letterheadBottom: '',
    remColor: '#2563eb',
    facColor: '#dc2626',
    restanteColor: '#16a34a',
    accentColor: '#1f2937',
    companyName: 'SITTEL'
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

function sanitizeJasperConfig(jasper = {}) {
  const normalized = { ...defaultSettings.jasper, ...jasper };
  const sanitizeString = value => (typeof value === 'string' ? value.trim() : '');
  normalized.compiledDir = sanitizeString(normalized.compiledDir);
  normalized.templatesDir = sanitizeString(normalized.templatesDir);
  normalized.fontsDir = sanitizeString(normalized.fontsDir);
  normalized.defaultReport = sanitizeString(normalized.defaultReport) || 'poSummary';
  normalized.dataSourceName = sanitizeString(normalized.dataSourceName) || 'poSummaryJson';
  normalized.jsonQuery = sanitizeString(normalized.jsonQuery) || 'summary.items';
  normalized.enabled = normalized.enabled !== false;
  return normalized;
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

function normalizeSettings(rawSettings = {}) {
  const merged = {
    ...defaultSettings,
    ...rawSettings,
    jasper: sanitizeJasperConfig(rawSettings.jasper),
    branding: sanitizeBrandingConfig(rawSettings.branding)
  };

  if (!ALLOWED_ENGINES.includes(merged.defaultEngine)) {
    merged.defaultEngine = defaultSettings.defaultEngine;
  }

  if (merged.jasper.enabled === false && merged.defaultEngine === 'jasper') {
    merged.defaultEngine = 'simple-pdf';
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
    jasper: {
      ...current.jasper,
      ...(partial.jasper || {})
    },
    branding: sanitizeBrandingConfig({
      ...current.branding,
      ...(partial.branding || {})
    })
  };
  return saveSettings(updated);
}

module.exports = {
  ALLOWED_ENGINES,
  SETTINGS_FILE,
  defaultSettings,
  loadSettings,
  getSettings,
  saveSettings,
  updateSettings
};
