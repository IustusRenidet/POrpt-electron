const fs = require('fs');

let PDFDocument;
let pdfkitLoadError = null;
let SVGtoPDF;
let svgToPdfLoadError = null;
let svgRenderErrorLogged = false;

try {
  PDFDocument = require('pdfkit');
  console.log('✓ pdfkit cargado correctamente');
} catch (err) {
  pdfkitLoadError = err;
  console.error('✗ Error cargando pdfkit:', err.message);
  console.error('El motor "PDF directo" requiere la dependencia opcional "pdfkit".');
  console.error('Ejecuta "npm install" para instalar las dependencias del proyecto.');
  console.error('Si estás en Windows y ves errores de compilación puedes usar "npm install --omit=optional".\n');
}

try {
  SVGtoPDF = require('svg-to-pdfkit');
  console.log('✓ svg-to-pdfkit cargado correctamente');
} catch (err) {
  svgToPdfLoadError = err;
  console.warn('⚠️ No se pudo cargar svg-to-pdfkit. Los íconos vectoriales no estarán disponibles:', err.message);
}

const PDF_UNAVAILABLE_MESSAGE =
  'El motor "PDF directo" no está disponible porque la dependencia opcional "pdfkit" no se pudo cargar. ' +
  'Ejecuta "npm install" (o "npm install --omit=optional" si necesitas omitir dependencias opcionales) y vuelve a iniciar la aplicación.';

const DEFAULT_BRANDING = {
  headerTitle: 'Reporte de POs - Consumo',
  headerSubtitle: '',
  footerText: 'POrpt • ICONET',
  letterheadEnabled: false,
  letterheadTop: '',
  letterheadBottom: '',
  remColor: '#2563eb',
  facColor: '#dc2626',
  restanteColor: '#16a34a',
  accentColor: '#1f2937',
  companyName: 'SSITEL'
};

const DEFAULT_CUSTOMIZATION = {
  includeSummary: true,
  includeDetail: true,
  includeCharts: true,
  includeMovements: true,
  includeObservations: true,
  includeUniverse: true
};

function ensurePdfkitAvailable() {
  if (PDFDocument) {
    return true;
  }
  const error = new Error(PDF_UNAVAILABLE_MESSAGE);
  error.code = 'PDFKIT_NOT_INSTALLED';
  if (pdfkitLoadError) {
    error.cause = pdfkitLoadError;
  }
  throw error;
}

function normalizeColor(value, fallback) {
  if (typeof value !== 'string') return fallback;
  const trimmed = value.trim();
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(trimmed) ? trimmed : fallback;
}

function normalizeBranding(branding = {}) {
  return {
    ...DEFAULT_BRANDING,
    ...branding,
    remColor: normalizeColor(branding.remColor, DEFAULT_BRANDING.remColor),
    facColor: normalizeColor(branding.facColor, DEFAULT_BRANDING.facColor),
    restanteColor: normalizeColor(branding.restanteColor, DEFAULT_BRANDING.restanteColor),
    accentColor: normalizeColor(branding.accentColor, DEFAULT_BRANDING.accentColor),
    headerTitle: typeof branding.headerTitle === 'string' && branding.headerTitle.trim()
      ? branding.headerTitle.trim()
      : DEFAULT_BRANDING.headerTitle,
    headerSubtitle: typeof branding.headerSubtitle === 'string' ? branding.headerSubtitle.trim() : '',
    footerText: typeof branding.footerText === 'string' && branding.footerText.trim()
      ? branding.footerText.trim()
      : DEFAULT_BRANDING.footerText,
    letterheadEnabled: branding.letterheadEnabled === true,
    letterheadTop: typeof branding.letterheadTop === 'string' ? branding.letterheadTop.trim() : '',
    letterheadBottom: typeof branding.letterheadBottom === 'string' ? branding.letterheadBottom.trim() : '',
    companyName:
      typeof branding.companyName === 'string' && branding.companyName.trim()
        ? branding.companyName.trim()
        : DEFAULT_BRANDING.companyName
  };
}

function normalizeCustomization(customization = {}) {
  return {
    ...DEFAULT_CUSTOMIZATION,
    includeSummary: customization.includeSummary !== false,
    includeDetail: customization.includeDetail !== false,
    includeCharts: customization.includeCharts !== false,
    includeMovements: customization.includeMovements !== false,
    includeObservations: customization.includeObservations !== false,
    includeUniverse: customization.includeUniverse !== false
  };
}

function fileExists(filePath) {
  if (!filePath) return false;
  try {
    return fs.existsSync(filePath);
  } catch (err) {
    return false;
  }
}

function formatCurrency(value) {
  const amount = Number(value || 0);
  return amount.toLocaleString('es-MX', {
    style: 'currency',
    currency: 'MXN',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function roundTo(value, decimals = 2) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) {
    return 0;
  }
  const factor = 10 ** decimals;
  return Math.round(number * factor) / factor;
}

function clampPercentage(value, max = 999.99) {
  const normalized = roundTo(value);
  if (!Number.isFinite(normalized)) {
    return 0;
  }
  const upperBound = Math.max(0, Number(max) || 0);
  const clamped = Math.min(upperBound > 0 ? upperBound : 0, Math.max(0, normalized));
  return roundTo(clamped);
}

function formatPercentage(value) {
  return `${roundTo(value, 2).toFixed(2)}%`;
}

function padTo2(value) {
  return value.toString().padStart(2, '0');
}

function formatDateLabel(value) {
  if (!value) {
    return '';
  }
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
    const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/u);
    if (isoMatch) {
      const [, year, month, day] = isoMatch;
      return `${year}-${month}-${day}`;
    }
    const normalized = trimmed.replace(/ /u, 'T');
    const parsed = new Date(normalized);
    if (!Number.isNaN(parsed.getTime())) {
      const year = parsed.getFullYear();
      const month = padTo2(parsed.getMonth() + 1);
      const day = padTo2(parsed.getDate());
      return `${year}-${month}-${day}`;
    }
  }
  const asDate = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(asDate.getTime())) {
    return '';
  }
  const year = asDate.getFullYear();
  const month = padTo2(asDate.getMonth() + 1);
  const day = padTo2(asDate.getDate());
  return `${year}-${month}-${day}`;
}

function normalizePoId(value) {
  if (typeof value !== 'string') return '';
  return value.trim();
}

function getBasePoId(poId) {
  const normalized = normalizePoId(poId);
  if (!normalized) {
    return '';
  }
  return normalized.replace(/-\d+$/u, '');
}

function normalizeAlertEntries(alerts = [], fallback = []) {
  const registry = new Set();
  const addEntry = value => {
    if (!value) return;
    const trimmed = value.replace(/\s+/gu, ' ').trim();
    if (!trimmed || trimmed.toLowerCase().startsWith('sin alertas')) {
      return;
    }
    const normalizedLabel = trimmed.replace(/\[warning\]/giu, '[ALERTA]');
    if (!registry.has(normalizedLabel)) {
      registry.add(normalizedLabel);
    }
  };

  alerts.forEach(alert => {
    if (!alert) return;
    if (typeof alert === 'string') {
      addEntry(alert);
      return;
    }
    const message = typeof alert.message === 'string' ? alert.message.trim() : '';
    if (!message) return;
    const type = typeof alert.type === 'string' ? alert.type.trim().toUpperCase() : 'INFO';
    addEntry(`[${type}] ${message}`);
  });

  const fallbackArray = Array.isArray(fallback)
    ? fallback
    : typeof fallback === 'string'
      ? fallback.split(/\n+/u)
      : [];
  fallbackArray.forEach(entry => addEntry(entry));

  return Array.from(registry);
}

const ALERT_SUMMARY_LABELS = {
  alerta: 'Alerta',
  warning: 'Alerta',
  critica: 'Crítica',
  critical: 'Crítica',
  danger: 'Crítica',
  info: 'Info',
  success: 'Éxito'
};

const ALERT_SUMMARY_SVGS = {
  alerta: `
    <svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 2L1 21h22L12 2z" fill="#F59E0B"/>
      <path d="M12 7a1 1 0 0 0-1 1v6a1 1 0 0 0 2 0V8a1 1 0 0 0-1-1zm-1 10v2h2v-2h-2z" fill="#FFFFFF"/>
    </svg>
  `,
  warning: `
    <svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 2L1 21h22L12 2z" fill="#F59E0B"/>
      <path d="M12 7a1 1 0 0 0-1 1v6a1 1 0 0 0 2 0V8a1 1 0 0 0-1-1zm-1 10v2h2v-2h-2z" fill="#FFFFFF"/>
    </svg>
  `,
  critica: `
    <svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 2a10 10 0 1 0 10 10A10.011 10.011 0 0 0 12 2z" fill="#DC2626"/>
      <path d="M12 7a1 1 0 0 0-1 1v6a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1zm-1 10v2h2v-2h-2z" fill="#FFFFFF"/>
    </svg>
  `,
  critical: `
    <svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 2a10 10 0 1 0 10 10A10.011 10.011 0 0 0 12 2z" fill="#DC2626"/>
      <path d="M12 7a1 1 0 0 0-1 1v6a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1zm-1 10v2h2v-2h-2z" fill="#FFFFFF"/>
    </svg>
  `,
  danger: `
    <svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 2a10 10 0 1 0 10 10A10.011 10.011 0 0 0 12 2z" fill="#DC2626"/>
      <path d="M12 7a1 1 0 0 0-1 1v6a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1zm-1 10v2h2v-2h-2z" fill="#FFFFFF"/>
    </svg>
  `
};

function formatAlertSeverityLabel(rawSeverity) {
  const key = typeof rawSeverity === 'string' ? rawSeverity.trim().toLowerCase() : '';
  if (!key) {
    return '';
  }
  if (ALERT_SUMMARY_LABELS[key]) {
    return ALERT_SUMMARY_LABELS[key];
  }
  return key.charAt(0).toUpperCase() + key.slice(1);
}

function summarizeAlertEntries(alerts, fallback) {
  const entries = normalizeAlertEntries(alerts, fallback);
  if (!entries.length) {
    return { text: 'Sin alertas registradas', severity: '', count: 0 };
  }
  const [first, ...rest] = entries;
  const match = first.match(/^\[([^\]]+)\]\s*(.*)$/u);
  const severityKey = match ? match[1].trim().toLowerCase() : '';
  const message = match ? match[2].trim() : first.trim();
  const severityLabel = formatAlertSeverityLabel(severityKey);
  const parts = [];
  if (severityLabel) {
    parts.push(`${severityLabel}:`);
  }
  if (message) {
    parts.push(message);
  }
  let summary = parts.join(' ').replace(/\s+/gu, ' ').trim();
  if (!summary) {
    summary = 'Alerta registrada';
  }
  if (rest.length > 0) {
    summary += ` (+${rest.length} más)`;
  }
  return { text: summary, severity: severityKey, count: entries.length, label: severityLabel };
}

function getAlertSummary(value) {
  if (value && typeof value === 'object') {
    const text = typeof value.text === 'string' ? value.text : String(value.text ?? '');
    const severity = typeof value.severity === 'string' ? value.severity.trim().toLowerCase() : '';
    return {
      text: text.trim(),
      severity,
      label: typeof value.label === 'string' ? value.label : '',
      count: Number.isFinite(value.count) ? value.count : undefined
    };
  }
  const text = typeof value === 'string' ? value : String(value ?? '');
  return { text: text.trim(), severity: '', label: '', count: undefined };
}

function getCellText(value) {
  if (value && typeof value === 'object') {
    const text = typeof value.text === 'string' ? value.text : String(value.text ?? '');
    return text;
  }
  return String(value ?? '');
}

function drawAlertCell(doc, value, cellX, cellY, cellWidth, rowHeight) {
  const summary = getAlertSummary(value);
  const paddingX = 10;
  const paddingY = 6;
  const availableWidth = Math.max(0, cellWidth - paddingX * 2);
  const baseX = cellX + paddingX;
  const baseY = cellY + paddingY;
  const iconSvg = summary.severity ? ALERT_SUMMARY_SVGS[summary.severity] : null;
  const hasSvgRenderer = typeof SVGtoPDF === 'function' && !svgToPdfLoadError;
  let textOffset = 0;

  if (hasSvgRenderer && iconSvg) {
    const iconSize = 12;
    const iconX = baseX;
    const iconY = cellY + Math.max(0, (rowHeight - iconSize) / 2);
    try {
      SVGtoPDF(doc, iconSvg, iconX, iconY, { width: iconSize, height: iconSize, assumePt: true });
      textOffset = iconSize + 4;
    } catch (err) {
      if (!svgRenderErrorLogged) {
        svgRenderErrorLogged = true;
        console.warn('⚠️ Error dibujando SVG en el PDF:', err.message);
      }
      textOffset = 0;
    }
  }

  const textX = baseX + textOffset;
  const textWidth = Math.max(0, availableWidth - textOffset);
  doc
    .font('Helvetica')
    .fontSize(11)
    .fillColor('#111827')
    .text(summary.text || '—', textX, baseY, { width: textWidth, align: 'left' });
}

function getContentBounds(doc) {
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const top = doc.page.margins.top;
  const bottom = doc.page.height - doc.page.margins.bottom;
  return {
    left,
    right,
    top,
    bottom,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top)
  };
}

function clampHorizontalRect(bounds, x, width) {
  const safeWidth = Math.max(0, Math.min(width, bounds.width));
  let safeX = Number.isFinite(x) ? x : bounds.left;
  if (safeX < bounds.left) {
    safeX = bounds.left;
  }
  if (safeX + safeWidth > bounds.right) {
    safeX = bounds.right - safeWidth;
  }
  if (safeX < bounds.left) {
    safeX = bounds.left;
  }
  return { x: safeX, width: safeWidth };
}

function computePercentages(totals = {}) {
  const total = Math.max(0, Number(totals.total || 0));
  const totalRem = Math.max(0, Number(totals.totalRem || 0));
  const totalFac = Math.max(0, Number(totals.totalFac || 0));
  const consumo = Math.max(0, totalRem + totalFac);
  const restanteFallback = total - consumo;
  const restanteAmountRaw = totals.restante != null ? Number(totals.restante) : restanteFallback;
  const restanteAmount = Number.isFinite(restanteAmountRaw) ? Math.max(0, restanteAmountRaw) : 0;
  if (total <= 0) {
    return { rem: 0, fac: 0, rest: 0, base: 0, overage: consumo };
  }
  const rem = clampPercentage((totalRem / total) * 100);
  const fac = clampPercentage((totalFac / total) * 100);
  const rest = clampPercentage((restanteAmount / total) * 100, 100);
  const overage = Math.max(0, consumo - total);
  return {
    rem,
    fac,
    rest,
    base: total,
    overage
  };
}

function computeSegmentScale(percentages = {}) {
  const rem = Math.max(0, Number(percentages.rem) || 0);
  const fac = Math.max(0, Number(percentages.fac) || 0);
  const rest = Math.max(0, Number(percentages.rest) || 0);
  const sum = rem + fac + rest;
  return Math.max(sum, 100);
}

function ensureSpace(doc, requiredHeight = 0, options = {}) {
  const bottomLimit = doc.page.height - doc.page.margins.bottom;
  if (doc.y + requiredHeight <= bottomLimit) {
    return false;
  }
  doc.addPage();
  if (typeof options.onAddPage === 'function') {
    options.onAddPage();
  }
  if (options.padding) {
    doc.moveDown(options.padding);
  }
  return true;
}

function drawLetterhead(doc, branding) {
  if (!branding.letterheadEnabled) return;
  const hasTop = fileExists(branding.letterheadTop);
  const hasBottom = fileExists(branding.letterheadBottom);
  if (!hasTop && !hasBottom) {
    return;
  }

  const drawTop = () => {
    if (!hasTop) return;
    try {
      const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      doc.image(branding.letterheadTop, doc.page.margins.left, doc.page.margins.top - 20, { width });
      doc.moveDown(1);
    } catch (err) {
      console.warn('No se pudo dibujar membrete superior:', err.message);
    }
  };

  const drawBottom = () => {
    if (!hasBottom) return;
    try {
      const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      const y = doc.page.height - doc.page.margins.bottom - 60;
      doc.image(branding.letterheadBottom, doc.page.margins.left, y, { width });
    } catch (err) {
      console.warn('No se pudo dibujar membrete inferior:', err.message);
    }
  };

  const drawLetterheadsForPage = () => {
    drawTop();
    drawBottom();
  };

  drawLetterheadsForPage();
  doc.on('pageAdded', drawLetterheadsForPage);
}

function drawHeader(doc, summary, branding) {
  const title = branding.headerTitle || DEFAULT_BRANDING.headerTitle;
  doc.font('Helvetica-Bold').fontSize(20).fillColor(branding.accentColor || '#111827').text(title, { align: 'center' });
  if (branding.headerSubtitle) {
    doc.moveDown(0.2);
    doc.font('Helvetica').fontSize(12).fillColor('#475569').text(branding.headerSubtitle, { align: 'center' });
  }
  doc.moveDown(0.4);
  const empresaLabel =
    branding.companyName || summary.companyName || summary.empresaLabel || summary.empresa || '';
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(`Empresa: ${empresaLabel}`, { align: 'center' });
  if (summary.universe?.isUniverse) {
    const titleText = summary.universe.title || 'Reporte del universo de POs';
    doc.moveDown(0.2);
    doc
      .font('Helvetica')
      .fontSize(12)
      .fillColor(branding.accentColor || '#1f2937')
      .text(titleText, { align: 'center' });
  }
  doc.moveDown(0.8);
}

function drawUniverseFilterInfo(doc, summary, branding) {
  ensureSpace(doc, 70);
  const universe = summary.universe || {};
  const label = universe.label || 'Global (todas las fechas)';
  const description = universe.description || '';
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(`Filtro aplicado: ${label}`, { align: 'center' });
  if (description) {
    doc.moveDown(0.2);
    doc.font('Helvetica').fontSize(11).fillColor('#475569').text(description, { align: 'center' });
  }
  doc.moveDown(0.8);
}

function drawTotalsTable(doc, summary, branding) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const remainderLabel = summary.universe?.isUniverse
    ? 'Remanente total del universo'
    : 'Remanente total de la selección';
  const rows = [
    { label: 'Total autorizado de POs', value: formatCurrency(totals.total), color: branding.accentColor },
    { label: 'Total remisiones', value: formatCurrency(totals.totalRem), color: branding.remColor },
    { label: 'Total facturas', value: formatCurrency(totals.totalFac), color: branding.facColor },
    { label: 'Total consumido (Rem + Fac)', value: formatCurrency(totals.totalConsumo), color: '#1f2937' },
    { label: remainderLabel, value: formatCurrency(totals.restante), color: branding.restanteColor }
  ];

  const labelColumnWidth = Math.max(0, width / 2 - 16);
  const valueColumnWidth = Math.max(0, width / 2 - 16);
  const verticalPadding = 20;

  const measuredRows = rows.map(row => {
    const labelHeight = doc.font('Helvetica-Bold').fontSize(12).heightOfString(row.label, {
      width: labelColumnWidth
    });
    const valueHeight = doc.font('Helvetica').fontSize(12).heightOfString(row.value, {
      width: valueColumnWidth,
      align: 'right'
    });
    const contentHeight = Math.max(labelHeight, valueHeight);
    const height = Math.max(32, Math.ceil(contentHeight + verticalPadding));
    return { ...row, height };
  });

  const totalHeight = measuredRows.reduce((sum, row) => sum + row.height, 0);
  ensureSpace(doc, totalHeight + 32);
  const adjustedTop = doc.y;
  doc.save();
  doc.lineWidth(1).rect(startX, adjustedTop, width, totalHeight).stroke('#d1d5db');
  let currentY = adjustedTop;
  measuredRows.forEach((row, index) => {
    if (index > 0) {
      doc.lineWidth(0.5).moveTo(startX, currentY).lineTo(startX + width, currentY).stroke('#e5e7eb');
    }
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor(row.color || branding.accentColor || '#1f2937')
      .text(row.label, startX + 12, currentY + 10, { width: labelColumnWidth });
    doc
      .font('Helvetica')
      .fontSize(12)
      .fillColor('#111827')
      .text(row.value, startX + width / 2, currentY + 10, { width: valueColumnWidth, align: 'right' });
    currentY += row.height;
  });
  doc.restore();
  doc.y = adjustedTop + totalHeight + 18;
}

function drawUniverseObservations(doc, summary) {
  const universe = summary.universe || {};
  const label = universe.label || 'Global';
  const company = summary.companyName || summary.empresaLabel || summary.empresa || 'SSITEL';
  const notes = [`Universo de ${company} con el filtro "${label}".`];
  if (summary.totals && summary.totals.total === 0) {
    notes.push('No se encontraron POs activas con el filtro aplicado.');
  }
  renderObservationBox(doc, Array.from(new Set(notes)), {
    alwaysShow: true,
    placeholderText: 'Sin observaciones registradas.'
  });
}

function drawUniverseGroupDetails(doc, summary, branding) {
  const groups = buildPoGroupDetails(summary);
  if (!groups.length) {
    return;
  }

  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;

  doc.moveDown(0.6);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text('Detalle por pedido del universo', startX, doc.y, { width });
  doc.moveDown(0.3);

  groups.forEach((group, index) => {
    drawGroupSection(doc, group, branding, { isLast: index === groups.length - 1 });
  });
}


function drawPoSummaryTable(doc, summary, branding) {
  const groups = buildPoGroupDetails(summary);
  const hasGroups = groups.length > 0;
  const startX = doc.page.margins.left;
  const tableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const headerHeight = 24;
  const rowPadding = 12;
  const columns = [
    { key: 'po', label: 'PO', width: tableWidth * 0.24, align: 'left' },
    { key: 'fecha', label: 'Fecha', width: tableWidth * 0.12, align: 'left' },
    { key: 'subtotal', label: 'Subtotal', width: tableWidth * 0.14, align: 'right' },
    { key: 'total', label: 'Total autorizado', width: tableWidth * 0.14, align: 'right' },
    { key: 'consumo', label: 'Consumido / Disponible', width: tableWidth * 0.20, align: 'right' },
    { key: 'alerts', label: 'Advertencias', width: tableWidth * 0.16, align: 'left' }
  ];
  const minimumRowHeight = rowPadding + 14;
  const minimumSectionContent = headerHeight + minimumRowHeight + 8;

  const drawHeader = nextRowHeight => {
    const requiredRowHeight = Math.max(minimumRowHeight, Number(nextRowHeight) || 0);
    ensureSpace(doc, headerHeight + requiredRowHeight + 6);
    const y = doc.y;
    doc.save();
    doc.lineWidth(1);
    doc.rect(startX, y, tableWidth, headerHeight).fillAndStroke('#111827', '#1f2937');
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(11);
    let offsetX = startX;
    columns.forEach(column => {
      doc.text(column.label, offsetX + 10, y + 6, { width: column.width - 20, align: column.align });
      offsetX += column.width;
    });
    doc.restore();
    doc.y = y + headerHeight;
  };

  const measureRowHeight = values => {
    const measuringFont = doc.font('Helvetica').fontSize(11);
    return values.reduce((height, value, index) => {
      const column = columns[index];
      const cellWidth = Math.max(0, column.width - 20);
      const text = getCellText(value);
      const cellHeight = measuringFont.heightOfString(text, { width: cellWidth });
      return Math.max(height, cellHeight + rowPadding);
    }, rowPadding);
  };

  ensureSpace(doc, minimumSectionContent);

  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text('Resumen de POs seleccionadas', startX, doc.y);
  doc.moveDown(0.3);

  const subtotalLookup = new Map();
  const rawItems = Array.isArray(summary.items) ? summary.items : [];
  rawItems.forEach(item => {
    if (!item || typeof item.id !== 'string') return;
    subtotalLookup.set(item.id, Number(item.subtotal || 0));
  });

  if (hasGroups) {
    let headerDrawn = false;
    groups.forEach(group => {
      group.items.forEach(item => {
        const totals = item.totals || {};
        const subtotalAmount = subtotalLookup.get(item.id) || 0;
        const totalAmount = totals.total;
        const consumido = Number.isFinite(totals.totalConsumo)
          ? totals.totalConsumo
          : roundTo((totals.totalRem || 0) + (totals.totalFac || 0));
        const disponible = Number.isFinite(totals.restante)
          ? totals.restante
          : roundTo(Math.max(totalAmount - consumido, 0));
        const isExtension = !item.isBase;
        const consumptionPercent = totalAmount > 0 ? clampPercentage((consumido / totalAmount) * 100) : 0;
        const alertSummary = `Consumido al ${formatPercentage(consumptionPercent)}`;
        const poLabel = isExtension
          ? `${item.id} · Ext. de ${group.baseId}`
          : `${item.id} (base)`;
        const rowValues = [
          poLabel,
          item.fecha || '-',
          subtotalAmount > 0 ? formatCurrency(subtotalAmount) : '—',
          formatCurrency(totalAmount),
          `${formatCurrency(consumido)} / ${formatCurrency(disponible)}`,
          alertSummary
        ];
        const rowHeight = measureRowHeight(rowValues);
        if (!headerDrawn) {
          drawHeader(rowHeight);
          headerDrawn = true;
        }
        if (ensureSpace(doc, rowHeight + 4)) {
          drawHeader(rowHeight);
          headerDrawn = true;
        }
        const y = doc.y;
        const background = isExtension ? '#ffffff' : '#f8fafc';
        doc.save().fillColor(background).rect(startX, y, tableWidth, rowHeight).fill().restore();
        doc.lineWidth(0.5).strokeColor('#d1d5db').rect(startX, y, tableWidth, rowHeight).stroke();
        let offsetX = startX;
        columns.forEach((column, index) => {
          const value = rowValues[index];
          if (column.key === 'alerts') {
            drawAlertCell(doc, value, offsetX, y, column.width, rowHeight);
          } else {
            doc
              .font('Helvetica')
              .fontSize(11)
              .fillColor('#111827')
              .text(getCellText(value), offsetX + 10, y + 6, { width: column.width - 20, align: column.align });
          }
          offsetX += column.width;
        });
        doc.y = y + rowHeight;
      });
    });
  } else {
    ensureSpace(doc, 40);
    doc
      .font('Helvetica')
      .fontSize(11)
      .fillColor('#4b5563')
      .text('No se encontraron POs para mostrar en la tabla general.', startX, doc.y, {
        width: tableWidth
      });
  }

  doc.moveDown(0.4);
  const totals = summary.totals || {};
  const totalConsumido = roundTo((totals.totalRem || 0) + (totals.totalFac || 0));
  doc.lineWidth(0.5).moveTo(startX, doc.y).lineTo(startX + tableWidth, doc.y).stroke('#d1d5db');
  doc.moveDown(0.3);
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor('#111827')
    .text('Totales del reporte', startX, doc.y);
  doc.moveDown(0.2);
  doc
    .font('Helvetica')
    .fontSize(11)
    .fillColor('#111827')
    .text(`Autorizado: ${formatCurrency(totals.total || 0)}`, startX, doc.y);
  doc.text(`Consumido: ${formatCurrency(totalConsumido)}`, startX, doc.y);
  doc.text(`Disponible: ${formatCurrency(totals.restante || 0)}`, startX, doc.y);
  doc.moveDown(0.6);
  return true;
}
function collectMovements(items, totals, key) {
  const grandTotal = totals.total || 0;
  const movements = [];
  items.forEach(item => {
    const list = item[key] || [];
    list.forEach(entry => {
      movements.push({
        id: entry.id,
        monto: Number(entry.monto || 0),
        porcentaje: grandTotal > 0 ? (Number(entry.monto || 0) / grandTotal) * 100 : 0
      });
    });
  });
  movements.sort((a, b) => b.monto - a.monto);
  const subtotal = movements.reduce((sum, entry) => sum + entry.monto, 0);
  const porcentaje = grandTotal > 0 ? (subtotal / grandTotal) * 100 : 0;
  return { movements, subtotal, porcentaje };
}

function drawMovements(doc, summary, branding) {
  const { totals } = summary;
  const rem = collectMovements(summary.items || [], totals, 'remisiones');
  const fac = collectMovements(summary.items || [], totals, 'facturas');
  const bounds = getContentBounds(doc);
  const gutter = 20;
  const columnWidth = Math.max(0, (bounds.width - gutter) / 2);
  const startX = bounds.left;
  const secondColumnX = columnWidth > 0 ? Math.min(bounds.right - columnWidth, startX + columnWidth + gutter) : startX;
  const availableHeight = doc.page.height - doc.page.margins.top - doc.page.margins.bottom - 120;
  const rowsPerChunk = Math.max(1, Math.floor((availableHeight - 36) / 16));
  const totalRows = Math.max(rem.movements.length, fac.movements.length);
  const totalChunks = Math.max(1, Math.ceil(totalRows / rowsPerChunk));

  for (let chunkIndex = 0; chunkIndex < totalChunks; chunkIndex += 1) {
    const start = chunkIndex * rowsPerChunk;
    const end = start + rowsPerChunk;
    const remSlice = rem.movements.slice(start, end);
    const facSlice = fac.movements.slice(start, end);
    const rowsInChunk = Math.max(remSlice.length, facSlice.length);
    const boxHeight = rowsInChunk * 16 + 36;
    const requiredHeight = boxHeight + 90;
    ensureSpace(doc, requiredHeight);
    const top = doc.y;

    doc
      .font('Helvetica-Bold')
      .fontSize(14)
      .fillColor(branding.remColor)
      .text('Remisiones', startX, top);
    doc
      .font('Helvetica-Bold')
      .fontSize(14)
      .fillColor(branding.facColor)
      .text('Facturas', secondColumnX, top);

    doc.lineWidth(1);
    doc.rect(startX, top + 18, columnWidth, boxHeight).stroke(branding.remColor);
    doc.rect(secondColumnX, top + 18, columnWidth, boxHeight).stroke(branding.facColor);

    doc.font('Helvetica').fontSize(11).fillColor('#111827');
    remSlice.forEach((item, index) => {
      const y = top + 30 + index * 16;
      doc.text(`${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`, startX + 10, y, {
        width: Math.max(0, columnWidth - 20)
      });
    });
    if (remSlice.length === 0) {
      doc
        .font('Helvetica')
        .fontSize(11)
        .fillColor('#6b7280')
        .text('Sin remisiones destacadas en este bloque.', startX + 10, top + 32, {
          width: Math.max(0, columnWidth - 20)
        });
    }

    doc.font('Helvetica').fillColor('#111827');
    facSlice.forEach((item, index) => {
      const y = top + 30 + index * 16;
      doc.text(
        `${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`,
        secondColumnX + 10,
        y,
        {
          width: Math.max(0, columnWidth - 20)
        }
      );
    });
    if (facSlice.length === 0) {
      doc
        .font('Helvetica')
        .fontSize(11)
        .fillColor('#6b7280')
        .text('Sin facturas destacadas en este bloque.', secondColumnX + 10, top + 32, {
          width: Math.max(0, columnWidth - 20)
        });
    }

    const footerY = top + 18 + boxHeight - 18;
    const isLastChunk = chunkIndex === totalChunks - 1;
    if (isLastChunk) {
      doc
        .font('Helvetica-Bold')
        .fillColor('#111827')
        .text(
          `Subtotal Remisiones: ${formatCurrency(rem.subtotal)} (${formatPercentage(rem.porcentaje)})`,
          startX + 10,
          footerY,
          { width: Math.max(0, columnWidth - 20) }
        );
      doc
        .font('Helvetica-Bold')
        .fillColor('#111827')
        .text(
          `Subtotal Facturas: ${formatCurrency(fac.subtotal)} (${formatPercentage(fac.porcentaje)})`,
          secondColumnX + 10,
          footerY,
          { width: Math.max(0, columnWidth - 20) }
        );
    } else {
      doc
        .font('Helvetica')
        .fontSize(10)
        .fillColor('#6b7280')
        .text('Continúa en la siguiente página…', startX + 10, footerY, { width: Math.max(0, columnWidth - 20) });
      doc
        .font('Helvetica')
        .fontSize(10)
        .fillColor('#6b7280')
        .text('Continúa en la siguiente página…', secondColumnX + 10, footerY, {
          width: Math.max(0, columnWidth - 20)
        });
    }
    doc.y = top + 18 + boxHeight + 12;
  }
}

function drawSummaryBox(doc, summary, branding = {}) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  ensureSpace(doc, 90);
  const top = doc.y;
  const percentages = computePercentages(totals);
  const totalAuthorized = Math.max(0, Number(totals.total || 0));
  const totalRem = Math.max(0, Number(totals.totalRem || 0));
  const totalFac = Math.max(0, Number(totals.totalFac || 0));
  const totalConsumo = roundTo(totalRem + totalFac);
  const consumoPercentage = totalAuthorized > 0 ? roundTo((totalConsumo / totalAuthorized) * 100) : 0;
  const accent = branding.accentColor || '#111827';
  const columnWidth = width / 3;
  const cards = [
    {
      title: 'Autorizado',
      value: formatCurrency(totals.total),
      detail: summary.selectedIds && summary.selectedIds.length > 1
        ? `${summary.selectedIds.length} bases combinadas`
        : 'Monto aprobado'
    },
    {
      title: 'Consumido',
      value: formatCurrency(totals.totalConsumo),
      detail: `${formatPercentage(percentages.rem)} Rem · ${formatPercentage(percentages.fac)} Fac`
    },
    {
      title: 'Restante',
      value: formatCurrency(totals.restante),
      detail: `${formatPercentage(consumoPercentage)} consumido · ${formatPercentage(percentages.rest)} disponible`
    }
  ];
  doc.save();
  doc.lineWidth(1);
  doc.roundedRect(startX, top, width, 76, 8).fillAndStroke('#f8fafc', '#d1d5db');
  cards.forEach((card, index) => {
    const colX = startX + index * columnWidth;
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor(accent)
      .text(card.title, colX + 14, top + 12, { width: columnWidth - 28, align: 'left' });
    doc
      .font('Helvetica-Bold')
      .fontSize(16)
      .fillColor('#111827')
      .text(card.value, colX + 14, top + 28, { width: columnWidth - 28, align: 'left' });
    doc
      .font('Helvetica')
      .fontSize(10)
      .fillColor('#475569')
      .text(card.detail, colX + 14, top + 52, { width: columnWidth - 28, align: 'left' });
  });
  doc.restore();
  doc.y = top + 88;
}

function buildLegendEntries(totals, percentages, branding) {
  return [
    { label: `Remisiones: ${formatPercentage(percentages.rem)} (${formatCurrency(totals.totalRem)})`, color: branding.remColor },
    { label: `Facturas: ${formatPercentage(percentages.fac)} (${formatCurrency(totals.totalFac)})`, color: branding.facColor },
    { label: `Restante: ${formatPercentage(percentages.rest)} (${formatCurrency(totals.restante)})`, color: branding.restanteColor }
  ];
}

function drawLegendRows(doc, entries, startX, startY, maxWidth) {
  let currentY = startY;
  const bounds = getContentBounds(doc);
  const requestedWidth = Number.isFinite(maxWidth) && maxWidth > 0 ? maxWidth : bounds.width;
  const clamped = clampHorizontalRect(bounds, startX, requestedWidth);
  const safeStartX = clamped.x;
  const availableWidth = Math.max(0, bounds.right - safeStartX);
  const minimumWidth = Math.min(200, availableWidth);
  const effectiveWidth = Math.min(Math.max(clamped.width, minimumWidth), availableWidth);
  entries.forEach(entry => {
    doc.save().rect(safeStartX, currentY, 12, 12).fill(entry.color).restore();
    doc
      .font('Helvetica')
      .fontSize(11)
      .fillColor('#111827')
      .text(entry.label, safeStartX + 18, currentY - 1, { width: Math.max(0, effectiveWidth - 18) });
    currentY += 18;
  });
  return currentY;
}

function drawCombinedConsumptionBar(doc, totals, branding, options = {}) {
  const bounds = getContentBounds(doc);
  const availableWidth = Math.max(0, bounds.width);
  if (availableWidth <= 0) {
    return;
  }
  const requestedWidth = typeof options.maxWidth === 'number' && options.maxWidth > 0
    ? Math.min(options.maxWidth, availableWidth)
    : availableWidth;
  const { x: startX, width } = clampHorizontalRect(bounds, bounds.left, requestedWidth);
  const barHeight = options.barHeight || 26;
  const title = options.title || 'Consumo combinado';
  ensureSpace(doc, barHeight + 130);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text(title, bounds.left, doc.y, { width: bounds.width, align: 'center' });
  doc.moveDown(0.35);
  const barTop = doc.y;
  doc.lineWidth(1).roundedRect(startX, barTop, width, barHeight, 6).stroke('#d1d5db');
  const percentages = computePercentages(totals);
  const segments = [
    { value: percentages.rem, color: branding.remColor },
    { value: percentages.fac, color: branding.facColor },
    { value: percentages.rest, color: branding.restanteColor }
  ];
  const scale = computeSegmentScale(percentages);
  let currentX = startX;
  segments.forEach(segment => {
    const safeValue = Math.max(0, Number(segment.value) || 0);
    const segmentWidth = scale > 0 ? (width * safeValue) / scale : 0;
    if (segmentWidth <= 0) return;
    doc.save().rect(currentX, barTop, segmentWidth, barHeight).fillOpacity(0.9).fill(segment.color).restore();
    currentX += segmentWidth;
  });
  doc.y = barTop + barHeight + 12;
  const legendBottom = drawLegendRows(doc, buildLegendEntries(totals, percentages, branding), startX, doc.y, width);
  doc.y = legendBottom + 14;
}

function drawSmallConsumptionBar(doc, totals, branding, options = {}) {
  const bounds = getContentBounds(doc);
  const startX = options.startX ?? bounds.left;
  const width = Math.min(options.width ?? bounds.width, bounds.width);
  if (width <= 0) {
    return;
  }
  const barHeight = options.barHeight || 16;
  const legendGap = 6;
  const percentages = computePercentages(totals);
  const legendEntries = [
    { label: `Remisiones: ${formatPercentage(percentages.rem)} (${formatCurrency(totals.totalRem)})`, color: branding.remColor },
    { label: `Facturas: ${formatPercentage(percentages.fac)} (${formatCurrency(totals.totalFac)})`, color: branding.facColor },
    { label: `Disponible: ${formatPercentage(percentages.rest)} (${formatCurrency(totals.restante)})`, color: branding.restanteColor }
  ];
  const measuring = doc.font('Helvetica').fontSize(10);
  const legendHeight = legendEntries.reduce((sum, entry) => sum + measuring.heightOfString(entry.label, { width }) + 2, 0);
  const requiredHeight = barHeight + legendGap + legendHeight + 6;
  ensureSpace(doc, requiredHeight);
  const barTop = doc.y;
  doc.lineWidth(0.6).roundedRect(startX, barTop, width, barHeight, 5).stroke('#cbd5f5');
  let cursorX = startX;
  const segments = [
    { value: percentages.rem, color: branding.remColor },
    { value: percentages.fac, color: branding.facColor },
    { value: percentages.rest, color: branding.restanteColor }
  ];
  const scale = computeSegmentScale(percentages);
  segments.forEach(segment => {
    const safeValue = Math.max(0, Number(segment.value) || 0);
    const segmentWidth = scale > 0 ? (width * safeValue) / scale : 0;
    if (segmentWidth <= 0) return;
    doc.save().rect(cursorX, barTop, segmentWidth, barHeight).fillOpacity(0.95).fill(segment.color).restore();
    cursorX += segmentWidth;
  });
  doc.y = barTop + barHeight + legendGap;
  legendEntries.forEach(entry => {
    doc.font('Helvetica').fontSize(10).fillColor(entry.color).text(entry.label, startX, doc.y, { width });
    doc.moveDown(0.1);
  });
  doc.moveDown(0.2);
}

function normalizeItemTotals(item) {
  const totals = item?.totals || {};
  const authorized = Number(item?.total ?? totals.total ?? 0);
  const subtotal = Number(item?.subtotal ?? 0);
  const totalRem = Number(totals.totalRem ?? totals.rem ?? 0);
  const totalFac = Number(totals.totalFac ?? totals.fac ?? 0);
  const restanteRaw = totals.restante != null ? Number(totals.restante) : authorized - (totalRem + totalFac);
  const restante = roundTo(Math.max(restanteRaw, 0));
  const normalized = {
    total: roundTo(authorized),
    subtotal: roundTo(subtotal),
    totalRem: roundTo(totalRem),
    totalFac: roundTo(totalFac),
    totalConsumo: roundTo(totalRem + totalFac),
    restante
  };
  const percentages = computePercentages(normalized);
  percentages.consumo = clampPercentage(percentages.rem + percentages.fac);
  return { totals: normalized, percentages };
}

function buildPoGroupDetails(summary) {
  const items = Array.isArray(summary.items) ? summary.items : [];
  if (!items.length) {
    return [];
  }
  const groups = new Map();
  items.forEach(item => {
    const id = normalizePoId(item?.id);
    if (!id) return;
    const providedBase = normalizePoId(item?.baseId);
    const baseId = providedBase || getBasePoId(id) || id;
    if (!groups.has(baseId)) {
      groups.set(baseId, {
        baseId,
        ids: new Set(),
        items: [],
        total: 0,
        totalRem: 0,
        totalFac: 0,
        alerts: []
      });
    }
    const group = groups.get(baseId);
    const { totals, percentages } = normalizeItemTotals(item);
    group.ids.add(id);
    group.total += totals.total;
    group.totalRem += totals.totalRem;
    group.totalFac += totals.totalFac;
    const itemAlerts = Array.isArray(item.alerts) ? item.alerts : [];
    if (itemAlerts.length) {
      group.alerts.push(...itemAlerts);
    }
    const isBase = providedBase ? providedBase === id : getBasePoId(id) === baseId;
    const rawDate =
      item.fecha ?? item.FECHA_DOC ?? item.fecha_doc ?? item.fechaDoc ?? item.fechaDocumento ?? '';
    const remisiones = Array.isArray(item?.remisiones) ? item.remisiones : [];
    const facturas = Array.isArray(item?.facturas) ? item.facturas : [];
    const remClaves = remisiones
      .map(rem => (typeof rem?.id === 'string' ? rem.id.trim() : ''))
      .filter(Boolean);
    const facClaves = facturas
      .map(fac => (typeof fac?.id === 'string' ? fac.id.trim() : ''))
      .filter(Boolean);
    const claveParts = [];
    if (remClaves.length) {
      claveParts.push(`REM: ${remClaves.join(', ')}`);
    }
    if (facClaves.length) {
      claveParts.push(`FAC: ${facClaves.join(', ')}`);
    }
    const claveFallback = item.docSig || item.doc_sig || item.tipDoc || item.tip_doc || '';
    const claveLabel = claveParts.length ? claveParts.join(' | ') : claveFallback;
    group.items.push({
      id,
      fecha: formatDateLabel(rawDate),
      clave: claveLabel,
      totals,
      percentages,
      isBase,
      alerts: itemAlerts,
      alertasTexto: typeof item.alertasTexto === 'string' ? item.alertasTexto : ''
    });
  });
  return Array.from(groups.values())
    .map(group => {
      const orderedItems = group.items.sort((a, b) => {
        if (a.isBase === b.isBase) {
          return a.id.localeCompare(b.id);
        }
        return a.isBase ? -1 : 1;
      });
      const fallbackDate = orderedItems
        .map(item => item.fecha)
        .filter(Boolean)
        .sort()[0] || '';
      const normalizedItems = orderedItems.map(item => {
        if (!item.fecha && fallbackDate && item.isBase) {
          return { ...item, fecha: fallbackDate };
        }
        return item;
      });
      const totals = {
        total: roundTo(group.total),
        totalRem: roundTo(group.totalRem),
        totalFac: roundTo(group.totalFac)
      };
      totals.totalConsumo = roundTo(totals.totalRem + totals.totalFac);
      totals.restante = roundTo(Math.max(totals.total - totals.totalConsumo, 0));
      const percentages = computePercentages(totals);
      percentages.consumo = clampPercentage(percentages.rem + percentages.fac);
      const ids = Array.from(group.ids.values());
      const extensionIds = ids.filter(id => id !== group.baseId);
      return {
        baseId: group.baseId,
        ids,
        extensionIds,
        totals,
        percentages,
        items: normalizedItems,
        alerts: group.alerts
      };
    })
    .sort((a, b) => a.baseId.localeCompare(b.baseId));
}

function drawGroupTable(doc, group, options = {}) {
  const startX = options.startX ?? doc.page.margins.left;
  const width = options.width ?? (doc.page.width - doc.page.margins.left - doc.page.margins.right);
  const columns = [
    { key: 'id', label: 'PO / Ext', width: width * 0.15 },
    { key: 'fecha', label: 'Fecha', width: width * 0.11 },
    { key: 'clave', label: 'Clave', width: width * 0.11 },
    { key: 'autorizado', label: 'Autorizado', width: width * 0.14 },
    { key: 'rem', label: 'Remisiones', width: width * 0.14 },
    { key: 'fac', label: 'Facturas', width: width * 0.14 },
    { key: 'consumo', label: 'Consumido', width: width * 0.10 },
    { key: 'rest', label: 'Disponible', width: width * 0.11 }
  ];
  const headerHeight = 24;
  const rowPadding = 12;
  const minimumRowHeight = rowPadding + 12;
  const measureRowHeight = values => {
    const measuringFont = doc.font('Helvetica').fontSize(10);
    return values.reduce((height, value, index) => {
      const column = columns[index];
      const cellWidth = Math.max(0, column.width - 16);
      const text = String(value ?? '');
      const cellHeight = measuringFont.heightOfString(text, { width: cellWidth });
      return Math.max(height, cellHeight + rowPadding);
    }, rowPadding);
  };
  const drawHeader = nextRowHeight => {
    const requiredRowHeight = Math.max(minimumRowHeight, Number(nextRowHeight) || 0);
    ensureSpace(doc, headerHeight + requiredRowHeight + 6);
    const y = doc.y;
    doc.save();
    doc.lineWidth(1);
    doc.rect(startX, y, width, headerHeight).fillAndStroke('#0f172a', '#1f2937');
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(10);
    let offsetX = startX;
    columns.forEach(column => {
      doc.text(column.label, offsetX + 8, y + 6, { width: column.width - 16 });
      offsetX += column.width;
    });
    doc.restore();
    doc.y = y + headerHeight;
  };
  let headerDrawn = false;
  group.items.forEach(item => {
    const percentages = item.percentages || {};
    const totals = item.totals || {};
    const consumido = Number.isFinite(totals.totalConsumo) ? totals.totalConsumo : 0;
    const restanteBase = Number.isFinite(totals.restante) ? totals.restante : totals.total - consumido;
    const restante = Math.max(Number(restanteBase) || 0, 0);
    const fullyConsumed = Number(restante) <= 0.01;
    const consumoLabel = `${formatCurrency(consumido)} (${formatPercentage(percentages.consumo)})`;
    const disponibleLabel = `${formatCurrency(restante)} (${formatPercentage(percentages.rest)})`;
    const rowValues = [
      item.isBase ? `${item.id} (base)` : item.id,
      item.fecha || '-',
      item.clave || '-',
      formatCurrency(totals.total),
      `${formatCurrency(totals.totalRem)} (${formatPercentage(percentages.rem)})`,
      `${formatCurrency(totals.totalFac)} (${formatPercentage(percentages.fac)})`,
      consumoLabel,
      disponibleLabel
    ];
    const rowHeight = measureRowHeight(rowValues);
    if (!headerDrawn) {
      drawHeader(rowHeight);
      headerDrawn = true;
    }
    if (ensureSpace(doc, rowHeight + 6)) {
      drawHeader(rowHeight);
      headerDrawn = true;
    }
    const y = doc.y;
    const background = fullyConsumed ? '#fef3c7' : item.isBase ? '#f8fafc' : '#ffffff';
    const textColor = fullyConsumed ? '#92400e' : '#111827';
    doc.save().fillColor(background).rect(startX, y, width, rowHeight).fill().restore();
    doc.lineWidth(0.5).strokeColor('#d1d5db').rect(startX, y, width, rowHeight).stroke();
    let offsetX = startX;
    rowValues.forEach((value, index) => {
      const align = index >= 3 ? 'right' : 'left';
      const cellWidth = columns[index].width - 16;
      doc
        .font('Helvetica')
        .fontSize(10)
        .fillColor(textColor)
        .text(value, offsetX + 8, y + 6, { width: cellWidth, align });
      offsetX += columns[index].width;
    });
    doc.y = y + rowHeight;
  });
  doc.moveDown(0.4);
}

function drawGroupSection(doc, group, branding, options = {}) {
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const title = group.extensionIds.length
    ? `PO base ${group.baseId} (incluye ${group.extensionIds.length} extensión${group.extensionIds.length === 1 ? '' : 'es'})`
    : `PO ${group.baseId}`;
  ensureSpace(doc, 60);
  doc
    .font('Helvetica-Bold')
    .fontSize(13)
    .fillColor(branding.accentColor || '#111827')
    .text(title, startX, doc.y, { width });
  if (group.extensionIds.length) {
    doc
      .font('Helvetica')
      .fontSize(10)
      .fillColor('#475569')
      .text(`Detalle: ${[group.baseId, ...group.extensionIds].join(', ')}`, startX, doc.y, { width });
  }
  doc.moveDown(0.2);
  drawSmallConsumptionBar(doc, group.totals, branding, { startX, width });
  drawGroupTable(doc, group, { startX, width });
  if (!options.isLast) {
    const addedPage = ensureSpace(doc, 12);
    if (!addedPage) {
      const separatorY = doc.y;
      doc
        .lineWidth(2.5)
        .strokeColor('#000000')
        .moveTo(startX, separatorY)
        .lineTo(startX + width, separatorY)
        .stroke();
      doc.moveDown(0.6);
    } else {
      doc.moveDown(0.3);
    }
  }
}

function renderObservationBox(doc, notes, options = {}) {
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const sanitized = Array.isArray(notes)
    ? notes
        .map(note => (typeof note === 'string' ? note.trim() : ''))
        .filter(Boolean)
    : [];
  const hasNotes = sanitized.length > 0;
  const alwaysShow = options.alwaysShow === true;
  if (!hasNotes && !alwaysShow) {
    return;
  }
  const placeholderText = options.placeholderText || 'Sin observaciones registradas.';
  const notesToRender = hasNotes ? sanitized : [placeholderText];
  const bullets = notesToRender.map(note => {
    const trimmed = note.trim();
    return trimmed.startsWith('•') ? trimmed : `• ${trimmed}`;
  });
  doc.font('Helvetica').fontSize(12);
  const textWidth = width - 24;
  const contentHeight = bullets.reduce((sum, note) => sum + doc.heightOfString(note, { width: textWidth }), 0);
  const boxHeight = Math.max(86, contentHeight + 24);
  ensureSpace(doc, boxHeight + 90);
  doc.lineWidth(0.5).moveTo(startX, doc.y).lineTo(startX + width, doc.y).stroke('#e5e7eb');
  doc.moveDown(0.5);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor('#111827')
    .text(options.title || 'Observaciones', startX, doc.y);
  doc.moveDown(0.3);
  const boxTop = doc.y;
  doc.roundedRect(startX, boxTop, width, boxHeight, 8).stroke('#d1d5db');
  let currentY = boxTop + 12;
  bullets.forEach(note => {
    const height = doc.heightOfString(note, { width: textWidth });
    doc.font('Helvetica').fontSize(12).fillColor('#4b5563').text(note, startX + 12, currentY, { width: textWidth });
    currentY += height + 6;
  });
  doc.y = boxTop + boxHeight + 18;
}

function drawChartsSection(doc, summary, branding) {
  const groups = buildPoGroupDetails(summary);
  if (!groups.length) {
    return;
  }
  const totals = summary.totals || {};
  const title = groups.length > 1
    ? 'Consumo combinado de la selección'
    : `Consumo de la PO ${groups[0].baseId}`;
  drawCombinedConsumptionBar(doc, totals, branding, { title });
  doc.moveDown(0.6);
}

function drawSelectionGroupDetails(doc, summary, branding) {
  const groups = buildPoGroupDetails(summary);
  if (!groups.length) {
    return;
  }
  const startX = doc.page.margins.left;
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text('Detalle por pedido', startX, doc.y);
  doc.moveDown(0.3);
  groups.forEach((group, index) => {
    drawGroupSection(doc, group, branding, { isLast: index === groups.length - 1 });
  });
}

function drawObservations(doc, summary) {
  const baseCount = (summary.selectedIds || []).length;
  const itemCount = (summary.items || []).length;
  const selectionDetails = Array.isArray(summary.selectionDetails) ? summary.selectionDetails : [];
  const notes = [];
  if (baseCount === 1) {
    const detail = selectionDetails.find(entry => entry.baseId === summary.selectedIds[0]);
    if (!detail && itemCount <= 1) {
      notes.push('La PO seleccionada no presenta extensiones activas en este reporte.');
    }
  }
  if ((summary.totals?.restante ?? 0) <= 0) {
    notes.push('El presupuesto autorizado se encuentra completamente consumido.');
  }
  renderObservationBox(doc, Array.from(new Set(notes)), {
    alwaysShow: true,
    placeholderText: 'Sin observaciones registradas.'
  });
}

function drawFooter(doc, branding) {
  const footerText = branding.footerText || DEFAULT_BRANDING.footerText;
  doc.font('Helvetica').fontSize(10).fillColor('#9ca3af');
  doc.text(footerText, doc.page.margins.left, doc.page.height - doc.page.margins.bottom - 14, {
    width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
    align: 'center'
  });
}

async function generate(summary, branding = {}, customization = {}) {
  // Aseguramos que la dependencia PDFKit esté disponible antes de continuar.
  ensurePdfkitAvailable();
  // Normalizamos la configuración de marca y personalización para homogenizar el uso posterior.
  const style = normalizeBranding(branding);
  const options = normalizeCustomization(customization);

  // Generamos el PDF de forma asíncrona, acumulando los datos emitidos por PDFKit.
  return await new Promise((resolve, reject) => {
    try {
      // Instanciamos el documento con tamaño carta y márgenes predefinidos.
      const doc = new PDFDocument({ size: 'LETTER', margin: 40 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      // Colocamos la papelería y la cabecera con los datos generales.
      drawLetterhead(doc, style);
      drawHeader(doc, summary, style);

      // Rama especializada cuando el reporte corresponde a un universo.
      const isUniverseReport = summary.universe?.isUniverse;
      if (isUniverseReport) {
        // 1. Header (ya dibujado arriba)
        // 2. Filtros aplicados al universo
        drawUniverseFilterInfo(doc, summary, style);

        // 3. Autorizado y consumido, disponible (summary box)
        if (options.includeSummary) {
          drawSummaryBox(doc, summary, style);
        }

        // 5. Gráfica global con stats (barra de consumo combinado)
        const shouldDrawResumenContent = options.includeSummary || options.includeDetail || options.includeUniverse;
        if (shouldDrawResumenContent) {
          drawCombinedConsumptionBar(doc, summary.totals || {}, style, { title: 'Consumo total del universo' });
        }

        // Tabla de totales generales (autorizado, remisiones, etc.)
        drawTotalsTable(doc, summary, style);

        if (Array.isArray(summary.items) && summary.items.length > 0) {
          drawPoSummaryTable(doc, summary, style);
        }

        if (options.includeObservations) {
          drawUniverseObservations(doc, summary);
        }

        if (options.includeDetail) {
          drawUniverseGroupDetails(doc, summary, style);
        }
      } else {
        // Rama para reportes de selección de pedidos individuales (no universo).
        // 1. Header (ya dibujado arriba)
        // 2. Filtros (no aplica para selección individual)

        // 3. Autorizado y consumido, disponible (summary box)
        if (options.includeSummary) {
          drawSummaryBox(doc, summary, style);
        }

        // 4. Gráfica global con stats (barra de consumo combinado)
        const shouldDrawResumenContent = options.includeSummary || options.includeDetail || options.includeCharts || options.includeUniverse;
        if (shouldDrawResumenContent) {
          drawCombinedConsumptionBar(doc, summary.totals || {}, style, { title: 'Consumo total de la selección' });
        }

        // Tabla de totales generales (autorizado, remisiones, etc.)
        drawTotalsTable(doc, summary, style);

        if (Array.isArray(summary.items) && summary.items.length > 0) {
          drawPoSummaryTable(doc, summary, style);
        }

        if (options.includeObservations) {
          drawObservations(doc, summary);
        }

        if (options.includeDetail) {
          drawSelectionGroupDetails(doc, summary, style);
        }

        if (options.includeMovements) {
          drawMovements(doc, summary, style);
        }
      }

      // Cerramos con el pie de página institucional y finalizamos el flujo del documento.
      drawFooter(doc, style);
      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}

module.exports = {
  generate,
  isAvailable: () => Boolean(PDFDocument),
  getUnavailableMessage: () => {
    if (pdfkitLoadError) {
      return `${PDF_UNAVAILABLE_MESSAGE} Detalle técnico: ${pdfkitLoadError.message}`;
    }
    return PDF_UNAVAILABLE_MESSAGE;
  }
};
