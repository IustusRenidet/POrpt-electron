const fs = require('fs');

let PDFDocument;
let pdfkitLoadError = null;

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

const PDF_UNAVAILABLE_MESSAGE =
  'El motor "PDF directo" no está disponible porque la dependencia opcional "pdfkit" no se pudo cargar. ' +
  'Ejecuta "npm install" (o "npm install --omit=optional" si necesitas omitir dependencias opcionales) y vuelve a iniciar la aplicación.';

const DEFAULT_BRANDING = {
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
  companyName: 'SSITEL'
};

const DEFAULT_CUSTOMIZATION = {
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

function clampPercentage(value) {
  const normalized = roundTo(value);
  const clamped = Math.min(100, Math.max(0, normalized));
  return roundTo(clamped);
}

function formatPercentage(value) {
  return `${roundTo(value, 2).toFixed(2)}%`;
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
    if (!registry.has(trimmed)) {
      registry.add(trimmed);
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
  const base = Math.max(total, consumo);
  if (base <= 0) {
    return { rem: 0, fac: 0, rest: 0, base: 0, overage: 0 };
  }
  const rem = clampPercentage((totalRem / base) * 100);
  const fac = clampPercentage((totalFac / base) * 100);
  let rest = clampPercentage((restanteAmount / base) * 100);
  if (roundTo(rem + fac + rest) !== 100) {
    rest = clampPercentage(100 - (rem + fac));
  }
  const overage = Math.max(0, consumo - total);
  return {
    rem,
    fac,
    rest,
    base,
    overage
  };
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
  const selectedIds = Array.isArray(summary.selectedIds) ? summary.selectedIds.filter(Boolean) : [];
  let seleccion = summary.selectedId || summary.baseId || '-';
  if (summary.universe?.isUniverse) {
    const universeLabel = summary.universe.label || 'Global (todas las POs)';
    seleccion = `Alcance: ${universeLabel}`;
  } else if (selectedIds.length > 1) {
    seleccion = `${selectedIds.length} bases (${selectedIds.join(', ')})`;
  } else if (selectedIds.length === 1) {
    seleccion = selectedIds[0];
  }
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(`Empresa: ${empresaLabel}`, { align: 'center' });
  doc
    .font('Helvetica')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(summary.universe?.isUniverse ? seleccion : `PO(s) seleccionada(s): ${seleccion}`, { align: 'center' });
  if (summary.universe?.isUniverse) {
    const titleText = summary.universe.title || 'Reporte del universo de POs';
    doc.moveDown(0.2);
    doc
      .font('Helvetica')
      .fontSize(12)
      .fillColor('#475569')
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

function drawUniverseTotalsTable(doc, summary, branding) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const rowHeight = 32;
  const rows = [
    { label: 'Total autorizado de POs', value: formatCurrency(totals.total), color: branding.accentColor },
    { label: 'Total remisiones', value: formatCurrency(totals.totalRem), color: branding.remColor },
    { label: 'Total facturas', value: formatCurrency(totals.totalFac), color: branding.facColor },
    { label: 'Total consumido (Rem + Fac)', value: formatCurrency(totals.totalConsumo), color: '#1f2937' },
    { label: 'Remanente total del universo', value: formatCurrency(totals.restante), color: branding.restanteColor }
  ];

  const totalHeight = rowHeight * rows.length;
  ensureSpace(doc, totalHeight + 32);
  const adjustedTop = doc.y;
  doc.save();
  doc.lineWidth(1).rect(startX, adjustedTop, width, totalHeight).stroke('#d1d5db');
  rows.forEach((row, index) => {
    const y = adjustedTop + index * rowHeight;
    if (index > 0) {
      doc.lineWidth(0.5).moveTo(startX, y).lineTo(startX + width, y).stroke('#e5e7eb');
    }
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor(row.color || branding.accentColor || '#1f2937')
      .text(row.label, startX + 12, y + 10, { width: width / 2 - 16 });
    doc
      .font('Helvetica')
      .fontSize(12)
      .fillColor('#111827')
      .text(row.value, startX + width / 2, y + 10, { width: width / 2 - 16, align: 'right' });
  });
  doc.restore();
  doc.y = adjustedTop + totalHeight + 18;
}

function drawUniverseObservations(doc, summary) {
  const universe = summary.universe || {};
  const label = universe.label || 'Global';
  const company = summary.companyName || summary.empresaLabel || summary.empresa || 'SSITEL';
  const notes = [
    `Universo de ${company} con el filtro "${label}".`,
    'Flujo del reporte: se inicia con el importe autorizado, continúa con el consumo acumulado (remisiones + facturas) y finaliza con el disponible.',
    'El consumo refleja toda la operación registrada en el periodo para dimensionar el avance del presupuesto.',
    'El disponible señala si es necesario liberar nuevas POs o ajustar el gasto antes de revisar cada pedido.'
  ];
  if (summary.alertasTexto && summary.alertasTexto !== 'Sin alertas generales') {
    notes.push(`Alertas: ${summary.alertasTexto}`);
  }
  if (summary.totals && summary.totals.total === 0) {
    notes.push('No se encontraron POs activas con el filtro aplicado.');
  }
  notes.push('Importes redondeados a dos decimales para facilitar la lectura.');
  renderObservationBox(doc, Array.from(new Set(notes)));
}

function drawPoBaseBreakdown(doc, summary, branding) {
  if (summary.universe?.isUniverse) return;
  const selectedIds = Array.isArray(summary.selectedIds) ? summary.selectedIds.filter(Boolean) : [];
  if (selectedIds.length <= 1) return;
  const items = summary.items || [];
  const groups = new Map();
  items.forEach(item => {
    const baseId = item.baseId || item.id;
    if (!groups.has(baseId)) {
      groups.set(baseId, {
        ids: [],
        total: 0,
        rem: 0,
        fac: 0
      });
    }
    const group = groups.get(baseId);
    group.ids.push(item.id);
    group.total += Number(item.total || 0);
    group.rem += Number(item.totals?.totalRem || 0);
    group.fac += Number(item.totals?.totalFac || 0);
  });

  const breakdown = selectedIds.map(baseId => {
    const group = groups.get(baseId) || { ids: [], total: 0, rem: 0, fac: 0 };
    const consumo = group.rem + group.fac;
    const restante = Math.max(group.total - consumo, 0);
    return {
      baseId,
      label: `${baseId} (${group.ids.length || 1} partidas)`,
      total: group.total,
      consumo,
      restante
    };
  });

  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const headerHeight = 26;
  const rowHeight = 24;
  const tableHeight = headerHeight + rowHeight * breakdown.length;
  ensureSpace(doc, tableHeight + 70);

  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text('Resumen por PEO base', startX, doc.y);
  doc.moveDown(0.4);

  const tableTop = doc.y;
  const columnPercents = [0.32, 0.34, 0.34];
  const columnXs = [startX, startX + width * columnPercents[0], startX + width * (columnPercents[0] + columnPercents[1])];

  doc.lineWidth(1).rect(startX, tableTop, width, tableHeight).stroke('#d1d5db');
  doc.save();
  doc.rect(startX, tableTop, width, headerHeight).fillAndStroke('#0f172a', '#1f2937');
  doc
    .fillColor('#ffffff')
    .font('Helvetica-Bold')
    .fontSize(11)
    .text('PEO base', columnXs[0] + 12, tableTop + 6, { width: width * columnPercents[0] - 24 });
  doc.text('Autorizado', columnXs[1] + 12, tableTop + 6, { width: width * columnPercents[1] - 24, align: 'right' });
  doc.text('Consumido / Restante', columnXs[2] + 12, tableTop + 6, {
    width: width * columnPercents[2] - 24,
    align: 'right'
  });
  doc.restore();

  breakdown.forEach((row, index) => {
    const y = tableTop + headerHeight + index * rowHeight;
    doc.lineWidth(0.5).moveTo(startX, y).lineTo(startX + width, y).stroke('#e5e7eb');
    doc
      .font('Helvetica')
      .fontSize(11)
      .fillColor('#111827')
      .text(row.label, columnXs[0] + 12, y + 6, { width: width * columnPercents[0] - 24 });
    doc.text(formatCurrency(row.total), columnXs[1] + 12, y + 6, {
      width: width * columnPercents[1] - 24,
      align: 'right'
    });
    doc.text(`${formatCurrency(row.consumo)} / ${formatCurrency(row.restante)}`, columnXs[2] + 12, y + 6, {
      width: width * columnPercents[2] - 24,
      align: 'right'
    });
  });

  doc.y = tableTop + tableHeight + 16;
}

function drawPoTable(doc, summary) {
  const items = summary.items || [];
  if (!items.length) return;
  const startX = doc.page.margins.left;
  const tableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const headerHeight = 26;
  const rowHeight = 22;
  const colWidths = [tableWidth * 0.38, tableWidth * 0.24, tableWidth * 0.38];
  const colXs = [startX, startX + colWidths[0], startX + colWidths[0] + colWidths[1]];

  const drawHeader = () => {
    ensureSpace(doc, headerHeight + 8);
    const y = doc.y;
    doc.save();
    doc.lineWidth(1);
    doc.rect(startX, y, tableWidth, headerHeight).fillAndStroke('#111827', '#1f2937');
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(11);
    doc.text('PEO ID', colXs[0] + 10, y + 6, { width: colWidths[0] - 20 });
    doc.text('Fecha', colXs[1] + 10, y + 6, { width: colWidths[1] - 20 });
    doc.text('Total', colXs[2] + 10, y + 6, { width: colWidths[2] - 20 });
    doc.restore();
    doc.y = y + headerHeight;
  };

  drawHeader();

  items.forEach(item => {
    if (ensureSpace(doc, rowHeight + 6)) {
      drawHeader();
    }
    const y = doc.y;
    doc.lineWidth(0.5).rect(startX, y, tableWidth, rowHeight).stroke('#d1d5db');
    doc.font('Helvetica').fontSize(11).fillColor('#1f2937');
    doc.text(item.id, colXs[0] + 10, y + 6, { width: colWidths[0] - 20 });
    doc.text(item.fecha || '-', colXs[1] + 10, y + 6, { width: colWidths[1] - 20 });
    const totalAmount = Number(item.total || 0);
    const subtotalAmount = Number(item.subtotal || 0);
    const showSubtotal = subtotalAmount > 0 && Math.abs(subtotalAmount - totalAmount) > 0.009;
    const totalText = showSubtotal
      ? `${formatCurrency(totalAmount)} (Sub: ${formatCurrency(subtotalAmount)})`
      : formatCurrency(totalAmount);
    doc.text(totalText, colXs[2] + 10, y + 6, {
      width: colWidths[2] - 20,
      align: 'left'
    });
    doc.y = y + rowHeight;
  });

  const totalsOnNewPage = ensureSpace(doc, 50);
  if (totalsOnNewPage) {
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor('#111827')
      .text('Resumen de totales', startX, doc.y);
    doc.moveDown(0.2);
  } else {
    doc.moveDown(0.3);
  }
  const baseCount = (summary.selectedIds || []).length;
  if (baseCount === 1) {
    const baseId = summary.selectedIds[0];
    const baseItem = items.find(item => item.id === baseId);
    const extensionItems = items.filter(item => item.id !== baseId);
    const originalTotal = baseItem ? Number(baseItem.total || 0) : 0;
    const extensionTotal = extensionItems.reduce((sum, item) => sum + Number(item.total || 0), 0);
    const total = originalTotal + extensionTotal;
    doc.lineWidth(0.5);
    doc.moveTo(startX, doc.y).lineTo(startX + tableWidth, doc.y).stroke('#d1d5db');
    doc.moveDown(0.4);
    if (extensionItems.length === 0) {
      doc
        .font('Helvetica-Bold')
        .fontSize(12)
        .fillColor('#111827')
        .text('Total PEO:', startX + 10, doc.y);
      doc
        .font('Helvetica')
        .fontSize(12)
        .text(formatCurrency(total), startX + 110, doc.y);
    } else {
      doc
        .font('Helvetica-Bold')
        .fontSize(12)
        .fillColor('#111827')
        .text('Totales', startX + 10, doc.y);
      doc
        .font('Helvetica')
        .fontSize(12)
        .text(`Original: ${formatCurrency(originalTotal)}`, startX + 110, doc.y);
      doc.text(`Extensión: ${formatCurrency(extensionTotal)}`, startX + 280, doc.y);
      doc
        .font('Helvetica-Bold')
        .text(`Total: ${formatCurrency(total)}`, startX + tableWidth - 180, doc.y, { width: 170, align: 'right' });
    }
  } else {
    doc.lineWidth(0.5);
    doc.moveTo(startX, doc.y).lineTo(startX + tableWidth, doc.y).stroke('#d1d5db');
    doc.moveDown(0.4);
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor('#111827')
      .text(`Totales (${baseCount} PEOs)`, startX + 10, doc.y);
    doc
      .font('Helvetica')
      .fontSize(12)
      .text(`Autorizado: ${formatCurrency(summary.totals.total)}`, startX + 200, doc.y);
    doc.text(`Consumido: ${formatCurrency(summary.totals.totalConsumo)}`, startX + 380, doc.y);
  }
  doc.moveDown(1);
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

function drawAlertList(doc, alerts, options = {}) {
  const entries = normalizeAlertEntries(alerts, options.fallback || []);
  if (!entries.length) {
    return;
  }
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const prepared = entries.map(entry => entry.replace(/\s+/gu, ' ').trim()).filter(Boolean);
  if (!prepared.length) {
    return;
  }
  const measuringFont = doc.font('Helvetica').fontSize(11);
  const contentHeight = prepared.reduce((sum, entry) => {
    const height = measuringFont.heightOfString(entry, { width });
    return sum + height + 4;
  }, 0);
  const requiredHeight = contentHeight + 28;
  ensureSpace(doc, requiredHeight);
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor('#b91c1c')
    .text(options.title || 'Alertas', startX, doc.y, { width });
  doc.moveDown(0.2);
  prepared.forEach(entry => {
    doc.font('Helvetica').fontSize(11).fillColor('#b91c1c').text(entry, startX, doc.y, { width });
    doc.moveDown(0.1);
  });
  doc.moveDown(0.4);
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
  let currentX = startX;
  segments.forEach(segment => {
    const segmentWidth = (width * segment.value) / 100;
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
  segments.forEach(segment => {
    const segmentWidth = (width * segment.value) / 100;
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
    group.items.push({
      id,
      fecha: item.fecha || '',
      clave: item.docSig || item.doc_sig || item.tipDoc || item.tip_doc || '',
      totals,
      percentages,
      isBase,
      alerts: itemAlerts
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
        items: orderedItems,
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
  const rowHeight = 22;
  const drawHeader = () => {
    ensureSpace(doc, headerHeight + 6);
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
  drawHeader();
  group.items.forEach(item => {
    if (ensureSpace(doc, rowHeight + 6)) {
      drawHeader();
    }
    const y = doc.y;
    const background = item.isBase ? '#f8fafc' : '#ffffff';
    doc.save().fillColor(background).rect(startX, y, width, rowHeight).fill().restore();
    doc.lineWidth(0.5).strokeColor('#d1d5db').rect(startX, y, width, rowHeight).stroke();
    doc.font('Helvetica').fontSize(10).fillColor('#111827');
    const rowValues = [
      item.isBase ? `${item.id} (base)` : item.id,
      item.fecha || '-',
      item.clave || '-',
      formatCurrency(item.totals.total),
      `${formatCurrency(item.totals.totalRem)} (${formatPercentage(item.percentages.rem)})`,
      `${formatCurrency(item.totals.totalFac)} (${formatPercentage(item.percentages.fac)})`,
      `${formatCurrency(item.totals.totalConsumo)} (${formatPercentage(item.percentages.consumo)})`,
      `${formatCurrency(item.totals.restante)} (${formatPercentage(item.percentages.rest)})`
    ];
    let offsetX = startX;
    rowValues.forEach((value, index) => {
      const align = index >= 3 ? 'right' : 'left';
      doc.text(value, offsetX + 8, y + 6, { width: columns[index].width - 16, align });
      offsetX += columns[index].width;
    });
    doc.y = y + rowHeight;
  });
  doc.moveDown(0.4);
}

function drawGroupSection(doc, group, branding) {
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
  const alertEntries = normalizeAlertEntries(group.alerts || []);
  if (alertEntries.length) {
    drawAlertList(doc, alertEntries, { title: 'Alertas del grupo' });
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
  if (!sanitized.length) {
    return;
  }
  const bullets = sanitized.map(note => (note.startsWith('•') ? note : `• ${note}`));
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
  doc.moveDown(0.3);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(branding.accentColor || '#111827')
    .text('Detalle por pedido', doc.page.margins.left, doc.y);
  doc.moveDown(0.3);
  groups.forEach(group => {
    drawGroupSection(doc, group, branding);
  });
}

function drawObservations(doc, summary) {
  const baseCount = (summary.selectedIds || []).length;
  const itemCount = (summary.items || []).length;
  const selectionDetails = Array.isArray(summary.selectionDetails) ? summary.selectionDetails : [];
  const notes = [];
  if (baseCount === 1) {
    const detail = selectionDetails.find(entry => entry.baseId === summary.selectedIds[0]);
    if (detail && (detail.count || detail.variants?.length || 0) > 1) {
      const totalVariants = detail.count || detail.variants.length;
      notes.push(`Se incluyeron ${totalVariants - 1} extensiones manuales para la base ${detail.baseId}.`);
    } else if (itemCount <= 1) {
      notes.push('La base seleccionada no presenta extensiones activas en este reporte.');
    }
  }
  if (baseCount > 1) {
    notes.push(`Se combinan ${baseCount} PEOs base; revisa las tarjetas por PEO para conocer su consumo individual.`);
  }
  const manualGroups = selectionDetails.filter(entry => entry.baseId && (entry.count || entry.variants?.length || 0) > 1);
  if (manualGroups.length > 0) {
    const detailText = manualGroups
      .map(entry => {
        const count = entry.count || entry.variants.length;
        return `${entry.baseId}: ${count - 1} ext.`;
      })
      .join(', ');
    notes.push(`Extensiones incluidas manualmente → ${detailText}.`);
  }
  if (summary.alertasTexto && summary.alertasTexto !== 'Sin alertas generales') {
    notes.push(`Alertas relevantes: ${summary.alertasTexto}`);
  }
  if ((summary.totals?.restante ?? 0) <= 0) {
    notes.push('El presupuesto autorizado se encuentra completamente consumido.');
  }
  notes.push('Importes redondeados a dos decimales para facilitar la lectura.');
  renderObservationBox(doc, Array.from(new Set(notes)));
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
  ensurePdfkitAvailable();
  const style = normalizeBranding(branding);
  const options = normalizeCustomization(customization);

  return await new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'LETTER', margin: 40 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      drawLetterhead(doc, style);
      drawHeader(doc, summary, style);
      if (summary.universe?.isUniverse) {
        drawUniverseFilterInfo(doc, summary, style);
      drawSummaryBox(doc, summary, style);
      const alertFallback = typeof summary.alertasTexto === 'string' ? summary.alertasTexto.split(/\n+/u) : [];
      const globalAlerts = normalizeAlertEntries(summary.alerts || [], alertFallback);
      if (globalAlerts.length) {
        drawAlertList(doc, globalAlerts, { title: 'Alertas del reporte' });
      }
      if (options.includeCharts) {
        drawCombinedConsumptionBar(doc, summary.totals || {}, style, { title: 'Consumo total del universo' });
      } else {
        doc.moveDown(1.2);
      }
        if (options.includeUniverse) {
          drawUniverseTotalsTable(doc, summary, style);
        }
        if (options.includeObservations) {
          drawUniverseObservations(doc, summary);
        }
      } else {
      drawSummaryBox(doc, summary, style);
      const alertFallback = typeof summary.alertasTexto === 'string' ? summary.alertasTexto.split(/\n+/u) : [];
      const globalAlerts = normalizeAlertEntries(summary.alerts || [], alertFallback);
      if (globalAlerts.length) {
        drawAlertList(doc, globalAlerts, { title: 'Alertas del reporte' });
      }
      if (options.includeCharts) {
        drawChartsSection(doc, summary, style);
      } else {
        doc.moveDown(1.2);
      }
        drawPoBaseBreakdown(doc, summary, style);
        drawPoTable(doc, summary);
        if (options.includeMovements) {
          drawMovements(doc, summary, style);
        }
        if (options.includeObservations) {
          drawObservations(doc, summary);
        }
      }
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
