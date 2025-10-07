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
  console.error('Si estás en Windows y no necesitas JasperReports puedes usar "npm install --omit=optional".\n');
}

const PDF_UNAVAILABLE_MESSAGE =
  'El motor "PDF directo" no está disponible porque la dependencia opcional "pdfkit" no se pudo cargar. ' +
  'Ejecuta "npm install" (o "npm install --omit=optional" si deseas omitir JasperReports) y vuelve a iniciar la aplicación.';

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
  companyName: 'SITTEL'
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

function computePercentages(totals = {}) {
  const total = Number(totals.total || 0);
  if (total <= 0) {
    return { rem: 0, fac: 0, rest: 0 };
  }
  const totalRem = Number(totals.totalRem || 0);
  const totalFac = Number(totals.totalFac || 0);
  const restanteAmount = Number(
    totals.restante != null ? totals.restante : total - (totalRem + totalFac)
  );
  const rem = clampPercentage((totalRem / total) * 100);
  const fac = clampPercentage((totalFac / total) * 100);
  let rest = clampPercentage((restanteAmount / total) * 100);
  if (roundTo(rem + fac + rest) !== 100) {
    rest = clampPercentage(100 - (rem + fac));
  }
  return { rem, fac, rest };
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
  const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  if (fileExists(branding.letterheadTop)) {
    try {
      doc.image(branding.letterheadTop, doc.page.margins.left, doc.page.margins.top - 20, {
        width: pageWidth
      });
      doc.moveDown(1);
    } catch (err) {
      console.warn('No se pudo dibujar membrete superior:', err.message);
    }
    doc.on('pageAdded', () => {
      try {
        const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
        doc.image(branding.letterheadTop, doc.page.margins.left, doc.page.margins.top - 20, { width });
        doc.moveDown(1);
      } catch (err) {
        console.warn('No se pudo dibujar membrete superior en página adicional:', err.message);
      }
    });
  }
  if (fileExists(branding.letterheadBottom)) {
    doc.on('pageAdded', () => {
      try {
        const y = doc.page.height - doc.page.margins.bottom - 60;
        doc.image(branding.letterheadBottom, doc.page.margins.left, y, { width: pageWidth });
      } catch (err) {
        console.warn('No se pudo dibujar membrete inferior:', err.message);
      }
    });
    try {
      const y = doc.page.height - doc.page.margins.bottom - 60;
      doc.image(branding.letterheadBottom, doc.page.margins.left, y, { width: pageWidth });
    } catch (err) {
      console.warn('No se pudo dibujar membrete inferior:', err.message);
    }
  }
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
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  ensureSpace(doc, 140);
  doc.lineWidth(0.5);
  doc.moveTo(startX, doc.y).lineTo(startX + width, doc.y).stroke('#e5e7eb');
  doc.moveDown(0.6);
  doc.font('Helvetica-Bold').fontSize(14).fillColor('#111827').text('Observaciones', startX, doc.y);
  doc.moveDown(0.4);
  const boxTop = doc.y;
  const boxHeight = 100;
  doc.lineWidth(1).rect(startX, boxTop, width, boxHeight).stroke('#d1d5db');
  doc.font('Helvetica').fontSize(12).fillColor('#555555');
  const universe = summary.universe || {};
  const label = universe.label || 'Global';
  const notes = [
    `*Resumen agregado del universo de POs (${label}).`,
    '*El consumo considera remisiones y facturas registradas en el periodo seleccionado.'
  ];
  if (summary.alertasTexto && summary.alertasTexto !== 'Sin alertas generales') {
    notes.push(`*Alertas: ${summary.alertasTexto}`);
  }
  if (summary.totals && summary.totals.total === 0) {
    notes.push('*No se encontraron POs activas con el filtro aplicado.');
  }
  const textStartY = boxTop + 12;
  notes.forEach((note, index) => {
    doc.text(note, startX + 10, textStartY + index * 20, { width: width - 20 });
  });
  doc.y = boxTop + boxHeight + 16;
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
  const startX = doc.page.margins.left;
  const columnWidth = (doc.page.width - doc.page.margins.left - doc.page.margins.right - 20) / 2;
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
      .text('Facturas', startX + columnWidth + 20, top);

    doc.lineWidth(1);
    doc.rect(startX, top + 18, columnWidth, boxHeight).stroke(branding.remColor);
    doc.rect(startX + columnWidth + 20, top + 18, columnWidth, boxHeight).stroke(branding.facColor);

    doc.font('Helvetica').fontSize(11).fillColor('#111827');
    remSlice.forEach((item, index) => {
      const y = top + 30 + index * 16;
      doc.text(`${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`, startX + 10, y, {
        width: columnWidth - 20
      });
    });
    if (remSlice.length === 0) {
      doc
        .font('Helvetica')
        .fontSize(11)
        .fillColor('#6b7280')
        .text('Sin remisiones destacadas en este bloque.', startX + 10, top + 32, {
          width: columnWidth - 20
        });
    }

    doc.font('Helvetica').fillColor('#111827');
    facSlice.forEach((item, index) => {
      const y = top + 30 + index * 16;
      doc.text(
        `${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`,
        startX + columnWidth + 30,
        y,
        {
          width: columnWidth - 20
        }
      );
    });
    if (facSlice.length === 0) {
      doc
        .font('Helvetica')
        .fontSize(11)
        .fillColor('#6b7280')
        .text('Sin facturas destacadas en este bloque.', startX + columnWidth + 30, top + 32, {
          width: columnWidth - 20
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
          { width: columnWidth - 20 }
        );
      doc
        .font('Helvetica-Bold')
        .fillColor('#111827')
        .text(
          `Subtotal Facturas: ${formatCurrency(fac.subtotal)} (${formatPercentage(fac.porcentaje)})`,
          startX + columnWidth + 30,
          footerY,
          { width: columnWidth - 20 }
        );
    } else {
      doc
        .font('Helvetica')
        .fontSize(10)
        .fillColor('#6b7280')
        .text('Continúa en la siguiente página…', startX + 10, footerY, { width: columnWidth - 20 });
      doc
        .font('Helvetica')
        .fontSize(10)
        .fillColor('#6b7280')
        .text('Continúa en la siguiente página…', startX + columnWidth + 30, footerY, {
          width: columnWidth - 20
        });
    }
    doc.y = top + 18 + boxHeight + 12;
  }
}

function drawSummaryBox(doc, summary) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  ensureSpace(doc, 80);
  const top = doc.y;
  const percentages = computePercentages(totals);
  const consumoPercentage = roundTo(percentages.rem + percentages.fac);
  doc.save();
  doc.lineWidth(1);
  doc.rect(startX, top, width, 62).fillAndStroke('#f7f7f7', '#d1d5db');
  doc.fillColor('#111827').font('Helvetica').fontSize(12);
  doc.text(
    `Consumido (Rem+Fac): ${formatCurrency(totals.totalConsumo)} (${formatPercentage(consumoPercentage)})`,
    startX + 10,
    top + 12,
    { width: width / 2 }
  );
  doc.text(
    `Restante: ${formatCurrency(totals.restante)} (${formatPercentage(percentages.rest)})`,
    startX + width / 2,
    top + 12,
    { width: width / 2 - 10 }
  );
  doc.fillColor('#6b7280').fontSize(11);
  doc.text(
    `Nota: Porcentajes recalculados sobre el total autorizado (${formatCurrency(totals.total)}).`,
    startX + 10,
    top + 32,
    { width: width - 20 }
  );
  doc.restore();
  doc.y = top + 74;
}

function drawStackedBar(doc, summary, branding) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left + 20;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right - 40;
  ensureSpace(doc, 180);
  const top = doc.y;
  const height = 34;
  const percentages = computePercentages(totals);
  const remWidth = (width * percentages.rem) / 100;
  const facWidth = (width * percentages.fac) / 100;
  let restWidth = width - remWidth - facWidth;
  if (restWidth < 0) {
    restWidth = 0;
  }

  doc
    .font('Helvetica-Bold')
    .fontSize(16)
    .fillColor('#111827')
    .text('Consumo del PEO (barra apilada 100%)', doc.page.margins.left, top, {
      width: width + 40,
      align: 'center'
    });
  doc.moveDown(0.4);

  const barTop = doc.y;
  doc.lineWidth(1).rect(startX, barTop, width, height).stroke('#d1d5db');
  if (remWidth > 0) {
    doc.save().rect(startX, barTop, remWidth, height).fillOpacity(0.9).fill(branding.remColor).restore();
  }
  if (facWidth > 0) {
    doc
      .save()
      .rect(startX + remWidth, barTop, facWidth, height)
      .fillOpacity(0.9)
      .fill(branding.facColor)
      .restore();
  }
  if (restWidth > 0) {
    doc
      .save()
      .rect(startX + remWidth + facWidth, barTop, restWidth, height)
      .fillOpacity(0.9)
      .fill(branding.restanteColor)
      .restore();
  }
  doc.lineWidth(2).strokeColor('#ffffff');
  doc.moveTo(startX + remWidth, barTop).lineTo(startX + remWidth, barTop + height).stroke();
  doc.moveTo(startX + remWidth + facWidth, barTop).lineTo(startX + remWidth + facWidth, barTop + height).stroke();

  doc.font('Helvetica').fontSize(11).fillColor('#111827');
  const remLabelX = remWidth > 0 ? startX + remWidth / 2 - 40 : startX;
  const facLabelX = facWidth > 0 ? startX + remWidth + facWidth / 2 - 40 : startX + remWidth - 40;
  const restLabelX = restWidth > 0
    ? startX + remWidth + facWidth + restWidth / 2 - 60
    : startX + remWidth + facWidth - 60;
  doc.text(`Rem: ${formatPercentage(percentages.rem)}`, remLabelX, barTop - 14, { width: 80, align: 'center' });
  doc.text(`Fac: ${formatPercentage(percentages.fac)}`, facLabelX, barTop - 14, { width: 80, align: 'center' });
  doc.text(`Restante: ${formatPercentage(percentages.rest)}`, restLabelX, barTop - 14, {
    width: 120,
    align: 'center'
  });
  doc.y = barTop + height + 30;

  const legendEntries = buildLegendEntries(totals, percentages, branding);
  drawLegend(doc, legendEntries, startX, doc.y);
  doc.moveDown(1.2);
}

function drawPieSlice(doc, centerX, centerY, radius, startAngle, endAngle, color) {
  doc.save();
  doc.moveTo(centerX, centerY);
  doc.lineTo(centerX + radius * Math.cos((Math.PI / 180) * startAngle), centerY + radius * Math.sin((Math.PI / 180) * startAngle));
  doc.arc(centerX, centerY, radius, startAngle, endAngle);
  doc.lineTo(centerX, centerY);
  doc.fillOpacity(0.9).fill(color);
  doc.restore();
}

function drawPieChart(doc, summary, branding) {
  const totals = summary.totals || {};
  ensureSpace(doc, 260);
  const centerX = doc.page.width / 2;
  const centerY = doc.y + 110;
  const radius = 95;
  const percentages = computePercentages(totals);
  doc
    .font('Helvetica-Bold')
    .fontSize(16)
    .fillColor('#111827')
    .text('Consumo Total (%) - PEO combinado', doc.page.margins.left, doc.y, {
      width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
      align: 'center'
    });
  doc.y += 18;

  const angles = [
    { value: percentages.rem, color: branding.remColor, label: 'Remisiones' },
    { value: percentages.fac, color: branding.facColor, label: 'Facturas' },
    { value: percentages.rest, color: branding.restanteColor, label: 'Restante' }
  ];
  let currentAngle = -90;
  angles.forEach(segment => {
    const sweep = (segment.value / 100) * 360;
    if (sweep <= 0) return;
    drawPieSlice(doc, centerX, centerY, radius, currentAngle, currentAngle + sweep, segment.color);
    currentAngle += sweep;
  });

  doc.font('Helvetica').fontSize(11).fillColor('#111827');
  const legendEntries = buildLegendEntries(totals, percentages, branding);
  const legendX = Math.min(centerX + radius + 24, doc.page.width - doc.page.margins.right - 240);
  drawLegend(doc, legendEntries, legendX, centerY - radius);
  doc.y = Math.max(doc.y, centerY + radius + 32);
}

function buildLegendEntries(totals, percentages, branding) {
  return [
    { label: `Remisiones: ${formatPercentage(percentages.rem)} (${formatCurrency(totals.totalRem)})`, color: branding.remColor },
    { label: `Facturas: ${formatPercentage(percentages.fac)} (${formatCurrency(totals.totalFac)})`, color: branding.facColor },
    { label: `Restante: ${formatPercentage(percentages.rest)} (${formatCurrency(totals.restante)})`, color: branding.restanteColor }
  ];
}

function drawLegend(doc, entries, startX, startY) {
  const availableWidth = Math.max(doc.page.width - doc.page.margins.right - startX, 160);
  const legendWidth = Math.min(availableWidth, 260);
  entries.forEach((entry, index) => {
    const y = startY + index * 22;
    doc.save().rect(startX, y, 14, 14).fill(entry.color).restore();
    doc.fillColor('#111827').font('Helvetica').fontSize(11).text(entry.label, startX + 22, y + 2, {
      width: legendWidth - 22
    });
  });
}

function drawObservations(doc, summary) {
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  ensureSpace(doc, 130);
  doc.lineWidth(0.5);
  doc.moveTo(startX, doc.y).lineTo(startX + width, doc.y).stroke('#e5e7eb');
  doc.moveDown(0.6);
  doc.font('Helvetica-Bold').fontSize(14).fillColor('#111827').text('Observaciones', startX, doc.y);
  doc.moveDown(0.4);
  const boxTop = doc.y;
  doc.lineWidth(1).rect(startX, boxTop, width, 92).stroke('#d1d5db');
  doc.font('Helvetica').fontSize(12).fillColor('#555555');

  const notes = [];
  const baseCount = (summary.selectedIds || []).length;
  const itemCount = (summary.items || []).length;
  if (baseCount === 1 && itemCount === 1) {
    notes.push('*El PEO no presenta extensiones activas en este periodo.');
  }
  if (baseCount === 1 && itemCount > 1) {
    notes.push('*El PEO incluye extensiones; los totales consideran todas las variantes listadas.');
  }
  if (baseCount > 1) {
    notes.push(`*Se combinan ${baseCount} PEOs base; cada fila detalla su consumo específico.`);
  }
  if (summary.alertasTexto && summary.alertasTexto !== 'Sin alertas generales') {
    notes.push(`*Alertas relevantes: ${summary.alertasTexto}`);
  }
  notes.push('*Las remisiones y facturas mostradas alimentan los totales de consumo.');
  const textStartY = boxTop + 12;
  Array.from(new Set(notes)).forEach((note, index) => {
    doc.text(note, startX + 10, textStartY + index * 20, { width: width - 20 });
  });
  const boxBottom = boxTop + 92;
  doc.y = boxBottom + 16;
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
      const doc = new PDFDocument({ size: 'A4', margin: 40 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      drawLetterhead(doc, style);
      drawHeader(doc, summary, style);
      if (summary.universe?.isUniverse) {
        drawUniverseFilterInfo(doc, summary, style);
        drawSummaryBox(doc, summary);
        if (options.includeCharts) {
          drawStackedBar(doc, summary, style);
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
        drawSummaryBox(doc, summary);
        if (options.includeCharts) {
          if (summary.selectedIds && summary.selectedIds.length > 1) {
            drawPieChart(doc, summary, style);
          } else if ((summary.items || []).length > 1) {
            drawPieChart(doc, summary, style);
          } else {
            drawStackedBar(doc, summary, style);
          }
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
