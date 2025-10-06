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
  accentColor: '#1f2937'
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
    letterheadBottom: typeof branding.letterheadBottom === 'string' ? branding.letterheadBottom.trim() : ''
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

function formatPercentage(value) {
  return `${Number(value || 0).toFixed(1)}%`;
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
  const empresaLabel = summary.empresaLabel || summary.companyName || summary.empresa || '';
  const seleccion = Array.isArray(summary.selectedIds) && summary.selectedIds.length > 1
    ? summary.selectedIds.join(', ')
    : summary.selectedId || summary.baseId || '-';
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(`Empresa: ${empresaLabel}`, { align: 'center' });
  doc
    .font('Helvetica')
    .fontSize(12)
    .fillColor(branding.accentColor || '#1f2937')
    .text(`PO(s) seleccionada(s): ${seleccion}`, { align: 'center' });
  doc.moveDown(0.8);
}

function drawPoTable(doc, summary) {
  const items = summary.items || [];
  if (!items.length) return;
  const startX = doc.page.margins.left;
  const tableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const headerHeight = 24;
  const rowHeight = 22;
  const tableTop = doc.y;
  const totalRowsHeight = headerHeight + rowHeight * items.length;

  doc.save();
  doc.lineWidth(1);
  doc.rect(startX, tableTop, tableWidth, totalRowsHeight).stroke('#1f2937');
  doc.moveTo(startX, tableTop + headerHeight).lineTo(startX + tableWidth, tableTop + headerHeight).stroke('#1f2937');
  const colXs = [startX, startX + tableWidth * 0.35, startX + tableWidth * 0.65];
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor('#111827')
    .text('PEO ID', colXs[0] + 10, tableTop + 6, { width: tableWidth * 0.35 - 20 });
  doc.text('Fecha', colXs[1] + 10, tableTop + 6, { width: tableWidth * 0.3 - 20 });
  doc.text('Total', colXs[2] + 10, tableTop + 6, { width: tableWidth * 0.35 - 20 });

  items.forEach((item, index) => {
    const y = tableTop + headerHeight + index * rowHeight;
    doc.moveTo(startX, y).lineTo(startX + tableWidth, y).stroke('#e5e7eb');
    doc
      .font('Helvetica')
      .fontSize(11)
      .fillColor('#1f2937')
      .text(item.id, colXs[0] + 10, y + 6, { width: tableWidth * 0.35 - 20 });
    doc.text(item.fecha || '-', colXs[1] + 10, y + 6, { width: tableWidth * 0.3 - 20 });
    doc.text(formatCurrency(item.total || 0), colXs[2] + 10, y + 6, {
      width: tableWidth * 0.35 - 20,
      align: 'left'
    });
  });
  doc.restore();
  doc.y = tableTop + totalRowsHeight + 12;

  if (summary.selectedIds && summary.selectedIds.length === 1) {
    const baseId = summary.selectedIds[0];
    const baseItem = items.find(item => item.id === baseId);
    const extensionItems = items.filter(item => item.id !== baseId);
    const originalTotal = baseItem ? baseItem.total : 0;
    const extensionTotal = extensionItems.reduce((sum, item) => sum + item.total, 0);
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
      .text(`Totales (${(summary.selectedIds || []).length} PEOs)`, startX + 10, doc.y);
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

  const boxHeight = Math.max(rem.movements.length, fac.movements.length) * 16 + 36;
  doc.lineWidth(1);
  doc.rect(startX, top + 18, columnWidth, boxHeight).stroke(branding.remColor);
  doc.rect(startX + columnWidth + 20, top + 18, columnWidth, boxHeight).stroke(branding.facColor);

  doc.font('Helvetica').fontSize(11).fillColor('#111827');
  rem.movements.forEach((item, index) => {
    const y = top + 30 + index * 16;
    doc.text(`${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`, startX + 10, y, {
      width: columnWidth - 20
    });
  });
  doc
    .font('Helvetica-Bold')
    .fillColor('#111827')
    .text(
      `Subtotal Remisiones: ${formatCurrency(rem.subtotal)} (${formatPercentage(rem.porcentaje)})`,
      startX + 10,
      top + 18 + boxHeight - 18,
      { width: columnWidth - 20 }
    );

  doc.font('Helvetica').fillColor('#111827');
  fac.movements.forEach((item, index) => {
    const y = top + 30 + index * 16;
    doc.text(`${item.id}: ${formatCurrency(item.monto)} (${formatPercentage(item.porcentaje)})`, startX + columnWidth + 30, y, {
      width: columnWidth - 20
    });
  });
  doc
    .font('Helvetica-Bold')
    .fillColor('#111827')
    .text(
      `Subtotal Facturas: ${formatCurrency(fac.subtotal)} (${formatPercentage(fac.porcentaje)})`,
      startX + columnWidth + 30,
      top + 18 + boxHeight - 18,
      { width: columnWidth - 20 }
    );
  doc.y = top + 18 + boxHeight + 12;
}

function drawSummaryBox(doc, summary) {
  const totals = summary.totals || {};
  const startX = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const top = doc.y;
  doc.save();
  doc.lineWidth(1);
  doc.rect(startX, top, width, 62).fillAndStroke('#f7f7f7', '#d1d5db');
  doc.fillColor('#111827').font('Helvetica').fontSize(12);
  doc.text(
    `Consumido (Rem+Fac): ${formatCurrency(totals.totalConsumo)} (${formatPercentage((totals.porcRem || 0) + (totals.porcFac || 0))})`,
    startX + 10,
    top + 12,
    { width: width / 2 }
  );
  doc.text(
    `Restante: ${formatCurrency(totals.restante)} (${formatPercentage(totals.porcRest || 0)})`,
    startX + width / 2,
    top + 12,
    { width: width / 2 - 10 }
  );
  doc.fillColor('#6b7280').fontSize(11);
  doc.text(
    `Nota: Porcentajes calculados sobre el total autorizado (${formatCurrency(totals.total)}).`,
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
  const top = doc.y;
  const height = 36;
  const remWidth = (width * (totals.porcRem || 0)) / 100;
  const facWidth = (width * (totals.porcFac || 0)) / 100;
  const restWidth = width - remWidth - facWidth;

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
  doc.text(`Rem: ${formatPercentage(totals.porcRem || 0)}`, remLabelX, barTop - 14, { width: 80, align: 'center' });
  doc.text(`Fac: ${formatPercentage(totals.porcFac || 0)}`, facLabelX, barTop - 14, { width: 80, align: 'center' });
  doc.text(`Restante: ${formatPercentage(totals.porcRest || 0)}`, restLabelX, barTop - 14, {
    width: 120,
    align: 'center'
  });
  doc.y = barTop + height + 30;

  drawLegend(doc, branding, totals, startX, doc.y);
  doc.moveDown(1.5);
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
  const centerX = doc.page.width / 2;
  const centerY = doc.y + 120;
  const radius = 105;
  doc
    .font('Helvetica-Bold')
    .fontSize(16)
    .fillColor('#111827')
    .text('Consumo Total (%) - PEO combinado', doc.page.margins.left, doc.y, {
      width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
      align: 'center'
    });
  doc.y += 20;

  const angles = [
    { value: totals.porcRem || 0, color: branding.remColor, label: 'Remisiones' },
    { value: totals.porcFac || 0, color: branding.facColor, label: 'Facturas' },
    { value: totals.porcRest || 0, color: branding.restanteColor, label: 'Restante' }
  ];
  let currentAngle = -90;
  angles.forEach(segment => {
    const sweep = (segment.value / 100) * 360;
    if (sweep <= 0) return;
    drawPieSlice(doc, centerX, centerY, radius, currentAngle, currentAngle + sweep, segment.color);
    currentAngle += sweep;
  });

  doc.font('Helvetica').fontSize(11).fillColor('#111827');
  drawLegend(doc, branding, totals, centerX + radius + 25, centerY - 80);
  doc.y = centerY + radius + 40;
}

function drawLegend(doc, branding, totals, startX, startY) {
  const availableWidth = Math.max(doc.page.width - doc.page.margins.right - startX, 160);
  const legendWidth = Math.min(availableWidth, 260);
  const entries = [
    { label: `Remisiones: ${formatPercentage(totals.porcRem || 0)} (${formatCurrency(totals.totalRem)})`, color: branding.remColor },
    { label: `Facturas: ${formatPercentage(totals.porcFac || 0)} (${formatCurrency(totals.totalFac)})`, color: branding.facColor },
    { label: `Restante: ${formatPercentage(totals.porcRest || 0)} (${formatCurrency(totals.restante)})`, color: branding.restanteColor }
  ];
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
  doc.lineWidth(0.5);
  doc.moveTo(startX, doc.y).lineTo(startX + width, doc.y).stroke('#e5e7eb');
  doc.moveDown(0.6);
  doc.font('Helvetica-Bold').fontSize(14).fillColor('#111827').text('Observaciones', startX, doc.y);
  doc.moveDown(0.4);
  const boxTop = doc.y;
  doc.lineWidth(1).rect(startX, boxTop, width, 92).stroke('#d1d5db');
  doc.font('Helvetica').fontSize(12).fillColor('#555555');

  const notes = [];
  if (summary.selectedIds && summary.selectedIds.length === 1 && (summary.items || []).length === 1) {
    notes.push('*Este formato se usa cuando solo existe el PEO original (sin extensiones “-2”).');
  }
  if (summary.selectedIds && summary.selectedIds.length === 1 && (summary.items || []).length > 1) {
    notes.push('*Este formato muestra el PEO original y sus extensiones combinadas.');
  }
  if (summary.selectedIds && summary.selectedIds.length > 1) {
    notes.push(`*Reporte combinado de ${summary.selectedIds.length} PEOs base.`);
  }
  notes.push('*Las remisiones y facturas listadas arriba alimentan directamente el consumo del PEO.');
  const textStartY = boxTop + 12;
  notes.forEach((note, index) => {
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

async function generate(summary, branding = {}) {
  ensurePdfkitAvailable();
  const style = normalizeBranding(branding);

  return await new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'A4', margin: 40 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      drawLetterhead(doc, style);
      drawHeader(doc, summary, style);
      drawPoTable(doc, summary);
      drawMovements(doc, summary, style);
      drawSummaryBox(doc, summary);

      if (summary.selectedIds && summary.selectedIds.length > 1) {
        drawPieChart(doc, summary, style);
      } else if ((summary.items || []).length > 1) {
        drawPieChart(doc, summary, style);
      } else {
        drawStackedBar(doc, summary, style);
      }

      drawObservations(doc, summary);
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
