const PDFDocument = require('pdfkit');

function formatCurrency(value) {
  const amount = Number(value || 0);
  return amount.toLocaleString('es-MX', {
    style: 'currency',
    currency: 'MXN',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function addSectionTitle(doc, text) {
  doc.moveDown(0.75);
  doc.fontSize(13).fillColor('#1f2937').text(text, { underline: true });
  doc.moveDown(0.25);
}

function renderItems(doc, title, entries = []) {
  addSectionTitle(doc, title);
  if (!entries.length) {
    doc.fontSize(11).fillColor('#475569').text('Sin registros');
    return;
  }
  entries.forEach(item => {
    const fecha = item.fecha || '-';
    const monto = formatCurrency(item.monto);
    doc.fontSize(11).fillColor('#0f172a').text(`${item.id} • ${fecha} • ${monto}`);
  });
}

function renderAlerts(doc, alerts = []) {
  addSectionTitle(doc, 'Alertas');
  if (!alerts.length) {
    doc.fontSize(11).fillColor('#15803d').text('Sin alertas registradas.');
    return;
  }
  alerts.forEach(alert => {
    const label = alert.type ? alert.type.toUpperCase() : 'INFO';
    const color = alert.type === 'danger' ? '#b91c1c' : alert.type === 'warning' ? '#d97706' : '#2563eb';
    doc.fontSize(11).fillColor(color).text(`[${label}] ${alert.message}`);
  });
}

function renderPoItem(doc, item) {
  const totals = item.totals || {};
  doc.moveDown(0.5);
  doc.fontSize(14).fillColor('#111827').text(`PO ${item.id}`, { continued: true });
  doc.fontSize(10).fillColor('#6366f1').text(`   (Extensión de ${item.baseId || item.id})`);
  doc.fontSize(11).fillColor('#1f2937').text(`Fecha: ${item.fecha || '-'}`);
  doc.fontSize(11).fillColor('#1f2937').text(`Total autorizado: ${formatCurrency(item.total)}`);
  doc.fontSize(11).fillColor('#1f2937').text(`Consumo remisiones: ${formatCurrency(totals.totalRem)}`);
  doc.fontSize(11).fillColor('#1f2937').text(`Consumo facturas: ${formatCurrency(totals.totalFac)}`);
  doc.fontSize(11).fillColor('#1f2937').text(`Disponible: ${formatCurrency(totals.restante)}`);
  renderItems(doc, 'Remisiones', item.remisiones || []);
  renderItems(doc, 'Facturas', item.facturas || []);
  renderAlerts(doc, item.alerts || []);
  doc.moveDown();
  doc.moveTo(doc.x, doc.y).lineTo(doc.page.width - doc.page.margins.right, doc.y).stroke('#e5e7eb');
}

function buildHeader(doc, summary) {
  const headerColor = '#0f172a';
  doc.fillColor(headerColor).fontSize(20).text('Reporte de seguimiento PO', { align: 'center' });
  doc.moveDown(0.5);
  doc.fontSize(12).fillColor('#1e293b').text(`Empresa: ${summary.companyName || summary.empresa || '-'}`, { align: 'center' });
  doc.fontSize(12).fillColor('#1e293b').text(`PO seleccionada: ${summary.selectedId || summary.baseId || '-'}`, { align: 'center' });
  doc.moveDown();
}

function renderTotals(doc, totals = {}) {
  addSectionTitle(doc, 'Resumen general');
  doc.fontSize(11).fillColor('#0f172a').text(`Total autorizado: ${formatCurrency(totals.total)}`);
  doc.fontSize(11).fillColor('#0f172a').text(`Total remisiones: ${formatCurrency(totals.totalRem)} (${(totals.porcRem || 0).toFixed(2)}%)`);
  doc.fontSize(11).fillColor('#0f172a').text(`Total facturas: ${formatCurrency(totals.totalFac)} (${(totals.porcFac || 0).toFixed(2)}%)`);
  doc.fontSize(11).fillColor('#0f172a').text(`Disponible: ${formatCurrency(totals.restante)} (${(totals.porcRest || 0).toFixed(2)}%)`);
}

async function generate(summary) {
  return await new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'A4', margin: 40 });
      const chunks = [];
      doc.on('data', chunk => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      buildHeader(doc, summary);
      renderTotals(doc, summary.totals || {});
      renderAlerts(doc, summary.alerts || []);

      (summary.items || []).forEach(item => {
        doc.addPage();
        renderPoItem(doc, item);
      });

      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}

module.exports = {
  generate
};
