const os = require('os');

function toCsvValue(value) {
  if (value === null || value === undefined) return '';
  const stringValue = String(value).replace(/\r?\n/gu, ' ');
  if (/[",]/u.test(stringValue)) {
    return `"${stringValue.replace(/"/gu, '""')}"`;
  }
  return stringValue;
}

function formatCurrency(value) {
  const number = Number(value);
  if (Number.isNaN(number)) return '0';
  return number.toFixed(2);
}

function createCsv(summary = {}) {
  const rows = [];
  rows.push(['tipo', 'po_id', 'referencia', 'fecha', 'descripcion', 'monto', 'comentarios'].join(','));
  const items = Array.isArray(summary.items) ? summary.items : [];
  items.forEach(item => {
    rows.push([
      'PO',
      toCsvValue(item.id || ''),
      toCsvValue(item.baseId || ''),
      toCsvValue(item.fecha || ''),
      toCsvValue(item.descripcion || item.concepto || ''),
      formatCurrency(item.total || 0),
      toCsvValue(`Rem: ${formatCurrency(item.totalRem || 0)} | Fac: ${formatCurrency(item.totalFac || 0)} | Rest: ${formatCurrency(item.restante || 0)}`)
    ].join(','));
    const remisiones = Array.isArray(item.remisiones) ? item.remisiones : [];
    remisiones.forEach(rem => {
      rows.push([
        'REMISION',
        toCsvValue(item.id || ''),
        toCsvValue(rem.id || rem.folio || ''),
        toCsvValue(rem.fecha || ''),
        toCsvValue(rem.descripcion || rem.concepto || ''),
        formatCurrency(rem.monto || 0),
        toCsvValue(rem.observaciones || '')
      ].join(','));
    });
    const facturas = Array.isArray(item.facturas) ? item.facturas : [];
    facturas.forEach(fac => {
      rows.push([
        'FACTURA',
        toCsvValue(item.id || ''),
        toCsvValue(fac.id || fac.folio || ''),
        toCsvValue(fac.fecha || ''),
        toCsvValue(fac.descripcion || fac.concepto || ''),
        formatCurrency(fac.monto || fac.total || 0),
        toCsvValue(fac.observaciones || '')
      ].join(','));
    });
  });
  rows.push('');
  const totals = summary.totals || {};
  rows.push(['RESUMEN', '', '', '', 'Total autorizado', formatCurrency(totals.total || 0), ''].join(','));
  rows.push(['RESUMEN', '', '', '', 'Total remisiones', formatCurrency(totals.totalRem || 0), ''].join(','));
  rows.push(['RESUMEN', '', '', '', 'Total facturas', formatCurrency(totals.totalFac || 0), ''].join(','));
  rows.push(['RESUMEN', '', '', '', 'Restante', formatCurrency(totals.restante || 0), ''].join(','));
  return Buffer.from(rows.join(os.EOL), 'utf8');
}

function createJson(summary = {}, metadata = {}) {
  const payload = {
    generatedAt: new Date().toISOString(),
    ...metadata,
    summary
  };
  return Buffer.from(JSON.stringify(payload, null, 2), 'utf8');
}

module.exports = {
  createCsv,
  createJson
};
