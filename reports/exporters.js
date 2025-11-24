const os = require('os');

let XLSX = null;
let xlsxLoadError = null;

try {
  XLSX = require('xlsx');
  console.log('V xlsx (excel) cargado correctamente');
} catch (err) {
  xlsxLoadError = err;
  console.warn('?? No se pudo cargar la dependencia opcional "xlsx" para exportar a Excel:', err.message);
}

function ensureXlsxAvailable() {
  if (XLSX) return true;
  const error = new Error(
    'El formato XLSX requiere la dependencia opcional "xlsx". Ejecuta "npm install" para habilitar la exportación a Excel.'
  );
  error.code = 'XLSX_UNAVAILABLE';
  if (xlsxLoadError) {
    error.cause = xlsxLoadError;
  }
  throw error;
}

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

function createCsv(summary = {}, options = {}) {
  const customization = options.customization || {};
  const csvCustomization = customization.csv || {};
  const includeMovements = customization.includeMovements !== false;
  const includePoResumen = csvCustomization.includePoResumen !== false;
  const includeRemisiones = includeMovements && csvCustomization.includeRemisiones !== false;
  const includeFacturas = includeMovements && csvCustomization.includeFacturas !== false;
  const includeTotals = csvCustomization.includeTotales !== false;
  const includeUniverseInfo = csvCustomization.includeUniverseInfo !== false;
  const includeObservations = customization.includeObservations !== false;

  const rows = [];
  rows.push(['tipo', 'po_id', 'referencia', 'fecha', 'descripcion', 'monto', 'comentarios'].join(','));
  const items = Array.isArray(summary.items) ? summary.items : [];
  items.forEach(item => {
    if (includePoResumen) {
      const commentParts = [];
      if (includeRemisiones) {
        commentParts.push(`Rem: ${formatCurrency(item.totalRem || item.totals?.totalRem || 0)}`);
      }
      if (includeFacturas) {
        commentParts.push(`Fac: ${formatCurrency(item.totalFac || item.totals?.totalFac || 0)}`);
      }
      const subtotal = Number(item.subtotal ?? 0);
      const total = Number(item.total ?? 0);
      if (includeTotals) {
        if (subtotal > 0 && Math.abs(subtotal - total) > 0.009) {
          commentParts.push(`Sub: ${formatCurrency(subtotal)}`);
        }
        const restante = item.restante ?? item.totals?.restante ?? 0;
        commentParts.push(`Rest: ${formatCurrency(restante)}`);
      }
      if (includeObservations && item.alertasTexto) {
        commentParts.push(item.alertasTexto);
      }
      rows.push([
        'PO',
        toCsvValue(item.id || ''),
        toCsvValue(item.baseId || ''),
        toCsvValue(item.fecha || ''),
        toCsvValue(item.descripcion || item.concepto || ''),
        formatCurrency(item.total || 0),
        toCsvValue(commentParts.join(' | '))
      ].join(','));
    }

    if (includeRemisiones) {
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
    }

    if (includeFacturas) {
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
    }
  });

  const totals = summary.totals || {};
  if (includeTotals) {
    if (rows.length > 1) {
      rows.push('');
    }
    rows.push(['RESUMEN', '', '', '', 'Total autorizado', formatCurrency(totals.total || 0), ''].join(','));
    rows.push(['RESUMEN', '', '', '', 'Total remisiones', formatCurrency(totals.totalRem || 0), ''].join(','));
    rows.push(['RESUMEN', '', '', '', 'Total facturas', formatCurrency(totals.totalFac || 0), ''].join(','));
    rows.push(['RESUMEN', '', '', '', 'Restante', formatCurrency(totals.restante || 0), ''].join(','));
  }

  if (includeUniverseInfo && summary.universe) {
    if (rows[rows.length - 1] !== '') {
      rows.push('');
    }
    const universe = summary.universe;
    rows.push(['UNIVERSO', '', '', '', 'Filtro seleccionado', toCsvValue(universe.label || universe.shortLabel || ''), ''].join(','));
    rows.push(['UNIVERSO', '', '', '', 'Descripción', toCsvValue(universe.description || ''), ''].join(','));
  }

  if (includeTotals || includeUniverseInfo) {
    rows.push(['UNIVERSO', '', '', '', 'Empresa', toCsvValue(summary.companyName || summary.empresaLabel || summary.empresa || ''), ''].join(','));
  }

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

function createXlsx(summary = {}, options = {}) {
  ensureXlsxAvailable();
  const customization = options.customization || {};
  const csvCustomization = customization.csv || {};
  const includeMovements = customization.includeMovements !== false;
  const includePoResumen = csvCustomization.includePoResumen !== false;
  const includeRemisiones = includeMovements && csvCustomization.includeRemisiones !== false;
  const includeFacturas = includeMovements && csvCustomization.includeFacturas !== false;
  const includeTotals = csvCustomization.includeTotales !== false;
  const includeUniverseInfo = csvCustomization.includeUniverseInfo !== false;
  const includeObservations = customization.includeObservations !== false;

  const rows = [];
  rows.push(['tipo', 'po_id', 'referencia', 'fecha', 'descripcion', 'monto', 'comentarios']);
  const items = Array.isArray(summary.items) ? summary.items : [];
  items.forEach(item => {
    if (includePoResumen) {
      const commentParts = [];
      if (includeRemisiones) {
        commentParts.push(`Rem: ${formatCurrency(item.totalRem || item.totals?.totalRem || 0)}`);
      }
      if (includeFacturas) {
        commentParts.push(`Fac: ${formatCurrency(item.totalFac || item.totals?.totalFac || 0)}`);
      }
      const subtotal = Number(item.subtotal ?? 0);
      const total = Number(item.total ?? 0);
      if (includeTotals) {
        if (subtotal > 0 && Math.abs(subtotal - total) > 0.009) {
          commentParts.push(`Sub: ${formatCurrency(subtotal)}`);
        }
        const restante = item.restante ?? item.totals?.restante ?? 0;
        commentParts.push(`Rest: ${formatCurrency(restante)}`);
      }
      if (includeObservations && item.alertasTexto) {
        commentParts.push(item.alertasTexto);
      }
      rows.push([
        'PO',
        item.id || '',
        item.baseId || '',
        item.fecha || '',
        item.descripcion || item.concepto || '',
        Number(item.total || 0),
        commentParts.join(' | ')
      ]);
    }

    if (includeRemisiones) {
      const remisiones = Array.isArray(item.remisiones) ? item.remisiones : [];
      remisiones.forEach(rem => {
        rows.push([
          'REMISION',
          item.id || '',
          rem.id || rem.folio || '',
          rem.fecha || '',
          rem.descripcion || rem.concepto || '',
          Number(rem.monto || 0),
          rem.observaciones || ''
        ]);
      });
    }

    if (includeFacturas) {
      const facturas = Array.isArray(item.facturas) ? item.facturas : [];
      facturas.forEach(fac => {
        rows.push([
          'FACTURA',
          item.id || '',
          fac.id || fac.folio || '',
          fac.fecha || '',
          fac.descripcion || fac.concepto || '',
          Number(fac.monto || fac.total || 0),
          fac.observaciones || ''
        ]);
      });
    }
  });

  const totals = summary.totals || {};
  if (includeTotals) {
    rows.push([]);
    rows.push(['RESUMEN', '', '', '', 'Total autorizado', Number(totals.total || 0), '']);
    rows.push(['RESUMEN', '', '', '', 'Total remisiones', Number(totals.totalRem || 0), '']);
    rows.push(['RESUMEN', '', '', '', 'Total facturas', Number(totals.totalFac || 0), '']);
    rows.push(['RESUMEN', '', '', '', 'Restante', Number(totals.restante || 0), '']);
  }

  if (includeUniverseInfo && summary.universe) {
    const universe = summary.universe;
    rows.push([]);
    rows.push(['UNIVERSO', '', '', '', 'Filtro seleccionado', universe.label || universe.shortLabel || '', '']);
    rows.push(['UNIVERSO', '', '', '', 'Descripción', universe.description || '', '']);
  }

  if (includeTotals || includeUniverseInfo) {
    rows.push([
      'UNIVERSO',
      '',
      '',
      '',
      'Empresa',
      summary.companyName || summary.empresaLabel || summary.empresa || '',
      ''
    ]);
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Reporte');

  const metaRows = [
    ['Generado', new Date().toISOString()],
    ['Formato', 'xlsx'],
    ['Empresa', summary.companyName || summary.empresaLabel || summary.empresa || ''],
    ['PO seleccionadas', (summary.selectedIds || []).join(', ') || 'N/D']
  ];
  if (summary.universe) {
    metaRows.push(['Universo', summary.universe.label || summary.universe.shortLabel || '']);
  }
  const metaSheet = XLSX.utils.aoa_to_sheet(metaRows);
  XLSX.utils.book_append_sheet(workbook, metaSheet, 'Meta');

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = {
  createCsv,
  createJson,
  createXlsx,
  isXlsxAvailable: () => Boolean(XLSX),
  getXlsxUnavailableMessage: () => {
    if (xlsxLoadError) {
      return `El formato XLSX requiere la dependencia opcional "xlsx". Detalle: ${xlsxLoadError.message}`;
    }
    return 'El formato XLSX requiere la dependencia opcional "xlsx". Ejecuta "npm install" para habilitarlo.';
  }
};
