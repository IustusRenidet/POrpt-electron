const path = require('path');

const basePath = __dirname;
const compiledPath = path.join(basePath, 'compiled');
const templatesPath = path.join(basePath, 'templates');

module.exports = {
  jasperPath: compiledPath,
  templatesPath,
  defaultReport: 'poSummary',
  reports: {
    poSummary: {
      jasper: path.join(compiledPath, 'po_summary.jasper'),
      jrxml: path.join(templatesPath, 'po_summary.jrxml'),
      conn: 'poSummaryJson'
    }
  },
  dataSourceName: 'poSummaryJson',
  dataSources: {
    poSummaryJson: {
      driver: 'json',
      jsonQuery: 'summary.items',
      data: []
    }
  },
  drivers: {
    jaybird: {
      jar: path.join(basePath, 'lib', 'jaybird-full-5.0.1.jar'),
      className: 'org.firebirdsql.jdbc.FBDriver'
    }
  }
};
