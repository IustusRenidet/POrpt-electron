# POrpt-electron

Dashboard y generador de reportes para control de Pedidos (PO), Remisiones y Facturas sobre Aspel SAE 9.

## Requisitos
- Node.js 20+
- Acceso a las bases Firebird de SAE 9 en la ruta configurada (`BASE_DIR` y `BASE_DIR_FB`)

## Dependencias destacadas
- [node-firebird](https://www.npmjs.com/package/node-firebird) para la conexión a Firebird
- [pdfkit](https://www.npmjs.com/package/pdfkit) para el motor de reporte "PDF directo"
- [Chart.js](https://www.chartjs.org/) para la visualización en el dashboard

## Configuración
1. Ajusta variables de entorno en `.env` si es necesario (host Firebird, rutas de empresa, etc.).
2. Instala dependencias: `npm install`.
3. Inicia la app con `npm start`.

## Funcionalidad
- Inicio de sesión con control de usuarios (admin: `admin` / `569OpEGvwh'8`).
- Selección de empresa y PO con búsqueda dinámica.
- Dashboard con gráficas separadas para remisiones y facturas, consumo acumulado y detalle de extensiones (`PO-xxx-2`).
- Alertas automáticas cuando el consumo supera el 10% del total.
- Exportación de reportes en PDF directo, CSV y JSON con opciones de personalización.
