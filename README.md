# POrpt-electron

Dashboard y generador de reportes para control de Pedidos (PO), Remisiones y Facturas sobre Aspel SAE 9.

## Requisitos
- Node.js 20+
- Java 21 LTS (para JasperReports)
- Driver Jaybird (`jaybird-full-5.x.x.jar`) ubicado en `reports/lib`
- Librerías Jasper compiladas (se generan automáticamente al iniciar el servidor)
- Acceso a las bases Firebird de SAE 9 en la ruta configurada (`BASE_DIR` y `BASE_DIR_FB`)

## Dependencias destacadas
- [node-firebird](https://www.npmjs.com/package/node-firebird) para la conexión a Firebird
- [node-jasper](https://www.npmjs.com/package/node-jasper) + Jaybird para renderizar JasperReports
- [Chart.js](https://www.chartjs.org/) para la visualización en el dashboard

## Configuración
1. Copia el `jaybird-full-5.x.x.jar` a `reports/lib`.
2. Ajusta variables de entorno en `.env` si es necesario (host Firebird, rutas de empresa, etc.).
3. Instala dependencias: `npm install --package-lock-only` (requiere `python3-distutils` para compilar la librería `java`).
4. Inicia la app con `npm start`.

## Funcionalidad
- Inicio de sesión con control de usuarios (admin: `admin` / `569OpEGvwh'8`).
- Selección de empresa y PO con búsqueda dinámica.
- Dashboard con gráficas separadas para remisiones y facturas, consumo acumulado y detalle de extensiones (`PO-xxx-2`).
- Alertas automáticas cuando el consumo supera el 10% del total.
- Reporte Jasper con totales, detalle de remisiones/facturas y observaciones.
