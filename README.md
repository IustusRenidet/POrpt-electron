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

### Variables de entorno para JasperReports
Si ves el mensaje `JasperReports no se inicializó. Verifica la configuración en reports/jasper.config.js.`, revisa las siguientes variables opcionales que permiten adaptar la ruta de los recursos:

| Variable | Descripción | Valor por defecto |
| --- | --- | --- |
| `JASPER_COMPILED_DIR` | Directorio donde se generan los `.jasper` compilados. Úsalo si el directorio por defecto es de solo lectura (por ejemplo dentro de `app.asar`). | `reports/compiled` (o `TMP/porpt-electron/jasper/compiled` si la app está empaquetada) |
| `JASPER_TEMPLATES_DIR` | Ruta a las plantillas `.jrxml`. | `reports/templates` |
| `JAYBIRD_JAR_PATH` | Ruta absoluta al `jaybird-full-5.x.x.jar`. | `reports/lib/jaybird-full-5.0.1.jar` |
| `JASPER_DATASOURCE_NAME` | Nombre del datasource JSON que se injecta en tiempo de ejecución. | `poSummaryJson` |
| `JASPER_JSON_QUERY` | Query JSON usada para el datasource. | `summary.items` |

También puedes definir `JASPER_WRITABLE_BASE` para controlar la raíz donde se crearán carpetas temporales y `JASPER_BASE_PATH` si deseas cargar la configuración desde otra carpeta.

## Funcionalidad
- Inicio de sesión con control de usuarios (admin: `admin` / `569OpEGvwh'8`).
- Selección de empresa y PO con búsqueda dinámica.
- Dashboard con gráficas separadas para remisiones y facturas, consumo acumulado y detalle de extensiones (`PO-xxx-2`).
- Alertas automáticas cuando el consumo supera el 10% del total.
- Reporte Jasper con totales, detalle de remisiones/facturas y observaciones.
