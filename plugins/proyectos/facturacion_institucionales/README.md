# Facturacion Institucionales

Plugin para SIS Hub orientado a ejecucion de flujos de facturacion institucional.

Fuente de configuracion:
- `Facturacion_Institucionales_Maestro.xlsx`

Credenciales Oracle:
- Se leen desde Windows Credential Manager via `keyring`
- Nombre por defecto de la credencial: `OperacionesBD`
- Conexion Oracle configurable por variables de entorno: `FIHUB_ORACLE_HOST`, `FIHUB_ORACLE_PORT`, `FIHUB_ORACLE_SERVICE`

Capacidades principales:
- Busqueda de cliente por RUT
- Descarga de factura PDF y XML (servicio SOAP)
- Resumen de transacciones en PDF
- Reporte Excel por query configurada
- Archivo plano por query configurada
- Envio manual (Outlook Display con adjuntos) o entrega de carpeta para TOMY/RPA

El plugin usa un workbook maestro normalizado y evita dependencias de imagenes para el resumen.
