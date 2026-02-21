# PDF Toolkit (PyMuPDF)

Plugin para SIS Hub orientado a operaciones PDF comunes y estables:

- Combinar varios PDFs.
- Dividir por pagina o por rangos.
- Editar (rotar, marca de agua de texto, metadatos).
- Organizar paginas (reordenar, invertir, duplicar o filtrar).
- Crear PDF desde imagenes.
- Extraer texto e imagenes embebidas.
- Exportar paginas PDF a imagenes (PNG/JPG).
- Proteger PDF y quitar contrasena.
- Previsualizar paginas y crear plantillas con hotspots.
- Cubrir contenido con cuadros blancos (overlay) o redaccion real.
- Digitalizar (flatten) PDF para convertirlo en imagen por pagina.

## Flujo de uso

1. Abre el plugin en SIS Hub.
2. Define carpeta de salida y nombre base.
3. (Opcional) usa `Guardar como (1 archivo)` para elegir ruta exacta del PDF de salida.
4. Elige la pestana de operacion.
5. Ejecuta la accion.
6. Revisa el log y los archivos generados.

## Plantillas y hotspots

1. Ve a `Plantilla/Hotspots`.
2. Carga un PDF en el preview.
3. Selecciona modo:
   - `Cuadro blanco`: tapa una zona (overlay visual).
   - `Hotspot texto`: define un campo para rellenar.
4. Arrastra el mouse sobre el area del PDF para crear la accion.
5. (Opcional) Marca:
   - `Digitalizar antes de aplicar` para flatten completo.
   - `Borrar texto real en cuadros blancos` para usar redaccion.
   - `Guardar hotspots como campos editables (AcroForm)` para que el PDF final se pueda llenar sin usar la plantilla JSON.
6. Completa `campo=valor`.
7. Usa `Validar plantilla`.
8. Ejecuta `Aplicar plantilla y guardar PDF`.

Tambien puedes guardar/cargar plantillas JSON para reutilizar campos.

## Notas de robustez

- Si un PDF esta protegido con contrasena, el plugin no lo procesa.
- Los metadatos se aplican solo si llenas los campos.
- La extraccion de imagenes usa xrefs unicos para evitar duplicados.
- El modo `cuadro blanco` solo cubre visualmente; para borrar contenido real usa redaccion o digitalizacion.
- Digitalizar elimina capas editables del PDF (incluyendo texto seleccionable) para dejar una version plana.

## Alcance

El plugin trabaja con PDFs en sentido general.
No valida cumplimiento normativo formal (PDF/A, PDF/X, PDF/UA, etc.).
