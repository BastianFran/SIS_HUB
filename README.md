# SIS Hub (Soporte Institucional Hub)

Aplicacion Tkinter para ejecutar proyectos (plugins) dentro de una sola ventana.
Funciona con Python 3.10 en desarrollo y se puede distribuir como ejecutable
para usuarios finales.

---

## Caracteristicas principales
- Navegacion simple: Inicio > Categorias > Funciones > Atras/Inicio.
- Plugins desacoplados: cada proyecto monta su UI dentro del contenedor del Hub.
- Contrato de plugin estable y minimo (plugin.json + funcion crear_interfaz).
- Sin dependencias externas (stdlib + tkinter).

---

## Requisitos

**Usuarios del ejecutable (.exe)**
- Windows 10/11, no requieren Python instalado.
- Crear previamente `C:\Users\<usuario>\Desktop\sishub_log.txt` (archivo vacio).
- Descargar la carpeta compartida completa (ejecutable + plugins + assets).

**Desarrolladores (ejecucion directa con Python)**
- Python 3.10.9 (Anaconda 2023 o CPython 3.10).
- `tkinter` disponible.
- Archivo `sishub_log.txt` creado en el Escritorio para capturar los registros.

---

## Estructura del proyecto

```
SIS_Hub/
|-- main.py
|-- sis_hub_core/
|   |-- app.py            # ventana principal, router, vistas basicas
|   |-- router.py         # pila de vistas (push/pop/home)
|   |-- catalog.py        # descubrimiento de plugins (plugin.json)
|   |-- loader.py         # import dinamico seguro de entrypoints
|   |-- logging_utils.py  # logger que escribe en Desktop\sishub_log.txt
|   `-- constants.py      # constantes de UI, version, paths
|-- plugins/
|   `-- <categoria>/<plugin>/plugin.json, src/
|-- assets/
|-- docs/
`-- scripts/
```

---

## Uso

Antes de ejecutar el Hub verifica que existe `Desktop\sishub_log.txt`. Si falta,
la aplicacion mostrara un mensaje y se cerrara.

### Opcion A: ejecutar con Python
1. Abrir una terminal con el entorno de Python 3.10 activo.
2. Ubicarse en la carpeta `SIS_Hub/`.
3. Ejecutar `python main.py`.

### Opcion B: generar ejecutable onefile
1.Agregar librerías al main.py y probar el import
2.Abrir anaconda prompt y usar el entorno sishub o donde tengas pyinstaller
3.cd a la ruta raíz en este caso C:\Users\BBRUNA\OneDrive - Banchile\Soporte Institucionales\Desarrollos\SIS_Hub
4.pyinstaller main.py --onefile --windowed --clean --name SIS_Hub
5.Mover el exe de la carpeta dist a la raíz
#################################################
1.Agregar librerías al main.py y probar el import
2.ejecutar build.bat
#################################################
---

## Distribucion mediante OneDrive
- Mantener juntos `SIS_Hub.exe`, `plugins/`, `assets/` y cualquier recurso
  adicional; el Hub escanea `plugins/` junto al ejecutable en cada arranque.
- Para actualizar basta con reemplazar archivos dentro de `plugins/`; no es
  necesario reconstruir el `.exe`.
- Evitar renombrar la carpeta compartida si existen accesos directos.
- Confirmar que todos los usuarios tienen permisos de lectura sobre la carpeta.

---

## Agregar un nuevo plugin
1. Crear `plugins/<categoria>/<nombre>/plugin.json` y la carpeta `src/`.
2. `plugin.json` minimo:
   ```json
   {
     "id": "mi_plugin",
     "name": "Mi Plugin",
     "category": "utilities",
     "description": "Descripcion corta.",
     "entrypoint": "main:crear_interfaz",
     "requires": [],
     "enabled": true
   }
   ```
3. En `src/main.py` definir `crear_interfaz(parent, context)`:
   - Construir los widgets dentro de `parent`.
   - No crear `Tk()` ni llamar a `mainloop()`.
   - Para tareas largas usar hilos/subprocesos y `after()` para actualizar UI.

Los plugins pueden obtener un logger mediante `context.logger()`; los mensajes
se escriben en `Desktop\sishub_log.txt`.

---

## Buenas practicas
- Seguir PEP8 y documentar brevemente las funciones publicas.
- Mantener las tareas pesadas fuera del hilo principal (usar hilos/after).
- Evitar rutas absolutas; usar `context.root` o `user_data_dir()` para datos.
- Registrar la actividad solo a traves de `context.logger()`.
- Documentar dependencias adicionales en `plugin.json`.

---

## Empaquetado y verificacion
- `scripts/build_pyinstaller.bat` sigue disponible para generar un build onedir
  durante el desarrollo.
- Para el build final usar el comando onefile y preparar la carpeta compartida.
- Probar el paquete en un equipo limpio o usuario distinto antes de publicar.

---

## Solucion de problemas
- **La app no abre**: ejecutar `python main.py` y revisar `Desktop\sishub_log.txt`.
- **Plugin no carga**: revisar `plugin.json` y `crear_interfaz`; el error queda en
  el log.
- **UI congelada**: asegurate de que la tarea pesada vaya en un hilo/subproceso.
- **Permisos**: confirmar acceso de lectura a la carpeta compartida y escritura en
  el Escritorio para el archivo de log.

---

## Licencia
Uso interno. Documentar responsables y el area de Soporte Institucional.
