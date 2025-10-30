# SIS Hub (Soporte Institucional Hub)

Aplicación **Tkinter** para centralizar proyectos (plugins) y ejecutarlos dentro de una sola ventana. Pensada para nuestro entorno corporativo, con ejecución directa en **Python 3.10.9 (Anaconda 2023)** o distribuible como **.exe** (PyInstaller).

## Características
- Navegación simple: Inicio > Categorías > Funciones > Atrás/Inicio.
- Plugins embebidos: cada mini‑proyecto monta su UI dentro del contenedor del Hub.
- Contrato de plugin mínimo y estable.
- Sin dependencias exóticas (stdlib + tkinter).

---

## Requisitos

**Usuarios .exe**
- Windows 10/11. Sin necesidad de Python instalado.
- Permisos de lectura/escritura en su carpeta de usuario (logs).
- Si un plugin requiere Outlook/COM u otros, debe estar presente en el equipo (no incluido por defecto).

**Usuarios Python (dev)**
- **Anaconda 2023** con **Python 3.10.9** (o CPython 3.10.9 en Windows).
- `tkinter` disponible (incluido en CPython para Windows).

---

## Estructura del proyecto

```
SIS_Hub/
├─ main.py
├─ sis_hub_core/
│  ├─ app.py                # ventana principal, router, vistas básicas
│  ├─ router.py             # administración de pila de vistas (push/pop/home)
│  ├─ catalog.py            # descubrimiento de plugins (plugin.json)
│  ├─ loader.py             # import dinámico seguro de entrypoints
│  ├─ logging_utils.py      # logger con rotación en %LOCALAPPDATA%\SISHub\logs
│  └─ constants.py          # constantes de UI, versión, paths
├─ plugins/
│  └─ utilities/
│     ├─ demo_ui/
│     │  ├─ plugin.json
│     │  └─ src/
│     │     └─ main.py      # crear_interfaz(parent, context)
│     └─ demo_task/
│        ├─ plugin.json
│        └─ src/
│           └─ main.py      # crear_interfaz(parent, context) con progreso
├─ assets/
├─ docs/
│  └─ CHANGELOG.md
├─ scripts/
│  ├─ build_pyinstaller.bat
│  └─ build.py
├─ dist/
└─ logs/
```

---

## Uso

### Opción A: Ejecutar con Python (Anaconda 2023, Python 3.10.9)
1. Abrir **Anaconda Prompt** (o terminal con ese Python activo).
2. Ir a la carpeta del proyecto `SIS_Hub/`.
3. Ejecutar: `python main.py`

### Opción B: Ejecutable (.exe, PyInstaller)
- Usar los scripts de `scripts/` para generar una carpeta **onedir** lista para distribuir.
- Entregar la carpeta `dist/SIS_Hub_x.y.z/` al usuario final. Abrir `SIS_Hub.exe`.

---

## Añadir un nuevo plugin

1. Crear una carpeta en `plugins/<categoria>/<nombre>/` con:
   - `plugin.json` (manifiesto con `id`, `name`, `category`, `entrypoint`, `requires[]` opcional)
   - `src/main.py` con **`crear_interfaz(parent, context)`** que monta la UI dentro de `parent`.

2. `plugin.json` (ejemplo mínimo):
```json
{
  "id": "mi_plugin",
  "name": "Mi Plugin",
  "category": "utilities",
  "description": "Descripción corta.",
  "entrypoint": "main:crear_interfaz",
  "requires": [],
  "enabled": true
}
```

3. `src/main.py` debe **NO** crear `Tk()` ni llamar `mainloop()`. Debe:
   - construir un `Frame` dentro de `parent`
   - montar widgets ahí
   - retornar opcionalmente el `Frame`
   - si usa tareas largas: ejecutarlas en thread/subprocess y actualizar la UI con `after()`
   - limpiar timers/hilos en `<Destroy>`

> Ver ejemplos incluidos: `demo_ui` y `demo_task`.

---

## Buenas prácticas (PEP8 y robustez)
- PEP8 y docstrings breves en funciones públicas.
- Nada de tareas largas en el hilo de UI; usar thread/subprocess y `after()`.
- Sin rutas absolutas; todo relativo a `__file__` o a la carpeta de datos del usuario: `%LOCALAPPDATA%\SISHub`.
- Logging con rotación: el Hub ya lo configura; desde el plugin usa `context.logger()`.
- Evitar dependencias externas: si un plugin las requiere, documentarlas en su `plugin.json` y canalizarlas por TI.

---

## Empaquetado (PyInstaller)
- Requiere instalar `pyinstaller` en un entorno de build.
- Script recomendado: `scripts/build_pyinstaller.bat` (genera *onedir*).
- Verificar el build en un equipo “limpio” antes de distribuir.

---

## Solución de problemas
- **La app no abre**: use `python main.py` desde consola para ver mensajes; revise `%LOCALAPPDATA%\SISHub\logs\sishub.log`.
- **Un plugin no carga**: revise su `plugin.json` y la función `crear_interfaz`; errores aparecen en el log.
- **UI congelada**: verifique que la tarea pesada esté en un thread y use `after()` para actualizar la UI.
- **Permisos/paths**: no usar rutas de red sin confirmar permisos; preferir `%LOCALAPPDATA%\SISHub`.

---

## Licencia y contacto
Uso interno. Documentar responsables y área de Soporte Institucional.
