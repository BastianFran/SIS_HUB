@echo off
REM Cambiar al directorio del proyecto
cd /d "C:/Users/csantos/OneDrive - Banchile/Escritorio/Proyectos\IRVI\src"

REM Verificar si la ruta de python.exe en Anaconda existe
IF EXIST "%USERPROFILE%\Anaconda3\python.exe" (
    "%USERPROFILE%\Anaconda3\python.exe" main.py
) ELSE (
    REM Usar la ruta alternativa si no se encuentra python.exe en Anaconda
    IF EXIST "C:\Users\%USERNAME%\AppData\Local\anaconda3\python.exe" (
        "C:\Users\%USERNAME%\AppData\Local\anaconda3\python.exe" main.py
    ) ELSE (
        echo No se encontró python.exe en ninguna de las rutas especificadas.
        pause
        exit /b 1
    )
)

REM Mantener la ventana abierta
pause
