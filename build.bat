@echo off
setlocal
call conda activate sishub
cd /d "%~dp0"
 
pyinstaller main.py ^
  --onefile ^
  --windowed ^
  --clean ^
  --name SIS_Hub
 
echo.
echo Ejecutable listo: "%~dp0dist\SIS_Hub.exe"
endlocal