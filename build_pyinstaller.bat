@echo off
REM Build onedir executable with PyInstaller (must be installed in the build env)
setlocal
set APP=SIS_Hub
set VERSION=0.1.0
set DISTDIR=dist\%APP%_%VERSION%

if exist dist rmdir /s /q dist
pyinstaller --noconfirm --clean ^
  --name %APP% ^
  --onedir ^
  --windowed ^
  main.py

if not exist "%DISTDIR%" mkdir "%DISTDIR%"
xcopy /e /i /y "dist\%APP%" "%DISTDIR%"
xcopy /e /i /y assets "%DISTDIR%\assets"
xcopy /e /i /y plugins "%DISTDIR%\plugins"
xcopy /e /i /y docs "%DISTDIR%\docs"
echo Build listo en %DISTDIR%
endlocal
