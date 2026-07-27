@echo off
rem ===========================================================
rem  CASCADAS HOTEL - AUDITORIA DIARIA
rem  Doble clic aqui para abrir la herramienta.
rem ===========================================================
title Auditoria Cascadas
cd /d "%~dp0"

rem Se prueba primero con el lanzador "py", que Windows instala junto a
rem Python y encuentra siempre. Si no esta, se cae a "python".
where py >nul 2>nul
if %errorlevel%==0 (
  py auditoria.py
  goto :fin
)

where python >nul 2>nul
if %errorlevel%==0 (
  python auditoria.py
  goto :fin
)

echo.
echo ==========================================================
echo  No encuentro Python en este computador.
echo.
echo  Instalalo desde python.org y marca la casilla
echo  "Add python.exe to PATH" antes de apretar Install.
echo ==========================================================
echo.
pause
exit /b 1

:fin
if errorlevel 1 (
  echo.
  echo ==========================================================
  echo  La herramienta se cerro con un error.
  echo  Saca una foto del texto de arriba y mandalo.
  echo ==========================================================
  echo.
  pause
)
