@echo off
title Cascadas Hotel - Informe RDS
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py "%~dp0vhf_rds.py" %*
  goto :fin
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "%~dp0vhf_rds.py" %*
  goto :fin
)

echo.
echo   No encuentro Python en este computador.
echo   Instalalo desde python.org marcando "Add python.exe to PATH"
echo   y despues vuelve a apretar este acceso directo.
echo.
pause

:fin
