@echo off
title Cascadas Hotel - Informe RDS
cd /d "%~dp0"

echo ========================================================================
echo   CASCADAS HOTEL - INFORME RDS
echo ========================================================================
echo   Carpeta: %cd%
echo.

rem --- Si el archivo llego como .txt (Windows esconde las extensiones y al
rem     renombrarlo suele quedar "vhf_rds.py.txt"), se copia con el nombre
rem     correcto y listo. ---
if not exist "vhf_rds.py" (
  if exist "vhf_rds.py.txt" (
    echo   El programa venia como .txt: lo dejo con el nombre correcto.
    copy /y "vhf_rds.py.txt" "vhf_rds.py" >nul
  )
)
if not exist "vhf_rds.py" (
  if exist "vhf-rds-py.txt" (
    echo   El programa venia como .txt: lo dejo con el nombre correcto.
    copy /y "vhf-rds-py.txt" "vhf_rds.py" >nul
  )
)

if not exist "vhf_rds.py" (
  echo.
  echo   NO ENCUENTRO EL ARCHIVO vhf_rds.py EN ESTA CARPETA.
  echo.
  echo   Lo que hay aqui adentro es:
  echo   ------------------------------------------------------------
  dir /b
  echo   ------------------------------------------------------------
  echo.
  echo   Copia esta lista y mandamela.
  echo.
  pause
  exit /b 1
)

rem --- Buscar Python. "py" es el lanzador que instala python.org; si no
rem     esta, se prueba "python" a secas. ---
set EJECUTAR=
where py >nul 2>nul
if %errorlevel%==0 set EJECUTAR=py

if "%EJECUTAR%"=="" (
  where python >nul 2>nul
  if %errorlevel%==0 set EJECUTAR=python
)

if "%EJECUTAR%"=="" (
  echo.
  echo   NO ENCUENTRO PYTHON EN ESTE COMPUTADOR.
  echo.
  echo   Instalalo desde python.org y marca la casilla
  echo   "Add python.exe to PATH" durante la instalacion.
  echo.
  pause
  exit /b 1
)

echo   Usando: %EJECUTAR%
echo.

%EJECUTAR% "%~dp0vhf_rds.py" %*

echo.
echo ========================================================================
echo   La ventana se queda abierta para que puedas leer lo de arriba.
echo ========================================================================
pause
