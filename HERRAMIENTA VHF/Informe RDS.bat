@echo off
title Cascadas Hotel - Informe RDS
cd /d "%~dp0"

echo ========================================================================
echo   CASCADAS HOTEL - INFORME RDS
echo ========================================================================
echo   Carpeta: %cd%
echo.

rem --- Se busca el programa aguantando como haya quedado el nombre: con guion
rem     bajo o sin el, y con el .txt pegado atras (Windows esconde las
rem     extensiones y al renombrar suele quedar "vhf_rds.py.txt"). ---
set PROGRAMA=

if exist "vhf_rds.py" set PROGRAMA=vhf_rds.py
if "%PROGRAMA%"=="" if exist "vhfrds.py"  set PROGRAMA=vhfrds.py
if "%PROGRAMA%"=="" if exist "vhf-rds.py" set PROGRAMA=vhf-rds.py

if "%PROGRAMA%"=="" (
  for %%A in ("vhf_rds.py.txt" "vhfrds.py.txt" "vhf-rds-py.txt" "vhf_rds-py.txt") do (
    if exist "%%~A" (
      echo   El programa venia como .txt: lo dejo con el nombre correcto.
      copy /y "%%~A" "vhf_rds.py" >nul
    )
  )
  if exist "vhf_rds.py" set PROGRAMA=vhf_rds.py
)

rem Ultimo intento: cualquier archivo .py que hable de RDS.
if "%PROGRAMA%"=="" (
  for %%A in (*rds*.py) do set PROGRAMA=%%A
)

if "%PROGRAMA%"=="" (
  echo.
  echo   NO ENCUENTRO EL PROGRAMA EN ESTA CARPETA.
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

echo   Programa: %PROGRAMA%

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

%EJECUTAR% "%PROGRAMA%" %*

echo.
echo ========================================================================
echo   La ventana se queda abierta para que puedas leer lo de arriba.
echo ========================================================================
pause
