@echo off
title Cascadas Hotel - Informe RDS
cd /d "%~dp0"

echo ========================================================================
echo   CASCADAS HOTEL - INFORME RDS
echo ========================================================================
echo   Carpeta: %cd%
echo.

rem =======================================================================
rem  1. BUSCAR EL PROGRAMA
rem     Se aguanta como haya quedado el nombre: con guion bajo o sin el, y
rem     con el .txt pegado atras (Windows esconde las extensiones y al
rem     renombrar suele quedar "vhf_rds.py.txt").
rem =======================================================================
set PROGRAMA=

if exist "vhf_rds.py"  set PROGRAMA=vhf_rds.py
if not "%PROGRAMA%"=="" goto :tengo_programa

if exist "vhfrds.py"   set PROGRAMA=vhfrds.py
if not "%PROGRAMA%"=="" goto :tengo_programa

if exist "vhf-rds.py"  set PROGRAMA=vhf-rds.py
if not "%PROGRAMA%"=="" goto :tengo_programa

if exist "vhf_rds.py.txt"  copy /y "vhf_rds.py.txt"  "vhf_rds.py" >nul
if exist "vhfrds.py.txt"   copy /y "vhfrds.py.txt"   "vhf_rds.py" >nul
if exist "vhf-rds-py.txt"  copy /y "vhf-rds-py.txt"  "vhf_rds.py" >nul
if exist "vhf_rds.py"  set PROGRAMA=vhf_rds.py
if not "%PROGRAMA%"=="" echo   El programa venia como .txt: lo deje con el nombre correcto.
if not "%PROGRAMA%"=="" goto :tengo_programa

rem Ultimo intento: cualquier archivo .py que hable de RDS.
for %%A in (*rds*.py) do set PROGRAMA=%%A
if not "%PROGRAMA%"=="" goto :tengo_programa

echo.
echo   NO ENCUENTRO EL PROGRAMA EN ESTA CARPETA.
echo.
echo   Lo que hay aqui adentro es:
echo   ------------------------------------------------------------
dir /b
echo   ------------------------------------------------------------
echo.
pause
exit /b 1

:tengo_programa
echo   Programa: %PROGRAMA%


rem =======================================================================
rem  2. BUSCAR PYTHON
rem     Visual Hotel es un programa de 32 bits: si esta el Python de 32 con
rem     pywinauto instalado, se usa ese, que lee bien el listado de informes.
rem     OJO: "python" a secas suele ser el atajo falso de la Microsoft Store,
rem     asi que se prueba SIEMPRE despues de "py".
rem =======================================================================
set EJECUTAR=

py -3-32 -c "import pywinauto" >nul 2>nul
if not errorlevel 1 set EJECUTAR=py -3-32
if not "%EJECUTAR%"=="" goto :tengo_python

py -c "import pywinauto" >nul 2>nul
if not errorlevel 1 set EJECUTAR=py
if not "%EJECUTAR%"=="" goto :tengo_python

python -c "import pywinauto" >nul 2>nul
if not errorlevel 1 set EJECUTAR=python
if not "%EJECUTAR%"=="" goto :tengo_python

rem Ninguno tiene pywinauto todavia: se usa el Python que exista y el
rem programa se encarga de explicar que falta instalarlo.
py -c "import sys" >nul 2>nul
if not errorlevel 1 set EJECUTAR=py
if not "%EJECUTAR%"=="" goto :falta_pywinauto

python -c "import sys" >nul 2>nul
if not errorlevel 1 set EJECUTAR=python
if not "%EJECUTAR%"=="" goto :falta_pywinauto

echo.
echo   NO ENCUENTRO PYTHON EN ESTE COMPUTADOR.
echo.
echo   Instalalo desde python.org y marca la casilla
echo   "Add python.exe to PATH" durante la instalacion.
echo.
pause
exit /b 1

:falta_pywinauto
echo.
echo   FALTA INSTALAR pywinauto, que es lo que mueve las ventanas.
echo   Escribe esto en esta misma ventana y despues vuelve a empezar:
echo.
echo       %EJECUTAR% -m pip install pywinauto
echo.
pause
exit /b 1

:tengo_python
echo   Usando: %EJECUTAR%
echo.

%EJECUTAR% "%PROGRAMA%" %*

echo.
echo ========================================================================
echo   La ventana se queda abierta para que puedas leer lo de arriba.
echo ========================================================================
pause
