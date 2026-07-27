@echo off
rem ===========================================================
rem  CASCADAS HOTEL - AUDITORIA DIARIA
rem  Doble clic aqui para abrir la herramienta.
rem ===========================================================
title Auditoria Cascadas
cd /d "%~dp0"

python auditoria.py
if errorlevel 1 (
  echo.
  echo ==========================================================
  echo  La herramienta se cerro con un error.
  echo  Copia el texto de arriba y mandalo para revisarlo.
  echo ==========================================================
  echo.
  pause
)
