@echo off
setlocal
cd /d "%~dp0.."
title Consulta BD - Actas de Alcance
set PYTHONUNBUFFERED=1
set PYTHONIOENCODING=utf-8
python -u scripts\consultar_datos_bd_alcance.py
if errorlevel 1 (
  echo.
  echo Hubo un error. Revisa el mensaje de arriba.
)
echo.
pause
