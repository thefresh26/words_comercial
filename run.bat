@echo off
setlocal
cd /d "%~dp0"
title Informe subasta Polybid
set PYTHONUNBUFFERED=1
set PYTHONIOENCODING=utf-8
python -u run.py %*
if errorlevel 1 (
  echo.
  echo La aplicacion no pudo iniciarse. Revisa el mensaje de arriba.
  pause
)
