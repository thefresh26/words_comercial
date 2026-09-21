@echo off
REM Abre Chrome con depuracion remota habilitada, usando un perfil dedicado
REM para la automatizacion de Clarity. La PRIMERA vez tendras que iniciar
REM sesion normalmente en Clarity dentro de esta ventana (con tu 2FA).
REM Las siguientes veces la sesion ya estara guardada y no te pedira nada.
REM
REM Deja esta ventana de Chrome ABIERTA mientras generas el informe con:
REM     python scripts\generar_informe.py <subasta>

set PERFIL=%USERPROFILE%\ClarityChromeDebug
set PUERTO=9222

echo Abriendo Chrome (perfil dedicado en %PERFIL%, puerto %PUERTO%)...
echo No cierres esta ventana mientras generas el informe.
echo.

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=%PUERTO% ^
    --user-data-dir="%PERFIL%" ^
    "https://clarity.microsoft.com/projects/view/t4qu8mqi5b/dashboard"

if errorlevel 1 (
    echo No se encontro Chrome en la ruta esperada.
    echo Ajusta la ruta de chrome.exe en este archivo si la tienes en otro lugar.
    pause
)
