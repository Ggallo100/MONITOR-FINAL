@echo off
REM ==========================================================================
REM  iniciar_servidor_web.bat - Inicia la version web multiusuario del Sistema POA - DPDI
REM  en un servidor Windows de la red institucional. Los usuarios ingresan desde su navegador:
REM      http://NOMBRE-DEL-SERVIDOR:8080
REM  Datos (base, logo, clave de sesion, respaldos): carpeta datos_web junto a este archivo.
REM ==========================================================================
setlocal
cd /d "%~dp0" || goto :error
if "%POA_DATA_DIR%"=="" set "POA_DATA_DIR=%~dp0datos_web"
if "%POA_PORT%"=="" set "POA_PORT=8080"

if not exist ".venv\Scripts\activate.bat" (
    echo Creando entorno virtual .venv ...
    python -m venv .venv || goto :error
)
call ".venv\Scripts\activate.bat" || goto :error
echo Instalando dependencias ...
pip install -r requirements-web.txt || goto :error

echo.
echo Sistema POA - DPDI web: http://%COMPUTERNAME%:%POA_PORT%
echo Datos en: %POA_DATA_DIR%
echo La primera vez, la clave del usuario admin queda en %POA_DATA_DIR%\credenciales_iniciales.txt
echo Para detener el servidor cierre esta ventana o pulse Ctrl+C.
echo.
python servidor_web.py --port %POA_PORT% || goto :error
exit /b 0

:error
echo.
echo *** No se pudo iniciar el servidor. Revise los mensajes anteriores. ***
pause
exit /b 1
