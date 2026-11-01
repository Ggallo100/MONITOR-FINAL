@echo off
REM ==========================================================================
REM  build_exe.bat - Genera Sistema_POA_DPDI.exe (ejecutable independiente, sin consola)
REM  Requisitos: Python 3.11+ para Windows en el PATH (con tcl/tk, opcion por defecto del instalador).
REM  Los datos NO se pierden: la base esta en %APPDATA%\Sistema_POA_DPDI\poa_dpdi.sqlite3
REM ==========================================================================
setlocal
cd /d "%~dp0"

echo [0/5] Cerrando el sistema si esta abierto ...
taskkill /f /im Sistema_POA_DPDI.exe >/dev/null 2>&1
timeout /t 2 /nobreak >nul
if exist "dist\Sistema_POA_DPDI.exe" del /f /q "dist\Sistema_POA_DPDI.exe" >/dev/null 2>&1
if exist "dist\Sistema_POA_DPDI.exe" (
    echo     El ejecutable anterior esta bloqueado; se renombra para poder reemplazarlo.
    move /y "dist\Sistema_POA_DPDI.exe" "dist\Sistema_POA_DPDI_anterior_%RANDOM%.exe" >/dev/null 2>&1
)
if exist "dist\Sistema_POA_DPDI.exe" (
    echo.
    echo *** No se puede reemplazar dist\Sistema_POA_DPDI.exe ***
    echo Cierre el Sistema POA ^(y cualquier ventana o antivirus que lo use^) y vuelva a ejecutar build_exe.bat.
    goto :error
)

if not exist .venv (
    echo [1/5] Creando entorno virtual .venv ...
    python -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat || goto :error

echo [2/5] Instalando dependencias ...
python -m pip install --upgrade pip || goto :error
pip install -r requirements.txt pyinstaller || goto :error

echo [3/5] Ejecutando pruebas automaticas ...
python -m unittest discover -s tests || goto :error

echo [4/5] Empaquetando con PyInstaller ...
pyinstaller --noconfirm --clean --onefile --windowed --name Sistema_POA_DPDI --collect-all customtkinter --collect-data reportlab --hidden-import openpyxl --hidden-import pandas main_app.py || goto :error

echo [5/5] Limpiando ejecutables anteriores renombrados ...
del /f /q "dist\Sistema_POA_DPDI_anterior_*.exe" >/dev/null 2>&1

echo.
echo Ejecutable generado: %~dp0dist\Sistema_POA_DPDI.exe
echo La base de datos se conserva en %%APPDATA%%\Sistema_POA_DPDI\poa_dpdi.sqlite3
pause
exit /b 0

:error
echo.
echo *** Error durante la construccion. Revise los mensajes anteriores. ***
pause
exit /b 1
