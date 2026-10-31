@echo off
REM ==========================================================================
REM  build_exe.bat - Genera Sistema_POA_DPDI.exe (ejecutable independiente, sin consola)
REM  Requisitos: Python 3.11+ para Windows en el PATH (con tcl/tk, opcion por defecto del instalador).
REM ==========================================================================
setlocal
cd /d "%~dp0"

if not exist .venv (
    echo [1/4] Creando entorno virtual .venv ...
    python -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat || goto :error

echo [2/4] Instalando dependencias ...
python -m pip install --upgrade pip || goto :error
pip install -r requirements.txt pyinstaller || goto :error

echo [3/4] Ejecutando pruebas automaticas ...
python -m unittest discover -s tests || goto :error

echo [4/4] Empaquetando con PyInstaller ...
pyinstaller --noconfirm --clean --onefile --windowed --name Sistema_POA_DPDI --collect-all customtkinter --collect-data reportlab --hidden-import openpyxl --hidden-import pandas main_app.py || goto :error

echo.
echo Ejecutable generado: %~dp0dist\Sistema_POA_DPDI.exe
echo La base de datos se crea en %%APPDATA%%\Sistema_POA_DPDI\poa_dpdi.sqlite3
pause
exit /b 0

:error
echo.
echo *** Error durante la construccion. Revise los mensajes anteriores. ***
pause
exit /b 1
