@echo off
REM ==========================================================================
REM  build_exe.bat - Genera Sistema_POA_DPDI.exe (ejecutable independiente, sin consola)
REM  Requisitos: Python 3.11+ para Windows en el PATH (con tcl/tk, opcion por defecto del instalador).
REM  Los datos NO se pierden: la base esta en %APPDATA%\Sistema_POA_DPDI\poa_dpdi.sqlite3
REM  (Sin redirecciones a "nul": en algunos equipos ese dispositivo no esta disponible.)
REM ==========================================================================
setlocal
cd /d "%~dp0" || goto :error

python "%~dp0preparar_build.py" || goto :error

if not exist ".venv\Scripts\activate.bat" (
    echo [1/5] Creando entorno virtual .venv ...
    python -m venv .venv || goto :error
)
call ".venv\Scripts\activate.bat" || goto :error

echo [2/5] Instalando dependencias ...
python -m pip install --upgrade pip || goto :error
pip install -r requirements.txt pyinstaller || goto :error

echo [3/5] Ejecutando pruebas automaticas ...
python -m unittest discover -s tests || goto :error

echo [4/5] Empaquetando con PyInstaller ...
pyinstaller --noconfirm --clean --onefile --windowed --name Sistema_POA_DPDI --collect-all customtkinter --collect-data reportlab --hidden-import openpyxl --hidden-import pandas main_app.py || goto :error

echo.
echo [5/5] Ejecutable generado: %~dp0dist\Sistema_POA_DPDI.exe
echo La base de datos se conserva en %%APPDATA%%\Sistema_POA_DPDI\poa_dpdi.sqlite3
pause
exit /b 0

:error
echo.
echo *** Error durante la construccion. Revise los mensajes anteriores. ***
pause
exit /b 1
