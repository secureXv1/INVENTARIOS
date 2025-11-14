@echo off
setlocal
cd /d "%~dp0"
call .\.venv\Scripts\activate
if errorlevel 1 (
  python -m venv .venv
  call .\.venv\Scripts\activate
  python -m pip install --upgrade pip
  pip install pyinstaller pandas openpyxl
)
pyinstaller --onefile --noconsole --collect-all pandas --collect-all openpyxl --name "ActualizadorInventario" actualizador_inventario_gui.py
echo.
echo Listo: dist\ActualizadorInventario.exe
pause
