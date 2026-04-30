@echo off
setlocal

echo ================================================
echo  PST a Mbox - Build EXE portable
echo ================================================

if not exist config\version_info.txt (
  echo.
  echo ERROR: Falta config\version_info.txt
  echo Crea/edita ese archivo para definir los campos de "Detalles".
  exit /b 1
)

echo.
echo [1/3] Instalando dependencias de compilacion...
python -m pip install --upgrade pyinstaller pywin32
if errorlevel 1 (
    echo.
    echo ERROR: No se pudieron instalar dependencias.
    exit /b 1
)

echo.
echo [2/3] Generando EXE...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name PST2MboxPortable ^
  --version-file config\version_info.txt ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32timezone ^
  --add-data "src;src" ^
  src\pst_to_mbox_gui.py

if errorlevel 1 (
    echo.
    echo ERROR: Fallo la compilacion del EXE.
    exit /b 1
)

echo.
echo [3/3] Listo.
echo EXE generado en: dist\PST2MboxPortable.exe
echo Campos de "Detalles" aplicados desde: config\version_info.txt
echo.
echo Puedes compartir solo ese .exe como version portable.
echo Requisito en la PC destino: Outlook de escritorio instalado.

endlocal
