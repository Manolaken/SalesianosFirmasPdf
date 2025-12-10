@echo off
echo Iniciando aplicacion de firmas PDF...
echo.

REM Verificar si existe el entorno virtual
if not exist "venv\Scripts\activate.bat" (
    echo ERROR: El entorno virtual no existe.
    echo Por favor, ejecuta primero: instalar.bat
    pause
    exit /b 1
)

REM Activar entorno virtual y ejecutar aplicación
call venv\Scripts\activate.bat
python insertar_firmas_pdf.py

pause
