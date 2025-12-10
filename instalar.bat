@echo off
echo ========================================
echo Instalador de Aplicacion de Firmas PDF
echo ========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado.
    echo Por favor, instale Python desde python.org
    pause
    exit /b 1
)

echo Python detectado correctamente.
echo.

REM Crear entorno virtual
echo Creando entorno virtual...
python -m venv venv
if errorlevel 1 (
    echo ERROR: No se pudo crear el entorno virtual.
    pause
    exit /b 1
)

echo Entorno virtual creado.
echo.

REM Activar entorno virtual e instalar dependencias
echo Activando entorno virtual e instalando dependencias...
call venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt

if errorlevel 1 (
    echo ERROR: No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Instalacion completada exitosamente!
echo ========================================
echo.
echo Para ejecutar la aplicacion:
echo   1. Activa el entorno virtual: venv\Scripts\activate
echo   2. Ejecuta: python insertar_firmas_pdf.py
echo.
echo O simplemente ejecuta: ejecutar.bat
echo.
pause
