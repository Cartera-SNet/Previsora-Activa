@echo off
:: ============================================================
::  Activa IT — Descargador de Cartas Glosa Previsora
::  Instancia dinámica — ingresa el puerto al ejecutar
:: ============================================================

setlocal

echo.
echo ============================================================
echo   Activa IT -- Descargador de Cartas Glosa Previsora
echo ============================================================
echo.

:: Pedir puerto al usuario
set /p PORT="Ingresa el puerto para esta instancia (ej: 5050): "

:: Validar que se ingresó algo
if "%PORT%"=="" (
    echo [ERROR] No ingresaste un puerto.
    pause
    exit /b 1
)

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"

echo.
echo   PUERTO: %PORT%
echo   Carpeta de descargas: downloads_%PORT%
echo ============================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado.
    pause
    exit /b 1
)

:: Crear entorno virtual si no existe
if not exist "%VENV_DIR%" (
    echo Creando entorno virtual Python...
    python -m venv "%VENV_DIR%"
)

:: Activar entorno virtual
call "%VENV_DIR%\Scripts\activate.bat"

:: Instalar dependencias
echo Instalando dependencias...
pip install --quiet flask playwright openpyxl

:: Instalar Chromium
echo Instalando Playwright Chromium...
python -m playwright install chromium

:: Crear carpeta de descargas independiente
if not exist "%SCRIPT_DIR%downloads_%PORT%" mkdir "%SCRIPT_DIR%downloads_%PORT%"

:: Abrir navegador en segundo plano
start "" "http://localhost:%PORT%"

:: Ejecutar aplicación
python "%SCRIPT_DIR%app.py"

pause