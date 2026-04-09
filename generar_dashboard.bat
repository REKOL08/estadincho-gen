@echo off
set "EXE=%~dp0generar_dashboard.exe"

if not exist "%EXE%" (
    echo ERROR: No se encontro generar_dashboard.exe en esta carpeta.
    pause
    exit /b 1
)

if "%~1"=="" (
    echo Arrastra tu archivo sobre este .bat o ejecutalo desde CMD:
    echo generar_dashboard.bat mi_archivo.xlsx
    pause
    exit /b 0
)

"%EXE%" "%~1"