@echo off
chcp 65001 >nul
title Dashboard Generator

if "%~1"=="" (
    echo.
    echo  Arrastra tu archivo Excel o CSV encima de este icono.
    echo  Formatos soportados: .xlsx .xls .csv
    echo.
    pause
    exit /b 0
)

"%~dp0generar_dashboard.exe" "%~1"
pause