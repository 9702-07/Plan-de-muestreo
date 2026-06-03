@echo off
title Plan de Muestreo - Puerto 5003
cd /d "C:\Plan de muestreo"

echo Iniciando Plan de Muestreo...
start /B python app.py

timeout /t 3 /nobreak >nul
start http://localhost:5003

echo.
echo Plan de Muestreo corriendo en: http://localhost:5003
echo Cierra esta ventana solo cuando termines de usarlo.
echo.
pause
