@echo off
rem Запуск программы двойным щелчком: поднимает интерфейс и открывает его в браузере.
cd /d "%~dp0"
title Претензии и иски - ООО Ромашка

where python >nul 2>&1
if %errorlevel%==0 (
    echo Запускаю интерфейс, откроется вкладка в браузере...
    python web.py
    if errorlevel 1 pause
    exit /b
)

echo Python не найден. Установите Python с python.org и повторите запуск.
pause
