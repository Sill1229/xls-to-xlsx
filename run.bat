@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

set "SCRIPT=%~dp0xls_to_xlsx.py"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PY=py -3"
) else (
    set "PY=python"
)

if "%~1"=="" (
    %PY% "%SCRIPT%"
) else (
    %PY% "%SCRIPT%" "%~1"
)

echo.
pause
