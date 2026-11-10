@echo off
setlocal
title Payroll Forensic Variance and Integrity Engine - Veritux Consulting Network

:: Switch to script directory
cd /d "%~dp0"

echo ===============================================================================
echo   PAYROLL FORENSIC VARIANCE AND INTEGRITY ENGINE
echo   Veritux Consulting Network - 100%% Continuous Controls Monitoring
echo ===============================================================================
echo.

:: Detect Python
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PY_CMD=python"
    goto :PYTHON_FOUND
)

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PY_CMD=py"
    goto :PYTHON_FOUND
)

echo [ERROR] Python was not found in your system PATH.
echo Please ensure Python 3.9+ is installed and added to PATH.
echo.
pause
exit /b 1

:PYTHON_FOUND
:: Check core libraries
%PY_CMD% -c "import pandas, openpyxl, rich" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] First-time setup: Installing required dependencies...
    %PY_CMD% -m pip install pandas openpyxl rich --quiet
)

:: Launch the Engine
echo Starting engine...
echo.
%PY_CMD% payroll_audit_engine.py %*

:: Keep window open after execution
echo.
echo ===============================================================================
echo   Execution session ended.
echo ===============================================================================
pause
