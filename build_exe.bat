@echo off
title Lebao QA - Label Printer Build Script
echo ============================================
echo   Lebao QA - Label Printer Tool v2.0
echo   Build EXE Script
echo ============================================
echo.
echo   This script auto-detects a compatible Python
echo   (3.10-3.13, NOT 3.14) and runs build.py
echo.

:: Try Python 3.13 first
py -3.13 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.13
    py -3.13 build.py
    goto :end
)

:: Try Python 3.12
py -3.12 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.12
    py -3.12 build.py
    goto :end
)

:: Try Python 3.11
py -3.11 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.11
    py -3.11 build.py
    goto :end
)

:: Try Python 3.10
py -3.10 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.10
    py -3.10 build.py
    goto :end
)

:: Fallback to default python
python -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [WARNING] Using default Python - may be 3.14 (unsupported)
    python build.py
    goto :end
)

echo [ERROR] No Python found!
echo Install Python 3.12 or 3.13 from python.org
pause
exit /b 1

:end
pause
