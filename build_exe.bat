@echo off
title Lebao QA - Label Printer Build Script
echo ============================================
echo   Lebao QA - Label Printer Tool v2.0
echo   Build EXE Script
echo ============================================
echo.
echo   This script auto-detects a compatible Python
echo   (3.8-3.13) and runs build.py
echo.

:: Detect Windows version - Win7 needs Python 3.8
ver | find "6.1" >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Windows 7 detected - will try Python 3.8 first
    echo.
    :: Try Python 3.8 first on Win7 (last version supporting Win7)
    py -3.8 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [OK] Using Python 3.8 (supports Windows 7!)
        py -3.8 build.py
        goto :end
    )
    :: Try Python 3.9 on Win7 (also works with some DLLs)
    py -3.9 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [WARNING] Python 3.8 not found, using 3.9 - build may need extra DLLs on Win7
        py -3.9 build.py
        goto :end
    )
    :: Try Python 3.10 on Win7 (may need api-ms-win-core-path DLL)
    py -3.10 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [WARNING] Using Python 3.10 - build will NOT run on other Win7 machines!
        py -3.10 build.py
        goto :end
    )
    :: Try any newer Python (build will NOT support Win7)
    py -3.11 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [WARNING] Using Python 3.11 - EXE will NOT run on Win7!
        py -3.11 build.py
        goto :end
    )
    py -3.12 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [WARNING] Using Python 3.12 - EXE will NOT run on Win7!
        py -3.12 build.py
        goto :end
    )
    py -3.13 -c "1" >nul 2>&1
    if not errorlevel 1 (
        echo [WARNING] Using Python 3.13 - EXE will NOT run on Win7!
        py -3.13 build.py
        goto :end
    )
    echo [ERROR] No Python found! On Win7, install Python 3.8 from python.org
    pause
    exit /b 1
)

:: Win8+ detected - try latest Python first
echo [INFO] Windows 8+ detected - trying latest Python first
echo.

py -3.13 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.13
    py -3.13 build.py
    goto :end
)

py -3.12 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.12
    py -3.12 build.py
    goto :end
)

py -3.11 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.11
    py -3.11 build.py
    goto :end
)

py -3.10 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.10
    py -3.10 build.py
    goto :end
)

py -3.9 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.9
    py -3.9 build.py
    goto :end
)

py -3.8 -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Using Python 3.8 (also supports Windows 7)
    py -3.8 build.py
    goto :end
)

:: Fallback to default python
python -c "1" >nul 2>&1
if not errorlevel 1 (
    echo [WARNING] Using default Python - version unknown
    python build.py
    goto :end
)

echo [ERROR] No Python found!
echo Install Python 3.8 (for Win7) or 3.13 (for Win8+) from python.org
pause
exit /b 1

:end
pause
