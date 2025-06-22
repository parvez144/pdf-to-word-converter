@echo off
title PDF to Word Converter - Inno Setup Builder
echo ========================================
echo    PDF to Word Converter - Inno Setup Builder
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

echo Python found. Checking requirements...
echo.

REM Check if executable exists
if not exist "dist\PDF to Word Converter.exe" (
    echo Error: Executable not found!
    echo Please run the PyInstaller build first:
    echo   build.bat
    echo.
    echo Current directory contents:
    dir dist\ 2>nul || echo dist\ folder not found
    pause
    exit /b 1
)

echo ✅ Executable found: dist\PDF to Word Converter.exe
echo.

REM Check if required files exist
if not exist "spk.ico" (
    echo ⚠️ Warning: spk.ico not found. Installer will use default icon.
    echo.
)

if not exist "LICENSE.txt" (
    echo ⚠️ Warning: LICENSE.txt not found. Creating default license...
    copy nul LICENSE.txt >nul 2>&1
)

if not exist "README.txt" (
    echo ⚠️ Warning: README.txt not found. Creating default readme...
    echo PDF to Word Converter > README.txt
)

echo ✅ Required files checked.
echo.

REM Try multiple methods to find Inno Setup
echo Searching for Inno Setup...
set INNO_PATH=

REM Method 1: Check registry for Inno Setup 6
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1" >nul 2>&1
if not errorlevel 1 (
    set INNO_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    echo Found Inno Setup 6 in registry
)

REM Method 2: Check registry for Inno Setup 5
if "%INNO_PATH%"=="" (
    reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 5_is1" >nul 2>&1
    if not errorlevel 1 (
        set INNO_PATH="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
        echo Found Inno Setup 5 in registry
    )
)

REM Method 3: Check common installation paths
if "%INNO_PATH%"=="" (
    if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
        set INNO_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
        echo Found Inno Setup 6 in Program Files
    )
)

if "%INNO_PATH%"=="" (
    if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" (
        set INNO_PATH="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
        echo Found Inno Setup 5 in Program Files
    )
)

REM Method 4: Check 64-bit Program Files
if "%INNO_PATH%"=="" (
    if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
        set INNO_PATH="C:\Program Files\Inno Setup 6\ISCC.exe"
        echo Found Inno Setup 6 in 64-bit Program Files
    )
)

if "%INNO_PATH%"=="" (
    if exist "C:\Program Files\Inno Setup 5\ISCC.exe" (
        set INNO_PATH="C:\Program Files\Inno Setup 5\ISCC.exe"
        echo Found Inno Setup 5 in 64-bit Program Files
    )
)

REM Method 5: Check if ISCC is in PATH
if "%INNO_PATH%"=="" (
    ISCC --version >nul 2>&1
    if not errorlevel 1 (
        set INNO_PATH=ISCC
        echo Found Inno Setup in PATH
    )
)

if "%INNO_PATH%"=="" (
    echo.
    echo ❌ Error: Inno Setup not found!
    echo.
    echo Please install Inno Setup from: https://jrsoftware.org/isinfo.php
    echo.
    echo Common installation paths checked:
    echo • C:\Program Files (x86)\Inno Setup 6\
    echo • C:\Program Files (x86)\Inno Setup 5\
    echo • C:\Program Files\Inno Setup 6\
    echo • C:\Program Files\Inno Setup 5\
    echo • PATH environment variable
    echo.
    pause
    exit /b 1
)

echo ✅ Found Inno Setup at: %INNO_PATH%
echo.

REM Test Inno Setup
echo Testing Inno Setup...
%INNO_PATH% --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Inno Setup test failed!
    echo The path %INNO_PATH% is not working.
    pause
    exit /b 1
)

echo ✅ Inno Setup test successful.
echo.

REM Create Output directory if it doesn't exist
if not exist "Output" (
    mkdir "Output"
    echo Created Output directory
)

REM Check if .iss file exists
if not exist "pdf_to_word_setup.iss" (
    echo ❌ Error: pdf_to_word_setup.iss not found!
    echo Please ensure the Inno Setup script exists.
    pause
    exit /b 1
)

echo ✅ Inno Setup script found: pdf_to_word_setup.iss
echo.

REM Build the installer
echo ========================================
echo Building Inno Setup installer...
echo ========================================
echo.

%INNO_PATH% "pdf_to_word_setup.iss"

if errorlevel 1 (
    echo.
    echo ❌ Inno Setup build failed!
    echo.
    echo Common issues:
    echo • Missing required files (spk.ico, LICENSE.txt, README.txt)
    echo • Syntax error in .iss file
    echo • Permission issues
    echo.
    echo Check the error messages above for details.
    pause
    exit /b 1
)

echo.
echo ========================================
echo    Build Completed Successfully!
echo ========================================
echo.

REM Check if installer was created
if exist "Output\PDF_to_Word_Converter_Setup.exe" (
    echo ✅ Installer created: Output\PDF_to_Word_Converter_Setup.exe
    
    REM Get file size
    for %%A in ("Output\PDF_to_Word_Converter_Setup.exe") do (
        echo 📦 Installer size: %%~zA bytes
    )
) else (
    echo ⚠️ Warning: Installer file not found in Output directory
    echo Checking Output directory contents:
    dir Output\ 2>nul || echo Output directory is empty
)

echo.
echo Features included:
echo • Professional installation wizard
echo • Desktop and start menu shortcuts
echo • File associations for PDF files
echo • Uninstaller
echo • License and information pages
echo.
echo You can now distribute the installer!
echo.
pause 