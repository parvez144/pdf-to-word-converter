@echo off
title PDF to Word Converter - Build Process
echo ========================================
echo    PDF to Word Converter - Build Process
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

echo Python found. Starting build process...
echo.

REM Run the build script
python build_installer.py

if errorlevel 1 (
    echo.
    echo Build process failed!
    pause
    exit /b 1
)

echo.
echo Build completed successfully!
echo.
echo Files created:
echo - dist/PDF to Word Converter.exe
echo - installer.bat
echo - PDF to Word Converter - Portable/
echo.
pause 