#!/usr/bin/env python3
"""
Build script for PDF to Word Converter
Creates a standalone executable using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def fix_pathlib_conflict():
    """Fix pathlib conflict with PyInstaller"""
    try:
        import pathlib
        # Check if it's the obsolete backport
        if hasattr(pathlib, '__version__'):
            print("⚠️ Detected obsolete pathlib package. Attempting to fix...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "uninstall", "pathlib", "-y"])
                print("✅ Removed obsolete pathlib package")
                return True
            except subprocess.CalledProcessError:
                print("⚠️ Could not remove pathlib package automatically")
                print("Please run manually: pip uninstall pathlib -y")
                return False
        else:
            print("✅ pathlib is standard library (no conflict)")
            return True
    except ImportError:
        print("✅ No pathlib package found (no conflict)")
        return True

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print("✅ PyInstaller is installed")
        return True
    except ImportError:
        print("❌ PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✅ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("❌ Failed to install PyInstaller")
            return False

def create_spec_file():
    """Create PyInstaller spec file"""
    # Check if icon exists
    icon_path = 'spk.ico'
    icon_line = f"icon='{icon_path}'" if os.path.exists(icon_path) else "icon=None"
    
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['pdf_to_word_gui_pro.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pdfplumber',
        'pytesseract',
        'pdf2docx',
        'PIL',
        'docx',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox'
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF to Word Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    {icon_line},
)
'''
    
    with open('pdf_to_word_converter.spec', 'w') as f:
        f.write(spec_content)
    
    if os.path.exists(icon_path):
        print(f"✅ Spec file created with icon: {icon_path}")
    else:
        print("⚠️ Spec file created without icon (spk.ico not found)")
    
    print("✅ Spec file created: pdf_to_word_converter.spec")

def build_executable():
    """Build the executable using PyInstaller"""
    print("🔨 Building executable...")
    
    # Create spec file
    create_spec_file()
    
    # Build command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        "pdf_to_word_converter.spec"
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✅ Executable built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return False

def create_installer_script():
    """Create a simple installer script"""
    installer_content = '''@echo off
title PDF to Word Converter Installer
echo ========================================
echo    PDF to Word Converter Installer
echo ========================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running as administrator...
) else (
    echo Please run this installer as administrator
    pause
    exit /b 1
)

REM Create installation directory
set INSTALL_DIR=C:\\Program Files\\PDF to Word Converter
echo Installing to: %INSTALL_DIR%

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy files
echo Copying files...
copy "dist\\PDF to Word Converter.exe" "%INSTALL_DIR%\\"
if exist "spk.ico" copy "spk.ico" "%INSTALL_DIR%\\"

REM Create desktop shortcut
echo Creating desktop shortcut...
set DESKTOP=%USERPROFILE%\\Desktop
set SHORTCUT=%DESKTOP%\\PDF to Word Converter.lnk

powershell "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT%'); $Shortcut.TargetPath = '%INSTALL_DIR%\\PDF to Word Converter.exe'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.IconLocation = '%INSTALL_DIR%\\spk.ico'; $Shortcut.Save()"

REM Create start menu shortcut
echo Creating start menu shortcut...
set START_MENU=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs
if not exist "%START_MENU%\\PDF to Word Converter" mkdir "%START_MENU%\\PDF to Word Converter"

set START_SHORTCUT=%START_MENU%\\PDF to Word Converter\\PDF to Word Converter.lnk
powershell "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_SHORTCUT%'); $Shortcut.TargetPath = '%INSTALL_DIR%\\PDF to Word Converter.exe'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.IconLocation = '%INSTALL_DIR%\\spk.ico'; $Shortcut.Save()"

echo.
echo ========================================
echo    Installation Complete!
echo ========================================
echo.
echo PDF to Word Converter has been installed to:
echo %INSTALL_DIR%
echo.
echo Shortcuts created on desktop and start menu.
echo.
echo Note: You may need to install Tesseract OCR
echo for scanned PDF support. Download from:
echo https://github.com/UB-Mannheim/tesseract/wiki
echo.
pause
'''
    
    with open('installer.bat', 'w') as f:
        f.write(installer_content)
    print("✅ Installer script created: installer.bat")

def create_portable_version():
    """Create a portable version"""
    portable_dir = "PDF to Word Converter - Portable"
    
    if os.path.exists(portable_dir):
        shutil.rmtree(portable_dir)
    
    os.makedirs(portable_dir)
    
    # Copy executable
    if os.path.exists("dist/PDF to Word Converter.exe"):
        shutil.copy2("dist/PDF to Word Converter.exe", portable_dir)
        print(f"✅ Copied executable to {portable_dir}")
    else:
        print("❌ Executable not found in dist/ folder")
        return False
    
    # Copy icon if exists (for manual use and troubleshooting)
    if os.path.exists("spk.ico"):
        shutil.copy2("spk.ico", portable_dir)
        print(f"✅ Copied icon file to {portable_dir}")
    
    # Create a batch file to run with icon (alternative method)
    batch_content = '''@echo off
title PDF to Word Converter
echo Starting PDF to Word Converter...
start "" "PDF to Word Converter.exe"
'''
    
    with open(f"{portable_dir}/Run Converter.bat", 'w') as f:
        f.write(batch_content)
    
    # Create README for portable version
    portable_readme = '''PDF to Word Converter - Portable Version

This is a portable version of the PDF to Word Converter.
No installation required - just run the executable!

USAGE:
1. Double-click "PDF to Word Converter.exe"
   OR
2. Double-click "Run Converter.bat"

REQUIREMENTS:
- Windows 10 or later
- Tesseract OCR (optional, for scanned PDFs)
  Download from: https://github.com/UB-Mannheim/tesseract/wiki

FEATURES:
- Convert PDF to Word documents
- Support for scanned PDFs with OCR
- Multiple conversion modes
- Professional dark theme GUI
- No installation required

ICON ISSUES:
If the icon doesn't show in Windows Explorer:
1. Right-click the .exe file
2. Select "Properties"
3. Click "Change Icon"
4. Browse to "spk.ico" in this folder
5. Click OK

© 2025 | by spk
'''
    
    with open(f"{portable_dir}/README.txt", 'w') as f:
        f.write(portable_readme)
    
    print(f"✅ Portable version created: {portable_dir}")
    return True

def main():
    """Main build process"""
    print("🚀 PDF to Word Converter - Build Process")
    print("=" * 50)
    
    # Fix pathlib conflict first
    if not fix_pathlib_conflict():
        print("⚠️ Please manually remove pathlib package and try again:")
        print("   pip uninstall pathlib -y")
        return False
    
    # Check PyInstaller
    if not check_pyinstaller():
        return False
    
    # Build executable
    if not build_executable():
        return False
    
    # Create installer script
    create_installer_script()
    
    # Create portable version
    create_portable_version()
    
    print("\n" + "=" * 50)
    print("🎉 Build completed successfully!")
    print("\nGenerated files:")
    print("- dist/PDF to Word Converter.exe (Main executable)")
    print("- installer.bat (Windows installer)")
    print("- PDF to Word Converter - Portable/ (Portable version)")
    print("\nNext steps:")
    print("1. Test the executable in dist/ folder")
    print("2. Run installer.bat as administrator to install")
    print("3. Or use the portable version")
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n❌ Build process failed!")
        sys.exit(1) 