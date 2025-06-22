# Build Guide - PDF to Word Converter

This guide explains how to create a standalone executable and installer for the PDF to Word Converter using PyInstaller.

## 🚀 Quick Build

### Option 1: Automatic Build (Recommended)
```bash
# Double-click or run:
build.bat
```

### Option 2: Manual Build
```bash
# Install PyInstaller
pip install pyinstaller

# Run build script
python build_installer.py
```

## 📋 Prerequisites

Before building, ensure you have:

1. **Python 3.7+** installed
2. **All dependencies** installed:
   ```bash
   pip install -r requirements.txt
   ```
3. **Your application files** ready:
   - `pdf_to_word_gui_pro.py` (main application)
   - `spk.ico` (optional icon file)

## 🔨 Build Process

### Step 1: Prepare Files
Ensure these files are in your project directory:
```
project/
├── pdf_to_word_gui_pro.py    # Main application
├── build_installer.py        # Build script
├── requirements.txt          # Dependencies
├── spk.ico                  # Icon (optional)
└── README.md                # Documentation
```

### Step 2: Run Build
```bash
python build_installer.py
```

### Step 3: Check Output
The build process creates:
```
project/
├── dist/
│   └── PDF to Word Converter.exe    # Standalone executable
├── installer.bat                    # Windows installer
├── PDF to Word Converter - Portable/ # Portable version
└── build/                          # Build cache (can be deleted)
```

## 📦 What Gets Created

### 1. Standalone Executable
- **Location**: `dist/PDF to Word Converter.exe`
- **Size**: ~50-100 MB (includes all dependencies)
- **Features**: 
  - No Python installation required
  - All libraries bundled
  - Windows executable

### 2. Windows Installer
- **Location**: `installer.bat`
- **Features**:
  - Installs to `C:\Program Files\PDF to Word Converter\`
  - Creates desktop shortcut
  - Creates start menu entry
  - Requires administrator privileges

### 3. Portable Version
- **Location**: `PDF to Word Converter - Portable/`
- **Features**:
  - No installation required
  - Can run from USB drive
  - Includes README with instructions

## 🎯 Installation Options

### For End Users

#### Option 1: Full Installation
1. Run `installer.bat` as administrator
2. Application installs to Program Files
3. Shortcuts created on desktop and start menu

#### Option 2: Portable Version
1. Copy `PDF to Word Converter - Portable/` folder
2. Run `PDF to Word Converter.exe` directly
3. No installation required

#### Option 3: Direct Executable
1. Copy `dist/PDF to Word Converter.exe`
2. Run directly (no installation)

## 🔧 Build Configuration

### PyInstaller Spec File
The build script creates `pdf_to_word_converter.spec` with:

```python
# Key settings:
- console=False          # No console window
- icon='spk.ico'         # Custom icon
- name='PDF to Word Converter'  # Executable name
- upx=True              # Compress executable
```

### Hidden Imports
The build includes these libraries:
- `pdfplumber` - PDF text extraction
- `pytesseract` - OCR functionality
- `pdf2docx` - Direct PDF conversion
- `PIL` - Image processing
- `docx` - Word document creation
- `tkinter` - GUI framework

## 🐛 Troubleshooting

### Common Build Issues

#### 1. **"Module not found" errors**
```bash
# Solution: Install missing dependencies
pip install -r requirements.txt
```

#### 2. **PyInstaller not found**
```bash
# Solution: Install PyInstaller
pip install pyinstaller
```

#### 3. **Large executable size**
- This is normal (50-100 MB)
- Includes all Python libraries
- No external dependencies needed

#### 4. **Executable doesn't run**
- Check Windows Defender/antivirus
- Try running as administrator
- Check if all files are present

#### 5. **Missing icon**
- Ensure `spk.ico` exists in project directory
- Or remove icon reference from spec file

### Build Optimization

#### Reduce Size
```python
# In spec file, add to excludes:
excludes=['matplotlib', 'numpy', 'scipy']
```

#### Add UPX Compression
```bash
# Install UPX for better compression
# Download from: https://upx.github.io/
```

## 📊 Build Statistics

### Typical Build Times
- **First build**: 2-5 minutes
- **Subsequent builds**: 30-60 seconds
- **Clean build**: 1-2 minutes

### Executable Sizes
- **Basic**: ~50 MB
- **With all features**: ~80-100 MB
- **Compressed**: ~30-50 MB

## 🔄 Updating the Build

### After Code Changes
1. Modify your Python files
2. Run `python build_installer.py`
3. New executable will be created

### Updating Dependencies
1. Update `requirements.txt`
2. Install new dependencies: `pip install -r requirements.txt`
3. Rebuild: `python build_installer.py`

## 📤 Distribution

### For Distribution
1. **Full installer**: `installer.bat`
2. **Portable version**: `PDF to Word Converter - Portable/` folder
3. **Direct executable**: `dist/PDF to Word Converter.exe`

### File Structure for Distribution
```
PDF to Word Converter/
├── PDF to Word Converter.exe
├── spk.ico (optional)
└── README.txt
```

## 🎉 Success Checklist

After building, verify:

- [ ] Executable runs without errors
- [ ] GUI opens correctly
- [ ] PDF conversion works
- [ ] OCR functionality works (if Tesseract installed)
- [ ] Installer creates shortcuts
- [ ] Portable version works on different computers

## 📞 Support

If you encounter build issues:

1. Check the troubleshooting section
2. Ensure all dependencies are installed
3. Try a clean build (delete `build/` and `dist/` folders)
4. Check Python and PyInstaller versions

---

**Happy Building! 🚀**

*© 2025 | by spk* 