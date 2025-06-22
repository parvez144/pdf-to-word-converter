# PDF to Word Converter

A professional Python application that converts PDF files to Word documents (DOCX format) with OCR support for scanned PDFs. Features a modern dark-themed GUI with multiple conversion modes.

![PDF to Word Converter](https://img.shields.io/badge/Python-3.7+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

## ✨ Features

- **🎨 Professional Dark Theme GUI** - Modern, user-friendly interface
- **📄 Multiple Conversion Modes**:
  - **Auto (Best Quality)** - Automatically chooses the best method
  - **Text-based PDF only** - Fast conversion for documents with selectable text
  - **Scanned PDF with OCR** - For scanned documents and images
- **🔍 OCR Support** - Uses Tesseract for scanned PDFs
- **📁 Smart Output** - Saves converted files in the same directory as input
- **⚡ Background Processing** - GUI stays responsive during conversion
- **🛡️ Error Handling** - Robust validation and user-friendly error messages
- **📊 Progress Tracking** - Real-time conversion progress
- **🔄 File Overwrite Protection** - Asks before overwriting existing files

## 🚀 Quick Start

### Prerequisites

1. **Python 3.7 or higher**
2. **Tesseract OCR** (for scanned PDF support)

### Installation

1. **Clone or download** this repository
2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Install Tesseract OCR** (optional but recommended):
   - **Windows**: Download from [Tesseract Releases](https://github.com/UB-Mannheim/tesseract/wiki)
   - **macOS**: `brew install tesseract`
   - **Linux**: `sudo apt-get install tesseract-ocr`

### Usage

Run the application:
```bash
python pdf_to_word_gui_pro.py
```

## 📖 Detailed Usage Guide

### 1. Launch the Application
- Run `python pdf_to_word_gui_pro.py`
- The GUI will open with a dark theme interface

### 2. Select PDF File
- Click **"Browse"** to select your PDF file
- The application will show where the output Word file will be saved

### 3. Choose Conversion Mode
- **Auto (Best Quality)** - Recommended for most cases
- **Text-based PDF only** - Faster for documents with selectable text
- **Scanned PDF with OCR** - For scanned documents

### 4. Convert
- Click **"Convert to Word"**
- Watch the progress bar during conversion
- Get notified when conversion completes

### 5. Access Results
- The Word file will be saved in the same folder as your PDF
- Option to open the output folder directly

## 🔧 Conversion Modes Explained

### Auto (Best Quality)
- **Best for**: Most PDF files
- **Process**: Tries text extraction first, falls back to OCR if needed
- **Speed**: Medium
- **Quality**: Highest

### Text-based PDF only
- **Best for**: PDFs with selectable text
- **Process**: Uses direct text extraction only
- **Speed**: Fastest
- **Quality**: High (preserves formatting)

### Scanned PDF with OCR
- **Best for**: Scanned documents, images, handwritten text
- **Process**: Uses OCR to extract text from images
- **Speed**: Slowest
- **Quality**: Good (depends on image quality)

## 📁 File Structure

```
pdf-to-word-converter/
├── pdf_to_word_gui_pro.py    # Main GUI application (RECOMMENDED)
├── pdf_to_word_allinone.py   # Alternative single-file version
├── requirements.txt          # Python dependencies
├── README.md                # This file
└── spk.ico                  # Application icon (optional)
```

## 🛠️ Installation Details

### Windows Installation

1. **Install Python**:
   - Download from [python.org](https://python.org)
   - Ensure "Add Python to PATH" is checked during installation

2. **Install Dependencies**:
   ```cmd
   pip install -r requirements.txt
   ```

3. **Install Tesseract** (for OCR):
   - Download from [Tesseract Releases](https://github.com/UB-Mannheim/tesseract/wiki)
   - Install to default location: `C:\Program Files\Tesseract-OCR\`

### macOS Installation

1. **Install Python** (if not already installed):
   ```bash
   brew install python
   ```

2. **Install Dependencies**:
   ```bash
   pip3 install -r requirements.txt
   ```

3. **Install Tesseract**:
   ```bash
   brew install tesseract
   ```

### Linux Installation

1. **Install Python** (if not already installed):
   ```bash
   sudo apt-get update
   sudo apt-get install python3 python3-pip
   ```

2. **Install Dependencies**:
   ```bash
   pip3 install -r requirements.txt
   ```

3. **Install Tesseract**:
   ```bash
   sudo apt-get install tesseract-ocr
   ```

## 📋 Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `pdfplumber` | ≥0.9.0 | PDF text extraction |
| `pytesseract` | ≥0.3.10 | OCR functionality |
| `pdf2docx` | ≥0.5.6 | Direct PDF to DOCX conversion |
| `Pillow` | ≥9.0.0 | Image processing |
| `python-docx` | ≥0.8.11 | DOCX file creation |

## 🔍 Troubleshooting

### Common Issues

#### 1. **"Tesseract not found" Warning**
- **Solution**: Install Tesseract OCR
- **Windows**: Download from [Tesseract Releases](https://github.com/UB-Mannheim/tesseract/wiki)
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt-get install tesseract-ocr`

#### 2. **"Module not found" Errors**
- **Solution**: Install dependencies
  ```bash
  pip install -r requirements.txt
  ```

#### 3. **GUI Not Opening**
- **Solution**: Check Python installation
  ```bash
  python --version
  ```
- **Alternative**: Try `python3` instead of `python`

#### 4. **Conversion Fails**
- **Check**: File is a valid PDF
- **Check**: File is not corrupted
- **Check**: Sufficient disk space
- **Try**: Different conversion mode

#### 5. **OCR Quality Issues**
- **Solution**: Ensure scanned PDF has good image quality (300+ DPI)
- **Solution**: Use "Scanned PDF with OCR" mode for scanned documents

### Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| `Tesseract not found` | OCR not available | Install Tesseract |
| `PDF file not found` | Invalid file path | Check file location |
| `Conversion failed` | Processing error | Try different mode |
| `Permission denied` | File access issue | Check file permissions |

## 🎯 Best Practices

### For Best Results:

1. **Use "Auto" mode** for most PDFs
2. **Ensure good image quality** for scanned documents (300+ DPI)
3. **Close the PDF** in other applications before converting
4. **Use descriptive filenames** for easier organization
5. **Backup important files** before conversion

### File Size Guidelines:

- **Text-based PDFs**: Usually convert quickly
- **Scanned PDFs**: May take longer, depending on page count
- **Large files**: Consider splitting into smaller parts

## 🔄 Updates and Maintenance

### Checking for Updates:
- Monitor this repository for new releases
- Update dependencies periodically:
  ```bash
  pip install --upgrade -r requirements.txt
  ```

### Reporting Issues:
- Check the troubleshooting section first
- Provide error messages and system information
- Include steps to reproduce the issue

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👨‍💻 Development

### Building from Source:
1. Clone the repository
2. Install development dependencies
3. Run the application

### Contributing:
- Fork the repository
- Create a feature branch
- Submit a pull request

## 🙏 Acknowledgments

- **Tesseract OCR** for text recognition capabilities
- **pdf2docx** for direct PDF conversion
- **pdfplumber** for PDF text extraction
- **Python community** for excellent libraries

## 📞 Support

For support and questions:
- Check the troubleshooting section
- Review the documentation
- Open an issue on GitHub

---

**Built with ❤️ for efficient PDF to Word conversion**

*© 2025 | by spk* 
