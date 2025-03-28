# pdf-to-word-converter
Convert pdf to Word file with python.

# 📄 PDF to Word Converter with OCR

## 📌 Overview
This Python project converts PDF files to Word (.docx) format while preserving both text and images. It supports **scanned PDFs** by extracting text using **Tesseract OCR**.

## 🚀 Features
- Converts text-based PDFs to editable Word format
- Extracts and preserves images in the Word document
- Uses OCR for scanned PDFs
- Maintains document formatting as much as possible

## 🛠 Requirements
Ensure you have the following installed before running the script:

### Python 3

### 📌 Install Dependencies
pip install pdfplumber pytesseract pillow python-docx

### 📌 Install Tesseract OCR (For Scanned PDFs)
- **Windows**: Download and install Tesseract OCR (https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux (Debian/Ubuntu)**:
  sudo apt install tesseract-ocr

## 🔧 Usage
Run the script with the following command:
python pdf_to_word.py
A file dialog box will open then select your desire PDF file.
After conversion you will find your file where was your PDF file.


## 📂 Project Structure
📁 PDF-to-Word-Converter
├── 📄 pdf_to_word.py   # Main script
├── 📄 README.md        # Project documentation

## 🛠 How It Works
1. Loads the input PDF
2. Extracts text directly if available
3. If text is not found, applies OCR to extract text from images
4. Saves the output as a .docx file

## 📜 License
This project is open-source under the **MIT License**.

---

🌟 **Contributions are welcome!** If you have any improvements, feel free to fork this repo and create a pull request. 😊
