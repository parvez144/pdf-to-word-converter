import os
import pdfplumber
import pytesseract
from PIL import Image, ImageEnhance
from docx import Document
from docx.shared import Pt

# Set Tesseract OCR Path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def pdf_to_word_with_ocr(pdf_path):
    """convert pdf file to DOCX same name"""
    
    # Separate directory and filename from PDF file name
    dir_name = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    docx_path = os.path.join(dir_name, f"{base_name}.docx")
    
    # If the previous file exists, delete it.
    if os.path.exists(docx_path):
        os.remove(docx_path)

    doc = Document()
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            if not text:
                print(f"🔹 Page {i+1}: OCR Using...")
                img = page.to_image(resolution=400).annotated
                img_path = f"temp_page_{i+1}.png"
                img.save(img_path)
                
                try:
                    processed_img = Image.open(img_path).convert('L')
                    text = pytesseract.image_to_string(processed_img, lang='eng+ben')
                except Exception as e:
                    print(f"⚠️ OCR Failed: {e}")
                    text = f"[Page {i+1}: Text could not be read by OCR.]"
                finally:
                    if os.path.exists(img_path):
                        os.remove(img_path)
            
            doc.add_paragraph(text)
            if i < len(pdf.pages) - 1:
                doc.add_page_break()

    doc.save(docx_path)
    print(f"✅ Conversion complete! File saved: {docx_path}")

# Usage examples
if __name__ == "__main__":
    pdf_file = r"Add your pdf path"
    pdf_to_word_with_ocr(pdf_file)
