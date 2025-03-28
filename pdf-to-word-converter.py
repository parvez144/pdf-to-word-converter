import os
import pdfplumber
import pytesseract
from tkinter import Tk, filedialog
from PIL import Image
from docx import Document

# Set Tesseract OCR Path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def select_pdf_file():
    """Open a file dialog to select a PDF file."""
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window
    file_path = filedialog.askopenfilename(title="Select a PDF File", filetypes=[("PDF Files", "*.pdf")])
    return file_path

def pdf_to_word_with_ocr(pdf_path):
    """Convert a PDF file to DOCX with OCR support."""
    
    if not pdf_path:
        print("❌ No file selected!")
        return
    
    # Extract directory and filename
    dir_name = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    docx_path = os.path.join(dir_name, f"{base_name}.docx")
    
    # Delete existing file
    if os.path.exists(docx_path):
        os.remove(docx_path)

    doc = Document()
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            if not text:
                print(f"🔹 Page {i+1}: Using OCR...")
                img = page.to_image(resolution=400).annotated
                img_path = f"temp_page_{i+1}.png"
                img.save(img_path)
                
                try:
                    processed_img = Image.open(img_path).convert('L')
                    text = pytesseract.image_to_string(processed_img, lang='eng+ben')
                except Exception as e:
                    print(f"⚠️ OCR Failed: {e}")
                    text = f"[Page {i+1}: OCR couldn't extract text.]"
                finally:
                    if os.path.exists(img_path):
                        os.remove(img_path)
            
            doc.add_paragraph(text)
            if i < len(pdf.pages) - 1:
                doc.add_page_break()

    doc.save(docx_path)
    print(f"✅ Conversion complete! File saved at: {docx_path}")

# Run the script
if __name__ == "__main__":
    pdf_file = select_pdf_file()
    pdf_to_word_with_ocr(pdf_file)
