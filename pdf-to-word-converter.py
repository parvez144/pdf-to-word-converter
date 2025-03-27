import os
import pdfplumber
import pytesseract
from PIL import Image, ImageEnhance
from docx import Document
from docx.shared import Pt

# Tesseract OCR পাথ সেট করুন
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def pdf_to_word_with_ocr(pdf_path):
    """PDF ফাইলকে একই নামে DOCX-এ কনভার্ট করুন"""
    
    # PDF ফাইলের নাম থেকে ডিরেক্টরি ও ফাইলনেম আলাদা করুন
    dir_name = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    docx_path = os.path.join(dir_name, f"{base_name}.docx")
    
    # যদি আগের ফাইল থাকে, ডিলিট করুন
    if os.path.exists(docx_path):
        os.remove(docx_path)

    doc = Document()
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            
            if not text:
                print(f"🔹 পৃষ্ঠা {i+1}: OCR ব্যবহার করা হচ্ছে...")
                img = page.to_image(resolution=400).annotated
                img_path = f"temp_page_{i+1}.png"
                img.save(img_path)
                
                try:
                    processed_img = Image.open(img_path).convert('L')
                    text = pytesseract.image_to_string(processed_img, lang='eng+ben')
                except Exception as e:
                    print(f"⚠️ OCR ব্যর্থ: {e}")
                    text = f"[পৃষ্ঠা {i+1}: OCR দ্বারা পাঠ্য পড়া যায়নি]"
                finally:
                    if os.path.exists(img_path):
                        os.remove(img_path)
            
            doc.add_paragraph(text)
            if i < len(pdf.pages) - 1:
                doc.add_page_break()

    doc.save(docx_path)
    print(f"✅ রূপান্তর সম্পূর্ণ! ফাইল সেভ করা হয়েছে: {docx_path}")

# ব্যবহারের উদাহরণ
if __name__ == "__main__":
    pdf_file = r"C:\Users\IT-PC\Desktop\Py Profect\PDF to word\FINAL DRAFT SECURITY SERVICES.pdf"
    pdf_to_word_with_ocr(pdf_file)