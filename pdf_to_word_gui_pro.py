import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import pdfplumber
import pytesseract
from pdf2docx import Converter
from PIL import Image
from docx import Document
from docx.shared import Pt

class PDFToWordConverter:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        self.validate_tesseract()  # Validate on startup

    def setup_ui(self):
        self.root.title("PDF to Word Converter")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        self.root.configure(bg="#2D2D2D")
        
        # Try to set icon
        try:
            icon_path = self.get_icon_path()
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Warning: Could not set icon: {str(e)}")
            
        # Header
        tk.Label(self.root, text="Select a PDF File:", 
                bg="#2D2D2D", fg="white", font=("Helvetica", 12, "bold")).pack(pady=5)

        # Input field
        self.input_entry = tk.Entry(self.root, bg="#F0F0F0", width=70)
        self.input_entry.pack(padx=10, pady=5)

        # Browse button
        tk.Button(self.root, text="Browse", command=self.browse_file,
                 bg="#008080", fg="white", font=("Inter", 12, "bold")).pack(pady=5)

        # Output path display
        tk.Label(self.root, text="Output will be saved to:",
                bg="#2D2D2D", fg="white", font=("Helvetica", 11, "bold")).pack(pady=5)
        
        self.output_label = tk.Label(self.root, text="", bg="#F0F0F0", fg="black", 
                                   font=("Helvetica", 9), padx=10, pady=5, anchor="w")
        self.output_label.pack(padx=10, pady=5, fill="x")

        # Conversion options
        tk.Label(self.root, text="Conversion Options:",
                bg="#2D2D2D", fg="white", font=("Helvetica", 11, "bold")).pack(pady=5)

        self.conversion_var = tk.StringVar(value="Auto (Best Quality)")
        conversion_options = [
            "Auto (Best Quality)",
            "Text-based PDF only",
            "Scanned PDF with OCR"
        ]
        conversion_dropdown = ttk.Combobox(self.root, textvariable=self.conversion_var, 
                                         values=conversion_options, font=("Helvetica", 11),
                                         state="readonly", width=40)
        conversion_dropdown.pack(pady=5)

        # Info label
        tk.Label(self.root, 
                 text="ℹ️ Auto mode tries text extraction first, then OCR if needed.\nText-based mode is faster for documents with selectable text.\nOCR mode is for scanned documents and images.",
                 bg="white", fg="black", font=("Helvetica", 9, "italic")).pack(pady=10)

        # Convert button
        self.convert_btn = tk.Button(self.root, text="Convert to Word", 
                                   command=self.start_conversion, bg="#2c3e50", 
                                   fg="white", font=("inter", 12, "bold"), 
                                   padx=20, pady=5, state=tk.NORMAL)
        self.convert_btn.pack(pady=15)

        # Progress bar (hidden initially)
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=400)
        
        # Status label
        self.status_label = tk.Label(self.root, text="", bg="#2D2D2D", fg="white")
        self.status_label.pack(pady=5)
        
        # Footer
        footer_label = tk.Label(self.root, text="© 2025 | by spk", 
                              bg="#2D2D2D", font=("inter", 8, "italic"), fg="white")
        footer_label.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)

    def get_icon_path(self):
        """Get path to icon file, works for both development and packaged app"""
        try:
            base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_dir, "spk.ico")
            if not os.path.exists(icon_path):
                # Try alternative paths if needed
                icon_path = os.path.join(os.path.dirname(sys.executable), "spk.ico")
            return icon_path
        except Exception as e:
            print(f"Warning: Could not find icon: {str(e)}")
            return None

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)
            self.update_output_path(file_path)

    def update_output_path(self, input_path):
        """Update the output path display"""
        if input_path and os.path.exists(input_path):
            pdf_file = Path(input_path)
            output_file = pdf_file.with_suffix('.docx')
            self.output_label.config(text=str(output_file))
        else:
            self.output_label.config(text="")

    def validate_tesseract(self):
        """Check if Tesseract OCR is available"""
        try:
            possible_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                "tesseract"
            ]
            
            for path in possible_paths:
                try:
                    pytesseract.pytesseract.tesseract_cmd = path
                    pytesseract.get_tesseract_version()
                    print(f"✅ Tesseract found at: {path}")
                    return True
                except Exception:
                    continue
            
            messagebox.showwarning("Warning", 
                                 "Tesseract OCR not found. OCR functionality will not be available.\n\n"
                                 "Please install Tesseract for scanned PDF support:\n"
                                 "Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki\n"
                                 "macOS: brew install tesseract\n"
                                 "Linux: sudo apt-get install tesseract-ocr")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Tesseract validation failed: {str(e)}")
            return False

    def get_output_path(self, input_path):
        """Generate output path that doesn't overwrite existing files"""
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}.docx"
        counter = 1
        
        while os.path.exists(output_path):
            output_path = f"{base}_{counter}.docx"
            counter += 1
            
        return output_path

    def pdf_to_word_best(self, pdf_path, docx_path, conversion_mode="auto"):
        """Convert PDF to DOCX with specified mode"""
        if not os.path.exists(pdf_path):
            print(f"❌ Error: PDF file not found: {pdf_path}")
            return False
        
        # Remove existing output file
        if os.path.exists(docx_path):
            try:
                os.remove(docx_path)
            except Exception as e:
                print(f"⚠️ Warning: Could not remove existing file: {e}")

        doc = Document()

        # Method 1: Try pdf2docx first (for auto and text-based modes)
        if conversion_mode in ["auto", "text"]:
            print("🔹 Trying pdf2docx...")
            try:
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()
                print("✅ Converted using pdf2docx!")
                return True
            except Exception as e:
                print(f"⚠️ pdf2docx failed: {e}")
                if conversion_mode == "text":
                    return False  # Text mode failed, don't try OCR

        # Method 2: Use pdfplumber + OCR (for auto and ocr modes)
        if conversion_mode in ["auto", "ocr"]:
            print("🔹 Using OCR for scanned PDF...")
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()

                        # If no text found, use OCR
                        if not text:
                            print(f"🔹 Page {i+1}: No text found, using OCR...")
                            
                            # Convert page to image
                            img = page.to_image(resolution=300).annotated
                            img_path = f"temp_page_{i+1}.png"
                            
                            try:
                                img.save(img_path)
                                # Extract text using OCR
                                text = pytesseract.image_to_string(Image.open(img_path))
                                # Clean up temporary file
                                os.remove(img_path)
                            except Exception as e:
                                print(f"⚠️ OCR failed for page {i+1}: {e}")
                                text = f"[OCR failed for page {i+1}]"

                        # Add text to document with formatting
                        if text.strip():  # Only add non-empty paragraphs
                            paragraph = doc.add_paragraph(text.strip())
                            paragraph.style.font.size = Pt(12)

                # Save final DOCX
                doc.save(docx_path)
                print(f"✅ OCR Conversion Successful: {docx_path}")
                return True
                
            except Exception as e:
                print(f"❌ Error during conversion: {e}")
                return False
        
        return False

    def convert_pdf(self, input_path, output_path, conversion_mode):
        """Convert PDF to Word document in a separate thread"""
        try:
            success = self.pdf_to_word_best(input_path, output_path, conversion_mode)
            self.root.after(0, self.conversion_finished, success, output_path)
        except Exception as e:
            self.root.after(0, self.conversion_error, str(e))

    def conversion_finished(self, success, output_path):
        """Handle conversion completion"""
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.config(state=tk.NORMAL)
        
        if success:
            self.status_label.config(text="✅ Conversion completed successfully!")
            result = messagebox.askyesno("Success", 
                                       f"PDF converted successfully!\n\n"
                                       f"Output: {output_path}\n\n"
                                       "Would you like to open the output folder?")
            if result:
                self.open_output_folder(output_path)
        else:
            self.status_label.config(text="❌ Conversion failed!")
            messagebox.showerror("Error", "Failed to convert PDF. Please check the file and try again.")

    def conversion_error(self, error_message):
        """Handle conversion errors"""
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text="❌ Conversion failed!")
        messagebox.showerror("Error", f"An error occurred during conversion:\n\n{error_message}")

    def open_output_folder(self, output_path):
        """Open the output folder in file explorer"""
        if output_path and os.path.exists(output_path):
            folder_path = os.path.dirname(output_path)
            try:
                if sys.platform == "win32":
                    os.startfile(folder_path)
                elif sys.platform == "darwin":  # macOS
                    os.system(f"open '{folder_path}'")
                else:  # Linux
                    os.system(f"xdg-open '{folder_path}'")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder: {e}")

    def start_conversion(self):
        input_path = self.input_entry.get().strip()
        
        if not input_path:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        if not os.path.isfile(input_path):
            messagebox.showwarning("Warning", "The specified file does not exist.")
            return

        if not input_path.lower().endswith('.pdf'):
            messagebox.showwarning("Warning", "Please select a valid PDF file.")
            return

        # Get conversion mode
        conversion_mode_map = {
            "Auto (Best Quality)": "auto",
            "Text-based PDF only": "text", 
            "Scanned PDF with OCR": "ocr"
        }
        conversion_mode = conversion_mode_map.get(self.conversion_var.get(), "auto")

        # Generate output path
        output_path = self.get_output_path(input_path)
        
        # Check if output file already exists
        if os.path.exists(output_path):
            result = messagebox.askyesno("File Exists", 
                                       f"Output file already exists:\n{output_path}\n\n"
                                       "Do you want to overwrite it?")
            if not result:
                return

        # Show progress and start conversion
        self.progress.pack(pady=10)
        self.progress.start()
        self.convert_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Converting PDF...")
        self.root.update()
        
        # Start conversion in separate thread
        thread = threading.Thread(target=self.convert_pdf, args=(input_path, output_path, conversion_mode))
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToWordConverter(root)
    root.mainloop() 