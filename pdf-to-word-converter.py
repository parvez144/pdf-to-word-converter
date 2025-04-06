import os
import pdfplumber
import pytesseract
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from docx import Document
import threading

# Set Tesseract OCR Path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Word Converter")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        self.root.configure(bg="#f0f0f0")
        
        # Set app icon if available
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Style configuration
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10), background="#f0f0f0")
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background="#f0f0f0")
        
        # Header
        header_label = ttk.Label(main_frame, text="PDF to Word Converter", style="Header.TLabel")
        header_label.pack(pady=(0, 20))
        
        # File selection frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        # File path entry
        self.file_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Browse button
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT)
        
        # Language selection
        lang_frame = ttk.Frame(main_frame)
        lang_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(lang_frame, text="OCR Language:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.lang_var = tk.StringVar(value="eng+ben")
        lang_combo = ttk.Combobox(lang_frame, textvariable=self.lang_var, values=["eng", "ben", "eng+ben"], state="readonly", width=10)
        lang_combo.pack(side=tk.LEFT)
        
        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="Convert to Word", command=self.start_conversion)
        self.convert_btn.pack(pady=20)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        
        # Instructions
        instructions = """
        Instructions:
        1. Click 'Browse' to select a PDF file
        2. Choose OCR language if needed
        3. Click 'Convert to Word' to start conversion
        4. The converted file will be saved in the same folder as the PDF
        """
        ttk.Label(main_frame, text=instructions, justify=tk.LEFT, wraplength=550).pack(pady=20)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a PDF File",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            self.status_var.set("File selected")
    
    def start_conversion(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file first!")
            return
        
        # Disable convert button during conversion
        self.convert_btn.config(state=tk.DISABLED)
        
        # Start conversion in a separate thread
        thread = threading.Thread(target=self.convert_pdf)
        thread.daemon = True
        thread.start()
    
    def convert_pdf(self):
        pdf_path = self.pdf_path.get()
        
        if not pdf_path:
            self.status_var.set("❌ No file selected!")
            self.convert_btn.config(state=tk.NORMAL)
            return
        
        # Extract directory and filename
        dir_name = os.path.dirname(pdf_path)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        docx_path = os.path.join(dir_name, f"{base_name}.docx")
        
        # Delete existing file
        if os.path.exists(docx_path):
            os.remove(docx_path)
        
        try:
            doc = Document()
            
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    # Update progress
                    progress = (i / total_pages) * 100
                    self.progress_var.set(progress)
                    self.status_var.set(f"Processing page {i+1} of {total_pages}")
                    self.root.update_idletasks()
                    
                    text = page.extract_text()
                    
                    if not text:
                        self.status_var.set(f"Page {i+1}: Using OCR...")
                        self.root.update_idletasks()
                        
                        img = page.to_image(resolution=400).annotated
                        img_path = f"temp_page_{i+1}.png"
                        img.save(img_path)
                        
                        try:
                            processed_img = Image.open(img_path).convert('L')
                            text = pytesseract.image_to_string(processed_img, lang=self.lang_var.get())
                        except Exception as e:
                            self.status_var.set(f"⚠️ OCR Failed: {e}")
                            text = f"[Page {i+1}: OCR couldn't extract text.]"
                        finally:
                            if os.path.exists(img_path):
                                os.remove(img_path)
                    
                    doc.add_paragraph(text)
                    if i < total_pages - 1:
                        doc.add_page_break()
            
            doc.save(docx_path)
            self.progress_var.set(100)
            self.status_var.set(f"✅ Conversion complete! File saved at: {docx_path}")
            messagebox.showinfo("Success", f"Conversion complete!\nFile saved at:\n{docx_path}")
        
        except Exception as e:
            self.status_var.set(f"❌ Error: {e}")
            messagebox.showerror("Error", f"An error occurred during conversion:\n{str(e)}")
        
        finally:
            # Re-enable convert button
            self.convert_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
