import tkinter as tk
from tkinter import filedialog, messagebox
import os
import fitz  # PyMuPDF library
from docx import Document

class PDFToWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF to Word Converter")
        master.geometry("500x300")
        master.configure(bg='#f0f0f0')

        # Style configuration
        self.font_style = ('Arial', 12)
        self.bg_color = '#f0f0f0'
        self.button_color = '#4CAF50'
        self.text_color = '#333333'

        # Create and set up GUI components
        self.create_widgets()

    def create_widgets(self):
        # PDF File Selection
        self.pdf_label = tk.Label(
            self.master, 
            text="Select PDF File to Convert", 
            font=self.font_style, 
            bg=self.bg_color, 
            fg=self.text_color
        )
        self.pdf_label.pack(pady=(20, 10))

        self.select_button = tk.Button(
            self.master, 
            text="Browse PDF", 
            command=self.select_pdf_file, 
            bg=self.button_color, 
            fg='white', 
            font=self.font_style
        )
        self.select_button.pack(pady=10)

        # Selected File Path Display
        self.file_path = tk.StringVar()
        self.path_entry = tk.Entry(
            self.master, 
            textvariable=self.file_path, 
            width=50, 
            font=self.font_style, 
            state='readonly'
        )
        self.path_entry.pack(pady=10)

        # Convert Button
        self.convert_button = tk.Button(
            self.master, 
            text="Convert to Word", 
            command=self.convert_pdf_to_word, 
            bg=self.button_color, 
            fg='white', 
            font=self.font_style, 
            state=tk.DISABLED
        )
        self.convert_button.pack(pady=10)

        # Status Label
        self.status_label = tk.Label(
            self.master, 
            text="", 
            font=self.font_style, 
            bg=self.bg_color, 
            fg='green'
        )
        self.status_label.pack(pady=10)

    def select_pdf_file(self):
        # Open file dialog to select PDF
        file_path = filedialog.askopenfilename(
            title="Select PDF File", 
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.convert_button.config(state=tk.NORMAL)

    def convert_pdf_to_word(self):
        pdf_path = self.file_path.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        try:
            # Generate output Word file path
            docx_path = os.path.splitext(pdf_path)[0] + ".docx"
            
            # Open the PDF
            pdf_document = fitz.open(pdf_path)
            
            # Create a new Word document
            doc = Document()

            # Extract text from each page
            for page in pdf_document:
                text = page.get_text()
                doc.add_paragraph(text)

            # Save the Word document
            doc.save(docx_path)
            
            # Close PDF document
            pdf_document.close()

            # Update status
            self.status_label.config(
                text=f"Conversion successful!\nSaved as {os.path.basename(docx_path)}", 
                fg='green'
            )
            
            # Optional: Open the folder containing the converted file
            os.startfile(os.path.dirname(docx_path))

        except Exception as e:
            messagebox.showerror("Conversion Error", str(e))
            self.status_label.config(
                text="Conversion failed.", 
                fg='red'
            )

def main():
    root = tk.Tk()
    app = PDFToWordConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
