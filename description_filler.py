import pymupdf as pdf
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

# Paste text at coordinates (1000, 720)

class PDFEditor:
    # Run when app starts
    def __init__(self, window):
        self.window = window
        self.window.title("Description Filler")
        self.window.geometry("300x400")
        
        # Define self class variables
        self.file_path = None
        self.output_path = None

        self.file_label = tk.Label(
            window,
            text = "No file loaded",
            font = ("Arial", 10),
            fg = "gray"
        )
        self.file_label.pack(pady=10)

        import_button = tk.Button(
            window,
            text = "Import PDF",
            command = self.import_file,
            width = 20,
            height = 2
        )
        import_button.pack(pady=20)

        run_button = tk.Button(
            window,
            text = "Run",
            command = self.process_pdf,
            width = 5,
            height = 1
        )
        run_button.pack(pady=2)

    def load_descriptions(self):

        titles = []

        doc = pdf.open(self.file_path)
        title_page = doc[0]

        tables = title_page.find_tables()

        if tables:
            target_table = None
            
            for table in tables:
                data = table.extract()
                if data and len(data[0]) >= 2:
                    if 'TITLE' in data[0]:
                        target_table = table
                        break
            if target_table:
                title_data = target_table.extract()
            
            for row in title_data[1:]:
                if len(row) >=2:
                    title = row[1].strip()
                    if title:
                        titles.append(title)

        doc.close()
        return titles

        
    def import_file(self):
        self.file_path = filedialog.askopenfilename(
            title = "Select PDF File",
            filetypes = [("PDF files", "*.pdf")]
        )
        if self.file_path:
            filename = os.path.basename(self.file_path)
            self.file_label.config(text=f"Loaded: {filename}", fg = "green")

        else:
            self.file_label.config(text = "No file loaded", fg = "gray")

    def process_pdf(self):

        titles = self.load_descriptions()
        self.output_path = filedialog.askdirectory(
            title = "Select Output Path"
        )

        doc = pdf.open(self.file_path)
        
        num = 0

        for page in doc:
            text = titles[num]
            page.insert_text(
                (1000, 730),
                text,
                fontname = "helv",
                fontsize = 14,
                color = (0,0,0)
            )
            num += 1

        og_filename = os.path.basename(self.file_path)
        new_filename = og_filename.replace('.pdf','_updated.pdf')

        save_path = os.path.join(self.output_path, new_filename )

        doc.save(save_path)
        doc.close()

if __name__ == "__main__":
    window = tk.Tk()
    app = PDFEditor(window)
    window.mainloop()
