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
        self.excel_path = None
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

        self.excel_label = tk.Label(
            window,
            text = "No Title Sheet Loaded",
            font = ("Arial, 10"),
            fg = "gray"
        )
        self.excel_label.pack(pady=10)

        load_excel_button = tk.Button(
            window,
            text = "Load Title Sheet",
            command = self.excel_import,
            width = 20,
            height = 2,
        )
        load_excel_button.pack(pady = 20)

        run_button = tk.Button(
            window,
            text = "Run",
            command = self.process_pdf,
            width = 5,
            height = 1
        )
        run_button.pack(pady=2)

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

    def excel_import(self):
        self.excel_path = filedialog.askopenfilename(
            title = "Select Excel File",
            filetypes = [("Excel files", "*.xlsx")]
        )
        if self.excel_path:
            filename = os.path.basename(self.excel_path)
            self.excel_label.config(text = f"Loaded: {filename}", fg = "green")

            workbook = openpyxl.load_workbook(self.excel_path)

            sheet = workbook.active

            data = []
            for row in sheet.iter_rows(min_col = 1, max_col = 1, values_only = True):
                if row[0] is not None:
                    data.append(row[0])
            
            self.names = data

            return data
        else:
            self.excel_label.config(text = "No file loaded", fg = "gray")

    def process_pdf(self):
        self.output_path = filedialog.askdirectory(
            title = "Select Output Path"
        )

        doc = pdf.open(self.file_path)
        
        num = 0

        for page in doc:
            text = self.names[num]
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
