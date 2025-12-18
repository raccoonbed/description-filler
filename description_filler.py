import pymupdf as pdf
import tkinter as tk
from tkinter import filedialog
import os
import openpyxl

class PDFEditor:
    # Run when app starts
    def __init__(self, window):
        self.window = window
        self.window.title("Description Filler")
        self.window.geometry("300x500")
        
        # Define self class variables
        self.file_path = []
        self.output_path = None

        self.file_label = tk.Label(
            window,
            text = "No file(s) loaded",
            font = ("Arial", 10),
            fg = "gray"
        )
        self.file_label.pack(pady=10)

        listbox_label = tk.Label(
            window,
            text="Selected File(s): ",
            font=("Arial", 9)
        )
        listbox_label.pack(pady=5)

        self.files_listbox = tk.Listbox(
            window,
            height=6,
            width=40
        )
        self.files_listbox.pack(pady=5)

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

    def load_descriptions(self, file_path):

        raw_titles = []

        doc = pdf.open(file_path)
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
                        title = title.replace('Ư', 'ff')
                        raw_titles.append(title)
        doc.close()

        if len(raw_titles) > 1:

            split_titles = [t.split() for t in raw_titles]
            common_words = []

            for words in zip(*split_titles):
                if len(set(words)) == 1:
                    common_words.append(words[0])
                else:
                    break

            # DEBUG: Print common prefix found
            print(f"\nCOMMON WORDS FOUND: {common_words}")
            print(f"Number of common words: {len(common_words)}")
            
            if len(common_words) >= 3:
                common_prefix_len = len(' '.join(common_words)) + 1
                titles = [t[common_prefix_len:].strip() for t in raw_titles]

                print(f"\nAFTER REMOVING PREFIX:")
                for i, title in enumerate(titles):
                    print(f"  {i}: '{title}'")
            else:
                titles = raw_titles
        else:
            titles = raw_titles
        
        print(f"\nFINAL TITLES RETURNED:")
        for i, title in enumerate(titles):
            print(f"  {i}: '{title}'")
            print(f"     Contains 'Office': {'Office' in title}, Contains 'ff': {'ff' in title}")

        return titles

    def import_file(self):
        file_paths = filedialog.askopenfilenames(
            title = "Select PDF Files",
            filetypes = [("PDF files", "*.pdf")]
        )

        if file_paths:
            self.file_path = list(file_paths)

            self.files_listbox.delete(0, tk.END)

            for path in self.file_path:
                filename = os.path.basename(path)
                self.files_listbox.insert(tk.END, filename)

            num_files = len(self.file_path)
            self.file_label.config(text=f"Loaded: {num_files} file(s)", fg = "green")

        else:
            self.file_label.config(text = "No file(s) loaded", fg = "gray")

    def process_pdf(self):
        self.output_path = filedialog.askdirectory(
            title = "Select Output Path"
        )
        
        if not self.output_path:
            return
        
        for i in range(len(self.file_path)):

            titles = self.load_descriptions(self.file_path[i])

            doc = pdf.open(self.file_path[i])
            num = 0

            for page in doc:
                if num >= len(titles):
                    break
                
        # Target coordinates (1087, 727); Center of description box
            for page in doc:
                text = titles[num]

                if len(text) <= 9:

                    page.insert_text(
                        (1087 - (3.5*len(text)), 727),
                        text,
                        fontname = "helv",
                        fontsize = 14,
                        color = (0,0,0)
                    )
                elif len(text) <= 19:

                    page.insert_text(
                        (1087 - (3*len(text)), 727),
                        text,
                        fontname = "helv",
                        fontsize = 14,
                        color = (0,0,0)
                    )
                elif len(text) <= 29:

                    page.insert_text(
                        (1087 - (3.25*len(text)), 727),
                        text,
                        fontname = "helv",
                        fontsize = 14,
                        color = (0,0,0)
                    )
                elif len(text) <= 39:
            
                    page.insert_text(
                        (1087 - (2.8*len(text)), 727),
                        text,
                        fontname = "helv",
                        fontsize = 12,
                        color = (0,0,0)
                    )
                else: 
                    long_title = text.split()
                    insert_index = len(long_title) // 2

                    bottom_text = long_title[insert_index:]
                    bottom_text = ' '.join(bottom_text)
                    top_text = long_title[:insert_index]
                    top_text = ' '.join(top_text)
                    offset = (len(bottom_text) - len(top_text)) // 2

                    long_title.insert(insert_index, '\n')

                    for j in range(offset):
                        long_title.insert(0, ' ')

                    text = ' '.join(long_title)

                    page.insert_text(
                        (1077 - (2.5*len(bottom_text)), 720),
                        text,
                        fontname = "helv",
                        fontsize = 11,
                        color = (0,0,0)
                    )
                num += 1

            og_filename = os.path.basename(self.file_path[i])
            new_filename = og_filename.replace('.pdf','_updated.pdf')

            save_path = os.path.join(self.output_path, new_filename )

            doc.save(save_path)
            doc.close()

if __name__ == "__main__":
    window = tk.Tk()
    app = PDFEditor(window)
    window.mainloop()
