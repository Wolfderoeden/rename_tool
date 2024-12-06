import os
import shutil
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, StringVar, messagebox, Checkbutton, IntVar
import PyPDF2
from docx import Document

# Funktion zum Lesen von SKUs aus einer Excel-Datei
def read_skus_from_excel(file_path):
    try:
        data = pd.read_excel(file_path, header=None)
        skus = data[0].dropna().astype(str).tolist()  # Konvertiere alle Werte zu Strings
        return skus
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return []

# Funktion zur Suche der SKU in einer PDF-Datei
def search_sku_in_pdf(file_path, sku):
    try:
        with open(file_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            for page in reader.pages:
                if sku in page.extract_text():
                    return True
    except Exception:
        return False
    return False

# Funktion zur Suche der SKU in einer Word-Datei
def search_sku_in_word(file_path, sku):
    try:
        sku = str(sku)  # Sicherstellen, dass SKU als String behandelt wird
        doc = Document(file_path)
        for table in doc.tables:  # Durch alle Tabellen in der Word-Datei iterieren
            for row in table.rows:  # Jede Zeile der Tabelle
                for cell in row.cells:  # Jede Zelle der Zeile
                    if sku in cell.text:  # Überprüfung, ob die SKU im Text der Zelle enthalten ist
                        return True  # SKU gefunden
    except Exception as e:
        print(f"Fehler beim Durchsuchen der Word-Datei: {file_path}, Fehler: {e}")
    return False

# Funktion zur Verarbeitung der SKUs
def process_skus(skus, source_folder, target_folder, name_pattern, include_subfolders, error_log_path):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    with open(error_log_path, 'w') as error_log:
        for sku in skus:
            found = False
            for root, _, files in os.walk(source_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    if file.lower().endswith('.pdf'):
                        found = search_sku_in_pdf(file_path, sku)
                    elif file.lower().endswith('.docx'):
                        found = search_sku_in_word(file_path, sku)
                    if found:
                        new_name = name_pattern.replace("{sku}", sku)
                        new_path = os.path.join(target_folder, new_name)
                        shutil.copy(file_path, new_path)
                        break
                if found and not include_subfolders:
                    break
            if not found:
                error_log.write(f"{sku}\n")

# Haupt-GUI-Anwendung
class RenameFilesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Rename Tool")

        # GUI-Variablen
        self.source_folder = StringVar()
        self.target_folder = StringVar()
        self.excel_file = StringVar()
        self.name_pattern = StringVar(value="p1001-{sku}.pdf")
        self.include_subfolders = IntVar()

        # Widgets
        Label(root, text="Source Folder:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        Entry(root, textvariable=self.source_folder, width=50).grid(row=0, column=1, padx=5, pady=5)
        Button(root, text="Browse", command=self.browse_source_folder).grid(row=0, column=2, padx=5, pady=5)

        Label(root, text="Target Folder:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        Entry(root, textvariable=self.target_folder, width=50).grid(row=1, column=1, padx=5, pady=5)
        Button(root, text="Browse", command=self.browse_target_folder).grid(row=1, column=2, padx=5, pady=5)

        Label(root, text="Excel File:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        Entry(root, textvariable=self.excel_file, width=50).grid(row=2, column=1, padx=5, pady=5)
        Button(root, text="Browse", command=self.browse_excel_file).grid(row=2, column=2, padx=5, pady=5)

        Label(root, text="Name Pattern:").grid(row=3, column=0, sticky='e', padx=5, pady=5)
        Entry(root, textvariable=self.name_pattern, width=50).grid(row=3, column=1, padx=5, pady=5)

        Checkbutton(root, text="Include Subfolders", variable=self.include_subfolders).grid(row=4, column=1, sticky='w', padx=5, pady=5)

        Button(root, text="Start", command=self.start_processing).grid(row=5, column=1, pady=10)

    def browse_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)

    def browse_target_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.target_folder.set(folder)

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.excel_file.set(file_path)

    def start_processing(self):
        source_folder = self.source_folder.get()
        target_folder = self.target_folder.get()
        excel_file = self.excel_file.get()
        name_pattern = self.name_pattern.get()
        include_subfolders = bool(self.include_subfolders.get())

        if not source_folder or not target_folder or not excel_file or not name_pattern:
            messagebox.showerror("Error", "All fields are required!")
            return

        skus = read_skus_from_excel(excel_file)
        if not skus:
            messagebox.showerror("Error", "No SKUs found in the provided Excel file.")
            return

        error_log_path = os.path.join(target_folder, "fehler_skus.txt")

        # Start processing
        process_skus(skus, source_folder, target_folder, name_pattern, include_subfolders, error_log_path)
        messagebox.showinfo("Completed", "Processing completed. Check the target folder and error log if needed.")

# Anwendung ausführen
if __name__ == "__main__":
    root = Tk()
    app = RenameFilesApp(root)
    root.mainloop()
