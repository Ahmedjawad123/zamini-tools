
import os
import sys

if hasattr(sys, '_MEIPASS'):
    # When running as a PyInstaller onefile exe, DLLs are extracted here
    os.environ['TKDND_LIBRARY'] = os.path.join(sys._MEIPASS, 'tkdnd2.8', 'libtkdnd2.8.dll')
else:
    # Running normally in Python environment
    os.environ['TKDND_LIBRARY'] = os.path.join(os.getcwd(), 'tkdnd2.8', 'libtkdnd2.8.dll')

from TkinterDnD2 import TkinterDnD, DND_FILES
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger
from PIL import Image
import os
import fitz  # PyMuPDF
import datetime
import platform
import tempfile
import win32com.client
from docx2pdf import convert as docx2pdf_convert


def convert_office_to_pdf(file_path):
    if platform.system() != "Windows":
        raise EnvironmentError("This only works on Windows with MS Office installed.")

    ext = os.path.splitext(file_path)[1].lower()
    temp_pdf = os.path.join(tempfile.gettempdir(), f"converted_{os.path.basename(file_path)}.pdf")

    excel = None
    powerpoint = None

    try:
        if ext in [".xlsx", ".xls"]:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.WindowState = -4140
            wb = excel.Workbooks.Open(os.path.abspath(file_path))
            for sheet in wb.Sheets:
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
                sheet.PageSetup.CenterHorizontally = True
                sheet.PageSetup.CenterVertically = True
            wb.ExportAsFixedFormat(0, temp_pdf)
            wb.Close(False)
            excel.Quit()

        elif ext in [".docx", ".doc"]:
            # Convert docx to pdf in the same folder first
            docx2pdf_convert(file_path)
            generated_pdf = os.path.splitext(file_path)[0] + ".pdf"
            if os.path.exists(generated_pdf):
                os.rename(generated_pdf, temp_pdf)
            else:
                raise RuntimeError(f"Failed to find converted PDF for {file_path}")

        elif ext in [".pptx", ".ppt"]:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.WindowState = 2
            powerpoint.Visible = True
            presentation = powerpoint.Presentations.Open(
                os.path.abspath(file_path),
                WithWindow=False
            )
            presentation.SaveAs(temp_pdf, 32)
            presentation.Close()
            powerpoint.Quit()

        else:
            raise ValueError("Unsupported Office file for conversion.")

        return temp_pdf

    except Exception as e:
        try:
            if excel:
                excel.Quit()
            if powerpoint:
                powerpoint.Quit()
        except:
            pass
        raise RuntimeError(f"Failed to convert {file_path} to PDF:\n{e}")


class PDFMergerGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Multi File to PDF Merger")
        self.master.geometry("900x650+0+0")
        self.master.configure(bg="#f7f9fc")
        self.files = []
        self.output_folder = None

        self.description_label = Label(
            master,
            text="Output filename format:\n[First selected file name] + '_' + [current datetime: YYYY-MM-DD_HH-MM-SS].pdf\n"
                 "File size depends on combined files.\n"
                 "You can edit the filename below before merging.",
            font=("Arial", 11),
            justify=LEFT,
            fg="#0f4c81",
            bg="#f7f9fc"
        )
        self.description_label.pack(padx=10, pady=(10, 5), anchor="w")

        self.output_name_label = Label(master, text="Output File Name (editable):", font=("Arial", 10, "bold"))
        self.output_name_label.pack(padx=10, anchor="w")
        self.output_name_entry = Entry(master, width=60, font=("Arial", 11))
        self.output_name_entry.pack(padx=10, pady=(0, 15), fill=X)

        self.label = Label(master, text="Drop files below or use 'Add Files'", font=("Arial", 12))
        self.label.pack(pady=5)

        tree_frame = Frame(master)
        tree_frame.pack(padx=10, pady=5, fill=BOTH, expand=False)

        style = ttk.Style()
        style.configure("Treeview",
                        background="#ffffff",
                        foreground="#333333",
                        rowheight=25,
                        fieldbackground="#ffffff")
        style.map("Treeview", background=[("selected", "#b3d4fc")])

        self.tree = ttk.Treeview(tree_frame, columns=("#1",), show="headings", height=10, selectmode="extended")
        self.tree.heading("#1", text="File Name")
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.master.drop_target_register(DND_FILES)
        self.master.dnd_bind('<<Drop>>', self.drop_files)

        self.button_frame = Frame(master)
        self.button_frame.pack(pady=10)

        Button(self.button_frame, text="Add Files", command=self.add_file).grid(row=0, column=0, padx=5)
        Button(self.button_frame, text="Delete Selected", command=self.delete_selected).grid(row=0, column=1, padx=5)
        Button(self.button_frame, text="Move Up", command=self.move_up).grid(row=0, column=2, padx=5)
        Button(self.button_frame, text="Move Down", command=self.move_down).grid(row=0, column=3, padx=5)
        Button(self.button_frame, text="Reset", command=self.reset_all, fg="red").grid(row=0, column=4, padx=5)

        Button(master, text="Merge to PDF", command=self.merge_files, font=("Arial", 12), bg="green", fg="white").pack(pady=10)

        self.status = Label(master, text="Total Files: 0", fg="blue", font=("Arial", 10))
        self.status.pack(pady=5)
        style.configure("TProgressbar", troughcolor="#e0e0e0", background="#4caf50", thickness=20)

        self.progress = ttk.Progressbar(master, orient=HORIZONTAL, length=600, mode='determinate')
        self.progress.pack(pady=10)

    def update_status(self):
        self.status.config(text=f"Total Files: {len(self.files)}")

    def add_file(self):
        files = filedialog.askopenfilenames(filetypes=[
            ("All Supported Files", "*.pdf *.jpg *.jpeg *.png *.webp *.docx *.doc *.xlsx *.xls *.pptx *.ppt *.txt *.csv"),
            ("PDF files", "*.pdf"),
            ("Word files", "*.docx *.doc"),
            ("Excel files", "*.xlsx *.xls"),
            ("PowerPoint files", "*.pptx *.ppt"),
            ("Image files", "*.jpg *.jpeg *.png *.webp"),
            ("Text files", "*.txt"),
            ("CSV files", "*.csv")
        ])
        self.process_files(files)
        self.update_output_filename()

    def drop_files(self, event):
        files = self.master.tk.splitlist(event.data)
        self.process_files(files)
        self.update_output_filename()

    def process_files(self, files):
        supported_exts = [
            ".pdf", ".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff",  # 🆕 added
            ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".txt", ".csv"
        ]
        
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in supported_exts and file not in self.files:
                self.tree.insert('', END, values=(os.path.basename(file),))
                self.files.append(file)
        self.update_status()

    def delete_selected(self):
        selected = list(self.tree.selection())
        indexes = [self.tree.index(item) for item in selected]
        for item in selected:
            self.tree.delete(item)
        for index in sorted(indexes, reverse=True):
            del self.files[index]
        self.update_status()
        self.update_output_filename()

    def move_up(self):
        selected = list(self.tree.selection())
        indexes = [self.tree.index(item) for item in selected]
        if not indexes:
            return
        for i in sorted(indexes):
            if i > 0:
                self.files[i - 1], self.files[i] = self.files[i], self.files[i - 1]
        self.refresh_tree()
        for i in [i - 1 if i > 0 else i for i in indexes]:
            self.tree.selection_add(self.tree.get_children()[i])

    def move_down(self):
        selected = list(self.tree.selection())
        indexes = [self.tree.index(item) for item in selected]
        if not indexes:
            return
        for i in sorted(indexes, reverse=True):
            if i < len(self.files) - 1:
                self.files[i + 1], self.files[i] = self.files[i], self.files[i + 1]
        self.refresh_tree()
        for i in [i + 1 if i < len(self.files) - 1 else i for i in indexes]:
            self.tree.selection_add(self.tree.get_children()[i])

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for file in self.files:
            self.tree.insert('', END, values=(os.path.basename(file),))

    def reset_all(self):
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.output_name_entry.delete(0, END)
        self.update_status()
        self.update_output_filename()

    def generate_auto_filename(self):
        if not self.files:
            base = "merged"
        else:
            base = os.path.splitext(os.path.basename(self.files[0]))[0]
        now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{base}_{now}.pdf"
        return filename

    def update_output_filename(self):
        # Only set default if entry is empty or matches previous auto filename
        current_text = self.output_name_entry.get()
        default_name = self.generate_auto_filename()
        if not current_text or current_text.endswith(".pdf") and (current_text.startswith(default_name[:-len(".pdf")]) or current_text == ""):
            self.output_name_entry.delete(0, END)
            self.output_name_entry.insert(0, default_name)

    def convert_image_to_pdf(self, image_path):
        ext = os.path.splitext(image_path)[1].lower()
        img = Image.open(image_path)
        if ext == ".webp":
            # convert webp to png in memory for PIL to save as PDF
            img = img.convert("RGB")
        else:
            if img.mode in ("RGBA", "LA"):
                img = img.convert("RGB")
        pdf_path = os.path.join(tempfile.gettempdir(), f"{os.path.basename(image_path)}.pdf")
        img.save(pdf_path, "PDF", resolution=100.0)
        return pdf_path

    def convert_text_to_pdf(self, text_path):
        # convert .txt or .csv to PDF (simple)
        pdf_path = os.path.join(tempfile.gettempdir(), f"{os.path.basename(text_path)}.pdf")
        with open(text_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        doc = fitz.open()
        page = doc.new_page()
        text = "".join(lines)
        text = text[:65000]  # limit to avoid crash
        page.insert_text((72, 72), text, fontsize=11)
        doc.save(pdf_path)
        doc.close()
        return pdf_path

    def merge_files(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add files first!")
            return
        output_name = self.output_name_entry.get().strip()
        if not output_name.lower().endswith(".pdf"):
            output_name += ".pdf"
        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            return
        output_path = os.path.join(output_folder, output_name)
        self.progress['maximum'] = len(self.files) + 1
        self.progress['value'] = 0
        self.master.update_idletasks()

        merger = PdfMerger()

        try:
            for i, file in enumerate(self.files, start=1):
                ext = os.path.splitext(file)[1].lower()

                self.progress['value'] = i
                self.status.config(text=f"Processing {os.path.basename(file)} ({i}/{len(self.files)})...")
                self.master.update_idletasks()

                if ext in [".pdf"]:
                    merger.append(file)
                elif ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"]:  # 🆕 added                   
                    pdf_img = self.convert_image_to_pdf(file)
                    merger.append(pdf_img)
                elif ext in [".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"]:
                    pdf_office = convert_office_to_pdf(file)
                    merger.append(pdf_office)
                elif ext in [".txt", ".csv"]:
                    pdf_text = self.convert_text_to_pdf(file)
                    merger.append(pdf_text)
                else:
                    messagebox.showwarning("Unsupported File", f"Skipping unsupported file: {file}")
                self.master.update_idletasks()

            self.progress['value'] = len(self.files) + 1
            merger.write(output_path)
            merger.close()
            self.status.config(text=f"Merged PDF saved: {output_path}")
            messagebox.showinfo("Success", f"PDF merged successfully:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error while merging:\n{e}")
        finally:
            self.progress['value'] = 0

def main():
    root = TkinterDnD.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
