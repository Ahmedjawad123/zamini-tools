import os
import sys

# Assuming your script folder contains "tcl" folder with tcl8.6 and tk8.6 subfolders
base_dir = os.path.dirname(os.path.abspath(__file__))
tcl_library = os.path.join(base_dir, "tcl", "tcl8.6")
tk_library = os.path.join(base_dir, "tcl", "tk8.6")

os.environ['TCL_LIBRARY'] = tcl_library
os.environ['TK_LIBRARY'] = tk_library

import tkinter
from tkinter import messagebox


import os
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
from TkinterDnD2 import TkinterDnD, DND_FILES
import win32com.client
import platform
import tempfile
from PIL import Image
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
import datetime
from PIL import Image, ImageTk  # add this import at the top of your script



        

class allinonepdf:
    def __init__(self, merger):
        self.merger = merger
        self.merger.geometry('900x650+0+0')  # size
        self.merger.title('Zamini Converter v1.0.0')
        self.merger.configure(bg='white')  # background color
        # Set window icon (use your .ico file, convert PNG to ICO if needed)
        try:
            self.merger.iconbitmap('Zamini_Musafir_logo.ico')
        except Exception as e:
            print("Icon file not found or invalid:", e)
        
        
        menubar = tk.Menu(self.merger)
        self.merger.config(menu=menubar)

        main_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Menu", menu=main_menu)

        main_menu.add_command(label="Help", command=self.show_help_info)
        main_menu.add_command(label="About Us", command=self.show_about_info)
        main_menu.add_separator()
        main_menu.add_command(label="Exit", command=self.merger.quit)     
        
        
        
        

        # === Make main window expandable ===
        self.merger.columnconfigure(0, weight=1)
        self.files = []

        self.description_label = tk.Label(
            self.merger,
            text="Output File name format:\n"
                 "[First selected file name Name]+'_'+[Current_Datetime:YYYY-MM-DD_HH-MM-SS].pdf",
            bg='white', fg='green')
        self.description_label.grid(row=0, column=0)

        self.lbl_output_file = tk.Label(
            self.merger, text="Output File Name (Editable)",
            font=('times new roman', 12, 'bold'), bg='white')
        self.lbl_output_file.grid(row=1, column=0, padx=10, pady=(10, 0), sticky='w')

        # === Change Entry sticky to expand horizontally and remove fixed width ===
        self.txt_output_file = tk.Entry(self.merger, font=('times new roman', 12,), bg='white')
        self.txt_output_file.grid(row=2, column=0, padx=10, sticky='ew')  # <-- changed sticky and removed width

        self.lbl_drop_or_select = tk.Label(
            self.merger, text="Drop Files Below or use 'Add Files'",font=('times new roman', 12, 'bold'), bg='white')
        self.lbl_drop_or_select.grid(row=3, column=0)

        supported_ext_text = (
            "Supported File Types:\n\n"
            "Office Files:  .pdf, .doc, .docx, .txt, .csv, .xls, .xlsx, .ppt, .pptx\n"
            "Image Files:   .jpg, .jpeg, .png, .webp, .bmp, .tiff, .gif, .ico, .jfif\n"
            "Other Files:   .md, .py, .c, .cpp, .java, .js, .sh, .bat, .html, .css, "
            ".json, .xml, .yaml, .yml, .ini, .conf, .log, .tsv"
        )

        self.lbl_supported_ext = tk.Label(
            self.merger,
            text=supported_ext_text,
            font=('times new roman', 10),
            fg='blue',
            bg='white',
            justify='left'
        )
        self.lbl_supported_ext.grid(row=4, column=0, padx=10, sticky='w')


        tree_frame = tk.Frame(self.merger, bd=10)
        # === Add sticky nsew so tree_frame expands both ways ===
        tree_frame.grid(row=5, column=0, pady=10, sticky='nsew')

        # === Make tree_frame's grid expandable ===
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # === Make treeview bigger and expandable ===
        self.tree = ttk.Treeview(
            tree_frame, columns=("#1",), show='headings',
            height=10,  # increased height for more rows visible
            selectmode='extended')
        self.tree.grid(row=0, column=0, sticky='nsew')  # expand in all directions
        self.tree.heading("#1", text="File Name")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky='ns')

        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self.drop_files)

        self.btn_frame = tk.Frame(self.merger, bd=10)
        self.btn_frame.grid(row=6, column=0, padx=10)

        self.btn_add = tk.Button(self.btn_frame, text='Add Files', bg='green', fg='white',command=self.add_file)
        self.btn_add.grid(row=0, column=0, padx=10)

        self.btn_delete = tk.Button(self.btn_frame, text='Delete Selected', bg='black', fg='white',command=self.delete_selected)
        self.btn_delete.grid(row=0, column=1, padx=10)

        self.btn_move_up = tk.Button(self.btn_frame, text='Move Up', bg='black', fg='white',command=self.move_up)
        self.btn_move_up.grid(row=0, column=2, padx=10)

        self.btn_move_down = tk.Button(self.btn_frame, text='Move Down', bg='black', fg='white',command=self.move_down)
        self.btn_move_down.grid(row=0, column=3, padx=10)

        self.btn_reset = tk.Button(self.btn_frame, text='Reset', bg='red', fg='white',command=self.reset_all)
        self.btn_reset.grid(row=0, column=4, padx=10)

        frame_progress = tk.Frame(self.merger)
        frame_progress.grid(row=7, column=0)

        self.btn_merge = tk.Button(frame_progress, text='Convert To PDF', bg='green', fg='white',command=self.merge_files)
        self.btn_merge.grid(row=0, column=0, padx=10)

        self.lbl_total = tk.Label(frame_progress, text="Total Files: 0")
        self.lbl_total.grid(row=0, column=1, padx=10)
        
        self.Progress= ttk.Progressbar(self.merger, orient='horizontal', length=600, mode='determinate')
        self.Progress.grid(row=8, column=0, pady=10)

        self.status = tk.Label(self.merger, text="", fg="blue", bg="white", font=("Arial", 10))
        self.status.grid(row=9, column=0, pady=5)



    # This function is triggered when files are dragged and dropped into the app
    def drop_files(self, event):
        files = self.merger.tk.splitlist(event.data)  # splitlist() turns it into a list like ['C:/file1.pdf', 'C:/file2.jpg']      # event.data contains the dropped file paths as a single string
        self.process_files(files)
        self.generate_output_filename()
        self.lbl_total.config(text=f"Total Files: {len(self.files)}")

    def process_files(self, files): # This function handles checking and adding the files to the GUI list
        # List of supported file extensions (PDF, images, Office docs, etc.)
        supported_exts = [
            ".pdf", ".jpg", ".jpeg", ".webp", ".png", ".bmp", ".tiff", ".txt",  # image and text files
            ".doc", ".docx", ".xls", ".xlsx", ".pptx", ".ppt", ".csv"           # Office documents
        ]
        for file in files:
            # Extract the file extension (like '.pdf') and convert it to lowercase
            ext = os.path.splitext(file)[1].lower()

            # Check if the extension is supported and the file has not already been added
            if ext in supported_exts and file not in self.files:
                # Insert just the file name (not full path) into the tree/list in the GUI
                self.tree.insert('', "end", values=(os.path.basename(file),))
                # Add the file to the internal list so it's not added again next time
                self.files.append(file)

                
 
    def add_file(self):
        files = filedialog.askopenfilenames(filetypes=[
            ("All Supported Files",
            "*.pdf *.jpg *.jpeg *.png *.webp *.bmp *.tiff *.gif *.ico *.jfif "
            "*.docx *.doc *.xlsx *.xls *.pptx *.ppt "
            "*.txt *.log *.csv *.tsv *.md "
            "*.json *.xml *.yaml *.yml *.ini *.conf "
            "*.py *.c *.cpp *.java *.js *.sh *.bat "
            "*.html *.css"),
            
            ("PDF files", "*.pdf"),
            ("Image files", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff *.gif *.ico *.jfif"),
            ("Word files", "*.docx *.doc"),
            ("Excel files", "*.xlsx *.xls"),
            ("PowerPoint files", "*.pptx *.ppt"),
            ("Text files", "*.txt *.log *.md"),
            ("CSV/TSV files", "*.csv *.tsv"),
            ("Code files", "*.py *.c *.cpp *.java *.js *.sh *.bat"),
            ("Config/Data files", "*.json *.xml *.yaml *.yml *.ini *.conf"),
            ("Web files", "*.html *.css")
        ])
        self.process_files(files)
        self.generate_output_filename()
        self.lbl_total.config(text=f"Total Files: {len(self.files)}")


    
    def convert_office_to_pdf(self,file_path): # Function to convert Office files (Word, Excel, PowerPoint) to PDF
        if platform.system() != "Windows": # If the system is not Windows, stop and throw an error
            raise OSError("This only works on Windows with MS Office Installed.")
        ext=os.path.splitext(file_path)[1].lower()# Get the file extension (like '.docx', '.ppt', etc.) and make it lowercase
        temp_folder=os.path.join(tempfile.gettempdir(),f"Converted_{os.path.basename(file_path)}.pdf")# Create a temporary path to save the converted PDF (in temp folder)    

        startexcel=None
        startppt=None
        startword=None
        # Excel Window States
        xlNormal = -4143       # Normal window state
        xlMinimized = -4140    # Minimized
        xlMaximized = -4137    # Maximized
        # Powerpoint Window States
        pptNormal = 1       # Normal window state
        pptMinimized = 2    # Minimized
        pptMaximized = 3    # Maximized
        # Word Window States
        wdWindowNormal = 0  # Normal window state
        wdWindowMinimized = 2  # Minimized
        wdWindowMaximized = 1  # Maximized
        
        try:
            if ext in [".xlsx",".xls"]:
                startexcel=win32com.client.Dispatch("Excel.Application")# === 1. Setup Word/Excel/PowerPoint app using COM
                startexcel.Visible=True
                startexcel.WindowState=xlMinimized
                # Open workbook (wb)
                try:
                    wb = startexcel.Workbooks.Open(os.path.abspath(file_path))
                except Exception as e:
                    raise RuntimeError(f"Excel failed to open the file:\n{file_path}\n\n{e}")

                if wb is None:
                    raise RuntimeError(f"Excel could not open the file: {file_path}")

                wb.ExportAsFixedFormat(0, temp_folder)

                wb.Close(False)
                startexcel.Quit()
                
            # Excel PDF export
    # wb.ExportAsFixedFormat(0, output_path)  # 0 = PDF
    # # Word PDF export (two ways)
    # doc.ExportAsFixedFormat(0, output_path)  # 0 = PDF
    # # OR
    # doc.SaveAs(output_path,fileformat= 17)               # 17 = PDF
    # # PowerPoint PDF export
    # presentation.SaveAs(output_path, 32)     # 32 = PDF
    # Excel close without saving changes
    # wb.Close(False)

    # # PowerPoint close
    # presentation.Close()

        # If file is PowerPoint
            elif ext in [".pptx",".ppt"]:
            # Launch PowerPoint via COM
                startppt=win32com.client.Dispatch("Powerpoint.Application")
                startppt.Visible=True
                startppt.WindowState=pptMinimized
                # Open presentation (presentation)
                presentation=startppt.Presentations.Open(os.path.abspath(file_path))
                presentation.SaveAs(temp_folder,32)
                presentation.Close()
                startppt.Quit()

        # If file is word
            elif ext in [".docx",".doc"]:
            # Launch Word via COM
                startword=win32com.client.Dispatch("Word.Application")
                startword.Visible=True
                startword.WindowState = wdWindowMinimized  # Word Minimized
                document=startword.Documents.Open(os.path.abspath(file_path))
                document.SaveAs(temp_folder,FileFormat=17)   # Save as PDF (format code 17)
                document.Close(False)
                startword.Quit()
            else:
                raise ValueError(f"Unsupported File Extension: {ext}")
            return temp_folder
        except Exception as e:
            # Try to quit apps if they were started
            try:
                if startexcel:
                    startexcel.Quit()
                if startppt:
                    startppt.Quit()
                if startword:
                    startword.Quit()
            except:
                pass
            raise RuntimeError(f"Failed to convert {file_path} to PDF:\n{e}")
                    
    def convert_images_to_pdf(self,file_path):
        """
        Converts a single image file into a PDF and saves it in the temporary directory.
        Supports JPG, PNG, BMP, GIF, TIFF, WEBP, ICO, JFIF.
        """
        ext = os.path.splitext(file_path)[1].lower()  # Get the file extension and lowercase it
        temp_folder = os.path.join(tempfile.gettempdir(),f"Converted_{os.path.basename(file_path)}.pdf")  # Create a temp file path for saving the PDF

        image_list = []

        try:
            if ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".ico", ".jfif"]:
                # === 1. Open the image file
                img = Image.open(os.path.abspath(file_path))

                # === 2. Convert images with transparency (RGBA, LA, or P) to RGB with white background
                if img.mode in ("RGBA", "LA", "P"):
                    background = Image.new("RGB", img.size, (255, 255, 255))  # white bg

                    # === 3. If not 'P' mode, apply mask using the alpha channel
                    mask = img.split()[-1] if img.mode != "P" else None
                    background.paste(img, mask=mask)

                    img = background
                else:
                    img = img.convert("RGB")  # Simple RGB conversion for other modes

                image_list.append(img)

                # === 4. Save as PDF
                if image_list:
                    image_list[0].save(temp_folder,save_all=True,append_images=image_list[1:])
                else:
                    raise ValueError("No valid image to convert.")

                return temp_folder  # === 5. Return the output path

            else:
                raise ValueError(f"Unsupported image file type: {ext}")

        except Exception as e:
            raise RuntimeError(f"Failed to convert image to PDF: {e}")
    
    
    

    def convert_text_to_pdf(self,file_path):
        """
        Converts text-based files (e.g., .txt, .py, .json) into a multi-page PDF.
        - Handles large files by automatically adding pages.
        - Clean layout using monospaced formatting and proper margins.
        """

        # STEP 1: Define which file types are allowed
        supported_extensions = [
            ".txt", ".log", ".csv", ".tsv", ".md",                   # Plain text types
            ".json", ".xml", ".yaml", ".yml", ".ini", ".conf",       # Config/data formats
            ".py", ".c", ".cpp", ".java", ".js", ".sh", ".bat",      # Code/script files
            ".html", ".css"                                          # Web-related text files
        ]

        # STEP 2: Check if the file extension is supported
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in supported_extensions:
            raise ValueError(f"Unsupported text format: {ext}")

        # STEP 3: Prepare the output file path (temporary folder)
        pdf_filename = f"Converted_{os.path.basename(file_path)}.pdf"
        temp_pdf_path = os.path.join(tempfile.gettempdir(), pdf_filename)

        # STEP 4: Read the content of the text file safely
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()  # Each line in the file becomes one item in the list

        # STEP 5: Setup PDF document
        doc = fitz.open()  # Create a new PDF document

        # Font settings
        font_size = 11                      # Font size in points (standard readable size)
        line_height = font_size + 2        # Extra space between lines for readability

        # Page dimensions and margins (A4 standard)
        page_height = 842                  # A4 page height in points (72 pt = 1 inch)
        top_margin = 72                    # 1 inch from top
        bottom_margin = 72                 # 1 inch from bottom

        # STEP 6: Calculate how many lines fit per page
        max_lines_per_page = int((page_height - top_margin - bottom_margin) / line_height)

        # STEP 7: Break the content into chunks and add them to pages
        for i in range(0, len(lines), max_lines_per_page):
            chunk = lines[i:i + max_lines_per_page]  # A group of lines that fit on one page
            page = doc.new_page()  # Create a new blank page
            y = top_margin  # Start from the top margin

            for line in chunk:
                # Normalize line: remove newlines, replace tabs with spaces for alignment
                line = line.strip("\n").replace("\t", "    ")

                # Insert the text line at position (72 pts from left, y pts from top)
                page.insert_text((72, y), line, fontsize=font_size)

                # Move cursor down for next line
                y += line_height

        # STEP 8: Save PDF to temp folder and close the document
        doc.save(temp_pdf_path)
        doc.close()

        # STEP 9: Return the full path of the saved PDF file
        return temp_pdf_path

            
                
    def merge_files(self):
        # Check if there are files to merge
        if not self.files:
            messagebox.showwarning("No files", "Please add files first!")
            return

        # Get output file name and ensure it ends with .pdf
        output_name = self.txt_output_file.get().strip()
        if not output_name.lower().endswith(".pdf"):
            output_name += ".pdf"

        # Ask user to select output folder
        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            return

        output_path = os.path.join(output_folder, output_name)

        # Setup progress bar maximum and reset value
        self.Progress['maximum'] = len(self.files)
        self.Progress['value'] = 0
        self.merger.update_idletasks()

        converted_files = []
        failed_files = []

        try:
            # Use context manager to ensure PdfMerger closes properly
            with PdfMerger() as merger:
                for i, file in enumerate(self.files, start=1):
                    ext = os.path.splitext(file)[1].lower()

                    # Update progress bar and status
                    self.Progress['value'] = i
                    self.status.config(text=f"Processing {os.path.basename(file)} ({i}/{len(self.files)})...")
                    self.merger.update_idletasks()

                    try:
                        # Handle file types accordingly
                        if ext == ".pdf":
                            # Directly append PDF files (no separate handler needed)
                            final_file = file
                        elif ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"]:
                            final_file = self.convert_images_to_pdf(file)
                        elif ext in [".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"]:
                            final_file = self.convert_office_to_pdf(file)
                        elif ext in [".txt", ".csv"]:
                            final_file = self.convert_text_to_pdf(file)
                        else:
                            raise ValueError("Unsupported file type")

                        # Append the converted file to merger
                        merger.append(final_file)
                        converted_files.append(os.path.basename(file))

                    except Exception as e:
                        # Log and track failed files without stopping process
                        error_msg = f"{os.path.basename(file)} - {str(e)}"
                        failed_files.append(error_msg)
                        print(f"[SKIPPED] {error_msg}")
                        continue

                # Check if any files were converted before writing output
                if not converted_files:
                    self.status.config(text="No files were merged.")
                    messagebox.showwarning("No Output", "No valid files to merge.")
                    return

                # Write the merged PDF to output path
                merger.write(output_path)
                self.status.config(text=f"Merged PDF saved: {output_path}")

        except Exception as e:
            # Handle errors during saving merged PDF
            messagebox.showerror("Error", f"Error while saving converting PDF:\n{e}")
            print(f"[ERROR] Merge failed: {e}")
            return

        finally:
            # Reset progress bar and update UI
            self.Progress['value'] = 0
            self.merger.update_idletasks()

        # Show summary popup with success and failure info
        result_msg = f"Converted PDF saved to:\n{output_path}\n\n"
        if converted_files:
            result_msg += "✅ Converted Files:\n" + "\n".join(converted_files) + "\n\n"
        if failed_files:
            result_msg += "❌ Failed Files:\n" + "\n".join(failed_files)

        messagebox.showinfo("Merge Complete", result_msg)

            
            
            
            

#     | Step  | What It Does             | Why It's Important                         |
# | ----- | ------------------------ | ------------------------------------------ |
# | 1–4   | Validates input/output   | Prevents crashes, ensures user intent      |
# | 5–6   | Prepares GUI feedback    | Keeps user informed during long tasks      |
# | 7     | Creates merger tool      | Core of combining files                    |
# | 8–9   | Processes each file      | Determines how to handle different formats |
# | 10    | Writes final file        | Delivers the result                        |
# | 11    | Confirms success         | Good user feedback                         |
# | 12–13 | Handles errors + cleanup | Avoids freeze or confusion                 |

    def reset_all(self):
        self.tree.delete(*self.tree.get_children())
        self.files.clear()
        self.txt_output_file.delete(0, tk.END)
        self.lbl_total.config(text="Total Files: 0")
        
        
        
        
        
        
        
        

    def generate_output_filename(self):
        if not self.files:
            base = "merged"
        else:
            base = os.path.splitext(os.path.basename(self.files[0]))[0]

        now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_name = f"{base}_{now}.pdf"

        current_text = self.txt_output_file.get()

        # Check if current name is empty or was based on old first file
        current_base = os.path.splitext(current_text)[0].split("_")[0] if current_text.endswith(".pdf") else ""
        if not current_text or current_base not in os.path.basename(self.files[0]):
            self.txt_output_file.delete(0, tk.END)
            self.txt_output_file.insert(0, default_name)

        return default_name


    def delete_selected(self):
        selected_items = self.tree.selection()

        for item in selected_items:
            filename = self.tree.item(item, 'values')[0]
            self.tree.delete(item)
            self.files = [f for f in self.files if os.path.basename(f) != filename]

        # Update total files count
        self.lbl_total.config(text=f"Total Files: {len(self.files)}")

        # ✅ Use your existing logic to update the output file name
        self.generate_output_filename()

    def move_up(self):
        selected_items = self.tree.selection()
        if not selected_items:
            return

        for item in selected_items:
            index = self.tree.index(item)
            if index > 0:
                self.tree.move(item, '', index - 1)
                # Swap in self.files
                self.files[index], self.files[index - 1] = self.files[index - 1], self.files[index]
                # Re-select the moved item
                self.tree.selection_set(item)

        self.generate_output_filename()


    def move_down(self):
        selected_items = self.tree.selection()
        if not selected_items:
            return

        for item in reversed(selected_items):  # reversed for proper multi-selection handling
            index = self.tree.index(item)
            if index < len(self.tree.get_children()) - 1:
                self.tree.move(item, '', index + 1)
                self.files[index], self.files[index + 1] = self.files[index + 1], self.files[index]
                self.tree.selection_set(item)

        self.generate_output_filename()


    def show_help_info(self):
        help_win = tk.Toplevel(self.merger)
        help_win.title("Zamini Converter Help")
        help_win.geometry("500x400")
        help_win.configure(bg='white')

        frame = tk.Frame(help_win, bg='white')
        frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

        help_win.grid_rowconfigure(0, weight=1)
        help_win.grid_columnconfigure(0, weight=1)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.grid(row=0, column=1, sticky='ns')

        text = tk.Text(
            frame, wrap='word', yscrollcommand=scrollbar.set,
            bg='white', font=('Arial', 11)
        )
        text.grid(row=0, column=0, sticky='nsew')

        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        scrollbar.config(command=text.yview)

        # Tags
        text.tag_configure("title", font=('Arial', 12, 'bold'), foreground="#1a73e8")
        text.tag_configure("bold", font=('Arial', 11, 'bold'))
        text.tag_configure("highlight", foreground="#d93025", font=('Arial', 11, 'bold'))
        text.tag_configure("center", justify="center", font=('Arial', 11, 'bold'))

        # Content
        text.insert("end", "How to Use:\n", "title")
        text.insert("end", "- Click ")
        text.insert("end", "Add Files", "highlight")
        text.insert("end", " to select files for conversion.\n")
        text.insert("end", "- Type a file name in the box for your final PDF.\n")
        text.insert("end", "- Click ")
        text.insert("end", "Convert to PDF", "highlight")
        text.insert("end", " to create and save the PDF.\n")
        text.insert("end", "- Use the ")
        text.insert("end", "progress bar", "bold")
        text.insert("end", " to track conversion progress.\n\n")

        text.insert("end", "Extra Features:\n", "title")
        text.insert("end", "- Use ")
        text.insert("end", "Move Up", "highlight")
        text.insert("end", " and ")
        text.insert("end", "Move Down", "highlight")
        text.insert("end", " to rearrange files.\n")
        text.insert("end", "- Click ")
        text.insert("end", "Delete Selected", "highlight")
        text.insert("end", " to remove unwanted files.\n")
        text.insert("end", "- Use ")
        text.insert("end", "Reset", "highlight")
        text.insert("end", " to clear all selections.\n\n")

        text.insert("end", "Enjoy converting your files offline!", "center")
        text.config(state='disabled')


    def show_about_info(self):
        about_win = tk.Toplevel(self.merger)
        about_win.title("About Zamini Converter")
        about_win.geometry("520x480")
        about_win.configure(bg='white')

        frame = tk.Frame(about_win, bg='white', padx=10, pady=10)
        frame.grid(row=0, column=0, sticky='nsew')
        about_win.grid_rowconfigure(0, weight=1)
        about_win.grid_columnconfigure(0, weight=1)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.grid(row=0, column=1, sticky='ns')

        text = tk.Text(
            frame, wrap='word', yscrollcommand=scrollbar.set,
            bg='#fafafa', font=('Segoe UI', 11), spacing3=8,
            relief='solid', borderwidth=1, padx=10, pady=10, fg='#333333'
        )
        text.grid(row=0, column=0, sticky='nsew')
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        scrollbar.config(command=text.yview)

        # Tags
        text.tag_configure("title", font=('Segoe UI', 16, 'bold'), foreground="#1a73e8", spacing3=12)
        text.tag_configure("section", font=('Segoe UI', 13, 'bold'), foreground="#d93025",
                        background="#ffe8e8", spacing3=8, lmargin1=5, lmargin2=5)
        text.tag_configure("bold", font=('Segoe UI', 12, 'bold'), foreground="#222222")
        text.tag_configure("normal", font=('Segoe UI', 11), foreground="#444444", spacing3=6)
        text.tag_configure("bullet", font=('Segoe UI', 11), foreground="#444444", lmargin1=25, lmargin2=45, spacing3=6)
        text.tag_configure("bullet_bold", font=('Segoe UI', 11, 'bold'), foreground="#444444", lmargin1=25, lmargin2=45, spacing3=6)
        text.tag_configure("center", justify="center", font=('Segoe UI', 12, 'bold'), foreground="#1a73e8", spacing3=15)

        # Header
        text.insert("end", "Zamini Converter v1.0.0\n", "title")

        text.insert("end", "Version Number Explanation\n", "section")
        text.insert("end", "v1.0.0 means:\n", "bold")
        text.insert("end",
                    "1 = Major version: big feature changes or redesigns.\n"
                    "0 = Minor version: smaller features or improvements.\n"
                    "0 = Patch version: bug fixes or small tweaks.\n\n", "normal")

        text.insert("end", "Developed by: ", "bold")
        text.insert("end", "Zamini Musafir Team\n", "normal")
        text.insert("end", "Contact: ", "bold")
        text.insert("end", "zamini.musafir123@gmail.com\n", "normal")
        text.insert("end", "Made in Pakistan | Built in Ajman\n", "bold")
        text.insert("end", "\nCategory: ZaminiTools Collection\n\n", "bold")

        text.insert("end", "Description\n", "section")
        text.insert("end", "Convert images, documents, and code files to PDF offline.\n", "normal")
        text.insert("end", "Microsoft Office", "bold")
        text.insert("end", " is required to convert Office files (.docx, .xlsx, etc.).\n\n", "normal")

        text.insert("end", "Best For\n", "section")

        best_for_items = [
            ("Teachers & Students", " — converting assignments, notes, or code to PDF."),
            ("Office workers", " — saving documents, spreadsheets, or presentations as PDFs."),
            ("Print shop staff", " — handling file submissions in various formats."),
            ("Anyone", " — needing a reliable, offline PDF converter.")
        ]
        for main, extra in best_for_items:
            text.insert("end", "• ", "bullet")
            text.insert("end", main, "bullet_bold")
            text.insert("end", extra + "\n", "bullet")

        text.insert("end", "\nThank you for using Zamini Converter!", "center")
        text.config(state='disabled')

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    all2pdf = allinonepdf(root)
    root.mainloop()
# Excel Window States
# xlNormal = -4143       # Normal window state
# xlMinimized = -4140    # Minimized
# xlMaximized = -4137    # Maximized

# # Page Orientation
# xlPortrait = 1         # Portrait orientation
# xlLandscape = 2        # Landscape orientation

# # Paper Sizes
# xlPaperLetter = 1      # US Letter (8.5" x 11")
# xlPaperA4 = 9          # A4 (210mm x 297mm)

# # Print Order
# xlDownThenOver = 1     # Print down, then over
# xlOverThenDown = 2     # Print over, then down

# # Alignment
# xlCenter = -4108       # Center alignment (horizontal/vertical)
# xlLeft = -4131         # Align left
# xlRight = -4152        # Align right
# xlTop = -4160          # Align top
# xlBottom = -4107       # Align bottom

# # Page Setup Zoom
# xlAutomatic = -4105    # Auto zoom
# # Zoom = False disables manual zoom (fit to page used instead)

# # Page Fit
# # PageSetup.FitToPagesWide = 1  # Fit sheet to 1 page wide
# # PageSetup.FitToPagesTall = 1  # Fit sheet to 1 page tall

# # File Format for ExportAsFixedFormat
# xlTypePDF = 0          # Export as PDF
# xlTypeXPS = 1          # Export as XPS

# # Direction
# xlToLeft = -4159       # Move/align to left
# xlToRight = -4161      # Move/align to right
# xlUp = -4162           # Move/align upward
# xlDown = -4121         # Move/align downward

# # Calculation
# xlCalculationAutomatic = -4105
# xlCalculationManual = -4135

# # Visibility
# xlSheetVisible = -1    # Sheet is visible
# xlSheetHidden = 0      # Sheet is hidden
# xlSheetVeryHidden = 2  # Sheet is very hidden (only visible via VBA)


# # Assume 'sheet' is an Excel worksheet object, e.g. from wb.Sheets

# # Fit sheet to 1 page wide and 1 page tall (disable zoom)
# PageSetup.Zoom = False
# PageSetup.FitToPagesWide = 1
# PageSetup.FitToPagesTall = 1

# # Page orientation
# PageSetup.Orientation = 1  # Portrait
# PageSetup.Orientation = 2  # Landscape

# # Center content on the page
# PageSetup.CenterHorizontally = True
# PageSetup.CenterVertically = True

# # Margins in points (1 point = 1/72 inch)
# PageSetup.LeftMargin = 36.0       # 0.5 inch
# PageSetup.RightMargin = 36.0      # 0.5 inch
# PageSetup.TopMargin = 36.0        # 0.5 inch
# PageSetup.BottomMargin = 36.0     # 0.5 inch
# PageSetup.HeaderMargin = 18.0     # 0.25 inch
# PageSetup.FooterMargin = 18.0     # 0.25 inch

# # Print options
# PageSetup.PrintGridlines = True   # Print gridlines on sheet
# PageSetup.PrintHeadings = False   # Do not print row/column headings

# # Black and white printing
# PageSetup.BlackAndWhite = False   # Print in color
# PageSetup.BlackAndWhite = True    # Print in black and white only

# # Other example settings
# # PageSetup.LeftHeader = "&\"Arial,Bold\"&14 My Header"  # Set left header with font Arial Bold 14
# # PageSetup.RightFooter = "&D &T"                      # Date and Time in right footer
