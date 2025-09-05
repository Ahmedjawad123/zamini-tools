import os
import win32com.client
import tempfile

def convert_excel_to_pdf_default(file_path):
    temp_pdf = os.path.join(
        tempfile.gettempdir(),
        f"WithoutPageSetup_{os.path.splitext(os.path.basename(file_path))[0]}.pdf"
    )
    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_app.Visible = False
    try:
        wb = excel_app.Workbooks.Open(os.path.abspath(file_path))

        # No PageSetup modifications here

        wb.ExportAsFixedFormat(0, temp_pdf)
        wb.Close(False)
        print(f"PDF with default print settings saved to: {temp_pdf}")
    finally:
        excel_app.Quit()

if __name__ == "__main__":
    input_file = r"C:\Users\Asus\Downloads\100-kb.xlsx"  # Change this to your Excel file path
    convert_excel_to_pdf_default(input_file)
