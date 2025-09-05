import os

os.system(
    'pyinstaller --onefile --windowed --icon=Zamini_Musafir_logo.ico --add-data "file;file" File_Based_Login.py'
)
