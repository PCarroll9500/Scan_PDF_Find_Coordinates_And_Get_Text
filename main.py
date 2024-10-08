# Really Great PDF Data Extractor  v0.01
#
# By: https://github.com/PCarroll9500
#
# Purpose: To create a user friendly way to mass extract data from PDF files into excel.

import tkinter as tk
from tkinter import filedialog, messagebox
import os
from Import_Document_To_Train import PDFViewer
import os
import Scan_Folder_Extract_Data as sfe

# Create the folders if they don't exist
os.makedirs("To Scan", exist_ok=True)
os.makedirs("Scanned", exist_ok=True)
os.makedirs("Form Templates", exist_ok=True)

print(f"Folders set up: 'To Scan', 'Scanned', 'Form Templates'")


# Function to import a document
def import_document():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("Image files", "*.png;*.jpg;*.jpeg")])
    if pdf_path:
        # Hide the main window
        root.withdraw()

        def on_close_pdf_viewer():
            # Show the main window again when PDFViewer is closed
            root.deiconify()

        # Open the PDFViewer window
        viewer = PDFViewer(pdf_path, on_close_pdf_viewer)
        viewer.mainloop()


# Function to scan a folder and extract data
def scan_folder_extract_data():
    sfe.main()


# Function to import and view in Excel
def view_excel():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        messagebox.showinfo("File Selected", f"CSV file {file_path} has been selected.")
        # Add your CSV handling and Excel viewing code here


if __name__ == "__main__":
    # Create the main window
    root = tk.Tk()

    # Set the window title
    root.title("Really Great PDF Data Extractor")

    # Set the window size
    root.geometry("500x200")

    # Set the background color to light black (dark gray)
    root.configure(bg="#333333")

    # Set resizable to False
    root.resizable(False, False)

    # Create buttons with a light background for contrast
    btn_style = {
        'width': 40,
        'height': 2,
        'bg': '#f0f0f0',
        'fg': 'black',
        'activebackground': '#d3d3d3',
        'font': ('Helvetica', 12, 'bold'),
        'relief': 'flat',
        'bd': 0,
    }

    # Create buttons
    btn_import_document = tk.Button(root, text="Import Document to Train", **btn_style, command=import_document)
    btn_import_document.pack(pady=10)

    btn_scan_folder = tk.Button(root, text="Scan Folder and Extract Data", **btn_style, command=scan_folder_extract_data)
    btn_scan_folder.pack(pady=10)

    btn_import_excel = tk.Button(root, text="View in Excel", **btn_style, command=view_excel)
    btn_import_excel.pack(pady=10)

    # TODO: exit button?

    # Run the main loop
    root.mainloop()
