# Really Great PDF Data Extractor  v0.01
#
# By: https://github.com/PCarroll9500
#
# Purpose: To create a user friendly way to mass extract data from PDF files into excel.

import tkinter as tk
from tkinter import filedialog, messagebox
from Import_Document_To_Train import PDFViewer
from Scan_Folder_Extract_Data import Create_OBJ_of_Scan_Folder 

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
    json_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if json_path:
        folder_path = filedialog.askdirectory()
    if folder_path:
        viewer = Create_OBJ_of_Scan_Folder(folder_path, json_path)
        viewer.scan_folder()

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

# Run the main loop
root.mainloop()
