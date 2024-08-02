# Really Great PDF Data Extractor  v0.01
#
# By: https://github.com/PCarroll9500
#
# Purpose: To create a user friendlyway to mass extract data from PDF files.

import tkinter as tk
from tkinter import filedialog, messagebox

# Function to import a document
def import_document():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("Image files", "*.png;*.jpg;*.jpeg")])
    if file_path:
        messagebox.showinfo("File Selected", f"Document {file_path} has been selected.")
        # Add your document handling code here

# Function to scan a folder and extract data
def scan_folder_extract_data():
    folder_path = filedialog.askdirectory()
    if folder_path:
        messagebox.showinfo("Folder Selected", f"Folder {folder_path} has been selected.")
        # Add your scanning and data extraction code here

# Function to import and view in Excel
def view_excel():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        messagebox.showinfo("File Selected", f"CSV file {file_path} has been selected.")
        # Add your CSV handling and Excel viewing code here

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

# Run the main loop
root.mainloop()