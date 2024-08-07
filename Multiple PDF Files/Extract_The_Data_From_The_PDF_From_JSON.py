import fitz  # PyMuPDF
import json
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_text_from_pdf(pdf_path, json_path):
    # Load the JSON data
    with open(json_path, 'r') as file:
        rectangles = json.load(file)

    # Open the PDF file
    doc = fitz.open(pdf_path)
    
    extracted_data = []

    for page_key, boxes in rectangles.items():
        # Extract page number from the key (assuming format: "page number: <page_number>")
        page_number = int(page_key.split(": ")[1]) - 1  # Convert to 0-based index

        if page_number < 0 or page_number >= len(doc):
            print(f"Page {page_number + 1} is out of range.")
            continue

        page = doc.load_page(page_number)
        
        for box in boxes:
            name = box['name']
            coords = box['coords']
            # Convert coordinates to fitz.Rect
            pdf_rect = fitz.Rect(coords[0], coords[1], coords[2], coords[3])
            text = page.get_text("text", clip=pdf_rect).strip()
            # Remove newline characters and excessive whitespace
            text = ' '.join(text.split())
            extracted_data.append({'name': name, 'text': text})

    doc.close()

    # Print the extracted data in the specified format
    for item in extracted_data:
        print(f"{item['name']} = {item['text']}")

def main():
    # Create a simple GUI to select files
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Prompt user to select JSON file
    json_path = filedialog.askopenfilename(
        title="Select JSON file",
        filetypes=[("JSON Files", "*.json")]
    )
    if not json_path:
        messagebox.showerror("Error", "No JSON file selected")
        return

    # Prompt user to select PDF file
    pdf_path = filedialog.askopenfilename(
        title="Select PDF file",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not pdf_path:
        messagebox.showerror("Error", "No PDF file selected")
        return

    # Extract text from PDF based on JSON data
    extract_text_from_pdf(pdf_path, json_path)

if __name__ == "__main__":
    main()
