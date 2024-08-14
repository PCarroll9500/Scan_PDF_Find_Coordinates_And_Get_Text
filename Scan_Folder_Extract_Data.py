import os
import json
import fitz  # PyMuPDF
from Import_Document_To_Train import PDFViewer  # Import the PDFViewer class from your existing module

class Create_OBJ_of_Scan_Folder:
    def __init__(self, folder_path, json_path):
        self.folder_path = folder_path
        self.json_path = json_path
        self.data = {"extracted_data": {}}  # Store extracted data separately
        self.page_coordinates = {}  # Store the initial JSON data separately
        self.load_json()

    def load_json(self):
        # Load the existing data from the JSON file
        try:
            with open(self.json_path, 'r') as json_file:
                self.page_coordinates = json.load(json_file)
        except FileNotFoundError:
            print(f"JSON file {self.json_path} not found. Starting with empty data.")
            self.page_coordinates = {}

    def print_data(self):
        # Print only the extracted data, not the initial JSON data
        print("The Extracted Data is:", self.data["extracted_data"])

    def process_pdf(self, pdf_path):
        # Open the PDF document
        doc = fitz.open(pdf_path)
        extracted_data = []

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_key = f"page number: {page_num + 1}"

            if page_key in self.page_coordinates:
                page_data = self.page_coordinates[page_key]
                for item in page_data:
                    name = item['name']
                    coords = item['coords']
                    rect = fitz.Rect(*coords)
                    text = page.get_text("text", clip=rect)
                    
                    # Append formatted data to the list
                    formatted_entry = {
                        "file_path": pdf_path,
                        "page_num": page_num + 1,  # 1-indexed page number
                        "field_name": name,
                        "extracted_text": text.strip()
                    }
                    extracted_data.append(formatted_entry)

        return extracted_data

    def scan_folder(self):
        # Process each PDF file in the folder
        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_path = os.path.join(root, file)
                    # Process the PDF and get extracted data
                    pdf_data = self.process_pdf(pdf_path)
                    # Store the extracted data under the filename key in self.data["extracted_data"]
                    self.data["extracted_data"][os.path.basename(pdf_path)] = pdf_data
        # Print the final data structure after processing all PDFs
        self.print_data()

# Example usage
if __name__ == "__main__":
    folder_path = os.path.join(os.getcwd(), "Multiple PDF Files")
    json_path = os.path.join(os.getcwd(), "TestBoxes.json")

    extractor = Create_OBJ_of_Scan_Folder(folder_path, json_path)
    extractor.scan_folder()
