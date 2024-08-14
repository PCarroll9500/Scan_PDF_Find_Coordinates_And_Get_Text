import os
import re
import json
import shutil
import openpyxl
import tkinter as tk
import pymupdf as pmu

from time import sleep
from os.path import abspath
from tkinter import filedialog
from openpyxl.styles import Font, Color, colors

from main import to_scan_folder, scanned_folder


# open folder dialog to select folder to scan, defaults to the created 'To Scan' folder
def open_folder_dialog():
    selected_folder = filedialog.askdirectory(initialdir="./To Scan", title="Select Folder", mustexist=True)
    if selected_folder:
        print("Selected Folder:", selected_folder)
    else:
        print("No folder selected. Exiting...")
        exit(0)
    return selected_folder


# TODO: add function to open file dialog to select json file
# TODO: add function to open file dialog to select folder to save scanned files
# TODO: remove commented out code before production
# move files from one directory to another
def move_file(file_name):
    # replace spaces with underscores
    file_name = file_name.replace(" ", "_")

    try:
        print('Moving file...')
        source_path = os.path.join(to_scan_folder, file_name)
        destination_path = os.path.join(scanned_folder, file_name)
        # shutil.move(source_path, destination_path)                # commented out during testing to keep test files in 'To Scan' folder
        print(f"Moved {file_name} to '{scanned_folder}'")
    except Exception as e:
        print(f"An error occurred moving the file: {e}")


# TODO: add function to move files to a 'Failed' folder if an error occurs during processing
# TODO: add feature to select to include subfolders or not
# manage the queue of files in the folder and process them
def queue_manager(working_directory):
    # Iterate through the files in the folder (not subfolders)
    for filename in os.listdir(working_directory):
    # This will process all files in the folder, including subfolders
    # for root, dirs, files in os.walk(working_directory):

        # Create the full path to the file by using os.path.join()
        file_path = os.path.join(working_directory, filename)

        # Check if the path is a file (not a directory)
        if os.path.isfile(file_path):

            # process pdf files
            if file_path.endswith(".pdf"):
                # announce processing file
                print(f"Processing file: {filename}")
                try:
                    # process the file, extract text and populate spreadsheet, then move the file to the scanned folder
                    fields = extract_text_from_pdf(file_path, "./Form Templates/1348.json")
                    populate_spreadsheet(fields, filename)
                    move_file(filename)
                except Exception as e:
                    # skip file and move to next file
                    print(f"Error processing file: {filename} \n\t{e}")
                    continue
            # skip non-pdf and move to next file
            else:
                print(f"File '{filename}' is not a pdf, skipping.")
                continue

# TODO: does this work with multiple 1348s in one pdf? - I don't think so, need to test
# extract text from pdf file
def extract_text_from_pdf(pdf_path, json_path):
    # Load the JSON file containing the coordinates of the fields to extract
    with open(json_path, 'r') as file:
        form_fields = json.load(file)

    # Open the PDF file
    doc = pmu.open(pdf_path)

    extracted_data = []

    for page_key, boxes in form_fields.items():
        # Extract page number from the key (assuming format: "page number: <page_number>")
        page_number = int(page_key.split(": ")[1]) - 1  # Convert to 0-based index

        if page_number < 0 or page_number >= len(doc):
            print(f"Page {page_number + 1} is out of range.")
            continue

        page = doc.load_page(page_number)

        for box in boxes:
            name = box['name']
            coords = box['coords']
            # Convert coordinates to a PyMuPDF Rect object
            pdf_rect = pmu.Rect(coords[0], coords[1], coords[2], coords[3])
            text = page.get_text("text", clip=pdf_rect).strip()
            # Remove newline characters and excessive whitespace
            text = ' '.join(text.split())
            extracted_data.append({'name': name, 'text': text})

    doc.close()

    # Print the extracted data in the specified format
    # for item in extracted_data:
    #     print(f"{item['name']} = {item['text']}")

    return extracted_data


def populate_spreadsheet(fields, pdf_name):
    # Define the hyperlink font style
    hyper_link_font = Font(color=colors.BLUE, underline='single')

    # Create full file path for hyperlink to the original file
    pdf_name = pdf_name.replace(" ", "_")
    # folder = abspath("./Scanned")
    folder = abspath("./To Scan")
    pdf_link_name = os.path.join(folder, pdf_name)

    if os.path.exists("./Scanned/scanned_data.xlsx"):
        # Load the existing workbook
        workbook = openpyxl.load_workbook("./Scanned/scanned_data.xlsx")
        sheet = workbook.active
    else:
        # Create a new workbook and select the active worksheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Add headers
        headers = ["filename"] + [item['name'] for item in fields]
        sheet.append(headers)

    # Add extracted field data to the sheet
    row = ['=HYPERLINK("{}","{}")'.format(pdf_link_name, pdf_name),
           # Document name with hyperlink
           fields[0]['text'],  # Document number
           fields[1]['text'],  # Nomenclature
           fields[2]['text'],  # NSN
           fields[3]['text'],  # UI
           fields[4]['text'],  # QTY
           re.sub(' ', '.', fields[5]['text']),  # Unit Price, replace space with decimal
           re.sub(' ', '.', fields[6]['text']),  # Total Price, replace space with decimal
           fields[7]['text'],  # Disposal Authorization Code
           fields[8]['text'],  # DEMIL Code
           fields[9]['text'],  # Supply Condition Code
           fields[10]['text'],  # Shipped From
           fields[11]['text'],  # Shipped To
           fields[12]['text'],  # POC name
           re.sub('Phone ', '', fields[13]['text']),  # POC Phone, strip 'Phone ' from the beginning
           re.sub('Email ','', fields[14]['text']),  # POC Email, strip 'Email ' from the beginning
           ]

    count = int(sheet.dimensions[4]) + 1
    sheet.append(row)
    # Apply hyperlink style to the pdf link
    sheet.cell(row=count, column=1).font = hyper_link_font

    # Save the workbook
    workbook.save("./Scanned/scanned_data.xlsx")
    print(f"Data from {pdf_name} added to the spreadsheet")


if __name__ == '__main__':
    folder_path = abspath(open_folder_dialog())
    queue_manager(folder_path)
