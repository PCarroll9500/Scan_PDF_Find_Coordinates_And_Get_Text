# Scan_PDF_Find_Coordinates_And_Get_Text
 Purpose: To let a user define the coordinates for information to be extracted from a PDF via a GUI. Extract Text for later use in other programs.

1. User uploads a document:
    The user interacts with the GUI to upload a document.
    The GUI displays the uploaded document to the user.
    User places rectangles on the document:

2. The user places rectangles on the document through the GUI.
    The GUI sends the coordinates and names of the rectangles to the JSON Handler.
    The JSON Handler saves this information and confirms the save back to the GUI.
    User runs the scanner:

3. The user initiates the scanning process through the Scanner.
    The Scanner retrieves the saved coordinates and names from the JSON Handler.
    The Scanner processes the documents and fills a CSV file with the scanned data.
    User uploads the CSV to Excel:

4. The user uploads the generated CSV file to Excel.
    Excel displays the tracked data from the CSV file to the user


![Diagram For PDF Coordinatesvg](https://github.com/user-attachments/assets/c637cc7e-18ba-4461-99ae-41752ec600dd)

# DLA disposition services 1348-1A form help page
https://www.dla.mil/Disposition-Services/DDSR/Turn-In/1348-Help/
