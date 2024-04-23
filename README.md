# Excel Sheets to Word Exporter

## Overview
This VBA script automates the process of exporting the contents of Excel sheets into separate Word documents. Each worksheet in a selected Excel file is exported to a separate Word document.

## Requirements
- Microsoft Excel
- Microsoft Word

## Usage
1. Open the Excel file containing the sheets you want to export.
2. Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
3. Insert a new module by selecting `Insert > Module`.
4. Copy and paste the provided VBA script into the module.
5. Close the VBA editor.
6. Press `Alt + F8` to open the "Macro" dialog box.
7. Select the `ExportSheetsToWord` macro.
8. Click `Run`.

## Instructions
1. Upon running the macro, a file picker dialog will appear.
2. Navigate to and select the Excel file (*.xlsx) you want to process.
3. The script will then loop through each worksheet in the selected Excel file.
4. For each worksheet, it will create a new Word document and copy the contents of non-empty cells into it.
5. Each Word document will be saved in the same directory as the Excel file, with the name of the corresponding worksheet.

## Note
- Ensure that macros are enabled in Excel for the script to run successfully.
- The script will save the Word documents in the same directory as the Excel file. Make sure the Excel file is saved before running the script.
- This script assumes that the content of each cell in the Excel sheet can be directly inserted into the Word document. You may need to adjust the script to suit your specific formatting requirements.
