# PPTs_TO_PDFs_AND_Merger
A script to convert MS Office `PPT/PPTX` files to `PDF files` and then `merge` all the PDF files to a single PDF file.

## Purpose
This script is intended to help students to convert multiple PPTs into PDFs and then merge all of them into single PDF file. This will also help the Linux users to view PPTs files in Linux.

## Dependencies
- Microsoft Windows 7 or higher
- Microsoft Office 2013 or higher
- Python > 2.5
- comtypes: `pip install comtypes`
- PyPDF2: `pip install PyPDF2`

## Usage
Run:

     python PPTtoPDF_MERGER.py input_Folder_name
 
The input_Folder_name directory will contain all the PDFs with the same name as that of their corresponding PPTs and a merged file named `MergedPDF.pdf`
