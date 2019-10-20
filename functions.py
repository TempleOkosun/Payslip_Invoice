# Requirements
import os  # Library needed to interact with Operating System for file handling.
from PyPDF2 import PdfFileReader, PdfFileMerger  # Libraries needed for PDF merging function


# Merge pdfs in a directory.
def merge_pdfs():
    # Get current working directory.
    files_dir = os.getcwd()
    # A list comprehension to get files ending with .pdf extension in the directory.
    pdf_files = [f for f in os.listdir(files_dir) if f.endswith('.pdf')]
    merger = PdfFileMerger()
    # Loop to append files from the list
    for filename in pdf_files:
        merger.append(PdfFileReader(os.path.join(files_dir, filename), 'rb'))
    # Output the merged file
    merger.write(os.path.join(files_dir, 'merged_pdfs.pdf'))