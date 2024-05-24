import os
from PyPDF2 import PdfReader
import sys

def extract_text_from_pdfs(folder_path):
    """
    Iterates through PDFs in a folder and extracts text from each file.

    Args:
        folder_path: The path to the folder containing the PDFs.
    """
    for foldername, subfolders, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith(".pdf"):
                filepath = os.path.join(foldername, filename)
                try:
                    with open(filepath, 'rb') as pdf_file:  # Open in binary mode
                        reader = PdfReader(pdf_file)
                        text = ""
                        for page_num in range(len(reader.pages)):  # Use len for page count
                            page = reader.pages[page_num]
                            text += page.extract_text()  # New extract_mode
          
                    # Create and write to TXT file (replace with your desired naming convention)
                    output_filename = os.path.splitext(filename)[0] + ".txt"  # Extract filename without extension
                    output_path = os.path.join(foldername, output_filename)
                    with open(output_path, 'w', encoding='utf-8') as txt_file:
                        txt_file.write(text)

                    print(f"Extracted text from {filename} and saved to {output_path}")
                except FileNotFoundError:
                    print(f"Error: File {filepath} not found.")

# Path to folder containing PDF data files
folder_path = sys.argv[1]  # Get folder path from command-line argument

# Call the function to start processing
extract_text_from_pdfs(folder_path)

# Write debug information to a text file with the script name and prefix "debug"
script_name = os.path.splitext(os.path.basename(__file__))[0]
debug_file_name = f"___debug_{script_name}.log"
debug_file_path = os.path.join(folder_path, debug_file_name)

# Print debug information
debug_info = f"Python script ran with the following arguments:\n"
debug_info += f"Arguments: {sys.argv}\n"
debug_info += f"Folder Path: {folder_path}\n\n"
with open(debug_file_path, 'w') as debug_file:
    debug_file.write(debug_info)
