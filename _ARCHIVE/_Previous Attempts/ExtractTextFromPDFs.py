import os
from PyPDF2 import PdfReader

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

# Specify the folder path containing your PDFs
folder_path = "C:/Experiments/Extract Texts from PDFs/PDFs"  # Replace with your actual path

# Call the function to start processing
extract_text_from_pdfs(folder_path)