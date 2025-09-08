import fitz  # PyMuPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def mm_to_points(mm):
    """Convert millimeters to points (1 mm = 2.83465 points)."""
    return mm * 2.83465

def calculate_bbox(page, offset_right_mm, offset_bottom_mm, width_mm, height_mm):
    """
    Calculate bounding box based on offsets and dimensions in mm.
    
    :param page: A PyMuPDF page object.
    :param offset_right_mm: Offset from the right edge in mm.
    :param offset_bottom_mm: Offset from the bottom edge in mm.
    :param width_mm: Width of the area in mm.
    :param height_mm: Height of the area in mm.
    :return: Bounding box (x0, y0, x1, y1) in points.
    """
    page_width = page.rect.width
    page_height = page.rect.height
    
    offset_right = mm_to_points(offset_right_mm)
    offset_bottom = mm_to_points(offset_bottom_mm)
    width = mm_to_points(width_mm)
    height = mm_to_points(height_mm)
    
    x1 = page_width - offset_right
    y1 = page_height - offset_bottom
    x0 = x1 - width
    y0 = y1 - height
    
    return (x0, y0, x1, y1)

def extract_text_from_region(pdf_path, page_number, offset_right_mm, offset_bottom_mm, width_mm, height_mm):
    """
    Extract text from a specified rectangular region in a PDF.
    
    :param pdf_path: Path to the PDF file.
    :param page_number: Page number to extract from (0-based index).
    :param offset_right_mm: Offset from the right edge in mm.
    :param offset_bottom_mm: Offset from the bottom edge in mm.
    :param width_mm: Width of the region in mm.
    :param height_mm: Height of the region in mm.
    :return: Extracted text.
    """
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_number]
        
        bbox = calculate_bbox(page, offset_right_mm, offset_bottom_mm, width_mm, height_mm)
        text = page.get_text("blocks", clip=bbox)
        doc.close()
        
        return "\n".join([block[4] for block in text])  # Combine text blocks
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return None

def process_folder():
    """
    Allow the user to select a folder and process all PDFs in it.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title="Select Folder with PDF Files")
    
    if not folder_path:
        messagebox.showwarning("No Folder Selected", "No folder was selected.")
        return
    
    # Input details for the region
    offset_right_mm = 0  # Offset from right edge
    offset_bottom_mm = 0  # Offset from bottom edge
    width_mm = 500  # Width of the region
    height_mm = 500  # Height of the region
    page_number = 0  # Extract from the first page (adjust if necessary)
    
    pdf_files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    if not pdf_files:
        messagebox.showinfo("No PDFs Found", "No PDF files found in the selected folder.")
        return
    
    output_folder = os.path.join(folder_path, "Extracted_Texts")
    os.makedirs(output_folder, exist_ok=True)
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        extracted_text = extract_text_from_region(
            pdf_path,
            page_number,
            offset_right_mm,
            offset_bottom_mm,
            width_mm,
            height_mm
        )
        
        if extracted_text:
            output_file = os.path.join(output_folder, f"{os.path.splitext(pdf_file)[0]}_extracted.txt")
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(extracted_text)
            except Exception as e:
                print(f"Error saving file {output_file}: {e}")
        else:
            print(f"No text extracted from {pdf_file}.")
    
    messagebox.showinfo("Success", f"Processed all PDFs. Extracted text saved in:\n{output_folder}")

# Entry Point
if __name__ == "__main__":
    process_folder()
