import os
import csv
import re
import json
from datetime import datetime
import argparse
import sys
from PyPDF2 import PdfReader

# Function to extract data between start and end keys
def extract_data(data, start_key, end_keys):
    for end_key in end_keys:
        if end_key == "\\n":
            regex_pattern = re.escape(start_key) + r'(.*?)(?:\n|$)'
        else:
            regex_pattern = re.escape(start_key) + r'(.*?)' + re.escape(end_key)

        match = re.search(regex_pattern, data, re.DOTALL)
        if match:
            return match.group(1).strip()
    return None

# Special case for extracting PO / Job Name
def extract_po_job_name(data, start_key):
    regex_pattern = r'(.*?)\s*' + re.escape(start_key)
    match = re.search(regex_pattern, data)
    if match:
        return match.group(1).strip()
    return None

# Function to format the data
def format_data(input_data):
    formatted_data = re.sub(r'\n(?!0\d)', ' ', input_data)
    return formatted_data

# Read search terms from JSON configuration file
def read_search_terms_from_json(json_file_path):
    try:
        with open(json_file_path, 'r', encoding='utf-8') as json_file:
            return json.load(json_file)['searchTerms']
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error reading JSON file: {e}")
        return []

# Write data to CSV file
def write_to_csv(csv_data, output_file_path):
    try:
        with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csv_data[0].keys())
            writer.writeheader()
            for data in csv_data:
                writer.writerow(data)
    except IOError as e:
        print(f"Error writing to CSV file: {e}")

# Main function
def main():
    # Argument parser
    parser = argparse.ArgumentParser(description="Extract data from PDF files.")
    parser.add_argument("folder_path", help="Path to the folder containing PDF files.")
    args = parser.parse_args()

    # Path to JSON configuration file (relative to the script's location)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_file_path = os.path.join(script_dir, "RegressionSEarchTerms_HomeDepotReceipts.json")

    # Path to folder containing text data files (relative to the script's location)
    folder_path = args.folder_path

    # Read search terms from JSON file
    search_terms = read_search_terms_from_json(json_file_path)

    # Create a list to hold data for CSV
    csv_data = []

    # Data extraction date
    data_extraction_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Iterate over PDF data files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'rb') as file:  # Open in binary mode for PyPDF2
                reader = PdfReader(file)
                data = ""
                for page in reader.pages:
                    data += page.extract_text()

            file_data = {}

            # Extract data from the file
            for term in search_terms:
                start_key = term.get('startKey')
                end_key = term.get('endKey')
                name = term.get('name', start_key)

                if not isinstance(end_key, list):
                    end_keys = [end_key]
                else:
                    end_keys = end_key

                if start_key == "(.*?)PO / Job Name":
                    extracted_value = extract_po_job_name(data, "PO / Job Name")
                elif start_key == "Payment Method":
                    # Extract data between "Payment Method" and "Charged" along with the entire line where "Charged" is found
                    payment_data = extract_data(data, start_key, ["Charged"])
                    charged_index = data.find("Charged")
                    if charged_index != -1:
                        charged_line = data[charged_index:].split('\n')[0]
                        payment_data = f"{payment_data} {charged_line}"
                    extracted_value = payment_data
                else:
                    extracted_value = extract_data(data, start_key, end_keys)

                if start_key == "QtySubtotal" and extracted_value:
                    extracted_value = format_data(extracted_value)

                # Only update if extracted_value is non-empty, otherwise keep existing value
                if extracted_value:
                    file_data[name] = extracted_value
                elif name not in file_data:
                    file_data[name] = ""

            # Add data extraction date and file path link to the file data
            file_data["Data Extraction Date"] = data_extraction_date
            file_data["File Path"] = file_path

            # Append file data to CSV data
            csv_data.append(file_data)

    # Write data to CSV file in the same folder as the text files
    # output_folder = os.path.dirname(folder_path)
    output_file_path = os.path.join(folder_path, "output.csv")
    write_to_csv(csv_data, output_file_path)

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

# Entry point of the script
if __name__ == "__main__":
    main()
