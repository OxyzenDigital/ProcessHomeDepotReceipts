import os
import csv
import re
import json
from datetime import datetime

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
    # Path to JSON configuration file (relative to the script's location)
    script_dir = os.path.dirname(__file__)
    json_file_path = os.path.join(script_dir, "RegressionSEarchTerms_HomeDepotReceipts.json")

    # Path to folder containing text data files (relative to the script's location)
    folder_path = os.path.join(script_dir, "PDFs")

    # Read search terms from JSON file
    search_terms = read_search_terms_from_json(json_file_path)

    # Create a list to hold data for CSV
    csv_data = []

    # Data extraction date
    data_extraction_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Iterate over text data files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                data = file.read()

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

            csv_data.append(file_data)

    # Write data to CSV file
    output_file_path = os.path.join(script_dir, "output.csv")
    write_to_csv(csv_data, output_file_path)

# Entry point of the script
if __name__ == "__main__":
    main()
