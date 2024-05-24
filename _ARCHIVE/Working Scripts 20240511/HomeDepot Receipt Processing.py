import os
import csv
import re
import json

# Function to extract data between start and end keys
def extract_data(data, start_key, end_key):
    if end_key == "\\n":
        regex_pattern = re.escape(start_key) + r'(.*?)(?:\n|$)'
    else:
        regex_pattern = re.escape(start_key) + r'(.*?)' + re.escape(end_key)

    #print(regex_pattern)
    match = re.search(regex_pattern, data, re.DOTALL)
    if match:
        return match.group(1).strip()
    else:
        return None

# Function to format the data
def format_data(input_data):
    # Replace new lines unless they are followed by the number pattern
    formatted_data = re.sub(r'\n(?!0\d)', ' ', input_data)
    return formatted_data


# Read search terms from JSON configuration file
def read_search_terms_from_json(json_file_path):
    with open(json_file_path, 'r', encoding='utf-8') as json_file:
        return json.load(json_file)['searchTerms']

# Path to JSON configuration file
json_file_path = os.path.join(os.path.dirname(__file__), "RegressionSEarchTerms_HomeDepotReceipts.json")

# Path to folder containing text data files
folder_path = os.path.join(os.path.dirname(__file__), "PDFs")

# Create a list to hold data for CSV
csv_data = []

# Iterate over text data files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        # Read text data from file with explicit UTF-8 encoding
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'r', encoding='utf-8') as file:
            data = file.read()

        # Read search terms from JSON configuration file
        search_terms = read_search_terms_from_json(json_file_path)

        # Extract and add data elements for each search term
        file_data = {}
        for term in search_terms:
            start_key = term['startKey']
            end_key = term['endKey']
            name = term.get('name', start_key)  # Use custom name if provided, otherwise use startKey
            
            # For PO / Job Name, use direct regex search
            #print(start_key)
            if start_key == "(.*?)PO / Job Name":
                match = re.search(rf"{start_key}", data)
                #print(match)
                if match:
                    file_data[name] = match.group(1).strip()
                else:
                    file_data[name] = ""  # Set to empty string if not found
            else:
                data_value = extract_data(data, start_key, end_key)
                if  start_key == "QtySubtotal":
                    data_value = format_data(data_value)
                else:
                    data_value
                file_data[name] = data_value if data_value is not None else ""
            #print(data_value)
        # Append file data to CSV data list
        csv_data.append(file_data)

# Get unique column names from all search terms
column_names = []
for data in csv_data:
    for key in data.keys():
        if key not in column_names:
            column_names.append(key)

# Write data to CSV file
output_file_path = os.path.join(os.path.dirname(__file__), "output.csv")
with open(output_file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=column_names)
    writer.writeheader()
    for data in csv_data:
        writer.writerow(data)
