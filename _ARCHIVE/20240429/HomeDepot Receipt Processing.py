import os
import re
import json
import xml.etree.ElementTree as ET

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Load the search terms from the JSON file
json_file_path = os.path.join(script_dir, 'RegressionSEarchTerms_HomeDepotReceipts.json')
with open(json_file_path, 'r') as file:
    search_terms_data = json.load(file)

# Function to extract data between start and end keys
def extract_data(data, start_key, end_key):
    if end_key == "\\n":
        regex_pattern = re.escape(start_key) + r'(.*?)\n'
    else:
        regex_pattern = re.escape(start_key) + r'(.*?)' + re.escape(end_key)
    
    match = re.search(regex_pattern, data, re.DOTALL)
    if match:
        return match.group(1).strip()
    else:
        return None

# Process each search term
search_terms = []
for term_data in search_terms_data['searchTerms']:
    start_key = term_data['startKey']
    end_key = term_data['endKey']
    name = term_data.get('name', start_key)  # Use custom name if provided, otherwise use startKey
    search_terms.append({'startKey': start_key, 'endKey': end_key, 'name': name})

# Get the path to the folder containing the text files
folder_path = os.path.join(script_dir, 'PDFs')

# Create the root element for the XML tree
root = ET.Element("data")

# Process each text file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'r') as file:
            data = file.read()
        
        # Create an element for the file data
        file_element = ET.SubElement(root, "file")
        file_element.set("name", filename)

        # Add data elements for each extracted data
        for term in search_terms:
            start_key = term['startKey']
            end_key = term['endKey']
            name = term['name']
            data_value = extract_data(data, start_key, end_key)
            
            # Create an element for the data
            data_element = ET.SubElement(file_element, "data")
            data_element.set("name", name)
            data_element.text = data_value if data_value is not None else ""

# Create an ElementTree object and write it to a file
tree = ET.ElementTree(root)
output_file_path = os.path.join(script_dir, 'output_data.xml')
tree.write(output_file_path)

print(f"Data written to: {output_file_path}")
