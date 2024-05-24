import os
import json
import csv

# Define the schema
SEARCH_SCHEMA = {
    "receipt_information": {
        "date_time": "text",
        "sales_person": "text",
        "store_phone_number": "text",
        "store_number": "text",
        "store_location": "text",
        "receipt_number": "text",
        "order_number": "text"
    },
    "customer_information": {
        "name": "text",
        "phone": "text",
        "email": "text",
        "address": "text"
    },
    "items_purchased": [
        {
            "description": "text",
            "model_number": "text",
            "sku_number": "text",
            "unit_price": "text",
            "quantity": "number",
            "subtotal": "text"
        }
    ],
    "payment_information": {
        "subtotal": "text",
        "discounts": "text",
        "sales_tax": "text",
        "order_total": "text",
        "balance_due": "text",
        "payment_method": "text",
        "payment_amount": "text"
    },
    "additional_information": {
        "return_policy": "text",
        "pro_xtra_member_statement": {
            "as_of": "text",
            "spend": "text",
            "savings": "text"
        },
        "survey_entry": {
            "description": "text",
            "user_id": "text",
            "password": "text"
        }
    }
}

def load_data_from_file(file_path):
    with open(file_path, 'r') as file:
        data = json.load(file)
    return data

def process_files_in_directory(directory):
    all_data = []
    for filename in os.listdir(directory):
        if filename.endswith(".json"):
            file_path = os.path.join(directory, filename)
            data = load_data_from_file(file_path)
            all_data.append(data)
    return all_data

def collect_data_in_csv(data, category, output_path):
    fieldnames = SEARCH_SCHEMA[category][0].keys()
    with open(output_path, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for item in data:
            if category in item:
                if isinstance(item[category], list):
                    for sub_item in item[category]:
                        writer.writerow(sub_item)
                else:
                    writer.writerow(item[category])

def main():
    directory = 'C:/Experiments/Extract Texts from PDFs/PDFs'
    all_data = process_files_in_directory(directory)
    
    for category in SEARCH_SCHEMA.keys():
        output_path = f"{category}.csv"
        collect_data_in_csv(all_data, category, output_path)
        print(f"Data collected for category '{category}' in '{output_path}'")

if __name__ == "__main__":
    main()
