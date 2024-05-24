import os
import csv
import re

def extract_data(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    # Extracting date
    date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', content)
    date = date_match.group(1) if date_match else "N/A"

    # Extracting PO Number
    po_match = re.search(r'Order #([A-Z0-9-]+)', content)
    po_number = po_match.group(1) if po_match else "N/A"

    # Extracting Order Total
    order_total_match = re.search(r'Order Total \$([\d.]+)', content)
    order_total = order_total_match.group(1) if order_total_match else "N/A"

    # Extracting Items Purchased
    items = re.findall(r'\d+\s+(.*?)\s+N/A\s+(\d+)\s+\$([\d.]+)\s+/\s+each\n\$(\d+\.\d{2})', content, re.S)
    items_purchased = [{"Item Description": item[0].strip(), "SKU #": "", "Unit Price": item[2], "Qty": item[1], "Subtotal": item[3]} for item in items]

    return {"Date": date, "PO Number": po_number, "Order Total": order_total, "Items Purchased": items_purchased}


def main():
    folder_path = "C:/Experiments/Extract Texts from PDFs/PDFs"
    output_file = "output.csv"

    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["Date", "PO Number", "Order Total", "Item Description", "SKU #", "Unit Price", "Qty", "Subtotal"]
        csv_writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        csv_writer.writeheader()

        for filename in os.listdir(folder_path):
            if filename.endswith(".txt"):
                file_path = os.path.join(folder_path, filename)
                print("Processing file:", file_path)
                data = extract_data(file_path)

                # Filter out any extra fields not in fieldnames
                filtered_data = {key: value for key, value in data.items() if key in fieldnames}

                csv_writer.writerow(filtered_data)

    print("Data extracted and written to", output_file)

if __name__ == "__main__":
    main()
