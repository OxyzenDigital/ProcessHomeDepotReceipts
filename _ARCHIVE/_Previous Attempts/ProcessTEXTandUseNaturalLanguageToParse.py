import os
import camelot

def extract_tables_from_pdf(pdf_file):
    tables = camelot.read_pdf(pdf_file, flavor='stream')
    return tables

def summarize_tables(tables):
    summary = []
    for table in tables:
        # Iterate through each table
        for row in table.df.itertuples(index=False):
            # Summarize data from each row
            # For demonstration purposes, let's just concatenate all cells in the row
            summary.append(" | ".join([str(cell) for cell in row]))
    return summary

def main():
    pdf_folder = "C:/Experiments/Extract Texts from PDFs/PDFs"  # Replace with the path to your PDF files folder

    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_file = os.path.join(pdf_folder, filename)
            print("Processing file:", pdf_file)
            tables = extract_tables_from_pdf(pdf_file)
            summary = summarize_tables(tables)
            
            # Print the summary
            for line in summary:
                print(line)
            print("---------------------------------------")

if __name__ == "__main__":
    main()
