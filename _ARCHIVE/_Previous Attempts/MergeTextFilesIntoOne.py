import os

def merge_text_files(folder_path, output_file):
    with open(output_file, 'w') as outfile:
        for filename in os.listdir(folder_path):
            if filename.endswith(".txt"):
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'r') as infile:
                    outfile.write(infile.read())
                # Add dashes as separator
                outfile.write("\n------------------------\n")

if __name__ == "__main__":
    folder_path = "C:/Experiments/Extract Texts from PDFs/PDFs"
    output_file = "merged_output.txt"
    merge_text_files(folder_path, output_file)
    print("Text files merged successfully into", output_file)
