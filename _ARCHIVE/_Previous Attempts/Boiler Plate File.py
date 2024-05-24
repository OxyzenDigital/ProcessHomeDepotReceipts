import re

# Sample text data
text_data = """
Receipt # 0563 00097 77038
 3701-108 PO / Job Name
"""

# Define the regex pattern
pattern = r'Receipt # \d+(.*)'

# Search for the pattern in the text data
match = re.search(pattern, text_data)

# Extract the matched data
if match:
    matched_data = match.group(1)
    print("Matched data:", matched_data)
else:
    print("Pattern not found.")
