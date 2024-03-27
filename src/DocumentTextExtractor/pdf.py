import sys
import pdfplumber

filename = sys.argv[1]
# print("Filename: " + filename)

with pdfplumber.open(filename) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        print(text)
