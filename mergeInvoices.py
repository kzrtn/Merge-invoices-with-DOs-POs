# importing required classes
from pypdf import PdfReader
from pypdf import PdfWriter
import os
from os import listdir
from os.path import isfile, join

# global variables
MY_PATH = "C:/Users/grosu/Desktop/Invoices (undone)"
DO_PATH = "C:/Users/grosu/Desktop/scanned files/new DOs"
PO_PATH = "C:/Users/grosu/Desktop/Invoices (undone)/POs for combining"
files_in_path = [f for f in listdir(MY_PATH) if isfile(join(MY_PATH, f))]

def PDFmerge(pdfs, output):
    # creating a pdf file writer object
    pdfWriter = PdfWriter()

    # appending pdfs one by one
    for pdf in pdfs:
        pdfWriter.append(pdf)

    #writing combined pdf to output pdf file
    with open(output, 'wb') as f:
        pdfWriter.write(f)

def find_substring_between(text, start, end):
    start_index = text.find(start) + len(start)
    
    if end != '\n':
        end_index = text.find(end)
    
    # Getting the company name is tricky, it's sandwiched between the date and a new line
    # e.g. 15 Jul 2025Company Name Here
    else:
        line_breaks = [i for i, char in enumerate(text) if char == '\n']
        if len(line_breaks) >= 7:
            start_index += 11
            end_index = line_breaks[6]
        else:
            print("Error: There aren't at 7 line breaks in document.")
    
    return text[start_index:end_index].strip()

print("Starting script...")

for filename in files_in_path:
    reader = PdfReader(MY_PATH + '/' + filename) #creating a pdf reader object
    page = reader.pages[0] #creating a page object
    extracted_string = page.extract_text() #extracts text from page

    #print(extracted_string)

    # finds DO number
    start = extracted_string.find("DO No: ")+7
    DO_no = extracted_string[start:start+4].strip()

    # finds invoice number
    invoice_number = find_substring_between(extracted_string, "Invoice No.", "Issue Date")
    invoice_number = invoice_number.replace("/","-") # Filenames cannot contain '/', replace with '-'

    # finds company name
    company_name = find_substring_between(extracted_string, "Invoice Date: ", '\n')
    if "Pte Ltd" in company_name:
        company_name = company_name.replace(" Pte Ltd","").strip()

    # Set pdf files to merge and output filename
    pdfs = [MY_PATH + '/' + filename, DO_PATH + '/DO '+ DO_no +'.pdf']
    output_filename = company_name + ' Inv ' + invoice_number + ' DO ' + DO_no

    # if Invoice has a PO number and is billed to Primech
    if "Primech A&P" in extracted_string and "PO No:" in extracted_string:
        PO_no = find_substring_between(extracted_string, "PO No: ", "Invoice No.")
        PO_filepath = [PO_PATH + '/PO ' + PO_no.replace("000","") + ' primech.pdf']

        pdfs = pdfs + PO_filepath

        output_filename = output_filename + ' PO' + PO_no # output pdf file
    
    output = MY_PATH + '/' + output_filename + '.pdf'
    PDFmerge(pdfs=pdfs, output=output) # call pdf merge function
    os.remove(MY_PATH + '/' + filename) # deletes the original invoice file

    print("Saved " + output)
print("Finished merging invoices with DOs and POs (if any).")

