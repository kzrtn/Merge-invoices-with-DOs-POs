# importing required classes
from pypdf import PdfReader
from pypdf import PdfWriter
import os
from os import listdir
from os.path import isfile, join

def PDFmerge(pdfs, output):
    # creating a pdf file writer object
    pdfWriter = PdfWriter()

    # appending pdfs one by one
    for pdf in pdfs:
        pdfWriter.append(pdf)

    #writing combined pdf to output pdf file
    with open(output, 'wb') as f:
        pdfWriter.write(f)

print("Starting script...")

mypath = "/Invoices"
DOpath = "/DOs"
POpath = "/POs"
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

for filename in onlyfiles:
    reader = PdfReader(mypath + '/' + filename) #creating a pdf reader object
    page = reader.pages[0] #creating a page object
    extracted_string = page.extract_text() #extracts text from page

    # finds DO number
    start = extracted_string.find("DO No: ")+7
    DO_no = extracted_string[start:start+4].strip()

    # finds invoice number
    start = extracted_string.find("Invoice No.")+11
    end = extracted_string.find("Issue Date")
    invoice_number = extracted_string[start:end].strip().replace("/","-")

    # finds company name
    start = extracted_string.find("Invoice Date: ")+25

    # this is to find the end of the string (it's a line break)
    line_breaks = [i for i, char in enumerate(extracted_string) if char == '\n']
    if len(line_breaks) >= 7:
        end = line_breaks[6]
    else:
        print("There aren't seven line breaks. Error.")

    company_name = extracted_string[start:end]

    #if Pte Ltd is in the name, remove it
    if "Pte Ltd" in company_name:
        company_name = company_name.replace(" Pte Ltd","").strip()

    # if Invoice has a PO number and is billed to Primech
    if "PO No:" in extracted_string:
        start = extracted_string.find("PO No: ")+7
        end = extracted_string.find("Invoice No.")
        PO_no = extracted_string[start:end].strip()
        POfilepath = POpath + '/PO ' + PO_no + '.pdf'

        pdfs = [mypath + '/' + filename, DOpath + '/DO '+DO_no+'.pdf', POfilepath] #pdf files to merge
        output = mypath + '/' + company_name + ' Inv ' + invoice_number + ' DO ' + DO_no + ' PO' + PO_no + '.pdf' #output pdf file
    else:
        pdfs = [mypath + '/' + filename, DOpath + '/DO '+DO_no+'.pdf'] #pdf files to merge
        output = mypath + '/' + company_name + ' Inv ' + invoice_number + ' DO ' + DO_no + '.pdf' #output pdf file

    PDFmerge(pdfs=pdfs, output=output) #calling pdf merge function
    os.remove(mypath + '/' + filename) #deletes the original invoice file

    print("Saved " + output)
print("Finished merging invoices with DOs and POs (if any).")

