# Merge invoice, delivery order (DO) and purchase order (PO) PDFs together
A simple Python script that merges PDF files together.

It reads the invoice PDF's text then then merges the file with the correct DO (and PO, if any). It also works with multiple invoices within the Invoices folder for mass merging. **Note: There is no OCR usage in this project. Invoice's PDF has to be text-selectable.**

The output filename follows the following format:
`(client name) Inv (invoice number) DO (delivery order number) PO(purchase order number).pdf`

Currently, the script reads the invoices from the `/Invoices` folder. Same with the respective `/DOs` (delivery order) and `/POs` (purchase order) folders. The script only reads the PDF text from the invoice in order to obtain DO and PO numbers, the DOs and POs have to follow a filename format of `DO (number).pdf` and `PO (number).pdf`. There are sample documents within this project for testing.

## Dependencies
PyPDF is required for this script. Execute the command below on your terminal.

```pip install pypdf```
