# Quickbooks-invoice-pdf-to-excel
A Python script to extract invoice data from QuickBooks PDF invoices and map it to an Excel import template.
# Mass Extract QuickBooks Invoice PDFs to Excel Import Template

This project provides a script to extract invoice data from multiple QuickBooks PDF invoices and map the extracted data to an Excel import template.

## Features

The script includes three main functions:

1. **Extract Invoice Data**: Extracts multiple QuickBooks invoice numbers, bill-to information, and invoice table data from PDF files. It also filters the data based on specified conditions which can be modified as needed.
2. **Create Excel Template**: Generates an Excel template based on the current system import mapping.
3. **Map Data to Excel Template**: Maps the extracted invoice data to the generated Excel template.

## Requirements

- Python 3.x
- `pdfplumber`
- `PyPDF2`
- `pandas`
- `openpyxl`
- `tk`

You can install the required packages using:

```bash
pip install -r requirements.txt
