# User Stories

## User Story 1: Extract Invoice Data

**As a** user  
**I want to** extract invoice data from multiple QuickBooks PDF invoices  
**So that** I can process and analyze the invoice information efficiently

### Acceptance Criteria
- The script should allow selecting multiple PDF files.
- The script should extract the invoice number, bill-to information, and invoice table data.
- The script should filter the data based on specified conditions which can be modified as needed.

## User Story 2: Create Excel Template

**As a** user  
**I want to** create an Excel template based on the current system import mapping  
**So that** I can ensure the extracted data is mapped correctly

### Acceptance Criteria
- The script should generate an Excel template with predefined column headers and formatting.
- The Excel template should include all necessary columns for order information and product details.

## User Story 3: Map Data to Excel Template

**As a** user  
**I want to** map the extracted invoice data to the generated Excel template  
**So that** I can have a structured and organized representation of the invoice data

### Acceptance Criteria
- The script should map the extracted data to the corresponding columns in the Excel template.
- The script should ensure that the data is aligned correctly without any empty rows.
- The script should auto-fit the column widths to the content.

## User Story 4: Provide Feedback on Completion

**As a** user  
**I want to** receive a notification upon the successful creation and saving of the Excel template  
**So that** I know the process has been completed

### Acceptance Criteria
- The script should display a message box confirming the successful creation and saving of the template.
