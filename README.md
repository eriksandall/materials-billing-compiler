# Materials Billing Compiler

[![Code Style: Google](https://img.shields.io/badge/code%20style-google-blueviolet.svg)](https://github.com/google/gts)
[![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)

A Google Apps Script project that automates the collection and processing of student material fees from multiple data sources, applies subsidies, and generates billing reports.

## Overview

This project helps Jacobs Institute for Design Innovation at UC Berkeley compile billing information for student material fees from multiple systems (Shopify and 3DPrinterOS), apply subsidies, handle opt-out requests, and prepare data for the CalCentral billing system.

## Features
- **Multi-source data integration**: Combines purchase data from Shopify and 3DPrinterOS
- **Multiple registration sources**: Combines student data from C-Cure and JPS systems
- **Student information lookup**: Maps purchases to student records using email and UID
- **Subsidies management**: Applies subsidies to eligible students by email or UID
- **Opt-out handling**: Separates students who opted out of CalCentral billing
- **Automated reporting**: Generates comprehensive billing reports in Google Sheets
- **Data validation**: Validates all data sources before processing

## Setup Instructions

1. Create a Google Sheet with tabs for the following data sources (note the sheet ID):
    - Shopify export data - name the tab 'Shopify'.
    - 3DPOS export data - name the tab '3DPOS'.
    - Subsidy eligibility list - name the tab 'Subsidy'.

2. Get the sheet IDs for other data sources:
    - C-Cure student registration data
    - JPS (Jacobs Project System) registration data
    - Output spreadsheet (the "Material Store Payment Tracking" sheet)
  
3. Update the `CONFIG` object in `Code.gs` with the spreadsheet IDs and your email address:
  ```javascript
  const CONFIG = {
    inputData: {
      sheetId: 'INPUT_DATA_SHEET_ID', // Update this
      tabs: {
        shopify: 'Shopify',
        pos: '3DPOS',
        subsidy: 'Subsidy'
      }
    },
    registration: { 
      ccure: {
        sheetId: 'REG_SHEET_ID', // Update this
        tabName: 'C-Cure' 
      },
      jps: {
        sheetId: 'JPS_REG_SHEET_ID', // Update this
        tabName: 'Approved'
      }
    },
    output: { 
      sheetId: 'OUTPUT_SHEET_ID', // Update this
      tabs: {
        billing: 'Total amounts spent',
        optOut: 'Opt-out still need to pay',
        optOutResponses: 'Opt-out form responses'
      }
    },
    ownerEmail: 'your_email@domain.com' // Update this
  };
  ```

## Data Format Requirements

### Shopify Export
Must include columns: `First Name`, `Last Name`, `Email`, `UID`, `Total`

### 3DPOS Export
Must include columns: `Account Email`, `Total Sum ($)`

### C-Cure Registration Data
Must include columns: `Last`, `First`, `SID/EID`, `UID`, `Email`

### JPS Registration Data
Must include columns: `Last Name`, `First Name`, `SID/EID`, `UID`, `Email Address`

### Subsidy List
Must include columns: `Email`, `UID` (students can be eligible by either identifier)

### Opt-out List
Must include columns: `Email Address`, `Student/Employee ID` (students can opt-out by either identifier)

## Usage

1. Open the output Google Spreadsheet
2. Look for the custom menu "Jacobs Tools" (if you don't see it, refresh the page)
3. Click "Run Billing Script" to start the billing compilation process
4. Review the output in the destination spreadsheet

## Output

The script generates two tabs in the output spreadsheet:

1. **Total amounts spent**: Students to be billed through CalCentral
2. **Opt-out still need to pay**: Students who opted out but still need to pay

Each output includes:
- Student information (name, email, ID)
- Original subtotal (before subsidy)
- Whether a subsidy was applied
- Grand total (after subsidy)

## Special Features

- **SID Format Preservation**: The script ensures Student IDs aren't incorrectly formatted as dates
- **Subsidy Application**: $25 subsidy applied to eligible students
- **Data Enrichment**: Student data from registration is used to fill in missing info
- **Authorization**: Only the owner email specified in the config can run the script
- **Data Validation**: All required sheets and columns are validated before processing

## Troubleshooting

If you encounter errors:
  - Verify that all required spreadsheets exist and contain the necessary columns
  - Check that the CONFIG object contains correct spreadsheet IDs and tab names
  - Ensure you have sufficient permissions for all spreadsheets
  - Review the script logs in the Apps Script editor for detailed error messages
  - For menu-related issues, make sure the script is bound to the output spreadsheet

## Development

This project is developed using [clasp](https://github.com/google/clasp), which allows for local development of Google Apps Script projects:

```bash
# Install clasp globally
npm install -g @google/clasp

# Login to your Google account
clasp login

# Clone this repository
git clone https://github.com/yourusername/materials_billing_compiler.git

# Push changes to Google Apps Script
clasp push
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.