# Materials Billing Compiler

[![Code Style: Google](https://img.shields.io/badge/code%20style-google-blueviolet.svg)](https://github.com/google/gts)
[![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)

A Google Apps Script project that automates the collection and processing of student material fees from multiple data sources, applies subsidies, and generates billing reports.

## Overview

This project helps Jacobs Institute for Design Innovation at UC Berkeley compile billing information for student material fees from multiple systems (Shopify and 3DPrinterOS), apply subsidies, handle opt-out requests, and prepare data for the CalCentral billing system.

## Features
- **Multi-source data integration**: Combines purchase data from Shopify and 3DPrinterOS
- **Student information lookup**: Maps purchases to student records using email and UID
- **Subsidies management**: Applies subsidies to eligible students
- **Opt-out handling**: Separates students who opted out of CalCentral billing
- **Automated reporting**: Generates comprehensive billing reports in Google Sheets

## Setup Instructions

1. Create a Google Sheet with tabs for the following some data sources (note the sheet ID):
    - Shopify export data - name the tab 'Shopify'.
    - 3DPOS export data - name the tab '3DPOS'.
    - Subsidy eligibility list - name the tab 'Subsidy'.

2. Get the sheet IDs for other data sources:
    - Student registration data ("CCURE Output" sheet)
    - Opt-out list (from the "Material Store Payment Tracking" sheet)
    - Output spreadsheet (the "Material Store Payment Tracking" sheet)
  
3. Update the `CONFIG` object in `Code.gs` with the spreadsheet IDs and tab names:
  ```javascript
  const CONFIG = {
    inputData: {
      sheetId: 'INPUT_DATA_SHEET_ID', // Single sheet ID for input data
      tabs: {
        shopify: 'Shopify',
        pos: '3DPOS',
        subsidy: 'Subsidy'
      }
    },
    registration: { sheetId: 'REG_SHEET_ID', tabName: 'C-Cure' },
    output: { 
      sheetId: 'OUTPUT_SHEET_ID', 
      tabs: {
        billing: 'Total amounts spent',
        optOut: 'Opt-out still need to pay',
        optOutResponses: 'Opt-out form responses' // Added the opt-out responses tab here
      }
    },
    ownerEmail: 'your_email@domain.com' // Add your email here
  };
  ```

## Data Format Requirements

### Shopify Export
Must include columns: `First Name`, `Last Name`, `Email`, `UID`, `Total`

### 3DPOS Export
Must include columns: `Account Email`, `Total Sum ($)`

### Registration Data
Must include columns: `Last`, `First`, `SID/EID`, `UID`, `Email`

### Subsidy List
Must include columns: `Email`, `UID`

### Opt-out List
Must include columns: `Email Address`, `Student/Employee ID`

## Usage

1. Open the Google Spreadsheet where you've deployed the script
2. Look for the custom menu "Jacobs Tools" 
3. Click "Run Billing Script" to start the billing compilation process
4. Review the output in the destination spreadsheet

## Output

The script generates two tabs in the output spreadsheet:

1. **Total amounts spent**: Students to be billed through CalCentral
2. **Opt-out still need to pay**: Students who opted out but still need to pay

## Troubleshooting

If you encounter errors:
  - Verify that all required spreadsheets exist and contain the necessary columns
  - Check that the CONFIG object contains correct spreadsheet IDs and tab names
  - Ensure you have sufficient permissions for all spreadsheets

## License

This project is licensed under the MIT License - see the LICENSE file for details.