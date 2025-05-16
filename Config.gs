/**
 * Configuration object for the billing compiler
 * This file contains all configuration settings for the application
 */
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
    sheetId: 'REG_SHEET_ID', // Update this
    tabName: 'C-Cure' 
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