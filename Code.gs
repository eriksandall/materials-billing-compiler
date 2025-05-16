/**
 * Opens a spreadsheet by ID and returns the specified sheet.
 * 
 * @param {string} sheetId - The ID of the spreadsheet
 * @param {string} tabName - The name of the sheet tab
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The specified sheet
 */
function getSheetByIdAndName(sheetId, tabName) {
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  return spreadsheet.getSheetByName(tabName);
}

/**
 * Reads Shopify export data and returns an array of student billing objects
 * 
 * @return {Array<Object>} Array of objects with first, last, email, uid, and total amount
 */
function readShopify() {
  // Get the Shopify sheet
  const sheet = getSheetByIdAndName(CONFIG.inputData.sheetId, CONFIG.inputData.tabs.shopify);
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Extract header row and find indices of required columns
  const headers = data[0];
  const firstNameIdx = headers.indexOf('First Name');
  const lastNameIdx = headers.indexOf('Last Name');
  const emailIdx = headers.indexOf('Email');
  const uidIdx = headers.indexOf('UID');
  const totalIdx = headers.indexOf('Total');
  
  // Ensure all required columns exist
  if (firstNameIdx === -1 || lastNameIdx === -1 || emailIdx === -1 || 
      uidIdx === -1 || totalIdx === -1) {
    throw new Error('Shopify sheet is missing required columns. Need: First Name, Last Name, Email, UID, Total');
  }
  
  // Process data rows (skip header)
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[firstNameIdx] && !row[lastNameIdx] && !row[emailIdx]) continue;
    
    // Create student object
    result.push({
      first: row[firstNameIdx],
      last: row[lastNameIdx],
      email: row[emailIdx],
      uid: row[uidIdx],
      total: parseFloat(row[totalIdx]) || 0 // Convert to number, default to 0 if NaN
    });
  }
  
  return result;
}

/**
 * Reads 3DPOS export data and returns an array of student billing objects
 * 
 * @return {Array<Object>} Array of objects with email and total amount
 */
function read3DPOS() {
  // Get the 3DPOS sheet
  const sheet = getSheetByIdAndName(CONFIG.inputData.sheetId, CONFIG.inputData.tabs.pos);
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Extract header row and find indices of required columns
  const headers = data[0];
  const emailIdx = headers.indexOf('Account Email');
  const totalIdx = headers.indexOf('Total Sum ($)');
  
  // Ensure all required columns exist
  if (emailIdx === -1 || totalIdx === -1) {
    throw new Error('3DPOS sheet is missing required columns. Need: Account Email, Total Sum ($)');
  }
  
  // Process data rows (skip header)
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[emailIdx]) continue;
    
    // Create student object
    result.push({
      email: row[emailIdx],
      total: parseFloat(row[totalIdx]) || 0 // Convert to number, default to 0 if NaN
    });
  }
  
  return result;
}

/**
 * Reads student registration data and creates lookup maps by email and UID
 * 
 * @return {Object} Object containing two Maps: emailToStudent and uidToStudent
 */
function readRegistration() {
  // Get the registration sheet
  const sheet = getSheetByIdAndName(CONFIG.registration.sheetId, CONFIG.registration.tabName);
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Extract header row and find indices of required columns
  const headers = data[0];
  const emailIdx = headers.indexOf('Email');
  const uidIdx = headers.indexOf('UID');
  const sidIdx = headers.indexOf('SID/EID');
  const firstIdx = headers.indexOf('First');
  const lastIdx = headers.indexOf('Last');
  
  // Ensure all required columns exist
  if (emailIdx === -1 || uidIdx === -1 || sidIdx === -1 || 
      firstIdx === -1 || lastIdx === -1) {
    throw new Error('Registration sheet is missing required columns. Need: Email, UID, SID/EID, First, Last');
  }
  
  // Create lookup maps
  const emailToStudent = new Map();
  const uidToStudent = new Map();
  
  // Process data rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[emailIdx] && !row[uidIdx]) continue;
    
    // Create student object
    const student = {
      sid: row[sidIdx],
      first: row[firstIdx],
      last: row[lastIdx]
    };
    
    // Add to maps if email/uid exists
    if (row[emailIdx]) {
      emailToStudent.set(row[emailIdx].toLowerCase(), student); // Store email lookup keys as lowercase
    }
    
    if (row[uidIdx]) {
      uidToStudent.set(String(row[uidIdx]), student); // Convert UID to string for consistent lookup
    }
  }
  
  return {
    emailToStudent,
    uidToStudent
  };
}

/**
 * Merges purchase data from Shopify and 3DPOS systems
 * 
 * @param {Array<Object>} shopifyArr - Array of Shopify purchase records
 * @param {Array<Object>} posArr - Array of 3DPOS purchase records
 * @return {Array<Object>} Array of merged purchase records
 */
function mergePurchases(shopifyArr, posArr) {
  // Create a map to store merged records, keyed by email
  const purchaseMap = new Map();
  
  // Process Shopify purchases first
  shopifyArr.forEach(record => {
    if (record.email) {
      const email = record.email.toLowerCase(); // Normalize email
      purchaseMap.set(email, {
        email: record.email,
        uid: record.uid || '',
        first: record.first || '',
        last: record.last || '',
        subtotal: record.total || 0
      });
    }
  });
  
  // Process 3DPOS purchases, adding to existing records or creating new ones
  posArr.forEach(record => {
    if (record.email) {
      const email = record.email.toLowerCase(); // Normalize email
      
      if (purchaseMap.has(email)) {
        // Add to existing record
        const existingRecord = purchaseMap.get(email);
        existingRecord.subtotal += record.total || 0;
      } else {
        // Create new record
        purchaseMap.set(email, {
          email: record.email,
          uid: '',
          first: '',
          last: '',
          subtotal: record.total || 0
        });
      }
    }
  });
  
  // Convert map values to array
  return Array.from(purchaseMap.values());
}

/**
 * Applies subsidy to eligible student records
 * 
 * @param {Array<Object>} records - Array of purchase records
 * @param {Set<string>} subsidyEmails - Set of emails eligible for subsidy
 * @param {Set<string>} subsidyUIDs - Set of UIDs eligible for subsidy
 * @return {Array<Object>} Updated records with subsidy applied
 */
function applySubsidy(records, subsidyEmails, subsidyUIDs) {
  // Loop through all records and apply subsidy where eligible
  records.forEach(record => {
    // Normalize email and UID for comparison
    const email = record.email ? record.email.toLowerCase() : '';
    const uid = record.uid ? String(record.uid) : '';
    
    // Check if student is eligible for subsidy by email or UID
    if ((email && subsidyEmails.has(email)) || (uid && subsidyUIDs.has(uid))) {
      // Apply subsidy (minimum 0)
      record.subtotal = Math.max(0, record.subtotal - 25);
      record.subsidyApplied = true;
    } else {
      record.subsidyApplied = false;
    }
  });
  
  return records;
}

/**
 * Splits records into billing and opt-out categories
 * 
 * @param {Array<Object>} records - Array of purchase records
 * @param {Set<string>} optOutSet - Set of emails or SIDs that have opted out
 * @return {Array<Array<Object>>} Array containing two arrays: [billingRecords, optOutRecords]
 */
function splitOptOut(records, optOutSet) {
  const billingRecords = [];
  const optOutRecords = [];
  
  // Process each record
  records.forEach(record => {
    // Normalize email for comparison
    const email = record.email ? record.email.toLowerCase() : '';
    const sid = record.sid ? String(record.sid) : '';
    
    // Check if student is in the opt-out list by email or SID
    if ((email && optOutSet.has(email)) || (sid && optOutSet.has(sid))) {
      optOutRecords.push(record);
    } else {
      billingRecords.push(record);
    }
  });
  
  return [billingRecords, optOutRecords];
}

/**
 * Writes data to a tab in a spreadsheet, creating or replacing it if needed
 * 
 * @param {string} sheetId - ID of the spreadsheet
 * @param {string} tabName - Name of the tab to write to
 * @param {Array<string>} headers - Array of column headers
 * @param {Array<Array<any>>} rows - 2D array of data to write
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet that was written to
 */
function writeTab(sheetId, tabName, headers, rows) {
  // Open the spreadsheet
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  
  // Check if tab exists and delete it if it does
  const existingSheet = spreadsheet.getSheetByName(tabName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create a new sheet with the specified name
  const sheet = spreadsheet.insertSheet(tabName);
  
  // Write headers to the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Make headers bold
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // Write data rows if any exist
  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, headers.length);
  
  return sheet;
}

/**
 * Creates a custom menu when the spreadsheet is opened
 */
function onOpen() {
  try {
    // Create custom menu without checking user email first
    SpreadsheetApp.getActiveSpreadsheet()
      .addMenu('Jacobs Tools', [
        {name: 'Run Billing Script', functionName: 'main'}
      ]);
  } catch (e) {
    // Handle any errors that might occur
    console.error('Error in onOpen function:', e);
  }
}

/**
 * Verifies user and runs the billing process
 * This function is called from the menu and can use Session and UI
 */
function main() {
  try {
    // First verify the user is authorized
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail !== CONFIG.ownerEmail) {
      SpreadsheetApp.getUi().alert('You are not authorized to run this script.');
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Starting billing process...');
    console.log('Starting billing compilation process');
    
    // Read data from all sources
    console.log('Reading Shopify purchase data...');
    const shopifyPurchases = readShopify();
    console.log(`Read ${shopifyPurchases.length} Shopify purchase records`);
    
    console.log('Reading 3DPOS purchase data...');
    const posPurchases = read3DPOS();
    console.log(`Read ${posPurchases.length} 3DPOS purchase records`);
    
    console.log('Reading student registration data...');
    const registrationData = readRegistration();
    console.log(`Read registration data with ${registrationData.emailToStudent.size} email records and ${registrationData.uidToStudent.size} UID records`);
    
    // Merge purchases from different systems
    console.log('Merging purchase records from both systems...');
    let records = mergePurchases(shopifyPurchases, posPurchases);
    console.log(`Created ${records.length} merged purchase records`);
    
    // Enrich records with student data from registration
    console.log('Enriching records with student registration data...');
    records = records.map(record => {
      const email = record.email ? record.email.toLowerCase() : '';
      const uid = record.uid ? String(record.uid) : '';
      
      // Try to find student in registration data
      let studentInfo = null;
      
      if (email && registrationData.emailToStudent.has(email)) {
        studentInfo = registrationData.emailToStudent.get(email);
      } else if (uid && registrationData.uidToStudent.has(uid)) {
        studentInfo = registrationData.uidToStudent.get(uid);
      }
      
      // Merge student info if found
      if (studentInfo) {
        return {
          ...record,
          sid: studentInfo.sid || '',
          first: record.first || studentInfo.first || '',
          last: record.last || studentInfo.last || ''
        };
      }
      
      return record;
    });
    console.log('Finished enriching purchase records with student data');
    
    // Read subsidy eligibility data
    console.log('Reading subsidy eligibility data...');
    const subsidySheet = getSheetByIdAndName(CONFIG.inputData.sheetId, CONFIG.inputData.tabs.subsidy);
    const subsidyData = subsidySheet.getDataRange().getValues();
    const subsidyHeader = subsidyData[0];
    const subsidyEmailIdx = subsidyHeader.indexOf('Email');
    const subsidyUIDIdx = subsidyHeader.indexOf('UID');
    
    if (subsidyEmailIdx === -1 || subsidyUIDIdx === -1) {
      throw new Error('Subsidy sheet is missing required Email or UID columns');
    }
    
    // Create sets of subsidy-eligible emails and UIDs
    const subsidyEmails = new Set();
    const subsidyUIDs = new Set();
    for (let i = 1; i < subsidyData.length; i++) {
      const email = subsidyData[i][subsidyEmailIdx];
      const uid = subsidyData[i][subsidyUIDIdx];
      if (email) {
        subsidyEmails.add(email.toLowerCase());
      }
      if (uid) {
        subsidyUIDs.add(String(uid));
      }
    }
    console.log(`Found ${subsidyEmails.size} students eligible for subsidy by email`);
    console.log(`Found ${subsidyUIDs.size} students eligible for subsidy by UID`);
    
    // Apply subsidies to eligible records
    console.log('Applying subsidies to eligible records...');
    records = applySubsidy(records, subsidyEmails, subsidyUIDs);
    console.log('Finished applying subsidies');
    
    // Read opt-out list
    console.log('Reading opt-out list...');
    const optOutSheet = getSheetByIdAndName(CONFIG.output.sheetId, CONFIG.output.tabs.optOutResponses);
    const optOutData = optOutSheet.getDataRange().getValues();
    const optOutHeader = optOutData[0];
    
    const optOutEmailIdx = optOutHeader.indexOf('Email Address');
    const optOutSidIdx = optOutHeader.indexOf('Student/Employee ID');
    
    if (optOutEmailIdx === -1 && optOutSidIdx === -1) {
      throw new Error('Opt-out sheet is missing required columns. Need either Email Address or Student/Employee ID column');
    }
    
    // Create set of opt-out identifiers
    const optOutSet = new Set();
    for (let i = 1; i < optOutData.length; i++) {
      const row = optOutData[i];
      
      // Add email to opt-out set if it exists
      if (optOutEmailIdx !== -1 && row[optOutEmailIdx]) {
        optOutSet.add(String(row[optOutEmailIdx]).toLowerCase());
      }
      
      // Add SID to opt-out set if it exists
      if (optOutSidIdx !== -1 && row[optOutSidIdx]) {
        optOutSet.add(String(row[optOutSidIdx]).toLowerCase());
      }
    }
    console.log(`Found ${optOutSet.size} unique identifiers in opt-out list`);
    
    // Split records into billing and opt-out categories
    console.log('Splitting records into billing and opt-out categories...');
    const [billingRecords, optOutRecords] = splitOptOut(records, optOutSet);
    console.log(`Split records: ${billingRecords.length} for billing, ${optOutRecords.length} opted out`);
    
    // Sort records by email for consistency
    console.log('Sorting records by email...');
    billingRecords.sort((a, b) => {
      const emailA = (a.email || '').toLowerCase();
      const emailB = (b.email || '').toLowerCase();
      return emailA.localeCompare(emailB);
    });
    
    optOutRecords.sort((a, b) => {
      const emailA = (a.email || '').toLowerCase();
      const emailB = (b.email || '').toLowerCase();
      return emailA.localeCompare(emailB);
    });
    console.log('Finished sorting records');
    
    // Prepare data for billing tab
    console.log('Preparing data for billing tab...');
    const billingHeaders = ['First Name', 'Last Name', 'Email', 'Student ID', 'Subtotal', 'Subsidy Applied', 'Grand Total', 'Billed to CalCentral', 'Note'];
    const billingRows = billingRecords.map(record => {
      const grandTotal = record.subtotal;
      return [
        record.first || '',
        record.last || '',
        record.email || '',
        record.sid || '',
        record.subtotal + (record.subsidyApplied ? 25 : 0),  // Show original amount before subsidy
        record.subsidyApplied ? 'Yes' : 'No',
        grandTotal,
        'Yes',  // Default to Yes for billing to CalCentral
        ''      // Empty note column
      ];
    });
    
    // Prepare data for opt-out tab
    console.log('Preparing data for opt-out tab...');
    const optOutHeaders = ['First Name', 'Last Name', 'Email', 'Student ID', 'Subtotal', 'Subsidy Applied', 'Grand Total', 'Payment Method', 'Note'];
    const optOutRows = optOutRecords.map(record => {
      const grandTotal = record.subtotal;
      return [
        record.first || '',
        record.last || '',
        record.email || '',
        record.sid || '',
        record.subtotal + (record.subsidyApplied ? 25 : 0),  // Show original amount before subsidy
        record.subsidyApplied ? 'Yes' : 'No',
        grandTotal,
        '',      // Empty payment method column
        ''       // Empty note column
      ];
    });
    
    // Write to output tabs
    console.log('Writing billing records to output spreadsheet...');
    writeTab(CONFIG.output.sheetId, CONFIG.output.tabs.billing, billingHeaders, billingRows);
    console.log(`Wrote ${billingRows.length} billing records`);
    
    console.log('Writing opt-out records to output spreadsheet...');
    writeTab(CONFIG.output.sheetId, CONFIG.output.tabNames.optOut, optOutHeaders, optOutRows);
    console.log(`Wrote ${optOutRows.length} opt-out records`);
    
    // Log final summary
    console.log(`
    ===== BILLING COMPILATION SUMMARY =====
    Total students processed: ${records.length}
    Students to bill via CalCentral: ${billingRecords.length}
    Students who opted out: ${optOutRecords.length}
    Subsidy-eligible students: ${subsidyEmails.size}
    Process completed successfully
    =====================================
    `);
    
    // Show completion message with summary
    ui.alert(`Billing process complete!\n\n` +
             `Total students processed: ${records.length}\n` +
             `Students to bill via CalCentral: ${billingRecords.length}\n` +
             `Students who opted out: ${optOutRecords.length}`);
    
    console.log('Billing compilation process completed successfully');
             
  } catch (e) {
    // Handle errors gracefully
    console.error('ERROR in billing process:', e);
    console.error(e.stack);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error: ' + e.message);
  }
}
