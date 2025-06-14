// core.gs - Google Apps Script for Parsonage Tenant Management System

/**
 * Global constants for sheet names.
 * These are used throughout the script to refer to specific sheets.
 */
const TENANTS_SHEET_NAME = 'Tenants';
const BUDGET_SHEET_NAME = 'Budget';

/**
 * Headers for the 'Tenants' sheet.
 * Define these clearly to ensure consistency.
 */
const TENANTS_HEADERS = [
  'Room Number',
  'Rental Price',
  'Negotiated Price',
  'Current Tenant Name',
  'Tenant Email',
  'Move-In Date',
  'Security Deposit Paid',
  'Room Status', // e.g., Occupied, Vacant, Pending
  'Last Payment Date',
  'Payment Status - Current Month', // e.g., Paid, Due, Overdue
  'Move-Out Date (Planned)',
  'Notes'
];

/**
 * Headers for the 'Budget' sheet.
 * Define these clearly to ensure consistency.
 */
const BUDGET_HEADERS = [
  'Date',
  'Type', // e.g., Rent Income, Utility Expense, Maintenance
  'Description',
  'Amount', // Positive for income, negative for expense
  'Category' // e.g., Rent, Electricity, Water, Repair
];

/**
 * This function runs automatically when the spreadsheet is opened.
 * It creates a custom menu in the Google Sheet UI, making it easier
 * to access the script's functions.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Parsonage Tools')
      .addItem('Initialize/Format Sheets', 'initializeSheets') // New item to format sheets
      .addSeparator()
      .addItem('Check Payment Status', 'checkAllPaymentStatus')
      .addItem('Send Rent Reminders (Test)', 'sendRentRemindersTest')
      .addToUi();
}

/**
 * Initializes and formats the 'Tenants' and 'Budget' sheets.
 * If the sheets don't exist, they are created. If they do exist,
 * their content is cleared, and headers and formatting are applied.
 * This function ensures a consistent and aesthetically pleasing layout.
 */
function initializeSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Initialize the Tenants sheet
    setupSheet(ss, TENANTS_SHEET_NAME, TENANTS_HEADERS);
    // Initialize the Budget sheet
    setupSheet(ss, BUDGET_SHEET_NAME, BUDGET_HEADERS);

    ui.alert('Sheets Initialized', 'Tenants and Budget sheets have been initialized and formatted.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', `Failed to initialize sheets: ${e.message}`, ui.ButtonSet.OK);
    console.error(`Error initializing sheets: ${e.message}`);
  }
}

/**
 * Helper function to create or get a sheet and apply basic formatting.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The active spreadsheet.
 * @param {string} sheetName The name of the sheet to set up.
 * @param {Array<string>} headers An array of strings representing the column headers.
 */
function setupSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  // If the sheet doesn't exist, create it.
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log(`Sheet '${sheetName}' created.`);
  } else {
    // If the sheet exists, clear its content to ensure a clean slate for re-formatting.
    sheet.clearContents();
    sheet.clearFormats(); // Clear existing formats
    console.log(`Sheet '${sheetName}' cleared for re-initialization.`);
  }

  // Set the header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Apply header formatting
  headerRange.setFontWeight('bold'); // Make headers bold
  headerRange.setBackground('#D9EAD3'); // Light green background for headers (aesthetically pleasing)
  headerRange.setHorizontalAlignment('center'); // Center align headers
  sheet.setFrozenRows(1); // Freeze the header row

  // Adjust column widths for readability and aesthetics.
  // This is a basic example; you might need to fine-tune these.
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i); // Auto-resize columns based on content
    // Add a little extra width for padding or longer potential text
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 20);
  }

  // Set text wrapping for all columns in the header row for better display of long titles.
  sheet.getRange(1, 1, 1, headers.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Set default text wrapping for new content (adjust as needed for specific columns)
  sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Optional: Set default row height for better spacing
  sheet.setDefaultRowHeight(25);
}


/**
 * A placeholder for checking all payment statuses.
 * This will be expanded later to read tenant data and update status.
 */
function checkAllPaymentStatus() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Payment Status Check', 'Checking payment statuses now...', ui.ButtonSet.OK);
  // --- Future logic to read tenant sheet and update status ---
  // For now, just a confirmation.
  ui.alert('Payment Status Check', 'Payment status check complete (placeholder).', ui.ButtonSet.OK);
}

/**
 * A test function to demonstrate sending a rent reminder email.
 * This will be expanded later to target specific overdue tenants.
 */
function sendRentRemindersTest() {
  const recipientEmail = Session.getActiveUser().getEmail(); // Sends to yourself for testing
  const subject = 'Rent Reminder - Parsonage (Test)';
  const body = `Dear Tenant (Test),\n\nThis is a friendly reminder that your rent is due soon (or is overdue).\n\nPlease ensure your payment is made as soon as possible.\n\nThank you,\nParsonage Management`;

  try {
    MailApp.sendEmail(recipientEmail, subject, body);
    SpreadsheetApp.getUi().alert('Test Email Sent', `A test rent reminder email has been sent to ${recipientEmail}. Check your inbox.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Email Error', `Failed to send test email: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(`Error sending test email: ${e.message}`);
  }
}

// Future functions to be added:
// - Functions to read/update specific tenant data
// - Logic for identifying overdue payments based on 'Last Payment Date' and current month
// - Sending targeted overdue notices to house manager
// - Generating and emailing PDF invoices
// - Functions to handle Google Form submissions (requires separate triggers)
// - More advanced budget analysis functions
