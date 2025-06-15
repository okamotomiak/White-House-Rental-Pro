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
 * Column index constants (1-based) for easier reference.
 */
const COL_ROOM_NUMBER = 1;
const COL_RENTAL_PRICE = 2;
const COL_NEGOTIATED_PRICE = 3;
const COL_TENANT_NAME = 4;
const COL_TENANT_EMAIL = 5;
const COL_MOVE_IN_DATE = 6;
const COL_SECURITY_DEPOSIT = 7;
const COL_ROOM_STATUS = 8;
const COL_LAST_PAYMENT = 9;
const COL_PAYMENT_STATUS = 10;
const COL_MOVE_OUT_PLANNED = 11;
const COL_NOTES = 12;

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
 * Email address for the house manager (receives overdue alerts).
 * Replace with the appropriate address for production use.
 */
const MANAGER_EMAIL = Session.getActiveUser().getEmail();

/**
 * This function runs automatically when the spreadsheet is opened.
 * It creates a custom menu in the Google Sheet UI, making it easier
 * to access the script's functions.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Parsonage Tools')
      .addItem('Initialize/Format Sheets', 'initializeSheets')
      .addSeparator()
      .addItem('Check Payment Status', 'checkAllPaymentStatus')
      .addItem('Send Rent Reminders', 'sendRentReminders')
      .addItem('Send Late Payment Alerts', 'sendLatePaymentAlerts')
      .addItem('Send Monthly Invoices', 'sendMonthlyInvoices')
      .addItem('Mark Selected Payment Received', 'markPaymentReceived')
      .addSeparator()
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
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#D9EAD3');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Adjust column widths for readability and aesthetics.
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
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
 * Checks payment status for all tenants and updates the sheet.
 * Determines whether each tenant is Paid, Due, or Overdue based on
 * the 'Last Payment Date' column.
 */
function checkAllPaymentStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, TENANTS_HEADERS.length);
  const data = dataRange.getValues();

  const today = new Date();
  const firstThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const firstLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);

  data.forEach(row => {
    const roomStatus = row[COL_ROOM_STATUS - 1];
    if (roomStatus !== 'Occupied') {
      row[COL_PAYMENT_STATUS - 1] = '';
      return;
    }

    const lastPayment = row[COL_LAST_PAYMENT - 1];
    let status = 'Overdue';

    if (lastPayment instanceof Date) {
      if (lastPayment >= firstThisMonth) {
        status = 'Paid';
      } else if (lastPayment >= firstLastMonth) {
        status = 'Due';
      }
    }

    row[COL_PAYMENT_STATUS - 1] = status;
  });

  dataRange.setValues(data);
}

/**
 * Sends rent reminder emails to tenants with Due or Overdue status.
 */
function sendRentReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No tenants found.');
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, TENANTS_HEADERS.length);
  const data = dataRange.getValues();
  const monthYear = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  let sent = 0;

  data.forEach(row => {
    const status = row[COL_PAYMENT_STATUS - 1];
    const email = row[COL_TENANT_EMAIL - 1];
    if ((status === 'Due' || status === 'Overdue') && email) {
      const tenantName = row[COL_TENANT_NAME - 1];
      const room = row[COL_ROOM_NUMBER - 1];
      const rent = row[COL_NEGOTIATED_PRICE - 1] || row[COL_RENTAL_PRICE - 1];
      const subject = `Rent Reminder - ${monthYear}`;
      const body = `Dear ${tenantName},\n\n` +
        `This is a friendly reminder that your rent of $${rent} for room ${room} is ${status.toLowerCase()} for ${monthYear}. ` +
        `Please make your payment as soon as possible.\n\nThank you,\nParsonage Management`;
      try {
        MailApp.sendEmail(email, subject, body);
        sent++;
      } catch (e) {
        console.error(`Failed to send reminder to ${email}: ${e.message}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert(`Rent reminders sent: ${sent}`);
}

/**
 * Sends an alert to the house manager listing tenants that are Overdue.
 */
function sendLatePaymentAlerts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, TENANTS_HEADERS.length).getValues();
  const overdueList = [];

  data.forEach(row => {
    if (row[COL_PAYMENT_STATUS - 1] === 'Overdue') {
      const tenant = row[COL_TENANT_NAME - 1];
      const email = row[COL_TENANT_EMAIL - 1];
      const room = row[COL_ROOM_NUMBER - 1];
      overdueList.push(`${tenant} (Room ${room}, ${email})`);
    }
  });

  if (overdueList.length === 0) return;

  const subject = 'Overdue Rent Alert';
  const body = 'The following tenants are overdue on rent:\n\n' + overdueList.join('\n') + '\n\nPlease follow up accordingly.';
  MailApp.sendEmail(MANAGER_EMAIL, subject, body);
}

/**
 * Generates PDF invoices for all occupied rooms and emails them to tenants.
 * Invoices are simple documents summarizing the amount due for the month.
 */
function sendMonthlyInvoices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, TENANTS_HEADERS.length).getValues();
  const monthYear = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');

  data.forEach(row => {
    const status = row[COL_ROOM_STATUS - 1];
    const email = row[COL_TENANT_EMAIL - 1];
    if (status === 'Occupied' && email) {
      const tenant = row[COL_TENANT_NAME - 1];
      const room = row[COL_ROOM_NUMBER - 1];
      const rent = row[COL_NEGOTIATED_PRICE - 1] || row[COL_RENTAL_PRICE - 1];

      const doc = DocumentApp.create(`Rent Invoice - ${tenant} - ${monthYear}`);
      const body = doc.getBody();
      body.appendParagraph(`Rent Invoice - ${monthYear}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(`Tenant: ${tenant}`);
      body.appendParagraph(`Room: ${room}`);
      body.appendParagraph(`Amount Due: $${rent}`);
      body.appendParagraph('\nPlease remit payment at your earliest convenience.');
      doc.saveAndClose();

      const pdf = doc.getAs('application/pdf');
      MailApp.sendEmail(email, `Rent Invoice - ${monthYear}`, `Please find attached your rent invoice for ${monthYear}.`, {attachments: [pdf]});
      DriveApp.getFileById(doc.getId()).setTrashed(true); // Clean up temporary doc
    }
  });
}

/**
 * Marks the currently selected tenant row as paid and logs the income
 * in the Budget sheet.
 */
function markPaymentReceived() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== TENANTS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select a row in the Tenants sheet.');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a tenant row.');
    return;
  }

  const tenantName = sheet.getRange(row, COL_TENANT_NAME).getValue();
  const rent = sheet.getRange(row, COL_NEGOTIATED_PRICE).getValue() || sheet.getRange(row, COL_RENTAL_PRICE).getValue();

  sheet.getRange(row, COL_LAST_PAYMENT).setValue(new Date());
  sheet.getRange(row, COL_PAYMENT_STATUS).setValue('Paid');

  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  if (budgetSheet) {
    budgetSheet.appendRow([new Date(), 'Rent Income', `Rent from ${tenantName}`, rent, 'Rent']);
  }
}

/**
 * A test function to demonstrate sending a rent reminder email to yourself.
 */
function sendRentRemindersTest() {
  const recipientEmail = Session.getActiveUser().getEmail();
  const subject = 'Rent Reminder - Parsonage (Test)';
  const body = 'This is a test rent reminder email.';

  try {
    MailApp.sendEmail(recipientEmail, subject, body);
    SpreadsheetApp.getUi().alert('Test Email Sent', `A test rent reminder email has been sent to ${recipientEmail}. Check your inbox.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Email Error', `Failed to send test email: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(`Error sending test email: ${e.message}`);
  }
}

/**
 * Triggered when the tenant application form is submitted.
 * Sends a welcome email with basic information.
 */
function onTenantApplicationSubmit(e) {
  const email = e.namedValues['Tenant Email'] ? e.namedValues['Tenant Email'][0] : '';
  const name = e.namedValues['Name'] ? e.namedValues['Name'][0] : 'Applicant';
  if (email) {
    const subject = 'Application Received';
    const body = `Dear ${name},\n\nThank you for your application. We will review it and get back to you shortly.\n\nRegards,\nParsonage Management`;
    MailApp.sendEmail(email, subject, body);
  }
}

/**
 * Triggered when the move-out form is submitted. Sends move-out instructions.
 */
function onMoveOutRequestSubmit(e) {
  const email = e.namedValues['Tenant Email'] ? e.namedValues['Tenant Email'][0] : '';
  const name = e.namedValues['Tenant Name'] ? e.namedValues['Tenant Name'][0] : 'Tenant';
  const moveOutDate = e.namedValues['Move-Out Date'] ? e.namedValues['Move-Out Date'][0] : '';
  if (email) {
    const subject = 'Move-Out Instructions';
    const body = `Hello ${name},\n\nWe have recorded your planned move-out date of ${moveOutDate}. Please ensure the room is left in good condition.\n\nThank you,\nParsonage Management`;
    MailApp.sendEmail(email, subject, body);
  }
}

// End of core.gs
