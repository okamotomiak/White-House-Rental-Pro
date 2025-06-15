// Complete Parsonage Tenant Management System - Google Apps Script
// This file contains all the code needed for the system

/**
 * Global constants for sheet names.
 * These are used throughout the script to refer to specific sheets.
 */
const TENANTS_SHEET_NAME = 'Tenants';
const BUDGET_SHEET_NAME = 'Budget';
const APPLICATION_SHEET_NAME = 'Tenant Applications';
const MOVEOUT_SHEET_NAME = 'Move-Out Requests';

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
      .addSubMenu(ui.createMenu('Payment Management')
          .addItem('Check Payment Status', 'checkAllPaymentStatus')
          .addItem('Send Rent Reminders', 'sendRentReminders')
          .addItem('Send Late Payment Alerts', 'sendLatePaymentAlerts')
          .addItem('Send Monthly Invoices', 'sendMonthlyInvoices')
          .addItem('Mark Selected Payment Received', 'markPaymentReceived'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Budget Analysis')
          .addItem('Generate Monthly Report', 'generateMonthlyReport')
          .addItem('Create Income/Expense Chart', 'createBudgetChart')
          .addItem('Calculate Occupancy Rate', 'calculateOccupancyRate')
          .addItem('View Financial Summary', 'showFinancialSummary'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Setup & Configuration')
          .addItem('Setup Triggers', 'setupTriggers')
          .addItem('Create Application Form', 'createApplicationForm')
          .addItem('Create Move-Out Form', 'createMoveOutForm')
          .addItem('Configure Email Templates', 'configureEmailTemplates'))
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
    
    // Add sample rooms to Tenants sheet
    addSampleRooms();
    
    // Add conditional formatting
    applyConditionalFormatting();

    ui.alert('Sheets Initialized', 'All sheets have been initialized and formatted successfully.', ui.ButtonSet.OK);

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
  headerRange.setBorder(true, true, true, true, false, false);
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

  // Optional: Set row heights for all rows for better spacing
  const numRows = sheet.getMaxRows() - 1;
  if (numRows > 0) {
    sheet.setRowHeights(2, numRows, 25);
  }
  
  // Format specific columns if it's the Budget sheet
  if (sheetName === BUDGET_SHEET_NAME) {
    // Format date column
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
    // Format amount column as currency
    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0.00');
  }
}

/**
 * Adds sample rooms to the Tenants sheet for initial setup
 */
function addSampleRooms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() > 1) return; // Don't add if data already exists
  
  const sampleRooms = [
    ['101', 800, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['102', 800, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['103', 900, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['104', 900, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['201', 850, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['202', 850, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['203', 950, '', '', '', '', '', 'Vacant', '', '', '', ''],
    ['204', 950, '', '', '', '', '', 'Vacant', '', '', '', '']
  ];
  
  sheet.getRange(2, 1, sampleRooms.length, sampleRooms[0].length).setValues(sampleRooms);
}

/**
 * Applies conditional formatting to the sheets
 */
function applyConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tenantsSheet = ss.getSheetByName(TENANTS_SHEET_NAME);
  
  if (tenantsSheet) {
    // Clear existing conditional formatting
    tenantsSheet.clearConditionalFormatRules();
    
    // Format Room Status column
    const statusRange = tenantsSheet.getRange(2, COL_ROOM_STATUS, tenantsSheet.getMaxRows() - 1, 1);
    
    // Occupied - Green
    const occupiedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Occupied')
      .setBackground('#D9EAD3')
      .setRanges([statusRange])
      .build();
    
    // Vacant - Yellow
    const vacantRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Vacant')
      .setBackground('#FFF2CC')
      .setRanges([statusRange])
      .build();
    
    // Payment Status column
    const paymentRange = tenantsSheet.getRange(2, COL_PAYMENT_STATUS, tenantsSheet.getMaxRows() - 1, 1);
    
    // Paid - Green
    const paidRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Paid')
      .setBackground('#D9EAD3')
      .setRanges([paymentRange])
      .build();
    
    // Due - Yellow
    const dueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Due')
      .setBackground('#FFF2CC')
      .setRanges([paymentRange])
      .build();
    
    // Overdue - Red
    const overdueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Overdue')
      .setBackground('#F4CCCC')
      .setRanges([paymentRange])
      .build();
    
    const rules = [occupiedRule, vacantRule, paidRule, dueRule, overdueRule];
    tenantsSheet.setConditionalFormatRules(rules);
  }
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
  
  SpreadsheetApp.getUi().alert('Payment Status Updated', 'All payment statuses have been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
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
      
      const emailTemplate = getEmailTemplate('rentReminder', {
        tenantName: tenantName,
        room: room,
        rent: rent,
        status: status,
        monthYear: monthYear
      });
      
      try {
        MailApp.sendEmail(email, emailTemplate.subject, emailTemplate.body);
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
      const lastPayment = row[COL_LAST_PAYMENT - 1];
      const lastPaymentStr = lastPayment ? Utilities.formatDate(lastPayment, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Never';
      overdueList.push(`• ${tenant} (Room ${room}, ${email}) - Last payment: ${lastPaymentStr}`);
    }
  });

  if (overdueList.length === 0) {
    SpreadsheetApp.getUi().alert('No overdue tenants found.');
    return;
  }

  const emailTemplate = getEmailTemplate('overdueAlert', {
    overdueList: overdueList.join('\n'),
    count: overdueList.length
  });
  
  MailApp.sendEmail(MANAGER_EMAIL, emailTemplate.subject, emailTemplate.body);
  SpreadsheetApp.getUi().alert('Late payment alert sent to manager.');
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
  let sent = 0;

  data.forEach(row => {
    const status = row[COL_ROOM_STATUS - 1];
    const email = row[COL_TENANT_EMAIL - 1];
    if (status === 'Occupied' && email) {
      const tenant = row[COL_TENANT_NAME - 1];
      const room = row[COL_ROOM_NUMBER - 1];
      const rent = row[COL_NEGOTIATED_PRICE - 1] || row[COL_RENTAL_PRICE - 1];

      try {
        const pdf = createInvoicePDF(tenant, room, rent, monthYear);
        const emailTemplate = getEmailTemplate('monthlyInvoice', {
          tenantName: tenant,
          monthYear: monthYear
        });
        
        MailApp.sendEmail(email, emailTemplate.subject, emailTemplate.body, {
          attachments: [pdf]
        });
        sent++;
      } catch (e) {
        console.error(`Failed to send invoice to ${email}: ${e.message}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert(`Monthly invoices sent: ${sent}`);
}

/**
 * Creates a PDF invoice for a tenant
 */
function createInvoicePDF(tenantName, room, rent, monthYear) {
  const doc = DocumentApp.create(`Rent Invoice - ${tenantName} - ${monthYear}`);
  const body = doc.getBody();
  
  // Clear default content
  body.clear();
  
  // Add header
  const header = body.appendParagraph('PARSONAGE RENTAL');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  // Add invoice title
  const title = body.appendParagraph(`Rent Invoice - ${monthYear}`);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  
  // Add invoice details
  body.appendParagraph(`Date: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`);
  body.appendParagraph(`Tenant: ${tenantName}`);
  body.appendParagraph(`Room Number: ${room}`);
  body.appendParagraph('');
  
  // Add amount section
  const amountSection = body.appendParagraph(`Amount Due: $${rent}`);
  amountSection.setFontSize(14);
  amountSection.setBold(true);
  body.appendParagraph('');
  
  // Add payment instructions
  body.appendParagraph('Payment Instructions:');
  body.appendParagraph('• Payment is due by the 5th of each month');
  body.appendParagraph('• Late payments may incur additional fees');
  body.appendParagraph('• Please include your room number with payment');
  body.appendParagraph('');
  
  // Add footer
  body.appendParagraph('Thank you for your prompt payment!');
  body.appendParagraph('Parsonage Management');
  
  doc.saveAndClose();
  
  const pdf = doc.getAs('application/pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true); // Clean up temporary doc
  
  return pdf;
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
  const room = sheet.getRange(row, COL_ROOM_NUMBER).getValue();
  const rent = sheet.getRange(row, COL_NEGOTIATED_PRICE).getValue() || sheet.getRange(row, COL_RENTAL_PRICE).getValue();

  sheet.getRange(row, COL_LAST_PAYMENT).setValue(new Date());
  sheet.getRange(row, COL_PAYMENT_STATUS).setValue('Paid');

  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  if (budgetSheet) {
    budgetSheet.appendRow([
      new Date(), 
      'Rent Income', 
      `Rent from ${tenantName} - Room ${room}`, 
      rent, 
      'Rent'
    ]);
  }
  
  SpreadsheetApp.getUi().alert('Payment Recorded', `Payment from ${tenantName} has been recorded.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Generates a monthly financial report
 */
function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  
  if (!budgetSheet || budgetSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No budget data found.');
    return;
  }
  
  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  
  const data = budgetSheet.getRange(2, 1, budgetSheet.getLastRow() - 1, 5).getValues();
  
  let totalIncome = 0;
  let totalExpenses = 0;
  const incomeByCategory = {};
  const expensesByCategory = {};
  
  data.forEach(row => {
    const date = row[0];
    const amount = row[3];
    const category = row[4];
    
    if (date >= firstOfMonth && date <= lastOfMonth) {
      if (amount > 0) {
        totalIncome += amount;
        incomeByCategory[category] = (incomeByCategory[category] || 0) + amount;
      } else {
        totalExpenses += Math.abs(amount);
        expensesByCategory[category] = (expensesByCategory[category] || 0) + Math.abs(amount);
      }
    }
  });
  
  const monthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');
  
  let report = `Monthly Financial Report - ${monthYear}\n`;
  report += `=====================================\n\n`;
  report += `Total Income: $${totalIncome.toFixed(2)}\n`;
  report += `Total Expenses: $${totalExpenses.toFixed(2)}\n`;
  report += `Net Profit: $${(totalIncome - totalExpenses).toFixed(2)}\n\n`;
  
  report += `Income by Category:\n`;
  Object.entries(incomeByCategory).forEach(([cat, amt]) => {
    report += `  ${cat}: $${amt.toFixed(2)}\n`;
  });
  
  report += `\nExpenses by Category:\n`;
  Object.entries(expensesByCategory).forEach(([cat, amt]) => {
    report += `  ${cat}: $${amt.toFixed(2)}\n`;
  });
  
  // Show report in a modal dialog
  const htmlOutput = HtmlService
      .createHtmlOutput(`<pre>${report}</pre>`)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Monthly Financial Report');
}

/**
 * Creates a visual chart of income vs expenses
 */
function createBudgetChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  
  if (!budgetSheet || budgetSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No budget data found.');
    return;
  }
  
  // Create or get chart sheet
  let chartSheet = ss.getSheetByName('Budget Charts');
  if (!chartSheet) {
    chartSheet = ss.insertSheet('Budget Charts');
  }
  
  // Prepare data for chart
  const sourceData = budgetSheet.getRange(1, 1, budgetSheet.getLastRow(), 5);
  
  // Create chart
  const chart = chartSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sourceData)
      .setPosition(1, 1, 0, 0)
      .setOption('title', 'Income vs Expenses')
      .setOption('width', 600)
      .setOption('height', 400)
      .build();
  
  chartSheet.insertChart(chart);
  
  SpreadsheetApp.getUi().alert('Chart Created', 'Budget chart has been created in the "Budget Charts" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Calculates and displays the current occupancy rate
 */
function calculateOccupancyRate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No room data found.');
    return;
  }
  
  const data = sheet.getRange(2, COL_ROOM_STATUS, sheet.getLastRow() - 1, 1).getValues();
  let totalRooms = 0;
  let occupiedRooms = 0;
  
  data.forEach(row => {
    if (row[0]) {
      totalRooms++;
      if (row[0] === 'Occupied') {
        occupiedRooms++;
      }
    }
  });
  
  const occupancyRate = totalRooms > 0 ? (occupiedRooms / totalRooms * 100).toFixed(1) : 0;
  
  SpreadsheetApp.getUi().alert(
    'Occupancy Rate',
    `Current Occupancy: ${occupiedRooms}/${totalRooms} rooms (${occupancyRate}%)`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows a comprehensive financial summary
 */
function showFinancialSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tenantsSheet = ss.getSheetByName(TENANTS_SHEET_NAME);
  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  
  if (!tenantsSheet || !budgetSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  // Calculate potential monthly income
  const tenantsData = tenantsSheet.getRange(2, 1, tenantsSheet.getLastRow() - 1, TENANTS_HEADERS.length).getValues();
  let potentialIncome = 0;
  let actualIncome = 0;
  
  tenantsData.forEach(row => {
    const rent = row[COL_NEGOTIATED_PRICE - 1] || row[COL_RENTAL_PRICE - 1];
    if (rent) {
      potentialIncome += rent;
      if (row[COL_ROOM_STATUS - 1] === 'Occupied') {
        actualIncome += rent;
      }
    }
  });
  
  // Get YTD financial data
  const budgetData = budgetSheet.getRange(2, 1, budgetSheet.getLastRow() - 1, 5).getValues();
  const yearStart = new Date(new Date().getFullYear(), 0, 1);
  let ytdIncome = 0;
  let ytdExpenses = 0;
  
  budgetData.forEach(row => {
    if (row[0] >= yearStart) {
      if (row[3] > 0) {
        ytdIncome += row[3];
      } else {
        ytdExpenses += Math.abs(row[3]);
      }
    }
  });
  
  const summary = `
    <h3>Financial Summary</h3>
    <hr>
    <h4>Monthly Projections:</h4>
    <p>Potential Income (100% occupancy): $${potentialIncome.toFixed(2)}</p>
    <p>Expected Income (current occupancy): $${actualIncome.toFixed(2)}</p>
    <hr>
    <h4>Year-to-Date (YTD):</h4>
    <p>Total Income: $${ytdIncome.toFixed(2)}</p>
    <p>Total Expenses: $${ytdExpenses.toFixed(2)}</p>
    <p>Net Profit: $${(ytdIncome - ytdExpenses).toFixed(2)}</p>
  `;
  
  const htmlOutput = HtmlService
      .createHtmlOutput(summary)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Financial Summary');
}

/**
 * Sets up automated triggers for daily checks and form submissions
 */
function setupTriggers() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Daily payment status check at 9 AM
  ScriptApp.newTrigger('checkAllPaymentStatus')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
  
  // Monthly reminder on the 1st at 10 AM
  ScriptApp.newTrigger('sendRentReminders')
      .timeBased()
      .onMonthDay(1)
      .atHour(10)
      .create();
  
  // Weekly overdue alert on Mondays at 2 PM
  ScriptApp.newTrigger('sendLatePaymentAlerts')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(14)
      .create();
  
  SpreadsheetApp.getUi().alert('Triggers Set Up', 'Automated triggers have been configured successfully.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates instructions for setting up the Application Form
 */
function createApplicationForm() {
  const instructions = `
    <h3>Create Tenant Application Form</h3>
    <p>Follow these steps to create your tenant application form:</p>
    <ol>
      <li>Go to <a href="https://forms.google.com" target="_blank">Google Forms</a></li>
      <li>Create a new form titled "Parsonage Tenant Application"</li>
      <li>Add these fields:
        <ul>
          <li>Name (Short answer, Required)</li>
          <li>Email (Short answer, Required, Validate as email)</li>
          <li>Phone Number (Short answer, Required)</li>
          <li>Desired Move-in Date (Date, Required)</li>
          <li>Preferred Room (Multiple choice: List available rooms)</li>
          <li>Employment Status (Short answer)</li>
          <li>References (Paragraph)</li>
          <li>Additional Information (Paragraph)</li>
          <li>Proof of Income (File upload)</li>
        </ul>
      </li>
      <li>In Form Settings > Responses, link to this spreadsheet</li>
      <li>Create a new sheet named "${APPLICATION_SHEET_NAME}"</li>
      <li>Set up a form submit trigger for onTenantApplicationSubmit</li>
    </ol>
  `;
  
  const htmlOutput = HtmlService
      .createHtmlOutput(instructions)
      .setWidth(500)
      .setHeight(400);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Application Form Setup');
}

/**
 * Creates instructions for setting up the Move-Out Form
 */
function createMoveOutForm() {
  const instructions = `
    <h3>Create Move-Out Request Form</h3>
    <p>Follow these steps to create your move-out request form:</p>
    <ol>
      <li>Go to <a href="https://forms.google.com" target="_blank">Google Forms</a></li>
      <li>Create a new form titled "Parsonage Move-Out Request"</li>
      <li>Add these fields:
        <ul>
          <li>Tenant Name (Short answer, Required)</li>
          <li>Tenant Email (Short answer, Required, Validate as email)</li>
          <li>Room Number (Short answer, Required)</li>
          <li>Planned Move-Out Date (Date, Required)</li>
          <li>Forwarding Address (Paragraph)</li>
          <li>Reason for Moving (Multiple choice or Paragraph)</li>
          <li>Feedback/Comments (Paragraph)</li>
        </ul>
      </li>
      <li>In Form Settings > Responses, link to this spreadsheet</li>
      <li>Create a new sheet named "${MOVEOUT_SHEET_NAME}"</li>
      <li>Set up a form submit trigger for onMoveOutRequestSubmit</li>
    </ol>
  `;
  
  const htmlOutput = HtmlService
      .createHtmlOutput(instructions)
      .setWidth(500)
      .setHeight(400);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Move-Out Form Setup');
}

/**
 * Email template management
 */
function getEmailTemplate(type, data) {
  const templates = {
    rentReminder: {
      subject: `Rent Reminder - ${data.monthYear}`,
      body: `Dear ${data.tenantName},

This is a friendly reminder that your rent of $${data.rent} for room ${data.room} is ${data.status.toLowerCase()} for ${data.monthYear}.

Please make your payment as soon as possible to avoid any late fees.

Payment can be made via:
• Bank transfer
• Check
• Cash (with receipt)

If you have already made your payment, please disregard this message and let us know so we can update our records.

Thank you for your cooperation.

Best regards,
Parsonage Management`
    },
    overdueAlert: {
      subject: `Overdue Rent Alert - ${data.count} Tenant(s)`,
      body: `Hello,

The following ${data.count} tenant(s) have overdue rent payments:

${data.overdueList}

Please follow up with these tenants as soon as possible. Consider:
• Personal follow-up calls
• Written notices if significantly overdue
• Review of payment plans if needed

This is an automated alert from the Parsonage Management System.

Best regards,
Parsonage Management System`
    },
    monthlyInvoice: {
      subject: `Rent Invoice - ${data.monthYear}`,
      body: `Dear ${data.tenantName},

Please find attached your rent invoice for ${data.monthYear}.

Payment is due by the 5th of the month. Thank you for your prompt payment.

If you have any questions about your invoice, please don't hesitate to contact us.

Best regards,
Parsonage Management`
    },
    applicationReceived: {
      subject: 'Application Received - Parsonage',
      body: `Dear ${data.name},

Thank you for your application to rent at our parsonage. We have received your application and will review it shortly.

What happens next:
1. We will review your application and references
2. We may contact you for additional information or to schedule a viewing
3. You will receive a response within 3-5 business days

House Rules & Cultural Vision:
• We maintain a quiet, respectful living environment
• Common areas should be kept clean and tidy
• Guest policies and quiet hours are enforced
• We foster a community atmosphere with optional shared activities

If you have any questions, please feel free to contact us.

Best regards,
Parsonage Management`
    },
    moveOutInstructions: {
      subject: 'Move-Out Instructions - Parsonage',
      body: `Hello ${data.name},

We have received your move-out request for ${data.moveOutDate}.

Move-Out Checklist:
1. Clean your room thoroughly (including windows, floors, and closets)
2. Remove all personal belongings
3. Return all keys and access cards
4. Schedule a move-out inspection
5. Provide forwarding address for security deposit return

Your security deposit will be returned within 30 days after move-out, minus any deductions for damages or cleaning.

Please contact us to schedule your move-out inspection at least 3 days before your move-out date.

Thank you for being a valued tenant.

Best regards,
Parsonage Management`
    }
  };
  
  return templates[type] || { subject: 'Parsonage Management', body: 'No template found.' };
}

/**
 * Configure email templates (shows current templates)
 */
function configureEmailTemplates() {
  const html = `
    <h3>Email Templates</h3>
    <p>The system uses the following email templates:</p>
    <ul>
      <li><b>Rent Reminder:</b> Sent monthly to tenants with due/overdue rent</li>
      <li><b>Overdue Alert:</b> Sent to manager about overdue tenants</li>
      <li><b>Monthly Invoice:</b> Accompanies PDF invoices</li>
      <li><b>Application Received:</b> Auto-response for new applications</li>
      <li><b>Move-Out Instructions:</b> Sent when move-out form submitted</li>
    </ul>
    <p>To modify templates, edit the getEmailTemplate() function in the script.</p>
  `;
  
  const htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Email Templates');
}

/**
 * A test function to demonstrate sending a rent reminder email to yourself.
 */
function sendRentRemindersTest() {
  const recipientEmail = Session.getActiveUser().getEmail();
  const testData = {
    tenantName: 'Test Tenant',
    room: '101',
    rent: '800',
    status: 'Due',
    monthYear: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy')
  };
  
  const emailTemplate = getEmailTemplate('rentReminder', testData);

  try {
    MailApp.sendEmail(recipientEmail, emailTemplate.subject + ' (TEST)', emailTemplate.body);
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
  if (!e || !e.namedValues) return;
  
  const email = e.namedValues['Email'] ? e.namedValues['Email'][0] : '';
  const name = e.namedValues['Name'] ? e.namedValues['Name'][0] : 'Applicant';
  
  if (email) {
    const emailTemplate = getEmailTemplate('applicationReceived', { name: name });
    try {
      MailApp.sendEmail(email, emailTemplate.subject, emailTemplate.body);
      console.log(`Application confirmation sent to ${email}`);
    } catch (error) {
      console.error(`Failed to send application confirmation: ${error.message}`);
    }
  }
}

/**
 * Triggered when the move-out form is submitted. Sends move-out instructions.
 */
function onMoveOutRequestSubmit(e) {
  if (!e || !e.namedValues) return;
  
  const email = e.namedValues['Tenant Email'] ? e.namedValues['Tenant Email'][0] : '';
  const name = e.namedValues['Tenant Name'] ? e.namedValues['Tenant Name'][0] : 'Tenant';
  const moveOutDate = e.namedValues['Planned Move-Out Date'] ? e.namedValues['Planned Move-Out Date'][0] : '';
  const room = e.namedValues['Room Number'] ? e.namedValues['Room Number'][0] : '';
  
  if (email) {
    const emailTemplate = getEmailTemplate('moveOutInstructions', {
      name: name,
      moveOutDate: moveOutDate
    });
    
    try {
      MailApp.sendEmail(email, emailTemplate.subject, emailTemplate.body);
      console.log(`Move-out instructions sent to ${email}`);
      
      // Update the tenant record in the main sheet
      updateTenantMoveOut(room, moveOutDate);
    } catch (error) {
      console.error(`Failed to send move-out instructions: ${error.message}`);
    }
  }
}

/**
 * Updates tenant record with move-out date
 */
function updateTenantMoveOut(roomNumber, moveOutDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TENANTS_SHEET_NAME);
  if (!sheet) return;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, TENANTS_HEADERS.length).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][COL_ROOM_NUMBER - 1] == roomNumber) {
      sheet.getRange(i + 2, COL_MOVE_OUT_PLANNED).setValue(moveOutDate);
      sheet.getRange(i + 2, COL_NOTES).setValue(`Move-out requested on ${new Date().toLocaleDateString()}`);
      break;
    }
  }
}

// End of Parsonage Tenant Management System
