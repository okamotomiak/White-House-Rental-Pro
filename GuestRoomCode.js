// GuestRoomCode.gs
/**
 * Guest Room Management Enhancement for Parsonage Tenant Management System
 * This code adds short-term rental functionality to the existing system
 * Add this code to your existing Google Apps Script
 */

// Additional constants for guest room management
const GUEST_BOOKINGS_SHEET_NAME = 'Guest Bookings';
const GUEST_ROOMS_SHEET_NAME = 'Guest Rooms';
const GUEST_BOOKING_FORM_NAME = 'Guest Room Booking Request';

// Guest booking status options
const BOOKING_STATUS = {
  PENDING: 'Pending',
  CONFIRMED: 'Confirmed',
  CHECKED_IN: 'Checked In',
  CHECKED_OUT: 'Checked Out',
  CANCELLED: 'Cancelled'
};

/**
 * Headers for the 'Guest Rooms' sheet
 */
const GUEST_ROOMS_HEADERS = [
  'Room Number',
  'Room Name',
  'Daily Rate',
  'Weekly Rate',
  'Room Type',
  'Max Occupancy',
  'Amenities',
  'Status', // Available, Occupied, Maintenance
  'Current Guest',
  'Check-In Date',
  'Check-Out Date',
  'Notes'
];

/**
 * Headers for the 'Guest Bookings' sheet
 */
const GUEST_BOOKINGS_HEADERS = [
  'Booking ID',
  'Guest Name',
  'Guest Email',
  'Guest Phone',
  'Room Number',
  'Check-In Date',
  'Check-Out Date',
  'Number of Nights',
  'Number of Guests',
  'Total Amount',
  'Amount Paid',
  'Payment Status',
  'Booking Status',
  'Special Requests',
  'Booking Date',
  'Notes'
];

/**
 * Enhanced onOpen function with guest room menu items
 */
function enhancedOnOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Parsonage Tools')
      .addItem('Initialize/Format Sheets', 'initializeAllSheets')
      .addSeparator()
      .addSubMenu(ui.createMenu('Tenant Management')
          .addItem('Check Payment Status', 'checkAllPaymentStatus')
          .addItem('Send Rent Reminders', 'sendRentReminders')
          .addItem('Send Late Payment Alerts', 'sendLatePaymentAlerts')
          .addItem('Send Monthly Invoices', 'sendMonthlyInvoices')
          .addItem('Mark Selected Payment Received', 'markPaymentReceived'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Guest Room Management')
          .addItem('View Today\'s Arrivals', 'showTodayArrivals')
          .addItem('View Today\'s Departures', 'showTodayDepartures')
          .addItem('Check Room Availability', 'checkGuestRoomAvailability')
          .addItem('Process Check-In', 'processCheckIn')
          .addItem('Process Check-Out', 'processCheckOut')
          .addItem('Send Booking Confirmation', 'sendBookingConfirmation')
          .addItem('Generate Guest Invoice', 'generateGuestInvoice')
          .addItem('View Occupancy Calendar', 'showOccupancyCalendar'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Budget & Reports')
          .addItem('Generate Monthly Report', 'generateMonthlyReport')
          .addItem('Guest Room Revenue Report', 'generateGuestRoomReport')
          .addItem('Combined Revenue Analysis', 'combinedRevenueAnalysis')
          .addItem('Create Income/Expense Chart', 'createBudgetChart')
          .addItem('Calculate Total Occupancy Rate', 'calculateTotalOccupancyRate'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Setup & Configuration')
          .addItem('Setup All Triggers', 'setupAllTriggers')
          .addItem('Auto-Create All Forms', 'autoCreateAllFormsEnhanced')
          .addItem('View Form URLs', 'showFormURLs')
          .addItem('Configure Email Templates', 'configureEmailTemplates'))
      .addSeparator()
      .addItem('Send Test Email', 'sendTestEmail')
      .addToUi();
}

/**
 * Initialize all sheets including guest room sheets
 */
function initializeAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Initialize original sheets
    setupSheet(ss, TENANTS_SHEET_NAME, TENANTS_HEADERS);
    setupSheet(ss, BUDGET_SHEET_NAME, BUDGET_HEADERS);
    
    // Initialize guest room sheets
    setupSheet(ss, GUEST_ROOMS_SHEET_NAME, GUEST_ROOMS_HEADERS);
    setupSheet(ss, GUEST_BOOKINGS_SHEET_NAME, GUEST_BOOKINGS_HEADERS);
    
    // Add sample data
    addSampleRooms();
    addSampleGuestRooms();
    
    // Apply formatting
    applyConditionalFormatting();
    applyGuestRoomFormatting();

    ui.alert('All Sheets Initialized', 'All sheets have been initialized and formatted successfully, including guest room management sheets.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', `Failed to initialize sheets: ${e.message}`, ui.ButtonSet.OK);
    console.error(`Error initializing sheets: ${e.message}`);
  }
}

/**
 * Add sample guest rooms
 */
function addSampleGuestRooms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_ROOMS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() > 1) return;
  
  const sampleGuestRooms = [
    ['G1', 'Guest Suite 1', 75, 450, 'Guest Room', 2, 'Queen bed, Private bath, Mini fridge', 'Available', '', '', '', ''],
    ['G2', 'Guest Suite 2', 65, 390, 'Guest Room', 2, 'Double bed, Shared bath', 'Available', '', '', '', 'Will be available after renovation']
  ];
  
  sheet.getRange(2, 1, sampleGuestRooms.length, sampleGuestRooms[0].length).setValues(sampleGuestRooms);
}

/**
 * Apply conditional formatting to guest room sheets
 */
function applyGuestRoomFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const guestRoomsSheet = ss.getSheetByName(GUEST_ROOMS_SHEET_NAME);
  const bookingsSheet = ss.getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  
  if (guestRoomsSheet) {
    // Clear existing rules
    guestRoomsSheet.clearConditionalFormatRules();
    
    // Status formatting
    const statusRange = guestRoomsSheet.getRange(2, 8, guestRoomsSheet.getMaxRows() - 1, 1);
    
    const availableRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Available')
      .setBackground('#D9EAD3')
      .setRanges([statusRange])
      .build();
    
    const occupiedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Occupied')
      .setBackground('#F4CCCC')
      .setRanges([statusRange])
      .build();
    
    guestRoomsSheet.setConditionalFormatRules([availableRule, occupiedRule]);
  }
  
  if (bookingsSheet) {
    // Booking status formatting
    const bookingStatusRange = bookingsSheet.getRange(2, 13, bookingsSheet.getMaxRows() - 1, 1);
    
    const confirmedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Confirmed')
      .setBackground('#D9EAD3')
      .setRanges([bookingStatusRange])
      .build();
    
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground('#FFF2CC')
      .setRanges([bookingStatusRange])
      .build();
    
    const checkedInRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Checked In')
      .setBackground('#CFE2F3')
      .setRanges([bookingStatusRange])
      .build();
    
    bookingsSheet.setConditionalFormatRules([confirmedRule, pendingRule, checkedInRule]);
  }
}

/**
 * Show today's arrivals
 */
function showTodayArrivals() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No bookings found.');
    return;
  }
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
  const arrivals = [];
  
  data.forEach(row => {
    const checkIn = new Date(row[5]);
    checkIn.setHours(0, 0, 0, 0);
    
    if (checkIn.getTime() === today.getTime() && row[12] === BOOKING_STATUS.CONFIRMED) {
      arrivals.push(`${row[1]} - Room ${row[4]} (${row[8]} guests)`);
    }
  });
  
  const message = arrivals.length > 0 
    ? `Today's Arrivals:\n\n${arrivals.join('\n')}` 
    : 'No arrivals scheduled for today.';
  
  SpreadsheetApp.getUi().alert('Today\'s Arrivals', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Show today's departures
 */
function showTodayDepartures() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No bookings found.');
    return;
  }
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
  const departures = [];
  
  data.forEach(row => {
    const checkOut = new Date(row[6]);
    checkOut.setHours(0, 0, 0, 0);
    
    if (checkOut.getTime() === today.getTime() && row[12] === BOOKING_STATUS.CHECKED_IN) {
      departures.push(`${row[1]} - Room ${row[4]} (Balance: $${row[9] - row[10]})`);
    }
  });
  
  const message = departures.length > 0 
    ? `Today's Departures:\n\n${departures.join('\n')}` 
    : 'No departures scheduled for today.';
  
  SpreadsheetApp.getUi().alert('Today\'s Departures', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Check guest room availability for a date range
 */
function checkGuestRoomAvailability() {
  const ui = SpreadsheetApp.getUi();
  
  // Get check-in date
  const checkInResponse = ui.prompt(
    'Check Availability',
    'Enter check-in date (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (checkInResponse.getSelectedButton() !== ui.Button.OK) return;
  
  // Get check-out date
  const checkOutResponse = ui.prompt(
    'Check Availability',
    'Enter check-out date (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (checkOutResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const checkIn = new Date(checkInResponse.getResponseText());
  const checkOut = new Date(checkOutResponse.getResponseText());
  
  if (isNaN(checkIn.getTime()) || isNaN(checkOut.getTime())) {
    ui.alert('Invalid dates entered.');
    return;
  }
  
  const availableRooms = getAvailableGuestRooms(checkIn, checkOut);
  
  if (availableRooms.length === 0) {
    ui.alert('No Rooms Available', 'No guest rooms are available for the selected dates.', ui.ButtonSet.OK);
  } else {
    const roomList = availableRooms.map(room => 
      `${room.name} (${room.number}) - $${room.dailyRate}/night`
    ).join('\n');
    
    ui.alert('Available Rooms', `Available rooms for ${checkInResponse.getResponseText()} to ${checkOutResponse.getResponseText()}:\n\n${roomList}`, ui.ButtonSet.OK);
  }
}

/**
 * Get available guest rooms for a date range
 */
function getAvailableGuestRooms(checkIn, checkOut) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roomsSheet = ss.getSheetByName(GUEST_ROOMS_SHEET_NAME);
  const bookingsSheet = ss.getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  
  if (!roomsSheet || !bookingsSheet) return [];
  
  // Get all guest rooms
  const roomsData = roomsSheet.getRange(2, 1, roomsSheet.getLastRow() - 1, GUEST_ROOMS_HEADERS.length).getValues();
  const bookingsData = bookingsSheet.getRange(2, 1, bookingsSheet.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
  
  const availableRooms = [];
  
  roomsData.forEach(room => {
    if (!room[0]) return;
    
    let isAvailable = true;
    
    // Check if room has any overlapping bookings
    bookingsData.forEach(booking => {
      if (booking[4] === room[0] && 
          (booking[12] === BOOKING_STATUS.CONFIRMED || booking[12] === BOOKING_STATUS.CHECKED_IN)) {
        
        const bookingCheckIn = new Date(booking[5]);
        const bookingCheckOut = new Date(booking[6]);
        
        // Check for date overlap
        if (!(checkOut <= bookingCheckIn || checkIn >= bookingCheckOut)) {
          isAvailable = false;
        }
      }
    });
    
    if (isAvailable && room[7] !== 'Maintenance') {
      availableRooms.push({
        number: room[0],
        name: room[1],
        dailyRate: room[2],
        weeklyRate: room[3],
        maxOccupancy: room[5],
        amenities: room[6]
      });
    }
  });
  
  return availableRooms;
}

/**
 * Process guest check-in
 */
function processCheckIn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getName() !== GUEST_BOOKINGS_SHEET_NAME) {
    ui.alert('Please select a booking in the Guest Bookings sheet.');
    return;
  }
  
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a booking row.');
    return;
  }
  
  const bookingStatus = sheet.getRange(row, 13).getValue();
  if (bookingStatus !== BOOKING_STATUS.CONFIRMED) {
    ui.alert('Only confirmed bookings can be checked in.');
    return;
  }
  
  // Update booking status
  sheet.getRange(row, 13).setValue(BOOKING_STATUS.CHECKED_IN);
  
  // Update guest room status
  const roomNumber = sheet.getRange(row, 5).getValue();
  const guestName = sheet.getRange(row, 2).getValue();
  const checkOut = sheet.getRange(row, 7).getValue();
  
  updateGuestRoomStatus(roomNumber, 'Occupied', guestName, new Date(), checkOut);
  
  ui.alert('Check-In Complete', `${guestName} has been checked into room ${roomNumber}.`, ui.ButtonSet.OK);
}

/**
 * Process guest check-out
 */
function processCheckOut() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== GUEST_BOOKINGS_SHEET_NAME) {
    ui.alert('Please select a booking in the Guest Bookings sheet.');
    return;
  }
  
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a booking row.');
    return;
  }
  
  const bookingStatus = sheet.getRange(row, 13).getValue();
  if (bookingStatus !== BOOKING_STATUS.CHECKED_IN) {
    ui.alert('Only checked-in bookings can be checked out.');
    return;
  }
  
  // Check payment status
  const totalAmount = sheet.getRange(row, 10).getValue();
  const amountPaid = sheet.getRange(row, 11).getValue();
  
  if (amountPaid < totalAmount) {
    const response = ui.alert(
      'Outstanding Balance',
      `There is an outstanding balance of $${totalAmount - amountPaid}. Proceed with check-out?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
  }
  
  // Update booking status
  sheet.getRange(row, 13).setValue(BOOKING_STATUS.CHECKED_OUT);
  
  // Update guest room status
  const roomNumber = sheet.getRange(row, 5).getValue();
  updateGuestRoomStatus(roomNumber, 'Available', '', '', '');
  
  // Log revenue
  const guestName = sheet.getRange(row, 2).getValue();
  logGuestRoomRevenue(guestName, roomNumber, amountPaid);
  
  ui.alert('Check-Out Complete', `Check-out completed for room ${roomNumber}.`, ui.ButtonSet.OK);
}

/**
 * Update guest room status
 */
function updateGuestRoomStatus(roomNumber, status, guestName, checkIn, checkOut) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_ROOMS_SHEET_NAME);
  if (!sheet) return;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === roomNumber) {
      const row = i + 2;
      sheet.getRange(row, 8).setValue(status);
      sheet.getRange(row, 9).setValue(guestName || '');
      sheet.getRange(row, 10).setValue(checkIn || '');
      sheet.getRange(row, 11).setValue(checkOut || '');
      break;
    }
  }
}

/**
 * Log guest room revenue to budget sheet
 */
function logGuestRoomRevenue(guestName, roomNumber, amount) {
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BUDGET_SHEET_NAME);
  if (!budgetSheet) return;
  
  budgetSheet.appendRow([
    new Date(),
    'Guest Room Income',
    `Guest room rental - ${guestName} (Room ${roomNumber})`,
    amount,
    'Guest Room'
  ]);
}

/**
 * Send booking confirmation email
 */
function sendBookingConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== GUEST_BOOKINGS_SHEET_NAME) {
    ui.alert('Please select a booking in the Guest Bookings sheet.');
    return;
  }
  
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a booking row.');
    return;
  }
  
  const bookingData = sheet.getRange(row, 1, 1, GUEST_BOOKINGS_HEADERS.length).getValues()[0];
  
  const emailData = {
    bookingId: bookingData[0],
    guestName: bookingData[1],
    guestEmail: bookingData[2],
    roomNumber: bookingData[4],
    checkIn: Utilities.formatDate(new Date(bookingData[5]), Session.getScriptTimeZone(), 'MMMM dd, yyyy'),
    checkOut: Utilities.formatDate(new Date(bookingData[6]), Session.getScriptTimeZone(), 'MMMM dd, yyyy'),
    nights: bookingData[7],
    totalAmount: bookingData[9],
    specialRequests: bookingData[13]
  };
  
  const emailTemplate = getGuestEmailTemplate('bookingConfirmation', emailData);
  
  try {
    MailApp.sendEmail(emailData.guestEmail, emailTemplate.subject, emailTemplate.body);
    
    // Update booking status to confirmed if it was pending
    if (bookingData[12] === BOOKING_STATUS.PENDING) {
      sheet.getRange(row, 13).setValue(BOOKING_STATUS.CONFIRMED);
    }
    
    ui.alert('Email Sent', 'Booking confirmation has been sent to ' + emailData.guestEmail, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Generate guest invoice
 */
function generateGuestInvoice() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== GUEST_BOOKINGS_SHEET_NAME) {
    ui.alert('Please select a booking in the Guest Bookings sheet.');
    return;
  }
  
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a booking row.');
    return;
  }
  
  const bookingData = sheet.getRange(row, 1, 1, GUEST_BOOKINGS_HEADERS.length).getValues()[0];
  
  const doc = DocumentApp.create(`Guest Invoice - ${bookingData[1]} - ${bookingData[0]}`);
  const body = doc.getBody();
  
  body.clear();
  
  // Header
  body.appendParagraph('PARSONAGE GUEST ACCOMMODATION')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('Guest Invoice')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  // Invoice details
  body.appendParagraph(`Invoice Date: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`);
  body.appendParagraph(`Booking ID: ${bookingData[0]}`);
  body.appendParagraph('');
  
  // Guest details
  body.appendParagraph('Guest Information:').setBold(true);
  body.appendParagraph(`Name: ${bookingData[1]}`);
  body.appendParagraph(`Email: ${bookingData[2]}`);
  body.appendParagraph(`Phone: ${bookingData[3]}`);
  body.appendParagraph('');
  
  // Booking details
  body.appendParagraph('Booking Details:').setBold(true);
  body.appendParagraph(`Room: ${bookingData[4]}`);
  body.appendParagraph(`Check-in: ${Utilities.formatDate(new Date(bookingData[5]), Session.getScriptTimeZone(), 'MMMM dd, yyyy')}`);
  body.appendParagraph(`Check-out: ${Utilities.formatDate(new Date(bookingData[6]), Session.getScriptTimeZone(), 'MMMM dd, yyyy')}`);
  body.appendParagraph(`Number of nights: ${bookingData[7]}`);
  body.appendParagraph(`Number of guests: ${bookingData[8]}`);
  body.appendParagraph('');
  
  // Payment details
  body.appendParagraph('Payment Summary:').setBold(true);
  body.appendParagraph(`Total Amount: $${bookingData[9]}`);
  body.appendParagraph(`Amount Paid: $${bookingData[10]}`);
  body.appendParagraph(`Balance Due: $${bookingData[9] - bookingData[10]}`);
  
  doc.saveAndClose();
  
  const pdf = doc.getAs('application/pdf');
  
  // Send invoice
  const emailTemplate = getGuestEmailTemplate('invoice', {
    guestName: bookingData[1],
    bookingId: bookingData[0]
  });
  
  try {
    MailApp.sendEmail(bookingData[2], emailTemplate.subject, emailTemplate.body, {
      attachments: [pdf]
    });
    
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    
    ui.alert('Invoice Sent', 'Invoice has been sent to ' + bookingData[2], ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to send invoice: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Generate guest room revenue report
 */
function generateGuestRoomReport() {
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BUDGET_SHEET_NAME);
  
  if (!budgetSheet || budgetSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No financial data found.');
    return;
  }
  
  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  
  const data = budgetSheet.getRange(2, 1, budgetSheet.getLastRow() - 1, 5).getValues();
  
  let guestRoomRevenue = 0;
  let bookingCount = 0;
  const revenueByRoom = {};
  
  data.forEach(row => {
    const date = row[0];
    const category = row[4];
    const amount = row[3];
    const description = row[2];
    
    if (date >= firstOfMonth && date <= lastOfMonth && category === 'Guest Room' && amount > 0) {
      guestRoomRevenue += amount;
      bookingCount++;
      
      // Extract room number from description
      const roomMatch = description.match(/Room (\w+)/);
      if (roomMatch) {
        const room = roomMatch[1];
        revenueByRoom[room] = (revenueByRoom[room] || 0) + amount;
      }
    }
  });
  
  const monthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');
  
  let report = `Guest Room Revenue Report - ${monthYear}\n`;
  report += `=====================================\n\n`;
  report += `Total Guest Room Revenue: $${guestRoomRevenue.toFixed(2)}\n`;
  report += `Number of Bookings: ${bookingCount}\n`;
  report += `Average Revenue per Booking: $${bookingCount > 0 ? (guestRoomRevenue / bookingCount).toFixed(2) : '0.00'}\n\n`;
  
  report += `Revenue by Room:\n`;
  Object.entries(revenueByRoom).forEach(([room, revenue]) => {
    report += `  Room ${room}: $${revenue.toFixed(2)}\n`;
  });
  
  // Calculate occupancy rate for guest rooms
  const guestBookings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  if (guestBookings && guestBookings.getLastRow() > 1) {
    const bookingsData = guestBookings.getRange(2, 1, guestBookings.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
    let totalNights = 0;
    
    bookingsData.forEach(booking => {
      const checkIn = new Date(booking[5]);
      const checkOut = new Date(booking[6]);
      
      if ((booking[12] === BOOKING_STATUS.CHECKED_IN || booking[12] === BOOKING_STATUS.CHECKED_OUT) &&
          checkIn <= lastOfMonth && checkOut >= firstOfMonth) {
        
        // Calculate nights within the month
        const startDate = checkIn < firstOfMonth ? firstOfMonth : checkIn;
        const endDate = checkOut > lastOfMonth ? lastOfMonth : checkOut;
        const nights = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));
        totalNights += nights;
      }
    });
    
    const daysInMonth = lastOfMonth.getDate();
    const totalRoomNights = 2 * daysInMonth; // 2 guest rooms
    const occupancyRate = (totalNights / totalRoomNights * 100).toFixed(1);
    
    report += `\nGuest Room Occupancy: ${occupancyRate}% (${totalNights}/${totalRoomNights} nights)`;
  }
  
  const htmlOutput = HtmlService
      .createHtmlOutput(`<pre>${report}</pre>`)
      .setWidth(400)
      .setHeight(400);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Guest Room Revenue Report');
}

/**
 * Combined revenue analysis for all rental types
 */
function combinedRevenueAnalysis() {
  const budgetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BUDGET_SHEET_NAME);
  
  if (!budgetSheet || budgetSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No financial data found.');
    return;
  }
  
  const now = new Date();
  const yearStart = new Date(now.getFullYear(), 0, 1);
  
  const data = budgetSheet.getRange(2, 1, budgetSheet.getLastRow() - 1, 5).getValues();
  
  const monthlyRevenue = {};
  const revenueByType = {
    'Long-term Rent': 0,
    'Guest Room': 0,
    'Other': 0
  };
  
  data.forEach(row => {
    const date = row[0];
    const amount = row[3];
    const category = row[4];
    
    if (date >= yearStart && amount > 0) {
      // Monthly breakdown
      const monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
      if (!monthlyRevenue[monthKey]) {
        monthlyRevenue[monthKey] = {
          'Long-term': 0,
          'Guest': 0,
          'Total': 0
        };
      }
      
      // Categorize revenue
      if (category === 'Rent') {
        revenueByType['Long-term Rent'] += amount;
        monthlyRevenue[monthKey]['Long-term'] += amount;
      } else if (category === 'Guest Room') {
        revenueByType['Guest Room'] += amount;
        monthlyRevenue[monthKey]['Guest'] += amount;
      } else {
        revenueByType['Other'] += amount;
      }
      
      monthlyRevenue[monthKey]['Total'] += amount;
    }
  });
  
  const totalRevenue = Object.values(revenueByType).reduce((sum, val) => sum + val, 0);
  
  let report = `Combined Revenue Analysis - ${now.getFullYear()} YTD\n`;
  report += `=========================================\n\n`;
  report += `Total Revenue: $${totalRevenue.toFixed(2)}\n\n`;
  
  report += `Revenue by Type:\n`;
  Object.entries(revenueByType).forEach(([type, revenue]) => {
    const percentage = totalRevenue > 0 ? (revenue / totalRevenue * 100).toFixed(1) : 0;
    report += `  ${type}: $${revenue.toFixed(2)} (${percentage}%)\n`;
  });
  
  report += `\nMonthly Breakdown:\n`;
  Object.entries(monthlyRevenue).sort().forEach(([month, revenues]) => {
    report += `  ${month}: Total: $${revenues.Total.toFixed(2)} (Long-term: $${revenues['Long-term'].toFixed(2)}, Guest: $${revenues.Guest.toFixed(2)})\n`;
  });
  
  const htmlOutput = HtmlService
      .createHtmlOutput(`<pre>${report}</pre>`)
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Combined Revenue Analysis');
}

/**
 * Calculate total occupancy rate (long-term + guest rooms)
 */
function calculateTotalOccupancyRate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Long-term rooms
  const tenantsSheet = ss.getSheetByName(TENANTS_SHEET_NAME);
  let longTermOccupied = 0;
  let longTermTotal = 0;
  
  if (tenantsSheet && tenantsSheet.getLastRow() > 1) {
    const data = tenantsSheet.getRange(2, COL_ROOM_STATUS, tenantsSheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      if (row[0]) {
        longTermTotal++;
        if (row[0] === 'Occupied') {
          longTermOccupied++;
        }
      }
    });
  }
  
  // Guest rooms
  const guestRoomsSheet = ss.getSheetByName(GUEST_ROOMS_SHEET_NAME);
  let guestOccupied = 0;
  let guestTotal = 0;
  
  if (guestRoomsSheet && guestRoomsSheet.getLastRow() > 1) {
    const data = guestRoomsSheet.getRange(2, 8, guestRoomsSheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      if (row[0]) {
        guestTotal++;
        if (row[0] === 'Occupied') {
          guestOccupied++;
        }
      }
    });
  }
  
  const totalRooms = longTermTotal + guestTotal;
  const totalOccupied = longTermOccupied + guestOccupied;
  const overallRate = totalRooms > 0 ? (totalOccupied / totalRooms * 100).toFixed(1) : 0;
  const longTermRate = longTermTotal > 0 ? (longTermOccupied / longTermTotal * 100).toFixed(1) : 0;
  const guestRate = guestTotal > 0 ? (guestOccupied / guestTotal * 100).toFixed(1) : 0;
  
  const message = `Overall Occupancy: ${totalOccupied}/${totalRooms} (${overallRate}%)\n\n` +
    `Long-term Rentals: ${longTermOccupied}/${longTermTotal} (${longTermRate}%)\n` +
    `Guest Rooms: ${guestOccupied}/${guestTotal} (${guestRate}%)`;
  
  SpreadsheetApp.getUi().alert('Total Occupancy Rate', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Show occupancy calendar
 */
function showOccupancyCalendar() {
  const html = HtmlService.createHtmlOutputFromFile('calendar')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Guest Room Occupancy Calendar');
}

/**
 * Get guest email templates
 */
function getGuestEmailTemplate(type, data) {
  const templates = {
    bookingConfirmation: {
      subject: `Booking Confirmation - ${data.bookingId}`,
      body: `Dear ${data.guestName},

Thank you for booking with Parsonage Guest Accommodation. Your booking has been confirmed!

Booking Details:
- Booking ID: ${data.bookingId}
- Room: ${data.roomNumber}
- Check-in: ${data.checkIn} (after 3:00 PM)
- Check-out: ${data.checkOut} (before 11:00 AM)
- Number of nights: ${data.nights}
- Total Amount: $${data.totalAmount}

${data.specialRequests ? `Special Requests: ${data.specialRequests}\n` : ''}

Check-in Instructions:
- Please check in at the main office between 3:00 PM and 8:00 PM
- Bring a valid photo ID
- Payment is due at check-in if not already paid

House Rules:
- Quiet hours: 10:00 PM - 7:00 AM
- No smoking in rooms
- No pets allowed
- Maximum 2 guests per room

If you need to modify or cancel your booking, please contact us at least 48 hours in advance.

We look forward to hosting you!

Best regards,
Parsonage Management`
    },
    invoice: {
      subject: `Invoice - Booking ${data.bookingId}`,
      body: `Dear ${data.guestName},

Please find attached your invoice for booking ${data.bookingId}.

Thank you for choosing Parsonage Guest Accommodation.

Best regards,
Parsonage Management`
    },
    checkInReminder: {
      subject: `Check-in Reminder - Tomorrow`,
      body: `Dear ${data.guestName},

This is a friendly reminder that your check-in at Parsonage Guest Accommodation is scheduled for tomorrow, ${data.checkIn}.

Room: ${data.roomNumber}
Check-in time: After 3:00 PM

Please remember to bring:
- Valid photo ID
- Payment (if not already paid)

We look forward to welcoming you!

Best regards,
Parsonage Management`
    }
  };
  
  return templates[type] || { subject: 'Parsonage Guest Accommodation', body: 'No template found.' };
}

/**
 * Enhanced form creation to include guest booking form
 */
function autoCreateAllFormsEnhanced() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Create original forms
    const appForm = createTenantApplicationForm();
    const moveOutForm = createMoveOutRequestForm();
    
    // Create guest booking form
    const guestForm = createGuestBookingForm();
    
    // Set up triggers
    setupEnhancedFormTriggers();
    
    // Success message
    const message = `All forms created successfully!\n\n` +
      `Tenant Application Form:\n${appForm.getEditUrl()}\n\n` +
      `Move-Out Form:\n${moveOutForm.getEditUrl()}\n\n` +
      `Guest Booking Form:\n${guestForm.getEditUrl()}`;
    
    ui.alert('Forms Created', message, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Failed to create forms: ${e.message}`, ui.ButtonSet.OK);
    console.error('Form creation error:', e);
  }
}

/**
 * Create guest booking form
 */
function createGuestBookingForm() {
  const form = FormApp.create('Parsonage Guest Room Booking Request');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  form.setDescription('Request a booking for our guest accommodation. We will review your request and contact you within 24 hours.');
  form.setCollectEmail(true);
  form.setRequireLogin(false);
  
  // Personal information
  form.addSectionHeaderItem()
    .setTitle('Guest Information')
    .setHelpText('Please provide your contact details');
  
  form.addTextItem()
    .setTitle('Full Name')
    .setRequired(true);
  
  form.addTextItem()
    .setTitle('Email Address')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextIsEmail()
        .build()
    );
  
  form.addTextItem()
    .setTitle('Phone Number')
    .setRequired(true);
  
  // Booking details
  form.addSectionHeaderItem()
    .setTitle('Booking Details');
  
  form.addDateItem()
    .setTitle('Check-in Date')
    .setRequired(true);
  
  form.addDateItem()
    .setTitle('Check-out Date')
    .setRequired(true);
  
  form.addTextItem()
    .setTitle('Number of Guests')
    .setHelpText('Maximum 2 per room')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireNumberBetween(1, 2)
        .build()
    );
  
  // Room preference
  const roomChoice = form.addMultipleChoiceItem()
    .setTitle('Room Preference')
    .setRequired(false);
  
  // Get guest rooms
  try {
    const guestRoomsSheet = ss.getSheetByName(GUEST_ROOMS_SHEET_NAME);
    if (guestRoomsSheet && guestRoomsSheet.getLastRow() > 1) {
      const rooms = guestRoomsSheet.getRange(2, 1, guestRoomsSheet.getLastRow() - 1, 4).getValues();
      const choices = [];
      
      rooms.forEach(room => {
        if (room[0] && room[1]) {
          choices.push(
            roomChoice.createChoice(`${room[1]} - $${room[2]}/night`)
          );
        }
      });
      
      choices.push(roomChoice.createChoice('No preference'));
      roomChoice.setChoices(choices);
    }
  } catch (e) {
    roomChoice.setChoices([
      roomChoice.createChoice('Guest Suite 1 - $75/night'),
      roomChoice.createChoice('Guest Suite 2 - $65/night'),
      roomChoice.createChoice('No preference')
    ]);
  }
  
  // Purpose of visit
  form.addMultipleChoiceItem()
    .setTitle('Purpose of Visit')
    .setChoices([
      FormApp.createChoice('Visiting family/friends'),
      FormApp.createChoice('Business/work'),
      FormApp.createChoice('Tourism/vacation'),
      FormApp.createChoice('Medical/healthcare'),
      FormApp.createChoice('Other')
    ])
    .setRequired(true);
  
  // Special requests
  form.addParagraphTextItem()
    .setTitle('Special Requests or Notes')
    .setHelpText('Any special requirements or additional information')
    .setRequired(false);
  
  // Agreement
  form.addCheckboxItem()
    .setTitle('Terms and Conditions')
    .setChoices([
      FormApp.createChoice('I understand this is a booking request and not a confirmed reservation'),
      FormApp.createChoice('I agree to the house rules and check-in/out times'),
      FormApp.createChoice('I understand payment is due at check-in')
    ])
    .setRequired(true)
    .setValidation(
      FormApp.createCheckboxValidation()
        .requireSelectAtLeast(3)
        .build()
    );
  
  // Link to spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // Rename the sheet
  Utilities.sleep(2000);
  const sheets = ss.getSheets();
  const newSheet = sheets[sheets.length - 1];
  newSheet.setName('Guest Booking Requests');
  
  return form;
}

/**
 * Handle guest booking form submission
 */
function onGuestBookingSubmit(e) {
  if (!e || !e.namedValues) return;
  
  const name = e.namedValues['Full Name'] ? e.namedValues['Full Name'][0] : '';
  const email = e.namedValues['Email Address'] ? e.namedValues['Email Address'][0] : '';
  const checkIn = e.namedValues['Check-in Date'] ? e.namedValues['Check-in Date'][0] : '';
  const checkOut = e.namedValues['Check-out Date'] ? e.namedValues['Check-out Date'][0] : '';
  
  if (email) {
    const subject = 'Guest Booking Request Received';
    const body = `Dear ${name},

Thank you for your booking request for Parsonage Guest Accommodation.

We have received your request for:
- Check-in: ${checkIn}
- Check-out: ${checkOut}

We will review availability and contact you within 24 hours to confirm your booking or discuss alternatives.

If you have any urgent questions, please don't hesitate to contact us.

Best regards,
Parsonage Management`;
    
    try {
      MailApp.sendEmail(email, subject, body);
      
      // Also notify manager
      const managerSubject = 'New Guest Booking Request';
      const managerBody = `New guest booking request received:\n\n` +
        `Guest: ${name}\n` +
        `Email: ${email}\n` +
        `Dates: ${checkIn} to ${checkOut}\n\n` +
        `Please review in the Guest Booking Requests sheet.`;
      
      MailApp.sendEmail(MANAGER_EMAIL, managerSubject, managerBody);
      
    } catch (error) {
      console.error(`Failed to send booking confirmation: ${error.message}`);
    }
  }
}

/**
 * Setup enhanced triggers including guest bookings
 */
function setupEnhancedFormTriggers() {
  // Set up the form submit router
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Remove existing form submit triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger for form submissions
  ScriptApp.newTrigger('onEnhancedFormSubmitRouter')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

/**
 * Enhanced form submit router
 */
function onEnhancedFormSubmitRouter(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  if (sheetName.includes('Application')) {
    onTenantApplicationSubmit(e);
  } else if (sheetName.includes('Move-Out')) {
    onMoveOutRequestSubmit(e);
  } else if (sheetName.includes('Guest Booking')) {
    onGuestBookingSubmit(e);
  }
}

/**
 * Setup all triggers including guest room daily checks
 */
function setupAllTriggers() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Original triggers
  ScriptApp.newTrigger('checkAllPaymentStatus')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
  
  ScriptApp.newTrigger('sendRentReminders')
      .timeBased()
      .onMonthDay(1)
      .atHour(10)
      .create();
  
  ScriptApp.newTrigger('sendLatePaymentAlerts')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(14)
      .create();
  
  // Guest room triggers
  ScriptApp.newTrigger('dailyGuestRoomCheck')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
  
  SpreadsheetApp.getUi().alert('Triggers Set Up', 'All automated triggers have been configured successfully.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Daily guest room check
 */
function dailyGuestRoomCheck() {
  // Check for today's arrivals and send reminders
  const bookingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_BOOKINGS_SHEET_NAME);
  if (!bookingsSheet || bookingsSheet.getLastRow() < 2) return;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  const data = bookingsSheet.getRange(2, 1, bookingsSheet.getLastRow() - 1, GUEST_BOOKINGS_HEADERS.length).getValues();
  
  data.forEach((row, index) => {
    const checkIn = new Date(row[5]);
    checkIn.setHours(0, 0, 0, 0);
    
    // Send check-in reminder for tomorrow's arrivals
    if (checkIn.getTime() === tomorrow.getTime() && row[12] === BOOKING_STATUS.CONFIRMED) {
      const emailData = {
        guestName: row[1],
        checkIn: Utilities.formatDate(checkIn, Session.getScriptTimeZone(), 'MMMM dd, yyyy'),
        roomNumber: row[4]
      };
      
      const emailTemplate = getGuestEmailTemplate('checkInReminder', emailData);
      
      try {
        MailApp.sendEmail(row[2], emailTemplate.subject, emailTemplate.body);
      } catch (e) {
        console.error(`Failed to send check-in reminder: ${e.message}`);
      }
    }
    
    // Auto-checkout for departures
    const checkOut = new Date(row[6]);
    checkOut.setHours(0, 0, 0, 0);
    
    if (checkOut.getTime() === today.getTime() && row[12] === BOOKING_STATUS.CHECKED_IN) {
      // Mark as checked out
      bookingsSheet.getRange(index + 2, 13).setValue(BOOKING_STATUS.CHECKED_OUT);
      
      // Update room status
      updateGuestRoomStatus(row[4], 'Available', '', '', '');
      
      // Log revenue if fully paid
      if (row[10] >= row[9]) {
        logGuestRoomRevenue(row[1], row[4], row[10]);
      }
    }
  });
}

// Calendar HTML template (create as a separate HTML file named 'calendar.html')

// End of Guest Room Enhancement Code
