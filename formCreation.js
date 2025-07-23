//formCreation.gs
/**
 * Automated Google Forms Creation Script
 * This script creates and configures the required forms for the Parsonage Tenant Management System
 * Add this code to your existing Google Apps Script
 */

/**
 * Creates both required forms (Application and Move-Out)
 * Run this function from the menu: Parsonage Tools > Setup & Configuration > Auto-Create All Forms
 */
function autoCreateAllForms() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Create Application Form
    const appForm = createTenantApplicationForm();
    
    // Create Move-Out Form
    const moveOutForm = createMoveOutRequestForm();
    
    // Set up triggers for both forms
    setupFormSubmitTriggers(appForm, moveOutForm);
    
    // Success message with form URLs
    const message = `Forms created successfully!\n\n` +
      `Application Form:\n${appForm.getEditUrl()}\n\n` +
      `Move-Out Form:\n${moveOutForm.getEditUrl()}\n\n` +
      `Share these forms with prospective and current tenants.`;
    
    ui.alert('Forms Created', message, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Failed to create forms: ${e.message}`, ui.ButtonSet.OK);
    console.error('Form creation error:', e);
  }
}

/**
 * Creates the Tenant Application Form with all required fields
 */
function createTenantApplicationForm() {
  const form = FormApp.create('Parsonage Tenant Application');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set form description
  form.setDescription('Thank you for your interest in our parsonage. Please complete all required fields. We will review your application and contact you within 3-5 business days.');
  
  // Form settings
  form.setCollectEmail(true);
  form.setRequireLogin(false);
  form.setShowLinkToRespondAgain(true);
  
  // Add form header image/title section
  form.addSectionHeaderItem()
    .setTitle('Personal Information')
    .setHelpText('Please provide your basic contact information');
  
  // Name field
  form.addTextItem()
    .setTitle('Full Name')
    .setHelpText('Enter your first and last name')
    .setRequired(true);
  
  // Email field
  form.addTextItem()
    .setTitle('Email Address')
    .setHelpText('We will use this to contact you about your application')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextIsEmail()
        .build()
    );
  
  // Phone field
  form.addTextItem()
    .setTitle('Phone Number')
    .setHelpText('Include area code (e.g., 555-123-4567)')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('[0-9\-\(\) \+]+')
        .setHelpText('Please enter a valid phone number')
        .build()
    );
  
  // Current address
  form.addParagraphTextItem()
    .setTitle('Current Address')
    .setHelpText('Street address, City, State, ZIP')
    .setRequired(true);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Rental Information')
    .setHelpText('Tell us about your rental needs');
  
  // Move-in date
  form.addDateItem()
    .setTitle('Desired Move-in Date')
    .setHelpText('When would you like to move in?')
    .setRequired(true);
  
  // Room preference - you'll need to update these room numbers
  const roomChoice = form.addMultipleChoiceItem()
    .setTitle('Preferred Room')
    .setHelpText('Select your preferred room (subject to availability)')
    .setRequired(true);
  
  // Get rooms from the Tenants sheet if possible
  try {
    const tenantsSheet = ss.getSheetByName(TENANTS_SHEET_NAME);
    if (tenantsSheet && tenantsSheet.getLastRow() > 1) {
      const rooms = tenantsSheet.getRange(2, 1, tenantsSheet.getLastRow() - 1, 3).getValues();
      const availableRooms = [];
      
      rooms.forEach(room => {
        if (room[0]) { // If room number exists
          const roomNum = room[0];
          const price = room[1] || room[2]; // Use rental price or negotiated price
          const status = tenantsSheet.getRange(rooms.indexOf(room) + 2, 8).getValue();
          
          if (status === 'Vacant' || !status) {
            availableRooms.push(
              roomChoice.createChoice(`Room ${roomNum} - $${price}/month`)
            );
          }
        }
      });
      
      if (availableRooms.length > 0) {
        roomChoice.setChoices(availableRooms);
      } else {
        roomChoice.setChoices([
          roomChoice.createChoice('Room 101 - $800/month'),
          roomChoice.createChoice('Room 102 - $800/month'),
          roomChoice.createChoice('Room 103 - $900/month'),
          roomChoice.createChoice('Room 104 - $900/month'),
          roomChoice.createChoice('No preference')
        ]);
      }
    }
  } catch (e) {
    // Default rooms if sheet not available
    roomChoice.setChoices([
      roomChoice.createChoice('Room 101 - $800/month'),
      roomChoice.createChoice('Room 102 - $800/month'),
      roomChoice.createChoice('Room 103 - $900/month'),
      roomChoice.createChoice('Room 104 - $900/month'),
      roomChoice.createChoice('No preference')
    ]);
  }
  
  // Length of stay
  form.addMultipleChoiceItem()
    .setTitle('Expected Length of Stay')
    .setChoices([
      FormApp.createChoice('6 months'),
      FormApp.createChoice('1 year'),
      FormApp.createChoice('2+ years'),
      FormApp.createChoice('Unsure')
    ])
    .setRequired(true);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Employment & Financial Information')
    .setHelpText('This information helps us ensure you can comfortably afford the rental');
  
  // Employment status
  form.addMultipleChoiceItem()
    .setTitle('Current Employment Status')
    .setChoices([
      FormApp.createChoice('Employed Full-Time'),
      FormApp.createChoice('Employed Part-Time'),
      FormApp.createChoice('Self-Employed'),
      FormApp.createChoice('Student'),
      FormApp.createChoice('Retired'),
      FormApp.createChoice('Other')
    ])
    .setRequired(true);
  
  // Employer/School
  form.addTextItem()
    .setTitle('Employer/School Name')
    .setHelpText('Name of your current employer or educational institution')
    .setRequired(true);
  
  // Monthly income
  form.addTextItem()
    .setTitle('Monthly Income (Gross)')
    .setHelpText('Your income before taxes')
    .setRequired(true);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('References')
    .setHelpText('Please provide two references (not family members)');
  
  // Reference 1
  form.addParagraphTextItem()
    .setTitle('Reference 1')
    .setHelpText('Name, Phone Number, Relationship (e.g., Current Landlord, Employer)')
    .setRequired(true);
  
  // Reference 2
  form.addParagraphTextItem()
    .setTitle('Reference 2')
    .setHelpText('Name, Phone Number, Relationship')
    .setRequired(true);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Additional Information');
  
  // Emergency contact
  form.addParagraphTextItem()
    .setTitle('Emergency Contact')
    .setHelpText('Name, Phone Number, Relationship')
    .setRequired(true);
  
  // Vehicle information
  form.addTextItem()
    .setTitle('Vehicle Information')
    .setHelpText('Make, Model, Year, License Plate (if you have a vehicle)');
  
  // About yourself
  form.addParagraphTextItem()
    .setTitle('Tell Us About Yourself')
    .setHelpText('Share a bit about yourself, your interests, and why you\'d like to live in our parsonage')
    .setRequired(false);
  
  // Special needs
  form.addParagraphTextItem()
    .setTitle('Special Needs or Requests')
    .setHelpText('Any accommodations or special requests we should know about?')
    .setRequired(false);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Documents')
    .setHelpText('Please upload required documents');
  
  // File upload for proof of income
  // Note: File upload requires Google Workspace account
  try {
    form.addTextItem()
      .setTitle('Proof of Income')
      .setHelpText('Please describe your proof of income (e.g., "Recent pay stubs uploaded", "Bank statements available upon request"). You may need to email documents separately.')
      .setRequired(true);
  } catch (e) {
    // If file upload fails, use text field
    console.log('File upload not available, using text field instead');
  }
  
  // Agreement checkbox
  form.addCheckboxItem()
    .setTitle('Application Agreement')
    .setChoices([
      FormApp.createChoice('I certify that all information provided is accurate and complete'),
      FormApp.createChoice('I understand that false information may result in denial of my application'),
      FormApp.createChoice('I agree to a background and credit check if my application is considered')
    ])
    .setRequired(true)
    .setValidation(
      FormApp.createCheckboxValidation()
        .requireSelectAtLeast(3)
        .build()
    );
  
  // Link to spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // Get the new sheet and rename it
  Utilities.sleep(2000); // Wait for sheet creation
  const sheets = ss.getSheets();
  const newSheet = sheets[sheets.length - 1]; // Get the last sheet (newly created)
  newSheet.setName(APPLICATION_SHEET_NAME);
  
  return form;
}

/**
 * Creates the Move-Out Request Form
 */
function createMoveOutRequestForm() {
  const form = FormApp.create('Parsonage Move-Out Request');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set form description
  form.setDescription('Please submit this form at least 30 days before your intended move-out date. We will contact you to schedule a move-out inspection.');
  
  // Form settings
  form.setCollectEmail(true);
  form.setRequireLogin(false);
  
  // Header
  form.addSectionHeaderItem()
    .setTitle('Move-Out Request Form')
    .setHelpText('Please complete all required fields');
  
  // Tenant name
  form.addTextItem()
    .setTitle('Your Full Name')
    .setHelpText('As it appears on your rental agreement')
    .setRequired(true);
  
  // Email
  form.addTextItem()
    .setTitle('Email Address')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextIsEmail()
        .build()
    );
  
  // Phone
  form.addTextItem()
    .setTitle('Phone Number')
    .setHelpText('Best number to reach you')
    .setRequired(true);
  
  // Room number
  const roomItem = form.addTextItem()
    .setTitle('Room Number')
    .setHelpText('Your current room number')
    .setRequired(true);
  
  // Move-out date
  form.addDateItem()
    .setTitle('Planned Move-Out Date')
    .setHelpText('Must be at least 30 days from today')
    .setRequired(true);
  
  // Forwarding address
  form.addParagraphTextItem()
    .setTitle('Forwarding Address')
    .setHelpText('Where should we send your security deposit and any correspondence?')
    .setRequired(true);
  
  // Reason for moving
  form.addMultipleChoiceItem()
    .setTitle('Primary Reason for Moving')
    .setChoices([
      FormApp.createChoice('Job relocation'),
      FormApp.createChoice('Found permanent housing'),
      FormApp.createChoice('Financial reasons'),
      FormApp.createChoice('Family reasons'),
      FormApp.createChoice('Dissatisfied with accommodation'),
      FormApp.createChoice('Other')
    ])
    .setRequired(true);
  
  // Additional details
  form.addParagraphTextItem()
    .setTitle('Additional Details')
    .setHelpText('Please elaborate on your reason for moving (optional)')
    .setRequired(false);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Move-Out Logistics');
  
  // Inspection availability
  form.addCheckboxItem()
    .setTitle('Availability for Move-Out Inspection')
    .setHelpText('Check all times you\'re generally available')
    .setChoices([
      FormApp.createChoice('Weekday mornings (9 AM - 12 PM)'),
      FormApp.createChoice('Weekday afternoons (12 PM - 5 PM)'),
      FormApp.createChoice('Weekday evenings (5 PM - 7 PM)'),
      FormApp.createChoice('Saturday mornings'),
      FormApp.createChoice('Saturday afternoons'),
      FormApp.createChoice('Sunday afternoons')
    ])
    .setRequired(true);
  
  // Key return
  form.addMultipleChoiceItem()
    .setTitle('Key Return Method')
    .setHelpText('How do you plan to return your keys?')
    .setChoices([
      FormApp.createChoice('In person during inspection'),
      FormApp.createChoice('Drop in office mailbox'),
      FormApp.createChoice('Hand to management'),
      FormApp.createChoice('Other arrangement needed')
    ])
    .setRequired(true);
  
  // Section break
  form.addSectionHeaderItem()
    .setTitle('Feedback (Optional)')
    .setHelpText('Your feedback helps us improve');
  
  // Rating
  form.addScaleItem()
    .setTitle('Overall Satisfaction')
    .setHelpText('How would you rate your experience living here?')
    .setBounds(1, 5)
    .setLabels('Very Unsatisfied', 'Very Satisfied');
  
  // What went well
  form.addParagraphTextItem()
    .setTitle('What aspects of living here did you appreciate?')
    .setRequired(false);
  
  // What could improve
  form.addParagraphTextItem()
    .setTitle('What could we improve for future tenants?')
    .setRequired(false);
  
  // Would recommend
  form.addMultipleChoiceItem()
    .setTitle('Would you recommend our parsonage to others?')
    .setChoices([
      FormApp.createChoice('Yes'),
      FormApp.createChoice('No'),
      FormApp.createChoice('Maybe')
    ]);
  
  // Acknowledgments
  form.addCheckboxItem()
    .setTitle('Move-Out Acknowledgments')
    .setChoices([
      FormApp.createChoice('I understand I must leave my room in clean, rentable condition'),
      FormApp.createChoice('I understand deductions may be made from my security deposit for damages or excessive cleaning'),
      FormApp.createChoice('I will remove all personal belongings by the move-out date'),
      FormApp.createChoice('I understand rent is due through my move-out date')
    ])
    .setRequired(true)
    .setValidation(
      FormApp.createCheckboxValidation()
        .requireSelectAtLeast(4)
        .build()
    );
  
  // Additional comments
  form.addParagraphTextItem()
    .setTitle('Any other comments or special circumstances?')
    .setRequired(false);
  
  // Link to spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // Get the new sheet and rename it
  Utilities.sleep(2000); // Wait for sheet creation
  const sheets = ss.getSheets();
  const newSheet = sheets[sheets.length - 1]; // Get the last sheet (newly created)
  newSheet.setName(MOVEOUT_SHEET_NAME);
  
  return form;
}

/**
 * Sets up form submit triggers for both forms
 */
function setupFormSubmitTriggers(appForm, moveOutForm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create trigger for application form
  ScriptApp.newTrigger('onTenantApplicationSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  
  // Note: Since both forms link to the same spreadsheet, we need to modify
  // the submit handlers to check which form was submitted
  
  // Store form URLs in script properties for reference
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('APPLICATION_FORM_URL', appForm.getPublishedUrl());
  scriptProperties.setProperty('MOVEOUT_FORM_URL', moveOutForm.getPublishedUrl());
  
  console.log('Form triggers set up successfully');
}

/**
 * Enhanced form submit handler that detects which form was submitted
 */
function onFormSubmitRouter(e) {
  // Check which sheet the submission came to
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  if (sheetName === APPLICATION_SHEET_NAME || sheetName.includes('Application')) {
    onTenantApplicationSubmit(e);
  } else if (sheetName === MOVEOUT_SHEET_NAME || sheetName.includes('Move-Out')) {
    onMoveOutRequestSubmit(e);
  }
}

/**
 * Updates the onOpen menu to include the form creation option
 */
function updateMenuWithFormCreation() {
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
          .addItem('Auto-Create All Forms', 'autoCreateAllForms')
          .addItem('View Form URLs', 'showFormURLs')
          .addItem('Configure Email Templates', 'configureEmailTemplates'))
      .addSeparator()
      .addItem('Send Rent Reminders (Test)', 'sendRentRemindersTest')
      .addToUi();
}

/**
 * Shows the URLs of created forms
 */
function showFormURLs() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const appFormUrl = scriptProperties.getProperty('APPLICATION_FORM_URL');
  const moveOutFormUrl = scriptProperties.getProperty('MOVEOUT_FORM_URL');
  
  let message = 'Form URLs:\n\n';
  
  if (appFormUrl) {
    message += `Application Form:\n${appFormUrl}\n\n`;
  } else {
    message += 'Application Form: Not yet created\n\n';
  }
  
  if (moveOutFormUrl) {
    message += `Move-Out Form:\n${moveOutFormUrl}`;
  } else {
    message += 'Move-Out Form: Not yet created';
  }
  
  SpreadsheetApp.getUi().alert('Form URLs', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Helper function to get room information from the Tenants sheet
 */
function getAvailableRooms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tenantsSheet = ss.getSheetByName(TENANTS_SHEET_NAME);
  
  if (!tenantsSheet || tenantsSheet.getLastRow() < 2) {
    return [];
  }
  
  const rooms = tenantsSheet.getRange(2, 1, tenantsSheet.getLastRow() - 1, 8).getValues();
  const availableRooms = [];
  
  rooms.forEach(room => {
    if (room[0] && (room[7] === 'Vacant' || !room[7])) {
      availableRooms.push({
        number: room[0],
        price: room[2] || room[1], // Negotiated price or standard price
        status: room[7]
      });
    }
  });
  
  return availableRooms;
}

// End of Form Creation Script
