/**
 * @fileoverview WebVeer HRMS - Enterprise Edition (vFinal)
 * @owner WEBVEER AUTOMATION AND SERVICE PRIVATE LIMITED
 */

const PROPERTIES = PropertiesService.getScriptProperties();

/**
 * Serves the App and handles URL Routing Parameters
 */
function doGet(e) {
  const dbId = PROPERTIES.getProperty('DB_ID');
  const template = HtmlService.createTemplateFromFile('index');
  
  // Pass Database Status and URL Page Parameter to Frontend
  template.isSetup = !!dbId;
  template.initialPage = (e.parameter && e.parameter.page) ? e.parameter.page : 'dashboard';
  
  return template.evaluate()
    .setTitle('WebVeer Automation HRMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * INITIAL SETUP - Generates the complete Database Schema
 */
function setupSystem() {
  let ss;
  let ssId = PROPERTIES.getProperty('DB_ID');
  let isNewSheet = false;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ssId = null; // Reset if sheet is not accessible
    }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create("WebVeer_HRMS_Database_Master");
    ssId = ss.getId();
    PROPERTIES.setProperty('DB_ID', ssId);
    isNewSheet = true;
  }

  // Defined Schema with Leave Balances
  const schemas = {
    'Employees': ['EmpID', 'Name', 'Email', 'Role', 'Department', 'Designation', 'DOJ', 'DOB', 'Mobile', 'Status', 'BalCL', 'BalSL', 'BalMat', 'BalPat', 'LastBalanceUpdate', 'ManagerID'],
    'Attendance': ['Date', 'EmpID', 'Name', 'Type', 'Time', 'Lat', 'Lng', 'MapLink', 'Device'],
    'Leaves': ['RequestID', 'EmpID', 'Name', 'Type', 'StartDate', 'EndDate', 'Reason', 'Status', 'Days'],
    'SystemLogs': ['Timestamp', 'UserEmail', 'Action', 'Details', 'Meta'],
    'Announcements': ['Date', 'Title', 'Message', 'PostedBy'],
    'Assets': ['AssetID', 'Type', 'Model', 'SerialNo', 'AssignedTo', 'Status'],
    'Payroll': ['PayslipID', 'EmpID', 'Month', 'Year', 'NetPay', 'GenDate'],
    'Documents': ['DocID', 'EmpID', 'DocumentType', 'FileName', 'FileURL', 'UploadDate'],
    'DocumentTemplates': ['TemplateID', 'Title', 'Content'],
    'GeneratedDocuments': ['GeneratedDocID', 'TemplateID', 'EmpID', 'Status', 'CreatedDate', 'ApprovedDate', 'Data', 'FileURL'],
    'Holidays': ['Date', 'Title', 'Type']
  };

  if (isNewSheet) {
    const sheets = ss.getSheets();
    if (sheets.length > 0) sheets[0].setName('Config_Dashboard');
  }

  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      // Header Styling (Corporate Blue)
      const rng = sheet.getRange(1, 1, 1, headers.length);
      rng.setValues([headers]);
      rng.setBackground("#003366").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  }

  if (isNewSheet) {
    // Create Super Admin with 0 CL to start pro-rata calculation
    const email = Session.getActiveUser().getEmail();
    const doj = new Date();
    const empSheet = ss.getSheetByName('Employees');
    empSheet.appendRow(['WV-ADMIN', 'Super Admin', email, 'ADMIN', 'Mgmt', 'Director', doj, '', '9999999999', 'Active', 0, 7, 0, 0, doj, '']);

    // Initial Log
    ss.getSheetByName('SystemLogs').appendRow([new Date(), 'System', 'Setup', 'Database Initialized', 'Server']);
    ss.getSheetByName('Announcements').appendRow([new Date(), 'Welcome', 'Welcome to WebVeer HRMS.', 'System']);
  }

  return { success: true, url: ss.getUrl() };
}

function getDbId_() { return PROPERTIES.getProperty('DB_ID'); }

/**
 * Run this function once to add missing tables to your existing database.
 */
function patchDatabase() {
  patchDatabase_();
}

/**
 * Run this ONCE to fix missing tables (Logs, Holidays)
 */
function fixDatabaseTables() {
  patchDatabase_();
}

function addDummyHolidays_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) return;
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('Holidays');
  if (sheet.getLastRow() === 1) { // Only add if no data exists
    sheet.appendRow(['2025-01-26', 'Republic Day', 'Public Holiday']);
    sheet.appendRow(['2025-08-15', 'Independence Day', 'Public Holiday']);
    Logger.log('Added dummy holiday data.');
  }
}

function addDummyTemplates_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) return;
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('DocumentTemplates');
  if (sheet.getLastRow() === 1) { // Only add if no data exists
    const offerLetterTemplate = `
      <h1>Offer Letter</h1>
      <p>Date: {{current_date}}</p>
      <br>
      <p>Dear {{employee_name}},</p>
      <p>We are pleased to offer you the position of <b>{{designation}}</b> at WebVeer Automation.</p>
      <p>Your start date will be {{start_date}}.</p>
      <p>Your annual salary will be {{salary}}.</p>
      <br>
      <p>Sincerely,</p>
      <p><b>The WebVeer Team</b></p>
    `;
    const expenseReportTemplate = `
      <h1>Expense Report</h1>
      <p><b>Employee:</b> {{employee_name}}</p>
      <p><b>Date of Expense:</b> {{expense_date}}</p>
      <p><b>Amount:</b> {{currency}} {{amount}}</p>
      <p><b>Description:</b></p>
      <p>{{description}}</p>
      <p><b>Receipt URL:</b> <a href="{{receipt_url}}">{{receipt_url}}</a></p>
    `;
    const wfhRequestTemplate = `
      <h1>Work From Home Request</h1>
      <p><b>Employee:</b> {{employee_name}}</p>
      <p><b>Start Date:</b> {{start_date}}</p>
      <p><b>End Date:</b> {{end_date}}</p>
      <p><b>Reason:</b></p>
      <p>{{reason}}</p>
    `;
    const addressVerificationTemplate = `
      <h1>To Whom It May Concern</h1>
      <br>
      <p>This is to certify that <b>{{employee_name}}</b> is an employee of WebVeer Automation.</p>
      <p>As per our records, their current residential address is:</p>
      <p><b>{{employee_address}}<br>{{city}}, {{state}} - {{zip_code}}</b></p>
      <p>This letter is issued upon the request of the employee for verification purposes.</p>
      <br>
      <p>Sincerely,</p>
      <p><b>HR Department, WebVeer Automation</b></p>
    `;
    const employmentVerificationTemplate = `
      <h1>To Whom It May Concern</h1>
      <br>
      <p>This is to certify that <b>{{employee_name}}</b> has been an employee of WebVeer Automation since <b>{{doj}}</b>.</p>
      <p>Their current designation is <b>{{designation}}</b> in the <b>{{department}}</b> department.</p>
      <p>This letter is issued upon the request of the employee for employment verification.</p>
      <br>
      <p>Sincerely,</p>
      <p><b>HR Department, WebVeer Automation</b></p>
    `;

    sheet.appendRow(['OFFER-001', 'Offer Letter', offerLetterTemplate]);
    sheet.appendRow(['EXPENSE-001', 'Expense Reimbursement Request', expenseReportTemplate]);
    sheet.appendRow(['WFH-001', 'Work From Home Request', wfhRequestTemplate]);
    sheet.appendRow(['ADDR-VER-001', 'Address Verification Letter', addressVerificationTemplate]);
    sheet.appendRow(['EMP-VER-001', 'Employment Verification Letter', employmentVerificationTemplate]);
    Logger.log('Added dummy document templates.');
  }
}

function patchDatabase_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) {
    Logger.log("No Database ID found. Please run Setup first.");
    return;
  }
  
  const ss = SpreadsheetApp.openById(id);
  
  const requiredSheets = {
    'Holidays': ['Date', 'Title', 'Type'],
    'SystemLogs': ['Timestamp', 'UserEmail', 'Action', 'Details', 'IP_UserAgent'],
    'Documents': ['DocID', 'EmpID', 'DocumentType', 'FileName', 'FileURL', 'UploadDate'],
    'DocumentTemplates': ['TemplateID', 'Title', 'Content'],
    'GeneratedDocuments': ['GeneratedDocID', 'TemplateID', 'EmpID', 'Status', 'CreatedDate', 'ApprovedDate', 'Data', 'FileURL'],
    'Employees': ['EmpID', 'Name', 'Email', 'Role', 'Department', 'Designation', 'DOJ', 'DOB', 'Mobile', 'Status', 'BalCL', 'BalSL', 'BalMat', 'BalPat', 'LastBalanceUpdate', 'ManagerID']
  };
  
  for (const [name, headers] of Object.entries(requiredSheets)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      const rng = sheet.getRange(1, 1, 1, headers.length);
      rng.setValues([headers]);
      rng.setBackground("#003366").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
      Logger.log(`✅ Created missing sheet: ${name}`);
      
      if(name === 'Holidays') {
        addDummyHolidays_();
      }
      if(name === 'DocumentTemplates') {
        addDummyTemplates_();
      }
    } else {
      Logger.log(`ℹ️ Sheet already exists: ${name}`);
    }
  }
  
  Logger.log("Database Patch Complete. You can now reload the Web App.");
}

// --- AUTOMATED LEAVE ACCRUAL ---

/**
 * Creates a daily trigger to run the leave accrual function.
 * Run this once from the script editor to set up the automation.
 */
function createLeaveAccrualTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExists = triggers.some(t => t.getHandlerFunction() === 'dailyLeaveAccrual_');

  if (!triggerExists) {
    ScriptApp.newTrigger('dailyLeaveAccrual_')
      .timeBased()
      .everyDays(1)
      .atHour(1) // Runs every day at 1 AM
      .create();
    Logger.log('✅ Daily leave accrual trigger created successfully.');
  } else {
    Logger.log('ℹ️ Daily leave accrual trigger already exists.');
  }
}

/**
 * This function is meant to be run by a daily trigger.
 * It resets leaves on Jan 1 and accrues 1.5 CL on the 1st of every month.
 */
function dailyLeaveAccrual_() {
  const today = new Date();
  
  // On January 1st, run the yearly reset process.
  if (today.getMonth() === 0 && today.getDate() === 1) {
    yearlyLeaveReset_();
  }
  
  // On the 1st day of any month, accrue 1.5 CL for all active employees.
  if (today.getDate() === 1) {
    const ss = SpreadsheetApp.openById(getDbId_());
    const sheet = ss.getSheetByName('Employees');
    const data = sheet.getRange("A2:N" + sheet.getLastRow()).getValues();
    const headers = sheet.getRange("A1:N1").getValues()[0];
    
    const clColIndex = headers.indexOf('BalCL');
    const statusColIndex = headers.indexOf('Status');
    const emailColIndex = headers.indexOf('Email');
    
    for (let i = 0; i < data.length; i++) {
      const employee = data[i];
      if (employee[statusColIndex] === 'Active') {
        const currentCL = parseFloat(employee[clColIndex]) || 0;
        const newCL = currentCL + 1.5;
        sheet.getRange(i + 2, clColIndex + 1).setValue(newCL);
        logAction_('Leave Accrual', `Added 1.5 CL for ${employee[emailColIndex]}`);
      }
    }
    Logger.log(`Leave accrual complete for ${today.toDateString()}`);
  }
}

/**
 * Resets leave balances at the start of the year (Jan 1).
 * TODO: Implement specific company policy (e.g., carry-over limits).
 */
function yearlyLeaveReset_() {
  // This is a placeholder for year-end leave balance reset logic.
  // For example, you could reset CL to 0 or apply a carry-over limit.
  const ss = SpreadsheetApp.openById(getDbId_());
  const sheet = ss.getSheetByName('Employees');
  const clColIndex = sheet.getRange("A1:N1").getValues()[0].indexOf('BalCL') + 1;
  
  // Example: Reset all CL to 0. A more complex rule could be applied here.
  // sheet.getRange(2, clColIndex, sheet.getLastRow() - 1, 1).setValue(0);
  
  logAction_('System', 'Yearly leave reset check performed.');
}


// --- UTILITY FUNCTIONS ---

function isWeekendOrHoliday_(date, holidays) {
    const day = date.getDay();
    if (day === 0) return true; // Sunday is a holiday

    if (day === 6) { // Saturday
        const weekOfMonth = Math.ceil(date.getDate() / 7);
        if (weekOfMonth === 1 || weekOfMonth === 3) {
            return true; // 1st and 3rd Saturdays are holidays
        }
        return false; // 2nd, 4th, and 5th Saturdays are working
    }

    const dateString = Utilities.formatDate(date, "Asia/Kolkata", "yyyy-MM-dd");
    if (holidays.includes(dateString)) {
        return true; // It's a holiday from the holiday list
    }

    return false; // It's a working day
}