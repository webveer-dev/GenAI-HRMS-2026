// --- UTILITIES ---

function logAction_(action, details) {
  try {
    const id = getDbId_();
    if(!id) return;
    const email = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.openById(id).getSheetByName('SystemLogs');
    sheet.appendRow([new Date(), email, action, details, 'Web']);
  } catch(e) { console.error("Log Error", e); }
}

function getData(sheetName) {
  const id = getDbId_();
  if (!id) throw new Error("Database Disconnected");
  
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(sheetName);
  
  // 1. Missing Sheet Protection
  if (!sheet) {
    console.error(`Missing Sheet: ${sheetName}`);
    return []; // Return empty array to prevent Frontend Crash
  }

  // 2. Empty Sheet Protection
  if (sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  let headers = values.shift().map(h => String(h).trim());

  // 3. Date Serialization Fix
  return values.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        if (val.getFullYear() < 1900) { // Sheets converts time-only to a date in 1899
          obj[h] = Utilities.formatDate(val, "Asia/Kolkata", "HH:mm:ss");
        } else {
          try { obj[h] = Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd"); } catch(e){ obj[h] = ""; }
        }
      } else {
         obj[h] = (val === undefined || val === null) ? "" : val;
      }
    });
    return obj;
  });
}

// --- CORE AUTHENTICATION ---

function getActiveUserContext() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    const employees = getData('Employees');
    
    // Filter valid rows
    const valid = employees.filter(r => r.Email && String(r.Email).trim() !== "");
    
    // Case-Insensitive Check
    const user = valid.find(e => String(e.Email).toLowerCase().trim() === email);
    
    if (!user) {
      logAction_('Auth Failed', 'Email not in DB: ' + email);
      return { error: "Access Denied: Email not found.", Email: email }; 
    }
    
    user.Role = user.Role ? user.Role.toUpperCase().trim() : '';

    // Round leave balances
    user.BalCL = Math.round((user.BalCL || 0) * 100) / 100;
    user.BalSL = Math.round((user.BalSL || 0) * 100) / 100;
    user.BalMat = Math.round((user.BalMat || 0) * 100) / 100;
    user.BalPat = Math.round((user.BalPat || 0) * 100) / 100;

    logAction_('Login', 'User accessed system');
    return user;
  } catch (e) { return { error: "System Error: " + e.message }; }
}

/**
 * Gets attendance history for the current user.
 * Called from attendance-user.html
 */
function getUserAttendanceHistory() {
    const user = getActiveUserContext();
    if (user.error) {
        console.error("getUserAttendanceHistory Error:", user.error);
        return [];
    }
    const attendance = getData('Attendance');
    return attendance.filter(rec => rec.EmpID === user.EmpID);
}

/**
 * Gets aggregated attendance data for the admin dashboard for a specific date.
 * Called from attendance-admin.html
 */
function getAdminAttendanceData(dateString) {
    try {
        const user = getActiveUserContext();
        if (user.error || (user.Role !== 'ADMIN' && user.Role !== 'HR')) {
            return { success: false, message: "Unauthorized" };
        }

        const allAttendance = getData('Attendance');
        const dailyAttendance = allAttendance.filter(rec => rec.Date === dateString);

        // 1. Map Data
        const mapData = [['Lat', 'Lng', 'Name']];
        dailyAttendance.forEach(rec => {
            if (rec.Lat && rec.Lng) {
                mapData.push([parseFloat(rec.Lat), parseFloat(rec.Lng), `${rec.Name} (${rec.Type})`]);
            }
        });

        // 2. Hours Data & Logs
        const employeeHours = {};
        const logs = [];
        dailyAttendance.forEach(rec => {
            if (!employeeHours[rec.EmpID]) {
                employeeHours[rec.EmpID] = { name: rec.Name, in: null, out: null };
            }
            if (rec.Type === 'CheckIn') employeeHours[rec.EmpID].in = new Date(`${rec.Date}T${rec.Time}`);
            if (rec.Type === 'CheckOut') employeeHours[rec.EmpID].out = new Date(`${rec.Date}T${rec.Time}`);
            logs.push({ empId: rec.EmpID, name: rec.Name, type: rec.Type === 'CheckIn' ? 'IN' : 'OUT' });
        });

        const hoursData = [['Employee', 'Hours']];
        for (const empId in employeeHours) {
            const emp = employeeHours[empId];
            if (emp.in && emp.out) {
                const diffMs = emp.out - emp.in;
                const hours = diffMs / 3600000; // Convert ms to hours
                hoursData.push([emp.name, hours]);
            }
        }

        return { success: true, mapData, hoursData, logs };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * Gets the detailed check-in/out log for a single employee on a given day.
 * Called from attendance-admin.html modal.
 */
function getEmployeeDailyDetails(empId, dateString) {
    const allAttendance = getData('Attendance');
    return allAttendance.filter(rec => rec.EmpID === empId && rec.Date === dateString)
                        .map(rec => ({ type: rec.Type === 'CheckIn' ? 'IN' : 'OUT', time: rec.Time, lat: rec.Lat, lng: rec.Lng }));
}

function getAttendanceHistory() {
    const user = getActiveUserContext();
    if (user.error) return [];

    const attendance = getData('Attendance');
  const adminRoles = ['ADMIN', 'HR', 'ACCOUNTENT'];
  const role = (user.Role || '').toUpperCase().trim();

  if (adminRoles.includes(role)) {
    logAction_('Attendance Access', `User ${user.Email} with role ${role} accessed all attendance records.`);
    return attendance;
  }

  // Default: return only records for the logged-in employee
  return attendance.filter(rec => String(rec.EmpID) === String(user.EmpID));
}

function getTodaysAttendance() {
    const user = getActiveUserContext();
    if (user.error) return { checkIn: null, checkOut: null };

    const today = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
    const attendance = getData('Attendance');
    const userAttendance = attendance.filter(rec => rec.EmpID === user.EmpID && rec.Date === today);

    const checkIn = userAttendance.find(rec => rec.Type === 'CheckIn');
    const checkOut = userAttendance.find(rec => rec.Type === 'CheckOut');

    return {
        checkIn: checkIn ? checkIn.Time : null,
        checkOut: checkOut ? checkOut.Time : null
    };
}

function getServerTime() {
    const now = new Date();
    return {
        date: Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd"),
        time: Utilities.formatDate(now, "Asia/Kolkata", "HH:mm:ss")
    };
}

function searchAttendance(searchCriteria) {
    const user = getActiveUserContext();
    if (user.error) return [];

    let attendance = getData('Attendance');
    if (user.Role !== 'ADMIN' && user.Role !== 'HR' && user.Role !== 'ACCOUNTENT') {
        attendance = attendance.filter(rec => rec.EmpID === user.EmpID);
    }

    let filtered = attendance;

    if (searchCriteria.date) {
        filtered = filtered.filter(rec => rec.Date === searchCriteria.date);
    }

    if (searchCriteria.name) {
        const lowerCaseName = searchCriteria.name.toLowerCase();
        filtered = filtered.filter(rec => rec.Name.toLowerCase().includes(lowerCaseName));
    }

    return filtered;
}

function exportAttendanceToSheet(data) {
    try {
        const sheet = SpreadsheetApp.create("Attendance Export").getActiveSheet();
        const headers = ["Date", "EmpID", "Name", "Type", "Time", "MapLink"];
        sheet.appendRow(headers);

        data.forEach(row => {
            sheet.appendRow([row.Date, row.EmpID, row.Name, row.Type, row.Time, row.MapLink]);
        });

        return { success: true, url: sheet.getParent().getUrl() };
    } catch (e) {
        return { success: false, msg: "Error exporting to sheet: " + e.message };
    }
}

function accrueMonthlyLeave() {
    const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const balCLIndex = headers.indexOf('BalCL');
    const lastUpdateIndex = headers.indexOf('LastBalanceUpdate');
    const dojIndex = headers.indexOf('DOJ');
    const today = new Date();
    const currentYear = today.getFullYear();

    const isLeapYear = (year) => (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
    const daysInYear = isLeapYear(currentYear) ? 366 : 365;
    const dailyRate = 18 / daysInYear;

    data.forEach((row, index) => {
        const doj = new Date(row[dojIndex]);
        let lastUpdate = row[lastUpdateIndex] ? new Date(row[lastUpdateIndex]) : doj;
        let currentBalance = parseFloat(row[balCLIndex]) || 0;
        currentBalance = Math.round(currentBalance * 100) / 100;

        let calculationStartDate;

        if (lastUpdate.getFullYear() < currentYear) {
            calculationStartDate = new Date(currentYear, 0, 1); // Jan 1st of current year
        } else {
            calculationStartDate = lastUpdate;
        }

        if (doj.getFullYear() === currentYear && doj > calculationStartDate) {
            calculationStartDate = doj;
        }
        
        // If the calculation start date is today, do nothing
        if (calculationStartDate.toDateString() === today.toDateString()) {
            return;
        }

        const timeDiff = today.getTime() - calculationStartDate.getTime();
        const daysDiff = Math.floor(timeDiff / (1000 * 3600 * 24));

        if (daysDiff > 0) {
            const leavesToAdd = daysDiff * dailyRate;
            let newBalance = currentBalance + leavesToAdd;
            newBalance = Math.round(newBalance * 100) / 100;
            sheet.getRange(index + 2, balCLIndex + 1).setValue(newBalance);
            sheet.getRange(index + 2, lastUpdateIndex + 1).setValue(today);
        }
    });

    return { success: true, msg: "Prorated monthly leave accrued successfully." };
}

function applyYearlyCarryOver() {
    const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const balCLIndex = headers.indexOf('BalCL');

    data.forEach((row, index) => {
        const currentBalance = parseFloat(row[balCLIndex]) || 0;
        // Carry over 50% and add 1 CL bonus
        let newBalance = (currentBalance * 0.5) + 1;
        newBalance = Math.round(newBalance * 100) / 100;
        sheet.getRange(index + 2, balCLIndex + 1).setValue(newBalance);
    });

    return { success: true, msg: "Yearly carry-over and bonus applied." };
}

function calculateLeaveDays(startDate, endDate) {
    let leaveDays = 0;
    let currentDate = new Date(startDate);
    const stopDate = new Date(endDate);
    const allHolidays = getData('Holidays').map(h => h.Date);

    while (currentDate <= stopDate) {
        if (!isWeekendOrHoliday_(currentDate, allHolidays)) {
            leaveDays++;
        }
        currentDate.setDate(currentDate.getDate() + 1);
    }
    return leaveDays;
}

// Returns details about leave calculation: days counted, excluded dates (weekends/holidays)
function calculateLeaveDaysDetailed(startDate, endDate) {
  let leaveDays = 0;
  let excluded = [];
  let currentDate = new Date(startDate);
  const stopDate = new Date(endDate);
  const allHolidays = getData('Holidays').map(h => h.Date);

  while (currentDate <= stopDate) {
    if (!isWeekendOrHoliday_(currentDate, allHolidays)) {
      leaveDays++;
    } else {
      excluded.push(Utilities.formatDate(new Date(currentDate), "Asia/Kolkata", "yyyy-MM-dd"));
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return { days: leaveDays, excludedCount: excluded.length, excludedDates: excluded };
}




// --- MODULES ---

function markAttendance(type, loc) {
  const now = new Date();
  const allHolidays = getData('Holidays').map(h => h.Date);
  if (isWeekendOrHoliday_(now, allHolidays)) {
    return {success: false, msg: "Today is a holiday or non-working day."};
  }

  const user = getActiveUserContext();
  if (user.error) return {success: false, msg: user.error};

  const id = getDbId_();
  const sheet = SpreadsheetApp.openById(id).getSheetByName('Attendance');
  const today = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd");
  const time = Utilities.formatDate(now, "Asia/Kolkata", "HH:mm:ss");

  // Duplicate Check
  const logs = getData('Attendance');
  const dup = logs.some(l => l.EmpID === user.EmpID && l.Date === today && l.Type === type);
  if (dup) return {success: false, msg: `You already marked ${type} today.`};

  const lat = loc && loc.lat ? loc.lat : '';
  const lng = loc && loc.lng ? loc.lng : '';
  const mapLink = (lat && lng) ? `https://maps.google.com/?q=${lat},${lng}` : '';
  
  sheet.appendRow([today, user.EmpID, user.Name, type, time, lat, lng, mapLink, 'Web']);
  logAction_('Attendance', `${type} - ${user.Name}`);
  return {success: true, msg: `${type} Recorded.`};
}

function submitLeave(form) {
  const user = getActiveUserContext();
  if (user.error) return {success: false, msg: user.error};

  let days;
  if (form.session === 'First Half' || form.session === 'Second Half') {
      days = 0.5;
  } else {
      days = calculateLeaveDays(form.start, form.end);
  }

  if(isNaN(days) || days <= 0) return {success: false, msg: "Invalid Dates or 0 leave days."};

  // Balance Check Logic
  let balKey = "";
  if(form.type.includes("Casual")) balKey = "BalCL";
  else if(form.type.includes("Sick")) balKey = "BalSL";
  else if(form.type.includes("Maternity")) balKey = "BalMat";
  else if(form.type.includes("Paternity")) balKey = "BalPat";

  const allLeaves = getData('Leaves');
  const pendingLeaves = allLeaves.filter(l => l.EmpID === user.EmpID && l.Status === 'Pending' && l.Type === form.type);
  const pendingDays = pendingLeaves.reduce((acc, l) => acc + l.Days, 0);
  
  const availableBalance = user[balKey] - pendingDays;

  // Check against User's Context
  if (balKey && availableBalance < days) {
     return {success: false, msg: `Insufficient Balance! Available: ${availableBalance}, Requested: ${days} (includes ${pendingDays} pending days)`};
  }

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Leaves');
  sheet.appendRow(['LR-'+Date.now(), user.EmpID, user.Name, form.type, form.start, form.end, form.reason, 'Pending', days]);
  logAction_('Leave Apply', `${days} days ${form.type}`);
  
  // Email Notification to Manager
  try {
    const appUrl = ScriptApp.getService().getUrl();
    const employees = getData('Employees');
    const manager = employees.find(e => e.EmpID === user.ManagerID);
    const emailOptions = {
      htmlBody: '',
      name: 'HRMS (No-Reply)'
    };

    // 1. Send notification to Manager
    if (manager && manager.Email) {
      const subject = `Leave Application from ${user.Name}`;
      emailOptions.htmlBody = `
        <p>Hi ${manager.Name},</p>
        <p>${user.Name} has applied for leave. Please review the details below:</p>
        <ul>
          <li><b>Leave Type:</b> ${form.type}</li>
          <li><b>Dates:</b> ${form.start} to ${form.end}</li>
          <li><b>Days:</b> ${days}</li>
          <li><b>Reason:</b> ${form.reason || 'N/A'}</li>
        </ul>
        <p>You can approve or reject this leave from the HRMS portal by clicking the link below:</p>
        <p><a href="${appUrl}?page=leaves">Open HRMS Portal</a></p>
        <p>Thanks,<br>HRMS</p>
      `;
      MailApp.sendEmail(manager.Email, subject, "", emailOptions);
    }

    // 2. Send confirmation to User
    const userSubject = `Your Leave Application has been submitted`;
    emailOptions.htmlBody = `
      <p>Hi ${user.Name},</p>
      <p>Your leave application has been successfully submitted and is pending approval. Here are the details:</p>
      <ul>
        <li><b>Leave Type:</b> ${form.type}</li>
        <li><b>Dates:</b> ${form.start} to ${form.end}</li>
        <li><b>Days:</b> ${days}</li>
        <li><b>Reason:</b> ${form.reason || 'N/A'}</li>
      </ul>
      <p>You can view the status of your application in the HRMS portal:</p>
      <p><a href="${appUrl}?page=leaves">Open HRMS Portal</a></p>
      <p>Thanks,<br>HRMS</p>
    `;
    MailApp.sendEmail(user.Email, userSubject, "", emailOptions);

  } catch(e) {
    logAction_('Email Failed', `Leave notification failed for ${user.Name}. Error: ${e.message}`);
  }

  return {success: true, msg: "Leave Applied successfully"};
}

// Admin Action: Approve & Deduct Balance
function getLeaveDataForUser() {
    const user = getActiveUserContext();
    if (user.error) return { my_leaves: [], team_leaves: [] };

    const allLeaves = getData('Leaves');
    const myLeaves = allLeaves.filter(l => l.EmpID === user.EmpID);

    let teamLeaves = [];
    if (user.Role === 'ADMIN' || user.Role === 'HR') {
        // Admins or HR can see all pending leaves for approval
        teamLeaves = allLeaves.filter(l => l.Status === 'Pending');
    } else {
        // Managers see pending leaves for their direct reports
        const employees = getData('Employees');
        const myReportees = employees.filter(e => e.ManagerID === user.EmpID).map(e => e.EmpID);
        teamLeaves = allLeaves.filter(l => l.Status === 'Pending' && myReportees.includes(l.EmpID));
    }

    return { my_leaves: myLeaves, team_leaves: teamLeaves };
}

// Admin/Manager Action: Approve & Deduct Balance
function approveLeave(reqId) {
    const approver = getActiveUserContext();
    if (approver.error) return {success: false, msg: "Unauthorized"};

    const ss = SpreadsheetApp.openById(getDbId_());
    const lSheet = ss.getSheetByName('Leaves');
    const lData = lSheet.getDataRange().getValues();
    const lHeaders = lData[0];
    const reqIdIndex = lHeaders.indexOf('RequestID');
    const empIdIndex = lHeaders.indexOf('EmpID');
    const statusIndex = lHeaders.indexOf('Status');
    const daysIndex = lHeaders.indexOf('Days');
    const typeIndex = lHeaders.indexOf('Type');

    let leaveRow = -1;
    let leaveData = null;
    for (let i = 1; i < lData.length; i++) {
        if (String(lData[i][reqIdIndex]) === String(reqId)) {
            leaveRow = i + 1;
            leaveData = lData[i];
            break;
        }
    }

    if (!leaveData) return {success:false, msg:"Request not found"};
    if (leaveData[statusIndex] !== 'Pending') return {success:false, msg:"Already processed"};

    const requesterId = leaveData[empIdIndex];
    const employees = getData('Employees');
    const requester = employees.find(e => e.EmpID === requesterId);

    if (approver.Role !== 'ADMIN' && approver.Role !== 'HR' && requester.ManagerID !== approver.EmpID) {
        logAction_('Leave Approve Failed', `Unauthorized attempt by ${approver.EmpID} for ${reqId}`);
        return {success: false, msg: "You are not authorized to approve this request."};
    }

    // Deduct Balance
    const type = leaveData[typeIndex];
    const days = leaveData[daysIndex];
    const eSheet = ss.getSheetByName('Employees');
    const eData = eSheet.getDataRange().getValues();
    const eHeaders = eData[0];
    const empIdIndex_e = eHeaders.indexOf('EmpID');

    let balKey = "";
    if(type.includes("Casual")) balKey = "BalCL";
    else if(type.includes("Sick")) balKey = "BalSL";
    else if(type.includes("Maternity")) balKey = "BalMat";
    else if(type.includes("Paternity")) balKey = "BalPat";
    
    const colIdx = eHeaders.indexOf(balKey);

    if(colIdx > -1) {
       for(let j=1; j<eData.length; j++) {
           if(String(eData[j][empIdIndex_e]) === String(requesterId)) {
               const cur = eData[j][colIdx];
               eSheet.getRange(j+1, colIdx + 1).setValue(cur - days);
               break;
           }
       }
    }

    lSheet.getRange(leaveRow, statusIndex + 1).setValue('Approved');
    logAction_('Leave Approve', `Approved ${reqId} by ${approver.Email}`);
    // Send notification to requester about approval
    try {
      const requesterEmail = requester && requester.Email ? requester.Email : null;
      if (requesterEmail) {
        const subject = `Leave Approved: ${reqId}`;
        const body = `${approver.Name || approver.Email} has approved your leave request (${reqId}).\n\nType: ${type}\nDays: ${days}`;
        MailApp.sendEmail(requesterEmail, subject, body, { name: 'HRMS (No-Reply)'});
      }
    } catch(e) { logAction_('Email Failed', `Approval notification failed for ${reqId}: ${e.message}`); }

    return {success: true, msg: "Leave Approved & Balance Deducted"};
}

function createEmployee(form) {
  const employees = getData('Employees');
  if (employees.some(e => e.EmpID === form.empId)) {
    return {success: false, msg: "Employee with this ID already exists."};
  }
  if (employees.some(e => e.Email.toLowerCase() === form.email.toLowerCase())) {
    return {success: false, msg: "Employee with this email already exists."};
  }
  // Basic validation for mobile and DOB
  if (!form.mobile || !form.dob) {
    return {success: false, msg: "Mobile number and DOB are required."};
  }

  if (!/^\d{10}$/.test(form.mobile)) {
    return {success: false, msg: "Invalid mobile number. Please enter a 10-digit number."};
  }

  const dob = new Date(form.dob);
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
  if (age < 18) return {success: false, msg: "Employee must be at least 18 years old."};

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Employees');
  // Add with 0 CL balance and set LastBalanceUpdate to DOJ for pro-rata calculation
  sheet.appendRow([form.empId, form.name, form.email, form.role, form.dept, form.desig, form.doj, form.dob, form.mobile, 'Active', 0, 7, 0, 0, form.doj, form.managerId]);
  logAction_('Add Emp', form.name);
  return {success: true, msg: "Employee Added"};
}

function updateUserProfile(data) {
  const user = getActiveUserContext();
  if (user.error) return {success: false, msg: user.error};

  if (!data.mobile || !data.dob) {
    return {success: false, msg: "Mobile number and DOB are required."};
  }

  // Mobile number validation (10 digits)
  if (!/^\d{10}$/.test(data.mobile)) {
    return {success: false, msg: "Invalid mobile number. Please enter a 10-digit number."};
  }

  // DOB validation (at least 18 years old)
  const dob = new Date(data.dob);
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) {
    age--;
  }
  if (age < 18) {
    return {success: false, msg: "You must be at least 18 years old."};
  }

  try {
    const ss = SpreadsheetApp.openById(getDbId_());
    const sheet = ss.getSheetByName('Employees');
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const empIdIndex = headers.indexOf('EmpID');
    const mobileIndex = headers.indexOf('Mobile');
    const dobIndex = headers.indexOf('DOB');

    for (let i = 1; i < values.length; i++) {
      if (values[i][empIdIndex] === user.EmpID) {
        sheet.getRange(i + 1, mobileIndex + 1).setValue(data.mobile);
        sheet.getRange(i + 1, dobIndex + 1).setValue(data.dob);
        logAction_('Profile Update', `User ${user.Email} updated their profile.`);
        return {success: true, msg: "Profile updated successfully!"};
      }
    }
    return {success: false, msg: "Could not find your employee record to update."};
  } catch (e) {
    logAction_('Profile Update Error', e.message);
    return {success: false, msg: "An error occurred while updating: " + e.message};
  }
}

function submitDocument(form) {
  if (!form.type || !form.name || !form.url) {
    return {success: false, msg: "Please fill out all fields."};
  }
  const user = getActiveUserContext();
  if (user.error) return {success: false, msg: user.error};

  const empId = (user.Role === 'ADMIN' || user.Role === 'HR' && form.empId) ? form.empId : user.EmpID;

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Documents');
  sheet.appendRow(['DOC-'+Date.now(), empId, form.type, form.name, form.url, new Date()]);
  logAction_('Document Upload', `${form.name} for ${empId}`);
  return {success: true, msg: "Document Added"};
}

function createAnnouncement(form) {
  if (!form.title || !form.message) {
    return {success: false, msg: "Please fill out all fields."};
  }
  const user = getActiveUserContext();
  if (user.error || (user.Role !== 'ADMIN' && user.Role !== 'HR')) return {success: false, msg: "Unauthorized"};

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Announcements');
  sheet.appendRow([new Date(), form.title, form.message, user.Name]);
  logAction_('Announcement Post', form.title);

    // Notify all employees via email about the announcement
    try {
      const employees = getData('Employees');
      const subject = `Announcement: ${form.title}`;
      const htmlBody = `<p>Hi,</p><p>${form.message}</p><p>Posted by: ${user.Name}</p>`;
      employees.forEach(e => {
        try { if (e.Email) MailApp.sendEmail(e.Email, subject, '', { htmlBody: htmlBody, name: 'HRMS (No-Reply)' }); } catch(ignore) {}
      });
    } catch(e) {
      logAction_('Announcement Email Failed', e.message);
    }
  return {success: true, msg: "Announcement Posted"};
}

function createAsset(form) {
  if (!form.type || !form.model || !form.serial) {
    return {success: false, msg: "Please fill out all fields."};
  }
  const user = getActiveUserContext();
  if (user.error || (user.Role !== 'ADMIN' && user.Role !== 'HR')) return {success: false, msg: "Unauthorized"};

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Assets');
  sheet.appendRow(['ASSET-'+Date.now(), form.type, form.model, form.serial, '', 'Available']);
  logAction_('Asset Create', form.model);
  return {success: true, msg: "Asset Added"};
}

function assignAsset(form) {
  const user = getActiveUserContext();
  if (user.error || (user.Role !== 'ADMIN' && user.Role !== 'HR')) return {success: false, msg: "Unauthorized"};

  const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Assets');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const assetIdCol = headers.indexOf('AssetID');
  const assignedToCol = headers.indexOf('AssignedTo');
  const statusCol = headers.indexOf('Status');

  for (let i = 0; i < data.length; i++) {
    if (data[i][assetIdCol] === form.assetId) {
      if (data[i][statusCol] !== 'Available') {
        return {success: false, msg: "Asset is not available for assignment."};
      }
      sheet.getRange(i + 2, assignedToCol + 1).setValue(form.empId);
      sheet.getRange(i + 2, statusCol + 1).setValue('Assigned');
      logAction_('Asset Assign', `${form.assetId} to ${form.empId}`);
      return {success: true, msg: "Asset Assigned"};
    }
  }
  return {success: false, msg: "Asset not found"};
}

function generatePayroll(form) {
    if (!form.empId || !form.month || !form.year || !form.netPay) {
        return {success: false, msg: "Please fill out all fields."};
    }
    const user = getActiveUserContext();
    if (user.error || (user.Role !== 'ADMIN' && user.Role !== 'HR')) return {success: false, msg: "Unauthorized"};

    const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('Payroll');
    sheet.appendRow(['PAYSLIP-'+Date.now(), form.empId, form.month, form.year, form.netPay, new Date()]);
    logAction_('Payroll Generate', `For ${form.empId} - ${form.month}/${form.year}`);
    return {success: true, msg: "Payroll Generated"};
}



function getDocuments() {
    const user = getActiveUserContext();
    if (user.error) return [];

    const documents = getData('Documents');
    if (user.Role === 'ADMIN' || user.Role === 'HR') {
        return documents;
    }

    return documents.filter(doc => doc.EmpID === user.EmpID);
}

function getEmployees() {
    const employees = getData('Employees');
    return employees.map(e => {
        e.BalCL = Math.round((e.BalCL || 0) * 100) / 100;
        e.BalSL = Math.round((e.BalSL || 0) * 100) / 100;
        e.BalMat = Math.round((e.BalMat || 0) * 100) / 100;
        e.BalPat = Math.round((e.BalPat || 0) * 100) / 100;
        return e;
    });
}

function getGeneratedDocumentDataForUser() {
    const user = getActiveUserContext();
    if (user.error) return { my_docs: [], team_docs: [] };

    const allGeneratedDocs = getData('GeneratedDocuments');
    const myDocs = allGeneratedDocs.filter(d => d.EmpID === user.EmpID);

    let teamDocs = [];
    if (user.Role === 'ADMIN' || user.Role === 'HR') {
        teamDocs = allGeneratedDocs.filter(d => d.Status === 'Pending');
    } else {
        // Managers see pending leaves for their direct reports
        const employees = getData('Employees');
        const myReportees = employees.filter(e => e.ManagerID === user.EmpID).map(e => e.EmpID);
        teamDocs = allGeneratedDocs.filter(d => d.Status === 'Pending' && myReportees.includes(d.EmpID));
    }

    return { my_docs: myDocs, team_docs: teamDocs };
}

function fillTemplate(form) {
    const user = getActiveUserContext();
    if (user.error) return {success: false, msg: user.error};
    
    const sheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('GeneratedDocuments');
    const docId = 'GENDOC-' + Date.now();
    const status = 'Pending';
    const createdDate = new Date();
    
    // Store the form data as a JSON string
    const formData = JSON.stringify(form.data);

    sheet.appendRow([docId, form.templateId, user.EmpID, status, createdDate, '', formData, '']);
    logAction_('HR Form Filled', `Template: ${form.templateId}`);
    return {success: true, msg: 'Form submitted for approval.'};
}

function approveGeneratedDocument(docId) {
    const approver = getActiveUserContext();
    if (approver.error) return {success: false, msg: "Unauthorized"};

    const ss = SpreadsheetApp.openById(getDbId_());
    const sheet = ss.getSheetByName('GeneratedDocuments');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const docIdCol = headers.indexOf('GeneratedDocID');
    const empIdCol = headers.indexOf('EmpID');
    const statusCol = headers.indexOf('Status');
    const approvedDateCol = headers.indexOf('ApprovedDate');

    for (let i = 0; i < data.length; i++) {
        if (data[i][docIdCol] === docId) {
            const requesterId = data[i][empIdCol];
            if (requesterId === approver.EmpID && approver.Role !== 'ADMIN') {
                return {success: false, msg: "You cannot approve your own request."};
            }

            const employees = getData('Employees');
            const requester = employees.find(e => e.EmpID === requesterId);

            if (approver.Role !== 'ADMIN' && approver.Role !== 'HR' && requester.ManagerID !== approver.EmpID) {
                logAction_('Doc Approve Failed', `Unauthorized attempt by ${approver.EmpID} for ${docId}`);
                return {success: false, msg: "You are not authorized to approve this request."};
            }

            sheet.getRange(i + 2, statusCol + 1).setValue('Approved');
            sheet.getRange(i + 2, approvedDateCol + 1).setValue(new Date());
            logAction_('HR Form Approved', `Doc ID: ${docId} by ${approver.Email}`);
            return {success: true, msg: "Document Approved"};
        }
    }
    return {success: false, msg: "Document not found"};
}

function generatePdf(docId) {
    const user = getActiveUserContext();
    if (user.error) return {success: false, msg: user.error};

    const genDocSheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('GeneratedDocuments');
    const genDocData = getData('GeneratedDocuments').find(d => d.GeneratedDocID === docId);

    if (!genDocData || genDocData.Status !== 'Approved') {
        return {success: false, msg: "Document not found or not approved."};
    }
    
    const templateSheet = SpreadsheetApp.openById(getDbId_()).getSheetByName('DocumentTemplates');
    const templateData = getData('DocumentTemplates').find(t => t.TemplateID === genDocData.TemplateID);

    if (!templateData) {
        return {success: false, msg: "Template not found."};
    }

    let htmlContent = templateData.Content;
    const formData = JSON.parse(genDocData.Data);

    // Replace placeholders
    for (const key in formData) {
        const placeholder = new RegExp(`{{${key}}}`, 'g');
        htmlContent = htmlContent.replace(placeholder, formData[key]);
    }
     htmlContent = htmlContent.replace(/{{current_date}}/g, new Date().toLocaleDateString());

    const pdfBlob = Utilities.newBlob(htmlContent, 'text/html', `${genDocData.TemplateID}-${genDocData.EmpID}.html`).getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob).setName(`${genDocData.TemplateID}-${genDocData.EmpID}.pdf`);
    const fileUrl = pdfFile.getUrl();

    // Update the sheet with the file URL
    const data = genDocSheet.getDataRange().getValues();
    const headers = data.shift();
    const docIdCol = headers.indexOf('GeneratedDocID');
    const fileUrlCol = headers.indexOf('FileURL');

    for (let i = 0; i < data.length; i++) {
        if (data[i][docIdCol] === docId) {
            genDocSheet.getRange(i + 2, fileUrlCol + 1).setValue(fileUrl);
            break;
        }
    }

    logAction_('PDF Generated', `Doc ID: ${docId}`);
    return {success: true, url: fileUrl};
}



function getDashboardStats() {
  const e = getData('Employees');
  const l = getData('Leaves');
  const logs = getData('SystemLogs');
  const ann = getData('Announcements');
  
  return {
    empCount: e.length,
    pendingLeaves: l.filter(r => r.Status === 'Pending').length,
    logs: logs.reverse().slice(0, 10),
    announcements: ann.reverse().slice(0, 5)
  };
}

function updateSettings(settings) {
  const user = getActiveUserContext();
  if (user.error || user.Role !== 'ADMIN') {
    return {success: false, msg: "Unauthorized"};
  }

  try {
    const ss = SpreadsheetApp.openById(getDbId_());
    const sheet = ss.getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const settingIndex = headers.indexOf('Setting');
    const valueIndex = headers.indexOf('Value');

    for (const key in settings) {
      let found = false;
      for (let i = 0; i < data.length; i++) {
        if (data[i][settingIndex] === key) {
          sheet.getRange(i + 2, valueIndex + 1).setValue(settings[key]);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, settings[key]]);
      }
    }
    logAction_('Settings Update', `User ${user.Email} updated settings.`);
    return {success: true, msg: "Settings updated successfully!"};
  } catch (e) {
    logAction_('Settings Update Error', e.message);
    return {success: false, msg: "An error occurred while updating settings: " + e.message};
  }
}