// ============================================================
// LEAVE & ATTENDANCE MANAGEMENT SYSTEM
// Main Backend Controller: Code.gs — OOP REFACTORED
// ============================================================

const SHEETS = {
  EMPLOYEES:     'Employees',
  ATTENDANCE:    'Attendance',
  LEAVE_APPS:    'LeaveApplications',
  LEAVE_BALANCE: 'LeaveBalance',
  SALARY_CALC:   'SalaryCalculations',
  AUDIT_LOG:     'AuditLog',
  CONFIG:        'Config'
};

const ROLES = { EMPLOYEE:'employee', OFFICE_STAFF:'office_staff', SUPER_ADMIN:'super_admin' };

const ADMIN_EMAILS = ['sharanyabaidya2020@gmail.com','khushi.srivastava@gyws.org','aaratrika.bhattacharya@gyws.org'];

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Leave & Attendance System')
    .addMetaTag('viewport','width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTodayIST() {
  const d = new Date();
  const formatted = Utilities.formatDate(d, 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss");
  return new Date(formatted);
}

function hashPassword(pw) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw));
}

function getSuperAdmins() {
  if (typeof SUPER_ADMINS !== 'undefined') {
    return SUPER_ADMINS;
  }
  return [];
}

// ============================================================
// AUTHENTICATION
// ============================================================
function ensureAllPasswordsSet() {
  try {
    const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
    const adminHash = hashPassword('admin123');
    const hrHash = hashPassword('hr123');
    const empHash = hashPassword('pass123');
    
    let updated = false;
    const emps = empRepo.records;
    
    for (let i = 0; i < emps.length; i++) {
      const d = emps[i].item;
      if (!d.employeeid) continue;
      
      if (!d.passwordhash || String(d.passwordhash).trim() === '') {
        const role = String(d.role).trim().toLowerCase();
        let correctHash = empHash;
        if (role === 'super_admin') {
          correctHash = adminHash;
        } else if (role === 'hr' || role === 'office_staff') {
          correctHash = hrHash;
        }
        empRepo.update(emps[i].rowIndex, { passwordhash: correctHash });
        updated = true;
      }
    }
    if (updated) {
      SpreadsheetApp.flush();
    }
  } catch(e) {
    Logger.log('Error in ensureAllPasswordsSet: ' + e.message);
  }
}

function login(empId, password) {
  try {
    ensureAllPasswordsSet();
    const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
    const empData = empRepo.findOne(e => String(e.employeeid) === String(empId));
    
    if (empData) {
      const emp = new Employee(empData.item);
      if (emp.passwordHash === hashPassword(password) && emp.isActive) {
        const user = {id: emp.id, name: emp.name, email: emp.email, department: emp.department, role: emp.role, salary: emp.salary};
        logAudit(user.id, user.name, 'LOGIN', 'User logged in');
        return {success: true, user};
      }
    }
    
    // Check SUPER_ADMINS
    const superAdminsList = getSuperAdmins();
    for (let i = 0; i < superAdminsList.length; i++) {
      if (superAdminsList[i].empId === empId && superAdminsList[i].password === password) {
        const user = {id: superAdminsList[i].empId, name: superAdminsList[i].name, email: superAdminsList[i].emails[0], department: 'Management', role: 'super_admin', salary: 0};
        logAudit(user.id, user.name, 'LOGIN', 'Super Admin logged in');
        return {success: true, user};
      }
    }
    
    return {success: false, message: 'Invalid Employee ID or password.'};
  } catch(e) { return {success: false, message: e.message}; }
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboardData(empId, role) {
  try {
    const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);
    const year = getTodayIST().getFullYear();
    
    const hasBalance = balRepo.findOne(b => String(b.employeeid) === String(empId) && safeParseInt(b.year) === parseInt(year));
    if (!hasBalance) {
      rebuildLeaveBalances();
      balRepo.load();
    }

    const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const attRepo = new SheetRepository(SHEETS.ATTENDANCE);

    const balanceData = balRepo.findOne(b => String(b.employeeid) === String(empId) && safeParseInt(b.year) === parseInt(year));
    const balance = balanceData ? new LeaveBalanceRecord(balanceData.item) : null;
    
    const attendanceData = attRepo.find(a => String(a.employeeid) === String(empId) && safeParseInt(a.year) === parseInt(year));
    const attendance = attendanceData.map(a => new AttendanceRecord(a.item));

    const myLeavesData = leaveRepo.find(l => String(l.employeeid) === String(empId));
    const myLeaves = myLeavesData.map(l => new LeaveApplication(l.item)).reverse();

    const isSuperOrHr = (role === 'super_admin' || role === 'hr');

    if (isSuperOrHr) {
      const allLeaves = leaveRepo.getAll().map(l => new LeaveApplication(l)).reverse();
      const pendingLeaves = allLeaves.filter(l => l.status === 'Pending');
      const allEmployees = empRepo.getAll().map(e => new Employee(e)).filter(emp => emp.isActive && emp.role !== 'super_admin');
      
      const today = getTodayIST();
      const currentMonth = ['January','February','March','April','May','June','July','August','September','October','November','December'][today.getMonth()];
      const currentMonthAttData = attRepo.find(a => parseInt(a.year) === year && a.month === currentMonth);
      
      const empMap = {};
      empRepo.getAll().forEach(e => {
        empMap[e.employeeid] = e.employeename;
      });

      const currentMonthAtt = currentMonthAttData.map(a => {
        const record = new AttendanceRecord(a.item);
        return {
          ...record,
          employeeName: empMap[record.employeeId] || record.employeeId
        };
      });

      return toPlainObject({
        balance, attendance, myLeaves,
        allLeaves, pendingLeaves, allEmployees, currentMonthAttendance: currentMonthAtt
      });
    } else {
      return toPlainObject({
        balance, attendance, myLeaves,
        allLeaves:[], pendingLeaves:[], allEmployees:[], currentMonthAttendance:[]
      });
    }
  } catch(err) {
    Logger.log(err);
    return {balance:{},attendance:[],myLeaves:[],allLeaves:[],pendingLeaves:[],allEmployees:[],currentMonthAttendance:[]};
  }
}

// ============================================================
// EMPLOYEES
// ============================================================
function getEmployeeList() {
  ensureAllPasswordsSet();
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  return empRepo.find(e => (e.isactive === true || e.isactive === 'TRUE' || String(e.isactive).toLowerCase() === 'true') && e.role !== 'super_admin')
    .map(wrapper => {
      const emp = new Employee(wrapper.item);
      return {id: emp.id, name: emp.name, email: emp.email, department: emp.department, role: emp.role};
    });
}

function getEmployeeById(empId) {
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  const match = empRepo.findOne(e => String(e.employeeid) === String(empId));
  if (match) {
    return new Employee(match.item);
  }
  
  // Check SUPER_ADMINS
  const superAdminsList = getSuperAdmins();
  for (let i = 0; i < superAdminsList.length; i++) {
    if (superAdminsList[i].empId === empId) {
      return new Employee({
        employeeid: superAdminsList[i].empId,
        employeename: superAdminsList[i].name,
        email: superAdminsList[i].emails[0],
        department: 'Management',
        role: 'super_admin',
        salary: 0,
        phone: ''
      });
    }
  }
  return null;
}

// ============================================================
// LEAVE APPLICATIONS
// ============================================================
function getLeaveApplications(filter) {
  try {
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const filterObj = filter || {};
    
    const fmt = (v) => v instanceof Date
      ? Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd')
      : (v ? String(v) : '');

    const apps = leaveRepo.find(l => {
      if (filterObj.employeeId && String(l.employeeid) !== String(filterObj.employeeId)) return false;
      if (filterObj.status && String(l.status).trim() !== String(filterObj.status).trim()) return false;
      return true;
    }).map(wrapper => {
      const leave = new LeaveApplication(wrapper.item);
      return {
        appId: leave.appId,
        employeeId: leave.employeeId,
        employeeName: leave.employeeName,
        appliedBy: leave.appliedBy,
        leaveType: leave.leaveType,
        fromDate: fmt(leave.fromDate),
        toDate: fmt(leave.toDate),
        totalDays: leave.totalDays,
        reason: leave.reason,
        status: leave.status,
        appliedOn: fmt(leave.appliedOn),
        approvedBy: leave.approvedBy,
        approvedOn: fmt(leave.approvedOn),
        adminNote: leave.adminNote,
        isHalfDay: leave.isHalfDay,
        referenceNo: leave.referenceNo,
        markedAbsent: leave.markedAbsent
      };
    });
    return apps.reverse();
  } catch(e) { Logger.log(e); return []; }
}

function submitLeaveApplication(appData) {
  try {
    if (!appData.referenceNo || !appData.referenceNo.trim()) {
      return {success: false, message: 'Reference Number is required.'};
    }

    const config = getConfig();
    const configAdvance = parseInt(config.advance_days_required) || 3;
    
    // Strategy based notice and balance validation
    const leaveStrategy = LeaveStrategyFactory.getStrategy(appData.leaveType);
    const advanceDays = leaveStrategy.getRequiredAdvanceNoticeDays(configAdvance);

    const fromDate = new Date(appData.fromDate);
    const today = getTodayIST();
    
    const fromMidnight = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
    const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const diffDays = Math.round((fromMidnight.getTime() - todayMidnight.getTime()) / 86400000);

    if (!['Medical Leave','Emergency Leave'].includes(appData.leaveType) && diffDays < advanceDays) {
      return {success: false, message: `Leave must be applied at least ${advanceDays} days in advance.`};
    }

    // Check balance
    const year = getTodayIST().getFullYear();
    const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);
    const balData = balRepo.findOne(b => String(b.employeeid) === String(appData.employeeId) && safeParseInt(b.year) === parseInt(year));
    if (!balData) {
      return {success: false, message: 'Leave balance record not found. Please contact admin to initialize balance.'};
    }
    
    const balance = new LeaveBalanceRecord(balData.item);
    const empDetails = getEmployeeById(appData.employeeId);
    if (!empDetails) {
      return {success: false, message: 'Employee details not found.'};
    }
    const emp = new Employee(empDetails);
    const maxLeaves = emp.getMaxAnnualLeaves(parseInt(config.max_leaves_per_year) || 14);

    if (leaveStrategy.requiresBalanceCheck()) {
      if (balance.approvedLeaveUsed + parseFloat(appData.totalDays) > maxLeaves) {
        return {success: false, message: `Insufficient leave balance. Available: ${maxLeaves-balance.approvedLeaveUsed} days.`};
      }
    }

    const appId = generateId('APP');
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    
    const newApp = {
      appid: appId,
      employeeid: appData.employeeId,
      employeename: appData.employeeName,
      appliedby: appData.appliedBy,
      leavetype: appData.leaveType,
      fromdate: appData.fromDate,
      todate: appData.toDate,
      totaldays: appData.totalDays,
      reason: appData.reason,
      status: 'Pending',
      appliedon: Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd'),
      approvedby: '',
      approvedon: '',
      adminnote: '',
      ishalfday: appData.isHalfDay || 'FALSE',
      referenceno: appData.referenceNo || '',
      markedabsent: ''
    };

    leaveRepo.append(newApp);

    sendLeaveNotification(appData, appId);
    logAudit(appData.appliedBy, appData.appliedBy, 'LEAVE_APPLY', `Applied ${appId} for ${appData.employeeName}`);
    return {success: true, appId, message: 'Leave application submitted successfully!'};
  } catch(e) { return {success: false, message: e.message}; }
}

function approveRejectLeave(appId, decision, adminId, adminNote, adminName, adminEmail, markAbsent) {
  try {
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const match = leaveRepo.findOne(l => String(l.appid) === String(appId));
    if (!match) {
      return {success: false, message: 'Application not found.'};
    }

    const leave = new LeaveApplication(match.item);
    const today = Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd');

    leave.status = decision;
    leave.approvedBy = adminId;
    leave.approvedOn = today;
    leave.adminNote = adminNote || '';

    const cleanDecision = decision ? String(decision).trim().toLowerCase() : '';
    const leaveStrategy = LeaveStrategyFactory.getStrategy(leave.leaveType);

    let countedIn14 = false;

    if (cleanDecision === 'approved') {
      leave.markedAbsent = ''; 
      leaveRepo.update(match.rowIndex, leave.toRowObject());
      updateLeaveBalance(leave.employeeId, leave.leaveType, leave.totalDays, 'Approved');
      countedIn14 = leaveStrategy.countsTowardsEntitlementLimit();
    } else if (cleanDecision === 'rejected') {
      if (!leaveStrategy.countsTowardsEntitlementLimit()) {
        leave.markedAbsent = ''; 
        leaveRepo.update(match.rowIndex, leave.toRowObject());
        updateLeaveBalance(leave.employeeId, leave.leaveType, leave.totalDays, 'Rejected');
        countedIn14 = true;
      } else {
        leave.markedAbsent = 'PENDING';
        leaveRepo.update(match.rowIndex, leave.toRowObject());
        rebuildLeaveBalances();
        countedIn14 = false; 
      }
    }

    const empDetails = getEmployeeById(leave.employeeId);
    const emp = empDetails ? new Employee(empDetails) : null;
    const dynamicLimit = emp ? emp.getMaxAnnualLeaves(parseInt(getConfig().max_leaves_per_year)||14) : 14;

    sendDecisionNotification({
      appId,
      decision,
      adminId,
      adminName:    adminName || adminId,
      adminEmail:   adminEmail || '',
      adminNote:    adminNote || '',
      employeeId:   leave.employeeId,
      employeeName: leave.employeeName,
      leaveType:    leave.leaveType,
      fromDate:     leave.fromDate,
      toDate:       leave.toDate,
      totalDays:    leave.totalDays,
      reason:       leave.reason,
      appliedOn:    leave.appliedOn,
      referenceNo:  leave.referenceNo,
      countedIn14:  countedIn14,
      limit:        dynamicLimit,
      markAbsent:   false
    });

    logAudit(adminId, adminId, `LEAVE_${decision.toUpperCase()}`, `${decision} leave ${appId}`);
    return {success: true, message: `Leave ${decision} successfully.`};
  } catch(e) { return {success: false, message: e.message}; }
}

function markRejectedCasualOutcome(appId, outcome, adminId) {
  try {
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const match = leaveRepo.findOne(l => String(l.appid) === String(appId));
    if (!match) {
      return {success: false, message: 'Application not found.'};
    }

    const leave = new LeaveApplication(match.item);
    if (leave.status !== 'Rejected') {
      return {success: false, message: 'Only rejected leaves can be marked.'};
    }

    leave.markedAbsent = outcome;
    leaveRepo.update(match.rowIndex, leave.toRowObject());
    SpreadsheetApp.flush();

    try {
      const fromDateObj = new Date(leave.fromDate);
      if (!isNaN(fromDateObj.getTime())) {
        const monthName = fromDateObj.toLocaleString('default', { month: 'long' });
        const yearVal = fromDateObj.getFullYear();

        const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
        const attMatch = attRepo.findOne(a => String(a.employeeid) === String(leave.employeeId) && String(a.month).toUpperCase() === monthName.toUpperCase() && safeParseInt(a.year) === parseInt(yearVal));

        if (attMatch) {
          const att = new AttendanceRecord(attMatch.item);
          const prefill = getAttendancePrefillData(leave.employeeId, monthName, yearVal);
          if (prefill && prefill.success) {
            const cl = prefill.approvedLeave || 0;
            const ml = prefill.medicalLeave || 0;
            const el = prefill.emergencyLeave || 0;
            const rejectedCasual = prefill.rejectedCasualLeave || 0;
            const rejectedMedOrEmg = prefill.rejectedMedOrEmgLeave || 0;

            const totalRejected = rejectedCasual + rejectedMedOrEmg;
            let remainingAbsent = att.workingDays - att.daysPresent - cl - ml - el - totalRejected;
            if (remainingAbsent < 0) remainingAbsent = 0;
            const calculatedAbsents = rejectedCasual + remainingAbsent;

            att.approvedLeave = cl;
            att.medicalLeave = ml;
            att.emergencyLeave = el;
            att.daysAbsent = calculatedAbsents;
            att.rejectedMedOrEmg = rejectedMedOrEmg;

            attRepo.update(attMatch.rowIndex, att.toRowObject());
          }
        }
      }
    } catch (attErr) {
      Logger.log('Failed to update attendance row during markOutcome: ' + attErr.message);
    }

    rebuildLeaveBalances();
    const label = outcome === 'PRESENT' ? 'Present (scrapped)' : 'Absent (counted in limit + penalty)';
    logAudit(adminId, adminId, 'MARK_REJECTED_CASUAL', `Marked ${appId} as ${label}`);
    return {success: true, message: `Marked as ${label} successfully.`};
  } catch(e) { return {success: false, message: e.message}; }
}

// ============================================================
// ATTENDANCE
// ============================================================
function saveAttendance(data) {
  try {
    const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
    const match = attRepo.findOne(a => a.employeeid === data.employeeId && a.month === data.month && parseInt(a.year) === parseInt(data.year));

    const record = {
      employeeid: data.employeeId,
      month: data.month,
      year: parseInt(data.year),
      workingdays: parseFloat(data.workingDays) || 0,
      dayspresent: parseFloat(data.daysPresent) || 0,
      approvedleave: parseFloat(data.approvedLeave) || 0,
      medicalleave: parseFloat(data.medicalLeave) || 0,
      emergencyleave: parseFloat(data.emergencyLeave) || 0,
      daysabsent: parseFloat(data.daysAbsent) || 0,
      rejectedmedoremg: parseFloat(data.rejectedMedOrEmg) || 0,
      enteredby: data.enteredBy,
      timestamp: new Date()
    };

    if (match) {
      attRepo.update(match.rowIndex, record);
      logAudit(data.enteredBy, data.enteredBy, 'ATTENDANCE_UPDATE', `Updated ${data.month} ${data.year} for ${data.employeeId}`);
    } else {
      record.attid = generateId('ATT');
      attRepo.append(record);
      logAudit(data.enteredBy, data.enteredBy, 'ATTENDANCE_SAVE', `Saved ${data.month} ${data.year} for ${data.employeeId}`);
    }

    rebuildLeaveBalances();
    return {success: true, message: 'Attendance saved.'};
  } catch(e) { return {success: false, message: e.message}; }
}

function getAttendance(employeeId, year) {
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const matched = attRepo.find(a => {
    if (employeeId && String(a.employeeid) !== String(employeeId)) return false;
    if (year && parseInt(a.year) !== parseInt(year)) return false;
    return true;
  }).map(wrapper => new AttendanceRecord(wrapper.item));
  return toPlainObject(matched);
}

function getAllAttendance(month, year) {
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  
  const empMap = {};
  empRepo.getAll().forEach(e => {
    empMap[e.employeeid] = e.employeename;
  });

  const records = attRepo.find(a => {
    if (year && parseInt(a.year) !== parseInt(year)) return false;
    if (month && String(a.month).toUpperCase() !== String(month).toUpperCase()) return false;
    return true;
  }).map(wrapper => {
    const r = new AttendanceRecord(wrapper.item);
    return {
      ...r,
      employeeName: empMap[r.employeeId] || r.employeeId,
      medicalLeave: r.medicalLeave || 0,
      emergencyLeave: r.emergencyLeave || 0,
      rejectedMedOrEmg: r.rejectedMedOrEmg || 0
    };
  });
  return records;
}

function getAttendancePrefillData(employeeId, month, year) {
  try {
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const leaves = leaveRepo.find(l => String(l.employeeid).trim() === String(employeeId).trim());

    let approvedLeave = 0;
    let medicalLeave = 0;
    let emergencyLeave = 0;
    let rejectedCasualLeave = 0;
    let rejectedMedOrEmgLeave = 0;

    leaves.forEach(wrapper => {
      const leave = new LeaveApplication(wrapper.item);
      const fromDate = new Date(leave.fromDate);
      if (isNaN(fromDate.getTime())) return;
      
      const rowMonth = fromDate.toLocaleString('default', { month: 'long' });
      const rowYear = fromDate.getFullYear();

      if (rowMonth !== month || String(rowYear) !== String(year)) {
        return;
      }

      const days = leave.totalDays;

      if (leave.status === 'Approved') {
        if (leave.leaveType === 'Medical Leave') {
          medicalLeave += days;
        } else if (leave.leaveType === 'Emergency Leave') {
          emergencyLeave += days;
        } else {
          approvedLeave += days;
        }
      } else if (leave.status === 'Rejected') {
        if (leave.leaveType === 'Medical Leave' || leave.leaveType === 'Emergency Leave') {
          rejectedMedOrEmgLeave += days;
        } else {
          if (leave.markedAbsent !== 'PRESENT') {
            rejectedCasualLeave += days;
          }
        }
      }
    });

    return {
      success: true,
      approvedLeave,
      medicalLeave,
      emergencyLeave,
      rejectedCasualLeave,
      rejectedMedOrEmgLeave
    };
  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}

// ============================================================
// LEAVE BALANCE
// ============================================================
function getLeaveBalance(employeeId, year) {
  const y = year || getTodayIST().getFullYear();
  const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);
  const match = balRepo.findOne(b => String(b.employeeid).trim() === String(employeeId).trim() && safeParseInt(b.year) === parseInt(y));
  
  if (match) {
    return new LeaveBalanceRecord(match.item);
  }
  
  let defaultMaxLeaves = 14;
  let empName = '';
  const empDetails = getEmployeeById(employeeId);
  if (empDetails) {
    const emp = new Employee(empDetails);
    empName = emp.name;
    defaultMaxLeaves = emp.getMaxAnnualLeaves(14);
  }
  
  const newBal = {
    employeeid: employeeId,
    employeename: empName,
    year: y,
    totalentitled: defaultMaxLeaves,
    approvedleaveused: 0,
    medicalleaveused: 0,
    emergencyleaveused: 0,
    unapprovedabsentused: 0,
    rejectedcasualused: 0,
    rejectedmedoremgused: 0,
    carryforward: 0,
    unusedleaves: defaultMaxLeaves,
    leavesurplus: 0,
    lastupdated: new Date()
  };
  
  balRepo.append(newBal);
  return new LeaveBalanceRecord(newBal);
}

function updateLeaveBalance(employeeId, leaveType, days, decision) {
  const y = getTodayIST().getFullYear();
  const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);
  const match = balRepo.findOne(b => String(b.employeeid) === String(employeeId) && safeParseInt(b.year) === parseInt(y));
  
  if (match) {
    const bal = new LeaveBalanceRecord(match.item);
    const cleanDecision = decision ? String(decision).trim().toLowerCase() : '';
    const cleanType = leaveType ? String(leaveType).trim().toLowerCase() : '';

    if (cleanDecision === 'approved') {
      if (cleanType === 'medical leave') {
        bal.medicalLeaveUsed += days;
      } else if (cleanType === 'emergency leave') {
        bal.emergencyLeaveUsed += days;
      } else {
        bal.approvedLeaveUsed += days;
      }
    } else if (cleanDecision === 'rejected') {
      if (cleanType === 'medical leave' || cleanType === 'emergency leave') {
        bal.rejectedMedOrEmgUsed += days;
      } else {
        bal.rejectedCasualUsed += days;
      }
    }

    const usedCount = bal.approvedLeaveUsed + bal.unapprovedAbsentUsed + bal.rejectedCasualUsed + bal.rejectedMedOrEmgUsed;
    const rawRemaining = bal.totalEntitled - usedCount;
    bal.unusedLeaves = Math.max(0, rawRemaining);
    bal.leaveSurplus = rawRemaining < 0 ? Math.abs(rawRemaining) : 0;
    bal.lastUpdated = new Date();

    balRepo.update(match.rowIndex, bal.toRowObject());
  }
}

// ============================================================
// SALARY CALCULATION
// ============================================================
function calculateSalary(employeeId, month, year) {
  try {
    const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
    const empData = empRepo.findOne(e => String(e.employeeid) === String(employeeId));
    if (!empData) return {success: false, message: 'Employee not found.'};
    
    const emp = new Employee(empData.item);
    const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
    const attRecordsData = attRepo.find(a => a.employeeid === employeeId && parseInt(a.year) === parseInt(year));
    const attRecords = attRecordsData.map(a => new AttendanceRecord(a.item));

    let att = attRecords.find(a => a.month === month && a.year == year);
    if (!att) {
      att = new AttendanceRecord({
        attid: '',
        employeeid: employeeId,
        month: month,
        year: year,
        workingdays: 30,
        dayspresent: 30,
        approvedleave: 0,
        medicalleave: 0,
        emergencyleave: 0,
        daysabsent: 0,
        rejectedmedoremg: 0
      });
      attRecords.push(att);
    }

    const base = emp.salary;
    const perDay = emp.getDailyRate();
    const approvedCasual = att.approvedLeave;
    const medLeave = att.medicalLeave;
    const emgLeave = att.emergencyLeave;
    const rawAbsent = att.daysAbsent;

    // Count rejected leave days (split by type) for this employee in this month/year
    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const leaveRecords = leaveRepo.find(l => l.employeeid === employeeId && l.status === 'Rejected');
    
    let rejectedCasual = 0;
    let rejectedMedicalEmergency = 0;

    leaveRecords.forEach(wrapper => {
      const leave = new LeaveApplication(wrapper.item);
      const fromDate = new Date(leave.fromDate);
      if (isNaN(fromDate.getTime())) return;
      const rowYear = fromDate.getFullYear();
      const rowMonth = fromDate.toLocaleString('default', { month: 'long' });

      if (rowYear == year && rowMonth === month) {
        const leaveStrategy = LeaveStrategyFactory.getStrategy(leave.leaveType);
        if (!leaveStrategy.countsTowardsEntitlementLimit()) {
          rejectedMedicalEmergency += leave.totalDays;
        } else {
          if (leave.markedAbsent !== 'PRESENT') {
            rejectedCasual += leave.totalDays;
          }
        }
      }
    });

    const pureAbsent = rawAbsent; 

    const MONTHS_ORDER = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const targetMonthIndex = MONTHS_ORDER.indexOf(month);

    let priorUsed = 0;
    attRecords.forEach(a => {
      const idx = MONTHS_ORDER.indexOf(a.month);
      if (idx !== -1 && idx < targetMonthIndex) {
        priorUsed += (a.approvedLeave + a.rejectedMedOrEmg + a.daysAbsent);
      }
    });

    const targetMonthUsed = approvedCasual + rejectedMedicalEmergency + pureAbsent;
    const totalUsed = priorUsed + targetMonthUsed;
    
    const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);
    const balData = balRepo.findOne(b => String(b.employeeid) === String(employeeId) && safeParseInt(b.year) === parseInt(year));
    const balance = balData ? new LeaveBalanceRecord(balData.item) : null;
    const remaining = balance ? balance.unusedLeaves : emp.getMaxAnnualLeaves(14);

    const maxLeaves = emp.getMaxAnnualLeaves(parseInt(getConfig().max_leaves_per_year) || 14);
    const extraDays = Math.max(0, totalUsed - maxLeaves) - Math.max(0, priorUsed - maxLeaves);
    
    attRecords.sort((x, y) => MONTHS_ORDER.indexOf(x.month) - MONTHS_ORDER.indexOf(y.month));
    
    let cumulativeUsed = 0;
    const historicalLeaves = attRecords.map(a => {
      const monthUsed = a.approvedLeave + a.rejectedMedOrEmg + a.daysAbsent;
      cumulativeUsed += monthUsed;
      const monthRemaining = Math.max(0, maxLeaves - cumulativeUsed);
      
      return {
        month: a.month,
        approvedLeave: a.approvedLeave,
        rejectedMedOrEmg: a.rejectedMedOrEmg,
        daysAbsent: a.daysAbsent,
        total: monthUsed,
        remaining: monthRemaining
      };
    });

    const extraDeduct = extraDays * perDay;
    const absentDeduct = pureAbsent * 0.5 * perDay; 
    const totalDeduct = extraDeduct + absentDeduct;
    const netPayable = base - totalDeduct;

    const salaryRepo = new SheetRepository(SHEETS.SALARY_CALC);
    salaryRepo.append({
      calcid: generateId('SAL'),
      employeeid: employeeId,
      employeename: emp.name,
      month: month,
      year: year,
      basesalary: base,
      approvedleavedays: approvedCasual,
      absentdays: pureAbsent,
      halfdays: 0, 
      medicaldays: medLeave,
      emergencydays: emgLeave,
      extradays: extraDays,
      absentdeduction: parseFloat(absentDeduct.toFixed(2)),
      extraleavededuction: parseFloat(extraDeduct.toFixed(2)),
      unusedleavebonus: remaining,
      netpayable: parseFloat(netPayable.toFixed(2)),
      calculatedon: new Date()
    });

    return {success: true, data: {
      employeeId, employeeName: emp.name, month, year, baseSalary: base,
      approvedCasual, rejectedMedicalEmergency, pureAbsent,
      medicalDays: medLeave, emergencyDays: emgLeave,
      totalUsed, remaining, extraDays,
      perDay: perDay.toFixed(2),
      absentDeduction: absentDeduct.toFixed(2),
      extraLeaveDeduction: extraDeduct.toFixed(2),
      netPayable: netPayable.toFixed(2),
      maxLeaves: maxLeaves,
      historicalLeaves: historicalLeaves
    }};
  } catch(e) { return {success: false, message: e.message}; }
}

function getSalaryCalcForMonth(month, year) {
  const salaryRepo = new SheetRepository(SHEETS.SALARY_CALC);
  const results = salaryRepo.find(s => String(s.month).toUpperCase() === String(month).toUpperCase() && parseInt(s.year) === parseInt(year))
    .map(wrapper => {
      const data = wrapper.item;
      return {
        calcId: data.calcid,
        employeeId: data.employeeid,
        employeeName: data.employeename,
        month: data.month,
        year: data.year,
        baseSalary: parseFloat(data.basesalary) || 0,
        approvedLeaveDays: parseFloat(data.approvedleavedays) || 0,
        absentDays: parseFloat(data.absentdays) || 0,
        halfDays: parseFloat(data.halfdays) || 0,
        medicalDays: parseFloat(data.medicaldays) || 0,
        emergencyDays: parseFloat(data.emergencydays) || 0,
        extraDays: parseFloat(data.extradays) || 0,
        absentDeduction: parseFloat(data.absentdeduction) || 0,
        extraLeaveDeduction: parseFloat(data.extraleavededuction) || 0,
        unusedLeaveBonus: parseFloat(data.unusedleavebonus) || 0,
        netPayable: parseFloat(data.netpayable) || 0
      };
    });
  return results;
}

function bulkCalculateSalary(month, year) {
  const emps = getEmployeeList();
  let count = 0;
  emps.forEach(emp => {
    if (emp.role !== 'super_admin') {
      calculateSalary(emp.id, month, year);
      count++;
    }
  });
  return {success: true, count};
}

// ============================================================
// AUDIT & CONFIG
// ============================================================
function logAudit(userId, userName, action, details) {
  try {
    const auditRepo = new SheetRepository(SHEETS.AUDIT_LOG);
    auditRepo.append({
      timestamp: new Date(),
      userid: userId,
      username: userName,
      action: action,
      details: details,
      extra: ''
    });
  } catch(e) {}
}

function getConfig() {
  const configRepo = new SheetRepository(SHEETS.CONFIG);
  const cfg = {};
  configRepo.getAll().forEach(row => {
    const keys = Object.keys(row);
    if (keys.length >= 2) {
      cfg[row[keys[0]]] = row[keys[1]];
    }
  });
  return cfg;
}

function generateId(prefix) {
  return prefix+Utilities.formatDate(new Date(),'Asia/Kolkata','yyyyMMddHHmmss')+Math.floor(Math.random()*100);
}

// ============================================================
// REBUILD LEAVE BALANCES
// ============================================================
function rebuildLeaveBalances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balSheet = ss.getSheetByName(SHEETS.LEAVE_BALANCE);
  
  const correctHeaders = ['EmployeeID','EmployeeName','Year','TotalEntitled','ApprovedLeaveUsed','MedicalLeaveUsed','EmergencyLeaveUsed','UnapprovedAbsentUsed','RejectedCasualUsed','RejectedMedOrEmgUsed','CarryForward','UnusedLeaves','LeaveSurplus','LastUpdated'];
  balSheet.getRange(1, 1, 1, correctHeaders.length).setValues([correctHeaders])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  
  const lastRow = balSheet.getLastRow();
  const maxCols = balSheet.getLastColumn() || 14;
  if (lastRow > 1) {
    balSheet.getRange(2, 1, lastRow - 1, maxCols).clear();
  }
  
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  const validEmps = empRepo.getAll().map(e => new Employee(e));
  const validEmpIds = new Set(validEmps.map(e => e.id));
  const empNamesMap = {};
  const empMaxLeavesMap = {};
  const config = getConfig();
  const defaultMaxLeaves = parseInt(config.max_leaves_per_year) || 14;

  validEmps.forEach(emp => {
    empNamesMap[emp.id] = emp.name;
    empMaxLeavesMap[emp.id] = emp.getMaxAnnualLeaves(defaultMaxLeaves);
  });
  
  const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
  const allLeaves = leaveRepo.getAll().map(l => new LeaveApplication(l));

  const empLeavesMap = {};
  allLeaves.forEach(leave => {
    if (!leave.employeeId || !leave.fromDate || !validEmpIds.has(leave.employeeId)) return;
    
    let year;
    try { year = new Date(leave.fromDate).getFullYear(); } catch(e) { return; }
    if (isNaN(year)) return;
    
    const key = leave.employeeId + '_' + year;
    if (!empLeavesMap[key]) {
      empLeavesMap[key] = { cl: 0, ml: 0, el: 0, rejectedCasualPending: 0, rejectedCasualAbsent: 0, rejectedMedOrEmg: 0 };
    }
    
    const cleanStatus = leave.status.toLowerCase();
    const leaveStrategy = LeaveStrategyFactory.getStrategy(leave.leaveType);
    
    if (cleanStatus === 'approved') {
      if (leaveStrategy.countsTowardsEntitlementLimit()) {
        empLeavesMap[key].cl += leave.totalDays;
      } else {
        if (leave.leaveType === 'Medical Leave') empLeavesMap[key].ml += leave.totalDays;
        else if (leave.leaveType === 'Emergency Leave') empLeavesMap[key].el += leave.totalDays;
      }
    } else if (cleanStatus === 'rejected') {
      if (leaveStrategy.countsTowardsEntitlementLimit()) {
        if (leave.markedAbsent === 'PRESENT') {
          // Present -> scrapped
        } else if (leave.markedAbsent === 'ABSENT' || leave.markedAbsent === 'TRUE' || leave.markedAbsent === 'YES') {
          empLeavesMap[key].rejectedCasualAbsent += leave.totalDays;
        } else {
          empLeavesMap[key].rejectedCasualPending += leave.totalDays;
        }
      } else {
        empLeavesMap[key].rejectedMedOrEmg += leave.totalDays;
      }
    }
  });
  
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const allAtt = attRepo.getAll().map(a => new AttendanceRecord(a));

  const empAttMap = {};
  allAtt.forEach(att => {
    if (!att.employeeId || !att.year || !validEmpIds.has(att.employeeId)) return;
    const key = att.employeeId + '_' + att.year;
    if (!empAttMap[key]) {
      empAttMap[key] = { daysAbsent: 0, approvedLeave: 0, medicalLeave: 0, emergencyLeave: 0, rejectedMedOrEmg: 0 };
    }
    empAttMap[key].daysAbsent += att.daysAbsent;
    empAttMap[key].approvedLeave += att.approvedLeave;
    empAttMap[key].medicalLeave += att.medicalLeave;
    empAttMap[key].emergencyLeave += att.emergencyLeave;
    empAttMap[key].rejectedMedOrEmg += att.rejectedMedOrEmg;
  });
  
  const allKeys = new Set([
    ...Object.keys(empLeavesMap),
    ...Object.keys(empAttMap)
  ]);
  
  const newRows = [];
  const processedKeys = new Set();
  
  for (const key of allKeys) {
    const [empId, yearStr] = key.split('_');
    const year = parseInt(yearStr);
    if (!validEmpIds.has(empId)) continue;
    
    const l = empLeavesMap[key] || { cl: 0, ml: 0, el: 0, rejectedCasualPending: 0, rejectedCasualAbsent: 0, rejectedMedOrEmg: 0 };
    const a = empAttMap[key] || { daysAbsent: 0, approvedLeave: 0, medicalLeave: 0, emergencyLeave: 0, rejectedMedOrEmg: 0 };
    
    const approvedCasual = Math.max(l.cl, a.approvedLeave);
    const medicalLeave = Math.max(l.ml, a.medicalLeave);
    const emergencyLeave = Math.max(l.el, a.emergencyLeave);
    const rejectedMedOrEmg = Math.max(l.rejectedMedOrEmg, a.rejectedMedOrEmg);
    
    const rejectedCasualTotal = l.rejectedCasualPending + l.rejectedCasualAbsent;
    const unapprovedAbsent = Math.max(0, a.daysAbsent - rejectedCasualTotal);
    
    const totalUsed = approvedCasual + unapprovedAbsent + rejectedCasualTotal + rejectedMedOrEmg;
    const maxLeaves = empMaxLeavesMap[empId] || defaultMaxLeaves;
    const rawRemaining = maxLeaves - totalUsed;
    const unusedLeaves = Math.max(0, rawRemaining);
    const leaveSurplus = rawRemaining < 0 ? Math.abs(rawRemaining) : 0;
    
    newRows.push([
      empId,
      empNamesMap[empId] || '',
      year,
      maxLeaves,
      approvedCasual,
      medicalLeave,
      emergencyLeave,
      unapprovedAbsent,
      rejectedCasualTotal,
      rejectedMedOrEmg,
      0,
      unusedLeaves,
      leaveSurplus,
      new Date()
    ]);
    processedKeys.add(empId + '_' + year);
  }
  
  const currentYear = getTodayIST().getFullYear();
  validEmps.forEach(emp => {
    const key = emp.id + '_' + currentYear;
    if (!processedKeys.has(key)) {
      const maxLeaves = empMaxLeavesMap[emp.id] || defaultMaxLeaves;
      newRows.push([
        emp.id, emp.name, currentYear, maxLeaves, 0, 0, 0, 0, 0, 0, 0, maxLeaves, 0, new Date()
      ]);
    }
  });
  
  if (newRows.length > 0) {
    const range = balSheet.getRange(2, 1, newRows.length, 14);
    range.setNumberFormat('@');
    range.setValues(newRows);
    balSheet.getRange(2, 1, newRows.length, 1).setNumberFormat('@');
    balSheet.getRange(2, 2, newRows.length, 1).setNumberFormat('@');
    balSheet.getRange(2, 3, newRows.length, 1).setNumberFormat('0');
    balSheet.getRange(2, 4, newRows.length, 1).setNumberFormat('0');
    balSheet.getRange(2, 5, newRows.length, 9).setNumberFormat('0.##');
    balSheet.getRange(2, 14, newRows.length, 1).setNumberFormat('M/d/yyyy HH:mm');
  }
  
  return {success: true, message: `Rebuilt leave balances for ${newRows.length} records.`};
}

// ============================================================
// ANALYSIS & NOTIFICATIONS
// ============================================================
function getEmployeeAnalysisData(employeeId, year) {
  try {
    const emp = getEmployeeById(employeeId);
    if (!emp) {
      return { success: false, message: 'Employee not found.' };
    }
    
    const attendance = getAttendance(employeeId, year);
    const leaveBalance = getLeaveBalance(employeeId, year);
    
    const monthsOrder = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    attendance.sort((a, b) => monthsOrder.indexOf(a.month) - monthsOrder.indexOf(b.month));
    
    return toPlainObject({
      success: true,
      employee: emp,
      attendance: attendance,
      leaveBalance: leaveBalance
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function generateEmployeeAiSummary(employeeId, year) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return { success: false, message: 'API key not configured.' };
    }
    
    const analysisData = getEmployeeAnalysisData(employeeId, year);
    if (!analysisData.success) {
      return { success: false, message: analysisData.message };
    }
    
    const emp = analysisData.employee;
    const att = analysisData.attendance;
    const lb = analysisData.leaveBalance;
    
    if (!att || att.length === 0) {
      return { success: false, message: 'No attendance data recorded for this employee in the selected year.' };
    }
    
    let attSummary = att.map(r => {
      const pureAbs = (parseFloat(r.daysAbsent) || 0); 
      return `- ${r.month}: Working Days: ${r.workingDays}, Present: ${r.daysPresent}, Approved Leaves: ${r.approvedLeave}, Medical Leaves: ${r.medicalLeave}, Emergency Leaves: ${r.emergencyLeave}, Absent/Penalized Days: ${pureAbs}, Rejected Med/Emg (Counts under 14, no penalty): ${r.rejectedMedOrEmg}`;
    }).join('\n');
    
    const prompt = `You are a professional HR intelligence analyst and executive director at GYWS (Gopali Youth Welfare Society). 
Analyze the following yearly attendance and leave history for the year ${year} of the employee:
Name: ${emp.name}
ID: ${emp.id}
Role: ${emp.role}
Department: ${emp.department || 'N/A'}
Monthly Attendance Records:\n${attSummary}

Leave Balance Info:
- Annual Entitlement: ${lb.totalEntitled} days
- Approved Casual Leaves Used: ${lb.approvedLeaveUsed} days
- Medical Leaves Used: ${lb.medicalLeaveUsed} days
- Emergency Leaves Used: ${lb.emergencyLeaveUsed} days
- Unapproved Absent Days: ${lb.unapprovedAbsentUsed} days
- Remaining Unused Entitled Leaves: ${lb.unusedLeaves} days

Provide a comprehensive, highly professional HR performance report and attendance assessment. 
Your analysis MUST be formatted in clean markdown, containing:
1. **Executive Attendance Summary**: A concise, objective assessment of their overall attendance consistency and dedication (highlighting their attendance percentage rate, total presents/absents).
2. **Leave Pattern Analysis**: Evaluate their leave usage patterns. Check for potential trends, like clustered leaves around weekends, spikes in medical leaves, frequent unapproved absences, or balance management.
3. **Engagement and Dedication Rating**: Provide a professional grading (choose from: 'Exceptional Connection', 'Reliable Performer', 'Needs Consistency Improvement', or 'Critical Concern') with strict justification based on the numbers.
4. **Actionable Recommendations**: 2-3 specific, encouraging, and actionable constructive steps for management or the employee (e.g. alignment review, reward recommendation, scheduling advisory).

Return ONLY the markdown analysis report. Do not include introductory text like "Here is the report...". Keep the tone executive, positive, constructive, and objective.`;

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }]
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { success: false, message: 'Gemini API Connection failed: ' + responseText };
    }
    
    const parsed = JSON.parse(responseText);
    if (!parsed.candidates || parsed.candidates.length === 0 || !parsed.candidates[0].content || !parsed.candidates[0].content.parts || parsed.candidates[0].content.parts.length === 0) {
      return { success: false, message: 'Gemini returned empty payload: ' + responseText };
    }
    
    const text = parsed.candidates[0].content.parts[0].text;
    return { success: true, summary: text };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteLeaveApplicationBackend(appId, adminId) {
  try {
    const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
    const adminData = empRepo.findOne(e => String(e.employeeid) === String(adminId));
    let isSuperAdmin = false;
    if (adminData) {
      const adminEmp = new Employee(adminData.item);
      if (adminEmp.role === 'super_admin') {
        isSuperAdmin = true;
      }
    }
    if (!isSuperAdmin) {
      const superAdminsList = getSuperAdmins();
      for (let i = 0; i < superAdminsList.length; i++) {
        if (superAdminsList[i].empId === adminId) {
          isSuperAdmin = true;
          break;
        }
      }
    }
    if (!isSuperAdmin) {
      return {success: false, message: 'Unauthorized. Only Super Admin can delete leave applications.'};
    }

    const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);
    const match = leaveRepo.findOne(l => String(l.appid) === String(appId));

    if (!match) {
      return {success: false, message: 'Leave application not found.'};
    }

    const leave = new LeaveApplication(match.item);
    
    sendDeletionNotification(leave, adminId);

    leaveRepo.deleteRow(match.rowIndex);
    SpreadsheetApp.flush();

    if (leave) {
      try {
        const fromDateObj = new Date(leave.fromDate);
        if (!isNaN(fromDateObj.getTime())) {
          const monthName = fromDateObj.toLocaleString('default', { month: 'long' });
          const yearVal = fromDateObj.getFullYear();

          const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
          const attMatch = attRepo.findOne(a => String(a.employeeid) === String(leave.employeeId) && String(a.month).toUpperCase() === monthName.toUpperCase() && safeParseInt(a.year) === parseInt(yearVal));

          if (attMatch) {
            const att = new AttendanceRecord(attMatch.item);
            const prefill = getAttendancePrefillData(leave.employeeId, monthName, yearVal);
            if (prefill && prefill.success) {
              const cl = prefill.approvedLeave || 0;
              const ml = prefill.medicalLeave || 0;
              const el = prefill.emergencyLeave || 0;
              const rejectedCasual = prefill.rejectedCasualLeave || 0;
              const rejectedMedOrEmg = prefill.rejectedMedOrEmgLeave || 0;

              const totalRejected = rejectedCasual + rejectedMedOrEmg;
              let remainingAbsent = att.workingDays - att.daysPresent - cl - ml - el - totalRejected;
              if (remainingAbsent < 0) remainingAbsent = 0;
              const calculatedAbsents = rejectedCasual + remainingAbsent;

              att.approvedLeave = cl;
              att.medicalLeave = ml;
              att.emergencyLeave = el;
              att.daysAbsent = calculatedAbsents;
              att.rejectedMedOrEmg = rejectedMedOrEmg;

              attRepo.update(attMatch.rowIndex, att.toRowObject());
            }
          }
        }
      } catch (attErr) {
        Logger.log('Failed to update attendance row during delete: ' + attErr.message);
      }
    }

    rebuildLeaveBalances();
    logAudit(adminId, adminId, 'DELETE_LEAVE_APP', `Deleted leave application ${appId}`);
    return {success: true, message: `Leave application ${appId} deleted successfully.`};
  } catch (e) {
    return {success: false, message: e.message};
  }
}

function sendDeletionNotification(leave, adminId) {
  try {
    const adminData = getEmployeeById(adminId);
    const sigName = adminData ? adminData.name : 'Super Admin';
    const sigDept = adminData ? adminData.department : 'Management';
    const sigEmail = adminData ? adminData.email : 'admin@gyws.org';

    const formatDateStr = (val) => {
      if (!val) return 'N/A';
      if (val instanceof Date) {
        return Utilities.formatDate(val, 'Asia/Kolkata', 'yyyy-MM-dd');
      }
      const parsed = new Date(val);
      if (!isNaN(parsed.getTime())) {
        return Utilities.formatDate(parsed, 'Asia/Kolkata', 'yyyy-MM-dd');
      }
      return String(val);
    };

    var fromDateIST  = formatDateStr(leave.fromDate);
    var toDateIST    = formatDateStr(leave.toDate);
    var appliedOnIST = formatDateStr(leave.appliedOn);

    var deletedOn = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a') + ' IST';
    var subject   = '[Leave System] Thread for Applications & Approvals';
    
    var body = 'Hello,\n\n' +
      'Please note that the following leave application has been DELETED / DISCARDED by Super Admin (' + sigName + ') because it was mistakenly applied for.\n\n' +
      'All relevant attendance entries and leave balance counts have been automatically reverted and corrected.\n\n' +
      '---------------------------------------------\n' +
      '  DELETED APPLICATION DETAILS\n' +
      '---------------------------------------------\n' +
      '  Application ID  : ' + leave.appId + '\n' +
      '  Reference No    : ' + (leave.referenceNo || 'N/A') + '\n' +
      '  Employee Name   : ' + leave.employeeName + '\n' +
      '  Employee ID     : ' + leave.employeeId + '\n' +
      '  Leave Type      : ' + leave.leaveType + '\n' +
      '  From Date       : ' + fromDateIST + '\n' +
      '  To Date         : ' + toDateIST + '\n' +
      '  Total Days      : ' + leave.totalDays + ' day(s)\n' +
      '  Reason          : ' + leave.reason + '\n' +
      '  Applied On      : ' + appliedOnIST + '\n' +
      '---------------------------------------------\n' +
      '  DELETION DETAILS\n' +
      '---------------------------------------------\n' +
      '  Status          : Deleted (Mistakenly Applied)\n' +
      '  Deleted By      : ' + sigName + ' (' + sigDept + ')\n' +
      '  Deleted On      : ' + deletedOn + '\n' +
      '---------------------------------------------\n\n' +
      'Regards,\n' +
      sigName + '\n' +
      sigDept + '\n' +
      'Gopali Youth Welfare Society | IIT Kharagpur\n' +
      '+91 70181 01642 | ' + sigEmail;

    var htmlBody = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">' +
      '<h2 style="color: #ea4335;">Leave Application Deleted / Discarded</h2>' +
      '<p>Hello,</p>' +
      '<p>Please note that the following leave application has been <strong>DELETED / DISCARDED</strong> by Super Admin (' + sigName + ') because it was mistakenly applied for.</p>' +
      '<p>All relevant attendance entries and leave balance counts have been automatically reverted and corrected.</p>' +
      '<table style="border-collapse: collapse; width: 100%; max-width: 600px; margin-top: 20px; border: 1px solid #ddd;">' +
        '<tr style="background-color: #f8f9fa;">' +
          '<th colspan="2" style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd; color: #1a73e8;">Application Details</th>' +
        '</tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold; width: 35%;">Application ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.appId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reference No</td><td style="padding: 10px; border-bottom: 1px solid #eee; font-family: monospace; font-weight: bold;">' + (leave.referenceNo || 'N/A') + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee Name</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.employeeName + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.employeeId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Leave Type</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.leaveType + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">From Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + fromDateIST + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">To Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + toDateIST + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Total Days</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.totalDays + ' day(s)</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reason</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + leave.reason + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Applied On</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appliedOnIST + '</td></tr>' +
        '<tr style="background-color: #f8f9fa;">' +
          '<th colspan="2" style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd; border-top: 2px solid #ddd; color: #ea4335;">Deletion Details</th>' +
        '</tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Status</td><td style="padding: 10px; border-bottom: 1px solid #eee; color: #ea4335; font-weight: bold;">Deleted (Mistakenly Applied)</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Deleted By</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + adminId + ' (' + sigName + ')</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Deleted On</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + deletedOn + '</td></tr>' +
      '</table>' +
      '<p style="margin-top: 20px;">Regards,<br><strong>' + sigName + '</strong><br>' + sigDept + '<br>Gopali Youth Welfare Society | IIT Kharagpur<br>+91 70181 01642 | ' + sigEmail + '</p>' +
      '</div>';

    var options = {
      htmlBody: htmlBody,
      name: sigName + ' (Leave System)',
      replyTo: sigEmail
    };

    var props = PropertiesService.getScriptProperties();
    var threadKey = 'LEAVE_JVM_LIVE_THREAD';
    var globalThreadId = props.getProperty(threadKey);

    if (globalThreadId) {
      try {
        var thread = GmailApp.getThreadById(globalThreadId);
        if (thread) {
          var replyOptions = { ...options, cc: ADMIN_EMAILS.join(',') };
          thread.replyAll(body, replyOptions);
          return;
        }
      } catch(threadErr) {
        Logger.log('Global thread delete reply failed: ' + threadErr.message);
      }
    }

    var recipients = ADMIN_EMAILS.join(',');
    try {
      GmailApp.sendEmail(recipients, subject, body, options);
      Utilities.sleep(2000);
      var threads = GmailApp.search('subject:"' + subject.replace(/"/g, '\\"') + '" in:sent newer_than:1m', 0, 1);
      if (threads.length > 0) {
        props.setProperty(threadKey, threads[0].getId());
      }
    } catch(gmailErr) {
      Logger.log('GmailApp delete fallback failed, using MailApp: ' + gmailErr.message);
      MailApp.sendEmail({ to: recipients, subject: subject, body: body, htmlBody: options.htmlBody, name: options.name, replyTo: options.replyTo });
    }
  } catch(e) { Logger.log('Deletion notify error: ' + e.message); }
}

function sendLeaveNotification(appData, appId) {
  try {
    var appliedOn = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a') + ' IST';
    var subject   = '[Leave System] Thread for Applications & Approvals';
    
    var appliedByName = '';
    if (appData.appliedBy) {
      var applicant = getEmployeeById(appData.appliedBy);
      if (applicant && applicant.name) {
        appliedByName = applicant.name;
      }
    }
    var appliedByDisplay = appData.appliedBy + (appliedByName ? ' (' + appliedByName + ')' : '');

    var body      = 'Hello,\n\n' +
      'A new leave application has been submitted and requires your review.\n\n' +
      '---------------------------------------------\n' +
      '  APPLICATION DETAILS\n' +
      '---------------------------------------------\n' +
      '  Application ID  : ' + appId                + '\n' +
      '  Reference No    : ' + (appData.referenceNo || 'N/A') + '\n' +
      '  Employee Name   : ' + appData.employeeName + '\n' +
      '  Employee ID     : ' + appData.employeeId   + '\n' +
      '  Leave Type      : ' + appData.leaveType    + '\n' +
      '  From Date       : ' + appData.fromDate     + '\n' +
      '  To Date         : ' + appData.toDate       + '\n' +
      '  Total Days      : ' + appData.totalDays    + ' day(s)\n' +
      '  Reason          : ' + appData.reason       + '\n' +
      '  Applied By      : ' + appliedByDisplay     + '\n' +
      '  Applied On      : ' + appliedOn            + '\n' +
      '---------------------------------------------\n\n' +
      'Please log in to the Leave Management System to approve or reject.\n\n' +
      'Regards,\nLeave & Attendance System';

    var htmlBody = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">' +
      '<h2 style="color: #1a73e8;">New Leave Application</h2>' +
      '<p>Hello,</p>' +
      '<p>A new leave application has been submitted and requires your review.</p>' +
      '<table style="border-collapse: collapse; width: 100%; max-width: 600px; margin-top: 20px; border: 1px solid #ddd;">' +
        '<tr style="background-color: #f8f9fa;">' +
          '<th colspan="2" style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd; color: #1a73e8;">Application Details</th>' +
        '</tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold; width: 35%;">Application ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reference No</td><td style="padding: 10px; border-bottom: 1px solid #eee; font-family: monospace; font-weight: bold;">' + (appData.referenceNo || 'N/A') + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee Name</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.employeeName + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.employeeId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Leave Type</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.leaveType + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">From Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.fromDate + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">To Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.toDate + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Total Days</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.totalDays + ' day(s)</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reason</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appData.reason + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Applied By</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appliedByDisplay + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Applied On</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appliedOn + '</td></tr>' +
      '</table>' +
      '<p style="margin-top: 20px;">Please log in to the Leave Management System to approve or reject.</p>' +
      '<p>Regards,<br><strong>Leave & Attendance System</strong></p>' +
      '</div>';

    var options = {
      name: appData.employeeName + ' (Leave System)',
      htmlBody: htmlBody
    };
    
    var props = PropertiesService.getScriptProperties();
    var threadKey = 'LEAVE_JVM_LIVE_THREAD';
    var globalThreadId = props.getProperty(threadKey);
    
    if (globalThreadId) {
      try {
        var thread = GmailApp.getThreadById(globalThreadId);
        if (thread) {
          var replyOptions = { ...options, cc: ADMIN_EMAILS.join(',') };
          thread.replyAll(body, replyOptions);
          return; 
        }
      } catch(e) { Logger.log('Global thread reply failed: ' + e.message); }
    }

    var recipients = ADMIN_EMAILS.join(',');
    try {
      GmailApp.sendEmail(recipients, subject, body, options);
      Utilities.sleep(2000);
      var threads = GmailApp.search('subject:"' + subject.replace(/"/g, '\\"') + '" in:sent newer_than:1m', 0, 1);
      if (threads.length > 0) {
        props.setProperty(threadKey, threads[0].getId());
      }
    } catch(gmailErr) {
      Logger.log('GmailApp failed, using MailApp fallback: ' + gmailErr.message);
      MailApp.sendEmail({ to: recipients, subject: subject, body: body, htmlBody: options.htmlBody, name: options.name });
    }
  } catch(e) { Logger.log('Leave notification error: ' + e.message); }
}

function sendDecisionNotification(d) {
  try {
    const formatDateStr = (val) => {
      if (!val) return 'N/A';
      if (val instanceof Date) {
        return Utilities.formatDate(val, 'Asia/Kolkata', 'yyyy-MM-dd');
      }
      const parsed = new Date(val);
      if (!isNaN(parsed.getTime())) {
        return Utilities.formatDate(parsed, 'Asia/Kolkata', 'yyyy-MM-dd');
      }
      return String(val);
    };

    var fromDateIST  = formatDateStr(d.fromDate);
    var toDateIST    = formatDateStr(d.toDate);
    var appliedOnIST = formatDateStr(d.appliedOn);

    var decidedOn = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a') + ' IST';
    var subject   = '[Leave System] Thread for Applications & Approvals';
    var note      = d.adminNote ? ('  Note            : ' + d.adminNote + '\n') : '';
    var limitLine = '';
    var limitHtml = '';
    
    if (d.decision === 'Rejected') {
      const cleanType = d.leaveType ? String(d.leaveType).trim().toLowerCase() : '';
      if (cleanType === 'casual leave' || cleanType === 'half day') {
        limitLine =
          '\n\u26A0\uFE0F  IMPORTANT NOTICE' + '\n' +
          '---------------------------------------------' + '\n' +
          '  Since this Casual Leave application has been REJECTED,' + '\n' +
          '  the ' + d.totalDays + ' day(s) will be counted under the' + '\n' +
          '  employee\'s 14-day annual leave limit if employee remains absent.' + '\n' +
          '  Please note that it will carry a 0.5 day salary penalty' + '\n' +
          '  if the employee still remains absent, and will be' + '\n' +
          '  marked as "Absent".' + '\n' +
          '---------------------------------------------' + '\n';
          
        limitHtml = '<div style="margin-top: 20px; padding: 15px; border-left: 4px solid #f4b400; background-color: #fff8e1;">' +
          '<strong>&#9888;&#65039; IMPORTANT NOTICE:</strong> Since this Casual Leave application has been REJECTED, the ' + d.totalDays + ' day(s) will be counted under the employee\'s 14-day annual leave limit if employee remains absent. ' +
          'Please note that it will carry a <strong>0.5 day salary penalty</strong> if the employee still remains absent, and will be marked as "Absent".' +
          '</div>';
      } else if (d.countedIn14) {
        limitLine =
          '\n\u26A0\uFE0F  IMPORTANT NOTICE' + '\n' +
          '---------------------------------------------' + '\n' +
          '  Since this ' + d.leaveType + ' application has been' + '\n' +
          '  REJECTED, the ' + d.totalDays + ' day(s) will be counted' + '\n' +
          '  under the employee\'s 14-day annual leave limit' + '\n' +
          '  as per organisation policy.' + '\n' +
          '---------------------------------------------' + '\n';
          
        limitHtml = '<div style="margin-top: 20px; padding: 15px; border-left: 4px solid #f4b400; background-color: #fff8e1;">' +
          '<strong>&#9888;&#65039; IMPORTANT NOTICE:</strong> Since this ' + d.leaveType + ' application has been REJECTED, the ' + d.totalDays + ' day(s) will be counted under the employee\'s 14-day annual leave limit as per organisation policy.' +
          '</div>';
      }
    }
    
    var adminData = getEmployeeById(d.adminId);
    var sigName   = adminData ? adminData.name : (d.adminName || 'Admin');
    var sigDept   = adminData ? adminData.department : 'Management';
    var sigEmail  = adminData ? adminData.email : (d.adminEmail || 'admin@gyws.org');
    
    var body      = 'Hello,' + '\n\n' +
      'The following leave application has been ' + d.decision + ' by ' + sigName +', ' + sigDept + '.' + '\n\n' +
      '---------------------------------------------' + '\n' +
      '  APPLICATION DETAILS'                          + '\n' +
      '---------------------------------------------' + '\n' +
      '  Application ID  : ' + d.appId                + '\n' +
      '  Reference No    : ' + (d.referenceNo || 'N/A') + '\n' +
      '  Employee Name   : ' + d.employeeName          + '\n' +
      '  Employee ID     : ' + d.employeeId            + '\n' +
      '  Leave Type      : ' + d.leaveType             + '\n' +
      '  From Date       : ' + fromDateIST             + '\n' +
      '  To Date         : ' + toDateIST               + '\n' +
      '  Total Days      : ' + d.totalDays + ' day(s)' + '\n' +
      '  Reason          : ' + d.reason                + '\n' +
      '  Applied On      : ' + appliedOnIST            + '\n' +
      '---------------------------------------------' + '\n' +
      '  DECISION'                                     + '\n' +
      '---------------------------------------------' + '\n' +
      '  Status          : ' + d.decision              + '\n' +
      '  Decided By      : ' + sigName + ' (' + sigDept + ')\n' +
      '  Decided On      : ' + decidedOn               + '\n' +
      note +
      '---------------------------------------------' + '\n' +
      limitLine +
      '\nRegards,' + '\n' +
      sigName + '\n' +
      sigDept + '\n' +
      'Gopali Youth Welfare Society | IIT Kharagpur' + '\n' +
      '+91 70181 01642 | ' + sigEmail;
      
    var statusColor = d.decision === 'Approved' ? '#0f9d58' : '#d23f31';

    var htmlBody = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">' +
      '<h2 style="color: ' + statusColor + ';">Leave Application ' + d.decision + '</h2>' +
      '<p>Hello,</p>' +
      '<p>The following leave application has been <strong>' + d.decision + '</strong> by ' + sigDept + '.</p>' +
      '<table style="border-collapse: collapse; width: 100%; max-width: 600px; margin-top: 20px; border: 1px solid #ddd;">' +
        '<tr style="background-color: #f8f9fa;">' +
          '<th colspan="2" style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd; color: #1a73e8;">Application Details</th>' +
        '</tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold; width: 35%;">Application ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.appId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reference No</td><td style="padding: 10px; border-bottom: 1px solid #eee; font-family: monospace; font-weight: bold;">' + (d.referenceNo || 'N/A') + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee Name</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.employeeName + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Employee ID</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.employeeId + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Leave Type</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.leaveType + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">From Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + fromDateIST + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">To Date</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + toDateIST + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Total Days</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.totalDays + ' day(s)</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Reason</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.reason + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Applied On</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + appliedOnIST + '</td></tr>' +
        '<tr style="background-color: #f8f9fa;">' +
          '<th colspan="2" style="padding: 12px; text-align: left; border-bottom: 2px solid #ddd; border-top: 2px solid #ddd; color: ' + statusColor + ';">Decision Details</th>' +
        '</tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Status</td><td style="padding: 10px; border-bottom: 1px solid #eee; color: ' + statusColor + '; font-weight: bold;">' + d.decision + '</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Decided By</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.adminId + ' (' + (d.adminName || 'Admin') + ')</td></tr>' +
        '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Decided On</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + decidedOn + '</td></tr>' +
        (d.adminNote ? '<tr><td style="padding: 10px; border-bottom: 1px solid #eee; font-weight: bold;">Note</td><td style="padding: 10px; border-bottom: 1px solid #eee;">' + d.adminNote + '</td></tr>' : '') +
      '</table>' +
      limitHtml +
      '<p style="margin-top: 20px;">Regards,<br><strong>' + sigName + '</strong><br>' + sigDept + '<br>Gopali Youth Welfare Society | IIT Kharagpur<br>+91 70181 01642 | ' + sigEmail + '</p>' +
      '</div>';

    var options = {
      htmlBody: htmlBody
    };
    if (d.adminName)  options.name = d.adminName + ' (Leave System)';
    if (d.adminEmail) options.replyTo = d.adminEmail;

    var props = PropertiesService.getScriptProperties();
    var threadKey = 'LEAVE_JVM_LIVE_THREAD';
    var globalThreadId = props.getProperty(threadKey);

    if (globalThreadId) {
      try {
        var thread = GmailApp.getThreadById(globalThreadId);
        if (thread) {
          var replyOptions = { ...options, cc: ADMIN_EMAILS.join(',') };
          thread.replyAll(body, replyOptions);
          return; 
        }
      } catch(threadErr) {
        Logger.log('Global thread reply failed: ' + threadErr.message);
      }
    }

    var recipients = ADMIN_EMAILS.join(',');
    try {
      GmailApp.sendEmail(recipients, subject, body, options);
      Utilities.sleep(2000);
      var threads = GmailApp.search('subject:"' + subject.replace(/"/g, '\\"') + '" in:sent newer_than:1m', 0, 1);
      if (threads.length > 0) {
        props.setProperty(threadKey, threads[0].getId());
      }
    } catch(gmailErr) {
      Logger.log('GmailApp fallback failed, using MailApp: ' + gmailErr.message);
      MailApp.sendEmail({ to: recipients, subject: subject, body: body, htmlBody: options.htmlBody, name: options.name, replyTo: options.replyTo });
    }
  } catch(e) { Logger.log('Decision notify error: ' + e.message); }
}

function sendTestMail() {
  try {
    const recipients = ADMIN_EMAILS.join(',');
    const subject = '[Leave System] Thread for Applications & Approvals';
    
    const body = 'Hello,\n\n' +
      'Please ignore any previous test email threads (including those with emoji/text display issues). This is the official thread for all future leave correspondence.\n\n' +
      'This mail thread has been created for leave applications and approval through the leave portal for the employees of Jagriti Vidya Mandir from June 2026. ' +
      'Office staffs and employees are requested to follow it for knowing about approvals/rejections of leaves by the Governing Body members.\n\n' +
      'Regards,\n' +
      'Governing Body\n' +
      'Gopali Youth Welfare Society';
      
    const htmlBody = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">' +
      '  <div style="text-align: center; margin-bottom: 20px; border-bottom: 2px solid #1a73e8; padding-bottom: 10px;">' +
      '    <h2 style="color: #1a73e8; margin-bottom: 5px;">&#127979; Jagriti Vidya Mandir</h2>' +
      '    <p style="color: #666; font-size: 14px; margin-top: 0;">Leave & Attendance Management Portal</p>' +
      '  </div>' +
      '  <div style="background-color: #fff3cd; border-left: 4px solid #ffb300; padding: 12px; margin: 15px 0; border-radius: 4px; font-size: 14px; color: #664d03;">' +
      '    <strong>Note:</strong> Please ignore any previous test/draft email threads (including those with emoji/text display issues). This is the official thread for all future leave applications and approvals.' +
      '  </div>' +
      '  <div style="background-color: #e8f0fe; border-left: 4px solid #1a73e8; padding: 15px; margin: 20px 0; border-radius: 4px;">' +
      '    <p style="margin: 0; font-size: 15px; color: #1557b0; line-height: 1.6;">' +
      '      This mail thread has been created for leave applications and approval through the leave portal for the employees of <strong>Jagriti Vidya Mandir</strong> from June 2026. ' +
      '      Office staffs and employees are requested to follow it for knowing about approvals/rejections of leaves by the Governing Body members.' +
      '    </p>' +
      '  </div>' +
      '  <div style="font-size: 14px; color: #5f6368; margin-top: 25px; border-top: 1px solid #eee; padding-top: 15px;">' +
      '    <strong>Status:</strong> Leave Portal is now Live and Operational.<br>' +
      '    Please keep this thread active for all future leave correspondence.' +
      '  </div>' +
      '  <p style="margin-top: 20px;">' +
      '    Regards,<br>' +
      '    <strong>Governing Body</strong><br>' +
      '    Gopali Youth Welfare Society' +
      '  </p>' +
      '</div>';

    const options = {
      name: 'Leave System Setup',
      htmlBody: htmlBody
    };

    const props = PropertiesService.getScriptProperties();
    const threadKey = 'LEAVE_JVM_LIVE_THREAD';
    const globalThreadId = props.getProperty(threadKey);

    if (globalThreadId) {
      try {
        const thread = GmailApp.getThreadById(globalThreadId);
        if (thread) {
          const replyOptions = { ...options, cc: ADMIN_EMAILS.join(',') };
          thread.replyAll(body, replyOptions);
          return { success: true, message: 'Official launch email replied to existing thread successfully!' };
        }
      } catch(e) {
        Logger.log('Launch thread reply failed: ' + e.message);
      }
    }

    try {
      GmailApp.sendEmail(recipients, subject, body, options);
      Utilities.sleep(2000);
      const threads = GmailApp.search('subject:"' + subject.replace(/"/g, '\\"') + '" in:sent newer_than:1m', 0, 1);
      if (threads.length > 0) {
        props.setProperty(threadKey, threads[0].getId());
      }
      return { success: true, message: 'Official launch email sent and new thread established successfully!' };
    } catch(gmailErr) {
      Logger.log('GmailApp launch failed, using MailApp fallback: ' + gmailErr.message);
      MailApp.sendEmail({ to: recipients, subject: subject, body: body, htmlBody: options.htmlBody, name: options.name });
      return { success: true, message: 'Official launch email sent via MailApp fallback.' };
    }
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// SYSTEM MIGRATIONS & BACKWARDS COMPATIBILITY TOOLS
// ============================================================
function resetAllPasswords() {
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  const pass123Hash = hashPassword('pass123');
  const hr123Hash = hashPassword('hr123');
  
  empRepo.records.forEach(wrapper => {
    const role = String(wrapper.item.role).trim().toLowerCase();
    if (role === 'employee') {
      empRepo.update(wrapper.rowIndex, { passwordhash: pass123Hash, joiningdate: '2026-01-01' });
    } else if (role === 'office_staff') {
      empRepo.update(wrapper.rowIndex, { passwordhash: hr123Hash, joiningdate: '2026-01-01' });
    }
  });
}

function checkAttendanceExists(month, year, empId) {
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const match = attRepo.findOne(a => {
    if (String(a.month).toUpperCase() !== month.toUpperCase()) return false;
    if (String(a.year) !== String(year)) return false;
    if (empId && String(a.employeeid) !== String(empId)) return false;
    return true;
  });
  return match !== null;
}

function importEmployeesFromOldSheet() {
  const oldSsId = '11WXhLh1UgWYkAkopTYaq4NYNdFd6TZpp22rgjL_jFMo';
  const oldSs = SpreadsheetApp.openById(oldSsId);
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);

  const sheets = oldSs.getSheets();
  let dataSheet = null;
  for (let s = 0; s < sheets.length; s++) {
    if (sheets[s].getName().toUpperCase().includes('2026')) {
      dataSheet = sheets[s];
      break;
    }
  }
  
  if (!dataSheet) return {success: false, message: 'No 2026 sheet found.'};
  
  const data = dataSheet.getDataRange().getValues();
  
  let nameIdx = -1;
  let idIdx = -1;
  for (let searchRow = 0; searchRow < 3; searchRow++) {
    if (!data[searchRow]) continue;
    for (let c = 0; c < data[searchRow].length; c++) {
      const h = String(data[searchRow][c]).toUpperCase().replace(/\n/g, ' ').trim();
      if (h.includes('NAME')) nameIdx = c;
      if (h.includes('ID NO') || h === 'ID') idIdx = c;
    }
    if (nameIdx !== -1) break;
  }
  
  if (nameIdx === -1) return {success: false, message: 'Name column not found.'};
  
  const existingNames = new Set(empRepo.getAll().map(e => String(e.employeename).trim().toUpperCase()));

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const rawName = String(row[nameIdx]).trim();
    if (!rawName) continue; 
    
    let namePart = rawName;
    let deptPart = 'General'; 
    
    if (rawName.includes('(') && rawName.includes(')')) {
      const match = rawName.match(/(.*?)\((.*?)\)/s); 
      if (match) {
        namePart = match[1].replace(/\n/g, ' ').trim();
        deptPart = match[2].replace(/\n/g, ' ').trim();
      }
    } else if (rawName.includes('\n')) {
      const parts = rawName.split('\n');
      namePart = parts[0].trim();
      deptPart = parts[1].trim();
    }
    
    namePart = namePart.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
    deptPart = deptPart.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

    if (existingNames.has(namePart.toUpperCase())) continue; 
    
    let empIdStr = '';
    const rawId = idIdx !== -1 ? String(row[idIdx]).trim() : '';
    if (rawId && !isNaN(parseInt(rawId))) {
      let num = parseInt(rawId);
      empIdStr = 'EMP' + num.toString().padStart(3, '0');
    } else {
      empIdStr = generateId('EMP'); 
    }
    
    const emailParts = namePart.toLowerCase().split(' ');
    const email = emailParts.join('.') + '@school.org'; 
    
    const defaultPasswordHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, '123456', Utilities.Charset.UTF_8));
    
    empRepo.append({
      employeeid: empIdStr,
      employeename: namePart,
      email: email,
      department: deptPart,
      role: 'employee',
      salary: 0,
      passwordhash: defaultPasswordHash,
      joiningdate: '2026-01-01',
      isactive: true,
      phone: ''
    });
    
    existingNames.add(namePart.toUpperCase());
  }
  
  return {success: true, message: 'Employees imported successfully!'};
}

function migrateHistoricalData() {
  const oldSsId = '11WXhLh1UgWYkAkopTYaq4NYNdFd6TZpp22rgjL_jFMo';
  const oldSs = SpreadsheetApp.openById(oldSsId);
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const balRepo = new SheetRepository(SHEETS.LEAVE_BALANCE);

  const empMap = {};
  const empRolesMap = {};
  empRepo.getAll().forEach(e => {
    const emp = new Employee(e);
    if (emp.name) {
      empMap[emp.name.trim().toUpperCase()] = emp.id;
    }
    if (emp.id) {
      empRolesMap[emp.id] = emp.role;
    }
  });

  const sheets = oldSs.getSheets();
  const monthNames = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  const leaveUsage = {}; 

  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const sheetName = sheet.getName().toUpperCase();
    
    let matchedMonth = '';
    for (let m = 0; m < monthNames.length; m++) {
      if (sheetName.includes(monthNames[m]) && sheetName.includes('2026')) {
        matchedMonth = monthNames[m];
        break;
      }
    }
    if (!matchedMonth) continue;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    
    let nameIdx = -1, presentIdx = -1, absentIdx = -1, leaveIdx = -1, holidayIdx = -1, offDayIdx = -1;
    
    for (let searchRow = 0; searchRow < 3; searchRow++) {
      if (!data[searchRow]) continue;
      const headerRow = data[searchRow];
      for (let c = 0; c < headerRow.length; c++) {
        const h = String(headerRow[c]).toUpperCase().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
        if (h.includes('NAME')) nameIdx = c;
        if (h.includes('PRESENT')) presentIdx = c;
        if (h.includes('ABSENT')) absentIdx = c;
        if (h.includes('LEAVE')) leaveIdx = c;
        if (h.includes('HOLIDAY')) holidayIdx = c;
        if (h.includes('OFF')) offDayIdx = c;
      }
      if (presentIdx !== -1) break;
    }

    if (presentIdx === -1 || nameIdx === -1) continue; 

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      let rawName = String(row[nameIdx]).trim();
      if (!rawName) continue;
      
      const cleanName = rawName.split('\n')[0].replace(/\(.*?\)/g, '').trim().toUpperCase();
      let empId = empMap[cleanName];
      if (!empId) continue; 
      
      const presents = parseFloat(row[presentIdx]) || 0;
      const absents = absentIdx !== -1 ? (parseFloat(row[absentIdx]) || 0) : 0;
      
      const rawLeaveVal = leaveIdx !== -1 ? String(row[leaveIdx]).trim() : '';
      let leaves = 0;
      let emergencyLeaves = 0;
      if (rawLeaveVal) {
        const lowerVal = rawLeaveVal.toLowerCase();
        if (lowerVal.includes('special leave')) {
          const match = rawLeaveVal.match(/^([0-9]+(?:\.[0-9]+)?)/);
          if (match) {
            emergencyLeaves = parseFloat(match[1]) || 0;
          }
        } else {
          leaves = parseFloat(rawLeaveVal) || 0;
        }
      }
      
      if (presents === 0 && absents === 0 && leaves === 0 && emergencyLeaves === 0) continue;
      
      const holidays = holidayIdx !== -1 ? (parseInt(row[holidayIdx]) || 0) : 0;
      const offDays = offDayIdx !== -1 ? (parseInt(row[offDayIdx]) || 0) : 0;
      
      const monthNamesList = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
      const mIdx = monthNamesList.indexOf(matchedMonth.toUpperCase());
      const daysInMonth = new Date(2026, mIdx + 1, 0).getDate();
      
      let workingDays = daysInMonth - offDays - holidays;
      if (workingDays < 0) workingDays = 0;
      if (workingDays === 0) continue; 
      
      let exists = attRepo.findOne(a => a.employeeid === empId && String(a.month).toUpperCase() === matchedMonth && String(a.year) === '2026');
      
      if (!exists) {
        const id = generateId('ATT');
        attRepo.append({
          attid: id,
          employeeid: empId,
          month: matchedMonth.charAt(0).toUpperCase() + matchedMonth.slice(1).toLowerCase(),
          year: 2026,
          workingdays: workingDays,
          dayspresent: presents,
          approvedleave: leaves,
          medicalleave: 0,
          emergencyleave: emergencyLeaves,
          daysabsent: absents,
          rejectedmedoremg: 0
        });
      }
      
      if (!leaveUsage[empId]) leaveUsage[empId] = { cl: 0, el: 0 };
      leaveUsage[empId].cl += leaves;
      leaveUsage[empId].el += emergencyLeaves;
    }
  }

  balRepo.records.forEach(wrapper => {
    const empId = wrapper.item.employeeid;
    if (leaveUsage[empId]) {
      const bal = new LeaveBalanceRecord(wrapper.item);
      bal.approvedLeaveUsed += (leaveUsage[empId].cl || 0);
      bal.emergencyLeaveUsed += (leaveUsage[empId].el || 0);
      
      const role = empRolesMap[empId];
      const maxL = role === 'office_staff' ? 20 : 14;
      bal.totalEntitled = maxL;
      
      const totalUsed = bal.approvedLeaveUsed;
      const raw = bal.totalEntitled - totalUsed;
      bal.unusedLeaves = Math.max(0, raw);
      bal.leaveSurplus = raw < 0 ? Math.abs(raw) : 0;
      bal.lastUpdated = new Date();

      balRepo.update(wrapper.rowIndex, bal.toRowObject());
    }
  });
  
  rebuildLeaveBalances();
  return {success: true, message: 'Migration completed!'};
}

function extractMonthlyAttendance(month, year, adminId) {
  const oldSsId = '11WXhLh1UgWYkAkopTYaq4NYNdFd6TZpp22rgjL_jFMo';
  const oldSs = SpreadsheetApp.openById(oldSsId);
  const empRepo = new SheetRepository(SHEETS.EMPLOYEES);
  const attRepo = new SheetRepository(SHEETS.ATTENDANCE);
  const leaveRepo = new SheetRepository(SHEETS.LEAVE_APPS);

  const empMap = {}; 
  empRepo.getAll().forEach(e => {
    if (e.employeename) {
      empMap[String(e.employeename).trim().toUpperCase()] = e.employeeid;
    }
  });

  const targetSheetName = (month + year).toUpperCase();
  const sheets = oldSs.getSheets();
  let dataSheet = null;
  for (let s = 0; s < sheets.length; s++) {
    if (sheets[s].getName().toUpperCase().replace(/\s/g, '') === targetSheetName) {
      dataSheet = sheets[s];
      break;
    }
  }

  if (!dataSheet) return {success: false, message: `Could not find sheet for ${month} ${year} in external tracker.`};

  const data = dataSheet.getDataRange().getValues();
  
  let nameIdx = -1, presentIdx = -1, workingDaysIdx = -1, offDayIdx = -1, holidayIdx = -1;
  for (let searchRow = 0; searchRow < 3; searchRow++) {
    if (!data[searchRow]) continue;
    const headerRow = data[searchRow];
    for (let c = 0; c < headerRow.length; c++) {
      const h = String(headerRow[c]).toUpperCase().replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
      if (h.includes('NAME')) nameIdx = c;
      if (h.includes('PRESENT')) presentIdx = c;
      if (h.includes('WORKING DAYS') || h.includes('WORKINGDAYS') || h.includes('TOTAL WORKING')) workingDaysIdx = c;
      if (h.includes('OFF DAY') || h.includes('OFFDAY')) offDayIdx = c;
      if (h.includes('HOLIDAY')) holidayIdx = c;
    }
    if (presentIdx !== -1 && nameIdx !== -1) break;
  }

  if (presentIdx === -1 || nameIdx === -1) return {success: false, message: `Could not find required columns (NAME, PRESENT) in ${targetSheetName} sheet.`};

  const allLeaves = leaveRepo.getAll().map(l => new LeaveApplication(l));
  const empLeavesMap = {};
  
  allLeaves.forEach(leave => {
    if (!leave.employeeId) return;
    const fromDate = new Date(leave.fromDate);
    if (isNaN(fromDate.getTime())) return;
    
    const rowMonth = fromDate.toLocaleString('default', { month: 'long' }).toUpperCase();
    const rowYear = String(fromDate.getFullYear());
    
    if (rowMonth === month.toUpperCase() && rowYear === String(year)) {
      const days = leave.totalDays;
      if (!empLeavesMap[leave.employeeId]) {
        empLeavesMap[leave.employeeId] = { cl: 0, ml: 0, el: 0, rejectedCasual: 0, rejectedMedOrEmg: 0 };
      }
      
      if (leave.status === 'Approved') {
        if (leave.leaveType === 'Medical Leave') {
          empLeavesMap[leave.employeeId].ml += days;
        } else if (leave.leaveType === 'Emergency Leave') {
          empLeavesMap[leave.employeeId].el += days;
        } else {
          empLeavesMap[leave.employeeId].cl += days;
        }
      } else if (leave.status === 'Rejected') {
        if (leave.leaveType === 'Medical Leave' || leave.leaveType === 'Emergency Leave') {
          empLeavesMap[leave.employeeId].rejectedMedOrEmg += days;
        } else {
          if (leave.markedAbsent !== 'PRESENT') {
            empLeavesMap[leave.employeeId].rejectedCasual += days;
          }
        }
      }
    }
  });

  let recordsAdded = 0;
  let recordsUpdated = 0;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    let rawName = String(row[nameIdx]).trim();
    if (!rawName) continue;
    
    const cleanName = rawName.split('\n')[0].replace(/\(.*?\)/g, '').trim().toUpperCase();
    let empId = empMap[cleanName];
    if (!empId) continue; 
    
    const presents = parseInt(row[presentIdx]) || 0;
    const holidays = holidayIdx !== -1 ? (parseInt(row[holidayIdx]) || 0) : 0;
    const offDays = offDayIdx !== -1 ? (parseInt(row[offDayIdx]) || 0) : 0;
    
    const monthNamesList = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const mIdx = monthNamesList.findIndex(m => m.toUpperCase() === month.toUpperCase());
    const daysInMonth = new Date(year, mIdx + 1, 0).getDate();
    
    let workingDays = daysInMonth - offDays - holidays;
    if (workingDays < 0) workingDays = 0;
    if (workingDays === 0) continue; 
    
    const l = empLeavesMap[empId] || { cl: 0, ml: 0, el: 0, rejectedCasual: 0, rejectedMedOrEmg: 0 };
    
    const totalRejected = (l.rejectedCasual || 0) + (l.rejectedMedOrEmg || 0);
    let remainingAbsent = workingDays - presents - l.cl - l.ml - l.el - totalRejected;
    if (remainingAbsent < 0) remainingAbsent = 0;
    
    let calculatedAbsents = (l.rejectedCasual || 0) + remainingAbsent;

    const existsMatch = attRepo.findOne(a => a.employeeid === empId && String(a.month).toUpperCase() === month.toUpperCase() && parseInt(a.year) == year);

    const record = {
      employeeid: empId,
      month: month.charAt(0).toUpperCase() + month.slice(1).toLowerCase(),
      year: parseInt(year),
      workingdays: workingDays,
      dayspresent: presents,
      approvedleave: l.cl,
      medicalleave: l.ml,
      emergencyleave: l.el,
      daysabsent: calculatedAbsents,
      rejectedmedoremg: l.rejectedMedOrEmg || 0,
      enteredby: adminId || 'System',
      timestamp: new Date()
    };

    if (existsMatch) {
      attRepo.update(existsMatch.rowIndex, record);
      recordsUpdated++;
    } else {
      record.attid = generateId('ATT');
      attRepo.append(record);
      recordsAdded++;
    }
  }

  rebuildLeaveBalances();
  return {success: true, message: `Extracted ${month} ${year}: Added ${recordsAdded}, Updated ${recordsUpdated}`};
}

// ============================================================
// DATABASE INITIALIZATION & DUMMY DATA SETUP
// ============================================================
function setupDummyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetConfigs = [
    {
      name: SHEETS.EMPLOYEES || 'Employees',
      headers: ['EmployeeID', 'EmployeeName', 'Email', 'Department', 'Role', 'Salary', 'PasswordHash', 'JoiningDate', 'IsActive', 'Phone'],
      rows: [
        ['EMP001', 'Alice Smith', 'alice@example.com', 'Engineering', 'employee', 60000, hashPassword('pass123'), '2026-01-01', true, '1234567890'],
        ['EMP002', 'Bob Jones', 'bob@example.com', 'Human Resources', 'employee', 70000, hashPassword('pass123'), '2026-01-01', true, '0987654321'],
        ['EMP003', 'Charlie Brown', 'charlie@example.com', 'Operations', 'office_staff', 50000, hashPassword('hr123'), '2026-01-01', true, '1122334455'],
        ['EMP004', 'Dave Admin', 'dave@example.com', 'Management', 'super_admin', 90000, hashPassword('admin123'), '2026-01-01', true, '5566778899']
      ]
    },
    {
      name: SHEETS.LEAVE_APPS || 'LeaveApplications',
      headers: ['AppID', 'EmployeeID', 'EmployeeName', 'AppliedBy', 'LeaveType', 'FromDate', 'ToDate', 'TotalDays', 'Reason', 'Status', 'AppliedOn', 'ApprovedBy', 'ApprovedOn', 'AdminNote', 'IsHalfDay', 'ReferenceNo', 'MarkedAbsent'],
      rows: [
        ['APP001', 'EMP001', 'Alice Smith', 'EMP001', 'Casual Leave', '2026-06-15', '2026-06-17', 3, 'Family trip', 'Approved', '2026-06-01', 'EMP004', '2026-06-02', 'Enjoy!', 'FALSE', 'GYWS/L/101/2026', ''],
        ['APP002', 'EMP001', 'Alice Smith', 'EMP001', 'Medical Leave', '2026-06-05', '2026-06-06', 2, 'Fever', 'Approved', '2026-06-05', 'EMP004', '2026-06-05', 'Approved', 'FALSE', 'GYWS/L/102/2026', ''],
        ['APP003', 'EMP003', 'Charlie Brown', 'EMP003', 'Casual Leave', '2026-07-20', '2026-07-20', 1, 'Personal work', 'Pending', '2026-06-08', '', '', '', 'FALSE', 'GYWS/L/103/2026', '']
      ]
    },
    {
      name: SHEETS.ATTENDANCE || 'Attendance',
      headers: ['AttID', 'EmployeeID', 'Month', 'Year', 'WorkingDays', 'DaysPresent', 'ApprovedLeave', 'MedicalLeave', 'EmergencyLeave', 'DaysAbsent', 'RejectedMedOrEmg'],
      rows: [
        ['ATT001', 'EMP001', 'May', 2026, 22, 20, 2, 0, 0, 0, 0],
        ['ATT002', 'EMP003', 'May', 2026, 22, 21, 0, 1, 0, 0, 0]
      ]
    },
    {
      name: SHEETS.CONFIG || 'Config',
      headers: ['Key', 'Value'],
      rows: [
        ['max_leaves_per_year', 14],
        ['advance_days_required', 3]
      ]
    },
    {
      name: SHEETS.AUDIT_LOG || 'AuditLog',
      headers: ['Timestamp', 'UserID', 'UserName', 'Action', 'Details'],
      rows: []
    },
    {
      name: SHEETS.SALARY_CALC || 'SalaryCalculations',
      headers: ['CalcID', 'EmployeeID', 'EmployeeName', 'Month', 'Year', 'BaseSalary', 'ApprovedLeaveDays', 'AbsentDays', 'HalfDays', 'MedicalDays', 'EmergencyDays', 'ExtraDays', 'AbsentDeduction', 'ExtraLeaveDeduction', 'UnusedLeaveBonus', 'NetPayable'],
      rows: []
    },
    {
      name: SHEETS.LEAVE_BALANCE || 'LeaveBalance',
      headers: ['EmployeeID','EmployeeName','Year','TotalEntitled','ApprovedLeaveUsed','MedicalLeaveUsed','EmergencyLeaveUsed','UnapprovedAbsentUsed','RejectedCasualUsed','RejectedMedOrEmgUsed','CarryForward','UnusedLeaves','LeaveSurplus','LastUpdated'],
      rows: []
    }
  ];

  sheetConfigs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    } else {
      sheet.clear();
    }
    
    // Set headers
    sheet.getRange(1, 1, 1, cfg.headers.length)
      .setValues([cfg.headers])
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
      
    // Freeze first row
    sheet.setFrozenRows(1);
    
    // Set rows if any
    if (cfg.rows.length > 0) {
      sheet.getRange(2, 1, cfg.rows.length, cfg.headers.length).setValues(cfg.rows);
    }
  });

  SpreadsheetApp.flush();
  
  // Rebuild balances using the newly configured tables
  rebuildLeaveBalances();
  
  Logger.log("Database initialized with dummy data successfully!");
  return "Database initialized with dummy data successfully!";
}
