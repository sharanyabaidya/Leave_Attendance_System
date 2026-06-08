// ============================================================
// LEAVE & ATTENDANCE MANAGEMENT SYSTEM
// OOP Domain Models & Repositories: OOP.gs
// ============================================================

// ------------------------------------------------------------
// SAFE PARSING HELPERS
// ------------------------------------------------------------
function safeParseFloat(val, fallback = 0) {
  if (val === undefined || val === null || String(val).trim() === '') return fallback;
  const parsed = parseFloat(val);
  return isNaN(parsed) ? fallback : parsed;
}

function safeParseInt(val, fallback = 0) {
  if (val === undefined || val === null || String(val).trim() === '') return fallback;
  const parsed = parseInt(val);
  return isNaN(parsed) ? fallback : parsed;
}

function toPlainObject(obj) {
  if (obj === null || obj === undefined) return obj;
  return JSON.parse(JSON.stringify(obj));
}

// ------------------------------------------------------------
// 1. DATA REPOSITORY PATTERN
// ------------------------------------------------------------
class SheetRepository {
  constructor(sheetName) {
    this.sheetName = sheetName;
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    this.headers = [];
    this.normalizedHeaders = [];
    this.records = []; // Array of { item: normalizedObj, rowIndex: number }
    this.load();
  }

  load() {
    if (!this.sheet) {
      throw new Error(`Sheet "${this.sheetName}" not found.`);
    }
    const range = this.sheet.getDataRange();
    const values = range.getValues();
    if (values.length > 0) {
      this.headers = values[0].map(h => String(h).trim());
      this.normalizedHeaders = this.headers.map(h => this.normalizeKey(h));
      this.records = values.slice(1).map((row, index) => {
        const data = {};
        this.normalizedHeaders.forEach((normHeader, colIndex) => {
          data[normHeader] = row[colIndex];
        });
        return { item: data, rowIndex: index + 2 }; // row 1 is header, so row data starts at 2
      });
    } else {
      this.headers = [];
      this.normalizedHeaders = [];
      this.records = [];
    }
  }

  normalizeKey(key) {
    return String(key).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  getAll() {
    return this.records.map(r => r.item);
  }

  find(filterFn) {
    return this.records.filter(r => filterFn(r.item));
  }

  findOne(filterFn) {
    const record = this.records.find(r => filterFn(r.item));
    return record ? record : null;
  }

  update(rowIndex, dataObject) {
    const normalizedData = {};
    Object.keys(dataObject).forEach(k => {
      normalizedData[this.normalizeKey(k)] = dataObject[k];
    });

    this.headers.forEach((rawHeader, colIndex) => {
      const normHeader = this.normalizedHeaders[colIndex];
      if (normalizedData[normHeader] !== undefined) {
        this.sheet.getRange(rowIndex, colIndex + 1).setValue(normalizedData[normHeader]);
      }
    });

    // Update local cache
    const record = this.records.find(r => r.rowIndex === rowIndex);
    if (record) {
      Object.assign(record.item, normalizedData);
    }
  }

  append(dataObject) {
    const normalizedData = {};
    Object.keys(dataObject).forEach(k => {
      normalizedData[this.normalizeKey(k)] = dataObject[k];
    });

    const row = this.headers.map((rawHeader, colIndex) => {
      const normHeader = this.normalizedHeaders[colIndex];
      return normalizedData[normHeader] !== undefined ? normalizedData[normHeader] : '';
    });

    this.sheet.appendRow(row);
    const newRowIndex = this.sheet.getLastRow();
    this.records.push({ item: normalizedData, rowIndex: newRowIndex });
    return newRowIndex;
  }

  deleteRow(rowIndex) {
    this.sheet.deleteRow(rowIndex);
    this.records = this.records
      .filter(r => r.rowIndex !== rowIndex)
      .map(r => {
        if (r.rowIndex > rowIndex) {
          r.rowIndex -= 1;
        }
        return r;
      });
  }

  overwriteAll(objectsList) {
    const lastRow = this.sheet.getLastRow();
    const maxCols = this.sheet.getLastColumn() || this.headers.length;
    if (lastRow > 1) {
      this.sheet.getRange(2, 1, lastRow - 1, maxCols).clear();
    }
    const normalizedObjects = objectsList.map(obj => {
      const normalizedObj = {};
      Object.keys(obj).forEach(k => {
        normalizedObj[this.normalizeKey(k)] = obj[k];
      });
      return normalizedObj;
    });

    const newRows = normalizedObjects.map(normalizedObj => {
      return this.normalizedHeaders.map(normHeader => {
        return normalizedObj[normHeader] !== undefined ? normalizedObj[normHeader] : '';
      });
    });

    if (newRows.length > 0) {
      const range = this.sheet.getRange(2, 1, newRows.length, this.headers.length);
      range.setValues(newRows);
    }
    this.load();
  }
}

// ------------------------------------------------------------
// 2. DOMAIN ENTITY CLASSES
// ------------------------------------------------------------

class Employee {
  constructor(data) {
    this.id = String(data.employeeid || data.id || '');
    this.name = String(data.employeename || data.name || '');
    this.email = String(data.email || '');
    this.department = String(data.department || '');
    this.role = String(data.role || '');
    this.salary = safeParseFloat(data.salary, 0);
    this.passwordHash = String(data.passwordhash || data.password || '');
    this.joiningDate = data.joiningdate || '';
    this.isActive = data.isactive === true || data.isactive === 'TRUE' || String(data.isactive).toLowerCase() === 'true';
    this.phone = String(data.phone || '');
  }

  getMaxAnnualLeaves(configLimit = 14) {
    if (this.role === 'office_staff') {
      return 20;
    }
    return configLimit;
  }

  getDailyRate() {
    return this.salary / 30.0;
  }

  toRowObject() {
    return {
      employeeid: this.id,
      employeename: this.name,
      email: this.email,
      department: this.department,
      role: this.role,
      salary: this.salary,
      passwordhash: this.passwordHash,
      joiningdate: this.joiningDate,
      isactive: this.isActive,
      phone: this.phone
    };
  }
}

class LeaveApplication {
  constructor(data) {
    this.appId = String(data.appid || data.appId || '');
    this.employeeId = String(data.employeeid || data.employeeId || '');
    this.employeeName = String(data.employeename || data.employeeName || '');
    this.appliedBy = String(data.appliedby || data.appliedBy || '');
    this.leaveType = String(data.leavetype || data.leaveType || '');
    this.fromDate = data.fromdate || data.fromDate || '';
    this.toDate = data.todate || data.toDate || '';
    this.totalDays = safeParseFloat(data.totaldays !== undefined ? data.totaldays : data.totalDays, 0);
    this.reason = String(data.reason || '');
    this.status = String(data.status || 'Pending').trim();
    this.appliedOn = data.appliedon || data.appliedOn || '';
    this.approvedBy = String(data.approvedby || data.approvedBy || '');
    this.approvedOn = data.approvedon || data.approvedOn || '';
    this.adminNote = String(data.adminnote || data.adminNote || '');
    this.isHalfDay = String(data.ishalfday !== undefined ? data.ishalfday : (data.isHalfDay || 'FALSE'));
    this.referenceNo = String(data.referenceno || data.referenceNo || '');
    this.markedAbsent = String(data.markedabsent !== undefined ? data.markedabsent : (data.markedAbsent || '')).trim().toUpperCase();
  }

  toRowObject() {
    return {
      appid: this.appId,
      employeeid: this.employeeId,
      employeename: this.employeeName,
      appliedby: this.appliedBy,
      leavetype: this.leaveType,
      fromdate: this.fromDate,
      todate: this.toDate,
      totaldays: this.totalDays,
      reason: this.reason,
      status: this.status,
      appliedon: this.appliedOn,
      approvedby: this.approvedBy,
      approvedon: this.approvedOn,
      adminnote: this.adminNote,
      ishalfday: this.isHalfDay,
      referenceno: this.referenceNo,
      markedabsent: this.markedAbsent
    };
  }
}

class AttendanceRecord {
  constructor(data) {
    this.attId = String(data.attid || data.attId || '');
    this.employeeId = String(data.employeeid || data.employeeId || '');
    this.month = String(data.month || '');
    this.year = safeParseInt(data.year, 0);
    this.workingDays = safeParseFloat(data.workingdays !== undefined ? data.workingdays : data.workingDays, 0);
    this.daysPresent = safeParseFloat(data.dayspresent !== undefined ? data.dayspresent : data.daysPresent, 0);
    this.approvedLeave = safeParseFloat(data.approvedleave !== undefined ? data.approvedleave : data.approvedLeave, 0);
    this.medicalLeave = safeParseFloat(data.medicalleave !== undefined ? data.medicalleave : data.medicalLeave, 0);
    this.emergencyLeave = safeParseFloat(data.emergencyleave !== undefined ? data.emergencyleave : data.emergencyLeave, 0);
    this.daysAbsent = safeParseFloat(data.daysabsent !== undefined ? data.daysabsent : data.daysAbsent, 0);
    this.rejectedMedOrEmg = safeParseFloat(data.rejectedmedoremg !== undefined ? data.rejectedmedoremg : data.rejectedMedOrEmg, 0);
  }

  toRowObject() {
    return {
      attid: this.attId,
      employeeid: this.employeeId,
      month: this.month,
      year: this.year,
      workingdays: this.workingDays,
      dayspresent: this.daysPresent,
      approvedleave: this.approvedLeave,
      medicalleave: this.medicalLeave,
      emergencyleave: this.emergencyLeave,
      daysabsent: this.daysAbsent,
      rejectedmedoremg: this.rejectedMedOrEmg
    };
  }
}

class LeaveBalanceRecord {
  constructor(data) {
    this.employeeId = String(data.employeeid || data.employeeId || '');
    this.employeeName = String(data.employeename || data.employeeName || '');
    this.year = safeParseInt(data.year, 0);
    this.totalEntitled = safeParseFloat(data.totalentitled !== undefined ? data.totalentitled : data.totalEntitled, 14);
    this.approvedLeaveUsed = safeParseFloat(data.approvedleaveused !== undefined ? data.approvedleaveused : data.approvedLeaveUsed, 0);
    this.medicalLeaveUsed = safeParseFloat(data.medicalleaveused !== undefined ? data.medicalleaveused : data.medicalLeaveUsed, 0);
    this.emergencyLeaveUsed = safeParseFloat(data.emergencyleaveused !== undefined ? data.emergencyleaveused : data.emergencyLeaveUsed, 0);
    this.unapprovedAbsentUsed = safeParseFloat(data.unapprovedabsentused !== undefined ? data.unapprovedabsentused : data.unapprovedAbsentUsed, 0);
    this.rejectedCasualUsed = safeParseFloat(data.rejectedcasualused !== undefined ? data.rejectedcasualused : data.rejectedCasualUsed, 0);
    this.rejectedMedOrEmgUsed = safeParseFloat(data.rejectedmedoremgused !== undefined ? data.rejectedmedoremgused : data.rejectedMedOrEmgUsed, 0);
    this.carryForward = safeParseFloat(data.carryforward !== undefined ? data.carryforward : data.carryForward, 0);
    this.unusedLeaves = safeParseFloat(data.unusedleaves !== undefined ? data.unusedleaves : data.unusedLeaves, 14);
    this.leaveSurplus = safeParseFloat(data.leavesurplus !== undefined ? data.leavesurplus : data.leaveSurplus, 0);
    this.lastUpdated = data.lastupdated || data.lastUpdated || new Date();
  }

  toRowObject() {
    return {
      employeeid: this.employeeId,
      employeename: this.employeeName,
      year: this.year,
      totalentitled: this.totalEntitled,
      approvedleaveused: this.approvedLeaveUsed,
      medicalleaveused: this.medicalLeaveUsed,
      emergencyleaveused: this.emergencyLeaveUsed,
      unapprovedabsentused: this.unapprovedAbsentUsed,
      rejectedcasualused: this.rejectedCasualUsed,
      rejectedmedoremgused: this.rejectedMedOrEmgUsed,
      carryforward: this.carryForward,
      unusedleaves: this.unusedLeaves,
      leavesurplus: this.leaveSurplus,
      lastupdated: this.lastUpdated
    };
  }
}

// ------------------------------------------------------------
// 3. LEAVE RULES STRATEGY PATTERN
// ------------------------------------------------------------
class LeaveTypeStrategy {
  constructor(name) {
    this.name = name;
  }
  
  getRequiredAdvanceNoticeDays(configAdvanceDays = 3) {
    return configAdvanceDays;
  }

  countsTowardsEntitlementLimit() {
    return true;
  }

  getRejectedAbsentPenaltyFactor() {
    return 0.5; // default 0.5 salary deduction for rejected casual leaves
  }

  requiresBalanceCheck() {
    return true;
  }
}

class CasualLeaveStrategy extends LeaveTypeStrategy {
  getRequiredAdvanceNoticeDays() {
    return 2;
  }
}

class HalfDayStrategy extends LeaveTypeStrategy {
  getRequiredAdvanceNoticeDays() {
    return 2;
  }
}

class MedicalLeaveStrategy extends LeaveTypeStrategy {
  getRequiredAdvanceNoticeDays() {
    return 0; // Emergency/medical leaves do not require advance notice
  }
  
  countsTowardsEntitlementLimit() {
    return false; // Handled outside the standard 14-day limit
  }

  requiresBalanceCheck() {
    return false; // Not constrained by the annual casual leave limit
  }
}

class EmergencyLeaveStrategy extends LeaveTypeStrategy {
  getRequiredAdvanceNoticeDays() {
    return 0;
  }
  
  countsTowardsEntitlementLimit() {
    return false;
  }

  requiresBalanceCheck() {
    return false;
  }
}

class LeaveStrategyFactory {
  static getStrategy(leaveType) {
    const cleanType = String(leaveType).trim();
    switch (cleanType) {
      case 'Casual Leave':      return new CasualLeaveStrategy(cleanType);
      case 'Half Day':          return new HalfDayStrategy(cleanType);
      case 'Medical Leave':     return new MedicalLeaveStrategy(cleanType);
      case 'Emergency Leave':   return new EmergencyLeaveStrategy(cleanType);
      default:                  return new LeaveTypeStrategy(cleanType);
    }
  }
}

// ------------------------------------------------------------
// 4. DRY-RUN TESTING AND UTILITY RUNNER
// ------------------------------------------------------------
function runOopTests() {
  Logger.log("Starting OOP Domain & Repository Verification...");
  try {
    const empRepo = new SheetRepository(SHEETS.EMPLOYEES || 'Employees');
    const allEmps = empRepo.getAll();
    Logger.log(`Loaded ${allEmps.length} employees successfully.`);
    if (allEmps.length > 0) {
      const firstEmp = new Employee(allEmps[0]);
      Logger.log(`Test Employee Mapping: ID=${firstEmp.id}, Name=${firstEmp.name}, DailyRate=${firstEmp.getDailyRate().toFixed(2)}`);
    }

    const casualStrat = LeaveStrategyFactory.getStrategy('Casual Leave');
    const medStrat = LeaveStrategyFactory.getStrategy('Medical Leave');

    Logger.log(`Casual Leave Notice Required: ${casualStrat.getRequiredAdvanceNoticeDays()} days`);
    Logger.log(`Medical Leave Notice Required: ${medStrat.getRequiredAdvanceNoticeDays()} days`);
    Logger.log(`Casual Leave Counts: ${casualStrat.countsTowardsEntitlementLimit()}`);
    Logger.log(`Medical Leave Counts: ${medStrat.countsTowardsEntitlementLimit()}`);

    Logger.log("OOP Test Run Completed Successfully.");
    return { success: true, message: "OOP Tests completed successfully." };
  } catch (err) {
    Logger.log("OOP Test Failure: " + err.message);
    return { success: false, message: err.message };
  }
}
