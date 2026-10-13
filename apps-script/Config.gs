/**
 * ============================================================================
 * ALLO INNOWARE - DAILY OPERATIONS & TASK TRACKER
 * File: Config.gs
 * Description: Core configuration constants and shared utility functions.
 *              All other modules depend on this foundational configuration.
 * Author: Operations Automation Team / B.Tech Portfolio Project
 * ============================================================================
 */

// ============================================================================
// 1. SHEET NAME CONSTANTS
// Centralized mapping of all tab names used across the Google Spreadsheet.
// Using constants prevents hardcoded string errors and simplifies refactoring.
// ============================================================================
var SHEET = {
  TASK_MASTER: 'Task Master',
  ACTION_CENTER: 'Action Center',
  COMPLETION_LOG: 'Completion Log',
  DASHBOARD: 'Manager Dashboard',
  ALERT_LOG: 'Alert & Escalation Log',
  SETTINGS: 'Settings'
};

// ============================================================================
// 2. COLUMN INDEX CONSTANTS (1-based indices for SpreadsheetApp Range methods)
// ============================================================================

/**
 * Task Master (TM) Column Definitions
 * Represents the master catalog of recurring operational tasks.
 */
var TM = {
  TASK_ID: 1,        // Column A: Unique identifier for task (e.g. TASK-001)
  NAME: 2,           // Column B: Human-readable task name
  EMPLOYEE: 3,       // Column C: Assigned primary employee name
  EMAIL: 4,          // Column D: Employee email address for notifications
  FREQ_TYPE: 5,      // Column E: Frequency type (Daily, Weekly, Monthly)
  FREQ_DAYS: 6,      // Column F: Days of week (e.g. Mon,Wed,Fri) or day of month
  DESC: 7,           // Column G: Detailed operational instructions / SOP
  PROOF_TYPE: 8,     // Column H: Required evidence (Ticket ID, Checkbox, Text, Link)
  DUE_TIME: 9,       // Column I: Cutoff time (e.g. 18:00)
  CURRENT_DUE: 10,   // Column J: Current cycle due timestamp
  CYCLE_ID: 11,      // Column K: Active cycle identifier (e.g. CYC-TASK-001-20261013)
  ACTIVE: 12,        // Column L: Active status flag (TRUE / FALSE)
  LAST_DONE_AT: 13,  // Column M: Timestamp of most recent completion
  LAST_DONE_BY: 14,  // Column N: User who completed the last cycle
  STATUS: 15,        // Column O: Calculated status (Upcoming, Due Today, Overdue, etc.)
  NEXT_DUE: 16       // Column P: Projected next due timestamp
};

/**
 * Action Center (AC) Column Definitions
 * Represents the live operational queue where employees submit task evidence.
 */
var AC = {
  CYCLE_ID: 1,       // Column A: Unique active cycle ID
  TASK_ID: 2,        // Column B: Task ID referenced from Task Master
  NAME: 3,           // Column C: Task name
  EMPLOYEE: 4,       // Column D: Assigned employee
  DUE_DATE: 5,       // Column E: Target completion date (yyyy-MM-dd)
  DUE_TIME: 6,       // Column F: Target completion time (HH:mm)
  STATUS: 7,         // Column G: Status (Upcoming, Due Today, Overdue, Completed, etc.)
  PROOF_TYPE: 8,     // Column H: Validation requirement type
  TICKET_ID: 9,      // Column I: Freshdesk / Support ticket ID (if applicable)
  CONFIRMED: 10,     // Column J: Employee confirmation checkbox (TRUE / FALSE)
  NOTES: 11,         // Column K: Execution notes / remarks
  EVIDENCE: 12,      // Column L: Drive URL / link proof of work
  DONE_BY: 13,       // Column M: User email who submitted completion
  DONE_AT: 14,       // Column N: Submission timestamp
  MGR_REVIEW: 15     // Column O: Manager sign-off status (Approved / Rejected)
};

/**
 * Completion Log (CL) Column Definitions
 * Immutable audit trail capturing every completed task submission.
 */
var CL = {
  RECORD_ID: 1,      // Column A: Unique completion log record ID (e.g. COMP-001)
  CYCLE_ID: 2,       // Column B: Associated cycle ID
  TASK_ID: 3,        // Column C: Task ID referenced
  NAME: 4,           // Column D: Task title
  EMPLOYEE: 5,       // Column E: Originally assigned employee
  DONE_BY: 6,        // Column F: User who actually submitted the work
  TIMESTAMP: 7,      // Column G: Exact submission timestamp
  TICKET_ID: 8,      // Column H: Submitted ticket ID (if applicable)
  DAILY_CONFIRM: 9,  // Column I: Checkbox confirmation flag
  NOTES: 10,         // Column J: Submitted notes
  EVIDENCE: 11,      // Column K: Submitted link or proof
  VALIDATION: 12     // Column L: Automated validation outcome (Valid, Overdue Submission)
};

/**
 * Alert & Escalation Log (AL) Column Definitions
 * Audit trail of automated notifications, overdue warnings, and manager escalations.
 */
var AL = {
  EVENT_ID: 1,       // Column A: Unique alert event ID (e.g. EVT-001)
  CYCLE_ID: 2,       // Column B: Cycle ID that triggered the alert
  TASK_ID: 3,        // Column C: Associated task ID
  EMPLOYEE: 4,       // Column D: Employee associated with alert
  EVENT_TYPE: 5,     // Column E: Alert category (Reminder, Overdue, Escalation)
  TIMESTAMP: 6,      // Column F: Exact time alert was dispatched
  RECIPIENT: 7,      // Column G: Email address(es) alerted
  SUBJECT: 8,        // Column H: Email subject line
  STATUS: 9,         // Column I: Dispatch status (Sent, Failed)
  MGR_ACK: 10,       // Column J: Manager acknowledgement checkbox (TRUE / FALSE)
  ACK_BY: 11,        // Column K: Manager user ID who acknowledged
  ACK_AT: 12,        // Column L: Acknowledged timestamp
  ESC_STATUS: 13,    // Column M: Escalation state (Pending, Resolved)
  NOTES: 14          // Column N: Escalation resolution notes
};

// ============================================================================
// 3. STATUS CONSTANTS
// Standardized lifecycle status tags used across sheets.
// ============================================================================
var STATUS = {
  UPCOMING: 'Upcoming',
  DUE_TODAY: 'Due Today',
  OVERDUE: 'Overdue',
  COMPLETED: 'Completed',
  AWAITING_REVIEW: 'Awaiting Manager Review',
  ESCALATED: 'Escalated'
};

// ============================================================================
// 4. SETTINGS FUNCTIONS
// Handles dynamic application configuration stored in the Settings sheet.
// ============================================================================

/**
 * Reads all configuration key-value pairs from the Settings sheet into a JavaScript object.
 * Assumes settings are arranged with Keys in Column A and Values in Column B,
 * starting at Row 2 (Row 1 contains table headers).
 *
 * @return {Object.<string, *>} Object containing all settings as key-value pairs.
 */
function getAllSettings() {
  var settings = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('getAllSettings: No active spreadsheet context found.');
    return settings;
  }

  var sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) {
    Logger.log('getAllSettings: Sheet "' + SHEET.SETTINGS + '" does not exist.');
    return settings;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return settings; // Sheet is empty or contains only header row
  }

  // Batch read all rows from column A (keys) and column B (values)
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var rawKey = data[i][0];
    if (rawKey !== null && rawKey !== undefined && String(rawKey).trim() !== '') {
      var key = String(rawKey).trim();
      var val = data[i][1];
      settings[key] = val;
    }
  }

  return settings;
}

/**
 * Retrieves a single configuration setting value by key.
 *
 * @param {string} key The setting name to look up (e.g. 'Project Time Zone')
 * @param {*} [defaultValue=null] Optional fallback value if the key does not exist
 * @return {*} The setting value from the sheet, or defaultValue if not found
 */
function getSetting(key, defaultValue) {
  var fallback = defaultValue !== undefined ? defaultValue : null;
  if (!key) {
    return fallback;
  }

  var settings = getAllSettings();
  var trimmedKey = String(key).trim();

  if (settings.hasOwnProperty(trimmedKey) &&
      settings[trimmedKey] !== null &&
      settings[trimmedKey] !== undefined &&
      settings[trimmedKey] !== '') {
    return settings[trimmedKey];
  }

  return fallback;
}

/**
 * Updates an existing setting or appends a new key-value row in the Settings sheet.
 * Uses LockService to prevent race conditions during concurrent script triggers.
 *
 * @param {string} key The setting key to update or create
 * @param {*} value The value to store
 * @return {boolean} True if the update succeeded, false otherwise
 */
function updateSetting(key, value) {
  if (!key) {
    Logger.log('updateSetting: Key cannot be empty.');
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('updateSetting: No active spreadsheet context.');
    return false;
  }

  var sheet = ss.getSheetByName(SHEET.SETTINGS);
  if (!sheet) {
    // Automatically create Settings sheet if missing
    sheet = ss.insertSheet(SHEET.SETTINGS);
    sheet.getRange(1, 1, 1, 2).setValues([['Setting Key', 'Setting Value']]);
  }

  var lock = LockService.getScriptLock();
  try {
    // Acquire lock, wait up to 10 seconds for concurrent write operations to clear
    lock.waitLock(10000);
  } catch (err) {
    Logger.log('updateSetting: Failed to obtain script lock: ' + err.message);
    return false;
  }

  try {
    var trimmedKey = String(key).trim();
    var lastRow = sheet.getLastRow();
    var foundRow = -1;

    if (lastRow >= 2) {
      var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === trimmedKey) {
          foundRow = i + 2; // Offset by 2 (1-based index and header row)
          break;
        }
      }
    }

    if (foundRow !== -1) {
      sheet.getRange(foundRow, 2).setValue(value);
    } else {
      sheet.appendRow([trimmedKey, value]);
    }

    SpreadsheetApp.flush();
    return true;
  } catch (err) {
    Logger.log('updateSetting: Exception while writing to Settings sheet: ' + err.message);
    return false;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 5. DATE / TIME UTILITY FUNCTIONS
// Ensures all date-time calculations are synchronized with the company timezone.
// ============================================================================

/**
 * Retrieves the project timezone from Settings, defaulting to 'Asia/Kolkata'.
 *
 * @return {string} Valid IANA timezone string (e.g. 'Asia/Kolkata')
 */
function getTimezone() {
  var tz = getSetting('Project Time Zone');
  if (tz && typeof tz === 'string' && tz.trim() !== '') {
    return tz.trim();
  }
  return 'Asia/Kolkata';
}

/**
 * Returns current datetime Date object.
 * Internally, Apps Script Date objects represent UTC timestamps.
 * When formatting for display or storing strings, formatDateTime()
 * applies the configured project timezone.
 *
 * @return {Date} Current datetime Date object
 */
function getNow() {
  return new Date();
}

/**
 * Checks if a given date falls on an active working day according to the
 * 'Working Days' setting (defaults to 'Mon,Tue,Wed,Thu,Fri').
 *
 * Evaluates the day name in the project timezone.
 * Supported day abbreviations: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
 *
 * @param {Date} date The date to test
 * @return {boolean} True if date is a configured working day, false otherwise
 */
function isWorkingDay(date) {
  if (!date) {
    return false;
  }

  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) {
    return false;
  }

  // Determine 3-letter day name formatted in project timezone
  var dayName = Utilities.formatDate(d, getTimezone(), 'EEE');

  var rawSetting = getSetting('Working Days');
  var workingDaysStr = (rawSetting && typeof rawSetting === 'string')
    ? rawSetting
    : 'Mon,Tue,Wed,Thu,Fri';

  var workingDays = workingDaysStr.split(',').map(function(item) {
    return item.trim().toLowerCase();
  });

  return workingDays.indexOf(dayName.toLowerCase()) !== -1;
}

/**
 * Finds the next working day strictly following the provided date.
 * Increments day by day until isWorkingDay() returns true.
 * Includes a safety ceiling of 10 iterations to prevent infinite loops.
 *
 * @param {Date} [fromDate] Starting date. Defaults to getNow() if omitted.
 * @return {Date} Date object representing the next working day
 */
function getNextWorkingDay(fromDate) {
  var base = fromDate ? fromDate : getNow();
  var current = new Date(base instanceof Date ? base.getTime() : new Date(base).getTime());

  var maxIterations = 10;
  var iterations = 0;

  do {
    // Advance by 1 calendar day
    current.setDate(current.getDate() + 1);
    iterations++;

    if (isWorkingDay(current)) {
      return current;
    }
  } while (iterations < maxIterations);

  // Fallback return if no working day found within limit (e.g. empty config)
  return current;
}

/**
 * Sets the HH:MM time on a date object, zeroing seconds and milliseconds.
 * Returns a new Date object without mutating the input date.
 * Handles string time ('18:00', '9:30'), Date objects, or fractional day numbers.
 *
 * @param {Date} date Base date object
 * @param {string|Date|number} timeStr Time representation (e.g. '18:00', Date, or day fraction)
 * @return {Date|null} New Date object with specified time applied, or null if date is invalid
 */
function setTimeOnDate(date, timeStr) {
  if (!date) {
    return null;
  }

  var result = new Date(date instanceof Date ? date.getTime() : new Date(date).getTime());
  if (isNaN(result.getTime())) {
    return null;
  }

  var hours = 0;
  var minutes = 0;
  var seconds = 0;

  if (timeStr instanceof Date) {
    hours = timeStr.getHours();
    minutes = timeStr.getMinutes();
    seconds = timeStr.getSeconds();
  } else if (typeof timeStr === 'string' && timeStr.indexOf(':') !== -1) {
    var parts = timeStr.split(':');
    hours = parseInt(parts[0], 10) || 0;
    minutes = parseInt(parts[1], 10) || 0;
    if (parts.length > 2) {
      seconds = parseInt(parts[2], 10) || 0;
    }
  } else if (typeof timeStr === 'number') {
    // Fractional day number from spreadsheet time formatting (e.g. 0.75 = 18:00)
    var totalSeconds = Math.round(timeStr * 24 * 3600);
    hours = Math.floor(totalSeconds / 3600) % 24;
    minutes = Math.floor((totalSeconds % 3600) / 60);
    seconds = totalSeconds % 60;
  }

  result.setHours(hours, minutes, seconds, 0);
  return result;
}

/**
 * Formats a Date object into 'yyyy-MM-dd HH:mm:ss' string in the project timezone.
 *
 * @param {Date} date The date to format
 * @return {string} Formatted datetime string, or empty string if date is invalid
 */
function formatDateTime(date) {
  if (!date) {
    return '';
  }
  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) {
    return '';
  }
  return Utilities.formatDate(d, getTimezone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Formats a Date object into 'yyyy-MM-dd' string in the project timezone.
 *
 * @param {Date} date The date to format
 * @return {string} Formatted date string, or empty string if date is invalid
 */
function formatDateOnly(date) {
  if (!date) {
    return '';
  }
  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) {
    return '';
  }
  return Utilities.formatDate(d, getTimezone(), 'yyyy-MM-dd');
}

/**
 * Strips the time component from a date, setting hours, minutes, seconds,
 * and milliseconds to 0 (midnight). Returns a new Date object.
 *
 * @param {Date} date The date to strip time from
 * @return {Date|null} New Date object set to midnight (00:00:00.000)
 */
function dateOnly(date) {
  if (!date) {
    return null;
  }
  var d = new Date(date instanceof Date ? date.getTime() : new Date(date).getTime());
  if (isNaN(d.getTime())) {
    return null;
  }
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Calculates the difference between two dates in hours: (later - earlier).
 *
 * @param {Date} earlier The earlier reference date
 * @param {Date} later The later reference date
 * @return {number} Time difference in decimal hours (e.g. 2.5 for 2 hours 30 mins)
 */
function hoursDiff(earlier, later) {
  if (!earlier || !later) {
    return 0;
  }
  var tEarlier = (earlier instanceof Date) ? earlier.getTime() : new Date(earlier).getTime();
  var tLater = (later instanceof Date) ? later.getTime() : new Date(later).getTime();
  if (isNaN(tEarlier) || isNaN(tLater)) {
    return 0;
  }
  return (tLater - tEarlier) / (1000 * 60 * 60);
}

/**
 * Scans existing IDs in a given sheet and column, extracts the numeric portion,
 * finds the maximum existing ID number, and returns the next sequential ID
 * zero-padded to at least 3 digits.
 *
 * Example: If 'COMP-003' exists in column 1 of 'Completion Log', returns 'COMP-004'.
 * If no IDs exist, returns 'COMP-001'.
 *
 * @param {string} prefix The ID prefix (e.g. 'TASK', 'CYC', 'COMP', 'EVT')
 * @param {string|Sheet} sheet The Sheet object or sheet name string
 * @param {number} column 1-based column index to scan
 * @return {string} Newly generated sequential ID string (e.g. 'COMP-004')
 */
function generateId(prefix, sheet, column) {
  var cleanPrefix = prefix ? String(prefix).trim().toUpperCase() : 'ID';
  var targetSheet = sheet;

  if (typeof sheet === 'string') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return cleanPrefix + '-001';
    }
    targetSheet = ss.getSheetByName(sheet);
  }

  if (!targetSheet) {
    return cleanPrefix + '-001';
  }

  var lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    return cleanPrefix + '-001';
  }

  var colIndex = column || 1;
  var values = targetSheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  var maxNum = 0;
  // Regex to match prefix followed by hyphen and digits (case-insensitive)
  var pattern = new RegExp('^' + cleanPrefix + '-(\\d+)$', 'i');

  for (var i = 0; i < values.length; i++) {
    var cellVal = values[i][0];
    if (cellVal !== null && cellVal !== undefined) {
      var strVal = String(cellVal).trim();
      var match = strVal.match(pattern);
      if (match) {
        var num = parseInt(match[1], 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  var nextNum = maxNum + 1;
  var numStr = String(nextNum);
  while (numStr.length < 3) {
    numStr = '0' + numStr;
  }

  return cleanPrefix + '-' + numStr;
}
