/**
 * ============================================================================
 * Allo Innoware - Recurring Task Tracking, Pending Reports & Escalation
 * ----------------------------------------------------------------------------
 * File: Setup.gs
 * Description: Initializes the complete Google Sheet workbook structure.
 *              Creates all 6 sheets, headers, data validations, formulas,
 *              sample tasks, initial operational cycles, and formatting.
 * Dependencies: Config.gs (SHEET, TM, AC, CL, AL, STATUS, getNow, dateOnly)
 * ============================================================================
 */

/**
 * Main entry point for workbook setup.
 * Orchestrates the creation and configuration of all sheets, baseline data,
 * initial cycles, formulas, and visual styling in the prescribed sequence.
 * Safe to execute repeatedly without duplicating rows or headers.
 */
function setupWorkbook() {
  var lock = LockService.getScriptLock();
  var hasLock = false;

  try {
    // Acquire a script lock to prevent concurrent initialization runs
    hasLock = lock.tryLock(30000);
  } catch (lockError) {
    Logger.log('Warning: Could not acquire script lock for setup: ' + lockError.message);
  }

  try {
    Logger.log('Starting Allo Innoware Workbook setup...');

    // 1. Create or get each of the 6 required sheets
    var taskMasterSheet    = createOrGetSheet(SHEET.TASK_MASTER);
    var actionCenterSheet  = createOrGetSheet(SHEET.ACTION_CENTER);
    var completionLogSheet = createOrGetSheet(SHEET.COMPLETION_LOG);
    var dashboardSheet     = createOrGetSheet(SHEET.DASHBOARD);
    var alertLogSheet      = createOrGetSheet(SHEET.ALERT_LOG);
    var settingsSheet      = createOrGetSheet(SHEET.SETTINGS);

    // Clean up empty default "Sheet1" created automatically in fresh Google Sheets
    cleanupDefaultSheet();

    // 2. Setup Task Master configuration sheet
    setupTaskMaster(taskMasterSheet);

    // 3. Setup Action Center employee submission interface
    setupActionCenter(actionCenterSheet);

    // 4. Setup Completion Log audit trail sheet
    setupCompletionLog(completionLogSheet);

    // 5. Setup Manager Dashboard with formula metrics and live QUERY
    setupDashboard(dashboardSheet);

    // 6. Setup Alert & Escalation Log
    setupAlertLog(alertLogSheet);

    // 7. Setup Settings key-value configuration sheet
    setupSettings(settingsSheet);

    // 8. Insert baseline recurring sample tasks into Task Master
    addSampleTasks();

    // 9. Generate initial operational cycles in Action Center
    createInitialCycles();

    // 10. Apply conditional formatting rules and set column widths
    applyFormatting();

    Logger.log('Workbook setup completed successfully.');

    // 11. Present confirmation dialog to user (or log if run in headless mode)
    try {
      SpreadsheetApp.getUi().alert('Workbook setup complete!');
    } catch (uiError) {
      Logger.log('Workbook setup complete! (UI alert suppressed in non-interactive environment)');
    }

  } catch (error) {
    Logger.log('Fatal error during setupWorkbook: ' + error.message);
    throw error;
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Retrieves an existing sheet by name or creates a new one at the end of the workbook.
 * 
 * @param {string} name - The exact name of the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The retrieved or newly inserted sheet
 */
function createOrGetSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    // Insert new sheet at the end of the tab list
    sheet = ss.insertSheet(name, ss.getNumSheets());
    Logger.log('Created sheet: ' + name);
  } else {
    Logger.log('Found existing sheet: ' + name);
  }

  return sheet;
}

/**
 * Removes the default empty "Sheet1" created in brand-new Google Spreadsheets
 * if other configured sheets now exist.
 */
function cleanupDefaultSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Sheet1');

  if (sheet1 && ss.getNumSheets() > 1) {
    if (sheet1.getLastRow() === 0 && sheet1.getLastColumn() === 0) {
      try {
        ss.deleteSheet(sheet1);
        Logger.log('Removed empty default Sheet1');
      } catch (e) {
        Logger.log('Could not remove Sheet1: ' + e.message);
      }
    }
  }
}

/**
 * Configures the Task Master sheet headers, data validation rules, and freeze rows.
 * Task Master is the central registry for all recurring operational tasks.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Task Master sheet
 */
function setupTaskMaster(sheet) {
  var headers = [
    'Task ID',               // Col 1 (TM.TASK_ID)
    'Task Name',             // Col 2 (TM.NAME)
    'Assigned Employee',     // Col 3 (TM.EMPLOYEE)
    'Employee Email',        // Col 4 (TM.EMAIL)
    'Frequency Type',        // Col 5 (TM.FREQ_TYPE)
    'Frequency Days',        // Col 6 (TM.FREQ_DAYS)
    'Task Description',      // Col 7 (TM.DESC)
    'Required Proof Type',   // Col 8 (TM.PROOF_TYPE)
    'Due Time',              // Col 9 (TM.DUE_TIME)
    'Current Due Date',      // Col 10 (TM.CURRENT_DUE)
    'Current Cycle ID',      // Col 11 (TM.CYCLE_ID)
    'Task Active',           // Col 12 (TM.ACTIVE)
    'Last Completed At',     // Col 13 (TM.LAST_DONE_AT)
    'Last Completed By',     // Col 14 (TM.LAST_DONE_BY)
    'Current Status',        // Col 15 (TM.STATUS)
    'Next Due Date'          // Col 16 (TM.NEXT_DUE)
  ];

  // Set header row values and styling
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  applySheetHeaderStyle(sheet, headers.length);
  sheet.setFrozenRows(1);
}

/**
 * Configures the Action Center sheet headers, column freezing, and date formats.
 * Action Center is the employee-facing workspace where pending cycles are completed.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Action Center sheet
 */
function setupActionCenter(sheet) {
  var headers = [
    'Cycle ID',              // Col 1 (AC.CYCLE_ID)
    'Task ID',               // Col 2 (AC.TASK_ID)
    'Task Name',             // Col 3 (AC.NAME)
    'Assigned Employee',     // Col 4 (AC.EMPLOYEE)
    'Due Date',              // Col 5 (AC.DUE_DATE)
    'Due Time',              // Col 6 (AC.DUE_TIME)
    'Current Status',        // Col 7 (AC.STATUS)
    'Required Proof',        // Col 8 (AC.PROOF_TYPE)
    'Ticket ID',             // Col 9 (AC.TICKET_ID)
    'Completion Confirmed',  // Col 10 (AC.CONFIRMED)
    'Completion Notes',      // Col 11 (AC.NOTES)
    'Evidence Link',         // Col 12 (AC.EVIDENCE)
    'Completed By',          // Col 13 (AC.DONE_BY)
    'Completed At',          // Col 14 (AC.DONE_AT)
    'Manager Review Status'  // Col 15 (AC.MGR_REVIEW)
  ];

  // Set header row values and styling
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  applySheetHeaderStyle(sheet, headers.length);
  sheet.setFrozenRows(1);

  var maxRows = Math.max(sheet.getMaxRows() - 1, 50);

  // Set number formats for date and datetime columns
  sheet.getRange(2, AC.DUE_DATE, maxRows, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, AC.DONE_AT, maxRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Apply conditional formatting rules for row status colors
  applyActionCenterConditionalFormatting(sheet);
}

/**
 * Configures the Completion Log sheet headers, column freezing, and timestamp formats.
 * Serves as an immutable audit trail for every completed task cycle.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Completion Log sheet
 */
function setupCompletionLog(sheet) {
  var headers = [
    'Record ID',             // Col 1 (CL.RECORD_ID)
    'Cycle ID',              // Col 2 (CL.CYCLE_ID)
    'Task ID',               // Col 3 (CL.TASK_ID)
    'Task Name',             // Col 4 (CL.NAME)
    'Assigned Employee',     // Col 5 (CL.EMPLOYEE)
    'Completed By',          // Col 6 (CL.DONE_BY)
    'Completion Timestamp',  // Col 7 (CL.TIMESTAMP)
    'Ticket ID',             // Col 8 (CL.TICKET_ID)
    'Daily Confirmation',    // Col 9 (CL.DAILY_CONFIRM)
    'Completion Notes',      // Col 10 (CL.NOTES)
    'Evidence Link',         // Col 11 (CL.EVIDENCE)
    'Validation Result'      // Col 12 (CL.VALIDATION)
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  applySheetHeaderStyle(sheet, headers.length);
  sheet.setFrozenRows(1);

  var maxRows = Math.max(sheet.getMaxRows() - 1, 50);
  sheet.getRange(2, CL.TIMESTAMP, maxRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/**
 * Creates the Manager Dashboard with formula-based metrics and dynamic detail view.
 * Displays high-level KPIs and live QUERY of active task status.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Dashboard sheet
 */
function setupDashboard(sheet) {
  // Clear any previous formatting and values to ensure clean layout
  sheet.clear();

  // Row 1: Dashboard Title banner (Merged A1:F1)
  var titleRange = sheet.getRange('A1:F1');
  titleRange.merge();
  titleRange.setValue('MANAGER DASHBOARD');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(16);
  titleRange.setFontColor('#FFFFFF');
  titleRange.setBackground('#1F4E79');
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 42);

  // Row 2: Empty spacer row
  sheet.setRowHeight(2, 14);

  // Rows 3-10: Executive Summary KPI Cards (Labels in Col A, Live Formulas in Col B)
  var summaryRows = [
    ['Total Active Tasks',      "=COUNTIF('Task Master'!L2:L,TRUE)"],
    ['Tasks Due Today',         '=COUNTIFS(\'Task Master\'!O2:O,"Due Today",\'Task Master\'!L2:L,TRUE)'],
    ['Pending / Overdue',       '=COUNTIFS(\'Task Master\'!O2:O,"Overdue",\'Task Master\'!L2:L,TRUE)'],
    ['Completed (Current)',     '=COUNTIFS(\'Task Master\'!O2:O,"Completed",\'Task Master\'!L2:L,TRUE)'],
    ['Awaiting Manager Review', '=COUNTIFS(\'Task Master\'!O2:O,"Awaiting Manager Review",\'Task Master\'!L2:L,TRUE)'],
    ['Escalated Tasks',         '=COUNTIFS(\'Task Master\'!O2:O,"Escalated",\'Task Master\'!L2:L,TRUE)'],
    ['Unacknowledged Alerts',   '=COUNTIFS(\'Alert & Escalation Log\'!E2:E,"Manager Alert",\'Alert & Escalation Log\'!J2:J,FALSE)'],
    ['Last Automation Run',    "='Settings'!B14"]
  ];

  for (var i = 0; i < summaryRows.length; i++) {
    var rowIdx = i + 3;
    sheet.getRange(rowIdx, 1).setValue(summaryRows[i][0]);
    sheet.getRange(rowIdx, 2).setFormula(summaryRows[i][1]);

    sheet.setRowHeight(rowIdx, 26);
    var labelCell = sheet.getRange(rowIdx, 1);
    var valCell   = sheet.getRange(rowIdx, 2);

    labelCell.setFontWeight('bold')
      .setFontSize(10)
      .setFontColor('#202124')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');

    valCell.setFontWeight('bold')
      .setFontSize(11)
      .setFontColor('#1F4E79')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // Alternating zebra striping
    var bg = (i % 2 === 0) ? '#F8F9FA' : '#FFFFFF';
    sheet.getRange(rowIdx, 1, 1, 2).setBackground(bg);
  }

  // Format borders around the KPI summary table (A3:B10)
  sheet.getRange('A3:B10').setBorder(true, true, true, true, true, true, '#D0D5DD', SpreadsheetApp.BorderStyle.SOLID);

  // Row 11 & 12: Empty spacer rows
  sheet.setRowHeight(11, 12);
  sheet.setRowHeight(12, 12);

  // Row 13: Section Header for Detail Table (Merged A13:F13)
  var sectionRange = sheet.getRange('A13:F13');
  sectionRange.merge();
  sectionRange.setValue('TASK STATUS DETAIL');
  sectionRange.setFontWeight('bold');
  sectionRange.setFontSize(12);
  sectionRange.setFontColor('#FFFFFF');
  sectionRange.setBackground('#2C3E50');
  sectionRange.setHorizontalAlignment('center');
  sectionRange.setVerticalAlignment('middle');
  sheet.setRowHeight(13, 32);

  // Row 14: Dynamic QUERY formula pulling live columns from Task Master
  sheet.getRange('A14').setFormula('=IFERROR(QUERY(\'Task Master\'!A1:P,"SELECT A, B, C, J, I, O WHERE L = TRUE",1),"No active tasks")');

  // Format header styling for QUERY output row (Row 14)
  sheet.getRange('A14:F14').setFontWeight('bold').setBackground('#EAEDED').setHorizontalAlignment('left');

  // Format date output column (Col D in detail view)
  sheet.getRange('D15:D100').setNumberFormat('yyyy-mm-dd');
}

/**
 * Configures the Alert & Escalation Log headers, column freezing, and timestamp formats.
 * Tracks all notification events, delivery status, and manager acknowledgements.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Alert Log sheet
 */
function setupAlertLog(sheet) {
  var headers = [
    'Event ID',              // Col 1 (AL.EVENT_ID)
    'Cycle ID',              // Col 2 (AL.CYCLE_ID)
    'Task ID',               // Col 3 (AL.TASK_ID)
    'Employee',              // Col 4 (AL.EMPLOYEE)
    'Event Type',            // Col 5 (AL.EVENT_TYPE)
    'Event Timestamp',       // Col 6 (AL.TIMESTAMP)
    'Recipient',             // Col 7 (AL.RECIPIENT)
    'Email Subject',         // Col 8 (AL.SUBJECT)
    'Alert Status',          // Col 9 (AL.STATUS)
    'Manager Acknowledged',  // Col 10 (AL.MGR_ACK)
    'Acknowledged By',       // Col 11 (AL.ACK_BY)
    'Acknowledged At',       // Col 12 (AL.ACK_AT)
    'Escalation Status',     // Col 13 (AL.ESC_STATUS)
    'Notes'                  // Col 14 (AL.NOTES)
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  applySheetHeaderStyle(sheet, headers.length);
  sheet.setFrozenRows(1);

  var maxRows = Math.max(sheet.getMaxRows() - 1, 50);
  sheet.getRange(2, AL.TIMESTAMP, maxRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, AL.ACK_AT, maxRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/**
 * Configures the Settings key-value sheet with defaults, freezes row 1,
 * and sets warning protection on Column A.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Settings sheet
 */
function setupSettings(sheet) {
  // Row 1 headers
  sheet.getRange(1, 1, 1, 2).setValues([['Setting Name', 'Setting Value']]);
  applySheetHeaderStyle(sheet, 2);
  sheet.setFrozenRows(1);

  // Baseline default settings configuration
  var defaultSettings = [
    ['Manager Email', 'manager@example.com'],
    ['Leadership Email', 'leadership@example.com'],
    ['Project Time Zone', 'Asia/Kolkata'],
    ['Working Days', 'Mon,Tue,Wed,Thu,Fri'],
    ['Daily Task Deadline', '17:00'],
    ['Manager Alert Delay (minutes)', 15],
    ['Manager Ack Deadline (hours)', 18],
    ['Employee Resolution Window (hours)', 24],
    ['Test Mode', true],
    ['Email Sending Enabled', false],
    ['Test Email Recipient', 'test@example.com'],
    ['Spreadsheet URL', ''],
    ['Last Automation Run', '']
  ];

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    sheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  } else {
    // If settings already exist, ensure any missing default setting key is appended
    var existingKeys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var keyMap = {};
    for (var k = 0; k < existingKeys.length; k++) {
      if (existingKeys[k][0]) {
        keyMap[String(existingKeys[k][0]).trim()] = true;
      }
    }
    for (var d = 0; d < defaultSettings.length; d++) {
      if (!keyMap[defaultSettings[d][0]]) {
        sheet.appendRow(defaultSettings[d]);
      }
    }
  }

  // Format setting names column as bold labels
  var currentLastRow = Math.max(sheet.getLastRow(), defaultSettings.length + 1);
  sheet.getRange(2, 1, currentLastRow - 1, 1).setFontWeight('bold');

  // Protect Column A (Setting Names) to prevent accidental edits
  try {
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var protectedAlready = false;
    for (var p = 0; p < protections.length; p++) {
      if (protections[p].getDescription() === 'Protected Setting Names') {
        protectedAlready = true;
        break;
      }
    }
    if (!protectedAlready) {
      var protection = sheet.getRange(2, 1, currentLastRow - 1, 1).protect();
      protection.setDescription('Protected Setting Names');
      protection.setWarningOnly(true);
    }
  } catch (protErr) {
    Logger.log('Note on Settings protection: ' + protErr.message);
  }
}

/**
 * Inserts the 4 standard recurring sample tasks into Task Master if they do not already exist.
 * Idempotent check ensures re-running setup does not duplicate rows.
 */
function addSampleTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!sheet) {
    sheet = createOrGetSheet(SHEET.TASK_MASTER);
    setupTaskMaster(sheet);
  }

  // Check if row 2 already contains TASK-001
  var firstTaskId = String(sheet.getRange(2, TM.TASK_ID).getValue() || '').trim();
  if (firstTaskId === 'TASK-001') {
    Logger.log('Sample tasks already exist at row 2 of Task Master. Skipping insert.');
    return;
  }

  // Clear any old, misplaced or blank rows from row 2 downwards
  if (sheet.getMaxRows() >= 2) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 16).clearContent().clearDataValidations();
  }

  // Define the 4 standard recurring operational tasks
  var sampleTasks = [
    {
      id: 'TASK-001',
      name: 'Daily Operations Update',
      employee: 'Amit Sharma',
      email: 'amit@example.com',
      freqType: 'Daily',
      freqDays: 1,
      desc: 'Post daily operations update in the team WhatsApp group',
      proofType: 'Daily Confirmation',
      dueTime: '17:00'
    },
    {
      id: 'TASK-002',
      name: 'Portal Task A - Weekly Compliance',
      employee: 'Priya Patel',
      email: 'priya@example.com',
      freqType: '7-Day',
      freqDays: 7,
      desc: 'Complete weekly compliance check on the portal',
      proofType: 'Ticket ID',
      dueTime: '17:00'
    },
    {
      id: 'TASK-003',
      name: 'Portal Task B - Bi-Weekly Audit',
      employee: 'Rahul Verma',
      email: 'rahul@example.com',
      freqType: '15-Day',
      freqDays: 15,
      desc: 'Submit bi-weekly audit report on the portal',
      proofType: 'Ticket ID',
      dueTime: '17:00'
    },
    {
      id: 'TASK-004',
      name: 'Portal Task C - Monthly Review',
      employee: 'Sneha Gupta',
      email: 'sneha@example.com',
      freqType: '30-Day',
      freqDays: 30,
      desc: 'Complete monthly review process on the portal',
      proofType: 'Ticket ID',
      dueTime: '17:00'
    }
  ];

  var rowsToAdd = [];
  for (var j = 0; j < sampleTasks.length; j++) {
    var t = sampleTasks[j];
    rowsToAdd.push([
      t.id,                 // Col 1: Task ID
      t.name,               // Col 2: Task Name
      t.employee,           // Col 3: Assigned Employee
      t.email,              // Col 4: Employee Email
      t.freqType,           // Col 5: Frequency Type
      t.freqDays,           // Col 6: Frequency Days
      t.desc,               // Col 7: Task Description
      t.proofType,          // Col 8: Required Proof Type
      t.dueTime,            // Col 9: Due Time
      '',                   // Col 10: Current Due Date
      '',                   // Col 11: Current Cycle ID
      true,                 // Col 12: Task Active (Checkbox)
      '',                   // Col 13: Last Completed At
      '',                   // Col 14: Last Completed By
      STATUS.UPCOMING,      // Col 15: Current Status
      ''                    // Col 16: Next Due Date
    ]);
  }

  // Insert cleanly at Row 2
  sheet.getRange(2, 1, rowsToAdd.length, 16).setValues(rowsToAdd);

  // Apply validations and formatting directly to the inserted rows (2 to 5)
  var freqRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Daily', '7-Day', '15-Day', '30-Day'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TM.FREQ_TYPE, rowsToAdd.length, 1).setDataValidation(freqRule);

  var proofRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ticket ID', 'Daily Confirmation'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TM.PROOF_TYPE, rowsToAdd.length, 1).setDataValidation(proofRule);

  var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, TM.ACTIVE, rowsToAdd.length, 1).setDataValidation(cbRule);

  sheet.getRange(2, TM.CURRENT_DUE, rowsToAdd.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, TM.LAST_DONE_AT, rowsToAdd.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, TM.NEXT_DUE, rowsToAdd.length, 1).setNumberFormat('yyyy-mm-dd');

  Logger.log('Cleanly inserted 4 sample tasks into Task Master starting at row 2.');
}

/**
 * Creates initial operational cycles in Action Center for all active tasks in Task Master.
 * Uses today's date as the initial due date so tasks are immediately testable.
 * Updates Task Master with Current Due Date, Current Cycle ID, and Status='Due Today'.
 * Prevents duplicate cycle creation if cycles already exist in Action Center.
 */
function createInitialCycles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  if (!tmSheet || !acSheet) {
    Logger.log('Error: Task Master or Action Center sheet missing for createInitialCycles.');
    return;
  }

  var tmLastRow = tmSheet.getLastRow();
  if (tmLastRow < 2) {
    Logger.log('No tasks found in Task Master to create cycles for.');
    return;
  }

  // Read all existing Cycle IDs from Action Center to prevent duplicate entries
  var existingCycles = {};
  var acLastRow = acSheet.getLastRow();
  if (acLastRow >= 2) {
    var acCycleVals = acSheet.getRange(2, AC.CYCLE_ID, acLastRow - 1, 1).getValues();
    for (var c = 0; c < acCycleVals.length; c++) {
      var cId = acCycleVals[c][0];
      if (cId) {
        existingCycles[String(cId).trim()] = true;
      }
    }
  }

  var now = (typeof getNow === 'function') ? getNow() : new Date();
  var today = (typeof dateOnly === 'function') ? dateOnly(now) : new Date(now.getFullYear(), now.getMonth(), now.getDate());

  var numTasks = tmLastRow - 1;
  var tmValues = tmSheet.getRange(2, 1, numTasks, 16).getValues();
  var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  for (var i = 0; i < tmValues.length; i++) {
    var row = tmValues[i];
    var rowIndex = i + 2; // 1-based sheet row

    var taskId    = row[TM.TASK_ID - 1];
    var taskName  = row[TM.NAME - 1];
    var employee  = row[TM.EMPLOYEE - 1];
    var proofType = row[TM.PROOF_TYPE - 1];
    var dueTime   = row[TM.DUE_TIME - 1] || '17:00';
    var isActive  = row[TM.ACTIVE - 1];

    if (isActive === true && taskId) {
      var cycleId = String(taskId).trim() + '-C001';

      // Always ensure Task Master is updated with active cycle info
      tmSheet.getRange(rowIndex, TM.CURRENT_DUE).setValue(today).setNumberFormat('yyyy-mm-dd');
      tmSheet.getRange(rowIndex, TM.CYCLE_ID).setValue(cycleId);
      tmSheet.getRange(rowIndex, TM.STATUS).setValue(STATUS.DUE_TODAY);

      if (!existingCycles[cycleId]) {
        var acRow = [
          cycleId,           // Col 1 (AC.CYCLE_ID)
          taskId,            // Col 2 (AC.TASK_ID)
          taskName,          // Col 3 (AC.NAME)
          employee,          // Col 4 (AC.EMPLOYEE)
          today,             // Col 5 (AC.DUE_DATE)
          dueTime,           // Col 6 (AC.DUE_TIME)
          STATUS.DUE_TODAY,  // Col 7 (AC.STATUS)
          proofType,         // Col 8 (AC.PROOF_TYPE)
          '',                // Col 9 (AC.TICKET_ID - for employee input)
          false,             // Col 10 (AC.CONFIRMED - checkbox)
          '',                // Col 11 (AC.NOTES - for employee input)
          '',                // Col 12 (AC.EVIDENCE - for employee input)
          '',                // Col 13 (AC.DONE_BY)
          '',                // Col 14 (AC.DONE_AT)
          ''                 // Col 15 (AC.MGR_REVIEW)
        ];

        acSheet.appendRow(acRow);
        var newAcRow = acSheet.getLastRow();

        acSheet.getRange(newAcRow, AC.DUE_DATE).setNumberFormat('yyyy-mm-dd');
        acSheet.getRange(newAcRow, AC.CONFIRMED).setDataValidation(checkboxRule).setValue(false);

        existingCycles[cycleId] = true;
        Logger.log('Created initial cycle ' + cycleId + ' for task ' + taskId);
      }
    }
  }
}

/**
 * Applies conditional formatting rules to Action Center and Task Master,
 * and sets readable column widths across all sheets in the workbook.
 */
function applyFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
  var dbSheet = ss.getSheetByName(SHEET.DASHBOARD);
  var alSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  var stSheet = ss.getSheetByName(SHEET.SETTINGS);

  // 1. Action Center Conditional Formatting & Column Widths
  if (acSheet) {
    applyActionCenterConditionalFormatting(acSheet);
    setColumnWidths(acSheet, [
      140, // Col 1: Cycle ID
      100, // Col 2: Task ID
      230, // Col 3: Task Name
      140, // Col 4: Assigned Employee
      110, // Col 5: Due Date
      90,  // Col 6: Due Time
      140, // Col 7: Current Status
      150, // Col 8: Required Proof
      130, // Col 9: Ticket ID
      120, // Col 10: Completion Confirmed
      240, // Col 11: Completion Notes
      190, // Col 12: Evidence Link
      140, // Col 13: Completed By
      160, // Col 14: Completed At
      160  // Col 15: Manager Review Status
    ]);
  }

  // 2. Task Master Conditional Formatting (Column O - Current Status) & Column Widths
  if (tmSheet) {
    var maxRows = Math.max(tmSheet.getMaxRows() - 1, 100);
    var statusRange = tmSheet.getRange(2, TM.STATUS, maxRows, 1);

    var ruleOverdue = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Overdue')
      .setBackground('#FCE8E6')
      .setFontColor('#C5221F')
      .setRanges([statusRange])
      .build();

    var ruleEscalated = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Escalated')
      .setBackground('#FAD2CF')
      .setFontColor('#A50E0E')
      .setRanges([statusRange])
      .build();

    var ruleDueToday = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Due Today')
      .setBackground('#FEF7E0')
      .setFontColor('#B06000')
      .setRanges([statusRange])
      .build();

    var ruleUpcoming = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Upcoming')
      .setBackground('#E8F0FE')
      .setFontColor('#1A73E8')
      .setRanges([statusRange])
      .build();

    var ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Completed')
      .setBackground('#E6F4EA')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build();

    var ruleAwaitingReview = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Awaiting Manager Review')
      .setBackground('#F3E8FD')
      .setFontColor('#7627BB')
      .setRanges([statusRange])
      .build();

    tmSheet.setConditionalFormatRules([
      ruleOverdue,
      ruleEscalated,
      ruleDueToday,
      ruleUpcoming,
      ruleCompleted,
      ruleAwaitingReview
    ]);

    setColumnWidths(tmSheet, [
      100, // Col 1: Task ID
      230, // Col 2: Task Name
      140, // Col 3: Assigned Employee
      190, // Col 4: Employee Email
      110, // Col 5: Frequency Type
      110, // Col 6: Frequency Days
      300, // Col 7: Task Description
      160, // Col 8: Required Proof Type
      90,  // Col 9: Due Time
      120, // Col 10: Current Due Date
      140, // Col 11: Current Cycle ID
      90,  // Col 12: Task Active
      160, // Col 13: Last Completed At
      140, // Col 14: Last Completed By
      140, // Col 15: Current Status
      120  // Col 16: Next Due Date
    ]);
  }

  // 3. Completion Log Column Widths
  if (clSheet) {
    setColumnWidths(clSheet, [
      110, // Col 1: Record ID
      140, // Col 2: Cycle ID
      100, // Col 3: Task ID
      230, // Col 4: Task Name
      140, // Col 5: Assigned Employee
      140, // Col 6: Completed By
      160, // Col 7: Completion Timestamp
      130, // Col 8: Ticket ID
      130, // Col 9: Daily Confirmation
      240, // Col 10: Completion Notes
      190, // Col 11: Evidence Link
      140  // Col 12: Validation Result
    ]);
  }

  // 4. Manager Dashboard Column Widths
  if (dbSheet) {
    setColumnWidths(dbSheet, [
      240, // Col A: KPI Metric Label / Task ID
      200, // Col B: Value / Task Name
      160, // Col C: Assigned Employee
      130, // Col D: Current Due Date
      110, // Col E: Due Time
      140  // Col F: Current Status
    ]);
  }

  // 5. Alert & Escalation Log Column Widths
  if (alSheet) {
    setColumnWidths(alSheet, [
      110, // Col 1: Event ID
      140, // Col 2: Cycle ID
      100, // Col 3: Task ID
      140, // Col 4: Employee
      160, // Col 5: Event Type
      160, // Col 6: Event Timestamp
      190, // Col 7: Recipient
      260, // Col 8: Email Subject
      130, // Col 9: Alert Status
      140, // Col 10: Manager Acknowledged
      140, // Col 11: Acknowledged By
      160, // Col 12: Acknowledged At
      130, // Col 13: Escalation Status
      240  // Col 14: Notes
    ]);
  }

  // 6. Settings Column Widths
  if (stSheet) {
    setColumnWidths(stSheet, [
      260, // Col 1: Setting Name
      240  // Col 2: Setting Value
    ]);
  }
}

/**
 * Applies row-level conditional formatting to Action Center based on Column G.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Action Center sheet
 */
function applyActionCenterConditionalFormatting(sheet) {
  var maxRows = Math.max(sheet.getMaxRows() - 1, 100);
  var range = sheet.getRange(2, 1, maxRows, 15); // Applies across columns A:O

  // Custom formula rules using absolute column $G2
  var ruleOverdue = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Overdue"')
    .setBackground('#FCE8E6')
    .setFontColor('#C5221F')
    .setRanges([range])
    .build();

  var ruleEscalated = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Escalated"')
    .setBackground('#FAD2CF')
    .setFontColor('#A50E0E')
    .setRanges([range])
    .build();

  var ruleDueToday = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Due Today"')
    .setBackground('#FEF7E0')
    .setFontColor('#B06000')
    .setRanges([range])
    .build();

  var ruleUpcoming = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Upcoming"')
    .setBackground('#E8F0FE')
    .setFontColor('#1A73E8')
    .setRanges([range])
    .build();

  var ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Completed"')
    .setBackground('#E6F4EA')
    .setFontColor('#137333')
    .setRanges([range])
    .build();

  var ruleAwaitingReview = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="Awaiting Manager Review"')
    .setBackground('#F3E8FD')
    .setFontColor('#7627BB')
    .setRanges([range])
    .build();

  sheet.setConditionalFormatRules([
    ruleOverdue,
    ruleEscalated,
    ruleDueToday,
    ruleUpcoming,
    ruleCompleted,
    ruleAwaitingReview
  ]);
}

/**
 * Applies a consistent header style (dark navy background, bold white text, centered)
 * to row 1 across the specified number of columns.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {number} numCols - Number of columns in header row
 */
function applySheetHeaderStyle(sheet, numCols) {
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setBackground('#1F4E79');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
}

/**
 * Sets specific widths in pixels for columns starting at column 1.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {Array<number>} widths - Array of pixel widths for each column
 */
function setColumnWidths(sheet, widths) {
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}
