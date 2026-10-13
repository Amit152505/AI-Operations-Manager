/**
 * ==============================================================================
 * ALLO INNOWARE TRACKER - CORE BUSINESS LOGIC (Core.gs)
 * ==============================================================================
 *
 * Description:
 *   Core operational logic for the Allo Innoware Recurring Task Tracker.
 *   Handles task completion processing from the Action Center, lifecycle
 *   and recurring cycle creation, deadline and status calculations, and
 *   synchronization between Task Master, Action Center, and Completion Log.
 *
 * Dependencies:
 *   Config.gs (SHEET, TM, AC, CL, STATUS, and date/time helper functions)
 * ==============================================================================
 */

/**
 * Handles task completion when an employee checks the 'Completion Confirmed'
 * checkbox (column J = AC.CONFIRMED) in the Action Center.
 *
 * Workflow:
 *   1. Obtains a script lock to prevent concurrent modifications.
 *   2. Validates task existence, active status, cycle alignment, and required proofs.
 *   3. If valid, marks the Action Center row as Completed with timestamp and user.
 *   4. Appends a verifiable record into the Completion Log.
 *   5. Updates Task Master with last completion details.
 *   6. Automatically provisions the next recurring cycle.
 *   7. Alerts the user of success or rolls back the checkbox on validation failure.
 *
 * @param {number} row The 1-based row index in the Action Center being confirmed.
 */
function processCompletionFromActionCenter(row) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    // Attempt to acquire lock for up to 30 seconds to prevent race conditions
    lockAcquired = lock.tryLock(30000);
  } catch (lockError) {
    lockAcquired = false;
  }

  if (!lockAcquired) {
    try {
      SpreadsheetApp.getUi().alert(
        'System Busy',
        'Unable to acquire lock. Another user or operation is currently modifying the tracker. Please try again shortly.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) {
      Logger.log('Lock acquisition timed out and UI is unavailable.');
    }
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

    if (!acSheet || !row || row < 2 || row > acSheet.getLastRow()) {
      Logger.log('Invalid Action Center sheet or row index: ' + row);
      return;
    }

    // Read Action Center row data (Columns 1 to 15)
    var acData = acSheet.getRange(row, 1, 1, 15).getValues()[0];
    var cycleId = String(acData[AC.CYCLE_ID - 1] || '').trim();
    var taskId = String(acData[AC.TASK_ID - 1] || '').trim();
    var taskName = acData[AC.NAME - 1] || '';
    var employee = acData[AC.EMPLOYEE - 1] || '';
    var currentStatus = acData[AC.STATUS - 1] || '';
    var proofType = String(acData[AC.PROOF_TYPE - 1] || '').trim();
    var ticketId = String(acData[AC.TICKET_ID - 1] || '').trim();
    var notes = String(acData[AC.NOTES - 1] || '').trim();
    var evidence = acData[AC.EVIDENCE - 1] || '';

    var validationError = null;
    var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
    var tmRow = -1;

    // --------------------------------------------------------------------------
    // Validation Checks
    // --------------------------------------------------------------------------

    // 4a. Validate Cycle ID and Task ID presence
    if (!cycleId || !taskId) {
      validationError = 'The selected row is missing a Cycle ID or Task ID.';
    }
    // 4b. Prevent double completion
    else if (currentStatus === STATUS.COMPLETED) {
      validationError = 'This task cycle has already been completed.';
    }
    // 4c. Verify Task exists in Task Master
    else {
      tmRow = findTaskRow(taskId);
      if (tmRow === -1 || !tmSheet) {
        validationError = 'Task ID "' + taskId + '" could not be found in Task Master.';
      } else {
        var tmData = tmSheet.getRange(tmRow, 1, 1, 16).getValues()[0];

        // 4d. Task must be active
        var isActive = (tmData[TM.ACTIVE - 1] === true || String(tmData[TM.ACTIVE - 1]).toUpperCase() === 'TRUE');
        if (!isActive) {
          validationError = 'Task "' + taskId + '" is marked as inactive in Task Master.';
        }
        // 4e. Cycle ID must match the task's current cycle in Task Master
        else {
          var activeTmCycleId = String(tmData[TM.CYCLE_ID - 1] || '').trim();
          if (activeTmCycleId !== cycleId) {
            validationError = 'Cycle ID mismatch: The cycle "' + cycleId + '" does not match the active cycle "' + activeTmCycleId + '" in Task Master.';
          }
          // 4f. Required proof verification
          else if (proofType === 'Ticket ID' && !ticketId) {
            validationError = 'A valid Ticket ID is required as proof of completion for this task.';
          } else if (proofType === 'Daily Confirmation' && !notes) {
            validationError = 'Notes are mandatory for Daily Confirmation tasks before confirming completion.';
          }
        }
      }
    }

    // --------------------------------------------------------------------------
    // Handle Validation Failure
    // --------------------------------------------------------------------------
    if (validationError) {
      // Revert the confirmed checkbox back to FALSE
      acSheet.getRange(row, AC.CONFIRMED).setValue(false);

      try {
        SpreadsheetApp.getUi().alert('Validation Error', validationError, SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (uiErr) {
        Logger.log('Validation Error in row ' + row + ': ' + validationError);
      }
      return;
    }

    // --------------------------------------------------------------------------
    // Process Valid Completion
    // --------------------------------------------------------------------------

    // 6a. Identify who completed the task
    var completedBy = '';
    try {
      var activeUserEmail = Session.getActiveUser().getEmail();
      if (activeUserEmail && activeUserEmail.trim() !== '') {
        completedBy = activeUserEmail.trim();
      }
    } catch (userErr) {
      completedBy = '';
    }
    if (!completedBy) {
      completedBy = employee;
    }

    // 6b. Current system timestamp
    var now = getNow();

    // 6c. Update Action Center row
    acSheet.getRange(row, AC.STATUS).setValue(STATUS.COMPLETED);
    acSheet.getRange(row, AC.DONE_BY).setValue(completedBy);
    acSheet.getRange(row, AC.DONE_AT).setValue(now);

    // 6d. Record in Completion Log (with duplicate check)
    var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
    if (clSheet) {
      if (!cycleExistsInCompletionLog(cycleId)) {
        var recordId = generateId('COMP', clSheet, CL.RECORD_ID);
        var dailyConfirmation = (proofType === 'Daily Confirmation') ? 'TRUE' : 'FALSE';

        var logRow = [
          recordId,          // CL.RECORD_ID (1)
          cycleId,           // CL.CYCLE_ID (2)
          taskId,            // CL.TASK_ID (3)
          taskName,          // CL.NAME (4)
          employee,          // CL.EMPLOYEE (5)
          completedBy,       // CL.DONE_BY (6)
          now,               // CL.TIMESTAMP (7)
          ticketId,          // CL.TICKET_ID (8)
          dailyConfirmation, // CL.DAILY_CONFIRM (9)
          notes,             // CL.NOTES (10)
          evidence,          // CL.EVIDENCE (11)
          'Valid'            // CL.VALIDATION (12)
        ];

        clSheet.appendRow(logRow);
      } else {
        Logger.log('Cycle ' + cycleId + ' already recorded in Completion Log. Duplicate log prevented.');
      }
    }

    // 6e. Update Task Master row
    tmSheet.getRange(tmRow, TM.LAST_DONE_AT).setValue(now);
    tmSheet.getRange(tmRow, TM.LAST_DONE_BY).setValue(completedBy);
    tmSheet.getRange(tmRow, TM.STATUS).setValue(STATUS.COMPLETED);

    // 6f. Provision next recurring cycle
    createNextCycle(taskId, tmRow);

    // 6g. Display success alert
    try {
      SpreadsheetApp.getUi().alert(
        'Task Completed',
        'Task completed successfully! Next cycle has been created.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) {
      Logger.log('Task ' + taskId + ' completed successfully. Next cycle created.');
    }

  } catch (error) {
    Logger.log('Error in processCompletionFromActionCenter: ' + error.message);
    try {
      SpreadsheetApp.getUi().alert(
        'Processing Error',
        'An unexpected error occurred while processing completion: ' + error.message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiErr) {
      // Ignore if UI cannot be displayed
    }
  } finally {
    // 7. Always release lock in finally block
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * Searches the Task Master sheet for a row matching the specified Task ID.
 *
 * @param {string} taskId The unique task identifier (e.g., 'TASK-001').
 * @return {number} The 1-based row number in Task Master, or -1 if not found.
 */
function findTaskRow(taskId) {
  if (!taskId) {
    return -1;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet) {
    return -1;
  }

  var lastRow = tmSheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }

  var targetId = String(taskId).trim().toUpperCase();
  var values = tmSheet.getRange(2, TM.TASK_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    var cellValue = String(values[i][0]).trim().toUpperCase();
    if (cellValue === targetId) {
      return i + 2; // Convert 0-based offset to 1-based sheet row (row 2 onward)
    }
  }

  return -1;
}

/**
 * Creates the next recurring cycle for a task in Task Master and Action Center.
 *
 * Handles:
 *   - Calculation of next due date based on frequency (Daily vs N-day interval).
 *   - Catch-up rule for tasks that were overdue or completed late.
 *   - Incremental cycle ID generation (e.g., 'TASK-001-C004').
 *   - Prevention of duplicate cycle entries.
 *   - Updating Task Master with new cycle details and following due date.
 *   - Appending a fresh actionable row into Action Center with an unchecked checkbox.
 *
 * @param {string} taskId The unique task identifier (e.g., 'TASK-001').
 * @param {number} tmRowNum The 1-based row number in Task Master.
 * @return {string} The newly generated Cycle ID, or empty string on failure.
 */
function createNextCycle(taskId, tmRowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  if (!tmSheet || !acSheet || !taskId || !tmRowNum || tmRowNum < 2) {
    Logger.log('Invalid arguments or sheets missing in createNextCycle.');
    return '';
  }

  // 1. Read task configuration from Task Master
  var tmData = tmSheet.getRange(tmRowNum, 1, 1, 16).getValues()[0];
  var taskName = tmData[TM.NAME - 1] || '';
  var employee = tmData[TM.EMPLOYEE - 1] || '';
  var freqType = String(tmData[TM.FREQ_TYPE - 1] || '').trim();
  var freqDays = parseInt(tmData[TM.FREQ_DAYS - 1], 10);
  if (isNaN(freqDays) || freqDays <= 0) {
    freqDays = 1;
  }
  var proofType = tmData[TM.PROOF_TYPE - 1] || '';
  var dueTime = tmData[TM.DUE_TIME - 1] || '17:00';
  var currentDueDate = tmData[TM.CURRENT_DUE - 1];
  var currentCycleId = String(tmData[TM.CYCLE_ID - 1] || '').trim();

  // 2. Determine base date for next cycle calculation
  var baseDate;
  if (currentDueDate instanceof Date && !isNaN(currentDueDate.getTime())) {
    baseDate = new Date(currentDueDate.getTime());
  } else {
    baseDate = getNow();
  }

  // 3. Calculate next due date
  var nextDue;
  if (freqType === 'Daily') {
    nextDue = getNextWorkingDay(baseDate);
  } else {
    nextDue = new Date(baseDate.getTime());
    nextDue.setDate(nextDue.getDate() + freqDays);
  }

  // CATCH-UP RULE: If the calculated next due date is already in the past (before today),
  // recalculate starting from today to keep the task relevant and actionable.
  var now = getNow();
  var today = dateOnly(now);
  var nextDueDay = dateOnly(nextDue);

  if (nextDueDay.getTime() < today.getTime()) {
    if (freqType === 'Daily') {
      nextDue = getNextWorkingDay(now);
    } else {
      nextDue = new Date(now.getTime());
      nextDue.setDate(nextDue.getDate() + freqDays);
    }
  }

  // 4. Determine next cycle number and Cycle ID
  var currentCycleNum = 0;
  if (currentCycleId) {
    var match = currentCycleId.match(/C(\d+)$/i);
    if (match) {
      currentCycleNum = parseInt(match[1], 10);
    }
  }
  var nextCycleNum = currentCycleNum + 1;
  var paddedNum = String(nextCycleNum);
  while (paddedNum.length < 3) {
    paddedNum = '0' + paddedNum;
  }
  var nextCycleId = taskId + '-C' + paddedNum;

  // 5. Prevent duplicate cycle insertion in Action Center
  if (cycleExistsInActionCenter(nextCycleId)) {
    Logger.log('Cycle ID "' + nextCycleId + '" already exists in Action Center. Skipping creation.');
    return nextCycleId;
  }

  // Calculate status for the new cycle
  var nextStatus = calculateStatus(nextDue, dueTime);

  // Calculate the due date AFTER nextDue (for display in Task Master NEXT_DUE)
  var followingDue;
  if (freqType === 'Daily') {
    followingDue = getNextWorkingDay(nextDue);
  } else {
    followingDue = new Date(nextDue.getTime());
    followingDue.setDate(followingDue.getDate() + freqDays);
  }

  // 6. Update Task Master with new cycle details
  tmSheet.getRange(tmRowNum, TM.CURRENT_DUE).setValue(nextDue);
  tmSheet.getRange(tmRowNum, TM.CYCLE_ID).setValue(nextCycleId);
  tmSheet.getRange(tmRowNum, TM.STATUS).setValue(nextStatus);
  tmSheet.getRange(tmRowNum, TM.NEXT_DUE).setValue(followingDue);

  // 7. Add new row to Action Center
  var newAcRow = [
    nextCycleId, // AC.CYCLE_ID (1)
    taskId,      // AC.TASK_ID (2)
    taskName,    // AC.NAME (3)
    employee,    // AC.EMPLOYEE (4)
    nextDue,     // AC.DUE_DATE (5)
    dueTime,     // AC.DUE_TIME (6)
    nextStatus,  // AC.STATUS (7)
    proofType,   // AC.PROOF_TYPE (8)
    '',          // AC.TICKET_ID (9)
    false,       // AC.CONFIRMED (10)
    '',          // AC.NOTES (11)
    '',          // AC.EVIDENCE (12)
    '',          // AC.DONE_BY (13)
    '',          // AC.DONE_AT (14)
    ''           // AC.MGR_REVIEW (15)
  ];

  acSheet.appendRow(newAcRow);

  var newRowIndex = acSheet.getLastRow();
  var checkboxCell = acSheet.getRange(newRowIndex, AC.CONFIRMED);
  checkboxCell.insertCheckboxes();
  checkboxCell.setValue(false);

  return nextCycleId;
}

/**
 * Determines the current status of a task based on its due date and due time.
 *
 * Rules:
 *   1. If taskDueDay > today: STATUS.UPCOMING
 *   2. If taskDueDay equals today:
 *      - If now < deadline: STATUS.DUE_TODAY
 *      - Else: STATUS.OVERDUE
 *   3. If taskDueDay < today: STATUS.OVERDUE
 *
 * @param {Date|string} dueDate The date when the task is due.
 * @param {string} [dueTime] The due time (e.g., '17:00'). Defaults to '17:00'.
 * @return {string} One of STATUS.UPCOMING, STATUS.DUE_TODAY, or STATUS.OVERDUE.
 */
function calculateStatus(dueDate, dueTime) {
  if (!dueDate) {
    return STATUS.UPCOMING;
  }

  var now = getNow();
  var today = dateOnly(now);
  var taskDueDay = dateOnly(dueDate);
  var targetDueTime = dueTime || '17:00';
  var deadline = setTimeOnDate(dueDate, targetDueTime);

  var dueTimeMs = taskDueDay.getTime();
  var todayMs = today.getTime();

  // Due date is in the future
  if (dueTimeMs > todayMs) {
    return STATUS.UPCOMING;
  }
  // Due date is today
  else if (dueTimeMs === todayMs) {
    if (now.getTime() < deadline.getTime()) {
      return STATUS.DUE_TODAY;
    } else {
      return STATUS.OVERDUE;
    }
  }
  // Due date is in the past
  else {
    return STATUS.OVERDUE;
  }
}

/**
 * Iterates through all active tasks in Task Master and updates their status.
 *
 * Skips:
 *   - Inactive tasks (col TM.ACTIVE != TRUE).
 *   - Tasks that are already COMPLETED for the current cycle.
 *   - Tasks that are AWAITING_REVIEW or ESCALATED (managed by notification workflow).
 *
 * Synchronizes any recalculated statuses to both Task Master and the Action Center.
 */
function updateAllTaskStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  if (!tmSheet || !acSheet) {
    Logger.log('Task Master or Action Center sheet not found in updateAllTaskStatuses.');
    return;
  }

  var tmLastRow = tmSheet.getLastRow();
  if (tmLastRow < 2) {
    return;
  }

  // Pre-load Action Center cycle rows to avoid repeated scans
  var acLastRow = acSheet.getLastRow();
  var acCycleMap = {};
  if (acLastRow >= 2) {
    var acCycles = acSheet.getRange(2, AC.CYCLE_ID, acLastRow - 1, 1).getValues();
    for (var k = 0; k < acCycles.length; k++) {
      var cId = String(acCycles[k][0]).trim().toUpperCase();
      if (cId) {
        acCycleMap[cId] = k + 2; // 1-based row number in Action Center
      }
    }
  }

  // Batch read all Task Master rows
  var tmData = tmSheet.getRange(2, 1, tmLastRow - 1, 16).getValues();

  for (var i = 0; i < tmData.length; i++) {
    var row = tmData[i];
    var tmRowIndex = i + 2;

    // Check if task is active
    var isActive = (row[TM.ACTIVE - 1] === true || String(row[TM.ACTIVE - 1]).toUpperCase() === 'TRUE');
    if (!isActive) {
      continue;
    }

    var currentStatus = row[TM.STATUS - 1];

    // Skip tasks already completed or undergoing manager escalation/review
    if (currentStatus === STATUS.COMPLETED ||
        currentStatus === STATUS.AWAITING_REVIEW ||
        currentStatus === STATUS.ESCALATED) {
      continue;
    }

    var dueDate = row[TM.CURRENT_DUE - 1];
    var dueTime = row[TM.DUE_TIME - 1] || '17:00';
    if (!dueDate) {
      continue;
    }

    var newStatus = calculateStatus(dueDate, dueTime);

    // If status changed in Task Master, update it
    if (newStatus !== currentStatus) {
      tmSheet.getRange(tmRowIndex, TM.STATUS).setValue(newStatus);
    }

    // Synchronize status to Action Center row if corresponding cycle exists
    var cycleId = String(row[TM.CYCLE_ID - 1] || '').trim().toUpperCase();
    if (cycleId && acCycleMap[cycleId]) {
      var acRowIndex = acCycleMap[cycleId];
      var acCurrentStatus = acSheet.getRange(acRowIndex, AC.STATUS).getValue();

      // Only synchronize if Action Center row is not completed or under review
      if (acCurrentStatus !== newStatus &&
          acCurrentStatus !== STATUS.COMPLETED &&
          acCurrentStatus !== STATUS.AWAITING_REVIEW &&
          acCurrentStatus !== STATUS.ESCALATED) {
        acSheet.getRange(acRowIndex, AC.STATUS).setValue(newStatus);
      }
    }
  }
}

/**
 * Rebuilds the Action Center from scratch based on active tasks in Task Master.
 *
 * WARNING:
 *   This clears all existing rows (rows 2+) in Action Center.
 *   Any unsaved manual employee inputs (notes, tickets, checkboxes) will be cleared.
 *   Use primarily for initial setup, system recovery, or after bulk Task Master updates.
 */
function refreshActionCenter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  if (!tmSheet || !acSheet) {
    Logger.log('Task Master or Action Center sheet not found in refreshActionCenter.');
    return;
  }

  // 1. Clear existing Action Center data rows (preserving headers in row 1)
  var acLastRow = acSheet.getLastRow();
  var acLastCol = Math.max(acSheet.getLastColumn(), 15);
  if (acLastRow >= 2) {
    var clearRange = acSheet.getRange(2, 1, acLastRow - 1, acLastCol);
    clearRange.clearContent();
    clearRange.clearFormat();
    clearRange.removeCheckboxes();
  }

  // 2. Read all data from Task Master
  var tmLastRow = tmSheet.getLastRow();
  if (tmLastRow < 2) {
    return;
  }

  var tmData = tmSheet.getRange(2, 1, tmLastRow - 1, 16).getValues();
  var newRows = [];

  // 3. For each active task with a current cycle ID, prepare an Action Center row
  for (var i = 0; i < tmData.length; i++) {
    var row = tmData[i];
    var isActive = (row[TM.ACTIVE - 1] === true || String(row[TM.ACTIVE - 1]).toUpperCase() === 'TRUE');
    var cycleId = String(row[TM.CYCLE_ID - 1] || '').trim();

    if (!isActive || !cycleId) {
      continue;
    }

    var taskId = row[TM.TASK_ID - 1];
    var taskName = row[TM.NAME - 1];
    var employee = row[TM.EMPLOYEE - 1];
    var dueDate = row[TM.CURRENT_DUE - 1];
    var dueTime = row[TM.DUE_TIME - 1] || '17:00';
    var proofType = row[TM.PROOF_TYPE - 1];

    // Determine status: preserve completed/review/escalated if applicable, else recalculate
    var status = row[TM.STATUS - 1];
    if (status !== STATUS.COMPLETED && status !== STATUS.AWAITING_REVIEW && status !== STATUS.ESCALATED) {
      status = calculateStatus(dueDate, dueTime);
    }

    var acRow = [
      cycleId,        // AC.CYCLE_ID (1)
      taskId,         // AC.TASK_ID (2)
      taskName,       // AC.NAME (3)
      employee,       // AC.EMPLOYEE (4)
      dueDate,        // AC.DUE_DATE (5)
      dueTime,        // AC.DUE_TIME (6)
      status,         // AC.STATUS (7)
      proofType,      // AC.PROOF_TYPE (8)
      '',             // AC.TICKET_ID (9)
      false,          // AC.CONFIRMED (10)
      '',             // AC.NOTES (11)
      '',             // AC.EVIDENCE (12)
      '',             // AC.DONE_BY (13)
      '',             // AC.DONE_AT (14)
      ''              // AC.MGR_REVIEW (15)
    ];

    newRows.push(acRow);
  }

  // 4. Batch populate the Action Center
  if (newRows.length > 0) {
    var insertRange = acSheet.getRange(2, 1, newRows.length, 15);
    insertRange.setValues(newRows);

    // Insert checkboxes for column J (AC.CONFIRMED)
    var checkboxRange = acSheet.getRange(2, AC.CONFIRMED, newRows.length, 1);
    checkboxRange.insertCheckboxes();
  }
}

/**
 * Finds the 1-based row number in Action Center matching the given Cycle ID.
 *
 * @param {string} cycleId The unique cycle identifier (e.g., 'TASK-001-C001').
 * @return {number} The 1-based row index in Action Center, or -1 if not found.
 */
function findActionCenterRow(cycleId) {
  if (!cycleId) {
    return -1;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  if (!acSheet) {
    return -1;
  }

  var lastRow = acSheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }

  var searchId = String(cycleId).trim().toUpperCase();
  var values = acSheet.getRange(2, AC.CYCLE_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    var cellValue = String(values[i][0]).trim().toUpperCase();
    if (cellValue === searchId) {
      return i + 2; // 1-based row number
    }
  }

  return -1;
}

/**
 * Checks whether the specified cycle ID already exists in the Action Center.
 *
 * @param {string} cycleId The cycle ID to look up.
 * @return {boolean} True if the cycle ID exists, false otherwise.
 */
function cycleExistsInActionCenter(cycleId) {
  return findActionCenterRow(cycleId) !== -1;
}

/**
 * Checks whether the specified cycle ID already exists in the Completion Log.
 * Used to ensure idempotent completions and prevent duplicate audit records.
 *
 * @param {string} cycleId The cycle ID to look up.
 * @return {boolean} True if the cycle ID is recorded in Completion Log, false otherwise.
 */
function cycleExistsInCompletionLog(cycleId) {
  if (!cycleId) {
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
  if (!clSheet) {
    return false;
  }

  var lastRow = clSheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var searchId = String(cycleId).trim().toUpperCase();
  var values = clSheet.getRange(2, CL.CYCLE_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    var cellValue = String(values[i][0]).trim().toUpperCase();
    if (cellValue === searchId) {
      return true;
    }
  }

  return false;
}
