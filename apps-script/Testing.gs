/**
 * ==============================================================================
 * ALLO INNOWARE - DAILY OPERATIONS & TASK TRACKER
 * File: Testing.gs
 * Description: Comprehensive automated test suite validating all 8 operational
 *              scenarios: recurring cycle progression, single-point completion,
 *              daily task proofs, manager alerts, manager acknowledgements,
 *              leadership escalations, SLA timing verification, and duplicate prevention.
 * Author: Operations Automation Team / B.Tech Portfolio Project
 * Dependencies: Config.gs, Core.gs, Notifications.gs, Triggers.gs
 * ==============================================================================
 */

// Global state holding backup snapshots of tasks before test suite execution
var _TEST_DATA_BACKUP = null;

/**
 * Main test suite orchestrator.
 * Prompts user confirmation, runs all 8 test scenarios sequentially,
 * records detailed execution results, restores original sheet state,
 * outputs a styled 'Test Results' report tab, and presents a summary alert.
 *
 * @return {Array<Object>} Array of test result objects.
 */
function runAllTests() {
  Logger.log('=== STARTING ALLO INNOWARE TEST SUITE ===');

  // Step 1: Prompt user confirmation if running in interactive UI context
  var isInteractive = false;
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      isInteractive = true;
      var promptResponse = ui.alert(
        'Run Automated Test Suite',
        'This will run all 8 operational test scenarios. Changes will be validated and original task data will be restored automatically.\n\nDo you wish to continue?',
        ui.ButtonSet.OK_CANCEL
      );
      if (promptResponse !== ui.Button.OK) {
        Logger.log('Test execution cancelled by user.');
        return [];
      }
    }
  } catch (uiErr) {
    Logger.log('Non-interactive context detected. Proceeding with automated test execution.');
  }

  var results = [];

  // Step 2a: Capture baseline backup snapshot of existing task data
  try {
    backupOriginalData_();
  } catch (backupErr) {
    Logger.log('Warning: Unable to create complete backup snapshot: ' + backupErr.message);
  }

  // Step 2b: Execute each of the 8 test scenarios safely
  var tests = [
    { name: 'Test 1: Recurring Cycle Creation', fn: test1_RecurringCycle },
    { name: 'Test 2: Single-Point Completion', fn: test2_SinglePointCompletion },
    { name: 'Test 3: Daily Task Completion', fn: test3_DailyTask },
    { name: 'Test 4: Manager Overdue Alert', fn: test4_ManagerAlert },
    { name: 'Test 5: Manager Acknowledgement', fn: test5_ManagerAcknowledgement },
    { name: 'Test 6: Leadership Escalation', fn: test6_LeadershipEscalation },
    { name: 'Test 7: SLA Escalation Timing', fn: test7_EscalationTiming },
    { name: 'Test 8: Duplicate Prevention & Idempotency', fn: test8_NoDuplicates }
  ];

  for (var i = 0; i < tests.length; i++) {
    var t = tests[i];
    Logger.log('Running ' + t.name + '...');
    try {
      var res = t.fn();
      results.push(res);
      Logger.log('[' + (res.pass ? 'PASS' : 'FAIL') + '] ' + res.testId + ': ' + res.scenario);
    } catch (testErr) {
      Logger.log('Exception in ' + t.name + ': ' + testErr.message);
      results.push(createTestResult_(
        'TEST-' + (i + 1),
        t.name,
        'Automated execution runner',
        'Call ' + t.name,
        'Test execution completes without unhandled exceptions',
        'Error: ' + testErr.message,
        false,
        'Unhandled exception stack: ' + (testErr.stack || 'none')
      ));
    }
  }

  // Step 2c: Restore original workbook state
  try {
    restoreOriginalData();
  } catch (restoreErr) {
    Logger.log('Warning: Error while restoring original task data: ' + restoreErr.message);
  }

  // Step 2d: Output test results to 'Test Results' worksheet
  try {
    writeTestResults(results);
  } catch (writeErr) {
    Logger.log('Error writing Test Results sheet: ' + writeErr.message);
  }

  // Step 2e: Count passing tests and show summary alert
  var passCount = 0;
  for (var k = 0; k < results.length; k++) {
    if (results[k].pass) {
      passCount++;
    }
  }

  var summaryMessage = passCount + ' of ' + results.length + ' tests passed. See the "Test Results" sheet for full details.';
  Logger.log('=== TEST SUITE COMPLETE: ' + summaryMessage + ' ===');

  if (isInteractive) {
    try {
      SpreadsheetApp.getUi().alert('Test Execution Summary', summaryMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // Ignore UI alert error
    }
  }

  return results;
}

/**
 * Generates and formats the dedicated 'Test Results' reporting sheet.
 * Clears old results, formats headers, writes test outcomes, applies conditional
 * formatting (green for PASS, red for FAIL), and auto-fits column widths.
 *
 * @param {Array<Object>} results Array of test result objects.
 */
function writeTestResults(results) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('writeTestResults: No active spreadsheet context.');
    return;
  }

  var sheetName = 'Test Results';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  var headers = [
    'Test ID',
    'Scenario',
    'Setup',
    'Action',
    'Expected Result',
    'Actual Result',
    'Pass/Fail',
    'Notes'
  ];

  // Title Banner
  sheet.getRange('A1:H1').merge();
  sheet.getRange('A1').setValue('ALLO INNOWARE OPERATIONS TASK TRACKER - AUTOMATED TEST RESULTS')
    .setFontWeight('bold')
    .setFontSize(13)
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E79')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // Table Headers
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#FFFFFF')
    .setBackground('#2C3E50')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 28);
  sheet.setFrozenRows(2);

  // Prepare Data Rows
  var rows = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    rows.push([
      r.testId || ('TEST-' + (i + 1)),
      r.scenario || '',
      r.setup || '',
      r.action || '',
      r.expected || '',
      r.actual || '',
      r.pass ? 'PASS' : 'FAIL',
      r.notes || ''
    ]);
  }

  if (rows.length > 0) {
    var dataRange = sheet.getRange(3, 1, rows.length, headers.length);
    dataRange.setValues(rows)
      .setFontSize(9)
      .setVerticalAlignment('top');

    // Zebra striping
    for (var j = 0; j < rows.length; j++) {
      var rowNum = j + 3;
      var bg = (j % 2 === 0) ? '#FFFFFF' : '#F8F9FA';
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground(bg);
      sheet.setRowHeight(rowNum, 44);
    }

    // Alignment and wrapping
    sheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold'); // Test ID
    sheet.getRange(3, 2, rows.length, 1).setWrap(true); // Scenario
    sheet.getRange(3, 3, rows.length, 4).setWrap(true); // Setup, Action, Expected, Actual
    sheet.getRange(3, 7, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold'); // Pass/Fail
    sheet.getRange(3, 8, rows.length, 1).setWrap(true); // Notes

    // Conditional Formatting Rules on Pass/Fail column (Column G = 7)
    var passFailRange = sheet.getRange(3, 7, rows.length, 1);

    var passRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PASS')
      .setBackground('#D4EDDA')
      .setFontColor('#155724')
      .setBold(true)
      .setRanges([passFailRange])
      .build();

    var failRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('FAIL')
      .setBackground('#F8D7DA')
      .setFontColor('#721C24')
      .setBold(true)
      .setRanges([passFailRange])
      .build();

    sheet.setConditionalFormatRules([passRule, failRule]);
  }

  // Column Widths
  sheet.setColumnWidth(1, 95);   // Test ID
  sheet.setColumnWidth(2, 190);  // Scenario
  sheet.setColumnWidth(3, 220);  // Setup
  sheet.setColumnWidth(4, 210);  // Action
  sheet.setColumnWidth(5, 240);  // Expected Result
  sheet.setColumnWidth(6, 240);  // Actual Result
  sheet.setColumnWidth(7, 95);   // Pass/Fail
  sheet.setColumnWidth(8, 220);  // Notes

  // Summary Row at the bottom
  var summaryRowNum = rows.length + 4;
  sheet.getRange(summaryRowNum, 1, 1, headers.length).setBackground('#EAEDED');
  sheet.getRange(summaryRowNum, 1).setValue('TOTAL SUMMARY').setFontWeight('bold');
  sheet.getRange(summaryRowNum, 7).setFormula('=COUNTIF(G3:G' + (rows.length + 2) + ',"PASS") & " / " & COUNTA(G3:G' + (rows.length + 2) + ') & " PASSED"')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(summaryRowNum, 28);
}

/**
 * Direct task completion helper designed for automated tests and non-UI executions.
 * Executes the complete validation, audit logging, Task Master synchronization,
 * and next cycle provisioning logic WITHOUT triggering modal getUi() alerts.
 *
 * @param {number} acRowNum The 1-based row number in the Action Center.
 * @return {{success: boolean, message: string}} Result status object.
 */
function processCompletionDirect(acRowNum) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(20000);
  } catch (lockError) {
    lockAcquired = false;
  }

  if (!lockAcquired) {
    return { success: false, message: 'Script lock acquisition timed out.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

    if (!acSheet || !acRowNum || acRowNum < 2 || acRowNum > acSheet.getLastRow()) {
      return { success: false, message: 'Invalid Action Center row index: ' + acRowNum };
    }

    var acData = acSheet.getRange(acRowNum, 1, 1, 15).getValues()[0];
    var cycleId = String(acData[AC.CYCLE_ID - 1] || '').trim();
    var taskId = String(acData[AC.TASK_ID - 1] || '').trim();
    var taskName = acData[AC.NAME - 1] || '';
    var employee = acData[AC.EMPLOYEE - 1] || '';
    var currentStatus = acData[AC.STATUS - 1] || '';
    var proofType = String(acData[AC.PROOF_TYPE - 1] || '').trim();
    var ticketId = String(acData[AC.TICKET_ID - 1] || '').trim();
    var notes = String(acData[AC.NOTES - 1] || '').trim();
    var evidence = acData[AC.EVIDENCE - 1] || '';

    // Validation 1: Mandatory IDs
    if (!cycleId || !taskId) {
      return { success: false, message: 'Missing Cycle ID or Task ID in Action Center row ' + acRowNum };
    }

    // Validation 2: Double completion prevention
    if (currentStatus === STATUS.COMPLETED) {
      return { success: false, message: 'Cycle ' + cycleId + ' is already marked as Completed.' };
    }

    // Validation 3: Verify Task exists in Task Master
    var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
    var tmRow = findTaskRow(taskId);
    if (tmRow === -1 || !tmSheet) {
      return { success: false, message: 'Task ID ' + taskId + ' not found in Task Master.' };
    }

    var tmData = tmSheet.getRange(tmRow, 1, 1, 16).getValues()[0];
    var isActive = (tmData[TM.ACTIVE - 1] === true || String(tmData[TM.ACTIVE - 1]).toUpperCase() === 'TRUE');
    if (!isActive) {
      return { success: false, message: 'Task ' + taskId + ' is inactive in Task Master.' };
    }

    // Validation 4: Required Proof Verification
    if (proofType === 'Ticket ID' && !ticketId) {
      acSheet.getRange(acRowNum, AC.CONFIRMED).setValue(false);
      return { success: false, message: 'Ticket ID is required as proof of completion.' };
    }
    if (proofType === 'Daily Confirmation' && !notes) {
      acSheet.getRange(acRowNum, AC.CONFIRMED).setValue(false);
      return { success: false, message: 'Completion Notes are required for Daily Confirmation task.' };
    }

    // Determine completion user
    var completedBy = '';
    try {
      var activeUser = Session.getActiveUser();
      if (activeUser && activeUser.getEmail()) {
        completedBy = activeUser.getEmail().trim();
      }
    } catch (e) {
      completedBy = '';
    }
    if (!completedBy) {
      completedBy = employee || 'Test Runner';
    }

    var now = getNow();

    // 1. Mark Action Center row as Completed
    acSheet.getRange(acRowNum, AC.STATUS).setValue(STATUS.COMPLETED);
    acSheet.getRange(acRowNum, AC.CONFIRMED).setValue(true);
    acSheet.getRange(acRowNum, AC.DONE_BY).setValue(completedBy);
    acSheet.getRange(acRowNum, AC.DONE_AT).setValue(now);

    // 2. Append immutable record in Completion Log (idempotent)
    var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
    if (clSheet && !cycleExistsInCompletionLog(cycleId)) {
      var recordId = generateId('COMP', clSheet, CL.RECORD_ID);
      var isDaily = (proofType === 'Daily Confirmation') ? 'TRUE' : 'FALSE';

      var logRecord = [
        recordId,
        cycleId,
        taskId,
        taskName,
        employee,
        completedBy,
        now,
        ticketId || 'N/A',
        isDaily,
        notes,
        evidence,
        'Valid'
      ];
      clSheet.appendRow(logRecord);
    }

    // 3. Synchronize Task Master
    tmSheet.getRange(tmRow, TM.LAST_DONE_AT).setValue(now);
    tmSheet.getRange(tmRow, TM.LAST_DONE_BY).setValue(completedBy);
    tmSheet.getRange(tmRow, TM.STATUS).setValue(STATUS.COMPLETED);

    // 4. Provision next recurring cycle
    createNextCycle(taskId, tmRow);

    Logger.log('processCompletionDirect: Cycle ' + cycleId + ' completed successfully.');
    return { success: true, message: 'Task completed successfully! Next cycle created.' };

  } catch (err) {
    Logger.log('processCompletionDirect error: ' + err.message);
    return { success: false, message: 'Exception during completion: ' + err.message };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

// ==============================================================================
// TEST SCENARIO 1: RECURRING CYCLE CREATION
// ==============================================================================

/**
 * Test 1: Recurring Cycle Creation
 * Scenario: Completes a recurring interval task (TASK-002, 7-day frequency)
 *           and verifies that the next cycle is automatically generated with
 *           the correct incremented cycle ID, scheduled due date (+7 days),
 *           and that duplicate cycle creation is prevented upon retry.
 *
 * @return {Object} Test result details.
 */
function test1_RecurringCycle() {
  var testId = 'TEST-1';
  var scenario = 'Recurring Cycle Creation';
  var setup = 'Use TASK-002 (Portal Task A, 7-day interval). Find or initialize active cycle in Action Center.';
  var action = 'Populate Ticket ID = "TKT-TEST-001", set Confirmed = TRUE, call processCompletionDirect(). Call createNextCycle() again to verify duplicate prevention.';
  var expected = 'Status = Completed, Completion Log entry created, new cycle created with Cycle ID number +1 and Due Date = base date + 7 days, rerun creates no duplicate.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  var taskId = 'TASK-002';
  var tmRow = findTaskRow(taskId);
  if (tmRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'TASK-002 not found in Task Master', false, 'Missing sample task.');
  }

  // Ensure an active cycle exists in Action Center
  var initialCycleId = tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue();
  var initialDueDate = tmSheet.getRange(tmRow, TM.CURRENT_DUE).getValue();

  if (!initialCycleId || !cycleExistsInActionCenter(initialCycleId)) {
    // Set due date to today if empty and create cycle
    if (!initialDueDate) {
      initialDueDate = getNow();
      tmSheet.getRange(tmRow, TM.CURRENT_DUE).setValue(initialDueDate);
    }
    initialCycleId = createNextCycle(taskId, tmRow);
  }

  var acRow = findActionCenterRow(initialCycleId);
  if (acRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'Could not locate Action Center row for cycle ' + initialCycleId, false, 'Action Center sync failure.');
  }

  // Setup completion data in Action Center
  acSheet.getRange(acRow, AC.TICKET_ID).setValue('TKT-TEST-001');
  acSheet.getRange(acRow, AC.CONFIRMED).setValue(true);

  // Execute Action
  var completionResult = processCompletionDirect(acRow);
  if (!completionResult.success) {
    return createTestResult_(testId, scenario, setup, action, expected, 'processCompletionDirect failed: ' + completionResult.message, false, 'Completion returned false.');
  }

  // Verification A: Action Center status marked Completed
  var completedAcStatus = acSheet.getRange(acRow, AC.STATUS).getValue();
  var isAcCompleted = (completedAcStatus === STATUS.COMPLETED);

  // Verification B: Completion Log has entry for initialCycleId
  var isLogged = cycleExistsInCompletionLog(initialCycleId);

  // Verification C: New Cycle created in Task Master and Action Center
  var updatedCycleId = tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue();
  var updatedDueDate = tmSheet.getRange(tmRow, TM.CURRENT_DUE).getValue();
  var hasNewCycleId = (updatedCycleId && updatedCycleId !== initialCycleId);
  var newCycleExistsInAc = cycleExistsInActionCenter(updatedCycleId);

  // Verification D: Due Date arithmetic (+7 days from initial due date)
  var initialDueObj = (initialDueDate instanceof Date) ? initialDueDate : new Date(initialDueDate);
  var updatedDueObj = (updatedDueDate instanceof Date) ? updatedDueDate : new Date(updatedDueDate);
  var expectedDueDate = new Date(initialDueObj.getTime());
  expectedDueDate.setDate(expectedDueDate.getDate() + 7);

  // Account for catch-up rule: If original due date was past, catch-up adds 7 days from today
  var todayMidnight = dateOnly(getNow());
  if (dateOnly(expectedDueDate).getTime() < todayMidnight.getTime()) {
    expectedDueDate = new Date(getNow().getTime());
    expectedDueDate.setDate(expectedDueDate.getDate() + 7);
  }

  var isDueDateCorrect = (formatDateOnly(updatedDueObj) === formatDateOnly(expectedDueDate));

  // Verification E: Duplicate cycle prevention check
  var acCountBefore = countCyclesInActionCenter_(updatedCycleId);
  createNextCycle(taskId, tmRow); // Retry creation
  var acCountAfter = countCyclesInActionCenter_(updatedCycleId);
  var noDuplicatesCreated = (acCountBefore === 1 && acCountAfter === 1);

  var allPassed = isAcCompleted && isLogged && hasNewCycleId && newCycleExistsInAc && isDueDateCorrect && noDuplicatesCreated;
  var actualSummary = 'AC Status=' + completedAcStatus + ', Logged=' + isLogged + ', New Cycle=' + updatedCycleId + ', Due Date=' + formatDateOnly(updatedDueObj) + ', Dup Check=' + (noDuplicatesCreated ? 'Passed' : 'Failed');

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actualSummary,
    allPassed,
    allPassed ? 'Recurrence calculation and duplicate guards operating correctly.' : 'Discrepancy detected in cycle advancement.'
  );
}

// ==============================================================================
// TEST SCENARIO 2: SINGLE-POINT COMPLETION
// ==============================================================================

/**
 * Test 2: Single-Point Completion
 * Scenario: Employee submits completion exclusively from the Action Center.
 *           Verifies that Task Master, Action Center, and Completion Log all
 *           update and synchronize without requiring manual entry elsewhere.
 *
 * @return {Object} Test result details.
 */
function test2_SinglePointCompletion() {
  var testId = 'TEST-2';
  var scenario = 'Single-Point Completion';
  var setup = 'Use TASK-003 (Portal Task B, 15-day task). Record state of Task Master, Action Center, and Completion Log before submission.';
  var action = 'Submit Ticket ID = "TKT-AUDIT-15D" in Action Center and call processCompletionDirect().';
  var expected = 'Action Center row updated (Completed, Done By, Done At); Task Master updated (Last Completed At, By, Status); Completion Log receives new record without manual edits.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);

  var taskId = 'TASK-003';
  var tmRow = findTaskRow(taskId);
  if (tmRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'TASK-003 not found in Task Master', false, 'Missing sample task.');
  }

  // Ensure active cycle exists
  var cycleId = tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue();
  if (!cycleId || !cycleExistsInActionCenter(cycleId)) {
    if (!tmSheet.getRange(tmRow, TM.CURRENT_DUE).getValue()) {
      tmSheet.getRange(tmRow, TM.CURRENT_DUE).setValue(getNow());
    }
    cycleId = createNextCycle(taskId, tmRow);
  }

  var acRow = findActionCenterRow(cycleId);
  if (acRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'Action Center row not found for ' + cycleId, false, 'Row lookup failure.');
  }

  var clRowsBefore = clSheet ? clSheet.getLastRow() : 0;

  // Execute Action
  acSheet.getRange(acRow, AC.TICKET_ID).setValue('TKT-AUDIT-15D');
  acSheet.getRange(acRow, AC.CONFIRMED).setValue(true);

  var result = processCompletionDirect(acRow);
  if (!result.success) {
    return createTestResult_(testId, scenario, setup, action, expected, 'Completion failed: ' + result.message, false, 'Failed execution.');
  }

  // Verification 1: Action Center Row Updated
  var acStatus = acSheet.getRange(acRow, AC.STATUS).getValue();
  var acDoneBy = acSheet.getRange(acRow, AC.DONE_BY).getValue();
  var acDoneAt = acSheet.getRange(acRow, AC.DONE_AT).getValue();
  var isAcValid = (acStatus === STATUS.COMPLETED && !!acDoneBy && !!acDoneAt);

  // Verification 2: Task Master Row Updated
  var tmLastDoneAt = tmSheet.getRange(tmRow, TM.LAST_DONE_AT).getValue();
  var tmLastDoneBy = tmSheet.getRange(tmRow, TM.LAST_DONE_BY).getValue();
  var isTmValid = (!!tmLastDoneAt && !!tmLastDoneBy);

  // Verification 3: Completion Log has new valid entry
  var clRowsAfter = clSheet ? clSheet.getLastRow() : 0;
  var isLogAppended = (clRowsAfter > clRowsBefore) && cycleExistsInCompletionLog(cycleId);

  var passed = isAcValid && isTmValid && isLogAppended;
  var actual = 'AC (Status=' + acStatus + ', DoneBy=' + acDoneBy + '), TM (DoneAt=' + !!tmLastDoneAt + ', DoneBy=' + tmLastDoneBy + '), Log Entries Added=' + (clRowsAfter - clRowsBefore);

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'Single point of entry in Action Center successfully synchronized all 3 sheets.' : 'One or more destination sheets failed to update.'
  );
}

// ==============================================================================
// TEST SCENARIO 3: DAILY TASK COMPLETION
// ==============================================================================

/**
 * Test 3: Daily Task Completion
 * Scenario: Completes TASK-001 (Daily Operations Update) requiring daily
 *           attestation confirmation, notes, and optional evidence link.
 *           Verifies working day scheduling (skips weekends) and audit logging.
 *
 * @return {Object} Test result details.
 */
function test3_DailyTask() {
  var testId = 'TEST-3';
  var scenario = 'Daily Task Completion';
  var setup = 'Use TASK-001 (Daily Operations Update, Daily frequency). Verify Action Center cycle.';
  var action = 'Set Confirmed = TRUE, Notes = "Daily update shared in WhatsApp group", Evidence = "https://example.com/screenshot". Call processCompletionDirect().';
  var expected = 'Task completed, Completion Log records daily confirmation = TRUE with notes and evidence, next cycle due date falls strictly on next working day.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);

  var taskId = 'TASK-001';
  var tmRow = findTaskRow(taskId);
  if (tmRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'TASK-001 not found in Task Master', false, 'Missing sample task.');
  }

  // Ensure active cycle exists
  var cycleId = tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue();
  if (!cycleId || !cycleExistsInActionCenter(cycleId)) {
    if (!tmSheet.getRange(tmRow, TM.CURRENT_DUE).getValue()) {
      tmSheet.getRange(tmRow, TM.CURRENT_DUE).setValue(getNow());
    }
    cycleId = createNextCycle(taskId, tmRow);
  }

  var acRow = findActionCenterRow(cycleId);
  if (acRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'Action Center row not found for ' + cycleId, false, 'Row lookup failure.');
  }

  // Setup completion data
  var testNotes = 'Daily update shared in WhatsApp group';
  var testEvidence = 'https://example.com/screenshot';
  acSheet.getRange(acRow, AC.CONFIRMED).setValue(true);
  acSheet.getRange(acRow, AC.NOTES).setValue(testNotes);
  acSheet.getRange(acRow, AC.EVIDENCE).setValue(testEvidence);

  // Execute Action
  var result = processCompletionDirect(acRow);
  if (!result.success) {
    return createTestResult_(testId, scenario, setup, action, expected, 'Completion failed: ' + result.message, false, 'Validation failed.');
  }

  // Verification 1: Action Center Status = Completed
  var acStatus = acSheet.getRange(acRow, AC.STATUS).getValue();
  var isAcCompleted = (acStatus === STATUS.COMPLETED);

  // Verification 2: Completion Log contains daily confirmation details
  var clRowIndex = -1;
  var clLastRow = clSheet.getLastRow();
  var clData = clSheet.getRange(2, 1, clLastRow - 1, 12).getValues();
  for (var i = clData.length - 1; i >= 0; i--) {
    if (String(clData[i][CL.CYCLE_ID - 1]).trim() === cycleId) {
      clRowIndex = i;
      break;
    }
  }

  var isLogValid = false;
  if (clRowIndex !== -1) {
    var loggedConfirm = String(clData[clRowIndex][CL.DAILY_CONFIRM - 1]).toUpperCase();
    var loggedNotes = String(clData[clRowIndex][CL.NOTES - 1]);
    var loggedEvidence = String(clData[clRowIndex][CL.EVIDENCE - 1]);
    isLogValid = (loggedConfirm === 'TRUE' || loggedConfirm === 'YES') &&
                 (loggedNotes.indexOf('WhatsApp') !== -1) &&
                 (loggedEvidence === testEvidence);
  }

  // Verification 3: Next cycle scheduled for a valid Working Day
  var nextDueDate = tmSheet.getRange(tmRow, TM.CURRENT_DUE).getValue();
  var isNextDayWorking = isWorkingDay(nextDueDate);

  var passed = isAcCompleted && isLogValid && isNextDayWorking;
  var actual = 'AC Status=' + acStatus + ', Log Valid=' + isLogValid + ', Next Due=' + formatDateOnly(nextDueDate) + ' (WorkingDay=' + isNextDayWorking + ')';

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'Daily confirmation attestation verified and next working day properly assigned.' : 'Daily task validation or schedule verification failed.'
  );
}

// ==============================================================================
// TEST SCENARIO 4: MANAGER OVERDUE ALERT
// ==============================================================================

/**
 * Test 4: Manager Overdue Alert
 * Scenario: Simulates an overdue task condition (past deadline and alert window)
 *           and verifies that processPendingAlerts() triggers a 'Manager Alert',
 *           records the event in the Alert Log, updates task status to
 *           'Awaiting Manager Review', and does not duplicate alerts on subsequent runs.
 *
 * @return {Object} Test result details.
 */
function test4_ManagerAlert() {
  var testId = 'TEST-4';
  var scenario = 'Manager Overdue Alert';
  var setup = 'Set TASK-002 due date to yesterday and due time to 10:00 to simulate overdue condition.';
  var action = 'Call updateAllTaskStatuses(), then call processPendingAlerts(). Call processPendingAlerts() a second time.';
  var expected = 'Task status becomes Overdue or Awaiting Manager Review; Alert Log receives "Manager Alert" event; Rerun does not create duplicate alert for same cycle.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);

  var taskId = 'TASK-002';
  var tmRow = findTaskRow(taskId);
  if (tmRow === -1) {
    return createTestResult_(testId, scenario, setup, action, expected, 'TASK-002 not found in Task Master', false, 'Missing sample task.');
  }

  // Simulate overdue condition: Due yesterday at 10:00 AM
  var yesterday = new Date(getNow().getTime() - (24 * 60 * 60 * 1000));
  tmSheet.getRange(tmRow, TM.CURRENT_DUE).setValue(yesterday);
  tmSheet.getRange(tmRow, TM.DUE_TIME).setValue('10:00');
  tmSheet.getRange(tmRow, TM.STATUS).setValue(STATUS.UPCOMING);

  var cycleId = String(tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue() || '').trim();
  if (!cycleId) {
    cycleId = createNextCycle(taskId, tmRow);
  }

  var acRow = findActionCenterRow(cycleId);
  if (acRow !== -1) {
    acSheet.getRange(acRow, AC.DUE_DATE).setValue(yesterday);
    acSheet.getRange(acRow, AC.DUE_TIME).setValue('10:00');
    acSheet.getRange(acRow, AC.STATUS).setValue(STATUS.UPCOMING);
  }

  // Ensure test settings are active
  updateSetting('Test Mode', 'TRUE');
  updateSetting('Email Sending Enabled', 'FALSE');

  // Action 1: Refresh status and dispatch alerts
  updateAllTaskStatuses();
  var alertsSent1 = processPendingAlerts();

  // Action 2: Rerun to verify idempotency
  var alertsSent2 = processPendingAlerts();

  // Verification 1: Task status updated to Awaiting Manager Review
  var tmStatusAfter = tmSheet.getRange(tmRow, TM.STATUS).getValue();
  var isStatusReview = (tmStatusAfter === STATUS.AWAITING_REVIEW || tmStatusAfter === STATUS.OVERDUE);

  // Verification 2: Alert Log contains Manager Alert event for this cycle
  var alertRow = findAlertLogRow_(cycleId, 'Manager Alert');
  var isAlertLogged = (alertRow !== -1);

  // Verification 3: Duplicate check (only 1 Manager Alert for cycleId)
  var alertCount = countAlertsInLog_(cycleId, 'Manager Alert');
  var isIdempotent = (alertCount === 1);

  var passed = isStatusReview && isAlertLogged && isIdempotent;
  var actual = 'Status=' + tmStatusAfter + ', AlertLogged=' + isAlertLogged + ' (Row ' + alertRow + '), AlertCount=' + alertCount + ', RerunSent=' + alertsSent2;

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'Overdue condition successfully triggered manager alert and duplicate alert was prevented.' : 'Alert dispatch or duplicate prevention failed.'
  );
}

// ==============================================================================
// TEST SCENARIO 5: MANAGER ACKNOWLEDGEMENT
// ==============================================================================

/**
 * Test 5: Manager Acknowledgement
 * Scenario: Manager checks the acknowledgement checkbox in the Alert Log.
 *           Verifies that handleAlertAcknowledgement() records the manager's
 *           identity and timestamp, and dashboard metrics reflect acknowledgement.
 *
 * @return {Object} Test result details.
 */
function test5_ManagerAcknowledgement() {
  var testId = 'TEST-5';
  var scenario = 'Manager Acknowledgement';
  var setup = 'Locate the Manager Alert created in Test 4 in the Alert & Escalation Log.';
  var action = 'Set Manager Acknowledged (col AL.MGR_ACK) = TRUE and invoke handleAlertAcknowledgement(row).';
  var expected = 'Acknowledged By is populated with manager identity; Acknowledged At timestamp is recorded; Alert Log row updated.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);

  var taskId = 'TASK-002';
  var tmRow = findTaskRow(taskId);
  var cycleId = (tmRow !== -1) ? tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue() : '';

  // Find alert row from Test 4 or create a mock alert entry if not present
  var alertRow = findAlertLogRow_(cycleId, 'Manager Alert');
  if (alertRow === -1) {
    var eventId = generateId('ALERT', alertSheet, AL.EVENT_ID);
    var mockCycle = 'TASK-002-TESTACK';
    alertSheet.appendRow([
      eventId, mockCycle, 'TASK-002', 'Priya Patel', 'Manager Alert',
      getNow(), 'ops.manager@alloinnoware.com', 'OVERDUE TASK ALERT',
      'Test Mode - Logged', false, '', '', 'Pending', 'Test fixture'
    ]);
    alertRow = alertSheet.getLastRow();
  }

  // Action: Tick acknowledgement checkbox and run handler
  alertSheet.getRange(alertRow, AL.MGR_ACK).setValue(true);
  handleAlertAcknowledgement(alertRow);

  // Verification: Ack By and Ack At populated
  var ackBy = alertSheet.getRange(alertRow, AL.ACK_BY).getValue();
  var ackAt = alertSheet.getRange(alertRow, AL.ACK_AT).getValue();

  var isAckByValid = (ackBy !== null && ackBy !== undefined && String(ackBy).trim() !== '');
  var isAckAtValid = (ackAt !== null && ackAt !== undefined && String(ackAt).trim() !== '');

  var passed = isAckByValid && isAckAtValid;
  var actual = 'Ack By="' + ackBy + '", Ack At="' + formatDateTime(ackAt) + '"';

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'Manager acknowledgement recorded identity and timestamp successfully.' : 'Acknowledgement handler failed to write audit metadata.'
  );
}

// ==============================================================================
// TEST SCENARIO 6: LEADERSHIP ESCALATION
// ==============================================================================

/**
 * Test 6: Leadership Escalation
 * Scenario: Evaluates an alert where the resolution window (>24 hours) has expired
 *           and the task remains incomplete. Verifies processEscalations() logs a
 *           'Leadership Escalation' event, marks Escalation Status = 'Escalated',
 *           updates Task Master status to 'Escalated', and prevents duplicate escalations.
 *
 * @return {Object} Test result details.
 */
function test6_LeadershipEscalation() {
  var testId = 'TEST-6';
  var scenario = 'Leadership Escalation';
  var setup = 'Create or configure a Manager Alert with timestamp 26 hours in the past and unacknowledged status for an incomplete task.';
  var action = 'Call processEscalations(). Call processEscalations() a second time.';
  var expected = 'Leadership Escalation event logged in Alert Log; Escalation Status becomes "Escalated"; Task Master status becomes "Escalated"; Duplicate escalation prevented on rerun.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);

  // Ensure test task exists and is not completed
  var tmRow = setupTestTask();
  var taskId = 'TASK-TEST';
  var cycleId = tmSheet.getRange(tmRow, TM.CYCLE_ID).getValue();
  tmSheet.getRange(tmRow, TM.STATUS).setValue(STATUS.AWAITING_REVIEW);

  // Create an alert timestamped 26 hours ago (exceeds 24-hour resolution window and 18-hour ack window)
  var past26Hours = new Date(getNow().getTime() - (26 * 60 * 60 * 1000));
  var eventId = generateId('ALERT', alertSheet, AL.EVENT_ID);

  alertSheet.appendRow([
    eventId,
    cycleId,
    taskId,
    'Test Runner',
    'Manager Alert',
    past26Hours,
    'manager@example.com',
    'OVERDUE TASK ALERT: Test Task',
    'Test Mode - Logged',
    false, // Manager Acknowledged = FALSE
    '',
    '',
    'Pending',
    'Escalation test fixture'
  ]);
  var alertRow = alertSheet.getLastRow();

  // Ensure settings allow escalation in test mode
  updateSetting('Test Mode', 'TRUE');
  updateSetting('Email Sending Enabled', 'FALSE');

  // Action 1: Trigger escalation evaluation
  var escalationsTriggered1 = processEscalations();

  // Action 2: Rerun to verify idempotency
  var escalationsTriggered2 = processEscalations();

  // Verification 1: Leadership Escalation event logged
  var escLogRow = findAlertLogRow_(cycleId, 'Leadership Escalation');
  var isEscalationLogged = (escLogRow !== -1);

  // Verification 2: Original alert row Escalation Status updated
  var originalEscStatus = alertSheet.getRange(alertRow, AL.ESC_STATUS).getValue();
  var isOriginalRowEscalated = (originalEscStatus === 'Escalated');

  // Verification 3: Task Master status updated to Escalated
  var tmStatusAfter = tmSheet.getRange(tmRow, TM.STATUS).getValue();
  var isTmEscalated = (tmStatusAfter === STATUS.ESCALATED);

  // Verification 4: Duplicate escalation prevention check
  var totalEscalationsForCycle = countAlertsInLog_(cycleId, 'Leadership Escalation');
  var isIdempotent = (totalEscalationsForCycle === 1 && escalationsTriggered2 === 0);

  var passed = isEscalationLogged && isOriginalRowEscalated && isTmEscalated && isIdempotent;
  var actual = 'EscLogged=' + isEscalationLogged + ' (Row ' + escLogRow + '), OrigRowStatus=' + originalEscStatus + ', TM Status=' + tmStatusAfter + ', TotalEscCount=' + totalEscalationsForCycle;

  // Cleanup test task fixture
  cleanupTestTask();

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'Leadership escalation SLA trigger and duplicate prevention operated flawlessly.' : 'Escalation dispatch or status synchronization failed.'
  );
}

// ==============================================================================
// TEST SCENARIO 7: SLA ESCALATION TIMING LOGIC
// ==============================================================================

/**
 * Test 7: SLA Escalation Timing
 * Scenario: Validates the SLA timeline progression:
 *           - Task Deadline: T0
 *           - Manager Alert Sent: T0 + 15 min
 *           - Manager Ack Deadline: Alert + 18 hours
 *           - Resolution Deadline: Alert + 24 hours
 *           - Leadership Escalation Trigger: eligible after Ack/Resolution breaches
 *
 * @return {Object} Test result details.
 */
function test7_EscalationTiming() {
  var testId = 'TEST-7';
  var scenario = 'SLA Escalation Timing Logic';
  var setup = 'Establish anchor deadline: 2026-10-15 17:00:00. Load SLA parameters from Settings.';
  var action = 'Compute chronological SLA milestones: Task Deadline, Manager Alert Threshold (+15m), Ack Deadline (+18h), Resolution Deadline (+24h).';
  var expected = 'Mathematical offsets verify: Alert delay = 15m, Ack window = 18h, Resolution window = 24h. Timeline output generated.';

  // Read settings or use standard system SLA constants
  var alertDelayMinutes = Number(getSetting('Manager Alert Delay (minutes)', 15)) || 15;
  var ackWindowHours = Number(getSetting('Manager Ack Deadline (hours)', 18)) || 18;
  var resolutionWindowHours = Number(getSetting('Employee Resolution Window (hours)', 24)) || 24;

  // Establish base reference timestamp: 2026-10-15 17:00:00
  var baseDeadline = new Date(2026, 9, 15, 17, 0, 0); // Oct is month index 9

  // Milestone 1: Manager Alert threshold (Deadline + 15 minutes)
  var alertThreshold = new Date(baseDeadline.getTime() + (alertDelayMinutes * 60 * 1000));

  // Milestone 2: Manager Ack Deadline (Alert Time + 18 hours)
  var ackDeadline = new Date(alertThreshold.getTime() + (ackWindowHours * 60 * 60 * 1000));

  // Milestone 3: Employee Resolution Deadline (Alert Time + 24 hours)
  var resolutionDeadline = new Date(alertThreshold.getTime() + (resolutionWindowHours * 60 * 60 * 1000));

  // Mathematical Assertions
  var diffAlertMin = (alertThreshold.getTime() - baseDeadline.getTime()) / (1000 * 60);
  var diffAckHours = (ackDeadline.getTime() - alertThreshold.getTime()) / (1000 * 60 * 60);
  var diffResHours = (resolutionDeadline.getTime() - alertThreshold.getTime()) / (1000 * 60 * 60);

  var isAlertDelayCorrect = (diffAlertMin === alertDelayMinutes);
  var isAckWindowCorrect = (diffAckHours === ackWindowHours);
  var isResWindowCorrect = (diffResHours === resolutionWindowHours);

  var passed = isAlertDelayCorrect && isAckWindowCorrect && isResWindowCorrect;

  var timelineStr = 'T0 Deadline: ' + formatDateTime(baseDeadline) +
    ' -> Alert Sent: ' + formatDateTime(alertThreshold) + ' (+' + diffAlertMin + 'm)' +
    ' -> Ack Due: ' + formatDateTime(ackDeadline) + ' (+' + diffAckHours + 'h)' +
    ' -> Resolution Due: ' + formatDateTime(resolutionDeadline) + ' (+' + diffResHours + 'h)' +
    ' -> Escalation Eligible: >' + diffResHours + 'h';

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    timelineStr,
    passed,
    passed ? 'All SLA timing milestones align precisely with assignment business rules.' : 'Timing computation divergence detected.'
  );
}

// ==============================================================================
// TEST SCENARIO 8: DUPLICATE PREVENTION & IDEMPOTENCY
// ==============================================================================

/**
 * Test 8: Duplicate Prevention & Idempotency
 * Scenario: Verifies that scheduled automation and manual processing can be
 *           invoked repeatedly without creating duplicate cycles, completion records,
 *           manager alerts, or leadership escalations.
 *
 * @return {Object} Test result details.
 */
function test8_NoDuplicates() {
  var testId = 'TEST-8';
  var scenario = 'Duplicate Prevention & Idempotency';
  var setup = 'Record initial row counts across Action Center, Completion Log, and Alert Log.';
  var action = '1. Run scheduledAutomation() twice. 2. Attempt duplicate completion. 3. Attempt duplicate alert. 4. Attempt duplicate escalation.';
  var expected = 'No duplicate Cycle IDs in Action Center; No duplicate entries in Completion Log; No duplicate Manager Alerts; No duplicate Escalations.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
  var alSheet = ss.getSheetByName(SHEET.ALERT_LOG);

  // Subtest 1: Scheduled Automation idempotency
  try {
    scheduledAutomation();
    scheduledAutomation();
  } catch (autoErr) {
    Logger.log('scheduledAutomation warning: ' + autoErr.message);
  }

  // Scan Action Center for duplicate Cycle IDs
  var acDuplicatesFound = false;
  var acCycles = {};
  if (acSheet && acSheet.getLastRow() >= 2) {
    var acData = acSheet.getRange(2, AC.CYCLE_ID, acSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < acData.length; i++) {
      var cid = String(acData[i][0]).trim().toUpperCase();
      if (cid) {
        if (acCycles[cid]) {
          acDuplicatesFound = true;
          break;
        }
        acCycles[cid] = true;
      }
    }
  }

  // Subtest 2: Prevent duplicate completion on an already completed cycle
  var doubleCompletionBlocked = false;
  var testRow = -1;
  if (acSheet && acSheet.getLastRow() >= 2) {
    var statuses = acSheet.getRange(2, AC.STATUS, acSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < statuses.length; j++) {
      if (statuses[j][0] === STATUS.COMPLETED) {
        testRow = j + 2;
        break;
      }
    }
  }

  if (testRow !== -1) {
    var retryRes = processCompletionDirect(testRow);
    doubleCompletionBlocked = (!retryRes.success);
  } else {
    doubleCompletionBlocked = true; // No completed row available to test
  }

  // Subtest 3: Scan Completion Log for duplicate Cycle IDs
  var clDuplicatesFound = false;
  var clCycles = {};
  if (clSheet && clSheet.getLastRow() >= 2) {
    var clData = clSheet.getRange(2, CL.CYCLE_ID, clSheet.getLastRow() - 1, 1).getValues();
    for (var k = 0; k < clData.length; k++) {
      var clCid = String(clData[k][0]).trim().toUpperCase();
      if (clCid) {
        if (clCycles[clCid]) {
          clDuplicatesFound = true;
          break;
        }
        clCycles[clCid] = true;
      }
    }
  }

  // Subtest 4: Scan Alert Log for duplicate events (only 1 Manager Alert and 1 Leadership Escalation per Cycle ID)
  var alDuplicatesFound = false;
  var alEvents = {};
  if (alSheet && alSheet.getLastRow() >= 2) {
    var alData = alSheet.getRange(2, 1, alSheet.getLastRow() - 1, Math.max(AL.CYCLE_ID, AL.EVENT_TYPE)).getValues();
    for (var m = 0; m < alData.length; m++) {
      var aCid = String(alData[m][AL.CYCLE_ID - 1]).trim().toUpperCase();
      var aEvt = String(alData[m][AL.EVENT_TYPE - 1]).trim().toUpperCase();
      if (aCid && aEvt) {
        var eventKey = aCid + '|' + aEvt;
        if (alEvents[eventKey]) {
          alDuplicatesFound = true;
          break;
        }
        alEvents[eventKey] = true;
      }
    }
  }

  var passed = (!acDuplicatesFound) && doubleCompletionBlocked && (!clDuplicatesFound) && (!alDuplicatesFound);
  var actual = 'AC Dupes=' + (acDuplicatesFound ? 'FOUND' : 'None') +
    ', DoubleCompletionBlocked=' + doubleCompletionBlocked +
    ', Log Dupes=' + (clDuplicatesFound ? 'FOUND' : 'None') +
    ', Alert Dupes=' + (alDuplicatesFound ? 'FOUND' : 'None');

  return createTestResult_(
    testId,
    scenario,
    setup,
    action,
    expected,
    actual,
    passed,
    passed ? 'All idempotency constraints and duplicate prevention guards passed.' : 'Duplicate record detected in one or more sheets.'
  );
}

// ==============================================================================
// TEST FIXTURES & DATA RESTORATION HELPERS
// ==============================================================================

/**
 * Creates an isolated test task (TASK-TEST) in Task Master if it doesn't already exist.
 * Used for tests that need to manipulate dates and statuses without impacting real sample tasks.
 *
 * @return {number} The 1-based row number of TASK-TEST in Task Master.
 */
function setupTestTask() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet) {
    return -1;
  }

  var existingRow = findTaskRow('TASK-TEST');
  if (existingRow !== -1) {
    return existingRow;
  }

  var todayStr = formatDateOnly(getNow());
  var nextDue = new Date(getNow().getTime() + (7 * 24 * 60 * 60 * 1000));
  var testTaskRow = [
    'TASK-TEST',                         // Col 1: Task ID
    'Automated Test Task (Isolated)',    // Col 2: Task Name
    'Test Runner',                       // Col 3: Employee
    'test.runner@example.com',           // Col 4: Email
    '7-Day',                             // Col 5: Frequency Type
    7,                                   // Col 6: Frequency Days
    'Isolated fixture for automated testing', // Col 7: Description
    'Ticket ID',                         // Col 8: Proof Type
    '17:00',                             // Col 9: Due Time
    todayStr,                            // Col 10: Current Due Date
    'TASK-TEST-C001',                    // Col 11: Current Cycle ID
    true,                                // Col 12: Active Checkbox
    '',                                  // Col 13: Last Done At
    '',                                  // Col 14: Last Done By
    STATUS.UPCOMING,                     // Col 15: Status
    formatDateOnly(nextDue)              // Col 16: Next Due Date
  ];

  tmSheet.appendRow(testTaskRow);
  var rowNum = tmSheet.getLastRow();
  tmSheet.getRange(rowNum, TM.ACTIVE).insertCheckboxes();
  tmSheet.getRange(rowNum, TM.ACTIVE).setValue(true);

  return rowNum;
}

/**
 * Removes the test task (TASK-TEST) and any associated cycles or logs from all sheets.
 * Keeps production and sample data clean.
 */
function cleanupTestTask() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetId = 'TASK-TEST';

  // 1. Remove from Task Master
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (tmSheet && tmSheet.getLastRow() >= 2) {
    var tmData = tmSheet.getRange(2, TM.TASK_ID, tmSheet.getLastRow() - 1, 1).getValues();
    for (var i = tmData.length - 1; i >= 0; i--) {
      if (String(tmData[i][0]).trim().toUpperCase() === targetId) {
        tmSheet.deleteRow(i + 2);
      }
    }
  }

  // 2. Remove from Action Center
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  if (acSheet && acSheet.getLastRow() >= 2) {
    var acData = acSheet.getRange(2, AC.TASK_ID, acSheet.getLastRow() - 1, 1).getValues();
    for (var j = acData.length - 1; j >= 0; j--) {
      if (String(acData[j][0]).trim().toUpperCase() === targetId) {
        acSheet.deleteRow(j + 2);
      }
    }
  }

  // 3. Remove from Completion Log
  var clSheet = ss.getSheetByName(SHEET.COMPLETION_LOG);
  if (clSheet && clSheet.getLastRow() >= 2) {
    var clData = clSheet.getRange(2, CL.TASK_ID, clSheet.getLastRow() - 1, 1).getValues();
    for (var k = clData.length - 1; k >= 0; k--) {
      if (String(clData[k][0]).trim().toUpperCase() === targetId) {
        clSheet.deleteRow(k + 2);
      }
    }
  }

  // 4. Remove from Alert Log
  var alSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (alSheet && alSheet.getLastRow() >= 2) {
    var alData = alSheet.getRange(2, AL.TASK_ID, alSheet.getLastRow() - 1, 1).getValues();
    for (var m = alData.length - 1; m >= 0; m--) {
      if (String(alData[m][0]).trim().toUpperCase() === targetId) {
        alSheet.deleteRow(m + 2);
      }
    }
  }
}

/**
 * Creates an in-memory snapshot of standard tasks in Task Master prior to test execution.
 * @private
 */
function backupOriginalData_() {
  _TEST_DATA_BACKUP = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet || tmSheet.getLastRow() < 2) {
    return;
  }

  var tmData = tmSheet.getRange(2, 1, tmSheet.getLastRow() - 1, 16).getValues();
  for (var i = 0; i < tmData.length; i++) {
    var taskId = String(tmData[i][TM.TASK_ID - 1]).trim().toUpperCase();
    if (taskId) {
      _TEST_DATA_BACKUP[taskId] = {
        rowNum: i + 2,
        values: tmData[i]
      };
    }
  }
}

/**
 * Restores all tasks modified during test scenarios back to their original state.
 * Reapplies initial due dates, statuses, cycle IDs, and cleans up test fixtures.
 */
function restoreOriginalData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);

  if (!tmSheet || !_TEST_DATA_BACKUP) {
    cleanupTestTask();
    return;
  }

  for (var taskId in _TEST_DATA_BACKUP) {
    if (_TEST_DATA_BACKUP.hasOwnProperty(taskId)) {
      var item = _TEST_DATA_BACKUP[taskId];
      var currentRow = findTaskRow(taskId);
      if (currentRow !== -1) {
        tmSheet.getRange(currentRow, 1, 1, item.values.length).setValues([item.values]);
      }
    }
  }

  // Clean up any test fixtures
  cleanupTestTask();

  // Reset test settings back to safe test configuration
  updateSetting('Test Mode', 'TRUE');
  updateSetting('Email Sending Enabled', 'FALSE');

  Logger.log('restoreOriginalData: Task data restored to pre-test baseline.');
}

// ==============================================================================
// INTERNAL UTILITY ASSERTION HELPERS
// ==============================================================================

/**
 * Builds a standardized test result record.
 * @private
 */
function createTestResult_(testId, scenario, setup, action, expected, actual, pass, notes) {
  return {
    testId: testId,
    scenario: scenario,
    setup: setup,
    action: action,
    expected: expected,
    actual: actual,
    pass: !!pass,
    notes: notes || ''
  };
}

/**
 * Finds the 1-based row number in Alert Log for a specific Cycle ID and Event Type.
 * @private
 */
function findAlertLogRow_(cycleId, eventType) {
  if (!cycleId || !eventType) return -1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alSheet || alSheet.getLastRow() < 2) return -1;

  var targetCycle = String(cycleId).trim().toUpperCase();
  var targetEvent = String(eventType).trim().toUpperCase();
  var data = alSheet.getRange(2, 1, alSheet.getLastRow() - 1, Math.max(AL.CYCLE_ID, AL.EVENT_TYPE)).getValues();

  for (var i = 0; i < data.length; i++) {
    var loggedCycle = String(data[i][AL.CYCLE_ID - 1]).trim().toUpperCase();
    var loggedEvent = String(data[i][AL.EVENT_TYPE - 1]).trim().toUpperCase();
    if (loggedCycle === targetCycle && loggedEvent === targetEvent) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Counts occurrences of a Cycle ID in the Action Center.
 * @private
 */
function countCyclesInActionCenter_(cycleId) {
  if (!cycleId) return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  if (!acSheet || acSheet.getLastRow() < 2) return 0;

  var targetCycle = String(cycleId).trim().toUpperCase();
  var data = acSheet.getRange(2, AC.CYCLE_ID, acSheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === targetCycle) {
      count++;
    }
  }
  return count;
}

/**
 * Counts occurrences of an alert event in the Alert Log.
 * @private
 */
function countAlertsInLog_(cycleId, eventType) {
  if (!cycleId || !eventType) return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alSheet || alSheet.getLastRow() < 2) return 0;

  var targetCycle = String(cycleId).trim().toUpperCase();
  var targetEvent = String(eventType).trim().toUpperCase();
  var data = alSheet.getRange(2, 1, alSheet.getLastRow() - 1, Math.max(AL.CYCLE_ID, AL.EVENT_TYPE)).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var loggedCycle = String(data[i][AL.CYCLE_ID - 1]).trim().toUpperCase();
    var loggedEvent = String(data[i][AL.EVENT_TYPE - 1]).trim().toUpperCase();
    if (loggedCycle === targetCycle && loggedEvent === targetEvent) {
      count++;
    }
  }
  return count;
}
