/**
 * @fileoverview Triggers, custom menu, and event handlers for Task Tracker.
 * Manages workbook menus, installable onEdit routing, time-driven background
 * automations, and trigger lifecycle management.
 *
 * Dependencies:
 * - Config.gs: SHEET, TM, AC, AL, STATUS constants, getNow(), formatDateTime(), updateSetting()
 * - Core.gs: updateAllTaskStatuses(), refreshActionCenter(), processCompletionFromActionCenter()
 * - Notifications.gs: processPendingAlerts(), processEscalations(), handleAlertAcknowledgement()
 * - Setup.gs: setupWorkbook()
 * - Testing.gs: runAllTests()
 */

/**
 * Simple trigger that runs automatically when the spreadsheet is opened.
 * Builds the 'Task Tracker' custom menu for manual operations,
 * setup utilities, and trigger management.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Task Tracker')
    .addItem('Setup Workbook', 'setupWorkbook')
    .addSeparator()
    .addItem('Refresh Task Statuses', 'updateAllTaskStatuses')
    .addItem('Refresh Action Center', 'refreshActionCenter')
    .addSeparator()
    .addItem('Run Pending Task Check', 'runManualCheck')
    .addItem('Run Escalation Check', 'runManualEscalationCheck')
    .addSeparator()
    .addItem('Install Automation Triggers', 'installTriggers')
    .addItem('Remove All Triggers', 'removeTriggers')
    .addSeparator()
    .addItem('Run All Tests', 'runAllTests')
    .addToUi();
}

/**
 * Installable onEdit trigger handler.
 * Routes sheet cell edits to the appropriate business logic handlers.
 * All logic is wrapped in a try/catch block to prevent unhandled exceptions
 * from surfacing confusing error dialogs to the end user.
 *
 * Steps:
 * 1. Get the edited sheet name
 * 2. Get the edited row and column
 * 3. Skip if row <= 1 (header row)
 * 4. If sheet is Action Center AND column is AC.CONFIRMED (10) AND new value is TRUE:
 *    Call processCompletionFromActionCenter(row)
 * 5. If sheet is Alert & Escalation Log AND column is AL.MGR_ACK (10) AND new value is TRUE:
 *    Call handleAlertAcknowledgement(row)
 *
 * @param {Object} e The Google Sheets edit event object.
 */
function onEditHandler(e) {
  try {
    if (!e || !e.range) {
      Logger.log('onEditHandler: No event or range object passed.');
      return;
    }

    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var row = range.getRow();
    var col = range.getColumn();

    // Skip header row
    if (row <= 1) {
      return;
    }

    // Resolve target sheet names and column positions from Config constants
    var actionCenterSheetName = (typeof SHEET !== 'undefined' && SHEET.ACTION_CENTER) ? SHEET.ACTION_CENTER : 'Action Center';
    var alertLogSheetName = (typeof SHEET !== 'undefined' && (SHEET.ALERT_LOG || SHEET.ALERTS || SHEET.ALERT_ESCALATION_LOG)) 
      ? (SHEET.ALERT_LOG || SHEET.ALERTS || SHEET.ALERT_ESCALATION_LOG) 
      : 'Alert & Escalation Log';

    var colAcConfirmed = (typeof AC !== 'undefined' && AC.CONFIRMED) ? AC.CONFIRMED : 10;
    var colAlMgrAck = (typeof AL !== 'undefined' && AL.MGR_ACK) ? AL.MGR_ACK : 10;

    // Check if the edited value is truthy / checked
    var editVal = e.value;
    var isTrue = (editVal === true || editVal === 'TRUE' || editVal === 'true');
    if (!isTrue && range) {
      var currentVal = range.getValue();
      isTrue = (currentVal === true || currentVal === 'TRUE');
    }

    // Step 4: Action Center sheet - Completion confirmed checkbox
    if ((sheetName === actionCenterSheetName || sheetName === 'Action Center') && col === colAcConfirmed && isTrue) {
      Logger.log('onEditHandler: Task completion confirmed at row ' + row + ' in ' + sheetName);
      processCompletionFromActionCenter(row);
      return;
    }

    // Step 5: Alert & Escalation Log sheet - Manager acknowledgement checkbox
    if ((sheetName === alertLogSheetName || sheetName === 'Alert & Escalation Log') && col === colAlMgrAck && isTrue) {
      Logger.log('onEditHandler: Manager acknowledgement recorded at row ' + row + ' in ' + sheetName);
      handleAlertAcknowledgement(row);
      return;
    }
  } catch (err) {
    Logger.log('Error in onEditHandler: ' + err.toString());
  }
}

/**
 * Main function called by the time-driven trigger.
 * Runs all scheduled automated checks: task statuses, pending alerts, and escalations.
 * Uses LockService with a 60-second timeout to prevent overlapping executions.
 *
 * Steps:
 * 1. Obtain script lock (tryLock 60000). Return if unable to acquire.
 * 2. Execute automated routines:
 *    a. Log 'Scheduled automation started at ' + now
 *    b. Call updateAllTaskStatuses()
 *    c. Call processPendingAlerts()
 *    d. Call processEscalations()
 *    e. Call updateSetting('Last Automation Run', formatDateTime(getNow()))
 *    f. Log 'Scheduled automation completed'
 * 3. Release script lock in finally block.
 */
function scheduledAutomation() {
  var lock = LockService.getScriptLock();
  var hasLock = false;

  try {
    hasLock = lock.tryLock(60000);
  } catch (lockError) {
    Logger.log('Could not obtain script lock: ' + lockError.toString());
    return;
  }

  if (!hasLock) {
    Logger.log('Could not obtain script lock within 60000ms. Another execution is in progress.');
    return;
  }

  try {
    var now = (typeof getNow === 'function') ? getNow() : new Date();
    Logger.log('Scheduled automation started at ' + now);

    // Run core automated processes
    updateAllTaskStatuses();
    processPendingAlerts();
    processEscalations();

    // Update settings table with latest execution timestamp
    var lastRunTimestamp = (typeof formatDateTime === 'function' && typeof getNow === 'function')
      ? formatDateTime(getNow())
      : ((typeof formatDateTime === 'function') ? formatDateTime(now) : now.toString());

    if (typeof updateSetting === 'function') {
      updateSetting('Last Automation Run', lastRunTimestamp);
    }

    Logger.log('Scheduled automation completed');
  } catch (err) {
    Logger.log('Error during scheduled automation: ' + err.toString());
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {
      Logger.log('Error releasing script lock: ' + releaseErr.toString());
    }
  }
}

/**
 * Manual menu action to execute task status refresh and pending alert check.
 * Similar to scheduled automation, but provides interactive UI feedback.
 *
 * Steps:
 * 1. Call updateAllTaskStatuses()
 * 2. Call processPendingAlerts() - get count of alerts sent
 * 3. Show alert: 'Task check complete. X new alerts sent.'
 */
function runManualCheck() {
  try {
    updateAllTaskStatuses();
    var result = processPendingAlerts();
    var count = 0;

    if (typeof result === 'number') {
      count = result;
    } else if (result && typeof result.count === 'number') {
      count = result.count;
    } else if (result && typeof result.alertsSent === 'number') {
      count = result.alertsSent;
    } else if (Array.isArray(result)) {
      count = result.length;
    } else if (result) {
      count = Number(result) || 0;
    }

    SpreadsheetApp.getUi().alert('Task check complete. ' + count + ' new alerts sent.');
  } catch (err) {
    Logger.log('Error in runManualCheck: ' + err.toString());
    try {
      SpreadsheetApp.getUi().alert('Error running task check: ' + err.message);
    } catch (uiErr) {
      Logger.log('Could not display UI alert: ' + uiErr.toString());
    }
  }
}

/**
 * Manual menu action to trigger escalation check.
 *
 * Steps:
 * 1. Call processEscalations()
 * 2. Show alert: 'Escalation check complete.'
 */
function runManualEscalationCheck() {
  try {
    processEscalations();
    SpreadsheetApp.getUi().alert('Escalation check complete.');
  } catch (err) {
    Logger.log('Error in runManualEscalationCheck: ' + err.toString());
    try {
      SpreadsheetApp.getUi().alert('Error running escalation check: ' + err.message);
    } catch (uiErr) {
      Logger.log('Could not display UI alert: ' + uiErr.toString());
    }
  }
}

/**
 * Installs all required project triggers for automated operations.
 * Clears existing project triggers first to prevent duplicate trigger executions.
 *
 * Steps:
 * 1. Call removeTriggers() to clear existing triggers
 * 2. Create an installable onEdit trigger for onEditHandler
 * 3. Create a time-driven trigger for scheduledAutomation (runs every 15 minutes)
 * 4. Show alert: 'Triggers installed! The automation will run every 15 minutes.\nYou may need to authorize permissions when prompted.'
 */
function installTriggers() {
  try {
    // 1. Remove existing triggers to clear duplicates
    removeTriggers(true);

    // 2. Create installable onEdit trigger
    ScriptApp.newTrigger('onEditHandler')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    // 3. Create time-driven trigger for scheduledAutomation (every 15 minutes)
    ScriptApp.newTrigger('scheduledAutomation')
      .timeBased()
      .everyMinutes(15)
      .create();

    // 4. Notify user of successful installation
    SpreadsheetApp.getUi().alert('Triggers installed! The automation will run every 15 minutes.\nYou may need to authorize permissions when prompted.');
  } catch (err) {
    Logger.log('Error in installTriggers: ' + err.toString());
    try {
      SpreadsheetApp.getUi().alert('Error installing triggers: ' + err.message);
    } catch (uiErr) {
      Logger.log('Could not display UI alert: ' + uiErr.toString());
    }
  }
}

/**
 * Removes all triggers created by this project.
 *
 * Steps:
 * 1. Get all triggers: ScriptApp.getProjectTriggers()
 * 2. For each trigger, check if handlerFunction is 'onEditHandler' or 'scheduledAutomation'
 * 3. Delete matching triggers: ScriptApp.deleteTrigger(trigger)
 * 4. Log how many triggers were removed
 * 5. Show alert: 'All automation triggers have been removed.'
 *
 * @param {boolean} [silent=false] If true, suppresses the interactive UI alert dialog.
 * @return {number} Count of triggers removed.
 */
function removeTriggers(silent) {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var removedCount = 0;

    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      var handlerFunction = trigger.getHandlerFunction();
      if (handlerFunction === 'onEditHandler' || handlerFunction === 'scheduledAutomation') {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
      }
    }

    Logger.log('Removed ' + removedCount + ' automation trigger(s).');

    if (silent !== true) {
      SpreadsheetApp.getUi().alert('All automation triggers have been removed.');
    }

    return removedCount;
  } catch (err) {
    Logger.log('Error in removeTriggers: ' + err.toString());
    if (silent !== true) {
      try {
        SpreadsheetApp.getUi().alert('Error removing triggers: ' + err.message);
      } catch (uiErr) {
        Logger.log('Could not display UI alert: ' + uiErr.toString());
      }
    }
    return 0;
  }
}
