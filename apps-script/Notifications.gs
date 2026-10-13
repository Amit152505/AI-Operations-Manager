/**
 * @file Notifications.gs
 * @description Notification and escalation engine for the Allo Innoware Task Tracker.
 * Handles overdue task detection, manager alert emails, leadership escalations,
 * manager acknowledgement processing, and audit logging into the Alert & Escalation Log.
 *
 * All functions rely on shared configuration constants and utilities defined in Config.gs
 * and lookup helpers defined in Core.gs.
 */

// ============================================================================
// 1. MANAGER ALERT PROCESSING
// ============================================================================

/**
 * Checks all active tasks in Task Master for overdue status and sends manager alerts
 * where the alert delay window has elapsed and an alert has not yet been dispatched.
 *
 * @return {number} Count of manager alert notifications dispatched/logged.
 */
function processPendingAlerts() {
  Logger.log('Starting processPendingAlerts check...');

  // Step 1: Retrieve configuration settings
  var settings = getAllSettings();
  var alertDelaySetting = Number(settings['Manager Alert Delay (minutes)']);
  var alertDelayMinutes = !isNaN(alertDelaySetting) ? alertDelaySetting : 15;

  // Step 2: Access the Task Master sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet) {
    Logger.log('Task Master sheet not found. Aborting processPendingAlerts.');
    return 0;
  }

  var lastRow = tmSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows found in Task Master.');
    return 0;
  }

  // Read all task data rows (row 2 to lastRow)
  var taskDataRows = tmSheet.getRange(2, 1, lastRow - 1, tmSheet.getLastColumn()).getValues();
  var alertsSentCount = 0;
  var now = getNow();

  // Step 3: Iterate through each task row
  for (var i = 0; i < taskDataRows.length; i++) {
    var row = taskDataRows[i];
    var tmRowNum = i + 2;

    // Check if task is active (Column L = TM.ACTIVE)
    var isActive = (row[TM.ACTIVE - 1] === true || String(row[TM.ACTIVE - 1]).toUpperCase() === 'TRUE');
    if (!isActive) {
      continue;
    }

    // Step 3a: Skip if status is COMPLETED or UPCOMING
    var status = row[TM.STATUS - 1];
    if (status === STATUS.COMPLETED || status === STATUS.UPCOMING) {
      continue;
    }

    var cycleId = row[TM.CYCLE_ID - 1];
    if (!cycleId) {
      continue;
    }

    // Step 3b: Get current due date and due time
    var dueDate = row[TM.CURRENT_DUE - 1];
    var dueTime = row[TM.DUE_TIME - 1] || '17:00';
    if (!dueDate) {
      continue;
    }

    var dueDateObj = (dueDate instanceof Date) ? dueDate : new Date(dueDate);

    // Step 3c: Calculate deadline datetime
    var deadline = setTimeOnDate(dueDateObj, dueTime);

    // Step 3d: Calculate alert threshold time (deadline + alertDelay minutes)
    var alertTime = new Date(deadline.getTime() + (alertDelayMinutes * 60 * 1000));

    // Step 3e: Determine if overdue and beyond the alert delay threshold
    var isPastAlertTime = (now.getTime() > alertTime.getTime());
    var isTaskOverdue = false;

    if (typeof calculateStatus === 'function') {
      isTaskOverdue = (calculateStatus(dueDateObj, dueTime) === STATUS.OVERDUE || status === STATUS.OVERDUE);
    } else {
      isTaskOverdue = (now.getTime() > deadline.getTime() || status === STATUS.OVERDUE);
    }

    if (isPastAlertTime && isTaskOverdue) {
      // Check if a 'Manager Alert' has already been recorded for this cycle ID
      if (hasAlertBeenSent(cycleId, 'Manager Alert')) {
        continue;
      }

      // Send the manager alert and record in the Alert Log
      var success = sendManagerAlert(row, tmRowNum);
      if (success) {
        alertsSentCount++;
        Logger.log('Manager alert successfully processed for cycle: ' + cycleId);
      }
    }
  }

  Logger.log('processPendingAlerts completed. Total alerts sent: ' + alertsSentCount);
  return alertsSentCount;
}

/**
 * Checks the Alert & Escalation Log to verify if an alert or escalation event
 * of a given type has already been logged for a specific task cycle.
 *
 * @param {string} cycleId The cycle ID to look up (e.g., 'TASK-001-C001').
 * @param {string} eventType The event type ('Manager Alert' or 'Leadership Escalation').
 * @return {boolean} True if matching event found, false otherwise.
 */
function hasAlertBeenSent(cycleId, eventType) {
  if (!cycleId || !eventType) {
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alertSheet) {
    return false;
  }

  var lastRow = alertSheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var cycleIdCol = AL.CYCLE_ID;
  var eventTypeCol = AL.EVENT_TYPE;
  var maxCol = Math.max(cycleIdCol, eventTypeCol);

  // Read columns up to the maximum required column for performance
  var alertData = alertSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var targetCycle = String(cycleId).trim().toUpperCase();
  var targetEvent = String(eventType).trim().toUpperCase();

  for (var i = 0; i < alertData.length; i++) {
    var loggedCycle = String(alertData[i][cycleIdCol - 1]).trim().toUpperCase();
    var loggedEvent = String(alertData[i][eventTypeCol - 1]).trim().toUpperCase();

    if (loggedCycle === targetCycle && loggedEvent === targetEvent) {
      return true;
    }
  }

  return false;
}

/**
 * Sends a manager alert email for an overdue task and logs the event to the
 * Alert & Escalation Log, then updates the task and action center statuses.
 *
 * @param {Array} taskData Array of column values representing the Task Master row.
 * @param {number} tmRowNum The 1-based row number in Task Master.
 * @return {boolean} True if processed successfully (or logged in test mode), false on failure.
 */
function sendManagerAlert(taskData, tmRowNum) {
  // Step 1: Retrieve settings
  var settings = getAllSettings();
  var testModeSetting = settings['Test Mode'];
  var isTestMode = (testModeSetting === true || String(testModeSetting).toUpperCase() === 'TRUE');

  var emailEnabledSetting = settings['Email Sending Enabled'];
  var isEmailEnabled = (emailEnabledSetting === true || String(emailEnabledSetting).toUpperCase() === 'TRUE');

  var managerEmail = settings['Manager Email'] || '';
  var testRecipient = settings['Test Email Recipient'] || '';

  // Step 2: Determine recipient based on testMode
  var recipient = isTestMode ? testRecipient : managerEmail;

  // Step 3: Build email subject
  var taskId = taskData[TM.TASK_ID - 1];
  var taskName = taskData[TM.NAME - 1];
  var employee = taskData[TM.EMPLOYEE - 1];
  var cycleId = taskData[TM.CYCLE_ID - 1];
  var subject = 'OVERDUE TASK ALERT: ' + taskName + ' (' + taskId + ')';

  // Step 4: Build HTML body
  var htmlBody = buildManagerEmailHtml(taskData);

  // Step 5: Email dispatch handling
  var alertStatus = '';
  var errorNote = '';

  if (isEmailEnabled) {
    try {
      if (!recipient || recipient.trim() === '') {
        throw new Error('No recipient email configured for Manager Alert');
      }

      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        htmlBody: htmlBody
      });
      alertStatus = 'Sent';
      Logger.log('Manager Alert email sent to: ' + recipient + ' for task ' + taskId);
    } catch (err) {
      Logger.log('Failed to send manager alert email: ' + err.message);
      alertStatus = 'Failed';
      errorNote = 'Email error: ' + err.message;
    }
  } else if (isTestMode) {
    alertStatus = 'Test Mode - Logged';
    Logger.log('[TEST MODE] Manager Alert Email generated:');
    Logger.log('To: ' + recipient);
    Logger.log('Subject: ' + subject);
  } else {
    alertStatus = 'Email Disabled';
    Logger.log('[EMAIL DISABLED] Manager Alert logged without sending: ' + subject);
  }

  // Step 6: Log the event into the Alert & Escalation Log
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alertSheet) {
    Logger.log('Alert Log sheet not found. Cannot log manager alert.');
    return false;
  }

  var now = getNow();
  var eventId = generateId('ALERT', alertSheet, AL.EVENT_ID);

  var newRowValues = [
    eventId,           // AL.EVENT_ID (1)
    cycleId,           // AL.CYCLE_ID (2)
    taskId,            // AL.TASK_ID (3)
    employee,          // AL.EMPLOYEE (4)
    'Manager Alert',   // AL.EVENT_TYPE (5)
    now,               // AL.TIMESTAMP (6)
    recipient,         // AL.RECIPIENT (7)
    subject,           // AL.SUBJECT (8)
    alertStatus,       // AL.STATUS (9)
    false,             // AL.MGR_ACK (10)
    '',                // AL.ACK_BY (11)
    '',                // AL.ACK_AT (12)
    'Pending',         // AL.ESC_STATUS (13)
    errorNote          // AL.NOTES (14)
  ];

  alertSheet.appendRow(newRowValues);

  var newRowNum = alertSheet.getLastRow();
  // Insert checkbox for column J (AL.MGR_ACK) and initialize to FALSE
  alertSheet.getRange(newRowNum, AL.MGR_ACK).insertCheckboxes();
  alertSheet.getRange(newRowNum, AL.MGR_ACK).setValue(false);

  // Step 7: Update Task Master status to STATUS.AWAITING_REVIEW
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (tmSheet && tmRowNum >= 2) {
    tmSheet.getRange(tmRowNum, TM.STATUS).setValue(STATUS.AWAITING_REVIEW);
  }

  // Step 8: Update Action Center status for this cycle to STATUS.AWAITING_REVIEW
  var acRowNum = findActionCenterRowSafe_(cycleId);
  if (acRowNum >= 2) {
    var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
    if (acSheet) {
      acSheet.getRange(acRowNum, AC.STATUS).setValue(STATUS.AWAITING_REVIEW);
      // If Manager Review column exists, update it as well
      if (acSheet.getLastColumn() >= AC.MGR_REVIEW) {
        acSheet.getRange(acRowNum, AC.MGR_REVIEW).setValue(STATUS.AWAITING_REVIEW);
      }
    }
  }

  // Step 9: Return true if email sent or logged, false only if hard failure
  return (alertStatus !== 'Failed');
}

// ============================================================================
// 2. LEADERSHIP ESCALATION PROCESSING
// ============================================================================

/**
 * Evaluates all unresolved manager alerts against SLA deadlines and triggers
 * leadership escalation emails if resolution or acknowledgement time limits expire.
 *
 * Conditions triggering escalation:
 * 1. Alert sent AND resolution window has expired AND task is still not completed.
 * 2. Alert sent AND acknowledgement deadline has passed AND manager has NOT acknowledged.
 *
 * @return {number} Count of leadership escalations triggered.
 */
function processEscalations() {
  Logger.log('Starting processEscalations check...');

  // Step 1: Retrieve configuration settings
  var settings = getAllSettings();
  var resolutionWindowSetting = Number(settings['Employee Resolution Window (hours)']);
  var resolutionWindow = !isNaN(resolutionWindowSetting) ? resolutionWindowSetting : 24;

  var ackDeadlineSetting = Number(settings['Manager Ack Deadline (hours)']);
  var ackDeadline = !isNaN(ackDeadlineSetting) ? ackDeadlineSetting : 18;

  // Step 2: Read all rows from Alert & Escalation Log
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alertSheet) {
    Logger.log('Alert Log sheet not found. Aborting escalation check.');
    return 0;
  }

  var lastRow = alertSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows found in Alert Log.');
    return 0;
  }

  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet) {
    Logger.log('Task Master sheet not found. Aborting escalation check.');
    return 0;
  }

  var alertRows = alertSheet.getRange(2, 1, lastRow - 1, alertSheet.getLastColumn()).getValues();
  var escalationsSent = 0;
  var now = getNow();

  // Step 3: Iterate through each alert row
  for (var i = 0; i < alertRows.length; i++) {
    var alertRowNum = i + 2;
    var alertRow = alertRows[i];
    var eventType = alertRow[AL.EVENT_TYPE - 1];

    // Filter only for rows with Event Type = 'Manager Alert'
    if (eventType !== 'Manager Alert') {
      continue;
    }

    // Step 3a: Skip if Escalation Status is already 'Escalated' or 'Not Required'
    var escStatus = String(alertRow[AL.ESC_STATUS - 1]).trim();
    if (escStatus === 'Escalated' || escStatus === 'Not Required') {
      continue;
    }

    // Step 3b: Get the cycle ID and task ID
    var cycleId = alertRow[AL.CYCLE_ID - 1];
    var taskId = alertRow[AL.TASK_ID - 1];
    if (!cycleId || !taskId) {
      continue;
    }

    // Step 3c: Check if the task is now completed
    var tmRowNum = findTaskRowSafe_(taskId);
    if (tmRowNum < 2) {
      continue;
    }

    var currentTmStatus = tmSheet.getRange(tmRowNum, TM.STATUS).getValue();
    var currentCycleId = tmSheet.getRange(tmRowNum, TM.CYCLE_ID).getValue();

    // The task is considered complete if status is Completed or if Task Master has already advanced to a later cycle
    var isTaskCompleted = (currentTmStatus === STATUS.COMPLETED ||
      (currentCycleId && String(currentCycleId).trim() !== String(cycleId).trim()));

    if (isTaskCompleted) {
      alertSheet.getRange(alertRowNum, AL.ESC_STATUS).setValue('Not Required');
      continue;
    }

    // Step 3d: Check if a Leadership Escalation has already been sent for this cycle
    if (hasAlertBeenSent(cycleId, 'Leadership Escalation')) {
      alertSheet.getRange(alertRowNum, AL.ESC_STATUS).setValue('Escalated');
      continue;
    }

    // Step 3e: Get alert timestamp
    var alertTimestamp = alertRow[AL.TIMESTAMP - 1];
    if (!alertTimestamp) {
      continue;
    }
    var alertDateObj = (alertTimestamp instanceof Date) ? alertTimestamp : new Date(alertTimestamp);

    // Step 3f: Calculate elapsed hours since the manager alert
    var hoursElapsed = hoursDiff(alertDateObj, now);

    // Step 3g: Condition 1: hoursElapsed > resolutionWindow AND task not completed
    var condition1 = (hoursElapsed > resolutionWindow && !isTaskCompleted);

    // Step 3h: Condition 2: hoursElapsed > ackDeadline AND Manager Acknowledged is FALSE
    var mgrAck = alertRow[AL.MGR_ACK - 1];
    var isManagerAck = (mgrAck === true || String(mgrAck).toUpperCase() === 'TRUE');
    var condition2 = (hoursElapsed > ackDeadline && !isManagerAck);

    // Step 3i: Trigger escalation if either condition is met
    if (condition1 || condition2) {
      Logger.log('Escalation triggered for ' + taskId + ' (' + cycleId + '). Condition1=' + condition1 + ', Condition2=' + condition2);

      var taskData = tmSheet.getRange(tmRowNum, 1, 1, tmSheet.getLastColumn()).getValues()[0];
      var success = sendLeadershipEscalation(taskData, alertRow, alertRowNum);

      if (success) {
        // Update Escalation Status in Alert Log
        alertSheet.getRange(alertRowNum, AL.ESC_STATUS).setValue('Escalated');

        // Update Task Master status to STATUS.ESCALATED
        tmSheet.getRange(tmRowNum, TM.STATUS).setValue(STATUS.ESCALATED);

        // Update Action Center status
        var acRowNum = findActionCenterRowSafe_(cycleId);
        if (acRowNum >= 2) {
          var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
          if (acSheet) {
            acSheet.getRange(acRowNum, AC.STATUS).setValue(STATUS.ESCALATED);
            if (acSheet.getLastColumn() >= AC.MGR_REVIEW) {
              acSheet.getRange(acRowNum, AC.MGR_REVIEW).setValue(STATUS.ESCALATED);
            }
          }
        }

        escalationsSent++;
      }
    }
  }

  Logger.log('processEscalations completed. Total escalations triggered: ' + escalationsSent);
  return escalationsSent;
}

/**
 * Sends a leadership escalation email for an unresolved task and records
 * the event into the Alert & Escalation Log.
 *
 * @param {Array} taskData Array of column values from Task Master.
 * @param {Array} alertRow Array of column values from the original Manager Alert row.
 * @param {number} alertRowNum The 1-based row number of the original Manager Alert.
 * @return {boolean} True if successfully sent/logged, false otherwise.
 */
function sendLeadershipEscalation(taskData, alertRow, alertRowNum) {
  // Step 1: Retrieve settings
  var settings = getAllSettings();
  var testModeSetting = settings['Test Mode'];
  var isTestMode = (testModeSetting === true || String(testModeSetting).toUpperCase() === 'TRUE');

  var emailEnabledSetting = settings['Email Sending Enabled'];
  var isEmailEnabled = (emailEnabledSetting === true || String(emailEnabledSetting).toUpperCase() === 'TRUE');

  var leadershipEmail = settings['Leadership Email'] || '';
  var testRecipient = settings['Test Email Recipient'] || '';

  // Step 2: Determine recipient
  var recipient = isTestMode ? testRecipient : leadershipEmail;

  // Step 3: Build subject
  var taskId = taskData[TM.TASK_ID - 1];
  var taskName = taskData[TM.NAME - 1];
  var cycleId = alertRow[AL.CYCLE_ID - 1];
  var employee = alertRow[AL.EMPLOYEE - 1];
  var subject = 'ESCALATION: ' + taskName + ' - Unresolved Overdue Task';

  // Step 4: Build HTML body
  var htmlBody = buildEscalationEmailHtml(taskData, alertRow);

  // Step 5: Send email or log
  var alertStatus = '';
  var errorNote = '';

  if (isEmailEnabled) {
    try {
      if (!recipient || recipient.trim() === '') {
        throw new Error('No recipient email configured for Leadership Escalation');
      }

      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        htmlBody: htmlBody
      });
      alertStatus = 'Sent';
      Logger.log('Leadership escalation email sent to: ' + recipient + ' for task ' + taskId);
    } catch (err) {
      Logger.log('Failed to send leadership escalation email: ' + err.message);
      alertStatus = 'Failed';
      errorNote = 'Email error: ' + err.message;
    }
  } else if (isTestMode) {
    alertStatus = 'Test Mode - Logged';
    Logger.log('[TEST MODE] Leadership Escalation Email generated:');
    Logger.log('To: ' + recipient);
    Logger.log('Subject: ' + subject);
  } else {
    alertStatus = 'Email Disabled';
    Logger.log('[EMAIL DISABLED] Leadership Escalation logged without sending: ' + subject);
  }

  // Step 6: Log event to Alert Log with eventType = 'Leadership Escalation'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alertSheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!alertSheet) {
    Logger.log('Alert Log sheet not found. Cannot log escalation.');
    return false;
  }

  var now = getNow();
  var eventId = generateId('ALERT', alertSheet, AL.EVENT_ID);
  var escalationNote = errorNote ? errorNote : ('Escalated from Manager Alert at row ' + alertRowNum);

  var newRowValues = [
    eventId,                   // AL.EVENT_ID (1)
    cycleId,                   // AL.CYCLE_ID (2)
    taskId,                    // AL.TASK_ID (3)
    employee,                  // AL.EMPLOYEE (4)
    'Leadership Escalation',   // AL.EVENT_TYPE (5)
    now,                       // AL.TIMESTAMP (6)
    recipient,                 // AL.RECIPIENT (7)
    subject,                   // AL.SUBJECT (8)
    alertStatus,               // AL.STATUS (9)
    false,                     // AL.MGR_ACK (10)
    '',                        // AL.ACK_BY (11)
    '',                        // AL.ACK_AT (12)
    'Escalated',               // AL.ESC_STATUS (13)
    escalationNote             // AL.NOTES (14)
  ];

  alertSheet.appendRow(newRowValues);

  var newRowNum = alertSheet.getLastRow();
  alertSheet.getRange(newRowNum, AL.MGR_ACK).insertCheckboxes();
  alertSheet.getRange(newRowNum, AL.MGR_ACK).setValue(false);

  // Step 7: Return success
  return (alertStatus !== 'Failed');
}

// ============================================================================
// 3. HTML EMAIL TEMPLATE BUILDERS
// ============================================================================

/**
 * Builds a clean, professional HTML email body for manager overdue alerts.
 * Uses inline styling and semantic table formatting for email client compatibility.
 *
 * @param {Array} taskData Task Master row data array.
 * @return {string} HTML string for the email body.
 */
function buildManagerEmailHtml(taskData) {
  var taskId = escapeHtml_(taskData[TM.TASK_ID - 1] || 'N/A');
  var taskName = escapeHtml_(taskData[TM.NAME - 1] || 'N/A');
  var employee = escapeHtml_(taskData[TM.EMPLOYEE - 1] || 'Unassigned');

  var rawDueDate = taskData[TM.CURRENT_DUE - 1];
  var dueDateStr = rawDueDate ? formatDateOnly(rawDueDate instanceof Date ? rawDueDate : new Date(rawDueDate)) : 'N/A';
  var dueTimeStr = escapeHtml_(taskData[TM.DUE_TIME - 1] || '17:00');

  var currentStatus = escapeHtml_(taskData[TM.STATUS - 1] || STATUS.OVERDUE);
  var proofType = escapeHtml_(taskData[TM.PROOF_TYPE - 1] || 'None specified');
  var cycleId = escapeHtml_(taskData[TM.CYCLE_ID - 1] || 'N/A');
  var detectedAtStr = formatDateTime(getNow());

  var sheetUrl = getSetting('Spreadsheet URL');
  if (!sheetUrl) {
    try {
      sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    } catch (e) {
      sheetUrl = '#';
    }
  }

  var html = '' +
    '<div style="font-family: Arial, Helvetica, sans-serif; max-width: 650px; margin: 0 auto; border: 1px solid #dcdcdc; border-radius: 8px; overflow: hidden; background-color: #ffffff;">' +
      // Blue Header Banner
      '<div style="background-color: #1a73e8; color: #ffffff; padding: 22px 26px;">' +
        '<h2 style="margin: 0 0 6px 0; font-size: 21px; font-weight: bold; color: #ffffff; letter-spacing: -0.3px;">Overdue Task Alert</h2>' +
        '<p style="margin: 0; font-size: 13px; color: #e8f0fe;">Action required: A scheduled recurring task has passed its completion deadline.</p>' +
      '</div>' +

      // Main Content Container
      '<div style="padding: 24px 26px;">' +
        '<p style="margin: 0 0 16px 0; font-size: 14px; line-height: 1.5; color: #333333;">' +
          'Hello Manager,<br><br>' +
          'The following recurring operational task is currently <strong>overdue</strong> and requires your immediate attention and follow-up.' +
        '</p>' +

        // Task Detail Table
        '<table style="width: 100%; border-collapse: collapse; margin: 18px 0; font-size: 13px;">' +
          '<tbody>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; width: 34%; color: #555555;">Task Name</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #111111; font-weight: bold;">' + taskName + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Task ID</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + taskId + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Assigned Employee</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + employee + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Due Date & Time</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #d93025; font-weight: bold;">' + dueDateStr + ' at ' + dueTimeStr + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Current Status</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #d93025; font-weight: bold;">' + currentStatus + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Required Proof</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + proofType + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Cycle ID</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + cycleId + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Time Detected</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + detectedAtStr + '</td>' +
            '</tr>' +
          '</tbody>' +
        '</table>' +

        // Acknowledgement Instructions Callout Box
        '<div style="background-color: #e8f0fe; border-left: 4px solid #1a73e8; padding: 14px 18px; margin: 22px 0; border-radius: 0 4px 4px 0;">' +
          '<p style="margin: 0; font-size: 13px; color: #1a73e8; font-weight: bold;">Manager Acknowledgement Instructions:</p>' +
          '<p style="margin: 5px 0 0 0; font-size: 13px; color: #333333; line-height: 1.4;">' +
            'Please open the Google Sheet and check the acknowledgement box in the Alert & Escalation Log.' +
          '</p>' +
        '</div>' +

        // CTA Link Button
        '<div style="text-align: center; margin: 24px 0 12px 0;">' +
          '<a href="' + sheetUrl + '" target="_blank" style="background-color: #1a73e8; color: #ffffff; text-decoration: none; padding: 11px 22px; border-radius: 4px; font-weight: bold; font-size: 14px; display: inline-block;">' +
            'Open Google Sheet Tracker' +
          '</a>' +
        '</div>' +
      '</div>' +

      // Footer
      '<div style="background-color: #f8f9fa; padding: 14px 26px; border-top: 1px solid #e0e0e0; text-align: center;">' +
        '<p style="margin: 0; font-size: 12px; color: #777777;">' +
          'Allo Innoware — Recurring Task Tracker & Operations Management' +
        '</p>' +
      '</div>' +
    '</div>';

  return html;
}

/**
 * Builds a high-priority HTML email for leadership escalation.
 * Features an alert banner, urgency indicators, manager notification history,
 * acknowledgement state, and resolution time tracking.
 *
 * @param {Array} taskData Task Master row data array.
 * @param {Array} alertRow Alert & Escalation Log row array for the original manager alert.
 * @return {string} HTML string for the escalation email body.
 */
function buildEscalationEmailHtml(taskData, alertRow) {
  var taskId = escapeHtml_(taskData[TM.TASK_ID - 1] || 'N/A');
  var taskName = escapeHtml_(taskData[TM.NAME - 1] || 'N/A');
  var employee = escapeHtml_(taskData[TM.EMPLOYEE - 1] || 'Unassigned');

  var rawDueDate = taskData[TM.CURRENT_DUE - 1];
  var dueDateStr = rawDueDate ? formatDateOnly(rawDueDate instanceof Date ? rawDueDate : new Date(rawDueDate)) : 'N/A';
  var dueTimeStr = escapeHtml_(taskData[TM.DUE_TIME - 1] || '17:00');

  var proofType = escapeHtml_(taskData[TM.PROOF_TYPE - 1] || 'None specified');
  var cycleId = escapeHtml_(alertRow[AL.CYCLE_ID - 1] || 'N/A');
  var detectedAtStr = formatDateTime(getNow());

  // Manager notification details from original alert row
  var rawAlertTime = alertRow[AL.TIMESTAMP - 1];
  var alertTimeStr = rawAlertTime ? formatDateTime(rawAlertTime instanceof Date ? rawAlertTime : new Date(rawAlertTime)) : 'N/A';

  // Manager acknowledgement status
  var rawMgrAck = alertRow[AL.MGR_ACK - 1];
  var isAck = (rawMgrAck === true || String(rawMgrAck).toUpperCase() === 'TRUE');
  var ackStatusHtml = '';

  if (isAck) {
    var ackAt = alertRow[AL.ACK_AT - 1];
    var ackAtStr = ackAt ? formatDateTime(ackAt instanceof Date ? ackAt : new Date(ackAt)) : '';
    var ackBy = escapeHtml_(alertRow[AL.ACK_BY - 1] || 'Manager');
    ackStatusHtml = '<span style="color: #0f9d58; font-weight: bold;">Yes</span> (' + ackAtStr + ' by ' + ackBy + ')';
  } else {
    ackStatusHtml = '<span style="color: #d93025; font-weight: bold;">No (Unacknowledged)</span>';
  }

  // Duration unresolved calculation
  var hoursElapsed = 0;
  if (rawAlertTime) {
    var alertDate = (rawAlertTime instanceof Date) ? rawAlertTime : new Date(rawAlertTime);
    hoursElapsed = hoursDiff(alertDate, getNow());
  }
  var hoursElapsedStr = hoursElapsed > 0 ? (hoursElapsed.toFixed(1) + ' hours') : 'N/A';

  var alertNotes = escapeHtml_(alertRow[AL.NOTES - 1] || 'None recorded');

  var sheetUrl = getSetting('Spreadsheet URL');
  if (!sheetUrl) {
    try {
      sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    } catch (e) {
      sheetUrl = '#';
    }
  }

  var html = '' +
    '<div style="font-family: Arial, Helvetica, sans-serif; max-width: 650px; margin: 0 auto; border: 1px solid #dcdcdc; border-radius: 8px; overflow: hidden; background-color: #ffffff;">' +
      // Red Urgent Header Banner
      '<div style="background-color: #d93025; color: #ffffff; padding: 22px 26px;">' +
        '<h2 style="margin: 0 0 6px 0; font-size: 21px; font-weight: bold; color: #ffffff; letter-spacing: -0.3px;">LEADERSHIP ESCALATION</h2>' +
        '<p style="margin: 0; font-size: 13px; color: #fce8e6;">Unresolved Overdue Operational Task Requiring Leadership Attention</p>' +
      '</div>' +

      // Main Content Container
      '<div style="padding: 24px 26px;">' +
        // Urgency Alert Callout Box
        '<div style="background-color: #fce8e6; border-left: 4px solid #d93025; padding: 14px 18px; margin-bottom: 20px; border-radius: 0 4px 4px 0;">' +
          '<p style="margin: 0; font-size: 14px; color: #c5221f; font-weight: bold;">Urgent Attention Required:</p>' +
          '<p style="margin: 5px 0 0 0; font-size: 13px; color: #333333; line-height: 1.4;">' +
            'This task has exceeded the configured resolution window.' +
          '</p>' +
        '</div>' +

        '<p style="margin: 0 0 16px 0; font-size: 14px; line-height: 1.5; color: #333333;">' +
          'Dear Leadership Team,<br><br>' +
          'The operational task described below has breached SLA resolution timelines and has not been marked as completed despite manager notification.' +
        '</p>' +

        // Escalation Details Table
        '<table style="width: 100%; border-collapse: collapse; margin: 18px 0; font-size: 13px;">' +
          '<tbody>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; width: 36%; color: #555555;">Task Name</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #111111; font-weight: bold;">' + taskName + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Task ID</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + taskId + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Assigned Employee</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + employee + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Original Due Date & Time</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #d93025; font-weight: bold;">' + dueDateStr + ' at ' + dueTimeStr + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Current Status</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #d93025; font-weight: bold;">' + STATUS.ESCALATED + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Cycle ID</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + cycleId + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Required Proof</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + proofType + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Manager Notified</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #333333;">' + alertTimeStr + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Manager Acknowledged</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0;">' + ackStatusHtml + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Time Unresolved Since Alert</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #d93025; font-weight: bold;">' + hoursElapsedStr + '</td>' +
            '</tr>' +
            '<tr style="background-color: #f8f9fa;">' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; font-weight: bold; color: #555555;">Alert Log Notes</td>' +
              '<td style="padding: 10px 14px; border: 1px solid #e0e0e0; color: #555555; font-style: italic;">' + alertNotes + '</td>' +
            '</tr>' +
          '</tbody>' +
        '</table>' +

        // CTA Button
        '<div style="text-align: center; margin: 24px 0 12px 0;">' +
          '<a href="' + sheetUrl + '" target="_blank" style="background-color: #d93025; color: #ffffff; text-decoration: none; padding: 11px 22px; border-radius: 4px; font-weight: bold; font-size: 14px; display: inline-block;">' +
            'Review in Tracker Sheet' +
          '</a>' +
        '</div>' +
      '</div>' +

      // Footer
      '<div style="background-color: #f8f9fa; padding: 14px 26px; border-top: 1px solid #e0e0e0; text-align: center;">' +
        '<p style="margin: 0; font-size: 12px; color: #777777;">' +
          'Allo Innoware — Recurring Task Tracker & Operations Management' +
        '</p>' +
      '</div>' +
    '</div>';

  return html;
}

// ============================================================================
// 4. MANAGER ACKNOWLEDGEMENT HANDLER
// ============================================================================

/**
 * Handles the onEdit trigger action when a manager checks the
 * 'Manager Acknowledged' checkbox (column AL.MGR_ACK) in the Alert & Escalation Log.
 *
 * Records the user identity and timestamp of acknowledgement, and alerts the user.
 *
 * @param {number} row The 1-based row number edited in the Alert Log.
 */
function handleAlertAcknowledgement(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET.ALERT_LOG);
  if (!sheet || row < 2) {
    return;
  }

  // Step 1: Read the edited row data
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Step 2: Verify that the Event Type is 'Manager Alert'
  var eventType = rowData[AL.EVENT_TYPE - 1];
  if (eventType !== 'Manager Alert') {
    Logger.log('Row ' + row + ' is not a Manager Alert event. Skipping acknowledgement.');
    return;
  }

  // Verify that the checkbox is actually checked (TRUE)
  var isChecked = rowData[AL.MGR_ACK - 1];
  if (isChecked !== true && String(isChecked).toUpperCase() !== 'TRUE') {
    Logger.log('Acknowledgement checkbox at row ' + row + ' is not checked. Skipping.');
    return;
  }

  // Step 3: Identify the acknowledger email, fallback to 'Manager'
  var acknowledger = '';
  try {
    var activeUser = Session.getActiveUser();
    if (activeUser && activeUser.getEmail()) {
      acknowledger = activeUser.getEmail();
    }
  } catch (err) {
    Logger.log('Unable to retrieve Session user email: ' + err.message);
  }

  if (!acknowledger || acknowledger.trim() === '') {
    acknowledger = 'Manager';
  }

  // Step 4 & 5: Record acknowledger email and timestamp
  var now = getNow();
  sheet.getRange(row, AL.ACK_BY).setValue(acknowledger);
  sheet.getRange(row, AL.ACK_AT).setValue(now);

  Logger.log('Alert acknowledged for row ' + row + ' by ' + acknowledger + ' at ' + formatDateTime(now));

  // Step 6: Show confirmation alert in UI if interactive session is available
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      var taskId = rowData[AL.TASK_ID - 1] || 'Task';
      var cycleId = rowData[AL.CYCLE_ID - 1] || '';
      ui.alert(
        'Alert Acknowledged',
        'Acknowledgement recorded for ' + taskId + ' (' + cycleId + ') by ' + acknowledger + ' at ' + formatDateTime(now) + '.',
        ui.ButtonSet.OK
      );
    }
  } catch (uiErr) {
    // In automated tests or background triggers, getUi() may not be available
    Logger.log('Interactive UI alert skipped: ' + uiErr.message);
  }
}

// ============================================================================
// 5. INTERNAL HELPER FUNCTIONS
// ============================================================================

/**
 * Searches Task Master for a matching Task ID.
 * Uses global findTaskRow if available, with a local fallback search.
 *
 * @param {string} taskId Task identifier (e.g., 'TASK-001').
 * @return {number} 1-based row number or -1 if not found.
 * @private
 */
function findTaskRowSafe_(taskId) {
  if (typeof findTaskRow === 'function') {
    return findTaskRow(taskId);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tmSheet = ss.getSheetByName(SHEET.TASK_MASTER);
  if (!tmSheet) return -1;
  var lastRow = tmSheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = tmSheet.getRange(2, TM.TASK_ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(taskId).trim()) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Searches Action Center for a matching Cycle ID.
 * Uses global findActionCenterRow if available, with a local fallback search.
 *
 * @param {string} cycleId Cycle identifier (e.g., 'TASK-001-C001').
 * @return {number} 1-based row number or -1 if not found.
 * @private
 */
function findActionCenterRowSafe_(cycleId) {
  if (typeof findActionCenterRow === 'function') {
    return findActionCenterRow(cycleId);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var acSheet = ss.getSheetByName(SHEET.ACTION_CENTER);
  if (!acSheet) return -1;
  var lastRow = acSheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = acSheet.getRange(2, AC.CYCLE_ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(cycleId).trim()) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Escapes unsafe HTML characters to prevent XSS or layout glitches in emails.
 *
 * @param {*} input Input value to sanitize.
 * @return {string} Sanitized string.
 * @private
 */
function escapeHtml_(input) {
  if (input === null || input === undefined) {
    return '';
  }
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
