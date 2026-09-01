// ==========================================
// RIS STATUS MENU
// ==========================================

const RIS_STATUS_CONFIG = {
  menuName: 'RIS Status',
  entriesSheetName: 'RIS Entries',
  medicalSupplyEmail: 'msdmedicalsuppliesdepot@gmail.com',
  statusCcEmails: ['lyn4logistics@gmail.com'],
  alternateUpdaterCcEmail: '',
  reminderHour: 17,
  reminderTriggerFunction: 'risStatusSendDailyReminders',
  headerRow: 1,
  defaultStatus: 'Request Submitted',
  requestStatuses: ['Approved', 'Rescheduled'],
  deliveryStatuses: ['RIS NOT Signed', 'RIS Discrepancy', 'RIS Signed Completed'],
  timeOptions: ['AM', 'PM', 'Whole Day']
};

const RIS_STATUS_FIELD_ALIASES = {
  recordId: ['Record ID', 'RecordId'],
  risNo: ['RIS Number', 'RIS No.', 'RIS No'],
  requestorEmail: ['Requestor Email', 'Requester Email', 'Email Address', 'Email'],
  requestorName: ['Requested By', 'Requestor Name', 'Requester Name', 'Name'],
  requestorProgram: ['Requestor Program', 'Program', 'Program / End User'],
  requestedDate: ['Date Requested', 'Date', 'RIS Date', 'Requested Delivery Date'],
  deliveryLocation: ['Delivery Location / Office', 'Delivery Location', 'Office'],
  risStatus: ['RIS Status', 'Status'],
  requestStatus: ['Request Status'],
  requestDate: ['Request Date'],
  requestTime: ['Request Time'],
  requestStatusUpdatedBy: ['Request Status Updated By'],
  requestStatusUpdatedAt: ['Request Status Updated At'],
  requestStatusEmailSentAt: ['Request Status Email Sent Date', 'Request Status Email Sent At'],
  requestStatusEmailSentBy: ['Request Status Email Sent By'],
  deliveryStatus: ['Delivery Status'],
  deliveryDate: ['Delivery Date'],
  deliveryTime: ['Delivery Time'],
  deliveryStatusUpdatedBy: ['Delivery Status Updated By'],
  deliveryStatusUpdatedAt: ['Delivery Status Updated At'],
  deliveryStatusEmailSentAt: ['Delivery Status Email Sent Date', 'Delivery Status Email Sent At'],
  deliveryStatusEmailSentBy: ['Delivery Status Email Sent By'],
  statusRemarks: ['Status Remarks', 'Remarks']
};

const RIS_STATUS_DEFAULT_HEADERS = {
  risStatus: 'RIS Status',
  requestStatus: 'Request Status',
  requestDate: 'Request Date',
  requestTime: 'Request Time',
  requestStatusUpdatedBy: 'Request Status Updated By',
  requestStatusUpdatedAt: 'Request Status Updated At',
  requestStatusEmailSentAt: 'Request Status Email Sent Date',
  requestStatusEmailSentBy: 'Request Status Email Sent By',
  deliveryStatus: 'Delivery Status',
  deliveryDate: 'Delivery Date',
  deliveryTime: 'Delivery Time',
  deliveryStatusUpdatedBy: 'Delivery Status Updated By',
  deliveryStatusUpdatedAt: 'Delivery Status Updated At',
  deliveryStatusEmailSentAt: 'Delivery Status Email Sent Date',
  deliveryStatusEmailSentBy: 'Delivery Status Email Sent By',
  statusRemarks: 'Status Remarks'
};

function risStatusAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(RIS_STATUS_CONFIG.menuName)
    .addItem('Show Current Status', 'risStatusShowCurrentPrompt')
    .addSeparator()
    .addItem('Request Status', 'risStatusShowRequestDialog')
    .addItem('Delivery Status', 'risStatusShowDeliveryDialog')
    .addSeparator()
    .addItem('Install 5:00 PM Status Reminders', 'risStatusInstallDailyReminderTrigger')
    .addItem('Run Status Reminder Check Now', 'risStatusRunReminderCheckNow')
    .addToUi();
}

function risStatusShowCurrentPrompt() {
  const ui = SpreadsheetApp.getUi();
  const risNo = risStatusPromptRisNo_(ui);
  if (!risNo) return;

  try {
    const context = risStatusGetTargetContext_({ targetMode: 'risNo', risNo: risNo });
    ui.alert(
      'Current RIS Status',
      [
        'RIS No.: ' + context.risNo,
        'Overall Status: ' + risStatusCurrentOverallStatus_(context),
        'Request Status: ' + (context.requestStatus || ''),
        'Request Schedule: ' + risStatusJoinSchedule_(context.requestDate, context.requestTime),
        'Delivery Status: ' + (context.deliveryStatus || ''),
        'Delivery Schedule: ' + risStatusJoinSchedule_(context.deliveryDate, context.deliveryTime),
        'Program: ' + (context.requestorProgram || ''),
        'Delivery Location: ' + (context.deliveryLocation || '')
      ].join('\n'),
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Could not load RIS status:\n\n' + risCoreErrorMessage_(error));
  }
}

function risStatusShowRequestDialog() {
  risStatusShowDialog_('request');
}

function risStatusShowDeliveryDialog() {
  risStatusShowDialog_('delivery');
}

function risStatusShowDialog_(mode) {
  const html = HtmlService.createHtmlOutput(risStatusBuildDialogHtml_(mode))
    .setWidth(500)
    .setHeight(mode === 'request' ? 690 : 720);
  SpreadsheetApp.getUi().showModalDialog(html, mode === 'request' ? 'Request Status' : 'Delivery Status');
}

function risStatusProcessRequest(form) {
  const status = risStatusRequireOneOf_(form.status, RIS_STATUS_CONFIG.requestStatuses, 'request status');
  const scheduleDate = risStatusParseDate_(form.date);
  const scheduleTime = risStatusRequireOneOf_(form.time, RIS_STATUS_CONFIG.timeOptions, 'time');
  const updater = risStatusResolveUpdater_(form);
  const context = risStatusGetTargetContext_(form);

  risStatusValidateDateNotBeforeSheetDate_(scheduleDate, context, 'Request Date');
  if (!risStatusIsValidEmail_(context.requestorEmail)) {
    throw new Error('Missing or invalid requestor email for row ' + context.row + '.');
  }

  risStatusSetMany_(context.sheet, context.row, {
    risStatus: risStatusOverallStatusForRequest_(status),
    requestStatus: status,
    requestDate: scheduleDate,
    requestTime: scheduleTime,
    requestStatusUpdatedBy: risStatusUpdaterLabel_(updater),
    requestStatusUpdatedAt: new Date()
  });

  const dateText = risStatusFormatLongDate_(scheduleDate);
  const email = status === 'Approved'
    ? risStatusBuildApprovedRequestEmail_(context, dateText, scheduleTime)
    : risStatusBuildRescheduledRequestEmail_(context, dateText, scheduleTime);

  risStatusSendEmail_(email, [updater.email]);
  risStatusSetMany_(context.sheet, context.row, {
    requestStatusEmailSentAt: new Date(),
    requestStatusEmailSentBy: risStatusUpdaterLabel_(updater)
  });

  SpreadsheetApp.getActive().toast('Request Status email sent for RIS ' + context.risNo + '.', RIS_STATUS_CONFIG.menuName, 5);
  return 'Request Status saved and email sent for RIS ' + context.risNo + '.';
}

function risStatusProcessDelivery(form) {
  const status = risStatusRequireOneOf_(form.status, RIS_STATUS_CONFIG.deliveryStatuses, 'delivery status');
  const deliveredDate = risStatusParseDate_(form.date);
  const deliveredTime = risStatusRequireOneOf_(form.time, RIS_STATUS_CONFIG.timeOptions, 'time');
  const updater = risStatusResolveUpdater_(form);
  const context = risStatusGetTargetContext_(form);

  risStatusValidateDateNotBeforeSheetDate_(deliveredDate, context, 'Delivery Date');
  risStatusSetMany_(context.sheet, context.row, {
    risStatus: status,
    deliveryStatus: status,
    deliveryDate: deliveredDate,
    deliveryTime: deliveredTime,
    deliveryStatusUpdatedBy: risStatusUpdaterLabel_(updater),
    deliveryStatusUpdatedAt: new Date()
  });

  const email = risStatusBuildDeliveryEmail_(status, context, risStatusFormatLongDate_(deliveredDate), deliveredTime);
  risStatusSendEmail_(email, [updater.email]);
  risStatusSetMany_(context.sheet, context.row, {
    deliveryStatusEmailSentAt: new Date(),
    deliveryStatusEmailSentBy: risStatusUpdaterLabel_(updater)
  });

  const reminderResult = risStatusEnsureDailyReminderTriggerSafely_(status);
  const reminderWarning = reminderResult.warning ? ' Reminder trigger warning: ' + reminderResult.warning : '';

  SpreadsheetApp.getActive().toast('Delivery Status email sent for RIS ' + context.risNo + '.', RIS_STATUS_CONFIG.menuName, 5);
  return 'Delivery Status saved and email sent for RIS ' + context.risNo + '.' + reminderWarning;
}

function risStatusGetBaseDate(form) {
  const context = risStatusGetTargetContext_(form);
  const sheetDate = risStatusGetSheetBaseDate_(context);
  return {
    risNo: context.risNo,
    baseDate: risStatusFormatInputDate_(sheetDate),
    displayDate: risStatusFormatLongDate_(sheetDate)
  };
}

function risStatusGetTargetContext_(form) {
  const ss = risCoreGetTransactionsSpreadsheet_();
  risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);
  const sheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName || RIS_STATUS_CONFIG.entriesSheetName);
  const row = risStatusResolveTargetRow_(ss, sheet, form || {});
  return risStatusBuildRowContext_(ss, sheet, row, String((form && form.risNo) || '').trim());
}

function risStatusBuildRowContext_(ss, sheet, row, fallbackRisNo, rowValues, displayValues, headerMap) {
  rowValues = rowValues || sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  displayValues = displayValues || sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  headerMap = headerMap || risStatusGetHeaderMap_(sheet);

  const requestorProgram = risStatusGetDisplayField_(displayValues, headerMap, 'requestorProgram');
  const requestorName = risStatusGetDisplayField_(displayValues, headerMap, 'requestorName') || requestorProgram || 'Requestor';

  return {
    ss: ss,
    sheet: sheet,
    row: row,
    risNo: risStatusGetDisplayField_(displayValues, headerMap, 'risNo') ||
      fallbackRisNo ||
      risStatusGetDisplayField_(displayValues, headerMap, 'recordId') ||
      'the selected RIS',
    requestorName: requestorName,
    requestorEmail: risStatusGetDisplayField_(displayValues, headerMap, 'requestorEmail'),
    requestorProgram: requestorProgram,
    deliveryLocation: risStatusGetDisplayField_(displayValues, headerMap, 'deliveryLocation'),
    requestedDate: risStatusGetDisplayField_(displayValues, headerMap, 'requestedDate'),
    requestedDateValue: risStatusGetRawField_(rowValues, headerMap, 'requestedDate'),
    risStatus: risStatusGetDisplayField_(displayValues, headerMap, 'risStatus'),
    requestStatus: risStatusGetDisplayField_(displayValues, headerMap, 'requestStatus'),
    requestDate: risStatusGetDisplayField_(displayValues, headerMap, 'requestDate'),
    requestTime: risStatusGetDisplayField_(displayValues, headerMap, 'requestTime'),
    deliveryStatus: risStatusGetDisplayField_(displayValues, headerMap, 'deliveryStatus'),
    deliveryDate: risStatusGetDisplayField_(displayValues, headerMap, 'deliveryDate'),
    deliveryTime: risStatusGetDisplayField_(displayValues, headerMap, 'deliveryTime'),
    deliveryStatusUpdatedBy: risStatusGetDisplayField_(displayValues, headerMap, 'deliveryStatusUpdatedBy')
  };
}

function risStatusResolveTargetRow_(ss, entriesSheet, form) {
  if (String(form.targetMode || 'selectedRow') === 'selectedRow') {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = activeSpreadsheet ? activeSpreadsheet.getActiveSheet() : null;
    const activeRange = activeSheet ? activeSheet.getActiveRange() : null;

    if (!activeSpreadsheet || activeSpreadsheet.getId() !== ss.getId()) {
      throw new Error('Please open the transactions spreadsheet and select a RIS row.');
    }
    if (!activeSheet || !activeRange || activeRange.getRow() <= RIS_STATUS_CONFIG.headerRow) {
      throw new Error('Please select a data row, not the header row.');
    }
    if (activeSheet.getName() === entriesSheet.getName()) return activeRange.getRow();

    const selectedRisNo = risStatusGetRowRisKey_(activeSheet, activeRange.getRow());
    if (selectedRisNo) return risStatusFindEntryRow_(entriesSheet, selectedRisNo);
    throw new Error('Please select a row in RIS Entries or a row with a RIS Number.');
  }

  const lookup = String(form.risNo || '').trim();
  if (!lookup) throw new Error('Please enter the RIS No.');
  return risStatusFindEntryRow_(entriesSheet, lookup);
}

function risStatusFindEntryRow_(entriesSheet, lookup) {
  const headerMap = risStatusGetHeaderMap_(entriesSheet);
  const risNoColumn = headerMap.risNo;
  const recordIdColumn = headerMap.recordId;
  const lastRow = entriesSheet.getLastRow();

  if (lastRow <= RIS_STATUS_CONFIG.headerRow) throw new Error('No RIS records were found below the header row.');
  if (!risNoColumn && !recordIdColumn) throw new Error('RIS Entries must include a RIS Number column.');

  const values = entriesSheet
    .getRange(RIS_STATUS_CONFIG.headerRow + 1, 1, lastRow - RIS_STATUS_CONFIG.headerRow, entriesSheet.getLastColumn())
    .getDisplayValues();
  const lookupKey = risStatusNormalizeLookup_(lookup);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowRisNo = risNoColumn ? row[risNoColumn - 1] : '';
    const rowRecordId = recordIdColumn ? row[recordIdColumn - 1] : '';
    if (risStatusNormalizeLookup_(rowRisNo) === lookupKey || risStatusNormalizeLookup_(rowRecordId) === lookupKey) {
      return RIS_STATUS_CONFIG.headerRow + 1 + i;
    }
  }

  throw new Error('No RIS Entries row found for: ' + lookup);
}

function risStatusGetHeaderMap_(sheet) {
  const columns = {};
  if (!sheet || sheet.getLastRow() < RIS_STATUS_CONFIG.headerRow || sheet.getLastColumn() < 1) return columns;

  const headers = sheet.getRange(RIS_STATUS_CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  Object.keys(RIS_STATUS_FIELD_ALIASES).forEach(function(key) {
    columns[key] = risFindHeaderColumn_(headers, RIS_STATUS_FIELD_ALIASES[key]);
  });
  return columns;
}

function risStatusSetMany_(sheet, row, valuesByField) {
  Object.keys(valuesByField).forEach(function(key) {
    const column = risStatusEnsureColumn_(sheet, key);
    sheet.getRange(row, column).setValue(valuesByField[key]);
  });
}

function risStatusEnsureColumn_(sheet, key) {
  const headerMap = risStatusGetHeaderMap_(sheet);
  if (headerMap[key]) return headerMap[key];

  const headerName = RIS_STATUS_DEFAULT_HEADERS[key];
  if (!headerName) throw new Error('No default header configured for: ' + key);

  const nextColumn = sheet.getLastColumn() + 1;
  sheet.getRange(RIS_STATUS_CONFIG.headerRow, nextColumn).setValue(headerName);
  return nextColumn;
}

function risStatusGetDisplayField_(displayValues, headerMap, key) {
  const column = headerMap[key] || 0;
  return column ? String(displayValues[column - 1] || '').trim() : '';
}

function risStatusGetRawField_(rowValues, headerMap, key) {
  const column = headerMap[key] || 0;
  return column ? rowValues[column - 1] : '';
}

function risStatusGetRowRisKey_(sheet, rowNumber) {
  const headers = sheet.getRange(RIS_STATUS_CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const risCol = risFindHeaderColumn_(headers, RIS_STATUS_FIELD_ALIASES.risNo);
  const recordCol = risFindHeaderColumn_(headers, RIS_STATUS_FIELD_ALIASES.recordId);
  if (!risCol && !recordCol) return '';

  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return risCleanText_(risCol ? row[risCol - 1] : '') || risCleanText_(recordCol ? row[recordCol - 1] : '');
}

function risStatusGetSelectedRisKey_() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet && sheet.getActiveRange();
    if (!sheet || !range || range.getRow() <= RIS_STATUS_CONFIG.headerRow || sheet.getLastColumn() < 1) return '';
    return risStatusGetRowRisKey_(sheet, range.getRow());
  } catch (error) {
    return '';
  }
}

function risStatusPromptRisNo_(ui) {
  const selectedRisNo = risStatusGetSelectedRisKey_();
  const helper = selectedRisNo
    ? '\n\nSelected row detected: ' + selectedRisNo + '\nLeave the prompt blank to use the selected row.'
    : '';
  const response = ui.prompt('Find RIS Record', 'Enter RIS No. or Record ID:' + helper, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return '';

  const entered = response.getResponseText().trim();
  if (entered) return entered;
  if (selectedRisNo) return selectedRisNo;

  ui.alert('Please enter the RIS No.');
  return '';
}

function risStatusCurrentOverallStatus_(context) {
  return context.risStatus || context.deliveryStatus || context.requestStatus || RIS_STATUS_CONFIG.defaultStatus;
}

function risStatusOverallStatusForRequest_(status) {
  if (status === 'Approved') return 'Request Approved';
  if (status === 'Rescheduled' || status === 'Reschedule') return 'Request Rescheduled';
  return status || RIS_STATUS_CONFIG.defaultStatus;
}

function risStatusJoinSchedule_(dateText, timeText) {
  if (!dateText && !timeText) return '';
  if (!dateText) return timeText;
  if (!timeText) return dateText;
  return dateText + ', ' + timeText;
}

function risStatusResolveUpdater_(form) {
  const username = risCleanText_(form && form.username);
  if (!username) throw new Error('Enter your admin username.');

  const user = risCoreValidateActiveUser_(username);
  const email = risStatusFindAdminEmail_(username, user) ||
    risStatusGetActiveUserEmail_() ||
    RIS_STATUS_CONFIG.alternateUpdaterCcEmail;

  return {
    name: user.fullName || user.username || username,
    username: user.username || username,
    email: risStatusIsValidEmail_(email) ? risStatusNormalizeEmail_(email) : ''
  };
}

function risStatusUpdaterLabel_(updater) {
  return updater.email || updater.name || updater.username || 'Unavailable';
}

function risStatusFindAdminEmail_(username, user) {
  const ss = risCoreGetTransactionsSpreadsheet_();
  const sheet = ss.getSheetByName(RIS_CONFIG.adminEmailsSheetName) || ss.getSheetByName(RIS_CONFIG.usersSheetName);
  if (!sheet || sheet.getLastRow() < 2) return '';

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const nameCol = risFindHeaderColumn_(headers, ['Name', 'Full Name']);
  const usernameCol = risFindHeaderColumn_(headers, ['Username', 'User Name']);
  const recipientCol = risFindHeaderColumn_(headers, ['Recipients', 'Receipients', 'Recipient', 'Email', 'Email Address', 'TO']);
  const senderCol = risFindHeaderColumn_(headers, ['Sender', 'From']);
  const wanted = [username, user && user.username, user && user.fullName].map(risCoreNormalizeText_).filter(String);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowKeys = [
      usernameCol ? row[usernameCol - 1] : '',
      nameCol ? row[nameCol - 1] : '',
      recipientCol ? row[recipientCol - 1] : '',
      senderCol ? row[senderCol - 1] : ''
    ].map(risCoreNormalizeText_);

    if (rowKeys.some(function(value) { return wanted.indexOf(value) !== -1; })) {
      return risStatusFirstEmail_([
        recipientCol ? row[recipientCol - 1] : '',
        senderCol ? row[senderCol - 1] : '',
        usernameCol ? row[usernameCol - 1] : ''
      ]);
    }
  }

  return '';
}

function risStatusFirstEmail_(values) {
  for (let i = 0; i < values.length; i++) {
    const emails = risStatusNormalizeEmailList_(values[i]);
    if (emails.length > 0) return emails[0];
  }
  return '';
}

function risStatusInstallDailyReminderTrigger() {
  const trigger = risStatusEnsureDailyReminderTrigger_();
  SpreadsheetApp.getUi().alert(
    'Daily status reminders are installed.\n\n' +
    'Reminder time: around 5:00 PM, ' + Session.getScriptTimeZone() + '\n' +
    'Trigger function: ' + trigger.getHandlerFunction()
  );
}

function risStatusRunReminderCheckNow() {
  const result = risStatusSendDailyReminders();
  const warningText = result.warnings.length ? '\n\nWarnings:\n' + result.warnings.join('\n') : '';
  SpreadsheetApp.getUi().alert(
    'Status reminder check complete.\n\n' +
    'Emails sent: ' + result.sent + '\n' +
    'Rows checked: ' + result.checked +
    warningText
  );
}

function risStatusSendDailyReminders() {
  const ss = risCoreGetTransactionsSpreadsheet_();
  risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);
  const sheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName || RIS_STATUS_CONFIG.entriesSheetName);
  const headerMap = risStatusGetHeaderMap_(sheet);
  const deliveryStatusColumn = headerMap.deliveryStatus;
  const result = {
    checked: Math.max(sheet.getLastRow() - RIS_STATUS_CONFIG.headerRow, 0),
    sent: 0,
    warnings: []
  };

  if (!deliveryStatusColumn || sheet.getLastRow() <= RIS_STATUS_CONFIG.headerRow) return result;

  const rowCount = sheet.getLastRow() - RIS_STATUS_CONFIG.headerRow;
  const rowValues = sheet.getRange(RIS_STATUS_CONFIG.headerRow + 1, 1, rowCount, sheet.getLastColumn()).getValues();
  const displayValues = sheet.getRange(RIS_STATUS_CONFIG.headerRow + 1, 1, rowCount, sheet.getLastColumn()).getDisplayValues();

  for (let i = 0; i < rowCount; i++) {
    const rowNumber = RIS_STATUS_CONFIG.headerRow + 1 + i;
    const status = String(displayValues[i][deliveryStatusColumn - 1] || '').trim();
    if (!risStatusIsSameStatus_(status, 'RIS NOT Signed') && !risStatusIsSameStatus_(status, 'RIS Discrepancy')) continue;

    try {
      const context = risStatusBuildRowContext_(ss, sheet, rowNumber, '', rowValues[i], displayValues[i], headerMap);
      risStatusSendEmail_(risStatusBuildDailyReminderEmail_(status, context), [context.deliveryStatusUpdatedBy]);
      result.sent++;
    } catch (error) {
      result.warnings.push('Row ' + rowNumber + ': ' + risCoreErrorMessage_(error));
    }
  }

  return result;
}

function risStatusEnsureDailyReminderTriggerSafely_(status) {
  if (!risStatusIsSameStatus_(status, 'RIS NOT Signed') && !risStatusIsSameStatus_(status, 'RIS Discrepancy')) {
    return { installed: false, skipped: true };
  }

  try {
    risStatusEnsureDailyReminderTrigger_();
    return { installed: true };
  } catch (error) {
    return { installed: false, warning: risCoreErrorMessage_(error) };
  }
}

function risStatusEnsureDailyReminderTrigger_() {
  const existingTrigger = risStatusGetDailyReminderTrigger_();
  if (existingTrigger) return existingTrigger;

  return ScriptApp.newTrigger(RIS_STATUS_CONFIG.reminderTriggerFunction)
    .timeBased()
    .everyDays(1)
    .atHour(RIS_STATUS_CONFIG.reminderHour)
    .nearMinute(0)
    .inTimezone(Session.getScriptTimeZone())
    .create();
}

function risStatusGetDailyReminderTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === RIS_STATUS_CONFIG.reminderTriggerFunction) return triggers[i];
  }
  return null;
}

function risStatusSendEmail_(email, extraCcEmails) {
  const options = {
    to: email.to,
    subject: email.subject,
    body: email.body,
    htmlBody: email.htmlBody
  };
  const ccEmails = risStatusBuildCcEmails_(email.to, extraCcEmails || []);
  if (ccEmails.length > 0) options.cc = ccEmails.join(',');
  MailApp.sendEmail(options);
}

function risStatusBuildCcEmails_(toEmails, extraCcEmails) {
  const recipients = {};
  const result = [];

  risStatusNormalizeEmailList_(toEmails).forEach(function(email) {
    recipients[email] = true;
  });

  RIS_STATUS_CONFIG.statusCcEmails.concat(extraCcEmails || []).forEach(function(email) {
    const normalized = risStatusNormalizeEmail_(email);
    if (!risStatusIsValidEmail_(normalized) || recipients[normalized] || result.indexOf(normalized) >= 0) return;
    result.push(normalized);
  });

  return result;
}

function risStatusBuildApprovedRequestEmail_(context, dateText, deliveryTime) {
  return risStatusBuildEmail_({
    to: context.requestorEmail,
    subject: 'Delivery of Requested Items is Approved - RIS ' + context.risNo,
    greetingName: context.requestorName,
    paragraphs: [
      'We are pleased to inform you that the delivery of requested items for RIS ' + context.risNo + ' has been approved.',
      'Kindly expect the delivery on ' + dateText + ', ' + deliveryTime + '.'
    ],
    closing: 'Thank you. Please coordinate with the depot personnel if further assistance is needed.'
  });
}

function risStatusBuildRescheduledRequestEmail_(context, dateText, deliveryTime) {
  return risStatusBuildEmail_({
    to: context.requestorEmail,
    subject: 'Delivery Reschedule Notice - RIS ' + context.risNo,
    greetingName: context.requestorName,
    paragraphs: [
      'Please be informed that the delivery schedule for RIS ' + context.risNo + ' needs to be rescheduled.',
      'The proposed rescheduled delivery date and time is ' + dateText + ', ' + deliveryTime + '.',
      'If this schedule is not applicable, kindly reply with your preferred delivery date and time so we can coordinate the update.'
    ],
    closing: 'We apologize for the inconvenience and appreciate your understanding.'
  });
}

function risStatusBuildDeliveryEmail_(status, context, dateText, deliveredTime) {
  if (status === 'RIS NOT Signed') {
    if (!risStatusIsValidEmail_(context.requestorEmail)) {
      throw new Error('Missing or invalid requestor email for row ' + context.row + '.');
    }
    return risStatusBuildEmail_({
      to: context.requestorEmail,
      subject: 'Action Required: RIS Document Signature - RIS ' + context.risNo,
      greetingName: context.requestorName,
      paragraphs: [
        'This is a reminder that the RIS documents for RIS ' + context.risNo + ' remain unsigned after delivery.',
        'The delivery was recorded on ' + dateText + ', ' + deliveredTime + '.',
        'Kindly settle the required documents with the depot personnel so the delivery status can be completed.'
      ],
      closing: 'Your prompt attention to this matter is appreciated.'
    });
  }

  if (status === 'RIS Discrepancy') {
    return risStatusBuildEmail_({
      to: RIS_STATUS_CONFIG.medicalSupplyEmail,
      subject: 'RIS Discrepancy for Settlement - RIS ' + context.risNo,
      greetingName: 'Medical Supply Depot Team',
      paragraphs: [
        'Please be informed that a discrepancy was reported for RIS ' + context.risNo + '.',
        'The delivery was recorded on ' + dateText + ', ' + deliveredTime + '.',
        'Kindly coordinate and settle the items with discrepancy as soon as possible.'
      ],
      closing: 'Thank you for your immediate attention.'
    });
  }

  if (status === 'RIS Signed Completed') {
    if (!risStatusIsValidEmail_(context.requestorEmail)) {
      throw new Error('Missing or invalid requestor email for row ' + context.row + '.');
    }
    return risStatusBuildEmail_({
      to: context.requestorEmail,
      subject: 'Delivery Completed - RIS ' + context.risNo,
      greetingName: context.requestorName,
      paragraphs: [
        'This is to confirm that the delivery for RIS ' + context.risNo + ' has been completed.',
        'The delivery was recorded on ' + dateText + ', ' + deliveredTime + '.'
      ],
      closing: 'Thank you for completing the required documentation.'
    });
  }

  throw new Error('Unsupported delivery status: ' + status);
}

function risStatusBuildDailyReminderEmail_(status, context) {
  if (risStatusIsSameStatus_(status, 'RIS NOT Signed')) {
    if (!risStatusIsValidEmail_(context.requestorEmail)) throw new Error('Missing or invalid requestor email.');
    return risStatusBuildEmail_({
      to: context.requestorEmail,
      subject: 'RIS Reminder: Document Signature Required - ' + context.risNo,
      greetingName: context.requestorName,
      paragraphs: [
        'This is a daily reminder that the RIS documents for ' + context.risNo + ' remain unsigned.',
        'Kindly settle the required documents with the depot personnel so the delivery status can be completed.'
      ],
      closing: 'Your prompt attention to this matter is appreciated.'
    });
  }

  if (risStatusIsSameStatus_(status, 'RIS Discrepancy')) {
    return risStatusBuildEmail_({
      to: RIS_STATUS_CONFIG.medicalSupplyEmail,
      subject: 'RIS Reminder: Discrepancy for Settlement - ' + context.risNo,
      greetingName: 'Medical Supply Depot Team',
      paragraphs: [
        'This is a daily reminder that ' + context.risNo + ' remains marked as RIS Discrepancy.',
        'Kindly coordinate and settle the items with discrepancy as soon as possible.'
      ],
      closing: 'Thank you for your immediate attention.'
    });
  }

  throw new Error('Unsupported reminder status: ' + status);
}

function risStatusBuildEmail_(details) {
  const paragraphs = details.paragraphs || [];
  const body = [
    'Dear ' + details.greetingName + ',',
    paragraphs.join('\n\n'),
    details.closing,
    'Respectfully,',
    'Medical Supply Depot'
  ].join('\n\n');
  const htmlParagraphs = paragraphs.map(function(paragraph) {
    return '<p style="margin:0 0 12px;">' + risCoreEscapeHtml_(paragraph) + '</p>';
  }).join('');

  return {
    to: details.to,
    subject: details.subject,
    body: body,
    htmlBody: [
      '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#1f2937;">',
      '<p style="margin:0 0 12px;">Dear ' + risCoreEscapeHtml_(details.greetingName) + ',</p>',
      htmlParagraphs,
      '<p style="margin:0 0 12px;">' + risCoreEscapeHtml_(details.closing) + '</p>',
      '<p style="margin:0;">Respectfully,<br>Medical Supply Depot</p>',
      '</div>'
    ].join('')
  };
}

function risStatusParseDate_(dateValue) {
  const text = String(dateValue || '').trim();
  if (!text) throw new Error('Please select a date.');

  const date = new Date(text + 'T00:00:00');
  if (isNaN(date.getTime())) throw new Error('The selected date is invalid.');
  return date;
}

function risStatusParseSheetDate_(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
  }

  const text = String(dateValue || '').trim();
  if (!text) return null;

  let parsed = new Date(text + 'T00:00:00');
  if (isNaN(parsed.getTime())) parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function risStatusValidateDateNotBeforeSheetDate_(selectedDate, context, label) {
  const sheetDate = risStatusGetSheetBaseDate_(context);
  const selectedDateOnly = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate());

  if (selectedDateOnly.getTime() < sheetDate.getTime()) {
    throw new Error(
      label + ' cannot be earlier than the Date Requested in the sheet (' +
      risStatusFormatLongDate_(sheetDate) +
      '). Please select the same date or a later date.'
    );
  }
}

function risStatusGetSheetBaseDate_(context) {
  const sheetDate = risStatusParseSheetDate_(context.requestedDateValue || context.requestedDate);
  if (!sheetDate) throw new Error('The selected RIS row does not have a valid Date Requested value.');
  return sheetDate;
}

function risStatusFormatInputDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function risStatusFormatLongDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM d, yyyy');
}

function risStatusRequireOneOf_(value, allowedValues, label) {
  let text = String(value || '').trim();
  if (label === 'request status' && text === 'Reschedule' && allowedValues.indexOf('Rescheduled') >= 0) {
    text = 'Rescheduled';
  }
  if (allowedValues.indexOf(text) === -1) throw new Error('Please select a valid ' + label + '.');
  return text;
}

function risStatusNormalizeEmailList_(value) {
  if (Array.isArray(value)) return value.map(risStatusNormalizeEmail_).filter(risStatusIsValidEmail_);
  return String(value || '').split(/[;,\n]+/).map(risStatusNormalizeEmail_).filter(risStatusIsValidEmail_);
}

function risStatusNormalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function risStatusIsValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function risStatusIsSameStatus_(actual, expected) {
  return risStatusNormalizeLookup_(actual) === risStatusNormalizeLookup_(expected);
}

function risStatusNormalizeLookup_(value) {
  return String(value || '').trim().toLowerCase();
}

function risStatusGetActiveUserEmail_() {
  try {
    return risStatusNormalizeEmail_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function risStatusBuildDialogHtml_(mode) {
  const isRequest = mode === 'request';
  const title = isRequest ? 'Request Status' : 'Delivery Status';
  const statusOptions = isRequest ? RIS_STATUS_CONFIG.requestStatuses : RIS_STATUS_CONFIG.deliveryStatuses;
  const dateLabel = isRequest ? 'Request Date' : 'Delivery Date';
  const serverFunction = isRequest ? 'risStatusProcessRequest' : 'risStatusProcessDelivery';
  const statusLegend = isRequest
    ? '<div class="legend">' +
      '<div><strong>Approved</strong> - The requested date is accepted.</div>' +
      '<div><strong>Rescheduled</strong> - The requested date will be rescheduled. Enter the proposed new date and time below. If the proposed schedule is not applicable, the requestor may reply with their preferred date and time.</div>' +
      '</div>'
    : '<div class="legend">' +
      '<div><strong>RIS NOT Signed</strong> - Sends a reminder email to the requestor to sign or complete the RIS document.</div>' +
      '<div><strong>RIS Discrepancy</strong> - Sends a reminder email to msdmedicalsuppliesdepot@gmail.com so the discrepancy can be coordinated and settled.</div>' +
      '<div><strong>RIS Signed Completed</strong> - Sends a completion email to the requestor.</div>' +
      '</div>';

  return `
    <!doctype html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            color: #202124;
            font-family: Arial, sans-serif;
            font-size: 14px;
            margin: 0;
            padding: 22px;
          }

          h2 {
            font-size: 18px;
            font-weight: 600;
            margin: 0 0 18px;
          }

          label {
            display: block;
            font-weight: 600;
            margin: 16px 0 6px;
          }

          select,
          input {
            border: 1px solid #dadce0;
            border-radius: 6px;
            box-sizing: border-box;
            font-size: 14px;
            height: 38px;
            padding: 8px 10px;
            width: 100%;
          }

          .help {
            color: #5f6368;
            font-size: 12px;
            line-height: 1.35;
            margin-top: 5px;
          }

          .legend {
            background: #f8fafc;
            border: 1px solid #dbe3ea;
            border-radius: 6px;
            color: #344054;
            display: grid;
            gap: 8px;
            line-height: 1.4;
            margin: 8px 0 10px;
            padding: 12px;
          }

          .radio-row {
            align-items: center;
            display: flex;
            gap: 8px;
            margin-top: 8px;
          }

          .radio-row input {
            height: auto;
            width: auto;
          }

          #risNoWrap {
            display: none;
          }

          .actions {
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            margin-top: 22px;
          }

          button {
            border: 0;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            min-width: 92px;
            padding: 10px 14px;
          }

          button:disabled {
            cursor: default;
            opacity: 0.65;
          }

          .primary {
            background: #1a73e8;
            color: #fff;
          }

          .secondary {
            background: #f1f3f4;
            color: #202124;
          }

          .message {
            border-radius: 6px;
            display: none;
            line-height: 1.35;
            margin-top: 16px;
            padding: 10px 12px;
          }

          .message.success {
            background: #e6f4ea;
            color: #137333;
          }

          .message.error {
            background: #fce8e6;
            color: #c5221f;
          }
        </style>
      </head>
      <body>
        <h2>${title}</h2>

        <form id="statusForm">
          <label>Find RIS Record</label>
          <div class="radio-row">
            <input checked id="selectedRow" name="targetMode" type="radio" value="selectedRow">
            <span>Use selected row</span>
          </div>
          <div class="radio-row">
            <input id="risNoMode" name="targetMode" type="radio" value="risNo">
            <span>Enter RIS No.</span>
          </div>

          <div id="risNoWrap">
            <label for="risNo">RIS No.</label>
            <input id="risNo" name="risNo" placeholder="Example: RIS-2026-08-003">
          </div>

          <label for="status">Status</label>
          ${statusLegend}
          <select id="status" name="status" required>
            ${statusOptions.map(function(option) {
              return '<option value="' + risCoreEscapeHtml_(option) + '">' + risCoreEscapeHtml_(option) + '</option>';
            }).join('')}
          </select>

          <label for="date">${dateLabel}</label>
          <input id="date" name="date" required type="date">
          <div class="help" id="dateHelp">Defaults to Date Requested in the sheet. Same date or later only.</div>

          <label for="time">Time</label>
          <select id="time" name="time" required>
            ${RIS_STATUS_CONFIG.timeOptions.map(function(option) {
              return '<option value="' + risCoreEscapeHtml_(option) + '">' + risCoreEscapeHtml_(option) + '</option>';
            }).join('')}
          </select>

          <label for="username">Admin Username</label>
          <input id="username" name="username" required autocomplete="username" placeholder="Username from Admin Emails">

          <div class="actions">
            <button class="secondary" onclick="google.script.host.close()" type="button">Cancel</button>
            <button class="primary" id="submitButton" type="submit">Update &amp; Send Email</button>
          </div>

          <div class="message" id="message"></div>
        </form>

        <script>
          const form = document.getElementById('statusForm');
          const message = document.getElementById('message');
          const submitButton = document.getElementById('submitButton');
          const risNoWrap = document.getElementById('risNoWrap');
          const risNoInput = document.getElementById('risNo');
          const dateInput = document.getElementById('date');
          const dateHelp = document.getElementById('dateHelp');
          const defaultDateHelp = dateHelp.textContent;

          function updateRisNoVisibility() {
            const mode = form.targetMode.value;
            risNoWrap.style.display = mode === 'risNo' ? 'block' : 'none';
            risNoInput.required = mode === 'risNo';
          }

          function getFormData() {
            return Object.fromEntries(new FormData(form).entries());
          }

          function applyBaseDate(result) {
            if (!result || !result.baseDate) return;

            dateInput.min = result.baseDate;
            if (!dateInput.value || dateInput.value < result.baseDate) dateInput.value = result.baseDate;
            dateHelp.textContent = defaultDateHelp + ' Base Date: ' + result.displayDate + '.';
          }

          function refreshBaseDate() {
            const formData = getFormData();
            if (formData.targetMode === 'risNo' && !String(formData.risNo || '').trim()) {
              dateInput.removeAttribute('min');
              dateHelp.textContent = defaultDateHelp;
              return;
            }

            google.script.run
              .withSuccessHandler(applyBaseDate)
              .withFailureHandler(error => {
                dateInput.removeAttribute('min');
                dateHelp.textContent = defaultDateHelp;
                showMessage(error.message || String(error), 'error');
              })
              .risStatusGetBaseDate(formData);
          }

          function showMessage(text, type) {
            message.textContent = text;
            message.className = 'message ' + type;
            message.style.display = 'block';
          }

          Array.prototype.forEach.call(form.targetMode, input => {
            input.addEventListener('change', () => {
              updateRisNoVisibility();
              refreshBaseDate();
            });
          });

          risNoInput.addEventListener('blur', refreshBaseDate);
          risNoInput.addEventListener('change', refreshBaseDate);

          form.addEventListener('submit', event => {
            event.preventDefault();
            if (dateInput.min && dateInput.value && dateInput.value < dateInput.min) {
              showMessage('The selected date cannot be earlier than Date Requested in the sheet.', 'error');
              return;
            }

            submitButton.disabled = true;
            submitButton.textContent = 'Sending...';
            message.style.display = 'none';

            google.script.run
              .withSuccessHandler(result => {
                showMessage(result, 'success');
                setTimeout(() => google.script.host.close(), 1300);
              })
              .withFailureHandler(error => {
                showMessage(error.message || String(error), 'error');
                submitButton.disabled = false;
                submitButton.textContent = 'Update & Send Email';
              })
              .${serverFunction}(getFormData());
          });

          updateRisNoVisibility();
          refreshBaseDate();
        </script>
      </body>
    </html>
  `;
}
