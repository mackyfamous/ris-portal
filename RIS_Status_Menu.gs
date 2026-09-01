// ==========================================
// RIS STATUS MENU
// ==========================================

const RIS_STATUS_CONFIG = {
  menuName: 'RIS Status',
  defaultStatus: 'Request Submitted',
  requestStatuses: ['Request Submitted', 'Request Approved', 'Request Rejected'],
  deliveryStatuses: ['RIS NOT Signed', 'RIS Discrepancy', 'RIS Signed Completed'],
  sendToConfiguredRecipients: true,
  sendToRequestor: true,
  notificationCc: ''
};

function risStatusAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(RIS_STATUS_CONFIG.menuName)
    .addItem('Show Current Status', 'risStatusShowCurrentPrompt')
    .addSeparator()
    .addItem('Update Request Status', 'risStatusUpdateRequestStatusPrompt')
    .addItem('Update Delivery Status', 'risStatusUpdateDeliveryStatusPrompt')
    .addSeparator()
    .addItem('Email Request Status', 'risStatusSendRequestStatusPrompt')
    .addItem('Email Delivery Status', 'risStatusSendDeliveryStatusPrompt')
    .addToUi();
}

function risStatusShowCurrentPrompt() {
  const ui = SpreadsheetApp.getUi();
  const risNo = risStatusPromptRisNo_(ui);
  if (!risNo) return;

  try {
    const bundle = risCoreGetRisBundle_(risNo);
    const entry = bundle.entry;
    ui.alert(
      'Current RIS Status',
      [
        'RIS No.: ' + (entry.risNo || entry.recordId || risNo),
        'Status: ' + risStatusCurrentStatus_(entry),
        'Program: ' + (entry.requestorProgram || ''),
        'Category: ' + (entry.category || entry.inventorySource || '')
      ].join('\n'),
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Could not load RIS status:\n\n' + risCoreErrorMessage_(error));
  }
}

function risStatusUpdateRequestStatusPrompt() {
  risStatusPromptAndApply_('request', false);
}

function risStatusUpdateDeliveryStatusPrompt() {
  risStatusPromptAndApply_('delivery', false);
}

function risStatusSendRequestStatusPrompt() {
  risStatusPromptAndApply_('request', true);
}

function risStatusSendDeliveryStatusPrompt() {
  risStatusPromptAndApply_('delivery', true);
}

function risStatusPromptAndApply_(kind, sendEmail) {
  const ui = SpreadsheetApp.getUi();
  const normalizedKind = risStatusNormalizeKind_(kind);
  const risNo = risStatusPromptRisNo_(ui);
  if (!risNo) return;

  const statuses = risStatusChoices_(normalizedKind);
  const statusResponse = ui.prompt(
    normalizedKind === 'delivery' ? 'Delivery Status' : 'Request Status',
    'Enter one status exactly:\n\n' + statuses.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (statusResponse.getSelectedButton() !== ui.Button.OK) return;

  let status;
  try {
    status = risStatusValidateStatus_(statusResponse.getResponseText(), normalizedKind);
  } catch (error) {
    ui.alert(risCoreErrorMessage_(error));
    return;
  }

  const username = risStatusPromptUsername_(ui);
  if (username === null) return;

  try {
    const result = sendEmail
      ? risStatusSendByRisNo_(risNo, status, normalizedKind, username)
      : risStatusUpdateByRisNo_(risNo, status, normalizedKind, username);
    ui.alert(result.message);
  } catch (error) {
    ui.alert((sendEmail ? 'Could not send status email:' : 'Could not update RIS status:') + '\n\n' + risCoreErrorMessage_(error));
  }
}

function risStatusPromptRisNo_(ui) {
  const selectedRisNo = risStatusGetSelectedRisKey_();
  const helper = selectedRisNo
    ? '\n\nSelected row detected: ' + selectedRisNo + '\nLeave the prompt blank to use the selected row.'
    : '';
  const risResponse = ui.prompt('Find RIS Record', 'Enter RIS No. or Record ID:' + helper, ui.ButtonSet.OK_CANCEL);
  if (risResponse.getSelectedButton() !== ui.Button.OK) return '';

  const entered = risResponse.getResponseText().trim();
  if (entered) return entered;
  if (selectedRisNo) return selectedRisNo;

  ui.alert('Please enter the RIS No.');
  return '';
}

function risStatusPromptUsername_(ui) {
  const userResponse = ui.prompt('Confirm Admin', 'Enter your admin username:', ui.ButtonSet.OK_CANCEL);
  if (userResponse.getSelectedButton() !== ui.Button.OK) return null;

  const username = userResponse.getResponseText().trim();
  if (!username) {
    ui.alert('Username is required.');
    return null;
  }
  return username;
}

function risStatusUpdateByRisNo(risNo, status, kind, username) {
  return risStatusUpdateByRisNo_(risNo, status, kind || 'request', username || '');
}

function risStatusSendByRisNo(risNo, status, kind, username) {
  return risStatusSendByRisNo_(risNo, status, kind || 'request', username || '');
}

function risStatusUpdateByRisNo_(risNo, status, kind, username) {
  const record = risStatusUpdateRecord_(risNo, status, kind, username);
  return risStatusPublicResult_(record, false);
}

function risStatusSendByRisNo_(risNo, status, kind, username) {
  const record = risStatusUpdateRecord_(risNo, status, kind, username);
  risStatusSendEmail_(record.bundle, record.status, record.kind);
  return risStatusPublicResult_(record, true);
}

function risStatusUpdateRecord_(risNo, status, kind, username) {
  const normalizedKind = risStatusNormalizeKind_(kind);
  const validStatus = risStatusValidateStatus_(status, normalizedKind);
  if (username) risCoreValidateActiveUser_(username);

  const bundle = risCoreGetRisBundle_(risNo);
  const entry = bundle.entry;
  risCoreUpdateRecordFields_(
    bundle.entriesSheet,
    entry.rowNumber,
    RIS_ENTRIES_DEFAULT_HEADERS,
    RIS_ENTRIES_ALIASES,
    { risStatus: validStatus }
  );
  entry.risStatus = validStatus;

  return {
    bundle: bundle,
    entry: entry,
    risNo: entry.risNo || entry.recordId || risNo,
    status: validStatus,
    kind: normalizedKind
  };
}

function risStatusPublicResult_(record, emailed) {
  return {
    success: true,
    risNo: record.risNo,
    status: record.status,
    emailed: emailed,
    message: emailed
      ? 'RIS status updated to "' + record.status + '" and email sent for ' + record.risNo + '.'
      : 'RIS status updated to "' + record.status + '" for ' + record.risNo + '.'
  };
}

function risStatusSendEmail_(bundle, status, kind) {
  const ss = bundle.ss;
  const entry = bundle.entry;
  const recipients = RIS_STATUS_CONFIG.sendToConfiguredRecipients
    ? risCoreReadEmailRecipients_(ss)
    : { to: [], cc: [] };
  const to = recipients.to.slice();
  const cc = recipients.cc.slice();

  if (RIS_STATUS_CONFIG.sendToRequestor && entry.requestorEmail) to.push(entry.requestorEmail);
  if (RIS_STATUS_CONFIG.notificationCc) cc.push(RIS_STATUS_CONFIG.notificationCc);
  if (to.length === 0) throw new Error('No email recipients found. Add a requestor email or configured recipients first.');

  const message = {
    to: risCoreUnique_(to).join(','),
    subject: risStatusSubject_(entry, status, kind),
    htmlBody: risStatusHtmlBody_(entry, status, kind),
    body: risStatusPlainBody_(entry, status, kind)
  };
  const uniqueCc = risCoreUnique_(cc);
  if (uniqueCc.length > 0) message.cc = uniqueCc.join(',');

  MailApp.sendEmail(message);
}

function risStatusSubject_(entry, status, kind) {
  const risNo = entry.risNo || entry.recordId || 'RIS request';
  if (kind === 'delivery') {
    if (status === 'RIS NOT Signed') return 'Action Required: RIS Document Signature - ' + risNo;
    if (status === 'RIS Discrepancy') return 'RIS Discrepancy for Settlement - ' + risNo;
    if (status === 'RIS Signed Completed') return 'Delivery Completed - ' + risNo;
  }
  return '[RIS] ' + status + ': ' + risNo;
}

function risStatusPlainBody_(entry, status, kind) {
  return [
    'RIS Status Update',
    '',
    'RIS No.: ' + (entry.risNo || entry.recordId || ''),
    'Status: ' + status,
    'Category: ' + (entry.category || entry.inventorySource || ''),
    'Program: ' + (entry.requestorProgram || ''),
    'Delivery Location: ' + (entry.deliveryLocation || ''),
    '',
    risStatusMessage_(status, kind)
  ].join('\n');
}

function risStatusHtmlBody_(entry, status, kind) {
  return [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#172033;">',
    '<h2 style="margin:0 0 12px;color:#0f766e;">RIS Status Update</h2>',
    '<table cellpadding="4" cellspacing="0">',
    '<tr><td><b>RIS No.</b></td><td>' + risCoreEscapeHtml_(entry.risNo || entry.recordId || '') + '</td></tr>',
    '<tr><td><b>Status</b></td><td>' + risCoreEscapeHtml_(status) + '</td></tr>',
    '<tr><td><b>Category</b></td><td>' + risCoreEscapeHtml_(entry.category || entry.inventorySource || '') + '</td></tr>',
    '<tr><td><b>Program</b></td><td>' + risCoreEscapeHtml_(entry.requestorProgram || '') + '</td></tr>',
    '<tr><td><b>Delivery Location</b></td><td>' + risCoreEscapeHtml_(entry.deliveryLocation || '') + '</td></tr>',
    '</table>',
    '<p>' + risCoreEscapeHtml_(risStatusMessage_(status, kind)) + '</p>',
    '</div>'
  ].join('');
}

function risStatusMessage_(status, kind) {
  if (kind === 'delivery') {
    if (status === 'RIS NOT Signed') {
      return 'The RIS documents remain unsigned after delivery. Please coordinate for completion.';
    }
    if (status === 'RIS Discrepancy') {
      return 'A discrepancy was reported for this RIS. Please review and coordinate for settlement.';
    }
    if (status === 'RIS Signed Completed') {
      return 'The delivery for this RIS has been completed and signed.';
    }
  }

  if (status === 'Request Approved') return 'Your RIS request has been approved.';
  if (status === 'Request Rejected') return 'Your RIS request was not approved. Please coordinate with the logistics team.';
  return 'Your RIS request has been submitted and is now being processed.';
}

function risStatusChoices_(kind) {
  return risStatusNormalizeKind_(kind) === 'delivery'
    ? RIS_STATUS_CONFIG.deliveryStatuses
    : RIS_STATUS_CONFIG.requestStatuses;
}

function risStatusNormalizeKind_(kind) {
  return risCoreNormalizeText_(kind) === 'delivery' ? 'delivery' : 'request';
}

function risStatusValidateStatus_(status, kind) {
  const choices = risStatusChoices_(kind);
  const choice = risStatusNormalizeChoice_(status, choices);
  if (choice) return choice;

  if (risCoreNormalizeText_(status) === 'submitted') return RIS_STATUS_CONFIG.defaultStatus;
  throw new Error('Please choose one of these statuses:\n\n' + choices.join('\n'));
}

function risStatusCurrentStatus_(entry) {
  const current = risCleanText_(entry.risStatus);
  if (risCoreNormalizeText_(current) === 'submitted') return RIS_STATUS_CONFIG.defaultStatus;
  return current || RIS_STATUS_CONFIG.defaultStatus;
}

function risStatusNormalizeChoice_(input, choices) {
  const wanted = risCoreNormalizeText_(input);
  for (let i = 0; i < choices.length; i++) {
    if (risCoreNormalizeText_(choices[i]) === wanted) return choices[i];
  }
  return '';
}

function risStatusGetSelectedRisKey_() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet && sheet.getActiveRange();
    if (!sheet || !range || range.getRow() <= 1 || sheet.getLastColumn() < 1) return '';

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const risAliases = RIS_ENTRIES_ALIASES.risNo.concat(RIS_ITEMS_ALIASES.risNo);
    const recordAliases = RIS_ENTRIES_ALIASES.recordId.concat(RIS_ITEMS_ALIASES.recordId);
    const risCol = risFindHeaderColumn_(headers, risAliases);
    const recordCol = risFindHeaderColumn_(headers, recordAliases);
    if (!risCol && !recordCol) return '';

    const row = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    return risCleanText_(risCol ? row[risCol - 1] : '') || risCleanText_(recordCol ? row[recordCol - 1] : '');
  } catch (error) {
    return '';
  }
}
