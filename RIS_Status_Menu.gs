// ==========================================
// RIS STATUS EMAIL MENU
// ==========================================

const RIS_STATUS_CONFIG = {
  menuName: 'RIS Status',
  requestStatuses: ['Request Submitted', 'Request Approved', 'Request Rejected'],
  deliveryStatuses: ['RIS NOT Signed', 'RIS Discrepancy', 'RIS Signed Completed']
};

function risStatusAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(RIS_STATUS_CONFIG.menuName)
    .addItem('Send Request Status email', 'risStatusSendRequestStatusPrompt')
    .addItem('Send Delivery Status email', 'risStatusSendDeliveryStatusPrompt')
    .addToUi();
}

function risStatusSendRequestStatusPrompt() {
  risStatusPromptAndSend_('request');
}

function risStatusSendDeliveryStatusPrompt() {
  risStatusPromptAndSend_('delivery');
}

function risStatusPromptAndSend_(kind) {
  const ui = SpreadsheetApp.getUi();
  const risResponse = ui.prompt('Find RIS Record', 'Enter RIS No. or Record ID:', ui.ButtonSet.OK_CANCEL);
  if (risResponse.getSelectedButton() !== ui.Button.OK) return;

  const risNo = risResponse.getResponseText().trim();
  if (!risNo) {
    ui.alert('Please enter the RIS No.');
    return;
  }

  const statuses = kind === 'delivery' ? RIS_STATUS_CONFIG.deliveryStatuses : RIS_STATUS_CONFIG.requestStatuses;
  const statusResponse = ui.prompt(
    kind === 'delivery' ? 'Delivery Status' : 'Request Status',
    'Enter one status exactly:\n\n' + statuses.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (statusResponse.getSelectedButton() !== ui.Button.OK) return;

  const status = risStatusNormalizeChoice_(statusResponse.getResponseText(), statuses);
  if (!status) {
    ui.alert('Please choose one of these statuses:\n\n' + statuses.join('\n'));
    return;
  }

  try {
    const result = risStatusSendByRisNo_(risNo, status, kind);
    ui.alert(result.message);
  } catch (error) {
    ui.alert('Could not send status email:\n\n' + risCoreErrorMessage_(error));
  }
}

function risStatusSendByRisNo(risNo, status, kind) {
  return risStatusSendByRisNo_(risNo, status, kind || 'request');
}

function risStatusSendByRisNo_(risNo, status, kind) {
  const bundle = risCoreGetRisBundle_(risNo);
  const ss = bundle.ss;
  const entry = bundle.entry;
  const recipients = risCoreReadEmailRecipients_(ss);
  const to = recipients.to.slice();

  if (entry.requestorEmail) to.push(entry.requestorEmail);
  if (to.length === 0) throw new Error('No email recipients found.');

  risCoreUpdateRecordFields_(
    bundle.entriesSheet,
    entry.rowNumber,
    RIS_ENTRIES_DEFAULT_HEADERS,
    RIS_ENTRIES_ALIASES,
    { risStatus: status }
  );

  const subject = risStatusSubject_(entry, status, kind);
  const htmlBody = risStatusHtmlBody_(entry, status, kind);

  MailApp.sendEmail({
    to: risCoreUnique_(to).join(','),
    cc: risCoreUnique_(recipients.cc).join(','),
    subject: subject,
    htmlBody: htmlBody,
    body: risStatusPlainBody_(entry, status, kind)
  });

  return {
    success: true,
    message: 'Status email sent for ' + (entry.risNo || entry.recordId) + '.'
  };
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

function risStatusNormalizeChoice_(input, choices) {
  const wanted = risCoreNormalizeText_(input);
  for (let i = 0; i < choices.length; i++) {
    if (risCoreNormalizeText_(choices[i]) === wanted) return choices[i];
  }
  return '';
}
