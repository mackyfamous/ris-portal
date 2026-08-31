// ==========================================
// RIS PDF GENERATOR ADD-ON
// ==========================================

const RIS_PDF_CONFIG = {
  outputFolderId: 'PASTE_PDF_OUTPUT_FOLDER_ID',
  pasigLogoFileId: '1LQ10a2zc3oxRRw6tUTWzDFuNdnOkSvLQ',
  notificationSubjectPrefix: '[RIS PDF] Generated: ',
  notificationCc: 'lyn4logistics@gmail.com',
  divisionName: 'CITY HEALTH DEPARTMENT',
  minimumItemRows: 18,
  stockMark: 'X',
  paperCss: '@page { size: 8.5in 14in; margin: 0.35in; }'
};

function risPdfAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('RIS PDF Tools')
    .addItem('Authorize PDF generator', 'risPdfAuthorize')
    .addItem('Generate PDF for selected RIS', 'risPdfGenerateForSelectedRIS')
    .addItem('Generate PDF by RIS No.', 'risPdfGenerateByRisNoPrompt')
    .addToUi();
}

function risPdfAuthorize() {
  risPdfAssertConfigured_();
  const ss = risCoreGetTransactionsSpreadsheet_();
  const recipients = risCoreReadEmailRecipients_(ss);
  const target = recipients.to[0] || Session.getActiveUser().getEmail();
  if (!target) throw new Error('Add an email address in Admin Emails or Client Emails before authorizing.');

  MailApp.sendEmail({
    to: target,
    subject: 'RIS PDF generator authorization test',
    body: 'This test email confirms that RIS PDF generation notifications are authorized.'
  });
}

function risPdfGenerateForSelectedRIS() {
  const ui = SpreadsheetApp.getUi();
  try {
    const risNo = risPdfGetSelectedRisNo_();
    const username = risPdfPromptUsername_('Generate PDF for ' + risNo);
    if (!username) return;

    const result = risPdfGenerateByRisNo_(risNo, username);
    ui.alert('RIS PDF generated.\n\nRIS No.: ' + result.risNo + '\nFile: ' + result.url);
  } catch (error) {
    ui.alert('Could not generate RIS PDF:\n\n' + risCoreErrorMessage_(error));
  }
}

function risPdfGenerateByRisNoPrompt() {
  const ui = SpreadsheetApp.getUi();
  const risResponse = ui.prompt('Generate RIS PDF', 'Enter RIS No. or Record ID:', ui.ButtonSet.OK_CANCEL);
  if (risResponse.getSelectedButton() !== ui.Button.OK) return;

  const risNo = risResponse.getResponseText().trim();
  if (!risNo) {
    ui.alert('Please enter a RIS No. or Record ID.');
    return;
  }

  const username = risPdfPromptUsername_('Generate PDF for ' + risNo);
  if (!username) return;

  try {
    const result = risPdfGenerateByRisNo_(risNo, username);
    ui.alert('RIS PDF generated.\n\nRIS No.: ' + result.risNo + '\nFile: ' + result.url);
  } catch (error) {
    ui.alert('Could not generate RIS PDF:\n\n' + risCoreErrorMessage_(error));
  }
}

function risPdfGenerateByRisNo(risNo, username) {
  return risPdfGenerateByRisNo_(risNo, username);
}

function risPdfGenerateByRisNo_(risNo, username) {
  risPdfAssertConfigured_();
  const user = risCoreValidateActiveUser_(username);
  const bundle = risCoreGetRisBundle_(risNo);
  const items = risPdfPrepareItems_(bundle.items);
  const html = risPdfBuildHtml_(bundle.entry, items, user);
  const folder = DriveApp.getFolderById(RIS_PDF_CONFIG.outputFolderId);
  const fileName = risPdfSafeFileName_(bundle.entry.risNo || bundle.entry.recordId) + ' - RIS.pdf';
  const blob = Utilities.newBlob(html, 'text/html', fileName.replace(/\.pdf$/i, '.html')).getAs(MimeType.PDF);
  blob.setName(fileName);
  const file = folder.createFile(blob);

  risCoreUpdateRecordFields_(
    bundle.entriesSheet,
    bundle.entry.rowNumber,
    RIS_ENTRIES_DEFAULT_HEADERS,
    RIS_ENTRIES_ALIASES,
    {
      pdfUrl: file.getUrl(),
      pdfGeneratedAt: new Date(),
      pdfGeneratedBy: user.fullName + ' (' + user.username + ')'
    }
  );

  risPdfSendNotification_(bundle.ss, bundle.entry, items, file, user);

  return {
    success: true,
    risNo: bundle.entry.risNo,
    url: file.getUrl(),
    generatedBy: user.fullName
  };
}

function risPdfGetSelectedRisNo_() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet || sheet.getName() !== RIS_CONFIG.entriesSheetName) {
    throw new Error('Please select a row in the "' + RIS_CONFIG.entriesSheetName + '" sheet.');
  }

  const range = sheet.getActiveRange();
  if (!range || range.getRow() <= 1) {
    throw new Error('Please select a RIS data row.');
  }

  const info = risCoreGetHeaderInfo_(sheet, RIS_ENTRIES_DEFAULT_HEADERS, RIS_ENTRIES_ALIASES);
  const row = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  const risNo = info.columns.risNo ? row[info.columns.risNo - 1] : '';
  const recordId = info.columns.recordId ? row[info.columns.recordId - 1] : '';
  return risNo || recordId;
}

function risPdfPromptUsername_(title) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(title, 'Enter your username:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return '';
  return response.getResponseText().trim();
}

function risPdfAssertConfigured_() {
  if (!RIS_PDF_CONFIG.outputFolderId || RIS_PDF_CONFIG.outputFolderId.indexOf('PASTE_') === 0) {
    throw new Error('Configure RIS_PDF_CONFIG.outputFolderId first.');
  }
}

function risPdfPrepareItems_(items) {
  return items.map(function(item) {
    let recovered = {};
    try {
      recovered = risCoreRecoverInventoryDataForItem_(item);
    } catch (error) {
      recovered = {};
    }

    const result = {};
    [
      'itemCode',
      'itemDescription',
      'uom',
      'poNumber',
      'supplier',
      'batch',
      'expiry',
      'unitCost',
      'remarks'
    ].forEach(function(field) {
      result[field] = item[field] || recovered[field] || '';
    });

    result.qtyRequested = risCoreParseNumber_(item.qtyRequested);
    result.issuedQty = risCoreParseNumber_(item.issuedQty || item.qtyRequested);
    result.currentSoh = risCoreParseNumber_(item.currentSoh || recovered.stock);
    result.totalCost = result.issuedQty * risCoreParseNumber_(result.unitCost);
    result.stockAvailable = result.currentSoh > 0;
    result.poWithSupplier = risPdfPoWithSupplier_(result);
    return result;
  });
}

function risPdfBuildHtml_(entry, items, user) {
  const logoDataUri = risPdfGetPasigLogoDataUri_();
  const logoHtml = logoDataUri
    ? '<img class="pasig-logo" src="' + logoDataUri + '" alt="Pasig City logo">'
    : '<div class="pasig-logo-fallback">PASIG</div>';
  const itemRows = risPdfBuildItemRows_(items);
  const grandTotal = items.reduce(function(sum, item) {
    return sum + risCoreParseNumber_(item.totalCost);
  }, 0);

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    RIS_PDF_CONFIG.paperCss +
    '*{box-sizing:border-box;}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:8px;color:#111;margin:0;}' +
    'table{border-collapse:collapse;table-layout:fixed;width:100%;}' +
    'td,th{border:1px solid #111;padding:2px 3px;vertical-align:middle;}' +
    'p,h1,h2{margin:0;}' +
    '.brand td{height:76px;}' +
    '.logo-cell{width:31%;text-align:center;padding:0;}' +
    '.pasig-logo{display:block;max-width:100%;max-height:74px;margin:0 auto;}' +
    '.pasig-logo-fallback{font-size:28px;font-weight:700;color:#164e9f;letter-spacing:1px;}' +
    '.title-cell{text-align:center;color:#9f9f9f;font-weight:700;}' +
    '.title-cell h1{font-size:18px;line-height:1.1;letter-spacing:.2px;}' +
    '.title-cell h2{font-size:14px;line-height:1.2;margin-top:9px;}' +
    '.meta{margin-top:-1px;}' +
    '.meta th,.meta td{height:20px;}' +
    '.meta th{text-align:left;font-weight:700;width:8%;}' +
    '.meta td{text-align:center;}' +
    '.meta .left-value{width:46%;}' +
    '.meta .label-small{width:8%;}' +
    '.meta .value-small{width:15%;}' +
    '.accent{color:#a10f45;}' +
    '.items{margin-top:-1px;}' +
    '.items th,.items td{font-size:7px;line-height:1.05;text-align:center;word-break:break-word;overflow-wrap:break-word;}' +
    '.items .group th{height:19px;font-weight:700;font-size:8px;}' +
    '.items .head th{height:39px;font-weight:700;}' +
    '.items tbody td{height:32px;}' +
    '.items .desc{text-align:left;}' +
    '.items .num{text-align:right;}' +
    '.items .stock{font-size:8px;}' +
    '.items .blank td{height:34px;}' +
    '.items .grand td{height:20px;font-weight:700;}' +
    '.items .grand-label{text-align:right;}' +
    '.items .grand-value{text-align:right;}' +
    '.purpose{margin-top:21px;}' +
    '.purpose td{height:29px;font-weight:700;text-align:left;}' +
    '.signatures{margin-top:-1px;}' +
    '.signatures th,.signatures td{height:24px;font-size:8px;}' +
    '.signatures th{text-align:center;font-weight:700;}' +
    '.signatures .row-label{text-align:left;font-weight:700;width:20%;}' +
    '.signatures .name{text-align:center;}' +
    '.prepared td{height:24px;font-weight:700;text-align:left;}' +
    '.prepared .value{height:27px;font-weight:400;text-align:center;}' +
    '</style></head><body>' +
    '<table class="brand"><tr><td class="logo-cell">' + logoHtml + '</td><td class="title-cell"><h1>REQUISITION AND ISSUE SLIP</h1><h2>CITY GOVERNMENT OF PASIG</h2></td></tr></table>' +
    '<table class="meta"><colgroup><col style="width:8%"><col style="width:46%"><col style="width:8%"><col style="width:15%"><col style="width:8%"><col style="width:15%"></colgroup>' +
    '<tr><th>Entity Name:</th><td class="left-value">' + risCoreEscapeHtml_(entry.requestorProgram || '') + '</td><th class="label-small">RIS No.:</th><td class="value-small accent">' + risCoreEscapeHtml_(entry.risNo || entry.recordId || '') + '</td><th class="label-small">Date:</th><td class="value-small accent">' + risCoreEscapeHtml_(risPdfDate_(entry.deliveryDate, '')) + '</td></tr>' +
    '<tr><th>Office:</th><td>' + risCoreEscapeHtml_(entry.deliveryLocation || '') + '</td><td colspan="4"></td></tr>' +
    '<tr><th>Division:</th><td><b>' + risCoreEscapeHtml_(RIS_PDF_CONFIG.divisionName) + '</b></td><td colspan="4"></td></tr>' +
    '</table>' +
    '<table class="items"><colgroup>' +
    '<col style="width:9%"><col style="width:25%"><col style="width:5%"><col style="width:8%"><col style="width:6%"><col style="width:7%"><col style="width:7%"><col style="width:3%"><col style="width:3%"><col style="width:7%"><col style="width:7%"><col style="width:8%"><col style="width:5%">' +
    '</colgroup><thead><tr class="group"><th colspan="7">Requisition</th><th colspan="6">Issuance</th></tr>' +
    '<tr class="head"><th>Item Code</th><th>Item Description</th><th>Unit of Measurement</th><th>Purchase Order #</th><th>Batch / Lot No.</th><th>Expiration Date</th><th>Quantity Requested</th><th>yes</th><th>no</th><th>Quantity Requested</th><th>Unit Cost</th><th>Total Amount</th><th>Remarks</th></tr></thead>' +
    '<tbody>' + itemRows +
    '<tr class="grand"><td colspan="11" class="grand-label">Grand Total:</td><td colspan="2" class="grand-value">' + risCoreEscapeHtml_(risPdfMoney_(grandTotal)) + '</td></tr>' +
    '</tbody></table>' +
    '<table class="purpose"><tr><td>Purpose: ' + risCoreEscapeHtml_(entry.purpose || '') + '</td></tr></table>' +
    '<table class="signatures"><colgroup><col style="width:20%"><col style="width:20%"><col style="width:20%"><col style="width:20%"><col style="width:20%"></colgroup>' +
    '<tr><td></td><th>Requested by:</th><th>Approved by:</th><th>Issued by:</th><th>Received by:</th></tr>' +
    '<tr><td class="row-label">Signature:</td><td></td><td></td><td></td><td></td></tr>' +
    '<tr><td class="row-label">Printed Name:</td><td class="name">' + risCoreEscapeHtml_(entry.requestedBy || '') + '</td><td class="name">' + risCoreEscapeHtml_(entry.approvedBy || '') + '</td><td class="name">' + risCoreEscapeHtml_(entry.issuedBy || '') + '</td><td class="name">' + risCoreEscapeHtml_(entry.receivedBy || '') + '</td></tr>' +
    '<tr><td class="row-label">Designation:</td><td></td><td></td><td></td><td></td></tr>' +
    '<tr><td class="row-label">Date:</td><td></td><td></td><td></td><td></td></tr>' +
    '</table>' +
    '<table class="prepared"><tr><td colspan="3">Prepared RIS by:</td><td colspan="2">Checked By:</td></tr>' +
    '<tr><td colspan="3" class="value">' + risCoreEscapeHtml_(user.fullName + ' / ' + user.username) + '</td><td colspan="2" class="value"></td></tr></table>' +
    '</body></html>';
}

function risPdfBuildItemRows_(items) {
  const rows = [];
  items.forEach(function(item) {
    rows.push('<tr>' +
      '<td>' + risCoreEscapeHtml_(item.itemCode) + '</td>' +
      '<td class="desc">' + risCoreEscapeHtml_(item.itemDescription) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.uom) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.poWithSupplier) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.batch || 'N/A') + '</td>' +
      '<td>' + risCoreEscapeHtml_(risPdfDate_(item.expiry, 'N/A')) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risPdfQuantity_(item.qtyRequested)) + '</td>' +
      '<td class="stock">' + (item.stockAvailable ? RIS_PDF_CONFIG.stockMark : '') + '</td>' +
      '<td class="stock">' + (item.stockAvailable ? '' : RIS_PDF_CONFIG.stockMark) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risPdfQuantity_(item.issuedQty)) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risPdfAmount_(item.unitCost)) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risPdfAmount_(item.totalCost)) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.remarks) + '</td>' +
      '</tr>');
  });

  const minimumRows = Math.max(RIS_PDF_CONFIG.minimumItemRows || 0, items.length);
  for (let i = items.length; i < minimumRows; i++) {
    rows.push('<tr class="blank"><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>');
  }

  return rows.join('');
}

function risPdfGetPasigLogoDataUri_() {
  const fileId = risPdfDriveFileId_(RIS_PDF_CONFIG.pasigLogoFileId);
  if (!fileId || fileId.indexOf('PASTE_') === 0) return '';

  try {
    return risPdfImageDataUriFromBlob_(DriveApp.getFileById(fileId).getBlob());
  } catch (driveError) {
    try {
      const response = UrlFetchApp.fetch('https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId), {
        muteHttpExceptions: true
      });
      if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return '';
      return risPdfImageDataUriFromBlob_(response.getBlob());
    } catch (fetchError) {
      return '';
    }
  }
}

function risPdfImageDataUriFromBlob_(blob) {
  const contentType = blob.getContentType() || 'image/png';
  if (contentType.indexOf('image/') !== 0) return '';
  return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function risPdfDriveFileId_(value) {
  const text = String(value || '').trim();
  const pathMatch = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (pathMatch) return pathMatch[1];
  const queryMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (queryMatch) return queryMatch[1];
  return text;
}

function risPdfDate_(value, blankText) {
  const date = risCoreParseDate_(value);
  if (date) return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  return value || blankText || '';
}

function risPdfQuantity_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const number = risCoreParseNumber_(value);
  return number.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function risPdfAmount_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const number = risCoreParseNumber_(value);
  return number.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function risPdfMoney_(value) {
  return String.fromCharCode(8369) + ' ' + risPdfAmount_(value);
}

function risPdfSendNotification_(ss, entry, items, file, user) {
  const to = risCoreUnique_([risPdfGeneratedByEmail_(ss, user)]);
  if (to.length === 0) return;

  MailApp.sendEmail({
    to: to.join(','),
    cc: risCoreUnique_([RIS_PDF_CONFIG.notificationCc]).join(','),
    subject: RIS_PDF_CONFIG.notificationSubjectPrefix + (entry.risNo || entry.recordId),
    htmlBody: '<p>RIS PDF generated: <b>' + risCoreEscapeHtml_(entry.risNo || entry.recordId) + '</b></p>' +
      '<p>Generated by: ' + risCoreEscapeHtml_(user.fullName + ' (' + user.username + ')') + '</p>' +
      '<p><a href="' + risCoreEscapeHtml_(file.getUrl()) + '">Open RIS PDF</a></p>',
    body: 'RIS PDF generated: ' + (entry.risNo || entry.recordId) + '\n' + file.getUrl()
  });
}

function risPdfGeneratedByEmail_(ss, user) {
  const adminEmail = risPdfFindAdminEmail_(ss, user);
  if (adminEmail) return adminEmail;

  const username = String(user.username || '').trim();
  if (risPdfLooksLikeEmail_(username)) return username;

  const activeEmail = Session.getActiveUser().getEmail();
  return activeEmail || '';
}

function risPdfFindAdminEmail_(ss, user) {
  const sheet = ss.getSheetByName(RIS_CONFIG.adminEmailsSheetName) || ss.getSheetByName(RIS_CONFIG.usersSheetName);
  if (!sheet || sheet.getLastRow() < 2) return '';

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const usernameCol = risFindHeaderColumn_(headers, ['Username', 'User Name']);
  const nameCol = risFindHeaderColumn_(headers, ['Name', 'Full Name']);
  const recipientCol = risFindHeaderColumn_(headers, ['Recipients', 'Receipients', 'Recipient', 'Email', 'Email Address', 'TO']);
  const senderCol = risFindHeaderColumn_(headers, ['Sender', 'From']);
  const wanted = [user.username, user.fullName].map(risCoreNormalizeText_).filter(String);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowUsername = usernameCol ? row[usernameCol - 1] : '';
    const rowName = nameCol ? row[nameCol - 1] : '';
    const rowRecipient = recipientCol ? row[recipientCol - 1] : '';
    const rowSender = senderCol ? row[senderCol - 1] : '';
    const rowKeys = [rowUsername, rowName, rowRecipient, rowSender].map(risCoreNormalizeText_);
    const matches = rowKeys.some(function(value) {
      return wanted.indexOf(value) !== -1;
    });

    if (matches) {
      return risPdfFirstEmail_(rowRecipient) || risPdfFirstEmail_(rowSender) ||
        (risPdfLooksLikeEmail_(rowUsername) ? String(rowUsername).trim() : '');
    }
  }

  return '';
}

function risPdfFirstEmail_(value) {
  const emails = risCoreSplitEmailList_(value).filter(risPdfLooksLikeEmail_);
  return emails[0] || '';
}

function risPdfLooksLikeEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function risPdfPoWithSupplier_(item) {
  const po = String(item.poNumber || '').trim();
  const supplier = String(item.supplier || '').trim();
  if (po && supplier) return po + ' (' + supplier + ')';
  return po || supplier;
}

function risPdfSafeFileName_(value) {
  return String(value || 'RIS')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}
