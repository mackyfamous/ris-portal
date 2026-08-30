// ==========================================
// RIS PDF GENERATOR ADD-ON
// ==========================================

const RIS_PDF_CONFIG = {
  outputFolderId: 'PASTE_PDF_OUTPUT_FOLDER_ID',
  notificationSubjectPrefix: '[RIS PDF] Generated: ',
  divisionName: 'CITY HEALTH DEPARTMENT',
  paperCss: '@page { size: 8.5in 14in; margin: 0.45in; }'
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
  if (!target) throw new Error('Add an email address in the Emails sheet before authorizing.');

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
  const rows = items.map(function(item) {
    return '<tr>' +
      '<td>' + risCoreEscapeHtml_(item.itemCode) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.itemDescription) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.uom) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.poWithSupplier) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.batch) + '</td>' +
      '<td>' + risCoreEscapeHtml_(risCoreFormatDate_(item.expiry)) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(item.qtyRequested) + '</td>' +
      '<td>' + (item.stockAvailable ? 'YES' : '') + '</td>' +
      '<td>' + (item.stockAvailable ? '' : 'NO') + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(item.issuedQty) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risCoreParseNumber_(item.unitCost).toFixed(2)) + '</td>' +
      '<td class="num">' + risCoreEscapeHtml_(risCoreParseNumber_(item.totalCost).toFixed(2)) + '</td>' +
      '<td>' + risCoreEscapeHtml_(item.remarks) + '</td>' +
      '</tr>';
  }).join('');

  const grandTotal = items.reduce(function(sum, item) {
    return sum + risCoreParseNumber_(item.totalCost);
  }, 0);

  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    RIS_PDF_CONFIG.paperCss +
    'body{font-family:Arial,sans-serif;font-size:10px;color:#111827;}' +
    'h1,h2,p{margin:0;}' +
    '.center{text-align:center;}' +
    '.title{font-size:16px;font-weight:700;}' +
    '.subtitle{font-size:12px;font-weight:700;margin-top:2px;}' +
    '.meta{width:100%;border-collapse:collapse;margin-top:16px;margin-bottom:8px;}' +
    '.meta td{padding:3px 4px;font-size:10px;}' +
    '.label{font-weight:700;width:120px;}' +
    'table.items{width:100%;border-collapse:collapse;table-layout:fixed;}' +
    '.items th,.items td{border:1px solid #111827;padding:3px;vertical-align:top;word-wrap:break-word;}' +
    '.items th{font-size:9px;background:#f3f4f6;}' +
    '.num{text-align:right;}' +
    '.total-row td{font-weight:700;}' +
    '.purpose{border:1px solid #111827;border-top:0;padding:6px;min-height:28px;}' +
    '.signatures{width:100%;border-collapse:collapse;margin-top:18px;}' +
    '.signatures td{border:1px solid #111827;padding:7px;height:34px;vertical-align:bottom;}' +
    '.small{font-size:9px;color:#374151;}' +
    '</style></head><body>' +
    '<div class="center"><h1 class="title">REQUISITION AND ISSUE SLIP</h1><h2 class="subtitle">CITY GOVERNMENT OF PASIG</h2></div>' +
    '<table class="meta">' +
    '<tr><td class="label">Entity Name:</td><td></td><td class="label">RIS #:</td><td>' + risCoreEscapeHtml_(entry.risNo || entry.recordId) + '</td></tr>' +
    '<tr><td class="label">Office:</td><td>' + risCoreEscapeHtml_(entry.deliveryLocation || entry.requestorProgram) + '</td><td class="label">Date:</td><td>' + risCoreEscapeHtml_(risCoreFormatDate_(entry.deliveryDate)) + '</td></tr>' +
    '<tr><td class="label">Division:</td><td>' + risCoreEscapeHtml_(RIS_PDF_CONFIG.divisionName) + '</td><td class="label">Category:</td><td>' + risCoreEscapeHtml_(entry.category) + '</td></tr>' +
    '</table>' +
    '<table class="items">' +
    '<thead><tr>' +
    '<th>Item Code</th><th>Item Description</th><th>Unit</th><th>Purchase Order #</th><th>Batch/Lot</th><th>Expiration</th><th>Qty Requested</th><th>Yes</th><th>No</th><th>Qty Issued</th><th>Unit Cost</th><th>Total Amount</th><th>Remarks</th>' +
    '</tr></thead><tbody>' + rows +
    '<tr class="total-row"><td colspan="11" class="num">TOTAL</td><td class="num">' + risCoreEscapeHtml_(grandTotal.toFixed(2)) + '</td><td></td></tr>' +
    '</tbody></table>' +
    '<div class="purpose"><b>Purpose:</b> ' + risCoreEscapeHtml_(entry.purpose || '') + '</div>' +
    '<table class="signatures">' +
    '<tr><td>Requested by:</td><td>Approved by:</td><td>Issued by:</td><td>Received by:</td></tr>' +
    '<tr><td>' + risCoreEscapeHtml_(entry.requestedBy) + '</td><td>' + risCoreEscapeHtml_(entry.approvedBy) + '</td><td>' + risCoreEscapeHtml_(entry.issuedBy) + '</td><td>' + risCoreEscapeHtml_(entry.receivedBy) + '</td></tr>' +
    '<tr><td class="small">Signature / Printed Name / Date</td><td class="small">Signature / Printed Name / Date</td><td class="small">Signature / Printed Name / Date</td><td class="small">Signature / Printed Name / Date</td></tr>' +
    '</table>' +
    '<p class="small" style="margin-top:10px;">Prepared by: ' + risCoreEscapeHtml_(user.fullName + ' / ' + user.username) + '</p>' +
    '</body></html>';
}

function risPdfSendNotification_(ss, entry, items, file, user) {
  const recipients = risCoreReadEmailRecipients_(ss);
  const to = recipients.to.slice();
  if (entry.requestorEmail) to.push(entry.requestorEmail);
  if (to.length === 0) return;

  MailApp.sendEmail({
    to: risCoreUnique_(to).join(','),
    cc: risCoreUnique_(recipients.cc).join(','),
    subject: RIS_PDF_CONFIG.notificationSubjectPrefix + (entry.risNo || entry.recordId),
    htmlBody: '<p>RIS PDF generated: <b>' + risCoreEscapeHtml_(entry.risNo || entry.recordId) + '</b></p>' +
      '<p>Generated by: ' + risCoreEscapeHtml_(user.fullName + ' (' + user.username + ')') + '</p>' +
      '<p><a href="' + risCoreEscapeHtml_(file.getUrl()) + '">Open RIS PDF</a></p>',
    body: 'RIS PDF generated: ' + (entry.risNo || entry.recordId) + '\n' + file.getUrl()
  });
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
