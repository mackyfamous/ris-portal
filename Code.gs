// ==========================================
// RIS PORTAL WEB APP ENTRYPOINTS
// ==========================================

const RIS_CONFIG = {
  testingMode: true,
  autoCreateSheets: false,
  transactionsSheetId: 'PASTE_TRANSACTIONS_SHEET_ID',
  inventorySheetId: 'PASTE_INVENTORY_SHEET_ID',
  portalDisplayName: 'Requisition Issuance Slip Portal',
  sourcesSheetName: 'RIS Sources',
  entriesSheetName: 'RIS Entries',
  itemsSheetName: 'RIS Items',
  adminEmailsSheetName: 'Admin Emails',
  clientEmailsSheetName: 'Client Emails',
  legacyEmailsSheetNames: ['Emails', 'emails', 'EMAILS'],
  usersSheetName: 'Users',
  minimumWorkingDays: 5,
  lowStockThreshold: 10,
  headerScanRows: 10
};

const RIS_DEFAULT_SOURCES = [
  ['TRUE', 'Medicine', 'same', 'DRUGS and MEDICINES', 'auto', 1, 'Medicine'],
  ['TRUE', 'Supplies', 'same', '2026 SUPPLIES', 'auto', 2, 'Supplies']
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(RIS_CONFIG.portalDisplayName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RIS Portal')
    .addItem('Setup required sheets', 'setupRISPortalSheets')
    .addItem('Authorize email notifications', 'authorizeEmailNotifications')
    .addItem('Show next RIS number', 'showNextRisNo')
    .addToUi();

  if (typeof risExcelAddMenu_ === 'function') risExcelAddMenu_();
  if (typeof risPdfAddMenu_ === 'function') risPdfAddMenu_();
  if (typeof risStatusAddMenu_ === 'function') risStatusAddMenu_();
}

function setupRISPortalSheets() {
  const ss = risCoreGetTransactionsSpreadsheet_();
  risCoreEnsureTransactionSheets_(ss, true);
  risEnsureSourcesSheet_(ss, true);
  SpreadsheetApp.getUi().alert('RIS Portal sheets are ready.');
}

function showNextRisNo() {
  const ss = risCoreGetTransactionsSpreadsheet_();
  risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);
  const sheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName);
  SpreadsheetApp.getUi().alert('Next RIS No.: ' + risGenerateNextRisNo_(sheet, new Date(), RIS_CONFIG.autoCreateSheets));
}

function authorizeEmailNotifications() {
  const ss = risCoreGetTransactionsSpreadsheet_();
  const recipients = risCoreReadEmailRecipients_(ss);
  const target = recipients.to[0] || Session.getActiveUser().getEmail();
  if (!target) throw new Error('Add at least one recipient in Admin Emails or Client Emails before authorizing.');

  MailApp.sendEmail({
    to: target,
    subject: 'RIS email notification authorization test',
    body: 'This test email confirms that RIS email notifications are authorized.'
  });
}

function getInitialData() {
  try {
    const ss = risCoreGetTransactionsSpreadsheet_();
    risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);

    const sources = risGetSources_(ss, RIS_CONFIG.autoCreateSheets);
    const programsBySource = {};
    const sourceWarnings = {};

    sources.forEach(function(source) {
      try {
        programsBySource[source.sourceKey] = risGetProgramsForSource_(source);
      } catch (error) {
        programsBySource[source.sourceKey] = [];
        sourceWarnings[source.sourceKey] = risCoreErrorMessage_(error);
      }
    });

    return {
      success: true,
      portalName: RIS_CONFIG.portalDisplayName,
      testingMode: RIS_CONFIG.testingMode,
      lowStockThreshold: RIS_CONFIG.lowStockThreshold,
      minimumWorkingDays: RIS_CONFIG.minimumWorkingDays,
      sources: sources.map(risPublicSource_),
      programsBySource: programsBySource,
      sourceWarnings: sourceWarnings,
      holidays: risGetHolidayDates_(ss),
      nextRisNo: risGenerateNextRisNo_(risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName), new Date(), RIS_CONFIG.autoCreateSheets)
    };
  } catch (error) {
    return { success: false, error: risCoreErrorMessage_(error) };
  }
}

function getNextRisNo() {
  try {
    const ss = risCoreGetTransactionsSpreadsheet_();
    risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);
    const sheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName);
    return { success: true, risNo: risGenerateNextRisNo_(sheet, new Date(), RIS_CONFIG.autoCreateSheets) };
  } catch (error) {
    return { success: false, error: risCoreErrorMessage_(error) };
  }
}

function getInventoryData(sourceKey, selectedProgram) {
  try {
    const ss = risCoreGetTransactionsSpreadsheet_();
    risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);
    const source = risFindSourceByKey_(risGetSources_(ss, RIS_CONFIG.autoCreateSheets), sourceKey);
    if (!source) throw new Error('Please select Medicine or Supplies first.');

    const inventory = risReadInventoryRows_(source, selectedProgram || '');
    return {
      success: true,
      source: risPublicSource_(source),
      selectedProgram: selectedProgram || '',
      data: inventory.rows,
      statusCounts: inventory.statusCounts,
      warnings: inventory.warnings || []
    };
  } catch (error) {
    return { success: false, error: risCoreErrorMessage_(error) };
  }
}

function submitTransaction(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'Another RIS request is being saved. Please try again.' };
  }

  try {
    const ss = risCoreGetTransactionsSpreadsheet_();
    risCoreEnsureTransactionSheets_(ss, RIS_CONFIG.autoCreateSheets);

    const source = risFindSourceByKey_(risGetSources_(ss, RIS_CONFIG.autoCreateSheets), payload && payload.sourceKey);
    if (!source) throw new Error('Choose Medicine or Supplies.');

    const entriesSheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.entriesSheetName);
    const itemsSheet = risCoreGetRequiredSheet_(ss, RIS_CONFIG.itemsSheetName);
    const now = new Date();
    const risNo = risGenerateNextRisNo_(entriesSheet, now);
    const items = risValidateAndPrepareItems_(payload, source);
    const grandTotal = items.reduce(function(sum, item) {
      return sum + risCoreParseNumber_(item.totalCost);
    }, 0);
    const deliveryDate = risCoreParseDate_(payload.deliveryDate);

    if (!deliveryDate || !risIsAllowedDeliveryDate_(deliveryDate, risGetHolidayDates_(ss))) {
      throw new Error('Choose an allowed delivery date.');
    }

    const friendlyInventorySource = source.buttonName || source.category || source.sourceSheetName;
    const entry = {
      timestamp: now,
      recordId: risNo,
      risNo: risNo,
      category: source.category,
      inventorySource: friendlyInventorySource,
      deliveryDate: deliveryDate,
      requestorProgram: risRequiredText_(payload.requestorProgram, 'Requestor Program'),
      requestorEmail: risRequiredText_(payload.requestorEmail, 'Requestor Email'),
      deliveryLocation: risRequiredText_(payload.deliveryLocation, 'Delivery Location / Office'),
      requestedBy: risRequiredText_(payload.requestedBy, 'Requested By'),
      approvedBy: risRequiredText_(payload.approvedBy, 'Approved By'),
      issuedBy: payload.issuedBy || '',
      receivedBy: risRequiredText_(payload.receivedBy, 'Received By'),
      purpose: payload.purpose || '',
      grandTotal: grandTotal,
      risStatus: 'Submitted'
    };

    risCoreAppendRecord_(entriesSheet, RIS_ENTRIES_DEFAULT_HEADERS, RIS_ENTRIES_ALIASES, RIS_ENTRIES_FIELD_ORDER, entry);

    items.forEach(function(item, index) {
      item.timestamp = now;
      item.itemId = risNo + '-' + String(index + 1).padStart(3, '0');
      item.recordId = risNo;
      item.risNo = risNo;
      item.category = source.category;
      item.inventorySource = friendlyInventorySource;
      risCoreAppendRecord_(itemsSheet, RIS_ITEMS_DEFAULT_HEADERS, RIS_ITEMS_ALIASES, RIS_ITEMS_FIELD_ORDER, item);
    });

    risSendSubmissionEmails_(ss, entry, items);
    return {
      success: true,
      risNo: risNo,
      recordId: risNo,
      grandTotal: grandTotal,
      testingMode: RIS_CONFIG.testingMode
    };
  } catch (error) {
    return { success: false, error: risCoreErrorMessage_(error) };
  } finally {
    lock.releaseLock();
  }
}

function risValidateAndPrepareItems_(payload, source) {
  if (!payload || !payload.items || payload.items.length === 0) throw new Error('Add at least one item.');

  const cache = {};
  return payload.items.map(function(item, index) {
    const sourceMember = risFindSourceMemberForItem_(source, item);
    const context = risGetValidationSourceContext_(sourceMember, cache);
    const sourceRow = risCoreParseInteger_(item.sourceRow);

    if (!sourceRow || sourceRow <= context.layout.headerRow) {
      throw new Error('Item ' + (index + 1) + ' is missing its inventory row.');
    }

    const current = risMapInventoryRow_(
      context.sheet.getRange(sourceRow, 1, 1, context.sheet.getLastColumn()).getValues()[0],
      context.layout,
      context.source,
      sourceRow
    );
    const qtyRequested = risCoreParseNumber_(item.qtyRequested);
    const issuedQty = risCoreParseNumber_(item.issuedQty || qtyRequested);
    const currentStock = risCoreParseNumber_(current.stock);
    const unitCost = risCoreParseNumber_(current.unitCost);

    if (!qtyRequested || qtyRequested <= 0) throw new Error('Enter a quantity for item ' + (index + 1) + '.');
    if (currentStock <= 0) throw new Error('Item "' + current.itemDescription + '" is out of stock.');
    if (qtyRequested > currentStock) throw new Error('Requested quantity for "' + current.itemDescription + '" is greater than available stock.');

    if (!RIS_CONFIG.testingMode) {
      context.sheet.getRange(sourceRow, context.layout.columns.stock).setValue(currentStock - qtyRequested);
    }

    return {
      itemCode: current.itemCode,
      itemDescription: current.itemDescription,
      uom: current.uom,
      poNumber: current.poNumber,
      supplier: current.supplier,
      batch: current.batch,
      expiry: current.expiry,
      currentSoh: currentStock,
      qtyRequested: qtyRequested,
      issuedQty: issuedQty,
      unitCost: unitCost,
      totalCost: issuedQty * unitCost,
      remarks: item.remarks || current.remarks || '',
      sourceSpreadsheetId: context.sourceSpreadsheetId,
      sourceSheet: context.sourceSheetName,
      sourceRow: sourceRow,
      sourceHeaderRow: context.layout.headerRow
    };
  });
}

function risGetValidationSourceContext_(source, cache) {
  const key = [risNormalizeSpreadsheetId_(source.sourceSpreadsheetId), source.sourceSheetName, source.headerRow].join('::');
  if (!cache[key]) {
    const spreadsheet = risCoreOpenSourceSpreadsheet_(source);
    const sheet = risResolveSourceSheet_(spreadsheet, source);
    const resolvedSource = Object.assign({}, source, { sourceSheetName: sheet.getName() });

    cache[key] = {
      spreadsheet: spreadsheet,
      sourceSpreadsheetId: spreadsheet.getId(),
      sourceSheetName: sheet.getName(),
      source: resolvedSource,
      sheet: sheet,
      layout: risGetInventoryLayout_(sheet, resolvedSource)
    };
  }
  return cache[key];
}

function risSendSubmissionEmails_(ss, entry, items) {
  const recipients = risCoreReadEmailRecipients_(ss);
  const to = recipients.to.slice();
  if (entry.requestorEmail) to.push(entry.requestorEmail);
  if (to.length === 0) return;

  const itemRows = items.map(function(item) {
    return '<tr><td>' + risCoreEscapeHtml_(item.itemCode) + '</td><td>' +
      risCoreEscapeHtml_(item.itemDescription) + '</td><td>' + risCoreEscapeHtml_(item.uom) +
      '</td><td style="text-align:right;">' + risCoreEscapeHtml_(item.qtyRequested) +
      '</td><td style="text-align:right;">' + risCoreEscapeHtml_(risCoreMoney_(item.totalCost)) + '</td></tr>';
  }).join('');

  MailApp.sendEmail({
    to: risCoreUnique_(to).join(','),
    cc: risCoreUnique_(recipients.cc).join(','),
    subject: '[RIS] Submitted: ' + entry.risNo,
    htmlBody: '<h2>New RIS Request Submitted</h2>' +
      '<p><b>RIS No.:</b> ' + risCoreEscapeHtml_(entry.risNo) + '</p>' +
      '<p><b>Category:</b> ' + risCoreEscapeHtml_(entry.category) + '</p>' +
      '<p><b>Program:</b> ' + risCoreEscapeHtml_(entry.requestorProgram) + '</p>' +
      '<table border="1" cellpadding="6" cellspacing="0"><tr><th>Code</th><th>Description</th><th>UOM</th><th>Qty</th><th>Total</th></tr>' +
      itemRows + '</table><p><b>Grand Total:</b> ' + risCoreEscapeHtml_(risCoreMoney_(entry.grandTotal)) + '</p>',
    body: 'New RIS request submitted: ' + entry.risNo
  });
}
