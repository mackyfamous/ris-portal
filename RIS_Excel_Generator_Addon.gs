// ==========================================
// RIS EXCEL GENERATOR ADD-ON
// ==========================================

const RIS_EXCEL_CONFIG = {
  templateFileId: 'PASTE_RIS_EXCEL_TEMPLATE_FILE_ID',
  outputFolderId: 'PASTE_EXCEL_OUTPUT_FOLDER_ID',
  notificationSubjectPrefix: '[RIS Excel] Generated: ',
  notificationCc: 'lyn4logistics@gmail.com',
  divisionName: 'CITY HEALTH DEPARTMENT',
  risNoLabelPrefix: 'RIS #: ',
  itemStartRow: 10,
  itemEndRow: 40,
  sheetName: '',
  stockAvailableMark: 'ü',
  stockUnavailableMark: 'ü',
  stockMarkFontFamily: 'Wingdings',
  itemFontSize: 60,
  itemRowHeight: 350,
  dateNumberFormat: 'mm-dd-yy',
  quantityNumberFormat: 'General',
  moneyNumberFormat: '_-[$₱-3409]* #,##0.00_-;\\-[$₱-3409]* #,##0.00_-;_-[$₱-3409]* "-"??_-;_-@',
  defaultSheetView: 'pageBreakPreview',
  nothingFollowsText: '***************** Nothing Follows ****************',
  xlsxMimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  headerCells: {
    office: 'C6',
    division: 'C7',
    risNo: 'I7',
    date: 'M7',
    purpose: 'C42',
    requestedBy: 'C46',
    approvedBy: 'D46',
    issuedBy: 'H46',
    receivedBy: 'L46',
    preparedBy: 'C49'
  },
  itemColumns: {
    itemCode: 2,
    itemDescription: 3,
    uom: 4,
    poWithSupplier: 5,
    batch: 6,
    expiry: 7,
    qtyRequested: 8,
    stockYes: 9,
    stockNo: 10,
    issuedQty: 11,
    unitCost: 12,
    totalCost: 13,
    remarks: 14
  },
  requiredFields: [
    'itemCode',
    'itemDescription',
    'uom',
    'qtyRequested',
    'issuedQty',
    'unitCost'
  ],
  recoverableFields: [
    'itemCode',
    'itemDescription',
    'uom',
    'poNumber',
    'supplier',
    'batch',
    'expiry',
    'unitCost',
    'remarks'
  ]
};

function risExcelAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('RIS Excel Tools')
    .addItem('Authorize Excel generator', 'risExcelAuthorize')
    .addItem('Generate Excel for selected RIS', 'risExcelGenerateForSelectedRIS')
    .addItem('Generate Excel by RIS No.', 'risExcelGenerateByRisNoPrompt')
    .addToUi();
}

function risExcelAuthorize() {
  risExcelAssertConfigured_();
  const ss = risCoreGetTransactionsSpreadsheet_();
  const recipients = risCoreReadEmailRecipients_(ss);
  const target = recipients.to[0] || Session.getActiveUser().getEmail();
  if (!target) throw new Error('Add an email address in Admin Emails or Client Emails before authorizing.');

  MailApp.sendEmail({
    to: target,
    subject: 'RIS Excel generator authorization test',
    body: 'This test email confirms that RIS Excel generation notifications are authorized.'
  });
}

function risExcelGenerateForSelectedRIS() {
  const ui = SpreadsheetApp.getUi();
  try {
    const risNo = risExcelGetSelectedRisNo_();
    const username = risExcelPromptUsername_('Generate Excel for ' + risNo);
    if (!username) return;

    const result = risExcelGenerateByRisNo_(risNo, username);
    ui.alert(risExcelResultMessage_(result));
  } catch (error) {
    ui.alert('Could not generate RIS Excel:\n\n' + risCoreErrorMessage_(error));
  }
}

function risExcelGenerateByRisNoPrompt() {
  const ui = SpreadsheetApp.getUi();
  const risResponse = ui.prompt('Generate RIS Excel', 'Enter RIS No. or Record ID:', ui.ButtonSet.OK_CANCEL);
  if (risResponse.getSelectedButton() !== ui.Button.OK) return;

  const risNo = risResponse.getResponseText().trim();
  if (!risNo) {
    ui.alert('Please enter a RIS No. or Record ID.');
    return;
  }

  const username = risExcelPromptUsername_('Generate Excel for ' + risNo);
  if (!username) return;

  try {
    const result = risExcelGenerateByRisNo_(risNo, username);
    ui.alert(risExcelResultMessage_(result));
  } catch (error) {
    ui.alert('Could not generate RIS Excel:\n\n' + risCoreErrorMessage_(error));
  }
}

function risExcelGenerateByRisNo(risNo, username) {
  return risExcelGenerateByRisNo_(risNo, username);
}

function risExcelGenerateByRisNo_(risNo, username) {
  risExcelAssertConfigured_();
  const user = risCoreValidateActiveUser_(username);
  const bundle = risCoreGetRisBundle_(risNo);
  const prepared = risExcelPrepareItems_(bundle.items);

  if (prepared.items.length === 0) {
    throw new Error('No RIS items were found for ' + risNo + '.');
  }

  if (prepared.items.length > RIS_EXCEL_CONFIG.itemEndRow - RIS_EXCEL_CONFIG.itemStartRow + 1) {
    throw new Error('This RIS has more item rows than the Excel template can fit.');
  }

  const workingSpreadsheet = risExcelCreateEditableTemplateCopy_(bundle.entry.risNo);
  let outputFile = null;

  try {
    const sheet = risExcelGetTargetSheet_(workingSpreadsheet);
    risExcelFillTemplate_(sheet, bundle.entry, prepared.items, user);
    SpreadsheetApp.flush();
    outputFile = risExcelExportToXlsx_(workingSpreadsheet, bundle.entry.risNo);

    risCoreUpdateRecordFields_(
      bundle.entriesSheet,
      bundle.entry.rowNumber,
      RIS_ENTRIES_DEFAULT_HEADERS,
      RIS_ENTRIES_ALIASES,
      {
        excelUrl: outputFile.getUrl(),
        excelGeneratedAt: new Date(),
        excelGeneratedBy: user.fullName + ' (' + user.username + ')'
      }
    );

    risExcelSendNotification_(bundle.ss, bundle.entry, prepared.items, outputFile, user, prepared.report);
  } finally {
    try {
      DriveApp.getFileById(workingSpreadsheet.getId()).setTrashed(true);
    } catch (cleanupError) {
      // Keep the generated file even if cleanup of the temporary Google Sheet fails.
    }
  }

  return {
    success: true,
    risNo: bundle.entry.risNo,
    url: outputFile.getUrl(),
    generatedBy: user.fullName,
    report: prepared.report
  };
}

function risExcelGetSelectedRisNo_() {
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

function risExcelPromptUsername_(title) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(title, 'Enter your username:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return '';
  return response.getResponseText().trim();
}

function risExcelAssertConfigured_() {
  if (!RIS_EXCEL_CONFIG.templateFileId || RIS_EXCEL_CONFIG.templateFileId.indexOf('PASTE_') === 0) {
    throw new Error('Configure RIS_EXCEL_CONFIG.templateFileId first.');
  }
  if (!RIS_EXCEL_CONFIG.outputFolderId || RIS_EXCEL_CONFIG.outputFolderId.indexOf('PASTE_') === 0) {
    throw new Error('Configure RIS_EXCEL_CONFIG.outputFolderId first.');
  }
}

function risExcelCreateEditableTemplateCopy_(risNo) {
  const title = risExcelSafeFileName_(risNo) + ' - RIS Excel Working Copy';
  const templateFile = DriveApp.getFileById(RIS_EXCEL_CONFIG.templateFileId);
  const folder = DriveApp.getFolderById(RIS_EXCEL_CONFIG.outputFolderId);

  if (templateFile.getMimeType() === MimeType.GOOGLE_SHEETS) {
    const copy = templateFile.makeCopy(title, folder);
    return SpreadsheetApp.openById(copy.getId());
  }

  if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.copy) {
    throw new Error('Enable the Advanced Drive service to use an Excel template file.');
  }

  const resource = {
    title: title,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: RIS_EXCEL_CONFIG.outputFolderId }]
  };
  const converted = Drive.Files.copy(resource, RIS_EXCEL_CONFIG.templateFileId);
  return SpreadsheetApp.openById(converted.id);
}

function risExcelGetTargetSheet_(spreadsheet) {
  if (RIS_EXCEL_CONFIG.sheetName) {
    const named = spreadsheet.getSheetByName(RIS_EXCEL_CONFIG.sheetName);
    if (!named) throw new Error('Template sheet not found: ' + RIS_EXCEL_CONFIG.sheetName);
    return named;
  }
  return spreadsheet.getSheets()[0];
}

function risExcelPrepareItems_(items) {
  const report = [];
  const prepared = items.map(function(item, index) {
    let recovered = {};
    try {
      recovered = risCoreRecoverInventoryDataForItem_(item);
    } catch (error) {
      report.push({
        row: index + 1,
        field: 'source inventory',
        status: 'missing',
        message: risCoreErrorMessage_(error)
      });
    }

    const finalItem = {};
    RIS_EXCEL_CONFIG.recoverableFields.forEach(function(field) {
      finalItem[field] = risExcelChooseValue_(item[field], recovered[field], report, index + 1, field);
    });

    finalItem.qtyRequested = risCoreParseNumber_(item.qtyRequested);
    finalItem.issuedQty = risCoreParseNumber_(item.issuedQty || item.qtyRequested);
    finalItem.currentSoh = risCoreParseNumber_(item.currentSoh || recovered.stock);
    finalItem.totalCost = finalItem.issuedQty * risCoreParseNumber_(finalItem.unitCost);
    finalItem.stockAvailable = finalItem.currentSoh > 0;
    finalItem.remarks = finalItem.remarks || item.remarks || '';

    RIS_EXCEL_CONFIG.requiredFields.forEach(function(field) {
      if (finalItem[field] === '' || finalItem[field] === null || finalItem[field] === undefined) {
        report.push({
          row: index + 1,
          field: field,
          status: 'missing',
          message: 'Required field is blank after database and inventory lookup.'
        });
      }
    });

    finalItem.poWithSupplier = risExcelPoWithSupplier_(finalItem);
    return finalItem;
  });

  const missingRequired = report.filter(function(row) {
    return row.status === 'missing' && RIS_EXCEL_CONFIG.requiredFields.indexOf(row.field) !== -1;
  });

  if (missingRequired.length > 0) {
    throw new Error('RIS Excel required fields are missing:\n' + risExcelReportLines_(missingRequired).join('\n'));
  }

  return { items: prepared, report: report };
}

function risExcelChooseValue_(databaseValue, inventoryValue, report, rowNumber, field) {
  const dbText = databaseValue === null || databaseValue === undefined ? '' : databaseValue;
  if (dbText !== '') {
    if (inventoryValue !== '' && inventoryValue !== null && inventoryValue !== undefined &&
        String(dbText) !== String(inventoryValue)) {
      report.push({
        row: rowNumber,
        field: field,
        status: 'mismatch',
        databaseValue: dbText,
        inventoryValue: inventoryValue,
        finalValue: dbText
      });
    }
    return dbText;
  }

  if (inventoryValue !== '' && inventoryValue !== null && inventoryValue !== undefined) {
    report.push({
      row: rowNumber,
      field: field,
      status: 'recovered',
      databaseValue: '',
      inventoryValue: inventoryValue,
      finalValue: inventoryValue
    });
    return inventoryValue;
  }

  return '';
}

function risExcelFillTemplate_(sheet, entry, items, user) {
  const cells = RIS_EXCEL_CONFIG.headerCells;
  sheet.getRange(cells.office).setValue(entry.deliveryLocation || entry.requestorProgram || '');
  sheet.getRange(cells.division).setValue(RIS_EXCEL_CONFIG.divisionName);
  sheet.getRange(cells.risNo).setValue(RIS_EXCEL_CONFIG.risNoLabelPrefix + (entry.risNo || entry.recordId || ''));
  sheet.getRange(cells.date).setValue(entry.deliveryDate || new Date());
  sheet.getRange(cells.date).setNumberFormat(RIS_EXCEL_CONFIG.dateNumberFormat);
  sheet.getRange(cells.purpose).setValue(entry.purpose || '');
  sheet.getRange(cells.requestedBy).setValue(entry.requestedBy || '');
  sheet.getRange(cells.approvedBy).setValue(entry.approvedBy || '');
  sheet.getRange(cells.issuedBy).setValue(entry.issuedBy || '');
  sheet.getRange(cells.receivedBy).setValue(entry.receivedBy || '');
  sheet.getRange(cells.preparedBy).setValue(user.fullName + ' / ' + user.username);

  const start = RIS_EXCEL_CONFIG.itemStartRow;
  const end = RIS_EXCEL_CONFIG.itemEndRow;
  const itemRange = sheet.getRange(start, 2, end - start + 1, 13);
  itemRange.clearContent();
  risExcelFormatItemArea_(sheet, start, end);

  items.forEach(function(item, index) {
    const row = start + index;
    sheet.getRange(row, 2, 1, 13).setValues([[
      item.itemCode,
      item.itemDescription,
      item.uom,
      item.poWithSupplier,
      item.batch,
      risExcelDateValue_(item.expiry),
      item.qtyRequested,
      item.stockAvailable ? RIS_EXCEL_CONFIG.stockAvailableMark : '',
      item.stockAvailable ? '' : RIS_EXCEL_CONFIG.stockUnavailableMark,
      item.issuedQty,
      item.unitCost,
      '',
      item.remarks
    ]]);
    sheet.getRange(row, RIS_EXCEL_CONFIG.itemColumns.totalCost).setFormula('=K' + row + '*L' + row);
  });

  risExcelPlaceNothingFollows_(sheet, start + items.length, end);
  sheet.getRange('M41')
    .setFormula('=SUM(M' + start + ':M' + end + ')')
    .setNumberFormat(RIS_EXCEL_CONFIG.moneyNumberFormat)
    .setFontSize(RIS_EXCEL_CONFIG.itemFontSize);
}

function risExcelFormatItemArea_(sheet, start, end) {
  const columns = RIS_EXCEL_CONFIG.itemColumns;
  const rowCount = end - start + 1;
  sheet.getRange(start, 2, rowCount, 13)
    .setFontSize(RIS_EXCEL_CONFIG.itemFontSize)
    .setVerticalAlignment('middle');
  sheet.setRowHeights(start, rowCount, RIS_EXCEL_CONFIG.itemRowHeight);
  sheet.getRange(start, columns.expiry, rowCount, 1).setNumberFormat(RIS_EXCEL_CONFIG.dateNumberFormat);
  sheet.getRange(start, columns.qtyRequested, rowCount, 1).setNumberFormat(RIS_EXCEL_CONFIG.quantityNumberFormat);
  sheet.getRange(start, columns.stockYes, rowCount, 1).setFontFamily(RIS_EXCEL_CONFIG.stockMarkFontFamily);
  sheet.getRange(start, columns.issuedQty, rowCount, 1).setNumberFormat(RIS_EXCEL_CONFIG.quantityNumberFormat);
  sheet.getRange(start, columns.unitCost, rowCount, 1).setNumberFormat(RIS_EXCEL_CONFIG.moneyNumberFormat);
  sheet.getRange(start, columns.totalCost, rowCount, 1).setNumberFormat(RIS_EXCEL_CONFIG.moneyNumberFormat);
}

function risExcelDateValue_(value) {
  const date = risCoreParseDate_(value);
  return date || value || '';
}

function risExcelPlaceNothingFollows_(sheet, markerRow, endRow) {
  if (markerRow > endRow) return;

  let movedDrawing = false;
  if (typeof sheet.getDrawings === 'function') {
    sheet.getDrawings().forEach(function(drawing) {
      const info = typeof drawing.getContainerInfo === 'function' ? drawing.getContainerInfo() : null;
      const anchorRow = info && typeof info.getAnchorRow === 'function' ? info.getAnchorRow() : 0;
      const anchorColumn = info && typeof info.getAnchorColumn === 'function' ? info.getAnchorColumn() : 0;
      const isItemAreaDrawing = (!anchorRow || anchorRow >= RIS_EXCEL_CONFIG.itemStartRow && anchorRow <= RIS_EXCEL_CONFIG.itemEndRow) &&
        (!anchorColumn || anchorColumn >= 2 && anchorColumn <= 14);
      if (!isItemAreaDrawing) return;

      try {
        drawing.setPosition(markerRow, RIS_EXCEL_CONFIG.itemColumns.itemDescription, 0, 0);
        movedDrawing = true;
      } catch (error) {
        // Fall back to a cell marker if Google Sheets cannot move the template textbox.
      }
    });
  }

  if (movedDrawing) return;

  sheet.getRange(markerRow, 2, 1, 13).clearContent();
  sheet.getRange(markerRow, RIS_EXCEL_CONFIG.itemColumns.itemDescription)
    .setValue(RIS_EXCEL_CONFIG.nothingFollowsText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function risExcelExportToXlsx_(spreadsheet, risNo) {
  const folder = DriveApp.getFolderById(RIS_EXCEL_CONFIG.outputFolderId);
  const fileName = risExcelSafeFileName_(risNo) + ' - RIS.xlsx';
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Excel export failed with HTTP ' + status + ': ' + response.getContentText().slice(0, 500));
  }

  const blob = risExcelBuildXlsxBlob_(response.getBlob(), fileName);
  return folder.createFile(blob);
}

function risExcelBuildXlsxBlob_(blob, fileName) {
  blob.setName(fileName);
  blob.setContentType(RIS_EXCEL_CONFIG.xlsxMimeType);

  if (RIS_EXCEL_CONFIG.defaultSheetView !== 'pageBreakPreview') return blob;

  try {
    const updatedParts = Utilities.unzip(blob).map(function(part) {
      const name = part.getName();
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) return part;

      const xml = part.getDataAsString('UTF-8');
      const updatedXml = risExcelSetPageBreakPreviewXml_(xml);
      if (updatedXml === xml) return part;
      return Utilities.newBlob(updatedXml, 'text/xml', name);
    });

    const updatedBlob = Utilities.zip(updatedParts, fileName);
    updatedBlob.setContentType(RIS_EXCEL_CONFIG.xlsxMimeType);
    return updatedBlob;
  } catch (error) {
    return blob;
  }
}

function risExcelSetPageBreakPreviewXml_(xml) {
  const sheetViewTag = xml.match(/<sheetView\b[^>]*\/?>/);
  if (sheetViewTag) {
    return xml.replace(sheetViewTag[0], risExcelSetXmlAttribute_(sheetViewTag[0], 'view', 'pageBreakPreview'));
  }

  const sheetViewsTag = xml.match(/<sheetViews\b[^>]*>/);
  if (sheetViewsTag) {
    return xml.replace(sheetViewsTag[0], sheetViewsTag[0] + '<sheetView workbookViewId="0" view="pageBreakPreview"/>');
  }

  return xml.replace(/<worksheet\b([^>]*)>/, '<worksheet$1><sheetViews><sheetView workbookViewId="0" view="pageBreakPreview"/></sheetViews>');
}

function risExcelSetXmlAttribute_(tag, name, value) {
  const attributePattern = new RegExp('\\s' + name + '="[^"]*"');
  if (attributePattern.test(tag)) {
    return tag.replace(attributePattern, ' ' + name + '="' + value + '"');
  }
  return tag.replace(/\/?>$/, function(end) {
    return ' ' + name + '="' + value + '"' + end;
  });
}

function risExcelSendNotification_(ss, entry, items, file, user, report) {
  const to = risCoreUnique_([risExcelGeneratedByEmail_(ss, user)]);
  if (to.length === 0) return;

  const warnings = report.filter(function(row) {
    return row.status === 'recovered' || row.status === 'mismatch' || row.status === 'missing';
  });

  const html = [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#172033;">',
    '<h2 style="margin:0 0 12px;color:#0f766e;">RIS Excel Generated</h2>',
    '<p>The RIS Excel file was generated.</p>',
    '<p><b>RIS No.:</b> ' + risCoreEscapeHtml_(entry.risNo || entry.recordId) + '</p>',
    '<p><b>Generated by:</b> ' + risCoreEscapeHtml_(user.fullName + ' (' + user.username + ')') + '</p>',
    '<p><b>File:</b> <a href="' + risCoreEscapeHtml_(file.getUrl()) + '">Open RIS Excel</a></p>',
    '<p><b>Items:</b> ' + items.length + '</p>',
    warnings.length ? '<p><b>Validation notes:</b><br>' + risCoreEscapeHtml_(risExcelReportLines_(warnings).join('\n')).replace(/\n/g, '<br>') + '</p>' : '',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: to.join(','),
    cc: risCoreUnique_([RIS_EXCEL_CONFIG.notificationCc]).join(','),
    subject: RIS_EXCEL_CONFIG.notificationSubjectPrefix + (entry.risNo || entry.recordId),
    htmlBody: html,
    body: 'RIS Excel generated: ' + (entry.risNo || entry.recordId) + '\n' + file.getUrl()
  });
}

function risExcelGeneratedByEmail_(ss, user) {
  const adminEmail = risExcelFindAdminEmail_(ss, user);
  if (adminEmail) return adminEmail;

  const username = String(user.username || '').trim();
  if (risExcelLooksLikeEmail_(username)) return username;

  const activeEmail = Session.getActiveUser().getEmail();
  return activeEmail || '';
}

function risExcelFindAdminEmail_(ss, user) {
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
      return risExcelFirstEmail_(rowRecipient) || risExcelFirstEmail_(rowSender) ||
        (risExcelLooksLikeEmail_(rowUsername) ? String(rowUsername).trim() : '');
    }
  }

  return '';
}

function risExcelFirstEmail_(value) {
  const emails = risCoreSplitEmailList_(value).filter(risExcelLooksLikeEmail_);
  return emails[0] || '';
}

function risExcelLooksLikeEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function risExcelResultMessage_(result) {
  const report = result.report || [];
  const notes = report.filter(function(row) {
    return row.status === 'recovered' || row.status === 'mismatch' || row.status === 'missing';
  });
  let message = 'RIS Excel generated.\n\nRIS No.: ' + result.risNo + '\nFile: ' + result.url;
  if (notes.length > 0) {
    message += '\n\nValidation notes:\n' + risExcelReportLines_(notes).slice(0, 12).join('\n');
    if (notes.length > 12) message += '\n...and ' + (notes.length - 12) + ' more.';
  }
  return message;
}

function risExcelReportLines_(rows) {
  return rows.map(function(row) {
    return 'Item ' + row.row + ' - ' + row.field + ': ' + row.status +
      (row.message ? ' (' + row.message + ')' : '');
  });
}

function risExcelPoWithSupplier_(item) {
  const po = String(item.poNumber || '').trim();
  const supplier = String(item.supplier || '').trim();
  if (po && supplier) return po + ' (' + supplier + ')';
  return po || supplier;
}

function risExcelSafeFileName_(value) {
  return String(value || 'RIS')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}
