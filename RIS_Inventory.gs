// ==========================================
// RIS SOURCE INVENTORY ROUTING
// ==========================================

const RIS_SOURCES_DEFAULT_HEADERS = [
  'Enabled', 'Button Name', 'Source Spreadsheet ID', 'Source Sheet Name', 'Header Row', 'Display Order', 'Category'
];

const RIS_INVENTORY_ALIASES = {
  program: ['Program', 'Program / End User', 'End User', 'Remarks / Program'],
  itemCode: ['Item Code', 'Code'],
  itemDescription: ['Item Description', 'Description'],
  uom: ['UOM', 'Unit of Measure', 'Unit of Measurement', 'Unit  of Measure'],
  poNumber: ['PO / PTR', 'PO/PTR Number', 'PO / PTR Number', 'P.O / PTR #', 'P.O.#/PTR', 'Purchase Order Number', 'PO Number', 'P.O. Number'],
  supplier: ['Supplier'],
  batch: ['Batch / Lot No.', 'Batch / Lot No', 'Batch / Lot Number', 'Batch / Lot #', 'Lot / Batch Number', 'Lot / Batch #', 'Lot Number'],
  expiry: ['Expiry Date', 'Expiration Date', 'Expiration Date (MMDDYYYY)', 'Expiry'],
  deliveryDate: ['Date Received / Delivery Date', 'Date Received', 'Delivery Date', 'Delivey Date'],
  stock: ['Stock On Hand / Balance', 'Stock On Hand', 'SOH', 'Physical Count', 'Balance', 'Ending Balance', 'Ending Bal.'],
  unitCost: ['Unit Cost'],
  totalCost: ['Total Cost', 'Total Amount', 'Total Price', 'Total Price (Based on Physical Count)'],
  remarks: ['Remarks', 'Remarks / Program']
};

function risEnsureSourcesSheet_(ss) {
  const sheet = risCoreGetOrCreateSheet_(ss, RIS_CONFIG.sourcesSheetName);
  risCoreEnsureColumns_(sheet, RIS_SOURCES_DEFAULT_HEADERS);

  if (sheet.getLastRow() === 1) {
    sheet.getRange(2, 1, RIS_DEFAULT_SOURCES.length, RIS_SOURCES_DEFAULT_HEADERS.length).setValues(RIS_DEFAULT_SOURCES);
  }

  return sheet;
}

function risGetSources_(ss) {
  const sheet = risEnsureSourcesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const columns = {
    enabled: risFindHeaderColumn_(headers, ['Enabled']),
    buttonName: risFindHeaderColumn_(headers, ['Button Name', 'RIS Button Name', 'Button']),
    sourceSpreadsheetId: risFindHeaderColumn_(headers, ['Source Spreadsheet ID', 'Spreadsheet ID']),
    sourceSheetName: risFindHeaderColumn_(headers, ['Source Sheet Name', 'Sheet Name', 'Tab Name']),
    headerRow: risFindHeaderColumn_(headers, ['Header Row']),
    displayOrder: risFindHeaderColumn_(headers, ['Display Order', 'Order']),
    category: risFindHeaderColumn_(headers, ['Category'])
  };

  return values.slice(1).map(function(row) {
    const buttonName = String(row[columns.buttonName - 1] || '').trim();
    const category = String(row[columns.category - 1] || buttonName).trim();
    return {
      sourceKey: risSlug_(category || buttonName),
      enabled: risCoreTruthy_(row[columns.enabled - 1]),
      buttonName: buttonName,
      sourceSpreadsheetId: risNormalizeSpreadsheetId_(row[columns.sourceSpreadsheetId - 1]),
      sourceSheetName: String(row[columns.sourceSheetName - 1] || '').trim(),
      headerRow: row[columns.headerRow - 1] || 'auto',
      displayOrder: risCoreParseNumber_(row[columns.displayOrder - 1]) || 999,
      category: category
    };
  }).filter(function(source) {
    return source.enabled && source.buttonName && source.sourceSheetName;
  }).sort(function(a, b) {
    return a.displayOrder - b.displayOrder || a.buttonName.localeCompare(b.buttonName);
  });
}

function risPublicSource_(source) {
  return {
    sourceKey: source.sourceKey,
    buttonName: source.buttonName,
    category: source.category,
    sourceSheetName: source.sourceSheetName,
    displayOrder: source.displayOrder
  };
}

function risFindSourceByKey_(sources, sourceKey) {
  const key = risSlug_(sourceKey);
  return sources.find(function(source) {
    return source.sourceKey === key || risSlug_(source.buttonName) === key || risSlug_(source.category) === key;
  });
}

function risReadInventoryRows_(source, selectedProgram) {
  const ss = risCoreOpenSourceSpreadsheet_(source);
  const sheet = ss.getSheetByName(source.sourceSheetName);
  if (!sheet) throw new Error('Source sheet not found: ' + source.sourceSheetName);

  const layout = risGetInventoryLayout_(sheet, source);
  const lastRow = sheet.getLastRow();
  const rows = [];
  const statusCounts = { available: 0, low: 0, out: 0 };
  const selectedKey = risCoreNormalizeText_(selectedProgram);

  if (lastRow <= layout.headerRow) return { rows: rows, statusCounts: statusCounts };

  const values = sheet.getRange(layout.headerRow + 1, 1, lastRow - layout.headerRow, sheet.getLastColumn()).getValues();
  values.forEach(function(row, offset) {
    const rowNumber = layout.headerRow + 1 + offset;
    const item = risMapInventoryRow_(row, layout, source, rowNumber);
    if (!item.itemDescription) return;
    if (selectedKey && risCoreNormalizeText_(item.program) !== selectedKey) return;

    item.stock = risCoreParseNumber_(item.stock);
    item.unitCost = risCoreParseNumber_(item.unitCost);
    item.totalCost = risCoreParseNumber_(item.totalCost);
    item.expiry = risCoreFormatDate_(item.expiry);
    item.deliveryDate = risCoreFormatDate_(item.deliveryDate);
    item.sourceRow = rowNumber;
    item.sourceHeaderRow = layout.headerRow;
    item.sourceSheet = source.sourceSheetName;
    item.sourceSpreadsheetId = source.sourceSpreadsheetId;
    item.status = risStockStatus_(item.stock);
    item.statusLabel = risStockStatusLabel_(item.status);
    statusCounts[item.status] += 1;
    rows.push(item);
  });

  return { rows: rows, statusCounts: statusCounts };
}

function risGetProgramsForSource_(source) {
  const inventory = risReadInventoryRows_(source, '');
  const seen = {};
  inventory.rows.forEach(function(item) {
    if (item.program) seen[String(item.program).trim()] = true;
  });
  return Object.keys(seen).sort();
}

function risGetInventoryLayout_(sheet, source) {
  const headerRow = risResolveHeaderRow_(sheet, source);
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columns = {};

  Object.keys(RIS_INVENTORY_ALIASES).forEach(function(field) {
    columns[field] = risFindHeaderColumn_(headers, RIS_INVENTORY_ALIASES[field]);
  });

  ['itemDescription', 'uom', 'stock', 'unitCost'].forEach(function(field) {
    if (!columns[field]) throw new Error('Could not find inventory column: ' + field + ' in ' + sheet.getName());
  });

  return { headerRow: headerRow, headers: headers, columns: columns };
}

function risResolveHeaderRow_(sheet, source) {
  const fixed = risCoreParseInteger_(source.headerRow);
  if (fixed) return fixed;

  const maxRows = Math.min(sheet.getLastRow(), RIS_CONFIG.headerScanRows);
  let bestRow = 1;
  let bestScore = -1;

  for (let row = 1; row <= maxRows; row++) {
    const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    let score = 0;
    Object.keys(RIS_INVENTORY_ALIASES).forEach(function(field) {
      if (risFindHeaderColumn_(values, RIS_INVENTORY_ALIASES[field])) score += 1;
    });
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }

  return bestRow;
}

function risMapInventoryRow_(row, layout, source, rowNumber) {
  return {
    rowNumber: rowNumber,
    sourceKey: source.sourceKey,
    category: source.category,
    program: risCell_(row, layout.columns.program),
    itemCode: risCell_(row, layout.columns.itemCode),
    itemDescription: risCell_(row, layout.columns.itemDescription),
    uom: risCell_(row, layout.columns.uom),
    poNumber: risCleanText_(risCell_(row, layout.columns.poNumber)),
    supplier: risCell_(row, layout.columns.supplier),
    batch: risCell_(row, layout.columns.batch),
    expiry: risCell_(row, layout.columns.expiry),
    deliveryDate: risCell_(row, layout.columns.deliveryDate),
    stock: risCell_(row, layout.columns.stock),
    unitCost: risCell_(row, layout.columns.unitCost),
    totalCost: risCell_(row, layout.columns.totalCost),
    remarks: risCell_(row, layout.columns.remarks)
  };
}

function risGenerateNextRisNo_(sheet, date) {
  risCoreEnsureColumns_(sheet, RIS_ENTRIES_DEFAULT_HEADERS);
  const info = risCoreGetHeaderInfo_(sheet, RIS_ENTRIES_DEFAULT_HEADERS, RIS_ENTRIES_ALIASES);
  const year = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy');
  const month = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM');
  const prefix = 'RIS-' + year + '-' + month + '-';
  let max = 0;

  if (sheet.getLastRow() > 1 && info.columns.risNo) {
    sheet.getRange(2, info.columns.risNo, sheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
      const value = String(row[0] || '');
      if (value.indexOf(prefix) === 0) max = Math.max(max, parseInt(value.slice(prefix.length), 10) || 0);
    });
  }

  return prefix + String(max + 1).padStart(3, '0');
}

function risGetHolidayDates_(ss) {
  const sheet = ss.getSheetByName(RIS_CONFIG.holidaysSheetName);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(function(row) {
    return risCoreDateKey_(row[0]);
  }).filter(String);
}

function risIsAllowedDeliveryDate_(date, holidays) {
  const day = date.getDay();
  if (day === 0 || day === 5 || day === 6) return false;
  if (holidays.indexOf(risCoreDateKey_(date)) !== -1) return false;
  return risCoreDateOnly_(date).getTime() >= risAddAllowedWorkingDays_(new Date(), RIS_CONFIG.minimumWorkingDays, holidays).getTime();
}

function risAddAllowedWorkingDays_(startDate, days, holidays) {
  const date = risCoreDateOnly_(startDate);
  let added = 0;
  while (added < days) {
    date.setDate(date.getDate() + 1);
    const day = date.getDay();
    if (day !== 0 && day !== 5 && day !== 6 && holidays.indexOf(risCoreDateKey_(date)) === -1) added += 1;
  }
  return date;
}

function risStockStatus_(stock) {
  if (stock <= 0) return 'out';
  if (stock <= RIS_CONFIG.lowStockThreshold) return 'low';
  return 'available';
}

function risStockStatusLabel_(status) {
  if (status === 'out') return 'Out of Stock';
  if (status === 'low') return 'Low Stock';
  return 'Available';
}
