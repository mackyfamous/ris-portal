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
  const rawSources = risGetRawSources_(ss);
  const groups = {};

  rawSources.forEach(function(source) {
    if (!groups[source.sourceKey]) {
      groups[source.sourceKey] = {
        sourceKey: source.sourceKey,
        buttonName: source.buttonName,
        category: source.category,
        displayOrder: source.displayOrder,
        sourceSheetName: '',
        sourceSheetNames: [],
        sourceCount: 0,
        sources: []
      };
    }

    const group = groups[source.sourceKey];
    group.sources.push(source);
    group.displayOrder = Math.min(group.displayOrder, source.displayOrder);
    if (!group.buttonName && source.buttonName) group.buttonName = source.buttonName;
    if (!group.category && source.category) group.category = source.category;
  });

  return Object.keys(groups).map(function(key) {
    const group = groups[key];
    group.sources.sort(risCompareSources_);
    group.sourceSheetNames = group.sources.map(function(source) { return source.sourceSheetName; });
    group.sourceCount = group.sources.length;
    group.sourceSheetName = group.sourceCount === 1 ? group.sourceSheetNames[0] : group.sourceCount + ' inventory sheets';
    return group;
  }).sort(risCompareSources_);
}

function risGetRawSources_(ss) {
  const sheet = risEnsureSourcesSheet_(ss);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

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

  return values.slice(1).map(function(row, index) {
    const sourceSheetName = columns.sourceSheetName ? String(row[columns.sourceSheetName - 1] || '').trim() : '';
    const category = columns.category ? String(row[columns.category - 1] || '').trim() : '';
    const buttonName = (columns.buttonName ? String(row[columns.buttonName - 1] || '').trim() : '') || category || sourceSheetName;
    const displayOrder = (columns.displayOrder ? risCoreParseNumber_(row[columns.displayOrder - 1]) : 0) || index + 1;

    return {
      sourceKey: risSlug_(buttonName || category || sourceSheetName),
      sourceId: risSlug_([buttonName, sourceSheetName, index + 2].join(' ')),
      enabled: columns.enabled ? risCoreTruthy_(row[columns.enabled - 1]) : true,
      buttonName: buttonName,
      sourceSpreadsheetId: risNormalizeSpreadsheetId_(columns.sourceSpreadsheetId ? row[columns.sourceSpreadsheetId - 1] : 'same'),
      sourceSheetName: sourceSheetName,
      headerRow: columns.headerRow ? (row[columns.headerRow - 1] || 'auto') : 'auto',
      displayOrder: displayOrder,
      category: category || buttonName,
      configRow: index + 2
    };
  }).filter(function(source) {
    return source.enabled && source.sourceKey && source.sourceSheetName;
  }).sort(risCompareSources_);
}

function risCompareSources_(a, b) {
  return (a.displayOrder || 999) - (b.displayOrder || 999) || String(a.buttonName || '').localeCompare(String(b.buttonName || ''));
}

function risPublicSource_(source) {
  return {
    sourceKey: source.sourceKey,
    buttonName: source.buttonName,
    category: source.category,
    sourceSheetName: source.sourceSheetName,
    sourceSheetNames: source.sourceSheetNames || [source.sourceSheetName],
    sourceCount: source.sourceCount || risSourceMembers_(source).length,
    displayOrder: source.displayOrder
  };
}

function risFindSourceByKey_(sources, sourceKey) {
  const key = risSlug_(sourceKey);
  return sources.find(function(source) {
    return source.sourceKey === key || risSlug_(source.buttonName) === key || risSlug_(source.category) === key;
  });
}

function risSourceMembers_(source) {
  if (!source) return [];
  return source.sources && source.sources.length ? source.sources : [source];
}

function risFindSourceMemberForItem_(source, item) {
  const members = risSourceMembers_(source);
  const wantedSheet = risNormalizeSheetName_(item && item.sourceSheet);
  const wantedId = risCoreNormalizeText_(risNormalizeSpreadsheetId_(item && item.sourceSpreadsheetId));

  if (wantedSheet) {
    const match = members.find(function(member) {
      if (risNormalizeSheetName_(member.sourceSheetName) !== wantedSheet) return false;
      if (!wantedId || wantedId === 'same') return true;
      const configuredId = risCoreNormalizeText_(risNormalizeSpreadsheetId_(member.sourceSpreadsheetId));
      return configuredId === 'same' || configuredId === wantedId;
    });
    if (match) return match;
  }

  if (members.length === 1) return members[0];

  throw new Error('The selected item does not include a valid source sheet for ' + (source.buttonName || source.category || 'this source') + '.');
}

function risResolveSourceSheet_(ss, source) {
  const configuredName = String(source && source.sourceSheetName || '').trim();
  const exact = ss.getSheetByName(configuredName);
  if (exact) return exact;

  const configuredKey = risNormalizeSheetName_(configuredName);
  const sheets = ss.getSheets();
  const normalized = sheets.find(function(sheet) {
    return risNormalizeSheetName_(sheet.getName()) === configuredKey;
  });
  if (normalized) return normalized;

  const available = sheets.map(function(sheet) { return sheet.getName(); }).join(', ') || 'none';
  throw new Error('Source sheet not found: ' + configuredName + '. Available tabs: ' + available);
}

function risNormalizeSheetName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function risReadInventoryRows_(source, selectedProgram) {
  const rows = [];
  const warnings = [];
  const statusCounts = { available: 0, low: 0, out: 0 };
  const members = risSourceMembers_(source);

  members.forEach(function(member) {
    try {
      const result = risReadInventoryRowsForSingleSource_(member, selectedProgram);
      rows.push.apply(rows, result.rows);
      statusCounts.available += result.statusCounts.available || 0;
      statusCounts.low += result.statusCounts.low || 0;
      statusCounts.out += result.statusCounts.out || 0;
    } catch (error) {
      warnings.push(member.sourceSheetName + ': ' + risCoreErrorMessage_(error));
    }
  });

  if (rows.length === 0 && warnings.length > 0 && warnings.length === members.length) {
    throw new Error('Could not load inventory for ' + (source.buttonName || source.category || 'selected source') + '. ' + warnings.join(' | '));
  }

  return { rows: rows, statusCounts: statusCounts, warnings: warnings };
}

function risReadInventoryRowsForSingleSource_(source, selectedProgram) {
  const ss = risCoreOpenSourceSpreadsheet_(source);
  const sourceSpreadsheetId = ss.getId();
  const sheet = risResolveSourceSheet_(ss, source);
  const resolvedSource = Object.assign({}, source, { sourceSheetName: sheet.getName() });

  const layout = risGetInventoryLayout_(sheet, resolvedSource);
  const lastRow = sheet.getLastRow();
  const rows = [];
  const statusCounts = { available: 0, low: 0, out: 0 };
  const selectedKey = risCoreNormalizeText_(selectedProgram);

  if (lastRow <= layout.headerRow) return { rows: rows, statusCounts: statusCounts };

  const values = sheet.getRange(layout.headerRow + 1, 1, lastRow - layout.headerRow, sheet.getLastColumn()).getValues();
  values.forEach(function(row, offset) {
    const rowNumber = layout.headerRow + 1 + offset;
    const item = risMapInventoryRow_(row, layout, resolvedSource, rowNumber);
    if (!item.itemDescription) return;
    if (selectedKey && risCoreNormalizeText_(item.program) !== selectedKey) return;

    item.stock = risCoreParseNumber_(item.stock);
    item.unitCost = risCoreParseNumber_(item.unitCost);
    item.totalCost = risCoreParseNumber_(item.totalCost);
    item.expiry = risCoreFormatDate_(item.expiry);
    item.deliveryDate = risCoreFormatDate_(item.deliveryDate);
    item.sourceRow = rowNumber;
    item.sourceHeaderRow = layout.headerRow;
    item.sourceSheet = sheet.getName();
    item.sourceSpreadsheetId = sourceSpreadsheetId;
    item.inventorySource = resolvedSource.buttonName || resolvedSource.category;
    item.status = risStockStatus_(item.stock);
    item.statusLabel = risStockStatusLabel_(item.status);
    statusCounts[item.status] = (statusCounts[item.status] || 0) + 1;
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
