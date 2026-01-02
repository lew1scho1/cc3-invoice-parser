// ============================================================================
// BarcodeUpload.js - Upload sheet processing to Active DB
// ============================================================================

var BARCODE_UPLOAD_CONFIG = {
  uploadSheet: 'Upload',
  outreNewSheet: 'Upload(Outre NEW)'
};

var COMPANY_HEADERS = {
  OUTRE: ['ITEM GROUP', 'ITEM NUMBER', 'ITEM NAME', 'COLOR', 'BARCODE'],
  SNG: ['Class', 'Old Item', 'Old Item Code', 'Item', 'Item Code', 'Color', 'Description', 'Barcode']
};

var OUTRE_NEW_HEADERS = ['Item Group', 'Item Name', 'Color', 'UPC Barcode', 'REMARK'];
var META_HEADERS = ['first_seen_ym', 'last_seen_ym', 'status'];

var UPLOAD_EVENTS_SHEETS = {
  OUTRE: 'UPLOAD_EVENTS_OUTRE',
  SNG: 'UPLOAD_EVENTS_SNG'
};

var UPLOAD_EVENTS_HEADERS = {
  OUTRE: COMPANY_HEADERS.OUTRE.concat(['upload_ym', 'source', 'ingested_at']),
  SNG: COMPANY_HEADERS.SNG.concat(['upload_ym', 'source', 'ingested_at'])
};

function processUploadToActiveDb() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var uploadSheet = ss.getSheetByName(BARCODE_UPLOAD_CONFIG.uploadSheet);
    if (!uploadSheet) {
      throw new Error('Upload sheet not found.');
    }

    var headerRow = uploadSheet.getRange(1, 1, 1, uploadSheet.getLastColumn()).getValues()[0];
    var headerMap = buildHeaderIndex(headerRow);

    var companyKey = detectCompanyByHeader(headerMap);
    if (!companyKey) {
      throw new Error('Failed to detect company from Upload headers.');
    }

    var uploadDate = readUploadDate(uploadSheet);

    var rawUploadRows = readUploadRows(uploadSheet, headerMap, COMPANY_HEADERS[companyKey]);

    var outreNewRows = [];
    var outreNewSheet = null;
    if (companyKey === 'OUTRE') {
      outreNewSheet = ss.getSheetByName(BARCODE_UPLOAD_CONFIG.outreNewSheet);
      outreNewRows = readOutreNewRows(outreNewSheet);
    }

    var uploadRows = mergeUploadRows(rawUploadRows, outreNewRows, companyKey);
    if (uploadRows.length === 0) {
      return {
        success: false,
        error: 'Upload data is empty.'
      };
    }

    var uploadSourceSet = buildBarcodeSet(rawUploadRows, COMPANY_HEADERS[companyKey]);
    var outreNewSourceSet = buildBarcodeSet(outreNewRows, COMPANY_HEADERS[companyKey]);

    var activeSheet = getActiveDbSheet(ss, companyKey);
    var activeData = readActiveRecords(activeSheet, COMPANY_HEADERS[companyKey]);
    var activeRecords = activeData.records;

    var uploadBarcodes = buildBarcodeSet(uploadRows, COMPANY_HEADERS[companyKey]);
    var activeBarcodes = buildBarcodeSetFromRecords(activeRecords);

    var outreNewSet = {};
    if (companyKey === 'OUTRE') {
      outreNewSet = buildOutreNewSet(ss);
    }

    var newRows = [];
    for (var i = 0; i < uploadRows.length; i++) {
      var barcode = normalizeBarcodeValue(uploadRows[i][COMPANY_HEADERS[companyKey].length - 1]);
      if (!barcode || activeBarcodes[barcode]) {
        continue;
      }
      if (companyKey === 'OUTRE' && !outreNewSet[barcode]) {
        continue;
      }
      newRows.push(uploadRows[i]);
    }

    var discontinuedRows = [];
    for (var activeBarcode in activeRecords) {
      if (!activeRecords.hasOwnProperty(activeBarcode) || uploadBarcodes[activeBarcode]) {
        continue;
      }
      discontinuedRows.push(activeRecords[activeBarcode].baseRow);
    }

    var currentActiveYm = updateCurrentActiveYm(companyKey, uploadDate);
    appendUploadEvents(ss, companyKey, uploadRows, uploadDate, uploadSourceSet, outreNewSourceSet);
    rebuildActiveSnapshotFromHistory(ss, companyKey, currentActiveYm);

    try {
      rebuildSearchIndex(companyKey, currentActiveYm);
    } catch (indexError) {
      logError('rebuildSearchIndex', indexError.toString(), { company: companyKey, uploadDate: currentActiveYm });
    }

    resetUploadSheets(uploadSheet, outreNewSheet);

    return {
      success: true,
      message: 'Upload processed',
      company: companyKey,
      uploadDate: uploadDate,
      currentActiveYm: currentActiveYm,
      uploadRows: uploadRows.length,
      newCount: newRows.length,
      discontinuedUnknownCount: discontinuedRows.length,
      activeDbUpdated: true
    };
  } catch (error) {
    logError('processUploadToActiveDb', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function detectCompanyByHeader(headerMap) {
  if (hasHeaders(headerMap, COMPANY_HEADERS.OUTRE)) {
    return 'OUTRE';
  }
  if (hasHeaders(headerMap, COMPANY_HEADERS.SNG)) {
    return 'SNG';
  }
  return null;
}

function buildHeaderIndex(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeHeaderValue(headers[i]);
    if (key) {
      map[key] = i;
    }
  }
  return map;
}

function normalizeHeaderValue(value) {
  if (value === null || value === undefined) return '';
  return value.toString().trim();
}

function hasHeaders(headerMap, requiredHeaders) {
  for (var i = 0; i < requiredHeaders.length; i++) {
    if (headerMap[requiredHeaders[i]] === undefined) {
      return false;
    }
  }
  return true;
}

function readUploadDate(sheet) {
  var yearRaw = sheet.getRange('J1').getValue();
  var monthRaw = sheet.getRange('K1').getValue();

  var yearValue = yearRaw !== null && yearRaw !== undefined ? yearRaw.toString().trim() : '';
  var monthValue = monthRaw !== null && monthRaw !== undefined ? monthRaw.toString().trim() : '';

  if (!/^\d{4}$/.test(yearValue)) {
    throw new Error('Upload J1 must be a 4-digit year (e.g. 2025).');
  }

  var monthNum = parseInt(monthValue, 10);
  if (!monthValue || isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
    throw new Error('Upload K1 must be a month between 1 and 12.');
  }

  var monthPadded = monthNum < 10 ? '0' + monthNum : monthNum.toString();
  return yearValue + monthPadded;
}

function readUploadRows(sheet, headerMap, headerList) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var row = extractRowByHeaders(data[i], headerMap, headerList);
    if (hasAnyValue(row)) {
      rows.push(row);
    }
  }
  return rows;
}
function readOutreNewRows(sheet) {
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderIndex(headerRow);

  if (headerMap['Item Group'] === undefined ||
      headerMap['Item Name'] === undefined ||
      headerMap['Color'] === undefined ||
      headerMap['UPC Barcode'] === undefined) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var barcode = normalizeBarcodeValue(data[i][headerMap['UPC Barcode']]);
    if (!barcode) continue;

    rows.push([
      data[i][headerMap['Item Group']] || '',
      '',
      data[i][headerMap['Item Name']] || '',
      data[i][headerMap['Color']] || '',
      barcode
    ]);
  }

  return rows;
}

function mergeUploadRows(uploadRows, outreNewRows, companyKey) {
  if (companyKey !== 'OUTRE') return uploadRows;

  var merged = [];
  var seen = {};
  var barcodeIdx = COMPANY_HEADERS[companyKey].length - 1;

  for (var i = 0; i < uploadRows.length; i++) {
    var barcode = normalizeBarcodeValue(uploadRows[i][barcodeIdx]);
    if (!barcode) continue;
    if (!seen[barcode]) {
      seen[barcode] = true;
      merged.push(uploadRows[i]);
    }
  }

  for (var j = 0; j < outreNewRows.length; j++) {
    var newBarcode = normalizeBarcodeValue(outreNewRows[j][barcodeIdx]);
    if (!newBarcode || seen[newBarcode]) continue;
    seen[newBarcode] = true;
    merged.push(outreNewRows[j]);
  }

  return merged;
}


function normalizeYearMonthValue(value) {
  if (value === null || value === undefined) return '';
  var text = value.toString().trim();
  return text;
}

function getUploadEventsSheet(ss, companyKey) {
  var sheetName = UPLOAD_EVENTS_SHEETS[companyKey];
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  ensureUploadEventsHeader(sheet, companyKey);
  return sheet;
}

function ensureUploadEventsHeader(sheet, companyKey) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    var headers = UPLOAD_EVENTS_HEADERS[companyKey];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function appendUploadEvents(ss, companyKey, uploadRows, uploadDate, uploadSourceSet, outreNewSourceSet) {
  var sheet = getUploadEventsSheet(ss, companyKey);
  var headers = UPLOAD_EVENTS_HEADERS[companyKey];
  var barcodeIdx = COMPANY_HEADERS[companyKey].length - 1;
  var rows = [];
  var now = new Date();

  for (var i = 0; i < uploadRows.length; i++) {
    var barcode = normalizeBarcodeValue(uploadRows[i][barcodeIdx]);
    if (!barcode) continue;

    var source = 'UPLOAD';
    if (companyKey === 'OUTRE' && !uploadSourceSet[barcode] && outreNewSourceSet[barcode]) {
      source = 'OUTRE_NEW';
    }

    rows.push(uploadRows[i].concat([uploadDate, source, now]));
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function rebuildActiveSnapshotFromHistory(ss, companyKey, currentActiveYm) {
  var historySheet = getUploadEventsSheet(ss, companyKey);
  var activeSheet = getActiveDbSheet(ss, companyKey);
  var activeHeaders = COMPANY_HEADERS[companyKey].concat(META_HEADERS);

  var lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    writeSheetData(activeSheet, activeHeaders, []);
    return;
  }

  var lastCol = historySheet.getLastColumn();
  var headerRow = historySheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderIndex(headerRow);

  var uploadYmIdx = headerMap['upload_ym'];
  if (uploadYmIdx === undefined) {
    throw new Error('upload_ym column missing in history sheet.');
  }

  var barcodeHeader = COMPANY_HEADERS[companyKey][COMPANY_HEADERS[companyKey].length - 1];
  var barcodeIdx = headerMap[barcodeHeader];
  if (barcodeIdx === undefined) {
    throw new Error('barcode column missing in history sheet.');
  }

  var data = historySheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var grouped = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var barcode = normalizeBarcodeValue(row[barcodeIdx]);
    if (!barcode) continue;

    var uploadYm = normalizeYearMonthValue(row[uploadYmIdx]);
    if (!uploadYm) continue;

    if (!grouped[barcode]) {
      grouped[barcode] = {
        firstSeen: uploadYm,
        lastSeen: uploadYm,
        latestYm: uploadYm,
        latestRow: row
      };
      continue;
    }

    var entry = grouped[barcode];
    if (compareYearMonth(uploadYm, entry.firstSeen) < 0) {
      entry.firstSeen = uploadYm;
    }
    if (compareYearMonth(uploadYm, entry.lastSeen) > 0) {
      entry.lastSeen = uploadYm;
    }
    if (compareYearMonth(uploadYm, entry.latestYm) >= 0) {
      entry.latestYm = uploadYm;
      entry.latestRow = row;
    }
  }

  var outputRows = [];
  for (var key in grouped) {
    if (!grouped.hasOwnProperty(key)) continue;
    var item = grouped[key];
    var baseRow = extractRowByHeaders(item.latestRow, headerMap, COMPANY_HEADERS[companyKey]);
    var status = calculateStatus(item.firstSeen, item.lastSeen, currentActiveYm);
    outputRows.push(baseRow.concat([item.firstSeen, item.lastSeen, status]));
  }

  writeSheetData(activeSheet, activeHeaders, outputRows);
}

function updateCurrentActiveYm(companyKey, uploadDate) {
  var props = PropertiesService.getScriptProperties();
  var key = companyKey === 'OUTRE' ? 'ACTIVE_DB_DATE_OUTRE' : 'ACTIVE_DB_DATE_SNG';
  var current = props.getProperty(key);

  if (!current || compareYearMonth(uploadDate, current) > 0) {
    props.setProperty(key, uploadDate);
    return uploadDate;
  }

  return current;
}

function readActiveRows(sheet, headerMap, headerList) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var row = extractRowByHeaders(data[i], headerMap, headerList);
    if (hasAnyValue(row)) {
      rows.push(row);
    }
  }
  return rows;
}

function extractRowByHeaders(row, headerMap, headerList) {
  var output = [];
  for (var i = 0; i < headerList.length; i++) {
    var idx = headerMap[headerList[i]];
    output.push(idx !== undefined ? row[idx] : '');
  }
  return output;
}

function hasAnyValue(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) {
      return true;
    }
  }
  return false;
}

function buildBarcodeSet(rows, headerList) {
  var set = {};
  var barcodeIdx = headerList.length - 1;
  for (var i = 0; i < rows.length; i++) {
    var barcode = normalizeBarcodeValue(rows[i][barcodeIdx]);
    if (barcode) {
      set[barcode] = true;
    }
  }
  return set;
}

function normalizeBarcodeValue(value) {
  if (value === null || value === undefined) return '';
  return value.toString().trim();
}

function readActiveRecords(sheet, headerList) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = buildHeaderIndex(headerRow);
  var records = {};

  if (lastRow < 2) {
    return { records: records, headerMap: headerMap };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var barcodeHeader = headerList[headerList.length - 1];
  var barcodeIdx = headerMap[barcodeHeader];

  for (var i = 0; i < data.length; i++) {
    var baseRow = extractRowByHeaders(data[i], headerMap, headerList);
    if (!hasAnyValue(baseRow)) {
      continue;
    }
    var barcode = '';
    if (barcodeIdx !== undefined) {
      barcode = normalizeBarcodeValue(data[i][barcodeIdx]);
    }
    if (!barcode) {
      barcode = normalizeBarcodeValue(baseRow[headerList.length - 1]);
    }
    if (!barcode) {
      continue;
    }
    records[barcode] = {
      baseRow: baseRow,
      firstSeen: getMetaValue(data[i], headerMap, META_HEADERS[0]),
      lastSeen: getMetaValue(data[i], headerMap, META_HEADERS[1]),
      status: getMetaValue(data[i], headerMap, META_HEADERS[2])
    };
  }

  return { records: records, headerMap: headerMap };
}

function getMetaValue(row, headerMap, headerName) {
  var idx = headerMap[headerName];
  return idx !== undefined ? row[idx] : '';
}

function buildOutreNewSet(ss) {
  var sheet = ss.getSheetByName(BARCODE_UPLOAD_CONFIG.outreNewSheet);
  if (!sheet) {
    return {};
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = buildHeaderIndex(headerRow);
  if (headerMap['UPC Barcode'] === undefined) {
    return {};
  }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var set = {};
  for (var i = 0; i < data.length; i++) {
    var barcode = normalizeBarcodeValue(data[i][headerMap['UPC Barcode']]);
    if (barcode) {
      set[barcode] = true;
    }
  }
  return set;
}

function buildBarcodeSetFromRecords(records) {
  var set = {};
  for (var barcode in records) {
    if (records.hasOwnProperty(barcode)) {
      set[barcode] = true;
    }
  }
  return set;
}

function getActiveDbSheet(ss, companyKey) {
  var name = CONFIG.COMPANIES[companyKey].dbSheet;
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(name + ' sheet not found.');
  }
  return sheet;
}

function buildActiveDbRows(uploadRows, uploadDate, companyKey, activeRecords, outreNewSet) {
  var rows = [];
  var uploadBarcodes = {};
  var barcodeIdx = COMPANY_HEADERS[companyKey].length - 1;

  // Step 1: barcodes present in upload (refresh last_seen)
  for (var i = 0; i < uploadRows.length; i++) {
    var barcode = normalizeBarcodeValue(uploadRows[i][barcodeIdx]);
    if (!barcode) {
      continue;
    }

    uploadBarcodes[barcode] = true;

    var existing = activeRecords[barcode];
    var firstSeen = existing && existing.firstSeen ? existing.firstSeen : uploadDate;
    var lastSeen = uploadDate; // Upload present -> refresh last_seen

    // Status calculation
    var status = calculateStatus(firstSeen, lastSeen, uploadDate);

    rows.push(uploadRows[i].concat([firstSeen, lastSeen, status]));
  }

  // Step 2: barcodes missing from upload (keep last_seen)
  for (var barcode in activeRecords) {
    if (!activeRecords.hasOwnProperty(barcode)) continue;
    if (uploadBarcodes[barcode]) continue; // Already processed
    var record = activeRecords[barcode];
    var firstSeen = record.firstSeen;
    var lastSeen = record.lastSeen; // Upload missing -> keep last_seen

    // Status calculation
    var status = calculateStatus(firstSeen, lastSeen, uploadDate);

    rows.push(record.baseRow.concat([firstSeen, lastSeen, status]));
  }

  return rows;
}

function ensureSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function writeSheetData(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function clearUploadSheetsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    clearSheetDataExceptHeader(ss.getSheetByName(BARCODE_UPLOAD_CONFIG.uploadSheet));
    clearSheetDataExceptHeader(ss.getSheetByName(BARCODE_UPLOAD_CONFIG.outreNewSheet));
    SpreadsheetApp.getUi().alert('Upload sheet data cleared.');
  } catch (error) {
    logError('clearUploadSheetsData', error.toString());
    SpreadsheetApp.getUi().alert('Upload sheet clear failed: ' + error.toString());
  }
}

function clearSheetDataExceptHeader(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

function resetUploadSheets(uploadSheet, outreNewSheet) {
  if (uploadSheet) {
    var yearValue = uploadSheet.getRange('J1').getValue();
    var monthValue = uploadSheet.getRange('K1').getValue();

    var maxRows = uploadSheet.getMaxRows();
    var maxCols = uploadSheet.getMaxColumns();
    uploadSheet.clear();
    uploadSheet.getRange(1, 1, maxRows, maxCols).setNumberFormat('@');

    uploadSheet.getRange('J1').setValue(yearValue);
    uploadSheet.getRange('K1').setValue(monthValue);
  }

  if (outreNewSheet) {
    var maxRowsNew = outreNewSheet.getMaxRows();
    var maxColsNew = outreNewSheet.getMaxColumns();
    outreNewSheet.clear();
    outreNewSheet.getRange(1, 1, maxRowsNew, maxColsNew).setNumberFormat('@');
  }
}

function shouldUpdateActiveDb(companyKey, uploadDate, activeSheet) {
  if (activeSheet && activeSheet.getLastRow() < 2) {
    return true;
  }
  var props = PropertiesService.getScriptProperties();
  var key = companyKey === 'OUTRE' ? 'ACTIVE_DB_DATE_OUTRE' : 'ACTIVE_DB_DATE_SNG';
  var current = props.getProperty(key);
  if (!current) return true;
  return compareYearMonth(uploadDate, current) >= 0;
}

function setActiveDbDate(companyKey, uploadDate) {
  var props = PropertiesService.getScriptProperties();
  var key = companyKey === 'OUTRE' ? 'ACTIVE_DB_DATE_OUTRE' : 'ACTIVE_DB_DATE_SNG';
  props.setProperty(key, uploadDate);
}

function compareYearMonth(a, b) {
  var ai = parseInt(a, 10);
  var bi = parseInt(b, 10);
  if (isNaN(ai) || isNaN(bi)) {
    throw new Error('Date compare failed: ' + a + ', ' + b);
  }
  if (ai === bi) return 0;
  return ai > bi ? 1 : -1;
}






