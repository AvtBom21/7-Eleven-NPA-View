// ========================================
// CONFIGURATION
// ========================================
var CONFIG = {
  SPREADSHEET_ID: '1W8oN5mha-GZ6wLIct-hsWQXSbFKreeICw2Y0-SFgQnY',
  SHEET_NAME_CATEGORY: 'Category',
  SHEET_NAME_MANAGER: 'Manager',
  SHEET_NAME_DIRECTOR: 'Director',
  DROPDOWN_SHEET: 'Dropdown_Master',
  USER_MASTER_SHEET: 'User_Master',
  AUDIT_LOG_SHEET: 'Audit_Log',
  UOM_IMAGE_FOLDER_ID: '1WDIup4xssu-HLNGTOND3G9bIFIJfQF0g',
  MAX_LOG_ENTRIES: 20,
  PAGE_SIZE: 50,
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

var PIC_EDITABLE_FIELDS = [
  'Present Date',
  'Status',
  'New product type',
  'Why Do We List This Product?',
  'Premium products',
  'Branded',
  'Sub Category',
  'PRIMARY USAGE SEGMENT',
  'TARGET STORE TYPES',
  'TARGET CUSTOMER STORE TYPES',
  'Preservation temperature',
  'PIC input Product Name (Dưới 40 kí tự)',
  'SAME PRICE PRODUCT',
  'REFERENCE PRODUCT ID (Nếu không có điền NO)',
  'Retail Price (+VAT) in Tier 6',
  'Other Income + Listing Fee',
  'Pricing Strategy',
  'Competitors (Retail Price) <Tên đối thủ: Giá tiền>',
  'Product description (Dưới 250 Kí tự)',
  'Pattern',
  'SSV Note Product',
  'Logistics Group South',
  'Logistics Group North',
  'FC UPSD of NEW SKU',
  'Display area (POG CONCEPT) Bắt Buộc',
  'Inventory Type',
  'POG_Choose Reduce product facing',
  'POG_Choose type Reduce product facing',
  'UOM Information'
];

var DIRECTOR_EDITABLE_FIELDS = [
  'Status',
  'Note'
];

var EDITABLE_FIELDS = PIC_EDITABLE_FIELDS.slice();

function _safeErrorText_(err, fallback) {
  var fb = fallback || 'Unknown error';
  try {
    if (err === null || err === undefined) return fb;
    if (typeof err === 'string') return err;
    if (typeof err.message === 'string' && err.message.trim()) return err.message;
    if (err.error && typeof err.error.message === 'string' && err.error.message.trim()) return err.error.message;
    var s = (typeof err.toString === 'function') ? err.toString() : '';
    if (s && s !== '[object Object]') return s;
    var json = JSON.stringify(err);
    return (json && json !== '{}') ? json : fb;
  } catch (e) {
    return fb;
  }
}

// ========================================
// ENTRY POINT — Role-based routing
// ========================================
// Access Login     : ?page=login
// Access PIC view  : ?page=index
// Access Director  : ?page=director
function doGet(e) {
  try {
    var page = (e && e.parameter && e.parameter.page)
      ? e.parameter.page.toLowerCase() : 'index';
    if (page !== 'login' && page !== 'index' && page !== 'director') {
      page = 'index';
    }

    if (page === 'login') {
      return _renderLoginPage('', _resolveRequestEmail_(e) || '');
    }

    if (page !== 'index' && page !== 'director') {
      return _renderLoginPage('Trang không hợp lệ. Vui lòng đăng nhập lại.', _resolveRequestEmail_(e) || '');
    }

    var token = e && e.parameter ? String(e.parameter.token || '').trim() : '';
    var access;
    if (!token) {
      // Force login for production - no bypass
      return _renderLoginPage('Vui lòng đăng nhập để tiếp tục.', _resolveRequestEmail_(e) || '');
    } else {
      access = _getAccessByToken_(token);
      if (!access.ok) {
        return _renderLoginPage(access.error.message, _resolveRequestEmail_(e) || '');
      }
    }

    // If non-director requests director page, always show index page.
    if (page === 'director' && access.data.role !== 'DIRECTOR') {
      page = 'index';
    }

    var fileName = page === 'director' ? 'Director_View' : 'Index';
    var title = page === 'director' ? 'Director – NPA Dashboard' : 'NPA Product Dashboard';
    var tpl = HtmlService.createTemplateFromFile(fileName);
    tpl.appUserJson = JSON.stringify({
      email: access.data.email,
      fullName: access.data.fullName,
      role: access.data.role,
      scope: access.data.scope,
      canEdit: access.data.canEdit,
      authToken: token
    }).replace(/</g, '\\u003c');
    var html = tpl.evaluate();
    html.setTitle(title);
    html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  } catch (err) {
    var msg = _safeErrorText_(err, 'Unknown error');
    return HtmlService.createHtmlOutput(
      '<h2 style="font-family:sans-serif;color:red">Error: ' + msg + '</h2>'
    );
  }
}

// ========================================
// HELPER: Safe sheet access
// ========================================
function _openSheet(access) {
  var sheetName = (access && access.role === 'MANAGER') ? CONFIG.SHEET_NAME_MANAGER : CONFIG.SHEET_NAME_CATEGORY;
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var names = ss.getSheets().map(function(s) { return s.getName(); }).join(', ');
    throw new Error('Sheet "' + sheetName + '" not found. Available: ' + names);
  }
  return sheet;
}

function _getHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  var vals = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  return vals.map(function(h) {
    return (h === null || h === undefined) ? '' : String(h).trim();
  });
}

function _trimTrailingEmptyHeaders_(headers) {
  var out = Array.isArray(headers) ? headers.slice() : [];
  while (out.length > 0 && out[out.length - 1] === '') out.pop();
  return out;
}

function _cellToString(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    // Format as YYYY-MM-DD
    var y = val.getFullYear();
    var m = String(val.getMonth() + 1).padStart('2', '0');
    var d = String(val.getDate()).padStart('2', '0');
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

function _normalizeAuditValue_(field, value) {
  if (field === 'Present Date' && value) {
    var d = value instanceof Date ? value : new Date(value);
    if (!isNaN(d.getTime())) return _cellToString(d);
  }
  return _cellToString(value);
}

function _buildRowObjectFromArray_(headers, rowValues, rowNumber) {
  var obj = { _rowNumber: rowNumber };
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = _cellToString(rowValues[i]);
  }
  return obj;
}

function _getRowSnapshotByNumber_(sheet, headers, rowNumber) {
  var rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  return _buildRowObjectFromArray_(headers, rowValues, rowNumber);
}

function _hasRowAccess_(rowObj, access) {
  if (!rowObj) return false;
  if (!access || !access.role || access.role === 'DIRECTOR') return true;
  return _applyAccessScopeToRows_([rowObj], access).length > 0;
}

function _getScopeDisplay_(access) {
  if (!access || access.role === 'DIRECTOR') return 'ALL';
  var scopes = Array.isArray(access.scope) ? access.scope : [];
  return scopes.length ? scopes.join(', ') : '';
}

function _getFormattedTimestamp_() {
  var tz = '';
  try {
    tz = Session.getScriptTimeZone();
  } catch (err) {
    tz = '';
  }
  return Utilities.formatDate(new Date(), tz || 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
}

function _getEditLogUser_(access) {
  if (access && access.email) return String(access.email);
  try {
    var activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) return String(activeEmail);
  } catch (err) {
    // Keep fallback below when session email is not available.
  }
  return 'unknown';
}

function _appendEditHistoryNote_(range, beforeValue, afterValue, userEmail) {
  if (!range) return;

  var sheetName = range.getSheet().getName();
  var oldValue = (beforeValue === null || beforeValue === undefined || String(beforeValue) === '')
    ? '(empty)'
    : String(beforeValue);
  var newValue = (afterValue === null || afterValue === undefined || String(afterValue) === '')
    ? '(empty)'
    : String(afterValue);
  var currentNote = range.getNote() || '';

  var newEntry = _getFormattedTimestamp_() + ' | Sheet: ' + sheetName + '\n'
    + 'Old: ' + oldValue + '\n'
    + 'New: ' + newValue + '\n'
    + 'By: ' + userEmail;

  var updatedNote = currentNote ? (newEntry + '\n\n' + currentNote) : newEntry;
  var entries = updatedNote.split('\n\n');
  var maxEntries = Math.max(1, parseInt(CONFIG.MAX_LOG_ENTRIES, 10) || 20);
  var limitedNote = entries.slice(0, maxEntries).join('\n\n');

  range.setNote(limitedNote);
}

function _resolveAuditSource_(payload, access, fallback) {
  var source = payload && payload.source ? String(payload.source).trim() : '';
  if (source) return source;
  if (fallback) return fallback;
  return (access && access.role === 'DIRECTOR') ? 'DIRECTOR_UPDATE' : 'PIC_UPDATE';
}

function _ensureAuditLogSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.AUDIT_LOG_SHEET);
  var headers = [
    'Timestamp',
    'Request ID',
    'Source',
    'User Email',
    'Full Name',
    'Role',
    'Scope',
    'Row Number',
    'Ticket ID',
    'Product Name',
    'Field',
    'Before',
    'After',
    'Sheet Name'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.AUDIT_LOG_SHEET);
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  if (sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('') !== headers.join('')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    var widths = [170, 240, 190, 220, 180, 120, 260, 110, 140, 260, 220, 260, 260, 120];
    for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);
  }

  return sheet;
}

function _buildAuditEntries_(changeDetails, rowSnapshot, access, context) {
  var entries = [];
  if (!changeDetails || !changeDetails.length) return entries;

  var now = new Date();
  var ticketId = rowSnapshot['Ticket ID'] || rowSnapshot['TICKET ID'] || '';
  var productName = rowSnapshot['Product Name']
    || rowSnapshot['PIC input Product Name (Duoi 40 ki tu)']
    || rowSnapshot['PIC input Product Name (Dưới 40 kí tự)']
    || '';

  for (var i = 0; i < changeDetails.length; i++) {
    var item = changeDetails[i];
    entries.push([
      now,
      context.requestId || '',
      context.source || '',
      access && access.email ? access.email : '',
      access && access.fullName ? access.fullName : '',
      access && access.role ? access.role : '',
      _getScopeDisplay_(access),
      rowSnapshot._rowNumber || '',
      ticketId,
      productName,
      item.field,
      item.before,
      item.after,
      CONFIG.SHEET_NAME
    ]);
  }
  return entries;
}

function _appendAuditLogEntries_(entries) {
  if (!entries || !entries.length) return;
  var sheet = _ensureAuditLogSheet_();
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);
}

// Polyfill padStart for GAS
if (!String.prototype.padStart) {
  String.prototype.padStart = function(targetLength, padString) {
    var str = String(this);
    padString = padString || ' ';
    while (str.length < targetLength) str = padString + str;
    return str;
  };
}

function _applyAccessScopeToRows_(rows, access) {
  if (!access || !access.role || access.role === 'DIRECTOR') return rows;
  var scopes = Array.isArray(access.scope) ? access.scope : [];
  if (!scopes.length) return [];
  var scopeFields = _detectCategoryScopeFieldNames_(rows);
  if (!scopeFields.length) return [];
  var scopeMap = {};
  for (var i = 0; i < scopes.length; i++) {
    var scopeKey = _normalizeCategoryKey_(scopes[i]);
    if (scopeKey) scopeMap[scopeKey] = true;
  }
  if (scopeMap['all']) return rows;

  return rows.filter(function(r) {
    if (!r) return false;
    for (var f = 0; f < scopeFields.length; f++) {
      var field = scopeFields[f];
      var rowCategory = _normalizeCategoryKey_(r[field] || '');
      if (rowCategory && scopeMap[rowCategory]) return true;
    }
    return false;
  });
}

function _normalizeCategoryKey_(text) {
  return String(text || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function _findSheetByLooseName_(ss, exactNames, requiredTokens) {
  var sheets = ss.getSheets();
  var names = sheets.map(function(s) { return s.getName(); });
  var normalize = function(v) { return _normalizeCategoryKey_(v); };

  for (var i = 0; i < exactNames.length; i++) {
    var hit = ss.getSheetByName(exactNames[i]);
    if (hit) return { sheet: hit, allNames: names };
  }

  var exactNorm = {};
  for (var j = 0; j < exactNames.length; j++) exactNorm[normalize(exactNames[j])] = true;
  for (var s = 0; s < sheets.length; s++) {
    if (exactNorm[normalize(sheets[s].getName())]) return { sheet: sheets[s], allNames: names };
  }

  if (requiredTokens && requiredTokens.length) {
    for (var k = 0; k < sheets.length; k++) {
      var n = normalize(sheets[k].getName());
      var ok = true;
      for (var t = 0; t < requiredTokens.length; t++) {
        var token = normalize(requiredTokens[t]);
        if (!token) continue;
        if (n.indexOf(token) < 0) { ok = false; break; }
      }
      if (ok) return { sheet: sheets[k], allNames: names };
    }
  }

  return { sheet: null, allNames: names };
}

function _normalizeSubcategoryKey_(text) {
  return String(text || '').trim().replace(/\s+/g, ' ');
}

function _isSubcategoryHeaderText_(text) {
  var n = _normalizeCategoryKey_(text);
  return n === 'sub category' || n === 'sub_category' || n === 'subcategory';
}

function _parseScopeList_(text) {
  if (!text) return [];
  return String(text)
    .split(/[,\n;]+/)
    .map(function(s) { return String(s || '').trim(); })
    .filter(Boolean);
}

function _normalizePersonName_(text) {
  return String(text || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function _parsePersonList_(text) {
  if (!text) return [];
  return String(text)
    .split(/[,\n;]+/)
    .map(function(s) { return String(s || '').trim(); })
    .filter(Boolean);
}

function _detectPicFieldName_(rows) {
  if (!rows || !rows.length) return '';
  var sample = rows[0] || {};
  var keys = Object.keys(sample);
  if (!keys.length) return '';

  var priority = ['PIC CATEGORY', 'PIC', 'PIC NAME', 'PIC FULL NAME'];
  var keyMap = {};
  for (var i = 0; i < keys.length; i++) {
    keyMap[String(keys[i]).trim().toLowerCase()] = keys[i];
  }
  for (var p = 0; p < priority.length; p++) {
    var hit = keyMap[priority[p].toLowerCase()];
    if (hit) return hit;
  }
  return '';
}

function _detectCategoryScopeFieldName_(rows) {
  var fields = _detectCategoryScopeFieldNames_(rows);
  return fields.length ? fields[0] : '';
}

function _detectCategoryScopeFieldNames_(rows) {
  if (!rows || !rows.length) return [];
  var sample = rows[0] || {};
  var keys = Object.keys(sample);
  if (!keys.length) return [];

  var keyMap = {};
  for (var i = 0; i < keys.length; i++) {
    keyMap[String(keys[i]).trim().toLowerCase()] = keys[i];
  }

  var ordered = ['pic category', 'category_ncc', 'product group', 'category'];
  var found = [];
  for (var j = 0; j < ordered.length; j++) {
    if (keyMap[ordered[j]]) found.push(keyMap[ordered[j]]);
  }
  return found;
}

function _resolveAccessForDataApi_(params) {
  if (typeof params === 'string') return _getAccessByToken_(params);
  if (!params || typeof params !== 'object') {
    return _err('MISSING_TOKEN', 'Vui long dang nhap de tiep tuc.');
  }
  var token = params.token ? String(params.token).trim() : '';
  if (!token) {
    return _err('MISSING_TOKEN', 'Vui long dang nhap de tiep tuc.');
  }
  return _getAccessByToken_(token);
}

function _getEditableFieldsForAccess_(access) {
  if (!access || access.canEdit !== true) return [];
  if (access.role === 'DIRECTOR') return DIRECTOR_EDITABLE_FIELDS.slice();
  return PIC_EDITABLE_FIELDS.slice();
}

// ========================================
// API: getInitData
// ========================================
function getInitData(params) {
  try {
    Logger.log('getInitData START');

    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;
    var editableFields = _getEditableFieldsForAccess_(accessRes.data);

    var sheet = _openSheet(accessRes.data);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    Logger.log('Dimensions: lastRow=' + lastRow + ' lastCol=' + lastCol);

    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var numCols = headers.length;

    Logger.log('Headers (' + numCols + '): ' + headers.slice(0, 5).join(', ') + '...');

    if (numCols === 0 || lastRow < CONFIG.DATA_START_ROW) {
      return _ok({
        headers: headers,
        rows: [],
        editableFields: editableFields,
        meta: { totalRows: 0, currentPage: 1, totalPages: 1, pageSize: CONFIG.PAGE_SIZE }
      });
    }

    // Read all data at once
    var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
    var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols).getValues();

    // Build rows with actual row numbers, skip totally empty rows
    var allRows = [];
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      var hasContent = false;
      for (var j = 0; j < row.length; j++) {
        if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
          hasContent = true;
          break;
        }
      }
      if (!hasContent) continue;

      var obj = { _rowNumber: CONFIG.DATA_START_ROW + i };
      for (var k = 0; k < numCols; k++) {
        obj[headers[k]] = _cellToString(row[k]);
      }
      allRows.push(obj);
    }

    if (accessRes.data) {
      allRows = _applyAccessScopeToRows_(allRows, accessRes.data);
    }

    var total = allRows.length;

    Logger.log('getInitData SUCCESS: total=' + total + ' loaded=' + allRows.length);

    return _ok({
      headers: headers,
      rows: allRows,
      editableFields: editableFields,
      meta: { totalRows: total, currentPage: 1, totalPages: 1, pageSize: total || CONFIG.PAGE_SIZE }
    });

  } catch (err) {
    Logger.log('getInitData ERROR: ' + err.toString());
    return _err('INIT_ERROR', err.message, err.stack);
  }
}

// ========================================
// API: getDropdownMaster
// ========================================
function getDropdownMaster(params) {
  try {
    Logger.log('getDropdownMaster START');
    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.DROPDOWN_SHEET);
    
    if (!sheet) {
      Logger.log('Dropdown_Master sheet not found, returning empty');
      return _ok({ dropdowns: {}, subCategoryMap: {}, subCategoryBenchMap: {} });
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow < 2 || lastCol < 1) {
      return _ok({ dropdowns: {}, subCategoryMap: {}, subCategoryBenchMap: {} });
    }
    
    // Read headers from row 1
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headers = headers.map(function(h) {
      return (h === null || h === undefined) ? '' : String(h).trim();
    });
    
    // Read all data starting from row 2
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Build unique values for each column
    var dropdowns = {};
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      if (!header) continue;
      
      var uniqueValues = {};
      for (var row = 0; row < data.length; row++) {
        var val = data[row][col];
        if (val !== null && val !== undefined && val !== '') {
          var strVal = String(val).trim();
          if (strVal) uniqueValues[strVal] = true;
        }
      }
      
      dropdowns[header] = Object.keys(uniqueValues).sort();
    }

    // Keep row-level Sub Category -> Group/Category and KPI mapping
    // so UI can auto-fill derived fields when Sub Category changes.
    var _normHdr = function(v) {
      return String(v || '').trim().replace(/\s+/g, ' ').toLowerCase();
    };
    var _findColIndex = function(candidates) {
      var normCandidates = {};
      for (var i = 0; i < candidates.length; i++) {
        normCandidates[_normHdr(candidates[i])] = true;
      }
      for (var c = 0; c < headers.length; c++) {
        if (normCandidates[_normHdr(headers[c])]) return c;
      }
      return -1;
    };

    var subCategoryMap = {};
    var subCategoryBenchMap = {};
    var idxSubCategory = _findColIndex(['Sub Category', 'Sub_Category', 'Subcategory']);
    var idxGroup = _findColIndex(['Group']);
    var idxCategory = _findColIndex(['Category']);
    var idxGpu = _findColIndex(['GP/Unit by Subcategory', 'GP/UNIT BY SUBCATEGORY']);
    var idxGpsc = _findColIndex(['%GP by Subcategory', '%GP BY SUBCATEGORY']);
    var idxUpsd = _findColIndex(['UPSD per SKU by Subcategory', 'UPSD PER SKU BY SUBCATEGORY']);
    var idxSku = _findColIndex(['#SKU by subcategory', '#SKU BY SUBCATEGORY']);
    var idxSp = _findColIndex(['%SP by Subcategory', '%SP BY SUBCATEGORY']);
    var subKeyMap = {};

    if (idxSubCategory >= 0) {
      for (var r = 0; r < data.length; r++) {
        var rawSub = data[r][idxSubCategory];
        var sub = (rawSub === null || rawSub === undefined) ? '' : String(rawSub).replace(/\s+/g, ' ').trim();
        if (!sub) continue;
        var subNorm = sub.toLowerCase();
        var subKey = subKeyMap[subNorm] || sub;
        if (!subKeyMap[subNorm]) subKeyMap[subNorm] = subKey;

        var group = idxGroup >= 0 ? String(data[r][idxGroup] || '').trim() : '';
        var category = idxCategory >= 0 ? String(data[r][idxCategory] || '').trim() : '';
        var gpu = idxGpu >= 0 ? _cellToString(data[r][idxGpu]) : '';
        var gpsc = idxGpsc >= 0 ? _cellToString(data[r][idxGpsc]) : '';
        var upsd = idxUpsd >= 0 ? _cellToString(data[r][idxUpsd]) : '';
        var sku = idxSku >= 0 ? _cellToString(data[r][idxSku]) : '';
        var sp = idxSp >= 0 ? _cellToString(data[r][idxSp]) : '';

        if (!subCategoryMap[subKey]) {
          subCategoryMap[subKey] = { group: group, category: category };
        } else {
          if (!subCategoryMap[subKey].group && group) subCategoryMap[subKey].group = group;
          if (!subCategoryMap[subKey].category && category) subCategoryMap[subKey].category = category;
        }

        if (!subCategoryBenchMap[subKey]) {
          subCategoryBenchMap[subKey] = { gpu: '', gpsc: '', upsd: '', sku: '', sp: '' };
        }
        var bm = subCategoryBenchMap[subKey];
        if (!bm.gpu && gpu) bm.gpu = gpu;
        if (!bm.gpsc && gpsc) bm.gpsc = gpsc;
        if (!bm.upsd && upsd) bm.upsd = upsd;
        if (!bm.sku && sku) bm.sku = sku;
        if (!bm.sp && sp) bm.sp = sp;
      }
    }
    
    Logger.log('getDropdownMaster SUCCESS: ' + Object.keys(dropdowns).length + ' columns, subCategoryMap=' + Object.keys(subCategoryMap).length + ', subCategoryBenchMap=' + Object.keys(subCategoryBenchMap).length);
    return _ok({ dropdowns: dropdowns, subCategoryMap: subCategoryMap, subCategoryBenchMap: subCategoryBenchMap });
    
  } catch (err) {
    Logger.log('getDropdownMaster ERROR: ' + err.toString());
    return _err('DROPDOWN_ERROR', err.message, err.stack);
  }
}

// ========================================
// API: searchRows
// ========================================
function searchRows(params) {
  try {
    params = params || {};
    Logger.log('searchRows params: ' + JSON.stringify(params));

    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;

    var sheet = _openSheet(accessRes.data);
    var lastRow = sheet.getLastRow();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var numCols = headers.length;

    if (numCols === 0 || lastRow < CONFIG.DATA_START_ROW) {
      return _ok({
        rows: [],
        meta: { totalRows: 0, currentPage: 1, totalPages: 1, pageSize: CONFIG.PAGE_SIZE }
      });
    }

    var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
    var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols).getValues();

    // Build full dataset
    var allRows = [];
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      var hasContent = false;
      for (var j = 0; j < row.length; j++) {
        if (row[j] !== null && row[j] !== undefined && row[j] !== '') { hasContent = true; break; }
      }
      if (!hasContent) continue;

      var obj = { _rowNumber: CONFIG.DATA_START_ROW + i };
      for (var k = 0; k < numCols; k++) {
        obj[headers[k]] = _cellToString(row[k]);
      }
      allRows.push(obj);
    }

    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;
    if (accessRes.data) {
      allRows = _applyAccessScopeToRows_(allRows, accessRes.data);
    }

    // Apply keyword filter
    var keyword = (params.keyword || '').toLowerCase().trim();
    if (keyword) {
      allRows = allRows.filter(function(row) {
        return headers.some(function(h) {
          var v = row[h] || '';
          return v.toLowerCase().indexOf(keyword) >= 0;
        });
      });
    }

    // Apply dropdown filters
    var filters = params.filters || {};
    if (filters.zone) {
      allRows = _filterByField(allRows, 'Zone', filters.zone);
    }
    if (filters.status) {
      allRows = _filterByField(allRows, 'Status', filters.status);
    }
    if (filters.productGroup) {
      allRows = _filterByField(allRows, 'Product Group', filters.productGroup);
    }
    if (filters.category) {
      allRows = _filterByField(allRows, 'Category', filters.category);
    }

    var total = allRows.length;
    var page = Math.max(1, parseInt(params.page) || 1);
    var pageSize = parseInt(params.pageSize) || CONFIG.PAGE_SIZE;
    var totalPages = Math.max(1, Math.ceil(total / pageSize));
    page = Math.min(page, totalPages);

    var start = (page - 1) * pageSize;
    var pageRows = allRows.slice(start, start + pageSize);

    Logger.log('searchRows: total=' + total + ' page=' + page + ' returning=' + pageRows.length);

    return _ok({
      rows: pageRows,
      meta: { totalRows: total, currentPage: page, totalPages: totalPages, pageSize: pageSize }
    });

  } catch (err) {
    Logger.log('searchRows ERROR: ' + err.toString());
    return _err('SEARCH_ERROR', err.message, err.stack);
  }
}

function _filterByField(rows, field, value) {
  var v = value.toLowerCase();
  return rows.filter(function(row) {
    return (row[field] || '').toLowerCase() === v;
  });
}

function _findHeaderIndexLoose_(headers, fieldName) {
  var target = String(fieldName || '').trim().replace(/\s+/g, ' ').toLowerCase();
  if (!target) return -1;
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim().replace(/\s+/g, ' ').toLowerCase();
    if (key === target) return i;
  }
  return -1;
}

// ========================================
// API: updateRow
// ========================================
function _requireWriteAccess_(token) {
  var accessRes = _getAccessByToken_(token);
  if (!accessRes.ok) return accessRes;
  if (!accessRes.data || accessRes.data.canEdit !== true) {
    return _err('FORBIDDEN', 'Tài khoản hiện không có quyền chỉnh sửa.');
  }
  return accessRes;
}

function _writeChangesToSheet_(sheet, headers, rowNumber, changes, allowedFields, rowSnapshot, access) {
  var allowedMap = {};
  var currentRow = rowSnapshot || _getRowSnapshotByNumber_(sheet, headers, rowNumber);
  var changeDetails = [];
  var editLogUser = _getEditLogUser_(access);
  for (var i = 0; i < allowedFields.length; i++) allowedMap[allowedFields[i]] = true;

  for (var field in changes) {
    if (!changes.hasOwnProperty(field)) continue;
    if (!allowedMap[field]) {
      Logger.log('Skipping non-editable: ' + field);
      continue;
    }

    var colIdx = headers.indexOf(field);
    if (colIdx < 0) colIdx = _findHeaderIndexLoose_(headers, field);
    if (colIdx < 0) {
      Logger.log('Field not in headers: ' + field);
      continue;
    }

    var before = currentRow[field] == null ? '' : String(currentRow[field]);
    var valueToWrite = changes[field];
    if (field === 'Present Date' && valueToWrite) {
      var d = new Date(valueToWrite);
      if (!isNaN(d.getTime())) valueToWrite = d;
    }
    var after = _normalizeAuditValue_(field, valueToWrite);
    if (before === after) continue;

    var cell = sheet.getRange(rowNumber, colIdx + 1);
    cell.setValue(valueToWrite);
    _appendEditHistoryNote_(cell, before, after, editLogUser);
    currentRow[field] = after;
    changeDetails.push({ field: field, before: before, after: after });
  }

  return {
    updatedCount: changeDetails.length,
    changeDetails: changeDetails,
    rowSnapshot: currentRow
  };
}

function updateRow(payload) {
  try {
    Logger.log('updateRow: row=' + payload.rowNumber + ' changes=' + JSON.stringify(payload.changes));

    if (!payload || !payload.rowNumber || !payload.changes) return _err('INVALID_PAYLOAD', 'Missing rowNumber or changes');

    var accessRes = _requireWriteAccess_(payload.token);
    if (!accessRes.ok) return accessRes;
    var validation = _validateChanges(payload.changes);
    if (!validation.ok) return validation;

    var sheet = _openSheet(accessRes.data);
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var allowedFields = _getEditableFieldsForAccess_(accessRes.data);

    var lastRow = sheet.getLastRow();
    if (payload.rowNumber < CONFIG.DATA_START_ROW || payload.rowNumber > lastRow) {
      return _err('INVALID_ROW', 'Row ' + payload.rowNumber + ' is out of range (' +
        CONFIG.DATA_START_ROW + '-' + lastRow + ')');
    }

    var rowSnapshot = _getRowSnapshotByNumber_(sheet, headers, payload.rowNumber);
    if (!_hasRowAccess_(rowSnapshot, accessRes.data)) {
      return _err('FORBIDDEN_SCOPE', 'Tai khoan hien khong duoc phep truy cap san pham nay.');
    }

    var requestId = Utilities.getUuid();
    var source = _resolveAuditSource_(payload, accessRes.data, accessRes.data.role === 'DIRECTOR' ? 'DIRECTOR_UPDATE_ROW' : 'PIC_UPDATE_ROW');
    var applied = _writeChangesToSheet_(sheet, headers, payload.rowNumber, payload.changes, allowedFields, rowSnapshot, accessRes.data);
    var user = '';
    try { user = Session.getActiveUser().getEmail(); } catch(e) {}
    Logger.log('updateRow SUCCESS: row=' + payload.rowNumber + ' fields=' + applied.updatedCount + ' by=' + user + ' source=' + source + ' requestId=' + requestId);

    return _ok({
      rowNumber: payload.rowNumber,
      updatedFields: applied.updatedCount,
      requestId: requestId,
      source: source,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    Logger.log('updateRow ERROR: ' + err.toString());
    return _err('UPDATE_ERROR', err.message, err.stack);
  }
}

// ========================================
// API: updateBatch
// ========================================
function updateBatch(payload) {
  try {
    if (!payload || !Array.isArray(payload.updates)) return _err('INVALID_PAYLOAD', 'Missing updates array');

    var accessRes = _requireWriteAccess_(payload.token);
    if (!accessRes.ok) return accessRes;

    var sheet = _openSheet(accessRes.data);
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var lastRow = sheet.getLastRow();
    var allowedFields = _getEditableFieldsForAccess_(accessRes.data);
    var requestId = Utilities.getUuid();
    var source = _resolveAuditSource_(payload, accessRes.data, accessRes.data.role === 'DIRECTOR' ? 'DIRECTOR_UPDATE_BATCH' : 'PIC_UPDATE_BATCH');

    var results = [];
    var successful = 0;
    var failed = 0;

    for (var i = 0; i < payload.updates.length; i++) {
      var update = payload.updates[i];
      if (!update || !update.rowNumber || !update.changes) {
        failed++;
        results.push({ rowNumber: update && update.rowNumber ? update.rowNumber : null, success: false, error: { code: 'INVALID_PAYLOAD', message: 'Missing rowNumber or changes' } });
        continue;
      }

      var validation = _validateChanges(update.changes);
      if (!validation.ok) {
        failed++;
        results.push({ rowNumber: update.rowNumber, success: false, error: validation.error });
        continue;
      }

      if (update.rowNumber < CONFIG.DATA_START_ROW || update.rowNumber > lastRow) {
        failed++;
        results.push({
          rowNumber: update.rowNumber,
          success: false,
          error: { code: 'INVALID_ROW', message: 'Row ' + update.rowNumber + ' is out of range (' + CONFIG.DATA_START_ROW + '-' + lastRow + ')' }
        });
        continue;
      }

      try {
        var rowSnapshot = _getRowSnapshotByNumber_(sheet, headers, update.rowNumber);
        if (!_hasRowAccess_(rowSnapshot, accessRes.data)) {
          failed++;
          results.push({
            rowNumber: update.rowNumber,
            success: false,
            error: { code: 'FORBIDDEN_SCOPE', message: 'Tai khoan hien khong duoc phep truy cap san pham nay.' }
          });
          continue;
        }

        var applied = _writeChangesToSheet_(sheet, headers, update.rowNumber, update.changes, allowedFields, rowSnapshot, accessRes.data);
        successful++;
        results.push({ rowNumber: update.rowNumber, success: true, updatedFields: applied.updatedCount });
      } catch (writeErr) {
        failed++;
        results.push({ rowNumber: update.rowNumber, success: false, error: { code: 'UPDATE_ERROR', message: writeErr.message } });
      }
    }

    Logger.log('updateBatch SUCCESS: total=' + payload.updates.length + ' successful=' + successful + ' failed=' + failed + ' source=' + source + ' requestId=' + requestId);

    return _ok({
      totalUpdates: payload.updates.length,
      successful: successful,
      failed: failed,
      results: results,
      requestId: requestId,
      source: source,
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    Logger.log('updateBatch ERROR: ' + err.toString());
    return _err('BATCH_ERROR', err.message, err.stack);
  }
}

function _sanitizeUploadFileName_(name) {
  var raw = String(name || '').trim();
  if (!raw) raw = 'uom_image_' + Date.now();
  raw = raw.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
  if (!raw) raw = 'uom_image_' + Date.now();
  return raw;
}

function uploadUomImage(payload) {
  try {
    if (!payload || typeof payload !== 'object') {
      return _err('INVALID_PAYLOAD', 'Missing upload payload');
    }

    var accessRes = _requireWriteAccess_(payload.token);
    if (!accessRes.ok) return accessRes;

    var base64 = String(payload.dataBase64 || '').trim();
    if (!base64) return _err('MISSING_FILE', 'Không có dữ liệu ảnh để upload');

    var mimeType = String(payload.mimeType || '').trim().toLowerCase();
    var allowed = { 'image/jpeg': true, 'image/jpg': true, 'image/png': true };
    if (!allowed[mimeType]) {
      return _err('UNSUPPORTED_FILE_TYPE', 'Chỉ hỗ trợ JPEG/PNG');
    }

    var bytes = Utilities.base64Decode(base64);
    if (!bytes || !bytes.length) {
      return _err('INVALID_FILE_DATA', 'Dữ liệu ảnh không hợp lệ');
    }

    var ext = mimeType === 'image/png' ? '.png' : '.jpg';
    var fileName = _sanitizeUploadFileName_(payload.fileName || ('uom_image_' + Date.now() + ext));
    if (fileName.toLowerCase().indexOf('.') < 0) fileName += ext;

    var folderId = String(CONFIG.UOM_IMAGE_FOLDER_ID || '').trim();
    if (!folderId) return _err('MISSING_UPLOAD_FOLDER', 'Chưa cấu hình folder upload ảnh UOM');

    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (folderErr) {
      return _err('UPLOAD_FOLDER_ACCESS_ERROR', 'Không truy cập được folder upload: ' + folderId + '. Vui lòng kiểm tra quyền Drive.');
    }

    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('uploadUomImage sharing warning: ' + shareErr.message);
    }

    var fileId = file.getId();
    var url = 'https://drive.google.com/file/d/' + fileId + '/view?usp=drivesdk';
    return _ok({
      fileId: fileId,
      name: file.getName(),
      mimeType: mimeType,
      size: file.getSize(),
      url: url
    });
  } catch (err) {
    return _err('UOM_UPLOAD_ERROR', err.message, err.stack);
  }
}

// ========================================
// VALIDATION
// ========================================
function _splitCsvListPreserve_(raw) {
  return String(raw == null ? '' : raw).split(',').map(function(item) {
    return String(item == null ? '' : item).trim();
  });
}

function _trimTrailingEmptyTokens_(list) {
  var out = Array.isArray(list) ? list.slice() : [];
  while (out.length && !String(out[out.length - 1] || '').trim()) out.pop();
  return out;
}

function _getPogActionIntent_(raw) {
  var txt = String(raw == null ? '' : raw).toLowerCase().replace(/\s+/g, ' ').trim();
  if (!txt) return '';
  if (txt.indexOf('cut product') >= 0 || txt === 'cut') return 'cut';
  if (txt.indexOf('reduce facing') >= 0 || txt.indexOf('reduce product facing') >= 0 || txt === 'reduce') return 'reduce';
  return '';
}

function _normalizePogActionForSlot_(raw, slotNumber) {
  var intent = _getPogActionIntent_(raw);
  if (!intent) return '';
  return intent === 'cut'
    ? ('Cut Product ' + slotNumber)
    : ('Reduce facing product ' + slotNumber);
}

function _validatePogFields_(productsRaw, actionsRaw) {
  var errors = [];
  var products = _trimTrailingEmptyTokens_(_splitCsvListPreserve_(productsRaw));
  var actions = _trimTrailingEmptyTokens_(_splitCsvListPreserve_(actionsRaw));
  var seen = {};

  if (products.length > 5) errors.push('POG reduce product facing supports up to 5 products.');
  if (actions.length > 5) errors.push('POG reduce action supports up to 5 slots.');

  for (var i = 0; i < products.length; i++) {
    var product = String(products[i] || '').trim();
    var action = String(actions[i] || '').trim();
    if (!product) continue;
    var productKey = product.toUpperCase();
    if (seen[productKey]) errors.push('POG reduce product cannot repeat the same product.');
    seen[productKey] = true;
    if (!action) {
      errors.push('POG action is missing for Product ' + (i + 1) + '.');
      continue;
    }
    if (_normalizePogActionForSlot_(action, i + 1) !== action) {
      errors.push('POG action must match the same slot number as its product.');
    }
  }

  for (var j = products.length; j < actions.length; j++) {
    if (String(actions[j] || '').trim()) errors.push('POG action cannot exist without a selected product.');
  }

  return errors;
}

function _validateChanges(changes) {
  var errors = [];

  var productName = changes['PIC input Product Name (Dưới 40 kí tự)'];
  if (productName !== undefined && productName !== null) {
    if (String(productName).length > 40) errors.push('Product Name phải ≤ 40 ký tự');
  }

  var desc = changes['Product description (Dưới 250 Kí tự)'];
  if (desc !== undefined && desc !== null) {
    if (String(desc).length > 250) errors.push('Product description phải ≤ 250 ký tự');
  }

  if ('Display area (POG CONCEPT) Bắt Buộc' in changes) {
    var display = changes['Display area (POG CONCEPT) Bắt Buộc'];
    if (!display || String(display).trim() === '') errors.push('Display area là bắt buộc');
  }

  if ('Present Date' in changes) {
    var date = changes['Present Date'];
    if (!date || String(date).trim() === '') errors.push('Present Date là bắt buộc');
  }

  if ('Status' in changes) {
    var status = changes['Status'];
    if (!status || String(status).trim() === '') errors.push('Status là bắt buộc');
  }

  if (changes['Retail Price (+VAT) in Tier 6'] !== undefined) {
    var price = parseFloat(String(changes['Retail Price (+VAT) in Tier 6']).replace(/[^0-9.-]/g, ''));
    if (isNaN(price) || price < 0) errors.push('Retail Price phải là số ≥ 0');
  }

  if (changes['Pricing Strategy'] !== undefined) {
    var pricing = String(changes['Pricing Strategy'] || '').trim();
    if (pricing && ['Tier 7', 'Tier 8', 'Tier 9'].indexOf(pricing) < 0) {
      errors.push('Pricing Strategy must be Tier 7, Tier 8, or Tier 9.');
    }
  }

  errors = errors.concat(_validatePogFields_(
    changes['POG_Choose Reduce product facing'],
    changes['POG_Choose type Reduce product facing']
  ));

  if (errors.length > 0) {
    return _err('VALIDATION_ERROR', errors.join('; '));
  }
  return _ok(true);
}

// ========================================
// API: getMDProductIDs
// Returns all product IDs from [Data] MD Product col A (A3:A)
// Falls back to fuzzy-match if exact name not found
// ========================================
function getMDProductIDs(params) {
  try {
    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;
    var md = _readMDProductTable_();
    var allNames = md.available || [];
    Logger.log('getMDProductIDs: available sheets = ' + allNames.join(' | '));
    if (!md.sheet) {
      Logger.log('getMDProductIDs: no matching sheet found. Available: ' + allNames.join(', '));
      return _ok({ ids: [], items: [], debug: 'Sheet not found. Available: ' + allNames.join(', ') });
    }
    if (!md.headers.length) return _ok({ ids: [], items: [] });

    var idCol = -1;
    var nameCol = -1;
    var nameAliases = ['product_short_name', 'product_name', 'product full name', 'item_name', 'name'];
    for (var h = 0; h < md.headers.length; h++) {
      var header = String(md.headers[h] || '').trim().toLowerCase();
      if (header === 'product_id') idCol = h;
      if (nameAliases.indexOf(header) >= 0 && nameCol < 0) nameCol = h;
    }
    if (idCol < 0) {
      Logger.log('getMDProductIDs: product_id column not found');
      return _ok({ ids: [], items: [], debug: 'product_id column not found' });
    }

    var ids = [];
    var items = [];
    var seen = {};
    for (var i = 0; i < md.rows.length; i++) {
      var row = md.rows[i] || [];
      var pid = String(row[idCol] == null ? '' : row[idCol]).trim();
      if (!pid) continue;
      var key = pid.toUpperCase();
      if (seen[key]) continue;
      seen[key] = true;
      ids.push(pid);
      var name = nameCol >= 0 ? String(row[nameCol] == null ? '' : row[nameCol]).trim() : '';
      items.push(name ? (pid + '_' + name) : pid);
    }
    Logger.log('getMDProductIDs: ' + ids.length + ' IDs loaded, ' + items.length + ' display labels built');
    return _ok({ ids: ids, items: items });
/*
    // Try exact name first, then fallback variations
    var candidates = [
      '[Data} MD Product',
      '[Data] MD Product',
      '[Data]MD Product',
      'MD Product',
      '[Data} MD Products',
      '[Data] MD Products'
    ];
    var sheet = null;
    for (var c = 0; c < candidates.length; c++) {
      sheet = ss.getSheetByName(candidates[c]);
      if (sheet) { Logger.log('getMDProductIDs: matched sheet "' + candidates[c] + '"'); break; }
    }
    // Last resort: fuzzy match — any sheet whose name contains "MD Product" (case-insensitive)
    if (!sheet) {
      for (var s = 0; s < allSheets.length; s++) {
        if (allSheets[s].getName().toLowerCase().indexOf('md product') >= 0) {
          sheet = allSheets[s];
          Logger.log('getMDProductIDs: fuzzy-matched sheet "' + sheet.getName() + '"');
          break;
        }
      }
    }
    if (!sheet) {
      Logger.log('getMDProductIDs: no matching sheet found. Available: ' + allNames.join(', '));
      return _ok({ ids: [], debug: 'Sheet not found. Available: ' + allNames.join(', ') });
    }

    var lastRow = sheet.getLastRow();
    Logger.log('getMDProductIDs: lastRow=' + lastRow);
    if (lastRow < 3) return _ok({ ids: [] });
    var vals = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
    var ids = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] === null || vals[i][0] === undefined ? '' : vals[i][0]).trim();
      if (v) ids.push(v);
    }
    Logger.log('getMDProductIDs: ' + ids.length + ' IDs loaded');
    return _ok({ ids: ids });
*/
  } catch (err) {
    Logger.log('getMDProductIDs ERROR: ' + err.toString());
    return _err('MD_PRODUCT_ERROR', err.message, err.stack);
  }
}

function _renderIndexAsGuest_() {
  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appUserJson = JSON.stringify({
    email: 'guest@npa.local',
    fullName: 'Guest User',
    role: 'GUEST',
    scope: ['ALL'],
    canEdit: false,
    authToken: ''
  }).replace(/</g, '\\u003c');
  var html = tpl.evaluate();
  html.setTitle('NPA Product Dashboard');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getMDProductDetail(payload) {
  try {
    var params = (payload && typeof payload === 'object' && !Array.isArray(payload))
      ? payload
      : { productId: payload };
    var accessRes = _resolveAccessForDataApi_(params);
    if (!accessRes.ok) return accessRes;

    var pid = String(params.productId || '').trim();
    if (!pid) return _err('MISSING_PRODUCT_ID', 'Thiếu Reference Product ID');

    var md = _readMDProductTable_();
    if (!md.sheet) return _err('MD_PRODUCT_SHEET_NOT_FOUND', 'Không tìm thấy sheet [Data} MD Product');
    if (!md.headers.length) return _err('MD_PRODUCT_HEADER_NOT_FOUND', 'Không đọc được header MD Product');

    var productIdCol = md.headers.findIndex(function(h) {
      return String(h || '').trim().toLowerCase() === 'product_id';
    });
    if (productIdCol < 0) return _err('MD_PRODUCT_SCHEMA_ERROR', 'Sheet MD Product thiếu cột product_id');

    var found = null;
    for (var i = 0; i < md.rows.length; i++) {
      if (String(md.rows[i][productIdCol] || '').trim() === pid) {
        found = md.rows[i];
        break;
      }
    }
    if (!found) return _err('MD_PRODUCT_NOT_FOUND', 'Không tìm thấy Reference Product ID: ' + pid);

    var row = {};
    for (var c = 0; c < md.headers.length; c++) row[md.headers[c]] = found[c];

    var retailT6 = _getMDNumber_(row, ['HCM_T6_retail_price_with_tax']);
    var retailT7 = _getMDNumber_(row, ['HCM_T7_retail_price_with_tax']);
    var retailT8 = _getMDNumber_(row, ['HCM_T8_retail_price_with_tax']);
    var retailT9 = _getMDNumber_(row, ['HCM_T9_retail_price_with_tax', 'HN_T9_retail_price_with_tax']);
    var retailIntl = _getMDNumber_(row, ['INTL_retail_price_with_tax']);
    var vat = _getMDNumber_(row, ['input_vat_percent', 'tax']);
    var purchasePriceNoVat = _getMDNumber_(row, ['hcm_purchase_price_no_tax', 'hcm_original_purchase_price_no_tax']);

    var data = {
      productId: pid,
      name: _getMDValue_(row, ['product_short_name']),
      vat: vat,
      purchasePriceNoVat: purchasePriceNoVat,
      retailT6: retailT6,
      retailT7: retailT7,
      retailT8: retailT8,
      retailT9: retailT9,
      retailIntl: retailIntl
    };

    data.gpValueT6 = _calcRefGpValue_(retailT6, purchasePriceNoVat, vat);
    data.gpValueT7 = _calcRefGpValue_(retailT7, purchasePriceNoVat, vat);
    data.gpValueT8 = _calcRefGpValue_(retailT8, purchasePriceNoVat, vat);
    data.gpValueT9 = _calcRefGpValue_(retailT9, purchasePriceNoVat, vat);
    data.gpValueIntl = _calcRefGpValue_(retailIntl, purchasePriceNoVat, vat);

    data.gpPctT6 = _calcRefGpPct_(data.gpValueT6, retailT6);
    data.gpPctT7 = _calcRefGpPct_(data.gpValueT7, retailT7);
    data.gpPctT8 = _calcRefGpPct_(data.gpValueT8, retailT8);
    data.gpPctT9 = _calcRefGpPct_(data.gpValueT9, retailT9);
    data.gpPctIntl = _calcRefGpPct_(data.gpValueIntl, retailIntl);

    return _ok(data);
  } catch (err) {
    Logger.log('getMDProductDetail ERROR: ' + err.toString());
    return _err('MD_PRODUCT_DETAIL_ERROR', err.message, err.stack);
  }
}

function _openMDProductSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var allSheets = ss.getSheets();
  var allNames = allSheets.map(function(s) { return s.getName(); });
  var candidates = [
    '[Data} MD Product',
    '[Data] MD Product',
    '[Data]MD Product',
    'MD Product',
    '[Data} MD Products',
    '[Data] MD Products'
  ];
  for (var i = 0; i < candidates.length; i++) {
    var candidate = ss.getSheetByName(candidates[i]);
    if (candidate) return { sheet: candidate, available: allNames };
  }
  for (var s = 0; s < allSheets.length; s++) {
    if (allSheets[s].getName().toLowerCase().indexOf('md product') >= 0) {
      return { sheet: allSheets[s], available: allNames };
    }
  }
  return { sheet: null, available: allNames };
}

function _readMDProductTable_() {
  var opened = _openMDProductSheet_();
  if (!opened.sheet) return { sheet: null, available: opened.available, headers: [], rows: [] };

  var sheet = opened.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { sheet: sheet, available: opened.available, headers: [], rows: [] };

  var sampleRows = Math.min(5, lastRow);
  var top = sheet.getRange(1, 1, sampleRows, lastCol).getValues();
  var headerRow = 1;
  for (var r = 0; r < top.length; r++) {
    for (var c = 0; c < top[r].length; c++) {
      if (String(top[r][c] || '').trim().toLowerCase() === 'product_id') {
        headerRow = r + 1;
        break;
      }
    }
  }

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var rows = lastRow > headerRow ? sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues() : [];
  return { sheet: sheet, available: opened.available, headers: headers, rows: rows };
}

function _getMDValue_(row, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var target = String(aliases[i] || '').trim().toLowerCase();
    for (var key in row) {
      if (String(key || '').trim().toLowerCase() === target) return row[key];
    }
  }
  return '';
}

function _getMDNumber_(row, aliases) {
  var value = _getMDValue_(row, aliases);
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  var raw = String(value).replace(/[^0-9.-]/g, '');
  var num = parseFloat(raw);
  return isNaN(num) ? 0 : num;
}

function _calcRefGpValue_(retail, purchasePriceNoVat, vat) {
  if (!retail || !purchasePriceNoVat) return 0;
  return retail - (purchasePriceNoVat * (1 + (vat || 0)));
}

function _calcRefGpPct_(gpValue, retail) {
  if (!gpValue || !retail) return 0;
  return (gpValue / retail) * 100;
}

function loginWithEmail(email, password) {
  try {
    var access = _getUserAccessByEmail_(email, password, true);
    if (!access.ok) return access;

    var page = access.data.role === 'DIRECTOR' ? 'director' : 'index';
    var token = _createAuthToken_(access.data);
    var baseUrl = ScriptApp.getService().getUrl();
    var targetUrl = baseUrl
      ? (baseUrl + '?page=' + page + '&token=' + encodeURIComponent(token))
      : ('?page=' + page + '&token=' + encodeURIComponent(token));
    return _ok({
      email: access.data.email,
      role: access.data.role,
      page: page,
      token: token,
      url: targetUrl
    });
  } catch (err) {
    return _err('LOGIN_ERROR', err.message, err.stack);
  }
}

function getCurrentUserAccess(token) {
  return _getAccessByToken_(token);
}

function _renderLoginPage(message, prefillEmail) {
  var tpl = HtmlService.createTemplateFromFile('Login');
  tpl.loginMessage = message || '';
  tpl.prefillEmail = prefillEmail || '';
  var out = tpl.evaluate();
  out.setTitle('NPA Login');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

function _resolveRequestEmail_(e) {
  var fromQuery = e && e.parameter && e.parameter.email ? String(e.parameter.email).trim() : '';
  if (fromQuery) return fromQuery.toLowerCase();
  try {
    var me = Session.getActiveUser().getEmail();
    return me ? String(me).trim().toLowerCase() : '';
  } catch (err) {
    return '';
  }
}

function _normalizeUserMasterHeader_(text) {
  return String(text == null ? '' : text)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _findUserMasterHeaderIndex_(headers, candidates) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[_normalizeUserMasterHeader_(headers[i])] = i;
  }
  for (var j = 0; j < candidates.length; j++) {
    var idx = map[_normalizeUserMasterHeader_(candidates[j])];
    if (idx !== undefined) return idx;
  }
  return -1;
}

function _normalizeLoginEmail_(value) {
  var raw = String(value == null ? '' : value);
  if (raw && typeof raw.normalize === 'function') raw = raw.normalize('NFKC');
  return raw
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function _getUserAccessByEmail_(email, password, enforcePassword) {
  try {
    var em = _normalizeLoginEmail_(email);
    if (!em) return _err('MISSING_EMAIL', 'Không tìm thấy email đăng nhập');
    var pwd = String(password || '');

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
    if (!sheet) return _err('USER_MASTER_NOT_FOUND', 'Chưa có sheet User_Master');

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return _err('USER_NOT_FOUND', 'User_Master chưa có dữ liệu');

    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    var hEmail = _findUserMasterHeaderIndex_(headers, ['Email']);
    var hPassword = _findUserMasterHeaderIndex_(headers, ['Password']);
    var hName = _findUserMasterHeaderIndex_(headers, ['Full Name', 'FullName']);
    var hRole = _findUserMasterHeaderIndex_(headers, ['Role']);
    var hStatus = _findUserMasterHeaderIndex_(headers, ['Status']);
    var hCategory = _findUserMasterHeaderIndex_(headers, ['Category']);
    var hManaged = _findUserMasterHeaderIndex_(headers, ['Managed Categories', 'Managed Category']);
    var hLanding = _findUserMasterHeaderIndex_(headers, ['Landing Page (Auto)', 'Landing Page']);
    var hCanEdit = _findUserMasterHeaderIndex_(headers, ['Can Edit', 'Permission']);
    var hScope = _findUserMasterHeaderIndex_(headers, ['Effective Scope (Auto)', 'Effective Scope']);

    if (hEmail < 0 || hRole < 0 || hStatus < 0) {
      return _err('USER_MASTER_SCHEMA_ERROR', 'Thiếu cột Email/Role/Status trong User_Master');
    }

    var found = null;
    for (var i = 0; i < data.length; i++) {
      var rowEmail = _normalizeLoginEmail_(data[i][hEmail]);
      if (rowEmail === em) {
        found = data[i];
        break;
      }
    }
    if (!found) {
      Logger.log('USER_NOT_FOUND lookup=' + em + ' headers=' + headers.join(' | '));
      return _err('USER_NOT_FOUND', 'Không tìm thấy tài khoản trong User_Master cho email: ' + em);
    }

    var status = String(found[hStatus] || '').trim().toUpperCase();
    if (status !== 'ACTIVE') return _err('USER_INACTIVE', 'Tài khoản đang bị khóa (INACTIVE)');

    if (enforcePassword) {
      if (hPassword < 0) return _err('MISSING_PASSWORD_COLUMN', 'Thiếu cột Password trong User_Master');
      if (!pwd) return _err('MISSING_PASSWORD', 'Vui lòng nhập mật khẩu');
      var sheetPwd = String(found[hPassword] || '');
      if (pwd !== sheetPwd) return _err('INVALID_PASSWORD', 'Mật khẩu không đúng');
    }

    var role = String(found[hRole] || '').trim().toUpperCase();
    if (role !== 'DIRECTOR' && role !== 'MANAGER' && role !== 'CATEGORY') {
      return _err('INVALID_ROLE', 'Role không hợp lệ: ' + role);
    }

    var fullName = hName >= 0 ? String(found[hName] || '').trim() : '';
    var canEditRaw = hCanEdit >= 0 ? found[hCanEdit] : true;
    var canEdit = (canEditRaw === true || String(canEditRaw).toLowerCase() === 'true');
    var landing = hLanding >= 0 ? String(found[hLanding] || '').trim() : '';
    var scopeRaw = hScope >= 0 ? String(found[hScope] || '').trim() : '';
    var category = hCategory >= 0 ? String(found[hCategory] || '').trim() : '';
    var managed = hManaged >= 0 ? String(found[hManaged] || '').trim() : '';

    // Merge all possible scope columns to avoid missing categories when one column is stale/empty.
    var scopeJoined = []
      .concat(_parseScopeList_(scopeRaw))
      .concat(_parseScopeList_(category))
      .concat(_parseScopeList_(managed));
    var scopeSeen = {};
    var scope = [];
    for (var si = 0; si < scopeJoined.length; si++) {
      var raw = scopeJoined[si];
      var key = _normalizeCategoryKey_(raw);
      if (!key || scopeSeen[key]) continue;
      scopeSeen[key] = true;
      scope.push(raw);
    }
    if (role === 'DIRECTOR') scope = ['ALL'];

    return _ok({
      email: em,
      fullName: fullName,
      role: role,
      status: status,
      canEdit: canEdit,
      landingPage: landing || (role === 'DIRECTOR' ? 'DIRECTOR_VIEW' : 'INDEX'),
      scope: scope
    });
  } catch (err) {
    return _err('USER_ACCESS_ERROR', err.message, err.stack);
  }
}

function _createAuthToken_(accessData) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var cache = CacheService.getScriptCache();
  cache.put('AUTH_' + token, JSON.stringify(accessData || {}), 21600); // 6h
  return token;
}

function _getAccessByToken_(token) {
  try {
    var tk = String(token || '').trim();
    if (!tk) return _err('MISSING_TOKEN', 'Phiên đăng nhập không hợp lệ. Vui lòng đăng nhập lại.');

    var raw = CacheService.getScriptCache().get('AUTH_' + tk);
    if (!raw) return _err('TOKEN_EXPIRED', 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
    var data = JSON.parse(raw);
    if (!data || !data.email) return _err('INVALID_TOKEN', 'Phiên đăng nhập không hợp lệ.');
    return _ok(data);
  } catch (err) {
    return _err('TOKEN_READ_ERROR', err.message, err.stack);
  }
}

// ========================================
// ADMIN: Setup User_Master for Role-based Login
// Roles: DIRECTOR > MANAGER > CATEGORY
// ========================================
function setupUserMasterSheet() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.USER_MASTER_SHEET);
    }

    var headers = [
      'Email',
      'Password',
      'Full Name',
      'Role',
      'Status',
      'Category',
      'Can Edit',
      'Notes',
      'Created At',
      'Updated At'
    ];

    var notes = [
      'Email đăng nhập (nên dùng email domain công ty), duy nhất',
      'Mật khẩu đăng nhập (hiện đang lưu dạng plain text)',
      'Tên hiển thị',
      'DIRECTOR | MANAGER | CATEGORY',
      'ACTIVE | INACTIVE',
      'Category/scope được phép xem (có thể nhiều giá trị, ngăn cách bằng dấu phẩy)',
      'TRUE/FALSE quyền chỉnh sửa',
      'Ghi chú nội bộ',
      'Thời điểm tạo tài khoản',
      'Thời điểm cập nhật gần nhất'
    ];

    var requiredCols = headers.length;
    if (sheet.getMaxColumns() < requiredCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < 200) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 200 - sheet.getMaxRows());
    }

    sheet.getRange(1, 1, 1, requiredCols).setValues([headers]);
    sheet.getRange(1, 1, 1, requiredCols).setNotes([notes]);
    sheet.setFrozenRows(1);

    var headerRange = sheet.getRange(1, 1, 1, requiredCols);
    headerRange
      .setFontWeight('bold')
      .setBackground('#1f4e78')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    var widths = [260, 140, 180, 130, 120, 360, 110, 280, 170, 170];
    for (var i = 0; i < widths.length; i++) {
      sheet.setColumnWidth(i + 1, widths[i]);
    }

    var maxRows = sheet.getMaxRows();
    var dataRows = maxRows - 1;

    // Data validation rules
    var roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['DIRECTOR', 'MANAGER', 'CATEGORY'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 4, dataRows, 1).setDataValidation(roleRule);

    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ACTIVE', 'INACTIVE'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 5, dataRows, 1).setDataValidation(statusRule);

    var editRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    sheet.getRange(2, 7, dataRows, 1).setDataValidation(editRule);

    // Remove legacy helper column if present
    var legacyHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var helperColIdx = legacyHeader.indexOf('Category_List_Helper');
    if (helperColIdx >= 0) {
      sheet.deleteColumn(helperColIdx + 1);
      // ensure header row is still correct after deleting old column
      sheet.getRange(1, 1, 1, requiredCols).setValues([headers]);
    }

    // Category validation from main data sheet
    var categories = _extractCategoriesFromMainSheet_();
    if (categories.length > 0) {
      var catRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(categories, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 6, dataRows, 1).setDataValidation(catRule);
    } else {
      sheet.getRange(2, 6, dataRows, 1).clearDataValidations();
    }

    // Formatting
    sheet.getRange(2, 1, dataRows, 10).setVerticalAlignment('middle');
    sheet.getRange(2, 6, dataRows, 1).setWrap(true); // Category scope
    sheet.getRange(2, 8, dataRows, 1).setWrap(true); // Notes
    sheet.getRange(2, 9, dataRows, 2).setNumberFormat('yyyy-mm-dd hh:mm');

    // Seed sample rows if empty
    var hasAnyUser = false;
    if (sheet.getLastRow() >= 2) {
      var emailVals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      hasAnyUser = emailVals.some(function(r) {
        return String(r[0] || '').trim() !== '';
      });
    }

    if (!hasAnyUser) {
      var now = new Date();
      var c1 = categories[0] || 'Non Alcoholic';
      var seed = [
        ['director@yourcompany.com', '7-Eleven', 'Director User', 'DIRECTOR', 'ACTIVE', '', true, 'Xem toàn bộ sản phẩm', now, now],
        ['manager@yourcompany.com', '7-Eleven', 'Manager User', 'MANAGER', 'ACTIVE', c1, true, 'Quản lý nhóm category', now, now],
        ['category@yourcompany.com', '7-Eleven', 'Category User', 'CATEGORY', 'ACTIVE', c1, true, 'Chỉ xem category được phân quyền', now, now]
      ];
      sheet.getRange(2, 1, seed.length, requiredCols).setValues(seed);
    }

    // Basic filter
    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, sheet.getMaxRows(), 10).createFilter();
    }

    return _ok({
      sheetName: CONFIG.USER_MASTER_SHEET,
      categoriesLoaded: categories.length,
      message: 'User_Master is ready'
    });
  } catch (err) {
    return _err('SETUP_USER_MASTER_ERROR', err.message, err.stack);
  }
}

function setupAuditLogSheet() {
  try {
    var sheet = _ensureAuditLogSheet_();
    return _ok({
      sheetName: sheet.getName(),
      message: 'Audit_Log is ready'
    });
  } catch (err) {
    return _err('SETUP_AUDIT_LOG_ERROR', err.message, err.stack);
  }
}

// Refresh only category helper list + validation in User_Master
function refreshUserMasterCategoryList() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
    if (!sheet) {
      return _err('USER_MASTER_NOT_FOUND', 'Run setupUserMasterSheet() first');
    }

    var categories = _extractCategoriesFromMainSheet_();
    var dataRows = Math.max(1, sheet.getMaxRows() - 1);

    if (categories.length > 0) {
      var catRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(categories, true)
        .setAllowInvalid(true)
        .build();
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var hCategory = headers.indexOf('Category');
      if (hCategory >= 0) {
        sheet.getRange(2, hCategory + 1, dataRows, 1).setDataValidation(catRule);
      }
    } else {
      var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idxCategory = hdr.indexOf('Category');
      if (idxCategory >= 0) {
        sheet.getRange(2, idxCategory + 1, dataRows, 1).clearDataValidations();
      }
    }

    return _ok({ categoriesLoaded: categories.length });
  } catch (err) {
    return _err('REFRESH_USER_MASTER_CATEGORY_ERROR', err.message, err.stack);
  }
}

function _extractCategoriesFromMainSheet_() {
  var sheet = _openSheet();
  var headers = _getHeaders(sheet);
  var idx = headers.indexOf('Category');
  if (idx < 0) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];

  var vals = sheet
    .getRange(CONFIG.DATA_START_ROW, idx + 1, lastRow - CONFIG.DATA_START_ROW + 1, 1)
    .getValues();

  var seen = {};
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] === null || vals[i][0] === undefined ? '' : vals[i][0]).trim();
    if (v) seen[v] = true;
  }
  return Object.keys(seen).sort();
}

// ========================================
// DEBUG: list all sheet names (call from Apps Script editor to diagnose)
// ========================================
function debugListSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var names = ss.getSheets().map(function(s) { return '"' + s.getName() + '"'; });
  Logger.log('All sheets (' + names.length + '):\n' + names.join('\n'));
  return names;
}

// ========================================
// RESPONSE BUILDERS
// ========================================
function _ok(data) {
  return { ok: true, data: data };
}

function _err(code, message, details) {
  var safeMessage = _safeErrorText_(message, 'Unknown error');
  var safeDetails = details || '';
  if (typeof safeDetails !== 'string') {
    safeDetails = _safeErrorText_(safeDetails, '');
  }
  return {
    ok: false,
    error: { code: code, message: safeMessage, details: safeDetails }
  };
}
