// ========================================
// DIRECTOR VIEW DATA ADAPTER
// ========================================

var DIRECTOR_VIEW_CONFIG = {
  SHEET_NAME: 'Director'
};

function _getDirectorProtectedWriteFieldNames_() {
  var out = ['SSV Note Product', 'CEO & OPs Note'];
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.CHAT_FIELD_NAME) out.push(CONFIG.CHAT_FIELD_NAME);
  } catch (err) {
    // Keep defaults when global CONFIG is unavailable.
  }
  return out;
}

function _isDirectorProtectedWriteField_(fieldName) {
  var key = _directorNormHeader_(fieldName);
  if (!key) return false;
  var protectedFields = _getDirectorProtectedWriteFieldNames_();
  for (var i = 0; i < protectedFields.length; i++) {
    if (_directorNormHeader_(protectedFields[i]) === key) return true;
  }
  return false;
}

function _assertDirectorWriteChangesAllowed_(changes) {
  var data = changes || {};
  for (var field in data) {
    if (!Object.prototype.hasOwnProperty.call(data, field)) continue;
    if (_isDirectorProtectedWriteField_(field)) {
      return _err('FORBIDDEN_FIELD', 'Truong "' + field + '" chi duoc cap nhat qua chat API.');
    }
  }
  return _ok(true);
}

var DIRECTOR_VIEW_ALIAS_MAP = {
  'Product Name': ['Product Name', 'PIC input Product Name (Dưới 40 kí tự)'],
  'PIC input Product Name (Dưới 40 kí tự)': ['PIC input Product Name (Dưới 40 kí tự)', 'Product Name'],
  'Ticket ID': ['Ticket ID', 'New Product ID'],
  'Present Status': ['Present Status', 'Status'],
  'Status': ['Status', 'Present Status'],
  'Note': ['Note'],
  'Product Group': ['Product Group', 'Group'],
  'Group': ['Group', 'Product Group'],
  'Zone': ['Zone', 'PIC CATEGORY', 'PIC Category'],
  'URL Image': ['URL Image', 'Url Image'],
  'Url Image': ['Url Image', 'URL Image'],
  'Image Path': ['Image Path', 'Image'],
  'Retail Price (+VAT) in Tier 6': ['Retail Price (+VAT) in Tier 6', 'Retail Price (+VAT) in Tier 7'],
  'Retail Price (+VAT) in Tier 7': ['Retail Price (+VAT) in Tier 7', 'Retail Price (+VAT) in Tier 6'],
  '% GP Retail Price Tier 6': ['% GP Retail Price Tier 6', '% GP Retail Price Tier 7'],
  '$ GP Retail Price Tier 6': ['$ GP Retail Price Tier 6', '$ GP Retail Price Tier 7'],
  'REF PRODUCT\nVAT input': ['REF PRODUCT\nVAT input', 'REF PRODUCT\nVAT'],
  'REF PRODUCT\nVAT': ['REF PRODUCT\nVAT', 'REF PRODUCT\nVAT input']
};

var DIRECTOR_WRITE_ALIAS_MAP = {
  'Status': ['Status', 'Present Status'],
  'Note': ['Note']
};

var DIRECTOR_VIEW_REQUIRED_KEYS = [
  'No.Present',
  'Ticket ID',
  'Product Name',
  'PIC input Product Name (Dưới 40 kí tự)',
  'Status',
  'Present Date',
  'Category',
  'Product Group',
  'Sub Category',
  'Brand Name',
  'Supplier Name',
  'Retail Price (+VAT) in Tier 6',
  'Retail Price (+VAT) in Tier 7',
  'Note',
  'URL Image',
  'Url Image',
  'Image Path',
  'New Product ID',
  'Other Income + Listing Fee',
  'Product description (Dưới 250 Kí tự)',
  'UOM Information',
  'Display area (POG CONCEPT) Bắt Buộc',
  'Preservation temperature',
  'Logistics Group South',
  'Logistics Group North'
];

var DIRECTOR_VIEW_PREFERRED_HEADERS = [
  'No.Present',
  'Ticket ID',
  'Product Name',
  'PIC input Product Name (Dưới 40 kí tự)',
  'Status',
  'Present Date',
  'Category',
  'Product Group',
  'Sub Category',
  'Brand Name',
  'Supplier Name',
  'Retail Price (+VAT) in Tier 6',
  'Retail Price (+VAT) in Tier 7',
  'Note',
  'Other Income + Listing Fee',
  'Product description (Dưới 250 Kí tự)',
  'UOM Information',
  'Display area (POG CONCEPT) Bắt Buộc',
  'Preservation temperature',
  'Logistics Group South',
  'Logistics Group North',
  'URL Image',
  'Url Image',
  'Image',
  'Image Path'
];

function _directorNormHeader_(text) {
  return _directorMobileSafeString_(text, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _directorHasValue_(value) {
  return !(value === null || value === undefined || _directorMobileSafeString_(value, '').trim() === '');
}

function _directorSafeErrorText_(err, fallback) {
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
  } catch (safeErr) {
    return fb;
  }
}

function _directorMobileSafeString_(value, fallback) {
  var fb = fallback || '';
  try {
    if (value === null || value === undefined) return fb;
    if (typeof value === 'string') return value;
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    if (value instanceof Date) return _cellToString(value);
    var s = (typeof value.toString === 'function') ? value.toString() : '';
    if (s && s !== '[object Object]') return s;
    var json = JSON.stringify(value);
    return (json && json !== '{}') ? json : fb;
  } catch (err) {
    return fb;
  }
}

function _directorBuildLookup_(rowObj) {
  var lookup = {};
  for (var key in rowObj) {
    if (!Object.prototype.hasOwnProperty.call(rowObj, key)) continue;
    if (key === '_rowNumber') continue;
    lookup[_directorNormHeader_(key)] = rowObj[key];
  }
  return lookup;
}

function _directorGetByAliases_(lookup, aliases) {
  var list = Array.isArray(aliases) ? aliases : [];
  for (var i = 0; i < list.length; i++) {
    var hit = lookup[_directorNormHeader_(list[i])];
    if (_directorHasValue_(hit)) return hit;
  }
  return '';
}

function _normalizeDirectorRowForView_(rowObj) {
  var out = {};
  for (var key in rowObj) {
    if (!Object.prototype.hasOwnProperty.call(rowObj, key)) continue;
    out[key] = rowObj[key];
  }

  var lookup = _directorBuildLookup_(out);
  for (var target in DIRECTOR_VIEW_ALIAS_MAP) {
    if (!Object.prototype.hasOwnProperty.call(DIRECTOR_VIEW_ALIAS_MAP, target)) continue;
    if (_directorHasValue_(out[target])) continue;
    var aliasedValue = _directorGetByAliases_(lookup, DIRECTOR_VIEW_ALIAS_MAP[target]);
    if (_directorHasValue_(aliasedValue)) {
      out[target] = aliasedValue;
      lookup[_directorNormHeader_(target)] = aliasedValue;
    }
  }

  if (!_directorHasValue_(out['No.Present']) && out._rowNumber) {
    out['No.Present'] = String(Math.max(1, parseInt(out._rowNumber, 10) - (CONFIG.DATA_START_ROW - 1)));
  }
  if (!_directorHasValue_(out['Ticket ID']) && _directorHasValue_(out['New Product ID'])) {
    out['Ticket ID'] = out['New Product ID'];
  }
  if (!_directorHasValue_(out['Present Status']) && _directorHasValue_(out['Status'])) {
    out['Present Status'] = out['Status'];
  }
  if (!_directorHasValue_(out['Product Group']) && _directorHasValue_(out['Group'])) {
    out['Product Group'] = out['Group'];
  }
  if (!_directorHasValue_(out['Zone']) && _directorHasValue_(out['PIC CATEGORY'])) {
    out['Zone'] = out['PIC CATEGORY'];
  }
  if (!_directorHasValue_(out['Retail Price (+VAT) in Tier 7']) && _directorHasValue_(out['Retail Price (+VAT) in Tier 6'])) {
    out['Retail Price (+VAT) in Tier 7'] = out['Retail Price (+VAT) in Tier 6'];
  }
  if (!_directorHasValue_(out['Retail Price (+VAT) in Tier 6']) && _directorHasValue_(out['Retail Price (+VAT) in Tier 7'])) {
    out['Retail Price (+VAT) in Tier 6'] = out['Retail Price (+VAT) in Tier 7'];
  }
  if (!_directorHasValue_(out['% GP Retail Price Tier 6']) && _directorHasValue_(out['% GP Retail Price Tier 7'])) {
    out['% GP Retail Price Tier 6'] = out['% GP Retail Price Tier 7'];
  }
  if (!_directorHasValue_(out['$ GP Retail Price Tier 6']) && _directorHasValue_(out['$ GP Retail Price Tier 7'])) {
    out['$ GP Retail Price Tier 6'] = out['$ GP Retail Price Tier 7'];
  }
  if (!_directorHasValue_(out['Url Image']) && _directorHasValue_(out['URL Image'])) {
    out['Url Image'] = out['URL Image'];
  }
  if (!_directorHasValue_(out['Image Path']) && _directorHasValue_(out['Image'])) {
    out['Image Path'] = out['Image'];
  }
  if (!_directorHasValue_(out['REF PRODUCT\nVAT input']) && _directorHasValue_(out['REF PRODUCT\nVAT'])) {
    out['REF PRODUCT\nVAT input'] = out['REF PRODUCT\nVAT'];
  }

  for (var i = 0; i < DIRECTOR_VIEW_REQUIRED_KEYS.length; i++) {
    var requiredKey = DIRECTOR_VIEW_REQUIRED_KEYS[i];
    if (!_directorHasValue_(out[requiredKey])) out[requiredKey] = '';
  }

  return out;
}

function _normalizeDirectorRowsForView_(rows) {
  if (!Array.isArray(rows)) return [];
  return rows.map(_normalizeDirectorRowForView_);
}

function _directorBuildViewHeaders_(rows, sourceHeaders) {
  var ordered = [];
  var seen = {};

  function addHeader(h) {
    if (!h || h === '_rowNumber') return;
    if (seen[h]) return;
    seen[h] = true;
    ordered.push(h);
  }

  (DIRECTOR_VIEW_PREFERRED_HEADERS || []).forEach(addHeader);
  (Array.isArray(sourceHeaders) ? sourceHeaders : []).forEach(addHeader);
  (Array.isArray(rows) ? rows : []).forEach(function(row) {
    Object.keys(row || {}).forEach(addHeader);
  });

  return ordered;
}

function _openDirectorSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var preferredName = CONFIG.SHEET_NAME_DIRECTOR || DIRECTOR_VIEW_CONFIG.SHEET_NAME;
  var found = _findSheetByLooseName_(ss, [preferredName, 'Director'], ['director']);
  if (!found.sheet) {
    throw new Error('Sheet "' + preferredName + '" not found. Available: ' + (found.allNames || []).join(', '));
  }
  return found.sheet;
}

function _resolveDirectorHeaderName_(headers, field) {
  var aliases = DIRECTOR_WRITE_ALIAS_MAP[field] || [field];
  for (var i = 0; i < aliases.length; i++) {
    var idx = headers.indexOf(aliases[i]);
    if (idx >= 0) return headers[idx];
    idx = _findHeaderIndexLoose_(headers, aliases[i]);
    if (idx >= 0) return headers[idx];
  }
  return '';
}

function _mapDirectorAllowedFields_(headers, allowedFields) {
  var out = [];
  var seen = {};
  for (var i = 0; i < allowedFields.length; i++) {
    var target = _resolveDirectorHeaderName_(headers, allowedFields[i]) || allowedFields[i];
    if (!seen[target]) {
      seen[target] = true;
      out.push(target);
    }
  }
  return out;
}

function _mapDirectorChangesToSheet_(headers, changes) {
  var out = {};
  for (var field in changes) {
    if (!Object.prototype.hasOwnProperty.call(changes, field)) continue;
    var target = _resolveDirectorHeaderName_(headers, field) || field;
    if (_isDirectorProtectedWriteField_(target)) continue;
    out[target] = changes[field];
  }
  return out;
}

function _assertDirectorAccess_(accessRes) {
  if (!accessRes || !accessRes.ok) return accessRes;
  if (!accessRes.data || _getSelectedViewCode_(accessRes.data) !== 'DIR') {
    return _err('FORBIDDEN', 'Chi co Director view moi duoc phep truy cap API nay.');
  }
  return accessRes;
}

function _requireDirectorAccessOrThrow_(params, requireWrite) {
  var accessRes = requireWrite ? _requireWriteAccess_(params && params.token) : _resolveAccessForDataApi_(params);
  accessRes = _assertDirectorAccess_(accessRes);
  if (!accessRes.ok) {
    throw new Error((accessRes.error && accessRes.error.message) || 'Access denied');
  }
  return accessRes.data;
}

function getDirectorInitData(params) {
  try {
    Logger.log('getDirectorInitData START');
    var accessRes = _resolveAccessForDataApi_(params);
    accessRes = _assertDirectorAccess_(accessRes);
    if (!accessRes.ok) return accessRes;

    var editableFields = _getEditableFieldsForAccess_(accessRes.data);
    var sheet = _openDirectorSheet_();
    var lastRow = sheet.getLastRow();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var numCols = headers.length;

    if (numCols === 0 || lastRow < CONFIG.DATA_START_ROW) {
      return _ok({
        headers: _directorBuildViewHeaders_([], headers),
        rows: [],
        editableFields: editableFields,
        meta: {
          totalRows: 0,
          currentPage: 1,
          totalPages: 1,
          pageSize: CONFIG.PAGE_SIZE,
          sheetName: sheet.getName(),
          selectedView: accessRes.data.selectedView || 'DIR',
          currentViewCanEdit: accessRes.data.currentViewCanEdit === true
        }
      });
    }

    var numDataRows = lastRow - CONFIG.DATA_START_ROW + 1;
    var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numDataRows, numCols).getValues();

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

    var mappedRows = _normalizeDirectorRowsForView_(allRows);
    var viewHeaders = _directorBuildViewHeaders_(mappedRows, headers);
    var total = mappedRows.length;

    Logger.log('getDirectorInitData SUCCESS: total=' + total);
    return _ok({
      headers: viewHeaders,
      rows: mappedRows,
      editableFields: editableFields,
      meta: {
        totalRows: total,
        currentPage: 1,
        totalPages: 1,
        pageSize: total || CONFIG.PAGE_SIZE,
        sheetName: sheet.getName(),
        selectedView: accessRes.data.selectedView || 'DIR',
        currentViewCanEdit: accessRes.data.currentViewCanEdit === true
      }
    });
  } catch (err) {
    var msg = _directorSafeErrorText_(err, 'Unknown error');
    Logger.log('getDirectorInitData ERROR: ' + msg);
    return _err('DIRECTOR_INIT_ERROR', msg, err && err.stack ? err.stack : '');
  }
}

function updateDirectorRow(payload) {
  try {
    if (!payload || !payload.rowNumber || !payload.changes) {
      return _err('INVALID_PAYLOAD', 'Missing rowNumber or changes');
    }

    var accessRes = _requireWriteAccess_(payload.token);
    accessRes = _assertDirectorAccess_(accessRes);
    if (!accessRes.ok) return accessRes;

    var protection = _assertDirectorWriteChangesAllowed_(payload.changes);
    if (!protection.ok) return protection;

    var validation = _validateChanges(payload.changes);
    if (!validation.ok) return validation;

    var sheet = _openDirectorSheet_();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var lastRow = sheet.getLastRow();

    if (payload.rowNumber < CONFIG.DATA_START_ROW || payload.rowNumber > lastRow) {
      return _err('INVALID_ROW', 'Row ' + payload.rowNumber + ' is out of range (' +
        CONFIG.DATA_START_ROW + '-' + lastRow + ')');
    }

    var rowSnapshot = _getRowSnapshotByNumber_(sheet, headers, payload.rowNumber);
    if (!_hasRowAccess_(rowSnapshot, accessRes.data)) {
      return _err('FORBIDDEN_SCOPE', 'Tai khoan hien khong duoc phep truy cap san pham nay.');
    }

    var allowedFields = _mapDirectorAllowedFields_(headers, _getEditableFieldsForAccess_(accessRes.data));
    var mappedChanges = _mapDirectorChangesToSheet_(headers, payload.changes);
    var requestId = Utilities.getUuid();
    var source = _resolveAuditSource_(payload, accessRes.data, 'DIRECTOR_UPDATE_ROW');

    var applied = _writeChangesToSheet_(sheet, headers, payload.rowNumber, mappedChanges, allowedFields, rowSnapshot, accessRes.data);

    Logger.log('updateDirectorRow SUCCESS: row=' + payload.rowNumber + ' fields=' + applied.updatedCount + ' source=' + source + ' requestId=' + requestId);

    return _ok({
      rowNumber: payload.rowNumber,
      updatedFields: applied.updatedCount,
      requestId: requestId,
      source: source,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    var msg = _directorSafeErrorText_(err, 'Unknown error');
    Logger.log('updateDirectorRow ERROR: ' + msg);
    return _err('DIRECTOR_UPDATE_ERROR', msg, err && err.stack ? err.stack : '');
  }
}

function updateDirectorBatch(payload) {
  try {
    if (!payload || !Array.isArray(payload.updates)) {
      return _err('INVALID_PAYLOAD', 'Missing updates array');
    }

    var accessRes = _requireWriteAccess_(payload.token);
    accessRes = _assertDirectorAccess_(accessRes);
    if (!accessRes.ok) return accessRes;

    var sheet = _openDirectorSheet_();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var lastRow = sheet.getLastRow();
    var allowedFields = _mapDirectorAllowedFields_(headers, _getEditableFieldsForAccess_(accessRes.data));
    var requestId = Utilities.getUuid();
    var source = _resolveAuditSource_(payload, accessRes.data, 'DIRECTOR_UPDATE_BATCH');

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

      var protection = _assertDirectorWriteChangesAllowed_(update.changes);
      if (!protection.ok) {
        failed++;
        results.push({ rowNumber: update.rowNumber, success: false, error: protection.error });
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

        var mappedChanges = _mapDirectorChangesToSheet_(headers, update.changes);
        var applied = _writeChangesToSheet_(sheet, headers, update.rowNumber, mappedChanges, allowedFields, rowSnapshot, accessRes.data);
        successful++;
        results.push({ rowNumber: update.rowNumber, success: true, updatedFields: applied.updatedCount });
      } catch (writeErr) {
        failed++;
        results.push({
          rowNumber: update.rowNumber,
          success: false,
          error: { code: 'UPDATE_ERROR', message: _directorSafeErrorText_(writeErr, 'Update failed') }
        });
      }
    }

    Logger.log('updateDirectorBatch SUCCESS: total=' + payload.updates.length + ' successful=' + successful + ' failed=' + failed + ' source=' + source + ' requestId=' + requestId);

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
    var msg = _directorSafeErrorText_(err, 'Unknown error');
    Logger.log('updateDirectorBatch ERROR: ' + msg);
    return _err('DIRECTOR_BATCH_ERROR', msg, err && err.stack ? err.stack : '');
  }
}

// ========================================
// DIRECTOR MOBILE VIEW API (Director_View.html)
// ========================================

var DIRECTOR_MOBILE_ROLE_DIRECTOR = 'DIRECTOR';

var DIRECTOR_MOBILE_FIELD_ALIASES = {
  noPresent: ['No.Present', 'No Present'],
  presentDate: ['Present Date'],
  status: ['Status', 'Present Status'],
  note: ['Note'],
  productName: [
    'PIC input Product Name (Duoi 40 ki tu)',
    'PIC input Product Name (Duoi 40 ky tu)',
    'PIC input Product Name',
    'Product Name'
  ],
  newProductType: ['New product type', 'New Product Type'],
  whyListVN: ['Why Do We List This Product?'],
  whyListEN: ['Why Do We List This Product? (Eng)', 'Why Do We List This Product? (EN)'],
  premiumProducts: ['Premium products'],
  branded: ['Branded'],
  picCategory: ['PIC CATEGORY', 'PIC Category'],
  group: ['Group', 'Product Group'],
  category: ['Category'],
  subCategory: ['Sub Category'],
  primaryUsageSegment: ['PRIMARY USAGE SEGMENT'],
  targetStoreTypes: ['TARGET STORE TYPES'],
  targetCustomerStoreTypes: ['TARGET CUSTOMER STORE TYPES'],
  supplierName: ['Supplier Name'],
  countryOfOrigin: ['Country of Origin'],
  brandName: ['Brand Name'],
  purchasePrice: ['Purchase Price Excluded VAT', 'Purchase Price (Excluded VAT)'],
  inboundVAT: ['Inbound VAT'],
  samePriceProduct: ['SAME PRICE PRODUCT'],
  refProductId: [
    'REFERENCE PRODUCT ID (Neu khong co dien NO)',
    'REFERENCE PRODUCT ID'
  ],
  refProductName: ['REFERENCE PRODUCT NAME'],
  priceT6: ['Retail Price (+VAT) in Tier 6'],
  priceT7: ['Retail Price (+VAT) in Tier 7'],
  priceT8: ['Retail Price (+VAT) in Tier 8'],
  priceT9: ['Retail Price (+VAT) in Tier 9'],
  priceINTL: ['Retail Price (+VAT) in INTL', 'Retail Price (+VAT) in Intl'],
  gpPctT6: ['% GP Retail Price Tier 6'],
  gpPctT7: ['% GP Retail Price Tier 7'],
  gpPctT8: ['% GP Retail Price Tier 8'],
  gpPctT9: ['% GP Retail Price Tier 9'],
  gpPctINTL: ['% GP Retail Price in INTL', '% GP Retail Price INTL'],
  gpAmtT6: ['$ GP Retail Price Tier 6'],
  gpAmtT7: ['$ GP Retail Price Tier 7'],
  gpAmtT8: ['$ GP Retail Price Tier 8'],
  gpAmtT9: ['$ GP Retail Price Tier 9'],
  gpAmtINTL: ['$ GP Retail Price INTL', '$ GP Retail Price in INTL'],
  gpUnitSubcat: ['GP/Unit by Subcategory'],
  gpPctSubcat: ['%GP by Subcategory', '% GP by Subcategory'],
  fcUpsd: ['FC UPSD of NEW SKU', 'UPSD per SKU by Subcategory'],
  skuCountSubcat: ['#SKU by subcategory', '# SKU by subcategory'],
  competitors: ['Competitors (Retail Price)', 'Competitors'],
  ssvNote: ['SSV Note Product'],
  newProductId: ['New Product ID', 'Ticket ID'],
  otherIncomeListingFee: ['Other Income + Listing Fee'],
  productDescription: ['Product description (Dưới 250 Kí tự)', 'Product description (Duoi 250 Ki tu)', 'Product description'],
  uomInformation: ['UOM Information', 'UOM info'],
  displayAreaPogConcept: ['Display area (POG CONCEPT) Bắt Buộc', 'Display area (POG CONCEPT) Bat Buoc', 'Display area (POG CONCEPT)', 'Display area'],
  preservationTemperature: ['Preservation temperature'],
  logisticsGroupSouth: ['Logistics Group South'],
  logisticsGroupNorth: ['Logistics Group North'],
  pattern: ['Pattern'],
  inventoryType: ['Inventory Type'],
  refPriceT6: ['REF PRODUCT Retail Price (+VAT) in Tier 6'],
  refPriceT7: ['REF PRODUCT Retail Price (+VAT) in Tier 7'],
  refPriceT8: ['REF PRODUCT Retail Price (+VAT) in Tier 8'],
  refPriceT9: ['REF PRODUCT Retail Price (+VAT) in Tier 9'],
  refPriceINTL: ['REF PRODUCT Retail Price (+VAT) in INTL', 'REF PRODUCT Retail Price (+VAT) in Intl'],
  refGpPctT6: ['REF PRODUCT % GP Retail Price Tier 6'],
  refGpPctT7: ['REF PRODUCT % GP Retail Price Tier 7'],
  refGpPctT8: ['REF PRODUCT % GP Retail Price Tier 8'],
  refGpPctT9: ['REF PRODUCT % GP Retail Price Tier 9'],
  refGpPctINTL: ['REF PRODUCT % GP Retail Price in INTL', 'REF PRODUCT % GP Retail Price INTL'],
  refGpAmtT6: ['REF PRODUCT $ GP Retail Price Tier 6'],
  refGpAmtT7: ['REF PRODUCT $ GP Retail Price Tier 7'],
  refGpAmtT8: ['REF PRODUCT $ GP Retail Price Tier 8'],
  refGpAmtT9: ['REF PRODUCT $ GP Retail Price Tier 9'],
  refGpAmtINTL: ['REF PRODUCT $ GP Retail Price INTL', 'REF PRODUCT $ GP Retail Price in INTL'],
  refVat: ['REF PRODUCT VAT', 'REF PRODUCT VAT input'],
  refPurchasePriceNoVat: ['REF PRODUCT Purchase Price -vat', 'REF PRODUCT Purchase Price - VAT'],
  targetStores: ['Target #Stores'],
  salesMonthlySubcat: ['Sales Monthly of Subcategory'],
  gpMonthlySubcat: ['GP Monthly of Subcategory'],
  pricingStrategy: ['Pricing strategy'],
  pogChooseReduce: ['POG choose to reduce'],
  pogChooseTypeReduce: ['POG choose Type to reduce'],
  totalSaleCut: ['Total Sales of Product to be CUT'],
  totalGpCut: ['Total GP of Product to be CUT'],
  fcSalePerMonth: ['FC Sale per month of new product'],
  fcGpAmountPerMonth: ['FC GP per month of new product'],
  saleImpactVsSubCategory: ['%Sale Impact vs Subcategory'],
  gpImpactVsSubCategory: ['%GP impact vs Subcategory'],
  urlImage: ['URL Image', 'Url Image'],
  imagePath: ['Image', 'Image Path']
};

function _directorMobileNormalizeToken_(value) {
  var s = _directorMobileSafeString_(value, '');
  try {
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (err) {
    // Keep original value when normalize is unavailable.
  }
  return s
    // Preserve semantic symbols so "% GP" and "$ GP" map to different columns.
    .replace(/[%％]/g, ' pct ')
    .replace(/[$＄]/g, ' usd ')
    .replace(/[#＃]/g, ' num ')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _directorMobileNormalizeHeaderExact_(value) {
  var s = _directorMobileSafeString_(value, '');
  try {
    s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (err) {
    // Keep original value when normalize is unavailable.
  }
  return s
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _directorMobileNormalizeKey_(value) {
  return _directorMobileSafeString_(value, '').trim().toLowerCase();
}

function _directorMobileHasValue_(value) {
  return !(value === null || value === undefined || _directorMobileSafeString_(value, '').trim() === '');
}

function _directorMobileToOutputValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return _cellToString(value);
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  return _directorMobileSafeString_(value, '');
}

function _directorMobileBuildHeaderMap_(headers) {
  var out = {};
  for (var i = 0; i < headers.length; i++) {
    var key = _directorMobileNormalizeToken_(headers[i]);
    if (!key) continue;
    // Prefer the rightmost column when headers are duplicated.
    out[key] = i;
  }
  return out;
}

function _directorMobileResolveIndex_(headers, headerMap, aliases) {
  var list = Array.isArray(aliases) ? aliases : [aliases];
  var headerList = Array.isArray(headers) ? headers : [];

  // First try exact header text matching (case/space insensitive, symbol-aware).
  for (var i = 0; i < list.length; i++) {
    var exactAlias = _directorMobileNormalizeHeaderExact_(list[i]);
    if (!exactAlias) continue;
    for (var h = headerList.length - 1; h >= 0; h--) {
      if (_directorMobileNormalizeHeaderExact_(headerList[h]) === exactAlias) return h;
    }
  }

  for (var i = 0; i < list.length; i++) {
    var key = _directorMobileNormalizeToken_(list[i]);
    if (key && headerMap[key] !== undefined) return headerMap[key];
  }
  for (var j = 0; j < list.length; j++) {
    var loose = _findHeaderIndexLoose_(headers, list[j]);
    if (loose >= 0) return loose;
  }
  return -1;
}

function _directorMobileBuildColumnMap_(headers) {
  var headerMap = _directorMobileBuildHeaderMap_(headers || []);
  var out = {};
  for (var key in DIRECTOR_MOBILE_FIELD_ALIASES) {
    if (!Object.prototype.hasOwnProperty.call(DIRECTOR_MOBILE_FIELD_ALIASES, key)) continue;
    out[key] = _directorMobileResolveIndex_(headers, headerMap, DIRECTOR_MOBILE_FIELD_ALIASES[key]);
  }
  return out;
}

function _directorMobileGetValue_(row, colMap, field) {
  var idx = colMap[field];
  if (idx === undefined || idx < 0 || idx >= row.length) return '';
  return _directorMobileToOutputValue_(row[idx]);
}

function _directorMobileCoalesce_(primary, fallback) {
  return _directorMobileHasValue_(primary) ? primary : fallback;
}

function _directorMobileNormalizeStatus_(value) {
  var s = _directorMobileNormalizeToken_(value);
  if (s === 'approved' || s === 'approve') return 'Approved';
  if (s === 'rejected' || s === 'reject') return 'Rejected';
  return 'Pending';
}

function _directorMobileReadUsers_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
  if (!sheet) throw new Error('Sheet "' + CONFIG.USER_MASTER_SHEET + '" not found');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sheet: sheet, headers: [], rows: [], idx: {} };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var headerMap = _directorMobileBuildHeaderMap_(headers);

  return {
    sheet: sheet,
    headers: headers,
    rows: rows,
    idx: {
      email: _directorMobileResolveIndex_(headers, headerMap, ['Email']),
      password: _directorMobileResolveIndex_(headers, headerMap, ['Password']),
      fullName: _directorMobileResolveIndex_(headers, headerMap, ['Full Name']),
      role: _directorMobileResolveIndex_(headers, headerMap, ['Role']),
      status: _directorMobileResolveIndex_(headers, headerMap, ['Status'])
    }
  };
}

function _directorMobileFindUserByLogin_(userPack, login) {
  if (!userPack || !Array.isArray(userPack.rows)) return null;
  var idx = userPack.idx || {};
  if (idx.email === undefined || idx.email < 0) return null;

  var loginNorm = _directorMobileNormalizeKey_(login);
  if (!loginNorm) return null;

  for (var i = 0; i < userPack.rows.length; i++) {
    var row = userPack.rows[i];
    var email = _directorMobileNormalizeKey_(row[idx.email]);
    if (!email) continue;
    var username = email.split('@')[0];
    if (loginNorm !== email && loginNorm !== username) continue;

    var role = idx.role >= 0 ? _directorMobileSafeString_(row[idx.role], '').trim().toUpperCase() : '';
    var status = idx.status >= 0 ? _directorMobileSafeString_(row[idx.status], '').trim().toUpperCase() : '';
    var fullName = idx.fullName >= 0 ? _directorMobileSafeString_(row[idx.fullName], '').trim() : '';
    var password = idx.password >= 0 ? _directorMobileSafeString_(row[idx.password], '') : '';

    return {
      rowNumber: i + 2,
      email: email,
      username: username,
      fullName: fullName,
      role: role,
      status: status,
      password: password
    };
  }
  return null;
}

function authenticateAccount(email, password) {
  try {
    var login = _directorMobileSafeString_(email, '').trim().toLowerCase();
    var pwd = _directorMobileSafeString_(password, '');
    if (!login || !pwd) {
      return { ok: false, message: 'Email and password are required.' };
    }

    var accessRes = _getUserAccessByEmail_(login, pwd, true);
    if (!accessRes.ok) {
      return { ok: false, message: accessRes.error && accessRes.error.message ? accessRes.error.message : 'Invalid email or password.' };
    }

    var selectedRes = _resolveSelectedAccess_(accessRes.data, 'DIR');
    if (!selectedRes.ok) {
      return { ok: false, message: 'Only accounts with DIR_View or DIR_Edit access can open this page.' };
    }

    var token = _createAuthToken_(selectedRes.data);
    return {
      ok: true,
      email: selectedRes.data.email,
      role: DIRECTOR_MOBILE_ROLE_DIRECTOR,
      authToken: token,
      currentViewCanEdit: selectedRes.data.currentViewCanEdit === true,
      allowedViews: selectedRes.data.allowedViews || []
    };
  } catch (err) {
    throw new Error('authenticateAccount error: ' + _directorSafeErrorText_(err, 'Unknown error'));
  }
}

function changeAccountPassword(email, currentPassword, newPassword) {
  try {
    var login = _directorMobileSafeString_(email, '').trim().toLowerCase();
    var oldPwd = _directorMobileSafeString_(currentPassword, '');
    var newPwd = _directorMobileSafeString_(newPassword, '').trim();

    if (!login || !oldPwd || !newPwd) {
      return { ok: false, message: 'Email, current password, and new password are required.' };
    }
    if (newPwd.length < 4) {
      return { ok: false, message: 'New password must be at least 4 characters.' };
    }

    var users = _directorMobileReadUsers_();
    var user = _directorMobileFindUserByLogin_(users, login);
    if (!user) return { ok: false, message: 'Account not found.' };
    if (user.role !== 'DIRECTOR') {
      return { ok: false, message: 'Only Director accounts can change password here.' };
    }
    if (user.status && user.status !== 'ACTIVE') {
      return { ok: false, message: 'Account is inactive.' };
    }
    if (oldPwd !== user.password) {
      return { ok: false, message: 'Current password is incorrect.' };
    }
    if (users.idx.password === undefined || users.idx.password < 0) {
      return { ok: false, message: 'Password column was not found.' };
    }

    users.sheet.getRange(user.rowNumber, users.idx.password + 1).setValue(newPwd);
    return { ok: true, message: 'Password updated successfully.' };
  } catch (err) {
    throw new Error('changeAccountPassword error: ' + _directorSafeErrorText_(err, 'Unknown error'));
  }
}

function _directorMobileBuildProduct_(row, rowNumber, colMap) {
  var productName = _directorMobileGetValue_(row, colMap, 'productName');
  var newProductId = _directorMobileGetValue_(row, colMap, 'newProductId');
  var noPresent = _directorMobileGetValue_(row, colMap, 'noPresent');
  var priceT6 = _directorMobileGetValue_(row, colMap, 'priceT6');
  var priceT7 = _directorMobileGetValue_(row, colMap, 'priceT7');
  var priceT8 = _directorMobileGetValue_(row, colMap, 'priceT8');
  var priceT9 = _directorMobileGetValue_(row, colMap, 'priceT9');
  var priceINTL = _directorMobileGetValue_(row, colMap, 'priceINTL');
  var gpPctT6 = _directorMobileGetValue_(row, colMap, 'gpPctT6');
  var gpPctT7 = _directorMobileGetValue_(row, colMap, 'gpPctT7');
  var gpPctT8 = _directorMobileGetValue_(row, colMap, 'gpPctT8');
  var gpPctT9 = _directorMobileGetValue_(row, colMap, 'gpPctT9');
  var gpPctINTL = _directorMobileGetValue_(row, colMap, 'gpPctINTL');
  var gpAmtT6 = _directorMobileGetValue_(row, colMap, 'gpAmtT6');
  var gpAmtT7 = _directorMobileGetValue_(row, colMap, 'gpAmtT7');
  var gpAmtT8 = _directorMobileGetValue_(row, colMap, 'gpAmtT8');
  var gpAmtT9 = _directorMobileGetValue_(row, colMap, 'gpAmtT9');
  var gpAmtINTL = _directorMobileGetValue_(row, colMap, 'gpAmtINTL');

  return {
    rowIndex: rowNumber,
    noPresent: _directorMobileCoalesce_(noPresent, _directorMobileSafeString_(rowNumber - CONFIG.DATA_START_ROW + 1, '')),
    presentDate: _directorMobileGetValue_(row, colMap, 'presentDate'),
    status: _directorMobileNormalizeStatus_(_directorMobileGetValue_(row, colMap, 'status')),
    note: _directorMobileGetValue_(row, colMap, 'note'),
    productName: _directorMobileCoalesce_(productName, _directorMobileCoalesce_(newProductId, '')),
    newProductType: _directorMobileGetValue_(row, colMap, 'newProductType'),
    whyDoWeListThisProduct_VN: _directorMobileGetValue_(row, colMap, 'whyListVN'),
    whyDoWeListThisProduct: _directorMobileGetValue_(row, colMap, 'whyListVN'),
    whyDoWeListThisProduct_EN: _directorMobileGetValue_(row, colMap, 'whyListEN'),
    premiumProducts: _directorMobileGetValue_(row, colMap, 'premiumProducts'),
    branded: _directorMobileGetValue_(row, colMap, 'branded'),
    picCategory: _directorMobileGetValue_(row, colMap, 'picCategory'),
    group: _directorMobileGetValue_(row, colMap, 'group'),
    category: _directorMobileGetValue_(row, colMap, 'category'),
    subCategory: _directorMobileGetValue_(row, colMap, 'subCategory'),
    primaryUsageSegment: _directorMobileGetValue_(row, colMap, 'primaryUsageSegment'),
    targetStoreTypes: _directorMobileGetValue_(row, colMap, 'targetStoreTypes'),
    targetCustomerStoreTypes: _directorMobileGetValue_(row, colMap, 'targetCustomerStoreTypes'),
    supplierName: _directorMobileGetValue_(row, colMap, 'supplierName'),
    countryOfOrigin: _directorMobileGetValue_(row, colMap, 'countryOfOrigin'),
    brandName: _directorMobileGetValue_(row, colMap, 'brandName'),
    purchasePrice: _directorMobileGetValue_(row, colMap, 'purchasePrice'),
    inboundVAT: _directorMobileGetValue_(row, colMap, 'inboundVAT'),
    samePriceProduct: _directorMobileGetValue_(row, colMap, 'samePriceProduct'),
    refProductId: _directorMobileGetValue_(row, colMap, 'refProductId'),
    refProductName: _directorMobileGetValue_(row, colMap, 'refProductName'),
    priceT6: priceT6,
    priceT7: priceT7,
    priceT8: priceT8,
    priceT9: priceT9,
    priceINTL: priceINTL,
    gpPctT6: gpPctT6,
    gpPctT7: gpPctT7,
    gpPctT8: gpPctT8,
    gpPctT9: gpPctT9,
    gpPctINTL: gpPctINTL,
    gpAmtT6: gpAmtT6,
    gpAmtT7: gpAmtT7,
    gpAmtT8: gpAmtT8,
    gpAmtT9: gpAmtT9,
    gpAmtINTL: gpAmtINTL,
    gpUnitSubcat: _directorMobileGetValue_(row, colMap, 'gpUnitSubcat'),
    gpPctSubcat: _directorMobileGetValue_(row, colMap, 'gpPctSubcat'),
    fcUpsd: _directorMobileGetValue_(row, colMap, 'fcUpsd'),
    skuCountSubcat: _directorMobileGetValue_(row, colMap, 'skuCountSubcat'),
    competitors: _directorMobileGetValue_(row, colMap, 'competitors'),
    ssvNote: _directorMobileGetValue_(row, colMap, 'ssvNote'),
    newProductId: _directorMobileCoalesce_(newProductId, _directorMobileCoalesce_(productName, '')),
    otherIncomeListingFee: _directorMobileGetValue_(row, colMap, 'otherIncomeListingFee'),
    productDescription: _directorMobileGetValue_(row, colMap, 'productDescription'),
    uomInformation: _directorMobileGetValue_(row, colMap, 'uomInformation'),
    displayAreaPogConcept: _directorMobileGetValue_(row, colMap, 'displayAreaPogConcept'),
    preservationTemperature: _directorMobileGetValue_(row, colMap, 'preservationTemperature'),
    logisticsGroupSouth: _directorMobileGetValue_(row, colMap, 'logisticsGroupSouth'),
    logisticsGroupNorth: _directorMobileGetValue_(row, colMap, 'logisticsGroupNorth'),
    pattern: _directorMobileGetValue_(row, colMap, 'pattern'),
    inventoryType: _directorMobileGetValue_(row, colMap, 'inventoryType'),
    refPriceT6: _directorMobileGetValue_(row, colMap, 'refPriceT6'),
    refPriceT7: _directorMobileGetValue_(row, colMap, 'refPriceT7'),
    refPriceT8: _directorMobileGetValue_(row, colMap, 'refPriceT8'),
    refPriceT9: _directorMobileGetValue_(row, colMap, 'refPriceT9'),
    refPriceINTL: _directorMobileGetValue_(row, colMap, 'refPriceINTL'),
    refGpPctT6: _directorMobileGetValue_(row, colMap, 'refGpPctT6'),
    refGpPctT7: _directorMobileGetValue_(row, colMap, 'refGpPctT7'),
    refGpPctT8: _directorMobileGetValue_(row, colMap, 'refGpPctT8'),
    refGpPctT9: _directorMobileGetValue_(row, colMap, 'refGpPctT9'),
    refGpPctINTL: _directorMobileGetValue_(row, colMap, 'refGpPctINTL'),
    refGpAmtT6: _directorMobileGetValue_(row, colMap, 'refGpAmtT6'),
    refGpAmtT7: _directorMobileGetValue_(row, colMap, 'refGpAmtT7'),
    refGpAmtT8: _directorMobileGetValue_(row, colMap, 'refGpAmtT8'),
    refGpAmtT9: _directorMobileGetValue_(row, colMap, 'refGpAmtT9'),
    refGpAmtINTL: _directorMobileGetValue_(row, colMap, 'refGpAmtINTL'),
    refVAT: _directorMobileGetValue_(row, colMap, 'refVat'),
    refPurchasePrice: _directorMobileGetValue_(row, colMap, 'refPurchasePriceNoVat'),
    targetStores: _directorMobileGetValue_(row, colMap, 'targetStores'),
    salesMonthlySubcat: _directorMobileGetValue_(row, colMap, 'salesMonthlySubcat'),
    gpMonthlySubcat: _directorMobileGetValue_(row, colMap, 'gpMonthlySubcat'),
    pricingStrategy: _directorMobileGetValue_(row, colMap, 'pricingStrategy'),
    pogChooseReduce: _directorMobileGetValue_(row, colMap, 'pogChooseReduce'),
    pogChooseTypeReduce: _directorMobileGetValue_(row, colMap, 'pogChooseTypeReduce'),
    totalSaleCut: _directorMobileGetValue_(row, colMap, 'totalSaleCut'),
    totalGpCut: _directorMobileGetValue_(row, colMap, 'totalGpCut'),
    fcSalePerMonth: _directorMobileGetValue_(row, colMap, 'fcSalePerMonth'),
    fcGpAmountPerMonth: _directorMobileGetValue_(row, colMap, 'fcGpAmountPerMonth'),
    saleImpactVsSubCategory: _directorMobileGetValue_(row, colMap, 'saleImpactVsSubCategory'),
    gpImpactVsSubCategory: _directorMobileGetValue_(row, colMap, 'gpImpactVsSubCategory'),
    urlImage: _directorMobileGetValue_(row, colMap, 'urlImage'),
    imagePath: _directorMobileGetValue_(row, colMap, 'imagePath')
  };
}

function _directorMobileIsEmptyRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (_directorMobileHasValue_(row[i])) return false;
  }
  return true;
}

function _directorMobileFindRowByKey_(rows, colMap, targetKey, productName) {
  var keyNorm = _directorMobileNormalizeKey_(targetKey);
  var nameNorm = _directorMobileNormalizeKey_(productName);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (_directorMobileIsEmptyRow_(row)) continue;

    var rowNewId = _directorMobileNormalizeKey_(_directorMobileGetValue_(row, colMap, 'newProductId'));
    var rowName = _directorMobileNormalizeKey_(_directorMobileGetValue_(row, colMap, 'productName'));

    if (keyNorm && (rowNewId === keyNorm || rowName === keyNorm)) return i;
    if (!keyNorm && nameNorm && (rowNewId === nameNorm || rowName === nameNorm)) return i;
  }
  return -1;
}

function getProductsAndResults(params) {
  try {
    _requireDirectorAccessOrThrow_(params, false);
    var sheet = _openDirectorSheet_();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var lastRow = sheet.getLastRow();

    if (headers.length === 0 || lastRow < CONFIG.DATA_START_ROW) {
      return { products: [], results: [] };
    }

    var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
    var rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();
    var colMap = _directorMobileBuildColumnMap_(headers);
    var products = [];
    var results = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (_directorMobileIsEmptyRow_(row)) continue;

      var rowNumber = CONFIG.DATA_START_ROW + i;
      var product = _directorMobileBuildProduct_(row, rowNumber, colMap);
      if (!_directorMobileHasValue_(product.productName) && !_directorMobileHasValue_(product.newProductId)) continue;

      products.push(product);

      var status = _directorMobileNormalizeStatus_(_directorMobileGetValue_(row, colMap, 'status'));
      var includeVote = status !== 'Pending';
      if (!includeVote) continue;

      var key = _directorMobileSafeString_(product.newProductId || product.productName || '', '').trim();
      var date = _directorMobileGetValue_(row, colMap, 'presentDate');
      results.push({
        date: date,
        newProductId: key,
        productName: product.productName,
        status: status,
        note: '',
        pic: DIRECTOR_MOBILE_ROLE_DIRECTOR,
        result: '',
        uniqueKey: ''
      });
    }

    return { products: products, results: results };
  } catch (err) {
    throw new Error('getProductsAndResults error: ' + _directorSafeErrorText_(err, 'Unknown error'));
  }
}

function saveResult(payload) {
  try {
    var data = payload || {};
    _requireDirectorAccessOrThrow_(data, true);
    var key = _directorMobileSafeString_(data.newProductId || data.productName || '', '').trim();
    if (!key) throw new Error('Missing product identifier');

    var status = _directorMobileNormalizeStatus_(data.status);
    if (status !== 'Approved' && status !== 'Rejected' && status !== 'Pending') {
      throw new Error('Status must be Approved, Rejected or Pending');
    }

    var sheet = _openDirectorSheet_();
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var lastRow = sheet.getLastRow();
    if (headers.length === 0 || lastRow < CONFIG.DATA_START_ROW) {
      throw new Error('Director sheet has no data rows');
    }

    var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
    var rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();
    var colMap = _directorMobileBuildColumnMap_(headers);

    var rowOffset = _directorMobileFindRowByKey_(rows, colMap, key, data.productName);
    if (rowOffset < 0) throw new Error('Product not found in Director sheet');

    if (colMap.status === undefined || colMap.status < 0) {
      throw new Error('Status column not found in Director sheet');
    }

    var targetRow = CONFIG.DATA_START_ROW + rowOffset;
    sheet.getRange(targetRow, colMap.status + 1).setValue(status);

    return { success: true, rowNumber: targetRow, status: status };
  } catch (err) {
    throw new Error('saveResult error: ' + _directorSafeErrorText_(err, 'Unknown error'));
  }
}
