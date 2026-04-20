// ========================================
// STAGE-AWARE PRODUCT CHAT SERVICE
// ========================================

var PRODUCT_CHAT_KEY_ALIASES = [
  'New Product ID',
  'New product ID',
  'Ticket ID',
  'Danh sach chinh'
];

var PRODUCT_CHAT_ALLOWED_SHEET_CONTEXTS = [
  'Category',
  'Manager',
  'Director',
  'Present List',
  'Final Information',
  'Setup Product ID'
];

var PRODUCT_CHAT_LOG_HEADERS = [
  'Timestamp',
  'Action',
  'New Product ID',
  'Message ID',
  'Message PIC',
  'Actor Email',
  'Sheet Context',
  'View Context',
  'Old Raw Entry',
  'New Raw Entry',
  'Storage Sheet',
  'Storage Row',
  'Request ID'
];

function getProductChatHistory(payload) {
  try {
    payload = payload || {};
    var accessRes = _resolveAccessForDataApi_(payload);
    if (!accessRes.ok) return accessRes;

    var productId = _normalizeProductChatProductId_(payload.newProductId || payload.productId || payload.key || payload.query);
    if (!productId) return _err('INVALID_PAYLOAD', 'Thiếu New Product ID để tải chat sản phẩm.');

    var target = _resolveProductChatStorageTarget_(payload, accessRes.data, productId);
    var readRes = _requireProductChatReadAccess_(payload, accessRes.data, target);
    if (!readRes.ok) return readRes;

    var parsed = _parseProductChatTranscript_(target.rawValue);
    var actorEmail = _getProductChatActorEmail_(accessRes.data);
    var sheetContext = _resolveProductChatSheetContext_(payload, accessRes.data);
    var viewContext = _resolveProductChatViewContext_(payload);
    var placement = target.placement || _buildProductChatPlacement_(payload, accessRes.data);

    return _ok(_buildProductChatHistoryPayload_(target, parsed, actorEmail, accessRes.data, sheetContext, viewContext, placement));
  } catch (err) {
    Logger.log('getProductChatHistory ERROR: ' + _safeErrorText_(err, 'Unknown error'));
    return _productChatExceptionToErr_(err, 'CHAT_READ_ERROR');
  }
}

function addProductChatMessage(payload) {
  var lock = null;
  var locked = false;
  try {
    payload = payload || {};
    var accessRes = _resolveAccessForDataApi_(payload);
    if (!accessRes.ok) return accessRes;

    var productId = _normalizeProductChatProductId_(payload.newProductId || payload.productId || payload.key || payload.query);
    if (!productId) return _err('INVALID_PAYLOAD', 'Thiếu New Product ID để gửi chat sản phẩm.');

    var content = _normalizeProductChatContent_(payload.content || payload.message || '');
    if (!content) return _err('INVALID_PAYLOAD', 'Vui lòng nhập nội dung tin nhắn.');

    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    locked = true;

    var target = _resolveProductChatStorageTarget_(payload, accessRes.data, productId);
    var writeRes = _requireProductChatWriteAccess_(payload, accessRes.data, target);
    if (!writeRes.ok) return writeRes;

    var parsed = _parseProductChatTranscript_(target.rawValue);
    var actorEmail = _getProductChatActorEmail_(accessRes.data);
    if (!actorEmail || actorEmail === 'unknown') return _err('FORBIDDEN', 'Không xác định được email người gửi chat.');

    var now = new Date();
    var uniqueDate = _getProductChatUniqueDate_(target.newProductId, now, parsed);
    var message = {
      type: 'message',
      pic: actorEmail,
      content: content,
      timestamp: _formatProductChatMessageTimestamp_(uniqueDate),
      id: _buildProductChatMessageId_(target.newProductId, uniqueDate)
    };
    var rawEntry = _serializeProductChatMessage_(message);

    parsed.entries.push({
      type: 'message',
      raw: rawEntry,
      message: message
    });

    var newRaw = _serializeProductChatEntries_(parsed.entries);
    var sheetContext = _resolveProductChatSheetContext_(payload, accessRes.data);
    var viewContext = _resolveProductChatViewContext_(payload);
    var placement = target.placement || _buildProductChatPlacement_(payload, accessRes.data);
    var requestId = Utilities.getUuid();

    var commitRes = _commitProductChatMutation_(target, target.rawValue, newRaw, {
      noteOld: '',
      noteNew: rawEntry,
      actorEmail: actorEmail,
      sheetContext: sheetContext,
      requestId: requestId,
      logEntry: _buildProductChatLogEntry_({
        action: 'ADD',
        newProductId: target.newProductId,
        messageId: message.id,
        messagePic: message.pic,
        actorEmail: actorEmail,
        sheetContext: sheetContext,
        viewContext: viewContext,
        oldRawEntry: '',
        newRawEntry: rawEntry,
        storageSheet: target.targetSheetName,
        storageRow: target.rowNumber,
        requestId: requestId
      })
    });
    if (!commitRes.ok) return commitRes;

    target.rawValue = newRaw;
    var updated = _parseProductChatTranscript_(newRaw);
    return _ok({
      message: _toProductChatMessageDto_(message, actorEmail),
      history: _buildProductChatHistoryPayload_(target, updated, actorEmail, accessRes.data, sheetContext, viewContext, placement),
      requestId: requestId
    });
  } catch (err) {
    Logger.log('addProductChatMessage ERROR: ' + _safeErrorText_(err, 'Unknown error'));
    return _productChatExceptionToErr_(err, 'CHAT_WRITE_ERROR');
  } finally {
    if (locked && lock) {
      try { lock.releaseLock(); } catch (releaseErr) {}
    }
  }
}

function editProductChatMessage(payload) {
  var lock = null;
  var locked = false;
  try {
    payload = payload || {};
    var accessRes = _resolveAccessForDataApi_(payload);
    if (!accessRes.ok) return accessRes;

    var productId = _normalizeProductChatProductId_(payload.newProductId || payload.productId || payload.key || payload.query);
    var messageId = String(payload.messageId || payload.id || '').trim();
    var content = _normalizeProductChatContent_(payload.content || payload.message || '');
    if (!productId || !messageId) return _err('INVALID_PAYLOAD', 'Thiếu New Product ID hoặc message ID.');
    if (!content) return _err('INVALID_PAYLOAD', 'Vui lòng nhập nội dung tin nhắn.');

    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    locked = true;

    var target = _resolveProductChatStorageTarget_(payload, accessRes.data, productId);
    var writeRes = _requireProductChatWriteAccess_(payload, accessRes.data, target);
    if (!writeRes.ok) return writeRes;

    var parsed = _parseProductChatTranscript_(target.rawValue);
    var found = _findProductChatMessageEntry_(parsed.entries, messageId);
    if (!found) return _err('MESSAGE_NOT_FOUND', 'Không tìm thấy tin nhắn chat.');

    var actorEmail = _getProductChatActorEmail_(accessRes.data);
    var ownerEmail = _normalizeLoginEmail_(found.entry.message.pic || '');
    if (!actorEmail || ownerEmail !== actorEmail) return _err('NOT_OWNER', 'Chỉ người tạo tin nhắn mới được sửa tin nhắn này.');

    if (found.entry.message.content === content) {
      var noChangeSheetContext = _resolveProductChatSheetContext_(payload, accessRes.data);
      var noChangeViewContext = _resolveProductChatViewContext_(payload);
      var noChangePlacement = target.placement || _buildProductChatPlacement_(payload, accessRes.data);
      return _ok({
        message: _toProductChatMessageDto_(found.entry.message, actorEmail),
        history: _buildProductChatHistoryPayload_(target, parsed, actorEmail, accessRes.data, noChangeSheetContext, noChangeViewContext, noChangePlacement),
        updated: false
      });
    }

    var oldRawEntry = found.entry.raw;
    var updatedMessage = {
      type: 'message',
      pic: found.entry.message.pic,
      content: content,
      timestamp: found.entry.message.timestamp,
      id: found.entry.message.id
    };
    var newRawEntry = _serializeProductChatMessage_(updatedMessage);
    parsed.entries[found.index] = {
      type: 'message',
      raw: newRawEntry,
      message: updatedMessage
    };

    var newRaw = _serializeProductChatEntries_(parsed.entries);
    var sheetContext = _resolveProductChatSheetContext_(payload, accessRes.data);
    var viewContext = _resolveProductChatViewContext_(payload);
    var placement = target.placement || _buildProductChatPlacement_(payload, accessRes.data);
    var requestId = Utilities.getUuid();

    var commitRes = _commitProductChatMutation_(target, target.rawValue, newRaw, {
      noteOld: oldRawEntry,
      noteNew: newRawEntry,
      actorEmail: actorEmail,
      sheetContext: sheetContext,
      requestId: requestId,
      logEntry: _buildProductChatLogEntry_({
        action: 'EDIT',
        newProductId: target.newProductId,
        messageId: messageId,
        messagePic: found.entry.message.pic,
        actorEmail: actorEmail,
        sheetContext: sheetContext,
        viewContext: viewContext,
        oldRawEntry: oldRawEntry,
        newRawEntry: newRawEntry,
        storageSheet: target.targetSheetName,
        storageRow: target.rowNumber,
        requestId: requestId
      })
    });
    if (!commitRes.ok) return commitRes;

    target.rawValue = newRaw;
    var updated = _parseProductChatTranscript_(newRaw);
    return _ok({
      message: _toProductChatMessageDto_(updatedMessage, actorEmail),
      history: _buildProductChatHistoryPayload_(target, updated, actorEmail, accessRes.data, sheetContext, viewContext, placement),
      requestId: requestId,
      updated: true
    });
  } catch (err) {
    Logger.log('editProductChatMessage ERROR: ' + _safeErrorText_(err, 'Unknown error'));
    return _productChatExceptionToErr_(err, 'CHAT_WRITE_ERROR');
  } finally {
    if (locked && lock) {
      try { lock.releaseLock(); } catch (releaseErr) {}
    }
  }
}

function deleteProductChatMessage(payload) {
  var lock = null;
  var locked = false;
  try {
    payload = payload || {};
    var accessRes = _resolveAccessForDataApi_(payload);
    if (!accessRes.ok) return accessRes;

    var productId = _normalizeProductChatProductId_(payload.newProductId || payload.productId || payload.key || payload.query);
    var messageId = String(payload.messageId || payload.id || '').trim();
    if (!productId || !messageId) return _err('INVALID_PAYLOAD', 'Thiếu New Product ID hoặc message ID.');

    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    locked = true;

    var target = _resolveProductChatStorageTarget_(payload, accessRes.data, productId);
    var writeRes = _requireProductChatWriteAccess_(payload, accessRes.data, target);
    if (!writeRes.ok) return writeRes;

    var parsed = _parseProductChatTranscript_(target.rawValue);
    var found = _findProductChatMessageEntry_(parsed.entries, messageId);
    if (!found) return _err('MESSAGE_NOT_FOUND', 'Không tìm thấy tin nhắn chat.');

    var actorEmail = _getProductChatActorEmail_(accessRes.data);
    var ownerEmail = _normalizeLoginEmail_(found.entry.message.pic || '');
    if (!actorEmail || ownerEmail !== actorEmail) return _err('NOT_OWNER', 'Chỉ người tạo tin nhắn mới được xóa tin nhắn này.');

    var oldRawEntry = found.entry.raw;
    parsed.entries.splice(found.index, 1);

    var newRaw = _serializeProductChatEntries_(parsed.entries);
    var sheetContext = _resolveProductChatSheetContext_(payload, accessRes.data);
    var viewContext = _resolveProductChatViewContext_(payload);
    var placement = target.placement || _buildProductChatPlacement_(payload, accessRes.data);
    var requestId = Utilities.getUuid();

    var commitRes = _commitProductChatMutation_(target, target.rawValue, newRaw, {
      noteOld: oldRawEntry,
      noteNew: '',
      actorEmail: actorEmail,
      sheetContext: sheetContext,
      requestId: requestId,
      logEntry: _buildProductChatLogEntry_({
        action: 'DELETE',
        newProductId: target.newProductId,
        messageId: messageId,
        messagePic: found.entry.message.pic,
        actorEmail: actorEmail,
        sheetContext: sheetContext,
        viewContext: viewContext,
        oldRawEntry: oldRawEntry,
        newRawEntry: '',
        storageSheet: target.targetSheetName,
        storageRow: target.rowNumber,
        requestId: requestId
      })
    });
    if (!commitRes.ok) return commitRes;

    target.rawValue = newRaw;
    var updated = _parseProductChatTranscript_(newRaw);
    return _ok({
      deletedMessageId: messageId,
      history: _buildProductChatHistoryPayload_(target, updated, actorEmail, accessRes.data, sheetContext, viewContext, placement),
      requestId: requestId
    });
  } catch (err) {
    Logger.log('deleteProductChatMessage ERROR: ' + _safeErrorText_(err, 'Unknown error'));
    return _productChatExceptionToErr_(err, 'CHAT_WRITE_ERROR');
  } finally {
    if (locked && lock) {
      try { lock.releaseLock(); } catch (releaseErr) {}
    }
  }
}

function _normalizeProductChatProductId_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}

function _normalizeProductChatLookupKey_(value) {
  return _normalizeProductChatProductId_(value).toLowerCase();
}

function _normalizeProductChatContent_(value) {
  return String(value == null ? '' : value)
    .replace(/\r\n?/g, '\n')
    .trim();
}

function _getProductChatTimezone_() {
  try {
    return Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  } catch (err) {
    return 'Asia/Ho_Chi_Minh';
  }
}

function _formatProductChatMessageTimestamp_(date) {
  return Utilities.formatDate(date, _getProductChatTimezone_(), 'dd-MM-yyyy HH:mm');
}

function _formatProductChatIdTimestamp_(date) {
  return Utilities.formatDate(date, _getProductChatTimezone_(), 'ddMMyyyyHHmmssSSS');
}

function _buildProductChatMessageId_(productId, date) {
  return String(productId || '').trim() + '-' + _formatProductChatIdTimestamp_(date);
}

function _getProductChatUniqueDate_(productId, seedDate, parsed) {
  var date = new Date(seedDate.getTime());
  var guard = 0;
  while (_productChatMessageIdExists_(parsed, _buildProductChatMessageId_(productId, date))) {
    date = new Date(date.getTime() + 1);
    guard++;
    if (guard > 1000) {
      _throwProductChatError_('CHAT_WRITE_ERROR', 'Cannot generate a unique chat message ID.');
    }
  }
  return date;
}

function _productChatMessageIdExists_(parsed, id) {
  var entries = parsed && Array.isArray(parsed.entries) ? parsed.entries : [];
  for (var i = 0; i < entries.length; i++) {
    if (entries[i] && entries[i].type === 'message' && entries[i].message && entries[i].message.id === id) return true;
  }
  return false;
}

function _encodeProductChatContent_(content) {
  return encodeURIComponent(String(content == null ? '' : content).replace(/\r\n?/g, '\n'));
}

function _decodeProductChatContent_(encoded) {
  var raw = String(encoded == null ? '' : encoded).replace(/\r\n?/g, '\n');
  if (!raw) return '';

  if (!/%[0-9A-Fa-f]{2}/.test(raw)) return raw;

  var firstPass = raw;
  try {
    firstPass = decodeURIComponent(raw);
  } catch (err) {
    return raw;
  }

  // Legacy data may have been encoded twice. Only decode a second pass when
  // the first pass still looks like UTF-8 byte escapes (%C3, %E1, ...).
  if (!/%[0-9A-Fa-f]{2}/.test(firstPass)) return firstPass;
  if (!/%(?:C[2-9A-F]|D[0-9A-F]|E[0-9A-F]|F[0-4])[0-9A-F]/i.test(firstPass)) return firstPass;

  try {
    return decodeURIComponent(firstPass);
  } catch (secondErr) {
    return firstPass;
  }
}

function _parseProductChatSerializedParts_(line) {
  var raw = String(line == null ? '' : line).replace(/\r\n?/g, '\n');
  var parts = raw.split('|');
  if (parts.length < 4) return null;

  var pic = _normalizeLoginEmail_(parts[0]);
  var timestamp = String(parts[parts.length - 2] || '').trim();
  var id = String(parts[parts.length - 1] || '').trim();
  if (!pic || !timestamp || !id) return null;
  if (!/^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$/.test(timestamp)) return null;

  return {
    raw: raw,
    pic: pic,
    encodedContent: parts.slice(1, parts.length - 2).join('|'),
    timestamp: timestamp,
    id: id
  };
}

function _buildProductChatLegacyDisplayFromRaw_(rawValue) {
  var text = String(rawValue == null ? '' : rawValue).replace(/\r\n?/g, '\n');
  if (!text) return '';

  var lines = text.split('\n');
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var parsed = _parseProductChatSerializedParts_(line);
    if (parsed) {
      out.push(parsed.pic + ' [' + parsed.timestamp + ']: ' + _decodeProductChatContent_(parsed.encodedContent));
    } else {
      out.push(_decodeProductChatContent_(line));
    }
  }
  return out.join('\n');
}

function _parseProductChatLine_(line) {
  var parsed = _parseProductChatSerializedParts_(line);
  if (!parsed) return null;

  var decoded = _decodeProductChatContent_(parsed.encodedContent);

  return {
    type: 'message',
    raw: parsed.raw,
    message: {
      type: 'message',
      pic: parsed.pic,
      content: decoded,
      timestamp: parsed.timestamp,
      id: parsed.id
    }
  };
}

function _parseProductChatTranscript_(value) {
  var text = String(value == null ? '' : value).replace(/\r\n?/g, '\n');
  var out = {
    entries: [],
    messages: [],
    legacy: []
  };
  if (!text) return out;

  var lines = text.split('\n');
  var legacyLines = [];
  function flushLegacy() {
    if (!legacyLines.length) return;
    var raw = legacyLines.join('\n');
    var entry = {
      type: 'legacy',
      raw: raw,
      content: _buildProductChatLegacyDisplayFromRaw_(raw)
    };
    out.entries.push(entry);
    out.legacy.push(entry);
    legacyLines = [];
  }

  for (var i = 0; i < lines.length; i++) {
    var parsedLine = _parseProductChatLine_(lines[i]);
    if (parsedLine) {
      flushLegacy();
      out.entries.push(parsedLine);
      out.messages.push(parsedLine.message);
    } else {
      legacyLines.push(lines[i]);
    }
  }
  flushLegacy();
  return out;
}

function _serializeProductChatMessage_(message) {
  return [
    _normalizeLoginEmail_(message.pic || ''),
    _encodeProductChatContent_(message.content || ''),
    String(message.timestamp || '').trim(),
    String(message.id || '').trim()
  ].join('|');
}

function _serializeProductChatEntries_(entries) {
  var lines = [];
  var list = Array.isArray(entries) ? entries : [];
  for (var i = 0; i < list.length; i++) {
    var entry = list[i];
    if (!entry) continue;
    if (entry.type === 'message' && entry.message) {
      lines.push(_serializeProductChatMessage_(entry.message));
    } else if (entry.raw !== null && entry.raw !== undefined) {
      lines.push(String(entry.raw));
    }
  }
  return lines.join('\n').replace(/\n+$/g, '');
}

function _findProductChatMessageEntry_(entries, messageId) {
  var id = String(messageId || '').trim();
  var list = Array.isArray(entries) ? entries : [];
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].type === 'message' && list[i].message && list[i].message.id === id) {
      return { index: i, entry: list[i] };
    }
  }
  return null;
}

function _getProductChatActorEmail_(access) {
  var email = access && access.email ? access.email : '';
  if (!email) {
    try { email = Session.getActiveUser().getEmail(); } catch (err) {}
  }
  return _normalizeLoginEmail_(email || 'unknown');
}

function _openProductChatDatabaseSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.SHEET_NAME_DATABASE || 'Database';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) _throwProductChatError_('DATABASE_SCHEMA_ERROR', 'Sheet "' + sheetName + '" was not found.');
  return sheet;
}

function _openProductChatCategorySheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.SHEET_NAME_CATEGORY || 'Category';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) _throwProductChatError_('CATEGORY_SCHEMA_ERROR', 'Sheet "' + sheetName + '" was not found.');
  return sheet;
}

function _findProductChatHeaderIndexes_(headers, aliases) {
  var out = [];
  var list = Array.isArray(aliases) ? aliases : [aliases];
  for (var i = 0; i < list.length; i++) {
    var idx = _findHeaderIndexLoose_(headers, list[i]);
    if (idx >= 0 && out.indexOf(idx) < 0) out.push(idx);
  }
  return out;
}

function _getProductChatTargetFromSheet_(sheet, productId, options) {
  var opts = options || {};
  var storageMode = opts.storageMode || 'DATABASE';
  var sheetLabel = opts.targetSheetName || sheet.getName();
  var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
  if (!headers.length) _throwProductChatError_(storageMode + '_SCHEMA_ERROR', sheetLabel + ' sheet has no headers.');

  var chatCol = _findHeaderIndexLoose_(headers, CONFIG.CHAT_FIELD_NAME || 'SSV Note Product');
  if (chatCol < 0) _throwProductChatError_(storageMode + '_SCHEMA_ERROR', sheetLabel + ' thiếu header "' + (CONFIG.CHAT_FIELD_NAME || 'SSV Note Product') + '".');

  var keyIndexes = _findProductChatHeaderIndexes_(headers, PRODUCT_CHAT_KEY_ALIASES);
  if (!keyIndexes.length) _throwProductChatError_(storageMode + '_SCHEMA_ERROR', sheetLabel + ' thiếu header mã sản phẩm.');

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    if (opts.allowMissing) return null;
    _throwProductChatError_(storageMode + '_ROW_NOT_FOUND', 'Không tìm thấy sản phẩm trong ' + sheetLabel + '.');
  }

  var targetKey = _normalizeProductChatLookupKey_(productId);
  var requestedRow = parseInt(opts.rowNumber, 10);
  if (requestedRow && requestedRow >= CONFIG.DATA_START_ROW && requestedRow <= lastRow) {
    var rowValues = sheet.getRange(requestedRow, 1, 1, headers.length).getValues()[0];
    var rowObjByNumber = _buildRowObjectFromArray_(headers, rowValues, requestedRow);
    if (_rowHasProductChatProductId_(rowObjByNumber, productId)) {
      return _buildProductChatStorageTarget_(sheet, headers, rowValues, requestedRow, rowObjByNumber, chatCol, productId, opts);
    }
    Logger.log('Product chat rowNumber hint mismatch. sheet=' + sheetLabel + ' row=' + requestedRow + ' product=' + productId + '. Falling back to product-id scan.');
  }

  var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  var rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    for (var k = 0; k < keyIndexes.length; k++) {
      var idx = keyIndexes[k];
      var candidate = _normalizeProductChatLookupKey_(row[idx]);
      if (!candidate || candidate !== targetKey) continue;

      var rowNumber = CONFIG.DATA_START_ROW + r;
      var rowObj = _buildRowObjectFromArray_(headers, row, rowNumber);
      return _buildProductChatStorageTarget_(sheet, headers, row, rowNumber, rowObj, chatCol, row[idx] || productId, opts);
    }
  }

  if (opts.allowMissing) return null;
  _throwProductChatError_(storageMode + '_ROW_NOT_FOUND', 'Không tìm thấy sản phẩm "' + productId + '" trong ' + sheetLabel + '.');
}

function _buildProductChatStorageTarget_(sheet, headers, rowValues, rowNumber, rowObj, chatCol, productId, options) {
  var opts = options || {};
  var canonicalId = _getProductChatCanonicalProductIdFromRow_(rowObj, productId);
  var storageMode = opts.storageMode || 'DATABASE';
  var targetSheetName = opts.targetSheetName || sheet.getName();
  return {
    sheet: sheet,
    headers: headers,
    rowNumber: rowNumber,
    rowObj: rowObj,
    chatColIndex: chatCol,
    rawValue: _cellToString(rowValues[chatCol]),
    newProductId: canonicalId,
    storageMode: storageMode,
    targetSheetName: targetSheetName,
    targetFieldName: CONFIG.CHAT_FIELD_NAME || 'SSV Note Product',
    businessSheet: opts.businessSheet || '',
    uiPlacement: opts.uiPlacement || '',
    flowState: opts.flowState || null,
    placement: opts.placement || null
  };
}

function _getProductChatDatabaseTarget_(productId, options) {
  var opts = Object.assign({}, options || {}, {
    storageMode: 'DATABASE',
    targetSheetName: CONFIG.SHEET_NAME_DATABASE || 'Database'
  });
  return _getProductChatTargetFromSheet_(_openProductChatDatabaseSheet_(), productId, opts);
}

function _getProductChatCategoryTarget_(productId, options) {
  var opts = Object.assign({}, options || {}, {
    storageMode: 'CATEGORY',
    targetSheetName: CONFIG.SHEET_NAME_CATEGORY || 'Category'
  });
  return _getProductChatTargetFromSheet_(_openProductChatCategorySheet_(), productId, opts);
}

function _shouldUseCategoryProductChatStorage_(access, flowState) {
  if (_getSelectedViewCode_(access) !== 'CAT') return false;
  if (!flowState || flowState.effectiveSheet !== 'Category') return false;
  return flowState.shouldUseWorkspaceChat === true;
}

function _resolveProductChatStorageTarget_(payload, access, productId) {
  var placement = _buildProductChatPlacement_(payload || {}, access || {});
  var flowState = placement.flowState || _getProductChatFlowState_(access || {}, payload || {});
  var requestedStorageMode = String(payload && (payload.storageMode || payload.storagePreference || '') || '').trim().toUpperCase();
  var sheetContext = _resolveProductChatSheetContext_(payload || {}, access || {});
  var uiPlacement = flowState && flowState.shouldUseWorkspaceChat ? 'WORKSPACE' : 'PROCESS';
  if (sheetContext === 'Director') uiPlacement = 'DIRECTOR';
  var baseOptions = {
    rowNumber: payload && payload.rowNumber,
    businessSheet: sheetContext,
    uiPlacement: uiPlacement,
    flowState: flowState,
    placement: placement
  };

  if (requestedStorageMode === 'DATABASE') {
    return _getProductChatDatabaseTarget_(productId, baseOptions);
  }

  if (_shouldUseCategoryProductChatStorage_(access, flowState)) {
    var categoryTarget = _getProductChatCategoryTarget_(productId, Object.assign({}, baseOptions, { allowMissing: true }));
    if (categoryTarget) return categoryTarget;
    Logger.log('Product chat CATEGORY storage requested but Category row was not found; falling back to Database for product=' + productId);
  }

  return _getProductChatDatabaseTarget_(productId, baseOptions);
}

function _getProductChatLooseRowValue_(rowObj, aliases) {
  if (!rowObj) return '';
  var keys = Object.keys(rowObj);
  var list = Array.isArray(aliases) ? aliases : [aliases];
  for (var a = 0; a < list.length; a++) {
    var wanted = _normalizeCategoryKey_(list[a]);
    if (!wanted) continue;
    for (var i = 0; i < keys.length; i++) {
      if (_normalizeCategoryKey_(keys[i]) !== wanted) continue;
      var value = _normalizeProductChatProductId_(rowObj[keys[i]]);
      if (value) return value;
    }
  }
  return '';
}

function _getProductChatCanonicalProductIdFromRow_(rowObj, fallback) {
  return _getProductChatLooseRowValue_(rowObj, ['New Product ID', 'New product ID'])
    || _getProductChatLooseRowValue_(rowObj, PRODUCT_CHAT_KEY_ALIASES)
    || _normalizeProductChatProductId_(fallback);
}

function _rowHasProductChatProductId_(rowObj, productId) {
  if (!rowObj) return false;
  var target = _normalizeProductChatLookupKey_(productId);
  for (var i = 0; i < PRODUCT_CHAT_KEY_ALIASES.length; i++) {
    var value = _getProductChatLooseRowValue_(rowObj, [PRODUCT_CHAT_KEY_ALIASES[i]]);
    if (_normalizeProductChatLookupKey_(value) === target) return true;
  }
  var processKeys = _collectProcessCandidateKeysFromRow_(rowObj);
  var list = processKeys && Array.isArray(processKeys.list) ? processKeys.list : [];
  for (var j = 0; j < list.length; j++) {
    if (_normalizeProductChatLookupKey_(list[j].raw) === target) return true;
  }
  return false;
}

function _hasProductChatCurrentRowAccess_(payload, access, productId) {
  var rowNumber = parseInt(payload && payload.rowNumber, 10);
  if (!rowNumber) return false;
  try {
    var sheet = _openSheet(access);
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    if (!headers.length || rowNumber < CONFIG.DATA_START_ROW || rowNumber > sheet.getLastRow()) return false;
    var rowSnapshot = _getRowSnapshotByNumber_(sheet, headers, rowNumber);
    return _rowHasProductChatProductId_(rowSnapshot, productId) && _hasRowAccess_(rowSnapshot, access);
  } catch (err) {
    return false;
  }
}

function _hasProductChatProcessAccess_(access, productId) {
  try {
    var viewCode = _getSelectedViewCode_(access);
    if (viewCode !== 'CAT' && viewCode !== 'MAN') return false;
    var processTable = _readProcessRecords_();
    if (processTable.warning) return false;

    var queryNorm = _normalizeProcessExactKey_(productId);
    if (!queryNorm) return false;
    var matches = [];
    for (var i = 0; i < processTable.records.length; i++) {
      if (processTable.records[i] && processTable.records[i].keyNorm === queryNorm) matches.push(processTable.records[i]);
    }
    if (!matches.length) return false;

    var allowedGroupInfo = _getAllowedProcessGroupsForAccess_(access, matches);
    var verifiedMatches = _filterProcessByAccess_(matches, access, {
      requireVerifiedAccess: true,
      enforceProductGroup: true,
      allowBlankGroupWhenVerified: true,
      allowVerifiedAccessBypassGroup: true,
      allowedGroupInfo: allowedGroupInfo
    });
    return verifiedMatches.length > 0;
  } catch (err) {
    return false;
  }
}

function _requireProductChatReadAccess_(payload, access, target) {
  if (!access) return _err('FORBIDDEN', 'Thiếu ngữ cảnh truy cập.');
  if (_isDirectorViewAccess_(access)) return _ok(true);
  if (_hasRowAccess_(target.rowObj, access)) return _ok(true);
  if (_hasProductChatCurrentRowAccess_(payload, access, target.newProductId)) return _ok(true);
  if (_hasProductChatProcessAccess_(access, target.newProductId)) return _ok(true);
  return _err('FORBIDDEN_SCOPE', 'Tai khoan hien tai khong duoc phep xem chat cua san pham nay.');
}

function _requireProductChatWriteAccess_(payload, access, target) {
  var readRes = _requireProductChatReadAccess_(payload, access, target);
  if (!readRes.ok) return readRes;
  if (!access || access.currentViewCanEdit !== true) {
    return _err('FORBIDDEN', 'Tai khoan hien tai khong co quyen gui chat o view da chon.');
  }
  return _ok(true);
}

function _resolveProductChatSheetContext_(payload, access) {
  var raw = String(
    (payload && (payload.sheetContext || payload.currentSheet || payload.currentSheetContext)) ||
    (access && access.selectedSheetName) ||
    ''
  ).trim();
  var normalized = _normalizeCategoryKey_(raw);

  var aliases = {
    'cat': 'Category',
    'category': 'Category',
    'category view': 'Category',
    'man': 'Manager',
    'manager': 'Manager',
    'manager view': 'Manager',
    'dir': 'Director',
    'director': 'Director',
    'director view': 'Director',
    'present list': 'Present List',
    'finallist': 'Final Information',
    'final information': 'Final Information',
    'setup product id': 'Setup Product ID'
  };

  var resolved = aliases[normalized] || '';
  if (!resolved && access) {
    var viewCode = _getSelectedViewCode_(access);
    if (viewCode === 'CAT') resolved = 'Category';
    if (viewCode === 'MAN') resolved = 'Manager';
    if (viewCode === 'DIR') resolved = 'Director';
  }
  if (!resolved || resolved === (CONFIG.SHEET_NAME_DATABASE || 'Database')) resolved = 'Category';
  if (PRODUCT_CHAT_ALLOWED_SHEET_CONTEXTS.indexOf(resolved) < 0) resolved = 'Category';
  return resolved;
}

function _resolveProductChatViewContext_(payload) {
  var raw = String(payload && payload.viewContext ? payload.viewContext : '').trim();
  if (!raw) return 'PRODUCT_CHAT';
  return raw.replace(/[\r\n\t]+/g, ' ').slice(0, 120);
}

function _getProductChatSelectedSheetContext_(access, payload) {
  return _resolveProductChatSheetContext_(payload || {}, access || {});
}

function _getProductChatStoredSheetFromSummary_(rowOrProductSummary, fallbackSheet) {
  var data = rowOrProductSummary || {};
  if (data.flowState && data.flowState.storedSheet) {
    return _normalizeProcessSheetName_(data.flowState.storedSheet) || String(data.flowState.storedSheet || '').trim();
  }
  var candidates = [
    data.storedSheet,
    data.currentStoredSheet,
    data.currentSheet,
    data.CurrentSheet,
    data['Current Sheet'],
    data.currentProductSheet,
    data.businessCurrentSheet,
    data.sheet,
    data.Sheet,
    data.sheetName,
    fallbackSheet
  ];
  for (var i = 0; i < candidates.length; i++) {
    var canonical = _normalizeProcessSheetName_(candidates[i]);
    if (canonical) return canonical;
    var raw = String(candidates[i] == null ? '' : candidates[i]).replace(/\s+/g, ' ').trim();
    if (raw) return raw;
  }
  return '';
}

function _getProductChatStatusFromSummary_(rowOrProductSummary) {
  var data = rowOrProductSummary || {};
  if (data.flowState && (data.flowState.statusRaw || data.flowState.statusNormalized)) {
    return data.flowState.statusRaw || data.flowState.statusNormalized;
  }
  return data.currentStatus || data.status || data.Status || '';
}

function _getProductChatFlowState_(access, rowOrProductSummary) {
  var data = rowOrProductSummary || {};
  var selectedSheet = _getProductChatSelectedSheetContext_(access || {}, data);
  if (data.flowState && data.flowState.storedSheet) {
    return _resolveBusinessFlowState_(data.flowState.storedSheet, data.flowState.statusRaw || data.flowState.statusNormalized || '', {
      selectedSheet: selectedSheet
    });
  }
  var storedSheet = _getProductChatStoredSheetFromSummary_(data, selectedSheet);
  var status = _getProductChatStatusFromSummary_(data);
  return _resolveBusinessFlowState_(storedSheet, status, { selectedSheet: selectedSheet });
}

function _getProductChatCurrentSheetFromSummary_(rowOrProductSummary, access) {
  return _getProductChatFlowState_(access || {}, rowOrProductSummary || {}).effectiveSheet || '';
}

function _isWorkspaceChatContext_(access, rowOrProductSummary) {
  return _getProductChatFlowState_(access || {}, rowOrProductSummary || {}).shouldUseWorkspaceChat === true;
}

function _shouldRenderWorkspaceChat_(access, rowOrProductSummary) {
  return _isWorkspaceChatContext_(access, rowOrProductSummary);
}

function _shouldRenderProcessChat_(access, rowOrProductSummary) {
  return _getProductChatFlowState_(access || {}, rowOrProductSummary || {}).shouldUseProcessChat === true;
}

function _buildProductChatPlacement_(payload, access) {
  var flowState = _getProductChatFlowState_(access || {}, payload || {});
  return {
    selectedSheet: flowState.selectedSheet,
    storedSheet: flowState.storedSheet,
    currentSheet: flowState.effectiveSheet,
    effectiveSheet: flowState.effectiveSheet,
    statusNormalized: flowState.statusNormalized,
    isRejected: flowState.isRejected,
    isProcessingForward: flowState.isProcessingForward,
    isTerminal: flowState.isTerminal,
    nextSheet: flowState.nextSheet,
    isWorkspaceChat: flowState.shouldUseWorkspaceChat,
    isProcessChat: flowState.shouldUseProcessChat,
    flowState: flowState
  };
}

function _appendProductChatHistoryNote_(range, beforeValue, afterValue, userEmail, sheetContext) {
  var oldPreview = _buildProductChatLegacyDisplayFromRaw_(beforeValue);
  var newPreview = _buildProductChatLegacyDisplayFromRaw_(afterValue);
  var oldValue = (!oldPreview)
    ? '(empty)'
    : oldPreview;
  var newValue = (!newPreview)
    ? '(empty)'
    : newPreview;
  var currentNote = range.getNote() || '';
  var context = sheetContext || 'Category';
  if (context === (CONFIG.SHEET_NAME_DATABASE || 'Database')) context = 'Category';

  var newEntry = _getFormattedTimestamp_() + ' | Sheet: ' + context + '\n'
    + 'Old: ' + oldValue + '\n'
    + 'New: ' + newValue + '\n'
    + 'By: ' + userEmail;

  var updatedNote = currentNote ? (newEntry + '\n\n' + currentNote) : newEntry;
  var entries = updatedNote.split('\n\n');
  var maxEntries = Math.max(1, parseInt(CONFIG.MAX_LOG_ENTRIES, 10) || 20);
  range.setNote(entries.slice(0, maxEntries).join('\n\n'));
}

function _ensureProductChatLogSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.CHAT_LOG_SHEET || 'Log_Chat';
  var sheet = ss.getSheetByName(sheetName);
  var headers = PRODUCT_CHAT_LOG_HEADERS;
  if (!sheet) sheet = ss.insertSheet(sheetName);

  var lastCol = Math.max(1, sheet.getLastColumn());
  var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var oldDatabaseRowIndex = existingHeaders.indexOf('Database Row');
  var storageSheetIndex = existingHeaders.indexOf('Storage Sheet');
  if (oldDatabaseRowIndex >= 0 && storageSheetIndex < 0) {
    sheet.insertColumnBefore(oldDatabaseRowIndex + 1);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, oldDatabaseRowIndex + 1, sheet.getLastRow() - 1, 1).setValue(CONFIG.SHEET_NAME_DATABASE || 'Database');
    }
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (current.join('') !== headers.join('')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    var widths = [170, 100, 180, 260, 220, 220, 140, 170, 360, 360, 140, 120, 260];
    for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);
  }
  return sheet;
}

function _buildProductChatLogEntry_(data) {
  return [
    new Date(),
    data.action || '',
    data.newProductId || '',
    data.messageId || '',
    data.messagePic || '',
    data.actorEmail || '',
    data.sheetContext || '',
    data.viewContext || '',
    data.oldRawEntry || '',
    data.newRawEntry || '',
    data.storageSheet || data.targetSheetName || data.databaseSheet || '',
    data.storageRow || data.databaseRow || '',
    data.requestId || ''
  ];
}

function _appendProductChatLogEntry_(entry) {
  if (!entry) return;
  var sheet = _ensureProductChatLogSheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, PRODUCT_CHAT_LOG_HEADERS.length).setValues([entry]);
}

function _commitProductChatMutation_(target, oldRaw, newRaw, options) {
  var opts = options || {};
  var cell = target.sheet.getRange(target.rowNumber, target.chatColIndex + 1);
  var oldNote = cell.getNote() || '';

  try {
    if (opts.logEntry) _ensureProductChatLogSheet_();
    cell.setValue(newRaw);
    _appendProductChatHistoryNote_(cell, opts.noteOld, opts.noteNew, opts.actorEmail, opts.sheetContext);
    if (opts.logEntry) _appendProductChatLogEntry_(opts.logEntry);
    return _ok(true);
  } catch (err) {
    try {
      cell.setValue(oldRaw);
      cell.setNote(oldNote);
    } catch (rollbackErr) {
      Logger.log('Product chat rollback failed: ' + _safeErrorText_(rollbackErr, 'Unknown error'));
    }
    var code = opts.logEntry ? 'LOG_CHAT_ERROR' : 'CHAT_WRITE_ERROR';
    return _err(code, _safeErrorText_(err, 'Không thể ghi chat sản phẩm.'), err && err.stack ? err.stack : '');
  }
}

function _toProductChatMessageDto_(message, actorEmail) {
  var owner = _normalizeLoginEmail_(message.pic || '');
  return {
    type: 'message',
    id: message.id,
    pic: owner,
    content: String(message.content || ''),
    timestamp: String(message.timestamp || ''),
    isOwn: owner && owner === actorEmail
  };
}

function _toProductChatLegacyDto_(entry, index) {
  var content = _buildProductChatLegacyDisplayFromRaw_(entry && entry.raw);
  if (!content) content = String(entry && entry.content || '');
  return {
    type: 'legacy',
    id: 'legacy-' + index,
    content: content,
    readOnly: true
  };
}

function _buildProductChatHistoryPayload_(target, parsed, actorEmail, access, sheetContext, viewContext, placement) {
  var items = [];
  var messages = [];
  var legacy = [];
  var entries = parsed && Array.isArray(parsed.entries) ? parsed.entries : [];
  var legacyIndex = 0;

  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    if (!entry) continue;
    if (entry.type === 'message' && entry.message) {
      var msg = _toProductChatMessageDto_(entry.message, actorEmail);
      items.push(msg);
      messages.push(msg);
    } else {
      var legacyDto = _toProductChatLegacyDto_(entry, legacyIndex++);
      items.push(legacyDto);
      legacy.push(legacyDto);
    }
  }

  return {
    newProductId: target.newProductId,
    databaseRow: target.storageMode === 'DATABASE' ? target.rowNumber : '',
    storageMode: target.storageMode || '',
    storageSheet: target.targetSheetName || '',
    storageRow: target.rowNumber || '',
    storageField: target.targetFieldName || (CONFIG.CHAT_FIELD_NAME || 'SSV Note Product'),
    businessSheet: target.businessSheet || sheetContext,
    uiPlacement: target.uiPlacement || '',
    sheetContext: sheetContext,
    viewContext: viewContext,
    placement: placement || target.placement || _buildProductChatPlacement_({ sheetContext: sheetContext }, access || {}),
    actorEmail: actorEmail,
    canWrite: !!(access && access.currentViewCanEdit === true),
    items: items,
    messages: messages,
    legacy: legacy
  };
}

function _throwProductChatError_(code, message, details) {
  var err = new Error(message || code || 'Product chat error');
  err.code = code || 'CHAT_ERROR';
  err.details = details || '';
  throw err;
}

function _productChatExceptionToErr_(err, fallbackCode) {
  return _err(
    (err && err.code) || fallbackCode || 'CHAT_ERROR',
    _safeErrorText_(err, 'Product chat error'),
    (err && (err.details || err.stack)) || ''
  );
}
