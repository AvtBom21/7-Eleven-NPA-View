// ========================================
// CONFIGURATION
// ========================================
var CONFIG = {
  SPREADSHEET_ID: '1W8oN5mha-GZ6wLIct-hsWQXSbFKreeICw2Y0-SFgQnY',
  SHEET_NAME_CATEGORY: 'Category',
  SHEET_NAME_MANAGER: 'Manager',
  SHEET_NAME_DIRECTOR: 'Director',
  SHEET_NAME_PROCESS: 'Process',
  DROPDOWN_SHEET: 'Dropdown_Master',
  USER_MASTER_SHEET: 'User_Master',
  AUDIT_LOG_SHEET: 'Audit_Log',
  SHEET_NAME_DATABASE: 'Database',
  CHAT_LOG_SHEET: 'Log_Chat',
  CHAT_FIELD_NAME: 'SSV Note Product',
  UOM_IMAGE_FOLDER_ID: '1WDIup4xssu-HLNGTOND3G9bIFIJfQF0g',
  MAX_LOG_ENTRIES: 20,
  PAGE_SIZE: 50,
  HEADER_ROW: 1,
  DATA_START_ROW: 2
};

var AUTH_CACHE_TTL_SECONDS = 21600;

var VIEW_DEFINITIONS = {
  DIR: {
    code: 'DIR',
    name: 'Director View',
    page: 'director',
    htmlFile: 'Director_View',
    title: 'Director - NPA Dashboard',
    dataMode: 'director',
    sheetName: CONFIG.SHEET_NAME_DIRECTOR,
    legacyRole: 'DIRECTOR',
    description: 'Director review workspace'
  },
  MAN: {
    code: 'MAN',
    name: 'Manager View',
    page: 'index',
    htmlFile: 'Index',
    title: 'Manager - NPA Dashboard',
    dataMode: 'manager',
    sheetName: CONFIG.SHEET_NAME_MANAGER,
    legacyRole: 'MANAGER',
    description: 'Manager workspace'
  },
  CAT: {
    code: 'CAT',
    name: 'Category View',
    page: 'index',
    htmlFile: 'Index',
    title: 'Category - NPA Dashboard',
    dataMode: 'category',
    sheetName: CONFIG.SHEET_NAME_CATEGORY,
    legacyRole: 'CATEGORY',
    description: 'Category workspace'
  }
};

function _getViewDefinition_(viewCode) {
  var code = String(viewCode || '').trim().toUpperCase();
  return VIEW_DEFINITIONS[code] || null;
}

function _legacyRoleToViewCode_(role) {
  var normalized = String(role || '').trim().toUpperCase();
  if (normalized === 'DIRECTOR') return 'DIR';
  if (normalized === 'MANAGER') return 'MAN';
  if (normalized === 'CATEGORY') return 'CAT';
  return '';
}

var ACCESS_TOKEN_TO_VIEW_CODE = {
  CAT_VIEW: 'CAT',
  CAT_EDIT: 'CAT',
  MAN_VIEW: 'MAN',
  MAN_EDIT: 'MAN',
  DIR_VIEW: 'DIR',
  DIR_EDIT: 'DIR'
};

var ACCESS_VIEW_CODE_TO_VIEW_TOKEN = {
  CAT: 'CAT_VIEW',
  MAN: 'MAN_VIEW',
  DIR: 'DIR_VIEW'
};

var ACCESS_VIEW_CODE_TO_EDIT_TOKEN = {
  CAT: 'CAT_EDIT',
  MAN: 'MAN_EDIT',
  DIR: 'DIR_EDIT'
};

function _getAccessTokenMeta_(token) {
  var normalized = _normalizeCanEditToken_(token);
  var viewCode = ACCESS_TOKEN_TO_VIEW_CODE[normalized] || '';
  if (!viewCode) return null;
  return {
    token: normalized,
    viewCode: viewCode,
    canView: true,
    canEdit: /_EDIT$/.test(normalized),
    viewToken: ACCESS_VIEW_CODE_TO_VIEW_TOKEN[viewCode] || '',
    editToken: ACCESS_VIEW_CODE_TO_EDIT_TOKEN[viewCode] || ''
  };
}

function _normalizeViewCode_(value) {
  var normalized = String(value == null ? '' : value)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '_');
  if (!normalized) return '';
  if (normalized === 'DIR' || normalized === 'DIRECTOR') return 'DIR';
  if (normalized === 'MAN' || normalized === 'MANAGER') return 'MAN';
  if (normalized === 'CAT' || normalized === 'CATEGORY') return 'CAT';
  return '';
}

function _toAccessTokenFromViewCode_(viewCode, canEdit) {
  var code = _normalizeViewCode_(viewCode);
  if (!code) return '';
  if (canEdit === true) return ACCESS_VIEW_CODE_TO_EDIT_TOKEN[code] || '';
  return ACCESS_VIEW_CODE_TO_VIEW_TOKEN[code] || '';
}

function _splitAccessTokenList_(raw) {
  if (raw === null || raw === undefined) return [];
  var text = String(raw)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\r\n?/g, '\n')
    .replace(/[;\n]/g, ',');
  return text
    .split(',')
    .map(function(item) {
      return String(item == null ? '' : item).replace(/\s+/g, ' ').trim();
    })
    .filter(function(item) {
      return !!item;
    });
}

function _normalizeCanEditToken_(value) {
  var token = String(value == null ? '' : value);
  if (token && typeof token.normalize === 'function') {
    try {
      token = token.normalize('NFKC');
    } catch (err) {}
  }
  return token
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\s-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase()
    .trim();
}

function _parseCanEditTokens_(raw) {
  var items = _splitAccessTokenList_(raw);
  var normalizedTokens = [];
  var invalidTokens = [];
  var seen = {};

  for (var i = 0; i < items.length; i++) {
    var rawToken = items[i];
    var normalized = _normalizeCanEditToken_(rawToken);
    if (!normalized) continue;
    if (!_getAccessTokenMeta_(normalized)) {
      invalidTokens.push({ raw: rawToken, normalized: normalized });
      continue;
    }
    if (seen[normalized]) continue;
    seen[normalized] = true;
    normalizedTokens.push(normalized);
  }

  return {
    raw: String(raw == null ? '' : raw),
    normalizedTokens: normalizedTokens,
    invalidTokens: invalidTokens
  };
}

function _buildAllowedViewsFromCanEditTokens_(canEditTokens) {
  var out = [];
  var byViewCode = {};
  var orderedViewCodes = [];

  var list = Array.isArray(canEditTokens) ? canEditTokens : [];
  for (var i = 0; i < list.length; i++) {
    var token = _normalizeCanEditToken_(list[i]);
    var tokenMeta = _getAccessTokenMeta_(token);
    if (!tokenMeta) continue;

    var viewCode = tokenMeta.viewCode;
    var viewEntry = byViewCode[viewCode];
    if (!viewEntry) {
      var viewDef = _getViewDefinition_(viewCode);
      if (!viewDef) continue;

      viewEntry = {
        code: viewDef.code,
        name: viewDef.name,
        page: viewDef.page,
        htmlFile: viewDef.htmlFile,
        title: viewDef.title,
        dataMode: viewDef.dataMode,
        sheetName: viewDef.sheetName,
        legacyRole: viewDef.legacyRole,
        description: viewDef.description,
        accessToken: tokenMeta.viewToken || token,
        editRule: '',
        editRuleNormalized: '',
        canView: true,
        canEdit: false
      };

      byViewCode[viewCode] = viewEntry;
      orderedViewCodes.push(viewCode);
    }

    if (tokenMeta.canEdit) {
      viewEntry.canEdit = true;
      viewEntry.accessToken = tokenMeta.editToken || token;
      viewEntry.editRule = tokenMeta.editToken || token;
      viewEntry.editRuleNormalized = tokenMeta.editToken || token;
    }
  }

  for (var j = 0; j < orderedViewCodes.length; j++) {
    var code = orderedViewCodes[j];
    if (byViewCode[code]) out.push(byViewCode[code]);
  }

  return out;
}

function _collectCanEditTokensFromAllowedViews_(allowedViews, selectedViewCode, selectedCanEdit) {
  var out = [];
  var seen = {};
  var selected = _normalizeViewCode_(selectedViewCode);
  var list = Array.isArray(allowedViews) ? allowedViews : [];

  for (var i = 0; i < list.length; i++) {
    var view = list[i] || {};
    var viewCode = _normalizeViewCode_(view.code || view.viewCode || '');
    if (!viewCode) continue;

    var canEdit = view.canEdit === true;
    if (!canEdit && view.editRule) {
      var editMeta = _getAccessTokenMeta_(view.editRule);
      canEdit = !!(editMeta && editMeta.canEdit);
    }
    if (!canEdit && selectedCanEdit === true && viewCode === selected && view.canEdit !== false) canEdit = true;

    var token = _toAccessTokenFromViewCode_(viewCode, canEdit);
    if (!token || seen[token]) continue;
    seen[token] = true;
    out.push(token);
  }

  if (!out.length && selected) {
    var selectedToken = _toAccessTokenFromViewCode_(selected, selectedCanEdit === true);
    if (selectedToken) out.push(selectedToken);
  }

  return out;
}

function _resolveCanEditRawFromAccessData_(source) {
  var data = source || {};
  if (typeof data.canEditRaw === 'string' && data.canEditRaw.trim()) return data.canEditRaw;
  if (typeof data.canEditTokensRaw === 'string' && data.canEditTokensRaw.trim()) return data.canEditTokensRaw;
  if (Array.isArray(data.canEditTokensNormalized) && data.canEditTokensNormalized.length) {
    return data.canEditTokensNormalized.join(', ');
  }
  if (typeof data.editRulesRaw === 'string' && data.editRulesRaw.trim()) return data.editRulesRaw;

  var fromViews = _collectCanEditTokensFromAllowedViews_(
    data.allowedViews,
    data.selectedView || data.viewCode || '',
    data.currentViewCanEdit === true || data.canEdit === true
  );
  if (fromViews.length) return fromViews.join(', ');

  var fallbackView = _normalizeViewCode_(data.selectedView || data.viewCode || '')
    || _legacyRoleToViewCode_(data.role || data.baseRole || '');
  if (fallbackView) {
    var fallbackToken = _toAccessTokenFromViewCode_(fallbackView, data.currentViewCanEdit === true || data.canEdit === true);
    if (fallbackToken) return fallbackToken;
  }

  return '';
}

function _clonePlainObject_(value) {
  return JSON.parse(JSON.stringify(value || {}));
}

function _findAllowedView_(access, viewCode) {
  var code = _normalizeViewCode_(viewCode);
  var allowedViews = access && Array.isArray(access.allowedViews) ? access.allowedViews : [];
  for (var i = 0; i < allowedViews.length; i++) {
    if (allowedViews[i] && allowedViews[i].code === code) return allowedViews[i];
  }
  return null;
}

function _applySelectedViewToAccess_(access, viewCode) {
  var selected = _findAllowedView_(access, viewCode);
  if (!selected) return null;

  var out = _clonePlainObject_(access);
  out.selectedView = selected.code;
  out.selectedViewName = selected.name;
  out.selectedPage = selected.page;
  out.selectedDataMode = selected.dataMode;
  out.selectedSheetName = selected.sheetName;
  out.selectedEditRule = selected.editRule || '';
  out.selectedEditRuleNormalized = selected.editRuleNormalized || '';
  out.currentViewCanEdit = selected.canEdit === true;
  out.canEdit = out.currentViewCanEdit;
  out.role = selected.legacyRole;
  out.landingPage = selected.page === 'director' ? 'DIRECTOR_VIEW' : 'INDEX';
  out.selectionRequired = false;
  out.scope = selected.code === 'DIR'
    ? ['ALL']
    : (Array.isArray(out.categoryScope) ? out.categoryScope.slice() : []);
  return out;
}

function _resolveDefaultView_(allowedViews) {
  return allowedViews && allowedViews.length ? allowedViews[0].code : '';
}

function _getSelectedViewCode_(access) {
  if (!access) return '';
  var direct = _normalizeViewCode_(access.selectedView || access.viewCode || '');
  if (direct) return direct;
  var allowedViews = access && Array.isArray(access.allowedViews) ? access.allowedViews : [];
  if (allowedViews.length === 1 && allowedViews[0]) {
    return _normalizeViewCode_(allowedViews[0].code || '');
  }
  return '';
}

function _hasSelectedView_(access) {
  return !!_getSelectedViewCode_(access);
}

function _isDirectorViewAccess_(access) {
  return _getSelectedViewCode_(access) === 'DIR';
}

function _buildAccessModel_(options) {
  var opts = options || {};
  var rawRole = String(opts.role || '').trim().toUpperCase();
  var canEditRaw = String(opts.canEditRaw == null ? '' : opts.canEditRaw);
  var parsedCanEdit = _parseCanEditTokens_(canEditRaw);
  if (parsedCanEdit.invalidTokens.length) {
    Logger.log('ACCESS_CAN_EDIT_INVALID tokens=' + JSON.stringify(parsedCanEdit.invalidTokens) + ' email=' + _normalizeLoginEmail_(opts.email));
  }

  var allowedViews = _buildAllowedViewsFromCanEditTokens_(parsedCanEdit.normalizedTokens);
  var categoryScope = Array.isArray(opts.categoryScope) ? opts.categoryScope.slice() : [];

  var access = {
    email: _normalizeLoginEmail_(opts.email),
    fullName: String(opts.fullName || '').trim(),
    role: rawRole,
    baseRole: rawRole,
    status: String(opts.status || '').trim().toUpperCase(),
    scope: categoryScope.slice(),
    categoryScope: categoryScope.slice(),
    canEditRaw: parsedCanEdit.raw,
    canEditTokensNormalized: parsedCanEdit.normalizedTokens.slice(),
    invalidCanEditTokens: parsedCanEdit.invalidTokens.slice(),
    allowedViews: allowedViews,
    defaultView: _resolveDefaultView_(allowedViews),
    selectedView: '',
    selectedViewName: '',
    selectedPage: '',
    selectedDataMode: '',
    selectedSheetName: '',
    selectedEditRule: '',
    selectedEditRuleNormalized: '',
    currentViewCanEdit: false,
    canEdit: false,
    landingPage: '',
    hasMultipleViews: allowedViews.length > 1,
    selectionRequired: allowedViews.length > 1
  };

  if (allowedViews.length === 1) {
    access = _applySelectedViewToAccess_(access, allowedViews[0].code) || access;
  } else if (opts.selectedView) {
    access = _applySelectedViewToAccess_(access, opts.selectedView) || access;
  }

  return access;
}

function _normalizeLegacyAccessData_(data) {
  var source = data || {};
  var selectedView = _normalizeViewCode_(source.selectedView || source.viewCode || '');
  var categoryScope = Array.isArray(source.categoryScope)
    ? source.categoryScope.slice()
    : (Array.isArray(source.scope) ? source.scope.slice() : []);
  var access = _buildAccessModel_({
    email: source.email,
    fullName: source.fullName,
    role: source.baseRole || source.role,
    status: source.status,
    canEditRaw: _resolveCanEditRawFromAccessData_(source),
    categoryScope: categoryScope,
    selectedView: source.selectedView || selectedView
  });

  if (!access.selectedView && selectedView) {
    access = _applySelectedViewToAccess_(access, selectedView) || access;
  }
  if (source.selectedPage && access.selectedView && !access.selectedPage) {
    access.selectedPage = String(source.selectedPage || '').trim();
  }
  return access;
}

function _normalizeAccessData_(data) {
  if (!data || typeof data !== 'object') return null;
  return _normalizeLegacyAccessData_(data);
}

function _storeAccessByToken_(token, accessData) {
  CacheService.getScriptCache().put('AUTH_' + token, JSON.stringify(accessData || {}), AUTH_CACHE_TTL_SECONDS);
}

function _removeAccessByToken_(token) {
  var tk = String(token || '').trim();
  if (!tk) return false;
  CacheService.getScriptCache().remove('AUTH_' + tk);
  return true;
}

function _buildAppUrl_(page, token, viewCode) {
  var parts = ['page=' + encodeURIComponent(page || 'login')];
  if (token) parts.push('token=' + encodeURIComponent(token));
  if (viewCode) parts.push('view=' + encodeURIComponent(viewCode));
  var qs = parts.join('&');
  var baseUrl = '';
  try {
    baseUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    baseUrl = '';
  }
  return baseUrl ? (baseUrl + '?' + qs) : ('?' + qs);
}

function _buildCleanLoginUrl_() {
  return _buildAppUrl_('login', '', '');
}

function _renderClientRedirect_(url) {
  var rawUrl = String(url || '?page=login');
  var htmlUrl = rawUrl
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  var jsUrl = rawUrl
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0; url=' + htmlUrl + '"></head>'
    + '<body><script>window.top.location.replace("' + jsUrl + '");</script></body></html>'
  );
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function _buildLoginBootstrap_(access, token, message) {
  var allowedViews = access && Array.isArray(access.allowedViews) ? access.allowedViews : [];
  return {
    mode: allowedViews.length > 1 ? 'selector' : 'login',
    message: String(message || ''),
    email: access && access.email ? access.email : '',
    fullName: access && access.fullName ? access.fullName : '',
    token: String(token || ''),
    defaultView: access && access.defaultView ? access.defaultView : '',
    selectedView: access && access.selectedView ? access.selectedView : '',
    availableViews: allowedViews.map(function(view) {
      return {
        code: view.code,
        name: view.name,
        page: view.page,
        dataMode: view.dataMode,
        sheetName: view.sheetName,
        editRule: view.editRule,
        canEdit: view.canEdit === true,
        description: view.description
      };
    })
  };
}

var PIC_EDITABLE_FIELDS = [
  'Present Date',
  'Status',
  'Product Group',
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

var PRODUCT_GROUP_OPTIONS = [
  'Non Alcoholic',
  'Alcoholic',
  'Tobacco',
  'Packaged Snacks',
  'Health and Beauty',
  'Instant Food',
  'Condiments',
  'Groceries',
  'Confectionery',
  'Toys',
  'Stationery',
  'Household',
  'Home Care',
  'Clothing and Accessories'
];

function _isAllowedProductGroup_(value) {
  var v = String(value || '').trim();
  return !!v && PRODUCT_GROUP_OPTIONS.indexOf(v) >= 0;
}

function _stripVietnameseToneForPattern_(value) {
  var txt = String(value == null ? '' : value).toLowerCase().trim();
  try {
    if (txt && txt.normalize) txt = txt.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (err) {
    // Apps Script may run older V8 contexts; keep the original text if normalize fails.
  }
  return txt.replace(/đ/g, 'd');
}

function _normalizePatternToken_(value) {
  var raw = _stripVietnameseToneForPattern_(value)
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!raw) return '';
  var hasSouth = raw.indexOf('mien nam') >= 0 || raw.indexOf('south') >= 0 || /(^| )nam( |$)/.test(raw);
  var hasNorth = raw.indexOf('mien bac') >= 0 || raw.indexOf('north') >= 0 || /(^| )bac( |$)/.test(raw);
  if (
    raw.indexOf('hai mien') >= 0 ||
    raw.indexOf('ca hai') >= 0 ||
    raw.indexOf('ca 2') >= 0 ||
    raw.indexOf('2 mien') >= 0 ||
    raw.indexOf('both') >= 0 ||
    (hasSouth && hasNorth)
  ) return 'both';
  if (hasSouth) return 'south';
  if (hasNorth) return 'north';
  return '';
}

function _normalizePatternValue_(value) {
  var token = _normalizePatternToken_(value);
  if (token === 'south') return 'South';
  if (token === 'north') return 'North';
  if (token === 'both') return 'South, North';
  return String(value == null ? '' : value).trim();
}

function _isAllowedPatternValue_(value) {
  var v = _normalizePatternValue_(value);
  return !v || v === 'South' || v === 'North' || v === 'South, North';
}

function _normalizeProductWriteValue_(field, value) {
  if (field === 'Pattern') return _normalizePatternValue_(value);
  return value;
}

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

function _safeJsonForInlineScript_(value) {
  var json = '{}';
  try {
    json = JSON.stringify(value);
    if (json === undefined) json = 'null';
  } catch (err) {
    json = '{}';
  }
  return String(json)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
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
      ? String(e.parameter.page || '').toLowerCase() : 'index';
    var requestedView = e && e.parameter ? _normalizeViewCode_(e.parameter.view || '') : '';
    var token = e && e.parameter ? String(e.parameter.token || '').trim() : '';

    if (page !== 'login' && page !== 'index' && page !== 'director') {
      page = 'index';
    }

    if (page === 'login') {
      if (token) {
        var loginAccessRes = _getAccessByToken_(token);
        if (loginAccessRes.ok) {
          var loginAccess = loginAccessRes.data;
          if (loginAccess.allowedViews && loginAccess.allowedViews.length > 1) {
            return _renderLoginPage('', loginAccess.email || _resolveRequestEmail_(e) || '', _buildLoginBootstrap_(loginAccess, token, 'Choose a view to continue.'));
          }
          if (_hasSelectedView_(loginAccess)) {
            return _renderClientRedirect_(_buildAppUrl_(loginAccess.selectedPage, token, loginAccess.selectedView));
          }
        }
      }
      return _renderLoginPage('', _resolveRequestEmail_(e) || '', null);
    }

    if (!token) {
      return _renderLoginPage('Please sign in to continue.', _resolveRequestEmail_(e) || '', null);
    }

    var accessRes = _getAccessByToken_(token);
    if (!accessRes.ok) {
      return _renderLoginPage(accessRes.error.message, _resolveRequestEmail_(e) || '', null);
    }

    var access = accessRes.data;
    if (!_hasSelectedView_(access) && access.allowedViews && access.allowedViews.length === 1) {
      var autoSelected = _applySelectedViewToAccess_(access, access.allowedViews[0].code);
      if (autoSelected) {
        access = autoSelected;
        _storeAccessByToken_(token, access);
      }
    }

    if (!_hasSelectedView_(access)) {
      return _renderLoginPage('Choose a view to continue.', access.email || _resolveRequestEmail_(e) || '', _buildLoginBootstrap_(access, token, 'Choose a view to continue.'));
    }

    if ((requestedView && requestedView !== access.selectedView) || page !== access.selectedPage) {
      return _renderClientRedirect_(_buildAppUrl_(access.selectedPage, token, access.selectedView));
    }

    return _renderAppPage_(access, token);
  } catch (err) {
    var msg = _safeErrorText_(err, 'Unknown error');
    return HtmlService.createHtmlOutput(
      '<h2 style="font-family:sans-serif;color:red">Error: ' + msg + '</h2>'
    );
  }
}

function _renderAppPage_(access, token) {
  var viewDef = _getViewDefinition_(_getSelectedViewCode_(access)) || _getViewDefinition_('CAT');
  var tpl = HtmlService.createTemplateFromFile(viewDef.htmlFile);
  tpl.appUserJson = _safeJsonForInlineScript_({
    email: access.email,
    fullName: access.fullName,
    role: access.role,
    baseRole: access.baseRole,
    status: access.status,
    scope: Array.isArray(access.scope) ? access.scope : [],
    categoryScope: Array.isArray(access.categoryScope) ? access.categoryScope : [],
    canEdit: access.currentViewCanEdit === true,
    currentViewCanEdit: access.currentViewCanEdit === true,
    canEditRaw: access.canEditRaw || '',
    canEditTokensNormalized: Array.isArray(access.canEditTokensNormalized) ? access.canEditTokensNormalized : [],
    invalidCanEditTokens: Array.isArray(access.invalidCanEditTokens) ? access.invalidCanEditTokens : [],
    allowedViews: access.allowedViews || [],
    defaultView: access.defaultView || '',
    selectedView: access.selectedView || '',
    selectedViewName: access.selectedViewName || '',
    selectedPage: access.selectedPage || '',
    selectedDataMode: access.selectedDataMode || '',
    selectedSheetName: access.selectedSheetName || '',
    selectedEditRule: access.selectedEditRule || '',
    authToken: token
  });
  var html = tpl.evaluate();
  html.setTitle(viewDef.title);
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// ========================================
// HELPER: Safe sheet access
// ========================================
function _openSheet(access) {
  var viewCode = _getSelectedViewCode_(access);
  var sheetName = CONFIG.SHEET_NAME_CATEGORY;
  if (viewCode === 'DIR') sheetName = CONFIG.SHEET_NAME_DIRECTOR;
  if (viewCode === 'MAN') sheetName = CONFIG.SHEET_NAME_MANAGER;
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
  if (!access) return false;
  if (_isDirectorViewAccess_(access)) return true;
  return _applyAccessScopeToRows_([rowObj], access).length > 0;
}

function _getScopeDisplay_(access) {
  if (!access) return '';
  if (_isDirectorViewAccess_(access)) return 'ALL';
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
  return _isDirectorViewAccess_(access) ? 'DIRECTOR_UPDATE' : 'PIC_UPDATE';
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
  if (!Array.isArray(rows)) return [];
  if (!access) return [];
  if (_isDirectorViewAccess_(access)) return rows;
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
  var token = '';
  if (typeof params === 'string') {
    token = String(params || '').trim();
  } else if (params && typeof params === 'object') {
    token = params.token ? String(params.token).trim() : '';
  } else {
    return _err('MISSING_TOKEN', 'Vui long dang nhap de tiep tuc.');
  }

  if (!token) {
    return _err('MISSING_TOKEN', 'Vui long dang nhap de tiep tuc.');
  }
  var accessRes = _getAccessByToken_(token);
  if (!accessRes.ok) return accessRes;

  var access = accessRes.data || {};
  var hasManyViews = Array.isArray(access.allowedViews) && access.allowedViews.length > 1;
  if (hasManyViews && !_hasSelectedView_(access)) {
    return _err('VIEW_SELECTION_REQUIRED', 'Tai khoan co nhieu view. Vui long chon workspace truoc khi tiep tuc.');
  }

  return accessRes;
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
        meta: {
          totalRows: 0,
          currentPage: 1,
          totalPages: 1,
          pageSize: CONFIG.PAGE_SIZE,
          sheetName: sheet.getName(),
          selectedView: accessRes.data && accessRes.data.selectedView ? accessRes.data.selectedView : '',
          selectedDataMode: accessRes.data && accessRes.data.selectedDataMode ? accessRes.data.selectedDataMode : '',
          currentViewCanEdit: accessRes.data && accessRes.data.currentViewCanEdit === true
        }
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
    allRows = _applyConversationMetaToRows_(allRows, accessRes.data);

    var total = allRows.length;

    Logger.log('getInitData SUCCESS: total=' + total + ' loaded=' + allRows.length);

    return _ok({
      headers: headers,
      rows: allRows,
      editableFields: editableFields,
      meta: {
        totalRows: total,
        currentPage: 1,
        totalPages: 1,
        pageSize: total || CONFIG.PAGE_SIZE,
        sheetName: sheet.getName(),
        selectedView: accessRes.data && accessRes.data.selectedView ? accessRes.data.selectedView : '',
        selectedDataMode: accessRes.data && accessRes.data.selectedDataMode ? accessRes.data.selectedDataMode : '',
        currentViewCanEdit: accessRes.data && accessRes.data.currentViewCanEdit === true
      }
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
    allRows = _applyConversationMetaToRows_(allRows, accessRes.data);

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

function _isProtectedChatField_(fieldName) {
  var target = String(CONFIG.CHAT_FIELD_NAME || 'SSV Note Product').trim().replace(/\s+/g, ' ').toLowerCase();
  var candidate = String(fieldName || '').trim().replace(/\s+/g, ' ').toLowerCase();
  return !!target && candidate === target;
}

function _hasProductChatConversationValue_(value) {
  var text = String(value == null ? '' : value).replace(/\r\n?/g, '\n');
  if (!text.replace(/\s+/g, '').length) return false;

  try {
    var parsed = _parseProductChatTranscript_(text);
    if (parsed && Array.isArray(parsed.messages) && parsed.messages.length) return true;
    var legacy = parsed && Array.isArray(parsed.legacy) ? parsed.legacy : [];
    for (var i = 0; i < legacy.length; i++) {
      var content = String((legacy[i] && (legacy[i].content || legacy[i].raw)) || '').replace(/\s+/g, '');
      if (content) return true;
    }
    return false;
  } catch (err) {
    Logger.log('Conversation presence parser fallback: ' + err);
    return !!text.trim();
  }
}

function _getProductChatKeyFromRowObject_(rowObj) {
  if (!rowObj) return '';
  return _getProcessAliasValueFromObject_(rowObj, PROCESS_TRACKING_KEY_ALIASES);
}

function _buildDatabaseConversationFlagMap_(keys) {
  var wanted = {};
  var hasWanted = false;
  var list = Array.isArray(keys) ? keys : [];
  for (var i = 0; i < list.length; i++) {
    var normalized = _normalizeProcessExactKey_(list[i]);
    if (!normalized || wanted[normalized]) continue;
    wanted[normalized] = true;
    hasWanted = true;
  }

  var out = {};
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME_DATABASE || 'Database');
    if (!sheet) return out;

    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var chatCol = _findHeaderIndexLoose_(headers, CONFIG.CHAT_FIELD_NAME || 'SSV Note Product');
    var keyIndexes = _findProcessHeaderIndexes_(headers, PROCESS_TRACKING_KEY_ALIASES);
    var lastRow = sheet.getLastRow();
    if (!headers.length || chatCol < 0 || !keyIndexes.length || lastRow < CONFIG.DATA_START_ROW) return out;

    var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
    var rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var hasConversation = _hasProductChatConversationValue_(row[chatCol]);
      for (var k = 0; k < keyIndexes.length; k++) {
        var key = _normalizeProcessExactKey_(row[keyIndexes[k]]);
        if (!key || (hasWanted && !wanted[key])) continue;
        if (!out.hasOwnProperty(key)) out[key] = false;
        if (hasConversation) out[key] = true;
      }
    }
  } catch (err) {
    Logger.log('Database conversation flag map failed: ' + err);
  }
  return out;
}

function _applyConversationMetaToRows_(rows, access) {
  var list = Array.isArray(rows) ? rows : [];
  if (!list.length) return list;

  var viewCode = _getSelectedViewCode_(access);
  if (viewCode === 'CAT') {
    for (var i = 0; i < list.length; i++) {
      list[i].hasConversation = _hasProductChatConversationValue_(list[i][CONFIG.CHAT_FIELD_NAME || 'SSV Note Product']);
    }
    return list;
  }

  var keys = [];
  for (var r = 0; r < list.length; r++) {
    var rawKey = _getProductChatKeyFromRowObject_(list[r]);
    if (rawKey) keys.push(rawKey);
  }
  var flags = _buildDatabaseConversationFlagMap_(keys);
  for (var j = 0; j < list.length; j++) {
    var normalizedKey = _normalizeProcessExactKey_(_getProductChatKeyFromRowObject_(list[j]));
    list[j].hasConversation = !!(normalizedKey && flags[normalizedKey]);
  }
  return list;
}

function _attachDatabaseConversationMetaToProcessRecords_(records) {
  var list = Array.isArray(records) ? records : [];
  if (!list.length) return list;
  var keys = [];
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].key) keys.push(list[i].key);
  }
  var flags = _buildDatabaseConversationFlagMap_(keys);
  for (var r = 0; r < list.length; r++) {
    var normalizedKey = _normalizeProcessExactKey_(list[r] && list[r].key);
    list[r].hasConversation = !!(normalizedKey && flags[normalizedKey]);
  }
  return list;
}

// ========================================
// PROCESS TRACKING API
// ========================================
var PROCESS_TRACKING_STAGES = [
  'Category',
  'Manager',
  'Director',
  'Present List',
  'Final Information',
  'Setup Product ID'
];

var PROCESS_TRACKING_KEY_ALIASES = [
  'Danh sach chinh',
  'New Product ID',
  'New product ID',
  'Ticket ID'
];

var PROCESS_TRACKING_CARD_LIMIT_DEFAULT = 5000;
var PROCESS_TRACKING_CARD_LIMIT_MAX = 5000;

function _requireProcessTrackingAccess_(params) {
  var accessRes = _resolveAccessForDataApi_(params);
  if (!accessRes.ok) return accessRes;
  var viewCode = _getSelectedViewCode_(accessRes.data);
  if (viewCode !== 'CAT' && viewCode !== 'MAN') {
    return _err('FORBIDDEN', 'Chi co Category va Manager moi duoc su dung process tracking.');
  }
  return accessRes;
}

function _normalizeProcessLooseText_(value) {
  var text = String(value == null ? '' : value);
  if (text && typeof text.normalize === 'function') {
    try { text = text.normalize('NFKC'); } catch (err) {}
  }
  text = text.replace(/[\u200B-\u200D\uFEFF]/g, '');
  var basic = text;
  if (basic && typeof basic.normalize === 'function') {
    try { basic = basic.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (err2) {}
  }
  return basic
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _normalizeProcessExactKey_(value) {
  var text = String(value == null ? '' : value);
  if (text && typeof text.normalize === 'function') {
    try { text = text.normalize('NFKC'); } catch (err) {}
  }
  return text
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function _normalizeProcessProductGroup_(value) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  return _normalizeProcessLooseText_(raw);
}

function _getProcessAliasValueFromObject_(obj, aliases) {
  if (!obj || !aliases || !aliases.length) return '';
  var aliasMap = {};
  for (var a = 0; a < aliases.length; a++) {
    aliasMap[_normalizeProcessLooseText_(aliases[a])] = true;
  }
  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    if (aliasMap[_normalizeProcessLooseText_(key)]) {
      var value = String(obj[key] == null ? '' : obj[key]).trim();
      if (value) return value;
    }
  }
  return '';
}

function _findProcessHeaderIndexes_(headers, aliases) {
  var out = [];
  var seen = {};
  if (!Array.isArray(headers) || !Array.isArray(aliases)) return out;

  var aliasMap = {};
  for (var a = 0; a < aliases.length; a++) {
    aliasMap[_normalizeProcessLooseText_(aliases[a])] = true;
  }

  for (var i = 0; i < headers.length; i++) {
    if (aliasMap[_normalizeProcessLooseText_(headers[i])] && !seen[i]) {
      seen[i] = true;
      out.push(i);
    }
  }
  return out;
}

function _collectProcessCandidateKeysFromRow_(rowObj) {
  var groups = [
    ['Danh sach chinh'],
    ['New Product ID', 'New product ID'],
    ['Ticket ID']
  ];
  var list = [];
  var map = {};

  for (var i = 0; i < groups.length; i++) {
    var raw = _getProcessAliasValueFromObject_(rowObj, groups[i]);
    var normalized = _normalizeProcessExactKey_(raw);
    if (!raw || !normalized || map[normalized]) continue;
    list.push({ raw: raw, normalized: normalized });
    map[normalized] = { raw: raw, normalized: normalized };
  }

  return {
    list: list,
    map: map,
    primary: list.length ? list[0].raw : ''
  };
}

function _openProcessSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var found = _findSheetByLooseName_(ss, [CONFIG.SHEET_NAME_PROCESS || 'Process', 'Process'], ['process']);
  if (!found.sheet) {
    throw new Error('Sheet "' + (CONFIG.SHEET_NAME_PROCESS || 'Process') + '" not found. Available: ' + found.allNames.join(', '));
  }
  return found.sheet;
}

function _normalizeProcessSheetName_(value) {
  var loose = _normalizeProcessLooseText_(value);
  var compact = loose.replace(/\s+/g, '');
  if (!compact) return '';
  if (compact === 'category') return 'Category';
  if (compact === 'manager') return 'Manager';
  if (compact === 'director') return 'Director';
  if (compact === 'presentlist') return 'Present List';
  if (compact === 'finalinformation') return 'Final Information';
  if (compact === 'setupproductid') return 'Setup Product ID';
  return '';
}

function _normalizeProcessStatus_(value) {
  return _normalizeBusinessFlowStatus_(value).statusNormalized;
}

function _getProcessStageOrder_(sheetName) {
  var canonical = _normalizeProcessSheetName_(sheetName);
  for (var i = 0; i < PROCESS_TRACKING_STAGES.length; i++) {
    if (PROCESS_TRACKING_STAGES[i] === canonical) return i;
  }
  return -1;
}

function _normalizeBusinessFlowStatus_(value) {
  var raw = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  var loose = _normalizeProcessLooseText_(raw);
  var compact = loose.replace(/\s+/g, '');
  var normalized = raw;
  var category = 'OTHER';

  if (!compact) {
    normalized = '';
    category = 'BLANK';
  } else if (compact.indexOf('reject') >= 0) {
    normalized = 'Rejected';
    category = 'REJECTED';
  } else if (compact === 'processing' || compact === 'process' || compact === 'inprogress' || compact === 'inprocess' || compact === 'progressing') {
    normalized = 'Processing';
    category = 'PROCESSING_FORWARD';
  } else if (compact === 'done' || compact === 'complete' || compact === 'completed') {
    normalized = 'Done';
    category = 'DONE';
  } else if (compact === 'pending' || compact === 'wait' || compact === 'waiting') {
    normalized = 'Pending';
    category = 'PENDING';
  } else if (compact === 'approved' || compact === 'approve') {
    normalized = 'Approved';
    category = 'APPROVED';
  }

  return {
    raw: raw,
    statusNormalized: normalized,
    statusCategory: category
  };
}

function _getSelectedBusinessSheetForAccess_(access) {
  var selected = String(access && access.selectedSheetName ? access.selectedSheetName : '').trim();
  if (selected) return _normalizeProcessSheetName_(selected) || selected;
  var viewCode = _getSelectedViewCode_(access);
  if (viewCode === 'MAN') return 'Manager';
  if (viewCode === 'DIR') return 'Director';
  return 'Category';
}

function _resolveBusinessFlowState_(storedSheet, status, options) {
  var opts = options || {};
  var rawStoredSheet = String(storedSheet == null ? '' : storedSheet).replace(/\s+/g, ' ').trim();
  var normalizedStoredSheet = _normalizeProcessSheetName_(rawStoredSheet);
  var stored = normalizedStoredSheet || rawStoredSheet;
  var statusInfo = _normalizeBusinessFlowStatus_(status);
  var storedIndex = _getProcessStageOrder_(stored);
  var lastIndex = PROCESS_TRACKING_STAGES.length - 1;
  var nextSheet = (storedIndex >= 0 && storedIndex < lastIndex) ? PROCESS_TRACKING_STAGES[storedIndex + 1] : '';
  var isRejected = statusInfo.statusCategory === 'REJECTED';
  var isProcessingForward = statusInfo.statusCategory === 'PROCESSING_FORWARD';
  var effectiveSheet = stored;
  var effectiveIndex = storedIndex;
  var warnings = [];

  if (!stored) {
    warnings.push('MISSING_STORED_SHEET');
  } else if (storedIndex < 0) {
    warnings.push('UNKNOWN_STORED_SHEET');
  }

  if (isProcessingForward && nextSheet) {
    effectiveSheet = nextSheet;
    effectiveIndex = storedIndex + 1;
  } else if (isProcessingForward && !nextSheet) {
    warnings.push('PROCESSING_AT_LAST_STAGE');
  }

  var selectedSheet = String(opts.selectedSheet == null ? '' : opts.selectedSheet).replace(/\s+/g, ' ').trim();
  selectedSheet = _normalizeProcessSheetName_(selectedSheet) || selectedSheet;
  var selectedMatchesEffective = !!(selectedSheet && effectiveSheet && _normalizeCategoryKey_(selectedSheet) === _normalizeCategoryKey_(effectiveSheet));

  var state = {
    storedSheet: stored,
    storedSheetRaw: rawStoredSheet,
    storedStageIndex: storedIndex,
    effectiveSheet: effectiveSheet,
    effectiveStageIndex: effectiveIndex,
    statusRaw: statusInfo.raw,
    statusNormalized: statusInfo.statusNormalized,
    statusCategory: statusInfo.statusCategory,
    isRejected: isRejected,
    isProcessingForward: isProcessingForward,
    isTerminal: isRejected || statusInfo.statusCategory === 'DONE' || (isProcessingForward && !nextSheet),
    nextSheet: nextSheet,
    selectedSheet: selectedSheet,
    shouldUseWorkspaceChat: selectedMatchesEffective,
    shouldUseProcessChat: selectedSheet ? !selectedMatchesEffective : false,
    warnings: warnings
  };
  state.displayStageLabel = _getBusinessFlowDisplayLabel_(state);
  state.stageTimeline = _buildBusinessFlowStageTimeline_(state);
  return state;
}

function _getBusinessFlowDisplayLabel_(state) {
  var st = state || {};
  var sheet = st.effectiveSheet || st.storedSheet || '';
  if (st.isRejected) return st.storedSheet ? ('Rejected at ' + st.storedSheet) : 'Rejected';
  if (st.isProcessingForward && st.nextSheet) return st.nextSheet + ' Pending';
  if (st.isProcessingForward && !st.nextSheet) return st.storedSheet ? (st.storedSheet + ' Processing') : 'Processing';
  if (st.statusCategory === 'DONE') return st.storedSheet ? (st.storedSheet + ' Done') : 'Done';
  if (st.statusCategory === 'PENDING') return sheet ? (sheet + ' Pending') : 'Pending';
  if (sheet && st.statusNormalized) return sheet + ' | ' + st.statusNormalized;
  if (sheet) return sheet;
  if (st.statusNormalized) return st.statusNormalized;
  return 'Process unavailable';
}

function _buildBusinessFlowStageTimeline_(state) {
  var st = state || {};
  var timeline = [];
  for (var i = 0; i < PROCESS_TRACKING_STAGES.length; i++) {
    timeline.push({ name: PROCESS_TRACKING_STAGES[i], state: 'pending' });
  }

  var storedIndex = parseInt(st.storedStageIndex, 10);
  if (isNaN(storedIndex) || storedIndex < 0 || storedIndex >= timeline.length) return timeline;

  if (st.isRejected) {
    for (i = 0; i < storedIndex; i++) timeline[i].state = 'done';
    timeline[storedIndex].state = 'rejected';
    return timeline;
  }

  if (st.isProcessingForward) {
    for (i = 0; i <= storedIndex; i++) timeline[i].state = 'done';
    var effectiveIndex = parseInt(st.effectiveStageIndex, 10);
    if (!isNaN(effectiveIndex) && effectiveIndex > storedIndex && effectiveIndex < timeline.length) {
      timeline[effectiveIndex].state = 'current';
    } else {
      timeline[storedIndex].state = 'current';
    }
    return timeline;
  }

  if (st.statusCategory === 'DONE') {
    for (i = 0; i <= storedIndex; i++) timeline[i].state = 'done';
    return timeline;
  }

  for (i = 0; i < storedIndex; i++) timeline[i].state = 'done';
  timeline[storedIndex].state = 'current';
  return timeline;
}

function _getDisplayedProcessLabel_(sheetName, status) {
  return _resolveBusinessFlowState_(sheetName, status).displayStageLabel;
}

function _buildProcessStageTimeline_(sheetName, status) {
  return _resolveBusinessFlowState_(sheetName, status).stageTimeline;
}

function _extractGoogleDriveFileId_(url) {
  var raw = String(url || '').trim();
  if (!raw) return '';
  var match = raw.match(/\/d\/([^/?#]+)/i);
  if (match && match[1]) return match[1];
  match = raw.match(/[?&]id=([^&#]+)/i);
  if (match && match[1]) return match[1];
  return '';
}

function _buildProcessPreviewImageUrl_(url) {
  var raw = String(url || '').trim();
  if (!raw) return '';
  var fileId = _extractGoogleDriveFileId_(raw);
  if (fileId) return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';
  return raw;
}

function _extractPrimaryImageUrl_(rawUrl) {
  if (rawUrl === null || rawUrl === undefined) return '';
  var parts = String(rawUrl).split(/[;\n]+/);
  for (var i = 0; i < parts.length; i++) {
    var item = String(parts[i] || '').trim();
    if (item) return item;
  }
  return '';
}

function _formatProcessLastSeen_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var tz = '';
    try { tz = Session.getScriptTimeZone(); } catch (err) { tz = ''; }
    return Utilities.formatDate(value, tz || 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value == null ? '' : value).trim();
}

function _parseProcessLastSeenTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && isFinite(value)) return value;
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return -1;
  var parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return parsed.getTime();
  return -1;
}

function _isProcessHeaderLikeRow_(rowValues) {
  if (!Array.isArray(rowValues) || !rowValues.length) return false;
  var checks = 0;
  var col0 = _normalizeProcessLooseText_(rowValues[0]);
  var col1 = _normalizeProcessLooseText_(rowValues[1]);
  var col2 = _normalizeProcessLooseText_(rowValues[2]);
  var col3 = _normalizeProcessLooseText_(rowValues[3]);
  var col4 = _normalizeProcessLooseText_(rowValues[4]);
  var col5 = _normalizeProcessLooseText_(rowValues[5]);
  var col6 = _normalizeProcessLooseText_(rowValues[6]);
  var col7 = _normalizeProcessLooseText_(rowValues[7]);
  if (col0 === 'danh sach chinh' || col0 === 'main list') checks++;
  if (col1 === 'sheet') checks++;
  if (col2 === 'last seen' || col2 === 'last_seen') checks++;
  if (col3 === 'status') checks++;
  if (col4 === 'product name') checks++;
  if (col5.indexOf('uom') >= 0) checks++;
  if (col6.indexOf('image') >= 0 || col6.indexOf('url') >= 0) checks++;
  if (col7 === 'product group') checks++;
  return checks >= 3;
}

function _buildProcessRecordFromRow_(rowValues, rowNumber) {
  if (!Array.isArray(rowValues) || rowValues.length < 8) return null;

  var key = String(rowValues[0] == null ? '' : rowValues[0]).trim();
  var keyNorm = _normalizeProcessExactKey_(key);
  if (!keyNorm) return null;

  var normalizedSheet = _normalizeProcessSheetName_(rowValues[1]);
  var normalizedStatus = _normalizeProcessStatus_(rowValues[3]);
  var displaySheet = normalizedSheet || String(rowValues[1] == null ? '' : rowValues[1]).trim();
  var displayStatus = normalizedStatus || String(rowValues[3] == null ? '' : rowValues[3]).trim();
  var flowState = _resolveBusinessFlowState_(displaySheet, displayStatus);
  var imageOriginalUrl = _extractPrimaryImageUrl_(rowValues[6]);
  var productGroup = String(rowValues[7] == null ? '' : rowValues[7]).trim();

  return {
    rowNumber: rowNumber,
    key: key,
    keyNorm: keyNorm,
    currentSheet: displaySheet,
    currentStatus: displayStatus,
    lastSeen: _formatProcessLastSeen_(rowValues[2]),
    lastSeenTs: _parseProcessLastSeenTimestamp_(rowValues[2]),
    productName: String(rowValues[4] == null ? '' : rowValues[4]).trim(),
    barcodeOrUom: String(rowValues[5] == null ? '' : rowValues[5]).trim(),
    imageOriginalUrl: imageOriginalUrl,
    imageUrl: _buildProcessPreviewImageUrl_(imageOriginalUrl),
    productGroup: productGroup,
    productGroupNorm: _normalizeProcessProductGroup_(productGroup),
    storedSheet: flowState.storedSheet || displaySheet,
    effectiveSheet: flowState.effectiveSheet || displaySheet,
    flowState: flowState,
    displayStageLabel: flowState.displayStageLabel,
    stageTimeline: flowState.stageTimeline
  };
}

function _readProcessRecords_() {
  var opened;
  try {
    opened = _openProcessSheet_();
  } catch (err) {
    return {
      sheet: null,
      records: [],
      warning: { code: 'PROCESS_SHEET_NOT_FOUND', message: err.message }
    };
  }

  var sheet = opened;
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return { sheet: sheet, records: [], warning: null };
  }

  var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 3, rowCount, 8).getValues();
  var records = [];

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
    if (_isProcessHeaderLikeRow_(row)) continue;

    var record = _buildProcessRecordFromRow_(row, CONFIG.DATA_START_ROW + i);
    if (record) records.push(record);
  }

  return { sheet: sheet, records: records, warning: null };
}

function _readAccessScopedRowsFromCurrentSheet_(access) {
  var sheet = _openSheet(access);
  var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
  var lastRow = sheet.getLastRow();
  if (!headers.length || lastRow < CONFIG.DATA_START_ROW) {
    return { sheet: sheet, headers: headers, rows: [] };
  }

  var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();
  var rows = [];

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
    rows.push(_buildRowObjectFromArray_(headers, row, CONFIG.DATA_START_ROW + i));
  }

  return {
    sheet: sheet,
    headers: headers,
    rows: _applyAccessScopeToRows_(rows, access)
  };
}

function _pushAllowedProcessGroup_(collector, rawValue, source) {
  var raw = String(rawValue == null ? '' : rawValue).trim();
  var normalized = _normalizeProcessProductGroup_(raw);
  if (!raw || !normalized) return;
  if (!collector.map[normalized]) {
    collector.map[normalized] = raw;
    collector.list.push(raw);
  }
  if (source) collector.sources[source] = true;
}

function _getAllowedProcessGroupsForAccess_(access, records) {
  var info = {
    list: [],
    map: {},
    sources: {},
    warning: null
  };

  if (!access) {
    info.warning = {
      code: 'PROCESS_GROUP_SCOPE_EMPTY',
      message: 'Khong the xac dinh Product Group hop le cho tai khoan hien tai.'
    };
    return info;
  }

  var accessibleRows = [];
  try {
    accessibleRows = _readAccessScopedRowsFromCurrentSheet_(access).rows || [];
  } catch (err) {
    accessibleRows = [];
  }

  for (var i = 0; i < accessibleRows.length; i++) {
    var row = accessibleRows[i] || {};
    _pushAllowedProcessGroup_(info, row['Product Group'] || row['PRODUCT GROUP'] || '', 'currentSheet');
  }

  var processGroupCatalog = {};
  var processRows = Array.isArray(records) ? records : [];
  for (var r = 0; r < processRows.length; r++) {
    var record = processRows[r];
    if (!record || !record.productGroupNorm || processGroupCatalog[record.productGroupNorm]) continue;
    processGroupCatalog[record.productGroupNorm] = record.productGroup;
  }

  var scopes = Array.isArray(access.scope) ? access.scope : [];
  for (var s = 0; s < scopes.length; s++) {
    var normalizedScope = _normalizeProcessProductGroup_(scopes[s]);
    if (!normalizedScope || !processGroupCatalog[normalizedScope]) continue;
    _pushAllowedProcessGroup_(info, processGroupCatalog[normalizedScope], 'scope');
  }

  if (!info.list.length) {
    info.warning = {
      code: 'PROCESS_GROUP_SCOPE_EMPTY',
      message: 'Khong xac dinh duoc Product Group hop le. Danh sach process se an toan va tra ve rong cho den khi exact code lookup duoc xac minh scope.'
    };
  }

  return info;
}

function _isProcessRecordAllowedByGroup_(record, allowedGroupInfo) {
  return !!(record && record.productGroupNorm && allowedGroupInfo && allowedGroupInfo.map && allowedGroupInfo.map[record.productGroupNorm]);
}

function _filterProcessRowsByProductGroup_(rows, access, options) {
  var list = Array.isArray(rows) ? rows : [];
  var opts = options || {};
  var allowedGroupInfo = opts.allowedGroupInfo || _getAllowedProcessGroupsForAccess_(access, list);
  if (!allowedGroupInfo.list.length) return [];

  return list.filter(function(record) {
    return _isProcessRecordAllowedByGroup_(record, allowedGroupInfo);
  });
}

function _getProcessSheetLookupTokens_(sheetName) {
  var canonical = _normalizeProcessSheetName_(sheetName);
  if (!canonical) return [];
  var parts = _normalizeProcessLooseText_(canonical).split(' ');
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    if (parts[i]) out.push(parts[i]);
  }
  return out;
}

function _openProcessSourceSheet_(sheetName) {
  var canonical = _normalizeProcessSheetName_(sheetName);
  if (!canonical) return { sheet: null, allNames: [] };
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return _findSheetByLooseName_(ss, [canonical], _getProcessSheetLookupTokens_(canonical));
}

function _readProcessSourceSheetRows_(sheetName) {
  var opened = _openProcessSourceSheet_(sheetName);
  if (!opened.sheet) {
    return { sheet: null, available: opened.allNames || [], headers: [], rows: [] };
  }

  var sheet = opened.sheet;
  var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
  var lastRow = sheet.getLastRow();
  if (!headers.length || lastRow < CONFIG.DATA_START_ROW) {
    return { sheet: sheet, available: opened.allNames || [], headers: headers, rows: [] };
  }

  var rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  var rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, headers.length).getValues();
  return { sheet: sheet, available: opened.allNames || [], headers: headers, rows: rows };
}

function _verifyProcessRecordAccessStrict_(record, access) {
  if (!record || !access) return false;
  if (_isDirectorViewAccess_(access)) return true;

  var sourceTable = _readProcessSourceSheetRows_(record.currentSheet);
  if (!sourceTable.sheet || !sourceTable.headers.length || !sourceTable.rows.length) return false;

  var keyIndexes = _findProcessHeaderIndexes_(sourceTable.headers, PROCESS_TRACKING_KEY_ALIASES);
  if (!keyIndexes.length) return false;

  for (var i = 0; i < sourceTable.rows.length; i++) {
    var row = sourceTable.rows[i];
    var matched = false;
    for (var k = 0; k < keyIndexes.length; k++) {
      var idx = keyIndexes[k];
      if (_normalizeProcessExactKey_(row[idx]) === record.keyNorm) {
        matched = true;
        break;
      }
    }
    if (!matched) continue;

    var rowObj = _buildRowObjectFromArray_(sourceTable.headers, row, CONFIG.DATA_START_ROW + i);
    if (_hasRowAccess_(rowObj, access)) return true;
  }

  return false;
}

function _filterProcessByAccess_(records, access, options) {
  var list = Array.isArray(records) ? records : [];
  var opts = options || {};
  var allowedKeys = opts.allowedKeys || null;
  var rowSnapshot = opts.rowSnapshot || null;
  var requireVerifiedAccess = opts.requireVerifiedAccess === true;
  var enforceProductGroup = opts.enforceProductGroup === true;
  var allowBlankGroupWhenVerified = opts.allowBlankGroupWhenVerified === true;
  var allowVerifiedAccessBypassGroup = opts.allowVerifiedAccessBypassGroup === true;
  var allowedGroupInfo = opts.allowedGroupInfo || null;
  var out = [];

  if (rowSnapshot && !_hasRowAccess_(rowSnapshot, access)) return [];
  if (enforceProductGroup && !allowedGroupInfo) {
    allowedGroupInfo = _getAllowedProcessGroupsForAccess_(access, list);
  }

  for (var i = 0; i < list.length; i++) {
    var record = list[i];
    if (!record) continue;
    if (allowedKeys && !allowedKeys[record.keyNorm]) continue;
    if (enforceProductGroup) {
      var groupAllowed = _isProcessRecordAllowedByGroup_(record, allowedGroupInfo);
      var verifiedAccess = null;
      if (!groupAllowed && (allowBlankGroupWhenVerified || allowVerifiedAccessBypassGroup) && requireVerifiedAccess) {
        verifiedAccess = _verifyProcessRecordAccessStrict_(record, access);
      }
      if (!groupAllowed && allowBlankGroupWhenVerified && !record.productGroupNorm && verifiedAccess === true) {
        groupAllowed = true;
      }
      if (!groupAllowed && allowVerifiedAccessBypassGroup && verifiedAccess === true) {
        groupAllowed = true;
      }
      if (!groupAllowed) continue;
      if (requireVerifiedAccess && verifiedAccess !== true && !_verifyProcessRecordAccessStrict_(record, access)) continue;
    } else if (requireVerifiedAccess && !_verifyProcessRecordAccessStrict_(record, access)) {
      continue;
    }
    out.push(record);
  }

  return out;
}

function _getProcessRecordStageRank_(record) {
  if (!record) return -1;
  return _getProcessStageOrder_(record.currentSheet);
}

function _compareProcessRecords_(candidate, current, keyRankMap) {
  if (!candidate && !current) return 0;
  if (candidate && !current) return 1;
  if (!candidate && current) return -1;

  if (keyRankMap) {
    var candidateKeyRank = keyRankMap.hasOwnProperty(candidate.keyNorm) ? keyRankMap[candidate.keyNorm] : 999999;
    var currentKeyRank = keyRankMap.hasOwnProperty(current.keyNorm) ? keyRankMap[current.keyNorm] : 999999;
    if (candidateKeyRank !== currentKeyRank) return candidateKeyRank < currentKeyRank ? 1 : -1;
  }

  var candidateTs = typeof candidate.lastSeenTs === 'number' ? candidate.lastSeenTs : -1;
  var currentTs = typeof current.lastSeenTs === 'number' ? current.lastSeenTs : -1;
  if (candidateTs !== currentTs) return candidateTs > currentTs ? 1 : -1;

  var candidateStage = _getProcessRecordStageRank_(candidate);
  var currentStage = _getProcessRecordStageRank_(current);
  if (candidateStage !== currentStage) return candidateStage > currentStage ? 1 : -1;

  var candidateRow = parseInt(candidate.rowNumber, 10) || 0;
  var currentRow = parseInt(current.rowNumber, 10) || 0;
  if (candidateRow !== currentRow) return candidateRow > currentRow ? 1 : -1;

  return 0;
}

function _buildProcessKeyRankMap_(candidateKeys) {
  var map = {};
  var list = Array.isArray(candidateKeys) ? candidateKeys : [];
  for (var i = 0; i < list.length; i++) {
    var item = list[i];
    var normalized = typeof item === 'string' ? item : (item && item.normalized);
    if (!normalized || map.hasOwnProperty(normalized)) continue;
    map[normalized] = i;
  }
  return map;
}

function _findBestProcessMatch_(records, candidateKeys) {
  var list = Array.isArray(records) ? records : [];
  var rankMap = candidateKeys ? _buildProcessKeyRankMap_(candidateKeys) : null;
  var hasRankMap = rankMap && Object.keys(rankMap).length > 0;
  var best = null;

  for (var i = 0; i < list.length; i++) {
    var record = list[i];
    if (!record) continue;
    if (hasRankMap && !rankMap.hasOwnProperty(record.keyNorm)) continue;
    if (_compareProcessRecords_(record, best, rankMap) > 0) best = record;
  }

  return best;
}

function _buildBestProcessIndexByKey_(records) {
  var out = {};
  var list = Array.isArray(records) ? records : [];
  for (var i = 0; i < list.length; i++) {
    var record = list[i];
    if (!record || !record.keyNorm) continue;
    if (!out[record.keyNorm] || _compareProcessRecords_(record, out[record.keyNorm], null) > 0) {
      out[record.keyNorm] = record;
    }
  }
  return out;
}

function _getRowSnapshotsByNumber_(sheet, headers, rowNumbers) {
  var out = {};
  var unique = [];
  var seen = {};
  var lastRow = sheet.getLastRow();
  var i;

  for (i = 0; i < rowNumbers.length; i++) {
    var rowNumber = parseInt(rowNumbers[i], 10);
    if (!rowNumber || rowNumber < CONFIG.DATA_START_ROW || rowNumber > lastRow || seen[rowNumber]) continue;
    seen[rowNumber] = true;
    unique.push(rowNumber);
  }

  if (!unique.length) return out;

  unique.sort(function(a, b) { return a - b; });
  var minRow = unique[0];
  var maxRow = unique[unique.length - 1];
  var raw = sheet.getRange(minRow, 1, maxRow - minRow + 1, headers.length).getValues();

  for (i = 0; i < unique.length; i++) {
    var rn = unique[i];
    out[rn] = _buildRowObjectFromArray_(headers, raw[rn - minRow], rn);
  }

  return out;
}

function _getProcessRowFallbackSummary_(rowObj) {
  var rawImage = _extractPrimaryImageUrl_(_getProcessAliasValueFromObject_(rowObj, ['URL Image', 'Url Image', 'URL IMAGE']));
  return {
    productName: _getProcessAliasValueFromObject_(rowObj, ['Product Name', 'PIC input Product Name (Duoi 40 ki tu)']),
    barcodeOrUom: _getProcessAliasValueFromObject_(rowObj, ['UOM Information', 'UOM INFORMATION']),
    productGroup: _getProcessAliasValueFromObject_(rowObj, ['Product Group']),
    imageOriginalUrl: rawImage,
    imageUrl: _buildProcessPreviewImageUrl_(rawImage)
  };
}

function _buildProcessCardSummary_(record, fallback) {
  var base = fallback || {};
  var current = record || {};
  var fallbackFlowState = base.flowState || null;
  var currentFlowState = current.flowState || null;
  return {
    key: current.key || base.key || '',
    productName: current.productName || base.productName || '',
    imageUrl: current.imageUrl || base.imageUrl || '',
    imageOriginalUrl: current.imageOriginalUrl || base.imageOriginalUrl || '',
    barcodeOrUom: current.barcodeOrUom || base.barcodeOrUom || '',
    productGroup: current.productGroup || base.productGroup || '',
    currentSheet: current.currentSheet || '',
    storedSheet: current.storedSheet || current.currentSheet || '',
    effectiveSheet: current.effectiveSheet || (currentFlowState && currentFlowState.effectiveSheet) || (fallbackFlowState && fallbackFlowState.effectiveSheet) || '',
    currentStatus: current.currentStatus || '',
    displayStageLabel: current.displayStageLabel || (currentFlowState && currentFlowState.displayStageLabel) || base.displayStageLabel || (fallbackFlowState && fallbackFlowState.displayStageLabel) || '',
    lastSeen: current.lastSeen || '',
    hasConversation: current.hasConversation === true || base.hasConversation === true,
    flowState: currentFlowState || fallbackFlowState || null
  };
}

function _buildProcessDetailModel_(record, fallback) {
  var summary = _buildProcessCardSummary_(record, fallback);
  return {
    key: summary.key,
    productName: summary.productName,
    imageUrl: summary.imageUrl,
    imageOriginalUrl: summary.imageOriginalUrl,
    barcodeOrUom: summary.barcodeOrUom,
    productGroup: summary.productGroup,
    currentSheet: summary.currentSheet,
    storedSheet: summary.storedSheet,
    effectiveSheet: summary.effectiveSheet,
    currentStatus: summary.currentStatus,
    displayStageLabel: summary.displayStageLabel,
    lastSeen: summary.lastSeen,
    hasConversation: summary.hasConversation === true,
    stageTimeline: record ? record.stageTimeline : (summary.flowState ? summary.flowState.stageTimeline : _buildProcessStageTimeline_('', '')),
    flowState: summary.flowState || null
  };
}

function _buildProcessLookupEntry_(record, rowSnapshot, primaryKey, access) {
  var fallback = rowSnapshot ? _getProcessRowFallbackSummary_(rowSnapshot) : {};
  var selectedSheet = _getSelectedBusinessSheetForAccess_(access);
  if (primaryKey && !fallback.key) fallback.key = primaryKey;
  if (!record) {
    fallback.flowState = _resolveBusinessFlowState_(selectedSheet, '', { selectedSheet: selectedSheet });
    fallback.displayStageLabel = fallback.flowState.displayStageLabel;
    var emptySummary = _buildProcessCardSummary_(null, fallback);
    return {
      exists: false,
      key: emptySummary.key,
      label: emptySummary.displayStageLabel || '',
      currentSheet: emptySummary.currentSheet || selectedSheet || '',
      storedSheet: emptySummary.storedSheet || selectedSheet || '',
      effectiveSheet: emptySummary.effectiveSheet || selectedSheet || '',
      currentStatus: '',
      lastSeen: '',
      imageUrl: emptySummary.imageUrl,
      imageOriginalUrl: emptySummary.imageOriginalUrl,
      productName: emptySummary.productName,
      barcodeOrUom: emptySummary.barcodeOrUom,
      productGroup: emptySummary.productGroup,
      hasConversation: emptySummary.hasConversation === true,
      flowState: emptySummary.flowState || null
    };
  }

  var summary = _buildProcessCardSummary_(record, fallback);
  var selectedFlowState = _resolveBusinessFlowState_(summary.currentSheet, summary.currentStatus, { selectedSheet: selectedSheet });
  return {
    exists: true,
    key: summary.key || primaryKey || '',
    label: selectedFlowState.displayStageLabel || record.displayStageLabel || '',
    currentSheet: summary.currentSheet || '',
    storedSheet: summary.storedSheet || summary.currentSheet || '',
    effectiveSheet: selectedFlowState.effectiveSheet || summary.effectiveSheet || summary.currentSheet || '',
    currentStatus: summary.currentStatus || '',
    lastSeen: summary.lastSeen || '',
    imageUrl: summary.imageUrl || '',
    imageOriginalUrl: summary.imageOriginalUrl || '',
    productName: summary.productName || '',
    barcodeOrUom: summary.barcodeOrUom || '',
    productGroup: summary.productGroup || '',
    hasConversation: summary.hasConversation === true,
    flowState: selectedFlowState
  };
}

function _buildProcessDetailPayload_(record, rowSnapshot, primaryKey, access) {
  var fallback = rowSnapshot ? _getProcessRowFallbackSummary_(rowSnapshot) : {};
  if (primaryKey && !fallback.key) fallback.key = primaryKey;
  if (!record) {
    var selectedSheet = _getSelectedBusinessSheetForAccess_(access);
    fallback.flowState = _resolveBusinessFlowState_(selectedSheet, '', { selectedSheet: selectedSheet });
    fallback.displayStageLabel = fallback.flowState.displayStageLabel;
  }
  var model = _buildProcessDetailModel_(record, fallback);
  var selectedSheet = _getSelectedBusinessSheetForAccess_(access);
  if (model.currentSheet || selectedSheet) {
    var flowState = _resolveBusinessFlowState_(model.currentSheet || selectedSheet, model.currentStatus || '', { selectedSheet: selectedSheet });
    model.flowState = flowState;
    model.storedSheet = flowState.storedSheet || model.currentSheet || '';
    model.effectiveSheet = flowState.effectiveSheet || model.currentSheet || '';
    model.displayStageLabel = flowState.displayStageLabel || model.displayStageLabel || '';
    model.stageTimeline = flowState.stageTimeline || model.stageTimeline || [];
  }
  return model;
}

function _isProcessListEligible_(record) {
  if (!record || !record.keyNorm) return false;
  return !(_normalizeProcessSheetName_(record.currentSheet) === 'Setup Product ID' && _normalizeProcessStatus_(record.currentStatus) === 'Done');
}

function _buildProcessSearchText_(record) {
  if (!record) return '';
  return _normalizeProcessLooseText_([
    record.key || '',
    record.productName || '',
    record.barcodeOrUom || '',
    record.productGroup || '',
    record.currentSheet || '',
    record.currentStatus || '',
    record.displayStageLabel || ''
  ].join(' '));
}

function _matchesProcessingProductsQuery_(record, query) {
  var raw = String(query == null ? '' : query).trim();
  if (!raw) return true;
  var exact = _normalizeProcessExactKey_(raw);
  if (exact && record && record.keyNorm === exact) return true;
  var loose = _normalizeProcessLooseText_(raw);
  if (!loose) return true;
  return _buildProcessSearchText_(record).indexOf(loose) >= 0;
}

function _getProcessingProductsLimit_(value) {
  var limit = parseInt(value, 10);
  if (!limit || limit < 1) return PROCESS_TRACKING_CARD_LIMIT_DEFAULT;
  return Math.min(limit, PROCESS_TRACKING_CARD_LIMIT_MAX);
}

function getProcessingProducts(params) {
  try {
    params = params || {};
    var accessRes = _requireProcessTrackingAccess_(params);
    if (!accessRes.ok) return accessRes;

    var processTable = _readProcessRecords_();
    if (processTable.warning) {
      return _ok({
        items: [],
        total: 0,
        returned: 0,
        limit: _getProcessingProductsLimit_(params.limit),
        allowedGroups: [],
        warning: processTable.warning
      });
    }

    var allowedGroupInfo = _getAllowedProcessGroupsForAccess_(accessRes.data, processTable.records);
    var filtered = _filterProcessRowsByProductGroup_(processTable.records, accessRes.data, {
      allowedGroupInfo: allowedGroupInfo
    });
    var bestByKey = _buildBestProcessIndexByKey_(filtered);
    var deduped = [];
    for (var keyNorm in bestByKey) {
      if (!bestByKey.hasOwnProperty(keyNorm)) continue;
      deduped.push(bestByKey[keyNorm]);
    }

    deduped = deduped.filter(function(record) {
      return _isProcessListEligible_(record);
    });
    _attachDatabaseConversationMetaToProcessRecords_(deduped);

    var query = String(params.query || '').trim();
    if (query) {
      deduped = deduped.filter(function(record) {
        return _matchesProcessingProductsQuery_(record, query);
      });
    }

    deduped.sort(function(a, b) {
      var compared = _compareProcessRecords_(a, b, null);
      if (compared > 0) return -1;
      if (compared < 0) return 1;
      return 0;
    });

    var total = deduped.length;
    var limit = params.fullDataset === true ? PROCESS_TRACKING_CARD_LIMIT_MAX : _getProcessingProductsLimit_(params.limit);
    var selectedSheet = _getSelectedBusinessSheetForAccess_(accessRes.data);
    var items = deduped.slice(0, limit).map(function(record) {
      var item = _buildProcessCardSummary_(record, null);
      var flowState = _resolveBusinessFlowState_(item.currentSheet, item.currentStatus, { selectedSheet: selectedSheet });
      item.flowState = flowState;
      item.storedSheet = flowState.storedSheet || item.currentSheet || '';
      item.effectiveSheet = flowState.effectiveSheet || item.currentSheet || '';
      item.displayStageLabel = flowState.displayStageLabel || item.displayStageLabel || '';
      return item;
    });

    return _ok({
      items: items,
      total: total,
      returned: items.length,
      limit: limit,
      fullDataset: params.fullDataset === true,
      query: query,
      allowedGroups: allowedGroupInfo.list,
      warning: allowedGroupInfo.warning
    });
  } catch (err) {
    Logger.log('getProcessingProducts ERROR: ' + err.toString());
    return _err('PROCESS_PRODUCTS_ERROR', err.message, err.stack);
  }
}

function getProcessLookup(params) {
  try {
    params = params || {};
    var accessRes = _requireProcessTrackingAccess_(params);
    if (!accessRes.ok) return accessRes;

    var items = Array.isArray(params.items) ? params.items : [];
    var rowNumbers = [];
    for (var i = 0; i < items.length; i++) {
      var rowNumber = parseInt(items[i] && items[i].rowNumber, 10);
      if (rowNumber) rowNumbers.push(rowNumber);
    }

    if (!rowNumbers.length) {
      return _ok({ byRowNumber: {}, byKey: {} });
    }

    var sheet = _openSheet(accessRes.data);
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    if (!headers.length || sheet.getLastRow() < CONFIG.DATA_START_ROW) {
      return _ok({ byRowNumber: {}, byKey: {} });
    }

    var rowSnapshots = _getRowSnapshotsByNumber_(sheet, headers, rowNumbers);
    var processTable = _readProcessRecords_();
    if (processTable.warning) {
      return _ok({
        byRowNumber: {},
        byKey: {},
        warning: processTable.warning
      });
    }

    var bestByKey = _buildBestProcessIndexByKey_(processTable.records);
    var byRowNumber = {};
    var byKey = {};

    for (var r = 0; r < rowNumbers.length; r++) {
      var rn = rowNumbers[r];
      var rowSnapshot = rowSnapshots[rn];
      if (!rowSnapshot || !_hasRowAccess_(rowSnapshot, accessRes.data)) continue;

      var keyInfo = _collectProcessCandidateKeysFromRow_(rowSnapshot);
      if (!keyInfo.list.length) continue;

      var candidates = [];
      for (var k = 0; k < keyInfo.list.length; k++) {
        var keyNorm = keyInfo.list[k].normalized;
        if (bestByKey[keyNorm]) candidates.push(bestByKey[keyNorm]);
      }

      var matched = _findBestProcessMatch_(candidates, keyInfo.list);
      var entry = _buildProcessLookupEntry_(matched, rowSnapshot, keyInfo.primary, accessRes.data);
      byRowNumber[String(rn)] = entry;
      if (entry.key) {
        var existing = byKey[entry.key];
        if (!existing || (entry.exists && !existing.exists)) byKey[entry.key] = entry;
      }
    }

    return _ok({
      byRowNumber: byRowNumber,
      byKey: byKey
    });
  } catch (err) {
    Logger.log('getProcessLookup ERROR: ' + err.toString());
    return _err('PROCESS_LOOKUP_ERROR', err.message, err.stack);
  }
}

function getProcessDetail(params) {
  try {
    params = params || {};
    var accessRes = _requireProcessTrackingAccess_(params);
    if (!accessRes.ok) return accessRes;

    var processTable = _readProcessRecords_();
    if (processTable.warning) {
      return _err(processTable.warning.code || 'PROCESS_SHEET_NOT_FOUND', processTable.warning.message || 'Không thể đọc sheet Process.');
    }

    if (params.rowNumber) {
      var rowNumber = parseInt(params.rowNumber, 10);
      if (!rowNumber) return _err('INVALID_ROW', 'Số dòng không hợp lệ.');

      var sheet = _openSheet(accessRes.data);
      var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
      var lastRow = sheet.getLastRow();
      if (!headers.length || rowNumber < CONFIG.DATA_START_ROW || rowNumber > lastRow) {
        return _err('INVALID_ROW', 'Số dòng nằm ngoài phạm vi dữ liệu.');
      }

      var rowSnapshot = _getRowSnapshotByNumber_(sheet, headers, rowNumber);
      if (!_hasRowAccess_(rowSnapshot, accessRes.data)) {
        return _err('FORBIDDEN_SCOPE', 'Tai khoan hien tai khong duoc phep xem process cua san pham nay.');
      }

      var rowKeys = _collectProcessCandidateKeysFromRow_(rowSnapshot);
      if (!rowKeys.list.length) {
        return _err('MISSING_PROCESS_KEY', 'San pham chua co ma hop le de tra process.');
      }

      var allowedRecords = _filterProcessByAccess_(processTable.records, accessRes.data, {
        rowSnapshot: rowSnapshot,
        allowedKeys: rowKeys.map
      });
      var matchedRowRecord = _findBestProcessMatch_(allowedRecords, rowKeys.list);
      if (!matchedRowRecord) {
        return _err('PROCESS_NOT_FOUND', 'Khong tim thay du lieu process cho ma: ' + rowKeys.primary);
      }
      _attachDatabaseConversationMetaToProcessRecords_([matchedRowRecord]);

      return _ok(_buildProcessDetailPayload_(matchedRowRecord, rowSnapshot, rowKeys.primary, accessRes.data));
    }

    var query = String(params.query || '').trim();
    var queryNorm = _normalizeProcessExactKey_(query);
    if (!queryNorm) return _err('MISSING_QUERY', 'Vui long nhap ma san pham de tra process.');

    var queryMatches = [];
    for (var i = 0; i < processTable.records.length; i++) {
      if (processTable.records[i].keyNorm === queryNorm) queryMatches.push(processTable.records[i]);
    }

    if (!queryMatches.length) {
      return _err('PROCESS_NOT_FOUND', 'Khong tim thay du lieu process cho ma: ' + query);
    }

    var allowedGroupInfo = _getAllowedProcessGroupsForAccess_(accessRes.data, queryMatches.length ? queryMatches : processTable.records);
    var verifiedMatches = _filterProcessByAccess_(queryMatches, accessRes.data, {
      requireVerifiedAccess: true,
      enforceProductGroup: true,
      allowBlankGroupWhenVerified: true,
      allowVerifiedAccessBypassGroup: true,
      allowedGroupInfo: allowedGroupInfo
    });
    var matchedQueryRecord = _findBestProcessMatch_(verifiedMatches, [{ raw: query, normalized: queryNorm }]);
    if (!matchedQueryRecord) {
      return _err('FORBIDDEN_SCOPE', 'Khong the hien thi process cho ma nay trong pham vi quyen hien tai.');
    }
    _attachDatabaseConversationMetaToProcessRecords_([matchedQueryRecord]);

    return _ok(_buildProcessDetailPayload_(matchedQueryRecord, null, query, accessRes.data));
  } catch (err) {
    Logger.log('getProcessDetail ERROR: ' + err.toString());
    return _err('PROCESS_DETAIL_ERROR', err.message, err.stack);
  }
}

// ========================================
// API: updateRow
// ========================================
function _getEditableFieldsForAccess_(access) {
  var canEdit = access && access.currentViewCanEdit === true;
  if (!canEdit) return [];
  if (_isDirectorViewAccess_(access)) return DIRECTOR_EDITABLE_FIELDS.slice();
  return PIC_EDITABLE_FIELDS.slice();
}

function _requireWriteAccess_(token) {
  var accessRes = _getAccessByToken_(token);
  if (!accessRes.ok) return accessRes;
  if (!accessRes.data || accessRes.data.currentViewCanEdit !== true) {
    return _err('FORBIDDEN', 'Tai khoan hien tai khong co quyen chinh sua o view da chon.');
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
    if (_isProtectedChatField_(field)) {
      Logger.log('Skipping protected chat field from product save API: ' + field);
      continue;
    }
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
    var valueToWrite = _normalizeProductWriteValue_(field, changes[field]);
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

    if (!payload || !payload.rowNumber || !payload.changes) return _err('INVALID_PAYLOAD', 'Thiếu rowNumber hoặc dữ liệu thay đổi.');

    var accessRes = _requireWriteAccess_(payload.token);
    if (!accessRes.ok) return accessRes;
    var validation = _validateChanges(payload.changes);
    if (!validation.ok) return validation;

    var sheet = _openSheet(accessRes.data);
    var headers = _trimTrailingEmptyHeaders_(_getHeaders(sheet));
    var allowedFields = _getEditableFieldsForAccess_(accessRes.data);

    var lastRow = sheet.getLastRow();
    if (payload.rowNumber < CONFIG.DATA_START_ROW || payload.rowNumber > lastRow) {
      return _err('INVALID_ROW', 'Dòng ' + payload.rowNumber + ' nằm ngoài phạm vi (' +
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
    if (!payload || !Array.isArray(payload.updates)) return _err('INVALID_PAYLOAD', 'Thiếu danh sách cập nhật.');

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

  if ('Product Group' in changes) {
    if (!_isAllowedProductGroup_(changes['Product Group'])) {
      errors.push('Product Group khong hop le.');
    }
  }

  if ('Pattern' in changes) {
    if (!_isAllowedPatternValue_(changes['Pattern'])) {
      errors.push('Pattern phải là South, North hoặc South, North.');
    } else {
      changes['Pattern'] = _normalizePatternValue_(changes['Pattern']);
    }
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
  tpl.appUserJson = _safeJsonForInlineScript_({
    email: 'guest@npa.local',
    fullName: 'Guest User',
    role: 'GUEST',
    scope: ['ALL'],
    canEdit: false,
    authToken: ''
  });
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

function _collectScopeFromUserRow_(scopeRaw, category, managed) {
  var scopeJoined = []
    .concat(_parseScopeList_(scopeRaw))
    .concat(_parseScopeList_(category))
    .concat(_parseScopeList_(managed));
  var scopeSeen = {};
  var scope = [];

  for (var i = 0; i < scopeJoined.length; i++) {
    var raw = scopeJoined[i];
    var key = _normalizeCategoryKey_(raw);
    if (!key || scopeSeen[key]) continue;
    scopeSeen[key] = true;
    scope.push(raw);
  }

  return scope;
}

function _resolveSelectedAccess_(access, viewCode) {
  var selected = _applySelectedViewToAccess_(access, viewCode);
  if (!selected) {
    return _err('FORBIDDEN_VIEW', 'View khong hop le hoac khong nam trong pham vi duoc cap.');
  }
  return _ok(selected);
}

function loginWithEmail(email, password) {
  try {
    var access = _getUserAccessByEmail_(email, password, true);
    if (!access.ok) return access;

    var token = _createAuthToken_(access.data);
    var requiresViewSelection = access.data.allowedViews && access.data.allowedViews.length > 1 && !access.data.selectedView;
    var targetUrl = requiresViewSelection
      ? _buildAppUrl_('login', token, '')
      : _buildAppUrl_(access.data.selectedPage, token, access.data.selectedView);

    return _ok({
      email: access.data.email,
      fullName: access.data.fullName,
      role: access.data.role || access.data.baseRole || '',
      token: token,
      url: targetUrl,
      page: access.data.selectedPage || '',
      selectedView: access.data.selectedView || '',
      requiresViewSelection: requiresViewSelection,
      availableViews: access.data.allowedViews || [],
      defaultView: access.data.defaultView || ''
    });
  } catch (err) {
    return _err('LOGIN_ERROR', err.message, err.stack);
  }
}

function selectUserView(token, viewCode) {
  try {
    var tk = String(token || '').trim();
    var desiredView = _normalizeViewCode_(viewCode);
    if (!tk) return _err('MISSING_TOKEN', 'Vui long dang nhap de tiep tuc.');
    if (!desiredView) return _err('INVALID_VIEW', 'View khong hop le.');

    var accessRes = _getAccessByToken_(tk);
    if (!accessRes.ok) return accessRes;

    var selectedRes = _resolveSelectedAccess_(accessRes.data, desiredView);
    if (!selectedRes.ok) return selectedRes;

    _storeAccessByToken_(tk, selectedRes.data);

    return _ok({
      token: tk,
      url: _buildAppUrl_(selectedRes.data.selectedPage, tk, selectedRes.data.selectedView),
      page: selectedRes.data.selectedPage,
      selectedView: selectedRes.data.selectedView,
      selectedDataMode: selectedRes.data.selectedDataMode,
      selectedEditRule: selectedRes.data.selectedEditRule,
      currentViewCanEdit: selectedRes.data.currentViewCanEdit === true
    });
  } catch (err) {
    return _err('SELECT_VIEW_ERROR', err.message, err.stack);
  }
}

function getCurrentUserAccess(token) {
  return _getAccessByToken_(token);
}

function logoutSession(payload) {
  try {
    var data = payload || {};
    var token = String(data.token || data.authToken || '').trim();
    var loginUrl = _buildCleanLoginUrl_();
    if (!token) {
      return _ok({
        loggedOut: true,
        alreadyLoggedOut: true,
        loginUrl: loginUrl,
        message: 'No active session token was provided.'
      });
    }

    var accessRes = _getAccessByToken_(token);
    if (!accessRes.ok) {
      var code = accessRes.error && accessRes.error.code ? String(accessRes.error.code).toUpperCase() : '';
      if (code === 'MISSING_TOKEN' || code === 'TOKEN_EXPIRED' || code === 'INVALID_TOKEN' || code === 'TOKEN_READ_ERROR') {
        _removeAccessByToken_(token);
        return _ok({
          loggedOut: true,
          alreadyLoggedOut: true,
          loginUrl: loginUrl,
          message: accessRes.error && accessRes.error.message ? accessRes.error.message : 'Session was already invalid.'
        });
      }
      return accessRes;
    }

    _removeAccessByToken_(token);
    return _ok({
      loggedOut: true,
      email: accessRes.data && accessRes.data.email ? accessRes.data.email : '',
      loginUrl: loginUrl,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    return _err('LOGOUT_ERROR', _safeErrorText_(err, 'Logout failed'), err && err.stack ? err.stack : '');
  }
}

function _renderLoginPage(message, prefillEmail, bootstrap) {
  var tpl = HtmlService.createTemplateFromFile('Login');
  var initialBootstrap = bootstrap || {};
  tpl.loginMessage = message || initialBootstrap.message || '';
  tpl.prefillEmail = prefillEmail || initialBootstrap.email || '';
  tpl.loginBootstrapJson = _safeJsonForInlineScript_(initialBootstrap);
  var out = tpl.evaluate();
  out.setTitle('NPA Login');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

function _getUserAccessByEmail_(email, password, enforcePassword) {
  try {
    var em = _normalizeLoginEmail_(email);
    if (!em) return _err('MISSING_EMAIL', 'Khong tim thay email dang nhap.');
    var pwd = String(password || '');

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
    if (!sheet) return _err('USER_MASTER_NOT_FOUND', 'Chua co sheet User_Master.');

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return _err('USER_NOT_FOUND', 'User_Master chua co du lieu.');

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
    var hCanEdit = _findUserMasterHeaderIndex_(headers, ['Can Edit']);
    var hScope = _findUserMasterHeaderIndex_(headers, ['Effective Scope (Auto)', 'Effective Scope']);

    if (hEmail < 0 || hRole < 0 || hStatus < 0 || hCanEdit < 0) {
      return _err('USER_MASTER_SCHEMA_ERROR', 'Thieu cot bat buoc Email/Role/Status/Can Edit trong User_Master.');
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
      return _err('USER_NOT_FOUND', 'Khong tim thay tai khoan trong User_Master cho email: ' + em);
    }

    var status = String(found[hStatus] || '').trim().toUpperCase();
    if (status !== 'ACTIVE') return _err('USER_INACTIVE', 'Tai khoan dang bi khoa (INACTIVE).');

    if (enforcePassword) {
      if (hPassword < 0) return _err('MISSING_PASSWORD_COLUMN', 'Thieu cot Password trong User_Master.');
      if (!pwd) return _err('MISSING_PASSWORD', 'Vui long nhap mat khau.');
      var sheetPwd = String(found[hPassword] || '');
      if (pwd !== sheetPwd) return _err('INVALID_PASSWORD', 'Mat khau khong dung.');
    }

    var rawRole = String(found[hRole] || '').trim().toUpperCase();
    if (rawRole !== 'DIRECTOR' && rawRole !== 'MANAGER' && rawRole !== 'CATEGORY') {
      return _err('INVALID_ROLE', 'Role khong hop le: ' + rawRole);
    }

    var fullName = hName >= 0 ? String(found[hName] || '').trim() : '';
    var canEditRaw = String(found[hCanEdit] || '').trim();
    var scopeRaw = hScope >= 0 ? String(found[hScope] || '').trim() : '';
    var category = hCategory >= 0 ? String(found[hCategory] || '').trim() : '';
    var managed = hManaged >= 0 ? String(found[hManaged] || '').trim() : '';
    var scope = _collectScopeFromUserRow_(scopeRaw, category, managed);
    if (rawRole === 'DIRECTOR') scope = ['ALL'];

    var access = _buildAccessModel_({
      email: em,
      fullName: fullName,
      role: rawRole,
      status: status,
      canEditRaw: canEditRaw,
      categoryScope: scope
    });

    if (!access.allowedViews || !access.allowedViews.length) {
      return _err('INVALID_ACCESS', 'Can Edit rong hoac khong co token hop le. Su dung token: CAT_View/CAT_Edit, MAN_View/MAN_Edit, DIR_View/DIR_Edit.');
    }

    if (!access.selectedView && access.allowedViews.length === 1) {
      access = _applySelectedViewToAccess_(access, access.allowedViews[0].code) || access;
    }

    return _ok(access);
  } catch (err) {
    return _err('USER_ACCESS_ERROR', err.message, err.stack);
  }
}

function _createAuthToken_(accessData) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  _storeAccessByToken_(token, _normalizeAccessData_(accessData || {}));
  return token;
}

function _getAccessByToken_(token) {
  try {
    var tk = String(token || '').trim();
    if (!tk) return _err('MISSING_TOKEN', 'Phien dang nhap khong hop le. Vui long dang nhap lai.');

    var raw = CacheService.getScriptCache().get('AUTH_' + tk);
    if (!raw) return _err('TOKEN_EXPIRED', 'Phien dang nhap da het han. Vui long dang nhap lai.');

    var parsed = JSON.parse(raw);
    if (!parsed || !parsed.email) return _err('INVALID_TOKEN', 'Phien dang nhap khong hop le.');

    var normalized = _normalizeAccessData_(parsed);
    if (!normalized || !normalized.email) return _err('INVALID_TOKEN', 'Khong the phuc hoi thong tin truy cap.');

    _storeAccessByToken_(tk, normalized);
    return _ok(normalized);
  } catch (err) {
    return _err('TOKEN_READ_ERROR', err.message, err.stack);
  }
}

// ========================================
// ADMIN: Setup User_Master for Role-based Login
// Roles: DIRECTOR > MANAGER > CATEGORY
// ========================================
function _ensureUserMasterColumn_(sheet, headerName, previousHeaderName) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var idx = _findUserMasterHeaderIndex_(headers, [headerName]);
  if (idx >= 0) return idx;

  var insertAt = headers.length + 1;
  if (previousHeaderName) {
    var previousIdx = _findUserMasterHeaderIndex_(headers, [previousHeaderName]);
    if (previousIdx >= 0) insertAt = previousIdx + 2;
  }

  sheet.insertColumnBefore(insertAt);
  sheet.getRange(1, insertAt).setValue(headerName);

  headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), insertAt)).getDisplayValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  return _findUserMasterHeaderIndex_(headers, [headerName]);
}

function setupUserMasterSheet() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.USER_MASTER_SHEET);
    if (!sheet) sheet = ss.insertSheet(CONFIG.USER_MASTER_SHEET);

    if (sheet.getMaxRows() < 200) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 200 - sheet.getMaxRows());
    }

    var specs = [
      { header: 'Email', note: 'Login email. Must be unique.', width: 260 },
      { header: 'Password', note: 'Current login password.', width: 140 },
      { header: 'Full Name', note: 'Display name.', width: 180 },
      { header: 'Role', note: 'Legacy role label for identity and scope context.', width: 130 },
      { header: 'Status', note: 'ACTIVE or INACTIVE.', width: 120 },
      { header: 'Category', note: 'Category scope for Manager/Category views.', width: 360 },
      { header: 'Can Edit', note: 'Access tokens. Use CAT_View/CAT_Edit, MAN_View/MAN_Edit, DIR_View/DIR_Edit. Multiple tokens separated by comma/new line/semicolon.', width: 280 },
      { header: 'Notes', note: 'Internal note.', width: 280 },
      { header: 'Created At', note: 'Created timestamp.', width: 170 },
      { header: 'Updated At', note: 'Updated timestamp.', width: 170 }
    ];

    for (var s = 0; s < specs.length; s++) {
      _ensureUserMasterColumn_(sheet, specs[s].header, s > 0 ? specs[s - 1].header : '');
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
      return String(h || '').trim();
    });

    var permissionIdx = _findUserMasterHeaderIndex_(headers, ['Permission', 'Permissions']);
    while (permissionIdx >= 0) {
      sheet.deleteColumn(permissionIdx + 1);
      lastCol = sheet.getLastColumn();
      headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(h) {
        return String(h || '').trim();
      });
      permissionIdx = _findUserMasterHeaderIndex_(headers, ['Permission', 'Permissions']);
    }

    var headerNotes = sheet.getRange(1, 1, 1, lastCol).getNotes()[0];

    for (var i = 0; i < specs.length; i++) {
      var idx = _findUserMasterHeaderIndex_(headers, [specs[i].header]);
      if (idx < 0) continue;
      headers[idx] = specs[i].header;
      headerNotes[idx] = specs[i].note;
      sheet.setColumnWidth(idx + 1, specs[i].width);
    }

    sheet.getRange(1, 1, 1, lastCol).setValues([headers]);
    sheet.getRange(1, 1, 1, lastCol).setNotes([headerNotes]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#1f4e78')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    var dataRows = Math.max(1, sheet.getMaxRows() - 1);
    var idxRole = _findUserMasterHeaderIndex_(headers, ['Role']);
    var idxStatus = _findUserMasterHeaderIndex_(headers, ['Status']);
    var idxCategory = _findUserMasterHeaderIndex_(headers, ['Category']);
    var idxNotes = _findUserMasterHeaderIndex_(headers, ['Notes']);
    var idxCreated = _findUserMasterHeaderIndex_(headers, ['Created At']);
    var idxUpdated = _findUserMasterHeaderIndex_(headers, ['Updated At']);
    var idxCanEdit = _findUserMasterHeaderIndex_(headers, ['Can Edit']);

    if (idxRole >= 0) {
      var roleRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['DIRECTOR', 'MANAGER', 'CATEGORY'], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, idxRole + 1, dataRows, 1).setDataValidation(roleRule);
    }

    if (idxStatus >= 0) {
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['ACTIVE', 'INACTIVE'], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, idxStatus + 1, dataRows, 1).setDataValidation(statusRule);
    }

    if (idxCanEdit >= 0) {
      sheet.getRange(2, idxCanEdit + 1, dataRows, 1).clearDataValidations();
      sheet.getRange(2, idxCanEdit + 1, dataRows, 1).setWrap(true);
    }

    var categories = _extractCategoriesFromMainSheet_();
    if (idxCategory >= 0) {
      if (categories.length > 0) {
        var catRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(categories, true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, idxCategory + 1, dataRows, 1).setDataValidation(catRule);
      } else {
        sheet.getRange(2, idxCategory + 1, dataRows, 1).clearDataValidations();
      }
      sheet.getRange(2, idxCategory + 1, dataRows, 1).setWrap(true);
    }

    sheet.getRange(2, 1, dataRows, lastCol).setVerticalAlignment('middle');
    if (idxNotes >= 0) sheet.getRange(2, idxNotes + 1, dataRows, 1).setWrap(true);
    if (idxCreated >= 0) sheet.getRange(2, idxCreated + 1, dataRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    if (idxUpdated >= 0) sheet.getRange(2, idxUpdated + 1, dataRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');

    var hasAnyUser = false;
    var idxEmail = _findUserMasterHeaderIndex_(headers, ['Email']);
    if (idxEmail >= 0 && sheet.getLastRow() >= 2) {
      var emailVals = sheet.getRange(2, idxEmail + 1, sheet.getLastRow() - 1, 1).getValues();
      hasAnyUser = emailVals.some(function(r) {
        return String(r[0] || '').trim() !== '';
      });
    }

    if (!hasAnyUser) {
      var now = new Date();
      var c1 = categories[0] || 'Non Alcoholic';
      var seedSpecs = [
        {
          email: 'director@yourcompany.com',
          password: '7-Eleven',
          fullName: 'Director User',
          role: 'DIRECTOR',
          status: 'ACTIVE',
          category: '',
          canEdit: 'DIR_Edit',
          notes: 'Director workspace'
        },
        {
          email: 'manager@yourcompany.com',
          password: '7-Eleven',
          fullName: 'Manager User',
          role: 'MANAGER',
          status: 'ACTIVE',
          category: c1,
          canEdit: 'MAN_Edit',
          notes: 'Manager workspace'
        },
        {
          email: 'category@yourcompany.com',
          password: '7-Eleven',
          fullName: 'Category User',
          role: 'CATEGORY',
          status: 'ACTIVE',
          category: c1,
          canEdit: 'CAT_Edit',
          notes: 'Category workspace'
        }
      ];

      var seedRows = seedSpecs.map(function(item) {
        var row = [];
        for (var col = 0; col < lastCol; col++) row.push('');
        if (idxEmail >= 0) row[idxEmail] = item.email;
        var idxPassword = _findUserMasterHeaderIndex_(headers, ['Password']);
        var idxFullName = _findUserMasterHeaderIndex_(headers, ['Full Name']);
        if (idxPassword >= 0) row[idxPassword] = item.password;
        if (idxFullName >= 0) row[idxFullName] = item.fullName;
        if (idxRole >= 0) row[idxRole] = item.role;
        if (idxStatus >= 0) row[idxStatus] = item.status;
        if (idxCategory >= 0) row[idxCategory] = item.category;
        if (idxCanEdit >= 0) row[idxCanEdit] = item.canEdit;
        if (idxNotes >= 0) row[idxNotes] = item.notes;
        if (idxCreated >= 0) row[idxCreated] = now;
        if (idxUpdated >= 0) row[idxUpdated] = now;
        return row;
      });

      sheet.getRange(2, 1, seedRows.length, lastCol).setValues(seedRows);
    }

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, sheet.getMaxRows(), lastCol).createFilter();
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
