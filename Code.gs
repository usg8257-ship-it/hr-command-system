// ============================================================
//  HR COMMAND SYSTEM v2  —  Core / Entry Point
//  Branch: v2-rebuild
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

var TABS = {
  MASTER:       'Master_Data',
  INACTIVE:     'Inactive_Data',
  PREONBOARD:   'PreOnboarding',
  TRACKER:      'Strategy_Tracker',
  STEPS:        'Steps_Config',
  HR_DOCS:      'HR_Docs',
  USERS:        'Users',
  CONFIG:       'Config',
  AUDIT:        'Audit_Log',
  NOTIFS:       'Notifications',
  SUMMARY:      'Summary',
  JOBS:         'Jobs',
  APPLICATIONS: 'Applications',
  LEAVE:        'Leave'
};

// ── doGet ────────────────────────────────────────────────────
function doGet(e) {
  var view = e && e.parameter && e.parameter.view;
  if (view === 'jobs') {
    return HtmlService.createTemplateFromFile('jobs-portal')
      .evaluate()
      .setTitle('Careers')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('HR Command System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── getHeaderMap ─────────────────────────────────────────────
/**
 * Returns { map: {colName: colIndex (1-based)}, sheet, lastRow }
 * NEVER use hardcoded column indices — always call this.
 */
function getHeaderMap(sheetName) {
  var sh = SS.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return { map: {}, sheet: sh, lastRow: sh.getLastRow() };
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) {
    if (h !== '' && h !== null && h !== undefined) {
      map[String(h).trim()] = i + 1;
    }
  });
  return { map: map, sheet: sh, lastRow: sh.getLastRow() };
}

// ── getOrCreate ──────────────────────────────────────────────
function getOrCreate(name, headers) {
  var sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }
  var existing = sh.getLastColumn() > 0
    ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    : [];
  var existingSet = {};
  existing.forEach(function(h) { existingSet[String(h).trim()] = true; });
  var missing = headers.filter(function(h) { return !existingSet[h]; });
  if (missing.length) {
    var startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
  return sh;
}

// ── _buildRow ────────────────────────────────────────────────
function _buildRow(hm, data) {
  var len = Object.keys(hm.map).length;
  var row = new Array(len).fill('');
  Object.keys(data).forEach(function(key) {
    var col = hm.map[key];
    if (col) row[col - 1] = (data[key] !== undefined && data[key] !== null) ? data[key] : '';
  });
  return row;
}

// ── Date utilities ───────────────────────────────────────────
function formatDate(d) {
  if (!d) return '';
  var dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  var dd = String(dt.getDate()).padStart(2, '0');
  var mm = String(dt.getMonth() + 1).padStart(2, '0');
  var yyyy = dt.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

function _parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var parts = String(v).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (parts) return new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]));
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function addDays(dateVal, days) {
  var d = _parseDate(dateVal);
  if (!d) return '';
  d.setDate(d.getDate() + days);
  return formatDate(d);
}

// ── Text utilities ───────────────────────────────────────────
function toProperCase(text) {
  if (!text) return '';
  return String(text).toLowerCase().replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function genId(prefix) {
  var ts = new Date().getTime().toString(36).toUpperCase();
  var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return (prefix || 'ID') + '-' + ts + rand;
}

// ── Config helpers ───────────────────────────────────────────
function getConfig() {
  var sh = SS.getSheetByName(TABS.CONFIG);
  if (!sh || sh.getLastRow() < 2) return {};
  var hm = getHeaderMap(TABS.CONFIG);
  var lastRow = hm.lastRow;
  var cfg = {};
  var rows = sh.getRange(2, 1, lastRow - 1, Math.max(sh.getLastColumn(), 2)).getValues();
  rows.forEach(function(r) {
    var k = String(r[0]).trim();
    if (k) cfg[k] = r[1];
  });
  return cfg;
}

// Upsert a single config key — NEVER clears the sheet
function setConfig(key, value, userEmail) {
  var hm = getHeaderMap(TABS.CONFIG);
  var sh = hm.sheet;
  var lastRow = hm.lastRow;
  var keyCol = hm.map['key'] || 1;
  var valCol = hm.map['value'] || 2;
  if (lastRow >= 2) {
    var keys = sh.getRange(2, keyCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        var rowNum = i + 2;
        sh.getRange(rowNum, valCol).setValue(value);
        if (hm.map['updated_at']) sh.getRange(rowNum, hm.map['updated_at']).setValue(new Date());
        if (hm.map['updated_by'] && userEmail) sh.getRange(rowNum, hm.map['updated_by']).setValue(userEmail);
        return;
      }
    }
  }
  sh.appendRow(_buildRow(hm, { key: key, value: value, updated_at: new Date(), updated_by: userEmail || '' }));
}

// Save multiple config key/value pairs (upsert each)
function saveConfig(data, token) {
  _setSessionFromToken(token);
  var user = _SESSION_EMAIL;
  var keys = Object.keys(data);
  for (var i = 0; i < keys.length; i++) {
    setConfig(keys[i], data[keys[i]], user);
  }
  return { success: true };
}

// ── Init all sheets ──────────────────────────────────────────
function initSheets() {
  getOrCreate(TABS.MASTER, [
    'ID','NAME','GENDER','DESIGNATION','DOB','DOJ',
    'probation_end_date','main_dept','nationality',
    'personal_e_mail','STATUS','20DS'
  ]);
  getOrCreate(TABS.INACTIVE, [
    'ID','NAME','termination_date','DOJ','employee_status'
  ]);
  getOrCreate(TABS.PREONBOARD, [
    'ob_id','NAME','GENDER','DESIGNATION','main_dept','nationality',
    'personal_e_mail','ENTITY','DOB','expected_doj','mobile',
    'STATUS','ASSIGNED_TO','created_at','notes'
  ]);
  getOrCreate(TABS.TRACKER, _trackerHeaders());
  getOrCreate(TABS.STEPS, ['step_num','step_name','sla_days','responsible_role','is_active']);
  getOrCreate(TABS.HR_DOCS, [
    'doc_id','emp_id','emp_name','letter_type','sub_type',
    'entity','issued_by','issued_at','drive_file_id','notes'
  ]);
  getOrCreate(TABS.USERS, [
    'email','display_name','password_hash','role','entities',
    'is_active','email_notifications','last_login','created_at'
  ]);
  getOrCreate(TABS.CONFIG, ['key','value','updated_at','updated_by']);
  getOrCreate(TABS.AUDIT, [
    'timestamp','user_email','module','action',
    'record_id','field','old_value','new_value'
  ]);
  getOrCreate(TABS.NOTIFS, [
    'notif_id','recipient_email','message',
    'link_module','is_read','created_at','created_by'
  ]);
  getOrCreate(TABS.SUMMARY, ['metric_key','metric_value','updated_at']);
  getOrCreate(TABS.JOBS, [
    'job_id','title','entity','department','description',
    'salary_range','status','created_by','created_at'
  ]);
  getOrCreate(TABS.APPLICATIONS, [
    'app_id','job_id','applicant_name','email','mobile',
    'cv_drive_id','stage','notes','applied_at','updated_at'
  ]);
  getOrCreate(TABS.LEAVE, [
    'leave_id','emp_id','emp_name','leave_type',
    'start_date','end_date','days_count','status',
    'applied_by','reviewed_by','reviewed_at','notes'
  ]);
  _seedDefaultSteps();
  return { success: true, message: 'All sheets initialised.' };
}

function _trackerHeaders() {
  var base = ['emp_id','emp_name','designation','doj','STATUS','initiated_at','completed_at','approved_by'];
  for (var i = 1; i <= 12; i++) {
    base.push('Step'+i+'_Done','Step'+i+'_By','Step'+i+'_At','Step'+i+'_Notes');
  }
  return base;
}

function _seedDefaultSteps() {
  var sh = SS.getSheetByName(TABS.STEPS);
  if (!sh || sh.getLastRow() > 1) return;
  var defaults = [
    [1,  'Documents Collected',        2, 'HR_OFFICER', true],
    [2,  'Visa Application Submitted', 3, 'HR_OFFICER', true],
    [3,  'Emirates ID Applied',        5, 'HR_OFFICER', true],
    [4,  'SIRA/PSBD License Applied',  7, 'HR_OFFICER', true],
    [5,  'Medical Fitness Test',       2, 'HR_OFFICER', true],
    [6,  'Uniform + Kit Issued',       1, 'HR_OFFICER', true],
    [7,  'Bank Account Opened',        3, 'HR_OFFICER', true],
    [8,  'WPS Enrolled',               2, 'HR_OFFICER', true],
    [9,  'System Access Granted',      1, 'HR_OFFICER', true],
    [10, 'HR Manager Final Sign-off',  1, 'HR_MANAGER', true]
  ];
  sh.getRange(2, 1, defaults.length, 5).setValues(defaults);
}
