// ============================================================
// UNITED GROUP HOLDING — HR COMMAND SYSTEM v4
// Code.gs — Google Apps Script Backend
// Concept, Designed & Developed by Mohammad Sanish
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

var TABS = {
  MASTER:       'Master Data',
  ONBOARDING:   'Onboarding',
  HR_DOCS:      'HR Docs Tracker',
  DEL_LOG:      'Deletion_Log',
  ACTIVE_MP:    'Active_Manpower',
  SUMMARY:      'Summary',
  CONFIG:       'AppConfig',
  ACTIVITY:     'ActivityLog',
  USERS:        'Users',
  LEAVE:        'Leave',
  JOBS:         'Jobs',
  APPLICATIONS: 'Applications',
  DS_TRACKER:   '20 Days Strategy Tracker',
  DS_AUDIT:     'Onboarding Audit Log',
  STEPS:        'Steps Config'
};

// ============================================================
// WEB APP ENTRY
// ============================================================
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  if (view === 'jobs') {
    return HtmlService
      .createTemplateFromFile('jobs-portal')
      .evaluate()
      .setTitle('Careers — United Group Holding')
      .addMetaTag('viewport','width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('United Group Holding — HR System')
    .addMetaTag('viewport','width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// AUTH — PROFILE, ROLE GUARD, ENTITY FILTER
// ============================================================

// Per-request session email (set by _setSessionFromToken at start of each call)
var _SESSION_EMAIL = '';

function _setSessionFromToken(token) {
  if (!token) return '';
  var email = CacheService.getScriptCache().get('ughr_sess_' + token);
  if (email) _SESSION_EMAIL = String(email).toLowerCase().trim();
  return _SESSION_EMAIL;
}

function getMyProfile() {
  // Token-based auth takes priority (set by _setSessionFromToken)
  var email = _SESSION_EMAIL;
  // Fall back to Google session auth
  if (!email) {
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e2) {}
  }
  if (!email) return {email:'',name:'',role:'NONE',entities:[],active:false};

  // Cache per-email for 60 seconds to reduce sheet reads
  var cache = CacheService.getScriptCache();
  var cacheKey = 'profile_' + email;
  var cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch(e2) {} }

  // Ensure Users sheet exists (with PASSWORD column)
  var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE','PASSWORD']);
  var vals = sh.getDataRange().getValues();

  // Bootstrap: if only the header row exists, make first user SUPER_ADMIN
  if (vals.length === 1) {
    var lock = LockService.getScriptLock();
    lock.tryLock(3000);
    try {
      var vals2 = sh.getDataRange().getValues();
      if (vals2.length === 1) {
        sh.appendRow([email, email.split('@')[0], 'SUPER_ADMIN', 'ALL', 'TRUE', 'Admin@1234']);
        vals = sh.getDataRange().getValues();
      } else {
        vals = vals2;
      }
    } finally { lock.releaseLock(); }
  }

  var hdrs = vals[0].map(function(h){ return String(h).trim(); });
  var emailIdx = hdrs.indexOf('EMAIL');
  for (var i = 1; i < vals.length; i++) {
    var rowEmail = String(vals[i][emailIdx]||'').toLowerCase().trim();
    if (rowEmail === email) {
      var role     = String(vals[i][hdrs.indexOf('ROLE')]||'').trim();
      var entRaw   = String(vals[i][hdrs.indexOf('ENTITIES')]||'').trim();
      var active   = String(vals[i][hdrs.indexOf('ACTIVE')]||'').toUpperCase();
      var name     = String(vals[i][hdrs.indexOf('DISPLAY_NAME')]||'').trim();
      if (active === 'FALSE' || active === 'NO') {
        return {email:email,name:name,role:'NONE',entities:[],active:false};
      }
      var entities = entRaw.toUpperCase() === 'ALL' ? 'ALL' : entRaw.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
      var mcIdx = hdrs.indexOf('MUST_CHANGE');
      var mustChange = mcIdx >= 0 && String(vals[i][mcIdx]||'').toUpperCase() === 'TRUE';
      var profile = {email:email, name:name||email.split('@')[0], role:role, entities:entities, active:true, mustChange:mustChange};
      cache.put(cacheKey, JSON.stringify(profile), 60);
      return profile;
    }
  }
  return {email:email,name:'',role:'NONE',entities:[],active:false};
}

// ── Public login functions ──────────────────────────────────

// PUBLIC — validates email+password against Users sheet, returns session token
function loginUser(email, password) {
  try {
    if (!email || !password) return {success:false, error:'Email and password are required'};
    var normEmail = String(email).toLowerCase().trim();
    var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE','PASSWORD']);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:false, error:'No users configured. Contact administrator.'};
    var hdrs  = vals[0].map(function(h){ return String(h).trim(); });
    var eIdx  = hdrs.indexOf('EMAIL');
    var pIdx  = hdrs.indexOf('PASSWORD');
    var nIdx  = hdrs.indexOf('DISPLAY_NAME');
    var rIdx  = hdrs.indexOf('ROLE');
    var entIdx= hdrs.indexOf('ENTITIES');
    var aIdx  = hdrs.indexOf('ACTIVE');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][eIdx]||'').toLowerCase().trim() !== normEmail) continue;
      var active = String(vals[i][aIdx]||'').toUpperCase();
      if (active !== 'TRUE' && active !== 'YES') {
        return {success:false, error:'Account is disabled. Contact administrator.'};
      }
      var stored = String(vals[i][pIdx]||'');
      if (!stored) return {success:false, error:'Password not set for this account. Contact administrator.'};
      if (stored !== String(password)) return {success:false, error:'Invalid email or password'};
      // Issue 8-hour session token
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('ughr_sess_' + token, normEmail, 28800);
      var entRaw   = String(vals[i][entIdx]||'ALL').trim();
      var entities = entRaw.toUpperCase() === 'ALL' ? 'ALL' : entRaw.split(',').map(function(e){ return e.trim(); }).filter(Boolean);
      var mcIdx2 = hdrs.indexOf('MUST_CHANGE');
      var mustChange = mcIdx2 >= 0 && String(vals[i][mcIdx2]||'').toUpperCase() === 'TRUE';
      var profile  = {
        email:      normEmail,
        name:       String(vals[i][nIdx]||'').trim() || normEmail.split('@')[0],
        role:       String(vals[i][rIdx]||'STAFF').trim(),
        entities:   entities,
        active:     true,
        mustChange: mustChange
      };
      logActivity('AuthAgent', 'LOGIN', normEmail, 'SUCCESS');
      return {success:true, token:token, profile:profile};
    }
    return {success:false, error:'Invalid email or password'};
  } catch(e) { return {success:false, error:e.message}; }
}

// PUBLIC — validates an existing session token and refreshes its TTL
function validateSession(token) {
  try {
    if (!token) return {success:false};
    var email = CacheService.getScriptCache().get('ughr_sess_' + token);
    if (!email) return {success:false, error:'SESSION_EXPIRED'};
    _SESSION_EMAIL = String(email).toLowerCase().trim();
    var profile = getMyProfile();
    if (!profile || profile.role === 'NONE') {
      CacheService.getScriptCache().remove('ughr_sess_' + token);
      return {success:false, error:'Account access revoked'};
    }
    // Refresh TTL
    CacheService.getScriptCache().put('ughr_sess_' + token, email, 28800);
    return {success:true, profile:profile};
  } catch(e) { return {success:false, error:e.message}; }
}

// PUBLIC — invalidates a session token
function logoutUser(token) {
  try {
    if (token) {
      CacheService.getScriptCache().remove('ughr_sess_' + token);
      logActivity('AuthAgent', 'LOGOUT', String(token).slice(0,8)+'…', 'SUCCESS');
    }
    return {success:true};
  } catch(e) { return {success:true}; }
}

// Protected dispatcher — all non-public GAS calls route here with the session token
function runProtected(token, funcName, args) {
  try {
    if (!token) return {success:false, error:'SESSION_EXPIRED'};
    var email = _setSessionFromToken(token);
    if (!email) return {success:false, error:'SESSION_EXPIRED'};
    var a = args || [];
    var fns = {
      // Employee
      addEmployee:                    function(){ return addEmployee(a[0]); },
      updateEmployee:                 function(){ return updateEmployee(a[0]); },
      deleteEmployee:                 function(){ return deleteEmployee(a[0], a[1]); },
      // Onboarding
      addOnboarding:                  function(){ return addOnboarding(a[0]); },
      updateOnboarding:               function(){ return updateOnboarding(a[0]); },
      deleteOnboarding:               function(){ return deleteOnboarding(a[0]); },
      transferToMaster:               function(){ return transferToMaster(a[0], a[1]); },
      // HR Docs / Letters
      issueHRDoc:                     function(){ return issueHRDoc(a[0]); },
      generateAndIssueLetter:         function(){ return generateAndIssueLetter(a[0]); },
      generateExperienceLetterForEmp: function(){ return generateExperienceLetterForEmp(a[0]); },
      getOfferLetterTemplate:         function(){ return getOfferLetterTemplate(a[0]); },
      saveTemplateFileIds:            function(){ return saveTemplateFileIds(a[0], a[1]); },
      // Config
      saveConfig:                     function(){ return saveConfig(a[0]); },
      getActivityLog:                 function(){ return getActivityLog(); },
      // Step Config
      getStepConfig:                  function(){ return getStepConfig(); },
      saveStepConfig:                 function(){ return saveStepConfig(a[0]); },
      mergeSteps:                     function(){ return mergeSteps(a[0], a[1], a[2]); },
      // Users
      getUsers:                       function(){ return getUsers(); },
      saveUser:                       function(){ return saveUser(a[0]); },
      deleteUser:                     function(){ return deleteUser(a[0]); },
      toggleUserActive:               function(){ return toggleUserActive(a[0], a[1]); },
      resetUserPassword:              function(){ return resetUserPassword(a[0]); },
      changeMyPassword:               function(){ return changeMyPassword(a[0], a[1]); },
      updateMyProfile:                function(){ return updateMyProfile(a[0]); },
      // Self-service
      getMyEmployeeRecord:            function(){ return getMyEmployeeRecord(); },
      updateMyExpiryDates:            function(){ return updateMyExpiryDates(a[0]); },
      // Leave
      getLeave:                       function(){ return getLeave(a[0]); },
      addLeave:                       function(){ return addLeave(a[0]); },
      updateLeaveStatus:              function(){ return updateLeaveStatus(a[0], a[1], a[2]); },
      // Recruitment
      getJobs:                        function(){ return getJobs(); },
      saveJob:                        function(){ return saveJob(a[0]); },
      closeJob:                       function(){ return closeJob(a[0]); },
      getApplications:                function(){ return getApplications(a[0]); },
      updateApplicationStage:         function(){ return updateApplicationStage(a[0], a[1], a[2]); },
      transferAppToOnboarding:        function(){ return transferAppToOnboarding(a[0]); },
      // Resignations
      getResignations:                function(){ return getResignations(); },
      // 20DS Tracker
      get20DSTracker:                 function(){ return get20DSTracker(); },
      update20DSStep:                 function(){ return update20DSStep(a[0], a[1], a[2]); },
      completeDSStep:                 function(){ return completeDSStep(a[0], a[1], a[2], a[3]); },
      recalculate20DSTotals:          function(){ return recalculate20DSTotals(); },
      update20DSResponsible:          function(){ return update20DSResponsible(a[0], a[1]); },
      cancel20DSRecord:               function(){ return cancel20DSRecord(a[0], a[1]); },
      get20DSAuditLog:                function(){ return get20DSAuditLog(); },
      get20DSAnalytics:               function(){ return get20DSAnalytics(); },
    };
    if (!fns[funcName]) return {success:false, error:'Unknown function: ' + funcName};
    return fns[funcName]();
  } catch(e) { return {success:false, error:e.message}; }
}

function _requireRole(allowedRoles) {
  var profile = getMyProfile();
  if (!profile || allowedRoles.indexOf(profile.role) < 0) {
    throw new Error('ACCESS_DENIED: Role ' + (profile ? profile.role : 'NONE') + ' not permitted');
  }
  return profile;
}

function _filterByEntity(rows, profile, entityField) {
  var field = entityField || 'ENTITY';
  if (!profile || profile.entities === 'ALL') return rows;
  var allowed = profile.entities;
  return rows.filter(function(r) {
    return allowed.indexOf(normaliseEntity(r[field])) >= 0;
  });
}

// Server-side ID generator (mirrors frontend genId)
function genId_(prefix) {
  return (prefix||'ID') + '-' + new Date().getTime().toString().slice(-8);
}

// ============================================================
// LOAD ALL DATA
// ============================================================
function loadAllData(token) {
  if (token) _setSessionFromToken(token);
  var profile = getMyProfile();
  return {
    user:       profile,
    master:     getMasterData(profile),
    deletions:  getDeletionLog(),
    onboarding: getOnboarding(profile),
    hrDocs:     getHRDocs(profile),
    summary:    getSummary(),
    config:     getConfig(),
    dsTracker:  get20DSTracker(profile),
    stepConfig: getStepConfig()
  };
}

// ============================================================
// CONFIG
// ============================================================
function getConfig() {
  try {
    var sh = getOrCreate(TABS.CONFIG, ['KEY','VALUE']);
    var d = sh.getDataRange().getValues();
    var cfg = {};
    for (var i = 1; i < d.length; i++) { if(d[i][0]) cfg[String(d[i][0])] = d[i][1]; }
    return cfg;
  } catch(e) { return {}; }
}
function saveConfig(cfg) {
  var sh = getOrCreate(TABS.CONFIG, ['KEY','VALUE']);
  sh.clearContents(); sh.appendRow(['KEY','VALUE']);
  Object.keys(cfg).forEach(function(k){ sh.appendRow([k, cfg[k]]); });
  return {success:true};
}

// ============================================================
// STEP CONFIG
// ============================================================

// Hardcoded defaults — used as seed when Steps Config sheet is empty
var _DEFAULT_STEPS = [
  { STEP_KEY:'STEP_VISA',      LABEL:'Visa/Entry Issued',        SHORT:'VI', SLA_HOURS:24,  STATUSES:'Pending,Done,Problem',        SUBSTEPS:'{"mol_done":"MOL Done","entry_permit":"Entry Permit Submitted"}',                                                                                      ORDER:1, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_LABOR',     LABEL:'Tawjeeh & Labor Card',     SHORT:'LC', SLA_HOURS:24,  STATUSES:'Pending,Done',                SUBSTEPS:'{"tawjeeh":"Tawjeeh Completed","labor_card":"Labor Card Received"}',                                                                                    ORDER:2, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_MEDICAL',   LABEL:'Visa Medical',             SHORT:'MD', SLA_HOURS:24,  STATUSES:'Pending,Fit,Unfit',           SUBSTEPS:'{"medical_done":"Medical Done"}',                                                                                                                        ORDER:3, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_INSURANCE', LABEL:'Medical Insurance',        SHORT:'MI', SLA_HOURS:72,  STATUSES:'Pending,Done,Problem',        SUBSTEPS:'{"app_submitted":"Application Submitted","card_received":"Card Received"}',                                                                              ORDER:4, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_NSI',       LABEL:'NSI Training',             SHORT:'NS', SLA_HOURS:168, STATUSES:'Not Started,Pending,Scheduled,Done', SUBSTEPS:'{"app_submitted":"Application Submitted","date_scheduled":"Training Date Scheduled","training_completed":"Training Completed","cert_received":"Certificate Received"}', ORDER:5, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_EID',       LABEL:'EID & Residency Stamping', SHORT:'EI', SLA_HOURS:48,  STATUSES:'Pending,Done,Problem',        SUBSTEPS:'{"eid_app_submitted":"EID Application Submitted","biometrics":"Biometrics / Stamping Done","card_received":"Card Received"}',                           ORDER:6, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_ASSD',      LABEL:'ASSD',                     SHORT:'AS', SLA_HOURS:48,  STATUSES:'Locked,Pending,Done,Problem',  SUBSTEPS:'{}', ORDER:7, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_SIRA_CERT', LABEL:'SIRA Certificate',         SHORT:'SC', SLA_HOURS:72,  STATUSES:'Locked,Pending,Done,Problem',  SUBSTEPS:'{}', ORDER:8, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_PCC',       LABEL:'Police Clearance (PCC)',   SHORT:'PC', SLA_HOURS:72,  STATUSES:'Locked,Pending,Done,Problem',  SUBSTEPS:'{}', ORDER:9, ACTIVE:'TRUE', MERGED_INTO:'' },
  { STEP_KEY:'STEP_SIRA_LIC',  LABEL:'SIRA License',             SHORT:'SL', SLA_HOURS:72,  STATUSES:'Locked,Pending,Done,Problem',  SUBSTEPS:'{}', ORDER:10,ACTIVE:'TRUE', MERGED_INTO:'' }
];

var _STEPS_HDRS = ['STEP_KEY','LABEL','SHORT','SLA_HOURS','STATUSES','SUBSTEPS','ORDER','ACTIVE','MERGED_INTO'];

function getStepConfig() {
  try {
    var sh = getOrCreate(TABS.STEPS, _STEPS_HDRS);
    var vals = sh.getDataRange().getValues();
    // If only header row exists, seed with defaults
    if (vals.length < 2) {
      _DEFAULT_STEPS.forEach(function(s){
        sh.appendRow(_STEPS_HDRS.map(function(h){ return s[h] !== undefined ? s[h] : ''; }));
      });
      vals = sh.getDataRange().getValues();
    }
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var steps = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {};
      hdrs.forEach(function(h, j){ row[h] = String(vals[i][j]||''); });
      if (row.ACTIVE === 'FALSE') continue; // skip inactive/merged steps
      var substepsObj = {};
      try { substepsObj = JSON.parse(row.SUBSTEPS || '{}'); } catch(e2) {}
      steps.push({
        key:       row.STEP_KEY,
        label:     row.LABEL,
        short:     row.SHORT,
        slaHours:  Number(row.SLA_HOURS) || 24,
        statuses:  (row.STATUSES || 'Pending,Done').split(',').map(function(s){ return s.trim(); }),
        substeps:  substepsObj,
        order:     Number(row.ORDER) || 99,
        mergedInto: row.MERGED_INTO || ''
      });
    }
    // Ensure every default step is present; append any that were added after the sheet was first seeded
    var existingKeys = {};
    steps.forEach(function(s){ existingKeys[s.key] = true; });
    _DEFAULT_STEPS.forEach(function(def){
      if (existingKeys[def.STEP_KEY]) return;
      sh.appendRow(_STEPS_HDRS.map(function(h){ return def[h] !== undefined ? def[h] : ''; }));
      var substepsObj2 = {};
      try { substepsObj2 = JSON.parse(def.SUBSTEPS || '{}'); } catch(e2){}
      steps.push({
        key:      def.STEP_KEY,  label:    def.LABEL,  short:    def.SHORT,
        slaHours: def.SLA_HOURS, substeps: substepsObj2,
        statuses: (def.STATUSES||'Pending,Done').split(',').map(function(s){ return s.trim(); }),
        order:    def.ORDER || 99, mergedInto: ''
      });
    });
    steps.sort(function(a,b){ return a.order - b.order; });
    return {success:true, data:steps};
  } catch(e) { return {success:false, error:e.message}; }
}

function saveStepConfig(stepsArray) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    if (!Array.isArray(stepsArray) || !stepsArray.length) return {success:false, error:'No steps provided'};
    // Validate
    var seenKeys = {};
    for (var i = 0; i < stepsArray.length; i++) {
      var s = stepsArray[i];
      if (!s.key || !s.label || !s.short) return {success:false, error:'Each step requires key, label, and short code'};
      if (!/^STEP_[A-Z0-9_]+$/.test(s.key)) return {success:false, error:'Invalid step key: '+s.key+'. Must start with STEP_ and use uppercase letters/numbers/underscores'};
      if (seenKeys[s.key]) return {success:false, error:'Duplicate step key: '+s.key};
      seenKeys[s.key] = true;
    }
    var sh = getOrCreate(TABS.STEPS, _STEPS_HDRS);
    // Preserve existing inactive/merged rows
    var existing = sh.getDataRange().getValues();
    var existHdrs = existing.length > 0 ? existing[0].map(function(h){ return String(h).trim(); }) : _STEPS_HDRS;
    var activeIdx = existHdrs.indexOf('ACTIVE');
    var inactiveRows = [];
    for (var j = 1; j < existing.length; j++) {
      var active = activeIdx >= 0 ? String(existing[j][activeIdx]).toUpperCase() : 'TRUE';
      if (active === 'FALSE') inactiveRows.push(existing[j]);
    }
    sh.clearContents();
    sh.appendRow(_STEPS_HDRS);
    stepsArray.forEach(function(s, idx){
      var substepsStr = typeof s.substeps === 'object' ? JSON.stringify(s.substeps) : (s.substeps || '{}');
      var statusesStr = Array.isArray(s.statuses) ? s.statuses.join(',') : (s.statuses || 'Pending,Done');
      sh.appendRow([s.key, s.label, s.short, Number(s.slaHours)||24, statusesStr, substepsStr, idx+1, 'TRUE', s.mergedInto||'']);
    });
    inactiveRows.forEach(function(r){ sh.appendRow(r); });
    logActivity('StepConfig','SAVE','','SUCCESS - '+stepsArray.length+' steps');
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; }
}

function mergeSteps(keyA, keyB, newStep) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    if (!keyA || !keyB || !newStep || !newStep.key) return {success:false, error:'keyA, keyB and newStep.key are required'};
    if (!/^STEP_[A-Z0-9_]+$/.test(newStep.key)) return {success:false, error:'Invalid new step key format'};
    if (keyA === keyB) return {success:false, error:'Cannot merge a step with itself'};

    // 1. Update Steps Config sheet
    var sh = getOrCreate(TABS.STEPS, _STEPS_HDRS);
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var keyIdx   = hdrs.indexOf('STEP_KEY');
    var activeIdx= hdrs.indexOf('ACTIVE');
    var mergedIdx= hdrs.indexOf('MERGED_INTO');
    var orderIdx = hdrs.indexOf('ORDER');
    var maxOrder = 0;
    var rowA = -1, rowB = -1;
    for (var i = 1; i < vals.length; i++) {
      var k = String(vals[i][keyIdx]||'');
      if (k === keyA) rowA = i;
      if (k === keyB) rowB = i;
      var ord = Number(vals[i][orderIdx]||0);
      if (ord > maxOrder) maxOrder = ord;
    }
    if (rowA < 0) return {success:false, error:'Step '+keyA+' not found'};
    if (rowB < 0) return {success:false, error:'Step '+keyB+' not found'};

    // Mark A and B as inactive, merged into new key
    sh.getRange(rowA+1, activeIdx+1).setValue('FALSE');
    sh.getRange(rowA+1, mergedIdx+1).setValue(newStep.key);
    sh.getRange(rowB+1, activeIdx+1).setValue('FALSE');
    sh.getRange(rowB+1, mergedIdx+1).setValue(newStep.key);

    // Add the new merged step
    var substepsStr = typeof newStep.substeps === 'object' ? JSON.stringify(newStep.substeps) : (newStep.substeps||'{}');
    var statusesStr = Array.isArray(newStep.statuses) ? newStep.statuses.join(',') : (newStep.statuses||'Pending,Done');
    sh.appendRow([newStep.key, newStep.label, newStep.short, Number(newStep.slaHours)||24, statusesStr, substepsStr, maxOrder+1, 'TRUE', '']);

    // 2. Migrate tracker data — combine A+B into new key on each record
    var tsh = SS.getSheetByName(TABS.DS_TRACKER);
    var migratedCount = 0;
    if (tsh) {
      var tVals = tsh.getDataRange().getValues();
      var tHdrs = tVals[0].map(function(h){ return String(h).trim(); });
      var colA   = tHdrs.indexOf(keyA);
      var colB   = tHdrs.indexOf(keyB);
      // Add new key column if not present
      var colNew = tHdrs.indexOf(newStep.key);
      if (colNew < 0) {
        tsh.getRange(1, tHdrs.length+1).setValue(newStep.key);
        colNew = tHdrs.length;
        tHdrs.push(newStep.key);
      }
      for (var r = 1; r < tVals.length; r++) {
        var dataA = {}; var dataB = {};
        try { dataA = colA >= 0 ? JSON.parse(String(tVals[r][colA]||'{}')) : {}; } catch(e2) {}
        try { dataB = colB >= 0 ? JSON.parse(String(tVals[r][colB]||'{}')) : {}; } catch(e2) {}
        // Skip if both are empty
        if (!Object.keys(dataA).length && !Object.keys(dataB).length) continue;
        // Merge: use whichever is more progressed; combine substeps
        var statusOrder = ['Done','Fit','Pending','Not Started','Problem','Unfit',''];
        var stA = dataA.status||'', stB = dataB.status||'';
        var mergedStatus = statusOrder.indexOf(stA) <= statusOrder.indexOf(stB) ? stA : stB;
        var mergedSubsteps = {};
        Object.keys(dataA.substeps||{}).forEach(function(sk){ mergedSubsteps[sk] = dataA.substeps[sk]; });
        Object.keys(dataB.substeps||{}).forEach(function(sk){ mergedSubsteps[sk] = dataB.substeps[sk]; });
        var mergedData = {
          status:        mergedStatus || 'Pending',
          responsible:   dataA.responsible || dataB.responsible || '',
          start_date:    dataA.start_date  || dataB.start_date  || '',
          complete_date: dataA.complete_date || dataB.complete_date || '',
          notes:         [dataA.notes, dataB.notes].filter(Boolean).join(' | '),
          substeps:      mergedSubsteps
        };
        tsh.getRange(r+1, colNew+1).setValue(JSON.stringify(mergedData));
        migratedCount++;
      }
    }
    logActivity('StepConfig','MERGE', keyA+'+'+keyB+'->'+newStep.key, 'SUCCESS - '+migratedCount+' records');
    return {success:true, migratedCount:migratedCount};
  } catch(e) { return {success:false, error:e.message}; }
}

// ============================================================
// MASTER DATA
// ============================================================
function getMasterData(profile) {
  try {
    var sh = SS.getSheetByName(TABS.MASTER);
    if (!sh) return {success:false, error:'Master Data sheet not found'};
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var headers = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) row[headers[j]] = vals[i][j] !== undefined ? String(vals[i][j]) : '';
      rows.push(cleanMasterRow(row));
    }
    var p = profile || getMyProfile();
    // EMPLOYEE role: only return their own record
    if (p.role === 'EMPLOYEE') {
      var empEmail = p.email.toLowerCase();
      rows = rows.filter(function(r){ return String(r.EMAIL||'').toLowerCase().trim() === empEmail; });
    } else {
      rows = _filterByEntity(rows, p);
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

function cleanMasterRow(row) {
  var natMap = {'INDIAN':'INDIA','PAKISTANI':'PAKISTAN','EGYPTIAN':'EGYPT','UGANDAN':'UGANDA','FILIPINO':'PHILIPPINES','NEPALI':'NEPAL'};
  var nat = String(row['NATIONALITY']||'').trim().toUpperCase();
  row['NATIONALITY'] = natMap[nat] || nat;

  var lic = String(row['LIC AUTH']||'').trim().toUpperCase();
  if (!lic || lic==='NAN'||lic==='NIL'||lic==='') row['LIC AUTH']='NIL';
  else if (lic.indexOf('PSBD')>=0) row['LIC AUTH']='PSBD';
  else if (lic.indexOf('SIRA')>=0) row['LIC AUTH']='SIRA';
  else row['LIC AUTH']='OTHER';

  var smap={'ABUDHABI':'ABU DHABI','ABU-DHABI':'ABU DHABI','DUABI':'DUBAI','DUBAII':'DUBAI','NE':'OTHER'};
  var sts=String(row['STATUS']||'').trim().toUpperCase();
  row['STATUS']=smap[sts]||sts;
  var visa=String(row['VISA']||'').trim().toUpperCase();
  row['VISA']=smap[visa]||visa||row['STATUS'];

  row['ENTITY']=normaliseEntity(row['ENTITY']);
  row['NAME']=String(row['NAME']||'').trim().toUpperCase();
  return row;
}

function normaliseEntity(raw) {
  var v=String(raw||'').trim().toUpperCase();
  if (v==='UGU'||v==='USU') return 'UG';
  if (v==='USG-M') return 'USG-M';
  if (v==='USG') return 'USG';
  if (v==='UG') return 'UG';
  if (v==='UST') return 'UST';
  return v||'USG';
}

// ============================================================
// ACTIVE MANPOWER — Master minus Deletion_Log, deduped
// Writes result to Active_Manpower sheet
// ============================================================
function buildActiveManpower() {
  try {
    var masterRes = getMasterData();
    if (!masterRes.success) return {success:false, error:masterRes.error};
    var delRes = getDeletionLog();
    var deletedIds = {};
    if (delRes.success) {
      delRes.data.forEach(function(d){ if(d.EMP_ID) deletedIds[String(d.EMP_ID).trim()]=true; });
    }
    var active = masterRes.data.filter(function(r){
      return r.STATUS!=='DELETED' && !deletedIds[String(r.ID||'').trim()];
    });
    // Deduplicate by ID
    var seen={}, deduped=[];
    active.forEach(function(r){
      var k=String(r.ID||'').trim();
      if(k && !seen[k]){ seen[k]=true; deduped.push(r); }
    });
    // Write to Active_Manpower sheet
    var hdrs=['ID','NAME','VISA','STATUS','ENTITY','LIC AUTH','DESIGNATION','DATE OF JOIN','NATIONALITY','BIRTH DATE','PASSPORT NO','EID NO','AGE','Days'];
    var sh=SS.getSheetByName(TABS.ACTIVE_MP);
    if(sh){ sh.clearContents(); } else { sh=SS.insertSheet(TABS.ACTIVE_MP); }
    sh.appendRow(hdrs);
    deduped.forEach(function(r){ sh.appendRow(hdrs.map(function(h){ return r[h]||''; })); });
    logActivity('ActiveManpower','BUILD',deduped.length+' active records','SUCCESS');
    return {success:true, count:deduped.length, data:deduped};
  } catch(e) { return {success:false, error:e.message}; }
}

// ============================================================
// ADD / UPDATE / DELETE EMPLOYEE
// ============================================================
function addEmployee(data) {
  try {
    var sh=SS.getSheetByName(TABS.MASTER);
    if(!sh) return {success:false, error:'Master Data sheet not found'};
    var vals=sh.getDataRange().getValues();
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var idCol=hdrs.indexOf('ID'), ppCol=hdrs.indexOf('PASSPORT NO');
    for(var i=1;i<vals.length;i++){
      if(idCol>=0&&String(vals[i][idCol]).trim()===String(data.ID).trim()) return {success:false,error:'Employee ID already exists'};
      if(ppCol>=0&&String(vals[i][ppCol]).trim()===String(data['PASSPORT NO']).trim()) return {success:false,error:'Passport No already in system'};
    }
    sh.appendRow(buildRow(hdrs,data));
    logActivity('EmployeeAgent','ADD',data.ID+'--'+data.NAME,'SUCCESS');
    return {success:true};
  } catch(e){ return {success:false,error:e.message}; }
}

function updateEmployee(data) {
  try {
    var sh=SS.getSheetByName(TABS.MASTER); if(!sh) return {success:false,error:'Sheet not found'};
    var vals=sh.getDataRange().getValues();
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var idCol=hdrs.indexOf('ID');
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][idCol]).trim()===String(data.ID).trim()){
        sh.getRange(i+1,1,1,hdrs.length).setValues([buildRow(hdrs,data)]);
        logActivity('EmployeeAgent','EDIT',data.ID+' updated','SUCCESS');
        return {success:true};
      }
    }
    return {success:false,error:'Employee not found'};
  } catch(e){ return {success:false,error:e.message}; }
}

function deleteEmployee(empId, reason, lastWorkingDate) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var deletedBy = profile.email;
    var sh=SS.getSheetByName(TABS.MASTER); if(!sh) return {success:false,error:'Sheet not found'};
    var vals=sh.getDataRange().getValues();
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var idCol=hdrs.indexOf('ID'), stsCol=hdrs.indexOf('STATUS');
    var ppCol=hdrs.indexOf('PASSPORT NO'), nmCol=hdrs.indexOf('NAME');
    var desigCol=hdrs.indexOf('DESIGNATION'), joinCol=hdrs.indexOf('DATE OF JOIN'), entCol=hdrs.indexOf('ENTITY');
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][idCol]).trim()===String(empId).trim()){
        var delSh=getOrCreate(TABS.DEL_LOG,['LOG_ID','EMP_ID','FULL_NAME','PASSPORT_NO','REASON','DELETED_DATE','DELETED_BY','DESIGNATION','DATE_OF_JOIN','GROUP','LAST_WORKING_DATE']);
        // One-time column migration: add LAST_WORKING_DATE header if missing
        var delHdrs=delSh.getRange(1,1,1,delSh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
        if(delHdrs.indexOf('LAST_WORKING_DATE')<0){
          delSh.getRange(1, delSh.getLastColumn()+1).setValue('LAST_WORKING_DATE');
        }
        delSh.appendRow([
          'DEL-'+new Date().getTime(), empId,
          vals[i][nmCol]||'', vals[i][ppCol]||'',
          reason, formatDate(new Date()), deletedBy,
          desigCol>=0 ? vals[i][desigCol]||'' : '',
          joinCol>=0  ? vals[i][joinCol]||''  : '',
          entCol>=0   ? normaliseEntity(vals[i][entCol]) : '',
          lastWorkingDate || ''
        ]);
        if(stsCol>=0) sh.getRange(i+1,stsCol+1).setValue('DELETED');
        logActivity('EmployeeAgent','DELETE',empId+'|'+reason,'SUCCESS');
        return {success:true};
      }
    }
    return {success:false,error:'Employee not found'};
  } catch(e){ return {success:false,error:e.message}; }
}

function buildRow(headers, data) {
  return headers.map(function(h){ return data[h]!==undefined?data[h]:''; });
}

// ============================================================
// DELETION LOG
// ============================================================
function getDeletionLog() {
  try {
    var sh=SS.getSheetByName(TABS.DEL_LOG); if(!sh) return {success:true,data:[]};
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=1;i<vals.length;i++){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

// Returns Deletion_Log sorted newest-first, enriched with DESIGNATION & DATE_OF_JOIN from Master
function getResignations() {
  try {
    var sh=SS.getSheetByName(TABS.DEL_LOG); if(!sh) return {success:true,data:[]};
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=1;i<vals.length;i++){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }

    // Build Master lookup: ID → { DESIGNATION, DATE_OF_JOIN }
    // Employees keep their row in Master (STATUS=DELETED), so we can join on ID = EMP_ID
    var masterMap = {};
    var mSh = SS.getSheetByName(TABS.MASTER);
    if (mSh) {
      var mVals = mSh.getDataRange().getValues();
      if (mVals.length > 1) {
        var mHdrs    = mVals[0].map(function(h){ return String(h).trim(); });
        var idIdx    = mHdrs.indexOf('ID');
        var desigIdx = mHdrs.indexOf('DESIGNATION');
        var joinIdx  = mHdrs.indexOf('DATE OF JOIN');
        for (var mi = 1; mi < mVals.length; mi++) {
          var eid = String(mVals[mi][idIdx]||'').trim();
          if (!eid) continue;
          masterMap[eid] = {
            DESIGNATION:  desigIdx >= 0 ? String(mVals[mi][desigIdx]||'') : '',
            DATE_OF_JOIN: joinIdx  >= 0 ? String(mVals[mi][joinIdx] ||'') : ''
          };
        }
      }
    }

    // Enrich each row: Master values take priority; fall back to Deletion_Log values
    rows.forEach(function(row) {
      var m = masterMap[String(row.EMP_ID||'').trim()];
      if (m) {
        if (m.DESIGNATION)  row.DESIGNATION  = m.DESIGNATION;
        if (m.DATE_OF_JOIN) row.DATE_OF_JOIN = m.DATE_OF_JOIN;
      }
    });

    rows.reverse();
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

// ============================================================
// ONBOARDING
// ============================================================
function getOnboarding(profile) {
  try {
    var sh=getOrCreate(TABS.ONBOARDING,['OB_ID','FULL_NAME','PASSPORT_NO','POSITION_TYPE','MOBILE','VISA_STATUS','EXP_JOIN_DATE','DATE_ADDED','STATUS','NOTES','ENTITY','ASSIGNED_TO']);
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=1;i<vals.length;i++){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }
    var p = profile || getMyProfile();
    rows = _filterByEntity(rows, p);
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

function addOnboarding(data) {
  try {
    var sh=getOrCreate(TABS.ONBOARDING,['OB_ID','FULL_NAME','PASSPORT_NO','POSITION_TYPE','MOBILE','VISA_STATUS','EXP_JOIN_DATE','DATE_ADDED','STATUS','NOTES','ENTITY','ASSIGNED_TO']);
    sh.appendRow([data.OB_ID,data.FULL_NAME,data.PASSPORT_NO,data.POSITION_TYPE,data.MOBILE||'',data.VISA_STATUS||'',data.EXP_JOIN_DATE||'',formatDate(new Date()),'Pending',data.NOTES||'',data.ENTITY||'',data.ASSIGNED_TO||'']);
    logActivity('OnboardingAgent','ADD',data.OB_ID+'--'+data.FULL_NAME,'SUCCESS');
    return {success:true};
  } catch(e){ return {success:false,error:e.message}; }
}

function updateOnboarding(data) {
  try {
    var sh=SS.getSheetByName(TABS.ONBOARDING); if(!sh) return {success:false,error:'Onboarding sheet not found'};
    var vals=sh.getDataRange().getValues();
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var idCol=hdrs.indexOf('OB_ID');
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][idCol])===String(data.OB_ID)){
        var row=hdrs.map(function(h){ return data[h]!==undefined?data[h]:String(vals[i][hdrs.indexOf(h)]||''); });
        sh.getRange(i+1,1,1,row.length).setValues([row]);
        logActivity('OnboardingAgent','UPDATE',data.OB_ID,'SUCCESS');
        return {success:true};
      }
    }
    return {success:false,error:'Not found'};
  } catch(e){ return {success:false,error:e.message}; }
}

function transferToMaster(obId, empData) {
  try {
    var result=addEmployee(empData); if(!result.success) return result;
    var sh=SS.getSheetByName(TABS.ONBOARDING);
    var obRow = null;
    if(sh){
      var vals=sh.getDataRange().getValues();
      var hdrs=vals[0].map(function(h){ return String(h).trim(); });
      var idCol=hdrs.indexOf('OB_ID'), stsCol=hdrs.indexOf('STATUS');
      for(var i=1;i<vals.length;i++){
        if(String(vals[i][idCol])===String(obId)){
          if(stsCol>=0) sh.getRange(i+1,stsCol+1).setValue('Transferred');
          obRow = {};
          hdrs.forEach(function(h,idx){ obRow[h] = vals[i][idx]; });
          break;
        }
      }
    }
    // Auto-create 20DS Tracker record
    create20DSRecord(obId, empData, obRow);
    logActivity('OnboardingAgent','TRANSFER',obId+'->'+empData.ID,'SUCCESS');
    return {success:true};
  } catch(e){ return {success:false,error:e.message}; }
}

// ============================================================
// 20 DAYS STRATEGY TRACKER MODULE
// ============================================================
var DS_TRACKER_HEADERS = [
  'DS_ID','EMP_ID','EMP_NAME','DESIGNATION','PHONE','EMAIL',
  'EXP_JOIN_DATE','PIPELINE_ADDED_DATE','TRANSFER_DATE','TRANSFERRED_BY','RESPONSIBLE_HR',
  'LIC_TYPE',
  'STEP_VISA','STEP_LABOR','STEP_MEDICAL','STEP_INSURANCE','STEP_NSI','STEP_EID',
  'STEP_ASSD',
  'STEP_SIRA_CERT','STEP_PCC','STEP_SIRA_LIC',
  'CANCELLED','CANCEL_REASON','CANCELLED_BY','CANCELLED_ON','TOTAL_DAYS_ELAPSED','OB_COMPLETE'
];

var DS_AUDIT_HEADERS = [
  'TIMESTAMP','EMP_ID','EMP_NAME','STEP_KEY','STEP_LABEL',
  'OLD_STATUS','NEW_STATUS','DATE_COMPLETED','REASON_NOTES',
  'UPDATED_BY','ROLE','DAYS_SINCE_TRANSFER'
];

// UAE timezone datetime string
function _uaeDT_() {
  return Utilities.formatDate(new Date(), 'Asia/Dubai', 'dd/MM/yyyy HH:mm:ss');
}

// Empty step JSON — simple: just status, complete_date, notes, substeps
function _emptyStep_(status) {
  return JSON.stringify({
    status: status || 'Pending',
    responsible: '', substeps: {},
    complete_date: '',
    notes: '', reason: ''
  });
}

function create20DSRecord(obId, empData, obRow) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = getOrCreate(TABS.DS_TRACKER, DS_TRACKER_HEADERS);
      var transferrer = getMyProfile();
      var dsId = genId_('DS');
      var now = _uaeDT_();
      var pipelineDate = (obRow && obRow.DATE_ADDED) ? formatDate(new Date(obRow.DATE_ADDED)) : now;
      var licType = _licTypeFromAuth_(empData['LIC_AUTH'] || empData.LIC_AUTH || '');
      var isSIRA = licType === 'SIRA';
      sh.appendRow([
        dsId,
        empData.ID,
        empData.NAME || (obRow && obRow.FULL_NAME) || '',
        empData.DESIGNATION || '',
        (obRow && obRow.MOBILE) || '',
        '',
        (obRow && obRow.EXP_JOIN_DATE) ? formatDate(new Date(obRow.EXP_JOIN_DATE)) : '',
        pipelineDate,
        now,
        (transferrer && transferrer.name) || '',
        (transferrer && transferrer.name) || '',
        licType,                            // LIC_TYPE
        _emptyStep_('Pending'),             // STEP_VISA
        _emptyStep_('Pending'),             // STEP_LABOR
        _emptyStep_('Pending'),             // STEP_MEDICAL
        _emptyStep_('Pending'),             // STEP_INSURANCE
        _emptyStep_('Not Started'),         // STEP_NSI
        _emptyStep_('Pending'),             // STEP_EID
        isSIRA ? _emptyStep_('Locked') : _emptyStep_('Locked'), // STEP_ASSD (locked for both; unlocked after EID for ASSD type)
        isSIRA ? _emptyStep_('Locked') : _emptyStep_('Locked'), // STEP_SIRA_CERT
        isSIRA ? _emptyStep_('Locked') : _emptyStep_('Locked'), // STEP_PCC
        isSIRA ? _emptyStep_('Locked') : _emptyStep_('Locked'), // STEP_SIRA_LIC
        'FALSE','','','','0','FALSE'
      ]);
    } finally { lock.releaseLock(); }
    logActivity('20DSTracker','CREATE',dsId+'->'+empData.ID,'SUCCESS');
    return {success:true, dsId:dsId};
  } catch(e){ logActivity('20DSTracker','CREATE_ERROR','',e.message); return {success:false,error:e.message}; }
}

function get20DSTracker(profile) {
  try {
    var sh = getOrCreate(TABS.DS_TRACKER, DS_TRACKER_HEADERS);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true,data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    // One-time migration: add STEP_ASSD column
    if (hdrs.indexOf('STEP_ASSD') < 0) {
      var eidPos = hdrs.indexOf('STEP_EID');
      var insertCol = eidPos >= 0 ? eidPos + 2 : sh.getLastColumn() + 1;
      sh.insertColumnBefore(insertCol);
      sh.getRange(1, insertCol).setValue('STEP_ASSD');
      if (vals.length > 1) sh.getRange(2, insertCol, vals.length - 1, 1).setValue(_emptyStep_('Locked'));
      vals = sh.getDataRange().getValues();
      hdrs = vals[0].map(function(h){ return String(h).trim(); });
    }
    // One-time migration: add LIC_TYPE column (existing rows default to 'ASSD')
    if (hdrs.indexOf('LIC_TYPE') < 0) {
      var ltInsert = sh.getLastColumn() + 1;
      // Insert after RESPONSIBLE_HR if present
      var rhrPos = hdrs.indexOf('RESPONSIBLE_HR');
      if (rhrPos >= 0) { ltInsert = rhrPos + 2; sh.insertColumnBefore(ltInsert); }
      else { sh.insertColumnAfter(sh.getLastColumn()); }
      sh.getRange(1, ltInsert).setValue('LIC_TYPE');
      if (vals.length > 1) sh.getRange(2, ltInsert, vals.length - 1, 1).setValue('ASSD');
      vals = sh.getDataRange().getValues();
      hdrs = vals[0].map(function(h){ return String(h).trim(); });
    }
    // One-time migration: add SIRA step columns after STEP_ASSD
    ['STEP_SIRA_CERT','STEP_PCC','STEP_SIRA_LIC'].forEach(function(col){
      if (hdrs.indexOf(col) >= 0) return;
      var assdPos = hdrs.indexOf('STEP_ASSD');
      var ins = assdPos >= 0 ? assdPos + 2 : sh.getLastColumn() + 1;
      sh.insertColumnBefore(ins);
      sh.getRange(1, ins).setValue(col);
      if (vals.length > 1) sh.getRange(2, ins, vals.length - 1, 1).setValue(_emptyStep_('Locked'));
      vals = sh.getDataRange().getValues();
      hdrs = vals[0].map(function(h){ return String(h).trim(); });
    });
    var data = [];
    var nowMs = new Date().getTime();
    var ALL_STEP_COLS = ['STEP_VISA','STEP_LABOR','STEP_MEDICAL','STEP_INSURANCE','STEP_NSI',
                         'STEP_EID','STEP_ASSD','STEP_SIRA_CERT','STEP_PCC','STEP_SIRA_LIC'];
    for (var i = 1; i < vals.length; i++) {
      var row = {};
      hdrs.forEach(function(h,idx){
        var v = vals[i][idx];
        row[h] = (typeof v === 'boolean') ? (v ? 'TRUE' : 'FALSE') : String(v||'');
      });
      // Parse JSON step fields
      ALL_STEP_COLS.forEach(function(k){
        try { row[k] = JSON.parse(row[k]); }
        catch(e){ row[k] = {status:'Pending',responsible:'',substeps:{},start_date:'',complete_date:'',start_dt:'',end_dt:'',duration_hours:'',notes:'',reason:'',blocker_reason:''}; }
      });
      // Live-compute days elapsed from TRANSFER_DATE
      var tDate = row['TRANSFER_DATE'];
      if (tDate && row['CANCELLED'] !== 'TRUE' && row['OB_COMPLETE'] !== 'TRUE') {
        var d = _parseDT_(tDate);
        if (d) row['TOTAL_DAYS_ELAPSED'] = String(Math.floor((nowMs - d.getTime()) / 86400000));
      }
      data.push(row);
    }
    return {success:true, data:data};
  } catch(e){ return {success:false,error:e.message}; }
}

// Parse dd/MM/yyyy HH:mm:ss or dd/MM/yyyy into a Date
function _parseDT_(s) {
  if (!s) return null;
  var m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}):(\d{2}))?/);
  if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]),
    Number(m[4]||0), Number(m[5]||0), Number(m[6]||0));
  var d = new Date(s);
  return isNaN(d) ? null : d;
}

// SERVER-SIDE COMPLETE: record UAE timestamp as end_dt, compute duration_hours
function completeDSStep(dsId, stepKey, notes, blockerReason) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:false,error:'20 Days Strategy Tracker sheet not found'};
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var vals = sh.getDataRange().getValues();
      var hdrs = vals[0].map(function(h){ return String(h).trim(); });
      var idCol = hdrs.indexOf('DS_ID');
      var stepCol = hdrs.indexOf(stepKey);
      if (stepCol < 0) return {success:false,error:'Unknown step: '+stepKey};
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]) !== String(dsId)) continue;
        // Enforce prerequisite steps for locked steps
        var prereqMap = { STEP_ASSD:'STEP_EID', STEP_SIRA_CERT:'STEP_EID', STEP_PCC:'STEP_SIRA_CERT', STEP_SIRA_LIC:'STEP_PCC' };
        var prereqKey = prereqMap[stepKey];
        if (prereqKey) {
          var prereqCol = hdrs.indexOf(prereqKey);
          if (prereqCol >= 0) {
            try {
              var prereqStep = JSON.parse(String(vals[i][prereqCol]||'{}'));
              if (prereqStep.status !== 'Done' && prereqStep.status !== 'Fit') {
                return {success:false, error:_dsStepLabel_(stepKey)+' can only be updated after '+_dsStepLabel_(prereqKey)+' is completed.'};
              }
            } catch(e2){ return {success:false, error:'Cannot verify prerequisite step status.'}; }
          }
        }
        var step = {};
        try { step = JSON.parse(String(vals[i][stepCol]||'{}')); } catch(e){}
        var nowDT = _uaeDT_();
        step.complete_date = nowDT.substring(0,10);
        step.status        = (stepKey === 'STEP_MEDICAL') ? 'Fit' : 'Done';
        if (notes) step.notes = notes;
        var newStepJson = JSON.stringify(step);
        sh.getRange(i+1, stepCol+1).setValue(newStepJson);
        vals[i][stepCol] = newStepJson; // keep in-memory array in sync so _refreshDSTotals_ sees the new status
        _refreshDSTotals_(sh, vals, hdrs, i);
        _logDSAudit_(vals[i][hdrs.indexOf('EMP_ID')], vals[i][hdrs.indexOf('EMP_NAME')],
          stepKey, _dsStepLabel_(stepKey), 'Pending', step.status,
          nowDT.substring(0,10), notes||'', profile, sh, vals, hdrs, i);
        logActivity('20DSTracker','STEP_COMPLETE',dsId+':'+stepKey,'SUCCESS');
        return {success:true, complete_date: nowDT.substring(0,10)};
      }
      return {success:false,error:'Record not found'};
    } finally { lock.releaseLock(); }
  } catch(e){ return {success:false,error:e.message}; }
}

function update20DSStep(dsId, stepKey, stepData) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:false,error:'20 Days Strategy Tracker sheet not found'};
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var vals = sh.getDataRange().getValues();
      var hdrs = vals[0].map(function(h){ return String(h).trim(); });
      var idCol = hdrs.indexOf('DS_ID');
      var stepCol = hdrs.indexOf(stepKey);
      if (stepCol < 0) return {success:false,error:'Unknown step: '+stepKey};
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]) !== String(dsId)) continue;
        // Enforce prerequisite steps (skip check when explicitly locking a step)
        var prereqMap2 = { STEP_ASSD:'STEP_EID', STEP_SIRA_CERT:'STEP_EID', STEP_PCC:'STEP_SIRA_CERT', STEP_SIRA_LIC:'STEP_PCC' };
        var prereqKey2 = prereqMap2[stepKey];
        if (prereqKey2 && stepData.status !== 'Locked') {
          var prereqCol2 = hdrs.indexOf(prereqKey2);
          if (prereqCol2 >= 0) {
            try {
              var prereqStep2 = JSON.parse(String(vals[i][prereqCol2]||'{}'));
              if (prereqStep2.status !== 'Done' && prereqStep2.status !== 'Fit') {
                return {success:false, error:_dsStepLabel_(stepKey)+' can only be updated after '+_dsStepLabel_(prereqKey2)+' is completed.'};
              }
            } catch(e3){ return {success:false, error:'Cannot verify prerequisite step status.'}; }
          }
        }
        var oldStep = {};
        try { oldStep = JSON.parse(String(vals[i][stepCol]||'{}')); } catch(e){}
        var oldStatus = oldStep.status || 'Pending';
        var newStepJson2 = JSON.stringify(stepData);
        sh.getRange(i+1, stepCol+1).setValue(newStepJson2);
        vals[i][stepCol] = newStepJson2; // keep in-memory array in sync
        _refreshDSTotals_(sh, vals, hdrs, i);
        _logDSAudit_(vals[i][hdrs.indexOf('EMP_ID')], vals[i][hdrs.indexOf('EMP_NAME')],
          stepKey, _dsStepLabel_(stepKey), oldStatus, stepData.status||'',
          stepData.complete_date||'', stepData.reason||stepData.notes||'', profile, sh, vals, hdrs, i);
        logActivity('20DSTracker','STEP_UPDATE',dsId+':'+stepKey+'->'+stepData.status,'SUCCESS');
        return {success:true};
      }
      return {success:false,error:'Record not found'};
    } finally { lock.releaseLock(); }
  } catch(e){ return {success:false,error:e.message}; }
}

function update20DSResponsible(dsId, responsibleHR) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:false,error:'Sheet not found'};
    var lock = LockService.getScriptLock(); lock.waitLock(5000);
    try {
      var vals = sh.getDataRange().getValues();
      var hdrs = vals[0].map(function(h){ return String(h).trim(); });
      var idCol = hdrs.indexOf('DS_ID'), hrCol = hdrs.indexOf('RESPONSIBLE_HR');
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]) === String(dsId)) {
          if (hrCol >= 0) sh.getRange(i+1, hrCol+1).setValue(responsibleHR||'');
          return {success:true};
        }
      }
      return {success:false,error:'Record not found'};
    } finally { lock.releaseLock(); }
  } catch(e){ return {success:false,error:e.message}; }
}

function cancel20DSRecord(dsId, reason) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:false,error:'Sheet not found'};
    var lock = LockService.getScriptLock(); lock.waitLock(5000);
    try {
      var vals = sh.getDataRange().getValues();
      var hdrs = vals[0].map(function(h){ return String(h).trim(); });
      var idCol = hdrs.indexOf('DS_ID');
      var cancelCol = hdrs.indexOf('CANCELLED'), reasonCol = hdrs.indexOf('CANCEL_REASON');
      var byCol = hdrs.indexOf('CANCELLED_BY'), onCol = hdrs.indexOf('CANCELLED_ON');
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]) === String(dsId)) {
          if (cancelCol>=0) sh.getRange(i+1,cancelCol+1).setValue('TRUE');
          if (reasonCol>=0) sh.getRange(i+1,reasonCol+1).setValue(reason||'');
          if (byCol>=0)     sh.getRange(i+1,byCol+1).setValue((profile&&profile.name)||'');
          if (onCol>=0)     sh.getRange(i+1,onCol+1).setValue(_uaeDT_());
          logActivity('20DSTracker','CANCEL',dsId,'SUCCESS');
          return {success:true};
        }
      }
      return {success:false,error:'Record not found'};
    } finally { lock.releaseLock(); }
  } catch(e){ return {success:false,error:e.message}; }
}

// Recompute TOTAL_DAYS_ELAPSED + OB_COMPLETE for every non-cancelled row
function recalculate20DSTotals() {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:false, error:'Sheet not found'};
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      var vals = sh.getDataRange().getValues();
      if (vals.length < 2) return {success:true, updated:0};
      var hdrs = vals[0].map(function(h){ return String(h).trim(); });
      var cancelCol = hdrs.indexOf('CANCELLED');
      var updated = 0;
      for (var i = 1; i < vals.length; i++) {
        var cv = vals[i][cancelCol];
        if (cv === true || String(cv||'').toUpperCase() === 'TRUE') continue;
        _refreshDSTotals_(sh, vals, hdrs, i);
        updated++;
      }
      logActivity('20DSTracker','RECALCULATE','','SUCCESS - '+updated+' rows');
      return {success:true, updated:updated};
    } finally { lock.releaseLock(); }
  } catch(e){ return {success:false, error:e.message}; }
}

/**
 * Run this directly from the Apps Script editor (no login needed).
 * It fixes OB_COMPLETE and TOTAL_DAYS_ELAPSED for every non-cancelled
 * DS tracker row — useful for employees who finished after day 20 but
 * were never flagged as complete due to the stale-array bug.
 */
function adminFixOBComplete() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('20DS Tracker');
  if (!sh) { Logger.log('Sheet "20DS Tracker" not found'); return; }
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { Logger.log('No data'); return; }
  var hdrs = vals[0].map(function(h){ return String(h).trim(); });
  var cancelCol  = hdrs.indexOf('CANCELLED');
  var totalCol   = hdrs.indexOf('TOTAL_DAYS_ELAPSED');
  var obCol      = hdrs.indexOf('OB_COMPLETE');
  var tCol       = hdrs.indexOf('TRANSFER_DATE');
  var licTypeCol = hdrs.indexOf('LIC_TYPE');
  var fixed = 0;
  for (var i = 1; i < vals.length; i++) {
    var cv2 = vals[i][cancelCol];
    if (cv2 === true || String(cv2||'').toUpperCase() === 'TRUE') continue;
    var licType = licTypeCol >= 0 ? String(vals[i][licTypeCol] || '') : 'ASSD';
    var stepKeys = _stepKeysForType_(licType);
    var allDone = stepKeys.every(function(k) {
      var c = hdrs.indexOf(k);
      if (c < 0) return false;
      try { var s = JSON.parse(String(vals[i][c]||'{}')); return s.status==='Done'||s.status==='Fit'; }
      catch(e){ return false; }
    });
    var tRaw = String(vals[i][tCol]||'');
    var tD = null;
    var m = tRaw.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (m) tD = new Date(+m[3], +m[2]-1, +m[1]);
    else { var d2 = new Date(tRaw); if (!isNaN(d2)) tD = d2; }
    var days = 0;
    if (tD) {
      if (allDone) {
        var maxComplete = null;
        stepKeys.forEach(function(k) {
          var c = hdrs.indexOf(k);
          if (c < 0) return;
          try {
            var s = JSON.parse(String(vals[i][c]||'{}'));
            if (s.complete_date) {
              var cd; var m2 = s.complete_date.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
              if (m2) cd = new Date(+m2[3],+m2[2]-1,+m2[1]);
              else cd = new Date(s.complete_date);
              if (!isNaN(cd) && (!maxComplete || cd > maxComplete)) maxComplete = cd;
            }
          } catch(e2){}
        });
        days = Math.floor(((maxComplete||new Date()) - tD) / 86400000);
      } else {
        days = Math.floor((new Date() - tD) / 86400000);
      }
    }
    if (totalCol >= 0) sh.getRange(i+1, totalCol+1).setValue(days);
    if (obCol    >= 0) sh.getRange(i+1, obCol+1).setValue(allDone ? 'TRUE' : 'FALSE');
    if (allDone) fixed++;
  }
  Logger.log('adminFixOBComplete done. Marked complete: ' + fixed + ' row(s).');
}

function get20DSAuditLog() {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = getOrCreate(TABS.DS_AUDIT, DS_AUDIT_HEADERS);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true,data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var data = [];
    for (var i = vals.length-1; i >= 1; i--) {
      var row = {};
      hdrs.forEach(function(h,idx){ row[h] = String(vals[i][idx]||''); });
      data.push(row);
    }
    return {success:true, data:data};
  } catch(e){ return {success:false,error:e.message}; }
}

// Analytics: average days to complete onboarding, per HR and overall
function get20DSAnalytics() {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.DS_TRACKER);
    if (!sh) return {success:true, hrReport:[], overall:{avgDays:0, totalCompleted:0, totalActive:0}};
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, hrReport:[], overall:{avgDays:0, totalCompleted:0, totalActive:0}};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var hrMap = {};
    var totalCompleted = 0, totalDays = 0, totalActive = 0;
    var individuals = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {};
      hdrs.forEach(function(h,idx){
        var v = vals[i][idx];
        row[h] = (typeof v === 'boolean') ? (v ? 'TRUE' : 'FALSE') : String(v||'');
      });
      if (row.CANCELLED === 'TRUE') continue;
      var hrName = row.RESPONSIBLE_HR || 'Unassigned';
      if (!hrMap[hrName]) hrMap[hrName] = {completed:0, totalDays:0, active:0};
      if (row.OB_COMPLETE === 'TRUE') {
        var days = parseInt(row.TOTAL_DAYS_ELAPSED||0, 10);
        hrMap[hrName].completed++;
        hrMap[hrName].totalDays += days;
        totalCompleted++;
        totalDays += days;
        individuals.push({
          empId:        row.EMP_ID,
          empName:      row.EMP_NAME,
          designation:  row.DESIGNATION,
          hrName:       hrName,
          transferDate: row.TRANSFER_DATE ? String(row.TRANSFER_DATE).substring(0,10) : '',
          days:         days
        });
      } else {
        hrMap[hrName].active++;
        totalActive++;
      }
    }
    // Sort individuals fastest first
    individuals.sort(function(a,b){ return a.days - b.days; });
    var hrReport = Object.keys(hrMap).map(function(name){
      var d = hrMap[name];
      var avg = d.completed > 0 ? Math.round(d.totalDays / d.completed * 10) / 10 : null;
      return {name:name, avgDays:avg, completed:d.completed, active:d.active};
    }).sort(function(a,b){
      if (a.avgDays === null) return 1;
      if (b.avgDays === null) return -1;
      return a.avgDays - b.avgDays;
    });
    var overallAvg = totalCompleted > 0 ? Math.round(totalDays / totalCompleted * 10) / 10 : null;
    return {success:true, hrReport:hrReport, overall:{avgDays:overallAvg, totalCompleted:totalCompleted, totalActive:totalActive}, individuals:individuals};
  } catch(e){ return {success:false,error:e.message}; }
}

function _dsStepLabel_(key) {
  var labels = {
    STEP_VISA:      'Visa/Entry Issued',
    STEP_LABOR:     'Tawjeeh & Labor Card',
    STEP_MEDICAL:   'Visa Medical',
    STEP_INSURANCE: 'Medical Insurance',
    STEP_NSI:       'NSI Training',
    STEP_EID:       'EID & Residency Stamping',
    STEP_ASSD:      'ASSD',
    STEP_SIRA_CERT: 'SIRA Certificate',
    STEP_PCC:       'Police Clearance (PCC)',
    STEP_SIRA_LIC:  'SIRA License'
  };
  return labels[key] || key;
}

// Returns the correct ordered step keys based on employee LIC_TYPE
function _stepKeysForType_(licType) {
  if (String(licType || '').toUpperCase() === 'SIRA') {
    return ['STEP_VISA','STEP_LABOR','STEP_MEDICAL','STEP_INSURANCE','STEP_NSI','STEP_EID',
            'STEP_SIRA_CERT','STEP_PCC','STEP_SIRA_LIC'];
  }
  return ['STEP_VISA','STEP_LABOR','STEP_MEDICAL','STEP_INSURANCE','STEP_NSI','STEP_EID','STEP_ASSD'];
}

// Derives LIC_TYPE ('ASSD' | 'SIRA') from the LIC_AUTH master data column value
function _licTypeFromAuth_(licAuth) {
  return String(licAuth || '').toUpperCase().trim() === 'SIRA' ? 'SIRA' : 'ASSD';
}

function _refreshDSTotals_(sh, vals, hdrs, rowIdx) {
  try {
    var cancelCol   = hdrs.indexOf('CANCELLED');
    var totalCol    = hdrs.indexOf('TOTAL_DAYS_ELAPSED');
    var completeCol = hdrs.indexOf('OB_COMPLETE');
    var cancelRaw = vals[rowIdx][cancelCol];
    if (cancelRaw === true || String(cancelRaw||'').toUpperCase() === 'TRUE') return;
    var licTypeCol = hdrs.indexOf('LIC_TYPE');
    var licType = licTypeCol >= 0 ? String(vals[rowIdx][licTypeCol] || '') : 'ASSD';
    var stepKeys = _stepKeysForType_(licType);
    var allDone = stepKeys.every(function(k){
      var col = hdrs.indexOf(k);
      if (col < 0) return false;
      try {
        var s = JSON.parse(String(vals[rowIdx][col]||'{}'));
        return s.status === 'Done' || s.status === 'Fit';
      } catch(e){ return false; }
    });
    var transferCol = hdrs.indexOf('TRANSFER_DATE');
    var tD = _parseDT_(String(vals[rowIdx][transferCol]||''));
    var days = 0;
    if (tD) {
      if (allDone) {
        // Use the latest complete_date across all steps as the end point
        var maxComplete = null;
        stepKeys.forEach(function(k) {
          var col = hdrs.indexOf(k);
          if (col < 0) return;
          try {
            var s = JSON.parse(String(vals[rowIdx][col]||'{}'));
            if (s.complete_date) {
              var cd = _parseDT_(s.complete_date);
              if (cd && (!maxComplete || cd > maxComplete)) maxComplete = cd;
            }
          } catch(e2){}
        });
        days = Math.floor(((maxComplete || new Date()) - tD) / 86400000);
      } else {
        days = Math.floor((new Date() - tD) / 86400000);
      }
    }
    if (totalCol    >= 0) sh.getRange(rowIdx+1, totalCol+1).setValue(days);
    if (completeCol >= 0) sh.getRange(rowIdx+1, completeCol+1).setValue(allDone ? 'TRUE' : 'FALSE');
  } catch(e){}
}

function _logDSAudit_(empId, empName, stepKey, stepLabel, oldStatus, newStatus, completeDt, reason, profile, sh, vals, hdrs, rowIdx) {
  try {
    var auditSh = getOrCreate(TABS.DS_AUDIT, DS_AUDIT_HEADERS);
    var transferCol = hdrs.indexOf('TRANSFER_DATE');
    var daysSince = '';
    try {
      if (transferCol >= 0) {
        var tD = _parseDT_(String(vals[rowIdx][transferCol]||''));
        if (tD) daysSince = Math.floor((new Date() - tD) / 86400000);
      }
    } catch(e2){}
    auditSh.appendRow([
      _uaeDT_(),
      empId, empName, stepKey, stepLabel,
      oldStatus, newStatus, completeDt, reason,
      (profile&&profile.name)||'', (profile&&profile.role)||'',
      daysSince
    ]);
  } catch(e){}
}

// ============================================================
// HR DOCS
// ============================================================
function getHRDocs(profile) {
  try {
    var sh=getOrCreate(TABS.HR_DOCS,['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES','ENTITY']);
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=1;i<vals.length;i++){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }
    var p = profile || getMyProfile();
    rows = _filterByEntity(rows, p);
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

function issueHRDoc(data) {
  try {
    var sh=getOrCreate(TABS.HR_DOCS,['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES']);
    sh.appendRow([data.REF_NO,data.EMP_ID,data.EMP_NAME,data.LETTER_TYPE,data.ISSUE_DATE,data.ISSUED_BY||'HR',data.NOTES||'']);
    logActivity('LetterAgent','ISSUE',data.REF_NO+'--'+data.LETTER_TYPE,'SUCCESS');
    return {success:true};
  } catch(e){ return {success:false,error:e.message}; }
}

// ============================================================
// SUMMARY
// ============================================================
function getSummary() {
  try {
    var sh=getOrCreate(TABS.SUMMARY,['MONTH','TOTAL_EMP','PSBD','SIRA','NIL','ABU_DHABI','DUBAI','NEW_JOINS','EXITS','ONBOARDING','LETTERS_ISSUED','NOTES']);
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=1;i<vals.length;i++){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

// ============================================================
// ACTIVITY LOG
// ============================================================
function logActivity(agent,action,detail,status) {
  try {
    var actorEmail = '';
    try { actorEmail = Session.getActiveUser().getEmail(); } catch(e2) {}
    var sh=getOrCreate(TABS.ACTIVITY,['TIMESTAMP','AGENT','ACTION','DETAIL','STATUS','EMAIL']);
    sh.appendRow([new Date().toLocaleString(),agent,action,detail||'',status||'SUCCESS',actorEmail||agent]);
  } catch(e){}
}

function getActivityLog() {
  try {
    var sh=SS.getSheetByName(TABS.ACTIVITY); if(!sh) return {success:true,data:[]};
    var vals=sh.getDataRange().getValues();
    if(vals.length<2) return {success:true,data:[]};
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var rows=[];
    for(var i=vals.length-1;i>=1&&rows.length<200;i--){
      var row={}; for(var j=0;j<hdrs.length;j++) row[hdrs[j]]=String(vals[i][j]||'');
      rows.push(row);
    }
    return {success:true,data:rows};
  } catch(e){ return {success:false,error:e.message}; }
}

// ============================================================
// HELPERS
// ============================================================
function getOrCreate(name,headers){
  var sh=SS.getSheetByName(name);
  if(!sh){ sh=SS.insertSheet(name); if(headers) sh.appendRow(headers); }
  return sh;
}
function formatDate(d){
  return ('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();
}

// ============================================================
// LETTER TEMPLATES — HR Letters (Warning, Experience, etc.)
// Templates stored in AppConfig sheet as:
//   LTEMPL_{TYPE_KEY}_DRIVE  — Google Doc/Drive File ID
//   LTEMPL_{TYPE_KEY}_BODY   — in-app text body with placeholders
// Supported placeholders: {{NAME}} {{EMP_ID}} {{PASSPORT_NO}}
//   {{DESIGNATION}} {{COMPANY}} {{ENTITY}} {{DATE_OF_JOIN}}
//   {{SALARY}} {{DATE}} {{REF_NO}} {{ISSUED_BY}}
//   {{HR_OFFICER}} {{HR_DESIGNATION}}
// ============================================================
function getTypeKey_(type) {
  return String(type).toUpperCase().replace(/[^A-Z0-9]/g, '_');
}

function saveLetterTemplate(type, driveId, bodyText) {
  try {
    var sh = getOrCreate(TABS.CONFIG, ['KEY','VALUE']);
    var d = sh.getDataRange().getValues();
    var driveKey = 'LTEMPL_' + getTypeKey_(type) + '_DRIVE';
    var bodyKey  = 'LTEMPL_' + getTypeKey_(type) + '_BODY';
    var keyMap = {};
    for (var i = 1; i < d.length; i++) { if (d[i][0]) keyMap[String(d[i][0])] = i + 1; }
    if (keyMap[driveKey]) sh.getRange(keyMap[driveKey], 2).setValue(driveId || '');
    else sh.appendRow([driveKey, driveId || '']);
    if (keyMap[bodyKey]) sh.getRange(keyMap[bodyKey], 2).setValue(bodyText || '');
    else sh.appendRow([bodyKey, bodyText || '']);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function _numToWords_(n) {
  var ones=['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
            'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
            'Seventeen','Eighteen','Nineteen'];
  var tens=['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  if(n===0) return 'Zero';
  if(n<0) return 'Minus '+_numToWords_(-n);
  var w='';
  if(n>=1000){ w+=_numToWords_(Math.floor(n/1000))+' Thousand '; n=n%1000; }
  if(n>=100) { w+=ones[Math.floor(n/100)]+' Hundred '; n=n%100; }
  if(n>=20)  { w+=tens[Math.floor(n/10)]+' '; n=n%10; }
  if(n>0)    { w+=ones[n]+' '; }
  return w.trim();
}
function _salaryInWords_(salary) {
  var num = parseFloat(String(salary).replace(/,/g,''));
  if(isNaN(num)||num<0) return '';
  var whole = Math.floor(num);
  var fils  = Math.round((num-whole)*100);
  var result = _numToWords_(whole)+' Dirhams';
  if(fils>0) result += ' and '+_numToWords_(fils)+' Fils';
  return result+' Only';
}

function _fillPlaceholders(text, data, cfg) {
  var firstName = toProperCase(String(data.EMP_NAME||'').trim().split(' ')[0]);
  var map = {
    '{{NAME}}':           data.EMP_NAME    || '',
    '{{FIRSTNAME}}':      firstName,
    '{{ID}}':             data.EMP_ID      || '',
    '{{EMP_ID}}':         data.EMP_ID      || '',
    '{{PASSPORT_NO}}':    data.PASSPORT_NO || '',
    '{{DESIGNATION}}':    data.DESIGNATION || '',
    '{{COMPANY}}':        cfg.company_name || 'United Group Holding',
    '{{ENTITY}}':         data.ENTITY      || '',
    '{{DOJ}}':            data.DATE_OF_JOIN|| '',
    '{{DATE_OF_JOIN}}':   data.DATE_OF_JOIN|| '',
    '{{SALARY}}':         data.SALARY      || '',
    '{{SALARY_WORDS}}':   _salaryInWords_(data.SALARY||''),
    '{{BANK_NAME}}':      data.BANK_NAME   || '',
    '{{ADDRESS}}':        data.ADDRESS     || '',
    '{{DATE}}':           data.ISSUE_DATE  || '',
    '{{REF_NO}}':         data.REF_NO      || '',
    '{{ISSUED_BY}}':      data.ISSUED_BY   || '',
    '{{HR_OFFICER}}':     cfg.hr_officer   || 'HR Manager',
    '{{HR_DESIGNATION}}': cfg.designation  || 'HR Manager'
  };
  var result = text;
  Object.keys(map).forEach(function(k) { result = result.split(k).join(map[k]); });
  return result;
}

function _generateFromDriveTemplate(driveId, data, cfg) {
  var file = DriveApp.getFileById(driveId);
  var copy = file.makeCopy('_HR_TEMP_' + data.REF_NO);
  try {
    var doc = DocumentApp.openById(copy.getId());
    var body = doc.getBody();
    var firstName = toProperCase(String(data.EMP_NAME||'').trim().split(' ')[0]);
    var placeholders = {
      '{{NAME}}':           data.EMP_NAME    || '',
      '{{FIRSTNAME}}':      firstName,
      '{{ID}}':             data.EMP_ID      || '',
      '{{EMP_ID}}':         data.EMP_ID      || '',
      '{{PASSPORT_NO}}':    data.PASSPORT_NO || '',
      '{{DESIGNATION}}':    data.DESIGNATION || '',
      '{{COMPANY}}':        cfg.company_name || 'United Group Holding',
      '{{ENTITY}}':         data.ENTITY      || '',
      '{{DOJ}}':            data.DATE_OF_JOIN|| '',
      '{{DATE_OF_JOIN}}':   data.DATE_OF_JOIN|| '',
      '{{SALARY}}':         data.SALARY      || '',
      '{{SALARY_WORDS}}':   _salaryInWords_(data.SALARY||''),
      '{{BANK_NAME}}':      data.BANK_NAME   || '',
      '{{ADDRESS}}':        data.ADDRESS     || '',
      '{{DATE}}':           data.ISSUE_DATE  || '',
      '{{REF_NO}}':         data.REF_NO      || '',
      '{{ISSUED_BY}}':      data.ISSUED_BY   || '',
      '{{HR_OFFICER}}':     cfg.hr_officer   || 'HR Manager',
      '{{HR_DESIGNATION}}': cfg.designation  || 'HR Manager'
    };
    Object.keys(placeholders).forEach(function(k) { body.replaceText(k, placeholders[k]); });
    doc.saveAndClose();
    var pdfBytes = DriveApp.getFileById(copy.getId()).getAs('application/pdf').getBytes();
    return Utilities.base64Encode(pdfBytes);
  } finally {
    try { copy.setTrashed(true); } catch(e2) {}
  }
}

function _generateFromTextTemplate(bodyText, data, cfg) {
  var filled = _fillPlaceholders(bodyText, data, cfg);
  var doc = DocumentApp.create('_HR_TEMP_' + data.REF_NO);
  try {
    var body = doc.getBody();
    var lines = filled.split('\n');
    body.setText(lines[0] || ' ');
    for (var i = 1; i < lines.length; i++) { body.appendParagraph(lines[i]); }
    doc.saveAndClose();
    var pdfBytes = DriveApp.getFileById(doc.getId()).getAs('application/pdf').getBytes();
    return Utilities.base64Encode(pdfBytes);
  } finally {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(e2) {}
  }
}

function generateAndIssueLetter(data) {
  try {
    var cfg = getConfig();
    var typeKey  = getTypeKey_(data.LETTER_TYPE);
    var driveId  = String(cfg['LTEMPL_' + typeKey + '_DRIVE']  || '').trim();
    var bodyText = String(cfg['LTEMPL_' + typeKey + '_BODY']   || '').trim();
    var pdfBase64 = null;
    if (driveId)       pdfBase64 = _generateFromDriveTemplate(driveId, data, cfg);
    else if (bodyText) pdfBase64 = _generateFromTextTemplate(bodyText, data, cfg);
    var sh = getOrCreate(TABS.HR_DOCS, ['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES','ENTITY']);
    sh.appendRow([data.REF_NO, data.EMP_ID, data.EMP_NAME, data.LETTER_TYPE,
                  data.ISSUE_DATE, data.ISSUED_BY||'HR', data.NOTES||'', data.ENTITY||'']);
    logActivity('LetterAgent', 'ISSUE', data.REF_NO + '--' + data.LETTER_TYPE, 'SUCCESS');
    return { success: true, pdf: pdfBase64 };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
// EXPERIENCE LETTER — for resigned / deleted employees
// Template key reuses LTEMPL_EXPERIENCE_LETTER_DRIVE from Setup
// Supports placeholders: {{DATE}} {{ID}} {{NAME}} {{FIRSTNAME}}
//   {{DOJ}} {{DOL}} {{EMP_ID}} {{REF_NO}} {{DESIGNATION}}
//   {{HR_OFFICER}} {{HR_DESIGNATION}} {{COMPANY}}
// ============================================================
function _generateExpLetter(templateId, data, cfg) {
  var file = DriveApp.getFileById(templateId);
  var copy = file.makeCopy('_EXP_TEMP_' + data.REF_NO);
  try {
    var doc  = DocumentApp.openById(copy.getId());
    var body = doc.getBody();
    body.setFontFamily('Tahoma');
    body.setFontSize(12);
    var map = {
      '{{DATE}}': 'Date: ' +data.ISSUE_DATE  || formatDate(new Date()),
      '{{ID}}':             data.EMP_ID      || '',
      '{{NAME}}':           data.EMP_NAME    || '',
      '{{FIRSTNAME}}':      data.FIRSTNAME   || '',
      '{{DOJ}}':            data.DOJ         || '',
      '{{DOL}}':            data.DOL         || '',
      '{{EMP_ID}}':         data.EMP_ID      || '',
      '{{REF_NO}}':         data.REF_NO      || '',
      '{{DESIGNATION}}':    data.DESIGNATION || '',
      '{{ISSUED_BY}}':      data.ISSUED_BY   || '',
      '{{HR_OFFICER}}':     cfg.hr_officer   || 'HR Manager',
      '{{HR_DESIGNATION}}': cfg.designation  || 'HR Manager',
      '{{COMPANY}}':        cfg.company_name || 'United Group Holding'
    };
    Object.keys(map).forEach(function(k){ body.replaceText(k, map[k]); });
    doc.saveAndClose();
    var pdfBytes = DriveApp.getFileById(copy.getId()).getAs('application/pdf').getBytes();
    return Utilities.base64Encode(pdfBytes);
  } finally {
    try { copy.setTrashed(true); } catch(e2) {}
  }
}

function toProperCase(text) {
  if (!text) return '';
  return text.toLowerCase().replace(/\b\w/g, function(char) {
    return char.toUpperCase();
  });
}

function formatDate(date) {
  // If date is provided and valid, use it; otherwise use current date
  var d = date ? new Date(date) : new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// Robust date parser for sheet values — handles Date objects, DD/MM/YYYY strings,
// ISO strings, and guards against epoch / 1899-1900 serial-zero errors.
function _parseSheetDate(raw) {
  if (raw === null || raw === undefined || raw === '') return '';

  var d;

  if (raw instanceof Date) {
    d = raw;
  } else {
    var s = String(raw).trim();
    if (!s) return '';

    // Already DD/MM/YYYY — parse manually to avoid JS MM/DD/YYYY misinterpretation
    var dm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dm) {
      d = new Date(parseInt(dm[3], 10), parseInt(dm[2], 10) - 1, parseInt(dm[1], 10));
    } else {
      // ISO YYYY-MM-DD or YYYY/MM/DD
      var im = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})/);
      if (im) {
        d = new Date(parseInt(im[1], 10), parseInt(im[2], 10) - 1, parseInt(im[3], 10));
      } else {
        d = new Date(s);
      }
    }
  }

  if (!d || isNaN(d.getTime())) return '';
  var yr = d.getFullYear();
  // Reject clearly wrong years (epoch=1970, Sheets serial-0=1899, future junk)
  if (yr < 1950 || yr > 2100) return '';

  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function generateExperienceLetterForEmp(empId) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER']);

    // Read employee row from Deletion_Log
    var sh = SS.getSheetByName(TABS.DEL_LOG);
    if (!sh) return {success:false, error:'Deletion_Log sheet not found'};
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:false, error:'Deletion_Log is empty'};
    var hdrs  = vals[0].map(function(h){ return String(h).trim(); });
    var idCol = hdrs.indexOf('EMP_ID');
    var emp   = null;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idCol]).trim() === String(empId).trim()) {
        emp = {};
        for (var j = 0; j < hdrs.length; j++) {
          var _v = vals[i][j];
          // Keep Date objects intact so formatDate() can use them directly.
          // Converting a Date to String then back to new Date() loses the value in GAS.
          emp[hdrs[j]] = (_v instanceof Date) ? _v : String(_v||'');
        }
        break;
      }
    }
    if (!emp) return {success:false, error:'Employee ' + empId + ' not found in Deletion_Log'};

    // Look up DATE_OF_JOIN (and DESIGNATION as fallback) from Master using EMP_ID → ID
    // Employees remain in Master with STATUS=DELETED, so this join is always possible.
    var mSh = SS.getSheetByName(TABS.MASTER);
    if (mSh) {
      var mVals = mSh.getDataRange().getValues();
      var mHdrs = mVals[0].map(function(h){ return String(h).trim(); });
      var mIdIdx    = mHdrs.indexOf('ID');
      var mJoinIdx  = mHdrs.indexOf('DATE OF JOIN');
      var mDesigIdx = mHdrs.indexOf('DESIGNATION');
      for (var k = 1; k < mVals.length; k++) {
        if (String(mVals[k][mIdIdx]).trim() === String(empId).trim()) {
          if (mJoinIdx >= 0) {
            var rawJoin = mVals[k][mJoinIdx];
            // Keep Date objects intact; otherwise stringify
            emp.DATE_OF_JOIN = (rawJoin instanceof Date) ? rawJoin : String(rawJoin||'');
          }
          if (mDesigIdx >= 0 && !emp.DESIGNATION) {
            emp.DESIGNATION = String(mVals[k][mDesigIdx]||'');
          }
          break;
        }
      }
    }

    // Get template Drive ID from AppConfig
    var cfg = getConfig();
    var templateId = String(cfg['EXP_LETTER_TEMPLATE_ID'] || '').trim();
    if (!templateId) return {success:false,
      error:'Experience Letter template not configured. Go to Setup → Experience Letter Template ID and enter the Google Doc File ID.'};

    // Build letter data
    var cleanName = emp.FULL_NAME.trim().replace(/\s+/g,' ');
    var firstName = toProperCase(cleanName.split(' ')[0]);
    var refNo     = 'EXP-' + new Date().getTime().toString().slice(-8);
    var letterData = {
      REF_NO:      refNo,
      EMP_ID:      empId,
      EMP_NAME:    cleanName,
      FIRSTNAME:   firstName,
      DOJ:         _parseSheetDate(emp.DATE_OF_JOIN),
      DOL:         _parseSheetDate(emp.DELETED_DATE),
      DESIGNATION: emp.DESIGNATION  || '',
      ISSUE_DATE:  formatDate(new Date()),
      ISSUED_BY:   cfg.hr_officer   || 'HR'
    };

    var pdfBase64 = _generateExpLetter(templateId, letterData, cfg);

    // Log to HR Docs Tracker
    var hrSh = getOrCreate(TABS.HR_DOCS,
      ['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES','ENTITY']);
    hrSh.appendRow([
      refNo, empId, cleanName, 'Experience Letter',
      formatDate(new Date()), letterData.ISSUED_BY,
      'DOJ: ' + letterData.DOJ + ' | LWD: ' + letterData.DOL,
      emp.GROUP || emp.ENTITY || ''
    ]);
    logActivity('LetterAgent', 'EXP_LETTER', refNo + '--' + empId, 'SUCCESS');

    return {success:true, pdf:pdfBase64, refNo:refNo, name:cleanName};
  } catch(e) { return {success:false, error:e.message}; }
}

// ============================================================
// USER MANAGEMENT (SUPER_ADMIN only)
// ============================================================
function getUsers() {
  try {
    _requireRole(['SUPER_ADMIN']);
    var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE','PASSWORD']);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var pwdIdx = hdrs.indexOf('PASSWORD');
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {};
      for (var j = 0; j < hdrs.length; j++) {
        if (j === pwdIdx) continue; // never send passwords to client
        row[hdrs[j]] = String(vals[i][j]||'');
      }
      row.HAS_PASSWORD = pwdIdx >= 0 && String(vals[i][pwdIdx]||'').length > 0;
      rows.push(row);
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

function saveUser(data) {
  try {
    _requireRole(['SUPER_ADMIN']);
    if (!data.EMAIL) return {success:false, error:'Email required'};
    var validRoles = ['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','STAFF','EMPLOYEE'];
    if (validRoles.indexOf(data.ROLE) < 0) return {success:false, error:'Invalid role'};
    var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE','PASSWORD']);
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL');
    var pwdIdx   = hdrs.indexOf('PASSWORD');
    var emailLower = String(data.EMAIL).toLowerCase().trim();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === emailLower) {
        var row = hdrs.map(function(h){ return data[h]!==undefined ? data[h] : String(vals[i][hdrs.indexOf(h)]||''); });
        // If no new password supplied, preserve existing password
        if (!data.PASSWORD && pwdIdx >= 0) row[pwdIdx] = String(vals[i][pwdIdx]||'');
        sh.getRange(i+1, 1, 1, row.length).setValues([row]);
        try { CacheService.getScriptCache().remove('profile_' + emailLower); } catch(e2) {}
        logActivity('UserMgmt','UPDATE', emailLower, 'SUCCESS');
        return {success:true};
      }
    }
    var newPwd = data.PASSWORD || 'Admin@1234';
    sh.appendRow([emailLower, data.DISPLAY_NAME||'', data.ROLE, data.ENTITIES||'ALL', data.ACTIVE||'TRUE', newPwd]);
    logActivity('UserMgmt','ADD', emailLower, 'SUCCESS');
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; }
}

function deleteUser(email) {
  try {
    _requireRole(['SUPER_ADMIN']);
    var sh = SS.getSheetByName(TABS.USERS); if (!sh) return {success:false, error:'Users sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL'), activeIdx = hdrs.indexOf('ACTIVE');
    var emailLower = String(email).toLowerCase().trim();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === emailLower) {
        if (activeIdx >= 0) sh.getRange(i+1, activeIdx+1).setValue('FALSE');
        try { CacheService.getScriptCache().remove('profile_' + emailLower); } catch(e2) {}
        logActivity('UserMgmt','DEACTIVATE', emailLower, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'User not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function toggleUserActive(email, active) {
  try {
    _requireRole(['SUPER_ADMIN']);
    var emailLower = String(email).toLowerCase().trim();
    var callerProfile = getMyProfile();
    if (callerProfile.email === emailLower) return {success:false, error:'Cannot deactivate your own account'};
    var sh = SS.getSheetByName(TABS.USERS); if (!sh) return {success:false, error:'Users sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL'), activeIdx = hdrs.indexOf('ACTIVE');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === emailLower) {
        if (activeIdx >= 0) sh.getRange(i+1, activeIdx+1).setValue(active ? 'TRUE' : 'FALSE');
        try { CacheService.getScriptCache().remove('profile_' + emailLower); } catch(e2) {}
        logActivity('UserMgmt', active ? 'ACTIVATE' : 'DEACTIVATE', emailLower, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'User not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function resetUserPassword(email) {
  try {
    _requireRole(['SUPER_ADMIN']);
    var emailLower = String(email).toLowerCase().trim();
    var sh = SS.getSheetByName(TABS.USERS); if (!sh) return {success:false, error:'Users sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL'), pwdIdx = hdrs.indexOf('PASSWORD');
    var mcIdx = hdrs.indexOf('MUST_CHANGE');
    // Add MUST_CHANGE column if missing
    if (mcIdx < 0) {
      sh.getRange(1, hdrs.length + 1).setValue('MUST_CHANGE');
      mcIdx = hdrs.length;
      hdrs.push('MUST_CHANGE');
    }
    // Generate temp password: 3 uppercase + 3 digits + 2 special
    var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; var nums = '23456789';
    var tempPwd = '';
    for (var k = 0; k < 3; k++) tempPwd += chars.charAt(Math.floor(Math.random()*chars.length));
    for (var k = 0; k < 3; k++) tempPwd += nums.charAt(Math.floor(Math.random()*nums.length));
    tempPwd += '@' + nums.charAt(Math.floor(Math.random()*nums.length));
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === emailLower) {
        if (pwdIdx >= 0) sh.getRange(i+1, pwdIdx+1).setValue(tempPwd);
        sh.getRange(i+1, mcIdx+1).setValue('TRUE');
        try { CacheService.getScriptCache().remove('profile_' + emailLower); } catch(e2) {}
        logActivity('UserMgmt', 'RESET_PASSWORD', emailLower, 'SUCCESS');
        return {success:true, tempPassword:tempPwd};
      }
    }
    return {success:false, error:'User not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function changeMyPassword(currentPwd, newPwd) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','STAFF','EMPLOYEE']);
    if (!newPwd || String(newPwd).length < 6) return {success:false, error:'New password must be at least 6 characters'};
    var sh = SS.getSheetByName(TABS.USERS); if (!sh) return {success:false, error:'Users sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL'), pwdIdx = hdrs.indexOf('PASSWORD');
    var mcIdx = hdrs.indexOf('MUST_CHANGE');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === profile.email) {
        var stored = String(vals[i][pwdIdx]||'');
        var mustChangeFlag = mcIdx >= 0 && String(vals[i][mcIdx]||'').toUpperCase() === 'TRUE';
        // Allow bypass of current-pw check only when MUST_CHANGE is TRUE and sentinel is passed
        if (!(mustChangeFlag && String(currentPwd) === '__MUST_CHANGE__')) {
          if (stored !== String(currentPwd)) return {success:false, error:'Current password is incorrect'};
        }
        sh.getRange(i+1, pwdIdx+1).setValue(String(newPwd));
        if (mcIdx >= 0) sh.getRange(i+1, mcIdx+1).setValue('FALSE');
        try { CacheService.getScriptCache().remove('profile_' + profile.email); } catch(e2) {}
        logActivity('UserMgmt', 'CHANGE_PASSWORD', profile.email, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'User record not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function updateMyProfile(displayName) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','STAFF','EMPLOYEE']);
    if (!displayName || !String(displayName).trim()) return {success:false, error:'Name cannot be empty'};
    var sh = SS.getSheetByName(TABS.USERS); if (!sh) return {success:false, error:'Users sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL'), nameIdx = hdrs.indexOf('DISPLAY_NAME');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === profile.email) {
        sh.getRange(i+1, nameIdx+1).setValue(String(displayName).trim());
        try { CacheService.getScriptCache().remove('profile_' + profile.email); } catch(e2) {}
        logActivity('UserMgmt', 'UPDATE_PROFILE', profile.email, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'User record not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

// ============================================================
// EMPLOYEE SELF-SERVICE
// ============================================================
function getMyEmployeeRecord() {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','STAFF','EMPLOYEE']);
    return getMasterData(profile);
  } catch(e) { return {success:false, error:e.message}; }
}

function updateMyExpiryDates(data) {
  try {
    var profile = _requireRole(['EMPLOYEE']);
    var ALLOWED = ['PASSPORT EXPIRY','EID EXPIRY','VISA EXPIRY','BIRTH DATE'];
    var sh = SS.getSheetByName(TABS.MASTER); if (!sh) return {success:false, error:'Sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailCol = hdrs.indexOf('EMAIL');
    var email = profile.email.toLowerCase();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailCol]||'').toLowerCase().trim() === email) {
        ALLOWED.forEach(function(field) {
          var col = hdrs.indexOf(field);
          if (col >= 0 && data[field] !== undefined && data[field] !== '') {
            sh.getRange(i+1, col+1).setValue(data[field]);
          }
        });
        logActivity('SelfService','UPDATE_EXPIRY', email, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'Your employee record was not found. Contact HR to link your email.'};
  } catch(e) { return {success:false, error:e.message}; }
}

// Helper: find employee row by email
function _getEmpByEmail(email) {
  var sh = SS.getSheetByName(TABS.MASTER); if (!sh) return null;
  var vals = sh.getDataRange().getValues();
  var hdrs = vals[0].map(function(h){ return String(h).trim(); });
  var emailCol = hdrs.indexOf('EMAIL');
  if (emailCol < 0) return null;
  var emailLower = String(email).toLowerCase().trim();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][emailCol]||'').toLowerCase().trim() === emailLower) {
      var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
      return row;
    }
  }
  return null;
}

// ============================================================
// LEAVE MANAGEMENT
// ============================================================
function getLeave(empId) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','STAFF','EMPLOYEE']);
    var sh = getOrCreate(TABS.LEAVE, ['LEAVE_ID','EMP_ID','EMP_NAME','LEAVE_TYPE','START_DATE','END_DATE','DAYS','STATUS','APPROVED_BY','NOTES','DATE_ADDED']);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
      rows.push(row);
    }
    if (profile.role === 'EMPLOYEE') {
      var emp = _getEmpByEmail(profile.email);
      if (!emp) return {success:true, data:[]};
      rows = rows.filter(function(r){ return r.EMP_ID === emp.ID; });
    } else if (empId) {
      rows = rows.filter(function(r){ return r.EMP_ID === empId; });
    } else if (profile.entities !== 'ALL') {
      var masterRes = getMasterData(profile);
      var allowedIds = {};
      (masterRes.data||[]).forEach(function(e){ allowedIds[e.ID] = true; });
      rows = rows.filter(function(r){ return allowedIds[r.EMP_ID]; });
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

function addLeave(data) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','EMPLOYEE']);
    if (profile.role === 'EMPLOYEE') {
      var emp = _getEmpByEmail(profile.email);
      if (!emp) return {success:false, error:'No linked employee record. Contact HR.'};
      data.EMP_ID = emp.ID; data.EMP_NAME = emp.NAME;
      data.STATUS = 'Pending';
    }
    data.LEAVE_ID = genId_('LV');
    data.DATE_ADDED = formatDate(new Date());
    if (!data.STATUS) data.STATUS = 'Pending';
    var sh = getOrCreate(TABS.LEAVE, ['LEAVE_ID','EMP_ID','EMP_NAME','LEAVE_TYPE','START_DATE','END_DATE','DAYS','STATUS','APPROVED_BY','NOTES','DATE_ADDED']);
    var hdrs = ['LEAVE_ID','EMP_ID','EMP_NAME','LEAVE_TYPE','START_DATE','END_DATE','DAYS','STATUS','APPROVED_BY','NOTES','DATE_ADDED'];
    sh.appendRow(hdrs.map(function(h){ return data[h]||''; }));
    logActivity('LeaveAgent','ADD', data.LEAVE_ID+'--'+data.EMP_ID, 'SUCCESS');
    return {success:true, leaveId: data.LEAVE_ID};
  } catch(e) { return {success:false, error:e.message}; }
}

function updateLeaveStatus(leaveId, status, notes) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER']);
    var sh = SS.getSheetByName(TABS.LEAVE); if (!sh) return {success:false, error:'Leave sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var idCol = hdrs.indexOf('LEAVE_ID'), stsCol = hdrs.indexOf('STATUS');
    var apprCol = hdrs.indexOf('APPROVED_BY'), notesCol = hdrs.indexOf('NOTES');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idCol]).trim() === String(leaveId).trim()) {
        if (stsCol >= 0)  sh.getRange(i+1, stsCol+1).setValue(status);
        if (apprCol >= 0) sh.getRange(i+1, apprCol+1).setValue(profile.email);
        if (notesCol >= 0 && notes) sh.getRange(i+1, notesCol+1).setValue(notes);
        logActivity('LeaveAgent','UPDATE_STATUS', leaveId+'->'+status, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'Leave record not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

// ============================================================
// RECRUITMENT / JOB PORTAL & ATS
// ============================================================

// PUBLIC — no auth (intentional)
function getPublicJobs() {
  try {
    var sh = SS.getSheetByName(TABS.JOBS); if (!sh) return {success:true, data:[]};
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
      if (row.STATUS === 'Active') rows.push(row);
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

// Helper: get or create the CV uploads folder in Google Drive
function _getCVFolder() {
  var name = 'UG_HR_CVs';
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// PUBLIC — no auth (intentional) — upload PDF CV to Drive, return sharable URL
function uploadCV(base64Data, fileName) {
  try {
    if (!base64Data) return {success:false, error:'No file data provided'};
    var safeName = String(fileName||'cv.pdf').replace(/[^a-zA-Z0-9._\- ]/g,'_');
    var bytes  = Utilities.base64Decode(base64Data);
    var blob   = Utilities.newBlob(bytes, 'application/pdf', safeName);
    var folder = _getCVFolder();
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {success:true, url: file.getUrl()};
  } catch(e) { return {success:false, error:e.message}; }
}

// PUBLIC — no auth (intentional)
function submitApplication(data) {
  try {
    if (!data.FULL_NAME || !data.EMAIL || !data.JOB_ID) return {success:false, error:'Required fields missing'};
    var emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRe.test(data.EMAIL)) return {success:false, error:'Invalid email address'};
    data.APP_ID = genId_('APP');
    data.APPLIED_DATE = formatDate(new Date());
    data.STAGE = 'New';
    var sh = getOrCreate(TABS.APPLICATIONS, ['APP_ID','JOB_ID','JOB_TITLE','FULL_NAME','EMAIL','PHONE','NATIONALITY','PASSPORT_NO','EXPERIENCE_YEARS','CURRENT_COMPANY','COVER_NOTE','CV_DRIVE_LINK','APPLIED_DATE','STAGE','NOTES','REVIEWED_BY']);
    var hdrs = ['APP_ID','JOB_ID','JOB_TITLE','FULL_NAME','EMAIL','PHONE','NATIONALITY','PASSPORT_NO','EXPERIENCE_YEARS','CURRENT_COMPANY','COVER_NOTE','CV_DRIVE_LINK','APPLIED_DATE','STAGE','NOTES','REVIEWED_BY'];
    sh.appendRow(hdrs.map(function(h){ return data[h]||''; }));
    return {success:true, appId: data.APP_ID};
  } catch(e) { return {success:false, error:e.message}; }
}

function getJobs() {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = getOrCreate(TABS.JOBS, ['JOB_ID','TITLE','ENTITY','LOCATION','JOB_TYPE','DESCRIPTION','REQUIREMENTS','STATUS','POSTED_DATE','POSTED_BY','SALARY_RANGE']);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
      rows.push(row);
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

function saveJob(data) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = getOrCreate(TABS.JOBS, ['JOB_ID','TITLE','ENTITY','LOCATION','JOB_TYPE','DESCRIPTION','REQUIREMENTS','STATUS','POSTED_DATE','POSTED_BY','SALARY_RANGE']);
    var hdrs = ['JOB_ID','TITLE','ENTITY','LOCATION','JOB_TYPE','DESCRIPTION','REQUIREMENTS','STATUS','POSTED_DATE','POSTED_BY','SALARY_RANGE'];
    if (!data.JOB_ID) {
      data.JOB_ID = genId_('JOB');
      data.POSTED_DATE = formatDate(new Date());
      data.POSTED_BY = profile.email;
      if (!data.STATUS) data.STATUS = 'Active';
      sh.appendRow(hdrs.map(function(h){ return data[h]||''; }));
    } else {
      var vals = sh.getDataRange().getValues();
      var idCol = vals[0].map(function(h){ return String(h).trim(); }).indexOf('JOB_ID');
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]).trim() === data.JOB_ID) {
          sh.getRange(i+1, 1, 1, hdrs.length).setValues([hdrs.map(function(h){ return data[h]!==undefined?data[h]:String(vals[i][hdrs.indexOf(h)]||''); })]);
          break;
        }
      }
    }
    logActivity('RecruitmentAgent','SAVE_JOB', data.JOB_ID+'--'+data.TITLE, 'SUCCESS');
    return {success:true, jobId: data.JOB_ID};
  } catch(e) { return {success:false, error:e.message}; }
}

function closeJob(jobId) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.JOBS); if (!sh) return {success:false, error:'Jobs sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var idCol = hdrs.indexOf('JOB_ID'), stsCol = hdrs.indexOf('STATUS');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idCol]).trim() === String(jobId).trim()) {
        if (stsCol >= 0) sh.getRange(i+1, stsCol+1).setValue('Closed');
        logActivity('RecruitmentAgent','CLOSE_JOB', jobId, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'Job not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function getApplications(jobId) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = getOrCreate(TABS.APPLICATIONS, ['APP_ID','JOB_ID','JOB_TITLE','FULL_NAME','EMAIL','PHONE','NATIONALITY','PASSPORT_NO','EXPERIENCE_YEARS','CURRENT_COMPANY','COVER_NOTE','CV_DRIVE_LINK','APPLIED_DATE','STAGE','NOTES','REVIEWED_BY']);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {success:true, data:[]};
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
      if (!jobId || row.JOB_ID === jobId) rows.push(row);
    }
    return {success:true, data:rows};
  } catch(e) { return {success:false, error:e.message}; }
}

function updateApplicationStage(appId, stage, notes) {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.APPLICATIONS); if (!sh) return {success:false, error:'Applications sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var idCol = hdrs.indexOf('APP_ID'), stageCol = hdrs.indexOf('STAGE');
    var notesCol = hdrs.indexOf('NOTES'), reviewedCol = hdrs.indexOf('REVIEWED_BY');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idCol]).trim() === String(appId).trim()) {
        if (stageCol >= 0)   sh.getRange(i+1, stageCol+1).setValue(stage);
        if (notesCol >= 0 && notes) sh.getRange(i+1, notesCol+1).setValue(notes);
        if (reviewedCol >= 0) sh.getRange(i+1, reviewedCol+1).setValue(profile.email);
        logActivity('RecruitmentAgent','STAGE_UPDATE', appId+'->'+stage, 'SUCCESS');
        return {success:true};
      }
    }
    return {success:false, error:'Application not found'};
  } catch(e) { return {success:false, error:e.message}; }
}

function transferAppToOnboarding(appId) {
  try {
    _requireRole(['SUPER_ADMIN','HR_OFFICER']);
    var sh = SS.getSheetByName(TABS.APPLICATIONS); if (!sh) return {success:false, error:'Applications sheet not found'};
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var idCol = hdrs.indexOf('APP_ID'), stageCol = hdrs.indexOf('STAGE');
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][idCol]).trim() === String(appId).trim()) {
        var row = {}; for (var j = 0; j < hdrs.length; j++) row[hdrs[j]] = String(vals[i][j]||'');
        var obData = {
          OB_ID:         genId_('OB'),
          FULL_NAME:     row.FULL_NAME,
          PASSPORT_NO:   row.PASSPORT_NO||'',
          POSITION_TYPE: '',
          MOBILE:        row.PHONE||'',
          VISA_STATUS:   '',
          EXP_JOIN_DATE: '',
          NOTES:         'Transferred from ATS: '+row.JOB_TITLE+' ('+appId+')',
          ENTITY:        ''
        };
        var obResult = addOnboarding(obData);
        if (!obResult.success) return obResult;
        if (stageCol >= 0) sh.getRange(i+1, stageCol+1).setValue('Transferred');
        logActivity('RecruitmentAgent','TRANSFER_TO_OB', appId+'->'+obData.OB_ID, 'SUCCESS');
        return {success:true, obId: obData.OB_ID};
      }
    }
    return {success:false, error:'Application not found'};
  } catch(e) { return {success:false, error:e.message}; }
}
