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
  APPLICATIONS: 'Applications'
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
function getMyProfile() {
  var email = '';
  try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e2) {}
  if (!email) return {email:'',name:'',role:'NONE',entities:[],active:false};

  // Cache per-email for 60 seconds to reduce sheet reads
  var cache = CacheService.getScriptCache();
  var cacheKey = 'profile_' + email;
  var cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch(e2) {} }

  // Ensure Users sheet exists
  var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE']);
  var vals = sh.getDataRange().getValues();

  // Bootstrap: if only the header row exists, make first user SUPER_ADMIN
  if (vals.length === 1) {
    var lock = LockService.getScriptLock();
    lock.tryLock(3000);
    try {
      // Re-read inside lock to prevent race
      var vals2 = sh.getDataRange().getValues();
      if (vals2.length === 1) {
        sh.appendRow([email, email.split('@')[0], 'SUPER_ADMIN', 'ALL', 'TRUE']);
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
      var profile = {email:email, name:name||email.split('@')[0], role:role, entities:entities, active:true};
      cache.put(cacheKey, JSON.stringify(profile), 60);
      return profile;
    }
  }
  return {email:email,name:'',role:'NONE',entities:[],active:false};
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
function loadAllData() {
  var profile = getMyProfile();
  return {
    user:       profile,
    master:     getMasterData(profile),
    deletions:  getDeletionLog(),
    onboarding: getOnboarding(profile),
    hrDocs:     getHRDocs(profile),
    summary:    getSummary(),
    config:     getConfig()
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

function deleteEmployee(empId, reason) {
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
        var delSh=getOrCreate(TABS.DEL_LOG,['LOG_ID','EMP_ID','FULL_NAME','PASSPORT_NO','REASON','DELETED_DATE','DELETED_BY','DESIGNATION','DATE_OF_JOIN','GROUP']);
        delSh.appendRow([
          'DEL-'+new Date().getTime(), empId,
          vals[i][nmCol]||'', vals[i][ppCol]||'',
          reason, formatDate(new Date()), deletedBy,
          desigCol>=0 ? vals[i][desigCol]||'' : '',
          joinCol>=0  ? vals[i][joinCol]||''  : '',
          entCol>=0   ? normaliseEntity(vals[i][entCol]) : ''
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
    var sh=getOrCreate(TABS.ONBOARDING,['OB_ID','FULL_NAME','PASSPORT_NO','POSITION_TYPE','MOBILE','VISA_STATUS','EXP_JOIN_DATE','DATE_ADDED','STATUS','NOTES','ENTITY']);
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
    var sh=getOrCreate(TABS.ONBOARDING,['OB_ID','FULL_NAME','PASSPORT_NO','POSITION_TYPE','MOBILE','VISA_STATUS','EXP_JOIN_DATE','DATE_ADDED','STATUS','NOTES']);
    sh.appendRow([data.OB_ID,data.FULL_NAME,data.PASSPORT_NO,data.POSITION_TYPE,data.MOBILE||'',data.VISA_STATUS||'',data.EXP_JOIN_DATE||'',formatDate(new Date()),'Pending',data.NOTES||'']);
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
    if(sh){
      var vals=sh.getDataRange().getValues();
      var hdrs=vals[0].map(function(h){ return String(h).trim(); });
      var idCol=hdrs.indexOf('OB_ID'), stsCol=hdrs.indexOf('STATUS');
      for(var i=1;i<vals.length;i++){
        if(String(vals[i][idCol])===String(obId)){ if(stsCol>=0) sh.getRange(i+1,stsCol+1).setValue('Transferred'); break; }
      }
    }
    logActivity('OnboardingAgent','TRANSFER',obId+'->'+empData.ID,'SUCCESS');
    return {success:true};
  } catch(e){ return {success:false,error:e.message}; }
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

function _fillPlaceholders(text, data, cfg) {
  var map = {
    '{{NAME}}':           data.EMP_NAME    || '',
    '{{EMP_ID}}':         data.EMP_ID      || '',
    '{{PASSPORT_NO}}':    data.PASSPORT_NO || '',
    '{{DESIGNATION}}':    data.DESIGNATION || '',
    '{{COMPANY}}':        cfg.company_name || 'United Group Holding',
    '{{ENTITY}}':         data.ENTITY      || '',
    '{{DATE_OF_JOIN}}':   data.DATE_OF_JOIN|| '',
    '{{SALARY}}':         data.SALARY      || '',
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
    var placeholders = {
      '{{NAME}}':           data.EMP_NAME    || '',
      '{{EMP_ID}}':         data.EMP_ID      || '',
      '{{PASSPORT_NO}}':    data.PASSPORT_NO || '',
      '{{DESIGNATION}}':    data.DESIGNATION || '',
      '{{COMPANY}}':        cfg.company_name || 'United Group Holding',
      '{{ENTITY}}':         data.ENTITY      || '',
      '{{DATE_OF_JOIN}}':   data.DATE_OF_JOIN|| '',
      '{{SALARY}}':         data.SALARY      || '',
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
    var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE']);
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

function saveUser(data) {
  try {
    _requireRole(['SUPER_ADMIN']);
    if (!data.EMAIL) return {success:false, error:'Email required'};
    var validRoles = ['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','VIEWER','EMPLOYEE'];
    if (validRoles.indexOf(data.ROLE) < 0) return {success:false, error:'Invalid role'};
    var sh = getOrCreate(TABS.USERS, ['EMAIL','DISPLAY_NAME','ROLE','ENTITIES','ACTIVE']);
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var emailIdx = hdrs.indexOf('EMAIL');
    var emailLower = String(data.EMAIL).toLowerCase().trim();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][emailIdx]||'').toLowerCase().trim() === emailLower) {
        sh.getRange(i+1, 1, 1, hdrs.length).setValues([hdrs.map(function(h){ return data[h]!==undefined?data[h]:''; })]);
        // Invalidate cache
        try { CacheService.getScriptCache().remove('profile_' + emailLower); } catch(e2) {}
        logActivity('UserMgmt','UPDATE', emailLower, 'SUCCESS');
        return {success:true};
      }
    }
    sh.appendRow([emailLower, data.DISPLAY_NAME||'', data.ROLE, data.ENTITIES||'ALL', data.ACTIVE||'TRUE']);
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

// ============================================================
// EMPLOYEE SELF-SERVICE
// ============================================================
function getMyEmployeeRecord() {
  try {
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','VIEWER','EMPLOYEE']);
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
    var profile = _requireRole(['SUPER_ADMIN','HR_OFFICER','ENTITY_MANAGER','VIEWER','EMPLOYEE']);
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
