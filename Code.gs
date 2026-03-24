// ============================================================
// UNITED GROUP HOLDING — HR COMMAND SYSTEM v4
// Code.gs — Google Apps Script Backend
// Concept, Designed & Developed by Mohammad Sanish
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

var TABS = {
  MASTER:     'Master Data',
  ONBOARDING: 'Onboarding',
  HR_DOCS:    'HR Docs Tracker',
  DEL_LOG:    'Deletion_Log',
  ACTIVE_MP:  'Active_Manpower',
  SUMMARY:    'Summary',
  CONFIG:     'AppConfig',
  ACTIVITY:   'ActivityLog'
};

// ============================================================
// WEB APP ENTRY
// ============================================================
function doGet(e) {
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
// LOAD ALL DATA
// ============================================================
function loadAllData() {
  return {
    master:     getMasterData(),
    deletions:  getDeletionLog(),
    onboarding: getOnboarding(),
    hrDocs:     getHRDocs(),
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
function getMasterData() {
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

function deleteEmployee(empId, reason, deletedBy) {
  try {
    var sh=SS.getSheetByName(TABS.MASTER); if(!sh) return {success:false,error:'Sheet not found'};
    var vals=sh.getDataRange().getValues();
    var hdrs=vals[0].map(function(h){ return String(h).trim(); });
    var idCol=hdrs.indexOf('ID'), stsCol=hdrs.indexOf('STATUS');
    var ppCol=hdrs.indexOf('PASSPORT NO'), nmCol=hdrs.indexOf('NAME');
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][idCol]).trim()===String(empId).trim()){
        var delSh=getOrCreate(TABS.DEL_LOG,['LOG_ID','EMP_ID','FULL_NAME','PASSPORT_NO','REASON','DELETED_DATE','DELETED_BY']);
        delSh.appendRow(['DEL-'+new Date().getTime(),empId,vals[i][nmCol]||'',vals[i][ppCol]||'',reason,formatDate(new Date()),deletedBy||'HR']);
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

// ============================================================
// ONBOARDING
// ============================================================
function getOnboarding() {
  try {
    var sh=getOrCreate(TABS.ONBOARDING,['OB_ID','FULL_NAME','PASSPORT_NO','POSITION_TYPE','MOBILE','VISA_STATUS','EXP_JOIN_DATE','DATE_ADDED','STATUS','NOTES']);
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
function getHRDocs() {
  try {
    var sh=getOrCreate(TABS.HR_DOCS,['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES']);
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
    var sh=getOrCreate(TABS.ACTIVITY,['TIMESTAMP','AGENT','ACTION','DETAIL','STATUS']);
    sh.appendRow([new Date().toLocaleString(),agent,action,detail||'',status||'SUCCESS']);
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
    var sh = getOrCreate(TABS.HR_DOCS, ['REF_NO','EMP_ID','EMP_NAME','LETTER_TYPE','ISSUE_DATE','ISSUED_BY','NOTES']);
    sh.appendRow([data.REF_NO, data.EMP_ID, data.EMP_NAME, data.LETTER_TYPE,
                  data.ISSUE_DATE, data.ISSUED_BY||'HR', data.NOTES||'']);
    logActivity('LetterAgent', 'ISSUE', data.REF_NO + '--' + data.LETTER_TYPE, 'SUCCESS');
    return { success: true, pdf: pdfBase64 };
  } catch(e) { return { success: false, error: e.message }; }
}
