// ============================================================
//  HR COMMAND SYSTEM v2  —  Auth
// ============================================================

var _SESSION_EMAIL = '';
var _SESSION_PROFILE = null;

// ── Login ────────────────────────────────────────────────────
function loginUser(email, password) {
  try {
    var hm = getHeaderMap(TABS.USERS);
    var sh = hm.sheet;
    var lastRow = hm.lastRow;
    if (lastRow < 2) return { success: false, error: 'No users configured.' };

    var emailCol   = hm.map['email'];
    var hashCol    = hm.map['password_hash'];
    var nameCol    = hm.map['display_name'];
    var roleCol    = hm.map['role'];
    var entCol     = hm.map['entities'];
    var activeCol  = hm.map['is_active'];
    var loginCol   = hm.map['last_login'];

    var rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var found = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][emailCol - 1]).trim().toLowerCase() === email.trim().toLowerCase()) {
        found = { row: rows[i], rowNum: i + 2 };
        break;
      }
    }
    if (!found) return { success: false, error: 'Account not found.' };

    var r = found.row;
    if (activeCol && String(r[activeCol - 1]).toUpperCase() !== 'TRUE' && r[activeCol - 1] !== true) {
      return { success: false, error: 'Account is inactive. Contact administrator.' };
    }

    var storedHash = String(r[hashCol - 1]).trim();
    // Migration: if stored value is not a 64-char hex, treat as plain-text first login
    if (storedHash.length < 64) {
      if (password !== storedHash) return { success: false, error: 'Invalid password.' };
      // Upgrade to hash
      sh.getRange(found.rowNum, hashCol).setValue(_hashPassword(password));
    } else {
      if (_hashPassword(password) !== storedHash) return { success: false, error: 'Invalid password.' };
    }

    // Update last login
    if (loginCol) sh.getRange(found.rowNum, loginCol).setValue(new Date());

    // Issue token
    var token = Utilities.getUuid();
    var profile = {
      email:    String(r[emailCol - 1]).trim(),
      name:     nameCol ? String(r[nameCol - 1]).trim() : email,
      role:     roleCol ? String(r[roleCol - 1]).trim() : 'VIEWER',
      entities: entCol  ? String(r[entCol - 1]).trim()  : 'ALL'
    };
    var cache = CacheService.getScriptCache();
    cache.put('v2_sess_' + token, JSON.stringify(profile), 28800); // 8 hours

    return { success: true, token: token, profile: profile };
  } catch(e) {
    return handleError(e, 'loginUser');
  }
}

// ── Validate session ─────────────────────────────────────────
function validateSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var raw = cache.get('v2_sess_' + token);
  if (!raw) return null;
  try {
    var profile = JSON.parse(raw);
    // Refresh TTL
    cache.put('v2_sess_' + token, raw, 28800);
    return profile;
  } catch(e) {
    return null;
  }
}

// ── Logout ───────────────────────────────────────────────────
function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('v2_sess_' + token);
  return { success: true };
}

// ── Set session from token (server-side) ─────────────────────
function _setSessionFromToken(token) {
  var profile = validateSession(token);
  if (!profile) throw new Error('Session expired. Please log in again.');
  _SESSION_EMAIL   = profile.email;
  _SESSION_PROFILE = profile;
}

// ── Role guard ───────────────────────────────────────────────
function _requireRole(allowedRoles) {
  if (!_SESSION_PROFILE) throw new Error('Not authenticated.');
  var role = _SESSION_PROFILE.role;
  if (allowedRoles.indexOf(role) === -1) {
    throw new Error('Access denied. Required role: ' + allowedRoles.join(' or '));
  }
}

// ── runProtected dispatcher ──────────────────────────────────
var _FUNC_MAP = {
  // Auth
  'logoutUser':            function(a) { return logoutUser(a[0]); },

  // Config
  'getConfig':             function()  { return getConfig(); },
  'saveConfig':            function(a) { return saveConfig(a[0], a[1]); },

  // Users (Admin_Controller.gs)
  'getUsers':              function(a) { return getUsers(a[0]); },
  'saveUser':              function(a) { return saveUser(a[0], a[1]); },
  'deleteUser':            function(a) { return deleteUser(a[0], a[1]); },

  // preOnBoard (Lifecycle_Manager.gs)
  'getPreOnboard':         function(a) { return getPreOnboard(a[0]); },
  'addPreOnboard':         function(a) { return addPreOnboard(a[0], a[1]); },
  'updatePreOnboard':      function(a) { return updatePreOnboard(a[0], a[1]); },
  'deletePreOnboard':      function(a) { return deletePreOnboard(a[0], a[1]); },
  'transferToMaster':      function(a) { return transferToMaster(a[0], a[1], a[2]); },

  // Master Data (Lifecycle_Manager.gs)
  'getMasterData':         function(a) { return getMasterData(a[0]); },
  'updateEmployee':        function(a) { return updateEmployee(a[0], a[1]); },
  'resignEmployee':        function(a) { return resignEmployee(a[0], a[1], a[2]); },
  'getInactiveData':       function(a) { return getInactiveData(a[0]); },

  // Strategy Tracker (Strategy_Controller.gs)
  'getTrackerData':        function(a) { return getTrackerData(a[0]); },
  'getStepsConfig':        function(a) { return getStepsConfig(a[0]); },
  'saveStepsConfig':       function(a) { return saveStepsConfig(a[0], a[1]); },
  'completeStep':          function(a) { return completeStep(a[0], a[1], a[2]); },
  'approveTracker':        function(a) { return approveTracker(a[0], a[1]); },
  'get20DSDashboard':      function(a) { return get20DSDashboard(a[0]); },

  // Notifications (Notification_Manager.gs)
  'getNotifications':      function(a) { return getNotifications(a[0]); },
  'markNotifRead':         function(a) { return markNotifRead(a[0], a[1]); },
  'markAllNotifsRead':     function(a) { return markAllNotifsRead(a[0]); },

  // Letters (Doc_Factory.gs)
  'generateLetter':        function(a) { return generateLetter(a[0], a[1]); },
  'getHRDocs':             function(a) { return getHRDocs(a[0]); },

  // Analytics (Analytics_Controller.gs)
  'getAnalytics':          function(a) { return getAnalytics(a[0]); },

  // Jobs & Applications (Recruitment)
  'getJobs':               function(a) { return getJobs(a[0]); },
  'saveJob':               function(a) { return saveJob(a[0], a[1]); },
  'closeJob':              function(a) { return closeJob(a[0], a[1]); },
  'getApplications':       function(a) { return getApplications(a[0], a[1]); },
  'updateApplicationStage':function(a) { return updateApplicationStage(a[0], a[1], a[2]); },
  'transferAppToOnboard':  function(a) { return transferAppToOnboard(a[0], a[1]); },

  // Leave
  'getLeave':              function(a) { return getLeave(a[0]); },
  'addLeave':              function(a) { return addLeave(a[0], a[1]); },
  'updateLeaveStatus':     function(a) { return updateLeaveStatus(a[0], a[1], a[2]); },

  // Init
  'initSheets':            function()  { return initSheets(); }
};

function runProtected(token, funcName, args) {
  try {
    _setSessionFromToken(token);
    var fn = _FUNC_MAP[funcName];
    if (!fn) return { success: false, error: 'Unknown function: ' + funcName };
    var result = fn(args || []);
    return (result && typeof result === 'object' && 'success' in result)
      ? result
      : { success: true, data: result };
  } catch(e) {
    return handleError(e, 'runProtected:' + funcName);
  }
}

// ── My profile ───────────────────────────────────────────────
function getMyProfile(token) {
  var profile = validateSession(token);
  if (!profile) return { success: false, error: 'Session expired.' };
  return { success: true, data: profile };
}
