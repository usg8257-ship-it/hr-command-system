// ============================================================
//  HR COMMAND SYSTEM v2  —  Utils
// ============================================================

// ── Error handler ────────────────────────────────────────────
function handleError(e, context) {
  var msg = '[' + (context || 'GAS') + '] ' + e.message;
  console.error(msg);
  try { logAudit('SYSTEM', 'ERROR', 'SYSTEM_ERROR', '', 'error', '', msg); } catch(x) {}
  return { success: false, error: msg };
}

// ── LockService wrapper ──────────────────────────────────────
function withLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Could not acquire lock. Please retry.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ── Password hashing (SHA-256) ───────────────────────────────
function _hashPassword(plain) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

// ── Audit log ────────────────────────────────────────────────
function logAudit(userEmail, module, action, recordId, field, oldVal, newVal) {
  try {
    var sh = SS.getSheetByName(TABS.AUDIT);
    if (!sh) return;
    sh.appendRow([
      new Date(),
      userEmail || '',
      module || '',
      action || '',
      recordId || '',
      field || '',
      oldVal !== undefined ? oldVal : '',
      newVal !== undefined ? newVal : ''
    ]);
  } catch(e) {
    console.error('logAudit failed: ' + e.message);
  }
}

// ── Salary in words ──────────────────────────────────────────
function _numToWords_(n) {
  var ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
              'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
              'Seventeen','Eighteen','Nineteen'];
  var tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  if (n === 0) return 'Zero';
  if (n < 0) return 'Minus ' + _numToWords_(-n);
  var result = '';
  if (Math.floor(n / 1000000) > 0) {
    result += _numToWords_(Math.floor(n / 1000000)) + ' Million ';
    n %= 1000000;
  }
  if (Math.floor(n / 1000) > 0) {
    result += _numToWords_(Math.floor(n / 1000)) + ' Thousand ';
    n %= 1000;
  }
  if (Math.floor(n / 100) > 0) {
    result += _numToWords_(Math.floor(n / 100)) + ' Hundred ';
    n %= 100;
  }
  if (n > 0) {
    if (n < 20) {
      result += ones[n] + ' ';
    } else {
      result += tens[Math.floor(n / 10)] + ' ';
      if (n % 10 > 0) result += ones[n % 10] + ' ';
    }
  }
  return result.trim();
}

function _salaryInWords_(salary) {
  if (!salary) return '';
  var num = parseFloat(String(salary).replace(/,/g, ''));
  if (isNaN(num)) return '';
  var dirhams = Math.floor(num);
  var fils = Math.round((num - dirhams) * 100);
  var result = _numToWords_(dirhams) + ' Dirhams';
  if (fils > 0) result += ' and ' + _numToWords_(fils) + ' Fils';
  result += ' Only';
  return result;
}

// ── Summary sheet helpers ────────────────────────────────────
function updateSummaryMetric(key, value) {
  try {
    var sh = SS.getSheetByName(TABS.SUMMARY);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === key) {
          sh.getRange(i + 2, 2).setValue(value);
          sh.getRange(i + 2, 3).setValue(new Date());
          return;
        }
      }
    }
    sh.appendRow([key, value, new Date()]);
  } catch(e) {
    console.error('updateSummaryMetric failed: ' + e.message);
  }
}

function recalcSummary() {
  var masterSh = SS.getSheetByName(TABS.MASTER);
  var activeCnt = masterSh ? Math.max(0, masterSh.getLastRow() - 1) : 0;
  updateSummaryMetric('TOTAL_ACTIVE', activeCnt);

  var inactiveSh = SS.getSheetByName(TABS.INACTIVE);
  var inactiveCnt = inactiveSh ? Math.max(0, inactiveSh.getLastRow() - 1) : 0;
  updateSummaryMetric('TOTAL_INACTIVE', inactiveCnt);

  var trackerSh = SS.getSheetByName(TABS.TRACKER);
  if (trackerSh && trackerSh.getLastRow() > 1) {
    var thm = getHeaderMap(TABS.TRACKER);
    var statCol = thm.map['STATUS'];
    if (statCol) {
      var statuses = trackerSh.getRange(2, statCol, trackerSh.getLastRow() - 1, 1).getValues();
      var completed = statuses.filter(function(r) { return r[0] === 'Completed'; }).length;
      var active    = statuses.filter(function(r) { return r[0] === 'In Progress'; }).length;
      var pending   = statuses.filter(function(r) { return r[0] === 'Pending HR Approval'; }).length;
      updateSummaryMetric('20DS_TOTAL',     statuses.length);
      updateSummaryMetric('20DS_COMPLETED', completed);
      updateSummaryMetric('20DS_ACTIVE',    active);
      updateSummaryMetric('20DS_PENDING',   pending);
    }
  }
}
