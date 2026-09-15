// ============================================================
// Story 3.3 — Reporting (Finance and ED only)
// ============================================================

function getStaffUsers() {
  assertReviewer();
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  return jsonSafe(data.slice(1)
    .filter(function(r) { return r[ix.active]; })
    .map(function(r) {
      return { user_id: r[ix.user_id], name: r[ix.name] };
    }));
}

function assertReviewer() {
  var user = getCurrentUser();
  if (user.role !== 'finance' && user.role !== 'ed' && user.role !== 'exec') {
    throw new Error('Reports are available to Finance and ED only.');
  }
  return user;
}

function getApprovedRecons_(fromDate) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var ix = colIndexes_(headers);
  var results = [];
  var from = fromDate ? new Date(fromDate) : null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.status] !== 'approved') continue;
    if (from && new Date(data[i][ix.ed_approved_at]) < from) continue;
    results.push(rowToObj_(headers, data[i]));
  }
  return results;
}

function getMonthlyReport(month) {
  assertReviewer();

  var target = normalizeMonth(month);
  var recons = getApprovedRecons_().filter(function(r) { return normalizeMonth(r.month) === target; });
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);

  // Get budget code suffix → name map
  var codeSheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);
  var codeData = codeSheet.getDataRange().getValues();
  var ccx = colIndexes_(codeData[0]);
  var suffixNames = {};
  for (var c = 1; c < codeData.length; c++) {
    var suffix = codeData[c][ccx.department_suffix];
    if (!suffixNames[suffix]) suffixNames[suffix] = codeData[c][ccx.description];
  }

  var reconIds = {};
  recons.forEach(function(r) { reconIds[r.recon_id] = r; });

  var depts = {};
  for (var t = 1; t < txData.length; t++) {
    var rid = txData[t][tcx.recon_id];
    if (!reconIds[rid]) continue;
    if (!txNeedsCode_(txTypeOf_(txData[t][tcx.tx_type]))) continue; // payments and refunds have no budget code
    var code = String(txData[t][tcx.budget_code] || '');
    var suffix = code.split('-')[1] || 'unknown';
    if (!depts[suffix]) depts[suffix] = { department_suffix: suffix, total_spend: 0, tx_count: 0, recon_ids: {} };
    depts[suffix].total_spend += txData[t][tcx.total] || 0;
    depts[suffix].tx_count++;
    depts[suffix].recon_ids[rid] = true;
  }

  return jsonSafe(Object.keys(depts).map(function(s) {
    return {
      department_suffix: s,
      department_name: suffixNames[s] || s,
      total_spend: depts[s].total_spend,
      tx_count: depts[s].tx_count,
      reconciliation_count: Object.keys(depts[s].recon_ids).length
    };
  }).sort(function(a, b) { return b.total_spend - a.total_spend; }));
}

function getYTDReport(fiscalYearStart) {
  assertReviewer();

  var recons = getApprovedRecons_(fiscalYearStart);
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);

  var codeSheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);
  var codeData = codeSheet.getDataRange().getValues();
  var ccx = colIndexes_(codeData[0]);
  var codeMeta = {};
  for (var c = 1; c < codeData.length; c++) {
    codeMeta[codeData[c][ccx.code]] = {
      description: codeData[c][ccx.description],
      department_suffix: codeData[c][ccx.department_suffix]
    };
  }

  var reconIds = {};
  recons.forEach(function(r) { reconIds[r.recon_id] = true; });

  var codes = {};
  for (var t = 1; t < txData.length; t++) {
    if (!reconIds[txData[t][tcx.recon_id]]) continue;
    if (!txNeedsCode_(txTypeOf_(txData[t][tcx.tx_type]))) continue; // payments and refunds have no budget code
    var code = String(txData[t][tcx.budget_code] || '');
    if (!codes[code]) codes[code] = { code: code, total_spend_ytd: 0, tx_count: 0 };
    codes[code].total_spend_ytd += txData[t][tcx.total] || 0;
    codes[code].tx_count++;
  }

  return jsonSafe(Object.keys(codes).map(function(code) {
    var meta = codeMeta[code] || { description: '', department_suffix: '' };
    return {
      code: code,
      description: meta.description,
      department_suffix: meta.department_suffix,
      total_spend_ytd: codes[code].total_spend_ytd,
      tx_count: codes[code].tx_count
    };
  }).sort(function(a, b) { return b.total_spend_ytd - a.total_spend_ytd; }));
}

function getStaffReport(userId) {
  assertReviewer();

  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var rcptSheet = getOrCreateSheet_(SHEET_NAMES.RECEIPTS);

  var reconData = reconSheet.getDataRange().getValues();
  var rccx = colIndexes_(reconData[0]);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);
  var rcptData = rcptSheet.getDataRange().getValues();
  var rcx = colIndexes_(rcptData[0]);

  // tx counts per recon
  var txStats = {};
  for (var t = 1; t < txData.length; t++) {
    var rid = txData[t][tcx.recon_id];
    if (!txStats[rid]) txStats[rid] = { count: 0 };
    txStats[rid].count++;
  }

  // receipt stats per recon
  var rcptStats = {};
  for (var r = 1; r < rcptData.length; r++) {
    var rid2 = rcptData[r][rcx.recon_id];
    if (!rcptStats[rid2]) rcptStats[rid2] = { matched: 0, unmatched: 0 };
    if (rcptData[r][rcx.match_status] === 'matched') rcptStats[rid2].matched++;
    else rcptStats[rid2].unmatched++;
  }

  var results = [];
  for (var i = 1; i < reconData.length; i++) {
    if (reconData[i][rccx.user_id] !== userId) continue;
    var rId = reconData[i][rccx.recon_id];
    var stats = txStats[rId] || { count: 0 };
    var rStats = rcptStats[rId] || { matched: 0, unmatched: 0 };
    results.push({
      recon_id: rId,
      month: reconData[i][rccx.month],
      status: reconData[i][rccx.status],
      reconciled_total: reconData[i][rccx.reconciled_total],
      tx_count: stats.count,
      receipts_matched: rStats.matched,
      receipts_missing: rStats.unmatched
    });
  }

  results.sort(function(a, b) { return String(b.month).localeCompare(String(a.month)); });
  return jsonSafe(results);
}
