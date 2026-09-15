// ============================================================
// Fiscal year-end close
//
// The fiscal year runs May 1 – April 30. Closing a fiscal year:
//   1. Blocks unless every reconciliation in the year is approved.
//   2. Copies the year's Reconciliations / Transactions / Receipts / AuditLog
//      rows into a "Visa Rec Archive FY XXXX-YYYY" spreadsheet stored under
//      VISA_REC_APP/Archives/FY XXXX-YYYY, then deletes them from the live
//      sheets — dashboard totals reset naturally.
//   3. Moves the year's Drive content into the archive folder, mirroring the
//      working tree:
//        Visa Statements root files       -> Archives/FY…/Visa Statements
//        Approved Visa Recs/<Month YYYY>  -> Archives/FY…/Approved Visa Recs
//        Receipts/<Dept>/<Month YYYY>     -> Archives/FY…/Receipts/<Dept>
//      Department folders themselves stay put so users' drive_folder_id keeps
//      working, and loose files in department roots stay (they may be the new
//      year's receipts waiting for the Drive scan).
//
// Safe to re-run after a timeout or partial failure: archive appends dedupe by
// row id, live-row deletion only targets rows confirmed in the archive, and
// Drive moves skip anything already moved. Transactions with recon_id
// 'HISTORICAL' and rows with a blank recon_id are never touched.
// ============================================================

var FY_MONTH_NAMES_ = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// 'FY 2025-2026' = May 1 2025 – April 30 2026.
function fyLabelFor_(endYear) {
  return 'FY ' + (endYear - 1) + '-' + endYear;
}

// End year of the most recently ENDED fiscal year: from May onward that is the
// FY that closed on April 30 this calendar year, before May it's last year's.
function mostRecentEndedFy_() {
  var now = new Date();
  return now.getMonth() >= 4 ? now.getFullYear() : now.getFullYear() - 1;
}

// Parse a month label into { m: 0-11, y } via normalizeMonth; null if unparseable.
function parseMonthLabel_(label) {
  var s = String(normalizeMonth(label)).trim();
  var match = /^([A-Za-z]+)\s+(\d{4})$/.exec(s);
  if (!match) return null;
  for (var i = 0; i < FY_MONTH_NAMES_.length; i++) {
    if (FY_MONTH_NAMES_[i].toLowerCase() === match[1].toLowerCase()) {
      return { m: i, y: Number(match[2]) };
    }
  }
  return null;
}

// True if the month falls inside the FY ending April 30 of endYear.
function monthInFy_(label, endYear) {
  var p = parseMonthLabel_(label);
  if (!p) return false;
  return (p.y === endYear - 1 && p.m >= 4) || (p.y === endYear && p.m <= 3);
}

// Strip the "24. " style ordering prefix users put on month folders.
function cleanedMonthName_(name) {
  return String(name || '').replace(/^\s*\d+[.)]?\s*/, '').trim();
}

// A statements-folder file leaves with the closing year unless its name parses
// to a month OUTSIDE the FY — so the new year's statements stay, and old or
// unrecognisable files are swept into the archive with everything else.
function statementFileLeaves_(fileName, endYear) {
  var match = STATEMENT_NAME_RE.exec(String(fileName || ''));
  if (!match) return true;
  var monthLabel = match[2].trim();
  if (!parseMonthLabel_(monthLabel)) return true;
  return monthInFy_(monthLabel, endYear);
}

// List a folder's direct children via the Drive API (Shared Drive-safe).
// kind: 'folders' or 'files' (non-folders). Returns [{ id, name }].
function fyListChildren_(parentId, kind) {
  var q = "trashed=false and '" + parentId + "' in parents and mimeType" +
          (kind === 'folders' ? "='" : "!='") + FOLDER_MIME + "'";
  var out = [];
  var pageToken = null;
  do {
    var res = Drive.Files.list({
      q: q,
      corpora: 'allDrives',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      fields: 'nextPageToken,files(id,name)',
      pageSize: 200,
      pageToken: pageToken
    });
    out = out.concat(res.files || []);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

// Shared FY scan of the Reconciliations sheet: the FY's recon ids plus any
// non-approved blockers (with user names resolved for display).
function gatherFyRecons_(endYear) {
  var data = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS).getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var uData = getOrCreateSheet_(SHEET_NAMES.USERS).getDataRange().getValues();
  var ux = colIndexes_(uData[0]);
  var userNames = {};
  for (var u = 1; u < uData.length; u++) userNames[uData[u][ux.user_id]] = uData[u][ux.name];

  var reconIds = {};
  var count = 0;
  var blockers = [];
  for (var i = 1; i < data.length; i++) {
    if (!monthInFy_(data[i][ix.month], endYear)) continue;
    reconIds[data[i][ix.recon_id]] = true;
    count++;
    if (data[i][ix.status] !== 'approved') {
      blockers.push({
        name: userNames[data[i][ix.user_id]] || String(data[i][ix.user_id]),
        month: normalizeMonth(data[i][ix.month]),
        status: data[i][ix.status]
      });
    }
  }
  return { reconIds: reconIds, count: count, blockers: blockers };
}

// Count a sheet's rows belonging to the FY recon set (preview only).
function countFyRows_(sheetName, reconIds) {
  var data = getOrCreateSheet_(sheetName).getDataRange().getValues();
  if (data.length <= 1) return 0;
  var ix = colIndexes_(data[0]);
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (reconIds[data[i][ix.recon_id]]) n++;
  }
  return n;
}

// Read-only preview for the confirmation dialog: what would the close archive,
// and is anything blocking it?
function getFiscalYearClosePreview() {
  requireRole(['finance', 'ed']);
  var endYear = mostRecentEndedFy_();
  var label = fyLabelFor_(endYear);
  var fy = gatherFyRecons_(endYear);

  var statementFiles = 0;
  try {
    var files = fyListChildren_(ensureStatementsFolder_(), 'files');
    for (var i = 0; i < files.length; i++) {
      if (statementFileLeaves_(files[i].name, endYear)) statementFiles++;
    }
  } catch (e) {
    Logger.log('Statement preview failed: ' + e.message);
  }

  return jsonSafe({
    fy_label: label,
    period: 'May 1, ' + (endYear - 1) + ' – April 30, ' + endYear,
    blockers: fy.blockers,
    counts: {
      recons: fy.count,
      transactions: countFyRows_(SHEET_NAMES.TRANSACTIONS, fy.reconIds),
      receipts: countFyRows_(SHEET_NAMES.RECEIPTS, fy.reconIds),
      audit_rows: countFyRows_(SHEET_NAMES.AUDIT_LOG, fy.reconIds),
      statement_files: statementFiles
    },
    nothing_to_archive: fy.count === 0 && statementFiles === 0
  });
}

// Find-or-create the archive spreadsheet inside the FY archive folder.
function getOrCreateArchiveSpreadsheet_(label, fyFolderId) {
  var name = 'Visa Rec Archive ' + label;
  var q = "trashed=false and '" + fyFolderId + "' in parents and name='" +
          escapeDriveQuery_(name) + "' and mimeType='application/vnd.google-apps.spreadsheet'";
  var res = Drive.Files.list({
    q: q,
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'files(id)',
    pageSize: 5
  });
  if (res.files && res.files.length > 0) return SpreadsheetApp.openById(res.files[0].id);
  var ss = SpreadsheetApp.create(name);
  moveFileToFolder_(ss.getId(), fyFolderId);
  return ss;
}

// Copy the FY rows of one live sheet into the same-named archive sheet (deduped
// by keyCol so re-runs never duplicate), then delete them from the live sheet in
// contiguous bottom-up batches. Never deletes before the archive write succeeds.
// Returns the number of rows removed from the live sheet this run.
function archiveAndDeleteRows_(archiveSs, sheetName, keyCol, reconIds) {
  var liveSheet = getOrCreateSheet_(sheetName);
  var data = liveSheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  var headers = data[0];
  var ix = colIndexes_(headers);

  var arch = archiveSs.getSheetByName(sheetName);
  if (!arch) {
    arch = archiveSs.insertSheet(sheetName);
    arch.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var archData = arch.getDataRange().getValues();
  var archIx = colIndexes_(archData[0]);
  var width = archData[0].length;
  var have = {};
  for (var a = 1; a < archData.length; a++) have[String(archData[a][archIx[keyCol]])] = true;

  var toAppend = [];
  var rowsToDelete = []; // ascending sheet row numbers
  for (var i = 1; i < data.length; i++) {
    if (!reconIds[data[i][ix.recon_id]]) continue;
    rowsToDelete.push(i + 1);
    if (have[String(data[i][ix[keyCol]])]) continue;
    var row = data[i].slice(0, width);
    while (row.length < width) row.push('');
    toAppend.push(row);
  }
  if (toAppend.length) {
    arch.getRange(arch.getLastRow() + 1, 1, toAppend.length, width).setValues(toAppend);
  }

  // Delete bottom-up in contiguous batches — per-row deleteRow is far too slow
  // for a full year of transactions.
  var d = rowsToDelete.length - 1;
  while (d >= 0) {
    var end = rowsToDelete[d];
    var start = end;
    while (d > 0 && rowsToDelete[d - 1] === start - 1) { d--; start--; }
    liveSheet.deleteRows(start, end - start + 1);
    d--;
  }
  return rowsToDelete.length;
}

// The year-end close. Finance/ED only; safe to re-run (see header comment).
function closeFiscalYear() {
  var user = requireRole(['finance', 'ed']);
  getRootFolderId_(); // fail fast with a clear message if unconfigured

  var endYear = mostRecentEndedFy_();
  var label = fyLabelFor_(endYear);
  var fy = gatherFyRecons_(endYear);
  if (fy.blockers.length) {
    return jsonSafe({ success: false, fy_label: label, blockers: fy.blockers });
  }

  // Archive Drive skeleton (idempotent find-or-create all the way down).
  var archivesId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.ARCHIVES);
  var fyFolderId = findOrCreateChildFolder_(archivesId, label);
  var archStatementsId = findOrCreateChildFolder_(fyFolderId, FOLDER_NAMES.STATEMENTS);
  var archApprovedId = findOrCreateChildFolder_(fyFolderId, FOLDER_NAMES.APPROVED);
  var archReceiptsId = findOrCreateChildFolder_(fyFolderId, FOLDER_NAMES.RECEIPTS);

  // 1) Sheet data: copy into the archive spreadsheet, then delete from live.
  var archiveSs = getOrCreateArchiveSpreadsheet_(label, fyFolderId);
  var archived = {
    recons: archiveAndDeleteRows_(archiveSs, SHEET_NAMES.RECONCILIATIONS, 'recon_id', fy.reconIds),
    transactions: archiveAndDeleteRows_(archiveSs, SHEET_NAMES.TRANSACTIONS, 'tx_id', fy.reconIds),
    receipts: archiveAndDeleteRows_(archiveSs, SHEET_NAMES.RECEIPTS, 'receipt_id', fy.reconIds),
    audit_rows: archiveAndDeleteRows_(archiveSs, SHEET_NAMES.AUDIT_LOG, 'log_id', fy.reconIds),
    statement_files: 0,
    approved_folders: 0,
    receipt_folders: 0
  };

  // 2) Visa Statements: the working folder empties into the archive; only the
  //    new fiscal year's statements stay behind.
  var stFiles = fyListChildren_(ensureStatementsFolder_(), 'files');
  for (var s = 0; s < stFiles.length; s++) {
    if (!statementFileLeaves_(stFiles[s].name, endYear)) continue;
    moveFileToFolder_(stFiles[s].id, archStatementsId);
    archived.statement_files++;
  }

  // 3) Approved Visa Recs/<Month YYYY> folders inside the FY.
  var approvedId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.APPROVED);
  var apFolders = fyListChildren_(approvedId, 'folders');
  for (var p = 0; p < apFolders.length; p++) {
    if (!monthInFy_(cleanedMonthName_(apFolders[p].name), endYear)) continue;
    moveFileToFolder_(apFolders[p].id, archApprovedId);
    archived.approved_folders++;
  }

  // 4) Receipts/<Dept>/<Month YYYY> folders inside the FY. Department folders
  //    stay in place so users' drive_folder_id keeps working. Unmapped users'
  //    folders sit one level deeper: Receipts/<ALL bucket>/<email>/<Month YYYY>.
  var receiptsId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.RECEIPTS);
  var ownerFolders = [];
  fyListChildren_(receiptsId, 'folders').forEach(function(dept) {
    ownerFolders.push({ id: dept.id, name: dept.name, archiveUnder: '' });
    if (dept.name !== getAllFolderName_()) return;
    fyListChildren_(dept.id, 'folders').forEach(function(sub) {
      if (parseMonthLabel_(cleanedMonthName_(sub.name))) return; // month folder, handled with its parent
      ownerFolders.push({ id: sub.id, name: sub.name, archiveUnder: dept.name });
    });
  });
  for (var df = 0; df < ownerFolders.length; df++) {
    var monthFolders = fyListChildren_(ownerFolders[df].id, 'folders');
    var archDeptId = null;
    for (var mf = 0; mf < monthFolders.length; mf++) {
      if (!monthInFy_(cleanedMonthName_(monthFolders[mf].name), endYear)) continue;
      if (!archDeptId) {
        var archParentId = ownerFolders[df].archiveUnder
          ? findOrCreateChildFolder_(archReceiptsId, ownerFolders[df].archiveUnder)
          : archReceiptsId;
        archDeptId = findOrCreateChildFolder_(archParentId, ownerFolders[df].name);
      }
      moveFileToFolder_(monthFolders[mf].id, archDeptId);
      archived.receipt_folders++;
    }
  }

  logAudit_(user.user_id, '', '', 'fiscal_year_closed', '',
    label + ' — recons:' + archived.recons + ' tx:' + archived.transactions +
    ' receipts:' + archived.receipts + ' audit:' + archived.audit_rows +
    ' statements:' + archived.statement_files);

  return jsonSafe({
    success: true,
    fy_label: label,
    archived: archived,
    archive_url: archiveSs.getUrl()
  });
}
