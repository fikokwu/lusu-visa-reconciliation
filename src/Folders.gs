// ============================================================
// Shared Drive folder provisioning
//
// Builds and maintains the folder tree inside the VISA_REC_APP parent folder
// so admins never have to create folders or copy Drive IDs by hand:
//
//   VISA_REC_APP  (ROOT_FOLDER_ID)
//   ├── Visa Statements
//   │   └── <Month YYYY>          (created on demand; filing stays manual)
//   ├── Receipts
//   │   └── <Department>          (e.g. VPFIN, VPO — from DEPARTMENT_MAP; unmapped
//   │                              users get Admin (ALL)/<email>)
//   │       └── <Month YYYY>      (receipts filed here)
//   └── Approved Visa Recs
//       └── <Month YYYY>          (approved reconciliation exports archived here)
//
// All Drive operations go through the Advanced Drive Service (Drive.Files.*)
// with supportsAllDrives/includeItemsFromAllDrives so they work reliably inside
// a Shared Drive — plain DriveApp create/move is unreliable there.
// The web app runs as USER_DEPLOYING, so the deploying owner must be a member of
// the Shared Drive with at least Content Manager access.
// ============================================================

var APP_ROOT_NAME = 'VISA_REC_APP';

var FOLDER_NAMES = {
  RECEIPTS: 'Receipts',
  STATEMENTS: 'Visa Statements',
  APPROVED: 'Approved Visa Recs',
  ARCHIVES: 'Archives'
};

var FOLDER_MIME = 'application/vnd.google-apps.folder';

// The parent VISA_REC_APP folder that everything is created under.
function getRootFolderId_() {
  var id = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
  if (!id) {
    throw new Error('ROOT_FOLDER_ID script property is not set. Set it to the ID of the ' +
      'VISA_REC_APP parent folder in the Shared Drive.');
  }
  return id;
}

// Resolve the effective VISA_REC_APP root. ROOT_FOLDER_ID may point either at the
// VISA_REC_APP folder itself or at its parent (e.g. the Shared Drive root). If the
// configured folder isn't already named VISA_REC_APP, find-or-create that folder
// underneath it so the tree is always SharedDrive/VISA_REC_APP/... as intended.
var _appRootCache = null;
function getAppRootId_() {
  if (_appRootCache) return _appRootCache;
  var configured = getRootFolderId_();
  var name = '';
  try {
    name = Drive.Files.get(configured, { supportsAllDrives: true, fields: 'name' }).name || '';
  } catch (e) {
    Logger.log('Could not read ROOT_FOLDER_ID metadata: ' + e.message);
  }
  _appRootCache = (name === APP_ROOT_NAME)
    ? configured
    : findOrCreateChildFolder_(configured, APP_ROOT_NAME);
  return _appRootCache;
}

// Map of user email -> receipt folder name, supplied by the admin as a JSON script
// property, e.g. {"president@example.org":"President","events@example.org":"Events"}.
// Keyed by email (not department_suffix) because several people share a budget-code
// suffix (e.g. all execs are 1500) but each needs their own folder. {} if unset.
function getDepartmentMap_() {
  var raw = PropertiesService.getScriptProperties().getProperty('DEPARTMENT_MAP');
  if (!raw) return {};
  try {
    var parsed = JSON.parse(raw);
    return (parsed && typeof parsed === 'object') ? parsed : {};
  } catch (e) {
    Logger.log('DEPARTMENT_MAP is not valid JSON: ' + e.message);
    return {};
  }
}

// Folder name to hold users with no mapped department (unmapped/auto-provisioned).
function getAllFolderName_() {
  return PropertiesService.getScriptProperties().getProperty('ALL_FOLDER_NAME') || 'Admin (ALL)';
}

// Case-insensitive lookup of a folder name for an email; '' if not mapped.
function folderNameByEmail_(email) {
  var target = String(email || '').trim().toLowerCase();
  if (!target) return '';
  var map = getDepartmentMap_();
  for (var k in map) {
    if (map.hasOwnProperty(k) && String(k).trim().toLowerCase() === target) return map[k];
  }
  return '';
}

// Escape a value for use inside a Drive query string literal (name='...').
function escapeDriveQuery_(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

// Sanitize a name into something safe/tidy for a Drive folder.
function sanitizeFolderName_(name) {
  return String(name || '').replace(/[\/\\]/g, '-').replace(/\s+/g, ' ').trim();
}

// Idempotent find-or-create of a child folder by name under a given parent.
// Searches first (so re-runs never duplicate) then creates. Returns the folder ID.
function findOrCreateChildFolder_(parentId, rawName) {
  var name = sanitizeFolderName_(rawName);
  var q = "mimeType='" + FOLDER_MIME + "' and trashed=false and '" +
          parentId + "' in parents and name='" + escapeDriveQuery_(name) + "'";
  var res = Drive.Files.list({
    q: q,
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    orderBy: 'createdTime', // deterministic pick if same-name duplicates ever exist
    fields: 'files(id,name)',
    pageSize: 10
  });
  if (res.files && res.files.length > 0) return res.files[0].id;

  var created = Drive.Files.create({
    name: name,
    mimeType: FOLDER_MIME,
    parents: [parentId]
  }, null, { supportsAllDrives: true });
  return created.id;
}

// Ensure the user's receipts folder exists and return its ID (what gets stored in
// their drive_folder_id). Mapped users get VISA_REC_APP/Receipts/<department>,
// a clean position name like "President" or "VPO". Everyone else gets their own
// Receipts/Admin (ALL)/<email>, so receipt scanning never mixes people's receipts.
// `user` is an object with at least { email }.
function provisionUserReceiptsFolder_(user) {
  // Statement filing is manual, so "Visa Statements" has no on-demand trigger —
  // materialise it alongside the receipts tree whenever provisioning runs.
  // Best-effort: a failure on this tangential folder must never block the
  // essential receipts provisioning (memoised inside ensureStatementsFolder_,
  // so batch callers like backfill/repair don't re-check it per user).
  try { ensureStatementsFolder_(); }
  catch (e) { Logger.log('Visa Statements folder ensure failed: ' + e.message); }
  var receiptsId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.RECEIPTS);
  var email = String((user && user.email) || '').trim().toLowerCase();
  var mapped = folderNameByEmail_(email);
  if (mapped) return findOrCreateChildFolder_(receiptsId, mapped);
  var allId = findOrCreateChildFolder_(receiptsId, getAllFolderName_());
  return email ? findOrCreateChildFolder_(allId, email) : allId;
}

// Ensure a "<Month YYYY>" subfolder exists under a given parent; returns its ID.
// Falls back to the parent itself if the month can't be normalised.
function ensureMonthSubfolder_(parentFolderId, monthLabel) {
  var month = normalizeMonth(monthLabel);
  if (!month) return parentFolderId;
  return findOrCreateChildFolder_(parentFolderId, month);
}

// Ensure VISA_REC_APP/Visa Statements exists; returns its ID. Filing is manual.
// Memoised per execution so batch provisioning doesn't re-query Drive per user.
var _statementsFolderCache = null;
function ensureStatementsFolder_() {
  if (!_statementsFolderCache) {
    _statementsFolderCache = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.STATEMENTS);
  }
  return _statementsFolderCache;
}

// Ensure VISA_REC_APP/Approved Visa Recs/<Month YYYY> exists; returns its ID.
// This is the central archive for finalised reconciliations.
function ensureApprovedMonthFolder_(monthLabel) {
  var approvedId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.APPROVED);
  return ensureMonthSubfolder_(approvedId, monthLabel);
}

// Trash any existing non-folder files with the given name directly under a folder.
// Used before archiving a fresh export so re-approvals don't pile up duplicates.
function trashExistingByName_(folderId, rawName) {
  var name = sanitizeFolderName_(rawName);
  var q = "trashed=false and mimeType!='" + FOLDER_MIME + "' and '" +
          folderId + "' in parents and name='" + escapeDriveQuery_(name) + "'";
  var res = Drive.Files.list({
    q: q,
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'files(id)',
    pageSize: 50
  });
  var files = res.files || [];
  for (var i = 0; i < files.length; i++) {
    Drive.Files.update({ trashed: true }, files[i].id, null, { supportsAllDrives: true });
  }
}

// Create a file from a blob inside a Shared Drive folder (reliable path).
// Returns a DriveApp File handle so callers keep .getId()/.getUrl().
function createFileInFolder_(folderId, blob, fileName) {
  var created = Drive.Files.create({
    name: fileName,
    parents: [folderId]
  }, blob, { supportsAllDrives: true });
  return DriveApp.getFileById(created.id);
}

// Grant a user direct access to a Drive item (folder or file) in the Shared
// Drive. Best-effort: returns true on success, false on failure (logged), so
// callers never fail user creation over an ACL error. Idempotent: the Drive
// API upserts an equal-or-greater existing permission, so re-runs are safe.
// Prerequisite: shared drives default to "Only managers can share folders",
// so the deploying account must be a Manager of the Shared Drive (or that
// drive setting relaxed) for folder grants to succeed. Per-file grants only
// need Content Manager.
function grantDriveAccess_(itemId, email, role) {
  var addr = String(email || '').trim().toLowerCase();
  if (!itemId || !addr) return false;
  try {
    Drive.Permissions.create(
      { type: 'user', role: role || 'writer', emailAddress: addr },
      itemId,
      { supportsAllDrives: true, sendNotificationEmail: false }
    );
    return true;
  } catch (e) {
    Logger.log('Drive access grant failed (' + (role || 'writer') + ' on ' + itemId + ' for ' + addr + '): ' + e.message);
    return false;
  }
}

// Remove a user's direct permission on a Drive item. Best-effort; only the
// matching user's permission is touched (others sharing the folder keep theirs).
function revokeDriveAccess_(itemId, email) {
  var addr = String(email || '').trim().toLowerCase();
  if (!itemId || !addr) return false;
  try {
    var res = Drive.Permissions.list(itemId, {
      supportsAllDrives: true,
      fields: 'permissions(id,emailAddress)'
    });
    var perms = (res.permissions || []).filter(function(p) {
      return String(p.emailAddress || '').toLowerCase() === addr;
    });
    perms.forEach(function(p) {
      Drive.Permissions.remove(itemId, p.id, { supportsAllDrives: true });
    });
    return perms.length > 0;
  } catch (e) {
    Logger.log('Drive access revoke failed (' + itemId + ' for ' + addr + '): ' + e.message);
    return false;
  }
}

// Reparent an existing Drive file into a Shared Drive folder (reliable move).
function moveFileToFolder_(fileId, folderId) {
  var meta = Drive.Files.get(fileId, { supportsAllDrives: true, fields: 'parents' });
  var removeParents = (meta.parents || []).join(',');
  Drive.Files.update({}, fileId, null, {
    addParents: folderId,
    removeParents: removeParents,
    supportsAllDrives: true
  });
}

// ---- One-time config seeding (run from the Apps Script editor) ----

// Seeds the folder-provisioning script properties in one shot so nothing has to
// be hand-pasted into the Properties UI. Safe to re-run (it just overwrites the
// same values). Edit here and re-run when positions/emails change.
function setupFolderConfig() {
  requireOwner_();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ROOT_FOLDER_ID', 'YOUR_SHARED_DRIVE_FOLDER_ID'); // Shared Drive root; app creates VISA_REC_APP under it
  props.setProperty('ALLOWED_DOMAINS', 'lusu.ca,outpostpub.ca,thestudycoffeehouse.ca');
  props.setProperty('DEPARTMENT_MAP', JSON.stringify({
    'president@example.org':  'President',
    'vpfinance@example.org':  'VP Finance',
    'events@example.org':     'Events'
  }));
  Logger.log('Folder config seeded: ROOT_FOLDER_ID, ALLOWED_DOMAINS, DEPARTMENT_MAP (example mappings).');
}

// ---- One-time backfill for existing users (Finance/ED only) ----

// Creates a department folder for every active user that lacks a drive_folder_id
// and writes the ID back, then grants each active user Drive access to their
// folder (rows that already have an ID are exactly the ones missing ACLs from
// before grants existed). Idempotent and safe to re-run: find-or-create never
// duplicates folders and permission grants upsert.
function backfillUserFolders() {
  requireRole(['finance', 'ed']);
  getRootFolderId_(); // fail fast with a clear message if unconfigured

  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var results = { created: 0, skipped: 0, failed: 0, granted: 0, grant_failed: 0, details: [] };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[ix.active]) { results.skipped++; continue; }
    var id = String(row[ix.drive_folder_id] || '').trim();
    if (!id) {
      try {
        id = provisionUserReceiptsFolder_({ email: row[ix.email], department_suffix: row[ix.department_suffix] });
        sheet.getRange(i + 1, ix.drive_folder_id + 1).setValue(id);
        results.created++;
        results.details.push({ email: row[ix.email], folder_id: id });
      } catch (e) {
        results.failed++;
        results.details.push({ email: row[ix.email], error: e.message });
        continue;
      }
    }
    if (grantDriveAccess_(id, row[ix.email], 'writer')) {
      results.granted++;
    } else {
      results.grant_failed++;
      results.details.push({ email: row[ix.email], grant_failed: id });
    }
  }
  return results;
}

// ---- One-off statement cleanup (run from the Apps Script editor) ----

// Trash statement PDFs in the Visa Statements folder whose card-holder first
// name is in `firstNames` (lowercased) and whose month label matches `month`.
// Uses setTrashed — recoverable from Drive trash for 30 days. Returns the
// exact filenames trashed; if the count isn't what you expected, check the
// folder before re-running (names may differ from what you assumed).
function trashStatementPdfs_(firstNames, month) {
  requireRole(['finance', 'ed']);
  var folderId = getStatementsFolderId_();
  var wantMonth = String(normalizeMonth(month) || month).trim().toLowerCase();
  var results = { trashed: [], skipped: 0 };

  var files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    var file = files.next();
    var m = file.getName().match(STATEMENT_NAME_RE);
    if (!m) { results.skipped++; continue; }
    var holderFirst = m[1].trim().toLowerCase().split(/\s+/)[0];
    var fileMonth = String(normalizeMonth(m[2].trim()) || m[2]).trim().toLowerCase();
    if (firstNames.indexOf(holderFirst) === -1 || fileMonth !== wantMonth) {
      results.skipped++;
      results.skipped_names = results.skipped_names || [];
      results.skipped_names.push(file.getName());
      continue;
    }
    file.setTrashed(true);
    results.trashed.push(file.getName());
  }
  Logger.log('trashStatementPdfs_: ' + JSON.stringify(results));
  return results;
}


// Diagnostic: list every PDF in the statements folder (and how each one
// parses) so cleanup targets can be checked before trashing anything.
function listStatementFiles() {
  requireRole(['finance', 'ed']);
  var results = [];
  var files = DriveApp.getFolderById(getStatementsFolderId_()).getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    var name = files.next().getName();
    var m = name.match(STATEMENT_NAME_RE);
    results.push({
      name: name,
      cardholder: m ? m[1].trim() : '(does not match statement pattern)',
      month: m ? m[2].trim() : ''
    });
  }
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

// ---- One-time repair for mapped users with stale folder IDs (Finance/ED only) ----

// Re-resolves the department folder for every active user whose email is in
// DEPARTMENT_MAP and overwrites their stored drive_folder_id when it points
// elsewhere (e.g. the "Admin (ALL)" bucket, from before their mapping existed or
// before folders were keyed by email). backfillUserFolders() can't fix these rows
// because it skips anyone who already has a drive_folder_id. Existing files are
// not moved — only future receipts/exports file into the corrected folder.
// Also guarantees the "Visa Statements" folder exists (reported in the results),
// and reports DEPARTMENT_MAP keys that matched no Users-sheet row at all
// (unmatched_map_keys) so an email mismatch can't hide among expected skips.
// Note: DEPARTMENT_MAP is treated as the source of truth for mapped users — any
// hand-set drive_folder_id override on a mapped row is reverted. Safe to re-run.
function repairMappedUserFolders() {
  requireRole(['finance', 'ed']);
  getRootFolderId_(); // fail fast with a clear message if unconfigured

  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var results = {
    // Created up front so bug-1 is fixed even if every user row below is skipped.
    statements_folder_id: ensureStatementsFolder_(),
    repaired: 0, unchanged: 0, skipped: 0, failed: 0,
    granted: 0, grant_failed: 0,
    unmatched_map_keys: [], details: []
  };

  var matchedEmails = {}; // lowercased sheet emails found in DEPARTMENT_MAP (active or not)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = row[ix.email];
    var deptName = folderNameByEmail_(email);
    if (deptName) matchedEmails[String(email).trim().toLowerCase()] = true;
    if (!row[ix.active] || !deptName) {
      results.skipped++;
      results.details.push({ email: email, skipped: !row[ix.active] ? 'inactive' : 'not in DEPARTMENT_MAP' });
      continue;
    }
    try {
      var id = provisionUserReceiptsFolder_({ email: email });
      // Compare the RAW cell: a whitespace-padded ID breaks downstream Drive
      // queries (Receipts/Export read the cell untrimmed), so rewrite it too.
      if (String(row[ix.drive_folder_id] || '') === id) {
        results.unchanged++;
      } else {
        sheet.getRange(i + 1, ix.drive_folder_id + 1).setValue(id);
        results.repaired++;
        results.details.push({ email: email, folder: deptName, folder_id: id });
      }
      // Grant in both branches — unchanged rows may still lack the ACL.
      if (grantDriveAccess_(id, email, 'writer')) {
        results.granted++;
      } else {
        results.grant_failed++;
        results.details.push({ email: email, grant_failed: id });
      }
    } catch (e) {
      results.failed++;
      results.details.push({ email: email, error: e.message });
    }
  }

  // Map keys with no matching Users row — likely an email mismatch (e.g. the
  // sheet says comms@ but the map says communications@). These would otherwise
  // hide among the expected 'not in DEPARTMENT_MAP' skips of ALL-bucket users.
  var map = getDepartmentMap_();
  for (var k in map) {
    if (map.hasOwnProperty(k) && !matchedEmails[String(k).trim().toLowerCase()]) {
      results.unmatched_map_keys.push(k);
    }
  }
  return results;
}

// ---- One-time split of the shared "Admin (ALL)" receipts folder ----

// Run once from the Apps Script editor. Every active user whose receipts folder
// is the shared Admin (ALL) folder gets their own Admin (ALL)/<email> subfolder,
// and their access moves from the shared folder to it. Until then, receipt
// scanning skips shared folders. Files already in the shared folder stay put;
// move them into the right person's folder by hand. Safe to re-run.
function splitSharedReceiptFolders() {
  requireRole(['finance', 'ed']);
  var receiptsId = findOrCreateChildFolder_(getAppRootId_(), FOLDER_NAMES.RECEIPTS);
  var allId = findOrCreateChildFolder_(receiptsId, getAllFolderName_());

  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var results = { moved: [], failed: [] };
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][ix.email] || '').trim();
    if (!data[i][ix.active] || !email || folderNameByEmail_(email)) continue;
    if (String(data[i][ix.drive_folder_id] || '').trim() !== allId) continue;
    try {
      var ownId = provisionUserReceiptsFolder_({ email: email });
      sheet.getRange(i + 1, ix.drive_folder_id + 1).setValue(ownId);
      grantDriveAccess_(ownId, email, 'writer');
      revokeDriveAccess_(allId, email);
      results.moved.push(email);
    } catch (e) {
      results.failed.push({ email: email, error: e.message });
    }
  }
  Logger.log('splitSharedReceiptFolders: ' + JSON.stringify(results));
  return results;
}
