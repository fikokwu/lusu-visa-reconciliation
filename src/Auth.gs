// ============================================================
// Story 1.2 — Authentication and access control
// ============================================================

var VALID_ROLES = ['staff', 'finance', 'exec', 'ed'];

// Domains that may auto-provision a staff account on first sign-in. The
// ALLOWED_DOMAINS script property (comma-separated) can add more on top of these.
var DEFAULT_ALLOWED_DOMAINS = ['lusu.ca', 'outpostpub.ca', 'thestudycoffeehouse.ca'];

function getAllowedDomains_() {
  var domains = DEFAULT_ALLOWED_DOMAINS.slice();
  var extra = PropertiesService.getScriptProperties().getProperty('ALLOWED_DOMAINS') || '';
  extra.split(',').forEach(function(d) {
    d = d.trim().toLowerCase().replace(/^@/, '');
    if (d && domains.indexOf(d) === -1) domains.push(d);
  });
  return domains;
}

// Resolve the signed-in (accessing) user's email, normalised to lowercase + trimmed.
// The web app executes as USER_DEPLOYING (the owner), so the script can read the
// private database. getActiveUser().getEmail() still returns the *visitor's* email
// because they share the owner's Workspace domain. Do NOT fall back to
// getEffectiveUser() here: under "execute as me" that returns the OWNER's email,
// which would silently authenticate an unresolved visitor as the owner (a
// privileged finance/ed account). An empty result must stay empty so the caller
// denies access.
function getActiveEmail_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  return String(email).trim().toLowerCase();
}

function emailDomainAllowed_(email) {
  var at = email.lastIndexOf('@');
  if (at === -1) return false;
  return getAllowedDomains_().indexOf(email.slice(at + 1)) !== -1;
}

function rowToUser_(headers, row) {
  var user = {};
  headers.forEach(function(h, j) { user[h] = row[j]; });
  user.role = String(user.role || '').toLowerCase().trim();
  return jsonSafe(user);
}

function getCurrentUser() {
  var email = getActiveEmail_();
  if (!email) {
    throw new Error(
      'We could not detect which Google account you are signed in with. ' +
      'Open this link in an Incognito window, or use Chrome\'s account switcher ' +
      '(profile picture, top-right) to sign in with your organisation account.'
    );
  }

  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var emailIdx = headers.indexOf('email');
  var activeIdx = headers.indexOf('active');

  // Case-insensitive, whitespace-tolerant match against the Users sheet.
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[emailIdx] || '').trim().toLowerCase() === email) {
      if (!row[activeIdx]) {
        throw new Error('Your account is inactive. Contact your Finance officer.');
      }
      return rowToUser_(headers, row);
    }
  }

  // Auto-provision anyone signing in with an allowed organisation domain.
  if (emailDomainAllowed_(email)) {
    var name = email.split('@')[0];
    // Best-effort Drive folder; a provisioning failure must never block sign-in.
    var autoFolderId = '';
    try { autoFolderId = provisionUserReceiptsFolder_({ email: email, department_suffix: 'ALL' }); }
    catch (e) { Logger.log('Folder provision failed for ' + email + ': ' + e.message); }
    if (autoFolderId) grantDriveAccess_(autoFolderId, email, 'writer');
    addUserRow_(name, email, 'staff', 'ALL', autoFolderId);
    var freshData = sheet.getDataRange().getValues();
    for (var j = 1; j < freshData.length; j++) {
      if (String(freshData[j][emailIdx] || '').trim().toLowerCase() === email) {
        var autoUser = rowToUser_(headers, freshData[j]);
        // A silently created account is the #1 cause of "my statements don't
        // show" (its name is the email local-part, e.g. 'vpa') — leave a trail.
        Logger.log('Auto-provisioned new account for ' + email + ' (name "' + name + '")');
        try { logAudit_(autoUser.user_id, '', '', 'auto_provisioned', '', email); } catch (e) {}
        return autoUser;
      }
    }
  }

  throw new Error(
    'Access denied for ' + email + '. ' +
    'This app is only available to authorised staff. If you have an organisation ' +
    'account (' + getAllowedDomains_().join(', ') + '), try opening this link in an ' +
    'Incognito window or use Chrome\'s account switcher (profile picture, top-right) ' +
    'to sign in with that account.'
  );
}

function getUserBudgetCodes_(user) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);

  // department_suffix supports comma-separated values (e.g. "1000,1500") or "ALL"
  var suffixes = String(user.department_suffix || '').split(',').map(function(s) { return s.trim(); });
  var isAll = suffixes.indexOf('ALL') !== -1;

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[ix.active]) continue;
    if (row[ix.type] !== 'expense') continue;
    if (isAll || suffixes.indexOf(String(row[ix.department_suffix])) !== -1) {
      results.push({
        code: row[ix.code],
        description: row[ix.description],
        department_suffix: row[ix.department_suffix]
      });
    }
  }
  return results;
}

function getAllBudgetCodes() {
  var sheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[ix.active]) continue;
    if (row[ix.type] !== 'expense') continue;
    results.push({ code: row[ix.code], description: row[ix.description], department_suffix: row[ix.department_suffix] });
  }
  return jsonSafe(results);
}

// Editor-only bootstrap for the first Finance user; the app's Users tab goes
// through addUserFromUI instead.
function addUser(name, email, role, department_suffix, drive_folder_id) {
  requireOwner_();
  addUserRow_(name, email, role, department_suffix, drive_folder_id);
}

// Setup and migration functions stay callable from the browser like any other
// public function, so they refuse anyone but the script owner.
function requireOwner_() {
  var active = getActiveEmail_();
  var owner = '';
  try { owner = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  if (!active || active !== owner) {
    throw new Error('Only the app owner can run this, from the Apps Script editor.');
  }
}

// Scheduled jobs: allowed when started by one of this project's own triggers
// (the event carries its trigger id) or run by the owner from the editor.
function requireTriggerOrOwner_(e) {
  var uid = e && e.triggerUid;
  if (uid && ScriptApp.getProjectTriggers().some(function(t) { return t.getUniqueId() === uid; })) return;
  requireOwner_();
}

function addUserRow_(name, email, role, department_suffix, drive_folder_id) {
  if (VALID_ROLES.indexOf(role) === -1) {
    throw new Error('Invalid role: ' + role + '. Must be one of: ' + VALID_ROLES.join(', '));
  }
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  sheet.appendRow([
    generateId('USR'),
    name,
    email,
    role,
    department_suffix,
    drive_folder_id || '',
    true
  ]);
}

function addTestUsers() {
  throw new Error('addTestUsers() is disabled in production. Add users directly via addUser() or the Users sheet.');
}

// ---- In-app user management (Finance + ED only) ----

function requireRole(roles) {
  var user = getCurrentUser();
  if (roles.indexOf(user.role) === -1) {
    throw new Error('You do not have permission for this action.');
  }
  return user;
}

function listUsers() {
  requireRole(['finance', 'ed']);
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return jsonSafe(data.slice(1).map(function(row) {
    return rowToObj_(headers, row);
  }));
}

function addUserFromUI(name, email, role, department_suffix, drive_folder_id) {
  requireRole(['finance', 'ed']);
  if (!name || !email) throw new Error('Name and email are required.');
  if (VALID_ROLES.indexOf(role) === -1) {
    throw new Error('Invalid role: ' + role + '. Must be one of: ' + VALID_ROLES.join(', '));
  }
  // Check for duplicate email
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var emailIdx = data[0].indexOf('email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailIdx] || '').trim().toLowerCase() === String(email).trim().toLowerCase()) {
      throw new Error('A user with this email already exists.');
    }
  }
  // Auto-provision the department Drive folder when the admin didn't supply one.
  // Best-effort: a Drive failure shouldn't block adding the user (backfill later).
  var folderId = drive_folder_id || '';
  if (!folderId) {
    try { folderId = provisionUserReceiptsFolder_({ email: email, department_suffix: department_suffix }); }
    catch (e) { Logger.log('Folder provision failed for ' + email + ': ' + e.message); }
  }
  // Grant even when the admin supplied the folder ID by hand — those rows are
  // the ones most likely to be missing an ACL today.
  if (folderId) grantDriveAccess_(folderId, email, 'writer');
  addUserRow_(name, email, role, department_suffix || '', folderId);
  var mail = sendWelcomeEmail_(name, email, role);
  return { success: true, email: mail };
}

function resendInvite(userId) {
  requireRole(['finance', 'ed']);
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.user_id] === userId) {
      var name = data[i][ix.name];
      var email = data[i][ix.email];
      var role = String(data[i][ix.role] || '').toLowerCase().trim();
      var mail = sendWelcomeEmail_(name, email, role);
      return { success: true, email: mail };
    }
  }
  throw new Error('User not found: ' + userId);
}

function sendWelcomeEmail_(name, email, role) {
  var appUrl = getAppUrl();
  var roleBlurb = '';
  if (role === 'staff' || role === 'exec') {
    roleBlurb = 'You\'ll review your monthly Visa statement, code each transaction to a budget code, attach receipts, and submit for review.';
  } else if (role === 'finance') {
    roleBlurb = 'You\'ll review staff submissions, verify totals against the bank statement, and approve them for the Executive Director.';
  } else if (role === 'ed') {
    roleBlurb = 'You\'ll provide final approval on reconciliations after Finance has reviewed them.';
  }

  var subject = 'You\'ve been added to LUSU Visa Reconciliation App';
  var body =
    'Hi ' + (name || 'there') + ',\n\n' +
    'You\'ve been added to the LUSU Visa Reconciliation system as ' + role + '.\n\n' +
    roleBlurb + '\n\n' +
    'Sign in here to get started:\n' + appUrl + '\n\n' +
    'If you have any questions, contact your VP Operations & Finance. ' +
    'Please Note this is just an alpha product and may have system bugs. ' +
    'If you note any errors please report them to VPOF. ' +
    'At this point the manual Visa Rec process still takes precedence over this app.';

  return sendEmail_(email, subject, body);
}

function updateUser(userId, fields) {
  requireRole(['finance', 'ed']);
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var userIdIdx = headers.indexOf('user_id');
  var allowed = ['name', 'role', 'department_suffix', 'drive_folder_id', 'active'];
  for (var i = 1; i < data.length; i++) {
    if (data[i][userIdIdx] === userId) {
      Object.keys(fields).forEach(function(f) {
        if (allowed.indexOf(f) === -1) return;
        if (f === 'role' && VALID_ROLES.indexOf(fields[f]) === -1) {
          throw new Error('Invalid role: ' + fields[f]);
        }
        var col = headers.indexOf(f);
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(fields[f]);
      });
      return { success: true };
    }
  }
  throw new Error('User not found: ' + userId);
}

function setUserActive(userId, active) {
  return updateUser(userId, { active: !!active });
}

// Permanently removes a user's row (and revokes their Drive folder access).
// Their reconciliations and transactions are deliberately kept for audit
// history — the dashboard Staff column falls back to the raw user_id.
function deleteUser(userId) {
  var admin = requireRole(['finance', 'ed']);
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.user_id] === userId) {
      var email = String(data[i][ix.email] || '');
      if (email.trim().toLowerCase() === String(admin.email || '').trim().toLowerCase()) {
        throw new Error('You cannot delete your own account.');
      }
      var folderId = String(data[i][ix.drive_folder_id] || '').trim();
      if (folderId) revokeDriveAccess_(folderId, email);
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error('User not found: ' + userId);
}

function doGet(e) {
  try {
    var user = getCurrentUser();
    return HtmlService.createTemplateFromFile('Page')
      .evaluate()
      .setTitle('LUSU Visa Reconciliation')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    var logo = logoDataUri();
    var logoHtml = logo
      ? '<img src="' + logo + '" alt="LUSU" style="height:44px;margin-bottom:28px">'
      : '<div style="font-size:30px;font-weight:800;color:#0A2A5E;letter-spacing:.5px;margin-bottom:28px">LUSU</div>';
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<style>body{font-family:system-ui,-apple-system,"Segoe UI",sans-serif;background:#F4F6F9;margin:0;' +
      'display:flex;align-items:center;justify-content:center;min-height:100vh;color:#1A2B45}' +
      '.box{background:#fff;border:1px solid #E2E6EC;border-radius:12px;padding:40px 36px;' +
      'max-width:420px;width:90%;text-align:center;box-shadow:0 2px 8px rgba(10,42,94,.08)}' +
      'h2{color:#0A2A5E;margin:0 0 10px;font-size:20px}' +
      'p{color:#6B7A90;font-size:14px;line-height:1.55;margin:0}</style></head><body>' +
      '<div class="box">' + logoHtml +
      '<h2>Access Denied</h2>' +
      '<p>' + err.message + '</p>' +
      '</div></body></html>'
    ).setTitle('Access Denied — LUSU Visa Reconciliation');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
