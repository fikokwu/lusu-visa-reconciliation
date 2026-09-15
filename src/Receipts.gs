// ============================================================
// Story 2.1 — Receipt attachment
// Story 3.1 — Auto Drive scanning
// Story 3.2 — Duplicate detection
// ============================================================

// ---- Manual attach / remove ----

function attachReceipt(txId, base64File, fileName, mimeType) {
  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);

  var txData = txSheet.getDataRange().getValues();
  var txHeaders = txData[0];
  var tcx = colIndexes_(txHeaders);
  var tx = null, txRow = -1;
  for (var i = 1; i < txData.length; i++) {
    if (txData[i][tcx.tx_id] === txId) {
      tx = rowToObj_(txHeaders, txData[i]); txRow = i + 1;
      break;
    }
  }
  if (!tx) throw new Error('Transaction not found.');

  // Access + status check; fails closed when the transaction has no reconciliation.
  var reconData = reconSheet.getDataRange().getValues();
  var rcx = colIndexes_(reconData[0]);
  var reconFound = false;
  for (var r = 1; r < reconData.length; r++) {
    if (reconData[r][rcx.recon_id] === tx.recon_id) {
      reconFound = true;
      var status = reconData[r][rcx.status];
      var ownerId = reconData[r][rcx.user_id];
      if (['finance', 'ed'].indexOf(user.role) === -1 && ownerId !== user.user_id) throw new Error('Access denied.');
      if (['submitted','finance_review','ed_review','approved'].indexOf(status) !== -1) {
        throw new Error('Reconciliation is locked.');
      }
      break;
    }
  }
  if (!reconFound) throw new Error('Reconciliation not found.');

  // Save to Drive: the user's department folder, in the <Month YYYY> subfolder.
  var userSheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var userData = userSheet.getDataRange().getValues();
  var ucx = colIndexes_(userData[0]);
  var deptFolderId = '';
  for (var u = 1; u < userData.length; u++) {
    if (userData[u][ucx.user_id] === user.user_id) {
      deptFolderId = userData[u][ucx.drive_folder_id];
      break;
    }
  }

  // Resolve this reconciliation's month so the receipt files under the right month.
  var reconMonth = '';
  for (var rm = 1; rm < reconData.length; rm++) {
    if (reconData[rm][rcx.recon_id] === tx.recon_id) { reconMonth = reconData[rm][rcx.month]; break; }
  }

  var bytes = Utilities.base64Decode(base64File);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var file;
  if (deptFolderId) {
    var targetFolderId = ensureMonthSubfolder_(deptFolderId, reconMonth);
    file = createFileInFolder_(targetFolderId, blob, fileName);
  } else {
    // No folder configured for this user — fall back to the owner's Drive root.
    file = DriveApp.createFile(blob);
  }

  // Receipts row
  var receiptSheet = getOrCreateSheet_(SHEET_NAMES.RECEIPTS);
  var receiptId = generateId('RCP');
  receiptSheet.appendRow([
    receiptId, tx.recon_id, file.getId(), fileName, txId,
    100, 'matched', tx.total, tx.date, tx.vendor, new Date().toISOString()
  ]);

  // Update transaction
  txSheet.getRange(txRow, tcx.receipt_file_id + 1).setValue(file.getId());
  txSheet.getRange(txRow, tcx.receipt_status + 1).setValue('matched');

  logAudit_(user.user_id, tx.recon_id, txId, 'receipt_attached', '', file.getId());

  return { success: true, file_id: file.getId(), view_url: 'https://drive.google.com/file/d/' + file.getId() + '/view' };
}

function removeReceipt(txId) {
  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = txSheet.getDataRange().getValues();
  var headers = data[0];
  var ix = colIndexes_(headers);

  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.tx_id] === txId) {
      var reconId = data[i][ix.recon_id];
      var recon = getReconById_(reconId);
      if (!recon) throw new Error('Reconciliation not found.');
      if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) throw new Error('Access denied.');
      if (['submitted', 'finance_review', 'ed_review', 'approved'].indexOf(recon.status) !== -1) {
        throw new Error('Reconciliation is locked.');
      }
      var oldFileId = data[i][ix.receipt_file_id];
      txSheet.getRange(i + 1, ix.receipt_file_id + 1).setValue('');
      txSheet.getRange(i + 1, ix.receipt_status + 1).setValue('unmatched');

      // Mark receipt row as removed
      var rcptSheet = getOrCreateSheet_(SHEET_NAMES.RECEIPTS);
      var rcptData = rcptSheet.getDataRange().getValues();
      var rcx = colIndexes_(rcptData[0]);
      for (var r = 1; r < rcptData.length; r++) {
        if (rcptData[r][rcx.tx_id] === txId && rcptData[r][rcx.match_status] === 'matched') {
          rcptSheet.getRange(r + 1, rcx.match_status + 1).setValue('removed');
        }
      }

      logAudit_(user.user_id, reconId, txId, 'receipt_removed', oldFileId, '');
      return { success: true };
    }
  }
  throw new Error('Transaction not found.');
}

// ---- Auto receipt reading ----
//
// Receipts dropped anywhere in a user's receipts folder (month subfolders
// included) are read by Claude, which also picks the card charge each one pays
// for from the user's open reconciliations (draft or sent back). Runs when a
// reconciliation is opened, from the Scan Drive button, and hourly
// (setupReceiptScanTrigger). Each file is read once; its scanned details are
// kept, so a receipt that didn't match is re-matched from those details —
// without reading the file again — whenever the user's statement charges change.
// A file holding several receipts (a multi-page scan) becomes one receipt row per
// receipt (file_part), each matched to its own charge. Matching goes by what is
// printed on the receipt; file names are never used.

var RECEIPT_READABLE_TYPES = ['application/pdf', 'image/jpeg', 'image/png', 'image/gif', 'image/webp'];
var RECEIPT_UNREADABLE_TYPES = ['image/heic', 'image/heif']; // iPhone default; Claude can't read these
var RECEIPT_MAX_BYTES = 7 * 1024 * 1024;                    // under Claude's 10 MB base64 image limit
var RECEIPT_LOOKBACK_DAYS = 90;
var RECEIPT_AUTO_ATTACH_CONFIDENCE = 80;
var RECEIPT_SUGGEST_CONFIDENCE = 40;
var RECEIPT_SCAN_BUDGET_MS = 4.5 * 60 * 1000;               // under Apps Script's 6-minute limit

var RECEIPT_MATCH_PROMPT =
  'You match receipts to Visa card charges for a university student union. ' +
  'A file can hold one receipt or several (for example a multi-page scan); find every receipt in it. ' +
  'For each receipt, pick the card charge it pays for from the list, or null if none fits, ' +
  'and never pick the same charge for two receipts. ' +
  'Card vendor names are often abbreviated (e.g. "SQ *BREW CO" for Brew Co.), a restaurant charge ' +
  'can include a tip the receipt does not show, a US-dollar receipt is charged in Canadian dollars, ' +
  'and the card date can be a few days after the receipt date. ' +
  'Return ONLY valid JSON: {"receipts": [{"vendor": string, "date": "YYYY-MM-DD", "total": number, ' +
  '"tx_id": string or null, "confidence": integer 0-100, "reason": string under 12 words}]}. ' +
  'Use null for any receipt field you cannot read, and an empty list if the file holds no receipt.';

// Scan Drive button, and the automatic scan when a reconciliation is opened.
function scanReceiptsForRecon(reconId) {
  var user = getCurrentUser();
  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) {
    throw new Error('Access denied.');
  }
  var owners = loadReceiptOwners_();
  var owner = owners.byId[recon.user_id];
  if (!owner) return jsonSafe(emptyScanCounts_('This reconciliation\'s owner is no longer an active user.'));
  return jsonSafe(scanReceiptsForUser_(owner, owners.sharedFolders, Date.now() + RECEIPT_SCAN_BUDGET_MS));
}

// Hourly trigger: every active user's folder, within one execution budget.
function scanAllReceipts(e) {
  requireTriggerOrOwner_(e);
  var deadline = Date.now() + RECEIPT_SCAN_BUDGET_MS;
  var owners = loadReceiptOwners_();
  owners.list.forEach(function(u) {
    if (Date.now() > deadline) return;
    try { scanReceiptsForUser_(u, owners.sharedFolders, deadline); }
    catch (e) { Logger.log('Receipt scan failed for ' + u.email + ': ' + e.message); }
  });
}

// Run once from the Apps Script editor to start hourly receipt reading.
// Also removes the old daily dailyDriveScan trigger if it's still installed.
function setupReceiptScanTrigger() {
  requireRole(['finance', 'ed']);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'scanAllReceipts' || fn === 'dailyDriveScan') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanAllReceipts').timeBased().everyHours(1).create();
  Logger.log('Hourly receipt scan trigger set.');
}

function emptyScanCounts_(skipped) {
  var counts = { matched_count: 0, review_count: 0, unmatched_count: 0, unreadable_count: 0 };
  if (skipped) counts.skipped = skipped;
  return counts;
}

// Active users with their receipts folder, plus the folder IDs shared by more
// than one of them. Scanning a shared folder could match one person's receipt
// to someone else's charges, so those are skipped until splitSharedReceiptFolders().
function loadReceiptOwners_() {
  var data = getOrCreateSheet_(SHEET_NAMES.USERS).getDataRange().getValues();
  var ux = colIndexes_(data[0]);
  var list = [], byId = {}, folderUsers = {}, sharedFolders = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][ux.active]) continue;
    var u = {
      user_id: data[i][ux.user_id],
      email: String(data[i][ux.email] || '').trim(),
      drive_folder_id: String(data[i][ux.drive_folder_id] || '').trim()
    };
    list.push(u);
    byId[u.user_id] = u;
    if (!u.drive_folder_id) continue;
    folderUsers[u.drive_folder_id] = (folderUsers[u.drive_folder_id] || 0) + 1;
    if (folderUsers[u.drive_folder_id] > 1) sharedFolders[u.drive_folder_id] = true;
  }
  return { list: list, byId: byId, sharedFolders: sharedFolders };
}

function scanReceiptsForUser_(owner, sharedFolders, deadline) {
  if (!owner.drive_folder_id) return emptyScanCounts_('No receipts folder is set for this user.');
  if (sharedFolders[owner.drive_folder_id]) {
    return emptyScanCounts_('This receipts folder is shared with other users, so it isn\'t scanned. ' +
      'Finance can run splitSharedReceiptFolders() to give each user their own.');
  }
  // Soft per-user guard so the hourly run and a scan from the app don't read
  // the same new files at the same time.
  var cache = CacheService.getScriptCache();
  var busyKey = 'receipt-scan:' + owner.user_id;
  if (cache.get(busyKey)) return emptyScanCounts_('Receipts are already being read for this user. Try again in a minute.');
  cache.put(busyKey, '1', 360);
  try {
    return readAndMatchUserReceipts_(owner, deadline);
  } finally {
    cache.remove(busyKey);
  }
}

function readAndMatchUserReceipts_(owner, deadline) {
  var counts = emptyScanCounts_();

  // Candidates: charges still missing a receipt in the owner's open reconciliations.
  var reconData = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS).getDataRange().getValues();
  var rcx = colIndexes_(reconData[0]);
  var openMonths = {};
  for (var i = 1; i < reconData.length; i++) {
    var status = reconData[i][rcx.status];
    if (reconData[i][rcx.user_id] === owner.user_id && (status === 'draft' || status === 'sent_back')) {
      openMonths[reconData[i][rcx.recon_id]] = normalizeMonth(reconData[i][rcx.month]);
    }
  }

  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);
  var candidates = [];
  var chargeIds = [];
  for (var t = 1; t < txData.length; t++) {
    var reconId = txData[t][tcx.recon_id];
    if (!openMonths.hasOwnProperty(reconId)) continue;
    if (txTypeOf_(txData[t][tcx.tx_type]) !== 'charge') continue;
    chargeIds.push(txData[t][tcx.tx_id]);
    if (txData[t][tcx.receipt_status] !== 'unmatched') continue;
    candidates.push({
      tx_id: txData[t][tcx.tx_id],
      recon_id: reconId,
      row: t + 1,
      total: Number(txData[t][tcx.total]) || 0,
      vendor: String(txData[t][tcx.vendor] || ''),
      date: promptDate_(txData[t][tcx.date]) + ' (' + openMonths[reconId] + ' statement)'
    });
  }
  // Fingerprint of the owner's open charges: a waiting receipt is re-matched
  // only when this changes (e.g. a new statement was imported).
  var chargesKey = Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, chargeIds.sort().join(',')));

  var rcptSheet = ensureColumns_(SHEET_NAMES.RECEIPTS);
  var rcptData = rcptSheet.getDataRange().getValues();
  var rh = rcptData[0];
  var rx = colIndexes_(rh);
  var ctx = { candidates: candidates, txSheet: txSheet, tcx: tcx, counts: counts };

  // 1) Receipts already read that are still waiting: re-match from their
  //    scanned details when the charges have changed. The file isn't re-read.
  var knownFiles = {};
  for (var r = 1; r < rcptData.length; r++) {
    knownFiles[rcptData[r][rx.file_id]] = true;
    var waiting = rcptData[r][rx.match_status] === 'unmatched' || rcptData[r][rx.match_status] === 'review';
    if (!waiting || rcptData[r][rx.user_id] !== owner.user_id) continue;
    if (!candidates.length || rcptData[r][rx.matched_against] === chargesKey) continue;
    if (Date.now() > deadline) continue;
    var fields = rowToObj_(rh, rcptData[r]);
    if (!fields.scanned_vendor && !hasValue_(fields.scanned_amount)) continue; // nothing was read to match on
    var pick;
    try {
      pick = callReceiptMatcher_([{ type: 'text', text:
        'Receipt already read. Vendor: ' + (fields.scanned_vendor || 'unknown') +
        '. Date: ' + (promptDate_(fields.scanned_date) || 'unknown') +
        '. Total: ' + (hasValue_(fields.scanned_amount) ? '$' + fields.scanned_amount : 'unknown') + '.\n\n' +
        candidateListText_(candidates) }])[0];
    } catch (e) {
      Logger.log('Receipt re-match failed for ' + fields.file_name + ': ' + e.message);
      continue;
    }
    applyReceiptDecision_(fields, pick, ctx);
    fields.matched_against = chargesKey;
    rcptSheet.getRange(r + 1, 1, 1, rh.length).setValues([rowFromFields_(rh, fields)]);
  }

  // 2) New files anywhere in the folder tree: read and match in one Claude call each.
  var files = listReceiptFiles_(owner.drive_folder_id);
  for (var f = 0; f < files.length; f++) {
    var file = files[f];
    if (knownFiles[file.id]) continue;
    if (Date.now() > deadline) break;
    var base = {
      user_id: owner.user_id, file_id: file.id, file_name: file.name,
      uploaded_at: new Date().toISOString(), matched_against: chargesKey
    };
    var readable = RECEIPT_READABLE_TYPES.indexOf(file.mimeType) !== -1;
    var rows = [];
    if (!readable || Number(file.size) > RECEIPT_MAX_BYTES) {
      // Listed in the app so the owner can match it by hand.
      rows.push(receiptRow_(base, { match_status: 'unreadable', match_reason: readable ? 'file too large' : 'HEIC photo' }));
      counts.unreadable_count++;
    } else {
      var parts;
      try {
        parts = readReceiptFile_(file, candidates);
      } catch (e) {
        Logger.log('Receipt read failed for ' + file.name + ': ' + e.message);
        continue; // not recorded, so it's retried on the next scan
      }
      if (!parts.length) {
        rows.push(receiptRow_(base, { match_status: 'unmatched', match_reason: 'No receipt found in this file' }));
        counts.unmatched_count++;
      }
      // A multi-receipt scan becomes one row per receipt, each matched on its own.
      parts.forEach(function(part, p) {
        var fields = receiptRow_(base, {
          file_part: parts.length > 1 ? p + 1 : '',
          scanned_vendor: part.vendor == null ? '' : part.vendor,
          scanned_date: part.date == null ? '' : part.date,
          scanned_amount: part.total == null ? '' : part.total
        });
        applyReceiptDecision_(fields, part, ctx);
        rows.push(fields);
      });
    }
    rows.forEach(function(fields) { rcptSheet.appendRow(rowFromFields_(rh, fields)); });
    knownFiles[file.id] = true;
  }
  return counts;
}

function receiptRow_(base, extra) {
  var row = { receipt_id: generateId('RCP') };
  Object.keys(base).forEach(function(k) { row[k] = base[k]; });
  Object.keys(extra).forEach(function(k) { row[k] = extra[k]; });
  return row;
}

// Records Claude's pick on a receipt row (fields) and attaches the receipt when
// sure: auto-attach needs high confidence AND the amounts to agree to the cent.
// Any other pick (a tip, exchange rate, or lower confidence) is left for review.
function applyReceiptDecision_(fields, pick, ctx) {
  pick = pick || {};
  var idx = -1;
  for (var i = 0; i < ctx.candidates.length; i++) {
    if (ctx.candidates[i].tx_id === pick.tx_id) { idx = i; break; }
  }
  var confidence = Number(pick.confidence) || 0;
  fields.match_confidence = confidence;
  fields.match_reason = pick.reason || '';
  if (idx === -1 || confidence < RECEIPT_SUGGEST_CONFIDENCE) {
    fields.match_status = 'unmatched';
    fields.tx_id = '';
    fields.recon_id = '';
    ctx.counts.unmatched_count++;
    return;
  }
  var tx = ctx.candidates[idx];
  fields.tx_id = tx.tx_id;
  fields.recon_id = tx.recon_id;
  var exact = hasValue_(fields.scanned_amount) && Math.abs(Number(fields.scanned_amount) - tx.total) < 0.01;
  if (exact && confidence >= RECEIPT_AUTO_ATTACH_CONFIDENCE) {
    fields.match_status = 'matched';
    ctx.txSheet.getRange(tx.row, ctx.tcx.receipt_file_id + 1).setValue(fields.file_id);
    ctx.txSheet.getRange(tx.row, ctx.tcx.receipt_status + 1).setValue('matched');
    ctx.candidates.splice(idx, 1); // one receipt per charge
    logAudit_('', tx.recon_id, tx.tx_id, 'receipt_auto_matched', '', fields.file_id);
    ctx.counts.matched_count++;
  } else {
    fields.match_status = 'review';
    ctx.counts.review_count++;
  }
}

function readReceiptFile_(file, candidates) {
  var data = Utilities.base64Encode(DriveApp.getFileById(file.id).getBlob().getBytes());
  var block = file.mimeType === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: data } }
    : { type: 'image', source: { type: 'base64', media_type: file.mimeType, data: data } };
  return callReceiptMatcher_([block, { type: 'text', text: candidateListText_(candidates) }]);
}

function candidateListText_(candidates) {
  if (!candidates.length) return 'Card charges: none to match yet, so just read the receipt (tx_id null).';
  return 'Card charges (tx_id | date | vendor | amount):\n' + candidates.map(function(c) {
    return c.tx_id + ' | ' + c.date + ' | ' + c.vendor + ' | $' + c.total.toFixed(2);
  }).join('\n');
}

function callReceiptMatcher_(content) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  var response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CLAUDE_MODEL_PARSE,
      max_tokens: 2048,
      system: RECEIPT_MATCH_PROMPT,
      messages: [{ role: 'user', content: content }]
    }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Receipt read failed: ' + response.getResponseCode());
  }
  var text = JSON.parse(response.getContentText()).content[0].text.trim();
  text = text.replace(/^```json?\s*/i, '').replace(/\s*```$/, '');
  var reply = JSON.parse(text) || {};
  // One entry per receipt found; tolerate a bare single-receipt reply.
  if (Array.isArray(reply.receipts)) return reply.receipts;
  return (reply.vendor !== undefined || reply.tx_id !== undefined) ? [reply] : [];
}

// Every receipt-type file under a folder, month subfolders included (up to
// three levels deep), modified within the lookback window.
function listReceiptFiles_(rootId) {
  var cutoff = new Date(Date.now() - RECEIPT_LOOKBACK_DAYS * 86400000).toISOString();
  var out = [];
  var queue = [{ id: rootId, depth: 0 }];
  while (queue.length) {
    var folder = queue.shift();
    var pageToken = null;
    do {
      var res = Drive.Files.list({
        q: "trashed=false and '" + folder.id + "' in parents",
        corpora: 'allDrives',
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        fields: 'nextPageToken,files(id,name,mimeType,size,modifiedTime)',
        pageSize: 200,
        pageToken: pageToken
      });
      (res.files || []).forEach(function(f) {
        if (f.mimeType === FOLDER_MIME) {
          if (folder.depth < 3) queue.push({ id: f.id, depth: folder.depth + 1 });
        } else if (f.modifiedTime >= cutoff &&
                   (RECEIPT_READABLE_TYPES.indexOf(f.mimeType) !== -1 ||
                    RECEIPT_UNREADABLE_TYPES.indexOf(f.mimeType) !== -1)) {
          out.push(f);
        }
      });
      pageToken = res.nextPageToken;
    } while (pageToken);
  }
  return out;
}

// Sheets turns date text into Date objects; give Claude a plain YYYY-MM-DD.
function promptDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  return String(v == null ? '' : v).trim();
}

function hasValue_(v) {
  return v !== '' && v !== null && v !== undefined;
}

function setupReminderTrigger() {
  requireRole(['finance', 'ed']);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendReceiptReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReceiptReminders')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  Logger.log('Weekly receipt reminder trigger set for Monday 9am.');
}

function sendReceiptReminders(e) {
  requireTriggerOrOwner_(e);
  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var txSheet    = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var userSheet  = getOrCreateSheet_(SHEET_NAMES.USERS);

  var reconData = reconSheet.getDataRange().getValues();
  var rcx = colIndexes_(reconData[0]);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);
  var userData = userSheet.getDataRange().getValues();
  var ucx = colIndexes_(userData[0]);

  var userMap = {};
  for (var u = 1; u < userData.length; u++) {
    userMap[userData[u][ucx.user_id]] = {
      email: userData[u][ucx.email],
      name:  userData[u][ucx.name]
    };
  }

  // Count unmatched receipts per recon
  var missing = {};
  for (var t = 1; t < txData.length; t++) {
    if (txData[t][tcx.receipt_status] === 'unmatched') {
      var rid = txData[t][tcx.recon_id];
      missing[rid] = (missing[rid] || 0) + 1;
    }
  }

  var appUrl = getAppUrl();
  for (var i = 1; i < reconData.length; i++) {
    var status = reconData[i][rcx.status];
    if (status !== 'draft' && status !== 'sent_back') continue;
    var rId = reconData[i][rcx.recon_id];
    if (!missing[rId]) continue;
    var owner = userMap[reconData[i][rcx.user_id]];
    if (!owner) continue;
    var month = reconData[i][rcx.month];
    var count = missing[rId];
    sendEmail_(owner.email,
      'Reminder: ' + count + ' missing receipt' + (count > 1 ? 's' : '') + ' — ' + month,
      'Hi ' + owner.name + ',\n\n' +
      'Your ' + month + ' reconciliation has ' + count + ' transaction' + (count > 1 ? 's' : '') +
      ' still missing a receipt.\n\nPlease attach your receipts and submit: ' + appUrl
    );
  }
}

function manualMatchReceipt(receiptId, txId) {
  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var tcx = colIndexes_(txData[0]);
  var txRow = -1;
  for (var t = 1; t < txData.length; t++) {
    if (txData[t][tcx.tx_id] === txId) { txRow = t + 1; break; }
  }
  if (txRow === -1) throw new Error('Transaction not found.');

  var reconId = txData[txRow - 1][tcx.recon_id];
  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) {
    throw new Error('Access denied.');
  }
  if (['submitted', 'finance_review', 'ed_review', 'approved'].indexOf(recon.status) !== -1) {
    throw new Error('Reconciliation is locked.');
  }

  var rcptSheet = getOrCreateSheet_(SHEET_NAMES.RECEIPTS);
  var rcptData = rcptSheet.getDataRange().getValues();
  var rcx = colIndexes_(rcptData[0]);
  for (var r = 1; r < rcptData.length; r++) {
    if (rcptData[r][rcx.receipt_id] !== receiptId) continue;
    // The receipt must come from this reconciliation's owner.
    var receiptOwner = rcptData[r][rcx.user_id];
    if (receiptOwner && receiptOwner !== recon.user_id) throw new Error('Access denied.');
    var fileId = rcptData[r][rcx.file_id];
    rcptSheet.getRange(r + 1, rcx.tx_id + 1).setValue(txId);
    rcptSheet.getRange(r + 1, rcx.recon_id + 1).setValue(reconId);
    rcptSheet.getRange(r + 1, rcx.match_status + 1).setValue('matched');
    rcptSheet.getRange(r + 1, rcx.match_confidence + 1).setValue(100);
    txSheet.getRange(txRow, tcx.receipt_file_id + 1).setValue(fileId);
    txSheet.getRange(txRow, tcx.receipt_status + 1).setValue('matched');
    logAudit_(user.user_id, reconId, txId, 'receipt_manually_matched', '', receiptId);
    return { success: true };
  }
  throw new Error('Receipt not found.');
}

// ---- Duplicate detection ----

// Receipts to look at for an open reconciliation: Claude's suggested matches,
// plus recent receipts from the owner's folder it couldn't match or read.
function getReceiptsForReview(reconId) {
  var user = getCurrentUser();
  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) {
    throw new Error('Access denied.');
  }
  if (recon.status !== 'draft' && recon.status !== 'sent_back') return [];

  var txData = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS).getDataRange().getValues();
  var txHeaders = txData[0];
  var tcx = colIndexes_(txHeaders);
  var unmatchedTxs = [];
  for (var t = 1; t < txData.length; t++) {
    if (txData[t][tcx.recon_id] === reconId && txData[t][tcx.receipt_status] === 'unmatched') {
      unmatchedTxs.push(rowToObj_(txHeaders, txData[t]));
    }
  }

  var cutoff = new Date(Date.now() - RECEIPT_LOOKBACK_DAYS * 86400000);
  var rcptData = getOrCreateSheet_(SHEET_NAMES.RECEIPTS).getDataRange().getValues();
  var rcptHeaders = rcptData[0];
  var rcx = colIndexes_(rcptHeaders);
  var results = [];
  for (var r = 1; r < rcptData.length; r++) {
    var status = rcptData[r][rcx.match_status];
    var suggestedHere = status === 'review' && rcptData[r][rcx.recon_id] === reconId;
    var ownersLoose = (status === 'unmatched' || status === 'unreadable') &&
      rcptData[r][rcx.user_id] === recon.user_id &&
      new Date(rcptData[r][rcx.uploaded_at]) >= cutoff;
    if (!suggestedHere && !ownersLoose) continue;
    var rcpt = rowToObj_(rcptHeaders, rcptData[r]);
    rcpt.available_txs = unmatchedTxs;
    results.push(rcpt);
  }
  return jsonSafe(results);
}

function checkDuplicates_(reconId) {
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = txSheet.getDataRange().getValues();
  var headers = data[0];
  var ix = colIndexes_(headers);

  var txs = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.recon_id] === reconId) {
      txs.push({ row: i + 1, data: data[i] });
    }
  }

  var dupIdx = ix.duplicate_flag;
  var txDups = [];

  for (var a = 0; a < txs.length; a++) {
    for (var b = a + 1; b < txs.length; b++) {
      var ra = txs[a].data, rb = txs[b].data;
      var vendorA = ra[ix.vendor], totalA = ra[ix.total], dateA = ra[ix.date];

      if (vendorA === rb[ix.vendor] &&
          Math.abs(totalA - rb[ix.total]) < 0.01 &&
          String(dateA) === String(rb[ix.date])) {
        txSheet.getRange(txs[a].row, dupIdx + 1).setValue(true);
        txSheet.getRange(txs[b].row, dupIdx + 1).setValue(true);
        logAudit_('', reconId, ra[ix.tx_id], 'duplicate_flagged', 'false', 'true');
        logAudit_('', reconId, rb[ix.tx_id], 'duplicate_flagged', 'false', 'true');
        txDups.push({
          tx_id_1: ra[ix.tx_id],
          tx_id_2: rb[ix.tx_id],
          vendor: vendorA, total: totalA, date: dateA
        });
      }
    }
  }

  // Receipt duplicates
  var rcptSheet = getOrCreateSheet_(SHEET_NAMES.RECEIPTS);
  var rcptData = rcptSheet.getDataRange().getValues();
  var rx = colIndexes_(rcptData[0]);
  var fileIds = {}, rcptDups = [];

  for (var rr = 1; rr < rcptData.length; rr++) {
    if (rcptData[rr][rx.recon_id] !== reconId) continue;
    // A multi-receipt scan has one row per receipt (file_part), so only the same
    // receipt of the same file counts as a duplicate.
    var fid = rcptData[rr][rx.file_id] + '|' + (rx.file_part === undefined ? '' : rcptData[rr][rx.file_part]);
    if (fileIds[fid]) {
      rcptSheet.getRange(rr + 1, rx.match_status + 1).setValue('duplicate');
      rcptDups.push({ receipt_id: rcptData[rr][rx.receipt_id], file_name: rcptData[rr][rx.file_name] });
    } else {
      fileIds[fid] = true;
    }
  }

  return { tx_duplicates: txDups, receipt_duplicates: rcptDups };
}

function dismissDuplicate(txId) {
  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = txSheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  var dupIdx = ix.duplicate_flag;

  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.tx_id] === txId) {
      var reconId = data[i][ix.recon_id];

      // Permissions: finance/ed always; owner only on their own unlocked recon.
      if (user.role !== 'finance' && user.role !== 'ed') {
        var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
        var reconData = reconSheet.getDataRange().getValues();
        var rcx = colIndexes_(reconData[0]);
        var allowed = false;
        for (var r = 1; r < reconData.length; r++) {
          if (reconData[r][rcx.recon_id] === reconId) {
            var status = reconData[r][rcx.status];
            var ownerId = reconData[r][rcx.user_id];
            if (ownerId === user.user_id && ['submitted','finance_review','ed_review','approved'].indexOf(status) === -1) {
              allowed = true;
            }
            break;
          }
        }
        if (!allowed) throw new Error('You cannot dismiss this flag.');
      }

      txSheet.getRange(i + 1, dupIdx + 1).setValue(false);
      logAudit_(user.user_id, reconId, txId, 'duplicate_dismissed', 'true', 'false');
      return { success: true };
    }
  }
  throw new Error('Transaction not found.');
}
