// ============================================================
// Story 2.2 — Submit and approval workflow
// ============================================================

function getReconById_(reconId) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var ridIdx = headers.indexOf('recon_id');
  for (var i = 1; i < data.length; i++) {
    if (data[i][ridIdx] === reconId) {
      var obj = rowToObj_(headers, data[i]);
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

function updateRecon_(reconId, fields) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var ix = colIndexes_(headers);
    var lastEditedIdx = ix.last_edited_at !== undefined ? ix.last_edited_at : -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][ix.recon_id] === reconId) {
        Object.keys(fields).forEach(function(f) {
          var col = ix[f];
          if (col !== undefined) sheet.getRange(i + 1, col + 1).setValue(fields[f]);
        });
        if (lastEditedIdx !== -1 && !fields.hasOwnProperty('last_edited_at')) {
          sheet.getRange(i + 1, lastEditedIdx + 1).setValue(new Date().toISOString());
        }
        return;
      }
    }
    throw new Error('Reconciliation not found: ' + reconId);
  } finally {
    lock.releaseLock();
  }
}

function getOwnerEmail_(userId) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.USERS);
  var data = sheet.getDataRange().getValues();
  var ix = colIndexes_(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (data[i][ix.user_id] === userId) {
      return { email: data[i][ix.email], name: data[i][ix.name] };
    }
  }
  return null;
}

function getReconTransactions_(reconId) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var ridIdx = headers.indexOf('recon_id');
  return jsonSafe(data.slice(1)
    .filter(function(r) { return r[ridIdx] === reconId; })
    .map(function(r) { return rowToObj_(headers, r); }));
}

function sendEmail_(to, subject, body) {
  var from = '';
  try { from = Session.getActiveUser().getEmail(); } catch (e) {}
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'LUSU Visa Reconciliation' });
    var remaining = MailApp.getRemainingDailyQuota();
    Logger.log('Email queued: from=' + from + ' to=' + to + ' subject="' + subject + '" remaining-quota=' + remaining);
    return { sent: true, from: from, to: to, remaining_quota: remaining };
  } catch (e) {
    Logger.log('Email FAILED: from=' + from + ' to=' + to + ' error=' + e.message);
    return { sent: false, from: from, to: to, error: e.message };
  }
}

// Strip the /a/macros/{domain}/ tenant prefix so links work for any Google
// account (not just signed-in lusu.ca users). The tenant-scoped form forces
// domain auth regardless of the deployment's "access: ANYONE" setting.
function getAppUrl() {
  var url = ScriptApp.getService().getUrl() || '';
  return url.replace(/\/a\/macros\/[^\/]+\//, '/macros/');
}

// ---- Notification helpers (shared by real handlers and adminSetStatus) ----

function notifyFinanceOfSubmission_(submitterName, recon, total, txCount) {
  var financeEmails = (PropertiesService.getScriptProperties().getProperty('FINANCE_EMAILS') || '').split(',');
  var appUrl = getAppUrl();
  var subject = 'Reconciliation submitted — ' + submitterName + ' ' + recon.month;
  var body = submitterName + ' has submitted their ' + recon.month + ' reconciliation.\n\n' +
    'Total: $' + Number(total || 0).toFixed(2) + ' | Transactions: ' + (txCount || 0) + '\n\n' +
    'Review it here: ' + appUrl;
  financeEmails.forEach(function(email) {
    if (email.trim()) sendEmail_(email.trim(), subject, body);
  });
}

function notifyEdOfFinanceApproval_(ownerName, recon, variance, notes) {
  var edEmail = PropertiesService.getScriptProperties().getProperty('ED_EMAIL') || '';
  if (!edEmail) return;
  var appUrl = getAppUrl();
  var v = Number(variance) || 0;
  var varianceStr = v >= 0 ? '+$' + v.toFixed(2) : '-$' + Math.abs(v).toFixed(2);
  var subject = 'Ready for final approval — ' + ownerName + ' ' + recon.month;
  var body = 'Finance has approved ' + ownerName + '\'s ' + recon.month + ' reconciliation.\n\n' +
    'Variance: ' + varianceStr + '\nFinance notes: ' + (notes || 'None') + '\n\n' +
    'Review here: ' + appUrl;
  sendEmail_(edEmail, subject, body);
}

function notifyFinalApproval_(owner, recon) {
  var ownerName = owner ? owner.name : '';
  var financeEmails = (PropertiesService.getScriptProperties().getProperty('FINANCE_EMAILS') || '').split(',');
  var subject = 'Reconciliation approved — ' + ownerName + ' ' + recon.month;
  var body = ownerName + '\'s ' + recon.month + ' reconciliation has been finally approved by the ED.';
  if (owner && owner.email) sendEmail_(owner.email, subject, body);
  financeEmails.forEach(function(e) { if (e.trim()) sendEmail_(e.trim(), subject, body); });
}

// ---- Submit ----

function submitReconciliation(reconId) {
  var user = getCurrentUser();
  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (recon.user_id !== user.user_id) throw new Error('Access denied.');
  if (['submitted','finance_review','ed_review','approved'].indexOf(recon.status) !== -1) {
    throw new Error('This reconciliation has already been submitted.');
  }

  var txs = getReconTransactions_(reconId);
  var errors = [];

  // Only charges and fees need a code; only charges need a receipt.
  var requireReceipts = (user.role === 'staff');
  txs.forEach(function(tx) {
    var type = txTypeOf_(tx.tx_type);
    if (txNeedsCode_(type) && !tx.budget_code) {
      errors.push('Missing budget code: ' + tx.date + ' – ' + tx.vendor + ' ($' + tx.total + ')');
    }
    if (requireReceipts && type === 'charge' && tx.receipt_status !== 'matched') {
      errors.push('Missing receipt: ' + tx.date + ' – ' + tx.vendor + ' ($' + tx.total + ')');
    }
  });

  if (errors.length > 0) return { success: false, errors: errors };

  var total = reconSpendTotal_(txs);
  var now = new Date().toISOString();

  updateRecon_(reconId, {
    status: 'submitted',
    reconciled_total: total,
    submitted_at: now
  });
  logAudit_(user.user_id, reconId, '', 'submitted', '', 'submitted');

  notifyFinanceOfSubmission_(user.name, recon, total, txs.length);

  return { success: true };
}

// ---- Finance approve ----

function financeApprove(reconId, bankStatementTotal, notes) {
  var user = getCurrentUser();
  if (user.role !== 'finance') throw new Error('Finance role required.');

  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (recon.status !== 'submitted') throw new Error('Only a submitted reconciliation can be approved by Finance.');

  var variance = statementVariance_(bankStatementTotal, recon.previous_balance, getReconTransactions_(reconId));
  updateRecon_(reconId, {
    status: 'ed_review',
    bank_statement_total: bankStatementTotal,
    variance: variance,
    finance_notes: notes || '',
    finance_approved_at: new Date().toISOString()
  });
  logAudit_(user.user_id, reconId, '', 'finance_approved', '', '');

  var owner = getOwnerEmail_(recon.user_id);
  notifyEdOfFinanceApproval_(owner ? owner.name : '', recon, variance, notes);

  return { success: true };
}

// ---- Finance send back ----

function financeSendBack(reconId, notes) {
  var user = getCurrentUser();
  if (user.role !== 'finance') throw new Error('Finance role required.');
  if (!notes || !notes.trim()) throw new Error('Notes are required when sending back.');

  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (recon.status !== 'submitted') throw new Error('Only a submitted reconciliation can be sent back by Finance.');
  updateRecon_(reconId, { status: 'sent_back', finance_notes: notes });
  logAudit_(user.user_id, reconId, '', 'finance_sent_back', '', notes);

  var owner = getOwnerEmail_(recon.user_id);
  var appUrl = getAppUrl();
  if (owner) {
    sendEmail_(owner.email,
      'Action needed: Reconciliation sent back — ' + recon.month,
      'Your ' + recon.month + ' reconciliation has been sent back by Finance.\n\n' +
      'Notes: ' + notes + '\n\nPlease update and resubmit: ' + appUrl
    );
  }

  return { success: true };
}

// ---- ED approve ----

function edApprove(reconId, notes) {
  var user = getCurrentUser();
  if (user.role !== 'ed') throw new Error('ED role required.');

  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (recon.status !== 'ed_review') throw new Error('Only a reconciliation in ED review can be given final approval.');
  updateRecon_(reconId, {
    status: 'approved',
    ed_notes: notes || '',
    ed_approved_at: new Date().toISOString()
  });
  logAudit_(user.user_id, reconId, '', 'ed_approved', '', '');

  // Archive the approved reconciliation to "Approved Visa Recs/<Month>". The recon
  // is now 'approved', so exportReconciliation files it in the central archive.
  // Best-effort: never let an export/Drive hiccup block final approval.
  try { exportReconciliation(reconId); }
  catch (e) { Logger.log('Auto-archive on approval failed for ' + reconId + ': ' + e.message); }

  var owner = getOwnerEmail_(recon.user_id);
  notifyFinalApproval_(owner, recon);

  return { success: true };
}

// ---- Admin status override (testing) ----

function adminSetStatus(reconId, newStatus, sendEmails) {
  var actor = requireRole(['finance', 'ed']);
  var allowed = ['draft', 'submitted', 'ed_review', 'approved', 'sent_back'];
  if (allowed.indexOf(newStatus) === -1) throw new Error('Invalid status: ' + newStatus);

  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  var oldStatus = recon.status;
  var owner = getOwnerEmail_(recon.user_id);
  var ownerName = owner ? owner.name : 'Unknown';

  var fields = { status: newStatus };
  var now = new Date().toISOString();
  var total = recon.reconciled_total;
  var txCount = 0;

  if (newStatus === 'submitted') {
    var txs = getReconTransactions_(reconId);
    total = reconSpendTotal_(txs);
    fields.reconciled_total = total;
    fields.submitted_at = now;
    txCount = txs.length;
  } else if (newStatus === 'ed_review') {
    fields.finance_approved_at = now;
  } else if (newStatus === 'approved') {
    fields.ed_approved_at = now;
  }

  updateRecon_(reconId, fields);
  logAudit_(actor.user_id, reconId, '', 'admin_status_override', oldStatus, newStatus);

  // Keep the central archive in step with an admin approval override. Best-effort.
  if (newStatus === 'approved') {
    try { exportReconciliation(reconId); }
    catch (e) { Logger.log('Auto-archive on admin approval failed for ' + reconId + ': ' + e.message); }
  }

  if (sendEmails) {
    var freshRecon = getReconById_(reconId);
    if (newStatus === 'submitted') {
      notifyFinanceOfSubmission_(ownerName, freshRecon, total, txCount);
    } else if (newStatus === 'ed_review') {
      notifyEdOfFinanceApproval_(ownerName, freshRecon, Number(freshRecon.variance) || 0, freshRecon.finance_notes);
    } else if (newStatus === 'approved') {
      notifyFinalApproval_(owner, freshRecon);
    }
  }

  return { success: true, old_status: oldStatus, new_status: newStatus };
}

// ---- ED send back ----

function edSendBack(reconId, notes) {
  var user = getCurrentUser();
  if (user.role !== 'ed') throw new Error('ED role required.');
  if (!notes || !notes.trim()) throw new Error('Notes are required when sending back.');

  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (recon.status !== 'ed_review') throw new Error('Only a reconciliation in ED review can be sent back by the ED.');
  updateRecon_(reconId, { status: 'sent_back', ed_notes: notes });
  logAudit_(user.user_id, reconId, '', 'ed_sent_back', '', notes);

  var owner = getOwnerEmail_(recon.user_id);
  var financeEmails = (PropertiesService.getScriptProperties().getProperty('FINANCE_EMAILS') || '').split(',');
  var appUrl = getAppUrl();
  var body = 'The ' + recon.month + ' reconciliation has been sent back by the ED.\n\nNotes: ' + notes + '\n\n' + appUrl;

  if (owner) sendEmail_(owner.email, 'Action needed: Reconciliation sent back — ' + recon.month, body);
  financeEmails.forEach(function(e) { sendEmail_(e.trim(), 'Reconciliation sent back — ' + recon.month, body); });

  return { success: true };
}
