// ============================================================
// Stories 1.4 & 1.5 — Statement upload, AI coding, edit & save
// ============================================================

var CLAUDE_MODEL_PARSE = 'claude-haiku-4-5-20251001';  // PDF extraction
var CLAUDE_MODEL_CODE  = 'claude-sonnet-4-6';          // Budget code suggestions
var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';

// ---- Claude API ----

function callClaudeApi_(systemPrompt, userMessage, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  var payload = {
    model: CLAUDE_MODEL_CODE,
    max_tokens: maxTokens || 1024,
    system: systemPrompt,
    messages: [{ role: 'user', content: userMessage }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  var code = response.getResponseCode();
  if (code !== 200) {
    Logger.log('Claude API error ' + code + ': ' + response.getContentText());
    throw new Error('Claude API error: ' + code);
  }
  var result = JSON.parse(response.getContentText());
  return result.content[0].text;
}

// ---- Statement parsing ----

function parseStatement_(base64Pdf) {
  var systemPrompt = 'You are a financial data extraction assistant for a university student union. ' +
    'Extract this Visa statement. Return ONLY valid JSON with no explanation or markdown, shaped as ' +
    '{"previous_balance": number, "balance_owing": number, "transactions": [...]}. ' +
    'previous_balance and balance_owing come from the statement summary (null if not shown). ' +
    'Include every transaction line: purchases, payments, refunds/credits, fees and interest. ' +
    'Each transaction must have: date (string as shown on statement), vendor (string, uppercase), ' +
    'total (number, negative when the statement shows a minus sign or CR), ' +
    'type ("charge", "payment", "refund" or "fee"; interest counts as "fee"), hst (number, 0 if not shown).';

  var userMessage = {
    type: 'document',
    source: {
      type: 'base64',
      media_type: 'application/pdf',
      data: base64Pdf
    }
  };

  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  var payload = {
    model: CLAUDE_MODEL_PARSE,
    max_tokens: 4096,
    system: systemPrompt,
    messages: [{ role: 'user', content: [userMessage] }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to parse statement: ' + response.getContentText());
  }
  var text = JSON.parse(response.getContentText()).content[0].text.trim();
  // Strip markdown fences if present
  text = text.replace(/^```json?\s*/i, '').replace(/\s*```$/, '');
  var parsed = JSON.parse(text);
  if (Array.isArray(parsed)) parsed = { transactions: parsed };
  return {
    previous_balance: toAmount_(parsed.previous_balance),
    balance_owing: toAmount_(parsed.balance_owing),
    transactions: (parsed.transactions || []).map(normalizeStatementLine_)
      .filter(function(tx) { return tx.total !== 0; })
  };
}

// ---- Statement line types ----

// charge/fee add to the card balance; payment/refund reduce it (stored negative).
var STATEMENT_LINE_TYPES = ['charge', 'payment', 'refund', 'fee'];

// Transactions imported before payments/refunds were recorded have no tx_type;
// they were all charges.
function txTypeOf_(value) {
  var t = String(value || '').trim().toLowerCase();
  return STATEMENT_LINE_TYPES.indexOf(t) === -1 ? 'charge' : t;
}

// Only charges and fees get a budget code; payments and refunds don't.
function txNeedsCode_(type) {
  return type === 'charge' || type === 'fee';
}

// Statement amounts can come back as numbers or text like "-$2,276.06" or
// "2,276.06 CR". Returns a number rounded to cents, or '' when unreadable.
function toAmount_(v) {
  if (v === null || v === undefined || v === '') return '';
  var s = String(v).replace(/[$,\s]/g, '');
  var credit = /CR$/i.test(s);
  var n = Number(s.replace(/CR$/i, ''));
  if (isNaN(n)) return '';
  if (credit) n = -Math.abs(n);
  return Math.round(n * 100) / 100;
}

// The printed minus sign wins over the model's label, and a "PAYMENT" line is
// always a payment, so a payment or refund can never be recorded as a charge.
function normalizeStatementLine_(tx) {
  var vendor = String(tx.vendor || '').trim().toUpperCase();
  var total = Number(toAmount_(tx.total)) || 0;
  var type = String(tx.type || '').trim().toLowerCase();
  if (/^PAYMENT\b/.test(vendor)) type = 'payment';
  if (STATEMENT_LINE_TYPES.indexOf(type) === -1) type = total < 0 ? 'refund' : 'charge';
  if (total < 0 && (type === 'charge' || type === 'fee')) type = 'refund';
  if (type === 'payment' || type === 'refund') total = -Math.abs(total);
  return { date: tx.date, vendor: vendor, total: total, type: type };
}

// Spending on the card: charges, fees and refunds. Payments only pay the card
// down, so they're left out of totals and reports.
function reconSpendTotal_(txs) {
  var sum = txs.reduce(function(s, t) {
    return txTypeOf_(t.tx_type) === 'payment' ? s : s + (Number(t.total) || 0);
  }, 0);
  return Math.round(sum * 100) / 100;
}

// Statement balance check: previous balance + every line (payments included)
// should equal the balance owing. Returns balance owing minus that, in cents
// precision (0 when it balances).
function statementVariance_(balanceOwing, previousBalance, txs) {
  var calculated = (Number(previousBalance) || 0) +
    txs.reduce(function(s, t) { return s + (Number(t.total) || 0); }, 0);
  return Math.round((Number(balanceOwing) - calculated) * 100) / 100;
}

// ---- AI budget code suggestion ----

function suggestBudgetCode_(vendor, description, total, availableCodes, history) {
  var codeList = availableCodes.map(function(c) {
    return c.code + ' — ' + c.description;
  }).join('\n');

  var historyText = '';
  if (history && history.length > 0) {
    historyText = '\n\nPast confirmed coding for this vendor (weight heavily):\n' +
      history.map(function(h) {
        return h.code + ' — ' + h.description + ' (used ' + h.count + ' time' + (h.count > 1 ? 's' : '') + ')';
      }).join('\n');
  }

  var systemPrompt = 'You are a budget coding assistant for Lakehead University Student Union. ' +
    'Pick the best matching budget code for this transaction from the provided list. ' +
    'Return ONLY valid JSON: {"code": string, "confidence": integer 0-100, "reason": string}. ' +
    'Keep reason under 10 words. Set confidence below 75 if genuinely unsure. ' +
    'When past confirmed coding is provided, favour it unless the description clearly suggests a different category.';

  var userMessage = 'Vendor: ' + vendor + '\nDescription: ' + (description || '') +
    '\nAmount: $' + total + historyText + '\n\nAvailable codes:\n' + codeList;

  var text = callClaudeApi_(systemPrompt, userMessage, 256);
  text = text.trim().replace(/^```json?\s*/i, '').replace(/\s*```$/, '');
  return JSON.parse(text);
}

function getVendorHistory_(vendor, availableCodes) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var ix = colIndexes_(data[0]);
  var vendorIdx = ix.vendor;
  var codeIdx = ix.budget_code;
  var sourceIdx = ix.code_source;
  var confIdx = ix.ai_confidence;

  var allowedSet = {};
  availableCodes.forEach(function(c) { allowedSet[c.code] = c.description; });

  var vendorUp = vendor.toUpperCase();
  var counts = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var v = String(row[vendorIdx] || '').toUpperCase();
    var code = String(row[codeIdx] || '');
    if (!code || !allowedSet[code]) continue;
    if (row[sourceIdx] !== 'manual' && Number(row[confIdx]) < 85) continue;
    if (v !== vendorUp && v.indexOf(vendorUp) === -1 && vendorUp.indexOf(v) === -1) continue;
    counts[code] = (counts[code] || 0) + 1;
  }

  return Object.keys(counts)
    .sort(function(a, b) { return counts[b] - counts[a]; })
    .slice(0, 3)
    .map(function(code) { return { code: code, description: allowedSet[code], count: counts[code] }; });
}

// ---- Month normalization ----

// Sheets often auto-converts "April 2026" into a Date value.
// Normalise to "Month YYYY" so lookups match regardless of how the cell was stored.
function normalizeMonth(m) {
  if (m === null || m === undefined || m === '') return '';
  var names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  if (Object.prototype.toString.call(m) === '[object Date]') {
    return names[m.getMonth()] + ' ' + m.getFullYear();
  }
  var s = String(m).trim();
  // "August 2026", "Aug 2026" and "Sept. 2026" all become "August 2026" —
  // statement file names often abbreviate the month.
  var named = /^([A-Za-z]+)\.?\s+(\d{4})$/.exec(s);
  if (named) {
    var word = named[1].toLowerCase();
    for (var i = 0; i < names.length; i++) {
      if (word.length >= 3 && names[i].toLowerCase().indexOf(word) === 0) return names[i] + ' ' + named[2];
    }
    return s;
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return names[d.getMonth()] + ' ' + d.getFullYear();
  return s;
}

// ---- Upload and process statement ----

function uploadStatement_(base64Pdf, month) {
  var user = getCurrentUser();
  var codes = getUserBudgetCodes_(user);

  // Read and code the statement before taking the lock or touching the sheets:
  // a Claude error or timeout then leaves any existing reconciliation intact,
  // and other users' saves and approvals aren't blocked while the AI runs.
  var statement = parseStatement_(base64Pdf);
  var transactions = statement.transactions;
  var flaggedCount = 0;
  var BATCH = 5;
  var txFields = [];

  for (var b = 0; b < transactions.length; b += BATCH) {
    var batch = transactions.slice(b, b + BATCH);
    batch.forEach(function(tx) {
      // Only charges and fees get a budget code (and an AI suggestion).
      var suggestion = { code: '', confidence: '' };
      if (txNeedsCode_(tx.type)) {
        var history = getVendorHistory_(tx.vendor, codes);
        try {
          suggestion = suggestBudgetCode_(tx.vendor, tx.vendor, tx.total, codes, history);
        } catch (e) {
          suggestion = { code: '', confidence: 0, reason: 'AI unavailable' };
        }
        if (suggestion.confidence < 75) flaggedCount++;
      }

      // 13% HST on purchases and refunds; payments and card fees/interest carry none.
      var total = tx.total;
      var hst = (tx.type === 'charge' || tx.type === 'refund')
        ? Math.round((total - (total / 1.13)) * 100) / 100 : 0;
      var subtotal = Math.round((total - hst) * 100) / 100;

      txFields.push({
        tx_id: generateId('TX'),
        date: tx.date, vendor: tx.vendor, description: tx.vendor,
        total: total, hst: hst, subtotal: subtotal,
        budget_code: suggestion.code,
        code_source: txNeedsCode_(tx.type) ? 'ai' : '',
        ai_confidence: suggestion.confidence,
        receipt_file_id: '',
        // Only charges need a receipt.
        receipt_status: tx.type === 'charge' ? 'unmatched' : 'not_required',
        duplicate_flag: false,
        created_at: new Date().toISOString(),
        tx_type: tx.type
      });
    });

    // Brief pause between batches to avoid GAS timeout
    if (b + BATCH < transactions.length) Utilities.sleep(500);
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var reconSheet = ensureColumns_(SHEET_NAMES.RECONCILIATIONS);
    var txSheet = ensureColumns_(SHEET_NAMES.TRANSACTIONS);

    // Find or create reconciliation for this user + month
    var reconData = reconSheet.getDataRange().getValues();
    var reconHeaders = reconData[0];
    var rcx = colIndexes_(reconHeaders);
    var reconId = null;
    var reconRow = -1;

    var normalizedTarget = normalizeMonth(month);
    for (var i = 1; i < reconData.length; i++) {
      var row = reconData[i];
      if (row[rcx.user_id] === user.user_id &&
          normalizeMonth(row[rcx.month]) === normalizedTarget) {
        reconId = row[rcx.recon_id];
        reconRow = i + 1;
        break;
      }
    }

    if (!reconId) {
      reconId = generateId('REC');
      var now = new Date().toISOString();
      reconSheet.appendRow(rowFromFields_(reconHeaders, {
        recon_id: reconId, user_id: user.user_id, month: normalizedTarget || month, status: 'draft',
        created_at: now, last_edited_at: now
      }));
      reconRow = reconSheet.getLastRow();
    }

    // Statement summary for the balance check: previous balance + all lines = balance owing.
    reconSheet.getRange(reconRow, rcx.previous_balance + 1).setValue(statement.previous_balance);
    reconSheet.getRange(reconRow, rcx.statement_balance + 1).setValue(statement.balance_owing);

    // Clear any existing transactions for this recon before re-importing
    var existingData = txSheet.getDataRange().getValues();
    var txHeaders = existingData[0];
    var reconIdIdx = txHeaders.indexOf('recon_id');
    for (var d = existingData.length - 1; d >= 1; d--) {
      if (existingData[d][reconIdIdx] === reconId) {
        txSheet.deleteRow(d + 1);
      }
    }

    // Write all parsed transactions in a single bulk operation.
    var newRows = txFields.map(function(f) {
      f.recon_id = reconId;
      return rowFromFields_(txHeaders, f);
    });
    if (newRows.length > 0) {
      txSheet.getRange(txSheet.getLastRow() + 1, 1, newRows.length, txHeaders.length)
        .setValues(newRows);
    }

    checkDuplicates_(reconId);

    return { recon_id: reconId, tx_count: newRows.length, flagged_count: flaggedCount };
  } finally {
    lock.releaseLock();
  }
}

// ---- Read reconciliation ----

function getReconciliation(reconId) {
  var user = getCurrentUser();
  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var codeSheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);

  var reconData = reconSheet.getDataRange().getValues();
  var reconHeaders = reconData[0];
  var ridIdx = reconHeaders.indexOf('recon_id');
  var recon = null;

  for (var i = 1; i < reconData.length; i++) {
    if (reconData[i][ridIdx] === reconId) {
      recon = rowToObj_(reconHeaders, reconData[i]);
      break;
    }
  }

  if (!recon) throw new Error('Reconciliation not found.');

  // Access check: only finance/ed may open other users' reconciliations
  // (exec is scoped to their own, same as staff).
  if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) {
    throw new Error('Access denied.');
  }

  // Build code description lookup
  var codeData = codeSheet.getDataRange().getValues();
  var ccx = colIndexes_(codeData[0]);
  var codeMap = {};
  for (var c = 1; c < codeData.length; c++) {
    codeMap[codeData[c][ccx.code]] = codeData[c][ccx.description];
  }

  // Fetch transactions using the same retrieval used by submitReconciliation
  var txs = getReconTransactions_(reconId);
  txs.forEach(function(tx) {
    tx.budget_code_description = codeMap[tx.budget_code] || '';
  });
  txs.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
  recon.transactions = txs;
  recon.available_codes = getUserBudgetCodes_(user);
  return jsonSafe(recon);
}

// ---- Save a single field on a transaction ----

function saveTransaction(txId, field, newValue, overrideReason) {
  if (field !== 'description' && field !== 'budget_code') {
    throw new Error('Invalid field: ' + field);
  }

  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = txSheet.getDataRange().getValues();
  var headers = data[0];
  var txIdIdx = headers.indexOf('tx_id');
  var reconIdIdx = headers.indexOf('recon_id');
  var fieldIdx = headers.indexOf(field === 'description' ? 'description' : 'budget_code');
  var codeSourceIdx = headers.indexOf('code_source');

  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var reconData = reconSheet.getDataRange().getValues();
  var rcx = colIndexes_(reconData[0]);

  for (var i = 1; i < data.length; i++) {
    if (data[i][txIdIdx] === txId) {
      var reconId = data[i][reconIdIdx];

      // Find recon to check status and ownership; fails closed if there is none.
      var reconFound = false;
      for (var r = 1; r < reconData.length; r++) {
        if (reconData[r][rcx.recon_id] === reconId) {
          reconFound = true;
          var status = reconData[r][rcx.status];
          var ownerId = reconData[r][rcx.user_id];

          if (['finance', 'ed'].indexOf(user.role) === -1 && ownerId !== user.user_id) {
            throw new Error('Access denied.');
          }
          var isFinanceEd = (user.role === 'finance' || user.role === 'ed');
          if (['submitted', 'finance_review', 'ed_review', 'approved'].indexOf(status) !== -1) {
            if (isFinanceEd && field === 'budget_code' && status !== 'approved') {
              // Finance/ED can correct codes on in-review reconciliations
            } else {
              throw new Error('Reconciliation is locked and cannot be edited.');
            }
          }
          break;
        }
      }
      if (!reconFound) throw new Error('Reconciliation not found.');

      var oldValue = data[i][fieldIdx];

      if (field === 'budget_code' && newValue) {
        var allowedCodes = getUserBudgetCodes_(user).map(function(c) { return c.code; });
        if (allowedCodes.indexOf(newValue) === -1) {
          if (!overrideReason || !overrideReason.trim()) {
            throw new Error('Budget code ' + newValue + ' is outside your department. An override reason is required.');
          }
          logAudit_(user.user_id, reconId, txId, 'budget_code_override', oldValue, overrideReason);
        }
      }

      txSheet.getRange(i + 1, fieldIdx + 1).setValue(newValue);

      if (field === 'budget_code') {
        txSheet.getRange(i + 1, codeSourceIdx + 1).setValue('manual');
      }

      logAudit_(user.user_id, data[i][reconIdIdx], txId, field + '_changed', oldValue, newValue);
      try { updateRecon_(reconId, { last_edited_at: new Date().toISOString() }); } catch (e) {}
      return { success: true };
    }
  }
  throw new Error('Transaction not found: ' + txId);
}

// ---- Delete a transaction (used by "confirm duplicate") ----

function deleteTransaction(txId) {
  var user = getCurrentUser();
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var data = txSheet.getDataRange().getValues();
  var headers = data[0];
  var txIdIdx = headers.indexOf('tx_id');
  var rowIdx = -1;
  var deletedTx = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][txIdIdx] === txId) {
      rowIdx = i;
      deletedTx = rowToObj_(headers, data[i]);
      break;
    }
  }
  if (!deletedTx) throw new Error('Transaction not found.');

  var reconId = deletedTx.recon_id;

  // Permissions: finance/ed always; owner only on their own unlocked recon
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
    if (!allowed) throw new Error('You cannot delete this transaction.');
  }

  txSheet.deleteRow(rowIdx + 1);
  logAudit_(user.user_id, reconId, txId, 'transaction_deleted', deletedTx.vendor + ' $' + deletedTx.total, '');

  // If the deleted tx had a duplicate pair and only one peer remains, clear that peer's flag.
  if (deletedTx.duplicate_flag) {
    var freshData = txSheet.getDataRange().getValues();
    var fHeaders = freshData[0];
    var vIdx = fHeaders.indexOf('vendor');
    var tIdx = fHeaders.indexOf('total');
    var dIdx = fHeaders.indexOf('date');
    var dupIdx = fHeaders.indexOf('duplicate_flag');
    var rIdx = fHeaders.indexOf('recon_id');
    var peers = [];
    for (var k = 1; k < freshData.length; k++) {
      if (freshData[k][rIdx] !== reconId) continue;
      if (freshData[k][vIdx] === deletedTx.vendor &&
          Math.abs(Number(freshData[k][tIdx]) - Number(deletedTx.total)) < 0.01 &&
          String(freshData[k][dIdx]) === String(deletedTx.date)) {
        peers.push(k + 1);
      }
    }
    if (peers.length === 1) txSheet.getRange(peers[0], dupIdx + 1).setValue(false);
  }

  try { updateRecon_(reconId, { last_edited_at: new Date().toISOString() }); } catch (e) {}
  return { success: true };
}

// ---- Drive statement import ----

// Statement filenames: "[Name] VISA Statement [Month] [Year].pdf". The space
// before VISA is optional, since long names get typed as "MarteenaVISA Statement".
var STATEMENT_NAME_RE = /^(.+?)\s*VISA\s+Statement\s+(.+)\.pdf$/i;

// Zero-width and formatting characters that survive trim() and pasteboards —
// stripped before any name comparison so a copy-pasted Users-sheet name can't
// silently break matching.
var INVISIBLE_CHARS_RE = /[\u200B-\u200F\u2060-\u2064\u00AD\u061C\uFEFF]/g;

function cleanNameForMatch_(s) {
  return String(s == null ? '' : s).replace(INVISIBLE_CHARS_RE, '').trim().toLowerCase();
}

// True when a statement's card-holder label refers to this user. Accepts an
// exact full-name match, a word-boundary prefix in either direction (file
// "Fei" vs user "Fei Kokwu", or file "Fei Kokwu Jr" vs user "Fei Kokwu" —
// but never "Fei" vs "Feith"), a first-name match (keeps finance's
// first-name-only filenames working), and finally the first segment of the
// user's email local part (file "Jordan" matches jordan.lee@example.org even
// when the sheet name is wrong).
function statementMatchesUser_(cardHolder, user) {
  var holder = cleanNameForMatch_(cardHolder);
  var full = cleanNameForMatch_(user && user.name);
  if (!holder) return false;
  if (full) {
    if (holder === full) return true;
    if (full.indexOf(holder + ' ') === 0) return true;
    if (holder.indexOf(full + ' ') === 0) return true;
    if (holder.split(/\s+/)[0] === full.split(/\s+/)[0]) return true;
  }
  var local = cleanNameForMatch_(user && user.email).split('@')[0];
  var localFirst = local.split(/[._-]/)[0];
  return !!localFirst && holder.split(/\s+/)[0] === localFirst;
}

// Resolve the statements folder the same way everywhere: the optional
// STATEMENTS_FOLDER_ID admin override, else the auto-provisioned folder.
function getStatementsFolderId_() {
  var folderId = PropertiesService.getScriptProperties().getProperty('STATEMENTS_FOLDER_ID');
  return folderId || ensureStatementsFolder_();
}

// Make every non-ASCII character visible as \uXXXX. JSON.stringify passes
// zero-width characters and homoglyphs through invisibly, so name/email
// diagnostics must go through this instead.
function asciiEscape_(s) {
  return String(s == null ? '' : s).replace(/[^\x20-\x7E]/g, function(c) {
    return '\\u' + ('000' + c.charCodeAt(0).toString(16)).slice(-4);
  });
}

function getAvailableStatements() {
  var user = getCurrentUser();
  // Executions-panel breadcrumb: shows WHO the server resolved this visitor
  // to (auto-provisioned local-part names and invisible characters included).
  console.log('getAvailableStatements for name="' + asciiEscape_(user.name) +
              '" email="' + asciiEscape_(user.email) + '"');
  // STATEMENTS_FOLDER_ID stays supported as an admin override; by default use
  // the auto-provisioned VISA_REC_APP/Visa Statements folder (flat — finance
  // drops '[Name] VISA Statement [Month] [Year].pdf' files in the root).
  // Best-effort: a Drive failure must never break the dashboard.
  var folderId = null;
  try { folderId = getStatementsFolderId_(); }
  catch (e) { Logger.log('Could not resolve statements folder: ' + e.message); }
  if (!folderId) return [];

  // Build set of months already imported for this user
  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var reconData = reconSheet.getDataRange().getValues();
  var rcx = colIndexes_(reconData[0]);
  var importedMonths = {};
  for (var i = 1; i < reconData.length; i++) {
    if (reconData[i][rcx.user_id] === user.user_id) {
      importedMonths[normalizeMonth(reconData[i][rcx.month])] = true;
    }
  }

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.PDF);
  var results = [];

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    var match = name.match(STATEMENT_NAME_RE);
    if (!match) continue;
    if (!statementMatchesUser_(match[1], user)) continue;
    var month = match[2].trim();
    results.push({
      fileId: file.getId(),
      fileName: name,
      month: month,
      imported: !!importedMonths[normalizeMonth(month)]
    });
  }

  // Not-yet-imported statements first, then alphabetical for a stable order.
  results.sort(function(a, b) {
    if (a.imported !== b.imported) return a.imported ? 1 : -1;
    return a.fileName.localeCompare(b.fileName);
  });

  console.log('getAvailableStatements matched ' + results.length + ' file(s)');
  return jsonSafe(results);
}

// Throws unless the file lives in the statements folder AND its card-holder
// name matches the caller. Guards both import and open against arbitrary
// fileIds (the web app runs as the deploying owner, who can read far more
// of Drive than any individual caller should).
function assertCallersStatement_(fileId, user) {
  var folderId = getStatementsFolderId_();
  var meta = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,name,parents,webViewLink'
  });
  if ((meta.parents || []).indexOf(folderId) === -1) {
    throw new Error('That file is not in the statements folder.');
  }
  var m = String(meta.name || '').match(STATEMENT_NAME_RE);
  if (!m || !statementMatchesUser_(m[1], user)) {
    throw new Error('Access denied: this statement is not yours.');
  }
  return meta;
}

// Grants the caller read access to their own statement PDF and returns its
// Drive URL for the client to open in a new tab.
function openStatement(fileId) {
  var user = getCurrentUser();
  var meta = assertCallersStatement_(fileId, user);
  grantDriveAccess_(fileId, user.email, 'reader');
  return jsonSafe({ url: meta.webViewLink || '' });
}

function importStatementFromDrive(fileId) {
  var user = getCurrentUser();
  var meta = assertCallersStatement_(fileId, user);
  // The month comes from the statement's file name, never from the browser.
  var month = normalizeMonth(String(meta.name).match(STATEMENT_NAME_RE)[2].trim());
  var file = DriveApp.getFileById(fileId);
  var base64 = Utilities.base64Encode(file.getBlob().getBytes());
  return uploadStatement_(base64, month);
}

// ---- New-statement emails ----

// Emails each cardholder once when a statement PDF matching them lands in the
// statements folder. Run hourly by setupStatementNotifyTrigger(). "Already
// emailed" is stored on the file itself (Drive appProperties), so it survives
// renames and fiscal-year archive moves. The first run only marks statements
// already in the folder, so existing ones don't all send an email.
var STATEMENT_EMAILED_KEY = 'visaRecEmailed';

function notifyNewStatements(e) {
  requireTriggerOrOwner_(e);
  var props = PropertiesService.getScriptProperties();
  var seeding = !props.getProperty('STATEMENT_EMAILS_SEEDED');
  var folderId = getStatementsFolderId_();

  var userData = getOrCreateSheet_(SHEET_NAMES.USERS).getDataRange().getValues();
  var ux = colIndexes_(userData[0]);
  var users = [];
  for (var i = 1; i < userData.length; i++) {
    var email = String(userData[i][ux.email] || '').trim();
    if (userData[i][ux.active] && email) users.push({ name: userData[i][ux.name], email: email });
  }

  var appUrl = getAppUrl();
  var marker = {};
  marker[STATEMENT_EMAILED_KEY] = 'true';
  var pageToken = null;
  do {
    var res = Drive.Files.list({
      q: "trashed=false and mimeType='application/pdf' and '" + folderId + "' in parents",
      corpora: 'allDrives',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      fields: 'nextPageToken,files(id,name,appProperties)',
      pageSize: 200,
      pageToken: pageToken
    });
    (res.files || []).forEach(function(f) {
      if ((f.appProperties || {})[STATEMENT_EMAILED_KEY]) return;
      if (!seeding && !emailStatementHolders_(f.name, users, appUrl)) return;
      Drive.Files.update({ appProperties: marker }, f.id, null, { supportsAllDrives: true });
    });
    pageToken = res.nextPageToken;
  } while (pageToken);

  if (seeding) props.setProperty('STATEMENT_EMAILS_SEEDED', 'true');
}

// Emails every active user the card-holder name matches (the same rule that
// lists the statement on their dashboard). Returns true if any email was sent;
// false leaves the file unmarked so it's retried next run — e.g. after a
// misnamed file or a Users-sheet name is fixed.
function emailStatementHolders_(fileName, users, appUrl) {
  var m = String(fileName || '').match(STATEMENT_NAME_RE);
  if (!m) return false;
  var month = normalizeMonth(m[2].trim());
  var sent = false;
  users.forEach(function(u) {
    if (!statementMatchesUser_(m[1], u)) return;
    var result = sendEmail_(u.email,
      'Your ' + month + ' Visa statement is ready',
      'Hi ' + (u.name || 'there') + ',\n\n' +
      'Your ' + month + ' Visa statement has been added. Sign in to import and reconcile it:\n' + appUrl);
    if (result.sent) sent = true;
  });
  return sent;
}

// Run once from the Apps Script editor to start the hourly statement emails.
function setupStatementNotifyTrigger() {
  requireRole(['finance', 'ed']);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'notifyNewStatements') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('notifyNewStatements').timeBased().everyHours(1).create();
  Logger.log('Hourly new-statement email trigger set.');
}

// Finance/ED diagnostic: every statement PDF in the folder and exactly which
// user account(s) it matches, plus per-user flags that explain "I can't see
// my statement": auto-provisioned local-part names, duplicate emails, and
// invisible characters (names/emails are ASCII-escaped so they become
// visible). Uses the same STATEMENT_NAME_RE + statementMatchesUser_ as the
// staff-facing list, so this view can never disagree with what staff see.
function getStatementsOverview() {
  requireRole(['finance', 'ed']);

  var userData = getOrCreateSheet_(SHEET_NAMES.USERS).getDataRange().getValues();
  var ux = colIndexes_(userData[0]);
  var rawUsers = [];   // exact sheet values, fed to statementMatchesUser_
  var users = [];      // escaped display copies, returned to the client
  var emailsSeen = {};
  var duplicateEmails = [];
  for (var i = 1; i < userData.length; i++) {
    var row = userData[i];
    var name = String(row[ux.name] || '');
    var email = String(row[ux.email] || '');
    var normEmail = email.trim().toLowerCase();
    if (normEmail) {
      if (emailsSeen[normEmail]) duplicateEmails.push(asciiEscape_(email));
      emailsSeen[normEmail] = true;
    }
    rawUsers.push({ name: name, email: email });
    users.push({
      user_id: row[ux.user_id],
      name: asciiEscape_(name),
      email: asciiEscape_(email),
      role: row[ux.role],
      active: !!row[ux.active],
      looks_auto_provisioned: !!name && name.trim().toLowerCase() === normEmail.split('@')[0],
      matched_file_count: 0
    });
  }

  var overview = {
    files: [],
    users: users,
    duplicate_emails: duplicateEmails,
    statements_folder_prop: PropertiesService.getScriptProperties().getProperty('STATEMENTS_FOLDER_ID') || '',
    app_url: getAppUrl(),
    folder_id: ''
  };

  var folderId = null;
  try { folderId = getStatementsFolderId_(); }
  catch (e) { overview.error = 'Could not resolve statements folder: ' + e.message; }
  overview.folder_id = folderId || '';

  if (folderId) {
    var files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      var fname = files.next().getName();
      var m = fname.match(STATEMENT_NAME_RE);
      var entry = {
        fileName: fname,
        unparsed: !m,
        cardholder: m ? m[1].trim() : '',
        month: m ? m[2].trim() : '',
        matches: [],
        inactive_matches: []
      };
      if (m) {
        for (var u = 0; u < rawUsers.length; u++) {
          if (!statementMatchesUser_(m[1], rawUsers[u])) continue;
          var label = users[u].name + ' <' + users[u].email + '>';
          if (users[u].active) {
            entry.matches.push(label);
            users[u].matched_file_count++;
          } else {
            entry.inactive_matches.push(label);
          }
        }
      }
      overview.files.push(entry);
    }
    overview.files.sort(function(a, b) { return a.fileName.localeCompare(b.fileName); });
  }

  return jsonSafe(overview);
}

// ---- Dashboard ----

function getDashboard() {
  var user = getCurrentUser();
  var reconSheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var userSheet = getOrCreateSheet_(SHEET_NAMES.USERS);

  var reconData = reconSheet.getDataRange().getValues();
  var reconHeaders = reconData[0];
  var txData = txSheet.getDataRange().getValues();
  var txHeaders = txData[0];
  var userData = userSheet.getDataRange().getValues();
  var userHeaders = userData[0];

  // Hoist column indexes out of the row loops below.
  var uUserIdIdx = userHeaders.indexOf('user_id');
  var uNameIdx = userHeaders.indexOf('name');
  var txReconIdx = txHeaders.indexOf('recon_id');
  var txConfIdx = txHeaders.indexOf('ai_confidence');
  var txTotalIdx = txHeaders.indexOf('total');
  var txHstIdx = txHeaders.indexOf('hst');
  var txSrcIdx = txHeaders.indexOf('code_source');
  var txCodeIdx = txHeaders.indexOf('budget_code');
  var txRcptIdx = txHeaders.indexOf('receipt_status');
  var txDupIdx = txHeaders.indexOf('duplicate_flag');
  var txTypeIdx = txHeaders.indexOf('tx_type');
  var rUserIdIdx = reconHeaders.indexOf('user_id');
  var rReconIdIdx = reconHeaders.indexOf('recon_id');

  // Build user name lookup
  var userMap = {};
  for (var u = 1; u < userData.length; u++) {
    userMap[userData[u][uUserIdIdx]] = userData[u][uNameIdx];
  }

  // Build tx stats per recon
  var txStats = {};
  for (var t = 1; t < txData.length; t++) {
    var rId = txData[t][txReconIdx];
    if (!txStats[rId]) txStats[rId] = { count: 0, flagged: 0 };
    txStats[rId].count++;
    var conf = txData[t][txConfIdx];
    var rcpt = txData[t][txRcptIdx];
    var dup = txData[t][txDupIdx];
    // Same rule as the editor: low AI confidence (only on rows that need a
    // code, never manually coded ones), a missing receipt, or a possible duplicate.
    var lowConf = txNeedsCode_(txTypeOf_(txData[t][txTypeIdx])) &&
      txData[t][txSrcIdx] !== 'manual' && (Number(conf) || 0) < 75;
    if (lowConf || rcpt === 'unmatched' || dup) txStats[rId].flagged++;
  }

  // Split recons into the viewer's own vs everyone else's. Only finance/ed
  // get the staff set (their Staff Recs tab); staff and exec see own-only.
  var isReviewer = ['finance', 'ed'].indexOf(user.role) !== -1;
  var results = [];
  var staffResults = [];
  for (var i = 1; i < reconData.length; i++) {
    var row = reconData[i];
    var rUserId = row[rUserIdIdx];
    var isMine = rUserId === user.user_id;
    if (!isMine && !isReviewer) continue;

    var rId2 = row[rReconIdIdx];
    var stats = txStats[rId2] || { count: 0, flagged: 0 };
    var entry = {};
    reconHeaders.forEach(function(h, j) { entry[h] = row[j]; });
    entry.user_name = userMap[rUserId] || rUserId;
    entry.tx_count = stats.count;
    entry.flagged_count = stats.flagged;
    (isMine ? results : staffResults).push(entry);
  }

  var byCreatedDesc = function(a, b) {
    return String(b.created_at).localeCompare(String(a.created_at));
  };
  results.sort(byCreatedDesc);
  staffResults.sort(byCreatedDesc);

  var txIx = {
    recon: txReconIdx, total: txTotalIdx, hst: txHstIdx,
    src: txSrcIdx, conf: txConfIdx, code: txCodeIdx,
    rcpt: txRcptIdx, dup: txDupIdx, type: txTypeIdx
  };
  var payload = {
    recons: results,
    stats: computeReconSetStats_(txData, txIx, results)
  };
  if (isReviewer) {
    payload.staffRecons = staffResults;
    payload.staffStats = computeReconSetStats_(txData, txIx, staffResults);
  }
  return jsonSafe(payload);
}

// Aggregate transaction stats over one set of reconciliations.
function computeReconSetStats_(txData, txIx, recons) {
  var accessible = {};
  recons.forEach(function(r) { accessible[r.recon_id] = true; });
  var stats = { total_charges: 0, hst: 0, auto_coded: 0, needs_review: 0 };
  for (var k = 1; k < txData.length; k++) {
    if (!accessible[txData[k][txIx.recon]]) continue;
    var typeK = txTypeOf_(txData[k][txIx.type]);
    if (typeK !== 'payment') stats.total_charges += Number(txData[k][txIx.total]) || 0;
    stats.hst += Number(txData[k][txIx.hst]) || 0;
    var srcK = txData[k][txIx.src];
    if (srcK === 'ai' && txData[k][txIx.code]) stats.auto_coded++;
    var lowConfK = txNeedsCode_(typeK) && srcK !== 'manual' && (Number(txData[k][txIx.conf]) || 0) < 75;
    if (lowConfK || txData[k][txIx.rcpt] === 'unmatched' || txData[k][txIx.dup]) stats.needs_review++;
  }
  return stats;
}
