// ============================================================
// Story 2.3 — Export reconciliation in LUSU format
// ============================================================

function getFormatTemplateUrl() {
  return PropertiesService.getScriptProperties().getProperty('FORMAT_TEMPLATE_URL') || null;
}

function createFormatTemplate() {
  requireOwner_();
  var sampleTxs = [
    { date: '2026-04-01', vendor: 'STAPLES', description: 'Office supplies', total: 45.20, hst: 5.21, subtotal: 39.99, budget_code: '58352-1000' },
    { date: '2026-04-03', vendor: 'BEST WESTERN', description: 'Conference travel', total: 226.00, hst: 26.02, subtotal: 199.98, budget_code: '58510-1000' },
    { date: '2026-04-07', vendor: 'DOLLARAMA', description: 'Event supplies', total: 33.90, hst: 3.90, subtotal: 30.00, budget_code: '58000-1200' }
  ];
  var ownerName = 'Sample Cardholder';
  var month = 'April 2026';

  var ss = SpreadsheetApp.create('LUSU Visa Reconciliation — Format Reference');
  var sheet = ss.getActiveSheet();
  sheet.setName('Reconciliation');

  var n = sampleTxs.length;
  var dataStart = 5;
  var dataEnd = dataStart + n - 1;
  var blankRows = 2;
  var totalsRow = dataEnd + blankRows + 1;
  var footerRow = totalsRow + 3;

  sheet.setColumnWidth(1, 96); sheet.setColumnWidth(2, 208);
  sheet.setColumnWidth(3, 384); sheet.setColumnWidth(4, 96);
  sheet.setColumnWidth(5, 80); sheet.setColumnWidth(6, 96); sheet.setColumnWidth(7, 112);

  sheet.setRowHeight(1, 30);
  sheet.getRange('A1:G1').merge().setValue('VISA STATEMENT RECONCILIATION')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#595959')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('A2:C2').merge().setValue('Claimant: ' + ownerName)
    .setBackground('#D9D9D9').setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange('D2:E2').merge().setValue('Month / Year:')
    .setBackground('#D9D9D9').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F2:G2').merge().setValue(month).setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setRowHeight(3, 6);
  sheet.setRowHeight(4, 34);
  var colHeaders = ['DATE (d/m/year)','PURCHASE PLACE','PURPOSE OF EXPENSE / DESCRIPTION','TOTAL','HST','SUB-TOTAL','BUDGET CODE'];
  sheet.getRange(4, 1, 1, 7).setValues([colHeaders])
    .setBackground('#D9D9D9').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,true,true,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var moneyFmt = '$#,##0.00';
  sampleTxs.forEach(function(tx, i) {
    var row = dataStart + i;
    sheet.setRowHeight(row, 20);
    sheet.getRange(row, 1, 1, 7).setBackground(i % 2 === 0 ? '#FFFFFF' : '#FAFAFA')
      .setBorder(true,true,true,true,true,true,'#EEEEEE',SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, 1).setValue(tx.date);
    sheet.getRange(row, 2).setValue(tx.vendor);
    sheet.getRange(row, 3).setValue(tx.description);
    sheet.getRange(row, 4).setValue(tx.total).setNumberFormat(moneyFmt);
    sheet.getRange(row, 5).setValue(tx.hst).setNumberFormat(moneyFmt);
    sheet.getRange(row, 6).setValue(tx.subtotal).setNumberFormat(moneyFmt);
    sheet.getRange(row, 7).setValue(tx.budget_code);
  });

  for (var br = dataEnd + 1; br <= dataEnd + blankRows; br++) {
    sheet.getRange(br, 1, 1, 7).setBorder(true,true,true,true,true,true,'#EEEEEE',SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.getRange(totalsRow, 1, 1, 3).merge().setValue('TOTALS')
    .setFontWeight('bold').setBackground('#D9D9D9')
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ['D','E','F'].forEach(function(col) {
    sheet.getRange(col + totalsRow)
      .setFormula('=SUM(' + col + dataStart + ':' + col + (dataEnd + blankRows) + ')')
      .setFontWeight('bold').setBackground('#D9D9D9').setNumberFormat(moneyFmt)
      .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange('G' + totalsRow).setBackground('#D9D9D9')
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var fr = footerRow;
  sheet.getRange(fr, 1, 5, 4).merge().setValue('NOTES:').setVerticalAlignment('top')
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('E' + fr).setValue('TOTAL').setBackground('#D9D9D9').setFontWeight('bold')
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('F' + fr + ':G' + fr).merge().setFormula('=D' + totalsRow).setNumberFormat(moneyFmt)
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('E' + (fr+1)).setValue('Statement').setBackground('#D9D9D9').setFontWeight('bold');
  sheet.getRange('F' + (fr+1) + ':G' + (fr+1)).merge().setBackground('#FFFF00')
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('E' + (fr+2)).setValue('Variance').setBackground('#D9D9D9').setFontWeight('bold');
  sheet.getRange('F' + (fr+2) + ':G' + (fr+2)).merge()
    .setFormula('=F' + (fr+1) + '-F' + fr).setBackground('#FFFF00').setNumberFormat(moneyFmt)
    .setBorder(true,true,true,true,null,null,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('E' + (fr+4) + ':G' + (fr+4)).merge().setValue('Claimant Signature: ___________________________');
  sheet.getRange('E' + (fr+6) + ':G' + (fr+6)).merge().setValue('Approval Signature: ___________________________');
  sheet.getRange('E' + (fr+7) + ':G' + (fr+7)).merge().setValue('Approval Signature: ___________________________');

  var url = ss.getUrl();
  PropertiesService.getScriptProperties().setProperty('FORMAT_TEMPLATE_URL', url);
  Logger.log('Format template created: ' + url);
  return url;
}

function exportReconciliation(reconId) {
  var user = getCurrentUser();
  var recon = getReconById_(reconId);
  if (!recon) throw new Error('Reconciliation not found.');
  if (['finance', 'ed'].indexOf(user.role) === -1 && recon.user_id !== user.user_id) throw new Error('Access denied.');

  var owner = getOwnerEmail_(recon.user_id);
  var ownerName = owner ? owner.name : 'Unknown';
  var txs = getReconTransactions_(reconId);
  txs.sort(function(a, b) { return String(a.date).localeCompare(String(b.date)); });

  var ssName = 'Visa Reconciliation — ' + ownerName + ' — ' + recon.month;
  var ss = SpreadsheetApp.create(ssName);
  var sheet = ss.getActiveSheet();
  sheet.setName('Reconciliation');

  var n = txs.length;
  var dataStart = 5;
  var dataEnd = dataStart + n - 1;
  var blankRows = 2;
  var totalsRow = dataEnd + blankRows + 1;
  var footerRow = totalsRow + 3;

  // ---- Column widths ----
  sheet.setColumnWidth(1, 96);   // A DATE
  sheet.setColumnWidth(2, 208);  // B VENDOR
  sheet.setColumnWidth(3, 384);  // C DESCRIPTION
  sheet.setColumnWidth(4, 96);   // D TOTAL
  sheet.setColumnWidth(5, 80);   // E HST
  sheet.setColumnWidth(6, 96);   // F SUB-TOTAL
  sheet.setColumnWidth(7, 112);  // G BUDGET CODE

  // ---- ROW 1 — title ----
  sheet.setRowHeight(1, 30);
  var titleRange = sheet.getRange('A1:G1');
  titleRange.merge()
    .setValue('VISA STATEMENT RECONCILIATION')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#595959')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // ---- ROW 2 — claimant / month ----
  sheet.getRange('A2:C2').merge().setValue('Claimant: ' + ownerName)
    .setBackground('#D9D9D9').setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange('D2:E2').merge().setValue('Month / Year:')
    .setBackground('#D9D9D9').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F2:G2').merge().setValue(recon.month)
    .setFontWeight('bold').setHorizontalAlignment('center');

  // ---- ROW 3 — spacer ----
  sheet.setRowHeight(3, 6);

  // ---- ROW 4 — headers ----
  sheet.setRowHeight(4, 34);
  var colHeaders = ['DATE (d/m/year)','PURCHASE PLACE','PURPOSE OF EXPENSE / DESCRIPTION','TOTAL','HST','SUB-TOTAL','BUDGET CODE'];
  var hdrRange = sheet.getRange(4, 1, 1, 7);
  hdrRange.setValues([colHeaders])
    .setBackground('#D9D9D9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ---- DATA ROWS ----
  var moneyFmt = '$#,##0.00';
  txs.forEach(function(tx, i) {
    var row = dataStart + i;
    var bg = i % 2 === 0 ? '#FFFFFF' : '#FAFAFA';
    sheet.setRowHeight(row, 20);
    var r = sheet.getRange(row, 1, 1, 7);
    r.setBackground(bg)
      .setBorder(true, true, true, true, true, true, '#EEEEEE', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(row, 1).setValue(tx.date);
    sheet.getRange(row, 2).setValue(tx.vendor);
    sheet.getRange(row, 3).setValue(tx.description);
    sheet.getRange(row, 4).setValue(tx.total).setNumberFormat(moneyFmt);
    if (tx.hst) sheet.getRange(row, 5).setValue(tx.hst).setNumberFormat(moneyFmt);
    sheet.getRange(row, 6).setValue(tx.subtotal).setNumberFormat(moneyFmt);
    sheet.getRange(row, 7).setValue(tx.budget_code);
  });

  // ---- BLANK ROWS (for manual additions) ----
  for (var br = dataEnd + 1; br <= dataEnd + blankRows; br++) {
    sheet.getRange(br, 1, 1, 7)
      .setBorder(true, true, true, true, true, true, '#EEEEEE', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ---- TOTALS ROW ----
  var sumRange = 'D' + dataStart + ':D' + (dataEnd + blankRows);
  sheet.getRange(totalsRow, 1, 1, 3).merge().setValue('TOTALS')
    .setFontWeight('bold').setBackground('#D9D9D9')
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ['D','E','F'].forEach(function(col, ci) {
    var srcCol = ['D','E','F'][ci];
    var formula = '=SUM(' + srcCol + dataStart + ':' + srcCol + (dataEnd + blankRows) + ')';
    sheet.getRange(col + totalsRow).setFormula(formula)
      .setFontWeight('bold').setBackground('#D9D9D9').setNumberFormat(moneyFmt)
      .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange('G' + totalsRow).setBackground('#D9D9D9')
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ---- FOOTER ----
  var fr = footerRow;

  // Notes area (A:D, 5 rows)
  sheet.getRange(fr, 1, 5, 4).merge().setValue('NOTES:')
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Right side: statement balance check. TOTAL covers every line, payments
  // included, so previous balance + TOTAL should equal the balance owing.
  var totalRef = 'D' + totalsRow;
  var owing = (recon.bank_statement_total !== '' && recon.bank_statement_total != null)
    ? recon.bank_statement_total : recon.statement_balance;

  sheet.getRange('E' + fr + ':E' + fr).setValue('TOTAL').setBackground('#D9D9D9').setFontWeight('bold')
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('F' + fr + ':G' + fr).merge().setFormula('=' + totalRef).setNumberFormat(moneyFmt)
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange('E' + (fr + 1)).setValue('Previous balance').setBackground('#D9D9D9').setFontWeight('bold');
  sheet.getRange('F' + (fr + 1) + ':G' + (fr + 1)).merge().setValue(recon.previous_balance == null ? '' : recon.previous_balance)
    .setNumberFormat(moneyFmt)
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange('E' + (fr + 2)).setValue('Balance owing').setBackground('#D9D9D9').setFontWeight('bold');
  sheet.getRange('F' + (fr + 2) + ':G' + (fr + 2)).merge().setValue(owing == null ? '' : owing)
    .setNumberFormat(moneyFmt).setBackground('#FFFF00')
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange('E' + (fr + 3)).setValue('Variance').setBackground('#D9D9D9').setFontWeight('bold');
  sheet.getRange('F' + (fr + 3) + ':G' + (fr + 3)).merge()
    .setFormula('=F' + (fr + 2) + '-(F' + fr + '+F' + (fr + 1) + ')').setBackground('#FFFF00').setNumberFormat(moneyFmt)
    .setBorder(true, true, true, true, null, null, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange('E' + (fr + 5) + ':G' + (fr + 5)).merge().setValue('Claimant Signature: ___________________________');
  sheet.getRange('E' + (fr + 7) + ':G' + (fr + 7)).merge().setValue('Approval Signature: ___________________________');
  sheet.getRange('E' + (fr + 8) + ':G' + (fr + 8)).merge().setValue('Approval Signature: ___________________________');

  // File the export into the Shared Drive (reparent via the Drive API so it lands
  // reliably inside the Shared Drive). Approved reconciliations are archived
  // centrally under "Approved Visa Recs/<Month>"; everything else goes to the
  // staff department folder for that month.
  try {
    if (recon.status === 'approved') {
      var approvedFolderId = ensureApprovedMonthFolder_(recon.month);
      trashExistingByName_(approvedFolderId, ssName); // avoid duplicates on re-export
      moveFileToFolder_(ss.getId(), approvedFolderId);
    } else {
      var userSheet = getOrCreateSheet_(SHEET_NAMES.USERS);
      var userData = userSheet.getDataRange().getValues();
      var ucx = colIndexes_(userData[0]);
      for (var u = 1; u < userData.length; u++) {
        if (userData[u][ucx.user_id] === recon.user_id) {
          var folderId = userData[u][ucx.drive_folder_id];
          if (folderId) {
            var monthFolderId = ensureMonthSubfolder_(folderId, recon.month);
            moveFileToFolder_(ss.getId(), monthFolderId);
          }
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('Export file failed for recon ' + reconId + ': ' + e.message);
  }

  return ss.getUrl();
}
