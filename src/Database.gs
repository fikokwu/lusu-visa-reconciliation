// ============================================================
// Story 1.1 — Database setup
// ============================================================

var SHEET_NAMES = {
  USERS: 'Users',
  BUDGET_CODES: 'BudgetCodes',
  RECONCILIATIONS: 'Reconciliations',
  TRANSACTIONS: 'Transactions',
  AUDIT_LOG: 'AuditLog',
  RECEIPTS: 'Receipts'
};

var SHEET_HEADERS = {
  Users: ['user_id','name','email','role','department_suffix','drive_folder_id','active'],
  BudgetCodes: ['code','description','department_suffix','type','active'],
  Reconciliations: ['recon_id','user_id','month','status','bank_statement_total','reconciled_total','variance','submitted_at','finance_approved_at','ed_approved_at','finance_notes','ed_notes','created_at','last_edited_at','previous_balance','statement_balance'],
  Transactions: ['tx_id','recon_id','date','vendor','description','total','hst','subtotal','budget_code','code_source','ai_confidence','receipt_file_id','receipt_status','duplicate_flag','created_at','tx_type'],
  AuditLog: ['log_id','timestamp','user_id','recon_id','tx_id','action','old_value','new_value'],
  Receipts: ['receipt_id','recon_id','file_id','file_name','tx_id','match_confidence','match_status','scanned_amount','scanned_date','scanned_vendor','uploaded_at','user_id','match_reason','matched_against','file_part']
};

// ---- Helpers ----

// Memoise the spreadsheet for the duration of one execution. Apps Script globals
// are reset between requests, so this is a safe per-request cache that avoids the
// repeated PropertiesService lookup + openById on every getOrCreateSheet call.
var _ssCache = null;

function getSpreadsheet_() {
  if (_ssCache) return _ssCache;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID script property is not set. Run "Initialize Database" first.');
  _ssCache = SpreadsheetApp.openById(id);
  return _ssCache;
}

function getOrCreateSheet_(name) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function generateId(prefix) {
  // Append a short random suffix so IDs generated within the same millisecond
  // (e.g. when batch-inserting rows) remain unique.
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 100000);
}

// Build a { headerName: columnIndex } map once so row loops can do idx.foo
// lookups instead of repeatedly scanning the header array with indexOf().
function colIndexes_(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) idx[headers[i]] = i;
  return idx;
}

// Convert a single sheet row into an object keyed by header name.
function rowToObj_(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
  return obj;
}

// Build a sheet row from { headerName: value }, in the live sheet's column order
// (blank for anything not given), so writes never depend on column positions.
function rowFromFields_(headers, fields) {
  return headers.map(function(h) { return fields.hasOwnProperty(h) ? fields[h] : ''; });
}

// Append any SHEET_HEADERS columns missing from a live sheet. New columns always
// go at the end, so existing column positions never move. Returns the sheet.
function ensureColumns_(name) {
  var sheet = getOrCreateSheet_(name);
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = SHEET_HEADERS[name].filter(function(h) { return headers.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing])
      .setFontWeight('bold').setBackground('#D9D9D9');
  }
  return sheet;
}

function jsonSafe(obj) {
  // google.script.run silently returns null to the client when the payload
  // contains Date objects. Round-trip guarantees pure primitives.
  return JSON.parse(JSON.stringify(obj));
}

function logAudit_(userId, reconId, txId, action, oldValue, newValue) {
  var sheet = getOrCreateSheet_(SHEET_NAMES.AUDIT_LOG);
  sheet.appendRow([
    generateId('LOG'),
    new Date().toISOString(),
    userId || '',
    reconId || '',
    txId || '',
    action,
    oldValue !== undefined ? String(oldValue) : '',
    newValue !== undefined ? String(newValue) : ''
  ]);
}

// ---- Database initialisation ----

function initializeDatabase() {
  requireOwner_();
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');

  if (!ssId) {
    var ss = SpreadsheetApp.create('LUSU Visa Recon - Database');
    props.setProperty('SPREADSHEET_ID', ss.getId());
    Logger.log('Created spreadsheet: ' + ss.getUrl());
  }

  // Create all sheets with headers
  Object.keys(SHEET_HEADERS).forEach(function(name) {
    var sheet = getOrCreateSheet_(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SHEET_HEADERS[name]);
      sheet.getRange(1, 1, 1, SHEET_HEADERS[name].length)
        .setFontWeight('bold')
        .setBackground('#D9D9D9');
      sheet.setFrozenRows(1);
    }
  });

  populateBudgetCodes_();
  Logger.log('Database initialised.');
}

function populateBudgetCodes_() {
  var sheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);

  // Only populate if sheet is empty (header row only)
  if (sheet.getLastRow() > 1) {
    Logger.log('BudgetCodes already populated — skipping.');
    return;
  }

  var codes = getBudgetCodeData_();
  var rows = codes.map(function(c) {
    return [c[0], c[1], c[2], 'expense', !isPayrollCode_(c[0])];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  Logger.log('Loaded ' + rows.length + ' budget codes.');
}

// The 56xxx range is employment costs (salaries, wages, honoraria, commissions,
// benefits). Those go through payroll, never a Visa card, so they stay inactive
// and out of the coding dropdowns and AI suggestions.
function isPayrollCode_(code) {
  return /^56\d{3}-/.test(String(code || '').trim());
}

// One-off (run from the Apps Script editor): deactivate the payroll codes already
// in the BudgetCodes sheet. Safe to re-run. Returns the codes it switched off.
function deactivatePayrollCodes() {
  requireRole(['finance', 'ed']);
  var sheet = getOrCreateSheet_(SHEET_NAMES.BUDGET_CODES);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var ix = colIndexes_(data[0]);
  var deactivated = [];
  var activeCol = data.slice(1).map(function(row) {
    if (row[ix.active] && isPayrollCode_(row[ix.code])) {
      deactivated.push(row[ix.code] + ' ' + row[ix.description]);
      return [false];
    }
    return [row[ix.active]];
  });
  if (deactivated.length) {
    sheet.getRange(2, ix.active + 1, activeCol.length, 1).setValues(activeCol);
  }
  Logger.log('Deactivated ' + deactivated.length + ' payroll codes: ' + deactivated.join(', '));
  return deactivated;
}

function getBudgetCodeData_() {
  // [code, description, department_suffix]
  return [
    ['58000-1200','Events','1200'],
    ['58020-1200','Orientation','1200'],
    ['58025-1200','Winter Carnival','1200'],
    ['58041-1200','Swag','1200'],
    ['58046-1200','Member Outreach','1200'],
    ['58056-1200','Hockey Tickets','1200'],
    ['57100-1200','Campaigns','1200'],
    ['57145-1200','Personal Items','1200'],
    ['56320-1200','Wages - LWSP','1200'],
    ['56325-1200','Wages - Commissioner','1200'],
    ['56500-1200','Benefits - CPP/EI','1200'],
    ['56505-1200','Benefits - EHT','1200'],
    ['56510-1200','Benefits - WSIB','1200'],
    ['56520-1200','Benefits - Vacation Pay','1200'],
    ['56000-1000','Salaries - LUSU Staff','1000'],
    ['56325-1000','Wages - Other','1000'],
    ['56500-1000','Benefits - CPP/EI','1000'],
    ['56505-1000','Benefits - EHT','1000'],
    ['56510-1000','Benefits - WSIB','1000'],
    ['56520-1000','Benefits - Vacation Pay','1000'],
    ['56550-1000','Benefits - Pension','1000'],
    ['56561-1000','Benefits - Parking','1000'],
    ['56560-1000','Benefits - Meals/Misc','1000'],
    ['56562-1000','Benefits - Health & Dental','1000'],
    ['57041-1000','Nanabijou Childcare Bursary','1000'],
    ['57220-1000','Staff Appreciation','1000'],
    ['57290-1000','Scholarships & Bursaries','1000'],
    ['57500-1000','Off Campus Storage','1000'],
    ['58030-1000','Advertising','1000'],
    ['58035-1000','Argus Advertising','1000'],
    ['58210-1000','Bad Debts','1000'],
    ['58215-1000','Bank Charges','1000'],
    ['58230-1000','Collaborative Nursing','1000'],
    ['58235-1000','Merchant Fees','1000'],
    ['58250-1000','Audit','1000'],
    ['58255-1000','Legal','1000'],
    ['58307-1000','Insurance','1000'],
    ['58340-1000','E-Mail Service','1000'],
    ['58351-1000','Software Support','1000'],
    ['58352-1000','Office Supplies','1000'],
    ['58360-1000','Postage','1000'],
    ['58365-1000','Printing & Copying','1000'],
    ['58370-1000','Telephone','1000'],
    ['58385-1000','Public Relations & Community Engagement','1000'],
    ['58401-1000','Services Contracts','1000'],
    ['58405-1000','Repairs & Maintenance','1000'],
    ['58423-1000','Freight/Shipping','1000'],
    ['58440-1000','Cellular Phone','1000'],
    ['58500-1000','Membership & Dues','1000'],
    ['58530-1000','Recruitment','1000'],
    ['58535-1000','Professional Development','1000'],
    ['58550-1000','Training','1000'],
    ['58600-1000','Miscellaneous','1000'],
    ['56000-1500','Salaries - LUSU Executives','1500'],
    ['56500-1500','Benefits - CPP/EI','1500'],
    ['56505-1500','Benefits - EHT','1500'],
    ['56510-1500','Benefits - WSIB','1500'],
    ['56561-1500','Benefits - Parking','1500'],
    ['56562-1500','Benefits - Health and Dental','1500'],
    ['58101-1500','Executive Discretionary','1500'],
    ['58102-1500','President Discretionary','1500'],
    ['58103-1500','VP Finance Discretionary','1500'],
    ['58104-1500','VPA Discretionary','1500'],
    ['58440-1500','Executive Cell Phone Allowance','1500'],
    ['58525-1500','Executive Travel','1500'],
    ['56114-1700','Honorarium - Elections Committee','1700'],
    ['56115-1700','Honorarium - Secretary','1700'],
    ['56116-1700','Honorarium - Chairperson','1700'],
    ['56119-1700','Honorarium - Board of Directors','1700'],
    ['57106-1700','General Meetings & Town Halls','1700'],
    ['57125-1700','Finals Weekend Subsidy','1700'],
    ['57180-1700','Elections - Online Voting','1700'],
    ['57200-1700','Executive and Board Orientation','1700'],
    ['57210-1700','Pow Wow Funding','1700'],
    ['57225-1700','Volunteer Appreciation','1700'],
    ['57310-1700','Ombudsperson','1700'],
    ['58030-1700','Elections - Advertising','1700'],
    ['58046-1700','Meetings','1700'],
    ['58100-1700','Discretionary','1700'],
    ['58105-1700','Board of Directors Discretionary','1700'],
    ['58520-1700','Travel','1700'],
    ['58521-1700','CFS - Delegate Fees & Travel','1700'],
    ['58711-1700','Sponsorships','1700'],
    ['56060-3000','Centre Coordinator Wage','3000'],
    ['56320-3000','Wage - LWSP','3000'],
    ['56325-3000','Centre Assistant','3000'],
    ['56500-3000','Benefits - CPP/EI','3000'],
    ['56505-3000','Benefits - EHT','3000'],
    ['56510-3000','Benefits - WSIB','3000'],
    ['56520-3000','Benefits - Vacation Pay','3000'],
    ['56561-3000','Benefits - Parking','3000'],
    ['56562-3000','Benefits - Health and Dental','3000'],
    ['57100-3000','Campaigns','3000'],
    ['57150-3000','Resources','3000'],
    ['57160-3000','Speakers & Workshops','3000'],
    ['57225-3000','Volunteer Appreciation','3000'],
    ['58000-3000','Events','3000'],
    ['58021-3000','Dis-Orientation','3000'],
    ['58030-3000','Advertising','3000'],
    ['58100-3000','Discretionary','3000'],
    ['58352-3000','Office Supply','3000'],
    ['58510-3000','Conference & Travel','3000'],
    ['58711-3000','Sponsorships','3000'],
    ['56060-3100','Centre Coordinator Wage','3100'],
    ['56320-3100','Wages - LWSP','3100'],
    ['56325-3100','Wage - Centre Assistant','3100'],
    ['56500-3100','Benefits - CPP/EI','3100'],
    ['56505-3100','Benefits - EHT','3100'],
    ['56510-3100','Benefits - WSIB','3100'],
    ['56520-3100','Benefits - Vacation Pay','3100'],
    ['56561-3100','Benefits - Parking','3100'],
    ['56562-3100','Benefits - Health and Dental','3100'],
    ['57100-3100','Campaigns','3100'],
    ['57112-3100','Food Purchases','3100'],
    ['57113-3100','Free Meal Programs','3100'],
    ['57116-3100','Chopped U','3100'],
    ['57117-3100','Winter Support','3100'],
    ['57118-3100','Good Food Box','3100'],
    ['57150-3100','Resources','3100'],
    ['57160-3100','Speakers & Workshops','3100'],
    ['57225-3100','Volunteer Appreciation','3100'],
    ['58000-3100','Events','3100'],
    ['58021-3100','Dis-Orientation','3100'],
    ['58030-3100','Advertising','3100'],
    ['58100-3100','Discretionary','3100'],
    ['58352-3100','Office Supply','3100'],
    ['58485-3100','Greenhouse/Garden Project','3100'],
    ['58510-3100','Conference & Travel','3100'],
    ['58711-3100','Sponsorships','3100'],
    ['56060-3200','Centre Coordinator Wage','3200'],
    ['56115-3200','Honoraria','3200'],
    ['56320-3200','Wages - LWSP','3200'],
    ['56325-3200','Wages - Assistant','3200'],
    ['56500-3200','Benefits - CPP/EI','3200'],
    ['56505-3200','Benefits - EHT','3200'],
    ['56510-3200','Benefits - WSIB','3200'],
    ['56520-3200','Benefits - Vacation Pay','3200'],
    ['56561-3200','Benefits - Parking','3200'],
    ['56562-3200','Benefits - Health and Dental','3200'],
    ['57100-3200','Campaigns','3200'],
    ['57120-3200','DOA - Take Back the Night','3200'],
    ['57150-3200','Resources','3200'],
    ['57160-3200','Speakers & Workshops','3200'],
    ['57165-3200','Feminism at Lakehead','3200'],
    ['57175-3200','Groups and Workshops','3200'],
    ['57225-3200','Volunteer Appreciation','3200'],
    ['58000-3200','Events','3200'],
    ['58021-3200','Dis-Orientation','3200'],
    ['58030-3200','Advertising','3200'],
    ['58100-3200','Discretionary','3200'],
    ['58352-3200','Office Supply','3200'],
    ['58510-3200','Conference & Travel','3200'],
    ['58711-3200','Sponsorships','3200'],
    ['56060-3300','Centre Coordinator Wage','3300'],
    ['56115-3300','Honoraria','3300'],
    ['56320-3300','Wages - LWSP','3300'],
    ['56325-3300','Wages - Assistant','3300'],
    ['56500-3300','Benefits - CPP/EI','3300'],
    ['56505-3300','Benefits - EHT','3300'],
    ['56510-3300','Benefits - WSIB','3300'],
    ['56520-3300','Benefits - Vacation Pay','3300'],
    ['56561-3300','Benefits - Parking','3300'],
    ['56562-3300','Benefits - Health and Dental','3300'],
    ['57100-3300','Campaigns','3300'],
    ['57150-3300','Resources','3300'],
    ['57160-3300','Speakers & Workshops','3300'],
    ['57225-3300','Volunteer Appreciation','3300'],
    ['58000-3300','Events','3300'],
    ['58021-3300','Dis-Orientation','3300'],
    ['58030-3300','Advertising','3300'],
    ['58100-3300','Discretionary','3300'],
    ['58352-3300','Office Supply','3300'],
    ['58510-3300','Conference & Travel','3300'],
    ['58711-3300','Sponsorships','3300'],
    ['56060-3500','Centre Coordinator Wage','3500'],
    ['56115-3500','Honoraria','3500'],
    ['56320-3500','Wages - LWSP','3500'],
    ['56325-3500','Wage - Assistant','3500'],
    ['56500-3500','Benefits - CPP/EI','3500'],
    ['56505-3500','Benefits - EHT','3500'],
    ['56510-3500','Benefits - WSIB','3500'],
    ['56520-3500','Benefits - Vacation Pay','3500'],
    ['56561-3500','Benefits - Parking','3500'],
    ['56562-3500','Benefits - Health and Dental','3500'],
    ['57100-3500','Campaigns','3500'],
    ['57120-3500','Take Back the Night','3500'],
    ['57150-3500','Resources','3500'],
    ['57155-3500','Pride in the North','3500'],
    ['57156-3500','Gender-Affirming Gear','3500'],
    ['57160-3500','Speakers & Workshops','3500'],
    ['57165-3500','Feminism at Lakehead','3500'],
    ['57225-3500','Volunteer Appreciation','3500'],
    ['58000-3500','Events','3500'],
    ['58021-3500','Dis-Orientation','3500'],
    ['58030-3500','Advertising','3500'],
    ['58100-3500','Discretionary','3500'],
    ['58352-3500','Office Supply','3500'],
    ['58510-3500','Conference & Travel','3500'],
    ['58711-3500','Sponsorships','3500'],
    ['56060-1900','Centre Coordinator Wage','1900'],
    ['56115-1900','Honoraria','1900'],
    ['56320-1900','Wage - Assistant','1900'],
    ['56325-1900','Wage - LWSP','1900'],
    ['56500-1900','Benefits - CPP/EI','1900'],
    ['56505-1900','Benefits - EHT','1900'],
    ['56510-1900','Benefits - WSIB','1900'],
    ['56520-1900','Benefits - Vacation Pay','1900'],
    ['56561-1900','Benefits - Parking','1900'],
    ['56562-1900','Benefits - Health and Dental','1900'],
    ['57100-1900','Campaigns','1900'],
    ['57150-1900','Resources','1900'],
    ['57160-1900','Speakers & Workshops','1900'],
    ['57225-1900','Volunteer Appreciation','1900'],
    ['58000-1900','Events','1900'],
    ['58021-1900','Dis-Orientation','1900'],
    ['58030-1900','Advertising','1900'],
    ['58100-1900','Discretionary','1900'],
    ['58352-1900','Office Supply','1900'],
    ['58510-1900','Conference & Travel','1900'],
    ['58711-1900','Sponsorships','1900'],
    ['58712-1900','LUSU Sustainability Initiatives','1900'],
    ['55000-8000','COS - Bottled Beer','8000'],
    ['55001-8000','COS - Draft','8000'],
    ['55002-8000','COS - Liquor','8000'],
    ['55004-8000','COS - Non Alcoholic','8000'],
    ['55005-8000','COS - Pop','8000'],
    ['55010-8000','COS - Hot Drinks','8000'],
    ['55020-8000','COS - Food','8000'],
    ['55030-8000','Cost of Waste','8000'],
    ['55100-8000','Kitchen Supplies','8000'],
    ['55110-8000','Bar Supplies','8000'],
    ['56060-8000','Salaries/Wages','8000'],
    ['56300-8000','Wages - Kitchen','8000'],
    ['56315-8000','Wages - Bar','8000'],
    ['56328-8000','Wages - Door','8000'],
    ['56330-8000','Wages - Runner','8000'],
    ['56500-8000','Benefits - CPP/EI','8000'],
    ['56505-8000','Benefits - EHT','8000'],
    ['56510-8000','Benefits - WSIB','8000'],
    ['56520-8000','Benefits - Vacation Pay','8000'],
    ['56550-8000','Benefits - Pension','8000'],
    ['56561-8000','Benefits - Parking','8000'],
    ['56562-8000','Benefits - Health & Dental','8000'],
    ['57220-8000','Staff Appreciation','8000'],
    ['58030-8000','Advertising','8000'],
    ['58035-8000','Advertising - Argus','8000'],
    ['58037-8000','Advertising - Radio','8000'],
    ['58045-8000','Promotions','8000'],
    ['58063-8000','Artist Fees','8000'],
    ['58065-8000','Riders','8000'],
    ['58070-8000','DJ Contract','8000'],
    ['58085-8000','Sound/Technical Support','8000'],
    ['58215-8000','Bank Charges & Interest','8000'],
    ['58235-8000','Merchant Fees','8000'],
    ['58307-8000','Insurance','8000'],
    ['58315-8000','Memberships/Licences','8000'],
    ['58325-8000','SOCAN','8000'],
    ['58351-8000','Software Support','8000'],
    ['58352-8000','Office Supplies','8000'],
    ['58355-8000','Equipment/Furniture','8000'],
    ['58356-8000','Computer Supplies','8000'],
    ['58365-8000','Printing and Photocopier','8000'],
    ['58370-8000','Telephone','8000'],
    ['58375-8000','Outpost Utilities','8000'],
    ['58380-8000','Shaw - Internet','8000'],
    ['58401-8000','Service Maintenance Contracts','8000'],
    ['58402-8000','Outpost Cleaning','8000'],
    ['58405-8000','Repairs & Maintenance','8000'],
    ['58420-8000','Designated Driver Program','8000'],
    ['58440-8000','Cellphone Allowance','8000'],
    ['58510-8000','Conference & Travel','8000'],
    ['58540-8000','Staff Uniforms','8000'],
    ['58550-8000','Staff Training','8000'],
    ['58620-8000','Fuel Charge','8000'],
    ['58700-8000','Draft Line Clean/Liquor Gun','8000'],
    ['58710-8000','Thunderwolves Sponsorship','8000'],
    ['58715-8000','Glassware','8000'],
    ['58720-8000','Entertainment Fund','8000'],
    ['58725-8000','Technology Upgrade','8000'],
    ['55000-6000','COS - Beer','6000'],
    ['55003-6000','COS - Wine','6000'],
    ['55005-6000','COS - Pop and Juice','6000'],
    ['55009-6000','COS - Tea/Syrups','6000'],
    ['55010-6000','COS - Hot Drinks','6000'],
    ['55011-6000','COS - Coffee','6000'],
    ['55020-6000','COS - Food','6000'],
    ['55100-6000','Kitchen Supplies','6000'],
    ['55105-6000','Supplies','6000'],
    ['56310-6000','Wages - Barista','6000'],
    ['56500-6000','Benefits - CPP/EI','6000'],
    ['56505-6000','Benefits - EHT','6000'],
    ['56510-6000','Benefits - WSIB','6000'],
    ['56520-6000','Benefits - Vacation Pay','6000'],
    ['56550-6000','Benefits - Pension','6000'],
    ['56561-6000','Benefits - Parking','6000'],
    ['56562-6000','Benefits - Health & Dental','6000'],
    ['57220-6000','Staff Appreciation','6000'],
    ['58030-6000','Advertising','6000'],
    ['58215-6000','Bank Charges and Interest','6000'],
    ['58235-6000','Merchant Fees','6000'],
    ['58307-6000','Insurance','6000'],
    ['58351-6000','Software Support','6000'],
    ['58352-6000','Office Supplies','6000'],
    ['58370-6000','Telephone','6000'],
    ['58405-6000','Repairs & Maintenance','6000'],
    ['58440-6000','Cellular Phone','6000'],
    ['58510-6000','Conference and Travel','6000'],
    ['58600-6000','Miscellaneous','6000'],
    ['58655-6000','Study Rent','6000'],
    ['58720-6000','Entertainment Fund','6000'],
    ['55005-6500','COS - Pop and Juice','6500'],
    ['55009-6500','COS - Tea/Syrups','6500'],
    ['55011-6500','COS - Coffee','6500'],
    ['55020-6500','COS - Food','6500'],
    ['55100-6500','Kitchen Supplies','6500'],
    ['55105-6500','Supplies','6500'],
    ['56310-6500','Wages - Barista','6500'],
    ['56500-6500','Benefits - CPP/EI','6500'],
    ['56505-6500','Benefits - EHT','6500'],
    ['56510-6500','Benefits - WSIB','6500'],
    ['56520-6500','Benefits - Vacation Pay','6500'],
    ['56550-6500','Benefits - Pension','6500'],
    ['56561-6500','Benefits - Parking','6500'],
    ['56562-6500','Benefits - Health & Dental','6500'],
    ['58215-6500','Bank Charges and Interest','6500'],
    ['58235-6500','Merchant Fees','6500'],
    ['58307-6500','Insurance','6500'],
    ['58351-6500','Software Support','6500'],
    ['58352-6500','Office Supplies','6500'],
    ['58405-6500','Repairs & Maintenance','6500'],
    ['58440-6500','Cellular Phone','6500'],
    ['56000-7000','Salaries - LUSU Staff','7000'],
    ['56107-7000','Wage - Programming and Services Coordinator','7000'],
    ['56320-7000','Wage - Office Assistant','7000'],
    ['56500-7000','Benefits - CPP/EI','7000'],
    ['56505-7000','Benefits - EHT','7000'],
    ['56510-7000','Benefits - WSIB','7000'],
    ['56520-7000','Benefits - Vacation Pay','7000'],
    ['56550-7000','Benefits - Pension','7000'],
    ['56561-7000','Benefits - Parking','7000'],
    ['56562-7000','Benefits - Health & Dental','7000'],
    ['57100-7000','Campaigns','7000'],
    ['57115-7000','Food Resource Collective','7000'],
    ['57145-7000','Personal Items','7000'],
    ['57195-7000','Club Discretionary','7000'],
    ['57196-7000','Clubs Funding Transfer','7000'],
    ['57500-7000','Off Campus Storage','7000'],
    ['58000-7000','Events','7000'],
    ['58020-7000','Orientation','7000'],
    ['58046-7000','Membership Outreach','7000'],
    ['58066-7000','Programming','7000'],
    ['58100-7000','Discretionary','7000'],
    ['58215-7000','Bank Charges','7000'],
    ['58352-7000','Office Supplies','7000'],
    ['58360-7000','Postage','7000'],
    ['58370-7000','Telephone','7000'],
    ['58385-7000','Public Relations & Community Engagement','7000'],
    ['58405-7000','Repairs & Maintenance','7000'],
    ['58520-7000','Travel','7000'],
    ['58521-7000','Delegate Fees & Travel - CFS','7000'],
    ['58535-7000','Professional Development','7000'],
    ['56000-7100','Salaries - LUSU Executives','7100'],
    ['56500-7100','Benefits - CPP/EI','7100'],
    ['56505-7100','Benefits - EHT','7100'],
    ['56510-7100','Benefits - WSIB','7100'],
    ['56561-7100','Benefits - Parking','7100'],
    ['56562-7100','Benefits - Health and Dental','7100'],
    ['58106-7100','VP Orillia Discretionary','7100'],
    ['58440-7100','Executive Cell Phone Allowance','7100'],
    ['58525-7100','Executive Travel','7100'],
    ['56117-2000','Honoraria - Members','2000'],
    ['56200-2000','Commissions','2000'],
    ['56310-2000','Wages','2000'],
    ['56500-2000','Benefits - CPP/EI','2000'],
    ['56505-2000','Benefits - EHT','2000'],
    ['56510-2000','Benefits - WSIB','2000'],
    ['56520-2000','Benefits - Vacation Pay','2000'],
    ['58030-2000','Advertising - Social Media','2000'],
    ['58330-2000','Book Binding','2000'],
    ['58340-2000','E-Mail Service','2000'],
    ['58350-2000','Website','2000'],
    ['58352-2000','Office Supplies','2000'],
    ['58365-2000','Printing','2000'],
    ['58370-2000','Telephone','2000'],
    ['58405-2000','Repairs & Maintenance','2000'],
    ['58500-2000','Dues & Memberships','2000'],
    ['58510-2000','Conference & Travel','2000'],
    ['58600-2000','Miscellaneous','2000'],
    ['58620-2000','Fuel Charges','2000']
  ];
}

// ---- Schema migration ----

function ensureLastEditedAtColumn() {
  requireOwner_();
  var sheet = getOrCreateSheet_(SHEET_NAMES.RECONCILIATIONS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  if (headers.indexOf('last_edited_at') !== -1) return false;
  var newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue('last_edited_at').setFontWeight('bold').setBackground('#D9D9D9');
  var createdAtIdx = headers.indexOf('created_at');
  for (var i = 1; i < data.length; i++) {
    var val = createdAtIdx !== -1 ? data[i][createdAtIdx] : new Date().toISOString();
    sheet.getRange(i + 1, newCol).setValue(val);
  }
  Logger.log('Added last_edited_at column to Reconciliations and backfilled ' + (data.length - 1) + ' rows.');
  return true;
}

// ---- Historical data import ----

function importHistoricalSheets() {
  requireOwner_();
  var folderId = PropertiesService.getScriptProperties().getProperty('HISTORICAL_SHEETS_FOLDER_ID');
  if (!folderId) throw new Error('Set HISTORICAL_SHEETS_FOLDER_ID in Script Properties first.');

  // Remove any previously imported historical rows
  var txSheet = getOrCreateSheet_(SHEET_NAMES.TRANSACTIONS);
  var txData = txSheet.getDataRange().getValues();
  var txHeaders = txData[0];
  var reconIdIdx = txHeaders.indexOf('recon_id');
  for (var d = txData.length - 1; d >= 1; d--) {
    if (txData[d][reconIdIdx] === 'HISTORICAL') txSheet.deleteRow(d + 1);
  }

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var imported = 0;
  var skipped = 0;

  while (files.hasNext()) {
    var file = files.next();
    var ss = SpreadsheetApp.openById(file.getId());

    ss.getSheets().forEach(function(sheet) {
      var data = sheet.getDataRange().getValues();

      // Find the column-header row (contains "DATE" in col A)
      var dataStart = -1;
      for (var r = 0; r < Math.min(data.length, 10); r++) {
        if (String(data[r][0]).toUpperCase().indexOf('DATE') !== -1) {
          dataStart = r + 1;
          break;
        }
      }
      if (dataStart === -1) return;

      for (var i = dataStart; i < data.length; i++) {
        var row = data[i];
        var colA = String(row[0] || '').trim().toUpperCase();
        if (colA === 'TOTALS' || colA === 'NOTES:') break;

        var vendor     = String(row[1] || '').trim();
        var desc       = String(row[2] || '').trim();
        var total      = parseFloat(row[3]) || 0;
        var hst        = parseFloat(row[4]) || 0;
        var subtotal   = parseFloat(row[5]) || 0;
        var budgetCode = String(row[6] || '').trim();

        if (!vendor || !budgetCode || total === 0) { skipped++; continue; }

        txSheet.appendRow([
          generateId('HIST'), 'HISTORICAL',
          String(row[0] || '').trim(),
          vendor.toUpperCase(), desc,
          total, hst, subtotal,
          budgetCode, 'historical', 100,
          '', 'historical', false,
          new Date().toISOString()
        ]);
        imported++;
      }
    });
  }

  Logger.log('Historical import done: ' + imported + ' rows imported, ' + skipped + ' skipped.');
  return { imported: imported, skipped: skipped };
}

// ---- Admin menu ----

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LUSU Admin')
    .addItem('Initialize Database', 'initializeDatabase')
    .addItem('Import Historical Sheets', 'importHistoricalSheets')
    .addItem('Migrate: add last_edited_at', 'ensureLastEditedAtColumn')
    .addToUi();
}
