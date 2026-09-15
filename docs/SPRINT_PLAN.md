# LUSU Visa Reconciliation App — Sprint Plan

**Builder:** Felix (solo, using Claude Code)
**Go-live target:** June 30, 2026
**Start date:** May 10, 2026
**Sprint length:** 2 weeks

---

## Timeline overview

| Sprint | Dates | End-of-sprint deliverable |
|--------|-------|--------------------------|
| Sprint 1 | May 10 – May 23 | Staff can log in, upload a statement, get AI-coded transactions, edit and save |
| Sprint 2 | May 24 – June 6 | Full approval chain works — Finance and ED can review, approve, and export |
| Sprint 3 | June 7 – June 20 | Auto receipt matching from Drive, duplicate detection, reporting |
| Buffer | June 21 – June 30 | Real user testing, bug fixes, go live |

> Each sprint ends with something genuinely usable — not a demo, a working tool. Staff can start using Sprint 1 output to do reconciliations manually while Sprint 2 is being built.

---

## Sprint 1 — May 10 to May 23
### Goal: Staff can log in, upload a bank statement, see AI-coded transactions, and edit them

### Story 1.1 — Database setup
- [ ] Create GAS project "LUSU Visa Reconciliation"
- [ ] Create spreadsheet "LUSU Visa Recon - Database"
- [ ] initializeDatabase() creates all 6 sheets with correct headers
- [ ] BudgetCodes populated with all codes
- [ ] Helper functions: getOrCreateSheet, generateId, logAudit
- [ ] Admin menu item wired

### Story 1.2 — Authentication and access control
- [ ] getCurrentUser() — looks up by email, returns user object
- [ ] getUserBudgetCodes() — filters by department_suffix or ALL
- [ ] addUser() — validates role, appends to Users sheet
- [ ] doGet() — access denied page or main app
- [ ] Test users seeded (one per role)

### Story 1.3 — Core UI shell
- [ ] Page.html and Stylesheet.html created
- [ ] VIEW 1 — Dashboard with stat cards and history table
- [ ] VIEW 2 — Reconciliation editor with transaction table
- [ ] VIEW 3 — Finance/ED review (read-only layout)
- [ ] View switching with no page reload
- [ ] Status badge colours correct
- [ ] Variance bar pinned to bottom

### Story 1.4 — Statement upload and AI coding
- [ ] callClaudeApi() helper
- [ ] parseStatement() — extracts transactions from PDF
- [ ] suggestBudgetCode() — AI code assignment with confidence
- [ ] uploadStatement() — creates reconciliation, saves transactions
- [ ] Client: file input, month picker, loading overlay
- [ ] VIEW 2 wired to real data

### Story 1.5 — Edit and save transactions
- [ ] getReconciliation() — access-checked, returns full data
- [ ] saveTransaction() — description or budget_code, with audit log
- [ ] getDashboard() — role-filtered reconciliation list
- [ ] Description blur → save
- [ ] Dropdown change → save
- [ ] "Saved ✓" / "Save failed" flash

### Sprint 1 review checklist
- [ ] All 6 sheets with correct headers
- [ ] All budget codes loaded
- [ ] Staff user sees only their department codes
- [ ] Finance user sees all codes
- [ ] PDF parses correctly, AI assigns codes
- [ ] Transactions over $50 flagged amber
- [ ] Edit description/budget code → saves to Sheets + audit log
- [ ] Dashboard loads and shows history
- [ ] App URL accessible in browser

---

## Sprint 2 — May 24 to June 6
### Goal: Full approval chain — receipts, submit, Finance review, ED approval, export

### Story 2.1 — Receipt attachment
- [ ] attachReceipt() — saves to Drive, creates Receipts row, updates transaction
- [ ] removeReceipt() — reverts status, does NOT delete from Drive
- [ ] Client: red Attach button / green tick / remove button
- [ ] Missing receipt count updates live

### Story 2.2 — Submit and approval workflow
- [ ] submitReconciliation() — validation, emails Finance
- [ ] financeApprove() — variance calc, emails ED
- [ ] financeSendBack() — emails staff with notes
- [ ] edApprove() — final sign-off, emails all
- [ ] edSendBack() — emails Finance and staff
- [ ] Submit blocked with specific error messages
- [ ] Role-gated action panels in VIEW 3

### Story 2.3 — Export
- [ ] exportReconciliation() — creates Google Sheet in exact LUSU format
- [ ] Correct layout: header, claimant row, column headers, data rows, totals, footer
- [ ] Variance formula and signature lines
- [ ] File saved to staff Drive folder
- [ ] Export button in VIEW 2 header

### Story 2.4 — Stabilisation
- [ ] All 10 end-to-end scenarios pass
- [ ] Claude API timeout handling (retry once, then manual fallback)
- [ ] Drive permission error handling
- [ ] LockService on all Sheets writes
- [ ] Mobile browser layout tested

### Sprint 2 review checklist
- [ ] Receipt attach/remove works
- [ ] Submit blocked when missing receipts or codes — specific error
- [ ] Full staff → finance → ED chain works
- [ ] Emails at every stage
- [ ] Variance correct when Finance enters bank total
- [ ] Export matches LUSU format with working formulas
- [ ] Claude API errors handled gracefully
- [ ] Works on mobile

---

## Sprint 3 — June 7 to June 20
### Goal: Automation and reporting

### Story 3.1 — Automatic Drive receipt scanning
- [ ] scanReceipt() — Claude extracts vendor/date/total from file
- [ ] matchReceiptToTransaction() — scoring algorithm (total 60, date 25, vendor 15)
- [ ] scanDriveFolder() — processes new files, auto-matches or flags for review
- [ ] manualMatchReceipt() — Finance or staff links receipt to transaction
- [ ] Daily trigger at 8am for all active users
- [ ] "Scan Drive folder" button in VIEW 2
- [ ] "Receipts needing review" section

### Story 3.2 — Duplicate detection
- [ ] checkDuplicates() — flags same vendor+total+date pairs and duplicate file_ids
- [ ] dismissDuplicate() — Finance/ED only
- [ ] Auto-run after upload and after Drive scan
- [ ] Orange badge on duplicate rows
- [ ] "Duplicates" filter tab
- [ ] Warning banner in Finance VIEW 3

### Story 3.3 — Reporting
- [ ] getMonthlyReport() — approved recons grouped by department
- [ ] getYTDReport() — approved recons grouped by budget code, fiscal year May–Apr
- [ ] getStaffReport() — all recons for one user
- [ ] VIEW 4 with three tabs: Monthly / YTD / Staff history
- [ ] CSV download for each report
- [ ] Visible to finance and ed only

### Sprint 3 review checklist
- [ ] Daily Drive scan trigger running
- [ ] Receipts scanned, high-confidence auto-matched
- [ ] Low-confidence in review section
- [ ] Manual match works
- [ ] Duplicates flagged (transactions and receipts)
- [ ] Finance can dismiss duplicates
- [ ] All three reports correct
- [ ] CSV downloads work

---

## Buffer — June 21 to June 30
### Goal: Real user testing, final fixes, go live

- [ ] Real PDF statement parses correctly
- [ ] AI codes sensible for real LUSU vendors
- [ ] Real staff completes full workflow
- [ ] Export matches existing LUSU format exactly
- [ ] Drive scan finds receipts in real staff folder
- [ ] iPad tested
- [ ] Error messages clear and non-technical
- [ ] All test users removed
- [ ] All real staff added
- [ ] PropertiesService values confirmed (SPREADSHEET_ID, CLAUDE_API_KEY, FINANCE_EMAILS, ED_EMAIL)
- [ ] Redeployed as new version
- [ ] URL sent to all staff
- [ ] Missing receipt reminder trigger set up (optional)
