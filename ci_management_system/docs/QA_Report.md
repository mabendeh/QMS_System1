# Quality Assurance Report
**C.H. Guenther & Son, Duncanville: CI Management System**

Build v1.0 2026-11-19. **87 of 87 automated tests passed**; 0 failed.

## How testing was done
`test_workbook.py` runs four kinds of tests:

1. **Full recalculation.** Both workbooks are recalculated in headless LibreOffice. Any formula error value (REF, VALUE, NAME, DIV/0, N/A) is a failure.
2. **Structural checks.** Sheets, tables, named ranges, dropdown targets, protection and locking, freeze panes, conditional formatting, print setup, hyperlinks and fonts. The checks also confirm there are no macros or external links.
3. **Production-content checks.** Only the confirmed assignments are present, with no invented dates, baselines, savings or records.
4. **Independent recomputation and scenarios.** Every key KPI is recomputed in Python from the demo dataset and compared with the Excel results. Scenario tests then edit a temporary copy, recalculate it and check how the dashboard responds:
   * Changing the filters
   * Completing an action
   * Finance approving a savings line
   * Changing the As-Of date
   * Adding a supervisor and a project

The same results are written into the workbook on tab **14 QA Report**.

## Test results

| # | Area | Test | Result |
|---|---|---|---|
| 1 | Formulas (production) | LibreOffice full recalculation of 41244 formulas: zero errors (no REF, VALUE, NAME, DIV/0 or N/A error values) | PASS |
| 2 | Formulas (demo) | LibreOffice full recalculation of 41244 formulas: zero errors (no REF, VALUE, NAME, DIV/0 or N/A error values) | PASS |
| 3 | Structure (production) | All 17 worksheets present in the specified order (01-14 plus 03A/06A/07A support sheets) | PASS |
| 4 | Structure (production) | Excel tables exist for every register (filters/structured layout) | PASS |
| 5 | Structure (production) | All 135 named ranges used in formulas are defined | PASS |
| 6 | Dropdowns (production) | 144 data-validation rules; every list dropdown points to a defined dynamic list | PASS |
| 7 | Dropdowns (production) | Required editable lists exist (supervisors, departments, shifts, equipment, stages, statuses, priorities, Finance statuses, categories, IDs) | PASS |
| 8 | Dropdowns (production) | Dropdown lists are dynamic (OFFSET/COUNTA) so new supervisors/values need no formula edits | PASS |
| 9 | Protection (production) | Sheet protection enabled on every sheet (formula cells locked, filters allowed) | PASS |
| 10 | Protection (production) | Formula cells locked (2981 checked) and yellow input cells unlocked (5322 checked) | PASS |
| 11 | Usability (production) | Frozen header rows on every register | PASS |
| 12 | Formatting (production) | Conditional formatting present for overdue actions, blocked projects, missing evidence/links and leadership escalations | PASS |
| 13 | Printing (production) | Dashboard WOR view prints on one landscape page (fit 1x1) | PASS |
| 14 | Printing (production) | Project charter prints on one portrait page (fit 1x1) | PASS |
| 15 | Printing (production) | WOR report: landscape, 1 page wide, title rows repeat, page break every 5 projects | PASS |
| 16 | Hyperlinks (production) | Dashboard navigation: 14 internal links all target existing sheets | PASS |
| 17 | Hyperlinks (production) | Evidence Register 'Open Link' column uses HYPERLINK() on the entered SharePoint/Teams URL | PASS |
| 18 | Dashboard (production) | Trend charts present (4: completion, overdue, savings, Pareto) | PASS |
| 19 | Dashboard (production) | Every KPI card and project-status cell is a formula (no typed metrics) | PASS |
| 20 | Branding (production) | No logo image embedded (official logo not supplied - placeholder text only) | PASS |
| 21 | Formatting (production) | Consistent typography (Arial) on dashboard, register and WOR report | PASS |
| 22 | Compatibility (production) | No macros (VBA) and no external workbook links | PASS |
| 23 | Structure (demo) | All 17 worksheets present in the specified order (01-14 plus 03A/06A/07A support sheets) | PASS |
| 24 | Structure (demo) | Excel tables exist for every register (filters/structured layout) | PASS |
| 25 | Structure (demo) | All 135 named ranges used in formulas are defined | PASS |
| 26 | Dropdowns (demo) | 144 data-validation rules; every list dropdown points to a defined dynamic list | PASS |
| 27 | Dropdowns (demo) | Required editable lists exist (supervisors, departments, shifts, equipment, stages, statuses, priorities, Finance statuses, categories, IDs) | PASS |
| 28 | Dropdowns (demo) | Dropdown lists are dynamic (OFFSET/COUNTA) so new supervisors/values need no formula edits | PASS |
| 29 | Protection (demo) | Sheet protection enabled on every sheet (formula cells locked, filters allowed) | PASS |
| 30 | Protection (demo) | Formula cells locked (2981 checked) and yellow input cells unlocked (5322 checked) | PASS |
| 31 | Usability (demo) | Frozen header rows on every register | PASS |
| 32 | Formatting (demo) | Conditional formatting present for overdue actions, blocked projects, missing evidence/links and leadership escalations | PASS |
| 33 | Printing (demo) | Dashboard WOR view prints on one landscape page (fit 1x1) | PASS |
| 34 | Printing (demo) | Project charter prints on one portrait page (fit 1x1) | PASS |
| 35 | Printing (demo) | WOR report: landscape, 1 page wide, title rows repeat, page break every 5 projects | PASS |
| 36 | Hyperlinks (demo) | Dashboard navigation: 14 internal links all target existing sheets | PASS |
| 37 | Hyperlinks (demo) | Evidence Register 'Open Link' column uses HYPERLINK() on the entered SharePoint/Teams URL | PASS |
| 38 | Dashboard (demo) | Trend charts present (4: completion, overdue, savings, Pareto) | PASS |
| 39 | Dashboard (demo) | Every KPI card and project-status cell is a formula (no typed metrics) | PASS |
| 40 | Branding (demo) | No logo image embedded (official logo not supplied - placeholder text only) | PASS |
| 41 | Formatting (demo) | Consistent typography (Arial) on dashboard, register and WOR report | PASS |
| 42 | Compatibility (demo) | No macros (VBA) and no external workbook links | PASS |
| 43 | Production data | Exactly 5 register rows: 3 confirmed projects + Maintenance and Warehouse placeholders | PASS |
| 44 | Production data | Confirmed owners preloaded exactly as assigned (leader + co-owner) | PASS |
| 45 | Production data | Placeholders for Maintenance / Warehouse have NO owner and status Pending Assignment | PASS |
| 46 | Production data | No invented dates, baselines, targets, KPIs, sponsors or shifts (all blank) | PASS |
| 47 | Production data | Department left blank for the 3 projects where it was not confirmed | PASS |
| 48 | Production data | No fabricated actions (table empty) | PASS |
| 49 | Production data | No fabricated savings lines (table empty) | PASS |
| 50 | Production data | No fabricated weekly updates (table empty) | PASS |
| 51 | Production data | No fabricated risks (table empty) | PASS |
| 52 | Production data | No fabricated evidence records (table empty) | PASS |
| 53 | Production dashboard | KPI cards: 3 active, 0 complete, action completion N/A (no actions yet), 0 overdue, $0 forecast, $0 verified, 5 requiring attention, 0 escalations | PASS |
| 54 | Production dashboard | All 12 live reconciliation checks = PASS | PASS |
| 55 | Production dashboard | RAG: 3 AMBER (target dates not yet set) and 2 PENDING (awaiting assignment) | PASS |
| 56 | Production data | Settings: As-Of date blank (uses TODAY()), ID prefix DUN-CI-, 52-week annualization labelled as assumption | PASS |
| 57 | Calculation (demo) | Register per-project KPIs (actions, completion %, milestones, overdue, risks, RAG, forecast, verified) match independent Python recomputation for all 6 projects | PASS |
| 58 | Calculation (demo) | Savings lines: material, changeover, annualization, gross/net, attribution and verified values match the documented formulas | PASS |
| 59 | Finance integrity (demo) | Unapproved line with a realized amount entered (Under Review, $9,999) contributes $0 to verified | PASS |
| 60 | Finance integrity (demo) | Changeover line NOT demonstrated values $0 cash but reports capacity hours (5.8 h/wk -> 303.3 h/yr) | PASS |
| 61 | Finance integrity (demo) | Shared benefit pool attributed 60% + 50% is flagged OVER-ATTRIBUTED (double counting prevented) | PASS |
| 62 | Finance integrity (demo) | Rejected line excluded from forecast KPI and flagged for missing data source | PASS |
| 63 | Finance integrity (demo) | Volume shift above 20% tolerance is flagged / within tolerance reads 'Comparable' | PASS |
| 64 | Dashboard reconciliation (demo) | All KPI cards equal totals recomputed independently from the source records | PASS |
| 65 | Dashboard reconciliation (demo) | Action completion card = credited / non-cancelled actions | PASS |
| 66 | Dashboard reconciliation (demo) | Live integrity panel passes every check except the intentionally seeded over-attribution | PASS |
| 67 | Dashboard reconciliation (demo) | Department and shift breakdowns sum to total projects, actions and savings | PASS |
| 68 | Trends (demo) | 12-week planned / completed / overdue trend matches date-based recomputation | PASS |
| 69 | Trends (demo) | Cumulative forecast run-rate for the current month matches recomputation | PASS |
| 70 | Actions (demo) | 'Complete' without completed date is NOT credited and is flagged 'Enter completed date' | PASS |
| 71 | Actions (demo) | Days remaining = due - As-Of date; overdue flagged; on-time/late closure classified | PASS |
| 72 | Actions (demo) | Missing evidence flagged on credited actions without evidence; evidence-linked action shows Verified | PASS |
| 73 | Actions (demo) | Action with no due date flagged NO DUE DATE | PASS |
| 74 | Evidence (demo) | Link check: https link OK, malformed link flagged; Action ID on another project flagged | PASS |
| 75 | Risks (demo) | Escalated High-impact blocker flagged LEADERSHIP ACTION; closed item not flagged; days open computed | PASS |
| 76 | Development (demo) | Competency average excludes unscored areas (3.375 with 8 of 9 scored); level band and focus area derived | PASS |
| 77 | Development (demo) | Participation summary: latest review average and closure % per supervisor (no savings, no ranking) | PASS |
| 78 | Charter (demo) | Charter pulls register, charter-record and milestone data for the selected project | PASS |
| 79 | WOR report (demo) | Default filters show 5 open projects (completed excluded) with goal, KPI actual and latest update | PASS |
| 80 | WOR report (demo) | Project with no weekly update shows 'No update recorded' and STALE warning | PASS |
| 81 | Scenario (demo copy) | WOR filter Department = Maintenance + include completed -> only DEMO-004 shown | PASS |
| 82 | Scenario (demo copy) | WOR owner filter matches leader OR co-owner (Demo Supervisor C -> DEMO-002 co-owner, DEMO-003 leader) | PASS |
| 83 | Scenario (demo copy) | WOR shift filter = B -> DEMO-002 and DEMO-005 | PASS |
| 84 | Scenario (demo copy) | Entering a completed date on an overdue action: credited 7->8, overdue 6->5, reconciliation still PASS | PASS |
| 85 | Scenario (demo copy) | Finance status change to Approved moves the realized amount into verified savings (3,900 -> 13,899) | PASS |
| 86 | Scenario (demo copy) | Changing the As-Of date to 09/01/2026 re-evaluates overdue actions (6 -> 0) | PASS |
| 87 | Scenario (demo copy) | Adding a supervisor (Master Data) and a project (Register) works with no formula changes: ID auto-generated, owner valid, dashboard and development summary update | PASS |

## Known limitations

- Recalculation was verified with LibreOffice; the workbook targets desktop Excel 2019 / Microsoft 365 (MAXIFS/MINIFS). Open once in Excel and confirm the dashboard before first WOR use - Excel recalculates fully on open.
- Sheet protection has no password and is not a security control; Excel has no row-level ownership security. Anyone with edit access can change any input (use SharePoint version history for audit).
- Excel cannot upload, store or secure documents, and there is no live SharePoint synchronization. The Evidence Register stores hyperlinks only; viewers need permission to the target folder. The approved folder URL must be entered by the CI Manager.
- Trends are reconstructed from recorded dates (due/completed/Finance review). If dates are edited retroactively the history changes; for a frozen record, save a PDF of the dashboard each WOR week.
- IDs derive from row position: do not sort or delete rows in the data sheets (use Status = Cancelled). Filtering is safe.
- Capacity: 150 projects, 1,000 actions, 300 savings lines, 2,000 weekly updates, 300 risks, 1,000 evidence links, 300 reviews; dashboard lists the first 15 projects and the WOR report 10 per filter.
- Annualization uses 52 operating weeks by default - an assumption pending Finance confirmation (Master Data C9).
- Data-sheet print areas are dynamic named ranges (grow with the records entered); some non-Microsoft viewers ignore them.
- No official C.H. Guenther logo was supplied, so none is embedded; a labelled placeholder marks where to insert it.
- Primary owner for each initial project is the first-named person in the assignment; confirm or swap leader/co-owner in the register.

## Manual checks recommended in Microsoft Excel before first use
The build environment doesn't have desktop Excel, so these checks still need a person:

1. **Open the production file.**
   * Confirm there is no repair prompt.
   * Confirm the dashboard shows 3 active projects and 5 requiring attention.
   * Confirm all 12 integrity checks read PASS.
2. **Dropdowns and links.** Try one dropdown on each register, and open the evidence link on a test row.
3. **Print preview.** Check that:
   * The dashboard fits on one page.
   * The charter fits on one page.
   * The WOR report prints in landscape.
   * One data sheet prints as expected. Data-sheet print areas grow as records are added.
4. **Set up storage.** Save the file in the approved SharePoint folder, and enter the evidence folder URL in Master Data C13.
