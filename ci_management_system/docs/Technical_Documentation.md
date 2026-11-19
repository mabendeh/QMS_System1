# Technical Documentation
**C.H. Guenther & Son, Duncanville: Supervisor Productivity & Continuous Improvement Management System, v1.0**

## 1. Files

| File | Purpose |
|---|---|
| `output/CHG_Duncanville_CI_Management_System.xlsx` | **Production workbook.** Contains only the confirmed assignments; every unknown field is blank. |
| `output/CHG_Duncanville_CI_System_DEMO_TRAINING.xlsx` | Training and test copy with **fictitious** data. Red "DEMO" banner on every sheet; ID prefix `DEMO-`. Never use it for WOR or Finance reporting. |
| `build_workbook.py` | Generator script (Python 3 + openpyxl). It is the single source of truth for the structure. |
| `test_workbook.py` | Automated QA. Recalculates in LibreOffice, recomputes the KPIs in Python and runs scenario tests. Writes `qa_results.json`. |
| `docs/` | Quick-start guide, this document and the QA report. |

To rebuild:

```bash
python3 build_workbook.py
python3 test_workbook.py
python3 build_workbook.py --qa qa_results.json
```

## 2. Architecture

| Sheet | Type | Key contents |
|---|---|---|
| 01 Executive Dashboard | Output | 8 KPI cards and 8 secondary metrics; project status for the first 15 projects; breakdowns by department and by shift; Finance pipeline; 12 live reconciliation checks; 12-week and 12-month trend tables with 3 charts. The print area B1:P57 prints on one page. |
| 02 Project Register | Table `tblProjects`, prefix `reg_` | 25 input columns and 24 AUTO columns: KPI actual, attainment, action/milestone completion, overdue, risks, savings, RAG, attention reasons, owner check, WOR helpers. |
| 03 Project Charter | Output / print | Project ID selector. Pulls data from `reg_`, `chr_` and the milestone actions. Prints on one portrait page. |
| 03A Charter Records | Table `chr_` | Charter narrative fields and approval dates. Calculates completeness %. |
| 04 Action Tracker | Table `act_` | Actions with one accountable owner. Calculates credit, days remaining, schedule status, evidence check, data check and milestone keys. |
| 05 Financial Savings | Table `fin_` | Benefit lines: method inputs, Finance approval, annualization, attribution, double-count control, missing-data flags. |
| 06 Weekly WOR Report | Output / print | Filters for department, shift, owner and completed projects. Shows 10 auto-filled project cards. |
| 06A WOR Weekly Updates | Table `wor_` | One row per project per week, keyed `ProjectID|WeekEnding`. |
| 07 Supervisor Development | Table `dev_` | Competency scores 1–4 (nine areas), coaching notes, development actions, next review. |
| 07A Rubric and Summary | Reference / output | Rubric definitions and a participation summary per person (not ranked; no dollar figures). |
| 08 CI Tools | Templates | 5 Why, PDCA, A3, 5S audit, SMED, waste Pareto (auto-sort plus chart), before/after evidence, standard work verification. |
| 09 Risks and Escalations | Table `risk_` | Blockers, owners, escalation level and dates. Flags open items and items needing leadership action. |
| 10 Evidence Register | Table `evd_` | Hyperlinks to SharePoint/Teams evidence, with verification status, link-format check and action check. |
| 11 Master Data | Settings / lists | System settings, people table, 22 dropdown lists, report filter lists, owner-assignment and project-ownership lookups. |
| 12 User Guide | Documentation | Supervisor quick-start, CI Manager and Finance workflows, limitations. |
| 13 Data Dictionary | Documentation | KPI definitions, data states, maintenance. |
| 14 QA Report | Documentation | Test results for this build and known limitations. |

### Named ranges

* **Table columns.** Every table column has a workbook-level name `<prefix>_<Key>` covering rows 6 up to capacity, e.g. `act_Due` or `fin_Verified`.
  Formulas use these names, so they read like sentences and are easy to audit in **Formulas > Name Manager**.
* **Settings.** `AsOfDate`, `IDPrefix`, `OpWeeks`, `RedOverdueThr`, `DueSoonDays`, `VolTol`, `EvidenceRoot`.
* **Dropdowns.** `lst_*` names are dynamic, e.g. `OFFSET(first cell,0,0,COUNTA(list),1)`. Adding a value directly below a list extends it with no formula edits.
  `lst_ProjectID` and `lst_ActionID` count generated IDs with `COUNTIF(range,"?*")`.
* **Report filters.** `WOR_FDept`, `WOR_FShift`, `WOR_FOwner`, `WOR_FIncl`. Charter selectors: `CH_Sel`, `CH_Idx`, `CH_CIdx`.

### IDs

IDs are built from the row position: `IDPrefix & TEXT(ROW()-5,"000")`, `"A-"&TEXT(ROW()-5,"0000")`, and so on. A new ID appears as soon as the key field is filled.
Because IDs depend on position, **rows must never be sorted or deleted**. Set the status to Cancelled instead.

### Functions used

`COUNTIFS`, `SUMIFS`, `SUMPRODUCT`, `INDEX/MATCH`, `IFERROR`, `OFFSET`, `EOMONTH`, `HYPERLINK`, `TEXT`, and `MAXIFS/MINIFS`.
`MAXIFS/MINIFS` need Excel 2019 or Microsoft 365.
The workbook avoids dynamic-array functions, macros, Power Query and external links.

## 3. Calculation rules

### Action, milestone and outcome (three separate measures)

* **Credited complete** = `Status = "Complete"` and a numeric Completed Date.
* **Action completion %** = credited actions ÷ non-cancelled actions. It is count-based.
  The *Reported % Complete* column is informational and is **never averaged**.
* **Milestone completion %** = the same calculation, restricted to actions where `Milestone? = Yes`.
* **Outcome attainment %** = (Baseline − Latest Actual) ÷ (Baseline − Target).
  It shows N/A if the baseline, target or actual is missing, or if the baseline equals the target. The latest actual comes from the most recent 06A row that has a KPI value.

### Schedule

* **OVERDUE**: not credited, not cancelled, and Due < `AsOfDate`.
* **DUE SOON**: due within `DueSoonDays`.
* **Days remaining** = Due − `AsOfDate`.
* `AsOfDate` = the override date in Master Data C6 if one is entered, otherwise `TODAY()`.

### RAG

| Status | Condition |
|---|---|
| **CLOSED** | Project status is Complete or Cancelled. |
| **PENDING** | Project status is Pending Assignment. |
| **RED** | Blocked, **or** overdue actions ≥ `RedOverdueThr` (default 3), **or** past the target date. |
| **AMBER** | Any overdue action, an open risk, no primary owner, no target date, or a missing baseline after Define. |
| **GREEN** | Otherwise. |

**Requiring attention** = RED + AMBER + PENDING.

### Savings

| Step | Formula |
|---|---|
| Waste | (Baseline loss rate − Current loss rate) × **comparison** volume × validated unit cost. Using the loss rate normalizes for volume. A volume shift above `VolTol` is flagged. |
| Changeover | Minutes reduced × changeovers per period × Finance-approved $/min, **only if** "benefit demonstrated" = Yes. Otherwise the cash value is $0. Capacity hours (minutes × frequency ÷ 60) are always reported separately as non-cash. |
| SAP / other | Labor, downtime and error/rework savings per period, entered from data. A missing data source is flagged. |
| Recurring per period | Sum of whichever components are present. N/A if there are none. |
| Annualized | Recurring × (`OpWeeks` ÷ period weeks). Default 52 weeks; this is an **assumption** for Finance to confirm. |
| Gross | Annualized + one-time benefit. |
| Net | Gross − implementation cost. |
| Attributed forecast | Net × attribution % (blank means 100%). |
| Verified | Realized (Finance-validated) **only when** Finance Status = Approved; otherwise 0. |
| Double counting | Lines that share a *Benefit Source Key* may not total more than 100% attribution (Rejected lines excluded). Breaches are flagged **OVER-ATTRIBUTED**. |

Hard Savings, Cost Avoidance and Productivity Capacity are summed separately everywhere. The dashboard **Projected Annual Savings** card shows Hard Savings that are not Rejected.

### Trends

**Weekly trend.** The 12 weekly points end on `AsOfDate`.

| Series | Definition |
|---|---|
| Planned | `COUNTIFS(due ≤ week end)` |
| Completed | `COUNTIFS(completed date ≤ week end, credited)` |
| Overdue | Due < week end and not completed by the week end |

**Monthly savings trend.** The forecast run-rate is `SUMPRODUCT(monthly hard run-rate × months since start)`. Verified savings are dated by the Finance review date.

## 4. Data integrity controls

* Unknown values are left blank. Ratios return **N/A** and are never set to zero. Sums exclude N/A, and the record is flagged.
* Data checks run on every action:
  * Unknown project
  * Complete without a date
  * Date entered but status not Complete
  * No owner
  * Completion date in the future
* The register has an owner check: leader missing, leader equal to co-owner, or leader not in the people list.
* The dashboard runs 12 live reconciliation checks:
  * Department and shift totals equal registered projects.
  * Tracker totals equal the sum of project totals for actions, completions, forecast and verified savings.
  * Verified dollars on non-Approved lines = 0.
  * No over-attribution.
  * The latest trend point equals the current KPI.
  * No duplicate weekly updates.
  * No orphan IDs.
* Sheet protection (no password) locks every formula cell. Filters and column/row formatting stay available.

## 5. Maintenance

| Task | How |
|---|---|
| Add a supervisor or list value | Type it directly below the last entry in **11 Master Data**. Leave no blank rows. |
| Add a project | Type the title in the next empty register row. The ID is generated automatically. |
| Change settings | Master Data C6–C13: As-Of date, ID prefix, weeks per year, RED threshold, due-soon window, volume tolerance, evidence folder URL. |
| Extend capacity | Review > Unprotect Sheet. Copy the last table row down, resize the table (Table Design > Resize Table), then update the matching `<prefix>_*` names in Name Manager. Or change the `*_CAP` constants in `build_workbook.py` and regenerate. |
| Change a structure or formula | Edit `build_workbook.py`, regenerate, run `test_workbook.py`, then paste the existing data (**values only, input columns only**) into the new file. |
| System lists | Existing values are locked because formulas reference them. New values can be added; they behave as "other" states. Don't rename the existing ones. |

## 6. Multi-user and integration notes

Excel limits what the workbook can do on its own. It does **not** do any of the following:

* **Secure multi-user ownership.** Microsoft 365 co-authoring on SharePoint/OneDrive lets several people edit at the same time, but anyone with edit rights can change any input. Use SharePoint permissions and version history to control and audit changes.
* **Document uploads or live SharePoint synchronization.** The Evidence Register stores hyperlinks only, and viewers need access to the linked folder. Adding automated uploads or approvals would need a configured integration such as Power Automate or a SharePoint list. None is included.
* **Secure electronic signatures.** Approval dates record reviews only.
