# CHG Process Mapping & Continuous Improvement System

C.H. Guenther & Son — Duncanville, TX. First deployment: **Nugget Line**.

## What's here

| Path | What it is |
|---|---|
| `app/CHG_Process_Mapper.html` | The interactive app. Open it in Chrome, Edge or Safari (desktop or phone). |
| `templates/CHG_Nugget_Line_Process_Mapping_Workbook.xlsx` | **Primary data workbook** — blank Nugget Line template with live formulas, dropdowns and the floor-walk checklist. Also contains Rolled Dough Lines #3 and #4 as unverified templates. |
| `exports/Nugget_Line_Current_State_TEMPLATE_*` | The Nugget Line maps and report with Not Yet Collected placeholders (PowerPoint + PDF). |
| `exports/Nugget_Line_DEMO_*` | The same outputs filled with **demonstration data** — illustrative only, not plant measurements. For training and for checking the calculations. |
| `docs/ARCHITECTURE.md` | System design, data model, formulas, finance validation, extension steps, known limits. |
| `app/src/` · `tools/` · `tests/` | Source, build scripts and the end-to-end test. |

## Nugget Line first deployment — where each deliverable lives

| # | Deliverable | In the app | In the workbook |
|---|---|---|---|
| — | Floor-walk data collection checklist | Floor-Walk Checklist | Checklist |
| 1 | Current-state process flow | Process Flow | Flow Nodes / Flow Edges |
| 2 | Current-state value stream map | Value Stream | Step Metrics, KPI Calculations |
| 3 | Process activity analysis | Activity Map | Activities, PAM Summary |
| 4 | Waste & bottleneck opportunity register | Waste & Opportunities (with suggestions from data) | Opportunities |
| 5 | Future-state mapping template | State selector → Future state; Current vs Future | Steps / Measurements rows with State = Future |
| 6 | Improvement action tracker | Actions | Actions |
| 7 | Financial savings validation sheet | Finance Validation | Finance Validation, KPI Calculations (savings pipeline) |

## How to use

1. Print or open the **Checklist** (workbook or app) and walk the Nugget Line.
2. Enter values in the yellow cells of **Settings** and **Measurements** with Status (Verified / Estimate) and Source.
3. In the app: Lines & Export → Import the workbook. The maps, KPIs, bottleneck and suggestions update.
4. Create the future state, edit it, and compare. Record opportunities, actions and Finance validation.
5. Export to Excel (system of record), PowerPoint (editable maps) or PDF.

## Rebuild and test

```bash
python3 tools/build_app.py                                   # build the app
node tools/dump_project.js <dir>                             # dump template + demo data (needs Playwright)
python3 tools/build_workbook.py <dir>/project_dump.json templates/CHG_Nugget_Line_Process_Mapping_Workbook.xlsx
node tests/smoke_test.js <dir with xlsx, pptxgenjs, jszip, jspdf, jspdf-autotable> exports
```
