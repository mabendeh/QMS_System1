# CHG Process Mapper — system architecture

C.H. Guenther & Son — Duncanville, TX · Continuous Improvement

## 1. Purpose

One reusable system for mapping every production line: process flow, value stream, spaghetti, swimlane and process activity maps, built on a single set of step data. The system identifies waste and bottlenecks, tracks improvement actions, and keeps estimated savings separate from Finance-validated savings.

**Ground rule:** the system never invents measurements, equipment arrangements or savings. Every value is one of:

| Status | Meaning |
|---|---|
| **Verified** | Measured on the floor or taken from a system of record. Source recorded (who / when / how). |
| **Estimate** | Assumption, planning value or engineering estimate. Basis recorded. |
| **Not Yet Collected (NYC)** | Blank. Shown in amber everywhere. Never filled in by the system. |

Every calculated result inherits the worst status of its inputs: **Verified inputs** (all verified), **Estimate** (any estimate) or **Incomplete** (any input not collected; partial sums are flagged).

## 2. Components

```
                    ┌──────────────────────────────────────────────┐
  Floor walk ──►    │  Excel workbook (system of record)           │ ◄── Finance, Planning
  (Checklist)       │  data sheets + live formula sheets           │
                    └───────────────┬──────────────▲───────────────┘
                          import    │              │  export
                    ┌───────────────▼──────────────┴───────────────┐
                    │  CHG Process Mapper (single-page web app)    │
                    │  ┌────────────┐ ┌─────────────┐ ┌──────────┐ │
                    │  │ Data model │►│ Calculation │►│ 14 views │ │
                    │  │ (01_core)  │ │ (02_calc)   │ │ (03–05)  │ │
                    │  └────────────┘ └─────────────┘ └──────────┘ │
                    │  I/O (06_io): autosave · xlsx · pptx · pdf ·  │
                    │  svg · png · json                             │
                    └──────────────────────────────────────────────┘
```

| Layer | File | Responsibility |
|---|---|---|
| Data model & templates | `app/src/js/01_core.js` | Field definitions, line templates (Nugget, Rolled Dough #3/#4, any new line), default flow / activities / layout / swimlane, clearly labeled demonstration data |
| Calculation engine | `app/src/js/02_calc.js` | All KPIs with status propagation; `FORMULAS` is the single formula reference shown in the app, workbook and reports |
| Diagram renderers | `app/src/js/03_diagrams.js` | One SVG generator per map type; the same geometry drives the screen, SVG/PNG, PDF and the editable PowerPoint shapes |
| Views & editors | `app/src/js/04_ui.js` | Dashboard, 5 maps, step data, waste register, actions, finance validation, comparison, checklist, formulas, lines & export |
| Interaction | `app/src/js/05_events.js` | Drag-and-drop blocks / zones / swimlane activities, connectors and decision branches, route drawing, zoom, in-page confirmations |
| I/O | `app/src/js/06_io.js` | Autosave, Excel import/export, PowerPoint (native shapes + compatibility fixes), PDF, SVG/PNG, JSON backup |
| Workbook builder | `tools/build_workbook.py` | Styled Excel template with data validation and live formulas (Step Metrics, KPI Calculations, PAM Summary, route formulas) |
| Build | `tools/build_app.py` | Combines sources into `app/CHG_Process_Mapper.html` (standalone) and `app/dist/artifact_page.html` (claude.ai) |
| Tests | `tests/smoke_test.js` | End-to-end browser test: every tab, interactions, all exports, Excel round trip, phone width |

## 3. Data model

```
Project
 ├─ lines[]                      Nugget Line, Rolled Dough Line #3, #4, … (add / duplicate / delete)
 │   ├─ unit, notes
 │   ├─ settings{}               demand, shifts, shift length, breaks, days, weeks, labor rate,
 │   │                           material cost, walking & forklift speed — each {v, s, src}
 │   ├─ current: State
 │   └─ future:  State | null    created as a copy of current, then edited
 ├─ opps[]                       waste register (8 wastes) + fin{} validation record
 ├─ actions[]                    owner, due date, status, linked opportunity
 └─ checks{}                     floor-walk checklist ticks

State
 ├─ steps[]       name, department, owner, equipment, applies-to-line (To verify / Confirmed / N/A),
 │                VA class, food-safety & quality notes (CCP, allergen, sanitation, traceability, QC),
 │                f{16 fields}: cycle time, processing time, changeover, changeovers/week, downtime,
 │                uptime, labor, throughput, scrap, yield, WIP, walking, transport distance,
 │                waiting time, OEE, operating cost — each {v, s, src}
 ├─ flow          nodes (start, operation, decision, inspection, transfer, delay, storage, rework,
 │                scrap, end; x/y; department; linked step) + edges (labels for decision branches)
 ├─ acts[]        process activity map rows (O / T / I / D / S, time, distance, people, VA class)
 ├─ layout        building size (ft) + zones; routes[] (category, polyline in ft, trips, people)
 └─ swim[]        lane (7 departments), column, type, flags (duplicate / unnecessary approval /
                  delay / communication gap), next[]
```

Lines never share steps: each line's step list, equipment and layout are independent. A new line starts from the plant-wide sequence with every step marked **To verify**.

## 4. Calculations

All formulas are listed in the app (Formulas & Sources), the workbook (Formulas sheet) and every report. Key ones:

| Result | Formula |
|---|---|
| Takt time | Shifts × (shift length − breaks) × 60 ÷ demand per day |
| Effective cycle time | Cycle time ÷ uptime (uptime missing → cycle time, marked Estimate) |
| Bottleneck | Highest effective cycle time; provisional until every step is timed |
| Queue time | WIP ÷ demand per day × available minutes per day (or observed waiting time) |
| Lead time / VA / PCE | Σ queue + Σ processing; Σ processing of VA steps; VA ÷ lead time |
| Rolled throughput yield | Π yield (or 1 − scrap) |
| Travel | Route length (drawn, 1 unit = 1 ft) × trips ÷ speed; × people × shifts × days ÷ 60; × labor rate |
| Savings (estimates) | Travel, changeover, scrap-at-cost, throughput — current vs future; never reported as validated |

The workbook implements the same calculations as Excel formulas. On the demonstration data, all 12 KPIs in the workbook match the app (travel within 0.01% because the workbook stores route distance to 0.1 ft).

## 5. Finance validation

Each opportunity carries a benefit type (hard savings, cost avoidance, capacity, nonfinancial), baseline, valuation method, estimate, validation status (Not started → Method agreed with Finance → In verification → Validated by Finance / Rejected), validated amount, reviewer and date. Dashboards and reports show estimated-not-validated and validated totals separately; only **validated hard savings** are savings. The validated amount can only be entered once the status is Validated by Finance.

## 6. Persistence and sharing

* **Excel is the primary data format.** Export → collect / review in Excel → import. IDs keep rows linked.
* **Autosave** of the working copy: on claude.ai, privately to the signed-in user's account (one document per line); in a standalone browser, to that browser's local storage.
* **Backup**: JSON export / restore.

## 7. Exports

| Export | Content |
|---|---|
| Excel (.xlsx) | All lines and states: settings, steps, measurements, activities, layout, routes, flow, swimlane, opportunities, actions, finance validation, checklist, KPIs, comparison, formulas |
| PowerPoint (.pptx) | Native, editable shapes and connectors: process flow, VSM, spaghetti, swimlane, KPI and register tables, comparison. Post-processed for PowerPoint mobile: unique shape IDs, separate notes-master theme, no ZIP folder entries |
| PDF | Line report (KPIs, all maps, step data, activity map, register, actions, open checklist items, formulas) and current-vs-future comparison report, 16:9 pages |
| SVG / PNG | Any diagram, light print palette, from the diagram toolbar |

All exports carry a DEMONSTRATION DATA banner when demo data is loaded.

## 8. Extending to other lines

1. Lines & Export → Add line from plant template (or Duplicate an existing line).
2. Step Data: rename, add, remove, reorder steps; set Applies to line; enter equipment.
3. Process Flow → Rebuild from steps, then adjust blocks and branches.
4. Spaghetti: replace the placeholder zones with the plant drawing dimensions; draw routes.
5. Run the Floor-Walk Checklist; export to Excel.

## 9. Known limits of this prototype

* The factory layout is a schematic placeholder, not the Duncanville drawing.
* No official C.H. Guenther logo was supplied; the company name is shown as text (element marked `data-logo-placeholder`).
* Export libraries (SheetJS, PptxGenJS, jsPDF, JSZip) load from the jsDelivr CDN at the moment of export, so exporting needs an internet connection.
* Autosave on claude.ai is private to each user; team collaboration is through the Excel workbook.
* PowerPoint files are checked with the OOXML validator and LibreOffice rendering; they have not been opened in Microsoft PowerPoint from this environment.
