# CHG Duncanville — DunGPS Deployment Management System

Excel-based (no macros) management system for TPM / DunGPS pillar deployment and the Supervisor Productivity Program.

## Deliverables (`out/`)

| File | Purpose |
|---|---|
| `CHG_Bronze_Level_Tracker.xlsx` | **Start here.** Simple Bronze tracker: 32 official requirements in GPS pillar order, owner dropdown, status, % complete, evidence link, verified?, plus a one-page dashboard (by pillar and by the 3 GPS layers). |
| `CHG_Bronze_Level_Tracker_DEMO.xlsx` | Demo version for training. |
| `CHG_DunGPS_Deployment_Master.xlsx` | Optional full management system: 13 sheets (Executive Dashboard, Pillar Deployment Tracker, Process Deployment Matrix, Gap Analysis Dashboard, Gap Register, Owner Directory, Bronze Star Readiness, VO & Tiger Team Projects, Document & Evidence Register, Management Review, Configuration & Dropdown Lists, Instructions & Data Definitions, Calc Engine). Pre-loaded only with the 9 pillars and 32 Bronze requirements from the CHG source file plus the Pre-Weigh project (baseline pending). |
| `CHG_DunGPS_Deployment_Master_DEMO.xlsx` | Same system with clearly labelled fictitious data, used for testing and training. |
| `CHG_Supervisor_Productivity_Project_Tracker.xlsx` | **Separate** supervisor project tracker: project library (24 project types), assignments with 5 milestones and coach verification, gap & barrier log, 5S audit log, roster, dashboard (scorecard, completion rate, gaps). |
| `CHG_Supervisor_Productivity_Project_Tracker_DEMO.xlsx` | Demo version of the above. |
| `CHG_DunGPS_Dashboard_User_Guide.pdf` | Concise user guide (9 pages). |
| `MISSING_INFORMATION.md` | Information needed from CHG to configure the official system. |

## Rebuild & verify (`build/`)

```bash
cd build
python build_bronze_simple.py ../out   # simple Bronze tracker (+ demo)
python build_master.py ../out          # both master workbooks
python build_supervisor.py ../out      # both supervisor workbooks
python build_guide.py ../out           # user guide PDF
# after recalculating with LibreOffice (recalc.py), independent checks:
python test_master.py ../out/CHG_DunGPS_Deployment_Master_DEMO.xlsx
python test_supervisor.py ../out/CHG_Supervisor_Productivity_Project_Tracker_DEMO.xlsx
```
