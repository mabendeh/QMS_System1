# C.H. Guenther & Son, Duncanville: Supervisor Productivity & CI Management System

This folder holds an Excel workbook for managing continuous improvement (CI) projects. Supervisors lead and own the projects, the CI Manager coaches, and Finance validates the savings.

**Workbook files** (in `output/`):

| File | Use |
|---|---|
| `CHG_Duncanville_CI_Management_System.xlsx` | The **production** file. It contains only the confirmed assignments. |
| `CHG_Duncanville_CI_System_DEMO_TRAINING.xlsx` | A training and test copy filled with **fictitious** data. |

**Documents** (in `docs/`):

* [Supervisor Quick-Start Guide](docs/Supervisor_Quick_Start_Guide.md)
* [Technical Documentation](docs/Technical_Documentation.md) explains the calculations and how to maintain the workbook.
* [QA Report](docs/QA_Report.md) lists the tests performed and the remaining limitations.

To rebuild and retest, run these commands in order. They need Python 3, openpyxl and LibreOffice for recalculation.

```bash
python3 build_workbook.py                      # build both workbooks
python3 test_workbook.py                       # recalculate and run 87 automated tests -> qa_results.json
python3 build_workbook.py --qa qa_results.json # embed the QA results in tab 14
python3 make_qa_report.py                      # refresh docs/QA_Report.md
```
