#!/usr/bin/env python3
"""Render docs/QA_Report.md from qa_results.json."""
import json, os
HERE = os.path.dirname(os.path.abspath(__file__))
q = json.load(open(os.path.join(HERE, "qa_results.json")))["combined"]
L = ["# Quality Assurance Report", "**C.H. Guenther & Son, Duncanville: CI Management System**", "",
     f"Build {q['build']}. **{q['passed']} of {q['total']} automated tests passed**; {q['failed']} failed.", "",
     "## How testing was done",
     "`test_workbook.py` runs four kinds of tests:",
     "",
     "1. **Full recalculation.** Both workbooks are recalculated in headless LibreOffice. Any formula error value (REF, VALUE, NAME, DIV/0, N/A) is a failure.",
     "2. **Structural checks.** Sheets, tables, named ranges, dropdown targets, protection and locking, freeze panes, conditional formatting, print setup, hyperlinks and fonts. The checks also confirm there are no macros or external links.",
     "3. **Production-content checks.** Only the confirmed assignments are present, with no invented dates, baselines, savings or records.",
     "4. **Independent recomputation and scenarios.** Every key KPI is recomputed in Python from the demo dataset and compared with the Excel results. Scenario tests then edit a temporary copy, recalculate it and check how the dashboard responds:",
     "   * Changing the filters",
     "   * Completing an action",
     "   * Finance approving a savings line",
     "   * Changing the As-Of date",
     "   * Adding a supervisor and a project",
     "",
     "The same results are written into the workbook on tab **14 QA Report**.", "",
     "## Test results", "", "| # | Area | Test | Result |", "|---|---|---|---|"]
for i, t in enumerate(q["tests"], 1):
    L.append(f"| {i} | {t['area']} | {t['test'].replace('|', '/')} | {t['result']} |")
L += ["", "## Known limitations", ""] + [f"- {x}" for x in q["limitations"]]
L += ["", "## Manual checks recommended in Microsoft Excel before first use",
      "The build environment doesn't have desktop Excel, so these checks still need a person:",
      "",
      "1. **Open the production file.**",
      "   * Confirm there is no repair prompt.",
      "   * Confirm the dashboard shows 3 active projects and 5 requiring attention.",
      "   * Confirm all 12 integrity checks read PASS.",
      "2. **Dropdowns and links.** Try one dropdown on each register, and open the evidence link on a test row.",
      "3. **Print preview.** Check that:",
      "   * The dashboard fits on one page.",
      "   * The charter fits on one page.",
      "   * The WOR report prints in landscape.",
      "   * One data sheet prints as expected. Data-sheet print areas grow as records are added.",
      "4. **Set up storage.** Save the file in the approved SharePoint folder, and enter the evidence folder URL in Master Data C13."]
open(os.path.join(HERE, "docs", "QA_Report.md"), "w").write("\n".join(L) + "\n")
print("written")
