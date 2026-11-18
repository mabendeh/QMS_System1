"""Builds CHG_DunGPS_Dashboard_User_Guide.pdf (concise user instructions). Usage: python build_guide.py <out_dir>"""
import os
import sys
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
                                Image, PageBreak, KeepTogether)

NAVY = colors.HexColor("#1B2A4A")
BURG = colors.HexColor("#8C1D40")
GRAY = colors.HexColor("#595959")
LG = colors.HexColor("#F2F2F2")
HERE = os.path.dirname(os.path.abspath(__file__))

ss = {
    "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=15, textColor=NAVY, spaceBefore=4, spaceAfter=6),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11, textColor=BURG, spaceBefore=8, spaceAfter=3),
    "b": ParagraphStyle("b", fontName="Helvetica", fontSize=9, leading=12.5, textColor=colors.black, spaceAfter=3),
    "s": ParagraphStyle("s", fontName="Helvetica", fontSize=8, leading=10.5, textColor=GRAY),
    "c": ParagraphStyle("c", fontName="Helvetica", fontSize=8, leading=10.5),
    "ch": ParagraphStyle("ch", fontName="Helvetica-Bold", fontSize=8, leading=10.5, textColor=colors.white),
    "title": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=26, textColor=NAVY, leading=31),
    "sub": ParagraphStyle("sub", fontName="Helvetica", fontSize=13, textColor=GRAY, leading=18),
}
P = lambda t, s="b": Paragraph(t, ss[s])


def tbl(rows, widths, head=True):
    data = [[P(c, "ch" if (head and i == 0) else "c") for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=widths, repeatRows=1 if head else 0)
    st = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")), ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
          ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    if head:
        st.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    t.setStyle(TableStyle(st))
    return t


def img(name, w):
    path = os.path.join(HERE, "assets", name)
    from PIL import Image as PI
    iw, ih = PI.open(path).size
    return Image(path, width=w, height=w * ih / iw)


def deco(c, doc):
    W, H = doc.pagesize
    c.setFillColor(BURG)
    c.rect(0, H - 6, W, 6, stroke=0, fill=1)
    c.setFont("Helvetica", 7.5)
    c.setFillColor(GRAY)
    c.drawString(0.6 * inch, 0.4 * inch, "CHG Duncanville · DunGPS Deployment Management System · User Guide")
    c.drawRightString(W - 0.6 * inch, 0.4 * inch, f"Page {doc.page}")


def build(out):
    path = os.path.join(out, "CHG_DunGPS_Dashboard_User_Guide.pdf")
    ps = landscape(letter)
    doc = BaseDocTemplate(path, pagesize=ps, leftMargin=0.6 * inch, rightMargin=0.6 * inch, topMargin=0.55 * inch,
                          bottomMargin=0.6 * inch, title="CHG DunGPS Dashboard User Guide", author="CHG Duncanville CI")
    fr = Frame(doc.leftMargin, doc.bottomMargin, ps[0] - 1.2 * inch, ps[1] - 1.15 * inch, id="f")
    doc.addPageTemplates([PageTemplate(id="p", frames=[fr], onPage=deco)])
    W = ps[0] - 1.2 * inch
    s = []
    # cover
    s += [Spacer(1, 1.2 * inch), P("DunGPS Pillar Deployment<br/>&amp; Performance Dashboard", "title"), Spacer(1, 8),
          P("User Guide · C.H. Guenther &amp; Son — Duncanville", "sub"), Spacer(1, 20),
          P("Three separate Excel workbooks – no macros, Excel 2010+ / Excel for the web:", "b"),
          tbl([["File", "Purpose"],
               ["CHG_Bronze_Level_Tracker.xlsx  (START HERE)", "Simple Bronze tracker: the 32 official requirements, one row each, in GPS pillar order (LEAD, EHS, FSQ, AM, PM, VO, E&amp;T, IM, SC). Owner dropdown, status, % complete, evidence link, verified? and a one-page dashboard by pillar and by the 3 GPS layers."],
               ["CHG_DunGPS_Deployment_Master.xlsx", "Full management system (optional, for deeper deployment control): tracker, gaps, Bronze Star readiness, process matrix, VO projects, evidence register, executive &amp; management-review dashboards."],
               ["CHG_Supervisor_Productivity_Project_Tracker.xlsx", "Separate supervisor project tracker: supervisors choose projects (5S, standard work, SMED, problem solving…) from a library; 5-milestone tracking, coach verification, scorecard and gaps dashboard."],
               ["*_DEMO.xlsx versions", "Same system loaded with clearly labelled fictitious data, used for testing and training. Never use for decisions."]],
              [3.0 * inch, W - 3.0 * inch]),
          Spacer(1, 16),
          P("The system answers five leadership questions immediately: <b>Where are we? What is complete? What is missing? Who owns it? When will we close the gaps?</b>", "b"),
          P("Official content loaded: the 9 pillars and 32 Bronze Level requirements from <i>GPS_Pillar_Bronze_Level_Requirements.xlsx</i>, imported word for word. "
            "Nothing else is presented as official CHG content. No owners, dates or savings have been invented.", "s"),
          PageBreak()]
    s += [P("Bronze Level Tracker – the simple way to run Bronze", "h1"),
          tbl([["Step", "What to do"],
               ["1. Owners", "Owners sheet: type each person once (Active? = Yes). They appear in every Owner / Verified By dropdown."],
               ["2. Assign", "Bronze Tracker: choose Owner, Start Date and Due Date for each requirement."],
               ["3. Update weekly", "Status, % Complete, Gap / What's Missing, Next Step, Last Updated."],
               ["4. Upload evidence (mandatory)", "Click <b>Upload ↗</b> on the row – it opens the pillar's Teams folder (Teams → DunGPS → General → Shared → pillar). Upload the file(s), use <i>Copy link</i>, paste it in <b>Evidence Link</b> and describe what was uploaded. Tip: one sub-folder per requirement (e.g. FSQ-3)."],
               ["4b. Complete", "'Complete' only appears in the Status dropdown (and 'Yes' in Verified?) after an evidence link is pasted – no evidence, no completion. Rows marked Complete without evidence (e.g. pasted data) show <b>Evidence Missing</b> in red."],
               ["5. Verify", "Verifier sets Verified? = Yes. Only then is the requirement Done."],
               ["Health colours", "Done (green) · On Track (green) · Needs Verification (amber) · Evidence Missing (red) · At Risk – due within 14 days and under 75% (amber) · Overdue (red) · Not Started (gray)."],
               ["Dashboard", "Bronze-ready banner, % done, average progress, counts, progress by pillar, the 3 GPS layers (Foundation · Equipment Focus · Capability &amp; Governance), needs-attention list, owner summary, charts. Prints on one page."]],
              [1.3 * inch, W - 1.3 * inch]),
          PageBreak(), P("Bronze Level Tracker – dashboard", "h1"), img("bronze.png", W * 0.48), PageBreak()]
    # 1 quick start
    s += [P("1. Quick start (first 30 minutes)", "h1"),
          tbl([["Step", "Where", "What to do"],
               ["1", "Owner Directory", "Enter every person who will own, support, verify or escalate (name, department, position, shift, pillar, Active, email). They appear in all owner dropdowns immediately."],
               ["2", "Configuration", "Assign pillar leads, Teams sub-folders, RAG thresholds and the status-date setting. Add or edit areas, phases and categories at the end of each list."],
               ["3", "Pillar Deployment Tracker", "Each official Bronze requirement is pre-loaded as one task (T-AM-01 …). Break them into steps as needed (add rows), then set Owner, Approver/Verifier, Start and Target dates, and Evidence Required."],
               ["4", "Bronze Star Readiness", "Assign a Responsible Owner and Due Date per criterion. Enter CHG's official evidence requirement where provided. The 'Suggested Evidence' column is a CI recommendation only."],
               ["5", "Process Deployment Matrix", "Add equipment/shift rows per area. Mark each element N/A, Not Started, In Progress, Implemented or Verified. Nothing is assumed to apply (default Not Assessed)."],
               ["6", "VO &amp; Tiger Team Projects", "Complete the Pre-Weigh Room charter once the baseline is measured. Add further projects."],
               ["7", "Management Review", "Use it in the weekly meeting. Filter (yellow cells), review, record decisions, print on one landscape page."]],
              [0.45 * inch, 1.9 * inch, W - 2.35 * inch]),
          Spacer(1, 6),
          P("<b>Cell conventions.</b> White cells under a navy header are inputs. Gray cells under a gray header marked <b>ƒ</b> are formulas – never type over them. "
            "Dashboards and the Calc Engine are protected without a password (Review → Unprotect Sheet) so they cannot be overwritten by accident. "
            "Registers are left unprotected so you can sort, filter and add rows.", "b"),
          P("<b>Adding rows.</b> Tables are pre-formatted (tracker 500, gaps 300, documents 300, owners 200, projects 100, process rows 80). Type in the next empty row. "
            "To extend, drag the table's bottom-right handle; calculated columns copy automatically and all dashboards use full-table ranges, so new rows are counted with no formula edits.", "b"),
          PageBreak()]
    # 2 dashboards
    s += [P("2. Executive Dashboard", "h1"),
          P("Site view: 8 KPI tiles, the Bronze Star banner, the five leadership questions answered in sentences from live data, activity status, completion by pillar and phase, critical issues, and trend / readiness / closure-forecast charts.", "b"),
          img("exec.png", W * 0.80), PageBreak(),
          P("3. Management Review (weekly, one page)", "h1"),
          P("Pick a value in the yellow filter cells (Pillar, Owner, Process/Area, Shift, Phase); choose (All) to clear. Filters apply to tasks (all five), gaps and VO projects (pillar, owner, area). "
            "Covers overall and pillar completion, planned vs actual, top 5 gaps, overdue actions, owner accountability, Bronze readiness, VO status, Finance-verified savings and decisions needing leadership support. "
            "Record the meeting's decisions in the log at the bottom right. File → Print is preset to one landscape page.", "b"),
          img("mr.png", W * 0.62), PageBreak(),
          P("4. Gap Analysis Dashboard", "h1"),
          P("Separates the <b>open backlog</b> from the <b>cumulative closure rate</b>. Shows open gaps by pillar and severity, overdue gaps by owner, severity and aging, top recurring barriers (gap categories plus tracker barrier categories), top 5 gaps, Pareto, and new vs closed per month with the backlog line.", "b"),
          img("gap.png", W * 0.80), PageBreak()]
    # 5 weekly cadence
    s += [P("5. Weekly operating cadence &amp; responsibilities", "h1"),
          tbl([["When", "Who", "Action"],
               ["Mon–Wed", "Task owners", "Update Status, Work Progress %, Actual Completion, Evidence Link (Teams URL), Gap/Barrier, Next Action, Last Updated."],
               ["Thursday", "Approver / Verifier", "Open the evidence (Open ↗), set Verification Status = Verified or Rejected - Rework, and Verification Date. Only verified tasks count as complete."],
               ["Thursday", "Pillar leads", "Log new gaps; clear Data Quality Flags (missing owner, date, evidence, duplicate IDs, inactive owners); update the process matrix and Bronze criteria."],
               ["Friday", "Plant leadership", "Management Review: walk the filters, decide on escalations, record decisions. Executive Dashboard for the monthly / Rhonda Silbaugh review."],
               ["Monthly", "Admin / CI", "Review thresholds, add new owners/areas, archive obsolete documents, check document review dates."]],
              [1.0 * inch, 1.6 * inch, W - 2.6 * inch]),
          P("Evidence &amp; documents", "h2"),
          P("Store files in <b>Teams → DunGPS → General → Shared → &lt;Pillar folder&gt;</b>. In Teams/SharePoint use <i>Copy link</i> and paste the full URL into the Evidence Link or Document Link column. "
            "The workbook links to files – it never stores copies. The Document &amp; Evidence Register tracks owner, revision, last/next review and approval status; a document is complete only when approved/verified, linked and within its review cycle.", "b"),
          P("Gaps", "h2"),
          P("Record current vs required condition, severity, root cause, corrective action, owner and due date. A gap is closed only when Status = Closed - Verified, Verification = Verified and a Closure Date is entered; "
            "otherwise it stays in the open backlog. Tick <i>Leadership Support Needed</i> and describe the decision required to surface it on the Management Review.", "b"),
          PageBreak()]
    # 6 calculation rules
    s += [P("6. How the numbers are calculated", "h1"),
          tbl([["Metric", "Definition"],
               ["In scope (denominator)", "Pillar entered, Required? ≠ No and Applicability ≠ 'N/A - Approved'. Approved N/A and optional tasks are shown separately and never counted. 'N/A - Requested' counts until approved; an approved N/A requires an Approver."],
               ["Verified (task)", "Status = Complete AND evidence link present AND Verification Status = Verified. Selecting 'Complete' alone never counts as verified."],
               ["Verified Completion %", "Verified required tasks ÷ total applicable required tasks × 100 (weighted if 'Use task weights' = Yes). Rolls up task → Bronze requirement (step) → pillar → phase → overall."],
               ["Completion credit %", "Verified = 100%; Complete but unverified = max credit (default 80%); otherwise reported progress capped at the same max. Drives schedule RAG."],
               ["Work progress %", "Owner-reported progress, uncapped – transparency only."],
               ["Evidence submitted / verified %", "Share of in-scope tasks with a link / with verified evidence."],
               ["Planned % to date", "Linear from Start to Target date at the status date. The trend chart compares cumulative planned (target dates) with cumulative verified completion over 26 periods."],
               ["Overdue", "Target date before the status date and not verified. Escalated when days late exceed the escalation threshold (default 14)."],
               ["Gap closure rate", "Verified closed gaps ÷ total gaps identified (excluding cancelled). Open backlog reported separately."],
               ["Bronze readiness", "Criterion = 'Verified - Ready' only if verified, evidence linked, all linked tasks verified and no open gaps. The banner reads BRONZE STAR READY only when every mandatory, applicable criterion is ready."],
               ["VO savings", "Estimated opportunity and Finance-verified savings are separate; savings count only when Finance Approval = Approved."],
               ["Process matrix score", "Not Started 0 · In Progress 25% · Implemented 75% · Verified 100%, averaged over applicable elements; coverage = assessed ÷ 9."]],
              [1.9 * inch, W - 1.9 * inch]),
          P("RAG logic (thresholds on the Configuration sheet)", "h2"),
          tbl([["Colour", "Task", "Pillar / phase"],
               ["Green", "Verified, or completion credit within the Green tolerance (5 pts) of plan", "On plan and nothing overdue"],
               ["Yellow", "Within the Yellow tolerance (15 pts), On Hold, or verification rejected", "Behind plan within tolerance, or any overdue task"],
               ["Red", "Overdue, or further behind plan", "Critical open gap, overdue share ≥ 10%, or behind plan beyond tolerance"],
               ["Gray", "N/A, or not started and not yet planned", "No scope, or nothing started or planned"]],
              [0.8 * inch, (W - 0.8 * inch) / 2, (W - 0.8 * inch) / 2]),
          PageBreak()]
    # 7 supervisor program
    s += [P("7. Supervisor Productivity Project Tracker (separate file)", "h1"),
          tbl([["Step", "What to do"],
               ["Roster", "Enter supervisors (In Program? = Yes), managers/coaches and Finance."],
               ["Choose", "Project Assignments → select Supervisor and a Project Code. Name, category, pillar, KPI, duration and milestone definitions fill in from the Project Library (24 CI-recommended project types: 5S, 5 Why/A3, standard work, CIL, centerlines, SMED, Redzone loss reduction, FIT, SPIOC, giveaway, sanitation time, skill matrix, leader standard work …)."],
               ["Evidence", "Upload ↗ on each project row opens the Teams folder; paste the link. Complete / Verified only become selectable once a link is pasted."],
               ["Track", "Set M1–M5 to In Progress / Done. 'Current Milestone' shows the next definition of done. Log blockers on the Gap &amp; Barrier Log with the Assignment ID."],
               ["Verify", "When all 5 milestones are done: Status = Complete, evidence link, coach sets Verified. Savings count only if Finance Approved? = Yes."],
               ["Review", "Program Dashboard (filters: area, shift, category): scorecard per supervisor vs the program target (default 2 verified projects per period), completion by category/area, latest 5S score, gaps panel, overdue list, Pareto."]],
              [1.0 * inch, W - 1.0 * inch]),
          Spacer(1, 6), img("sup.png", W * 0.66), PageBreak()]
    # 8 admin & missing info
    s += [P("8. Administration &amp; technical notes", "h1"),
          P("• <b>Lists:</b> add values at the end of a list on the Configuration sheet; no blank rows inside a list. Gray-header lists are system values used by formulas – do not rename them.<br/>"
            "• <b>Status date:</b> TODAY() unless an override date is entered (useful to freeze a meeting pack).<br/>"
            "• <b>Capacity:</b> dashboards show up to 12 pillars and 24 supervisors; registers can be extended.<br/>"
            "• <b>Compatibility:</b> SUMPRODUCT, COUNTIFS, INDEX/MATCH, EDATE, EOMONTH, HYPERLINK only. No macros, Power Query, Power Pivot or Microsoft 365-only functions. Works in Excel for the web inside Teams (co-authoring).<br/>"
            "• <b>Validation performed:</b> every formula recalculated with zero errors; completion, overdue, gap, Bronze and scorecard results independently recomputed and matched on the demo data; zero-denominator, N/A, optional, missing owner, missing evidence, inactive owner, duplicate ID, overdue and new-row cases tested; filters tested.", "b"),
          P("9. Information required from CHG", "h1"),
          tbl([["#", "Item", "Why it is needed"],
               ["1", "Pillar names – RESOLVED from the CHG GPS poster (IM = Initiative Management)", "–"],
               ["2", "Per Bronze criterion: mandatory/optional, acceptance criteria, official evidence, verifier, approved exceptions", "Readiness logic and evidence columns"],
               ["3", "Silver and Gold criteria", "Load deployment phases 2–4"],
               ["4", "Owner names, departments, positions, shifts, emails, pillar leads, verifiers, escalation owners", "Owner dropdowns and accountability (none invented)"],
               ["5", "Program launch, Bronze assessment target and step-level target dates", "Planned %, overdue and trend"],
               ["6", "Whether steps are weighted, and the weights", "Weighted completion"],
               ["7", "Confirmed areas, equipment/line names and shifts; which elements apply where", "Process Deployment Matrix"],
               ["8", "Teams/SharePoint folder structure and site URL", "Evidence and document links"],
               ["9", "Finance savings-validation rules and approver; Pre-Weigh project baseline", "VO dashboard"],
               ["10", "RAG thresholds and escalation rules; supervisor program target per period", "Configuration defaults"]],
              [0.35 * inch, 5.2 * inch, W - 5.55 * inch])]
    doc.build(s)
    return path


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
