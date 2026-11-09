import "server-only";
import { prisma } from "@/lib/db";
import { loadPlant } from "@/lib/data/plant";
import { computeAnalytics, coverageFor } from "@/lib/data/analytics";
import { fmtDate, fmtDateTime } from "@/lib/utils";
import type { Report } from "./render";

const stamp = (who: string) => `Generated ${fmtDateTime(new Date())} by ${who} · CHG Duncanville · CONFIDENTIAL – internal use only`;

/** Promotion review report. Contains NO compensation data. */
export async function promotionReport(who: string): Promise<Report> {
  const reqs = await prisma.promotionRequest.findMany({
    orderBy: { submittedAt: "desc" },
    include: { employee: true, fromLevel: true, toLevel: true, ruleSet: true, approvals: { orderBy: { decidedAt: "asc" } } },
  });
  const users = new Map((await prisma.user.findMany({ select: { id: true, name: true } })).map((u) => [u.id, u.name]));
  return {
    title: "Promotion Review Report",
    subtitle: `${stamp(who)} · Recommendations only – no pay change is made by this system`,
    sheets: [{
      name: "Requests",
      columns: [
        { header: "Employee ID", width: 12 }, { header: "Employee", width: 22 }, { header: "From", width: 18 }, { header: "To", width: 20 },
        { header: "Rule version", width: 10 }, { header: "Status", width: 14 }, { header: "Submitted", width: 12 }, { header: "Submitted by", width: 20 },
        { header: "Decision date", width: 12 }, { header: "Effective date", width: 12 }, { header: "Decision history", width: 60 },
      ],
      rows: reqs.map((r) => [
        r.employee.employeeNumber, `${r.employee.firstName} ${r.employee.lastName}`, r.fromLevel?.name ?? "—", r.toLevel.name, `v${r.ruleSet.version}`,
        r.status.replace(/_/g, " "), fmtDate(r.submittedAt), users.get(r.submittedById) ?? "", fmtDate(r.decisionDate), fmtDate(r.effectiveDate),
        r.approvals.map((a) => `${fmtDate(a.decidedAt)} ${a.step.replace(/_/g, " ").toLowerCase()}: ${a.decision} by ${users.get(a.decidedById) ?? "?"}${a.comments ? ` – ${a.comments}` : ""}`).join(" | "),
      ]),
    }],
  };
}

export async function coverageReport(who: string): Promise<Report> {
  const plant = await loadPlant();
  const { shifts, rows } = await coverageFor(plant);
  return {
    title: "Shift Coverage Report",
    subtitle: `${stamp(who)} · Qualified = every competency in the department currently Competent and not expired`,
    notes: ["Scheduled and available headcount: UNKNOWN – no scheduling system is integrated. Figures show total qualified headcount only."],
    sheets: [{
      name: "Coverage",
      columns: [
        { header: "Department", width: 24 },
        ...shifts.flatMap((s) => [{ header: `Shift ${s.code} qualified`, width: 11 }, { header: `Shift ${s.code} target`, width: 10 }, { header: `Shift ${s.code} status`, width: 11 }]),
        { header: "Total qualified", width: 11 }, { header: "Total target", width: 10 }, { header: "Status", width: 10 },
      ],
      rows: rows.map((r) => [
        r.areaName,
        ...r.cells.flatMap((c) => [c.qualified, c.target ?? "—", c.status === "NO_TARGET" ? "No target" : c.status === "GAP" ? "Gap" : "Covered"]),
        r.totalQualified, r.totalTarget ?? "—", r.status === "NO_TARGET" ? "No target" : r.status === "GAP" ? "Gap" : "Covered",
      ]),
    }],
  };
}

export async function monthlyReport(who: string): Promise<Report> {
  const plant = await loadPlant();
  const a = await computeAnalytics(plant);
  return {
    title: "Monthly Skills & Training Report",
    subtitle: `${stamp(who)} · For plant leadership`,
    notes: [
      "All figures are counted from qualification and training records in this system. No productivity or financial savings are claimed.",
      "Training impact can only be reported where an actual baseline and post-training measurement have been recorded.",
    ],
    sheets: [
      {
        name: "Coverage by department",
        columns: [{ header: "Department", width: 24 }, { header: "Assigned employees", width: 12 }, { header: "Assigned & qualified", width: 12 }, { header: "Qualified (any shift)", width: 12 }, { header: "Open required-skill gaps", width: 14 }],
        rows: a.byArea.map((r) => [r.name, r.assigned, r.assignedQualified, r.qualified, r.gaps]),
      },
      {
        name: "Trend",
        columns: [{ header: "Month", width: 10 }, { header: "Qualifications verified", width: 14 }, { header: "Training completed", width: 14 }, { header: "Active qualifications (month end)", width: 18 }],
        rows: a.trend.map((t) => [t.month, t.qualificationsVerified, t.trainingCompleted, t.activeQualifications]),
      },
      {
        name: "Single-person dependencies",
        columns: [{ header: "Department", width: 22 }, { header: "Competency", width: 36 }, { header: "Only qualified person", width: 24 }, { header: "Shift", width: 6 }],
        rows: [...a.single.map((s) => [s.area, s.name, s.holder, s.shift]), ...a.none.map((n) => [n.area, n.name, "NO ONE QUALIFIED", ""])],
      },
      {
        name: "Cross-training",
        columns: [{ header: "Production areas qualified", width: 22 }, { header: "Employees", width: 12 }],
        rows: [["0", a.cross.dist[0]], ["1", a.cross.dist[1]], ["2", a.cross.dist[2]], ["3 or more", a.cross.dist[3]]],
      },
    ],
  };
}
