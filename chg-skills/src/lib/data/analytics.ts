import "server-only";
import { prisma } from "@/lib/db";
import { computeCoverage } from "@/lib/domain/coverage";
import type { Plant } from "./plant";

const monthKey = (d: Date) => d.toLocaleDateString("en-CA", { timeZone: "America/Chicago" }).slice(0, 7);

export async function coverageFor(plant: Plant) {
  const [shifts, targets] = await Promise.all([
    prisma.shift.findMany({ where: { isActive: true }, orderBy: { code: "asc" } }),
    prisma.shiftCoverageTarget.findMany(),
  ]);
  const prodAreas = plant.areas.filter((a) => a.kind === "PRODUCTION");
  const rows = computeCoverage({
    areas: prodAreas.map((a) => ({ id: a.id, name: a.name })),
    shifts: shifts.map((s) => ({ id: s.id, code: s.code })),
    employees: plant.employees.map((e) => ({ shiftId: e.shiftId, qualifiedAreaIds: plant.qualifiedAreaIds(e.id) })),
    targets,
  });
  return { shifts, rows };
}

/** All figures come from stored qualification and training records. Nothing is estimated. */
export async function computeAnalytics(plant: Plant) {
  const now = plant.now;
  const prodAreas = plant.areas.filter((a) => a.kind === "PRODUCTION");
  const empIds = plant.employees.map((e) => e.id);

  // Qualification coverage by department
  const byArea = prodAreas.map((a) => {
    const assigned = plant.employees.filter((e) => e.primaryAreaId === a.id);
    const qualified = plant.employees.filter((e) => plant.qualifiedAreaIds(e.id).has(a.id));
    const assignedQualified = assigned.filter((e) => plant.qualifiedAreaIds(e.id).has(a.id)).length;
    let gaps = 0;
    for (const e of plant.employees) for (const c of a.competencies) if (plant.required(e.id, c.id) && plant.display(e.id, c.id) !== "COMPETENT") gaps++;
    return { id: a.id, name: a.name, assigned: assigned.length, assignedQualified, qualified: qualified.length, gaps };
  });

  // Cross-training distribution
  const dist = [0, 0, 0, 0];
  for (const e of plant.employees) dist[Math.min(3, plant.employeeSummary(e.id).productionAreasQualified)]++;

  // Single-person dependencies and zero coverage
  const competentCount = new Map<string, number>();
  for (const e of plant.employees) for (const c of plant.competentIds(e.id)) competentCount.set(c, (competentCount.get(c) ?? 0) + 1);
  const single = plant.allComps.filter((c) => (competentCount.get(c.id) ?? 0) === 1).map((c) => {
    const holder = plant.employees.find((e) => plant.competentIds(e.id).has(c.id));
    return { id: c.id, name: c.name, area: c.areaName, holder: holder ? `${holder.firstName} ${holder.lastName}` : "", shift: holder?.shift.code ?? "" };
  });
  const none = plant.allComps.filter((c) => (competentCount.get(c.id) ?? 0) === 0 && plant.employees.some((e) => plant.required(e.id, c.id)))
    .map((c) => ({ id: c.id, name: c.name, area: c.areaName }));

  // Trends (last 12 months)
  const since = new Date(now);
  since.setMonth(since.getMonth() - 11, 1);
  since.setHours(0, 0, 0, 0);
  const months: string[] = [];
  for (let i = 0; i < 12; i++) { const d = new Date(since); d.setMonth(since.getMonth() + i); months.push(monthKey(d)); }
  const [verified, completedTraining, evals, completedAssign, promos] = await Promise.all([
    prisma.skillAssessment.findMany({ where: { employeeId: { in: empIds }, requestedStatus: "COMPETENT", state: { in: ["VERIFIED", "APPLIED"] }, createdAt: { gte: since } }, select: { createdAt: true, decidedAt: true } }),
    prisma.trainingAssignment.findMany({ where: { employeeId: { in: empIds }, status: "COMPLETED", completedDate: { gte: since } }, select: { completedDate: true } }),
    prisma.practicalEvaluation.findMany({ where: { employeeId: { in: empIds } }, select: { employeeId: true, competencyId: true, result: true, evaluatedAt: true }, orderBy: { evaluatedAt: "asc" } }),
    prisma.trainingAssignment.findMany({ where: { employeeId: { in: empIds }, status: "COMPLETED", startDate: { not: null }, completedDate: { not: null } }, select: { startDate: true, completedDate: true } }),
    prisma.promotionRequest.groupBy({ by: ["status"], _count: true }),
  ]);
  const count = (dates: (Date | null)[]) => {
    const m = new Map<string, number>();
    for (const d of dates) if (d) m.set(monthKey(d), (m.get(monthKey(d)) ?? 0) + 1);
    return m;
  };
  const vq = count(verified.map((v) => v.decidedAt ?? v.createdAt));
  const tc = count(completedTraining.map((t) => t.completedDate));
  const allRecs = [...plant.skillMap.values()].flatMap((m) => [...m.values()]).filter((r) => r.status === "COMPETENT" && r.qualifiedAt);
  const trend = months.map((m) => {
    const [y, mo] = m.split("-").map(Number);
    const end = new Date(Date.UTC(y, mo, 0, 23, 59));
    const active = allRecs.filter((r) => r.qualifiedAt! <= end && (!r.expiresAt || r.expiresAt > end)).length;
    return { month: m, qualificationsVerified: vq.get(m) ?? 0, trainingCompleted: tc.get(m) ?? 0, activeQualifications: active };
  });

  // Training effectiveness – only from real records
  const firstAttempt = new Map<string, string>();
  for (const e of evals) { const k = `${e.employeeId}:${e.competencyId}`; if (!firstAttempt.has(k)) firstAttempt.set(k, e.result); }
  const fa = [...firstAttempt.values()];
  const durations = completedAssign.map((a) => (a.completedDate!.getTime() - a.startDate!.getTime()) / 86_400_000).filter((d) => d >= 0);

  // Department priorities: coverage gaps first, then open required-skill gaps
  const cov = await coverageFor(plant);
  const priorities = byArea
    .map((a) => {
      const row = cov.rows.find((r) => r.areaId === a.id);
      const shiftGaps = row ? row.cells.filter((c) => c.status === "GAP").length : 0;
      return { ...a, shiftGaps, single: single.filter((s) => s.area === a.name).length };
    })
    .filter((a) => a.shiftGaps || a.gaps || a.single)
    .sort((x, y) => y.shiftGaps - x.shiftGaps || y.single - x.single || y.gaps - x.gaps);

  const readyNotSubmitted = plant.employees.filter((e) => plant.nextLevel(e.id)?.result?.readyForReview).length;

  return {
    byArea,
    cross: { dist, crossTrained: dist[2] + dist[3], total: plant.employees.length },
    single, none, trend, priorities, coverage: cov,
    effectiveness: {
      evaluations: fa.length,
      firstAttemptPassRate: fa.length ? Math.round((fa.filter((r) => r === "PASS").length / fa.length) * 100) : null,
      completedWithDates: durations.length,
      avgDaysToComplete: durations.length ? Math.round(durations.reduce((a, b) => a + b, 0) / durations.length) : null,
    },
    pipeline: { byStatus: Object.fromEntries(promos.map((p) => [p.status, p._count])), readyForReview: readyNotSubmitted },
  };
}
