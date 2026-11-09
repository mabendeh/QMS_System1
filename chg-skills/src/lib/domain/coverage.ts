/**
 * Shift coverage: how many currently-qualified people each area has per shift,
 * against the target set by managers. Scheduled/available headcount is only
 * reported when a scheduling integration provides it; otherwise "unknown".
 */
export interface CoverageInput {
  areas: { id: string; name: string }[];
  shifts: { id: string; code: string }[];
  /** For each active employee: shift and the areas they are currently qualified in. */
  employees: { shiftId: string; qualifiedAreaIds: Set<string> }[];
  targets: { areaId: string; shiftId: string; targetQualified: number }[];
}

export type CoverageStatus = "COVERED" | "GAP" | "NO_TARGET";

export interface CoverageCell {
  shiftId: string;
  qualified: number;
  target: number | null;
  status: CoverageStatus;
  scheduled: null; // null = unknown (no scheduling integration)
  available: null;
}

export interface CoverageRow {
  areaId: string;
  areaName: string;
  totalQualified: number;
  totalTarget: number | null;
  status: CoverageStatus;
  cells: CoverageCell[];
}

export function computeCoverage(input: CoverageInput): CoverageRow[] {
  const target = new Map(input.targets.map((t) => [`${t.areaId}:${t.shiftId}`, t.targetQualified]));
  return input.areas.map((a) => {
    const cells: CoverageCell[] = input.shifts.map((s) => {
      const qualified = input.employees.filter((e) => e.shiftId === s.id && e.qualifiedAreaIds.has(a.id)).length;
      const t = target.has(`${a.id}:${s.id}`) ? target.get(`${a.id}:${s.id}`)! : null;
      return { shiftId: s.id, qualified, target: t, status: t == null ? "NO_TARGET" : qualified >= t ? "COVERED" : "GAP", scheduled: null, available: null };
    });
    const totalQualified = cells.reduce((n, c) => n + c.qualified, 0);
    const withTarget = cells.filter((c) => c.target != null);
    const totalTarget = withTarget.length ? withTarget.reduce((n, c) => n + (c.target ?? 0), 0) : null;
    const status: CoverageStatus = withTarget.length === 0 ? "NO_TARGET" : cells.some((c) => c.status === "GAP") ? "GAP" : "COVERED";
    return { areaId: a.id, areaName: a.name, totalQualified, totalTarget, status, cells };
  });
}

/** Competencies where only one person in the plant is currently qualified. */
export function singlePersonDependencies(
  competencies: { id: string; name: string; areaName: string }[],
  competentCount: Map<string, number>,
): { id: string; name: string; areaName: string }[] {
  return competencies.filter((c) => (competentCount.get(c.id) ?? 0) === 1);
}
