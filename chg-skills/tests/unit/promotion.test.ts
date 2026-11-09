import { describe, expect, it } from "vitest";
import { evaluateRequirements, nextPromotionStatus, qualifiedAreas } from "@/lib/domain/promotion";
import { computeCoverage } from "@/lib/domain/coverage";

const lk = { areaName: (id: string) => ({ PW: "Pre-Weigh", MX: "Mixing", FS: "Food Safety" } as Record<string, string>)[id] ?? id, competencyName: (id: string) => id };

describe("qualified areas", () => {
  it("requires every competency in the area", () => {
    const by = new Map([["PW", ["a", "b"]], ["MX", ["c"]]]);
    expect([...qualifiedAreas(by, new Set(["a", "b"]))]).toEqual(["PW"]);
    expect([...qualifiedAreas(by, new Set(["a", "c"]))]).toEqual(["MX"]);
  });
});

describe("promotion evaluation", () => {
  const reqs = [
    { id: "1", type: "MIN_PRODUCTION_AREAS_QUALIFIED" as const, label: "2+ areas", minCount: 2 },
    { id: "2", type: "AREA_QUALIFIED" as const, label: "Food safety", areaId: "FS" },
    { id: "3", type: "MANUAL_VERIFICATION" as const, label: "Supervisor approval" },
  ];
  it("Scenario 3: Pre-Weigh + Mixing qualification makes the employee ready for review (not promoted)", () => {
    const r = evaluateRequirements(reqs, { primaryAreaId: "PW", currentLevelRank: 1, qualifiedAreaIds: new Set(["PW", "MX", "FS"]), productionAreaIds: new Set(["PW", "MX"]), competentCompetencyIds: new Set() }, lk);
    expect(r.readyForReview).toBe(true);
    expect(r.metCount).toBe(2);
    expect(r.total).toBe(3);
    expect(r.items.find((i) => i.manual)?.met).toBe(false);
  });
  it("plant-wide areas do not count toward cross-training and missing items are explained", () => {
    const r = evaluateRequirements(reqs, { primaryAreaId: "PW", currentLevelRank: 1, qualifiedAreaIds: new Set(["PW", "FS"]), productionAreaIds: new Set(["PW", "MX"]), competentCompetencyIds: new Set() }, lk);
    expect(r.readyForReview).toBe(false);
    expect(r.missing).toContain("1 of 2 production areas qualified");
  });
  it("does not use time employed", () => {
    const r = evaluateRequirements([{ id: "x", type: "CURRENT_LEVEL_AT_LEAST", label: "L2", minCount: 2 }], { primaryAreaId: null, currentLevelRank: 1, qualifiedAreaIds: new Set(), productionAreaIds: new Set(), competentCompetencyIds: new Set() }, lk);
    expect(r.readyForReview).toBe(false);
  });
});

describe("promotion workflow", () => {
  it("moves through manager → HR → final approval, and rejection is terminal", () => {
    expect(nextPromotionStatus("SUBMITTED", "APPROVED")).toBe("MANAGER_APPROVED");
    expect(nextPromotionStatus("MANAGER_APPROVED", "APPROVED")).toBe("HR_VERIFIED");
    expect(nextPromotionStatus("HR_VERIFIED", "APPROVED")).toBe("APPROVED");
    expect(nextPromotionStatus("HR_VERIFIED", "REJECTED")).toBe("REJECTED");
    expect(nextPromotionStatus("REJECTED", "APPROVED")).toBeNull();
    expect(nextPromotionStatus("APPROVED", "APPROVED")).toBeNull();
  });
});

describe("coverage", () => {
  it("counts qualified people per shift against targets and never estimates availability", () => {
    const rows = computeCoverage({
      areas: [{ id: "MX", name: "Mixing" }],
      shifts: [{ id: "A", code: "A" }, { id: "B", code: "B" }],
      employees: [{ shiftId: "A", qualifiedAreaIds: new Set(["MX"]) }, { shiftId: "A", qualifiedAreaIds: new Set(["MX"]) }, { shiftId: "B", qualifiedAreaIds: new Set() }],
      targets: [{ areaId: "MX", shiftId: "A", targetQualified: 2 }, { areaId: "MX", shiftId: "B", targetQualified: 1 }],
    });
    expect(rows[0].cells.map((c) => c.status)).toEqual(["COVERED", "GAP"]);
    expect(rows[0].status).toBe("GAP");
    expect(rows[0].cells[0].available).toBeNull();
  });
});
