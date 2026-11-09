import "server-only";
import { loadPlant, type Plant } from "@/lib/data/plant";
import { fmtDate, fmtDateTime } from "@/lib/utils";
import type { Cell, Report } from "./render";

export interface MatrixFilters {
  area?: string | null;
  shift?: string | null;
  dept?: string | null;
  job?: string | null;
  sup?: string | null;
}

const stamp = (who: string) => `Generated ${fmtDateTime(new Date())} by ${who} · CHG Duncanville · CONFIDENTIAL – internal use only`;

function filterEmployees(plant: Plant, f: MatrixFilters) {
  return plant.employees.filter(
    (e) => (!f.shift || e.shift.code === f.shift) && (!f.dept || e.primaryAreaId === f.dept) && (!f.job || e.jobRole?.title === f.job) && (!f.sup || e.supervisorUserId === f.sup),
  );
}

export async function matrixReport(f: MatrixFilters, who: string): Promise<Report> {
  const plant = await loadPlant();
  const view = plant.areas.find((a) => a.code === f.area) ?? null;
  const emps = filterEmployees(plant, f);
  const base = [
    { header: "Employee ID", width: 12 }, { header: "Employee", width: 22 }, { header: "Shift", width: 6 },
    { header: "Department", width: 16 }, { header: "Job title", width: 20 },
  ];
  const sampleNote = emps.some((e) => e.isSample) ? ["Rows with IDs starting DEMO- are fictional sample data."] : [];
  if (view) {
    return {
      title: `Skills Matrix – ${view.name}`,
      subtitle: stamp(who),
      legend: true,
      notes: sampleNote,
      sheets: [{
        name: view.name,
        frozenCols: 2,
        columns: [...base, ...view.competencies.map((c) => ({ header: `${c.name}${c.isCritical ? " ◆" : ""}`, width: 5, rotate: true }))],
        rows: emps.map((e) => [
          e.employeeNumber, `${e.firstName} ${e.lastName}`, e.shift.code, e.primaryArea?.name ?? "", e.jobRole?.title ?? "",
          ...view.competencies.map((c): Cell => ({ status: plant.display(e.id, c.id), pending: !!plant.rec(e.id, c.id)?.pendingStatus })),
        ]),
      }],
    };
  }
  // Complete matrix: overview sheet + one sheet per department
  return {
    title: "Complete Skills Matrix",
    subtitle: stamp(who),
    legend: true,
    notes: [...sampleNote, "Overview colours: blue = every competency Competent; green = all at least Trained; yellow = in training; red = needs training; grey = not assessed / N/A."],
    sheets: [
      {
        name: "Overview",
        frozenCols: 2,
        columns: [...base, ...plant.areas.map((a) => ({ header: a.name, width: 6, rotate: true }))],
        rows: emps.map((e) => [
          e.employeeNumber, `${e.firstName} ${e.lastName}`, e.shift.code, e.primaryArea?.name ?? "", e.jobRole?.title ?? "",
          ...plant.areas.map((a): Cell => {
            const c = plant.areaCell(e.id, a.id);
            return { status: c.status, pending: c.pending, text: `${c.competent}/${c.total} competent` };
          }),
        ]),
      },
      ...plant.areas.map((a) => ({
        name: a.name.replace(/[\\/?*[\]:]/g, "-"),
        frozenCols: 2,
        columns: [...base, ...a.competencies.map((c) => ({ header: `${c.name}${c.isCritical ? " ◆" : ""}`, width: 5, rotate: true }))],
        rows: emps.map((e) => [
          e.employeeNumber, `${e.firstName} ${e.lastName}`, e.shift.code, e.primaryArea?.name ?? "", e.jobRole?.title ?? "",
          ...a.competencies.map((c): Cell => ({ status: plant.display(e.id, c.id), pending: !!plant.rec(e.id, c.id)?.pendingStatus })),
        ]),
      })),
    ],
  };
}

export async function employeeReport(employeeId: string, who: string): Promise<Report> {
  const plant = await loadPlant({ includeInactiveEmployees: true, employeeIds: [employeeId] });
  const e = plant.employees[0];
  if (!e) throw new Error("Employee not found");
  const rows: Cell[][] = [];
  for (const a of plant.areas) {
    for (const c of a.competencies) {
      const r = plant.rec(e.id, c.id);
      const s = plant.display(e.id, c.id);
      if (s === "NOT_ASSESSED" && !r) continue;
      rows.push([a.name, c.name, c.isCritical ? "Yes" : "", { status: s, pending: !!r?.pendingStatus }, fmtDate(r?.qualifiedAt), fmtDate(r?.expiresAt)]);
    }
  }
  return {
    title: `Qualification Report – ${e.firstName} ${e.lastName} (${e.employeeNumber})`,
    subtitle: `${e.jobRole?.title ?? ""} · Shift ${e.shift.code} · ${e.primaryArea?.name ?? ""} · ${stamp(who)}`,
    legend: true,
    notes: e.isSample ? ["Fictional sample employee."] : [],
    sheets: [{
      name: "Qualifications",
      columns: [{ header: "Department", width: 18 }, { header: "Competency", width: 36 }, { header: "Critical", width: 8 }, { header: "Status", width: 6, rotate: true }, { header: "Qualified", width: 12 }, { header: "Requalify by", width: 12 }],
      rows,
    }],
  };
}

export async function gapsReport(f: MatrixFilters, who: string): Promise<Report> {
  const plant = await loadPlant();
  const emps = filterEmployees(plant, f);
  const rows: Cell[][] = [];
  for (const e of emps) {
    for (const a of plant.areas) {
      for (const c of a.competencies) {
        if (!plant.required(e.id, c.id)) continue;
        const s = plant.display(e.id, c.id);
        if (s === "COMPETENT") continue;
        rows.push([e.employeeNumber, `${e.firstName} ${e.lastName}`, e.shift.code, e.jobRole?.title ?? "", a.name, c.name, c.isCritical ? "Yes" : "", { status: s, pending: !!plant.rec(e.id, c.id)?.pendingStatus }]);
      }
    }
  }
  return {
    title: "Training Gaps",
    subtitle: `Required competencies not yet Competent (or expired) · ${stamp(who)}`,
    legend: true,
    sheets: [{
      name: "Gaps",
      frozenCols: 2,
      columns: [{ header: "Employee ID", width: 12 }, { header: "Employee", width: 22 }, { header: "Shift", width: 6 }, { header: "Job title", width: 20 }, { header: "Department", width: 18 }, { header: "Competency", width: 34 }, { header: "Critical", width: 8 }, { header: "Status", width: 6, rotate: true }],
      rows,
    }],
  };
}
