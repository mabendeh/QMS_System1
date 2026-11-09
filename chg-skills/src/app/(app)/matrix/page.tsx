import { pageUser } from "@/lib/auth";
import { loadPlant } from "@/lib/data/plant";
import { MatrixView, type MatrixPayload } from "@/components/matrix/MatrixView";

export const metadata = { title: "Skills Matrix" };

export default async function MatrixPage({ searchParams }: { searchParams: Promise<Record<string, string | undefined>> }) {
  const u = await pageUser("matrix:view");
  const sp = await searchParams;
  const plant = await loadPlant();
  const view = plant.areas.find((a) => a.code === sp.area) ?? null;

  const payload: MatrixPayload = {
    view: view ? { id: view.id, code: view.code, name: view.name, description: view.description } : null,
    areas: plant.areas.map((a) => ({ id: a.id, code: a.code, name: a.name, kind: a.kind, count: a.competencies.length })),
    columns: view
      ? view.competencies.map((c) => ({ id: c.id, code: c.code, name: c.name, critical: c.isCritical }))
      : plant.areas.map((a) => ({ id: a.id, code: a.code, name: a.name, critical: false })),
    employees: plant.employees.map((e) => {
      const s = plant.employeeSummary(e.id);
      const next = plant.nextLevel(e.id);
      return {
        id: e.id, number: e.employeeNumber, name: `${e.firstName} ${e.lastName}`, lastName: e.lastName, isSample: e.isSample,
        shift: e.shift.code, area: e.primaryArea?.name ?? "—", areaId: e.primaryAreaId, job: e.jobRole?.title ?? "—",
        supervisorId: e.supervisorUserId, supervisor: e.supervisor?.name ?? "—",
        level: e.careerLevel?.rank ?? null, levelName: e.careerLevel?.name ?? "—",
        eligible: !!next?.result?.readyForReview, nextLevel: next?.level.name ?? null,
        fullyQualified: s.fullyQualified, crossTrained: s.crossTrained,
        cells: view
          ? view.competencies.map((c) => {
              const r = plant.rec(e.id, c.id);
              return { s: plant.display(e.id, c.id), p: !!r?.pendingStatus, v: r?.version ?? 0 };
            })
          : plant.areas.map((a) => {
              const c = plant.areaCell(e.id, a.id);
              return { s: c.status, p: c.pending, v: 0, n: `${c.competent}/${c.total} competent` };
            }),
      };
    }),
    canExport: u.permissions.includes("export:data"),
    focus: sp.focus ?? null,
  };

  return <MatrixView data={payload} />;
}
