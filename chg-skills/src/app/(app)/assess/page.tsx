import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { PageHeader } from "@/components/ui";
import { QuickAssess } from "./QuickAssess";

export const metadata = { title: "Quick Assessment" };

export default async function AssessPage({ searchParams }: { searchParams: Promise<{ employee?: string }> }) {
  const u = await pageUser("skill:assess");
  const { employee } = await searchParams;
  const isManager = u.roles.includes("PRODUCTION_MANAGER");
  const trainerAreas = u.roles.includes("TRAINER")
    ? (await prisma.trainerAuthorization.findMany({ where: { userId: u.id, OR: [{ validUntil: null }, { validUntil: { gt: new Date() } }] } })).map((t) => t.areaId)
    : [];
  const employees = await prisma.employee.findMany({
    where: {
      status: "ACTIVE",
      ...(isManager || trainerAreas.length ? {} : { supervisorUserId: u.id }),
    },
    orderBy: [{ lastName: "asc" }, { firstName: "asc" }],
    select: { id: true, firstName: true, lastName: true, employeeNumber: true, isSample: true, supervisorUserId: true, primaryAreaId: true, shift: { select: { code: true } }, jobRole: { select: { title: true } } },
  });
  const areas = await prisma.productionArea.findMany({
    where: { isActive: true },
    orderBy: [{ kind: "asc" }, { sortOrder: "asc" }],
    select: { id: true, name: true, competencies: { where: { isActive: true }, orderBy: { sortOrder: "asc" }, select: { id: true, name: true, isCritical: true } } },
  });
  const allowedArea = (empSup: string | null, areaId: string) =>
    isManager || (u.roles.includes("SUPERVISOR") && empSup === u.id) || trainerAreas.includes(areaId);
  return (
    <div className="mx-auto max-w-2xl">
      <PageHeader title="Quick assessment" subtitle="Find the person, pick the skill, choose the level. Takes under a minute." />
      <QuickAssess
        initialEmployee={employee ?? null}
        employees={employees.map((e) => ({
          id: e.id, name: `${e.firstName} ${e.lastName}`, number: e.employeeNumber, isSample: e.isSample, shift: e.shift.code, job: e.jobRole?.title ?? "",
          primaryAreaId: e.primaryAreaId, areaIds: areas.filter((a) => allowedArea(e.supervisorUserId, a.id)).map((a) => a.id),
        })).filter((e) => e.areaIds.length > 0)}
        areas={areas}
      />
    </div>
  );
}
