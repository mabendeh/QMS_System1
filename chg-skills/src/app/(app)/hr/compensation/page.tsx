import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { Alert, PageHeader } from "@/components/ui";
import { CompensationClient } from "./CompensationClient";

export const metadata = { title: "HR Compensation" };

export default async function CompensationPage() {
  const u = await pageUser("comp:view");
  const [grades, premiums, levels, areas] = await Promise.all([
    prisma.payGrade.findMany({ include: { careerLevel: true }, orderBy: { code: "asc" } }),
    prisma.skillPremium.findMany({ orderBy: { name: "asc" } }),
    prisma.careerLevel.findMany({ orderBy: { rank: "asc" } }),
    prisma.productionArea.findMany({ where: { isActive: true }, orderBy: { sortOrder: "asc" }, select: { id: true, name: true } }),
  ]);
  const ser = (d: Date | null) => (d ? d.toISOString().slice(0, 10) : "");
  return (
    <div className="grid gap-5">
      <PageHeader title="Compensation configuration" subtitle="Restricted: HR administrators and plant leadership only. Every change is audited." />
      <Alert tone="warn" title="No amounts are pre-filled">
        Enter only pay bands, grades and premiums that are approved under CHG compensation policy. The system never applies pay changes automatically.
      </Alert>
      <CompensationClient
        canManage={u.permissions.includes("comp:manage")}
        levels={levels.map((l) => ({ id: l.id, name: `L${l.rank} ${l.name}` }))}
        areas={areas}
        grades={grades.map((g) => ({ id: g.id, code: g.code, name: g.name, careerLevelId: g.careerLevelId ?? "", level: g.careerLevel ? `L${g.careerLevel.rank} ${g.careerLevel.name}` : "—", minRate: g.minRate?.toString() ?? "", maxRate: g.maxRate?.toString() ?? "", effectiveDate: ser(g.effectiveDate), endDate: ser(g.endDate), approvalAuthority: g.approvalAuthority ?? "", notes: g.notes ?? "", isActive: g.isActive }))}
        premiums={premiums.map((p) => ({ id: p.id, name: p.name, areaId: p.areaId ?? "", amount: p.amount?.toString() ?? "", unit: p.unit, effectiveDate: ser(p.effectiveDate), endDate: ser(p.endDate), approvalAuthority: p.approvalAuthority ?? "", notes: p.notes ?? "", isActive: p.isActive }))}
      />
    </div>
  );
}
