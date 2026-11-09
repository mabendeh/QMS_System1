import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { PageHeader, Alert } from "@/components/ui";
import { RulesClient } from "./RulesClient";

export const metadata = { title: "Career levels & promotion rules" };

export default async function RulesPage() {
  await pageUser("rules:manage");
  const [levels, sets, areas, users] = await Promise.all([
    prisma.careerLevel.findMany({ orderBy: { rank: "asc" } }),
    prisma.promotionRuleSet.findMany({ include: { requirements: { orderBy: { sortOrder: "asc" } }, _count: { select: { requests: true } } }, orderBy: [{ version: "desc" }] }),
    prisma.productionArea.findMany({ where: { isActive: true }, orderBy: [{ kind: "asc" }, { sortOrder: "asc" }], include: { competencies: { where: { isActive: true }, orderBy: { sortOrder: "asc" } } } }),
    prisma.user.findMany({ select: { id: true, name: true } }),
  ]);
  const names = new Map(users.map((x) => [x.id, x.name]));
  return (
    <div className="grid gap-5">
      <PageHeader title="Career levels & promotion rules" subtitle="HR only. Rules are version-controlled: publishing creates a new version; earlier versions stay attached to requests that used them." />
      <Alert tone="warn">The starting levels and rules are demonstration assumptions, not official CHG pay grades or policy. Confirm them with HR leadership before use.</Alert>
      <RulesClient
        levels={levels.map((l) => ({ id: l.id, rank: l.rank, name: l.name, description: l.description ?? "", isActive: l.isActive }))}
        sets={sets.map((s) => ({ id: s.id, careerLevelId: s.careerLevelId, version: s.version, isActive: s.isActive, notes: s.notes ?? "", createdAt: s.createdAt.toISOString(), createdBy: s.createdById ? names.get(s.createdById) ?? "" : "seed", used: s._count.requests, requirements: s.requirements.map((r) => ({ type: r.type, label: r.label, areaId: r.areaId, competencyId: r.competencyId, minCount: r.minCount })) }))}
        areas={areas.map((a) => ({ id: a.id, name: a.name, kind: a.kind, competencies: a.competencies.map((c) => ({ id: c.id, name: c.name })) }))}
      />
    </div>
  );
}
