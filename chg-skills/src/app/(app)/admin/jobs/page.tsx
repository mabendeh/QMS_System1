import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { saveJobRole } from "@/app/actions/admin";
import { ActionForm } from "@/components/ActionForm";
import { Badge, Button, Card, CardHeader, Input, PageHeader } from "@/components/ui";

export const metadata = { title: "Job titles" };

export default async function JobsAdmin() {
  await pageUser("config:manage");
  const [jobs, areas] = await Promise.all([
    prisma.jobRole.findMany({ orderBy: { title: "asc" }, include: { requirements: { include: { competency: { select: { areaId: true } } } }, _count: { select: { employees: true } } } }),
    prisma.productionArea.findMany({ where: { isActive: true }, orderBy: [{ kind: "asc" }, { sortOrder: "asc" }] }),
  ]);
  const form = (j?: (typeof jobs)[number]) => {
    const req = new Set(j?.requirements.map((r) => r.competency.areaId));
    return (
      <ActionForm action={saveJobRole} className="grid gap-3 p-5" resetOnSuccess={!j}>
        <input type="hidden" name="id" value={j?.id ?? ""} />
        <Input name="title" defaultValue={j?.title} placeholder="Job title" aria-label="Job title" required />
        <fieldset>
          <legend className="mb-2 text-sm font-semibold">Must be qualified in (every active competency of each selected department)</legend>
          <div className="grid gap-1 sm:grid-cols-2 lg:grid-cols-3">
            {areas.map((a) => (
              <label key={a.id} className="flex items-center gap-2 text-sm"><input type="checkbox" name="areaIds" value={a.id} defaultChecked={req.has(a.id)} className="h-4 w-4" /> {a.name}</label>
            ))}
          </div>
        </fieldset>
        <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={j?.isActive ?? true} className="h-5 w-5" /> Active</label>
        <div><Button type="submit" size="sm">{j ? "Save" : "Add job title"}</Button></div>
      </ActionForm>
    );
  };
  return (
    <div className="grid gap-5">
      <PageHeader title="Job titles" subtitle="Required skills drive the red “Needs Training” cells and training gap lists." />
      {jobs.map((j) => (
        <Card key={j.id}>
          <CardHeader title={j.title} subtitle={`${j._count.employees} employees · ${j.requirements.length} required competencies`} action={!j.isActive && <Badge>Inactive</Badge>} />
          {form(j)}
        </Card>
      ))}
      <Card><CardHeader title="New job title" />{form()}</Card>
    </div>
  );
}
