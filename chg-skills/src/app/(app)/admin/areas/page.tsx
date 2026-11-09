import Link from "next/link";
import { ArrowDown, ArrowUp } from "lucide-react";
import { pageUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { moveItem, saveArea, saveAreaExpiry, saveCompetency } from "@/app/actions/admin";
import { ActionForm } from "@/components/ActionForm";
import { Badge, Button, Card, CardHeader, Field, Input, PageHeader, Select, Textarea } from "@/components/ui";
import { cn } from "@/lib/utils";

export const metadata = { title: "Departments & competencies" };

export default async function AreasAdmin({ searchParams }: { searchParams: Promise<{ area?: string }> }) {
  await pageUser("config:manage");
  const { area: sel } = await searchParams;
  const areas = await prisma.productionArea.findMany({
    orderBy: [{ sortOrder: "asc" }],
    include: { competencies: { orderBy: { sortOrder: "asc" }, include: { expirationRules: true } }, expirationRules: { where: { competencyId: null } }, _count: { select: { employees: true } } },
  });
  const area = areas.find((a) => a.id === sel) ?? areas[0];
  return (
    <div className="grid gap-5">
      <PageHeader title="Departments & competencies" subtitle="Starting checklists are assumptions – verify them against the Duncanville plant layout and procedures." />
      <div className="grid gap-5 lg:grid-cols-[320px_1fr]">
        <Card className="h-fit">
          <CardHeader title="Departments" />
          <ul className="divide-y divide-line">
            {areas.map((a) => (
              <li key={a.id} className={cn("flex items-center gap-1 pr-2", a.id === area?.id && "bg-navy-50")}>
                <Link href={`/admin/areas?area=${a.id}`} className="flex-1 px-4 py-2.5 text-sm">
                  <span className={cn("font-semibold", !a.isActive && "text-muted line-through")}>{a.name}</span>
                  <span className="block text-xs text-muted">{a.competencies.filter((c) => c.isActive).length} competencies{a.kind === "PLANT_WIDE" ? " · plant-wide" : ""}{!a.isActive ? " · disabled" : ""}</span>
                </Link>
                <ActionForm action={moveItem} className="flex">
                  <input type="hidden" name="kind" value="area" /><input type="hidden" name="id" value={a.id} />
                  <button name="dir" value="up" className="rounded p-1 text-muted hover:bg-white" aria-label={`Move ${a.name} up`}><ArrowUp className="h-4 w-4" /></button>
                  <button name="dir" value="down" className="rounded p-1 text-muted hover:bg-white" aria-label={`Move ${a.name} down`}><ArrowDown className="h-4 w-4" /></button>
                </ActionForm>
              </li>
            ))}
          </ul>
          <div className="border-t border-line p-4">
            <p className="mb-2 text-sm font-bold">Add department</p>
            <ActionForm action={saveArea} className="grid gap-2" resetOnSuccess>
              <Input name="name" placeholder="Department name" aria-label="New department name" required />
              <Select name="kind" aria-label="Type" defaultValue="PRODUCTION"><option value="PRODUCTION">Production department</option><option value="PLANT_WIDE">Plant-wide training</option></Select>
              <input type="hidden" name="isActive" value="on" />
              <Button type="submit" size="sm">Add</Button>
            </ActionForm>
          </div>
        </Card>

        {area && (
          <div className="grid content-start gap-5">
            <Card>
              <CardHeader title={area.name} subtitle={`Code ${area.code} · ${area._count.employees} employees assigned`} action={area.isActive ? <Badge tone="blue">Active</Badge> : <Badge>Disabled</Badge>} />
              <ActionForm action={saveArea} className="grid gap-4 p-5 md:grid-cols-2">
                <input type="hidden" name="id" value={area.id} />
                <Field label="Name" htmlFor="a-name"><Input id="a-name" name="name" defaultValue={area.name} required /></Field>
                <Field label="Type" htmlFor="a-kind"><Select id="a-kind" name="kind" defaultValue={area.kind}><option value="PRODUCTION">Production department</option><option value="PLANT_WIDE">Plant-wide training</option></Select></Field>
                <div className="md:col-span-2"><Field label="Description / operating notes" htmlFor="a-desc"><Textarea id="a-desc" name="description" defaultValue={area.description ?? ""} /></Field></div>
                <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={area.isActive} className="h-5 w-5" /> Active (disabled departments are hidden from the matrix; history is kept)</label>
                <div className="md:col-span-2"><Button type="submit">Save department</Button></div>
              </ActionForm>
              <ActionForm action={saveAreaExpiry} className="flex flex-wrap items-end gap-3 border-t border-line p-5">
                <input type="hidden" name="areaId" value={area.id} />
                <Field label="Requalify every (months)" htmlFor="a-exp" hint="Blank = does not expire. A competency-level setting overrides this."><Input id="a-exp" name="months" type="number" min={1} max={120} defaultValue={area.expirationRules[0]?.months ?? ""} className="w-40" /></Field>
                <Button type="submit" variant="secondary">Save interval</Button>
              </ActionForm>
            </Card>
            <Card>
              <CardHeader title="Competency checklist" subtitle="◆ Critical skills need a passed practical evaluation and uploaded evidence before Competent." />
              <ul className="divide-y divide-line">
                {area.competencies.map((c) => (
                  <li key={c.id} className="p-4">
                    <details>
                      <summary className="flex cursor-pointer items-center gap-3">
                        <span className={cn("flex-1 font-semibold", !c.isActive && "text-muted line-through")}>{c.isCritical && <span className="text-st-red">◆ </span>}{c.name}</span>
                        <span className="font-mono text-xs text-muted">{c.code}</span>
                        {c.expirationRules[0] && <Badge>{c.expirationRules[0].months} mo</Badge>}
                      </summary>
                      <ActionForm action={saveCompetency} className="mt-3 grid gap-3 md:grid-cols-2">
                        <input type="hidden" name="id" value={c.id} /><input type="hidden" name="areaId" value={area.id} />
                        <Field label="Name" htmlFor={`c-${c.id}`}><Input id={`c-${c.id}`} name="name" defaultValue={c.name} required /></Field>
                        <Field label="Requalify every (months)" htmlFor={`m-${c.id}`}><Input id={`m-${c.id}`} name="requalMonths" type="number" min={1} max={120} defaultValue={c.expirationRules[0]?.months ?? ""} /></Field>
                        <div className="md:col-span-2"><Field label="Standard / description" htmlFor={`d-${c.id}`}><Textarea id={`d-${c.id}`} name="description" defaultValue={c.description ?? ""} /></Field></div>
                        <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isCritical" defaultChecked={c.isCritical} className="h-5 w-5" /> Critical skill</label>
                        <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={c.isActive} className="h-5 w-5" /> Active</label>
                        <div className="flex gap-2 md:col-span-2"><Button type="submit" size="sm">Save</Button></div>
                      </ActionForm>
                      <ActionForm action={moveItem} className="mt-2 flex gap-2">
                        <input type="hidden" name="kind" value="competency" /><input type="hidden" name="id" value={c.id} />
                        <Button name="dir" value="up" size="sm" variant="ghost"><ArrowUp className="h-4 w-4" /> Move up</Button>
                        <Button name="dir" value="down" size="sm" variant="ghost"><ArrowDown className="h-4 w-4" /> Move down</Button>
                      </ActionForm>
                    </details>
                  </li>
                ))}
              </ul>
              <div className="border-t border-line p-5">
                <p className="mb-2 font-bold">Add competency</p>
                <ActionForm action={saveCompetency} className="grid gap-3 md:grid-cols-2" resetOnSuccess>
                  <input type="hidden" name="areaId" value={area.id} /><input type="hidden" name="isActive" value="on" />
                  <Field label="Name" htmlFor="new-c"><Input id="new-c" name="name" required /></Field>
                  <Field label="Requalify every (months)" htmlFor="new-m"><Input id="new-m" name="requalMonths" type="number" min={1} max={120} /></Field>
                  <div className="md:col-span-2"><Field label="Standard / description" htmlFor="new-d"><Textarea id="new-d" name="description" /></Field></div>
                  <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isCritical" className="h-5 w-5" /> Critical skill</label>
                  <div className="md:col-span-2"><Button type="submit">Add competency</Button></div>
                </ActionForm>
              </div>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}
