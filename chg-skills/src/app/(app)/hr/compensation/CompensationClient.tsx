"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { savePayGrade, saveSkillPremium } from "@/app/actions/compensation";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Badge, Button, Card, CardHeader, Field, Input, Select, Table, Td, Textarea, Th } from "@/components/ui";

type Grade = { id: string; code: string; name: string; careerLevelId: string; level: string; minRate: string; maxRate: string; effectiveDate: string; endDate: string; approvalAuthority: string; notes: string; isActive: boolean };
type Premium = { id: string; name: string; areaId: string; amount: string; unit: string; effectiveDate: string; endDate: string; approvalAuthority: string; notes: string; isActive: boolean };

export function CompensationClient({ canManage, levels, areas, grades, premiums }: { canManage: boolean; levels: { id: string; name: string }[]; areas: { id: string; name: string }[]; grades: Grade[]; premiums: Premium[] }) {
  const router = useRouter();
  const [edit, setEdit] = useState<{ kind: "grade"; g: Partial<Grade> } | { kind: "premium"; p: Partial<Premium> } | null>(null);
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  const save = (fn: (f: FormData) => Promise<ActionResult>, f: HTMLFormElement) => start(async () => { const r = await fn(new FormData(f)); setRes(r); if (r.ok) { setEdit(null); router.refresh(); } });
  const range = (g: Grade) => (g.minRate && g.maxRate ? `$${g.minRate} – $${g.maxRate}` : <span className="text-muted">Not configured</span>);
  return (
    <div className="grid gap-5">
      {res && <Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}</Alert>}
      <Card>
        <CardHeader title="Pay grades" subtitle="Linked to career levels" action={canManage && <Button size="sm" onClick={() => setEdit({ kind: "grade", g: { isActive: true } })}>Add pay grade</Button>} />
        <Table>
          <thead><tr><Th>Code</Th><Th>Name</Th><Th>Career level</Th><Th>Approved range (hourly)</Th><Th>Effective</Th><Th>Approval authority</Th><Th>Status</Th><Th /></tr></thead>
          <tbody>
            {grades.map((g) => (
              <tr key={g.id}>
                <Td className="font-mono">{g.code}</Td><Td>{g.name}</Td><Td>{g.level}</Td><Td>{range(g)}</Td>
                <Td>{g.effectiveDate || "—"}{g.endDate ? ` → ${g.endDate}` : ""}</Td><Td>{g.approvalAuthority || "—"}</Td>
                <Td>{g.isActive ? <Badge tone="blue">Active</Badge> : <Badge>Inactive</Badge>}</Td>
                <Td className="text-right">{canManage && <Button size="sm" variant="secondary" onClick={() => setEdit({ kind: "grade", g })}>Edit</Button>}</Td>
              </tr>
            ))}
          </tbody>
        </Table>
      </Card>
      <Card>
        <CardHeader title="Skill premiums" subtitle="Optional premiums for approved qualifications" action={canManage && <Button size="sm" onClick={() => setEdit({ kind: "premium", p: { isActive: true, unit: "per hour" } })}>Add premium</Button>} />
        {premiums.length === 0 ? <p className="px-5 py-6 text-sm text-muted">No skill premiums configured.</p> : (
          <Table>
            <thead><tr><Th>Name</Th><Th>Department</Th><Th>Amount</Th><Th>Effective</Th><Th>Approval authority</Th><Th>Status</Th><Th /></tr></thead>
            <tbody>
              {premiums.map((p) => (
                <tr key={p.id}>
                  <Td>{p.name}</Td><Td>{areas.find((a) => a.id === p.areaId)?.name ?? "Any"}</Td><Td>{p.amount ? `$${p.amount} ${p.unit}` : <span className="text-muted">Not configured</span>}</Td>
                  <Td>{p.effectiveDate || "—"}</Td><Td>{p.approvalAuthority || "—"}</Td><Td>{p.isActive ? <Badge tone="blue">Active</Badge> : <Badge>Inactive</Badge>}</Td>
                  <Td className="text-right">{canManage && <Button size="sm" variant="secondary" onClick={() => setEdit({ kind: "premium", p })}>Edit</Button>}</Td>
                </tr>
              ))}
            </tbody>
          </Table>
        )}
      </Card>

      {edit?.kind === "grade" && (
        <Card className="p-5">
          <form className="grid gap-4" onSubmit={(e) => { e.preventDefault(); save(savePayGrade, e.currentTarget); }}>
            <h3 className="text-lg font-bold">{edit.g.id ? `Edit ${edit.g.code}` : "New pay grade"}</h3>
            <input type="hidden" name="id" value={edit.g.id ?? ""} />
            <div className="grid gap-4 sm:grid-cols-3">
              <Field label="Code" htmlFor="g-code"><Input id="g-code" name="code" defaultValue={edit.g.code} required /></Field>
              <Field label="Name" htmlFor="g-name"><Input id="g-name" name="name" defaultValue={edit.g.name} required /></Field>
              <Field label="Career level" htmlFor="g-level"><Select id="g-level" name="careerLevelId" defaultValue={edit.g.careerLevelId}><option value="">None</option>{levels.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}</Select></Field>
              <Field label="Minimum hourly rate (USD)" htmlFor="g-min"><Input id="g-min" name="minRate" inputMode="decimal" defaultValue={edit.g.minRate} /></Field>
              <Field label="Maximum hourly rate (USD)" htmlFor="g-max"><Input id="g-max" name="maxRate" inputMode="decimal" defaultValue={edit.g.maxRate} /></Field>
              <Field label="Approval authority" htmlFor="g-auth"><Input id="g-auth" name="approvalAuthority" defaultValue={edit.g.approvalAuthority} placeholder="e.g. HR Director" /></Field>
              <Field label="Effective date" htmlFor="g-eff"><Input id="g-eff" type="date" name="effectiveDate" defaultValue={edit.g.effectiveDate} /></Field>
              <Field label="End date" htmlFor="g-end"><Input id="g-end" type="date" name="endDate" defaultValue={edit.g.endDate} /></Field>
              <label className="flex items-center gap-2 pt-7 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={edit.g.isActive} className="h-5 w-5" /> Active</label>
            </div>
            <Field label="Notes / policy reference" htmlFor="g-notes"><Textarea id="g-notes" name="notes" defaultValue={edit.g.notes} /></Field>
            <div className="flex gap-2"><Button type="submit" disabled={pending}>Save</Button><Button type="button" variant="secondary" onClick={() => setEdit(null)}>Cancel</Button></div>
          </form>
        </Card>
      )}
      {edit?.kind === "premium" && (
        <Card className="p-5">
          <form className="grid gap-4" onSubmit={(e) => { e.preventDefault(); save(saveSkillPremium, e.currentTarget); }}>
            <h3 className="text-lg font-bold">{edit.p.id ? `Edit ${edit.p.name}` : "New skill premium"}</h3>
            <input type="hidden" name="id" value={edit.p.id ?? ""} />
            <div className="grid gap-4 sm:grid-cols-3">
              <Field label="Name" htmlFor="p-name"><Input id="p-name" name="name" defaultValue={edit.p.name} required /></Field>
              <Field label="Department" htmlFor="p-area"><Select id="p-area" name="areaId" defaultValue={edit.p.areaId}><option value="">Any</option>{areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</Select></Field>
              <Field label="Amount (USD)" htmlFor="p-amt"><Input id="p-amt" name="amount" inputMode="decimal" defaultValue={edit.p.amount} /></Field>
              <Field label="Unit" htmlFor="p-unit"><Input id="p-unit" name="unit" defaultValue={edit.p.unit} /></Field>
              <Field label="Effective date" htmlFor="p-eff"><Input id="p-eff" type="date" name="effectiveDate" defaultValue={edit.p.effectiveDate} /></Field>
              <Field label="End date" htmlFor="p-end"><Input id="p-end" type="date" name="endDate" defaultValue={edit.p.endDate} /></Field>
              <Field label="Approval authority" htmlFor="p-auth"><Input id="p-auth" name="approvalAuthority" defaultValue={edit.p.approvalAuthority} /></Field>
              <label className="flex items-center gap-2 pt-7 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={edit.p.isActive} className="h-5 w-5" /> Active</label>
            </div>
            <Field label="Notes / policy reference" htmlFor="p-notes"><Textarea id="p-notes" name="notes" defaultValue={edit.p.notes} /></Field>
            <div className="flex gap-2"><Button type="submit" disabled={pending}>Save</Button><Button type="button" variant="secondary" onClick={() => setEdit(null)}>Cancel</Button></div>
          </form>
        </Card>
      )}
    </div>
  );
}
