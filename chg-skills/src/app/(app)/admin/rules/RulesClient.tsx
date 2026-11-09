"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { Plus, Trash2 } from "lucide-react";
import { saveCareerLevel, saveRuleVersion } from "@/app/actions/rules";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Badge, Button, Card, CardHeader, Field, Input, Select, Textarea } from "@/components/ui";
import { fmtDate } from "@/lib/utils";

type Req = { type: string; label: string; areaId: string | null; competencyId: string | null; minCount: number | null };
type Set = { id: string; careerLevelId: string; version: number; isActive: boolean; notes: string; createdAt: string; createdBy: string; used: number; requirements: Req[] };
type Area = { id: string; name: string; kind: string; competencies: { id: string; name: string }[] };
const TYPES: [string, string][] = [
  ["PRIMARY_AREA_QUALIFIED", "Qualified in assigned department"], ["AREA_QUALIFIED", "Qualified in a specific department"],
  ["MIN_PRODUCTION_AREAS_QUALIFIED", "Qualified in at least N production departments"], ["COMPETENCY_COMPETENT", "Specific competency is Competent"],
  ["CURRENT_LEVEL_AT_LEAST", "Currently at career level N or above"], ["MANUAL_VERIFICATION", "Manual confirmation during review"],
];

export function RulesClient({ levels, sets, areas }: { levels: { id: string; rank: number; name: string; description: string; isActive: boolean }[]; sets: Set[]; areas: Area[] }) {
  const router = useRouter();
  const [res, setRes] = useState<ActionResult | null>(null);
  const [editing, setEditing] = useState<{ levelId: string; reqs: Req[]; notes: string } | null>(null);
  const [pending, start] = useTransition();
  return (
    <div className="grid gap-5">
      {res && <Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}</Alert>}
      {levels.map((l) => {
        const active = sets.find((s) => s.careerLevelId === l.id && s.isActive);
        const history = sets.filter((s) => s.careerLevelId === l.id && !s.isActive);
        return (
          <Card key={l.id}>
            <CardHeader title={`Level ${l.rank} – ${l.name}`} subtitle={l.description} action={!l.isActive && <Badge>Inactive</Badge>} />
            <div className="grid gap-5 p-5 lg:grid-cols-2">
              <form className="grid content-start gap-3" onSubmit={(e) => { e.preventDefault(); const f = e.currentTarget; start(async () => { setRes(await saveCareerLevel(new FormData(f))); router.refresh(); }); }}>
                <input type="hidden" name="id" value={l.id} />
                <Field label="Level title" htmlFor={`n-${l.id}`}><Input id={`n-${l.id}`} name="name" defaultValue={l.name} /></Field>
                <Field label="Description" htmlFor={`d-${l.id}`}><Textarea id={`d-${l.id}`} name="description" defaultValue={l.description} /></Field>
                <label className="flex items-center gap-2 text-sm font-semibold"><input type="checkbox" name="isActive" defaultChecked={l.isActive} className="h-5 w-5" /> Active</label>
                <div><Button type="submit" size="sm" variant="secondary" disabled={pending}>Save title</Button></div>
              </form>
              <div>
                {l.rank === 0 ? <p className="text-sm text-muted">Starting level – no promotion rules.</p> : (
                  <>
                    <p className="text-sm font-bold">Requirements {active ? <>· v{active.version} <span className="font-normal text-muted">({active.createdBy}, {fmtDate(active.createdAt)})</span></> : "· none configured"}</p>
                    <ul className="mt-2 grid gap-1 text-sm">{active?.requirements.map((r, i) => <li key={i}>• {r.label}</li>)}</ul>
                    {history.length > 0 && <p className="mt-2 text-xs text-muted">Earlier versions: {history.map((h) => `v${h.version}${h.used ? ` (used by ${h.used})` : ""}`).join(", ")}</p>}
                    <Button size="sm" className="mt-3" onClick={() => setEditing({ levelId: l.id, reqs: active?.requirements.map((r) => ({ ...r })) ?? [], notes: "" })}>Edit rules (new version)</Button>
                  </>
                )}
              </div>
            </div>
            {editing?.levelId === l.id && (
              <div className="border-t border-line bg-canvas p-5">
                <h3 className="font-bold">New rule version for {l.name}</h3>
                <div className="mt-3 grid gap-3">
                  {editing.reqs.map((r, i) => {
                    const set = (patch: Partial<Req>) => setEditing({ ...editing, reqs: editing.reqs.map((x, k) => (k === i ? { ...x, ...patch } : x)) });
                    return (
                      <div key={i} className="grid gap-2 rounded-md border border-line bg-white p-3 md:grid-cols-[1.2fr_1fr_1.4fr_auto]">
                        <Select aria-label="Requirement type" value={r.type} onChange={(e) => set({ type: e.target.value })}>{TYPES.map(([v, t]) => <option key={v} value={v}>{t}</option>)}</Select>
                        {r.type === "AREA_QUALIFIED" ? (
                          <Select aria-label="Department" value={r.areaId ?? ""} onChange={(e) => set({ areaId: e.target.value })}><option value="">Department…</option>{areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</Select>
                        ) : r.type === "COMPETENCY_COMPETENT" ? (
                          <Select aria-label="Competency" value={r.competencyId ?? ""} onChange={(e) => set({ competencyId: e.target.value })}><option value="">Competency…</option>{areas.map((a) => <optgroup key={a.id} label={a.name}>{a.competencies.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</optgroup>)}</Select>
                        ) : r.type === "MIN_PRODUCTION_AREAS_QUALIFIED" || r.type === "CURRENT_LEVEL_AT_LEAST" ? (
                          <Input aria-label="Number" type="number" min={0} value={r.minCount ?? ""} onChange={(e) => set({ minCount: e.target.value === "" ? null : Number(e.target.value) })} />
                        ) : <span />}
                        <Input aria-label="Label shown to employees" value={r.label} onChange={(e) => set({ label: e.target.value })} placeholder="Label shown to employees" />
                        <Button type="button" variant="ghost" aria-label="Remove requirement" onClick={() => setEditing({ ...editing, reqs: editing.reqs.filter((_, k) => k !== i) })}><Trash2 className="h-4 w-4" /></Button>
                      </div>
                    );
                  })}
                  <div><Button type="button" variant="secondary" size="sm" onClick={() => setEditing({ ...editing, reqs: [...editing.reqs, { type: "AREA_QUALIFIED", label: "", areaId: null, competencyId: null, minCount: null }] })}><Plus className="h-4 w-4" /> Add requirement</Button></div>
                  <Field label="Change notes" htmlFor="rule-notes"><Textarea id="rule-notes" value={editing.notes} onChange={(e) => setEditing({ ...editing, notes: e.target.value })} placeholder="Why the rules changed and who approved it" /></Field>
                  <div className="flex gap-2">
                    <Button disabled={pending} onClick={() => start(async () => { const r = await saveRuleVersion(JSON.stringify({ careerLevelId: l.id, notes: editing.notes, requirements: editing.reqs })); setRes(r); if (r.ok) { setEditing(null); router.refresh(); } })}>Publish new version</Button>
                    <Button variant="secondary" onClick={() => setEditing(null)}>Cancel</Button>
                  </div>
                  <p className="text-xs text-muted">Time employed cannot be used as a requirement – only verified competencies, levels and manual review confirmations.</p>
                </div>
              </div>
            )}
          </Card>
        );
      })}
    </div>
  );
}
