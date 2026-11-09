"use client";
import { useEffect, useMemo, useState, useTransition } from "react";
import { getCellDetail, submitAssessment, type ActionResult } from "@/app/actions/skills";
import { SELECTABLE_STATUSES, STATUS_META, STORED_LABEL, type StoredStatus } from "@/lib/domain/status";
import { StatusChip } from "@/components/StatusCell";
import { Alert, Button, Card, Field, Input, SampleBadge, Select, Textarea } from "@/components/ui";
import { cn } from "@/lib/utils";

type Emp = { id: string; name: string; number: string; isSample: boolean; shift: string; job: string; primaryAreaId: string | null; areaIds: string[] };
type Area = { id: string; name: string; competencies: { id: string; name: string; isCritical: boolean }[] };
type Detail = Awaited<ReturnType<typeof getCellDetail>>;

export function QuickAssess({ employees, areas, initialEmployee }: { employees: Emp[]; areas: Area[]; initialEmployee: string | null }) {
  const [q, setQ] = useState("");
  const [emp, setEmp] = useState<Emp | null>(employees.find((e) => e.id === initialEmployee) ?? null);
  const [areaId, setAreaId] = useState("");
  const [compId, setCompId] = useState("");
  const [detail, setDetail] = useState<Detail | null>(null);
  const [status, setStatus] = useState<StoredStatus | null>(null);
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();

  const matches = useMemo(() => {
    const n = q.trim().toLowerCase();
    if (!n) return employees.slice(0, 8);
    return employees.filter((e) => e.name.toLowerCase().includes(n) || e.number.toLowerCase().includes(n)).slice(0, 12);
  }, [q, employees]);
  const myAreas = areas.filter((a) => emp?.areaIds.includes(a.id));
  const area = areas.find((a) => a.id === areaId);

  useEffect(() => {
    if (emp && !areaId && emp.primaryAreaId && emp.areaIds.includes(emp.primaryAreaId)) setAreaId(emp.primaryAreaId);
  }, [emp, areaId]);
  useEffect(() => {
    setDetail(null); setStatus(null); setRes(null);
    if (emp && compId) getCellDetail(emp.id, compId).then(setDetail).catch(() => setRes({ ok: false, error: "Could not load this skill." }));
  }, [emp, compId]);

  const reset = () => { setCompId(""); setStatus(null); setDetail(null); };

  return (
    <div className="grid gap-4">
      <Card className="p-4">
        <p className="mb-2 text-sm font-bold">1. Employee</p>
        {emp ? (
          <div className="flex items-center justify-between gap-3">
            <div>
              <p className="text-lg font-bold">{emp.name} {emp.isSample && <SampleBadge />}</p>
              <p className="text-sm text-muted">{emp.number} · Shift {emp.shift} · {emp.job}</p>
            </div>
            <Button variant="secondary" onClick={() => { setEmp(null); setAreaId(""); reset(); }}>Change</Button>
          </div>
        ) : (
          <>
            <Input autoFocus placeholder="Type a name or employee ID" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search employee" />
            <ul className="mt-2 divide-y divide-line rounded-md border border-line">
              {matches.map((e) => (
                <li key={e.id}>
                  <button className="flex min-h-[52px] w-full items-center justify-between px-3 py-2 text-left hover:bg-canvas" onClick={() => setEmp(e)}>
                    <span><span className="font-semibold">{e.name}</span> {e.isSample && <SampleBadge />}<br /><span className="text-sm text-muted">{e.number} · Shift {e.shift} · {e.job}</span></span>
                    <span className="text-navy-700">Select</span>
                  </button>
                </li>
              ))}
              {matches.length === 0 && <li className="px-3 py-4 text-sm text-muted">No one you can assess matches “{q}”.</li>}
            </ul>
          </>
        )}
      </Card>

      {emp && (
        <Card className="grid gap-4 p-4">
          <Field label="2. Production area" htmlFor="qa-area">
            <Select id="qa-area" value={areaId} onChange={(e) => { setAreaId(e.target.value); reset(); }}>
              <option value="">Choose an area…</option>
              {myAreas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </Select>
          </Field>
          {area && (
            <Field label="3. Competency" htmlFor="qa-comp">
              <Select id="qa-comp" value={compId} onChange={(e) => setCompId(e.target.value)}>
                <option value="">Choose a competency…</option>
                {area.competencies.map((c) => <option key={c.id} value={c.id}>{c.name}{c.isCritical ? " (critical)" : ""}</option>)}
              </Select>
            </Field>
          )}
        </Card>
      )}

      {detail && (
        <Card className="p-4">
          <form
            className="grid gap-4"
            onSubmit={(e) => {
              e.preventDefault();
              const f = e.currentTarget;
              start(async () => {
                const r = await submitAssessment(new FormData(f));
                setRes(r);
                if (r.ok || r.conflict) { setDetail(await getCellDetail(emp!.id, compId)); setStatus(null); f.reset(); }
              });
            }}
          >
            <div className="flex items-center gap-3 rounded-md bg-canvas p-3">
              <StatusChip status={detail.display} pending={!!detail.record?.pendingStatus} />
              <p className="text-sm">Current: <b>{STATUS_META[detail.display].label}</b>{detail.record?.pendingStatus ? " · verification pending" : ""}</p>
            </div>
            <input type="hidden" name="employeeId" value={detail.employee.id} />
            <input type="hidden" name="competencyId" value={detail.competency.id} />
            <input type="hidden" name="expectedVersion" value={detail.record?.version ?? 0} />
            <fieldset>
              <legend className="mb-2 text-sm font-bold">4. Rating</legend>
              <div className="grid gap-2">
                {SELECTABLE_STATUSES.map((s) => {
                  const m = STATUS_META[s === "NOT_APPLICABLE" ? "NOT_ASSESSED" : s];
                  return (
                    <label key={s} className={cn("flex min-h-[56px] cursor-pointer items-center gap-3 rounded-lg border-2 px-4 text-base font-semibold", status === s ? "border-navy-800 bg-navy-50" : "border-line")}>
                      <input type="radio" name="status" value={s} required checked={status === s} onChange={() => setStatus(s)} className="sr-only" />
                      <span className="inline-flex h-8 w-11 items-center justify-center rounded text-sm font-bold" style={{ background: m.color, color: m.text }} aria-hidden>{m.short}</span>
                      {STORED_LABEL[s]}
                    </label>
                  );
                })}
              </div>
            </fieldset>
            {status === "TRAINED" && <Field label="Training completion date" htmlFor="qa-date"><Input id="qa-date" type="date" name="trainingDate" required max={new Date().toISOString().slice(0, 10)} /></Field>}
            {status === "COMPETENT" && detail.competency.isCritical && (
              <Field label="Evidence (required for critical skills)" htmlFor="qa-ev"><Input id="qa-ev" type="file" name="evidence" accept="application/pdf,image/png,image/jpeg,image/webp" className="py-2" /></Field>
            )}
            <Field label="5. Comments (optional)" htmlFor="qa-notes"><Textarea id="qa-notes" name="notes" /></Field>
            {res && (
              <Alert tone={res.ok ? (res.outcome === "PENDING" ? "warn" : "success") : res.conflict ? "warn" : "error"}>
                <p className="font-semibold">{res.ok ? res.message : res.error}</p>
                {!res.ok && res.reasons && <ul className="mt-1 list-disc pl-5">{res.reasons.map((r) => <li key={r}>{r}</li>)}</ul>}
              </Alert>
            )}
            <Button type="submit" size="lg" disabled={pending || !status}>{pending ? "Saving…" : "6. Submit assessment"}</Button>
          </form>
        </Card>
      )}
      {res?.ok && <Button variant="secondary" onClick={() => { reset(); setRes(null); }}>Assess another skill for {emp?.name}</Button>}
    </div>
  );
}
