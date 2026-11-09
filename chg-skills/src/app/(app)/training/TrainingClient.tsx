"use client";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useMemo, useState, useTransition } from "react";
import { advanceTraining, assignTraining } from "@/app/actions/training";
import type { ActionResult } from "@/app/actions/skills";
import { AssessmentDialog } from "@/components/matrix/AssessmentDialog";
import { Alert, Badge, Button, Empty, Field, Input, SampleBadge, Select, Table, Td, Textarea, Th } from "@/components/ui";
import { fmtDate, fmtDateTime } from "@/lib/utils";

type Row = { id: string; version: number; employeeId: string; employee: string; isSample: boolean; shift: string; competencyId: string; competency: string; area: string; status: string; twiStep: string; trainer: string; due: string | null; overdue: boolean; completed: string | null };
export interface TrainingData {
  tab: string;
  canManage: boolean;
  active: Row[];
  overdue: Row[];
  done: Row[];
  verify: { id: string; employeeId: string; employee: string; isSample: boolean; competencyId: string; competency: string; area: string; by: string; at: string }[];
  requal: { employeeId: string; employee: string; isSample: boolean; shift: string; competencyId: string; competency: string; area: string; expiresAt: string | null; expired: boolean; assigned: boolean }[];
  gaps: { employeeId: string; employee: string; isSample: boolean; shift: string; competencyId: string; competency: string; area: string; critical: boolean }[];
  gapsTotal: number;
  employees: { id: string; name: string; supervisorId: string | null }[];
  areas: { id: string; name: string; competencies: { id: string; name: string }[] }[];
  trainers: { userId: string; name: string; areaId: string }[];
}

const TWI_LABEL: Record<string, string> = { PREPARE: "1 Prepare", PRESENT: "2 Present", TRY_OUT: "3 Try Out", FOLLOW_UP: "4 Follow Up" };
const STATUS_LABEL: Record<string, [string, "gray" | "yellow" | "green" | "blue" | "red"]> = {
  ASSIGNED: ["Assigned", "gray"], IN_PROGRESS: ["In progress", "yellow"], PRACTICAL_PASSED: ["Practical passed", "green"], COMPLETED: ["Completed", "blue"], CANCELLED: ["Cancelled", "gray"],
};

export function TrainingClient({ data }: { data: TrainingData }) {
  const router = useRouter();
  const [dialog, setDialog] = useState<{ e: string; c: string } | null>(null);
  const [msg, setMsg] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  const [prefill, setPrefill] = useState<{ employeeId: string; competencyId: string } | null>(null);

  const act = (id: string, version: number, action: string) =>
    start(async () => {
      const f = new FormData();
      f.set("id", id); f.set("expectedVersion", String(version)); f.set("action", action);
      setMsg(await advanceTraining(f));
      router.refresh();
    });

  const assignmentTable = (rows: Row[], closed = false) =>
    rows.length === 0 ? <Empty title="Nothing here">No training in this list.</Empty> : (
      <Table>
        <thead><tr><Th>Employee</Th><Th>Competency</Th><Th>Trainer</Th><Th>TWI step</Th><Th>Status</Th><Th>{closed ? "Completed" : "Due"}</Th><Th className="text-right">Actions</Th></tr></thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.id}>
              <Td><Link className="font-semibold text-navy-700 hover:underline" href={`/employees/${r.employeeId}`}>{r.employee}</Link> {r.isSample && <SampleBadge />}<div className="text-xs text-muted">Shift {r.shift}</div></Td>
              <Td>{r.competency}<div className="text-xs text-muted">{r.area}</div></Td>
              <Td>{r.trainer}</Td>
              <Td><TwiSteps step={r.twiStep} /></Td>
              <Td><Badge tone={STATUS_LABEL[r.status][1]}>{STATUS_LABEL[r.status][0]}</Badge></Td>
              <Td>{closed ? fmtDate(r.completed) : r.due ? <span className={r.overdue ? "font-bold text-st-red" : ""}>{fmtDate(r.due)}{r.overdue ? " · overdue" : ""}</span> : "—"}</Td>
              <Td className="text-right">
                {!closed && (
                  <div className="flex flex-wrap justify-end gap-1">
                    {data.canManage && r.status === "ASSIGNED" && <Button size="sm" disabled={pending} onClick={() => act(r.id, r.version, "start")}>Start</Button>}
                    {data.canManage && r.status === "IN_PROGRESS" && r.twiStep !== "FOLLOW_UP" && <Button size="sm" variant="secondary" disabled={pending} onClick={() => act(r.id, r.version, "next")}>Next step</Button>}
                    <Button size="sm" variant="secondary" onClick={() => setDialog({ e: r.employeeId, c: r.competencyId })}>{r.twiStep === "FOLLOW_UP" || r.status === "PRACTICAL_PASSED" ? "Evaluate / verify" : "Open"}</Button>
                    {data.canManage && <Button size="sm" variant="ghost" disabled={pending} onClick={() => act(r.id, r.version, "cancel")}>Cancel</Button>}
                  </div>
                )}
              </Td>
            </tr>
          ))}
        </tbody>
      </Table>
    );

  return (
    <div>
      {msg && <div className="p-4 pb-0"><Alert tone={msg.ok ? "success" : msg.conflict ? "warn" : "error"}>{msg.ok ? msg.message : msg.error}</Alert></div>}
      {data.tab === "active" && assignmentTable(data.active)}
      {data.tab === "overdue" && assignmentTable(data.overdue)}
      {data.tab === "done" && assignmentTable(data.done, true)}
      {data.tab === "verify" && (data.verify.length === 0 ? <Empty title="No qualifications waiting for approval" /> : (
        <Table>
          <thead><tr><Th>Employee</Th><Th>Competency</Th><Th>Requested by</Th><Th>Submitted</Th><Th /></tr></thead>
          <tbody>
            {data.verify.map((v) => (
              <tr key={v.id}>
                <Td><span className="font-semibold">{v.employee}</span> {v.isSample && <SampleBadge />}</Td>
                <Td>{v.competency}<div className="text-xs text-muted">{v.area}</div></Td>
                <Td>{v.by}</Td><Td>{fmtDateTime(v.at)}</Td>
                <Td className="text-right"><Button size="sm" onClick={() => setDialog({ e: v.employeeId, c: v.competencyId })}>Review</Button></Td>
              </tr>
            ))}
          </tbody>
        </Table>
      ))}
      {data.tab === "requal" && (data.requal.length === 0 ? <Empty title="No requalifications due" /> : (
        <Table>
          <thead><tr><Th>Employee</Th><Th>Competency</Th><Th>Expires</Th><Th>Training</Th><Th /></tr></thead>
          <tbody>
            {data.requal.map((r) => (
              <tr key={`${r.employeeId}:${r.competencyId}`}>
                <Td><Link className="font-semibold text-navy-700 hover:underline" href={`/employees/${r.employeeId}`}>{r.employee}</Link> {r.isSample && <SampleBadge />}</Td>
                <Td>{r.competency}<div className="text-xs text-muted">{r.area}</div></Td>
                <Td><Badge tone="red">{r.expired ? "Expired " : ""}{fmtDate(r.expiresAt)}</Badge></Td>
                <Td>{r.assigned ? <Badge tone="yellow">Assigned</Badge> : <span className="text-muted">Not assigned</span>}</Td>
                <Td className="text-right">{data.canManage && !r.assigned && <Link href={`/training?tab=assign&e=${r.employeeId}&c=${r.competencyId}`} onClick={() => setPrefill({ employeeId: r.employeeId, competencyId: r.competencyId })}><Button size="sm" variant="secondary">Assign</Button></Link>}</Td>
              </tr>
            ))}
          </tbody>
        </Table>
      ))}
      {data.tab === "gaps" && (data.gaps.length === 0 ? <Empty title="Every required skill gap already has training assigned" /> : (
        <>
          <p className="px-4 pt-4 text-sm text-muted">Required skills that are red and have no training assigned yet{data.gapsTotal > data.gaps.length ? ` (first ${data.gaps.length} of ${data.gapsTotal})` : ""}. Critical skills first.</p>
          <Table>
            <thead><tr><Th>Employee</Th><Th>Department</Th><Th>Competency</Th><Th /></tr></thead>
            <tbody>
              {data.gaps.map((g) => (
                <tr key={`${g.employeeId}:${g.competencyId}`}>
                  <Td><Link className="font-semibold text-navy-700 hover:underline" href={`/employees/${g.employeeId}`}>{g.employee}</Link> {g.isSample && <SampleBadge />}<div className="text-xs text-muted">Shift {g.shift}</div></Td>
                  <Td>{g.area}</Td>
                  <Td>{g.competency} {g.critical && <Badge tone="red">Critical</Badge>}</Td>
                  <Td className="text-right">{data.canManage && <Link href={`/training?tab=assign&e=${g.employeeId}&c=${g.competencyId}`} onClick={() => setPrefill({ employeeId: g.employeeId, competencyId: g.competencyId })}><Button size="sm" variant="secondary">Assign</Button></Link>}</Td>
                </tr>
              ))}
            </tbody>
          </Table>
        </>
      ))}
      {data.tab === "assign" && data.canManage && <AssignForm data={data} prefill={prefill} onDone={(r) => { setMsg(r); router.refresh(); }} />}
      {dialog && <AssessmentDialog employeeId={dialog.e} competencyId={dialog.c} onClose={() => setDialog(null)} onSaved={() => router.refresh()} />}
    </div>
  );
}

function TwiSteps({ step }: { step: string }) {
  const steps = ["PREPARE", "PRESENT", "TRY_OUT", "FOLLOW_UP"];
  const i = steps.indexOf(step);
  return (
    <span className="inline-flex items-center gap-1" aria-label={`TWI step ${TWI_LABEL[step]}`} title={TWI_LABEL[step]}>
      {steps.map((s, k) => <span key={s} className={`h-2 w-5 rounded ${k <= i ? "bg-navy-700" : "bg-line"}`} />)}
      <span className="ml-1 text-xs text-muted">{TWI_LABEL[step]}</span>
    </span>
  );
}

function AssignForm({ data, prefill, onDone }: { data: TrainingData; prefill: { employeeId: string; competencyId: string } | null; onDone: (r: ActionResult) => void }) {
  const params = typeof window !== "undefined" ? new URLSearchParams(window.location.search) : null;
  const initE = prefill?.employeeId ?? params?.get("e") ?? "";
  const initC = prefill?.competencyId ?? params?.get("c") ?? "";
  const initArea = data.areas.find((a) => a.competencies.some((c) => c.id === initC))?.id ?? "";
  const [areaId, setAreaId] = useState(initArea);
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  const area = data.areas.find((a) => a.id === areaId);
  const trainers = useMemo(() => data.trainers.filter((t) => t.areaId === areaId), [data.trainers, areaId]);
  const due = new Date(Date.now() + 30 * 86_400_000).toISOString().slice(0, 10);
  return (
    <form className="grid max-w-2xl gap-4 p-5" onSubmit={(e) => { e.preventDefault(); const f = e.currentTarget; start(async () => { const r = await assignTraining(new FormData(f)); setRes(r); if (r.ok) { onDone(r); } }); }}>
      <h2 className="text-lg font-bold">Assign training</h2>
      {res && <Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}</Alert>}
      <Field label="Employee" htmlFor="t-emp">
        <Select id="t-emp" name="employeeId" defaultValue={initE} required>
          <option value="">Choose…</option>
          {data.employees.map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
        </Select>
      </Field>
      <div className="grid gap-4 sm:grid-cols-2">
        <Field label="Department" htmlFor="t-area">
          <Select id="t-area" value={areaId} onChange={(e) => setAreaId(e.target.value)} required>
            <option value="">Choose…</option>
            {data.areas.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
          </Select>
        </Field>
        <Field label="Competency" htmlFor="t-comp">
          <Select id="t-comp" name="competencyId" defaultValue={initC} required key={areaId}>
            <option value="">Choose…</option>
            {area?.competencies.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </Select>
        </Field>
      </div>
      <div className="grid gap-4 sm:grid-cols-2">
        <Field label="Trainer" htmlFor="t-trainer" hint={areaId && trainers.length === 0 ? "No authorized trainer for this department yet." : "Only trainers authorized for this department are listed."}>
          <Select id="t-trainer" name="trainerUserId">
            <option value="">Assign later</option>
            {trainers.map((t) => <option key={t.userId} value={t.userId}>{t.name}</option>)}
          </Select>
        </Field>
        <Field label="Due date" htmlFor="t-due"><Input id="t-due" type="date" name="dueDate" defaultValue={due} /></Field>
      </div>
      <Field label="Notes" htmlFor="t-notes"><Textarea id="t-notes" name="notes" /></Field>
      <Button type="submit" size="lg" disabled={pending}>{pending ? "Assigning…" : "Assign training"}</Button>
    </form>
  );
}
