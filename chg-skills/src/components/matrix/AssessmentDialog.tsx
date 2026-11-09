"use client";
import * as Dialog from "@radix-ui/react-dialog";
import { useCallback, useEffect, useState, useTransition } from "react";
import { X, Paperclip, ShieldCheck, ClipboardCheck } from "lucide-react";
import { decideVerification, getCellDetail, recordPracticalEvaluation, submitAssessment, type ActionResult } from "@/app/actions/skills";
import { SELECTABLE_STATUSES, STATUS_META, STORED_LABEL, type DisplayStatus, type StoredStatus } from "@/lib/domain/status";
import { StatusChip } from "@/components/StatusCell";
import { Alert, Badge, Button, Field, Input, SampleBadge, Textarea } from "@/components/ui";
import { cn, fmtDate, fmtDateTime } from "@/lib/utils";

type Detail = Awaited<ReturnType<typeof getCellDetail>>;
const today = () => new Date().toLocaleDateString("en-CA", { timeZone: "America/Chicago" });
const asDisplay = (s: StoredStatus): DisplayStatus => (s === "NOT_APPLICABLE" ? "NOT_ASSESSED" : s);

export function AssessmentDialog({
  employeeId, competencyId, onClose, onSaved,
}: { employeeId: string; competencyId: string; onClose: () => void; onSaved: () => void }) {
  const [d, setD] = useState<Detail | null>(null);
  const [loadErr, setLoadErr] = useState<string | null>(null);
  const [res, setRes] = useState<ActionResult | null>(null);
  const [tab, setTab] = useState<"assess" | "practical">("assess");
  const [picked, setPicked] = useState<StoredStatus | null>(null);
  const [pending, start] = useTransition();

  const load = useCallback(async () => {
    try {
      const x = await getCellDetail(employeeId, competencyId);
      setD(x);
      setPicked(null);
      if (!x.permissions.canAssess && x.permissions.canEval) setTab("practical");
    } catch (e) {
      setLoadErr(e instanceof Error ? e.message : "Could not load this skill.");
    }
  }, [employeeId, competencyId]);
  useEffect(() => { load(); }, [load]);

  const run = (fn: (f: FormData) => Promise<ActionResult>, form: HTMLFormElement) =>
    start(async () => {
      const r = await fn(new FormData(form));
      setRes(r);
      if (r.ok || r.conflict) {
        await load();
        onSaved();
        if (r.ok) form.reset();
      }
    });

  return (
    <Dialog.Root open onOpenChange={(o) => !o && onClose()}>
      <Dialog.Portal>
        <Dialog.Overlay className="fixed inset-0 z-40 bg-navy-900/50" />
        <Dialog.Content
          className="fixed inset-x-0 bottom-0 z-50 max-h-[92vh] overflow-y-auto rounded-t-xl bg-white shadow-2xl sm:inset-x-auto sm:left-1/2 sm:top-1/2 sm:bottom-auto sm:w-[640px] sm:-translate-x-1/2 sm:-translate-y-1/2 sm:rounded-xl"
          aria-describedby={undefined}
        >
          <div className="sticky top-0 z-10 flex items-start justify-between gap-3 border-b border-line bg-white px-5 py-4">
            <div className="min-w-0">
              <Dialog.Title className="text-lg font-bold leading-tight">{d ? d.competency.name : "Loading…"}</Dialog.Title>
              {d && (
                <p className="mt-1 flex flex-wrap items-center gap-2 text-sm text-muted">
                  <span className="font-semibold text-ink">{d.employee.name}</span> · {d.employee.number} · {d.competency.area}
                  {d.employee.isSample && <SampleBadge />}
                  {d.competency.isCritical && <Badge tone="red">Critical skill</Badge>}
                </p>
              )}
            </div>
            <Dialog.Close className="rounded p-1.5 text-muted hover:bg-canvas" aria-label="Close"><X /></Dialog.Close>
          </div>

          {loadErr && <div className="p-5"><Alert tone="error">{loadErr}</Alert></div>}
          {!d && !loadErr && <p className="p-6 text-muted">Loading…</p>}
          {d && (
            <div className="grid gap-5 px-5 py-4">
              <div className="flex flex-wrap items-center gap-3 rounded-lg bg-canvas p-3">
                <StatusChip status={d.display} pending={!!d.record?.pendingStatus} />
                <div className="text-sm">
                  <p className="font-bold">{STATUS_META[d.display].label}{d.record?.pendingStatus && <span className="text-navy-700"> · {STORED_LABEL[d.record.pendingStatus]} pending verification</span>}</p>
                  <p className="text-muted">
                    {d.record?.lastAssessedAt ? <>Last assessed {fmtDate(d.record.lastAssessedAt)}{d.record.lastAssessedBy ? ` by ${d.record.lastAssessedBy}` : ""}</> : d.required ? "Required for this job – not assessed yet" : "Not required for this job"}
                    {d.record?.expiresAt && <> · Requalify by {fmtDate(d.record.expiresAt)}</>}
                  </p>
                </div>
              </div>
              {d.competency.description && <p className="text-sm text-muted">{d.competency.description}</p>}

              {res && (
                <Alert tone={res.ok ? (res.outcome === "PENDING" ? "warn" : "success") : res.conflict ? "warn" : "error"} title={res.ok ? (res.outcome === "PENDING" ? "Submitted for verification" : "Saved") : res.conflict ? "Changed by someone else" : "Not saved"}>
                  <p>{res.ok ? res.message : res.error}</p>
                  {!res.ok && res.reasons && <ul className="mt-1 list-disc pl-5">{res.reasons.map((r) => <li key={r}>{r}</li>)}</ul>}
                </Alert>
              )}

              {/* Verification of a pending Competent request */}
              {d.pending && (
                <section className="rounded-lg border-2 border-navy-200 p-4">
                  <h3 className="flex items-center gap-2 font-bold"><ShieldCheck className="h-5 w-5 text-navy-700" /> Pending verification</h3>
                  <p className="mt-1 text-sm text-muted">{d.pending.by} requested <b>{STORED_LABEL[d.pending.requested]}</b> on {fmtDateTime(d.pending.at)}.{d.pending.notes ? ` “${d.pending.notes}”` : ""}</p>
                  {d.permissions.canVerify ? (
                    <form className="mt-3 grid gap-3" onSubmit={(e) => { e.preventDefault(); const sub = (e.nativeEvent as SubmitEvent).submitter as HTMLButtonElement; const f = e.currentTarget; (f.elements.namedItem("decision") as HTMLInputElement).value = sub.value; run(decideVerification, f); }}>
                      <input type="hidden" name="assessmentId" value={d.pending.id} />
                      <input type="hidden" name="decision" value="" />
                      <Field label="Comments" htmlFor="v-comments" hint="Required when rejecting.">
                        <Textarea id="v-comments" name="comments" />
                      </Field>
                      <div className="flex flex-wrap gap-2">
                        <Button type="submit" value="APPROVED" variant="success" disabled={pending}>Verify as Competent</Button>
                        <Button type="submit" value="REJECTED" variant="secondary" disabled={pending}>Reject</Button>
                      </div>
                    </form>
                  ) : (
                    <p className="mt-2 text-sm">Waiting for a Production Manager or Quality evaluator to verify.</p>
                  )}
                </section>
              )}

              {(d.permissions.canAssess || d.permissions.canEval) && (
                <div className="flex gap-1 rounded-lg bg-canvas p-1" role="tablist">
                  {d.permissions.canAssess && <TabBtn on={tab === "assess"} onClick={() => setTab("assess")}>Update status</TabBtn>}
                  {d.permissions.canEval && <TabBtn on={tab === "practical"} onClick={() => setTab("practical")}>Practical evaluation</TabBtn>}
                </div>
              )}
              {!d.permissions.canAssess && !d.permissions.canEval && (
                <Alert tone="info">You can view this skill. Only the assigned supervisor, an authorized trainer or the Production Manager can change it.</Alert>
              )}

              {d.permissions.canAssess && tab === "assess" && (
                <form className="grid gap-4" onSubmit={(e) => { e.preventDefault(); run(submitAssessment, e.currentTarget); }}>
                  <input type="hidden" name="employeeId" value={d.employee.id} />
                  <input type="hidden" name="competencyId" value={d.competency.id} />
                  <input type="hidden" name="expectedVersion" value={d.record?.version ?? 0} />
                  <fieldset>
                    <legend className="mb-2 text-sm font-semibold">Skill level</legend>
                    <div className="grid grid-cols-2 gap-2 sm:grid-cols-3">
                      {SELECTABLE_STATUSES.map((s) => {
                        const m = STATUS_META[asDisplay(s)];
                        const on = picked === s;
                        return (
                          <label key={s} className={cn("flex min-h-[52px] cursor-pointer items-center gap-2 rounded-lg border-2 px-3 py-2 text-sm font-semibold", on ? "border-navy-800 bg-navy-50" : "border-line hover:border-navy-200")}>
                            <input type="radio" name="status" value={s} required className="sr-only" checked={on} onChange={() => setPicked(s)} />
                            <span className="inline-flex h-7 w-9 shrink-0 items-center justify-center rounded text-xs font-bold" style={{ background: m.color, color: m.text }} aria-hidden>{m.short}</span>
                            {STORED_LABEL[s]}
                          </label>
                        );
                      })}
                    </div>
                  </fieldset>
                  {picked === "COMPETENT" && (
                    <Alert tone="info">
                      Competent needs a passed practical evaluation{d.competency.isCritical ? " and supporting evidence" : ""}, then verification by an authorized evaluator.
                      {d.practical ? <> Latest practical: <b>{d.practical.result === "PASS" ? "Passed" : "Not passed"}</b> on {fmtDate(d.practical.evaluatedAt)} by {d.practical.evaluator}.</> : <> No practical evaluation recorded yet.</>}
                    </Alert>
                  )}
                  <div className="grid gap-4 sm:grid-cols-2">
                    <Field label="Training date" htmlFor="a-date" hint={picked === "TRAINED" ? "Required for Trained." : "Optional"}>
                      <Input id="a-date" name="trainingDate" type="date" max={today()} defaultValue={d.record?.trainingDate ? new Date(d.record.trainingDate).toISOString().slice(0, 10) : ""} />
                    </Field>
                    <Field label="Evaluator" htmlFor="a-eval">
                      <Input id="a-eval" name="evaluatorName" defaultValue={d.actorName} />
                    </Field>
                  </div>
                  <Field label="Assessment notes" htmlFor="a-notes" hint="Required when lowering a Competent qualification.">
                    <Textarea id="a-notes" name="notes" placeholder="What did you observe?" />
                  </Field>
                  <Field label="Supporting evidence" htmlFor="a-file" hint="PDF, PNG, JPEG or WEBP, up to 10 MB.">
                    <Input id="a-file" name="evidence" type="file" accept="application/pdf,image/png,image/jpeg,image/webp" className="py-2" />
                  </Field>
                  <Button type="submit" size="lg" disabled={pending || !picked}>
                    {pending ? "Saving…" : picked === "COMPETENT" && !d.permissions.canVerify ? "Submit for verification" : "Save assessment"}
                  </Button>
                </form>
              )}

              {d.permissions.canEval && tab === "practical" && (
                <form className="grid gap-4" onSubmit={(e) => { e.preventDefault(); run(recordPracticalEvaluation, e.currentTarget); }}>
                  <input type="hidden" name="employeeId" value={d.employee.id} />
                  <input type="hidden" name="competencyId" value={d.competency.id} />
                  <p className="text-sm text-muted">Record the Try-Out / Follow-Up demonstration. A pass must confirm the employee operated independently, following the approved procedure.</p>
                  <fieldset className="grid grid-cols-2 gap-2">
                    <legend className="mb-2 text-sm font-semibold">Result</legend>
                    {[["PASS", "Passed"], ["FAIL", "Not passed"]].map(([v, l]) => (
                      <label key={v} className="flex min-h-[52px] cursor-pointer items-center gap-3 rounded-lg border-2 border-line px-3 font-semibold has-[:checked]:border-navy-800 has-[:checked]:bg-navy-50">
                        <input type="radio" name="result" value={v} required className="h-5 w-5" /> {l}
                      </label>
                    ))}
                  </fieldset>
                  <label className="flex items-center gap-3 text-sm font-semibold">
                    <input type="checkbox" name="independentOperation" className="h-5 w-5" /> Operated independently without help
                  </label>
                  <Field label="Evaluation date" htmlFor="p-date">
                    <Input id="p-date" name="evaluatedAt" type="date" max={today()} defaultValue={today()} />
                  </Field>
                  <Field label="Notes" htmlFor="p-notes"><Textarea id="p-notes" name="notes" /></Field>
                  <Field label="Evidence" htmlFor="p-file" hint="Signed checklist or photo (optional; required before critical skills can be verified).">
                    <Input id="p-file" name="evidence" type="file" accept="application/pdf,image/png,image/jpeg,image/webp" className="py-2" />
                  </Field>
                  <Button type="submit" size="lg" disabled={pending}><ClipboardCheck className="h-5 w-5" /> {pending ? "Saving…" : "Record evaluation"}</Button>
                </form>
              )}

              <section className="grid gap-3 border-t border-line pt-4 text-sm">
                <div>
                  <h3 className="font-bold">Practical evaluation</h3>
                  <p className="text-muted">{d.practical ? `${d.practical.result === "PASS" ? "Passed" : "Not passed"}${d.practical.independentOperation ? " · independent operation" : ""} · ${fmtDate(d.practical.evaluatedAt)} · ${d.practical.evaluator}` : "None recorded"}</p>
                </div>
                <div>
                  <h3 className="font-bold">Evidence</h3>
                  {d.evidence.length ? (
                    <ul className="mt-1 grid gap-1">
                      {d.evidence.map((e) => (
                        <li key={e.id}><a className="inline-flex items-center gap-1 font-semibold text-navy-700 underline" href={`/api/evidence/${e.id}`} target="_blank" rel="noreferrer"><Paperclip className="h-4 w-4" />{e.name}</a> <span className="text-muted">· {fmtDate(e.at)}</span></li>
                      ))}
                    </ul>
                  ) : <p className="text-muted">None uploaded</p>}
                </div>
                <div>
                  <h3 className="font-bold">History</h3>
                  {d.history.length ? (
                    <ol className="mt-1 grid gap-1.5">
                      {d.history.map((h) => (
                        <li key={h.id} className="flex flex-wrap gap-x-2">
                          <span className="text-muted">{fmtDateTime(h.at)}</span>
                          <span><b>{STORED_LABEL[h.requested]}</b> · {h.state.replace(/_/g, " ").toLowerCase()} · {h.by}</span>
                          {h.notes && <span className="w-full text-muted">“{h.notes}”</span>}
                        </li>
                      ))}
                    </ol>
                  ) : <p className="text-muted">No assessments yet</p>}
                </div>
              </section>
            </div>
          )}
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function TabBtn({ on, children, onClick }: { on: boolean; children: React.ReactNode; onClick: () => void }) {
  return (
    <button type="button" role="tab" aria-selected={on} onClick={onClick} className={cn("flex-1 rounded-md px-3 py-2 text-sm font-semibold", on ? "bg-white shadow-sm" : "text-muted")}>
      {children}
    </button>
  );
}
