"use client";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { Paperclip } from "lucide-react";
import type { DisplayStatus } from "@/lib/domain/status";
import { StatusChip } from "@/components/StatusCell";
import { AssessmentDialog } from "@/components/matrix/AssessmentDialog";
import { Badge, Card, CardHeader } from "@/components/ui";
import { fmtDate } from "@/lib/utils";

export interface ProfileArea {
  id: string; name: string; kind: string; primary: boolean; qualified: boolean;
  rows: { id: string; name: string; critical: boolean; status: DisplayStatus; pending: boolean; required: boolean; assessedAt: string | null; evaluator: string | null; evidence: number; qualifiedAt: string | null; expiresAt: string | null; need: string | null }[];
}

export function ProfileSkills({ employeeId, areas }: { employeeId: string; areas: ProfileArea[] }) {
  const router = useRouter();
  const [open, setOpen] = useState<string | null>(null);
  return (
    <div className="grid content-start gap-4">
      {areas.map((a) => {
        const competent = a.rows.filter((r) => r.status === "COMPETENT").length;
        return (
          <Card key={a.id}>
            <CardHeader
              title={<span className="flex flex-wrap items-center gap-2">{a.name} {a.primary && <Badge tone="navy">Assigned department</Badge>} {a.kind === "PLANT_WIDE" && <Badge>Plant-wide</Badge>}</span>}
              subtitle={`${competent} of ${a.rows.length} Competent`}
              action={a.qualified ? <Badge tone="blue">Qualified</Badge> : <Badge tone="gray">Not yet qualified</Badge>}
            />
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs font-bold uppercase tracking-wide text-muted">
                    <th className="px-4 py-2">Competency</th><th className="px-2 py-2">Status</th><th className="px-2 py-2">Assessed</th>
                    <th className="px-2 py-2">Evaluator</th><th className="px-2 py-2">Evidence</th><th className="px-2 py-2">Qualified</th><th className="px-2 py-2">Requalify by</th><th className="px-4 py-2">Training need</th>
                  </tr>
                </thead>
                <tbody>
                  {a.rows.map((r) => (
                    <tr key={r.id} className="border-t border-line hover:bg-canvas">
                      <td className="px-4 py-2 font-medium">{r.name}{r.critical && <span className="ml-1 text-st-red" title="Critical skill">◆</span>}</td>
                      <td className="px-2 py-1.5">
                        <button onClick={() => setOpen(r.id)} aria-label={`Open ${r.name}`} className="rounded focus-visible:outline focus-visible:outline-2 focus-visible:outline-navy-800">
                          <StatusChip status={r.status} pending={r.pending} compact />
                        </button>
                      </td>
                      <td className="px-2 py-2 text-muted">{fmtDate(r.assessedAt)}</td>
                      <td className="px-2 py-2 text-muted">{r.evaluator ?? "—"}</td>
                      <td className="px-2 py-2">{r.evidence ? <span className="inline-flex items-center gap-1"><Paperclip className="h-3.5 w-3.5" />{r.evidence}</span> : <span className="text-muted">—</span>}</td>
                      <td className="px-2 py-2 text-muted">{fmtDate(r.qualifiedAt)}</td>
                      <td className="px-2 py-2">{r.expiresAt ? <span className={r.status === "EXPIRED" ? "font-semibold text-st-red" : ""}>{fmtDate(r.expiresAt)}</span> : <span className="text-muted">—</span>}</td>
                      <td className="px-4 py-2">{r.need ? <span className="text-st-red">{r.need}</span> : <span className="text-muted">—</span>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        );
      })}
      {open && <AssessmentDialog employeeId={employeeId} competencyId={open} onClose={() => setOpen(null)} onSaved={() => router.refresh()} />}
    </div>
  );
}
