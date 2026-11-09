"use client";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { saveTargets } from "@/app/actions/coverage";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Button, Card, Input } from "@/components/ui";

export function TargetsEditor({ shifts, rows }: { shifts: { id: string; code: string }[]; rows: { areaId: string; areaName: string; targets: (number | null)[] }[] }) {
  const router = useRouter();
  const [res, setRes] = useState<ActionResult | null>(null);
  const [pending, start] = useTransition();
  return (
    <Card className="p-5">
      <form onSubmit={(e) => { e.preventDefault(); const f = e.currentTarget; start(async () => { const r = await saveTargets(new FormData(f)); setRes(r); if (r.ok) router.push("/coverage"); }); }}>
        <p className="mb-3 text-sm text-muted">Minimum number of qualified people each department needs on each shift. Leave blank for no target.</p>
        {res && <div className="mb-3"><Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}</Alert></div>}
        <table className="w-full text-sm">
          <thead><tr className="text-left text-xs font-bold uppercase text-muted"><th className="py-2">Department</th>{shifts.map((s) => <th key={s.id} className="py-2">Shift {s.code}</th>)}</tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.areaId} className="border-t border-line">
                <td className="py-2 font-semibold">{r.areaName}</td>
                {shifts.map((s, i) => (
                  <td key={s.id} className="py-2 pr-3"><Input aria-label={`${r.areaName} shift ${s.code} target`} name={`t:${r.areaId}:${s.id}`} type="number" min={0} max={200} defaultValue={r.targets[i] ?? ""} className="max-w-[110px]" /></td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        <div className="mt-4 flex gap-2"><Button type="submit" disabled={pending}>Save targets</Button><Button type="button" variant="secondary" onClick={() => router.push("/coverage")}>Cancel</Button></div>
      </form>
    </Card>
  );
}
