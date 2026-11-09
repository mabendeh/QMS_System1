"use client";
import { useState, useTransition } from "react";
import { commitImport, previewImport, type ImportRow } from "@/app/actions/admin";
import type { ActionResult } from "@/app/actions/skills";
import { Alert, Badge, Button, Card, CardHeader, Field, Input } from "@/components/ui";

export function ImportClient() {
  const [rows, setRows] = useState<ImportRow[] | null>(null);
  const [summary, setSummary] = useState<{ create: number; update: number; errors: number } | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [res, setRes] = useState<ActionResult | null>(null);
  const [updateExisting, setUpdateExisting] = useState(false);
  const [pending, start] = useTransition();
  return (
    <div className="grid gap-5">
      <Card className="p-5">
        <form className="grid gap-4" onSubmit={(e) => { e.preventDefault(); const f = e.currentTarget; setRes(null); start(async () => { const r = await previewImport(new FormData(f)); if (r.ok) { setRows(r.rows); setSummary(r.summary); setErr(null); } else { setErr(r.error); setRows(null); } }); }}>
          <Field label="Employee list" htmlFor="imp-file" hint="Columns: Employee ID, First Name, Last Name, Shift (A/B/C). Optional: Department, Job Title, Hire Date (YYYY-MM-DD), Supervisor Email.">
            <Input id="imp-file" name="file" type="file" accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required className="py-2" />
          </Field>
          <div><Button type="submit" disabled={pending}>{pending ? "Checking…" : "Preview import"}</Button></div>
        </form>
        {err && <div className="mt-4"><Alert tone="error">{err}</Alert></div>}
      </Card>
      {rows && summary && (
        <Card>
          <CardHeader title="Preview" subtitle={`${summary.create} new · ${summary.update} already exist · ${summary.errors} with problems (will be skipped)`} />
          <div className="max-h-[480px] overflow-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-canvas"><tr className="text-left text-xs font-bold uppercase text-muted"><th className="px-3 py-2">Row</th><th className="px-3 py-2">Employee ID</th><th className="px-3 py-2">Name</th><th className="px-3 py-2">Shift</th><th className="px-3 py-2">Department</th><th className="px-3 py-2">Job title</th><th className="px-3 py-2">Result</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.row} className="border-t border-line">
                    <td className="px-3 py-2">{r.row}</td><td className="px-3 py-2 font-mono">{r.employeeNumber}</td><td className="px-3 py-2">{r.firstName} {r.lastName}</td>
                    <td className="px-3 py-2">{r.shift}</td><td className="px-3 py-2">{r.department}</td><td className="px-3 py-2">{r.jobTitle}</td>
                    <td className="px-3 py-2">{r.errors.length ? <span className="text-st-red">{r.errors.join("; ")}</span> : r.action === "update" ? <Badge tone="yellow">Exists</Badge> : <Badge tone="blue">New</Badge>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="grid gap-3 border-t border-line p-5">
            {summary.update > 0 && (
              <label className="flex items-start gap-2 text-sm">
                <input type="checkbox" className="mt-0.5 h-5 w-5" checked={updateExisting} onChange={(e) => setUpdateExisting(e.target.checked)} />
                <span><b>Also update details of the {summary.update} existing employees</b> (name, shift, department, job title, supervisor). Unticked, existing employees are left unchanged. Qualifications are never changed by an import.</span>
              </label>
            )}
            {res && <Alert tone={res.ok ? "success" : "error"}>{res.ok ? res.message : res.error}</Alert>}
            <div><Button disabled={pending || summary.create + summary.update === 0 || !!res?.ok} onClick={() => start(async () => setRes(await commitImport(JSON.stringify(rows), updateExisting)))}>
              {pending ? "Importing…" : `Import ${summary.create} new${updateExisting && summary.update ? ` and update ${summary.update}` : ""}`}
            </Button></div>
          </div>
        </Card>
      )}
    </div>
  );
}
