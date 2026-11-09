"use client";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useEffect, useMemo, useState } from "react";
import { Download, FileText, Printer, Search, Maximize2, Minimize2, ArrowUpDown } from "lucide-react";
import { STATUS_META, type DisplayStatus } from "@/lib/domain/status";
import { StatusChip, StatusLegend } from "@/components/StatusCell";
import { Badge, Button, Input, SampleBadge, Select } from "@/components/ui";
import { AssessmentDialog } from "./AssessmentDialog";
import { cn } from "@/lib/utils";

export interface MatrixPayload {
  view: { id: string; code: string; name: string; description: string | null } | null;
  areas: { id: string; code: string; name: string; kind: string; count: number }[];
  columns: { id: string; code: string; name: string; critical: boolean }[];
  employees: {
    id: string; number: string; name: string; lastName: string; isSample: boolean; shift: string; area: string; areaId: string | null;
    job: string; supervisorId: string | null; supervisor: string; level: number | null; levelName: string; eligible: boolean;
    nextLevel: string | null; fullyQualified: boolean; crossTrained: boolean;
    cells: { s: DisplayStatus; p: boolean; v: number; n?: string }[];
  }[];
  canExport: boolean;
  focus: string | null;
}

export function MatrixView({ data }: { data: MatrixPayload }) {
  const router = useRouter();
  const [q, setQ] = useState("");
  const [shift, setShift] = useState("");
  const [dept, setDept] = useState("");
  const [job, setJob] = useState("");
  const [sup, setSup] = useState("");
  const [status, setStatus] = useState("");
  const [level, setLevel] = useState("");
  const [elig, setElig] = useState("");
  const [sort, setSort] = useState<"name" | "number">("name");
  const [desc, setDesc] = useState(false);
  const [compact, setCompact] = useState(false);
  const [everyone, setEveryone] = useState(false);
  const [open, setOpen] = useState<{ employeeId: string; competencyId: string } | null>(null);

  useEffect(() => {
    try { setCompact(localStorage.getItem("matrix.compact") === "1"); } catch {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem("matrix.compact", compact ? "1" : "0"); } catch {}
  }, [compact]);
  useEffect(() => {
    if (data.focus) document.getElementById(`row-${data.focus}`)?.scrollIntoView({ block: "center" });
  }, [data.focus]);

  const uniq = (xs: string[]) => [...new Set(xs)].filter((x) => x && x !== "—").sort();
  const jobs = uniq(data.employees.map((e) => e.job));
  const sups = [...new Map(data.employees.filter((e) => e.supervisorId).map((e) => [e.supervisorId!, e.supervisor])).entries()];
  const levels = [...new Map(data.employees.filter((e) => e.level != null).map((e) => [String(e.level), e.levelName])).entries()].sort();

  const rows = useMemo(() => {
    const needle = q.trim().toLowerCase();
    const out = data.employees.filter((e) =>
      (!data.view || everyone || e.areaId === data.view.id || e.cells.some((c) => c.s !== "NOT_ASSESSED")) &&
      (!needle || e.name.toLowerCase().includes(needle) || e.number.toLowerCase().includes(needle)) &&
      (!shift || e.shift === shift) && (!dept || e.areaId === dept) && (!job || e.job === job) && (!sup || e.supervisorId === sup) &&
      (!level || String(e.level) === level) &&
      (!elig || (elig === "eligible" ? e.eligible : elig === "full" ? e.fullyQualified : elig === "cross" ? e.crossTrained : !e.eligible)) &&
      (!status || (status === "PENDING" ? e.cells.some((c) => c.p) : e.cells.some((c) => c.s === status))),
    );
    out.sort((a, b) => (sort === "name" ? a.lastName.localeCompare(b.lastName) || a.name.localeCompare(b.name) : a.number.localeCompare(b.number)) * (desc ? -1 : 1));
    return out;
  }, [data.employees, data.view, everyone, q, shift, dept, job, sup, status, level, elig, sort, desc]);

  const exportQs = new URLSearchParams({ ...(data.view ? { area: data.view.code } : {}), ...(shift ? { shift } : {}), ...(dept ? { dept } : {}), ...(job ? { job } : {}), ...(sup ? { sup } : {}) }).toString();
  const filtersOn = [q, shift, dept, job, sup, status, level, elig].some(Boolean);

  return (
    <div className={cn(compact && "compact")}>
      <div className="mb-4 flex flex-wrap items-end justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold tracking-tight">Production Skills Matrix</h1>
          <p className="mt-1 text-sm text-muted">
            {data.view ? (
              <>Showing every <b>{data.view.name}</b> competency. Click a cell to assess.</>
            ) : (
              <>One column per department. Click a cell to open that department&apos;s competencies.</>
            )}
          </p>
        </div>
        <div className="no-print flex flex-wrap gap-2">
          <Button variant="secondary" size="sm" onClick={() => setCompact(!compact)} aria-pressed={compact}>
            {compact ? <Minimize2 className="h-4 w-4" /> : <Maximize2 className="h-4 w-4" />} {compact ? "Standard view" : "Compact view"}
          </Button>
          <Button variant="secondary" size="sm" onClick={() => window.print()}><Printer className="h-4 w-4" /> Print</Button>
          {data.canExport && (
            <>
              <a href={`/api/export/matrix?format=xlsx&${exportQs}`}><Button variant="secondary" size="sm"><Download className="h-4 w-4" /> Excel</Button></a>
              <a href={`/api/export/matrix?format=pdf&${exportQs}`}><Button variant="secondary" size="sm"><FileText className="h-4 w-4" /> PDF</Button></a>
            </>
          )}
        </div>
      </div>

      {/* Department tabs */}
      <div className="no-print -mx-1 mb-3 flex gap-1 overflow-x-auto pb-1" role="tablist" aria-label="Department view">
        <Link role="tab" aria-selected={!data.view} href="/matrix" className={tabCls(!data.view)}>All departments</Link>
        {data.areas.map((a) => (
          <Link key={a.id} role="tab" aria-selected={data.view?.id === a.id} href={`/matrix?area=${a.code}`} className={tabCls(data.view?.id === a.id)}>
            {a.name}
          </Link>
        ))}
      </div>
      {data.view?.description && <p className="mb-3 rounded-md border border-navy-100 bg-navy-50 px-3 py-2 text-sm text-navy-800">{data.view.description}</p>}

      {/* Filters */}
      <div className="no-print mb-3 grid gap-2 rounded-lg border border-line bg-white p-3 sm:grid-cols-2 lg:grid-cols-4">
        <label className="relative">
          <span className="sr-only">Search employee</span>
          <Search className="pointer-events-none absolute left-3 top-3 h-5 w-5 text-muted" aria-hidden />
          <Input placeholder="Search name or ID" value={q} onChange={(e) => setQ(e.target.value)} className="pl-10" />
        </label>
        <Select aria-label="Shift" value={shift} onChange={(e) => setShift(e.target.value)}>
          <option value="">All shifts</option>
          {["A", "B", "C"].map((s) => <option key={s} value={s}>Shift {s}</option>)}
        </Select>
        <Select aria-label="Department" value={dept} onChange={(e) => setDept(e.target.value)}>
          <option value="">All departments</option>
          {data.areas.filter((a) => a.kind === "PRODUCTION").map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
        </Select>
        <Select aria-label="Job title" value={job} onChange={(e) => setJob(e.target.value)}>
          <option value="">All job titles</option>
          {jobs.map((j) => <option key={j}>{j}</option>)}
        </Select>
        <Select aria-label="Supervisor" value={sup} onChange={(e) => setSup(e.target.value)}>
          <option value="">All supervisors</option>
          {sups.map(([id, n]) => <option key={id} value={id}>{n}</option>)}
        </Select>
        <Select aria-label="Skill status" value={status} onChange={(e) => setStatus(e.target.value)}>
          <option value="">Any skill status</option>
          {(["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "EXPIRED"] as DisplayStatus[]).map((s) => <option key={s} value={s}>Has {STATUS_META[s].label}</option>)}
          <option value="PENDING">Has pending verification</option>
        </Select>
        <Select aria-label="Qualification level" value={level} onChange={(e) => setLevel(e.target.value)}>
            <option value="">Any level</option>
            {levels.map(([r, n]) => <option key={r} value={r}>L{r} {n}</option>)}
          </Select>
        <Select aria-label="Promotion eligibility" value={elig} onChange={(e) => setElig(e.target.value)}>
            <option value="">Any eligibility</option>
            <option value="eligible">Meets next-level requirements</option>
            <option value="not">Not yet eligible</option>
            <option value="full">Fully qualified for job</option>
            <option value="cross">Cross-trained (2+ areas)</option>
          </Select>
      </div>

      <div className="no-print mb-3 flex flex-wrap items-center justify-between gap-2 text-sm">
        <p className="text-muted">
          <b className="text-ink">{rows.length}</b> of {data.employees.length} employees
          {filtersOn && (
            <button className="ml-3 font-semibold text-navy-700 underline" onClick={() => { setQ(""); setShift(""); setDept(""); setJob(""); setSup(""); setStatus(""); setLevel(""); setElig(""); }}>
              Clear filters
            </button>
          )}
        </p>
        {data.view && (
          <label className="flex items-center gap-2 font-medium">
            <input type="checkbox" className="h-4 w-4" checked={everyone} onChange={(e) => setEveryone(e.target.checked)} />
            Show everyone (including people with no {data.view.name} skills)
          </label>
        )}
      </div>

      <div className="mb-3 rounded-lg border border-line bg-white p-3"><StatusLegend /></div>

      <div className="matrix-scroll max-h-[calc(100vh-220px)] overflow-auto rounded-lg border border-line bg-white shadow-sm">
        <table className="table-fixed border-separate border-spacing-0 text-sm" aria-label="Skills matrix">
          <thead className="sticky top-0 z-20 bg-white">
            <tr>
              {FIXED.map((col) => (
                <th key={col.label} scope="col" style={fixedStyle(col)}
                  className={cn("sticky z-30 border-b border-line bg-white px-2 pb-2 text-left align-bottom text-xs font-bold uppercase tracking-wide text-muted", col.last && "border-r-2 border-r-navy-200")}>
                  {col.sort ? (
                    <button className="inline-flex items-center gap-1 uppercase" onClick={() => { if (sort === col.sort) setDesc(!desc); else { setSort(col.sort!); setDesc(false); } }}
                      aria-label={`Sort by ${col.label}`}>
                      {col.label} <ArrowUpDown className="h-3 w-3" aria-hidden />
                    </button>
                  ) : col.label}
                </th>
              ))}
              {data.columns.map((c) => (
                <th key={c.id} scope="col" className={cn("th-angled border-b border-line bg-white", compact ? "w-[36px]" : "w-[50px]")} title={`${c.code} – ${c.name}${c.critical ? " (critical skill)" : ""}`}>
                  <div><span>{c.critical && <span className="text-st-red" aria-label="critical">◆ </span>}{data.view ? c.name : c.name}</span></div>
                </th>
              ))}
              <th className="w-6 border-b border-line bg-white" aria-hidden />
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={6 + data.columns.length} className="px-6 py-10 text-center text-muted">No employees match these filters.</td></tr>
            )}
            {rows.map((e) => (
              <tr key={e.id} id={`row-${e.id}`} className={cn("group", data.focus === e.id && "bg-yellow-50")}>
                <td style={fixedStyle(FIXED[0])} className={cn(stickyTd, "font-mono text-xs")}>{e.number}</td>
                <td style={fixedStyle(FIXED[1])} className={stickyTd}>
                  <Link href={`/employees/${e.id}`} className="font-semibold text-navy-700 hover:underline">{e.name}</Link>
                  <div className="mt-0.5 flex flex-wrap gap-1">
                    {e.isSample && <SampleBadge />}
                    {e.eligible && <Badge tone="blue" title={`Meets system-checked requirements for ${e.nextLevel}`}>Promotion-ready</Badge>}
                  </div>
                </td>
                <td style={fixedStyle(FIXED[2])} className={cn(stickyTd, "text-center font-semibold")}>{e.shift}</td>
                <td style={fixedStyle(FIXED[3])} className={cn(stickyTd, "text-xs")}>{e.area}</td>
                <td style={fixedStyle(FIXED[4])} className={cn(stickyTd, "border-r-2 border-r-navy-200 text-xs")}>{e.job}</td>
                {e.cells.map((cell, i) => {
                  const col = data.columns[i];
                  const label = `${e.name} – ${col.name}: ${STATUS_META[cell.s].label}${cell.p ? " (pending verification)" : ""}${cell.n ? `, ${cell.n}` : ""}`;
                  return (
                    <td key={col.id} className={cn("border-b border-line px-0.5 py-1 text-center group-hover:bg-canvas", compact ? "w-[36px]" : "w-[50px]")}>
                      {data.view ? (
                        <button className="rounded focus-visible:outline focus-visible:outline-2 focus-visible:outline-navy-800" aria-label={`${label}. Open assessment`}
                          onClick={() => setOpen({ employeeId: e.id, competencyId: col.id })}>
                          <StatusChip status={cell.s} pending={cell.p} compact={compact} />
                        </button>
                      ) : (
                        <Link href={`/matrix?area=${col.code}&focus=${e.id}`} aria-label={`${label}. Open ${col.name} competencies`} className="inline-block rounded focus-visible:outline focus-visible:outline-2 focus-visible:outline-navy-800">
                          <StatusChip status={cell.s} pending={cell.p} compact={compact} detail={cell.n} />
                        </Link>
                      )}
                    </td>
                  );
                })}
                <td className="border-b border-line" />
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!data.view && (
        <p className="mt-2 text-xs text-muted">Department colour: blue = every competency Competent · green = all at least Trained · yellow = training under way · red = needs training · grey = not assessed / not applicable.</p>
      )}

      {open && (
        <AssessmentDialog
          employeeId={open.employeeId}
          competencyId={open.competencyId}
          onClose={() => setOpen(null)}
          onSaved={() => router.refresh()}
        />
      )}
    </div>
  );
}

const FIXED: { label: string; w: number; left: number; sort?: "name" | "number"; last?: boolean }[] = [
  { label: "ID", w: 88, left: 0, sort: "number" },
  { label: "Employee", w: 190, left: 88, sort: "name" },
  { label: "Shift", w: 54, left: 278 },
  { label: "Department", w: 120, left: 332 },
  { label: "Job title", w: 140, left: 452, last: true },
];
const fixedStyle = (c: { w: number; left: number }) => ({ left: c.left, width: c.w, minWidth: c.w, maxWidth: c.w });
const stickyTd = "sticky z-10 border-b border-line bg-white px-2 py-1.5 group-hover:bg-canvas";
const tabCls = (on: boolean) =>
  cn("whitespace-nowrap rounded-md px-3 py-2 text-sm font-semibold", on ? "bg-navy-800 text-white" : "bg-white text-navy-700 ring-1 ring-line hover:bg-navy-50");
