"use client";
import { useRouter } from "next/navigation";
import { useEffect, useState } from "react";
import { Input, Select } from "@/components/ui";

export function EmployeeSearch({ q, shift }: { q: string; shift: string }) {
  const r = useRouter();
  const [v, setV] = useState(q);
  const [s, setS] = useState(shift);
  useEffect(() => {
    const t = setTimeout(() => r.replace(`/employees?${new URLSearchParams({ ...(v ? { q: v } : {}), ...(s ? { shift: s } : {}) })}`), 250);
    return () => clearTimeout(t);
  }, [v, s, r]);
  return (
    <div className="flex flex-wrap gap-2">
      <Input aria-label="Search employees" placeholder="Search name or employee ID" value={v} onChange={(e) => setV(e.target.value)} className="max-w-md" />
      <Select aria-label="Shift" value={s} onChange={(e) => setS(e.target.value)} className="max-w-[160px]">
        <option value="">All shifts</option>
        {["A", "B", "C"].map((x) => <option key={x} value={x}>Shift {x}</option>)}
      </Select>
    </div>
  );
}
