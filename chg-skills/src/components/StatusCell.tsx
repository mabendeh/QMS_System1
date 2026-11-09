import * as React from "react";
import { STATUS_META, type DisplayStatus } from "@/lib/domain/status";
import { cn } from "@/lib/utils";

/** One colour-coded matrix cell. Colour is never the only signal: it always shows a
 *  readable abbreviation and has an accessible label + tooltip. */
export function StatusChip({
  status, pending = false, compact = false, detail, className,
}: { status: DisplayStatus; pending?: boolean; compact?: boolean; detail?: string; className?: string }) {
  const m = STATUS_META[status];
  const label = `${m.label}${pending ? " – pending verification" : ""}${detail ? ` – ${detail}` : ""}`;
  return (
    <span
      role="img"
      aria-label={label}
      title={label}
      data-status={status}
      data-pending={pending ? "true" : undefined}
      className={cn(
        "relative inline-flex select-none items-center justify-center rounded font-bold leading-none",
        compact ? "h-6 w-8 text-[10px]" : "h-8 w-11 text-xs",
        pending && "ring-2 ring-offset-1 ring-navy-800 [background-image:repeating-linear-gradient(135deg,transparent_0_4px,rgba(255,255,255,.35)_4px_7px)]",
        className,
      )}
      style={{ backgroundColor: m.color, color: m.text }}
    >
      {m.short}
      {pending && <span aria-hidden className="absolute -right-1 -top-1 rounded-full bg-navy-800 px-1 text-[9px] text-white">P</span>}
    </span>
  );
}

export function StatusLegend({ className }: { className?: string }) {
  const order: DisplayStatus[] = ["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "NOT_ASSESSED", "EXPIRED"];
  return (
    <ul className={cn("flex flex-wrap gap-x-5 gap-y-2 text-sm", className)} aria-label="Status legend">
      {order.map((s) => (
        <li key={s} className="flex items-center gap-2">
          <StatusChip status={s} compact />
          <span>
            <span className="font-semibold">{STATUS_META[s].label}</span>
            <span className="hidden text-muted xl:inline"> – {STATUS_META[s].description}</span>
          </span>
        </li>
      ))}
      <li className="flex items-center gap-2">
        <StatusChip status="TRAINED" pending compact />
        <span className="font-semibold">Pending verification</span>
      </li>
    </ul>
  );
}
