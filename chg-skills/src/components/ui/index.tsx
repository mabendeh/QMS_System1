import * as React from "react";
import { cn } from "@/lib/utils";

export function Button({
  variant = "primary", size = "md", className, ...props
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: "primary" | "secondary" | "ghost" | "danger" | "success"; size?: "sm" | "md" | "lg" }) {
  return (
    <button
      className={cn(
        "inline-flex items-center justify-center gap-2 rounded-md font-semibold transition-colors focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-navy-600 disabled:cursor-not-allowed disabled:opacity-50",
        size === "sm" && "h-8 px-3 text-sm",
        size === "md" && "h-10 px-4 text-sm",
        size === "lg" && "h-12 px-5 text-base",
        variant === "primary" && "bg-navy-800 text-white hover:bg-navy-700",
        variant === "secondary" && "border border-line bg-white text-ink hover:bg-canvas",
        variant === "ghost" && "text-navy-700 hover:bg-navy-50",
        variant === "danger" && "bg-st-red text-white hover:bg-red-800",
        variant === "success" && "bg-st-blue text-white hover:bg-blue-800",
        className,
      )}
      {...props}
    />
  );
}

export function Card({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn("rounded-lg border border-line bg-white shadow-sm", className)} {...props} />;
}

export function CardHeader({ title, subtitle, action, className }: { title: React.ReactNode; subtitle?: React.ReactNode; action?: React.ReactNode; className?: string }) {
  return (
    <div className={cn("flex flex-wrap items-start justify-between gap-3 border-b border-line px-5 py-4", className)}>
      <div className="min-w-0">
        <h2 className="text-base font-bold text-ink">{title}</h2>
        {subtitle && <p className="mt-0.5 text-sm text-muted">{subtitle}</p>}
      </div>
      {action}
    </div>
  );
}

export function Badge({ tone = "gray", className, ...props }: React.HTMLAttributes<HTMLSpanElement> & { tone?: "gray" | "navy" | "red" | "yellow" | "green" | "blue" | "amber" }) {
  return (
    <span
      className={cn(
        "inline-flex items-center gap-1 whitespace-nowrap rounded px-2 py-0.5 text-xs font-semibold",
        tone === "gray" && "bg-gray-100 text-gray-700",
        tone === "navy" && "bg-navy-50 text-navy-700",
        tone === "red" && "bg-red-50 text-red-800",
        tone === "yellow" && "bg-yellow-50 text-yellow-900",
        tone === "amber" && "bg-amber-50 text-amber-900",
        tone === "green" && "bg-green-50 text-green-800",
        tone === "blue" && "bg-blue-50 text-blue-800",
        className,
      )}
      {...props}
    />
  );
}

export const SampleBadge = () => <Badge tone="amber" title="Fictional demonstration record">Sample</Badge>;

export function Label({ className, ...props }: React.LabelHTMLAttributes<HTMLLabelElement>) {
  return <label className={cn("text-sm font-semibold text-ink", className)} {...props} />;
}

const field = "w-full rounded-md border border-line bg-white px-3 text-base text-ink shadow-sm focus:border-navy-600 focus:outline-none focus:ring-2 focus:ring-navy-200 disabled:bg-canvas";

export const Input = React.forwardRef<HTMLInputElement, React.InputHTMLAttributes<HTMLInputElement>>(function Input({ className, ...p }, ref) {
  return <input ref={ref} className={cn(field, "h-11", className)} {...p} />;
});

export const Select = React.forwardRef<HTMLSelectElement, React.SelectHTMLAttributes<HTMLSelectElement>>(function Select({ className, ...p }, ref) {
  return <select ref={ref} className={cn(field, "h-11 pr-8", className)} {...p} />;
});

export const Textarea = React.forwardRef<HTMLTextAreaElement, React.TextareaHTMLAttributes<HTMLTextAreaElement>>(function Textarea({ className, ...p }, ref) {
  return <textarea ref={ref} className={cn(field, "min-h-[84px] py-2", className)} {...p} />;
});

export function Field({ label, hint, children, htmlFor }: { label: string; hint?: string; children: React.ReactNode; htmlFor?: string }) {
  return (
    <div className="grid gap-1.5">
      <Label htmlFor={htmlFor}>{label}</Label>
      {children}
      {hint && <p className="text-xs text-muted">{hint}</p>}
    </div>
  );
}

export function Alert({ tone = "info", title, children }: { tone?: "info" | "error" | "success" | "warn"; title?: string; children?: React.ReactNode }) {
  return (
    <div
      role={tone === "error" ? "alert" : "status"}
      className={cn(
        "rounded-md border px-4 py-3 text-sm",
        tone === "info" && "border-navy-100 bg-navy-50 text-navy-800",
        tone === "error" && "border-red-200 bg-red-50 text-red-900",
        tone === "success" && "border-blue-200 bg-blue-50 text-blue-900",
        tone === "warn" && "border-amber-200 bg-amber-50 text-amber-900",
      )}
    >
      {title && <p className="font-bold">{title}</p>}
      {children}
    </div>
  );
}

export function PageHeader({ title, subtitle, actions }: { title: string; subtitle?: React.ReactNode; actions?: React.ReactNode }) {
  return (
    <div className="mb-5 flex flex-wrap items-end justify-between gap-3">
      <div>
        <h1 className="text-2xl font-bold tracking-tight text-ink">{title}</h1>
        {subtitle && <p className="mt-1 text-sm text-muted">{subtitle}</p>}
      </div>
      {actions && <div className="flex flex-wrap gap-2">{actions}</div>}
    </div>
  );
}

export function Stat({ label, value, hint, tone = "navy", href }: { label: string; value: React.ReactNode; hint?: string; tone?: "navy" | "red" | "yellow" | "green" | "blue" | "gray"; href?: string }) {
  const inner = (
    <>
      <p className="text-xs font-bold uppercase tracking-wide text-muted">{label}</p>
      <p className={cn("mt-1 text-3xl font-bold tabular-nums",
        tone === "navy" && "text-navy-800", tone === "red" && "text-st-red", tone === "yellow" && "text-yellow-700",
        tone === "green" && "text-st-green", tone === "blue" && "text-st-blue", tone === "gray" && "text-muted")}>{value}</p>
      {hint && <p className="mt-1 text-xs text-muted">{hint}</p>}
    </>
  );
  return href ? (
    <a href={href} className="block rounded-lg border border-line bg-white p-4 shadow-sm hover:border-navy-600">{inner}</a>
  ) : (
    <div className="rounded-lg border border-line bg-white p-4 shadow-sm">{inner}</div>
  );
}

export function Empty({ title, children }: { title: string; children?: React.ReactNode }) {
  return (
    <div className="px-6 py-10 text-center">
      <p className="font-semibold text-ink">{title}</p>
      {children && <div className="mt-1 text-sm text-muted">{children}</div>}
    </div>
  );
}

export function Table({ className, ...p }: React.TableHTMLAttributes<HTMLTableElement>) {
  return (
    <div className="overflow-x-auto">
      <table className={cn("w-full border-collapse text-sm", className)} {...p} />
    </div>
  );
}
export const Th = ({ className, ...p }: React.ThHTMLAttributes<HTMLTableCellElement>) => (
  <th className={cn("border-b border-line bg-canvas px-3 py-2 text-left text-xs font-bold uppercase tracking-wide text-muted", className)} {...p} />
);
export const Td = ({ className, ...p }: React.TdHTMLAttributes<HTMLTableCellElement>) => (
  <td className={cn("border-b border-line px-3 py-2 align-top", className)} {...p} />
);
