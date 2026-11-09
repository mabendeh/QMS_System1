import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function fmtDate(d: Date | string | null | undefined) {
  if (!d) return "—";
  const x = new Date(d);
  return x.toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric", timeZone: "America/Chicago" });
}

export function fmtDateTime(d: Date | string | null | undefined) {
  if (!d) return "—";
  return new Date(d).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short", timeZone: "America/Chicago" });
}

export function fullName(e: { firstName: string; lastName: string }) {
  return `${e.firstName} ${e.lastName}`;
}

export function pct(n: number, d: number) {
  return d ? Math.round((n / d) * 100) : 0;
}
