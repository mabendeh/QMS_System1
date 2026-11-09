"use server";
import { redirect } from "next/navigation";
import { z } from "zod";
import { login, logout } from "@/lib/auth";

const schema = z.object({ email: z.string().email().max(200), password: z.string().min(1).max(200) });

export async function loginAction(_: unknown, form: FormData): Promise<{ error?: string }> {
  const parsed = schema.safeParse({ email: form.get("email"), password: form.get("password") });
  if (!parsed.success) return { error: "Enter your email and password." };
  const res = await login(parsed.data.email, parsed.data.password);
  if (!res.ok) return { error: res.error };
  const next = String(form.get("next") ?? "/");
  redirect(next.startsWith("/") && !next.startsWith("//") ? next : "/");
}

export async function logoutAction() {
  await logout();
  redirect("/login");
}
