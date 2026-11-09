import { NextResponse } from "next/server";
import { AuthError, requireUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { canViewEmployee } from "@/lib/scope";
import { readUpload } from "@/lib/uploads";

export const dynamic = "force-dynamic";

export async function GET(_: Request, { params }: { params: Promise<{ id: string }> }) {
  try {
    const u = await requireUser();
    const { id } = await params;
    const ev = await prisma.evidence.findUnique({ where: { id } });
    if (!ev || !canViewEmployee(u, ev.employeeId)) return NextResponse.json({ error: "Not found" }, { status: 404 });
    const buf = await readUpload(ev.storedName);
    return new NextResponse(new Uint8Array(buf), {
      headers: {
        "Content-Type": ev.mimeType,
        "Content-Disposition": `inline; filename="${ev.originalName.replace(/"/g, "")}"`,
        "X-Content-Type-Options": "nosniff",
        "Content-Security-Policy": "default-src 'none'; img-src 'self'; style-src 'unsafe-inline'; sandbox",
        "Cache-Control": "private, no-store",
      },
    });
  } catch (e) {
    if (e instanceof AuthError) return NextResponse.json({ error: e.message }, { status: 401 });
    return NextResponse.json({ error: "File unavailable" }, { status: 404 });
  }
}
