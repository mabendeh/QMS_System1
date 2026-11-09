import { NextResponse, type NextRequest } from "next/server";
import { AuthError, requireUser } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { audit } from "@/lib/audit";
import { canViewEmployee } from "@/lib/scope";
import { toPdf, toXlsx, type Report } from "@/lib/export/render";
import { employeeReport, gapsReport, matrixReport } from "@/lib/export/reports";
import { coverageReport, monthlyReport, promotionReport } from "@/lib/export/reports2";

export const dynamic = "force-dynamic";

export async function GET(req: NextRequest, { params }: { params: Promise<{ report: string }> }) {
  try {
    const { report } = await params;
    const sp = req.nextUrl.searchParams;
    const format = sp.get("format") === "pdf" ? "pdf" : "xlsx";
    const filters = { area: sp.get("area"), shift: sp.get("shift"), dept: sp.get("dept"), job: sp.get("job"), sup: sp.get("sup") };
    let r: Report;
    let u;
    switch (report) {
      case "matrix":
        u = await requireUser("export:data");
        r = await matrixReport(filters, u.name);
        break;
      case "gaps":
        u = await requireUser("export:data");
        r = await gapsReport(filters, u.name);
        break;
      case "employee": {
        u = await requireUser();
        const id = sp.get("id") ?? "";
        if (!canViewEmployee(u, id)) throw new AuthError("FORBIDDEN", "You cannot export this employee.");
        r = await employeeReport(id, u.name);
        break;
      }
      case "promotions":
        u = await requireUser("promotion:view");
        r = await promotionReport(u.name);
        break;
      case "coverage":
        u = await requireUser("coverage:view");
        r = await coverageReport(u.name);
        break;
      case "monthly":
        u = await requireUser("analytics:view");
        r = await monthlyReport(u.name);
        break;
      default:
        return NextResponse.json({ error: "Unknown report" }, { status: 404 });
    }
    const body = format === "pdf" ? await toPdf(r) : await toXlsx(r);
    await audit(prisma, u, "export", "Report", report, `${u.name} exported ${r.title} (${format.toUpperCase()})`);
    const name = `${r.title.replace(/[^\w\- ]+/g, "").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.${format}`;
    return new NextResponse(new Uint8Array(body), {
      headers: {
        "Content-Type": format === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${name}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (e) {
    if (e instanceof AuthError) return NextResponse.json({ error: e.message }, { status: e.code === "UNAUTHENTICATED" ? 401 : 403 });
    console.error(e);
    return NextResponse.json({ error: "Export failed" }, { status: 500 });
  }
}
