import "server-only";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
import { STATUS_META, type DisplayStatus } from "@/lib/domain/status";

export type Cell = string | number | null | { status: DisplayStatus; pending?: boolean; text?: string };
export interface Sheet {
  name: string;
  columns: { header: string; width?: number; rotate?: boolean }[];
  rows: Cell[][];
  frozenCols?: number;
}
export interface Report {
  title: string;
  subtitle: string;
  sheets: Sheet[];
  legend?: boolean;
  notes?: string[];
}

const cellText = (c: Cell) =>
  c == null ? "" : typeof c === "object" ? `${c.text ?? STATUS_META[c.status].label}${c.pending ? " (pending)" : ""}` : String(c);

export async function toXlsx(r: Report): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "CHG Skills & Career Progression";
  wb.created = new Date();
  for (const s of r.sheets) {
    const ws = wb.addWorksheet(s.name.slice(0, 31), { views: [{ state: "frozen", xSplit: s.frozenCols ?? 0, ySplit: 4 }] });
    ws.getCell("A1").value = r.title;
    ws.getCell("A1").font = { name: "Arial", size: 14, bold: true, color: { argb: "FF14284B" } };
    ws.getCell("A2").value = r.subtitle;
    ws.getCell("A2").font = { name: "Arial", size: 9, italic: true, color: { argb: "FF5D6B7E" } };
    const head = ws.getRow(4);
    s.columns.forEach((c, i) => {
      const cell = head.getCell(i + 1);
      cell.value = c.header;
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" }, size: 9 };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF14284B" } };
      cell.alignment = { vertical: "bottom", horizontal: c.rotate ? "center" : "left", wrapText: !c.rotate, textRotation: c.rotate ? 60 : 0 };
      ws.getColumn(i + 1).width = c.width ?? 14;
    });
    head.height = s.columns.some((c) => c.rotate) ? 130 : 30;
    s.rows.forEach((row, ri) => {
      const xr = ws.getRow(5 + ri);
      row.forEach((c, ci) => {
        const cell = xr.getCell(ci + 1);
        const thin = { style: "thin" as const, color: { argb: "FFD9DFE7" } };
        cell.border = { top: thin, left: thin, bottom: thin, right: thin };
        cell.font = { name: "Arial", size: 9 };
        if (c && typeof c === "object") {
          const m = STATUS_META[c.status];
          cell.value = `${m.short}${c.pending ? "*" : ""}`;
          cell.note = cellText(c);
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: m.argb } };
          cell.font = { name: "Arial", size: 9, bold: true, color: { argb: m.fontArgb } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else {
          cell.value = c ?? "";
        }
      });
    });
    let y = 6 + s.rows.length;
    if (r.legend) {
      ws.getCell(`A${y}`).value = "Legend";
      ws.getCell(`A${y}`).font = { name: "Arial", bold: true };
      for (const st of ["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "EXPIRED", "NOT_ASSESSED"] as DisplayStatus[]) {
        y++;
        const m = STATUS_META[st];
        const a = ws.getCell(`A${y}`);
        a.value = m.short;
        a.fill = { type: "pattern", pattern: "solid", fgColor: { argb: m.argb } };
        a.font = { name: "Arial", bold: true, color: { argb: m.fontArgb } };
        a.alignment = { horizontal: "center" };
        ws.getCell(`B${y}`).value = `${m.label} – ${m.description}`;
      }
      y++;
      ws.getCell(`B${y}`).value = "* = change pending verification";
    }
    for (const n of r.notes ?? []) {
      y += 2;
      ws.getCell(`A${y}`).value = n;
      ws.getCell(`A${y}`).font = { name: "Arial", size: 9, italic: true, color: { argb: "FF5D6B7E" } };
    }
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

export async function toPdf(r: Report): Promise<Buffer> {
  const doc = new PDFDocument({ size: "LETTER", layout: "landscape", margin: 28, info: { Title: r.title } });
  const chunks: Buffer[] = [];
  doc.on("data", (c: Buffer) => chunks.push(c));
  const done = new Promise<Buffer>((res) => doc.on("end", () => res(Buffer.concat(chunks))));
  const pageW = doc.page.width - 56;

  r.sheets.forEach((s, si) => {
    if (si > 0) doc.addPage();
    doc.font("Helvetica-Bold").fontSize(14).fillColor("#14284B").text(r.title + (r.sheets.length > 1 ? ` – ${s.name}` : ""));
    doc.font("Helvetica").fontSize(8).fillColor("#5D6B7E").text(r.subtitle);
    doc.moveDown(0.5);
    const statusCols = s.columns.filter((c) => c.rotate).length;
    const textCols = s.columns.length - statusCols;
    const statusW = statusCols ? Math.max(12, Math.min(34, (pageW * 0.62) / statusCols)) : 0;
    const textTotal = pageW - statusW * statusCols;
    const weights = s.columns.map((c) => (c.rotate ? 0 : c.width ?? 14));
    const wSum = weights.reduce((a, b) => a + b, 0) || 1;
    const widths = s.columns.map((c, i) => (c.rotate ? statusW : (textTotal * weights[i]) / wSum));
    const rowH = statusCols > 30 ? 11 : 14;
    const fs = statusCols > 30 ? 5.5 : 7;
    const headH = statusCols ? 90 : 18;

    const drawHead = () => {
      let x = 28;
      const y = doc.y;
      s.columns.forEach((c, i) => {
        doc.rect(x, y, widths[i], headH).fill("#14284B");
        doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(fs);
        if (c.rotate) {
          doc.save().rotate(-90, { origin: [x + widths[i] / 2 + 2, y + headH - 3] });
          doc.text(c.header.slice(0, 40), x + widths[i] / 2 + 2, y + headH - 3, { width: headH - 6, lineBreak: false });
          doc.restore();
        } else {
          doc.text(c.header, x + 3, y + headH - fs - 4, { width: widths[i] - 6, lineBreak: false, ellipsis: true });
        }
        x += widths[i];
      });
      doc.y = y + headH;
    };
    drawHead();
    for (const row of s.rows) {
      if (doc.y + rowH > doc.page.height - 40) {
        doc.addPage();
        drawHead();
      }
      let x = 28;
      const y = doc.y;
      row.forEach((c, i) => {
        if (c && typeof c === "object") {
          const m = STATUS_META[c.status];
          doc.rect(x + 0.5, y + 0.5, widths[i] - 1, rowH - 1).fill(m.color);
          doc.fillColor(m.text).font("Helvetica-Bold").fontSize(fs).text(`${m.short}${c.pending ? "*" : ""}`, x, y + (rowH - fs) / 2, { width: widths[i], align: "center", lineBreak: false });
        } else {
          doc.rect(x, y, widths[i], rowH).strokeColor("#D9DFE7").lineWidth(0.5).stroke();
          doc.fillColor("#1A2433").font("Helvetica").fontSize(fs).text(cellText(c), x + 3, y + (rowH - fs) / 2, { width: widths[i] - 6, lineBreak: false, ellipsis: true });
        }
        x += widths[i];
      });
      doc.y = y + rowH;
    }
    void textCols;
  });
  if (r.legend) {
    doc.moveDown(1);
    let x = 28;
    const y = Math.min(doc.y, doc.page.height - 50);
    for (const st of ["NEEDS_TRAINING", "IN_TRAINING", "TRAINED", "COMPETENT", "EXPIRED", "NOT_ASSESSED"] as DisplayStatus[]) {
      const m = STATUS_META[st];
      doc.rect(x, y, 22, 12).fill(m.color);
      doc.fillColor(m.text).font("Helvetica-Bold").fontSize(7).text(m.short, x, y + 3, { width: 22, align: "center" });
      doc.fillColor("#1A2433").font("Helvetica").fontSize(7).text(m.label, x + 26, y + 3, { lineBreak: false });
      x += 26 + doc.widthOfString(m.label) + 14;
    }
    doc.fillColor("#5D6B7E").text("* pending verification", x, y + 3, { lineBreak: false });
  }
  for (const n of r.notes ?? []) {
    doc.moveDown(1.2);
    doc.x = 28;
    doc.fillColor("#5D6B7E").font("Helvetica-Oblique").fontSize(7).text(n, 28, undefined, { width: pageW });
  }
  doc.end();
  return done;
}
