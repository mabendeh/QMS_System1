"""Shared styling and table helpers for the CHG Duncanville workbooks (openpyxl)."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

FONT = "Arial"
NAVY = "1B2A4A"
BURG = "8C1D40"
GRAY = "595959"
LGRAY = "F2F2F2"
MGRAY = "D9D9D9"
CALCHDR = "7F7F7F"
WHITE = "FFFFFF"
RAG = {
    "Green": ("E2F0D9", "1E5E20"),
    "Yellow": ("FFF2CC", "7F6000"),
    "Red": ("F8D7DA", "9C0006"),
    "Gray": ("EDEDED", "595959"),
}
DATE = "dd-mmm-yy"
PCT = "0%"

thin = Side(style="thin", color=MGRAY)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style="medium", color=NAVY))


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def L(n):
    return get_column_letter(n)


def title_band(ws, title, subtitle, last_col=16, legend=True):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 6
    for c in range(1, last_col + 1):
        ws.cell(1, c).fill = fill(BURG)
    ws["B2"] = title
    ws["B2"].font = f(16, True, NAVY)
    ws.row_dimensions[2].height = 26
    ws["B3"] = subtitle
    ws["B3"].font = f(9, False, GRAY, True)
    for c in range(1, last_col + 1):
        ws.cell(3, c).border = BOTTOM
    if legend:
        lc = min(max(last_col - 3, 5), 12)
        ws.cell(2, lc, "Input cell").font = f(8, False, GRAY)
        ws.cell(2, lc).fill = fill(WHITE)
        ws.cell(2, lc).border = BORDER
        ws.cell(2, lc + 1, "ƒ Formula – do not edit").font = f(8, False, GRAY)
        ws.cell(2, lc + 1).fill = fill(LGRAY)
        ws.cell(2, lc + 1).border = BORDER
    ws.column_dimensions["A"].width = 2


def define(wb, name, ref):
    dn = DefinedName(name, attr_text=ref)
    wb.defined_names[name] = dn


def qs(ws):
    return "'" + ws.title.replace("'", "''") + "'"


def dyn_list(wb, name, ws, col_letter, first_row, last_row, text_only=False):
    """Dynamic dropdown list (grows as rows are added, no formula edits needed)."""
    s = qs(ws)
    cnt = (f'COUNTIF({s}!${col_letter}${first_row}:${col_letter}${last_row},"?*")' if text_only
           else f"COUNTA({s}!${col_letter}${first_row}:${col_letter}${last_row})")
    define(wb, name, f"OFFSET({s}!${col_letter}${first_row},0,0,MAX(1,{cnt}),1)")


class Cols:
    """Ordered column spec -> letters, with {Key} templating to cell refs."""

    def __init__(self, spec, first_col=2):
        self.spec = spec
        self.letter = {}
        for i, c in enumerate(spec):
            self.letter[c["key"]] = L(first_col + i)
        self.first_col = first_col

    def tpl(self, s, r):
        out = s
        # longest keys first to avoid partial replacement
        for k in sorted(self.letter, key=len, reverse=True):
            out = out.replace("{" + k + "}", f"${self.letter[k]}{r}")
        return out


def build_table(ws, wb, cols, header_row, nrows, table_name, data=None, name_prefix=None,
                dv_map=None, freeze_col=None, row_height=None, style="TableStyleLight1"):
    """Writes a formatted Excel Table with input (white) and calculated (gray) columns.

    cols.spec entries: key, header, width, calc(bool), formula(str template), fmt, dv(list name or
    ('date'|'pct'|'num')), note(str), wrap(bool), default
    """
    data = data or []
    first = header_row + 1
    last = header_row + nrows
    for i, c in enumerate(cols.spec):
        col = cols.first_col + i
        letter = L(col)
        h = ws.cell(header_row, col, ("ƒ " if c.get("calc") else "") + c["header"])
        h.font = f(9, True, WHITE)
        h.fill = fill(CALCHDR if c.get("calc") else NAVY)
        h.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        h.border = BORDER
        if c.get("note"):
            cm = Comment(c["note"], "DunGPS System")
            cm.width, cm.height = 300, 120
            h.comment = cm
        ws.column_dimensions[letter].width = c.get("width", 14)
        for r in range(first, last + 1):
            cell = ws.cell(r, col)
            idx = r - first
            if c.get("calc"):
                cell.value = cols.tpl(c["formula"], r)
                cell.fill = fill(LGRAY)
                cell.font = f(9, False, "262626")
            else:
                v = None
                if idx < len(data):
                    v = data[idx].get(c["key"], c.get("default"))
                if v is not None:
                    cell.value = v
                cell.font = f(9)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=c.get("wrap", False),
                                       horizontal=c.get("align", "left" if not c.get("fmt") else "center"))
            if c.get("fmt"):
                cell.number_format = c["fmt"]
        rng = f"{letter}{first}:{letter}{last}"
        dv = c.get("dv")
        if dv:
            if dv == "date":
                d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)",
                                   formula2="DATE(2040,12,31)", allow_blank=True,
                                   errorTitle="Invalid date", error="Enter a date (e.g. 15-Jan-27).")
            elif dv == "pct":
                d = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                                   allow_blank=True, errorTitle="Invalid %",
                                   error="Enter a value from 0% to 100%.")
            elif dv == "num":
                d = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                                   allow_blank=True, error="Enter a number ≥ 0.")
            elif dv == "int15":
                d = DataValidation(type="whole", operator="between", formula1="0", formula2="5",
                                   allow_blank=True, error="Enter a whole score 0–5.")
            else:
                d = DataValidation(type="list", formula1=f"={dv}", allow_blank=True,
                                   showErrorMessage=True, errorTitle="Select from list",
                                   error="Please pick a value from the dropdown list. Lists are "
                                         "maintained on 'Configuration & Dropdown Lists' / directory sheets.")
            d.showErrorMessage = True
            ws.add_data_validation(d)
            d.add(rng)
        if name_prefix:
            define(wb, f"{name_prefix}_{c['key']}", f"{qs(ws)}!${letter}${first}:${letter}${last}")
    if row_height:
        for r in range(first, last + 1):
            ws.row_dimensions[r].height = row_height
    ws.row_dimensions[header_row].height = 36
    ref = f"{L(cols.first_col)}{header_row}:{L(cols.first_col + len(cols.spec) - 1)}{last}"
    t = Table(displayName=table_name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=False)
    ws.add_table(t)
    if freeze_col:
        ws.freeze_panes = f"{freeze_col}{first}"
    return first, last


def rag_cf(ws, rng):
    for k, (bg, fg) in RAG.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{k}"'],
                                                      fill=fill(bg), font=Font(name=FONT, color=fg, bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"N/A"'],
                                                  fill=fill("EDEDED"), font=Font(name=FONT, color="7F7F7F")))


def text_cf(ws, rng, mapping):
    """mapping: {text: 'Green'|'Yellow'|'Red'|'Gray'} exact match."""
    for txt, k in mapping.items():
        bg, fg = RAG[k]
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                                      fill=fill(bg), font=Font(name=FONT, color=fg)))


def bar_cf(ws, rng, color=NAVY):
    ws.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num",
                                                   end_value=1, color=color, showValue=True))


def hdr_row(ws, row, col, labels, widths=None, calc=False):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = f(9, True, WHITE)
        c.fill = fill(CALCHDR if calc else NAVY)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        c.border = BORDER


def section(ws, row, col, text, span=6):
    c = ws.cell(row, col, text)
    c.font = f(11, True, NAVY)
    for i in range(span):
        ws.cell(row, col + i).border = Border(bottom=Side(style="thin", color=BURG))


def calc_cell(ws, row, col, formula, fmt=None, bold=False, align="center", size=9, color="262626"):
    c = ws.cell(row, col, formula)
    c.font = f(size, bold, color)
    c.border = BORDER
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(align == "left"))
    if fmt:
        c.number_format = fmt
    return c


def tile(ws, row, col, label, formula, fmt, sub_formula=None, width=2, accent=NAVY):
    """KPI tile: label / big value / sub text, merged across `width` columns."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + width - 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + width - 1)
    ws.merge_cells(start_row=row + 2, start_column=col, end_row=row + 2, end_column=col + width - 1)
    a = ws.cell(row, col, label)
    a.font = f(8, True, WHITE)
    a.fill = fill(accent)
    a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    b = ws.cell(row + 1, col, formula)
    b.font = f(18, True, NAVY)
    b.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        b.number_format = fmt
    c = ws.cell(row + 2, col, sub_formula)
    c.font = f(8, False, GRAY)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(row, row + 3):
        for cc in range(col, col + width):
            ws.cell(r, cc).border = BORDER
            if r > row:
                ws.cell(r, cc).fill = fill(WHITE)
    return b


def style_chart(ch, title, ytitle=None, xtitle=None, w=16, h=7.5, legend_pos="b"):
    ch.title = title
    ch.roundedCorners = False
    ch.width, ch.height = w, h
    ch.visible_cells_only = False
    if ch.legend is not None:
        ch.legend.position = legend_pos
        ch.legend.overlay = False
    try:
        ch.title.overlay = False
    except Exception:
        pass
    if hasattr(ch, "x_axis") and ch.tagname not in ("pieChart", "doughnutChart"):
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        if ytitle:
            ch.y_axis.title = ytitle
        if xtitle:
            ch.x_axis.title = xtitle
        ch.y_axis.majorGridlines = None


def color_series(s, hexcolor, line=False):
    if line:
        s.graphicalProperties.line.solidFill = hexcolor
        s.graphicalProperties.line.width = 28000
        s.smooth = False
    else:
        s.graphicalProperties.solidFill = hexcolor
        s.graphicalProperties.line.solidFill = hexcolor


def protect(ws, unlocked_ranges=()):
    from openpyxl.styles import Protection
    for rng in unlocked_ranges:
        for row in ws[rng]:
            for c in (row if isinstance(row, tuple) else (row,)):
                c.protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.autoFilter = False
    ws.protection.sort = False


def landscape(ws, area=None, one_page=True):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if one_page else 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    if area:
        ws.print_area = area
