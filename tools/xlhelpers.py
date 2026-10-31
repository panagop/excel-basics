"""Shared styles and helpers for the lesson workbooks."""

from pathlib import Path

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FILES = Path(__file__).resolve().parent.parent / "files"

GREEN = "217346"
HEADER_FILL = PatternFill("solid", fgColor=GREEN)
FORMULA_FILL = PatternFill("solid", fgColor="E8F3EC")  # cells with a formula (examples)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")  # cells the student fills in
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def title(ws, text, subtitle=None, row=1):
    ws.cell(row, 1, text).font = Font(bold=True, size=14, color=GREEN)
    if subtitle:
        note(ws, row + 1, subtitle)


def section(ws, row, text):
    ws.cell(row, 1, text).font = Font(bold=True, size=12, color=GREEN)


def note(ws, row, text, col=1):
    ws.cell(row, col, text).font = Font(italic=True, color="595959")


def as_text(cell, text):
    """Store text that looks like a formula (e.g. '=2+3*4') without Excel evaluating it."""
    cell.value = text
    cell.data_type = "s"
    cell.quotePrefix = True


def table(ws, top, headers, rows, formulas=None, student=False, fmt="0.00", formula_fmt="0.0000", col0=1):
    """Write a table with its header at row `top`, starting at column `col0`.

    `formulas` maps column letter -> template with {r} for the row, e.g. {"E": "=B{r}*C{r}"}.
    With student=True those cells are left empty and coloured yellow for the student to fill in.
    """
    for j, h in enumerate(headers, start=col0):
        c = ws.cell(top, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BOX
    for i, row in enumerate(rows, start=top + 1):
        for j, v in enumerate(row, start=col0):
            c = ws.cell(i, j, v)
            c.border = BOX
            if isinstance(v, (int, float)):
                c.number_format = fmt
        for col, tpl in (formulas or {}).items():
            c = ws[f"{col}{i}"]
            c.value = None if student else tpl.format(r=i)
            c.fill = INPUT_FILL if student else FORMULA_FILL
            c.number_format = formula_fmt
            c.border = BOX
    return top + len(rows)


def input_cell(cell, fmt="0.00"):
    cell.fill = INPUT_FILL
    cell.border = BOX
    cell.number_format = fmt


def widths(ws, **cols):
    for col, w in cols.items():
        ws.column_dimensions[col].width = w


def save(wb, name):
    FILES.mkdir(exist_ok=True)
    out = FILES / name
    wb.save(out)
    print(f"Saved {out}")
