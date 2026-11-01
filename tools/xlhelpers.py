"""Shared styles and helpers for the lesson workbooks."""

from pathlib import Path

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_EMU

ROOT = Path(__file__).resolve().parent.parent
FILES = ROOT / "files"
LOGOS = [ROOT / "assets/img/ihu-logo.png", ROOT / "assets/img/civil-ihu-logo.png"]
LOGO_H = 40  # px; row 1 is made tall enough to hold the logos
LOGO_GAP = 16  # px between the logos, and between the logos and the title

GREEN = "217346"
HEADER_FILL = PatternFill("solid", fgColor=GREEN)
FORMULA_FILL = PatternFill("solid", fgColor="E8F3EC")  # cells with a formula (examples)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")  # cells the student fills in
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def logos(ws):
    """University and Department logos at the top left of the sheet, inside row 1."""
    x = 6
    for path in LOGOS:
        img = Image(path)
        w = round(LOGO_H * img.width / img.height)
        img.width, img.height = w, LOGO_H
        img.anchor = OneCellAnchor(
            _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(x), rowOff=pixels_to_EMU(8)),
            ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(LOGO_H)))
        ws.add_image(img)
        x += w + LOGO_GAP
    return x


def title(ws, text, subtitle=None, row=1):
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, size=14, color=GREEN)
    if row == 1:
        ws.logos_end = logos(ws)  # the title is moved right of the logos in save(), once column widths are known
        ws.row_dimensions[1].height = 42
        c.alignment = Alignment(vertical="center")
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


def _title_beside_logos(ws):
    """Move the title in A1 to the first column that starts right of the logos.

    Column widths (in characters) convert to pixels as 7·w + 5 in Excel; unset columns are 64 px.
    """
    end = getattr(ws, "logos_end", None)
    if end is None or ws["A1"].value is None:
        return
    left, col = 0, 1
    while left < end:
        dim = ws.column_dimensions.get(ws.cell(1, col).column_letter)
        left += round(7 * dim.width + 5) if dim is not None and dim.width else 64
        col += 1
    target = ws.cell(1, col)
    target.value, target.font, target.alignment = ws["A1"].value, ws["A1"].font.copy(), ws["A1"].alignment.copy()
    ws["A1"].value = None


def save(wb, name):
    for ws in wb.worksheets:
        _title_beside_logos(ws)
    FILES.mkdir(exist_ok=True)
    out = FILES / name
    wb.save(out)
    print(f"Saved {out}")
