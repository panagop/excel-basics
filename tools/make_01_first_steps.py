# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/01-first-steps.xlsx.  Run with:  uv run tools/make_01_first_steps.py"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(__file__).resolve().parent.parent / "files" / "01-first-steps.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="217346")
FORMULA_FILL = PatternFill("solid", fgColor="E8F3EC")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def table(ws, top, headers, rows, formulas, fill_formulas=True):
    """Write a table starting at row `top`; `formulas` maps column letter -> template with {r}."""
    for j, h in enumerate(headers, start=1):
        c = ws.cell(top, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
    for i, row in enumerate(rows, start=top + 1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            c.border = BOX
            if isinstance(v, float):
                c.number_format = "0.00"
        for col, tpl in formulas.items():
            c = ws[f"{col}{i}"]
            c.value = tpl.format(r=i) if fill_formulas else None
            c.fill = FORMULA_FILL if fill_formulas else INPUT_FILL
            c.number_format = "0.0000"
            c.border = BOX


wb = Workbook()

# ---- Example sheet ----
ws = wb.active
ws.title = "Example"
ws["A1"] = "Lesson 1 – Concrete beams: cross-section area and volume"
ws["A1"].font = Font(bold=True, size=14, color="217346")
ws["A2"] = "Click the green cells and look at the Formula Bar.  A = b·h,  V = A·L"
ws["A2"].font = Font(italic=True, color="595959")
table(
    ws, 4,
    ["Beam", "b (m)", "h (m)", "L (m)", "A (m²)", "V (m³)"],
    [["B1", 0.25, 0.50, 5.00], ["B2", 0.30, 0.60, 6.00], ["B3", 0.25, 0.45, 4.50]],
    {"E": "=B{r}*C{r}", "F": "=E{r}*D{r}"},
)

# ---- Exercise sheet ----
ex = wb.create_sheet("Exercise")
ex["A1"] = "Exercise – Concrete columns"
ex["A1"].font = Font(bold=True, size=14, color="217346")
ex["A2"] = "Fill in the yellow cells with formulas:  A = b·h  (m²)  and  V = A·H  (m³)."
ex["A2"].font = Font(italic=True, color="595959")
table(
    ex, 4,
    ["Column", "b (m)", "h (m)", "H (m)", "A (m²)", "V (m³)"],
    [["C1", 0.40, 0.40, 3.00], ["C2", 0.30, 0.50, 3.00],
     ["C3", 0.50, 0.50, 3.20], ["C4", 0.25, 0.60, 2.80]],
    {"E": "", "F": ""},
    fill_formulas=False,
)
ex["A10"] = "Check: the volume of column C1 must be 0.48 m³."
ex["A10"].font = Font(italic=True, color="595959")

for sheet in (ws, ex):
    sheet.column_dimensions["A"].width = 10
    for col in "BCDEF":
        sheet.column_dimensions[col].width = 11

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"Saved {OUT}")
