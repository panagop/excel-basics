# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/09-sort-filter.xlsx.  Run with:  uv run tools/make_09_sort_filter.py"""

import random

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

from xlhelpers import (BOX, FORMULA_FILL, HEADER_FILL, as_text, input_cell, note, save, section,
                       title, widths)

HEADERS = ["Κωδικός", "Όροφος", "Στοιχείο", "Δ1 (m)", "Δ2 (m)", "Δ3 (m)", "V (m³)"]
TYPES = ["Υποστύλωμα", "Δοκός", "Πλάκα"]


def take_off(floors, columns, beams, slabs, seed):
    """Element list (code, floor, type, d1, d2, d3) in shuffled order.

    columns(floor) -> list of (b, h, H); beams / slabs -> lists of (b, h, L) / (t, Lx, Ly).
    """
    rows = []
    for fl in floors:
        rows += [(f"Κ{i}", fl, "Υποστύλωμα", *d) for i, d in enumerate(columns(fl), start=1)]
        rows += [(f"Δ{i}", fl, "Δοκός", *d) for i, d in enumerate(beams, start=1)]
        rows += [(f"Π{i}", fl, "Πλάκα", *d) for i, d in enumerate(slabs, start=1)]
    random.Random(seed).shuffle(rows)
    return rows


def write_table(ws, top, rows, name, student=False):
    """Element list as an Excel Table; V = Δ1·Δ2·Δ3 (left empty for the student)."""
    for j, h in enumerate(HEADERS, start=1):
        ws.cell(top, j, h)
    for i, row in enumerate(rows, start=top + 1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            if j >= 4:
                c.number_format = "0.00"
        v = ws.cell(i, 7)
        if student:
            input_cell(v, "0.000")
        else:
            v.value = f"=D{i}*E{i}*F{i}"
            v.number_format = "0.000"
    last = top + len(rows)
    if not student:  # the student converts the range to a Table (Ctrl+T) as part of the exercise
        t = Table(displayName=name, ref=f"A{top}:G{last}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
        ws.add_table(t)
    else:
        for r in range(top, last + 1):
            for c in range(1, 8):
                ws.cell(r, c).border = BOX
        for c in range(1, 8):
            ws.cell(top, c).font = Font(bold=True, color="FFFFFF")
            ws.cell(top, c).fill = HEADER_FILL
    return last


def summary_header(ws, row, col, labels):
    for j, h in enumerate(labels, start=col):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.border = BOX


def out(ws, row, col, formula, fmt="0.000", student=False):
    c = ws.cell(row, col)
    if student:
        input_cell(c, fmt)
    else:
        c.value = formula
        c.fill = FORMULA_FILL
        c.number_format = fmt
        c.border = BOX


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 9 – Προμέτρηση σκυροδέματος: ταξινόμηση, φίλτρα, Πίνακες, SUMIF",
      "Ο πίνακας A4:G34 είναι «Πίνακας» του Excel. Ταξινομήστε και φιλτράρετε με τα βελάκια των επικεφαλίδων.")
ws["A3"] = "Δ1·Δ2·Δ3: υποστύλωμα b·h·H, δοκός b·h·L, πλάκα t·Lx·Ly"
ws["A3"].font = Font(italic=True, color="595959")
rows = take_off(
    floors=[0, 1, 2],
    columns=lambda fl: [(0.45, 0.45, 3.50)] * 4 if fl == 0 else [(0.40, 0.40, 3.00)] * 4,
    beams=[(0.25, 0.50, 5.00), (0.25, 0.50, 5.00), (0.30, 0.60, 6.00), (0.30, 0.60, 6.00)],
    slabs=[(0.18, 5.00, 6.00), (0.20, 6.00, 6.00)],
    seed=7,
)
last = write_table(ws, 4, rows, "Prometrisi")  # data rows 5..34
rng = {k: f"${c}$5:${c}${last}" for k, c in (("floor", "B"), ("type", "C"), ("V", "G"))}

# Totals: SUM vs SUBTOTAL
summary_header(ws, 4, 9, ["Σύνολα", "V (m³)", "Τύπος"])
for r, (label, f) in enumerate([("Όλα τα στοιχεία (SUM)", f"=SUM({rng['V']})"),
                                ("Μόνο τα ορατά (SUBTOTAL)", f"=SUBTOTAL(109,{rng['V']})")], start=5):
    ws.cell(r, 9, label)
    out(ws, r, 10, f)
    as_text(ws.cell(r, 11), f.replace(",", ";"))
note(ws, 7, "Φιλτράρετε π.χ. μόνο τις πλάκες: το SUM δεν αλλάζει, το SUBTOTAL δείχνει το άθροισμα των ορατών.", col=9)

# SUMIF / COUNTIF per floor
summary_header(ws, 9, 9, ["Όροφος", "V (m³)", "Πλήθος"])
for r, fl in enumerate([0, 1, 2], start=10):
    ws.cell(r, 9, fl).border = BOX
    out(ws, r, 10, f"=SUMIF({rng['floor']},I{r},{rng['V']})")
    out(ws, r, 11, f"=COUNTIF({rng['floor']},I{r})", "0")
note(ws, 13, f"J10: =SUMIF({rng['floor']};I10;{rng['V']})", col=9)

# SUMIF per element type
summary_header(ws, 15, 9, ["Στοιχείο", "V (m³)", "Πλήθος"])
for r, t in enumerate(TYPES, start=16):
    ws.cell(r, 9, t).border = BOX
    out(ws, r, 10, f"=SUMIF({rng['type']},I{r},{rng['V']})")
    out(ws, r, 11, f"=COUNTIF({rng['type']},I{r})", "0")

# SUMIFS matrix floor x type
summary_header(ws, 20, 9, ["Όροφος \\ Στοιχείο", *TYPES, "Σύνολο"])
for r, fl in enumerate([0, 1, 2], start=21):
    ws.cell(r, 9, fl).border = BOX
    for c in range(10, 13):
        col = ws.cell(20, c).column_letter
        out(ws, r, c, f"=SUMIFS({rng['V']},{rng['floor']},$I{r},{rng['type']},{col}$20)")
    out(ws, r, 13, f"=SUM(J{r}:L{r})")
note(ws, 24, f"J21: =SUMIFS({rng['V']};{rng['floor']};$I21;{rng['type']};J$20) – μικτές αναφορές (Μάθημα 3)!", col=9)

# Cost
ws["I26"] = "Τιμή σκυροδέματος (€/m³)"
input_cell(ws["J26"], "0")
ws["J26"].value = 95
ws["I27"] = "Κόστος (€)"
out(ws, 27, 10, "=ROUND(J5*J26,2)", "0.00")
widths(ws, A=10, B=9, C=13, D=9, E=9, F=9, G=10, H=3, I=26, J=13, K=13, L=12, M=10)
ws.freeze_panes = "A5"

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Προμέτρηση τετραώροφου κτιρίου",
      "Συμπληρώστε τα κίτρινα κελιά και ακολουθήστε τα βήματα στη στήλη I.")
ex["A3"] = "Δ1·Δ2·Δ3: υποστύλωμα b·h·H, δοκός b·h·L, πλάκα t·Lx·Ly"
ex["A3"].font = Font(italic=True, color="595959")
rows = take_off(
    floors=[0, 1, 2, 3],
    columns=lambda fl: [(0.50, 0.50, 3.20)] * 2 if fl == 0 else [(0.40, 0.50, 3.00)] * 2,
    beams=[(0.25, 0.60, 4.50), (0.25, 0.60, 4.50), (0.30, 0.60, 5.50)],
    slabs=[(0.20, 8.00, 10.00)],
    seed=11,
)
last = write_table(ex, 4, rows, "Askisi", student=True)  # data rows 5..28
steps = [
    "1) Υπολογίστε τη στήλη V (m³) = Δ1·Δ2·Δ3.",
    "2) Μετατρέψτε την περιοχή A4:G28 σε Πίνακα (Ctrl+T).",
    "3) Ταξινομήστε κατά Όροφο και μετά κατά Στοιχείο (Προσαρμοσμένη ταξινόμηση, 2 επίπεδα).",
    "4) Φιλτράρετε μόνο τις δοκούς και διαβάστε το άθροισμα V στη Γραμμή κατάστασης.",
    "5) Καθαρίστε το φίλτρο και συμπληρώστε τους κίτρινους πίνακες με SUMIF, COUNTIF, SUMIFS.",
]
for r, s in enumerate(steps, start=4):
    ex.cell(r, 9, s)
rng = {k: f"${c}$5:${c}${last}" for k, c in (("floor", "B"), ("type", "C"), ("V", "G"))}

summary_header(ex, 11, 9, ["Όροφος", "V (m³)", "Πλήθος στοιχείων"])
for r, fl in enumerate([0, 1, 2, 3], start=12):
    ex.cell(r, 9, fl).border = BOX
    out(ex, r, 10, None, student=True)
    out(ex, r, 11, None, "0", student=True)
summary_header(ex, 18, 9, ["Όροφος \\ Στοιχείο", *TYPES])
for r, fl in enumerate([0, 1, 2, 3], start=19):
    ex.cell(r, 9, fl).border = BOX
    for c in range(10, 13):
        out(ex, r, c, None, student=True)
ex["I24"] = "Συνολικός όγκος (m³)"
out(ex, 24, 10, None, student=True)
note(ex, 25, "Έλεγχος: συνολικός όγκος = 78,560 m³ · ισόγειο (0) = 19,940 m³ · δοκοί 2ου ορόφου = 2,340 m³.", col=9)
widths(ex, A=10, B=9, C=13, D=9, E=9, F=9, G=10, H=3, I=26, J=13, K=16, L=12)
ex.freeze_panes = "A5"

save(wb, "09-sort-filter.xlsx")
