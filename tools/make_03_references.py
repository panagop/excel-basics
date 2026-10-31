# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/03-references.xlsx.  Run with:  uv run tools/make_03_references.py"""

from openpyxl import Workbook
from openpyxl.styles import Font

from xlhelpers import (BOX, FORMULA_FILL, HEADER_FILL, input_cell, note, save, section, table,
                       title, widths)


def header_cell(cell, value, fmt=None):
    cell.value = value
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = HEADER_FILL
    cell.border = BOX
    if fmt:
        cell.number_format = fmt


def grid(ws, top, corner, row_vals, col_vals, formula=None, student=False, fmt="0.00"):
    """Two-way table: header row at `top` (col_vals from column B), header column A (row_vals).

    `formula` is a template with {r} (row) and {c} (column letter) of the cell being written.
    """
    header_cell(ws.cell(top, 1), corner)
    for j, v in enumerate(col_vals, start=2):
        header_cell(ws.cell(top, j), v, "0")
    for i, v in enumerate(row_vals, start=top + 1):
        header_cell(ws.cell(i, 1), v, "0")
        for j in range(2, len(col_vals) + 2):
            c = ws.cell(i, j)
            if student:
                input_cell(c, fmt)
            else:
                c.value = formula.format(r=i, c=c.column_letter)
                c.fill = FORMULA_FILL
                c.number_format = fmt
                c.border = BOX


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 3 – Σχετικές, απόλυτες και μικτές αναφορές",
      "Τα πράσινα κελιά περιέχουν τύπους. Κάντε κλικ πάνω τους και δείτε πώς άλλαξαν οι αναφορές κατά την αντιγραφή.")

# --- A. Relative vs absolute ---
section(ws, 4, "Α. Η πλευρά του κύβου σε κελί: σχετική (λάθος) και απόλυτη (σωστή) αναφορά")
ws["A5"] = "Πλευρά κύβου a (m)"
input_cell(ws["B5"])
ws["B5"].value = 0.15
table(
    ws, 7,
    ["Δοκίμιο", "F (kN)", "f (MPa) – ΛΑΘΟΣ\n=B8/B5^2/1000", "f (MPa) – ΣΩΣΤΟ\n=B8/$B$5^2/1000"],
    [["Κ1", 675], ["Κ2", 712], ["Κ3", 690], ["Κ4", 655], ["Κ5", 702]],
    {"D": "=B{r}/$B$5^2/1000"},
    fmt="0", formula_fmt="0.00",
)
for r in range(8, 13):  # the relative formula of C8 copied down: B5 drifts to B6, B7, ...
    c = ws.cell(r, 3, f"=B{r}/B{r - 3}^2/1000")
    c.fill = FORMULA_FILL
    c.number_format = "0.00"
    c.border = BOX
ws.row_dimensions[7].height = 32
note(ws, 13, "Στη στήλη C η αναφορά B5 «γλίστρησε» σε B6, B7, … (κενό κελί, κείμενο, άλλο φορτίο) → σφάλματα και λάθος τιμές.")
note(ws, 14, "Στη στήλη D το $B$5 έμεινε σταθερό σε όλες τις γραμμές. Αλλάξτε το B5 σε 0,10 και δείτε τη στήλη D.")

# --- B. Safety factors in cells ---
section(ws, 17, "Β. Φορτία σχεδιασμού πλακών:  p_d = γG·g + γQ·q  (οι συντελεστές σε κελιά)")
for r, (name, val) in enumerate([("γG", 1.35), ("γQ", 1.50)], start=18):
    ws.cell(r, 1, name)
    input_cell(ws.cell(r, 2))
    ws.cell(r, 2).value = val
table(
    ws, 21,
    ["Πλάκα", "g (kN/m²)", "q (kN/m²)", "p_d (kN/m²)"],
    [["Π1", 6.50, 2.00], ["Π2", 7.75, 2.00], ["Π3", 6.00, 5.00], ["Π4", 5.50, 3.50]],
    {"D": "=$B$18*B{r}+$B$19*C{r}"},
    formula_fmt="0.000",
)
note(ws, 26, "Τύπος στο D22: =$B$18*B22+$B$19*C22  – οι συντελεστές με απόλυτη αναφορά, τα φορτία με σχετική.")

# --- C. Mixed references ---
section(ws, 29, "Γ. Μικτές αναφορές: ροπή αμφιέρειστης δοκού  M = q·L²/8  (kNm)")
grid(ws, 30, "q (kN/m) \\ L (m)", [10, 15, 20, 25], [4, 5, 6, 7, 8], "=$A{r}*{c}$30^2/8")
note(ws, 35, "Ένας μόνο τύπος στο B31: =$A31*B$30^2/8, αντιγραμμένος δεξιά και κάτω. "
             "$A: η στήλη του q κλειδωμένη· $30: η γραμμή του L κλειδωμένη.")
widths(ws, A=24, B=13, C=20, D=20, E=11, F=11)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Αναφορές κελιών", "Συμπληρώστε τα κίτρινα κελιά. Γράψτε κάθε τύπο ΜΙΑ φορά και αντιγράψτε τον.")

# --- A. Concrete cost ---
section(ex, 4, "Α. Κόστος σκυροδέματος υποστυλωμάτων")
ex["A5"] = "Τιμή σκυροδέματος (€/m³)"
ex["B5"] = 95
ex["B5"].border = BOX
table(
    ex, 7,
    ["Υποστύλωμα", "b (m)", "h (m)", "H (m)", "V (m³)", "Κόστος (€)"],
    [["Κ1", 0.40, 0.40, 3.00], ["Κ2", 0.30, 0.50, 3.00],
     ["Κ3", 0.50, 0.50, 3.20], ["Κ4", 0.25, 0.60, 2.80]],
    {"E": "", "F": ""}, student=True,
)
note(ex, 12, "Γράψτε τους τύπους στα E8 και F8 και αντιγράψτε τους προς τα κάτω. Έλεγχος: Κ1 → 45,60 €.")
note(ex, 13, "Μετά αλλάξτε την τιμή στο B5 σε 110 €/m³: πρέπει να αλλάξουν ΟΛΑ τα κόστη.")

# --- B. Mixed references ---
section(ex, 16, "Β. Ροπή αντίστασης ορθογωνικής διατομής  W = b·h²/6  (cm³)")
grid(ex, 17, "b (cm) \\ h (cm)", [25, 30, 35, 40], [40, 50, 60, 70], student=True, fmt="0")
note(ex, 22, "Γράψτε ΕΝΑΝ τύπο στο B18 με μικτές αναφορές και αντιγράψτε τον σε όλο τον πίνακα.")
note(ex, 23, "Έλεγχος: b = 25 cm, h = 40 cm → W = 6667 cm³.  b = 40 cm, h = 70 cm → W = 32667 cm³.")
widths(ex, A=24, B=11, C=11, D=11, E=11, F=12)

save(wb, "03-references.xlsx")
