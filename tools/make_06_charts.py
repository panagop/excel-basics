# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/06-charts.xlsx.  Run with:  uv run tools/make_06_charts.py"""

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.trendline import Trendline

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, table,
                       title, widths)


def scatter(ws, anchor, chart_title, x_title, y_title, x_ref, y_refs, names,
            smooth=False, markers=True, x_max=None, trend=False, width=16, height=8.5):
    """XY scatter chart. y_refs/names are lists, one per series."""
    ch = ScatterChart()
    ch.title = chart_title
    ch.style = 2
    ch.width, ch.height = width, height
    ch.x_axis.title = x_title
    ch.y_axis.title = y_title
    ch.x_axis.delete = False  # without these, recent Excel versions hide the axes
    ch.y_axis.delete = False
    ch.x_axis.majorGridlines = None
    if x_max is not None:
        ch.x_axis.scaling.min = 0
        ch.x_axis.scaling.max = x_max
    for y_ref, name in zip(y_refs, names):
        s = Series(y_ref, x_ref, title=name)
        s.smooth = smooth
        if markers:
            s.marker.symbol = "circle"
            s.marker.size = 6
        else:
            s.marker.symbol = "none"
        if trend:
            s.trendline = Trendline(trendlineType="linear", dispEq=True, dispRSqr=True)
        ch.series.append(s)
    if len(y_refs) == 1:
        ch.legend = None
    ws.add_chart(ch, anchor)


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 6 – Διαγράμματα διασποράς XY",
      "Κάντε κλικ σε ένα διάγραμμα: οι περιοχές δεδομένων του χρωματίζονται στο φύλλο.")

# --- A. Steel tensile test ---
section(ws, 4, "Α. Δοκιμή εφελκυσμού χάλυβα (ενδεικτικές μετρήσεις)")
steel = [(0, 0), (0.5, 98), (1.0, 203), (1.5, 297), (2.0, 402), (2.5, 500), (3.0, 503), (5.0, 508),
         (10, 515), (20, 540), (40, 590), (60, 610), (80, 618), (100, 612)]
table(ws, 5, ["ε (‰)", "σ (MPa)"], [list(p) for p in steel], fmt="0.0")
last = 5 + len(steel)  # 19
scatter(ws, "D4", "Διάγραμμα σ–ε χάλυβα", "ε (‰)", "σ (MPa)",
        Reference(ws, min_col=1, min_row=6, max_row=last), [Reference(ws, min_col=2, min_row=6, max_row=last)],
        ["σ–ε"])
scatter(ws, "D22", "Ελαστική περιοχή – γραμμή τάσης", "ε (‰)", "σ (MPa)",
        Reference(ws, min_col=1, min_row=6, max_row=11), [Reference(ws, min_col=2, min_row=6, max_row=11)],
        ["Ελαστική περιοχή"], trend=True)
for r, (label, f, shown, fmt) in enumerate([
    ("Κλίση (MPa/‰)", "=SLOPE(B6:B11,A6:A11)", "=SLOPE(B6:B11;A6:A11)", "0.0"),
    ("E (MPa)", "=B22*1000", "=B22*1000", "0"),
], start=22):
    ws.cell(r, 1, label)
    c = ws.cell(r, 2, f)
    c.fill = FORMULA_FILL
    c.number_format = fmt
    c.border = BOX
    as_text(ws.cell(r, 3), shown)
note(ws, 24, "Η κλίση της ευθείας στην ελαστική περιοχή είναι το μέτρο ελαστικότητας.")
note(ws, 25, "ε σε ‰ → E = κλίση × 1000 (σε MPa).")

# --- B. Bending moment diagram ---
section(ws, 41, "Β. Διάγραμμα ροπών αμφιέρειστης δοκού:  M(x) = q·x·(L − x)/2")
for r, (label, val) in enumerate([("q (kN/m)", 20), ("L (m)", 6)], start=42):
    ws.cell(r, 1, label)
    input_cell(ws.cell(r, 2))
    ws.cell(r, 2).value = val
ws["A44"] = "Βήμα Δx (m)"
ws["B44"] = "=B43/12"
ws["B44"].fill = FORMULA_FILL
ws["B44"].number_format = "0.00"
as_text(ws["C44"], "=B43/12")
table(ws, 46, ["x (m)", "M (kNm)"], [[None, None]] * 13)
for r in range(47, 60):
    a = ws.cell(r, 1, 0 if r == 47 else f"=A{r - 1}+$B$44")
    a.number_format = "0.00"
    m = ws.cell(r, 2, f"=$B$42*A{r}*($B$43-A{r})/2")
    m.number_format = "0.00"
    for c in (a, m):
        if r > 47 or c is m:
            c.fill = FORMULA_FILL
note(ws, 60, "Στο A48: =A47+$B$44   Στο B47: =$B$42*A47*($B$43-A47)/2   – αντιγραμμένοι προς τα κάτω.")
note(ws, 61, "Αλλάξτε το q ή το L: ο πίνακας ΚΑΙ το διάγραμμα ενημερώνονται αυτόματα.")
scatter(ws, "D41", "Διάγραμμα ροπών M(x)", "x (m)", "M (kNm)",
        Reference(ws, min_col=1, min_row=47, max_row=59), [Reference(ws, min_col=2, min_row=47, max_row=59)],
        ["M(x)"], smooth=True, markers=False)
widths(ws, A=16, B=12, C=24)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Διαγράμματα διασποράς XY",
      "Δημιουργήστε τα διαγράμματα με Εισαγωγή → Διασπορά (X, Y). Βάλτε τίτλο και τίτλους αξόνων με μονάδες.")

# --- A. Concrete compression test ---
section(ex, 4, "Α. Δοκιμή θλίψης σκυροδέματος – σχεδιάστε το διάγραμμα σ–ε (διασπορά με ομαλές γραμμές)")
concrete = [(0, 0), (0.25, 7.4), (0.5, 14.2), (0.75, 19.9), (1.0, 24.3), (1.25, 27.5), (1.5, 29.5),
            (1.75, 30.6), (2.0, 31.0), (2.25, 30.8), (2.5, 30.0), (3.0, 28.3), (3.5, 24.9)]
table(ex, 5, ["ε (‰)", "σ (MPa)"], [list(p) for p in concrete], fmt="0.00")
note(ex, 19, "Από το διάγραμμα: ποια είναι η μέγιστη τάση και σε ποια παραμόρφωση εμφανίζεται;")

# --- B. M and V diagrams ---
section(ex, 22, "Β. Αμφιέρειστη δοκός: υπολογίστε M(x) = q·x·(L−x)/2 και V(x) = q·(L/2 − x), σχεδιάστε τα δύο διαγράμματα")
for r, (label, val) in enumerate([("q (kN/m)", 15), ("L (m)", 8), ("Βήμα Δx (m)", 0.5)], start=23):
    ex.cell(r, 1, label)
    ex.cell(r, 2, val).border = BOX
table(ex, 27, ["x (m)", "M (kNm)", "V (kN)"], [[0, None, None]] + [[None] * 3] * 16)
for r in range(28, 45):
    for col in (1, 2, 3):
        if (r, col) != (28, 1):
            input_cell(ex.cell(r, col))
note(ex, 45, "Το x ξεκινά από 0 με βήμα Δx (απόλυτη αναφορά!) μέχρι το L = 8 m.")
note(ex, 46, "Έλεγχος: M(4) = 120,00 kNm,  V(0) = 60,00 kN,  V(8) = −60,00 kN.")

# --- C. Modulus of elasticity ---
section(ex, 49, "Γ. Μέτρο ελαστικότητας από την ελαστική περιοχή μιας δοκιμής εφελκυσμού")
elastic = [(0, 0), (0.4, 83), (0.8, 163), (1.2, 247), (1.6, 329), (2.0, 410)]
table(ex, 50, ["ε (‰)", "σ (MPa)"], [list(p) for p in elastic], fmt="0.0")
for r, label in enumerate(["Κλίση με SLOPE (MPa/‰)", "E (MPa)"], start=58):
    ex.cell(r, 1, label)
    input_cell(ex.cell(r, 2), "0")
note(ex, 60, "1) Σχεδιάστε διάγραμμα διασποράς μόνο με σημεία και προσθέστε γραμμή τάσης με εξίσωση.")
note(ex, 61, "2) Υπολογίστε την κλίση και με τη συνάρτηση SLOPE. Έλεγχος: E ≈ 205 000 MPa.")
widths(ex, A=26, B=12, C=12)

save(wb, "06-charts.xlsx")
