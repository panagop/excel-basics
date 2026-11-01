# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/02-formulas.xlsx.  Run with:  uv run tools/make_02_formulas.py"""

from openpyxl import Workbook

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, table,
                       title, widths)

wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 2 – Τύποι και αριθμητικές πράξεις",
      "Τα πράσινα κελιά περιέχουν τύπους. Κάντε κλικ πάνω τους και δείτε τη Γραμμή τύπων.")

# --- A. Order of operations ---
section(ws, 4, "Α. Προτεραιότητα πράξεων")
# (formula shown as text in Greek notation, real formula, comment)
precedence = [
    ("=2+3*4", "=2+3*4", "πρώτα ο πολλαπλασιασμός"),
    ("=(2+3)*4", "=(2+3)*4", "πρώτα η παρένθεση"),
    ("=10/2*5", "=10/2*5", "ίδια προτεραιότητα: από αριστερά προς τα δεξιά"),
    ("=10/(2*5)", "=10/(2*5)", "η παρένθεση αλλάζει τη σειρά"),
    ("=9^0,5", "=9^0.5", "δύναμη 0,5 = τετραγωνική ρίζα"),
    ("=-2^2", "=-2^2", "ΠΡΟΣΟΧΗ: το πρόσημο εφαρμόζεται πρώτα, (−2)² = 4"),
    ("=-(2^2)", "=-(2^2)", "με παρένθεση: −(2²) = −4"),
    ("=2^3^2", "=2^3^2", "ΠΡΟΣΟΧΗ: (2³)² = 64, ενώ στα μαθηματικά 2^(3²) = 512"),
]
table(ws, 5, ["Τύπος (όπως γράφεται)", "Αποτέλεσμα", "Σχόλιο"], [[None] * 3] * len(precedence))
for i, (shown, formula, comment) in enumerate(precedence, start=6):
    as_text(ws.cell(i, 1), shown)
    c = ws.cell(i, 2, formula)
    c.fill = FORMULA_FILL
    c.number_format = "General"
    ws.cell(i, 3, comment)

# --- B. Slab loads ---
section(ws, 16, "Β. Φορτία πλάκας ανά m²")
rows = [
    ("Ειδικό βάρος οπλισμένου σκυροδέματος", "γ", 25, "kN/m³", None),
    ("Πάχος πλάκας", "h", 0.20, "m", None),
    ("Επικαλύψεις", "g_επ", 1.50, "kN/m²", None),
    ("Κινητό φορτίο", "q", 2.00, "kN/m²", None),
    ("Ίδιο βάρος πλάκας", "g_ιδ", "=C18*C19", "kN/m²", "=C18*C19"),
    ("Σύνολο μόνιμων φορτίων", "g", "=C22+C20", "kN/m²", "=C22+C20"),
    ("Φορτίο σχεδιασμού 1,35·g + 1,50·q", "p_d", "=1.35*C23+1.5*C21", "kN/m²", "=1,35*C23+1,5*C21"),
]
table(ws, 17, ["Μέγεθος", "Σύμβολο", "Τιμή", "Μονάδα", "Τύπος στο κελί Τιμή"], [[None] * 5] * len(rows))
for i, (name, sym, val, unit, shown) in enumerate(rows, start=18):
    ws.cell(i, 1, name)
    ws.cell(i, 2, sym)
    c = ws.cell(i, 3, val)
    c.number_format = "0.000" if shown else "0.00"
    if shown:
        c.fill = FORMULA_FILL
        as_text(ws.cell(i, 5), shown)
    ws.cell(i, 4, unit)

# --- C. Concrete cube tests ---
section(ws, 27, "Γ. Θλιπτική αντοχή κύβων 150 mm:  f = F / A,  A = 0,15² m²")
table(
    ws, 28,
    ["Δοκίμιο", "F (kN)", "f (MPa)"],
    [["Κ1", 675], ["Κ2", 712], ["Κ3", 690], ["Κ4", 655], ["Κ5", 702]],
    {"C": "=B{r}/0.15^2/1000"},
    fmt="0", formula_fmt="0.00",
)
note(ws, 34, "Ο τύπος =B29/0,15^2/1000 γράφτηκε μία φορά στο C29 και αντιγράφηκε προς τα κάτω "
             "με τη λαβή συμπλήρωσης. kN/m² ÷ 1000 = MPa.")
widths(ws, A=38, B=14, C=14, D=10, E=24)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Τύποι και αριθμητικές πράξεις",
      "Συμπληρώστε τα κίτρινα κελιά.")

# --- A. Predict, then check ---
section(ex, 4, "Α. Υπολογίστε πρώτα «με το μυαλό» και μετά ελέγξτε με τύπο στο Excel")
exprs = ["3+4*2", "(3+4)*2", "20/4/2", "20/(4/2)", "-3^2", "2^2^3", "(1+2)^2/3"]
table(ex, 5, ["Παράσταση", "Η πρόβλεψή μου", "Τύπος στο Excel"], [[e] for e in exprs],
      {"B": "", "C": ""}, student=True, formula_fmt="General")

# --- B. Simply supported beam ---
section(ex, 14, "Β. Αμφιέρειστη δοκός υπό ίδιο βάρος")
beam = [
    ("Ειδικό βάρος σκυροδέματος", "γ", 25, "kN/m³"),
    ("Πλάτος διατομής", "b", 0.30, "m"),
    ("Ύψος διατομής", "h", 0.60, "m"),
    ("Άνοιγμα", "L", 6.00, "m"),
    ("Ίδιο βάρος ανά μέτρο  g = γ·b·h", "g", None, "kN/m"),
    ("Μέγιστη ροπή  M = g·L²/8", "M", None, "kNm"),
    ("Ροπή αντίστασης  W = b·h²/6", "W", None, "m³"),
    ("Μέγιστη τάση  σ = M/W  (σε MPa)", "σ", None, "MPa"),
]
table(ex, 15, ["Μέγεθος", "Σύμβολο", "Τιμή", "Μονάδα"], [[a, b, c, d] for a, b, c, d in beam])
for r in range(20, 24):
    input_cell(ex.cell(r, 3), "0.000")
note(ex, 24, "Έλεγχος: M = 20,25 kNm και σ = 1,125 MPa. Αν το σ βγαίνει λάθος, ελέγξτε τις παρενθέσεις!")

# --- C. Cube tests ---
section(ex, 26, "Γ. Θλιπτική αντοχή κύβων 150 mm – γράψτε τον τύπο στο C28 και αντιγράψτε τον προς τα κάτω")
table(
    ex, 27,
    ["Δοκίμιο", "F (kN)", "f (MPa)"],
    [["Κ1", 598], ["Κ2", 640], ["Κ3", 615], ["Κ4", 572], ["Κ5", 633]],
    {"C": ""}, student=True, fmt="0", formula_fmt="0.00",
)
note(ex, 33, "Έλεγχος: για το Κ1 πρέπει να βγει f = 26,58 MPa.")
widths(ex, A=38, B=16, C=18, D=10)

save(wb, "02-formulas.xlsx")
