# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/04-functions.xlsx.  Run with:  uv run tools/make_04_functions.py"""

from openpyxl import Workbook

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, table,
                       title, widths)


def results(ws, top, rows, student=False):
    """Label | result | formula-as-text rows starting at `top`.

    rows: (label, formula for openpyxl, formula as shown with Greek separators, number format).
    With student=True the result cell is yellow and empty and the formula text is omitted.
    """
    for i, (label, formula, shown, fmt) in enumerate(rows, start=top):
        ws.cell(i, 1, label).border = BOX
        c = ws.cell(i, 2)
        if student:
            input_cell(c, fmt)
        else:
            c.value = formula
            c.fill = FORMULA_FILL
            c.number_format = fmt
            c.border = BOX
            as_text(ws.cell(i, 3), shown)
    return top + len(rows)


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 4 – Βασικές συναρτήσεις",
      "Στη στήλη C φαίνεται ο τύπος όπως γράφεται με ελληνικές ρυθμίσεις (ορίσματα με ;).")

# --- A. Statistics of cube results ---
section(ws, 4, "Α. Αποτελέσματα θλίψης κύβων σκυροδέματος")
table(ws, 5, ["Δοκίμιο", "f (MPa)"],
      [["Κ1", 31.2], ["Κ2", 29.8], ["Κ3", 33.5], ["Κ4", 30.4], ["Κ5", 32.1], ["Κ6", 28.9]], fmt="0.0")
results(ws, 13, [
    ("Πλήθος δοκιμίων", "=COUNT(B6:B11)", "=COUNT(B6:B11)", "0"),
    ("Μέση τιμή (MPa)", "=AVERAGE(B6:B11)", "=AVERAGE(B6:B11)", "0.0000"),
    ("Ελάχιστη τιμή (MPa)", "=MIN(B6:B11)", "=MIN(B6:B11)", "0.0"),
    ("Μέγιστη τιμή (MPa)", "=MAX(B6:B11)", "=MAX(B6:B11)", "0.0"),
    ("Μέση τιμή στρογγυλεμένη", "=ROUND(AVERAGE(B6:B11),1)", "=ROUND(AVERAGE(B6:B11);1)", "0.0"),
])

# --- B. Quantities and cost ---
section(ws, 20, "Β. Προμέτρηση σκυροδέματος δοκών")
ws["A21"] = "Τιμή (€/m³)"
input_cell(ws["B21"])
ws["B21"].value = 95
table(ws, 23, ["Δοκός", "b (m)", "h (m)", "L (m)", "V (m³)"],
      [["Δ1", 0.25, 0.50, 5.00], ["Δ2", 0.30, 0.60, 6.00], ["Δ3", 0.25, 0.45, 4.50], ["Δ4", 0.30, 0.50, 5.50]],
      {"E": "=B{r}*C{r}*D{r}"})
results(ws, 29, [
    ("Συνολικός όγκος (m³)", "=SUM(E24:E27)", "=SUM(E24:E27)", "0.0000"),
    ("Κόστος (€), 2 δεκαδικά", "=ROUND(B29*B21,2)", "=ROUND(B29*B21;2)", "0.00"),
])

# --- C. Reinforcement area ---
section(ws, 33, "Γ. Εμβαδόν οπλισμού  As = n·π·Φ²/4  (Φ σε mm → As σε cm²)")
table(ws, 34, ["Θέση", "n", "Φ (mm)", "As (cm²)"],
      [["Κάτω", 4, 16], ["Πάνω", 3, 20], ["Πλευρικά", 2, 12]],
      {"D": "=B{r}*PI()*C{r}^2/4/100"}, fmt="0", formula_fmt="0.00")
results(ws, 38, [("Σύνολο As (cm²)", "=SUM(D35:D37)", "=SUM(D35:D37)", "0.00")])
note(ws, 39, "Τύπος στο D35: =B35*PI()*C35^2/4/100  – η PI() δεν έχει ορίσματα, αλλά οι παρενθέσεις χρειάζονται.")

# --- D. Stirrups ---
section(ws, 42, "Δ. Πλήθος συνδετήρων:  n = ROUNDUP(L/s) + 1")
for r, (label, val) in enumerate([("Μήκος L (m)", 5.00), ("Απόσταση s (m)", 0.15)], start=43):
    ws.cell(r, 1, label)
    input_cell(ws.cell(r, 2))
    ws.cell(r, 2).value = val
results(ws, 45, [
    ("L/s (χωρίς στρογγύλευση)", "=B43/B44", "=B43/B44", "0.000"),
    ("Πλήθος συνδετήρων", "=ROUNDUP(B43/B44,0)+1", "=ROUNDUP(B43/B44;0)+1", "0"),
])

# --- E. Trigonometry ---
section(ws, 49, "Ε. Συνιστώσες δύναμης – οι τριγωνομετρικές συναρτήσεις θέλουν ακτίνια!")
for r, (label, val) in enumerate([("Δύναμη F (kN)", 50), ("Γωνία θ (μοίρες)", 30)], start=50):
    ws.cell(r, 1, label)
    input_cell(ws.cell(r, 2))
    ws.cell(r, 2).value = val
results(ws, 52, [
    ("Fx = F·cosθ  (kN)", "=B50*COS(RADIANS(B51))", "=B50*COS(RADIANS(B51))", "0.00"),
    ("Fy = F·sinθ  (kN)", "=B50*SIN(RADIANS(B51))", "=B50*SIN(RADIANS(B51))", "0.00"),
    ("ΛΑΘΟΣ: χωρίς RADIANS", "=B50*COS(B51)", "=B50*COS(B51)", "0.00"),
])
note(ws, 55, "Το COS(30) υπολογίζει το συνημίτονο 30 ακτινίων, όχι 30 μοιρών.")
widths(ws, A=28, B=14, C=30, D=12, E=12)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Βασικές συναρτήσεις", "Συμπληρώστε τα κίτρινα κελιά χρησιμοποιώντας συναρτήσεις.")

# --- A. Cube statistics ---
section(ex, 4, "Α. Στατιστικά αποτελεσμάτων κύβων")
table(ex, 5, ["Δοκίμιο", "f (MPa)"],
      [[f"Κ{i}", v] for i, v in enumerate([26.5, 28.1, 27.4, 25.9, 29.0, 27.8, 26.7, 28.4], start=1)],
      fmt="0.0")
results(ex, 15, [
    ("Πλήθος δοκιμίων", None, None, "0"),
    ("Μέση τιμή (MPa)", None, None, "0.000"),
    ("Ελάχιστη τιμή (MPa)", None, None, "0.0"),
    ("Μέγιστη τιμή (MPa)", None, None, "0.0"),
    ("Μέση τιμή με 1 δεκαδικό (ROUND)", None, None, "0.0"),
], student=True)
note(ex, 20, "Έλεγχος: μέση τιμή = 27,475 MPa.")

# --- B. Floor areas ---
section(ex, 23, "Β. Εμβαδά χώρων πλάκας")
table(ex, 24, ["Χώρος", "Μήκος (m)", "Πλάτος (m)", "Εμβαδόν (m²)"],
      [["Χ1", 4.20, 3.80], ["Χ2", 5.10, 4.00], ["Χ3", 3.20, 2.50], ["Χ4", 6.00, 4.50]],
      {"D": ""}, student=True, formula_fmt="0.00")
results(ex, 30, [
    ("Συνολικό εμβαδόν (m²)", None, None, "0.00"),
    ("Μεγαλύτερος χώρος (m²)", None, None, "0.00"),
    ("Μέσο εμβαδόν χώρου (m²)", None, None, "0.00"),
], student=True)
note(ex, 33, "Έλεγχος: συνολικό εμβαδόν = 71,36 m².")

# --- C. Reinforcement ---
section(ex, 36, "Γ. Εμβαδόν οπλισμού  As = n·π·Φ²/4  (cm²) – χρησιμοποιήστε την PI()")
table(ex, 37, ["Δοκός", "n", "Φ (mm)", "As (cm²)"],
      [["Δ1", 4, 14], ["Δ2", 5, 16], ["Δ3", 3, 18], ["Δ4", 2, 25]],
      {"D": ""}, student=True, fmt="0", formula_fmt="0.00")
note(ex, 42, "Έλεγχος: 4Φ14 → As = 6,16 cm².")

# --- D. Stirrups ---
section(ex, 45, "Δ. Πλήθος συνδετήρων  n = ROUNDUP(L/s) + 1")
ex["A46"] = "Απόσταση s (m)"
ex["B46"] = 0.25
ex["B46"].border = BOX
table(ex, 48, ["Δοκός", "L (m)", "Συνδετήρες"],
      [["Δ1", 4.50], ["Δ2", 6.00], ["Δ3", 5.20]],
      {"C": ""}, student=True, formula_fmt="0")
note(ex, 52, "Έλεγχος: Δ3 → 22 συνδετήρες. (Χρειάζεστε και απόλυτη αναφορά!)")

# --- E. Trigonometry ---
section(ex, 55, "Ε. Συνιστώσες δύναμης F = 80 kN")
table(ex, 56, ["θ (μοίρες)", "Fx (kN)", "Fy (kN)"],
      [[0], [30], [45], [60], [90]], {"B": "", "C": ""}, student=True, fmt="0", formula_fmt="0.00")
ex["E56"] = "F (kN)"
ex["F56"] = 80
ex["F56"].border = BOX
note(ex, 62, "Έλεγχος: θ = 60° → Fx = 40,00 kN, Fy = 69,28 kN.")
widths(ex, A=32, B=14, C=14, D=14, E=10, F=10)

save(wb, "04-functions.xlsx")
