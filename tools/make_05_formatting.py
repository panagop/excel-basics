# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/05-formatting.xlsx.  Run with:  uv run tools/make_05_formatting.py"""

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, table,
                       title, widths)

RED_FILL = PatternFill("solid", fgColor="F8CBAD")
RED_FONT = Font(color="9C0006")

wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 5 – Μορφοποίηση αριθμών και μονάδες",
      "Επιλέξτε ένα κελί και πατήστε Ctrl+1 για να δείτε τη μορφή του. Η Γραμμή τύπων δείχνει την πραγματική τιμή.")

# --- A. Same value, different formats ---
section(ws, 4, "Α. Η ίδια τιμή (0,1125) με διαφορετικές μορφές")
formats = [  # (description, number format in file, code as typed with Greek settings)
    ("Γενική (General)", "General", "Γενική"),
    ("Αριθμός, 2 δεκαδικά", "0.00", "0,00"),
    ("Αριθμός, 4 δεκαδικά", "0.0000", "0,0000"),
    ("Ποσοστό, 1 δεκαδικό", "0.0%", "0,0%"),
    ("Επιστημονική", "0.00E+00", "0,00E+00"),
    ("Προσαρμογή με μονάδα", '0.0000 "m²"', '0,0000 "m²"'),
]
table(ws, 5, ["Μορφή", "Εμφάνιση", "Κωδικός μορφής"], [[None] * 3] * len(formats))
for i, (desc, fmt, code) in enumerate(formats, start=6):
    ws.cell(i, 1, desc)
    c = ws.cell(i, 2, 0.1125)
    c.number_format = fmt
    as_text(ws.cell(i, 3), code)

# --- B. Display vs value ---
section(ws, 14, "Β. Η μορφή αλλάζει μόνο την εμφάνιση – όχι την τιμή")
table(ws, 15, ["", "Τιμή (2 δεκαδικά)", "Τύπος"], [["α", None], ["β", None], ["γ", None], ["Άθροισμα", None]])
for r in range(16, 19):
    ws.cell(r, 2, "=1/3").number_format = "0.00"
    as_text(ws.cell(r, 3), "=1/3")
c = ws.cell(19, 2, "=SUM(B16:B18)")
c.number_format = "0.00"
c.fill = FORMULA_FILL
as_text(ws.cell(19, 3), "=SUM(B16:B18)")
note(ws, 20, "Βλέπετε 0,33 + 0,33 + 0,33 = 1,00. Το Excel προσθέτει τις πλήρεις τιμές 0,3333…, όχι αυτές που εμφανίζονται.")

# --- C. Beam table with units in the formats ---
section(ws, 23, "Γ. Πίνακας δοκών: μονάδες μέσα στη μορφή των κελιών, ποσοστό οπλισμού ρ = As/(b·d)")
table(
    ws, 24,
    ["Δοκός", "b", "h", "d", "As", "A = b·h", "ρ"],
    [["Δ1", 0.25, 0.50, 0.46, 8.04], ["Δ2", 0.30, 0.60, 0.55, 12.57], ["Δ3", 0.25, 0.45, 0.41, 6.03]],
    {"F": "=B{r}*C{r}", "G": "=E{r}/10000/(B{r}*D{r})"},
)
for r in range(25, 28):
    for col in "BCD":
        ws[f"{col}{r}"].number_format = '0.00 "m"'
    ws[f"E{r}"].number_format = '0.00 "cm²"'
    ws[f"F{r}"].number_format = '0.0000 "m²"'
    ws[f"G{r}"].number_format = "0.00%"
note(ws, 28, "Οι μονάδες ΔΕΝ είναι πληκτρολογημένες: τα κελιά περιέχουν σκέτους αριθμούς (δείτε τη Γραμμή τύπων),")
note(ws, 29, "άρα οι τύποι δουλεύουν κανονικά. Στο G25: =E25/10000/(B25*D25)  (cm² → m² με /10000).")

# --- D. Unit conversions ---
section(ws, 32, "Δ. Μετατροπή μονάδων: το Excel δεν ξέρει μονάδες – εσείς ξέρετε!")
table(
    ws, 33,
    ["Μέγεθος", "Τιμή", "Μονάδα", "Μετατροπή", "Μονάδα", "Τύπος"],
    [["Πυκνότητα σκυροδέματος", 2500, "kg/m³", None, "kN/m³", None],
     ["Τάση", 1250, "kN/m²", None, "MPa", None],
     ["Εμβαδόν οπλισμού", 8.04, "cm²", None, "mm²", None]],
    fmt="0.00",
)
for r, (f, shown) in zip(range(34, 37), [("=B34*9.81/1000", "=B34*9,81/1000"),
                                          ("=B35/1000", "=B35/1000"),
                                          ("=B36*100", "=B36*100")]):
    c = ws.cell(r, 4, f)
    c.fill = FORMULA_FILL
    c.number_format = "0.00"
    as_text(ws.cell(r, 6), shown)
note(ws, 37, "Χρήσιμες ισοδυναμίες: 1 MPa = 1 N/mm² = 1000 kN/m² (kPa).  1 cm² = 100 mm² = 10⁻⁴ m².")

# --- E. Conditional formatting ---
section(ws, 40, "Ε. Μορφοποίηση υπό όρους: τιμές κάτω από 25 MPa με κόκκινο")
table(ws, 41, ["Δοκίμιο", "f (MPa)"],
      [["Κ1", 31.2], ["Κ2", 24.1], ["Κ3", 28.6], ["Κ4", 33.0], ["Κ5", 23.8], ["Κ6", 29.5]], fmt="0.0")
ws.conditional_formatting.add("B42:B47", CellIsRule(operator="lessThan", formula=["25"], fill=RED_FILL, font=RED_FONT))
note(ws, 48, "Κεντρική → Μορφοποίηση υπό όρους → Κανόνες επισήμανσης κελιών → Μικρότερο από… Αλλάξτε μια τιμή και δείτε το χρώμα να αλλάζει.")
widths(ws, A=26, B=18, C=18, D=12, E=12, F=16, G=10)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Μορφοποίηση αριθμών και μονάδες",
      "Μην αλλάζετε τις τιμές – αλλάξτε μόνο τη μορφή των κίτρινων κελιών (Ctrl+1).")

# --- A. Apply formats ---
section(ex, 4, "Α. Δώστε σε κάθε κίτρινο κελί τη μορφή της στήλης C")
table(ex, 5, ["Μέγεθος", "Τιμή", "Ζητούμενη εμφάνιση"], [
    ["Ροπή κάμψης", 20.25, "20,25 kNm"],
    ["Ποσοστό οπλισμού", 0.006991, "0,70%"],
    ["Μέτρο ελαστικότητας χάλυβα", 200000, "2,0E+05 MPa"],
    ["Οριακή παραμόρφωση σκυροδέματος", 0.0035, "3,50E-03"],
    ["Κόστος σκυροδέματος", 288.44375, "288,44 €"],
    ["Όγκος", 3.03625, "3,04 m³"],
], fmt="General")
for r in range(6, 12):
    input_cell(ex.cell(r, 2), "General")
note(ex, 12, "Υπόδειξη: Ctrl+1 → Αριθμός → Προσαρμογή. Το κείμενο της μονάδας μπαίνει σε εισαγωγικά, π.χ. 0,00 \"kNm\".")

# --- B. Reinforcement ratio ---
section(ex, 15, "Β. Ποσοστό οπλισμού ρ = As/(b·d): γράψτε τον τύπο, μετά μορφοποιήστε όλο τον πίνακα με μονάδες")
table(
    ex, 16,
    ["Δοκός", "b (m)", "h (m)", "d (m)", "As (cm²)", "ρ (%)"],
    [["Δ1", 0.25, 0.50, 0.46, 8.04], ["Δ2", 0.30, 0.60, 0.55, 12.57], ["Δ3", 0.25, 0.45, 0.41, 6.03]],
    {"F": ""}, student=True, fmt="General", formula_fmt="General",
)
note(ex, 20, "Έλεγχος: Δ1 → ρ = 0,70%. Θυμηθείτε: 1 cm² = 10⁻⁴ m².")

# --- C. Unit conversions ---
section(ex, 23, "Γ. Ειδικό βάρος υλικών: kg/m³ → kN/m³  (× 9,81 / 1000)")
table(ex, 24, ["Υλικό", "ρ (kg/m³)", "γ (kN/m³)"],
      [["Σκυρόδεμα", 2500], ["Χάλυβας", 7850], ["Ξύλο", 500], ["Νερό", 1000]],
      {"C": ""}, student=True, fmt="0", formula_fmt="0.00")
note(ex, 29, "Έλεγχος: χάλυβας → 77,01 kN/m³.")

# --- D. Conditional formatting ---
section(ex, 32, "Δ. Επισημάνετε με κόκκινο τις τιμές κάτω από 30 MPa (Μορφοποίηση υπό όρους)")
table(ex, 33, ["Δοκίμιο", "f (MPa)"],
      [["Κ1", 32.4], ["Κ2", 28.7], ["Κ3", 35.1], ["Κ4", 29.9], ["Κ5", 31.6], ["Κ6", 27.3]], fmt="0.0")
note(ex, 40, "Έλεγχος: πρέπει να χρωματιστούν 3 κελιά.")
widths(ex, A=32, B=14, C=20, D=10, E=10, F=10)

save(wb, "05-formatting.xlsx")
