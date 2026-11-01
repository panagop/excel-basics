# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/07-if.xlsx.  Run with:  uv run tools/make_07_if.py"""

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill

from xlhelpers import BOX, input_cell, note, save, section, table, title, widths

RED = dict(fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(color="9C0006", bold=True))
GREEN = dict(fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True))
AMBER = dict(fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C5700", bold=True))


def highlight(ws, rng, fails=("ΑΝΕΠΑΡΚΕΙΑ", "ΑΠΟΡΡΙΠΤΕΤΑΙ", "ΛΙΓΟΣ ΟΠΛΙΣΜΟΣ", "ΠΟΛΥΣ ΟΠΛΙΣΜΟΣ")):
    """Colour check results: OK green, ΟΡΙΑΚΑ amber, failures red."""
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], **GREEN))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"ΟΡΙΑΚΑ"'], **AMBER))
    for text in fails:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'], **RED))


def given(ws, row, label, value, fmt="0.00"):
    ws.cell(row, 1, label)
    input_cell(ws.cell(row, 2), fmt)
    ws.cell(row, 2).value = value


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 7 – Λογικοί έλεγχοι: IF, AND, OR",
      "Τα κελιά «Έλεγχος» χρωματίζονται αυτόματα με μορφοποίηση υπό όρους. Αλλάξτε τα δεδομένα και παρακολουθήστε.")

# --- A. Axial stress check, IF and nested IF ---
section(ws, 4, "Α. Έλεγχος τάσης υποστυλωμάτων:  σ = N/(b·h) ≤ f_d,  βαθμός εκμετάλλευσης η = σ/f_d")
given(ws, 5, "f_d (MPa) – δίνεται", 14.2)
table(
    ws, 7,
    ["Υποστ.", "N (kN)", "b (m)", "h (m)", "σ (MPa)", "η", "Έλεγχος (IF)", "Κατάταξη (εμφωλευμένο IF)"],
    [["Κ1", 1800, 0.40, 0.40], ["Κ2", 2400, 0.40, 0.40], ["Κ3", 2100, 0.35, 0.45],
     ["Κ4", 3000, 0.50, 0.50], ["Κ5", 1500, 0.30, 0.30]],
    {"E": "=B{r}/(C{r}*D{r})/1000", "F": "=E{r}/$B$5",
     "G": '=IF(E{r}<=$B$5,"OK","ΑΝΕΠΑΡΚΕΙΑ")',
     "H": '=IF(F{r}<=0.9,"OK",IF(F{r}<=1,"ΟΡΙΑΚΑ","ΑΝΕΠΑΡΚΕΙΑ"))'},
    formula_fmt="0.00",
)
for r in range(8, 13):
    ws[f"B{r}"].number_format = "0"
    ws[f"F{r}"].number_format = "0.000"
highlight(ws, "G8:H12")
note(ws, 13, 'G8: =IF(E8<=$B$5;"OK";"ΑΝΕΠΑΡΚΕΙΑ")')
note(ws, 14, 'H8: =IF(F8<=0,9;"OK";IF(F8<=1;"ΟΡΙΑΚΑ";"ΑΝΕΠΑΡΚΕΙΑ"))')

# --- B. Reinforcement ratio range, AND ---
section(ws, 17, "Β. Ποσοστό οπλισμού μέσα σε όρια:  ρ_min ≤ ρ ≤ ρ_max  (ενδεικτικά όρια – δίνονται)")
given(ws, 18, "ρ_min", 0.0015, "0.00%")
given(ws, 19, "ρ_max", 0.04, "0.00%")
table(
    ws, 21,
    ["Δοκός", "b (m)", "d (m)", "As (cm²)", "ρ", "Έλεγχος (AND)", "Μήνυμα (εμφωλευμένο IF)"],
    [["Δ1", 0.25, 0.46, 8.04], ["Δ2", 0.30, 0.55, 1.57], ["Δ3", 0.25, 0.41, 6.03], ["Δ4", 0.20, 0.36, 31.40]],
    {"E": "=D{r}/10000/(B{r}*C{r})",
     "F": '=IF(AND(E{r}>=$B$18,E{r}<=$B$19),"OK","ΑΝΕΠΑΡΚΕΙΑ")',
     "G": '=IF(E{r}<$B$18,"ΛΙΓΟΣ ΟΠΛΙΣΜΟΣ",IF(E{r}>$B$19,"ΠΟΛΥΣ ΟΠΛΙΣΜΟΣ","OK"))'},
)
for r in range(22, 26):
    ws[f"E{r}"].number_format = "0.00%"
highlight(ws, "F22:G25")
note(ws, 26, 'F22: =IF(AND(E22>=$B$18;E22<=$B$19);"OK";"ΑΝΕΠΑΡΚΕΙΑ")')

# --- C. Two criteria, AND vs OR ---
section(ws, 29, "Γ. Δύο κριτήρια μαζί: αντοχή (η ≤ 1) και βέλος (δ ≤ L/250)")
table(
    ws, 30,
    ["Δοκός", "L (m)", "δ (mm)", "η", "δ_lim = L/250 (mm)", "Με AND", "Με OR"],
    [["Δ1", 5.00, 15, 0.85], ["Δ2", 6.00, 26, 0.92], ["Δ3", 4.50, 17, 1.05]],
    {"E": "=B{r}*1000/250",
     "F": '=IF(AND(D{r}<=1,C{r}<=E{r}),"OK","ΑΠΟΡΡΙΠΤΕΤΑΙ")',
     "G": '=IF(OR(D{r}>1,C{r}>E{r}),"ΑΠΟΡΡΙΠΤΕΤΑΙ","OK")'},
    formula_fmt="0.0",
)
highlight(ws, "F31:G33")
note(ws, 34, "«Περνάει» = ικανοποιεί ΚΑΙ τα δύο κριτήρια.  «Απορρίπτεται» = αποτυγχάνει ΣΕ ΕΝΑ ΤΟΥΛΑΧΙΣΤΟΝ.")
note(ws, 35, "Γι’ αυτό το AND με τα ≤ και το OR με τα > δίνουν πάντα το ίδιο αποτέλεσμα.")

# --- D. IF returning a number ---
section(ws, 38, "Δ. Το IF μπορεί να επιστρέψει και αριθμό:  As = max(As,req ; As,min)")
given(ws, 39, "As,min (cm²) – δίνεται", 1.50)
table(
    ws, 41,
    ["Θέση", "As,req (cm²)", "As με IF", "As με MAX"],
    [["1", 3.20], ["2", 1.10], ["3", 5.60]],
    {"C": "=IF(B{r}<$B$39,$B$39,B{r})", "D": "=MAX(B{r},$B$39)"},
    formula_fmt="0.00",
)
note(ws, 45, "C42: =IF(B42<$B$39;$B$39;B42)    D42: =MAX(B42;$B$39)  – όταν υπάρχει έτοιμη συνάρτηση, είναι πιο απλή.")
widths(ws, A=22, B=12, C=12, D=12, E=18, F=16, G=26, H=26)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Λογικοί έλεγχοι", 'Συμπληρώστε τα κίτρινα κελιά. Τα κείμενα μπαίνουν σε ίσια εισαγωγικά: "OK".')

# --- A. Soil pressure ---
section(ex, 4, "Α. Τάση εδάφους κάτω από πέδιλα:  σ = N/(B·L) ≤ σ_επ  (αγνοούμε το ίδιο βάρος του πεδίλου)")
given(ex, 5, "σ_επ (kPa) – δίνεται", 250, "0")
table(
    ex, 7, ["Πέδιλο", "N (kN)", "B (m)", "L (m)", "σ (kPa)", "Έλεγχος"],
    [["Π1", 900, 2.0, 2.0], ["Π2", 1200, 2.0, 2.2], ["Π3", 650, 1.6, 1.8], ["Π4", 1500, 2.5, 2.5]],
    {"E": "", "F": ""}, student=True,
)
for r in range(8, 12):
    ex[f"B{r}"].number_format = "0"
note(ex, 12, 'Αποτέλεσμα: "OK" ή "ΑΝΕΠΑΡΚΕΙΑ". Έλεγχος: Π2 → σ = 272,73 kPa → ΑΝΕΠΑΡΚΕΙΑ.')

# --- B. Concrete conformity ---
section(ex, 15, "Β. Συμμόρφωση σκυροδέματος C25/30 – κριτήρια αρχικής παραγωγής (3 κύβοι ανά παρτίδα)")
given(ex, 16, "f_ck,cube (MPa)", 30, "0")
note(ex, 17, "Κριτήρια:  μέση τιμή f_m ≥ f_ck + 4   ΚΑΙ   κάθε τιμή f_i ≥ f_ck − 4")
table(
    ex, 19, ["Παρτίδα", "f1 (MPa)", "f2 (MPa)", "f3 (MPa)", "f_m", "f_min", "Αποδοχή"],
    [["1", 35.2, 33.8, 36.1], ["2", 32.5, 34.0, 33.1], ["3", 38.0, 25.5, 39.4], ["4", 36.5, 34.8, 35.9]],
    {"E": "", "F": "", "G": ""}, student=True, fmt="0.0",
)
note(ex, 24, 'Χρησιμοποιήστε AVERAGE, MIN και AND. Αποτέλεσμα: "OK" ή "ΑΠΟΡΡΙΠΤΕΤΑΙ". Έλεγχος: γίνονται δεκτές 2 παρτίδες.')

# --- C. Nested IF: seismic zones ---
section(ex, 27, "Γ. Εμφωλευμένο IF: σεισμική επιτάχυνση εδάφους από τη ζώνη σεισμικότητας")
note(ex, 28, "Ζώνη 1 → a_g = 0,16g    Ζώνη 2 → a_g = 0,24g    Ζώνη 3 → a_g = 0,36g")
table(
    ex, 30, ["Θέση", "Ζώνη", "a_g / g", "a_g (m/s²)"],
    [["Α", 1], ["Β", 3], ["Γ", 2], ["Δ", 1], ["Ε", 2]],
    {"C": "", "D": ""}, student=True, fmt="0",
)
note(ex, 36, "Έλεγχος: Θέση Β → a_g/g = 0,36 → a_g = 3,53 m/s² (g = 9,81 m/s²).")

# --- D. OR ---
section(ex, 39, "Δ. Απόρριψη δοκού αν αποτυγχάνει ΕΣΤΩ ΚΑΙ ΕΝΑ κριτήριο: η > 1 ή δ > L/250 – χρησιμοποιήστε OR")
table(
    ex, 40, ["Δοκός", "L (m)", "δ (mm)", "η", "δ_lim (mm)", "Αποτέλεσμα"],
    [["Δ1", 6.00, 20, 0.88], ["Δ2", 5.00, 21, 0.75], ["Δ3", 7.00, 25, 1.08], ["Δ4", 4.00, 12, 0.95]],
    {"E": "", "F": ""}, student=True,
)
note(ex, 45, 'Αποτέλεσμα: "OK" ή "ΑΠΟΡΡΙΠΤΕΤΑΙ". Έλεγχος: απορρίπτονται οι Δ2 και Δ3.')
for rng in ("F8:F11", "G20:G23", "F41:F44"):
    highlight(ex, rng)
widths(ex, A=22, B=12, C=12, D=12, E=14, F=14, G=16)

save(wb, "07-if.xlsx")
