# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/08-lookup.xlsx.  Run with:  uv run tools/make_08_lookup.py

XLOOKUP is stored as _xlfn.XLOOKUP: Excel needs this prefix for functions newer than Excel 2007.
"""

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, table,
                       title, widths)

M = "'Υλικά'!"  # sheet prefix for references to the materials sheet
CLASS, FCK, FCKC, FCM, FCTM, ECM = (f"{M}${c}$5:${c}$11" for c in "ABCDEF")
PHI, AS1, MASS = (f"{M}${c}$16:${c}$26" for c in "ABC")
ZONE, AG = f"{M}$A$31:$A$33", f"{M}$B$31:$B$33"


def xl(lookup, find_in, ret, *extra):
    return f"_xlfn.XLOOKUP({lookup},{find_in},{ret}{''.join(',' + e for e in extra)})"


def greek(formula):
    """Formula text as typed with Greek regional settings."""
    return formula.replace("_xlfn.", "").replace(",", ";").replace("0.", "0,")


def formula_cell(ws, ref, f, fmt="0.00"):
    c = ws[ref]
    c.value = f
    c.fill = FORMULA_FILL
    c.number_format = fmt
    c.border = BOX


wb = Workbook()

# ================= Materials sheet =================
mat = wb.active
mat.title = "Υλικά"
title(mat, "Πίνακες υλικών", "Οι πίνακες που «ψάχνουν» οι συναρτήσεις αναζήτησης.")
section(mat, 3, "Σκυρόδεμα – EN 1992-1-1, Πίνακας 3.1")
table(mat, 4, ["Κατηγορία", "f_ck (MPa)", "f_ck,cube (MPa)", "f_cm (MPa)", "f_ctm (MPa)", "E_cm (GPa)"],
      [["C20/25", 20, 25, None, 2.2, 30], ["C25/30", 25, 30, None, 2.6, 31], ["C30/37", 30, 37, None, 2.9, 33],
       ["C35/45", 35, 45, None, 3.2, 34], ["C40/50", 40, 50, None, 3.5, 35], ["C45/55", 45, 55, None, 3.8, 36],
       ["C50/60", 50, 60, None, 4.1, 37]],
      {"D": "=B{r}+8"}, fmt="0", formula_fmt="0")
for r in range(5, 12):
    mat[f"E{r}"].number_format = "0.0"
section(mat, 14, "Ράβδοι οπλισμού")
table(mat, 15, ["Φ (mm)", "As (mm²)", "Βάρος (kg/m)"],
      [[d] for d in (8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32)],
      {"B": "=PI()*A{r}^2/4", "C": "=B{r}*7850/10^6"}, fmt="0", formula_fmt="0.00")
for r in range(16, 27):
    mat[f"C{r}"].number_format = "0.000"
note(mat, 27, "As = π·Φ²/4,  βάρος = As·ρ με ρ = 7850 kg/m³ (As σε mm² → ×10⁻⁶ m²).")
section(mat, 29, "Ζώνες σεισμικότητας")
table(mat, 30, ["Ζώνη", "a_g / g"], [[1, 0.16], [2, 0.24], [3, 0.36]], fmt="0.00")
for r in range(31, 34):
    mat[f"A{r}"].number_format = "0"
widths(mat, A=12, B=14, C=16, D=14, E=14, F=14)

# ================= Example sheet =================
ws = wb.create_sheet("Παράδειγμα")
title(ws, "Μάθημα 8 – Πίνακες αναζήτησης: XLOOKUP και VLOOKUP",
      "Οι πίνακες βρίσκονται στο φύλλο «Υλικά». Στη στήλη D φαίνεται ο τύπος της στήλης B.")

# --- A. Material properties from the class ---
section(ws, 4, "Α. Ιδιότητες σκυροδέματος από την κατηγορία – επιλέξτε κατηγορία από τη λίστα στο B5")
ws["A5"] = "Κατηγορία"
input_cell(ws["B5"], "@")
ws["B5"].value = "C25/30"
dv = DataValidation(type="list", formula1=f"{M}$A$5:$A$11", allow_blank=False)
ws.add_data_validation(dv)
dv.add("B5")
table(ws, 6, ["Ιδιότητα", "XLOOKUP", "VLOOKUP", "Τύπος στο B"], [[None] * 4] * 4)
for r, (label, col, idx) in zip(range(7, 11), [("f_ck (MPa)", FCK, 2), ("f_cm (MPa)", FCM, 4),
                                               ("f_ctm (MPa)", FCTM, 5), ("E_cm (GPa)", ECM, 6)]):
    ws.cell(r, 1, label)
    f = "=" + xl("$B$5", CLASS, col)
    formula_cell(ws, f"B{r}", f, "0.0")
    formula_cell(ws, f"C{r}", f"=VLOOKUP($B$5,{M}$A$5:$F$11,{idx},0)", "0.0")
    as_text(ws.cell(r, 4), greek(f))

# --- B. One class per row, not-found message ---
section(ws, 13, "Β. Κατηγορία ανά στοιχείο – και τι γίνεται όταν δεν βρεθεί (Δ4)")
table(ws, 14, ["Δοκός", "Κατηγορία", "f_ck (MPa)", "f_ctm (MPa)"],
      [["Δ1", "C20/25"], ["Δ2", "C30/37"], ["Δ3", "C25/30"], ["Δ4", "C30/35"]])
for r in range(15, 19):
    formula_cell(ws, f"C{r}", "=" + xl(f"B{r}", CLASS, FCK, '"Άγνωστη"'), "0")
    formula_cell(ws, f"D{r}", "=" + xl(f"B{r}", CLASS, FCTM, '"Άγνωστη"'), "0.0")
note(ws, 19, 'C15: ' + greek("=" + xl("B15", CLASS, FCK, '"Άγνωστη"')))
note(ws, 20, "Η κατηγορία C30/35 δεν υπάρχει στον πίνακα: χωρίς το 4ο όρισμα θα εμφανιζόταν #N/A.")

# --- C. Reinforcement schedule ---
section(ws, 23, "Γ. Οπλισμός δοκού: εμβαδόν και βάρος από τον πίνακα ράβδων")
table(ws, 24, ["Θέση", "n", "Φ (mm)", "As (cm²)", "Βάρος (kg/m)"],
      [["Κάτω", 4, 16], ["Πάνω", 3, 20], ["Πλευρικά", 2, 12]], fmt="0")
for r in range(25, 28):
    formula_cell(ws, f"D{r}", f"=B{r}*" + xl(f"C{r}", PHI, AS1) + "/100")
    formula_cell(ws, f"E{r}", f"=B{r}*" + xl(f"C{r}", PHI, MASS))
note(ws, 28, "D25: " + greek("=B25*" + xl("C25", PHI, AS1) + "/100"))

# --- D. Bar selection, approximate match ---
section(ws, 31, "Δ. Επιλογή διαμέτρου: η μικρότερη ράβδος με εμβαδόν ≥ του απαιτούμενου")
rows = [("As,req (cm²)", 7.5, None), ("Πλήθος ράβδων n", 4, None),
        ("Απαιτούμενο ανά ράβδο (mm²)", "=B32*100/B33", "0.0"),
        ("Επιλεγμένη Φ (mm)", "=" + xl("B34", AS1, PHI, '"καμία"', "1"), "0"),
        ("As,prov (cm²)", "=B33*" + xl("B35", PHI, AS1) + "/100", "0.00")]
for r, (label, val, fmt) in zip(range(32, 37), rows):
    ws.cell(r, 1, label)
    if fmt:
        formula_cell(ws, f"B{r}", val, fmt)
        as_text(ws.cell(r, 4), greek(val))
    else:
        input_cell(ws.cell(r, 2), "0.0")
        ws.cell(r, 2).value = val
note(ws, 37, "Το 5ο όρισμα (1) σημαίνει: ακριβής τιμή ή, αν δεν υπάρχει, η αμέσως μεγαλύτερη. 187,5 mm² → Φ16 (201 mm²).")

# --- E. Seismic zone ---
section(ws, 40, "Ε. Ζώνη σεισμικότητας → a_g: αναζήτηση αντί για εμφωλευμένο IF (Μάθημα 7)")
table(ws, 41, ["Θέση", "Ζώνη", "a_g / g"], [["Α", 1], ["Β", 3], ["Γ", 2]], fmt="0")
for r in range(42, 45):
    formula_cell(ws, f"C{r}", "=" + xl(f"B{r}", ZONE, AG))
note(ws, 45, "C42: " + greek("=" + xl("B42", ZONE, AG)) + "  – για νέα ζώνη αρκεί μια γραμμή στον πίνακα.")
widths(ws, A=30, B=14, C=14, D=58, E=14)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Πίνακες αναζήτησης",
      "Συμπληρώστε τα κίτρινα κελιά με XLOOKUP (ή VLOOKUP σε παλαιότερο Excel), αναζητώντας στο φύλλο «Υλικά».")

section(ex, 4, "Α. Ιδιότητες σκυροδέματος υποστυλωμάτων")
table(ex, 5, ["Υποστ.", "Κατηγορία", "f_ck (MPa)", "f_cm (MPa)", "E_cm (GPa)"],
      [["Κ1", "C25/30"], ["Κ2", "C35/45"], ["Κ3", "C20/25"], ["Κ4", "C40/50"]],
      {"C": "", "D": "", "E": ""}, student=True, formula_fmt="0")
note(ex, 10, "Έλεγχος: Κ2 → f_cm = 43 MPa, E_cm = 34 GPa.")

section(ex, 13, "Β. Πίνακας οπλισμού δοκών")
table(ex, 14, ["Δοκός", "n", "Φ (mm)", "As (cm²)", "Βάρος (kg/m)"],
      [["Δ1", 3, 14], ["Δ2", 5, 18], ["Δ3", 2, 25], ["Δ4", 4, 20]],
      {"D": "", "E": ""}, student=True, fmt="0", formula_fmt="0.00")
note(ex, 19, "Έλεγχος: Δ2 (5Φ18) → As = 12,72 cm².")

section(ex, 22, "Γ. Επιλογή διαμέτρου ράβδων (XLOOKUP με 5ο όρισμα 1)")
table(ex, 23, ["Θέση", "As,req (cm²)", "n", "Ανά ράβδο (mm²)", "Φ (mm)", "As,prov (cm²)"],
      [["1", 6.0, 3], ["2", 10.5, 4], ["3", 3.2, 2], ["4", 15.0, 5]],
      {"D": "", "E": "", "F": ""}, student=True, formula_fmt="0.0")
for r in range(24, 28):
    ex[f"C{r}"].number_format = "0"
note(ex, 28, "Έλεγχος: θέση 2 → 262,5 mm² ανά ράβδο → Φ20, As,prov = 12,57 cm².  As,prov ≥ As,req σε όλες τις γραμμές;")

section(ex, 31, "Δ. Αναπτυσσόμενη λίστα")
ex["A35"] = "Κατηγορία"
input_cell(ex["B35"], "@")
ex["A36"] = "f_ctm (MPa)"
input_cell(ex["B36"], "0.0")
note(ex, 32, "1) Στο B35 δημιουργήστε λίστα κατηγοριών: Δεδομένα → Επικύρωση δεδομένων → Λίστα, Προέλευση: την περιοχή A5:A11 του φύλλου Υλικά.")
note(ex, 33, "2) Στο B36 υπολογίστε το f_ctm της επιλεγμένης κατηγορίας. Αλλάξτε κατηγορία και ελέγξτε ότι ενημερώνεται.")
widths(ex, A=16, B=14, C=12, D=16, E=14, F=14)

wb.active = 1  # open on the example sheet
save(wb, "08-lookup.xlsx")
