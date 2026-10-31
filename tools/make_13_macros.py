# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/13-macros.xlsx.  Run with:  uv run tools/make_13_macros.py

The file is a plain .xlsx with the data only: students paste the VBA code from the lesson page
and save as .xlsm themselves (a downloaded .xlsm would be blocked by Excel's internet-file protection).
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from xlhelpers import BOX, FORMULA_FILL, input_cell, note, save, section, title, widths

UNKNOWN_FILL = PatternFill("solid", fgColor="F4B183")  # the cell Goal Seek changes

wb = Workbook()

# ================= Formatting sheet (recorded macro) =================
ws = wb.active
ws.title = "Μορφοποίηση"
title(ws, "Μάθημα 13 – Α. Εγγραφή μακροεντολής",
      "Δύο αμορφοποίητοι πίνακες. Καταγράψτε μια μακροεντολή μορφοποιώντας τον πρώτο, και εκτελέστε τη στον δεύτερο.")
headers = ["Δοκός", "M_Ed (kNm)", "V_Ed (kN)", "As,req (cm²)", "As,prov (cm²)"]
data1 = [["Δ1", 85.2, 64.15, 4.613, 6.03], ["Δ2", 120.75, 88.4, 6.712, 8.04], ["Δ3", 64.8, 52.333, 3.418, 4.02],
         ["Δ4", 152.1, 101.95, 8.705, 9.42]]
data2 = [["Δ5", 98.4, 70.2, 5.371, 6.03], ["Δ6", 133.333, 92.85, 7.512, 8.04], ["Δ7", 45.6, 38.75, 2.384, 3.39]]
for top, rows in ((5, data1), (12, data2)):
    for j, h in enumerate(headers, start=1):
        ws.cell(top, j, h)
    for i, row in enumerate(rows, start=top + 1):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j, v)
note(ws, 4, "Πίνακας 1: επιλέξτε A5:E9 και καταγράψτε τη μορφοποίηση.")
note(ws, 11, "Πίνακας 2: επιλέξτε A12:E15 και εκτελέστε τη μακροεντολή.")
widths(ws, A=10, B=13, C=13, D=14, E=14)

# ================= UDF sheet =================
fs = wb.create_sheet("Συναρτήσεις")
title(fs, "Μάθημα 13 – Β. Δικές σας συναρτήσεις",
      "Αφού προσθέσετε τον κώδικα σε Λειτουργική μονάδα (Module), χρησιμοποιήστε τις συναρτήσεις στα κίτρινα κελιά.")
section(fs, 4, "=BAR_AREA(n; Φ) → εμβαδόν n ράβδων διαμέτρου Φ (mm), σε cm²")
for j, h in enumerate(["n", "Φ (mm)", "As (cm²)"], start=1):
    fs.cell(5, j, h).font = Font(bold=True)
for i, (n, d) in enumerate([(4, 16), (3, 20), (2, 12), (5, 25)], start=6):
    fs.cell(i, 1, n).border = BOX
    fs.cell(i, 2, d).border = BOX
    input_cell(fs.cell(i, 3), "0.00")
note(fs, 10, "Έλεγχος: 4Φ16 → 8,04 cm²,  3Φ20 → 9,42 cm².")
section(fs, 13, "=FCD(f_ck; α_cc) → αντοχή σχεδιασμού σκυροδέματος f_cd = α_cc·f_ck/1,5 (MPa)")
for j, h in enumerate(["f_ck (MPa)", "α_cc", "f_cd (MPa)"], start=1):
    fs.cell(14, j, h).font = Font(bold=True)
for i, (fck, acc) in enumerate([(20, 0.85), (25, 0.85), (30, 1.00)], start=15):
    fs.cell(i, 1, fck).border = BOX
    c = fs.cell(i, 2, acc)
    c.border = BOX
    c.number_format = "0.00"
    input_cell(fs.cell(i, 3), "0.00")
note(fs, 18, "Έλεγχος: f_ck = 25, α_cc = 0,85 → f_cd = 14,17 MPa.")
widths(fs, A=12, B=12, C=12)

# ================= Goal Seek sheet (loop macro) =================
gs = wb.create_sheet("Αναζήτηση")
title(gs, "Μάθημα 13 – Γ. Μακροεντολή που τρέχει την Αναζήτηση στόχου (Μάθημα 10)",
      "Τρεις διατομές 0,30×0,60 (d = 0,55 m), C25/30, B500C. Η μακροεντολή SolveAll λύνει το F_c − F_s = 0 για όλες.")
for c, h in ((3, "4Φ20"), (4, "6Φ20"), (5, "6Φ25")):
    gs.cell(4, c, h).font = Font(bold=True)
rows = [  # (label, value or formula template with {c} = column letter, number format, kind)
    ("b (m)", 0.30, "0.00", "in"), ("d (m)", 0.55, "0.00", "in"), ("As (cm²)", None, "0.00", "in"),
    ("f_ck (MPa)", 25, "0", "in"), ("α_cc", 0.85, "0.00", "in"), ("f_yk (MPa)", 500, "0", "in"),
    ("E_s (MPa)", 200000, "0", "in"), ("ε_cu", 0.0035, "0.0000", "in"), ("λ", 0.8, "0.00", "in"),
    ("f_cd (MPa)", "={c}9*{c}8/1.5", "0.00", "f"), ("f_yd (MPa)", "={c}10/1.15", "0.00", "f"),
    ("ε_yd", "={c}15/{c}11", "0.00000", "f"),
    ("x (m) – ΑΓΝΩΣΤΟΣ", 0.10, "0.0000", "x"),
    ("ε_s", "={c}12*({c}6-{c}17)/{c}17", "0.00000", "f"),
    ("σ_s (MPa)", "=MIN({c}11*{c}18,{c}15)", "0.00", "f"),
    ("F_c (kN)", "={c}13*{c}17*{c}5*{c}14*1000", "0.00", "f"),
    ("F_s (kN)", "={c}7/10000*{c}19*1000", "0.00", "f"),
    ("M_Rd (kNm)", "={c}21*({c}6-{c}13*{c}17/2)", "0.00", "f"),
    ("F_c − F_s (kN) – ΣΤΟΧΟΣ = 0", "={c}20-{c}21", "0.000", "f"),
    ("Διαρρέει ο χάλυβας;", '=IF({c}18>={c}16,"ΝΑΙ","ΟΧΙ")', "@", "f"),
]
As = {"C": 12.57, "D": 18.85, "E": 29.45}
for r, (lab, val, fmt, kind) in enumerate(rows, start=5):
    gs.cell(r, 1, lab)
    for col in "CDE":
        c = gs[f"{col}{r}"]
        v = As[col] if lab.startswith("As") else val
        if kind == "in":
            input_cell(c, fmt)
        else:
            c.border = BOX
            c.number_format = fmt
            c.fill = UNKNOWN_FILL if kind == "x" else FORMULA_FILL
        c.value = v.format(c=col) if isinstance(v, str) else v
for col in "CDE":
    gs[f"{col}17"].font = Font(bold=True)
    gs[f"{col}23"].font = Font(bold=True, color="C00000")
note(gs, 26, "Έλεγχος μετά την εκτέλεση (α_cc = 0,85):  4Φ20 → x = 0,1607 m, M_Rd = 265,4 kNm ·  "
             "6Φ20 → x = 0,2410 m, M_Rd = 371,7 kNm ·  6Φ25 → x = 0,3491 m, ΔΕΝ διαρρέει, M_Rd = 487,0 kNm.")
note(gs, 27, "Αλλάξτε ένα δεδομένο (π.χ. f_ck) και ξανατρέξτε τη μακροεντολή – ή πατήστε το κουμπί που θα φτιάξετε.")
widths(gs, A=30, B=3, C=12, D=12, E=12)

save(wb, "13-macros.xlsx")
