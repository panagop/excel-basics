# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/11-matrices.xlsx.  Run with:  uv run tools/make_11_matrices.py

Matrix results are stored as array formulas (Ctrl+Shift+Enter style), so they work in every
Excel version; Microsoft 365 shows them as ordinary formulas.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.formula import ArrayFormula

from xlhelpers import (BOX, FORMULA_FILL, HEADER_FILL, as_text, input_cell, note, save, section,
                       title, widths)


def label(ws, ref, text):
    c = ws[ref]
    c.value = text
    c.font = Font(bold=True)


def matrix(ws, top, left, values=None, shape=None, fmt="0.###", kind="in"):
    """Block of cells starting at (top, left).

    kind "in": data values (yellow) · "mixed": formulas green, numbers plain ·
    "student": empty yellow block of the given shape (rows, cols).
    """
    rows, cols = shape if shape else (len(values), len(values[0]))
    for i in range(rows):
        for j in range(cols):
            c = ws.cell(top + i, left + j)
            if kind == "student":
                input_cell(c, fmt)
                continue
            v = values[i][j]
            c.value = v
            c.border = BOX
            c.number_format = fmt
            if kind == "in":
                input_cell(c, fmt)
            elif isinstance(v, str):
                c.fill = FORMULA_FILL


def array(ws, top, left, text, rows, cols, fmt="0.###"):
    """One array formula (Ctrl+Shift+Enter) over a rows×cols block; green."""
    for i in range(rows):
        for j in range(cols):
            c = ws.cell(top + i, left + j)
            c.fill = FORMULA_FILL
            c.border = BOX
            c.number_format = fmt
    first = ws.cell(top, left).coordinate
    last = ws.cell(top + rows - 1, left + cols - 1).coordinate
    ws[first] = ArrayFormula(f"{first}:{last}", text)


def show(ws, ref, text):
    """Formula as typed with Greek settings, as a note next to a result."""
    as_text(ws[ref], text.replace(",", ";"))
    ws[ref].font = Font(italic=True, color="595959")


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 11 – Μητρώα και γραμμικά συστήματα",
      "Τα πράσινα μπλοκ είναι τύποι πίνακα: επιλέξτε ένα κελί τους και δείτε τη Γραμμή τύπων.")

# --- A. Basic operations ---
section(ws, 4, "Α. Πράξεις μητρώων και επίλυση του συστήματος A·x = b")
label(ws, "B5", "A")
matrix(ws, 6, 2, [[4, -2, 1], [-2, 4, -2], [1, -2, 4]])
label(ws, "F5", "b")
matrix(ws, 6, 6, [[11], [-16], [17]])
label(ws, "H5", "Ανάστροφο Aᵀ")
array(ws, 6, 8, "=TRANSPOSE(B6:D8)", 3, 3)
show(ws, "H9", "=TRANSPOSE(B6:D8)")
label(ws, "L5", "det(A)")
ws["L6"] = "=MDETERM(B6:D8)"
ws["L6"].fill = FORMULA_FILL
ws["L6"].border = BOX
show(ws, "L7", "=MDETERM(B6:D8)")

label(ws, "B11", "Αντίστροφο A⁻¹")
array(ws, 12, 2, "=MINVERSE(B6:D8)", 3, 3, "0.0000")
show(ws, "B15", "=MINVERSE(B6:D8)")
label(ws, "H11", "Έλεγχος A·A⁻¹ = I")
array(ws, 12, 8, "=MMULT(B6:D8,B12:D14)", 3, 3, "0.000")
show(ws, "H15", "=MMULT(B6:D8;B12:D14)")

label(ws, "B17", "Λύση x = A⁻¹·b")
array(ws, 18, 2, "=MMULT(MINVERSE(B6:D8),F6:F8)", 3, 1, "0.000")
show(ws, "C18", "=MMULT(MINVERSE(B6:D8),F6:F8)")
label(ws, "H17", "Έλεγχος A·x − b (≈ 0)")
array(ws, 18, 8, "=MMULT(B6:D8,B18:B20)-F6:F8", 3, 1, "0.000")
show(ws, "I18", "=MMULT(B6:D8,B18:B20)-F6:F8")
note(ws, 21, "Λύση: x = (1, −2, 3). det(A) = 36 ≠ 0, άρα το σύστημα έχει μοναδική λύση.")

# --- B. Three-storey shear frame ---
section(ws, 24, "Β. Τριώροφο διατμητικό πλαίσιο υπό οριζόντια φορτία:  K·u = F")
for j, h in enumerate(["Όροφος", "k (kN/m)", "F (kN)"], start=2):
    c = ws.cell(25, j, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
for i, (k, f) in enumerate([(60000, 50), (50000, 100), (40000, 150)]):
    ws.cell(26 + i, 2, i + 1).border = BOX
    for j, v in ((3, k), (4, f)):
        input_cell(ws.cell(26 + i, j), "0")
        ws.cell(26 + i, j).value = v
label(ws, "F25", "Μητρώο δυσκαμψίας K (kN/m)")
matrix(ws, 26, 6, [["=C26+C27", "=-C27", 0], ["=-C27", "=C27+C28", "=-C28"], [0, "=-C28", "=C28"]], fmt="0",
       kind="mixed")
show(ws, "F29", "Γραμμή 1: =C26+C27  =-C27  0")
label(ws, "J25", "u (mm)")
array(ws, 26, 10, "=MMULT(MINVERSE(F26:H28),D26:D28)*1000", 3, 1, "0.000")
show(ws, "K26", "=MMULT(MINVERSE(F26:H28),D26:D28)*1000")
note(ws, 30, "Έλεγχος χωρίς μητρώα: τέμνουσα ορόφου V_i = άθροισμα των F από πάνω· σχετική μετακίνηση Δ_i = V_i/k_i.")
table_rows = [("Όροφος", "V (kN)", "Δ = V/k (mm)", "u = ΣΔ (mm)")]
for j, h in enumerate(table_rows[0], start=2):
    c = ws.cell(31, j, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
for i in range(3):
    r = 32 + i
    ws.cell(r, 2, i + 1).border = BOX
    for j, f, fmt in ((3, f"=SUM(D{26 + i}:D28)", "0"), (4, f"=C{r}/C{26 + i}*1000", "0.000"),
                      (5, f"=SUM(D$32:D{r})", "0.000")):
        c = ws.cell(r, j, f)
        c.fill = FORMULA_FILL
        c.number_format = fmt
        c.border = BOX
note(ws, 35, "u = (5,000 ; 10,000 ; 13,750) mm και με τους δύο τρόπους.")

# --- C. Truss, method of joints in matrix form ---
section(ws, 38, "Γ. Τριγωνικό δικτύωμα – μέθοδος των κόμβων ως σύστημα 6×6")
note(ws, 39, "Κόμβοι A(0,0) άρθρωση, B(L,0) κύλιση, Γ(L/2,h) με φορτία P (κάτω) και H (δεξιά). Εφελκυσμός θετικός.")
geo = [("L (m)", 6), ("h (m)", 4), ("P (kN)", 100), ("H (kN)", 30)]
for i, (lab, v) in enumerate(geo):
    ws.cell(40 + i, 2, lab)
    input_cell(ws.cell(40 + i, 3), "0.00")
    ws.cell(40 + i, 3).value = v
for i, (lab, f) in enumerate([("Μήκος AΓ (m)", "=SQRT((C40/2)^2+C41^2)"), ("c = cos", "=C40/2/C44"),
                              ("s = sin", "=C41/C44")]):
    ws.cell(44 + i, 2, lab)
    c = ws.cell(44 + i, 3, f)
    c.fill = FORMULA_FILL
    c.number_format = "0.0000"
    c.border = BOX

unknowns = ["N_AB", "N_AΓ", "N_BΓ", "A_x", "A_y", "B_y"]
for j, u in enumerate(unknowns, start=6):
    c = ws.cell(40, j, u)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center")
for col, h in ((12, "δεξί μέλος"), (5, "Εξίσωση")):
    c = ws.cell(40, col, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
c_, s_ = "$C$45", "$C$46"
eqs = [
    ("A: ΣFx = 0", [1, f"={c_}", 0, 1, 0, 0], 0),
    ("A: ΣFy = 0", [0, f"={s_}", 0, 0, 1, 0], 0),
    ("B: ΣFx = 0", [-1, 0, f"=-{c_}", 0, 0, 0], 0),
    ("B: ΣFy = 0", [0, 0, f"={s_}", 0, 0, 1], 0),
    ("Γ: ΣFx = 0", [0, f"=-{c_}", f"={c_}", 0, 0, 0], "=-C43"),
    ("Γ: ΣFy = 0", [0, f"=-{s_}", f"=-{s_}", 0, 0, 0], "=C42"),
]
for i, (lab, coeffs, rhs) in enumerate(eqs):
    r = 41 + i
    ws.cell(r, 5, lab).border = BOX
    for j, v in enumerate(coeffs, start=6):
        c = ws.cell(r, j, v)
        c.border = BOX
        c.number_format = "0.###"
        if isinstance(v, str):
            c.fill = FORMULA_FILL
    c = ws.cell(r, 12, rhs)
    c.border = BOX
    c.number_format = "0.###"
    if isinstance(rhs, str):
        c.fill = FORMULA_FILL
note(ws, 47, "Κάθε γραμμή είναι μια εξίσωση ισορροπίας κόμβου· κάθε στήλη ο συντελεστής ενός αγνώστου.", col=5)

for j, h in enumerate(["Άγνωστος", "Τιμή (kN)", "Είδος"], start=6):
    c = ws.cell(49, j, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
for i, u in enumerate(unknowns):
    ws.cell(50 + i, 6, u).border = BOX
array(ws, 50, 7, "=MMULT(MINVERSE(F41:K46),L41:L46)", 6, 1, "0.00")
for i in range(3):
    c = ws.cell(50 + i, 8, f'=IF(G{50 + i}>0,"Εφελκυσμός","Θλίψη")')
    c.fill = FORMULA_FILL
    c.border = BOX
for i in range(3, 6):
    ws.cell(50 + i, 8, "αντίδραση").border = BOX
show(ws, "I50", "=MMULT(MINVERSE(F41:K46),L41:L46)")
note(ws, 56, "Λύση: N_AB = 52,5 (εφελκυσμός), N_AΓ = −37,5, N_BΓ = −87,5 (θλίψη), A_x = −30, A_y = 30, B_y = 70 kN.", col=5)
note(ws, 57, "Έλεγχος: A_y + B_y = P = 100 kN και A_x = −H. Αλλάξτε τα L, h, P, H – η λύση ενημερώνεται αμέσως.", col=5)
widths(ws, A=3, B=14, C=12, D=12, E=14, F=11, G=11, H=12, I=11, J=11, K=11, L=12)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Μητρώα και γραμμικά συστήματα",
      "Συμπληρώστε τα κίτρινα κελιά. Για αποτελέσματα-μητρώα: επιλέξτε όλη την περιοχή, γράψτε τον τύπο, Ctrl+Shift+Enter.")

section(ex, 4, "Α. Λύστε το σύστημα:  2x₁ + x₂ − x₃ = 1,   x₁ + 3x₂ + 2x₃ = 13,   3x₁ − x₂ + x₃ = 4")
label(ex, "B5", "A")
matrix(ex, 6, 2, shape=(3, 3), kind="student")
label(ex, "F5", "b")
matrix(ex, 6, 6, shape=(3, 1), kind="student")
label(ex, "H5", "det(A)")
input_cell(ex["H6"], "0.###")
label(ex, "J5", "x")
matrix(ex, 6, 10, shape=(3, 1), fmt="0.000", kind="student")
label(ex, "L5", "A·x − b")
matrix(ex, 6, 12, shape=(3, 1), fmt="0.000", kind="student")
note(ex, 10, "Έλεγχος: det(A) = 25, x = (1, 2, 3).")

section(ex, 13, "Β. Τριώροφο διατμητικό πλαίσιο – συντάξτε το K από τα k και λύστε K·u = F")
for j, h in enumerate(["Όροφος", "k (kN/m)", "F (kN)"], start=2):
    c = ex.cell(14, j, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
for i, (k, f) in enumerate([(80000, 40), (60000, 80), (40000, 120)]):
    ex.cell(15 + i, 2, i + 1).border = BOX
    for j, v in ((3, k), (4, f)):
        c = ex.cell(15 + i, j, v)
        c.border = BOX
        c.number_format = "0"
label(ex, "F14", "K (kN/m)")
matrix(ex, 15, 6, shape=(3, 3), fmt="0", kind="student")
label(ex, "J14", "u (mm)")
matrix(ex, 15, 10, shape=(3, 1), fmt="0.000", kind="student")
note(ex, 19, "Έλεγχος: u = (3,000 ; 6,333 ; 9,333) mm. Επιβεβαιώστε και με τέμνουσες ορόφων Δ = V/k.")

section(ex, 22, "Γ. Δικτύωμα του Παραδείγματος με L = 8 m, h = 3 m, P = 60 kN, H = 20 kN")
note(ex, 23, "Αντιγράψτε το μπλοκ του Παραδείγματος (Γ) εδώ και αλλάξτε τα δεδομένα, ή συντάξτε το μητρώο από την αρχή.")
for j, h in enumerate(["Άγνωστος", "Τιμή (kN)"], start=2):
    c = ex.cell(25, j, h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = HEADER_FILL
for i, u in enumerate(unknowns):
    ex.cell(26 + i, 2, u).border = BOX
    input_cell(ex.cell(26 + i, 3), "0.00")
note(ex, 32, "Έλεγχος: N_AB = 50,0 · N_AΓ = −37,5 · N_BΓ = −62,5 · A_x = −20,0 · A_y = 22,5 · B_y = 37,5 kN.")
widths(ex, A=3, B=12, C=12, D=12, E=4, F=11, G=11, H=11, I=4, J=11, K=4, L=11)

save(wb, "11-matrices.xlsx")
