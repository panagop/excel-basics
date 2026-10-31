# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/10-goal-seek.xlsx.  Run with:  uv run tools/make_10_goal_seek.py

The unknowns are stored with a starting guess; the student runs Goal Seek to solve.
"""

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.styles import Font, PatternFill

from xlhelpers import (BOX, FORMULA_FILL, as_text, input_cell, note, save, section, title, widths)

UNKNOWN_FILL = PatternFill("solid", fgColor="F4B183")  # the cell Goal Seek changes
TARGET_FONT = Font(bold=True, color="C00000")  # the cell Goal Seek sets


def block(ws, top, rows):
    """Calculation block: (label, symbol, value or formula, unit, kind, number format).

    kind: "in" data, "x" unknown (Goal Seek changes it), "f" formula, "t" target formula.
    Columns: A label, B symbol, D value, E unit, F formula as text.
    """
    for r, (label, sym, val, unit, kind, fmt) in enumerate(rows, start=top):
        ws.cell(r, 1, label)
        ws.cell(r, 2, sym)
        c = ws.cell(r, 4, val)
        c.number_format = fmt
        c.border = BOX
        ws.cell(r, 5, unit)
        if kind == "in":
            input_cell(c, fmt)
        elif kind == "x":
            c.fill = UNKNOWN_FILL
            c.font = Font(bold=True)
        else:
            c.fill = FORMULA_FILL
            if kind == "t":
                c.font = TARGET_FONT
            as_text(ws.cell(r, 6), greek(val))
    return top + len(rows)


def greek(formula):
    return formula.replace(",", ";").replace("0.", "0,").replace("1.", "1,")


def legend(ws, row):
    for i, (text, fill, font) in enumerate([("δεδομένο", "FFF2CC", None), ("άγνωστος (αλλάζει)", "F4B183", True),
                                            ("τύπος", "E8F3EC", None), ("στόχος (ορίζεται)", "E8F3EC", "C00000")]):
        c = ws.cell(row, 1 + i, text)
        c.fill = PatternFill("solid", fgColor=fill)
        c.border = BOX
        if font == "C00000":
            c.font = Font(bold=True, color="C00000")
        elif font:
            c.font = Font(bold=True)


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 10 – Αναζήτηση στόχου (Goal Seek)",
      "Δεδομένα → Ανάλυση πιθανοτήτων → Αναζήτηση στόχου. Ορισμός κελιού = κόκκινο, Με αλλαγή του κελιού = πορτοκαλί.")
legend(ws, 3)

# --- A. Timber beam height ---
section(ws, 5, "Α. Ύψος ξύλινης δοκού ώστε σ = M/W = f_d   (στόχος: D11 = 15 με αλλαγή του D9)")
block(ws, 6, [
    ("Ροπή σχεδιασμού", "M", 40, "kNm", "in", "0.00"),
    ("Πλάτος διατομής", "b", 0.20, "m", "in", "0.00"),
    ("Αντοχή σχεδιασμού (δίνεται)", "f_d", 15, "MPa", "in", "0.00"),
    ("Ύψος διατομής – ΑΓΝΩΣΤΟΣ", "h", 0.30, "m", "x", "0.0000"),
    ("Ροπή αντίστασης", "W", "=D7*D9^2/6", "m³", "f", "0.000000"),
    ("Τάση", "σ", "=D6/D10/1000", "MPa", "t", "0.000"),
    ("Έλεγχος: αναλυτική λύση √(6M/(b·f_d))", "h", "=SQRT(6*D6/(D7*D8*1000))", "m", "f", "0.0000"),
])
note(ws, 13, "Λύση: h = 0,2828 m. Εδώ υπάρχει και αναλυτική λύση – τη χρησιμοποιούμε για να εμπιστευτούμε το εργαλείο.")

# --- B. Manning ---
section(ws, 16, "Β. Βάθος ροής σε ορθογωνικό κανάλι (Manning)   (στόχος: D27 = 5 με αλλαγή του D21)")
block(ws, 17, [
    ("Παροχή σχεδιασμού", "Q", 5, "m³/s", "in", "0.00"),
    ("Πλάτος πυθμένα", "b", 3.0, "m", "in", "0.00"),
    ("Συντελεστής Manning", "n", 0.015, "s/m^(1/3)", "in", "0.000"),
    ("Κλίση πυθμένα", "S", 0.001, "–", "in", "0.0000"),
    ("Βάθος ροής – ΑΓΝΩΣΤΟΣ", "y", 0.50, "m", "x", "0.0000"),
    ("Υγρή διατομή", "A", "=D18*D21", "m²", "f", "0.000"),
    ("Βρεχόμενη περίμετρος", "P", "=D18+2*D21", "m", "f", "0.000"),
    ("Υδραυλική ακτίνα", "R", "=D22/D23", "m", "f", "0.000"),
    ("Ταχύτητα", "V", "=1/D19*D24^(2/3)*D20^0.5", "m/s", "f", "0.000"),
    ("Παροχή για το βάθος y", "Q(y)", "=D25*D22", "m³/s", "f", "0.000"),
    ("Διαφορά από τη ζητούμενη", "Q(y) − Q", "=D26-D17", "m³/s", "t", "0.0000"),
])
note(ws, 28, "Λύση: y = 1,079 m. Εδώ ΔΕΝ υπάρχει κλειστή λύση: ο άγνωστος y είναι μέσα στο A, στο P και στο R^(2/3).")
note(ws, 29, "Ο στόχος μπορεί να είναι η διαφορά = 0 (όπως εδώ) ή απευθείας Q(y) = 5 στο D26 – δοκιμάστε και τα δύο.")
# Q(y) curve for the chart
ws["H17"], ws["I17"] = "y (m)", "Q(y) (m³/s)"
for c in ("H17", "I17"):
    ws[c].font = Font(bold=True)
for i, r in enumerate(range(18, 28)):
    ws.cell(r, 8, 0.2 * (i + 1)).number_format = "0.0"
    c = ws.cell(r, 9, f"=$D$18*H{r}*($D$18*H{r}/($D$18+2*H{r}))^(2/3)*$D$20^0.5/$D$19")
    c.number_format = "0.00"
    c.fill = FORMULA_FILL
ch = ScatterChart()
ch.title = "Παροχή Q ως συνάρτηση του βάθους y"
ch.style, ch.width, ch.height = 2, 14, 7.5
ch.x_axis.title, ch.y_axis.title = "y (m)", "Q (m³/s)"
ch.x_axis.delete = ch.y_axis.delete = False
s = Series(Reference(ws, min_col=9, min_row=18, max_row=27), Reference(ws, min_col=8, min_row=18, max_row=27), title="Q(y)")
s.smooth = True
s.marker.symbol = "circle"
ch.series.append(s)
ch.legend = None
ws.add_chart(ch, "K16")

# --- C. Neutral axis depth ---
section(ws, 32, "Γ. Ύψος θλιβόμενης ζώνης x ορθογωνικής διατομής – ισορροπία F_c = F_s   (στόχος: D52 = 0 με αλλαγή του D45)")
block(ws, 33, [
    ("Πλάτος διατομής", "b", 0.30, "m", "in", "0.00"),
    ("Στατικό ύψος", "d", 0.55, "m", "in", "0.00"),
    ("Οπλισμός (4Φ20)", "As", 12.57, "cm²", "in", "0.00"),
    ("Χαρακτηριστική αντοχή σκυροδέματος", "f_ck", 25, "MPa", "in", "0"),
    ("Συντελεστής α_cc (συνιστώμενη τιμή EC2)", "α_cc", 1.00, "–", "in", "0.00"),
    ("Χαρακτηριστικό όριο διαρροής χάλυβα", "f_yk", 500, "MPa", "in", "0"),
    ("Μέτρο ελαστικότητας χάλυβα", "E_s", 200000, "MPa", "in", "0"),
    ("Οριακή βράχυνση σκυροδέματος", "ε_cu", 0.0035, "–", "in", "0.0000"),
    ("Ορθογωνικό διάγραμμα: λ", "λ", 0.8, "–", "in", "0.00"),
    ("f_cd = α_cc·f_ck/1,5", "f_cd", "=D37*D36/1.5", "MPa", "f", "0.00"),
    ("f_yd = f_yk/1,15", "f_yd", "=D38/1.15", "MPa", "f", "0.00"),
    ("ε_yd = f_yd/E_s", "ε_yd", "=D43/D39", "–", "f", "0.00000"),
    ("Ύψος θλιβόμενης ζώνης – ΑΓΝΩΣΤΟΣ", "x", 0.10, "m", "x", "0.0000"),
    ("Παραμόρφωση χάλυβα ε_s = ε_cu·(d − x)/x", "ε_s", "=D40*(D34-D45)/D45", "–", "f", "0.00000"),
    ("Τάση χάλυβα σ_s = min(E_s·ε_s ; f_yd)", "σ_s", "=MIN(D39*D46,D43)", "MPa", "f", "0.00"),
    ("Δύναμη σκυροδέματος F_c = λ·x·b·f_cd", "F_c", "=D41*D45*D33*D42*1000", "kN", "f", "0.00"),
    ("Δύναμη χάλυβα F_s = As·σ_s", "F_s", "=D35/10000*D47*1000", "kN", "f", "0.00"),
    ("Διαρρέει ο χάλυβας;", "", '=IF(D46>=D44,"ΝΑΙ","ΟΧΙ")', "", "f", "@"),
    ("Ροπή αντοχής M_Rd = F_s·(d − λ·x/2)", "M_Rd", "=D49*(D34-D41*D45/2)", "kNm", "f", "0.00"),
    ("Ισορροπία F_c − F_s", "ΔF", "=D48-D49", "kN", "t", "0.000"),
])
note(ws, 54, "Λύση: x = 0,1366 m, ο χάλυβας διαρρέει, M_Rd = 270,7 kNm.")
note(ws, 55, "Αφού ο χάλυβας διαρρέει, υπάρχει και αναλυτική λύση x = As·f_yd/(λ·b·f_cd) – ελέγξτε την!")
note(ws, 56, "Μην ξεκινάτε από x = 0: το ε_s = ε_cu·(d − x)/x δίνει διαίρεση με μηδέν.")
widths(ws, A=44, B=10, C=4, D=13, E=11, F=34, H=8, I=12)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Αναζήτηση στόχου",
      "Συμπληρώστε τους τύπους στα πράσινα κελιά και λύστε με Αναζήτηση στόχου (πορτοκαλί = άγνωστος).")
legend(ex, 3)


def student_block(ws, top, rows):
    """Like block(), but formula cells are left empty (green) for the student."""
    for r, (label, sym, val, unit, kind, fmt) in enumerate(rows, start=top):
        ws.cell(r, 1, label)
        ws.cell(r, 2, sym)
        c = ws.cell(r, 4)
        c.number_format = fmt
        c.border = BOX
        ws.cell(r, 5, unit)
        if kind == "in":
            input_cell(c, fmt)
            c.value = val
        elif kind == "x":
            c.value = val
            c.fill = UNKNOWN_FILL
            c.font = Font(bold=True)
        else:
            c.fill = FORMULA_FILL
            if kind == "t":
                c.font = TARGET_FONT


section(ex, 5, "Α. Πλευρά τετραγωνικού πεδίλου ώστε σ = N/B² = σ_επ")
student_block(ex, 6, [
    ("Αξονικό φορτίο", "N", 1200, "kN", "in", "0"),
    ("Επιτρεπόμενη τάση εδάφους (δίνεται)", "σ_επ", 250, "kPa", "in", "0"),
    ("Πλευρά πεδίλου – ΑΓΝΩΣΤΟΣ", "B", 1.50, "m", "x", "0.0000"),
    ("Τάση εδάφους σ = N/B²", "σ", None, "kPa", "t", "0.00"),
    ("Αναλυτική λύση B = √(N/σ_επ)", "B", None, "m", "f", "0.0000"),
])
note(ex, 11, "Έλεγχος: B = 2,191 m. (Στην πράξη θα στρογγυλεύατε προς τα πάνω, π.χ. σε 2,20 m.)")

section(ex, 14, "Β. Βάθος ροής σε ορθογωνικό κανάλι (Manning):  Q = (1/n)·A·R^(2/3)·S^(1/2)")
student_block(ex, 15, [
    ("Παροχή σχεδιασμού", "Q", 12, "m³/s", "in", "0.00"),
    ("Πλάτος πυθμένα", "b", 4.0, "m", "in", "0.00"),
    ("Συντελεστής Manning", "n", 0.014, "s/m^(1/3)", "in", "0.000"),
    ("Κλίση πυθμένα", "S", 0.0005, "–", "in", "0.0000"),
    ("Βάθος ροής – ΑΓΝΩΣΤΟΣ", "y", 1.00, "m", "x", "0.0000"),
    ("Υγρή διατομή A = b·y", "A", None, "m²", "f", "0.000"),
    ("Βρεχόμενη περίμετρος P = b + 2y", "P", None, "m", "f", "0.000"),
    ("Υδραυλική ακτίνα R = A/P", "R", None, "m", "f", "0.000"),
    ("Παροχή Q(y)", "Q(y)", None, "m³/s", "t", "0.000"),
    ("Μέση ταχύτητα V = Q/A", "V", None, "m/s", "f", "0.000"),
])
note(ex, 25, "Έλεγχος: y = 1,908 m, V = 1,572 m/s.")

section(ex, 28, "Γ. Ύψος θλιβόμενης ζώνης δύο διατομών 0,25 × 0,50 (d = 0,45 m), C20/25, B500C – μία Αναζήτηση στόχου ανά στήλη")
ex["D29"], ex["E29"] = "6Φ25", "3Φ20"
for c in ("D29", "E29"):
    ex[c].font = Font(bold=True)
rows = [
    ("Πλάτος", "b", 0.25, "in", "0.00"), ("Στατικό ύψος", "d", 0.45, "in", "0.00"),
    ("Οπλισμός", "As (cm²)", (29.45, 9.42), "in", "0.00"), ("f_ck (MPa)", "", 20, "in", "0"),
    ("α_cc", "", 1.00, "in", "0.00"), ("f_yk (MPa)", "", 500, "in", "0"), ("E_s (MPa)", "", 200000, "in", "0"),
    ("ε_cu", "", 0.0035, "in", "0.0000"), ("λ", "", 0.8, "in", "0.00"),
    ("f_cd (MPa)", "", None, "f", "0.00"), ("f_yd (MPa)", "", None, "f", "0.00"), ("ε_yd", "", None, "f", "0.00000"),
    ("x (m) – ΑΓΝΩΣΤΟΣ", "", 0.10, "x", "0.0000"),
    ("ε_s", "", None, "f", "0.00000"), ("σ_s = min(E_s·ε_s ; f_yd) (MPa)", "", None, "f", "0.00"),
    ("F_c (kN)", "", None, "f", "0.00"), ("F_s (kN)", "", None, "f", "0.00"),
    ("Διαρρέει ο χάλυβας;", "", None, "f", "@"), ("M_Rd (kNm)", "", None, "f", "0.00"),
    ("Ισορροπία F_c − F_s (kN)", "", None, "t", "0.000"),
]
for r, (label, sym, val, kind, fmt) in enumerate(rows, start=30):
    ex.cell(r, 1, label)
    ex.cell(r, 2, sym)
    for k, col in enumerate((4, 5)):
        c = ex.cell(r, col)
        c.number_format = fmt
        c.border = BOX
        v = val[k] if isinstance(val, tuple) else val
        if kind == "in":
            input_cell(c, fmt)
            c.value = v
        elif kind == "x":
            c.value = v
            c.fill = UNKNOWN_FILL
            c.font = Font(bold=True)
        else:
            c.fill = FORMULA_FILL
            if kind == "t":
                c.font = TARGET_FONT
note(ex, 50, "Έλεγχος: 6Φ25 → x = 0,3187 m, ο χάλυβας ΔΕΝ διαρρέει, M_Rd = 274,1 kNm.   3Φ20 → x = 0,1536 m, M_Rd = 159,1 kNm.")
note(ex, 51, "Ερώτηση: τριπλασιάσαμε σχεδόν τον οπλισμό· γιατί η ροπή αντοχής αυξήθηκε μόνο κατά ~70%;")
widths(ex, A=40, B=10, C=4, D=13, E=13, F=12)

save(wb, "10-goal-seek.xlsx")
