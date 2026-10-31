# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl"]
# ///
"""Builds files/12-more-charts.xlsx.  Run with:  uv run tools/make_12_more_charts.py"""

import random

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

from xlhelpers import (BOX, FORMULA_FILL, input_cell, note, save, section, table, title, widths)

MONTHS = ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μαΐ", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ"]


def axes_on(ch, x_title=None, y_title=None):
    ch.x_axis.delete = False  # recent Excel versions hide axes created by openpyxl otherwise
    ch.y_axis.delete = False
    if x_title:
        ch.x_axis.title = x_title
    if y_title:
        ch.y_axis.title = y_title


def labels_only(value=False, percent=False):
    """Data labels with only the value or only the percentage (other fields explicitly off)."""
    return DataLabelList(showVal=value, showPercent=percent, showCatName=False, showSerName=False,
                         showLegendKey=False, showLeaderLines=False)


def bar(ws, anchor, title_, cats, data, x_title=None, y_title=None, horizontal=False, stacked=False,
        gap=80, labels=False, legend=True):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title_
    ch.style = 10
    ch.width, ch.height = 15, 7.5
    ch.gapWidth = gap
    if stacked:
        ch.grouping = "stacked"
        ch.overlap = 100
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    axes_on(ch, x_title, y_title)
    if labels:
        ch.dataLabels = labels_only(value=True)
    if not legend:
        ch.legend = None
    ws.add_chart(ch, anchor)
    return ch


def cube_results(seed, n, mean, sd):
    rnd = random.Random(seed)
    return [round(rnd.gauss(mean, sd), 1) for _ in range(n)]


wb = Workbook()

# ================= Example sheet =================
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 12 – Γραφήματα στηλών, ράβδων, πίτας, γραμμής και ιστογράμματα",
      "Κάντε κλικ σε ένα γράφημα και δείτε την καρτέλα Σχεδίαση γραφήματος: Αλλαγή τύπου γραφήματος.")

# --- A. Concrete per floor: clustered vs stacked columns ---
section(ws, 4, "Α. Σκυρόδεμα ανά όροφο και είδος στοιχείου (m³) – από την προμέτρηση του Μαθήματος 9")
table(ws, 5, ["Όροφος", "Υποστυλώματα", "Δοκοί", "Πλάκες"],
      [["Ισόγειο", 2.835, 3.41, 12.6], ["1ος", 1.92, 3.41, 12.6], ["2ος", 1.92, 3.41, 12.6]], fmt="0.00")
cats = Reference(ws, min_col=1, min_row=6, max_row=8)
data = Reference(ws, min_col=2, max_col=4, min_row=5, max_row=8)
bar(ws, "G4", "Ομαδοποιημένες στήλες", cats, data, y_title="V (m³)")
bar(ws, "G20", "Στοιβαγμένες στήλες: σύνολο ανά όροφο", cats, data, y_title="V (m³)", stacked=True)
note(ws, 10, "Ομαδοποιημένες: σύγκριση των ειδών μέσα σε κάθε όροφο.")
note(ws, 11, "Στοιβαγμένες: το ύψος κάθε στήλης είναι το σύνολο του ορόφου.")

# --- B. Cost breakdown: pie and sorted bar ---
section(ws, 37, "Β. Κατανομή κόστους κατασκευής κτιρίου (ενδεικτικά ποσά)")
cost = [["Φέρων οργανισμός", 150000], ["Η/Μ εγκαταστάσεις", 110000], ["Τοιχοποιίες – επιχρίσματα", 75000],
        ["Δάπεδα – χρωματισμοί", 75000], ["Κουφώματα", 50000], ["Εκσκαφές – θεμελίωση", 40000]]
table(ws, 38, ["Κατηγορία εργασιών", "Κόστος (€)", "Ποσοστό"], cost, {"C": "=B{r}/SUM($B$39:$B$44)"},
      fmt="#,##0", formula_fmt="0.0%")
ws["A45"] = "Σύνολο"
ws["B45"] = "=SUM(B39:B44)"
ws["B45"].number_format = "#,##0"
ws["B45"].fill = FORMULA_FILL
ws["B45"].border = BOX
pie = PieChart()
pie.title = "Κατανομή κόστους"
pie.width, pie.height = 15, 8
pie.add_data(Reference(ws, min_col=2, min_row=38, max_row=44), titles_from_data=True)
pie.set_categories(Reference(ws, min_col=1, min_row=39, max_row=44))
pie.dataLabels = labels_only(percent=True)
ws.add_chart(pie, "G37")
ch = bar(ws, "G54", "Το ίδιο με ράβδους, ταξινομημένο", Reference(ws, min_col=1, min_row=39, max_row=44),
         Reference(ws, min_col=2, min_row=38, max_row=44), x_title=None, y_title="Κόστος (€)",
         horizontal=True, labels=True, legend=False)
ch.x_axis.scaling.orientation = "maxMin"  # largest category on top
note(ws, 47, "Η πίτα δείχνει μερίδια ενός συνόλου. Οι ράβδοι επιτρέπουν ακριβέστερη σύγκριση –")
note(ws, 48, "π.χ. ότι Τοιχοποιίες και Δάπεδα είναι ίσα φαίνεται αμέσως στις ράβδους, όχι στην πίτα.")

# --- C. Project S-curve: column + cumulative line (combo) ---
section(ws, 71, "Γ. Χρονοδιάγραμμα δαπανών έργου: μηνιαίες δαπάνες (στήλες) και αθροιστική καμπύλη S (γραμμή)")
monthly = [15, 25, 35, 50, 60, 65, 60, 55, 45, 35, 30, 25]
table(ws, 72, ["Μήνας", "Δαπάνη μήνα (χιλ. €)", "Αθροιστικά (χιλ. €)"],
      [[m, v] for m, v in zip(MONTHS, monthly)], {"C": "=SUM($B$73:B{r})"}, fmt="0", formula_fmt="0")
c1 = BarChart()
c1.title = "Μηνιαίες και αθροιστικές δαπάνες"
c1.style, c1.width, c1.height = 10, 17, 8.5
c1.add_data(Reference(ws, min_col=2, min_row=72, max_row=84), titles_from_data=True)
c1.set_categories(Reference(ws, min_col=1, min_row=73, max_row=84))
axes_on(c1, y_title="Μηνιαία (χιλ. €)")
c2 = LineChart()
c2.add_data(Reference(ws, min_col=3, min_row=72, max_row=84), titles_from_data=True)
c2.y_axis.axId = 200
c2.y_axis.title = "Αθροιστικά (χιλ. €)"
c2.y_axis.crosses = "max"  # secondary axis on the right
c2.y_axis.delete = False
c1 += c2
ws.add_chart(c1, "G71")
note(ws, 86, "Γράφημα συνδυασμού: δύο σειρές με πολύ διαφορετική κλίμακα → η αθροιστική στον δευτερεύοντα άξονα (δεξιά).")

# --- D. Histogram of cube results ---
section(ws, 90, "Δ. Ιστόγραμμα 40 αποτελεσμάτων θλίψης κύβων (MPa)")
cubes = cube_results(5, 40, 33, 2.8)
ws["A91"] = "f (MPa)"
for i, v in enumerate(cubes):
    c = ws.cell(92 + i % 20, 1 + i // 20, v)
    c.number_format = "0.0"
    c.border = BOX
rng = "$A$92:$B$111"
table(ws, 91, ["Από (≥)", "Έως (<)", "Πλήθος"], [[lo, lo + 2] for lo in range(26, 40, 2)],
      {"F": f'=COUNTIFS({rng},">="&D{{r}},{rng},"<"&E{{r}})'}, fmt="0", formula_fmt="0", col0=4)
ws["D99"] = "Σύνολο"
ws["F99"] = "=SUM(F92:F98)"
ws["F99"].fill = FORMULA_FILL
ws["D100"], ws["F100"] = "Μέση τιμή", f"=AVERAGE({rng})"
ws["D101"], ws["F101"] = "Τυπική απόκλιση", f"=_xlfn.STDEV.S({rng})"  # _xlfn: function newer than Excel 2007
for r in (100, 101):
    ws[f"F{r}"].number_format = "0.00"
    ws[f"F{r}"].fill = FORMULA_FILL
ws["G91"] = "Ετικέτα"
for i, r in enumerate(range(92, 99)):
    ws.cell(r, 7, f'=D{r}&"–"&E{r}').fill = FORMULA_FILL
bar(ws, "I90", "Ιστόγραμμα αντοχής κύβων", Reference(ws, min_col=7, min_row=92, max_row=98),
    Reference(ws, min_col=6, min_row=91, max_row=98), x_title="f (MPa)", y_title="Πλήθος δοκιμίων",
    gap=5, legend=False)
note(ws, 103, col=4, text='F92: =COUNTIFS(A92:B111;">="&D92;A92:B111;"<"&E92)  – πόσες τιμές είναι στο διάστημα [Από, Έως).')
note(ws, 104, col=4, text="Στο Excel 2016 και νεότερα υπάρχει και έτοιμο γράφημα: Εισαγωγή → Ιστόγραμμα (Histogram).")
widths(ws, A=26, B=18, C=18, D=10, E=10, F=10, G=10)

# ================= Exercise sheet =================
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Επιλογή και δημιουργία γραφημάτων",
      "Για κάθε ενότητα δημιουργήστε το γράφημα που ζητείται, με τίτλο και τίτλους αξόνων.")

section(ex, 4, "Α. Βάρος οπλισμού (kg) ανά όροφο και είδος στοιχείου → στοιβαγμένες στήλες")
table(ex, 5, ["Όροφος", "Θεμελίωση", "Υποστυλώματα", "Δοκοί", "Πλάκες"],
      [["Υπόγειο", 4200, 1350, 0, 980], ["Ισόγειο", 0, 1280, 1650, 1420], ["1ος", 0, 1150, 1600, 1420],
       ["2ος", 0, 1020, 1580, 1380]], fmt="#,##0")
note(ex, 10, "Ποιος όροφος έχει τον περισσότερο οπλισμό; (Έλεγχος: το υπόγειο, 6 530 kg.)")

section(ex, 13, "Β. Κατανομή χρόνου εργασιών ενός έργου (ημέρες) → πίτα ΚΑΙ ταξινομημένες ράβδοι")
table(ex, 14, ["Φάση", "Ημέρες"],
      [["Εκσκαφές", 15], ["Θεμελίωση", 25], ["Σκελετός", 90], ["Τοιχοποιίες", 40], ["Η/Μ", 60], ["Τελειώματα", 70]],
      fmt="0")
note(ex, 21, "Ποιο γράφημα σας βοηθά να συγκρίνετε Η/Μ και Τελειώματα; Ταξινομήστε τον πίνακα πριν φτιάξετε τις ράβδους.")

section(ex, 24, "Γ. Ιστόγραμμα 30 αποτελεσμάτων θλίψης κύβων – συμπληρώστε τα πλήθη με COUNTIFS")
for i, v in enumerate(cube_results(9, 30, 29.5, 2.5)):
    c = ex.cell(26 + i % 15, 1 + i // 15, v)
    c.number_format = "0.0"
    c.border = BOX
ex["A25"] = "f (MPa)"
table(ex, 25, ["Από (≥)", "Έως (<)", "Πλήθος"], [[lo, lo + 2] for lo in range(22, 36, 2)],
      {"F": ""}, student=True, fmt="0", formula_fmt="0", col0=4)
ex["D33"] = "Σύνολο"
input_cell(ex["F33"], "0")
note(ex, 42, "Έλεγχος: το σύνολο πρέπει να είναι 30· το διάστημα 28–30 έχει τα περισσότερα δοκίμια (9).")

section(ex, 45, "Δ. Σκυροδετήσεις έργου (m³ ανά εβδομάδα) → στήλες + αθροιστική γραμμή σε δευτερεύοντα άξονα")
weekly = [12, 30, 45, 0, 38, 52, 60, 25, 40, 18]
table(ex, 46, ["Εβδομάδα", "m³ εβδομάδας", "Αθροιστικά (m³)"], [[i + 1, v] for i, v in enumerate(weekly)],
      {"C": ""}, student=True, fmt="0", formula_fmt="0")
note(ex, 57, "Έλεγχος: συνολικά 320 m³. Γράφημα: Εισαγωγή → Γράφημα συνδυασμού (Combo).")
widths(ex, A=14, B=14, C=16, D=10, E=10, F=10)

save(wb, "12-more-charts.xlsx")
