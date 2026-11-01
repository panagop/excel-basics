# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl", "pillow"]
# ///
"""Builds files/01-first-steps.xlsx.  Run with:  uv run tools/make_01_first_steps.py"""

from openpyxl import Workbook

from xlhelpers import note, save, table, title, widths

wb = Workbook()

# ---- Example sheet ----
ws = wb.active
ws.title = "Παράδειγμα"
title(ws, "Μάθημα 1 – Δοκοί από σκυρόδεμα: εμβαδόν διατομής και όγκος",
      "Κάντε κλικ στα πράσινα κελιά και δείτε τη Γραμμή τύπων.  A = b·h,  V = A·L")
table(
    ws, 4,
    ["Δοκός", "b (m)", "h (m)", "L (m)", "A (m²)", "V (m³)"],
    [["Δ1", 0.25, 0.50, 5.00], ["Δ2", 0.30, 0.60, 6.00], ["Δ3", 0.25, 0.45, 4.50]],
    {"E": "=B{r}*C{r}", "F": "=E{r}*D{r}"},
)

# ---- Exercise sheet ----
ex = wb.create_sheet("Άσκηση")
title(ex, "Άσκηση – Υποστυλώματα από σκυρόδεμα",
      "Συμπληρώστε τα κίτρινα κελιά με τύπους:  A = b·h  (m²)  και  V = A·H  (m³).")
table(
    ex, 4,
    ["Υποστύλωμα", "b (m)", "h (m)", "H (m)", "A (m²)", "V (m³)"],
    [["Κ1", 0.40, 0.40, 3.00], ["Κ2", 0.30, 0.50, 3.00],
     ["Κ3", 0.50, 0.50, 3.20], ["Κ4", 0.25, 0.60, 2.80]],
    {"E": "", "F": ""},
    student=True,
)
note(ex, 10, "Έλεγχος: ο όγκος του υποστυλώματος Κ1 πρέπει να βγει 0,48 m³.")

for sheet in (ws, ex):
    widths(sheet, A=13, B=11, C=11, D=11, E=11, F=11)

save(wb, "01-first-steps.xlsx")
