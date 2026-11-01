# Εισαγωγή στο Excel για Πολιτικούς Μηχανικούς

Εισαγωγικά μαθήματα Excel για φοιτητές του 1ου εξαμήνου Πολιτικών Μηχανικών.
Στατικό site (HTML/CSS/JS, χωρίς build) – φιλοξενείται στο GitHub Pages:
https://panagop.github.io/excel-basics/

## Δομή

- `index.html` – αρχική σελίδα
- `lessons/` – μία σελίδα ανά μάθημα
- `guides/` – οδηγοί εκτός σειράς μαθημάτων (π.χ. τοπικές ρυθμίσεις)
- `files/` – αρχεία .xlsx παραδειγμάτων/ασκήσεων
- `assets/js/site.js` – λίστα μαθημάτων (μενού, κουμπιά Προηγούμενο/Επόμενο) και διαδραστικό φύλλο
- `tools/` – scripts που δημιουργούν τα .xlsx, π.χ. `uv run tools/make_02_formulas.py`

## Νέο μάθημα

1. Δημιουργήστε `lessons/NN-name.html` (αντιγράψτε ένα υπάρχον).
2. Ορίστε `ready: true` στην αντίστοιχη γραμμή του `LESSONS` στο `assets/js/site.js`.
3. Για οδηγό: `guides/name.html` και μια γραμμή στο `GUIDES` του `site.js`.
4. Προαιρετικά: `tools/make_NN_name.py` για το αρχείο .xlsx.
