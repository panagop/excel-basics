// Excel για Πολιτικούς Μηχανικούς – παρουσίαση ~60 λεπτών βασισμένη στο site.
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
// Requires pptxgenjs, react, react-dom, react-icons, sharp and the pptx skill apply_theme.js helper.
const { applyTheme } = require(
  "/root/.claude/skills/synced/d1ca36c3-25d8-49ed-9dbe-62c248701279_f501eb2f-e318-418a-916f-62a38034f65a/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "excel-civil.pptx";
const SITE = "panagop.github.io/excel-basics";
const AUTHOR = "Γιώργος Παναγόπουλος";
const YEAR = "2026";
const SIGN = [AUTHOR, YEAR].filter(Boolean).join(", ");
const LOGO_DIR = process.env.LOGO_DIR || "/home/user/excel-basics/assets/img";
const IHU = { path: `${LOGO_DIR}/ihu-logo.png`, ratio: 265 / 75 };
const CIVIL = { path: `${LOGO_DIR}/civil-ihu-logo.png`, ratio: 176 / 90 };
// Logos on a white panel (they are dark, the title slides are green)
function logoPanel(x, y, h) {
  const gap = 0.35, pad = 0.25, wI = h * IHU.ratio, wC = h * CIVIL.ratio;
  return [
    { rect: { x, y, w: wI + wC + gap + 2 * pad, h: h + 2 * pad, fill: { color: "FFFFFF" }, rectRadius: 0.1 } },
    { image: { x: x + pad, y: y + pad, w: wI, h, path: IHU.path } },
    { image: { x: x + pad + wI + gap, y: y + pad, w: wC, h, path: CIVIL.path } },
  ];
}

const THEME = {
  name: "Excel Civil",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2328", lt1: "FFFFFF", dk2: "185C37", lt2: "E8F3EC",
    accent1: "217346", accent2: "F4B183", accent3: "C00000", accent4: "4472C4",
    accent5: "7F7F7F", accent6: "FFC000", hlink: "185C37", folHlink: "5D6670",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 × 7.5 in
pres.title = "Excel για Πολιτικούς Μηχανικούς";
pres.author = "Τμήμα Πολιτικών Μηχανικών";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------------- Layouts ----------------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    ...logoPanel(0.8, 0.6, 0.8),
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.7, h: 1.5, fontSize: 44, bold: true,
      color: C.background1, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.9, w: 11.7, h: 1.4, fontSize: 20,
      color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.accent1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 11.7, h: 1.2, fontSize: 40, bold: true,
      color: C.background1, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.95, w: 11.7, h: 1.2, fontSize: 20,
      color: C.background2, valign: "top" }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: C.background2, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 10.9, h: 0.9, fontSize: 32, bold: true,
      color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Excel για Πολιτικούς Μηχανικούς", options: { x: 0.6, y: 6.95, w: 6, h: 0.35,
      fontSize: 11, color: C.accent5, margin: 0 } } },
    { text: { text: SIGN, options: { x: 6.85, y: 6.95, w: 3.2, h: 0.35, fontSize: 11, color: C.accent5, align: "right",
      margin: 0 } } },
    { image: { x: 10.25, y: 6.93, w: 0.36 * IHU.ratio, h: 0.36, path: IHU.path } },
    { image: { x: 11.55, y: 6.93, w: 0.36 * CIVIL.ratio, h: 0.36, path: CIVIL.path } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: C.accent5, align: "right" },
});

// ---------------- Helpers ----------------
let sectionTitle = "";
function section(title) { pres.addSection({ title }); sectionTitle = title; }
function content(title, lesson, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle });
  s.addText(title, { placeholder: "title" });
  if (lesson) {  // motif: lesson tag, like a cell address
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 11.75, y: 0.5, w: 0.98, h: 0.55, rectRadius: 0.08,
      fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Ετικέτα μαθήματος" });
    s.addText(lesson, { x: 11.75, y: 0.5, w: 0.98, h: 0.55, fontSize: 14, bold: true, color: C.background1,
      align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Αριθμός μαθήματος" });
  }
  if (notes) s.addNotes(notes);
  return s;
}
function divider(title, sub, notes) {
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "body" });
  if (notes) s.addNotes(notes);
  return s;
}
function text(s, t, o) {
  s.addText(t, { isTextBox: true, fontSize: 16, color: C.text1, valign: "top", margin: 0, ...o });
}
function bullets(s, items, o) {
  const runs = items.map((it, i) => {
    const parts = Array.isArray(it) ? it : [{ text: it }];
    return parts.map((p, j) => ({ text: p.text, options: { ...(p.options || {}), ...(j === 0 ? { bullet: true } : {}),
      ...(j === parts.length - 1 && i < items.length - 1 ? { breakLine: true } : {}) } }));
  }).flat();
  s.addText(runs, { isTextBox: true, fontSize: 16, color: C.text1, valign: "top", paraSpaceAfter: 8, margin: 0, ...o });
}
function b(t) { return { text: t, options: { bold: true } }; }
function card(s, x, y, w, h, title, body, kind = "green") {
  const fill = { green: { color: C.background2 }, warn: { color: C.accent2, transparency: 70 },
                 err: { color: C.accent3, transparency: 88 } }[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill, line: { color: C.background1, width: 0 },
    objectName: `Κάρτα: ${title}` });
  const lines = Array.isArray(body)
    ? body.map((t, i) => ({ text: t, options: { fontSize: 15, color: C.text1, bullet: true, ...(i < body.length - 1 ? { breakLine: true } : {}) } }))
    : [{ text: body, options: { fontSize: 15, color: C.text1 } }];
  s.addText([{ text: title, options: { bold: true, fontSize: 17, color: kind === "err" ? C.accent3 : C.text2, breakLine: true } },
             ...lines],
    { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: h - 0.4, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true });
}
function formula(s, f, x, y, w, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.6, rectRadius: 0.08, fill: { color: C.background2 },
    line: { color: C.accent1, width: 1 }, objectName: "Τύπος" });
  s.addText(f, { x, y, w, h: 0.6, fontFace: "Courier New", fontSize: o.fontSize || 18, bold: true, color: C.text2,
    align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function bigStat(s, x, y, w, value, label, color = C.text2, size = 48) {
  s.addText(value, { x, y, w, h: 1.0, fontSize: size, bold: true, color, fontFace: THEME.headFontFace, margin: 0,
    valign: "bottom", isTextBox: true });
  s.addText(label, { x, y: y + 1.05, w, h: 0.6, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
}
// Spreadsheet look-alike: header row of column letters and row numbers.
// cells: rows of strings or {t, k} with k in hdr | f (formula) | in (data) | x (unknown) | err | ok | bad
function sheet(s, x, y, colW, rows, o = {}) {
  const letters = colW.map((_, i) => String.fromCharCode(65 + i));
  const head = [{ text: "", options: { fill: { color: "F3F3F3" } } },
    ...letters.map((l) => ({ text: l, options: { fill: { color: "F3F3F3" }, color: "555555", align: "center" } }))];
  const style = {
    hdr: { bold: true, fill: { color: C.background2 }, color: C.text2 },
    f: { fill: { color: C.background2 } }, in: { fill: { color: "FFF2CC" } },
    x: { fill: { color: C.accent2 }, bold: true }, err: { color: C.accent3, bold: true },
    ok: { fill: { color: "C6EFCE" }, color: "006100", bold: true }, bad: { fill: { color: "F8CBAD" }, color: "9C0006", bold: true },
  };
  const body = rows.map((r, i) => [
    { text: String((o.firstRow || 1) + i), options: { fill: { color: "F3F3F3" }, color: "555555", align: "center" } },
    ...r.map((c) => {
      const cell = typeof c === "object" && c !== null ? c : { t: c };
      const num = /^[−\-]?[\d.,]+%?( \w+)?$/.test(String(cell.t)) && cell.k !== "hdr" && cell.k !== "txt";
      return { text: String(cell.t ?? ""), options: { align: num ? "right" : "left", ...(style[cell.k] || {}) } };
    }),
  ]);
  s.addTable([head, ...body], { x, y, colW: [0.45, ...colW], fontSize: o.fontSize || 14, rowH: o.rowH || 0.36,
    border: { type: "solid", pt: 0.75, color: "D4D4D4" }, valign: "middle", margin: [0, 0.08, 0, 0.08],
    objectName: o.name || "Φύλλο Excel" });
}
async function icon(Comp, color = "#FFFFFF") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: 256 }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  legendFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", catAxisTitleFontFace: "+mn-lt", valAxisTitleFontFace: "+mn-lt",
  catAxisLabelColor: "555555", valAxisLabelColor: "555555", catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  titleFontSize: 15, titleColor: HEX.dk2, legendFontSize: 12, valGridLine: { color: "E3E3E3", size: 0.75 },
  catGridLine: { style: "none" }, catAxisTitleFontSize: 13, valAxisTitleFontSize: 13, catAxisTitleColor: "333333",
  valAxisTitleColor: "333333" };

(async () => {
  const icons = {};
  for (const [k, comp] of Object.entries({ list: fa.FaClipboardList, flask: fa.FaFlask, check: fa.FaCheckCircle,
    chart: fa.FaChartLine, calc: fa.FaCalculator, net: fa.FaProjectDiagram })) icons[k] = await icon(comp);

  // ================= Εισαγωγή =================
  section("Εισαγωγή");
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle });
    s.addText("Excel για Πολιτικούς Μηχανικούς", { placeholder: "title" });
    s.addText([{ text: "Από το πρώτο κελί ως τις μακροεντολές", options: { breakLine: true } },
               { text: SIGN, options: { fontSize: 18, breakLine: true } },
               { text: SITE, options: { fontSize: 16 } }], { placeholder: "body" });
    s.addNotes("Καλωσόρισμα (1 λεπτό). Η διάλεξη είναι μια περιήγηση στα 13 μαθήματα του site. Δεν θα δούμε κάθε κουμπί – " +
      "στόχος είναι να ξέρετε τι μπορεί να κάνει το Excel για έναν μηχανικό και πού θα βρείτε τις λεπτομέρειες. " +
      "Κάθε διαφάνεια έχει πάνω δεξιά τον αριθμό του αντίστοιχου μαθήματος (π.χ. Μ3).");
  }
  {
    const s = content("Πού χρησιμοποιεί ο μηχανικός το Excel", null,
      "2 λεπτά. Ρωτήστε το ακροατήριο αν έχουν ήδη χρησιμοποιήσει Excel. Τονίστε ότι στην πράξη σχεδόν κάθε γραφείο μελετών " +
      "έχει δικά του φύλλα υπολογισμών, και ότι ένα καλό φύλλο είναι ελέγξιμο: δεδομένα σε κελιά, τύποι που διαβάζονται.");
    const items = [["list", "Προμετρήσεις", "όγκοι σκυροδέματος, βάρη οπλισμού, κόστος"],
      ["flask", "Εργαστήριο", "αποτελέσματα κύβων, δοκιμές εφελκυσμού"],
      ["check", "Έλεγχοι", "σ ≤ f_d, ρ_min ≤ ρ ≤ ρ_max, βέλη"],
      ["chart", "Διαγράμματα", "σ–ε, M(x), χρονοδιαγράμματα"],
      ["calc", "Επίλυση εξισώσεων", "Αναζήτηση στόχου, Manning, ύψος x"],
      ["net", "Μητρώα", "K·u = F, δικτυώματα"]];
    items.forEach(([ic, t, d], i) => {
      const x = 0.6 + (i % 3) * 4.1, y = 1.7 + Math.floor(i / 3) * 2.4;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 3.8, h: 2.1, rectRadius: 0.12, fill: { color: C.background2 },
        line: { color: C.background2 }, objectName: `Κάρτα ${t}` });
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.8, h: 0.8, fill: { color: C.accent1 }, line: { color: C.accent1 },
        objectName: `Κύκλος ${t}` });
      s.addImage({ data: icons[ic], x: x + 0.47, y: y + 0.47, w: 0.46, h: 0.46, altText: t });
      text(s, t, { x: x + 1.3, y: y + 0.42, w: 2.35, h: 0.6, fontSize: 19, bold: true, color: C.text2, valign: "middle" });
      text(s, d, { x: x + 0.3, y: y + 1.25, w: 3.3, h: 0.7, fontSize: 15 });
    });
  }
  {
    const s = content("Ο χάρτης των μαθημάτων", null,
      "1 λεπτό. Δείξτε το site σε browser: μενού αριστερά, κάθε μάθημα με διαδραστικά παραδείγματα και αρχείο .xlsx " +
      "με φύλλο Παράδειγμα και φύλλο Άσκηση. Τα πέντε μέρη είναι και η δομή της σημερινής διάλεξης.");
    const groups = [["1", "Βασικά", "1 Περιβάλλον · 2 Τύποι · 3 Αναφορές · 4 Συναρτήσεις · 5 Μορφοποίηση"],
      ["2", "Γραφήματα", "6 Διασπορά XY · 12 Στήλες, πίτα, ιστόγραμμα"],
      ["3", "Λογική και δεδομένα", "7 IF, AND, OR · 8 XLOOKUP · 9 Ταξινόμηση, Πίνακες, SUMIFS"],
      ["4", "Επίλυση προβλημάτων", "10 Αναζήτηση στόχου · 11 Μητρώα"],
      ["5", "Αυτοματοποίηση", "13 Μακροεντολές"]];
    groups.forEach(([n, t, d], i) => {
      const y = 1.5 + i * 0.97;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.75, h: 0.75, fill: { color: C.accent1 }, line: { color: C.accent1 },
        objectName: `Μέρος ${n}` });
      s.addText(n, { x: 0.6, y, w: 0.75, h: 0.75, fontSize: 22, bold: true, color: C.background1, align: "center",
        valign: "middle", margin: 0, isTextBox: true });
      text(s, t, { x: 1.6, y, w: 3.4, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      text(s, d, { x: 5.0, y, w: 7.7, h: 0.75, fontSize: 16, valign: "middle" });
    });
    text(s, `Υλικό, παραδείγματα και ασκήσεις: ${SITE}`, { x: 0.6, y: 6.45, w: 12, h: 0.4, fontSize: 14, color: C.accent1, bold: true });
  }

  // ================= Μέρος 1 =================
  section("Βασικά");
  divider("Μέρος 1 · Τα βασικά", "Κελιά, τύποι, αναφορές, συναρτήσεις και μορφοποίηση – Μαθήματα 1 έως 5",
    "Περίπου 15 λεπτά για αυτό το μέρος· είναι η βάση για όλα τα υπόλοιπα.");
  {
    const s = content("Κελιά και ο πρώτος τύπος", "Μ1",
      "3 λεπτά. ΕΠΙΔΕΙΞΗ στο Excel: φτιάξτε τον πίνακα των δοκών, γράψτε =B2*C2, αλλάξτε το h και δείξτε ότι το εμβαδόν " +
      "αλλάζει. Δείξτε Πλαίσιο ονόματος και Γραμμή τύπων. Τονίστε: γράφουμε διευθύνσεις, όχι αριθμούς, μέσα στον τύπο.");
    sheet(s, 0.6, 1.6, [1.0, 1.0, 1.0, 1.0, 1.15, 1.15], [
      [{ t: "Δοκός", k: "hdr" }, { t: "b (m)", k: "hdr" }, { t: "h (m)", k: "hdr" }, { t: "L (m)", k: "hdr" }, { t: "A (m²)", k: "hdr" }, { t: "V (m³)", k: "hdr" }],
      ["Δ1", "0,25", "0,50", "5,00", { t: "0,1250", k: "f" }, { t: "0,6250", k: "f" }],
      ["Δ2", "0,30", "0,60", "6,00", { t: "0,1800", k: "f" }, { t: "1,0800", k: "f" }],
      ["Δ3", "0,25", "0,45", "4,50", { t: "0,1125", k: "f" }, { t: "0,5063", k: "f" }]]);
    formula(s, "E2:  =B2*C2", 0.6, 3.55, 3.3);
    formula(s, "F2:  =E2*D2", 4.1, 3.55, 3.3);
    bullets(s, [[b("Διεύθυνση κελιού: "), { text: "στήλη + γραμμή, π.χ. C3" }],
      [b("Τύπος: "), { text: "ξεκινά πάντα με =" }],
      [b("Αναφορές, όχι αριθμοί: "), { text: "αλλάζει το h → αλλάζουν A και V" }],
      [b("Μονάδες "), { text: "στην επικεφαλίδα· στο κελί μόνο ο αριθμός" }]], { x: 0.6, y: 4.5, w: 7.0, h: 2.2 });
    card(s, 8.3, 1.6, 4.4, 2.3, "Προσοχή στο πληκτρολόγιο",
      "Το ελληνικό «Β» μοιάζει με το λατινικό «B». Ο τύπος =Β2*C2 με ελληνικό Β βγάζει σφάλμα. Τύποι πάντα με λατινικά (Alt+Shift).", "warn");
    card(s, 8.3, 4.15, 4.4, 2.3, "Κείμενο ή αριθμός;",
      "Οι αριθμοί στοιχίζονται δεξιά, το κείμενο αριστερά. Ένας «αριθμός» αριστερά δεν συμμετέχει σε πράξεις.", "green");
  }
  {
    const s = content("Τελεία ή κόμμα; Οι τοπικές ρυθμίσεις", "Οδηγός",
      "2 λεπτά. Πολύ συχνή πηγή λαθών. Τα αρχεία .xlsx δουλεύουν παντού· το πρόβλημα είναι η πληκτρολόγηση και η επικόλληση. " +
      "ΕΠΙΔΕΙΞΗ: πληκτρολογήστε 1.250 σε Excel με ελληνικές ρυθμίσεις και δείξτε ότι γίνεται 1250. Γρήγορος έλεγχος ρυθμίσεων: =1/4.");
    const rows = [["", "Ελληνικές ρυθμίσεις", "Αγγλικές ρυθμίσεις"], ["Υποδιαστολή", "0,25", "0.25"],
      ["Χιλιάδες", "1.234.567,89", "1,234,567.89"], ["Ορίσματα", "=ROUND(A1;2)", "=ROUND(A1,2)"], ["Σταθερά σε τύπο", "=A1*0,5", "=A1*0.5"]];
    s.addTable(rows.map((r, i) => r.map((c, j) => ({ text: c, options: {
      bold: i === 0 || j === 0, color: i === 0 ? HEX.lt1 : HEX.dk1, fill: { color: i === 0 ? HEX.accent1 : (i % 2 ? "FFFFFF" : HEX.lt2) },
      fontFace: i > 0 && j > 0 ? "Courier New" : undefined } }))),
      { x: 0.6, y: 1.6, w: 7.4, colW: [2.2, 2.6, 2.6], fontSize: 16, rowH: 0.55, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Σύγκριση ρυθμίσεων" });
    bigStat(s, 8.6, 1.5, 4.2, "1.250 → 1250", "με ελληνικές ρυθμίσεις, η τελεία διαβάζεται ως διαχωριστικό χιλιάδων – χωρίς κανένα μήνυμα σφάλματος", C.accent3);
    bullets(s, ["Το .xlsx δεν «κουβαλάει» ρυθμίσεις: ανοίγει σωστά παντού",
      "Λάθη γίνονται στην πληκτρολόγηση, στην επικόλληση από PDF και στα CSV",
      "Τύποι από αγγλόφωνες πηγές: , → ; και . → ,"], { x: 0.6, y: 4.65, w: 7.4, h: 1.9 });
    card(s, 8.6, 4.4, 4.1, 2.1, "Σύσταση",
      "Ελληνικές ρυθμίσεις (κόμμα και ;). Όποια κι αν είναι η επιλογή σας, να ξέρετε ποια έχετε.", "green");
  }
  {
    const s = content("Προτεραιότητα πράξεων: δύο παγίδες", "Μ2",
      "2 λεπτά. Ρωτήστε το ακροατήριο πριν αποκαλύψετε τα αποτελέσματα. Το παράδειγμα σ = M/W είναι το σημαντικότερο: " +
      "ένας παρονομαστής με πολλούς όρους θέλει πάντα παρένθεση.");
    sheet(s, 0.6, 1.6, [2.0, 1.4, 3.6], [
      [{ t: "Τύπος", k: "hdr" }, { t: "Αποτέλεσμα", k: "hdr" }, { t: "Γιατί", k: "hdr" }],
      ["=2+3*4", { t: "14", k: "f" }, "πρώτα ο πολλαπλασιασμός"],
      ["=10/2*5", { t: "25", k: "f" }, "ίδια προτεραιότητα: αριστερά → δεξιά"],
      ["=-2^2", { t: "4", k: "bad" }, "το πρόσημο πριν τη δύναμη: (−2)²"],
      ["=2^3^2", { t: "64", k: "bad" }, "(2³)² = 64, όχι 2⁹ = 512"]]);
    text(s, "Τάση σε διατομή:  σ = M / W,  W = b·h²/6", { x: 0.6, y: 4.0, w: 7.5, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
    sheet(s, 0.6, 4.5, [3.0, 4.0], [
      [{ t: "=B1/(B2*B3^2/6)", k: "ok" }, "M / (b·h²/6)  ✔"],
      [{ t: "=B1/B2*B3^2/6", k: "bad" }, "(M/b)·h²/6  ✘ λάθος αποτέλεσμα"]], { name: "Σωστός και λάθος τύπος" });
    card(s, 8.6, 1.6, 4.1, 2.4, "Κανόνας",
      "Όταν υπάρχει η παραμικρή αμφιβολία, βάλτε παρενθέσεις. Ο αριθμητής και ο παρονομαστής ενός κλάσματος μπαίνουν πάντα σε παρένθεση.", "green");
    card(s, 8.6, 4.25, 4.1, 2.2, "Σειρά πράξεων",
      "( )  →  πρόσημο  →  ^  →  * /  →  + −", "warn");
  }
  {
    const s = content("Σχετικές και απόλυτες αναφορές", "Μ3",
      "3 λεπτά. Το πιο σημαντικό εννοιολογικά βήμα. ΕΠΙΔΕΙΞΗ: γράψτε =B4/B1^2/1000, σύρετε τη λαβή συμπλήρωσης και δείξτε " +
      "#DIV/0!, #VALUE! και τις «σιωπηλές» λάθος τιμές. Μετά F4 → $B$1 και ξανά αντιγραφή. Οι λάθος τιμές χωρίς μήνυμα είναι το πιο " +
      "επικίνδυνο είδος λάθους σε φύλλο υπολογισμών.");
    sheet(s, 0.6, 1.6, [1.2, 1.0, 1.25, 1.25], [
      [{ t: "a (m)", k: "hdr" }, { t: "0,15", k: "in" }, "", ""],
      ["", "", "", ""],
      [{ t: "Δοκίμιο", k: "hdr" }, { t: "F (kN)", k: "hdr" }, { t: "ΛΑΘΟΣ", k: "hdr" }, { t: "ΣΩΣΤΟ", k: "hdr" }],
      ["Κ1", "675", { t: "30,00", k: "f" }, { t: "30,00", k: "f" }],
      ["Κ2", "712", { t: "#DIV/0!", k: "err" }, { t: "31,64", k: "f" }],
      ["Κ3", "690", { t: "#VALUE!", k: "err" }, { t: "30,67", k: "f" }],
      ["Κ4", "655", { t: "0,00", k: "bad" }, { t: "29,11", k: "f" }]]);
    formula(s, "ΛΑΘΟΣ  C4: =B4/B1^2/1000", 0.6, 4.6, 5.15, { fontSize: 16 });
    formula(s, "ΣΩΣΤΟ  D4: =B4/$B$1^2/1000", 0.6, 5.35, 5.15, { fontSize: 16 });
    text(s, "Το B1 «γλιστράει» σε B2, B3, B4… κατά την αντιγραφή. Το $ κλειδώνει στήλη ή γραμμή.",
      { x: 0.6, y: 6.15, w: 5.6, h: 0.6, fontSize: 14, color: C.accent5 });
    s.addTable([["Πατήματα F4", "Αναφορά", "Σταθερό"], ["–", "B1", "τίποτα"], ["1", "$B$1", "στήλη και γραμμή"],
      ["2", "B$1", "μόνο η γραμμή"], ["3", "$B1", "μόνο η στήλη"]].map((r, i) => r.map((c, j) => ({ text: c, options: {
        bold: i === 0, color: i === 0 ? HEX.lt1 : HEX.dk1, fill: { color: i === 0 ? HEX.accent1 : "FFFFFF" },
        fontFace: j === 1 && i > 0 ? "Courier New" : undefined } }))),
      { x: 6.9, y: 1.6, w: 5.8, colW: [1.6, 1.4, 2.8], fontSize: 15, rowH: 0.45, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Κύκλος F4" });
    card(s, 6.9, 4.1, 5.8, 2.35, "Μικτές αναφορές: πίνακας M = q·L²/8",
      "Ένας τύπος στο B2, αντιγραμμένος δεξιά και κάτω: =$A2*B$1^2/8. Το q είναι πάντα στη στήλη A ($A), το L πάντα στη γραμμή 1 (B$1).", "green");
  }
  {
    const s = content("Βασικές συναρτήσεις", "Μ4",
      "3 λεπτά. Ορίσματα με ; στις ελληνικές ρυθμίσεις. ROUNDUP για συνδετήρες: δεν στρογγυλεύουμε προς τα κάτω πλήθος ράβδων. " +
      "Το COS(30) δίνει συνημίτονο 30 ακτινίων – άλλο ένα λάθος χωρίς μήνυμα. ΕΠΙΔΕΙΞΗ: Alt+= για αυτόματη άθροιση και η Γραμμή κατάστασης.");
    bigStat(s, 0.6, 1.4, 3.8, "30,98 MPa", "=AVERAGE(B2:B7) – μέση τιμή 6 κύβων");
    bigStat(s, 4.6, 1.4, 3.8, "35", "=ROUNDUP(5/0,15;0)+1 – συνδετήρες σε 5 m ανά 15 cm");
    bigStat(s, 8.6, 1.4, 4.1, "8,04 cm²", "=4*PI()*16^2/4/100 – εμβαδόν 4Φ16");
    s.addTable([["Θέλω…", "Συνάρτηση"], ["άθροισμα / μέση τιμή", "SUM / AVERAGE"], ["ελάχιστο / μέγιστο / πλήθος", "MIN / MAX / COUNT"],
      ["στρογγύλευση", "ROUND / ROUNDUP / ROUNDDOWN"], ["π, ρίζα, απόλυτη τιμή", "PI() / SQRT / ABS"]].map((r, i) => r.map((c, j) => ({
        text: c, options: { bold: i === 0, color: i === 0 ? HEX.lt1 : HEX.dk1, fill: { color: i === 0 ? HEX.accent1 : "FFFFFF" },
          fontFace: j === 1 && i > 0 ? "Courier New" : undefined } }))),
      { x: 0.6, y: 3.6, w: 7.4, colW: [3.2, 4.2], fontSize: 15, rowH: 0.48, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Πίνακας συναρτήσεων" });
    card(s, 8.6, 3.6, 4.1, 2.8, "Τριγωνομετρία σε ακτίνια",
      "=50*COS(RADIANS(30)) → 43,30 kN ✔\n=50*COS(30) → 7,71 kN ✘\nΚαμία προειδοποίηση – απλώς λάθος αριθμός.", "err");
  }
  {
    const s = content("Μορφοποίηση: τιμή και εμφάνιση", "Μ5",
      "2 λεπτά. Ctrl+1 → Προσαρμογή. Η μονάδα μέσα στη μορφή κρατά το κελί αριθμό. Το Excel δεν ξέρει μονάδες: οι μετατροπές είναι " +
      "ευθύνη του μηχανικού και γίνονται ρητά. Μορφοποίηση υπό όρους: αυτόματη επισήμανση τιμών.");
    sheet(s, 0.6, 1.6, [2.6, 1.6, 2.0], [
      [{ t: "Μορφή", k: "hdr" }, { t: "Εμφάνιση", k: "hdr" }, { t: "Κωδικός", k: "hdr" }],
      ["Γενική", "0,1125", "Γενική"], ["2 δεκαδικά", "0,11", { t: "0,00", k: "txt" }], ["Ποσοστό", "11,3%", { t: "0,0%", k: "txt" }],
      ["Επιστημονική", "1,13E-01", { t: "0,00E+00", k: "txt" }], ["Με μονάδα", { t: "0,1125 m²", k: "f" }, { t: "0,0000 \"m²\"", k: "txt" }]],
      { rowH: 0.34, fontSize: 13 });
    text(s, "Η Γραμμή τύπων δείχνει πάντα 0,1125: η μορφή αλλάζει την εμφάνιση, όχι την τιμή.",
      { x: 0.6, y: 4.45, w: 7.0, h: 0.4, fontSize: 14, color: C.accent5 });
    s.addTable([["Από", "Σε", "Πράξη"], ["kg/m³", "kN/m³", "×9,81/1000"], ["kN/m² (kPa)", "MPa = N/mm²", "/1000"], ["cm²", "m²", "/10000"]]
      .map((r, i) => r.map((c) => ({ text: c, options: { bold: i === 0, color: i === 0 ? HEX.lt1 : HEX.dk1,
        fill: { color: i === 0 ? HEX.accent1 : "FFFFFF" } } }))),
      { x: 0.6, y: 5.0, w: 7.0, colW: [2.2, 2.4, 2.4], fontSize: 14, rowH: 0.36, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Μετατροπές μονάδων" });
    card(s, 8.3, 1.6, 4.4, 2.3, "Μονάδες χωρίς να χαλάσει ο αριθμός",
      "Ctrl+1 → Αριθμός → Προσαρμογή: 0,00 \"kN\". Το κελί δείχνει 12,50 kN αλλά παραμένει αριθμός για τους τύπους.", "green");
    card(s, 8.3, 4.15, 4.4, 2.3, "Το Excel δεν ξέρει μονάδες",
      "Ένα σύστημα μονάδων για όλο το φύλλο (kN, m) και κάθε μετατροπή ρητά, σε δικό της κελί ή με εμφανή συντελεστή.", "warn");
  }

  // ================= Μέρος 2 =================
  section("Γραφήματα");
  divider("Μέρος 2 · Γραφήματα", "Το σωστό γράφημα για κάθε είδος δεδομένων – Μαθήματα 6 και 12",
    "Περίπου 11 λεπτά. Το μήνυμα: η επιλογή του είδους γραφήματος δεν είναι αισθητική, αλλά θέμα ορθότητας.");
  const eps = [0, 0.5, 1, 1.5, 2, 2.5, 3, 5, 10, 20, 40, 60, 80, 100];
  const sig = [0, 98, 203, 297, 402, 500, 503, 508, 515, 540, 590, 610, 618, 612];
  {
    const s = content("Διασπορά XY, όχι «Γραμμή»", "Μ6",
      "3 λεπτά. Τα ίδια δεδομένα, δύο γραφήματα. Το «Γραμμή» βάζει το 0,5‰ και το 20‰ σε ίσες αποστάσεις και εξαφανίζει τον " +
      "ελαστικό κλάδο. Κανόνας: αν ο οριζόντιος άξονας έχει αριθμούς με μονάδες, θέλετε Διασπορά (X, Y).");
    s.addChart(pres.charts.LINE, [{ name: "σ", labels: eps.map((e) => String(e).replace(".", ",")), values: sig }], {
      x: 0.6, y: 1.5, w: 5.9, h: 4.3, ...chartText, showTitle: true, title: "✘ Γράφημα «Γραμμή»", chartColors: [HEX.accent3],
      lineSize: 2, lineDataSymbolSize: 6, showLegend: false, showCatAxisTitle: true, catAxisTitle: "ε (‰) ως ετικέτες",
      showValAxisTitle: true, valAxisTitle: "σ (MPa)", objectName: "Γράφημα γραμμής σ–ε" });
    s.addChart(pres.charts.SCATTER, [{ name: "ε", values: eps }, { name: "σ–ε", values: sig }], {
      x: 6.8, y: 1.5, w: 5.9, h: 4.3, ...chartText, showTitle: true, title: "✔ Διασπορά (X, Y)", chartColors: [HEX.accent1],
      lineSize: 2, lineDataSymbolSize: 6, showLegend: false, showCatAxisTitle: true, catAxisTitle: "ε (‰)",
      showValAxisTitle: true, valAxisTitle: "σ (MPa)", objectName: "Διάγραμμα διασποράς σ–ε" });
    text(s, "Ίδια δεδομένα δοκιμής εφελκυσμού χάλυβα. Μόνο η διασπορά δείχνει τον απότομο ελαστικό κλάδο, τη διαρροή και την κράτυνση.",
      { x: 0.6, y: 6.05, w: 12.1, h: 0.6, fontSize: 15 });
  }
  {
    const s = content("Διάγραμμα από τύπους και γραμμή τάσης", "Μ6",
      "2 λεπτά. Το διάγραμμα M(x) είναι συνδεδεμένο με τα κελιά q και L: ΕΠΙΔΕΙΞΗ αλλαγής q. Γραμμή τάσης στον ελαστικό κλάδο: " +
      "η κλίση είναι το E. Προσοχή στις μονάδες: με ε σε ‰ η κλίση πολλαπλασιάζεται επί 1000.");
    const xs = [0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6];
    s.addChart(pres.charts.SCATTER, [{ name: "x", values: xs }, { name: "M(x)", values: xs.map((x) => 20 * x * (6 - x) / 2) }], {
      x: 0.6, y: 1.5, w: 7.2, h: 4.8, ...chartText, showTitle: true, title: "M(x) = q·x·(L − x)/2,  q = 20 kN/m,  L = 6 m",
      chartColors: [HEX.accent1], lineSize: 3, lineDataSymbol: "none", lineSmooth: true, showLegend: false,
      showCatAxisTitle: true, catAxisTitle: "x (m)", showValAxisTitle: true, valAxisTitle: "M (kNm)", objectName: "Διάγραμμα ροπών" });
    bigStat(s, 8.3, 1.5, 4.4, "90 kNm", "M_max = qL²/8 στο μέσο – το διάγραμμα ξανασχεδιάζεται μόλις αλλάξει το q ή το L");
    bigStat(s, 8.3, 3.4, 4.4, "≈ 200 GPa", "μέτρο ελαστικότητας από τη γραμμή τάσης του ελαστικού κλάδου, ή =SLOPE(σ;ε)");
    card(s, 8.3, 5.3, 4.4, 1.25, "Πάντα τίτλοι αξόνων με μονάδες", "κουμπί + → Τίτλοι αξόνων", "green");
  }
  {
    const s = content("Στήλες και στοιβαγμένες στήλες", "Μ12",
      "2 λεπτά. Όταν ο οριζόντιος άξονας είναι κατηγορίες. Τα δεδομένα είναι από την προμέτρηση του Μαθήματος 9. " +
      "Στοιβαγμένες στήλες: το ύψος είναι το σύνολο του ορόφου, τα χρώματα η σύνθεσή του.");
    const floors = ["Ισόγειο", "1ος", "2ος"];
    s.addChart(pres.charts.BAR, [{ name: "Υποστυλώματα", labels: floors, values: [2.835, 1.92, 1.92] },
      { name: "Δοκοί", labels: floors, values: [3.41, 3.41, 3.41] }, { name: "Πλάκες", labels: floors, values: [12.6, 12.6, 12.6] }], {
      x: 0.6, y: 1.5, w: 7.2, h: 4.9, ...chartText, barDir: "col", barGrouping: "stacked", showTitle: true,
      title: "Σκυρόδεμα ανά όροφο (m³)", chartColors: [HEX.accent1, HEX.accent2, HEX.accent4], showLegend: true, legendPos: "b",
      showValAxisTitle: true, valAxisTitle: "V (m³)", barGapWidthPct: 60, objectName: "Στοιβαγμένες στήλες" });
    s.addTable([["Θέλω να δείξω…", "Γράφημα"], ["σχέση δύο αριθμών", "Διασπορά (X, Y)"], ["σύγκριση κατηγοριών", "Στήλες / Ράβδοι"],
      ["σύνολο και σύνθεση", "Στοιβαγμένες στήλες"], ["μερίδια ενός συνόλου", "Πίτα (λίγες κατηγορίες)"],
      ["εξέλιξη σε ίσα διαστήματα", "Γραμμή"], ["κατανομή μετρήσεων", "Ιστόγραμμα"]].map((r, i) => r.map((c) => ({ text: c,
        options: { bold: i === 0, color: i === 0 ? HEX.lt1 : HEX.dk1, fill: { color: i === 0 ? HEX.accent1 : (i % 2 ? "FFFFFF" : HEX.lt2) } } }))),
      { x: 8.1, y: 1.5, w: 4.6, colW: [2.4, 2.2], fontSize: 14, rowH: 0.55, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Επιλογή γραφήματος" });
  }
  {
    const s = content("Πίτα ή ράβδοι;", "Μ12",
      "2 λεπτά. Η πίτα δείχνει ότι ο φέρων οργανισμός είναι περίπου το 1/3. Ότι τοιχοποιίες και δάπεδα είναι ίσα φαίνεται μόνο στις " +
      "ράβδους: το μάτι συγκρίνει μήκη καλύτερα από γωνίες. Ποσά ενδεικτικά. Ποτέ τρισδιάστατη πίτα.");
    const cats = ["Φέρων οργανισμός", "Η/Μ εγκαταστάσεις", "Τοιχοποιίες – επιχρίσματα", "Δάπεδα – χρωματισμοί", "Κουφώματα", "Εκσκαφές – θεμελίωση"];
    const vals = [150, 110, 75, 75, 50, 40];
    s.addChart(pres.charts.PIE, [{ name: "Κόστος", labels: cats, values: vals }], {
      x: 0.6, y: 1.5, w: 5.6, h: 4.9, ...chartText, showTitle: true, title: "Κατανομή κόστους κατασκευής",
      chartColors: [HEX.accent1, HEX.accent2, HEX.accent4, HEX.accent6, HEX.accent5, "70AD47"], showPercent: true,
      showLegend: true, legendPos: "b", dataLabelColor: "FFFFFF", dataLabelFontSize: 12, objectName: "Πίτα κόστους" });
    s.addChart(pres.charts.BAR, [{ name: "Κόστος (χιλ. €)", labels: cats.slice().reverse(), values: vals.slice().reverse() }], {
      x: 6.5, y: 1.5, w: 6.2, h: 4.9, ...chartText, barDir: "bar", showTitle: true, title: "Το ίδιο, ταξινομημένο (χιλ. €)",
      chartColors: [HEX.accent1], showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12,
      dataLabelColor: "333333", barGapWidthPct: 40, objectName: "Ταξινομημένες ράβδοι κόστους" });
  }
  {
    const s = content("Καμπύλη S και ιστόγραμμα", "Μ12",
      "2 λεπτά. Αριστερά: γράφημα συνδυασμού με δευτερεύοντα άξονα – εδώ η «Γραμμή» είναι σωστή, γιατί οι μήνες είναι ισαπέχουσες " +
      "κατηγορίες. Δεξιά: ιστόγραμμα 40 κύβων με COUNTIFS. Η διασπορά συνδέεται με το f_ck ως ποσοστημόριο 5% (Δομικά Υλικά).");
    const months = ["Ι", "Φ", "Μ", "Α", "Μ", "Ι", "Ι", "Α", "Σ", "Ο", "Ν", "Δ"];
    const monthly = [15, 25, 35, 50, 60, 65, 60, 55, 45, 35, 30, 25];
    const cum = monthly.reduce((a, v) => [...a, (a.at(-1) || 0) + v], []);
    s.addChart([
      { type: pres.charts.BAR, data: [{ name: "Δαπάνη μήνα", labels: months, values: monthly }],
        options: { chartColors: [HEX.accent1], barGapWidthPct: 50 } },
      { type: pres.charts.LINE, data: [{ name: "Αθροιστικά", labels: months, values: cum }],
        options: { chartColors: [HEX.accent3], lineSize: 3, lineDataSymbol: "none", secondaryValAxis: true, secondaryCatAxis: true } },
    ], { x: 0.6, y: 1.5, w: 6.2, h: 4.9, ...chartText, showTitle: true, title: "Δαπάνες έργου (χιλ. €)", showLegend: true, legendPos: "b",
      valAxes: [{ showValAxisTitle: true, valAxisTitle: "Μηνιαία" }, { showValAxisTitle: true, valAxisTitle: "Αθροιστικά", valGridLine: { style: "none" } }],
      catAxes: [{ catAxisTitle: "Μήνας" }, { catAxisHidden: true }], objectName: "Καμπύλη S" });
    s.addChart(pres.charts.BAR, [{ name: "Πλήθος", labels: ["26–28", "28–30", "30–32", "32–34", "34–36", "36–38", "38–40"],
      values: [3, 4, 8, 13, 5, 2, 5] }], {
      x: 7.0, y: 1.5, w: 5.7, h: 4.9, ...chartText, barDir: "col", showTitle: true, title: "Ιστόγραμμα 40 κύβων",
      chartColors: [HEX.accent1], showLegend: false, barGapWidthPct: 5, showValue: true, dataLabelPosition: "outEnd",
      dataLabelFontSize: 12, dataLabelColor: "333333", showCatAxisTitle: true, catAxisTitle: "f (MPa)",
      showValAxisTitle: true, valAxisTitle: "Πλήθος", objectName: "Ιστόγραμμα κύβων" });
  }

  // ================= Μέρος 3 =================
  section("Λογική και δεδομένα");
  divider("Μέρος 3 · Λογική και δεδομένα", "Έλεγχοι επάρκειας, πίνακες υλικών και προμετρήσεις – Μαθήματα 7 έως 9",
    "Περίπου 9 λεπτά. Εδώ το Excel αρχίζει να μοιάζει με «εργαλείο μελέτης».");
  {
    const s = content("Έλεγχοι επάρκειας με IF, AND, OR", "Μ7",
      "3 λεπτά. Ο βαθμός εκμετάλλευσης η = σ/f_d. Εμφωλευμένο IF: η σειρά των ελέγχων μετράει. AND για τιμή μέσα σε όρια – " +
      "το Excel δεν καταλαβαίνει διπλή ανισότητα. Τα κείμενα σε ίσια εισαγωγικά, και OK με λατινικά – όχι ΟΚ με ελληνικά.");
    sheet(s, 0.6, 1.6, [0.75, 0.9, 1.0, 0.85, 1.5, 1.4], [
      [{ t: "Υποστ.", k: "hdr" }, { t: "N (kN)", k: "hdr" }, { t: "σ (MPa)", k: "hdr" }, { t: "η", k: "hdr" }, { t: "Έλεγχος", k: "hdr" }, { t: "Κατάταξη", k: "hdr" }],
      ["Κ1", "1800", "11,25", "0,792", { t: "OK", k: "ok" }, { t: "OK", k: "ok" }],
      ["Κ2", "2400", "15,00", "1,056", { t: "ΑΝΕΠΑΡΚΕΙΑ", k: "bad" }, { t: "ΑΝΕΠΑΡΚΕΙΑ", k: "bad" }],
      ["Κ3", "2100", "13,33", "0,939", { t: "OK", k: "ok" }, { t: "ΟΡΙΑΚΑ", k: "in" }],
      ["Κ5", "1500", "16,67", "1,174", { t: "ΑΝΕΠΑΡΚΕΙΑ", k: "bad" }, { t: "ΑΝΕΠΑΡΚΕΙΑ", k: "bad" }]], { fontSize: 13 });
    formula(s, '=IF(C2<=$f_d;"OK";"ΑΝΕΠΑΡΚΕΙΑ")', 0.6, 3.75, 6.85, { fontSize: 15 });
    formula(s, '=IF(η<=0,9;"OK";IF(η<=1;"ΟΡΙΑΚΑ";"ΑΝΕΠΑΡΚΕΙΑ"))', 0.6, 4.5, 6.85, { fontSize: 14 });
    formula(s, '=IF(AND(ρ>=ρmin;ρ<=ρmax);"OK";"ΑΝΕΠΑΡΚΕΙΑ")', 0.6, 5.25, 6.85, { fontSize: 14 });
    card(s, 7.9, 1.6, 4.8, 2.3, "AND ή OR – το ίδιο με δύο τρόπους",
      "Δεκτή αν AND(η ≤ 1; δ ≤ L/250).\nΑπορρίπτεται αν OR(η > 1; δ > L/250).", "green");
    card(s, 7.9, 4.15, 4.8, 2.3, "Χρώμα στα αποτελέσματα",
      "Μορφοποίηση υπό όρους → Ίσο με… → ΑΝΕΠΑΡΚΕΙΑ = κόκκινο. Δεκάδες στοιχεία ελέγχονται με μια ματιά.", "warn");
  }
  {
    const s = content("Πίνακες υλικών με XLOOKUP", "Μ8",
      "3 λεπτά. Ο πίνακας μπαίνει μία φορά, σε δικό του φύλλο. Αναπτυσσόμενη λίστα (Επικύρωση δεδομένων) για να μη γίνονται " +
      "τυπογραφικά λάθη. 5ο όρισμα 1: η αμέσως μεγαλύτερη τιμή → επιλογή διαμέτρου. Σε παλαιό Excel: VLOOKUP με 4ο όρισμα 0.");
    sheet(s, 0.6, 1.6, [1.15, 1.0, 1.0, 1.05, 1.05], [
      [{ t: "Κατηγορία", k: "hdr" }, { t: "f_ck", k: "hdr" }, { t: "f_cm", k: "hdr" }, { t: "f_ctm", k: "hdr" }, { t: "E_cm", k: "hdr" }],
      ["C20/25", "20", "28", "2,2", "30"], [{ t: "C25/30", k: "x" }, "25", "33", { t: "2,6", k: "x" }, "31"],
      ["C30/37", "30", "38", "2,9", "33"], ["C35/45", "35", "43", "3,2", "34"]], { fontSize: 14 });
    text(s, "Φύλλο «Υλικά»: EN 1992-1-1, Πίνακας 3.1 (απόσπασμα)", { x: 0.6, y: 3.85, w: 6, h: 0.4, fontSize: 13, color: C.accent5 });
    formula(s, "=XLOOKUP(B5;Κατηγορία;f_ctm)   → 2,6", 0.6, 4.4, 6.25, { fontSize: 16 });
    formula(s, '=XLOOKUP(187,5;As;Φ;"καμία";1)   → Φ16', 0.6, 5.15, 6.25, { fontSize: 16 });
    text(s, "Τι ψάχνω · πού το ψάχνω · τι επιστρέφω · [αν δεν βρεθεί] · [ακριβής ή αμέσως μεγαλύτερη]",
      { x: 0.6, y: 5.95, w: 6.4, h: 0.6, fontSize: 14, color: C.accent5 });
    card(s, 7.5, 1.6, 5.2, 2.3, "Επιλογή διαμέτρου",
      "As,req = 7,5 cm² με 4 ράβδους → 187,5 mm² ανά ράβδο. Δεν υπάρχει ακριβώς· η αμέσως μεγαλύτερη είναι 201 mm² → 4Φ16 = 8,04 cm².", "green");
    card(s, 7.5, 4.15, 5.2, 2.3, "#N/A ενώ η τιμή «φαίνεται»;",
      "Κενό στο τέλος («C25/30 »), ελληνικά γράμματα σε κωδικούς, αριθμός ως κείμενο. Λύση: αναπτυσσόμενη λίστα.", "warn");
  }
  {
    const s = content("Προμέτρηση: Πίνακες, φίλτρα και SUMIFS", "Μ9",
      "3 λεπτά. Ctrl+T για Πίνακα: φίλτρα, υπολογιζόμενες στήλες, γραμμή συνόλων. Με φίλτρο «Πλάκα» το SUM δεν αλλάζει, το SUBTOTAL ναι. " +
      "Ο πίνακας SUMIFS όροφος × στοιχείο είναι ένας τύπος με μικτές αναφορές. Όροφος ως αριθμός (0, 1, 2), όχι «Ισόγειο».");
    bigStat(s, 0.6, 1.5, 3.9, "54,705 m³", "SUM: όλα τα 30 στοιχεία, με ή χωρίς φίλτρο");
    bigStat(s, 4.7, 1.5, 3.9, "37,800 m³", "SUBTOTAL(109;…) με φίλτρο «Πλάκα»: μόνο οι ορατές γραμμές", C.accent1);
    sheet(s, 0.6, 3.6, [1.8, 1.55, 1.25, 1.25, 1.2], [
      [{ t: "Όροφος", k: "hdr" }, { t: "Υποστυλώματα", k: "hdr" }, { t: "Δοκοί", k: "hdr" }, { t: "Πλάκες", k: "hdr" }, { t: "Σύνολο", k: "hdr" }],
      ["0", { t: "2,835", k: "f" }, { t: "3,410", k: "f" }, { t: "12,600", k: "f" }, { t: "18,845", k: "f" }],
      ["1", { t: "1,920", k: "f" }, { t: "3,410", k: "f" }, { t: "12,600", k: "f" }, { t: "17,930", k: "f" }],
      ["2", { t: "1,920", k: "f" }, { t: "3,410", k: "f" }, { t: "12,600", k: "f" }, { t: "17,930", k: "f" }]], { fontSize: 14 });
    formula(s, "=SUMIFS(V;Όροφος;$A2;Στοιχείο;B$1)", 0.6, 5.65, 7.5, { fontSize: 16 });
    card(s, 8.9, 1.6, 3.8, 4.85, "Κανόνες λίστας", ["μία γραμμή επικεφαλίδων", "χωρίς κενές γραμμές",
      "χωρίς συγχωνευμένα κελιά", "ίδιο είδος δεδομένου σε κάθε στήλη", "ποτέ ταξινόμηση μίας στήλης μόνης της"], "green");
  }

  // ================= Μέρος 4 =================
  section("Επίλυση προβλημάτων");
  divider("Μέρος 4 · Επίλυση προβλημάτων", "Αναζήτηση στόχου και μητρώα – Μαθήματα 10 και 11",
    "Περίπου 11 λεπτά. Το πιο «μηχανικό» κομμάτι: εξισώσεις χωρίς κλειστή λύση και γραμμικά συστήματα.");
  {
    const s = content("Αναζήτηση στόχου", "Μ10",
      "3 λεπτά. Ξεκινήστε με το ύψος δοκού (υπάρχει αναλυτική λύση – έτσι εμπιστευόμαστε το εργαλείο), μετά Manning: το y είναι μέσα " +
      "στο A, P και R^(2/3). ΕΠΙΔΕΙΞΗ: Δεδομένα → Ανάλυση πιθανοτήτων → Αναζήτηση στόχου. Η λύση δεν ενημερώνεται αυτόματα.");
    const ys = [0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6, 1.8, 2.0];
    const Q = (y) => 3 * y * Math.pow(3 * y / (3 + 2 * y), 2 / 3) * Math.sqrt(0.001) / 0.015;
    s.addChart(pres.charts.SCATTER, [{ name: "y", values: ys }, { name: "Q(y) Manning", values: ys.map((y) => +Q(y).toFixed(3)) },
      { name: "Στόχος Q = 5", values: ys.map(() => 5) }], {
      x: 0.6, y: 1.5, w: 6.6, h: 4.9, ...chartText, showTitle: true, title: "Κανάλι b = 3 m, n = 0,015, S = 0,001",
      chartColors: [HEX.accent1, HEX.accent3], lineSize: 3, lineDataSymbol: "none", lineSmooth: true, showLegend: true,
      legendPos: "b", showCatAxisTitle: true, catAxisTitle: "Βάθος ροής y (m)", showValAxisTitle: true, valAxisTitle: "Q (m³/s)",
      objectName: "Καμπύλη Manning" });
    const fields = [["Ορισμός κελιού", "ο τύπος του αποτελέσματος, π.χ. Q(y)"], ["Στην τιμή", "ο στόχος, π.χ. 5"],
      ["Με αλλαγή του κελιού", "ο άγνωστος, π.χ. y"]];
    fields.forEach(([t, d], i) => {
      const y = 1.6 + i * 0.95;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y, w: 5.1, h: 0.8, rectRadius: 0.08,
        fill: { color: i === 2 ? C.accent2 : C.background2 }, line: { color: C.background1 }, objectName: `Πεδίο ${t}` });
      s.addText([{ text: t + ":  ", options: { bold: true, color: C.text2 } }, { text: d }],
        { x: 7.8, y, w: 4.8, h: 0.8, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
    });
    bigStat(s, 7.6, 4.5, 5.1, "y = 1,079 m", "η τομή της καμπύλης Q(y) με την ευθεία-στόχο Q = 5 m³/s");
  }
  {
    const s = content("Ύψος θλιβόμενης ζώνης x", "Μ10",
      "4 λεπτά. Ορθογωνικό διάγραμμα τάσεων EN 1992-1-1, α_cc = 0,85 (ελληνικό ΕΠ, επιλογή από λίστα στο αρχείο). " +
      "Η εξίσωση γράφεται ως διαφορά = 0. Η MIN κόβει την τάση στο f_yd, το IF ελέγχει διαρροή. Η ίδια διάταξη λύνει και διατομή " +
      "όπου ο χάλυβας δεν διαρρέει (άσκηση). Μην ξεκινάτε από x = 0: διαίρεση με μηδέν.");
    const xs = [0.04, 0.08, 0.12, 0.16, 0.20, 0.24];
    s.addChart(pres.charts.SCATTER, [{ name: "x", values: xs }, { name: "F_c = λ·x·b·f_cd", values: xs.map((x) => +(3400 * x).toFixed(1)) },
      { name: "F_s = A_s·σ_s", values: xs.map(() => 546.5) }], {
      x: 0.6, y: 1.5, w: 6.4, h: 4.9, ...chartText, showTitle: true, title: "4Φ20, 0,30×0,60, C25/30, α_cc = 0,85",
      chartColors: [HEX.accent4, HEX.accent2], lineSize: 3, lineDataSymbol: "none", showLegend: true, legendPos: "b",
      showCatAxisTitle: true, catAxisTitle: "x (m)", showValAxisTitle: true, valAxisTitle: "Δύναμη (kN)", objectName: "Ισορροπία δυνάμεων" });
    s.addTable([["ε_s", "=ε_cu·(d − x)/x"], ["σ_s", "=MIN(E_s·ε_s ; f_yd)"], ["F_c − F_s", "→ 0 με αλλαγή του x"],
      ["Διαρροή;", '=IF(ε_s>=ε_yd;"ΝΑΙ";"ΟΧΙ")']].map((r) => r.map((c, j) => ({ text: c, options: {
        bold: j === 0, color: HEX.dk1, fill: { color: j === 0 ? HEX.lt2 : "FFFFFF" }, fontFace: j === 1 ? "Courier New" : undefined } }))),
      { x: 7.4, y: 1.6, w: 5.3, colW: [1.4, 3.9], fontSize: 15, rowH: 0.5, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Τύποι ισορροπίας" });
    bigStat(s, 7.4, 4.1, 2.6, "0,161 m", "ύψος x", C.text2, 40);
    bigStat(s, 10.1, 4.1, 2.6, "265 kNm", "ροπή αντοχής M_Rd", C.text2, 40);
    text(s, "ε_s = 8,5‰ > ε_yd = 2,17‰: ο χάλυβας διαρρέει", { x: 7.4, y: 6.0, w: 5.3, h: 0.5, fontSize: 15, color: C.accent1, bold: true });
  }
  {
    const s = content("Μητρώα: τριώροφο διατμητικό πλαίσιο", "Μ11",
      "2 λεπτά. Η ιδέα της μεθόδου άμεσης δυσκαμψίας. Ένας τύπος πίνακα: =MMULT(MINVERSE(K);F). Στο παλαιό Excel: επιλογή περιοχής " +
      "και Ctrl+Shift+Enter. Έλεγχος με το χέρι: τέμνουσες ορόφων V και Δ = V/k – ίδιο αποτέλεσμα.");
    text(s, "K·u = F", { x: 0.6, y: 1.5, w: 4, h: 0.6, fontSize: 26, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    sheet(s, 0.6, 2.15, [1.35, 1.35, 1.35, 0.3, 1.0, 0.3, 1.2], [
      [{ t: "110000", k: "f" }, { t: "−50000", k: "f" }, { t: "0", k: "f" }, "", { t: "u₁", k: "x" }, "", { t: "50", k: "in" }],
      [{ t: "−50000", k: "f" }, { t: "90000", k: "f" }, { t: "−40000", k: "f" }, "×", { t: "u₂", k: "x" }, "=", { t: "100", k: "in" }],
      [{ t: "0", k: "f" }, { t: "−40000", k: "f" }, { t: "40000", k: "f" }, "", { t: "u₃", k: "x" }, "", { t: "150", k: "in" }]], { fontSize: 15, rowH: 0.42 });
    text(s, "K (kN/m) από k = 60 000 / 50 000 / 40 000 · F (kN)", { x: 0.6, y: 4.0, w: 7.6, h: 0.35, fontSize: 13, color: C.accent5 });
    formula(s, "=MMULT(MINVERSE(K);F)*1000", 0.6, 4.5, 7.4, { fontSize: 18 });
    text(s, "u = 5,00 · 10,00 · 13,75 mm", { x: 0.6, y: 5.35, w: 7.6, h: 0.7, fontSize: 32, bold: true, color: C.text2,
      fontFace: THEME.headFontFace, valign: "middle" });
    text(s, "μετακινήσεις ορόφων", { x: 0.6, y: 6.05, w: 7.6, h: 0.35, fontSize: 14, color: C.accent5 });
    card(s, 8.6, 1.6, 4.1, 2.3, "Έλεγχος με το χέρι",
      "V₃ = 150, V₂ = 250, V₁ = 300 kN\nΔᵢ = Vᵢ/kᵢ → 5,00 · 5,00 · 3,75 mm\nu = ΣΔ ✔", "green");
    card(s, 8.6, 4.15, 4.1, 2.3, "det(K) = 0;",
      "Το μητρώο δεν αντιστρέφεται (#NUM!). Στη Στατική: μηχανισμός – λείπει στήριξη.", "warn");
  }
  {
    const s = content("Δικτύωμα: σύστημα 6×6", "Μ11",
      "2 λεπτά. Δύο εξισώσεις ανά κόμβο, 3 κόμβοι → 6 εξισώσεις για 3 αξονικές και 3 αντιδράσεις. Τα c, s είναι τύποι από L, h: " +
      "αλλάζει η γεωμετρία, ενημερώνεται η λύση. Έλεγχος με ροπές ως προς A: B_y·6 = 100·3 + 30·4 → 70 kN.");
    // simple truss sketch
    const ox = 0.9, oy = 5.6, sc = 0.75; // 6 m × 4 m truss
    const P = (x, y) => [ox + x * sc, oy - y * sc];
    const [ax, ay] = P(0, 0), [bx, by] = P(6, 0), [gx, gy] = P(3, 4);
    const line = (x1, y1, x2, y2, name) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2),
      w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001, flipH: (x2 - x1) * (y2 - y1) < 0,
      line: { color: C.text1, width: 3 }, objectName: name });
    line(ax, ay, bx, by, "Ράβδος AB"); line(ax, ay, gx, gy, "Ράβδος AΓ"); line(bx, by, gx, gy, "Ράβδος BΓ");
    s.addShape(pres.shapes.LINE, { x: gx, y: gy - 1.0, w: 0.001, h: 0.85, line: { color: C.accent3, width: 2.5, endArrowType: "triangle" }, objectName: "Φορτίο P" });
    s.addShape(pres.shapes.LINE, { x: gx + 0.15, y: gy, w: 0.9, h: 0.001, line: { color: C.accent3, width: 2.5, endArrowType: "triangle" }, objectName: "Φορτίο H" });
    text(s, "P = 100 kN", { x: gx - 1.55, y: gy - 1.0, w: 1.4, h: 0.4, fontSize: 14, color: C.accent3, bold: true, align: "right" });
    text(s, "H = 30 kN", { x: gx + 1.15, y: gy - 0.2, w: 1.4, h: 0.4, fontSize: 14, color: C.accent3, bold: true });
    text(s, "A", { x: ax - 0.45, y: ay - 0.2, w: 0.4, h: 0.4, fontSize: 16, bold: true });
    text(s, "B", { x: bx + 0.12, y: by - 0.2, w: 0.4, h: 0.4, fontSize: 16, bold: true });
    text(s, "Γ", { x: gx - 0.45, y: gy - 0.05, w: 0.4, h: 0.4, fontSize: 16, bold: true });
    text(s, "L = 6 m,  h = 4 m", { x: ax, y: ay + 0.25, w: 4.5, h: 0.4, fontSize: 14, color: C.accent5 });
    formula(s, "=MMULT(MINVERSE(A);b)", 6.3, 1.6, 6.4, { fontSize: 18 });
    s.addTable([["Άγνωστος", "kN", "Είδος"], ["N_AB", "52,5", "εφελκυσμός"], ["N_AΓ", "−37,5", "θλίψη"], ["N_BΓ", "−87,5", "θλίψη"],
      ["A_x · A_y · B_y", "−30 · 30 · 70", "αντιδράσεις"]].map((r, i) => r.map((c, j) => ({ text: c, options: {
        bold: i === 0, color: i === 0 ? HEX.lt1 : (c === "θλίψη" ? HEX.accent3 : HEX.dk1), fill: { color: i === 0 ? HEX.accent1 : "FFFFFF" },
        align: j === 1 ? "right" : "left" } }))),
      { x: 6.3, y: 2.5, w: 6.4, colW: [2.3, 2.0, 2.1], fontSize: 15, rowH: 0.5, border: { type: "solid", pt: 0.75, color: "D4D4D4" },
        valign: "middle", objectName: "Λύση δικτυώματος" });
    card(s, 6.3, 5.25, 6.4, 1.2, "Ένας τύπος λύνει 6 εξισώσεις", "Αλλάξτε L, h, P, H – η λύση ενημερώνεται αμέσως.", "green");
  }

  // ================= Μέρος 5 =================
  section("Αυτοματοποίηση και σύνοψη");
  {
    const s = content("Μακροεντολές: μια πρώτη γνωριμία", "Μ13",
      "3 λεπτά. Εγγραφή μακροεντολής μορφοποίησης, δικές σας συναρτήσεις, βρόχος που τρέχει την Αναζήτηση στόχου για πολλές διατομές. " +
      "Στη VBA η υποδιαστολή είναι πάντα τελεία. Ασφάλεια: ποτέ ενεργοποίηση μακροεντολών από άγνωστη πηγή· αποθήκευση ως .xlsm.");
    const code = "Function BAR_AREA(n As Double, phi As Double) As Double\n" +
      "    BAR_AREA = n * WorksheetFunction.Pi() * phi ^ 2 / 4 / 100\nEnd Function\n\n" +
      "Sub SolveAll()\n    Dim c As Integer\n    For c = 3 To 5\n" +
      "        Cells(23, c).GoalSeek Goal:=0, ChangingCell:=Cells(17, c)\n    Next c\nEnd Sub";
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 1.55, w: 7.6, h: 4.0, rectRadius: 0.1, fill: { color: "1E1E1E" },
      line: { color: "1E1E1E" }, objectName: "Πλαίσιο κώδικα" });
    s.addText(code, { x: 0.85, y: 1.8, w: 7.2, h: 3.6, fontFace: "Courier New", fontSize: 12, color: "D4D4D4", valign: "top",
      margin: 0, isTextBox: true, objectName: "Κώδικας VBA" });
    const steps = [["1", "Εγγραφή", "Προγραμματιστής → Εγγραφή μακροεντολής"], ["2", "Συνάρτηση", "=BAR_AREA(4;16) → 8,04 cm²"],
      ["3", "Βρόχος", "Αναζήτηση στόχου για 3 διατομές με ένα κουμπί"]];
    steps.forEach(([n, t, d], i) => {
      const y = 1.6 + i * 1.15;
      s.addShape(pres.shapes.OVAL, { x: 8.6, y, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `Βήμα ${n}` });
      s.addText(n, { x: 8.6, y, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle",
        margin: 0, isTextBox: true });
      s.addText([{ text: t, options: { bold: true, color: C.text2, fontSize: 17, breakLine: true } }, { text: d, options: { fontSize: 14 } }],
        { x: 9.45, y: y - 0.05, w: 3.3, h: 0.95, color: C.text1, valign: "top", margin: 0, isTextBox: true });
    });
    card(s, 8.6, 5.1, 4.1, 1.35, "Ασφάλεια", "Μακροεντολές μόνο από πηγές που εμπιστεύεστε.", "err");
  }
  {
    const s = content("Τα λάθη που δεν βγάζουν μήνυμα", null,
      "2 λεπτά. Ανακεφαλαίωση με τον πιο πρακτικό τρόπο: τα λάθη που κάνουν το φύλλο να δίνει λάθος αριθμό χωρίς καμία προειδοποίηση. " +
      "Συμβουλή: ελέγχετε πάντα ένα αποτέλεσμα με το χέρι και κάθε νέο εργαλείο σε πρόβλημα με γνωστή λύση.");
    const errs = [["1.250 → 1250", "τελεία ως διαχωριστικό χιλιάδων", "Οδηγός"],
      ["=B4/B1 χωρίς $", "η αναφορά «γλιστράει» στην αντιγραφή", "Μ3"],
      ["=M/b*h^2/6", "λείπει η παρένθεση του παρονομαστή", "Μ2"],
      ["COS(30)", "ακτίνια, όχι μοίρες", "Μ4"],
      ["OK ≠ ΟΚ", "ελληνικά και λατινικά γράμματα", "Μ7"],
      ["μονάδες", "kN/m² σε MPa χωρίς /1000", "Μ5"]];
    errs.forEach(([t, d, l], i) => {
      const x = 0.6 + (i % 3) * 4.1, y = 1.6 + Math.floor(i / 3) * 2.45;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 3.8, h: 2.15, rectRadius: 0.1, fill: { color: C.accent3, transparency: 90 },
        line: { color: C.background1 }, objectName: `Λάθος ${i + 1}` });
      text(s, t, { x: x + 0.25, y: y + 0.25, w: 3.3, h: 0.7, fontSize: 22, bold: true, color: C.accent3, fontFace: "Courier New", valign: "middle" });
      text(s, d, { x: x + 0.25, y: y + 1.0, w: 3.3, h: 0.7, fontSize: 15 });
      text(s, l, { x: x + 2.9, y: y + 1.7, w: 0.75, h: 0.3, fontSize: 12, color: C.accent5, align: "right" });
    });
  }
  {
    const s = content("Πώς δουλεύετε με το υλικό", null,
      "2 λεπτά. Ενθαρρύνετε να κάνουν τις ασκήσεις μόνοι τους· οι τιμές ελέγχου είναι στα φύλλα. Πρόταση: κάθε άσκηση Στατικής ή " +
      "Αντοχής που λύνουν με το χέρι, να την ελέγχουν και σε Excel.");
    const steps = [["Διαβάστε", "τη σελίδα του μαθήματος, με τα διαδραστικά παραδείγματα"],
      ["Κατεβάστε", "το αρχείο .xlsx: φύλλο Παράδειγμα και φύλλο Άσκηση"],
      ["Λύστε", "την άσκηση μόνοι σας – κάθε ενότητα έχει τιμή ελέγχου"],
      ["Εφαρμόστε", "σε ασκήσεις άλλων μαθημάτων: Στατική, Αντοχή, Δομικά Υλικά"]];
    steps.forEach(([t, d], i) => {
      const x = 0.6 + i * 3.1;
      s.addShape(pres.shapes.OVAL, { x: x + 0.95, y: 1.7, w: 1.0, h: 1.0, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `Βήμα ${i + 1}` });
      s.addText(String(i + 1), { x: x + 0.95, y: 1.7, w: 1.0, h: 1.0, fontSize: 28, bold: true, color: C.background1, align: "center",
        valign: "middle", margin: 0, isTextBox: true });
      text(s, t, { x, y: 2.95, w: 2.9, h: 0.5, fontSize: 20, bold: true, color: C.text2, align: "center" });
      text(s, d, { x: x + 0.1, y: 3.55, w: 2.7, h: 1.2, fontSize: 15, align: "center" });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.9, y: 5.2, w: 7.5, h: 1.0, rectRadius: 0.1, fill: { color: C.background2 },
      line: { color: C.accent1, width: 1.5 }, objectName: "Διεύθυνση site" });
    s.addText(SITE, { x: 2.9, y: 5.2, w: 7.5, h: 1.0, fontSize: 26, bold: true, color: C.text2, align: "center", valign: "middle",
      margin: 0, isTextBox: true, hyperlink: { url: "https://" + SITE + "/" } });
  }
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle });
    s.addText("Ερωτήσεις;", { placeholder: "title" });
    s.addText([{ text: "13 μαθήματα · 13 αρχεία ασκήσεων · ένας οδηγός τοπικών ρυθμίσεων", options: { breakLine: true } },
               { text: SITE, options: { fontSize: 16 } }], { placeholder: "body" });
    s.addNotes("Χρόνος για ερωτήσεις (περίπου 5 λεπτά). Υπενθυμίστε ότι όλο το υλικό είναι διαθέσιμο online.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Saved", OUT);
})();
