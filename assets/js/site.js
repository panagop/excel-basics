// Excel για Πολιτικούς Μηχανικούς – κοινή πλοήγηση και διαδραστικό υπολογιστικό φύλλο.
// Νέο μάθημα: δημιουργήστε lessons/NN-name.html και ορίστε `ready: true` παρακάτω.
// Νέος οδηγός (σελίδα εκτός σειράς μαθημάτων): guides/name.html και μια γραμμή στο GUIDES.

const LESSONS = [
  { file: "01-first-steps.html", title: "Πρώτα βήματα: το περιβάλλον του Excel", ready: true },
  { file: "02-formulas.html", title: "Τύποι και αριθμητικές πράξεις", ready: true },
  { file: "03-references.html", title: "Σχετικές και απόλυτες αναφορές", ready: true },
  { file: "04-functions.html", title: "Βασικές συναρτήσεις (SUM, AVERAGE, …)", ready: true },
  { file: "05-formatting.html", title: "Μορφοποίηση αριθμών και μονάδες", ready: true },
  { file: "06-charts.html", title: "Διαγράμματα (διασπορά XY)" },
  { file: "07-if.html", title: "Λογικοί έλεγχοι: IF, AND, OR" },
  { file: "08-lookup.html", title: "Πίνακες αναζήτησης (XLOOKUP)" },
  { file: "09-sort-filter.html", title: "Ταξινόμηση, φίλτρα και Πίνακες" },
  { file: "10-goal-seek.html", title: "Αναζήτηση στόχου: επίλυση εξισώσεων" },
  { file: "11-matrices.html", title: "Μητρώα και γραμμικά συστήματα" },
];

const GUIDES = [
  { file: "regional-settings.html", title: "Τοπικές ρυθμίσεις: τελεία ή κόμμα;" },
];

(function buildLayout() {
  const inSubfolder = /\/(lessons|guides)\//.test(location.pathname);
  const root = inSubfolder ? "../" : "";
  const current = location.pathname.split("/").pop();

  // Header
  const header = document.createElement("header");
  header.className = "site-header";
  header.innerHTML = `
    <button class="menu-btn" aria-label="Εμφάνιση μαθημάτων">☰</button>
    <span class="logo">X</span>
    <a href="${root}index.html">Excel για Πολιτικούς Μηχανικούς</a>`;
  document.body.prepend(header);

  // Sidebar
  const items = LESSONS.map((l, i) => {
    const n = String(i + 1).padStart(2, "0");
    if (!l.ready) {
      return `<li><span><span class="num">${n}</span>${l.title}<span class="soon">σύντομα</span></span></li>`;
    }
    const active = l.file === current ? ' class="active"' : "";
    return `<li><a href="${root}lessons/${l.file}"${active}><span class="num">${n}</span>${l.title}</a></li>`;
  }).join("");
  const guideItems = GUIDES.map((g) => {
    const active = g.file === current ? ' class="active"' : "";
    return `<li><a href="${root}guides/${g.file}"${active}><span class="num">ⓘ</span>${g.title}</a></li>`;
  }).join("");
  const nav = document.createElement("nav");
  nav.className = "sidebar";
  nav.innerHTML = `
    <ol><li><a href="${root}index.html"${current === "index.html" || current === "" ? ' class="active"' : ""}>
      <span class="num">⌂</span>Αρχική</a></li></ol>
    <h2>Μαθήματα</h2><ol>${items}</ol>
    <h2>Οδηγοί</h2><ol>${guideItems}</ol>`;
  header.after(nav);

  header.querySelector(".menu-btn").addEventListener("click", () =>
    document.body.classList.toggle("nav-open"));
  nav.addEventListener("click", (e) => {
    if (e.target.closest("a")) document.body.classList.remove("nav-open");
  });

  // Prev / next buttons
  const pager = document.querySelector(".pager");
  const idx = LESSONS.findIndex((l) => l.file === current);
  if (pager && idx >= 0) {
    const prev = LESSONS[idx - 1];
    const next = LESSONS[idx + 1];
    pager.innerHTML =
      (prev ? `<a class="prev" href="${prev.file}"><small>← Προηγούμενο</small>${prev.title}</a>`
            : `<a class="prev" href="../index.html"><small>← Επιστροφή</small>Αρχική</a>`) +
      (next && next.ready ? `<a class="next" href="${next.file}"><small>Επόμενο →</small>${next.title}</a>` : "");
  }

  // Home page lesson cards
  const cards = document.querySelector(".cards");
  if (cards) {
    cards.innerHTML = LESSONS.map((l, i) => {
      const n = String(i + 1).padStart(2, "0");
      return l.ready
        ? `<a class="card" href="lessons/${l.file}"><span class="num">Μάθημα ${n}</span><h3>${l.title}</h3></a>`
        : `<div class="card disabled"><span class="num">Μάθημα ${n}</span><h3>${l.title}</h3><p>Σύντομα διαθέσιμο</p></div>`;
    }).join("");
  }
  const guideCards = document.querySelector(".guide-cards");
  if (guideCards) {
    guideCards.innerHTML = GUIDES.map((g) =>
      `<a class="card" href="guides/${g.file}"><span class="num">Οδηγός</span><h3>${g.title}</h3></a>`).join("");
  }
})();

// ---------- Mock spreadsheet ----------
// Usage: <div class="xl" data-sheet='{"cols":4,"rows":5,"cells":{"A1":{"v":"Beam","hdr":1}, "C2":{"v":0.15,"f":"=A2*B2"}}}'></div>
// Each cell: v = value, f = formula (optional), d = decimals shown (optional), hdr = header style.
// r = right-align a text value (e.g. a formatted number "20,25 kNm"; put the raw value in f).
// bg / color = cell fill and text colour (e.g. to show conditional formatting).
// Numbers are shown with a decimal comma, as in Excel with Greek regional settings.
const fmtNum = (v, d) =>
  typeof v !== "number" ? v ?? ""
    : v.toLocaleString("el-GR", { useGrouping: false, minimumFractionDigits: d ?? 0, maximumFractionDigits: d ?? 10 });

document.querySelectorAll(".xl[data-sheet]").forEach((box) => {
  const sheet = JSON.parse(box.dataset.sheet);
  const letters = Array.from({ length: sheet.cols }, (_, i) => String.fromCharCode(65 + i));

  let html = `<div class="xl-bar"><input class="xl-name" readonly value="A1">
    <span class="xl-fx">fx</span><span class="xl-formula"></span></div><table><tr><th></th>`;
  html += letters.map((c) => `<th data-col="${c}">${c}</th>`).join("") + "</tr>";
  for (let r = 1; r <= sheet.rows; r++) {
    html += `<tr><th data-row="${r}">${r}</th>`;
    for (const c of letters) {
      const cell = sheet.cells[c + r] || {};
      const cls = [
        typeof cell.v === "number" || cell.r ? "num" : "",
        cell.hdr ? "hdr" : "",
        cell.f ? "formula" : "",
      ].join(" ").trim();
      const style = cell.bg ? ` style="background:${cell.bg};color:${cell.color ?? "inherit"}"` : "";
      html += `<td data-ref="${c}${r}" class="${cls}"${style}>${fmtNum(cell.v, cell.d)}</td>`;
    }
    html += "</tr>";
  }
  html += "</table>";
  if (sheet.caption) html += `<div class="xl-caption">${sheet.caption}</div>`;
  box.innerHTML = html;

  const nameBox = box.querySelector(".xl-name");
  const formulaBar = box.querySelector(".xl-formula");
  const select = (ref) => {
    box.querySelectorAll(".sel").forEach((el) => el.classList.remove("sel"));
    const td = box.querySelector(`td[data-ref="${ref}"]`);
    td.classList.add("sel");
    const [, col, row] = ref.match(/([A-Z]+)(\d+)/);
    box.querySelector(`th[data-col="${col}"]`).classList.add("sel");
    box.querySelector(`th[data-row="${row}"]`).classList.add("sel");
    const cell = sheet.cells[ref] || {};
    nameBox.value = ref;
    formulaBar.textContent = cell.f ?? fmtNum(cell.v);
  };
  box.addEventListener("click", (e) => {
    const td = e.target.closest("td[data-ref]");
    if (td) select(td.dataset.ref);
  });
  select(sheet.start || "A1");
});
