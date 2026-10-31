// Excel για Πολιτικούς Μηχανικούς – κοινή πλοήγηση και διαδραστικό υπολογιστικό φύλλο.
// Νέο μάθημα: δημιουργήστε lessons/NN-name.html και ορίστε `ready: true` παρακάτω.
// Νέος οδηγός (σελίδα εκτός σειράς μαθημάτων): guides/name.html και μια γραμμή στο GUIDES.

const LESSONS = [
  { file: "01-first-steps.html", title: "Πρώτα βήματα: το περιβάλλον του Excel", ready: true },
  { file: "02-formulas.html", title: "Τύποι και αριθμητικές πράξεις", ready: true },
  { file: "03-references.html", title: "Σχετικές και απόλυτες αναφορές", ready: true },
  { file: "04-functions.html", title: "Βασικές συναρτήσεις (SUM, AVERAGE, …)", ready: true },
  { file: "05-formatting.html", title: "Μορφοποίηση αριθμών και μονάδες", ready: true },
  { file: "06-charts.html", title: "Διαγράμματα (διασπορά XY)", ready: true },
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

// ---------- XY chart (SVG) ----------
// Usage: <div class="xychart" data-chart='{"title":"…","xTitle":"…","yTitle":"…",
//   "series":[{"name":"…","x":[…],"y":[…],"markers":true,"line":true,"dash":false}]}'></div>
// Optional: xMin/xMax/yMin/yMax, height, categorical:true (x values as equally spaced categories,
// like an Excel Line chart), legend:true, legendLeft:true (legend top-left instead of top-right).
const CHART_COLORS = ["#4472C4", "#ED7D31", "#70AD47", "#7F7F7F"];

function niceTicks(lo, hi, target = 5) {
  if (hi === lo) hi = lo + 1;
  const raw = (hi - lo) / target;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((f) => f * mag).find((s) => s >= raw);
  const ticks = [];
  const first = Math.floor(lo / step + 1e-9) * step;
  for (let v = first; ticks.length < 2 || ticks[ticks.length - 1] < hi - step * 1e-9; v += step) ticks.push(+v.toFixed(10));
  return ticks;
}

document.querySelectorAll(".xychart[data-chart]").forEach((box) => {
  const spec = JSON.parse(box.dataset.chart);
  const W = 600, H = spec.height || 320;
  const m = { l: 62, r: 18, t: spec.title ? 36 : 14, b: 50 };
  const series = spec.series.map((s) => (spec.categorical ? { ...s, cat: s.x, x: s.x.map((_, i) => i) } : s));
  const allX = series.flatMap((s) => s.x), allY = series.flatMap((s) => s.y);

  let xt = spec.categorical ? allX.filter((v, i, a) => a.indexOf(v) === i)
                            : niceTicks(spec.xMin ?? Math.min(0, ...allX), spec.xMax ?? Math.max(...allX));
  const yt = niceTicks(spec.yMin ?? Math.min(0, ...allY), spec.yMax ?? Math.max(...allY));
  const [x0, x1] = spec.categorical ? [-0.5, Math.max(...allX) + 0.5] : [xt[0], xt[xt.length - 1]];
  const [y0, y1] = [yt[0], yt[yt.length - 1]];
  const px = (x) => m.l + ((x - x0) / (x1 - x0)) * (W - m.l - m.r);
  const py = (y) => H - m.b - ((y - y0) / (y1 - y0)) * (H - m.t - m.b);
  const lbl = (v) => fmtNum(+v.toPrecision(6));

  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" font-family="Calibri, Segoe UI, Arial" font-size="12">`;
  if (spec.title) svg += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="15" font-weight="600">${spec.title}</text>`;
  for (const y of yt) {
    svg += `<line x1="${m.l}" x2="${W - m.r}" y1="${py(y)}" y2="${py(y)}" stroke="#e3e3e3"/>`;
    svg += `<text x="${m.l - 6}" y="${py(y) + 4}" text-anchor="end" fill="#555">${lbl(y)}</text>`;
  }
  const xlabels = spec.categorical ? series[0].cat.map((c, i) => [i, c]) : xt.map((x) => [x, x]);
  for (const [x, text] of xlabels) {
    if (!spec.categorical) svg += `<line x1="${px(x)}" x2="${px(x)}" y1="${m.t}" y2="${H - m.b}" stroke="#f0f0f0"/>`;
    svg += `<text x="${px(x)}" y="${H - m.b + 16}" text-anchor="middle" fill="#555">${lbl(text)}</text>`;
  }
  svg += `<line x1="${m.l}" x2="${W - m.r}" y1="${py(Math.max(y0, Math.min(0, y1)))}" y2="${py(Math.max(y0, Math.min(0, y1)))}" stroke="#888"/>`;
  svg += `<line x1="${m.l}" x2="${m.l}" y1="${m.t}" y2="${H - m.b}" stroke="#888"/>`;
  if (spec.xTitle) svg += `<text x="${(m.l + W - m.r) / 2}" y="${H - 10}" text-anchor="middle" font-size="13">${spec.xTitle}</text>`;
  if (spec.yTitle) svg += `<text transform="translate(16 ${(m.t + H - m.b) / 2}) rotate(-90)" text-anchor="middle" font-size="13">${spec.yTitle}</text>`;

  series.forEach((s, k) => {
    const col = s.color || CHART_COLORS[k % CHART_COLORS.length];
    const pts = s.x.map((x, i) => `${px(x).toFixed(1)},${py(s.y[i]).toFixed(1)}`);
    if (s.line !== false)
      svg += `<polyline points="${pts.join(" ")}" fill="none" stroke="${col}" stroke-width="2.2"${s.dash ? ' stroke-dasharray="6 4"' : ""}/>`;
    if (s.markers !== false)
      svg += pts.map((p) => `<circle cx="${p.split(",")[0]}" cy="${p.split(",")[1]}" r="3.5" fill="${col}"/>`).join("");
  });
  if (spec.legend || series.length > 1) {
    series.forEach((s, k) => {
      const col = s.color || CHART_COLORS[k % CHART_COLORS.length];
      const y = m.t + 8 + k * 18;
      const lx = spec.legendLeft ? m.l + 14 : W - m.r - 150;
      svg += s.line === false
        ? `<circle cx="${lx + 11}" cy="${y}" r="3.5" fill="${col}"/>`
        : `<line x1="${lx}" x2="${lx + 22}" y1="${y}" y2="${y}" stroke="${col}" stroke-width="2.5"${s.dash ? ' stroke-dasharray="5 3"' : ""}/>`;
      svg += `<text x="${lx + 28}" y="${y + 4}">${s.name}</text>`;
    });
  }
  box.innerHTML = svg + "</svg>" + (spec.caption ? `<div class="xl-caption">${spec.caption}</div>` : "");
});
