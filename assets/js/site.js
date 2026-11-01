// Εισαγωγή στο Excel για Πολιτικούς Μηχανικούς – κοινή πλοήγηση και διαδραστικό υπολογιστικό φύλλο.
// Νέο μάθημα: δημιουργήστε lessons/NN-name.html και ορίστε `ready: true` παρακάτω.
// Νέος οδηγός (σελίδα εκτός σειράς μαθημάτων): guides/name.html και μια γραμμή στο GUIDES.

const LESSONS = [
  { file: "01-first-steps.html", title: "Πρώτα βήματα: το περιβάλλον του Excel", ready: true },
  { file: "02-formulas.html", title: "Τύποι και αριθμητικές πράξεις", ready: true },
  { file: "03-references.html", title: "Σχετικές και απόλυτες αναφορές", ready: true },
  { file: "04-functions.html", title: "Βασικές συναρτήσεις (SUM, AVERAGE, …)", ready: true },
  { file: "05-formatting.html", title: "Μορφοποίηση αριθμών και μονάδες", ready: true },
  { file: "06-charts.html", title: "Διαγράμματα (διασπορά XY)", ready: true },
  { file: "07-if.html", title: "Λογικοί έλεγχοι: IF, AND, OR", ready: true },
  { file: "08-lookup.html", title: "Πίνακες αναζήτησης (XLOOKUP)", ready: true },
  { file: "09-sort-filter.html", title: "Ταξινόμηση, φίλτρα και Πίνακες", ready: true },
  { file: "10-goal-seek.html", title: "Αναζήτηση στόχου: επίλυση εξισώσεων", ready: true },
  { file: "11-matrices.html", title: "Μητρώα και γραμμικά συστήματα", ready: true },
  { file: "12-more-charts.html", title: "Στήλες, ράβδοι, πίτα, ιστόγραμμα", ready: true },
  { file: "13-macros.html", title: "Μακροεντολές: σύντομη εισαγωγή", ready: true },
];

// Υπογραφή στο υποσέλιδο κάθε σελίδας.
const AUTHOR = "Γιώργος Παναγόπουλος";
const YEAR = "2026";

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
    <a href="${root}index.html">Εισαγωγή στο Excel για Πολιτικούς Μηχανικούς</a>`;
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

  // Footer with the logos of the University and the Department, author and year
  const main = document.querySelector("main");
  if (main) {
    const footer = document.createElement("footer");
    footer.className = "site-footer";
    footer.innerHTML = `
      <a href="https://www.ihu.gr/" target="_blank" rel="noopener"><img src="${root}assets/img/ihu-logo.png"
        alt="Διεθνές Πανεπιστήμιο της Ελλάδος" height="44"></a>
      <a href="https://civil.ihu.gr/" target="_blank" rel="noopener"><img src="${root}assets/img/civil-ihu-logo.png"
        alt="Τμήμα Πολιτικών Μηχανικών ΔιΠΑΕ" height="44"></a>
      <span class="sign">${[AUTHOR, YEAR].filter(Boolean).join(", ")}</span>`;
    main.append(footer);
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
const OFFICE_COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

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
  if (spec.type === "pie") return drawPie(box, spec);
  if (spec.type === "column" || spec.type === "bar") return drawBars(box, spec);
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

// ---------- Column / bar / combo and pie charts (SVG) ----------
// Column/bar: {"type":"column"|"bar","categories":[…],"series":[{"name","values":[…],"line":true?}],
//   "stacked":true?, "gap":0.4, "labels":true?, "thousands":true? (1.234 grouping), "yTitle",
//   "y2Title" (axis of line series), "xTitle"}
// Pie: {"type":"pie","categories":[…],"values":[…]}

function legendRow(series, x, y) {
  let svg = "";
  series.forEach((s, k) => {
    const col = s.color || OFFICE_COLORS[k % OFFICE_COLORS.length];
    svg += s.line
      ? `<line x1="${x}" x2="${x + 18}" y1="${y - 4}" y2="${y - 4}" stroke="${col}" stroke-width="2.5"/>`
      : `<rect x="${x + 3}" y="${y - 10}" width="12" height="12" fill="${col}"/>`;
    svg += `<text x="${x + 22}" y="${y}">${s.name}</text>`;
    x += 34 + s.name.length * 6.5;
  });
  return svg;
}

function drawBars(box, spec) {
  const W = 600, H = spec.height || 320, horiz = spec.type === "bar";
  const cats = spec.categories, n = cats.length;
  const bars = spec.series.filter((s) => !s.line), lines = spec.series.filter((s) => s.line);
  const showLegend = spec.series.length > 1;
  const fmtV = (v) => (spec.thousands ? v.toLocaleString("el-GR") : fmtNum(v));
  const m = { l: horiz ? spec.labelWidth || 160 : 62, r: lines.length ? 62 : horiz && spec.labels ? 70 : 22,
              t: (spec.title ? 36 : 14) + (showLegend ? 22 : 0), b: 50 };
  const tot = cats.map((_, i) => (spec.stacked ? bars.reduce((s, b) => s + b.values[i], 0)
                                               : Math.max(...bars.map((b) => b.values[i]))));
  const vt = niceTicks(0, Math.max(...tot)), vmax = vt[vt.length - 1];
  const gap = spec.gap ?? 0.4;
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" font-family="Calibri, Segoe UI, Arial" font-size="12">`;
  if (spec.title) svg += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="15" font-weight="600">${spec.title}</text>`;
  if (showLegend) svg += legendRow(spec.series, m.l, (spec.title ? 36 : 14) + 12);
  const x0 = m.l, x1 = W - m.r, y0 = m.t, y1 = H - m.b;

  if (!horiz) {
    const vy = (v) => y1 - (v / vmax) * (y1 - y0), band = (x1 - x0) / n, gw = band * (1 - gap);
    for (const v of vt) {
      svg += `<line x1="${x0}" x2="${x1}" y1="${vy(v)}" y2="${vy(v)}" stroke="#e3e3e3"/>`;
      svg += `<text x="${x0 - 6}" y="${vy(v) + 4}" text-anchor="end" fill="#555">${fmtNum(v)}</text>`;
    }
    cats.forEach((c, i) => {
      const cx = x0 + band * (i + 0.5);
      svg += `<text x="${cx}" y="${y1 + 16}" text-anchor="middle" fill="#555">${c}</text>`;
      let acc = 0;
      bars.forEach((s, k) => {
        const col = s.color || OFFICE_COLORS[spec.series.indexOf(s) % OFFICE_COLORS.length];
        const v = s.values[i], w = spec.stacked ? gw : gw / bars.length;
        const bx = spec.stacked ? cx - gw / 2 : cx - gw / 2 + k * w;
        const top = vy(spec.stacked ? acc + v : v), bottom = vy(spec.stacked ? acc : 0);
        svg += `<rect x="${bx.toFixed(1)}" y="${top.toFixed(1)}" width="${w.toFixed(1)}" height="${(bottom - top).toFixed(1)}" fill="${col}"/>`;
        if (spec.labels) svg += `<text x="${bx + w / 2}" y="${top - 4}" text-anchor="middle" fill="#333">${fmtNum(v)}</text>`;
        acc += v;
      });
    });
    if (lines.length) {
      const lt = niceTicks(0, Math.max(...lines.flatMap((s) => s.values))), lmax = lt[lt.length - 1];
      const ly = (v) => y1 - (v / lmax) * (y1 - y0);
      for (const v of lt) svg += `<text x="${x1 + 6}" y="${ly(v) + 4}" fill="#555">${fmtNum(v)}</text>`;
      lines.forEach((s) => {
        const col = s.color || OFFICE_COLORS[spec.series.indexOf(s) % OFFICE_COLORS.length];
        const pts = s.values.map((v, i) => `${(x0 + band * (i + 0.5)).toFixed(1)},${ly(v).toFixed(1)}`);
        svg += `<polyline points="${pts.join(" ")}" fill="none" stroke="${col}" stroke-width="2.5"/>`;
        svg += pts.map((p) => `<circle cx="${p.split(",")[0]}" cy="${p.split(",")[1]}" r="3" fill="${col}"/>`).join("");
      });
      if (spec.y2Title) svg += `<text transform="translate(${W - 10} ${(y0 + y1) / 2}) rotate(90)" text-anchor="middle" font-size="13">${spec.y2Title}</text>`;
    }
    svg += `<line x1="${x0}" x2="${x1}" y1="${y1}" y2="${y1}" stroke="#888"/>`;
    if (spec.yTitle) svg += `<text transform="translate(16 ${(y0 + y1) / 2}) rotate(-90)" text-anchor="middle" font-size="13">${spec.yTitle}</text>`;
  } else {
    const vx = (v) => x0 + (v / vmax) * (x1 - x0), band = (y1 - y0) / n, bh = band * (1 - gap);
    for (const v of vt) {
      svg += `<line x1="${vx(v)}" x2="${vx(v)}" y1="${y0}" y2="${y1}" stroke="#e3e3e3"/>`;
      svg += `<text x="${vx(v)}" y="${y1 + 16}" text-anchor="middle" fill="#555">${fmtV(v)}</text>`;
    }
    cats.forEach((c, i) => {
      const cy = y0 + band * (i + 0.5), v = bars[0].values[i];
      svg += `<text x="${x0 - 6}" y="${cy + 4}" text-anchor="end" fill="#333">${c}</text>`;
      svg += `<rect x="${x0}" y="${(cy - bh / 2).toFixed(1)}" width="${(vx(v) - x0).toFixed(1)}" height="${bh.toFixed(1)}" fill="${bars[0].color || OFFICE_COLORS[0]}"/>`;
      if (spec.labels) svg += `<text x="${vx(v) + 5}" y="${cy + 4}" fill="#333">${fmtV(v)}</text>`;
    });
    svg += `<line x1="${x0}" x2="${x0}" y1="${y0}" y2="${y1}" stroke="#888"/>`;
  }
  if (spec.xTitle) svg += `<text x="${(x0 + x1) / 2}" y="${H - 10}" text-anchor="middle" font-size="13">${spec.xTitle}</text>`;
  box.innerHTML = svg + "</svg>" + (spec.caption ? `<div class="xl-caption">${spec.caption}</div>` : "");
}

function drawPie(box, spec) {
  const W = 600, H = spec.height || 300, cx = 160, cy = (H + (spec.title ? 24 : 0)) / 2, r = Math.min(120, cy - 30);
  const total = spec.values.reduce((a, b) => a + b, 0);
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" font-family="Calibri, Segoe UI, Arial" font-size="12">`;
  if (spec.title) svg += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="15" font-weight="600">${spec.title}</text>`;
  let a0 = -Math.PI / 2;
  spec.values.forEach((v, k) => {
    const a1 = a0 + (v / total) * 2 * Math.PI, col = OFFICE_COLORS[k % OFFICE_COLORS.length];
    const p = (a) => `${(cx + r * Math.cos(a)).toFixed(1)},${(cy + r * Math.sin(a)).toFixed(1)}`;
    svg += `<path d="M${cx},${cy} L${p(a0)} A${r},${r} 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${p(a1)} z" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
    const am = (a0 + a1) / 2;
    svg += `<text x="${cx + 0.68 * r * Math.cos(am)}" y="${cy + 0.68 * r * Math.sin(am) + 4}" text-anchor="middle" fill="#fff" font-weight="600">${fmtNum(+(100 * v / total).toFixed(0))}%</text>`;
    const ly = cy - (spec.values.length * 22) / 2 + k * 22 + 10;
    svg += `<rect x="330" y="${ly - 10}" width="12" height="12" fill="${col}"/><text x="350" y="${ly}">${spec.categories[k]}</text>`;
    a0 = a1;
  });
  box.innerHTML = svg + "</svg>" + (spec.caption ? `<div class="xl-caption">${spec.caption}</div>` : "");
}
