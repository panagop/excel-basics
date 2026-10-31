// Excel Basics – shared navigation and the interactive mock spreadsheet.
// To add a lesson: create lessons/NN-name.html and set `ready: true` below.

const LESSONS = [
  { file: "01-first-steps.html", title: "First steps: the Excel window", ready: true },
  { file: "02-formulas.html", title: "Formulas and arithmetic" },
  { file: "03-references.html", title: "Relative and absolute references" },
  { file: "04-functions.html", title: "Basic functions (SUM, AVERAGE, …)" },
  { file: "05-formatting.html", title: "Formatting numbers and units" },
  { file: "06-charts.html", title: "Charts (XY scatter)" },
  { file: "07-if.html", title: "Logical tests: IF, AND, OR" },
  { file: "08-lookup.html", title: "Lookup tables (XLOOKUP)" },
  { file: "09-sort-filter.html", title: "Sorting, filtering and Tables" },
  { file: "10-goal-seek.html", title: "Goal Seek: solving equations" },
  { file: "11-matrices.html", title: "Matrices and linear systems" },
];

(function buildLayout() {
  const inLessons = location.pathname.includes("/lessons/");
  const root = inLessons ? "../" : "";
  const current = location.pathname.split("/").pop();

  // Header
  const header = document.createElement("header");
  header.className = "site-header";
  header.innerHTML = `
    <button class="menu-btn" aria-label="Show lessons">☰</button>
    <span class="logo">X</span>
    <a href="${root}index.html">Excel for Civil Engineers</a>`;
  document.body.prepend(header);

  // Sidebar
  const items = LESSONS.map((l, i) => {
    const n = String(i + 1).padStart(2, "0");
    if (!l.ready) {
      return `<li><span><span class="num">${n}</span>${l.title}<span class="soon">soon</span></span></li>`;
    }
    const active = l.file === current ? ' class="active"' : "";
    return `<li><a href="${root}lessons/${l.file}"${active}><span class="num">${n}</span>${l.title}</a></li>`;
  }).join("");
  const nav = document.createElement("nav");
  nav.className = "sidebar";
  nav.innerHTML = `
    <ol><li><a href="${root}index.html"${current === "index.html" || current === "" ? ' class="active"' : ""}>
      <span class="num">⌂</span>Home</a></li></ol>
    <h2>Lessons</h2><ol>${items}</ol>`;
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
      (prev ? `<a class="prev" href="${prev.file}"><small>← Previous</small>${prev.title}</a>`
            : `<a class="prev" href="../index.html"><small>← Back</small>Home</a>`) +
      (next && next.ready ? `<a class="next" href="${next.file}"><small>Next →</small>${next.title}</a>` : "");
  }

  // Home page lesson cards
  const cards = document.querySelector(".cards");
  if (cards) {
    cards.innerHTML = LESSONS.map((l, i) => {
      const n = String(i + 1).padStart(2, "0");
      return l.ready
        ? `<a class="card" href="lessons/${l.file}"><span class="num">Lesson ${n}</span><h3>${l.title}</h3></a>`
        : `<div class="card disabled"><span class="num">Lesson ${n}</span><h3>${l.title}</h3><p>Coming soon</p></div>`;
    }).join("");
  }
})();

// ---------- Mock spreadsheet ----------
// Usage: <div class="xl" data-sheet='{"cols":4,"rows":5,"cells":{"A1":{"v":"Beam","hdr":1}, "C2":{"v":0.15,"f":"=A2*B2"}}}'></div>
// Each cell: v = value shown, f = formula (optional), hdr = header style.
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
        typeof cell.v === "number" ? "num" : "",
        cell.hdr ? "hdr" : "",
        cell.f ? "formula" : "",
      ].join(" ").trim();
      html += `<td data-ref="${c}${r}" class="${cls}">${cell.v ?? ""}</td>`;
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
    formulaBar.textContent = cell.f ?? cell.v ?? "";
  };
  box.addEventListener("click", (e) => {
    const td = e.target.closest("td[data-ref]");
    if (td) select(td.dataset.ref);
  });
  select(sheet.start || "A1");
});
