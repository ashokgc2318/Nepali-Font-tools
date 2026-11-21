"use strict";
/* Case Expiry Monitor – web edition. Runs entirely in the browser: the Excel you upload is read here,
   nothing is sent to any server, and nothing is stored after you close the page. */

const ST = ["RED", "YELLOW", "GREEN", "OVERDUE", "CHECK"];
const ST_NAME = { RED: "Red · Urgent", YELLOW: "Yellow · Warning", GREEN: "Green · Normal", OVERDUE: "Overdue · Expired", CHECK: "Check Darta Miti" };
const ST_SHORT = { RED: "Red", YELLOW: "Yellow", GREEN: "Green", OVERDUE: "Overdue", CHECK: "Check date" };
const ST_HINT = {
  RED: "One month or less left before the one-year expiry (including the expiry day). Act on these first.",
  YELLOW: "More than one month and up to two months left. Follow up and prepare.",
  GREEN: "More than two months left. Normal monitoring.",
  OVERDUE: "The one-year period has passed.",
  CHECK: "The Darta Miti could not be read as a Bikram Sambat date. Open each case and pick the correct date.",
  ALL: "Every registered case, soonest expiry first."
};
const ST_COLOR = { RED: "#d92d20", YELLOW: "#f5a524", GREEN: "#12a150", OVERDUE: "#7a0916", CHECK: "#8a90a0" };
const ROUTES = { dashboard: "Dashboard", red: "RED", yellow: "YELLOW", green: "GREEN", overdue: "OVERDUE", check: "CHECK", all: "ALL", import: "Import", changes: "Changes", help: "Help" };
const WEEK = ["आइत", "सोम", "मंगल", "बुध", "बिहि", "शुक्र", "शनि"];
const WEEK_EN = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const S = { headers: [], dartaCol: "", cases: [], fileName: "", loaded: false, changes: [], dirty: 0, ev: [], counts: {}, q: "", sort: null, month: "", chart: {} };
const $ = s => document.querySelector(s);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const clean = v => (v === null || v === undefined) ? "" : String(v).trim();
const me = () => { try { return localStorage.getItem("cem-name") || ""; } catch (e) { return ""; } };
const today = () => BS.todayDayNum();
const xlsxHref = () => { let p = ""; try { p = localStorage.getItem("cem-pin") || ""; } catch (e) {} return "/api/export.xlsx" + (p ? "?pin=" + encodeURIComponent(p) : ""); };
const wd = n => ((n + 4) % 7 + 7) % 7;   // 0 = Sunday for an AD day number
const fmtNum = n => Number(n).toLocaleString("en-IN");

const ICON = {
  grid: '<path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"/>',
  list: '<path d="M8 6h12M8 12h12M8 18h12M4 6h.01M4 12h.01M4 18h.01"/>',
  up: '<path d="M12 19V5M5 12l7-7 7 7"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  alert: '<path d="M12 3 2 21h20L12 3z"/><path d="M12 10v5M12 18h.01"/>',
  bell: '<path d="M6 8a6 6 0 1 1 12 0c0 7 3 8 3 8H3s3-1 3-8"/><path d="M10 20a2 2 0 0 0 4 0"/>',
  check: '<path d="m5 12 5 5 9-10"/>',
  folder: '<path d="M3 6h6l2 2h10v11H3z"/>',
  file: '<path d="M6 3h8l5 5v13H6z"/><path d="M14 3v5h5"/>',
  activity: '<path d="M3 12h4l3 8 4-16 3 8h4"/>',
  help: '<circle cx="12" cy="12" r="9"/><path d="M9.5 9a2.5 2.5 0 1 1 3.5 2.3c-.7.3-1 .9-1 1.7M12 17h.01"/>',
  cal: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  users: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M18 14a6 6 0 0 1 3.5 6"/>',
  dl: '<path d="M12 4v11M7 10l5 5 5-5M4 20h16"/>',
  print: '<path d="M7 9V3h10v6M7 17H4v-7h16v7h-3M7 14h10v7H7z"/>',
  x: '<path d="M6 6l12 12M18 6 6 18"/>',
  edit: '<path d="M4 20h4L19 9l-4-4L4 16z"/>',
  trash: '<path d="M4 7h16M10 7V4h4v3M6 7l1 13h10l1-13"/>',
  hour: '<path d="M7 3h10M7 21h10M8 3c0 5 8 5 8 9s-8 4-8 9M16 3c0 5-8 5-8 9"/>'
};
const ic = (n, s = 18, w = 2) => `<svg class="ic" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[n]}</svg>`;

function toast(msg, ms = 2800) { const t = $("#toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), ms); }

// ------------------------------------------------------------- data helpers
const colSerial = () => S.headers[0];
const colName = () => S.headers.find(h => h.includes("जाहेरवाला") && h.includes("नाम")) || S.headers[2] || S.headers[1];
const colDef = () => S.headers.find(h => h.includes("प्रतिवादी") && h.includes("नाम")) || S.headers[5] || S.headers[3];
const colPhone = () => S.headers.find(h => h.includes("सम्पर्क")) || "";
const dateOf = c => clean(c.corr) || clean(c.f[S.dartaCol]);
const serialNum = c => { const n = parseFloat(clean(c.f[colSerial()]).replace(/[०-९]/g, ch => "०१२३४५६७८९".indexOf(ch))); return isNaN(n) ? 1e9 : n; };
function evaluate() {
  const t = today();
  S.ev = S.cases.map(c => ({ c, r: BS.evaluate(dateOf(c), t) }));
  S.counts = { ALL: S.ev.length }; ST.forEach(s => (S.counts[s] = 0)); S.ev.forEach(x => S.counts[x.r.status]++);
  S.within = d => S.ev.filter(x => x.r.total !== undefined && x.r.total >= 0 && x.r.total <= d).length;
}
const byExpiry = (a, b) => (a.r.expAD ?? 1e12) - (b.r.expAD ?? 1e12) || serialNum(a.c) - serialNum(b.c);
const elapsedPct = r => { if (!r.darta) return 0; const s = BS.toAD(r.darta), e = r.expAD; return Math.max(0, Math.min(100, (today() - s) / (e - s) * 100)); };
const matches = (x, q) => !q || Object.values(x.c.f).some(v => clean(v).toLowerCase().includes(q)) || (x.r.exp && BS.fmt(x.r.exp).includes(q));
const route = () => { const h = location.hash.slice(1); return ROUTES[h] ? h : "dashboard"; };

// ------------------------------------------------------------- shell
function renderShell() {
  const r = route();
  const n = S.counts;
  const item = (id, label, icon, count, hot) => `<a href="#${id}" class="${r === id ? "on" : ""}">${icon}${label}${count === undefined ? "" : `<span class="badge ${hot && count ? "hot" : ""}">${count}</span>`}</a>`;
  const dot = s => `<span class="dot" style="background:${ST_COLOR[s]}"></span>`;
  $("#nav").innerHTML = `
    <div class="navgroup">Overview</div>
    ${item("dashboard", "Dashboard", ic("grid"))}
    ${item("all", "All cases", ic("list"), n.ALL)}
    <div class="navgroup">By expiry</div>
    ${item("red", "Red · urgent", dot("RED"), n.RED, true)}
    ${item("yellow", "Yellow · warning", dot("YELLOW"), n.YELLOW)}
    ${item("green", "Green · normal", dot("GREEN"), n.GREEN)}
    ${item("overdue", "Overdue", dot("OVERDUE"), n.OVERDUE, true)}
    ${n.CHECK ? item("check", "Check date", dot("CHECK"), n.CHECK, true) : ""}
    <div class="navgroup">Excel</div>
    ${item("import", "Upload Excel", ic("file"))}
    <a href="#" data-act="download">${ic("dl")}Download Excel${S.dirty ? `<span class="badge hot">${S.dirty}</span>` : ""}</a>
    ${item("changes", "Changes this session", ic("activity"), S.changes.length || undefined)}
    ${item("help", "Help & privacy", ic("help"))}`;
  const t = BS.fromAD(today());
  $("#todayBS").textContent = `${BS.toDev(t.d)} ${BS.MONTHS_NP[t.m - 1]} ${BS.toDev(t.y)}, ${WEEK[wd(today())]}बार`;
  $("#todayAD").textContent = `${t.d} ${BS.MONTHS[t.m - 1]} ${t.y} BS · ${WEEK_EN[wd(today())]}, ${new Date(today() * 864e5).toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" })}`;
  $("#meName").textContent = S.fileName || "no file loaded";
  const st = $("#saveState");
  st.className = "live" + (S.dirty ? " off" : "");
  st.querySelector("span").textContent = S.dirty ? `${S.dirty} unsaved ${S.dirty === 1 ? "change" : "changes"} – download the Excel` : (S.loaded ? "Everything downloaded" : "Upload an Excel to begin");
}
function render() {
  evaluate(); renderShell();
  const r = route(), el = $("#content");
  Object.values(S.chart).forEach(c => c.destroy()); S.chart = {};
  if (S.q && !["all"].includes(r) && r !== "dashboard" && !ROUTES[r].match(/^[A-Z]+$/)) S.q = "";
  if (!S.loaded && r !== "help") { el.innerHTML = startView(); wireDrop(); return; }
  if (r === "dashboard") el.innerHTML = S.q ? listView("ALL") : dashView();
  else if (["red", "yellow", "green", "overdue", "check", "all"].includes(r)) el.innerHTML = listView(ROUTES[r]);
  else if (r === "import") el.innerHTML = importView();
  else if (r === "changes") el.innerHTML = changesView();
  else if (r === "help") el.innerHTML = helpView();
  if (r === "dashboard" && !S.q) drawCharts();
  if (r === "import") wireDrop();
}

// ------------------------------------------------------------- dashboard
function kpi(cls, href, label, val, icon, pct, foot) {
  return `<a class="kpi ${cls}" href="#${href}"><div class="top"><span class="lbl">${label}</span><span class="icon">${ic(icon, 18)}</span></div>
    <div class="val">${fmtNum(val)}</div><div class="meter"><i style="width:${Math.max(val ? 3 : 0, pct)}%"></i></div><div class="foot">${foot}</div></a>`;
}
function ring(x) {
  const r = x.r, total = r.total, col = ST_COLOR[r.status], C = 2 * Math.PI * 21;
  const frac = Math.max(0.04, Math.min(1, total / 365));
  return `<div class="ring"><svg width="56" height="56" viewBox="0 0 50 50"><circle cx="25" cy="25" r="21" fill="none" stroke="#eef0f4" stroke-width="5"/>
    <circle cx="25" cy="25" r="21" fill="none" stroke="${col}" stroke-width="5" stroke-linecap="round" stroke-dasharray="${C * frac} ${C}"/></svg>
    <b>${total}<small>${total === 1 ? "DAY" : "DAYS"}</small></b></div>`;
}
function dashView() {
  const n = S.counts, all = n.ALL || 1;
  if (!S.cases.length) return `<section class="panel"><div class="empty">${ic("folder", 48, 1.5)}<h2>No cases yet</h2><p>Register the first case with “New case”.</p></div></section>`;
  const red30 = S.within(30);
  const urgent = n.RED + n.OVERDUE;
  const ribbon = urgent
    ? `<div class="ribbon"><span class="pulse">${ic("bell", 20, 2.2)}</span><p><b>${n.RED} ${n.RED === 1 ? "case expires" : "cases expire"} within a month</b>${n.OVERDUE ? ` and <b>${n.OVERDUE} ${n.OVERDUE === 1 ? "is" : "are"} already overdue</b>` : ""}. <a href="#red">Open the Red list</a>${n.OVERDUE ? ` · <a href="#overdue">Overdue list</a>` : ""}</p></div>`
    : `<div class="ribbon calm"><span class="pulse">${ic("check", 20, 2.6)}</span><p><b>No case expires within the next month.</b> The next expiry is ${nextExpiryText()}.</p></div>`;
  const t = BS.fromAD(today());
  const next = S.ev.filter(x => ["RED", "YELLOW", "GREEN"].includes(x.r.status)).sort(byExpiry).slice(0, 8);
  return `${ribbon}
  <section class="kpis">
    ${kpi("k-all", "all", "Total registered", n.ALL, "folder", 100, `${fmtNum(regThisMonth())} registered in ${BS.MONTHS[t.m - 1]}`)}
    ${kpi("k-red", "red", "Red · ≤ 1 month", n.RED, "alert", n.RED / all * 100, "Immediate attention")}
    ${kpi("k-yel", "yellow", "Yellow · 1–2 months", n.YELLOW, "clock", n.YELLOW / all * 100, "Follow up")}
    ${kpi("k-grn", "green", "Green · > 2 months", n.GREEN, "check", n.GREEN / all * 100, "Normal monitoring")}
    ${kpi("k-ovd", "overdue", "Overdue", n.OVERDUE, "hour", n.OVERDUE / all * 100, "One year has passed")}
    ${kpi("k-30", "red", "Expiring ≤ 30 days", red30, "bell", red30 / all * 100, "Next 30 calendar days")}
    ${kpi("k-60", "yellow", "Expiring ≤ 60 days", S.within(60), "cal", S.within(60) / all * 100, "Next 60 calendar days")}
    ${n.CHECK ? kpi("k-chk", "check", "Darta Miti to check", n.CHECK, "edit", n.CHECK / all * 100, "Date could not be read") : ""}
  </section>
  <section class="grid">
    <div class="panel span-5">
      <div class="panel-h"><h2>Status of all cases</h2><span class="sub">tap a row to open the list</span></div>
      <div class="donutwrap"><div class="chartbox"><canvas id="cDonut" aria-label="Cases by status"></canvas><div class="donut-center"><div><b>${fmtNum(n.ALL)}</b><span>cases</span></div></div></div>
      <div class="legend">${ST.filter(s => s !== "CHECK" || n.CHECK).map(s => `<a href="#${s.toLowerCase()}"><i style="background:${ST_COLOR[s]}"></i><span>${ST_SHORT[s]}</span><b>${n[s]}</b><em>${(n[s] / all * 100).toFixed(1)}%</em></a>`).join("")}</div></div>
    </div>
    <div class="panel span-7">
      <div class="panel-h"><h2>Expiries in the next 12 BS months</h2><span class="sub">cases reaching their one-year date each month</span></div>
      <div class="chartbox"><canvas id="cMonths" aria-label="Expiries per BS month"></canvas></div>
    </div>
    <div class="panel span-12">
      <div class="panel-h"><h2>Next 60 days at a glance</h2><span class="sub">each block is one day · darker = more cases expiring that day</span></div>
      ${timeline()}
    </div>
    <div class="panel span-7">
      <div class="panel-h"><h2>Next cases to expire</h2><a class="sub" href="#all">see all →</a></div>
      <div class="nextlist">${next.map(x => `<div class="nx" data-id="${esc(x.c.id)}">${ring(x)}
        <div><div class="who">${esc(x.c.f[colName()]) || "–"}</div><div class="meta">${esc(colSerial())} ${esc(x.c.f[colSerial()])} · expires ${BS.fmt(x.r.exp)} · ${esc(x.r.label)}</div></div>
        <span class="pill p-${x.r.status}">${ST_SHORT[x.r.status]}</span></div>`).join("") || `<div class="empty">No upcoming expiries.</div>`}</div>
    </div>
    <div class="panel span-5">
      <div class="panel-h"><h2>New registrations</h2><span class="sub">cases by Darta Miti month</span></div>
      <div class="chartbox" style="height:210px"><canvas id="cIntake" aria-label="Registrations per month"></canvas></div>
      <div class="panel-h" style="margin:18px 0 6px"><h2>Changes this session</h2><a class="sub" href="#changes">all →</a></div>
      <div class="feed">${feedHTML(5)}</div>
    </div>
  </section>`;
}
function nextExpiryText() {
  const x = S.ev.filter(x => x.r.total >= 0).sort(byExpiry)[0];
  return x ? `${BS.fmt(x.r.exp)} (${x.r.label.replace(" Remaining", " from today")})` : "not scheduled";
}
function regThisMonth() { const t = BS.fromAD(today()); return S.ev.filter(x => x.r.darta && x.r.darta.y === t.y && x.r.darta.m === t.m).length; }
function timeline() {
  const t0 = today(), days = [];
  for (let i = 0; i < 60; i++) days.push({ n: t0 + i, c: 0 });
  S.ev.forEach(x => { if (x.r.total >= 0 && x.r.total < 60) days[x.r.total].c++; });
  const lvl = c => (c === 0 ? "" : c === 1 ? "n1" : c <= 3 ? "n2" : c <= 6 ? "n3" : "n4");
  const cells = days.map((d, i) => { const b = BS.fromAD(d.n);
    return `<div class="d ${lvl(d.c)} ${b.d === 1 && i ? "mark" : ""}" title="${BS.fmt(b)} (${b.d} ${BS.MONTHS[b.m - 1]}) · ${d.c} ${d.c === 1 ? "case" : "cases"}"></div>`; }).join("");
  const scale = days.map((d, i) => { const b = BS.fromAD(d.n);
    return `<span>${i === 0 ? "Today" : b.d === 1 ? BS.MONTHS[b.m - 1] + " 1" : ""}</span>`; }).join("");
  const tot = days.reduce((a, d) => a + d.c, 0);
  return `<div class="timeline" role="img" aria-label="${tot} cases expire in the next 60 days">${cells}</div><div class="tl-scale">${scale}</div>
    <div class="tl-key"><span><i style="background:var(--gray-soft)"></i>none</span><span><i style="background:#fde2a7"></i>1</span><span><i style="background:#f8b24a"></i>2–3</span><span><i style="background:#ef6c2f"></i>4–6</span><span><i style="background:#c81e1e"></i>7+</span>
    <span style="margin-left:auto"><b>${tot}</b> ${tot === 1 ? "case expires" : "cases expire"} in the next 60 days</span></div>`;
}
function drawCharts() {
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
  Chart.defaults.color = "#667089";
  const n = S.counts, shown = ST.filter(s => s !== "CHECK" || n.CHECK);
  S.chart.donut = new Chart($("#cDonut"), { type: "doughnut",
    data: { labels: shown.map(s => ST_SHORT[s]), datasets: [{ data: shown.map(s => n[s]), backgroundColor: shown.map(s => ST_COLOR[s]), borderWidth: 3, borderColor: "#fff", hoverOffset: 6 }] },
    options: { cutout: "72%", maintainAspectRatio: false, plugins: { legend: { display: false } }, onClick: (e, els) => { if (els[0]) location.hash = shown[els[0].index].toLowerCase(); } } });
  // next 12 BS months, stacked by the colour each case will be at expiry time? keep simple: one colour, current month highlighted
  const t = BS.fromAD(today()), months = [];
  for (let i = 0; i < 12; i++) { const k = t.m - 1 + i; months.push({ y: t.y + Math.floor(k / 12), m: k % 12 + 1, n: 0 }); }
  S.ev.forEach(x => { if (x.r.exp && x.r.total >= 0) { const o = months.find(o => o.y === x.r.exp.y && o.m === x.r.exp.m); if (o) o.n++; } });
  const ctx = $("#cMonths").getContext("2d"); const g = ctx.createLinearGradient(0, 0, 0, 250); g.addColorStop(0, "#3b5bdb"); g.addColorStop(1, "#8da2f0");
  S.chart.months = new Chart(ctx, { type: "bar",
    data: { labels: months.map(o => [BS.MONTHS[o.m - 1], String(o.y)]), datasets: [{ data: months.map(o => o.n), backgroundColor: months.map((o, i) => i === 0 ? ST_COLOR.RED : i === 1 ? ST_COLOR.YELLOW : g), borderRadius: 7, maxBarThickness: 38 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.raw} ${c.raw === 1 ? "case" : "cases"} expire` } } },
      scales: { x: { grid: { display: false }, ticks: { font: { size: 11 } } }, y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#eef0f4" }, border: { display: false } } } },
    plugins: [{ id: "vals", afterDatasetsDraw(ch) { const c = ch.ctx; c.save(); c.font = "600 11px " + Chart.defaults.font.family; c.fillStyle = "#18213a"; c.textAlign = "center";
      ch.getDatasetMeta(0).data.forEach((b, i) => { const v = months[i].n; if (v) c.fillText(v, b.x, b.y - 6); }); c.restore(); } }] });
  // intake: registrations per BS month for the last 8 months
  const inc = [];
  for (let i = 7; i >= 0; i--) { let k = t.m - 1 - i, y = t.y; while (k < 0) { k += 12; y--; } inc.push({ y, m: k + 1, n: 0 }); }
  S.ev.forEach(x => { if (x.r.darta) { const o = inc.find(o => o.y === x.r.darta.y && o.m === x.r.darta.m); if (o) o.n++; } });
  const ctx2 = $("#cIntake").getContext("2d"); const g2 = ctx2.createLinearGradient(0, 0, 0, 210); g2.addColorStop(0, "rgba(37,70,184,.28)"); g2.addColorStop(1, "rgba(37,70,184,0)");
  S.chart.intake = new Chart(ctx2, { type: "line",
    data: { labels: inc.map(o => BS.MONTHS[o.m - 1].slice(0, 3) + " " + String(o.y).slice(2)), datasets: [{ data: inc.map(o => o.n), borderColor: "#2546b8", backgroundColor: g2, fill: true, tension: .35, pointRadius: inc.map((o, i) => i === inc.length - 1 ? 5 : 3), pointBackgroundColor: "#2546b8" }] },
    options: { maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.raw} registered` } } },
      scales: { x: { grid: { display: false }, ticks: { font: { size: 11 } } }, y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#eef0f4" }, border: { display: false } } } } });
}
const avColor = s => { let h = 0; for (const ch of String(s)) h = (h * 31 + ch.charCodeAt(0)) % 360; return `hsl(${h} 45% 42%)`; };
// ------------------------------------------------------------- list views
function listView(status) {
  const q = S.q.trim().toLowerCase();
  let list = status === "ALL" ? S.ev : S.ev.filter(x => x.r.status === status);
  const months = [...new Set(list.filter(x => x.r.exp).map(x => x.r.exp.y * 100 + x.r.exp.m))].sort();
  if (S.month && !months.includes(+S.month)) S.month = "";
  if (S.month) list = list.filter(x => x.r.exp && x.r.exp.y * 100 + x.r.exp.m === +S.month);
  list = list.filter(x => matches(x, q)).sort(status === "CHECK" ? (a, b) => serialNum(a.c) - serialNum(b.c) : byExpiry);
  if (S.sort === "serial") list.sort((a, b) => serialNum(a.c) - serialNum(b.c));
  const ov = status === "OVERDUE";
  const color = status === "ALL" ? "#2546b8" : ST_COLOR[status];
  const title = S.q && route() === "dashboard" ? `Search results for “${esc(S.q)}”` : status === "ALL" ? "All cases" : ST_NAME[status];
  const rows = list.map(x => {
    const r = x.r, ok = r.status !== "CHECK", pct = elapsedPct(r);
    return `<tr class="s-${r.status}" data-id="${esc(x.c.id)}" tabindex="0">
      <td class="bar"></td>
      <td class="num"><b>${esc(x.c.f[colSerial()])}</b></td>
      <td class="who2"><b>${esc(x.c.f[colName()]) || "–"}</b><span>${esc(x.c.f[colDef()])}</span></td>
      <td class="date">${ok ? `<b>${BS.fmt(r.darta)}</b><span>Darta Miti</span>` : `<b style="color:var(--red-ink)">${esc(dateOf(x.c)) || "(blank)"}</b><span>cannot be read</span>`}</td>
      <td class="date">${ok ? `<b>${BS.fmt(r.exp)}</b><span>${BS.adStr(r.expAD)} AD</span>` : "–"}</td>
      <td><span class="remain" style="color:${ok ? (r.status === "GREEN" ? "var(--ink)" : ST_COLOR[r.status]) : "var(--muted)"}">${esc(r.label)}</span>
        ${ok ? `<div class="prog" title="${pct.toFixed(0)}% of the one-year period used"><i style="width:${pct}%;background:${ST_COLOR[r.status]}"></i></div>` : ""}</td>
      <td class="num">${ok ? (ov ? `${-r.total} overdue` : `${r.total} days`) : "–"}</td>
      <td><span class="pill p-${r.status}">${ST_SHORT[r.status]}</span></td></tr>`;
  }).join("");
  const mchips = status !== "CHECK" && months.length > 1 ? `<div class="filters"><button class="fchip ${S.month ? "" : "on"}" data-month="">All months</button>${months.map(k =>
    `<button class="fchip ${+S.month === k ? "on" : ""}" data-month="${k}">${BS.MONTHS[k % 100 - 1]} ${Math.floor(k / 100)}</button>`).join("")}</div>` : "";
  return `<div class="listhead"><div class="title"><span class="chip" style="background:${color}1a;color:${color}">${ic(status === "ALL" ? "list" : status === "RED" ? "alert" : status === "YELLOW" ? "clock" : status === "GREEN" ? "check" : status === "OVERDUE" ? "hour" : "edit", 22)}</span>
    <div><h1>${title} <span class="sub" style="font-family:var(--body);font-weight:500;color:var(--muted);font-size:1rem">· ${list.length}</span></h1><p>${ST_HINT[status]}</p></div></div>
    <div class="tools"><button class="btn ghost" data-act="print">${ic("print", 16)}Print</button><button class="btn" data-act="download">${ic("dl", 16)}Download Excel</button></div></div>
    ${mchips}
    ${list.length ? `<div class="tablewrap"><table class="list"><thead><tr><th></th><th><button data-act="sort-serial">${esc(colSerial())}</button></th><th>जाहेरवाला / प्रतिवादी</th><th>Darta Miti (BS)</th><th><button data-act="sort-exp">Expiry</button></th><th>Remaining</th><th>${ov ? "Overdue" : "Total"}</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></div>`
      : `<div class="panel"><div class="empty">${ic("check", 44, 1.6)}<p>${q ? "No cases match your search." : "No cases in this list right now."}</p></div></div>`}`;
}

// ------------------------------------------------------------- drawer (case detail)
function openDrawer(id) {
  const c = S.cases.find(x => x.id === id); if (!c) return;
  const r = BS.evaluate(dateOf(c), today()), ok = r.status !== "CHECK", pct = elapsedPct(r);
  $("#drawer").innerHTML = `<div class="mh"><h2>${esc(colSerial())} ${esc(c.f[colSerial()])} · ${esc(c.f[colName()])}</h2><button class="x" data-act="close" aria-label="Close">${ic("x", 16)}</button></div>
  <div class="mb">
    <div class="hero-status ${r.status}"><span style="opacity:.9;font-weight:600">${ST_NAME[r.status]}</span><b>${esc(r.label)}</b>
      ${ok ? `<div class="track"><div class="rail"><i style="width:${pct}%"></i></div><div class="ends"><span>Darta ${BS.fmt(r.darta)}</span><span>${pct.toFixed(0)}% of the year used</span><span>Expiry ${BS.fmt(r.exp)}</span></div></div>
      <span style="font-size:.86rem;opacity:.9">Expiry: ${r.exp.d} ${BS.MONTHS[r.exp.m - 1]} ${r.exp.y} BS · ${WEEK_EN[wd(r.expAD)]}, ${BS.adStr(r.expAD)} AD</span>` : `<span>Edit the case and pick the correct Darta Miti.</span>`}</div>
    <dl class="kv">${S.headers.map(h => `<dt>${esc(h)}</dt><dd>${esc(c.f[h]) || '<span class="stamp">–</span>'}</dd>`).join("")}</dl>
    ${c.createdAt ? `<div class="stamp">Added in this session at ${esc(c.createdAt)}</div>` : c.updatedAt ? `<div class="stamp">Edited in this session at ${esc(c.updatedAt)}</div>` : ""}
  </div>
  <div class="mf"><button class="btn danger" data-act="del" data-id="${esc(c.id)}">${ic("trash", 16)}Delete</button><button class="btn primary" data-act="edit" data-id="${esc(c.id)}">${ic("edit", 16)}Edit case</button></div>`;
  $("#drawer").classList.add("on"); $("#scrim").classList.add("on");
}
function closeAll() { $("#drawer").classList.remove("on"); $("#modal").classList.remove("on"); $("#scrim").classList.remove("on"); }

// ------------------------------------------------------------- case form with BS calendar
function openForm(id) {
  const c = id ? S.cases.find(x => x.id === id) : null;
  const t = BS.fromAD(today());
  const maxSerial = Math.max(0, ...S.cases.map(serialNum).filter(n => n < 1e9));
  const val = h => c ? clean(c.f[h]) : (h === colSerial() ? String(maxSerial + 1) : "");
  const sel = (c && BS.parse(dateOf(c))) || t;
  const F = { y: sel.y, m: sel.m, d: sel.d, raw: c ? dateOf(c) : BS.fmt(t) };
  const fields = S.headers.filter(h => h !== S.dartaCol).map((h, i) => `<div class="field ${h.length > 26 ? "wide" : ""}"><label for="f${i}">${esc(h)}</label><input id="f${i}" data-h="${esc(h)}" value="${esc(val(h))}" autocomplete="off"></div>`).join("");
  $("#modal").innerHTML = `<form id="cf" novalidate style="display:contents">
    <div class="mh"><h2>${c ? `Edit case ${esc(c.f[colSerial()])}` : "Register a new case"}</h2><button type="button" class="x" data-act="close" aria-label="Close">${ic("x", 16)}</button></div>
    <div class="mb">
      <div class="field wide"><label>${esc(S.dartaCol)} – Darta Miti (BS)</label>
        <div class="datebox"><div class="bscal" id="bscal"></div>
          <div style="display:grid;gap:10px;align-content:start">
            <div class="field"><label for="draw">Or type the date</label><input id="draw" value="${esc(F.raw)}" placeholder="२०८३।०८।०४ or 2083-08-04"></div>
            <div class="preview" id="pv" aria-live="polite"></div>
          </div></div></div>
      <div class="form">${fields}</div>
      <div id="formMsg"></div>
    </div>
    <div class="mf"><span class="stamp" style="align-self:center">Kept in this browser until you close the page – download the Excel to keep it</span><div style="display:flex;gap:10px"><button type="button" class="btn" data-act="close">Cancel</button><button type="submit" class="btn primary" id="saveBtn">${c ? "Save changes" : "Register case"}</button></div></div>
  </form>`;
  $("#modal").classList.add("on"); $("#scrim").classList.add("on");
  const cal = $("#bscal"), raw = $("#draw");
  const drawCal = () => {
    const first = BS.toAD({ y: F.y, m: F.m, d: 1 }), n = BS.dim(F.y, F.m), off = wd(first);
    const years = []; for (let y = t.y - 8; y <= t.y + 1; y++) years.push(y);
    if (!years.includes(F.y)) years.unshift(F.y);
    let cells = ""; for (let i = 0; i < off; i++) cells += "<span></span>";
    for (let d = 1; d <= n; d++) {
      const ad = first + d - 1, isT = F.y === t.y && F.m === t.m && d === t.d;
      cells += `<button type="button" data-day="${d}" class="${d === F.d ? "sel" : ""} ${isT ? "today" : ""} ${wd(ad) === 6 ? "sat" : ""}">${BS.toDev(d)}<small>${new Date(ad * 864e5).getUTCDate()}</small></button>`;
    }
    cal.innerHTML = `<div class="hd"><button type="button" data-nav="-1" aria-label="Previous month">‹</button>
      <select id="cm" aria-label="Month">${BS.MONTHS_NP.map((m, i) => `<option value="${i + 1}" ${i + 1 === F.m ? "selected" : ""}>${m}</option>`).join("")}</select>
      <select id="cy" aria-label="Year">${years.map(y => `<option ${y === F.y ? "selected" : ""} value="${y}">${BS.toDev(y)}</option>`).join("")}</select>
      <button type="button" data-nav="1" aria-label="Next month">›</button></div>
      <div class="wk">${WEEK.map(w => `<span>${w}</span>`).join("")}</div><div class="days">${cells}</div>`;
    $("#cm").onchange = e => { F.m = +e.target.value; F.d = Math.min(F.d, BS.dim(F.y, F.m)); pick(); };
    $("#cy").onchange = e => { F.y = +e.target.value; F.d = Math.min(F.d, BS.dim(F.y, F.m)); pick(); };
  };
  const preview = () => {
    const r = BS.evaluate(raw.value, today()), pv = $("#pv");
    pv.className = "preview " + r.status;
    pv.innerHTML = r.status === "CHECK" ? `<span class="big">Not a valid BS date yet</span><span>Pick a day on the calendar.</span>` :
      `<span class="stamp">One-year expiry</span><span class="big">${BS.fmt(r.exp)} · ${r.exp.d} ${BS.MONTHS[r.exp.m - 1]} ${r.exp.y}</span>
       <span>${WEEK_EN[wd(r.expAD)]}, ${BS.adStr(r.expAD)} AD</span><span><span class="pill p-${r.status}">${ST_SHORT[r.status]}</span> &nbsp;<b>${esc(r.label)}</b></span>`;
  };
  const pick = () => { raw.value = BS.fmt(F); drawCal(); preview(); };
  cal.addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    if (b.dataset.day) { F.d = +b.dataset.day; pick(); }
    if (b.dataset.nav) { let m = F.m + +b.dataset.nav, y = F.y; if (m < 1) { m = 12; y--; } if (m > 12) { m = 1; y++; } if (y < BS.Y0 || y > BS.Y1 - 1) return; F.y = y; F.m = m; F.d = Math.min(F.d, BS.dim(y, m)); pick(); }
  });
  raw.oninput = () => { const b = BS.parse(raw.value); if (b) { Object.assign(F, b); drawCal(); } preview(); };
  drawCal(); preview();
  setTimeout(() => { const f = $(`#cf input[data-h]:not([data-h="${CSS.escape(colSerial())}"])`); f && f.focus(); }, 60);
  $("#cf").onsubmit = async e => {
    e.preventDefault(); await submitForm(c, false);
  };
}

const nowStr = () => { const d = new Date(), p = n => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
function logChange(action, summary) { S.changes.unshift({ at: nowStr(), action, summary }); S.dirty++; }
const describe = f => `${colSerial()} ${f[colSerial()] || ""} · ${f[colName()] || ""}`;

async function submitForm(c, force) {
  const f = {}; document.querySelectorAll("#cf input[data-h]").forEach(i => (f[i.dataset.h] = i.value.trim()));
  f[S.dartaCol] = $("#draw").value.trim();
  const msg = $("#formMsg"); msg.innerHTML = "";
  if (!BS.parse(f[S.dartaCol])) { msg.innerHTML = `<div class="msg">Pick a valid Darta Miti on the calendar before saving.</div>`; return; }
  const ser = colSerial(), others = S.cases.filter(x => x !== c);
  if (!f[ser]) f[ser] = String(Math.max(0, ...S.cases.map(serialNum).filter(n => n < 1e9)) + 1);
  if (others.some(x => clean(x.f[ser]) === f[ser])) { msg.innerHTML = `<div class="msg">${esc(ser)} ${esc(f[ser])} is already used by another case. Use a different number.</div>`; return; }
  if (!force) {
    const p = BS.parse(f[S.dartaCol]), nm = colName();
    const dup = f[nm] && others.find(x => clean(x.f[nm]) === f[nm] && JSON.stringify(BS.parse(dateOf(x))) === JSON.stringify(p));
    if (dup) {
      msg.innerHTML = `<div class="msg">A case for ${esc(f[nm])} with the same Darta Miti already exists (${esc(ser)} ${esc(dup.f[ser])}). <button type="button" class="btn" id="forceBtn" style="margin-left:8px">Save anyway</button></div>`;
      $("#forceBtn").onclick = () => submitForm(c, true); return;
    }
  }
  if (c) { const changed = Object.keys(f).filter(h => clean(c.f[h]) !== f[h]); c.f = { ...c.f, ...f }; c.corr = ""; c.updatedAt = nowStr(); logChange("edit", describe(c.f) + (changed.length ? " – changed: " + changed.join(", ") : "")); }
  else { S.cases.push({ id: uid(), f, corr: "", createdAt: nowStr() }); logChange("add", describe(f) + " – Darta Miti " + f[S.dartaCol]); }
  closeAll(); render();
  const r = BS.evaluate(f[S.dartaCol], today());
  toast(c ? "Changes saved here. Download the Excel to keep them." : `Case added · ${ST_SHORT[r.status]} list · expires ${BS.fmt(r.exp)}`, 4000);
}
function confirmDelete(id) {
  const c = S.cases.find(x => x.id === id); if (!c) return;
  $("#drawer").classList.remove("on");
  showModal(`<div class="mh"><h2>Delete case ${esc(c.f[colSerial()])}?</h2></div>
    <div class="mb"><p style="margin:0">${esc(c.f[colName()])} · Darta Miti ${esc(dateOf(c))}. It is removed from every list and will not be in the Excel you download.</p></div>
    <div class="mf"><button class="btn" data-act="close">Keep it</button><button class="btn danger" data-act="del-yes" data-id="${esc(id)}">${ic("trash", 16)}Delete case</button></div>`);
}
function doDelete(id) {
  const c = S.cases.find(x => x.id === id); if (!c) return;
  S.cases = S.cases.filter(x => x.id !== id); logChange("delete", describe(c.f)); closeAll(); render(); toast("Case deleted.");
}

// ------------------------------------------------------------- reading Excel (in the browser)
const DEFAULT_HEADERS = ["सि.न.", "मुद्धा शिर्षक", "जाहेरवालाको नामथर", "जाहेरवालाको वतन", "सम्पर्क नम्वर", "प्रतिवादी नामथर",
  "प्रतिवादीको वतन", "सम्पर्क नं.", "लिङ्ग", "जाहेरीको रकम", "जाहेरीको मिति", "दर्ता मिती", "प्रहरी कारबाही", "मुद्दा दर्ता मिति"];
const COMPUTED = new Set(["Darta Miti (BS)", "Expiry Date (BS)", "Expiry Date (AD)", "Remaining Months", "Remaining Days", "Total Remaining Days",
  "Remaining Time", "Expiry Status", "Days Overdue", "Overdue Status", "#", "Status", "Entered By", "Entered On", "Last Edited By", "Last Edited On"]);
const CORR = "Darta Miti Correction (optional)";
function cellText(cell) {
  if (!cell || cell.v === undefined || cell.v === null) return "";
  if (cell.t === "e") return "";                                     // #REF!, #N/A …
  if (cell.t === "n") return Number.isInteger(cell.v) ? String(cell.v) : String(+cell.v.toFixed(6));
  if (cell.t === "d") return cell.v.toISOString().slice(0, 10);
  if (cell.t === "b") return cell.v ? "TRUE" : "FALSE";
  return String(cell.v).trim();
}
function guessDarta(hs) {
  const t = hs.map(h => h.replace(/\s+/g, " ").trim());
  let i = t.findIndex(h => h === "दर्ता मिती" || h === "दर्ता मिति" || /^darta\s*miti$/i.test(h));
  if (i < 0) i = t.findIndex(h => h.includes("दर्ता") && (h.includes("मिति") || h.includes("मिती")) && !h.includes("मुद्द") && !h.includes("मुद्ध"));
  if (i < 0) i = t.findIndex(h => /darta|registration/i.test(h));
  return i < 0 ? hs[0] : hs[i];
}
function parseWorkbook(wb) {
  const found = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]; if (!ws["!ref"]) continue;
    const rg = XLSX.utils.decode_range(ws["!ref"]);
    const row = r => { const out = []; for (let c = rg.s.c; c <= rg.e.c; c++) out.push(cellText(ws[XLSX.utils.encode_cell({ r, c })])); return out; };
    let hi = -1;
    for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 25); r++) {
      const short = row(r).filter(v => v.length > 0 && v.length <= 32);
      if (short.some(v => v.includes("दर्ता") && (v.includes("मिति") || v.includes("मिती"))) || short.some(v => /darta/i.test(v))) { hi = r; break; }
    }
    if (hi < 0) continue;
    const cols = [], seen = {};
    row(hi).forEach((h0, j) => {
      const h = h0.replace(/\s+/g, " ").trim();
      if (!h || COMPUTED.has(h) || h.startsWith("h_")) return;
      if (h === CORR) { cols.push({ j, h, corr: true }); return; }
      let n = h; if (seen[h]) n = `${h} (${++seen[h]})`; else seen[h] = 1;
      cols.push({ j, h: n });
    });
    const fields = cols.filter(c => !c.corr), data = [];
    let blankRun = 0;
    for (let r = hi + 1; r <= rg.e.r; r++) {
      const v = row(r), f = {}; let corr = "";
      for (const c of cols) { if (c.corr) corr = v[c.j]; else f[c.h] = v[c.j]; }
      if (!fields.slice(1).some(c => f[c.h])) { if (++blankRun > 2000) break; continue; }
      blankRun = 0; data.push({ f, corr });
    }
    found.push({ sheet: name, headers: fields.map(c => c.h), rows: data });
  }
  found.sort((a, b) => (b.sheet === "MASTER DATA") - (a.sheet === "MASTER DATA") || b.rows.length - a.rows.length);
  return found[0] || null;
}
function startView() {
  return `<section class="panel" style="padding:34px">
    <div style="display:grid;gap:22px;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));align-items:center">
      <div style="display:grid;gap:12px">
        <h1 style="font-size:1.8rem">Track every case's one-year expiry in Bikram Sambat</h1>
        <p style="margin:0;color:var(--muted);max-width:56ch">Upload your case register (.xlsx). Each case gets its exact BS expiry date and is sorted into Red, Yellow, Green or Overdue.
        Add or edit cases here, then download the updated master Excel.</p>
        <ul class="steps" style="color:var(--muted)"><li>Your file is read inside this browser. It is never uploaded to any server.</li><li>Nothing is saved when you close the page. Always download the Excel after changes.</li></ul>
        <div><button class="btn ghost" data-act="blank">or start an empty register</button></div>
      </div>
      <div class="drop" id="drop">${ic("file", 46, 1.4)}<b style="font-size:1.1rem">Drop your Excel file here</b><span class="stamp">.xlsx or .xls · the sheet with the दर्ता मिती column is found automatically</span>
        <button class="btn primary" data-act="pick">${ic("up", 16)}Choose Excel file</button><input type="file" id="fileIn" accept=".xlsx,.xlsm,.xls" hidden></div>
    </div></section>`;
}
function importView() {
  return `<div class="listhead"><div class="title"><span class="chip" style="background:#2546b81a;color:#2546b8">${ic("file", 22)}</span><div><h1>Upload Excel</h1><p>Load another register, or merge a colleague's file into the one open now.</p></div></div></div>
  <section class="grid">
    <div class="panel span-6"><div class="drop" id="drop">${ic("file", 40, 1.5)}<b>Drop an .xlsx file here</b><span class="stamp">Read in this browser only.</span>
      <button class="btn primary" data-act="pick">Choose file</button><input type="file" id="fileIn" accept=".xlsx,.xlsm,.xls" hidden></div></div>
    <div class="panel span-6"><div class="panel-h"><h2>Open now</h2></div>
      <p style="margin:0 0 6px"><b>${esc(S.fileName || "Empty register")}</b> · ${S.cases.length} cases</p>
      <p class="stamp" style="margin:0 0 14px">${S.dirty ? `${S.dirty} unsaved ${S.dirty === 1 ? "change" : "changes"} since it was opened.` : "No unsaved changes."}</p>
      <button class="btn primary" data-act="download">${ic("dl", 16)}Download updated Excel</button></div>
  </section>`;
}
function wireDrop() {
  const d = $("#drop"); if (!d) return;
  d.ondragover = e => { e.preventDefault(); d.classList.add("over"); };
  d.ondragleave = () => d.classList.remove("over");
  d.ondrop = e => { e.preventDefault(); d.classList.remove("over"); if (e.dataTransfer.files[0]) importFile(e.dataTransfer.files[0]); };
  $("#fileIn").onchange = e => { const f = e.target.files[0]; e.target.value = ""; if (f) importFile(f); };
}
async function importFile(file) {
  showModal(`<div class="mh"><h2>Reading ${esc(file.name)}…</h2></div><div class="mb"><p>Please wait.</p></div>`);
  let p;
  try { p = parseWorkbook(XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true })); }
  catch (e) { showModal(`<div class="mh"><h2>Could not open the file</h2><button class="x" data-act="close">${ic("x", 16)}</button></div><div class="mb"><div class="msg">${esc(file.name)} could not be read as an Excel workbook. Save it as .xlsx and try again.</div></div>`); return; }
  if (!p || !p.rows.length) { showModal(`<div class="mh"><h2>No cases found</h2><button class="x" data-act="close">${ic("x", 16)}</button></div><div class="mb"><div class="msg">No sheet in ${esc(file.name)} has a header row with a दर्ता मिती (Darta Miti) column and case rows below it.</div></div>`); return; }
  const has = S.loaded && S.cases.length, ser = p.headers[0];
  const keys = new Set(S.cases.map(c => clean(c.f[colSerial()])));
  const upd = has ? p.rows.filter(r => clean(r.f[ser]) && keys.has(clean(r.f[ser]))).length : 0;
  const draw = darta => {
    const ok = p.rows.filter(r => BS.parse(clean(r.corr) || r.f[darta])).length;
    showModal(`<div class="mh"><h2>Load ${p.rows.length} cases</h2><button class="x" data-act="close">${ic("x", 16)}</button></div>
    <div class="mb"><p style="margin:0">Sheet <b>${esc(p.sheet)}</b> of ${esc(file.name)} · ${p.headers.length} columns.</p>
      <div class="field"><label for="impD">Column with the Darta Miti (BS)</label><select id="impD">${p.headers.map(h => `<option ${h === darta ? "selected" : ""}>${esc(h)}</option>`).join("")}</select></div>
      <div class="msg ${ok === p.rows.length ? "info" : ""}">${ok} of ${p.rows.length} dates read as valid BS dates.${ok < p.rows.length ? ` The other ${p.rows.length - ok} will appear under “Check date” so you can correct them.` : ""}</div>
      ${has ? `<div class="choice"><label><input type="radio" name="mode" value="merge" checked><span><b>Merge into the open register</b> – match by ${esc(ser)}: ${upd} updated, ${p.rows.length - upd} added.</span></label>
      <label><input type="radio" name="mode" value="replace"><span><b>Open this file instead</b> – close the ${S.cases.length} cases open now${S.dirty ? ` (${S.dirty} unsaved changes will be lost)` : ""}.</span></label></div>` : ""}
    </div><div class="mf"><button class="btn" data-act="close">Cancel</button><button class="btn primary" id="impGo">Load cases</button></div>`);
    $("#impD").onchange = e => draw(e.target.value);
    $("#impGo").onclick = () => {
      const mode = (document.querySelector("input[name=mode]:checked") || { value: "replace" }).value;
      if (mode === "replace") {
        Object.assign(S, { headers: p.headers.slice(), dartaCol: darta, fileName: file.name, loaded: true, changes: [], dirty: 0 });
        S.cases = p.rows.map((r, i) => ({ id: uid() + i, f: r.f, corr: r.corr }));
        toast(`${S.cases.length} cases loaded from ${file.name}.`, 3500);
      } else {
        const map = h => (h === darta ? S.dartaCol : h);
        for (const h of p.headers) if (!S.headers.includes(map(h))) S.headers.push(map(h));
        const by = new Map(S.cases.map(c => [clean(c.f[colSerial()]), c])); let a = 0, u = 0;
        for (const r of p.rows) {
          const f = {}; for (const h of p.headers) f[map(h)] = r.f[h];
          const ex = clean(r.f[ser]) && by.get(clean(r.f[ser]));
          if (ex) { ex.f = { ...ex.f, ...f }; if (r.corr) ex.corr = r.corr; u++; } else { S.cases.push({ id: uid() + a, f, corr: r.corr, createdAt: nowStr() }); a++; }
        }
        logChange("import", `Merged ${file.name}: ${a} added, ${u} updated`);
        toast(`${a} added, ${u} updated.`, 3500);
      }
      closeAll(); location.hash = "dashboard"; render();
    };
  };
  draw(guessDarta(p.headers));
}
function showModal(html) { $("#modal").innerHTML = html; $("#modal").classList.add("on"); $("#scrim").classList.add("on"); }
function startBlank() {
  Object.assign(S, { headers: DEFAULT_HEADERS.slice(), dartaCol: "दर्ता मिती", cases: [], fileName: "New register", loaded: true, changes: [], dirty: 0 });
  render(); openForm(null);
}

// ------------------------------------------------------------- writing the master Excel (in the browser)
async function downloadExcel() {
  if (!S.loaded) { toast("Upload an Excel or start a register first."); return; }
  const t0 = today(), tb = BS.fromAD(t0), hs = S.headers;
  const ev = S.cases.map(c => ({ c, r: BS.evaluate(dateOf(c), t0) })).sort(byExpiry);
  const n = {}; ST.forEach(s => (n[s] = ev.filter(x => x.r.status === s).length));
  const within = d => ev.filter(x => x.r.total >= 0 && x.r.total <= d).length;
  const wb = new ExcelJS.Workbook(); wb.creator = "Case Expiry Monitor";
  const argb = h => "FF" + h.replace("#", "").toUpperCase();
  const thin = { style: "thin", color: { argb: "FFC9CED6" } }, border = { top: thin, left: thin, bottom: thin, right: thin };
  const DEV = "Kokila", LAT = "Arial";
  const isDev = s => /[ऀ-ॿ]/.test(String(s));
  const theme = { RED: ["#B3261E", "#FDECEA"], YELLOW: ["#B07A00", "#FFF6DA"], GREEN: ["#1E6B35", "#E8F5EC"], OVERDUE: ["#5E0B0B", "#F4E3E3"], CHECK: ["#4A4F58", "#EEEFF2"], MASTER: ["#14213D", "#EEF1F7"] };
  const adDate = dn => { const d = new Date(dn * 864e5); return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())); };
  // dashboard
  const d = wb.addWorksheet("DASHBOARD", { properties: { tabColor: { argb: argb("#14213D") } }, views: [{ showGridLines: false }] });
  d.getColumn(1).width = 3; d.getColumn(2).width = 34; d.getColumn(3).width = 12; d.getColumn(4).width = 3; d.getColumn(5).width = 46;
  d.getCell("B2").value = "CASE EXPIRY STATUS"; d.getCell("B2").font = { name: LAT, size: 18, bold: true, color: { argb: argb("#14213D") } };
  d.getCell("B3").value = `As of ${BS.fmt(tb, false)} BS (${BS.adStr(t0)} AD) · downloaded ${nowStr()} · source: ${S.fileName}`;
  d.getCell("B3").font = { name: LAT, size: 10, italic: true, color: { argb: argb("#5D6676") } };
  const rows = [["Total registered cases", ev.length, "#14213D"], ["RED – 1 month or less", n.RED, "#B3261E"], ["YELLOW – 1 to 2 months", n.YELLOW, "#B07A00"],
    ["GREEN – more than 2 months", n.GREEN, "#1E6B35"], ["OVERDUE – expired", n.OVERDUE, "#5E0B0B"], ["Expiring within 30 days", within(30), "#C2410C"],
    ["Expiring within 60 days", within(60), "#1D4ED8"], ["Darta Miti to check", n.CHECK, "#4A4F58"]];
  const max = Math.max(1, ...rows.slice(1, 5).map(r => r[1]));
  rows.forEach(([k, v, col], i) => {
    const r = 5 + i;
    Object.assign(d.getCell(r, 2), { value: k, font: { name: LAT, size: 11, bold: true }, border });
    Object.assign(d.getCell(r, 3), { value: v, font: { name: LAT, size: 14, bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: argb(col) } }, alignment: { horizontal: "center" }, border });
    if (i >= 1 && i <= 4) Object.assign(d.getCell(r, 5), { value: "█".repeat(Math.round(v / max * 40)) || "", font: { name: LAT, size: 10, color: { argb: argb(col) } } });
  });
  ["How to read this file:", "• Expiry = the same BS day and month one year after the Darta Miti (last day of the month if that day does not exist).",
   "• Remaining time is calculated for the date shown above. Upload this file again on any day to see the current position.",
   "• BS month lengths for 2084 and later are projected until the official Panchanga is published."].forEach((t, i) =>
    Object.assign(d.getCell(15 + i, 2), { value: t, font: { name: LAT, size: 10, bold: i === 0, color: { argb: argb("#5D6676") } } }));
  // case sheets
  const sheet = (name, th, list, overdue, title) => {
    const [dark, light] = theme[th];
    const extra = ["Darta Miti (BS)", "Expiry Date (BS)", "Expiry Date (AD)", ...(overdue ? ["Days Overdue"] : ["Remaining Months", "Remaining Days", "Total Remaining Days"]), "Remaining Time", "Expiry Status"];
    const cols = [...hs, ...extra];
    const ws = wb.addWorksheet(name, { properties: { tabColor: { argb: argb(dark) } }, views: [{ state: "frozen", xSplit: 2, ySplit: 4 }] });
    Object.assign(ws.getCell(1, 1), { value: title, font: { name: LAT, size: 14, bold: true, color: { argb: argb(dark) } } });
    Object.assign(ws.getCell(2, 1), { value: `${list.length} cases · as of ${BS.fmt(tb, false)} BS (${BS.adStr(t0)} AD)`, font: { name: LAT, size: 10, italic: true, color: { argb: argb("#5D6676") } } });
    cols.forEach((h, j) => Object.assign(ws.getCell(4, j + 1), { value: h, border, alignment: { wrapText: true, horizontal: "center", vertical: "middle" },
      font: { name: isDev(h) ? DEV : LAT, size: isDev(h) ? 13 : 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: argb(dark) } } }));
    ws.getRow(4).height = 36;
    list.forEach((x, i) => {
      const r = x.r, ok = r.status !== "CHECK";
      const vals = [...hs.map(h => x.c.f[h] ?? ""), ok ? BS.fmt(r.darta) : "INVALID: " + dateOf(x.c), ok ? BS.fmt(r.exp) : "", ok ? adDate(r.expAD) : "",
        ...(overdue ? [-r.total] : ok ? [r.months, r.days, r.total] : ["", "", ""]), r.label, r.status === "CHECK" ? "CHECK DATE" : r.status];
      vals.forEach((v, j) => {
        const cell = ws.getCell(5 + i, j + 1);
        cell.value = /^\d{1,15}$/.test(v) && j < hs.length && !/^0\d/.test(v) ? Number(v) : v;
        cell.border = border;
        cell.font = j < hs.length + 2 ? { name: DEV, size: 13 } : { name: LAT, size: 10 };
        if (i % 2) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(light) } };
        if (v instanceof Date) cell.numFmt = "dd-mmm-yyyy";
      });
      Object.assign(ws.getCell(5 + i, cols.length), { font: { name: LAT, size: 10, bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: argb((theme[r.status] || theme.CHECK)[0]) } } });
    });
    cols.forEach((h, j) => { let L = String(h).length; list.slice(0, 60).forEach((x, i) => { const v = ws.getCell(5 + i, j + 1).value; L = Math.max(L, v instanceof Date ? 11 : String(v ?? "").length); });
      ws.getColumn(j + 1).width = Math.max(8, Math.min(38, L * (j < hs.length ? 0.9 : 1) + 2)); });
    if (list.length) ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + list.length, column: cols.length } };
  };
  const by = s => ev.filter(x => x.r.status === s);
  sheet("RED - URGENT", "RED", by("RED"), false, "RED – URGENT CASES (1 month or less to expiry)");
  sheet("YELLOW - WARNING", "YELLOW", by("YELLOW"), false, "YELLOW – WARNING CASES (more than 1 month, up to 2 months)");
  sheet("GREEN - NORMAL", "GREEN", by("GREEN"), false, "GREEN – NORMAL CASES (more than 2 months)");
  sheet("OVERDUE - EXPIRED", "OVERDUE", by("OVERDUE"), true, "OVERDUE – EXPIRED CASES");
  if (n.CHECK) sheet("CHECK DATE", "CHECK", by("CHECK"), false, "DARTA MITI COULD NOT BE READ – correct these");
  sheet("MASTER DATA", "MASTER", ev.slice().sort((a, b) => serialNum(a.c) - serialNum(b.c)), false, "MASTER DATA – all registered cases");
  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = `MASTER_CASES_${BS.fmt(tb, false)}.xlsx`; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  S.dirty = 0; renderShell(); toast(`Downloaded ${a.download}`, 3500);
}

// ------------------------------------------------------------- session changes + help
function feedHTML(limit) {
  const list = S.changes.slice(0, limit);
  if (!list.length) return `<span class="sub">No changes yet. New, edited and deleted cases are listed here until you close the page.</span>`;
  return list.map(l => `<div><span class="av" style="background:${{ add: "#12a150", edit: "#2546b8", delete: "#d92d20", import: "#7a4cc2" }[l.action] || "#667089"}">${{ add: "+", edit: "✎", delete: "×", import: "⇣" }[l.action] || "•"}</span>
    <span><b>${{ add: "Added", edit: "Edited", delete: "Deleted", import: "Merged" }[l.action] || l.action}</b> ${esc(l.summary)}<div class="t">${esc(l.at)}</div></span></div>`).join("");
}
function changesView() {
  return `<div class="listhead"><div class="title"><span class="chip" style="background:#2546b81a;color:#2546b8">${ic("activity", 22)}</span><div><h1>Changes this session</h1><p>${S.dirty ? `${S.dirty} not yet downloaded.` : "All changes are in your last download."}</p></div></div>
    <div class="tools"><button class="btn primary" data-act="download">${ic("dl", 16)}Download Excel</button></div></div>
    <div class="panel"><div class="feed">${feedHTML(500)}</div></div>`;
}
function helpView() {
  return `<div class="listhead"><div class="title"><span class="chip" style="background:#2546b81a;color:#2546b8">${ic("help", 22)}</span><div><h1>Help &amp; privacy</h1><p>How the expiry is worked out, and where your data goes.</p></div></div></div>
  <section class="helpgrid">
    <div class="panel"><div class="panel-h"><h2>Daily routine</h2></div><ol class="steps">
      <li>Open this site and upload the latest master Excel.</li><li>Check the Red and Overdue lists.</li><li>Register new cases with <b>New case</b>.</li>
      <li>Click <b>Download Excel</b>. The downloaded file is your new master; upload it next time.</li></ol>
      <p class="stamp">Several staff? Agree who keeps the master. A colleague's file can be merged in with Upload Excel → “Merge”, matched by सि.न., so give each new case a unique सि.न.</p></div>
    <div class="panel"><div class="panel-h"><h2>Your data stays on this computer</h2></div><ul class="steps">
      <li>The Excel file is read by this page inside your browser. It is not uploaded to GitHub or any server.</li>
      <li>Nothing is stored. Closing or reloading the page clears everything that has not been downloaded.</li>
      <li>The page warns you before closing if there are changes you have not downloaded.</li></ul></div>
    <div class="panel"><div class="panel-h"><h2>How the expiry is calculated</h2></div><ul class="steps">
      <li>Expiry = the same BS day and month, one year after the Darta Miti (२०८३।०४।०४ → २०८४।०४।०४).</li>
      <li>If that day does not exist next year (e.g. Asar 32 when next Asar has 31 days), the last day of that month is used.</li>
      <li><b>Red</b>: 1 month or less (incl. expiry day) · <b>Yellow</b>: more than 1, up to 2 months · <b>Green</b>: more than 2 months · <b>Overdue</b>: expiry passed.</li>
      <li>BS month lengths for 2084 onward are projected until the official Panchanga is published.</li></ul></div>
  </section>`;
}

// ------------------------------------------------------------- events
document.addEventListener("click", e => {
  const m = e.target.closest("[data-month]"); if (m) { S.month = m.dataset.month; render(); return; }
  const a = e.target.closest("[data-act]");
  if (a) {
    const act = a.dataset.act;
    if (act === "download") { e.preventDefault(); downloadExcel().catch(er => toast("Could not create the Excel: " + er.message, 5000)); }
    else if (act === "close") closeAll();
    else if (act === "edit") { $("#drawer").classList.remove("on"); openForm(a.dataset.id); }
    else if (act === "del") confirmDelete(a.dataset.id);
    else if (act === "del-yes") doDelete(a.dataset.id);
    else if (act === "print") window.print();
    else if (act === "pick") $("#fileIn").click();
    else if (act === "blank") startBlank();
    else if (act === "sort-serial") { S.sort = "serial"; render(); }
    else if (act === "sort-exp") { S.sort = null; render(); }
    return;
  }
  const row = e.target.closest("tr[data-id], .nx[data-id]"); if (row) openDrawer(row.dataset.id);
});
document.addEventListener("keydown", e => {
  if (e.key === "Escape") closeAll();
  if (e.key === "Enter" && e.target.matches && e.target.matches("tr[data-id]")) openDrawer(e.target.dataset.id);
});
$("#scrim").onclick = closeAll;
$("#addBtn").onclick = () => { if (!S.loaded) startBlank(); else openForm(null); };
$("#dlBtn").onclick = () => downloadExcel().catch(er => toast("Could not create the Excel: " + er.message, 5000));
$("#q").oninput = e => { S.q = e.target.value; const r = route(); if (["import", "changes", "help"].includes(r)) location.hash = "all"; else render(); };
window.addEventListener("hashchange", () => { S.month = ""; S.sort = null; if (!["all", "red", "yellow", "green", "overdue", "check", "dashboard"].includes(route())) { S.q = ""; $("#q").value = ""; } closeAll(); render(); window.scrollTo(0, 0); });
window.addEventListener("beforeunload", e => { if (S.dirty) { e.preventDefault(); e.returnValue = ""; } });
let lastDay = today(); setInterval(() => { if (today() !== lastDay) { lastDay = today(); render(); } }, 60000);
render();
