// ---- Bikram Sambat engine (pure functions) ----
const BS = (() => {
  const Y0 = 2000, Y1 = 2099;
  const ENC = "243432221213334333212122334432212122343432221123243432221213334333212122334432212122343432221123333433122113334333212122334432212122343432221123333433122122334333212122334432212122343432221123333433122122334333212122343432212122343432221213333433212122334333212122343432221122343432221213333433212122334333212122343432221123243432221213334333212122334342212122343432221123243432221213334333212122334432212122343432221123243433122113334333212122334432212122343432221123333433122122334333212122334432212122343432221123333433122122334333212122343432212122343432221123333433212122334333212122343432221122343432221213333433212122334333212122343432221122343432221213334333212122334342212122343432221123243432221213334333212122334432212122343432221123333433121213334333212122334432212122343432221123333433122113334333212122334432212122343432221123333433122122334333212122343432212122343432221123333433212122334333212122343432221122343432221213333433212122334333212122343432221122343432221213334333212122334333212122334332221222343423221222243432221222334333212222234423221222243432221222243432221222334333221222234432221222243432221222334332221222334333212222234432212122343432221222334333121213334333211222";
  const DAYS = Array.from(ENC, ch => 28 + Number(ch));           // index (y-2000)*12+(m-1)
  const START = [];                                               // AD day number of each month's day 1
  const AD0 = Date.UTC(1943, 3, 14) / 864e5;                      // 2000-01-01 BS = 1943-04-14 AD
  let acc = AD0; for (const d of DAYS) { START.push(acc); acc += d; }
  const MONTHS = ["Baisakh","Jestha","Asar","Shrawan","Bhadra","Asoj","Kartik","Mangsir","Poush","Magh","Falgun","Chaitra"];
  const MONTHS_NP = ["बैशाख","जेठ","असार","साउन","भदौ","असोज","कात्तिक","मंसिर","पुस","माघ","फागुन","चैत"];
  const DEV = "०१२३४५६७८९";
  const idx = (y, m) => (y - Y0) * 12 + (m - 1);
  const dim = (y, m) => DAYS[idx(y, m)];
  const toDev = s => String(s).replace(/[0-9]/g, d => DEV[d]);
  const fmt = (b, dev = true) => {
    const s = b.y + "-" + String(b.m).padStart(2, "0") + "-" + String(b.d).padStart(2, "0");
    return dev ? toDev(s.replace(/-/g, "।")) : s;
  };
  function parse(raw) {
    if (raw === null || raw === undefined) return null;
    let s = String(raw).trim().replace(/[०-९]/g, ch => String(DEV.indexOf(ch)));
    s = s.replace(/[।|,\/.॥\\\-]/g, "-").replace(/\s+/g, "");
    const p = s.split("-");
    if (p.length !== 3 || p.some(x => !/^\d+$/.test(x))) return null;
    const [y, m, d] = p.map(Number);
    if (y < Y0 || y > Y1 - 1 || m < 1 || m > 12 || d < 1 || d > dim(y, m)) return null;
    return { y, m, d };
  }
  const toAD = b => START[idx(b.y, b.m)] + b.d - 1;               // AD day number (days since 1970-01-01)
  function fromAD(dayNum) {
    let lo = 0, hi = START.length - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (START[mid] <= dayNum) lo = mid; else hi = mid - 1; }
    return { y: Y0 + Math.floor(lo / 12), m: lo % 12 + 1, d: dayNum - START[lo] + 1 };
  }
  const expiry = b => ({ y: b.y + 1, m: b.m, d: Math.min(b.d, dim(b.y + 1, b.m)) });
  // calendar months + days from t to e (e >= t); month-end clamped
  function diff(t, e) {
    const a = Math.min(t.d, dim(e.y, e.m));
    let months = (e.y - t.y) * 12 + (e.m - t.m), days;
    if (e.d >= a) days = e.d - a;
    else {
      months -= 1;
      const py = e.m === 1 ? e.y - 1 : e.y, pm = e.m === 1 ? 12 : e.m - 1;
      const pd = dim(py, pm); days = pd - Math.min(t.d, pd) + e.d;
    }
    return { months, days };
  }
  function evaluate(raw, today) {
    const b = parse(raw);
    if (!b) return { status: "CHECK", label: "Darta Miti cannot be read" };
    const e = expiry(b), ead = toAD(e), total = ead - today;
    const out = { darta: b, exp: e, expAD: ead, total };
    if (total < 0) {
      out.status = "OVERDUE"; out.months = 0; out.days = 0;
      out.label = "EXPIRED – " + (-total) + (-total === 1 ? " Day" : " Days") + " Overdue";
      return out;
    }
    const r = diff(fromAD(today), e);
    out.months = r.months; out.days = r.days;
    out.status = (r.months < 1 || (r.months === 1 && r.days === 0)) ? "RED"
      : (r.months < 2 || (r.months === 2 && r.days === 0)) ? "YELLOW" : "GREEN";
    if (total === 0) out.label = "Expires Today";
    else out.label = [r.months ? r.months + (r.months === 1 ? " Month" : " Months") : "",
                      r.days ? r.days + (r.days === 1 ? " Day" : " Days") : ""].filter(Boolean).join(", ") + " Remaining";
    return out;
  }
  const todayDayNum = () => { const n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) / 864e5; };
  const adStr = n => new Date(n * 864e5).toISOString().slice(0, 10);
  return { Y0, Y1, MONTHS, MONTHS_NP, dim, parse, toAD, fromAD, expiry, diff, evaluate, fmt, toDev, todayDayNum, adStr };
})();
if (typeof module !== "undefined") module.exports = BS;
