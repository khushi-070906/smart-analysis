"""
Build dashboard.html (self-contained, works offline) from the sales CSV.
Usage:  python build_dashboard.py data/supermarket_sales_500_rows.csv
Then open dashboard.html in any browser.
"""
import json
import sys
import pandas as pd

path = sys.argv[1] if len(sys.argv) > 1 else "data/supermarket_sales_500_rows.csv"
df = pd.read_csv(path)
df["Sales"] = (df["Quantity"] * df["Unit Price"]).round(2)
df["Month"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m")

cols = ["Invoice ID", "Month", "Branch", "City", "Customer Type", "Product",
        "Category", "Quantity", "Unit Price", "Payment", "Rating", "Sales"]
records = df[cols].to_dict("records")

TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Supermarket Sales Dashboard</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2230;--mute:#6b7488;--line:#e4e7ee;--accent:#2f6fed;--accent2:#14a37f}
@media (prefers-color-scheme:dark){:root{--bg:#12151c;--card:#1b1f29;--ink:#e8ebf2;--mute:#9099ad;--line:#2a2f3d;--accent:#6c9bff;--accent2:#3cc9a2}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
header{padding:20px 20px 4px;max-width:1100px;margin:auto}
h1{margin:0 0 4px;font-size:22px}
.sub{color:var(--mute);font-size:13px}
.filters{display:flex;gap:10px;flex-wrap:wrap;padding:12px 20px;max-width:1100px;margin:auto}
select,button{background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:8px;padding:8px 10px;font:inherit}
.grid{display:grid;gap:14px;padding:6px 20px 28px;max-width:1100px;margin:auto;grid-template-columns:repeat(auto-fit,minmax(300px,1fr))}
.kpis{grid-template-columns:repeat(auto-fit,minmax(150px,1fr))}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .v{font-size:22px;font-weight:650}
.kpi .l{color:var(--mute);font-size:12px;text-transform:uppercase;letter-spacing:.04em}
h2{margin:0 0 10px;font-size:14px}
.row{display:grid;grid-template-columns:110px 1fr 84px;gap:8px;align-items:center;margin:6px 0;font-size:13px}
.bar{height:14px;background:var(--line);border-radius:7px;overflow:hidden}
.bar>i{display:block;height:100%;background:var(--accent);border-radius:7px}
.row .n{text-align:right;color:var(--mute);font-variant-numeric:tabular-nums}
.wide{grid-column:1/-1}
svg{width:100%;height:190px}
.empty{color:var(--mute);padding:20px 0}
</style>
</head>
<body>
<header><h1>Supermarket Sales Dashboard</h1><div class="sub" id="sub"></div></header>
<div class="filters">
  <select id="fBranch"></select>
  <select id="fCust"></select>
  <button id="reset">Reset</button>
</div>
<div class="grid kpis" id="kpis"></div>
<div class="grid">
  <div class="card"><h2>Top products by sales</h2><div id="prod"></div></div>
  <div class="card"><h2>Sales by branch</h2><div id="branch"></div></div>
  <div class="card"><h2>Sales by category</h2><div id="cat"></div></div>
  <div class="card"><h2>Transactions by payment method</h2><div id="pay"></div></div>
  <div class="card"><h2>Average transaction: Member vs Normal</h2><div id="cust"></div></div>
  <div class="card wide"><h2>Monthly sales</h2><div id="trend"></div></div>
</div>
<script>
const DATA = __DATA__;
const inr = n => "₹" + n.toLocaleString("en-IN", {maximumFractionDigits: 2, minimumFractionDigits: 2});
const $ = id => document.getElementById(id);
const uniq = k => [...new Set(DATA.map(r => r[k]))].sort();

function fillSelect(el, label, vals, fmt){
  el.innerHTML = `<option value="">${label}</option>` + vals.map(v => `<option value="${v}">${fmt ? fmt(v) : v}</option>`).join("");
}
const cityOf = Object.fromEntries(DATA.map(r => [r.Branch, r.City]));
fillSelect($("fBranch"), "All branches", uniq("Branch"), b => `${b} (${cityOf[b]})`);
fillSelect($("fCust"), "All customers", uniq("Customer Type"));

function group(rows, key, agg){
  const m = new Map();
  rows.forEach(r => { const k = r[key]; if(!m.has(k)) m.set(k, []); m.get(k).push(r); });
  return [...m].map(([k, v]) => [k, agg(v)]);
}
const sum = rs => rs.reduce((a, r) => a + r.Sales, 0);
const mean = rs => sum(rs) / rs.length;

function bars(el, pairs, fmt){
  if(!pairs.length){ el.innerHTML = '<div class="empty">No data for this filter.</div>'; return; }
  const max = Math.max(...pairs.map(p => p[1]));
  el.innerHTML = pairs.map(([k, v]) =>
    `<div class="row"><span>${k}</span><span class="bar"><i style="width:${(v / max * 100).toFixed(1)}%"></i></span><span class="n">${fmt(v)}</span></div>`).join("");
}

function trend(el, pairs){
  if(pairs.length < 2){ el.innerHTML = '<div class="empty">Not enough months for this filter.</div>'; return; }
  const W = 1000, H = 190, p = {l: 56, r: 40, t: 14, b: 28};
  const max = Math.max(...pairs.map(x => x[1])) * 1.1;
  const x = i => p.l + i * (W - p.l - p.r) / (pairs.length - 1);
  const y = v => H - p.b - v / max * (H - p.t - p.b);
  const pts = pairs.map((q, i) => `${x(i)},${y(q[1])}`).join(" ");
  const ticks = [0, .5, 1].map(f => max * f / 1.1);
  el.innerHTML = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">
    ${ticks.map(t => `<line x1="${p.l}" x2="${W - p.r}" y1="${y(t)}" y2="${y(t)}" stroke="var(--line)"/><text x="${p.l - 6}" y="${y(t) + 4}" text-anchor="end" font-size="11" fill="var(--mute)">${Math.round(t / 1000)}k</text>`).join("")}
    <polyline points="${pts}" fill="none" stroke="var(--accent)" stroke-width="3" stroke-linejoin="round"/>
    ${pairs.map((q, i) => `<circle cx="${x(i)}" cy="${y(q[1])}" r="4" fill="var(--accent)"><title>${q[0]}: ${inr(q[1])}</title></circle><text x="${x(i)}" y="${H - 8}" text-anchor="middle" font-size="11" fill="var(--mute)">${q[0]}</text>`).join("")}
  </svg>`;
}

function render(){
  const b = $("fBranch").value, c = $("fCust").value;
  const rows = DATA.filter(r => (!b || r.Branch === b) && (!c || r["Customer Type"] === c));
  $("sub").textContent = `${rows.length} of ${DATA.length} transactions`;
  const total = sum(rows), avg = rows.length ? total / rows.length : 0;
  const rating = rows.length ? rows.reduce((a, r) => a + r.Rating, 0) / rows.length : 0;
  $("kpis").innerHTML = [
    ["Total sales", inr(total)], ["Transactions", rows.length],
    ["Avg transaction", inr(avg)], ["Avg rating", rating.toFixed(2) + " / 5"]
  ].map(([l, v]) => `<div class="card kpi"><div class="l">${l}</div><div class="v">${v}</div></div>`).join("");

  const desc = (a, b2) => b2[1] - a[1];
  bars($("prod"), group(rows, "Product", sum).sort(desc).slice(0, 8), inr);
  bars($("branch"), group(rows, "Branch", sum).map(([k, v]) => [`${k} (${cityOf[k]})`, v]).sort(desc), inr);
  bars($("cat"), group(rows, "Category", sum).sort(desc), inr);
  bars($("pay"), group(rows, "Payment", r => r.length).sort(desc), v => v);
  bars($("cust"), group(rows, "Customer Type", mean).sort(desc), inr);
  trend($("trend"), group(rows, "Month", sum).sort((a, b2) => a[0].localeCompare(b2[0])));
}
$("fBranch").onchange = $("fCust").onchange = render;
$("reset").onclick = () => { $("fBranch").value = ""; $("fCust").value = ""; render(); };
render();
</script>
</body>
</html>
"""

html = TEMPLATE.replace("__DATA__", json.dumps(records, separators=(",", ":")))
with open("dashboard.html", "w", encoding="utf-8") as f:
    f.write(html)
print(f"dashboard.html written ({len(records)} records)")
