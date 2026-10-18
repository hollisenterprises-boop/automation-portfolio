"""Build an interactive sales dashboard (one HTML file) from the cleaned store data.

    python build_dashboard.py

Reads the Clean Data sheet that the sales report builder writes
(../sales-report/output/sales_report.xlsx) and writes output/dashboard.html:
KPI cards, net sales by month, top products, sales by state and the latest
orders, with a state filter and hover details. Open it in any browser.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
SOURCE = HERE.parent / "sales-report" / "output" / "sales_report.xlsx"
OUT = HERE / "output" / "dashboard.html"
SOLD = {"Shipped", "Delivered", "Completed"}


def load_orders(path: Path = SOURCE) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Clean Data")
    df["Order Date"] = pd.to_datetime(df["Order Date"])
    df["Month"] = df["Order Date"].dt.strftime("%Y-%m")
    return df


def records(df: pd.DataFrame) -> list[dict]:
    """The rows the page needs, kept small: one entry per order line."""
    return [
        {"id": r["Order ID"], "date": r["Order Date"].strftime("%Y-%m-%d"), "month": r["Month"],
         "product": r["Product Name"], "qty": int(r["Qty"]), "total": round(float(r["Line Total"]), 2),
         "status": r["Order Status"], "state": r["Ship State"]}
        for _, r in df.iterrows()
    ]


def build(df: pd.DataFrame) -> str:
    data = json.dumps(records(df), separators=(",", ":"))
    months = sorted(df["Month"].unique())
    first, last = pd.Timestamp(months[0] + "-01"), pd.Timestamp(months[-1] + "-01")
    period = f"{first:%B} to {last:%B %Y}"
    return TEMPLATE.replace("__DATA__", data).replace("__PERIOD__", period).replace(
        "__SOLD__", json.dumps(sorted(SOLD)))


TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Store sales dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root { --bg: #0b1220; --panel: #14223b; --panel-2: #1b2c4a; --ink: #f1f5f9; --ink-2: #a5b4cb;
          --muted: #6b7f9e; --grid: #23365a; --accent: #22c55e; --accent-2: #38bdf8;
          --good: #22c55e; --warn: #fab219; --bad: #f05252; }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--ink); font: 15px/1.45 "Segoe UI", system-ui, sans-serif; }
  .wrap { max-width: 1520px; margin: 0 auto; padding: 28px 32px 36px; }
  header { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
  h1 { margin: 0; font-size: 30px; font-weight: 650; }
  .sub { color: var(--ink-2); margin-top: 4px; }
  .filters { display: flex; gap: 12px; align-items: center; color: var(--ink-2); }
  select { background: var(--panel-2); color: var(--ink); border: 1px solid var(--grid); border-radius: 10px;
           padding: 8px 12px; font: inherit; }
  .kpis { display: grid; grid-template-columns: repeat(5, 1fr); gap: 16px; margin: 22px 0; }
  .card { background: var(--panel); border-radius: 16px; padding: 18px 20px; min-width: 0; }
  .table-wrap { overflow-x: auto; }
  .kpi .label { color: var(--ink-2); font-size: 14px; }
  .kpi .value { font-size: 32px; font-weight: 700; margin-top: 4px; }
  .kpi .note { color: var(--muted); font-size: 13px; margin-top: 2px; }
  .grid { display: grid; grid-template-columns: 1.5fr 1fr; gap: 16px; }
  .card h2 { margin: 0 0 12px; font-size: 16px; font-weight: 600; color: var(--ink-2); }
  .chart { position: relative; height: 300px; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; color: var(--muted); font-weight: 600; padding: 8px 6px; border-bottom: 1px solid var(--grid); }
  td { padding: 8px 6px; border-bottom: 1px solid #1a2a47; }
  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
  .badge { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink); }
  .dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
  footer { color: var(--muted); font-size: 13px; margin-top: 18px; }
  @media (max-width: 1100px) { .kpis { grid-template-columns: repeat(2, 1fr); } .grid { grid-template-columns: 1fr; } }
  @media (max-width: 600px) { .wrap { padding: 20px 16px 28px; } .kpis { gap: 10px; } .card { padding: 14px; }
                              .kpi .value { font-size: 26px; } h1 { font-size: 26px; } }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Store sales dashboard</h1>
      <div class="sub">__PERIOD__ &middot; built from the cleaned order export &middot; sample data</div>
    </div>
    <div class="filters">
      <label for="state">State</label>
      <select id="state"></select>
    </div>
  </header>

  <section class="kpis" id="kpis"></section>

  <section class="grid">
    <div class="card"><h2>Net sales by month</h2><div class="chart"><canvas id="byMonth"></canvas></div></div>
    <div class="card"><h2>Top products by net sales</h2><div class="chart"><canvas id="byProduct"></canvas></div></div>
    <div class="card"><h2>Latest orders</h2><div class="table-wrap"><table id="latest"></table></div></div>
    <div class="card"><h2>Net sales by state</h2><div class="chart"><canvas id="byState"></canvas></div></div>
  </section>
  <footer>Net sales count shipped and delivered orders only. Refunds, cancellations and pending orders are shown separately. All data is made up.</footer>
</div>

<script>
const ROWS = __DATA__;
const SOLD = new Set(__SOLD__);
const STATUS = { Shipped: "var(--good)", Delivered: "var(--good)", Pending: "var(--warn)",
                 Refunded: "var(--bad)", Cancelled: "var(--bad)" };
const css = name => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
const money = v => "$" + v.toLocaleString("en-AU", { maximumFractionDigits: 0 });
const money2 = v => "$" + v.toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const sum = (rows, f) => rows.reduce((a, r) => a + f(r), 0);
const groupSum = (rows, key) => { const m = new Map(); rows.forEach(r => m.set(r[key], (m.get(r[key]) || 0) + r.total)); return m; };

const states = [...new Set(ROWS.map(r => r.state))].sort();
const select = document.getElementById("state");
select.innerHTML = ['<option value="">All states</option>', ...states.map(s => `<option>${s}</option>`)].join("");

Chart.defaults.color = css("--ink-2");
Chart.defaults.font.family = "Segoe UI, system-ui, sans-serif";
Chart.defaults.borderColor = css("--grid");
const tooltip = { backgroundColor: "#0b1220", borderColor: css("--grid"), borderWidth: 1, padding: 10,
                  callbacks: { label: c => " " + money2(c.parsed.y ?? c.parsed.x) } };
const bar = (id, horizontal) => new Chart(document.getElementById(id), {
  type: "bar",
  data: { labels: [], datasets: [{ data: [], backgroundColor: horizontal ? css("--accent-2") : css("--accent"),
                                   borderRadius: 4, borderSkipped: false, maxBarThickness: horizontal ? 22 : 46 }] },
  options: { indexAxis: horizontal ? "y" : "x", maintainAspectRatio: false, animation: false,
             plugins: { legend: { display: false }, tooltip },
             scales: { [horizontal ? "x" : "y"]: { grid: { color: css("--grid") }, ticks: { callback: v => money(v) } },
                       [horizontal ? "y" : "x"]: { grid: { display: false } } } }
});
const charts = { byMonth: bar("byMonth", false), byProduct: bar("byProduct", true), byState: bar("byState", true) };

function setBars(chart, entries) {
  chart.data.labels = entries.map(e => e[0]);
  chart.data.datasets[0].data = entries.map(e => Math.round(e[1] * 100) / 100);
  chart.update();
}

function render() {
  const rows = ROWS.filter(r => !select.value || r.state === select.value);
  const sold = rows.filter(r => SOLD.has(r.status));
  const orders = new Set(sold.map(r => r.id));
  const net = sum(sold, r => r.total);
  const lost = rows.filter(r => r.status === "Refunded" || r.status === "Cancelled");
  const pending = sum(rows.filter(r => r.status === "Pending"), r => r.total);
  const kpis = [
    ["Net sales", money(net), `${orders.size} orders`],
    ["Units sold", sum(sold, r => r.qty).toLocaleString(), "shipped and delivered"],
    ["Average order", money2(orders.size ? net / orders.size : 0), "per order"],
    ["Refunded or cancelled", money(sum(lost, r => r.total)), `${new Set(lost.map(r => r.id)).size} orders`],
    ["Pending", money(pending), "not shipped yet"],
  ];
  document.getElementById("kpis").innerHTML = kpis.map(([l, v, n]) =>
    `<div class="card kpi"><div class="label">${l}</div><div class="value">${v}</div><div class="note">${n}</div></div>`).join("");

  const months = [...groupSum(sold, "month").entries()].sort((a, b) => a[0].localeCompare(b[0]))
    .map(([m, v]) => [new Date(m + "-01T00:00").toLocaleString("en-AU", { month: "short" }), v]);
  setBars(charts.byMonth, months);
  setBars(charts.byProduct, [...groupSum(sold, "product").entries()].sort((a, b) => b[1] - a[1]).slice(0, 6));
  setBars(charts.byState, [...groupSum(sold, "state").entries()].sort((a, b) => b[1] - a[1]));

  const latest = [...rows].sort((a, b) => b.date.localeCompare(a.date) || b.id.localeCompare(a.id)).slice(0, 8);
  document.getElementById("latest").innerHTML =
    "<tr><th>Order</th><th>Date</th><th>Product</th><th>State</th><th>Status</th><th class='num'>Total</th></tr>" +
    latest.map(r => `<tr><td>${r.id}</td><td>${r.date}</td><td>${r.product}</td><td>${r.state}</td>` +
      `<td><span class="badge"><span class="dot" style="background:${STATUS[r.status] || "var(--muted)"}"></span>${r.status}</span></td>` +
      `<td class="num">${money2(r.total)}</td></tr>`).join("");
}

const wanted = new URLSearchParams(location.search).get("state");  // dashboard.html?state=WA
if (wanted && states.includes(wanted)) select.value = wanted;
select.addEventListener("change", () => {
  history.replaceState(null, "", select.value ? `?state=${select.value}` : location.pathname);
  render();
});
render();
</script>
</body>
</html>
"""


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    df = load_orders()
    OUT.write_text(build(df), encoding="utf-8")
    print(f"{len(df)} order lines -> {OUT}")


if __name__ == "__main__":
    main()
