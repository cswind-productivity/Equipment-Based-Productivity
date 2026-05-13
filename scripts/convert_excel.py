#!/usr/bin/env python3
"""Generate the CS WIND Equipment-Based Productivity dashboard from Excel.

This script reads `GitHub_Export` first, falls back to `Raw(1)`, and writes
both `index.html` and `dashboard_template.html` as a complete standalone
HTML dashboard.
"""

import json
import math
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1] if Path(__file__).parent.name == "scripts" else Path.cwd()
EXCEL_FILE = ROOT / "Raw Data_Global Equipment-Based Productivity.xlsx"
OUTPUT_INDEX = ROOT / "index.html"
OUTPUT_TEMPLATE = ROOT / "dashboard_template.html"

FACTORY_ORDER = ["VN #1", "VN #2", "TW", "CN", "TR #1", "TR #2", "AM", "PT On", "PT Off"]

MACHINE_CATEGORIES = [
    "Roll Bending Machine",
    "L/W Machine",
    "C/W Machine",
    "Growing Line",
    "Paint Booth",
    "Paint Line",
]

PERFORMANCE_CATEGORIES = ["Bending", "L/W", "C/W", "BT GT", "WT GT"]

PRODUCTIVITY_METRICS = [
    {
        "key": "bending",
        "title": "Bending / Machine",
        "unit": "Can / Machine",
        "num": "Bending",
        "den": "Roll Bending Machine",
    },
    {
        "key": "lw",
        "title": "L/W / Machine",
        "unit": "Can / Machine",
        "num": "L/W",
        "den": "L/W Machine",
    },
    {
        "key": "cw",
        "title": "C/W / Machine",
        "unit": "CS Joint / Machine",
        "num": "C/W",
        "den": "C/W Machine",
    },
    {
        "key": "bt",
        "title": "BT GT / Growing Line",
        "unit": "BT Sec / Line",
        "num": "BT GT",
        "den": "Growing Line",
    },
    {
        "key": "wt_booth",
        "title": "WT GT / Paint Booth",
        "unit": "WT Sec / Booth",
        "num": "WT GT",
        "den": "Paint Booth",
    },
    {
        "key": "wt_line",
        "title": "WT GT / Paint Line",
        "unit": "WT Sec / Line",
        "num": "WT GT",
        "den": "Paint Line",
    },
]

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number(value):
    if pd.isna(value):
        return 0.0
    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if value in ("", "-", "#N/A"):
                return 0.0
        num = float(value)
        if math.isnan(num) or math.isinf(num):
            return 0.0
        return num
    except Exception:
        return 0.0


def read_excel_data():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE.name}")

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="GitHub_Export", engine="openpyxl")
    except Exception:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Raw(1)", engine="openpyxl")

    df.columns = [clean_text(c) for c in df.columns]

    required = ["Year", "Week", "Factory", "Type", "Category", "Q'ty"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    optional = [c for c in ["Month", "YearMonth", "Month_Sort"] if c in df.columns]
    df = df[required + optional].copy()

    df = df.dropna(subset=["Year", "Week", "Factory", "Type", "Category"], how="any")
    df["Year"] = df["Year"].apply(lambda x: int(to_number(x)) if to_number(x) else 0)
    df["Week"] = df["Week"].apply(lambda x: int(to_number(x)) if to_number(x) else 0)
    df["Factory"] = df["Factory"].apply(clean_text)
    df["Type"] = df["Type"].apply(clean_text)
    df["Category"] = df["Category"].apply(clean_text)
    df["Q'ty"] = df["Q'ty"].apply(to_number)

    if "Month" not in df.columns:
        df["Month"] = df["Week"].apply(
            lambda w: MONTH_ORDER[min(max(int((w - 1) / 4), 0), 11)] if w > 0 else ""
        )
    else:
        df["Month"] = df["Month"].apply(clean_text)

    if "YearMonth" not in df.columns:
        df["YearMonth"] = df.apply(
            lambda r: f"{r['Year']}-{r['Month']}" if r["Month"] else "",
            axis=1,
        )
    else:
        df["YearMonth"] = df["YearMonth"].apply(clean_text)

    if "Month_Sort" not in df.columns:
        month_index = {m: i + 1 for i, m in enumerate(MONTH_ORDER)}
        df["Month_Sort"] = df.apply(
            lambda r: r["Year"] * 100 + month_index.get(r["Month"], 0),
            axis=1,
        )
    else:
        df["Month_Sort"] = df["Month_Sort"].apply(to_number)

    df = df[(df["Year"] > 0) & (df["Week"] > 0) & (df["Factory"] != "")]
    return df


def sorted_factories(df):
    found = list(dict.fromkeys(df["Factory"].dropna().map(clean_text).tolist()))
    ordered = [f for f in FACTORY_ORDER if f in found]
    ordered += sorted([f for f in found if f not in ordered])
    return ordered


def pivot_values(df, type_name, categories, factories):
    subset = df[df["Type"].str.lower() == type_name.lower()]
    result = {factory: {cat: 0.0 for cat in categories} for factory in factories}

    if subset.empty:
        return result

    grouped = subset.groupby(["Factory", "Category"], dropna=False)["Q'ty"].sum()

    for (factory, cat), value in grouped.items():
        factory = clean_text(factory)
        cat = clean_text(cat)
        if factory in result and cat in categories:
            result[factory][cat] = round(float(value), 4)

    return result


def calc_productivity(machine, performance, factories):
    prod = {factory: {} for factory in factories}

    for factory in factories:
        for metric in PRODUCTIVITY_METRICS:
            numerator = performance.get(factory, {}).get(metric["num"], 0.0)
            denominator = machine.get(factory, {}).get(metric["den"], 0.0)
            value = numerator / denominator if denominator else 0.0
            prod[factory][metric["key"]] = round(value, 4)

    return prod


def build_period(df, factories):
    machine = pivot_values(df, "Machine", MACHINE_CATEGORIES, factories)
    perf = pivot_values(df, "Performance", PERFORMANCE_CATEGORIES, factories)
    prod = calc_productivity(machine, perf, factories)

    total_equipment = sum(sum(machine[f].values()) for f in factories)
    total_production = sum(sum(perf[f].values()) for f in factories)

    prod_values = [v for f in factories for v in prod[f].values() if v]
    avg_productivity = sum(prod_values) / len(prod_values) if prod_values else 0.0

    return {
        "machine": machine,
        "performance": perf,
        "productivity": prod,
        "summary": {
            "totalEquipment": round(total_equipment, 2),
            "totalProduction": round(total_production, 2),
            "avgProductivity": round(avg_productivity, 2),
        },
    }


def build_dashboard_data(df):
    factories = sorted_factories(df)
    years = sorted(df["Year"].unique().tolist())

    weekly = {}
    monthly = {}
    weeks_by_year = {}
    months_by_year = {}

    for year in years:
        ydf = df[df["Year"] == year]

        weekly[str(year)] = {}
        for week in sorted(ydf["Week"].unique().tolist()):
            wdf = ydf[ydf["Week"] == week]
            weekly[str(year)][f"WK{int(week):02d}"] = build_period(wdf, factories)

        weeks_by_year[str(year)] = list(weekly[str(year)].keys())

        monthly[str(year)] = {}
        mdf = ydf.copy()
        month_meta = (
            mdf[["Month", "YearMonth", "Month_Sort"]]
            .drop_duplicates()
            .sort_values(["Month_Sort", "YearMonth"])
            .to_dict("records")
        )

        for item in month_meta:
            month_label = clean_text(item.get("YearMonth")) or clean_text(item.get("Month"))
            if not month_label:
                continue
            one = mdf[mdf["YearMonth"] == month_label]
            monthly[str(year)][month_label] = build_period(one, factories)

        months_by_year[str(year)] = list(monthly[str(year)].keys())

    def last_nonzero(period_dict):
        last = None
        for key, value in period_dict.items():
            if value["summary"].get("totalProduction", 0) > 0:
                last = key
        return last or (list(period_dict.keys())[-1] if period_dict else "")

    default_year = str(years[-1]) if years else ""
    default_week = last_nonzero(weekly.get(default_year, {})) if default_year else ""
    default_month = last_nonzero(monthly.get(default_year, {})) if default_year else ""

    return {
        "factories": factories,
        "years": [str(y) for y in years],
        "weeksByYear": weeks_by_year,
        "monthsByYear": months_by_year,
        "weekly": weekly,
        "monthly": monthly,
        "defaultYear": default_year,
        "defaultWeek": default_week,
        "defaultMonth": default_month,
        "metrics": PRODUCTIVITY_METRICS,
        "machineCategories": MACHINE_CATEGORIES,
        "performanceCategories": PERFORMANCE_CATEGORIES,
    }


def html_template(data_json):
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CS WIND Global Equipment-Based Productivity Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root {{ --navy:#1f4e78; --bg:#eef2f6; --line:#d7e0ea; --card:#ffffff; --text:#111827; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font-family:Arial, 'Malgun Gothic', sans-serif; }}
.wrap {{ max-width:1480px; margin:0 auto; background:#fff; min-height:100vh; }}
.header {{ background:var(--navy); color:#fff; padding:26px 28px; text-align:center; }}
.header h1 {{ margin:0; font-size:28px; letter-spacing:.5px; font-weight:800; }}
.controls {{ display:flex; gap:14px; align-items:center; padding:16px 24px; border-bottom:1px solid var(--line); background:#f8fafc; flex-wrap:wrap; }}
.controls label {{ font-weight:700; font-size:13px; }}
.btn-group {{ display:inline-flex; border:1px solid #b9c7d6; border-radius:4px; overflow:hidden; background:#fff; }}
.btn-group button {{ min-width:72px; padding:8px 16px; border:0; background:#fff; cursor:pointer; font-weight:700; }}
.btn-group button.active {{ background:var(--navy); color:#fff; }}
select {{ min-width:96px; padding:8px 12px; border:1px solid #b9c7d6; border-radius:4px; background:#fff; font-weight:700; }}
.content {{ padding:18px 24px 40px; }}
.section-title {{ background:var(--navy); color:#fff; padding:12px 16px; border-radius:5px; font-weight:800; margin:12px 0 16px; font-size:17px; }}
table {{ width:100%; border-collapse:collapse; margin-bottom:22px; font-size:12px; background:#fff; }}
th {{ background:#315d84; color:#fff; padding:10px 8px; border:1px solid #d9e2ec; text-align:center; }}
td {{ padding:8px; border:1px solid #d9e2ec; text-align:center; }}
td:first-child, th:first-child {{ text-align:left; font-weight:700; }}
tfoot td {{ background:#fff1b8; font-weight:800; }}
.grid {{ display:grid; grid-template-columns:repeat(3, 1fr); gap:18px; margin-bottom:28px; }}
.chart-card {{ background:var(--card); border:1px solid #cfd9e5; border-radius:7px; padding:12px 14px; min-height:280px; box-shadow:0 1px 2px rgba(0,0,0,.04); }}
.chart-title {{ text-align:center; font-weight:800; margin:4px 0 0; font-size:13px; }}
.chart-unit {{ text-align:right; font-size:10px; color:#475569; margin-bottom:4px; }}
.chart-card canvas {{ width:100% !important; height:220px !important; }}
.trend-title {{ background:var(--navy); color:#fff; padding:10px 14px; border-radius:5px; font-weight:800; margin-top:18px; }}
.factory-box {{ border:1px solid #cfd9e5; background:#f8fafc; border-radius:7px; padding:12px; margin:12px 0 18px; }}
.factory-grid {{ display:grid; grid-template-columns:repeat(4, minmax(160px, 1fr)); gap:8px; }}
.factory-grid label {{ border:1px solid #89a7c5; border-radius:4px; padding:5px 8px; font-size:12px; background:#fff; font-weight:700; }}
.small-note {{ color:#64748b; font-size:11px; margin-top:6px; }}
@media (max-width:900px) {{ .grid {{ grid-template-columns:1fr; }} .factory-grid {{ grid-template-columns:1fr 1fr; }} }}
</style>
</head>
<body>
<div class="wrap">
  <div class="header"><h1>CS WIND Global Equipment-Based Productivity Dashboard</h1></div>

  <div class="controls">
    <label>View:</label>
    <div class="btn-group">
      <button id="weeklyBtn" onclick="setView('weekly')">Weekly</button>
      <button id="monthlyBtn" onclick="setView('monthly')">Monthly</button>
    </div>
    <label>Year:</label>
    <select id="yearSelect" onchange="onYearChange()"></select>
    <label id="periodLabel">Week:</label>
    <select id="periodSelect" onchange="renderAll()"></select>
  </div>

  <div class="content">
    <div class="section-title">1. 공장별 장비 현황 (Equipment Inventory)</div>
    <div id="machineTable"></div>

    <div class="section-title">2. 공장별 생산 실적 (Production Performance)</div>
    <div id="performanceTable"></div>

    <div class="section-title">3. 장비당 생산 효율성 (Production per Equipment)</div>
    <div class="grid" id="barCharts"></div>

    <div class="trend-title">Trend of Equipment-Based Productivity (최근 10주)</div>
    <div class="factory-box">
      <b>공장 선택 (Factory Selection)</b>
      <div class="factory-grid" id="factoryChecks"></div>
      <div class="small-note">체크한 공장만 하단 Trend 라인그래프에 표시됩니다.</div>
    </div>

    <div class="grid" id="trendCharts"></div>
  </div>
</div>

<script>
const DASHBOARD_DATA = {data_json};

let currentView = 'weekly';
let chartObjects = [];

const palette = [
  '#2563eb', '#38bdf8', '#0f766e', '#dc2626', '#f97316',
  '#8b5cf6', '#0891b2', '#64748b', '#111827', '#84cc16'
];

function fmt(n, digits = 2) {{
  const v = Number(n || 0);
  return v === 0 ? '0' : v.toLocaleString(undefined, {{
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  }});
}}

function destroyCharts() {{
  chartObjects.forEach(c => c.destroy());
  chartObjects = [];
}}

function getYear() {{
  return document.getElementById('yearSelect').value;
}}

function getPeriod() {{
  return document.getElementById('periodSelect').value;
}}

function getPeriodData() {{
  const y = getYear();
  const p = getPeriod();
  return (DASHBOARD_DATA[currentView] &&
          DASHBOARD_DATA[currentView][y] &&
          DASHBOARD_DATA[currentView][y][p]) || null;
}}

function setView(view) {{
  currentView = view;
  document.getElementById('weeklyBtn').classList.toggle('active', view === 'weekly');
  document.getElementById('monthlyBtn').classList.toggle('active', view === 'monthly');
  document.getElementById('periodLabel').textContent = view === 'weekly' ? 'Week:' : 'Month:';
  populatePeriods();
  renderAll();
}}

function populateYears() {{
  const sel = document.getElementById('yearSelect');
  sel.innerHTML = '';

  DASHBOARD_DATA.years.forEach(y => {{
    const o = document.createElement('option');
    o.value = y;
    o.textContent = y;
    sel.appendChild(o);
  }});

  if (DASHBOARD_DATA.defaultYear) {{
    sel.value = DASHBOARD_DATA.defaultYear;
  }}
}}

function populatePeriods() {{
  const y = getYear();
  const list = currentView === 'weekly'
    ? (DASHBOARD_DATA.weeksByYear[y] || [])
    : (DASHBOARD_DATA.monthsByYear[y] || []);

  const sel = document.getElementById('periodSelect');
  sel.innerHTML = '';

  list.forEach(p => {{
    const o = document.createElement('option');
    o.value = p;
    o.textContent = p;
    sel.appendChild(o);
  }});

  const def = currentView === 'weekly'
    ? DASHBOARD_DATA.defaultWeek
    : DASHBOARD_DATA.defaultMonth;

  if (list.includes(def)) {{
    sel.value = def;
  }}
}}

function onYearChange() {{
  populatePeriods();
  renderAll();
}}

function tableHtml(titleCols, rows, totalRow = false) {{
  let html = '<table><thead><tr>' + titleCols.map(c => `<th>${{c}}</th>`).join('') + '</tr></thead><tbody>';

  rows.forEach(r => {{
    html += '<tr>' + r.map(c => `<td>${{c}}</td>`).join('') + '</tr>';
  }});

  html += '</tbody>';

  if (totalRow && rows.length) {{
    const totals = ['CS Total'];
    for (let i = 1; i < titleCols.length; i++) {{
      let s = 0;
      rows.forEach(r => {{
        s += Number(String(r[i]).replace(/,/g, '')) || 0;
      }});
      totals.push(fmt(s, 2));
    }}
    html += '<tfoot><tr>' + totals.map(c => `<td>${{c}}</td>`).join('') + '</tr></tfoot>';
  }}

  return html + '</table>';
}}

function renderTables(data) {{
  const factories = DASHBOARD_DATA.factories;

  const machineCols = ['Factory'].concat(DASHBOARD_DATA.machineCategories).concat(['Remark']);
  const machineRows = factories.map(f =>
    [f].concat(DASHBOARD_DATA.machineCategories.map(c => fmt(data.machine[f]?.[c], 2))).concat([''])
  );

  document.getElementById('machineTable').innerHTML = tableHtml(machineCols, machineRows, true);

  const perfCols = ['Factory'].concat(DASHBOARD_DATA.performanceCategories).concat(['Remark']);
  const perfRows = factories.map(f =>
    [f].concat(DASHBOARD_DATA.performanceCategories.map(c => fmt(data.performance[f]?.[c], 2))).concat([''])
  );

  document.getElementById('performanceTable').innerHTML = tableHtml(perfCols, perfRows, false);
}}

function makeChart(canvas, config) {{
  const c = new Chart(canvas, config);
  chartObjects.push(c);
  return c;
}}

function renderBarCharts(data) {{
  const box = document.getElementById('barCharts');
  box.innerHTML = '';

  const labels = DASHBOARD_DATA.factories.concat(['CS Total']);

  DASHBOARD_DATA.metrics.forEach((m) => {{
    const card = document.createElement('div');
    card.className = 'chart-card';
    card.innerHTML = `
      <div class="chart-title">${{m.title}}</div>
      <div class="chart-unit">(${{m.unit}})</div>
      <canvas></canvas>
    `;
    box.appendChild(card);

    const vals = DASHBOARD_DATA.factories.map(f => Number(data.productivity[f]?.[m.key] || 0));
    const nonzero = vals.filter(v => v);
    const total = nonzero.length ? nonzero.reduce((a, b) => a + b, 0) / nonzero.length : 0;
    const allVals = vals.concat([Number(total.toFixed(2))]);

    makeChart(card.querySelector('canvas'), {{
      type: 'bar',
      data: {{
        labels: labels,
        datasets: [{{
          data: allVals,
          backgroundColor: labels.map((_, i) => palette[i % palette.length]),
          borderWidth: 0
        }}]
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{ display: false }}
        }},
        scales: {{
          y: {{ beginAtZero: true, grid: {{ color: '#e5e7eb' }} }},
          x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }}
        }}
      }}
    }});
  }});
}}

function renderFactoryChecks() {{
  const box = document.getElementById('factoryChecks');
  if (box.children.length) return;

  DASHBOARD_DATA.factories.concat(['CS Total']).forEach(f => {{
    const label = document.createElement('label');
    label.innerHTML = `<input type="checkbox" value="${{f}}" checked onchange="renderTrendCharts()"> ${{f}}`;
    box.appendChild(label);
  }});
}}

function selectedFactories() {{
  return Array.from(document.querySelectorAll('#factoryChecks input:checked')).map(i => i.value);
}}

function recentWeeklyKeys() {{
  const y = getYear();
  const list = DASHBOARD_DATA.weeksByYear[y] || [];
  const current = getPeriod();
  let idx = list.indexOf(current);

  if (idx < 0) {{
    idx = list.length - 1;
  }}

  return list.slice(Math.max(0, idx - 9), idx + 1);
}}

function renderTrendCharts() {{
  const box = document.getElementById('trendCharts');
  box.innerHTML = '';

  const year = getYear();
  const weeks = recentWeeklyKeys();
  const selected = selectedFactories();

  DASHBOARD_DATA.metrics.forEach((m) => {{
    const card = document.createElement('div');
    card.className = 'chart-card';
    card.innerHTML = `
      <div class="chart-title">${{m.title}}</div>
      <div class="chart-unit">(${{m.unit}})</div>
      <canvas></canvas>
    `;
    box.appendChild(card);

    const datasets = selected.map((f, i) => {{
      const values = weeks.map(w => {{
        const d = DASHBOARD_DATA.weekly[year]?.[w];

        if (!d) {{
          return 0;
        }}

        if (f === 'CS Total') {{
          const vals = DASHBOARD_DATA.factories
            .map(ff => Number(d.productivity[ff]?.[m.key] || 0))
            .filter(v => v);

          return vals.length
            ? Number((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2))
            : 0;
        }}

        return Number(d.productivity[f]?.[m.key] || 0);
      }});

      return {{
        label: f,
        data: values,
        borderColor: palette[i % palette.length],
        backgroundColor: palette[i % palette.length],
        tension: 0.25,
        pointRadius: 2,
        fill: false
      }};
    }});

    makeChart(card.querySelector('canvas'), {{
      type: 'line',
      data: {{
        labels: weeks,
        datasets: datasets
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        plugins: {{
          legend: {{
            display: true,
            position: 'bottom',
            labels: {{ boxWidth: 10, font: {{ size: 10 }} }}
          }}
        }},
        scales: {{
          y: {{ beginAtZero: true, grid: {{ color: '#e5e7eb' }} }},
          x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }}
        }}
      }}
    }});
  }});
}}

function renderAll() {{
  destroyCharts();

  const data = getPeriodData();

  if (!data) {{
    document.getElementById('machineTable').innerHTML = '<p>No data.</p>';
    document.getElementById('performanceTable').innerHTML = '<p>No data.</p>';
    document.getElementById('barCharts').innerHTML = '';
    document.getElementById('trendCharts').innerHTML = '';
    return;
  }}

  renderTables(data);
  renderBarCharts(data);
  renderFactoryChecks();
  renderTrendCharts();
}}

populateYears();
setView('weekly');
</script>
</body>
</html>
'''


def main():
    df = read_excel_data()
    data = build_dashboard_data(df)

    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = html_template(data_json)

    OUTPUT_INDEX.write_text(html, encoding="utf-8")
    OUTPUT_TEMPLATE.write_text(html, encoding="utf-8")

    print(f"Generated {OUTPUT_INDEX.name} and {OUTPUT_TEMPLATE.name}")
    print(f"Years: {', '.join(data['years'])}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
