#!/usr/bin/env python3
"""CS WIND old-style dashboard generator.

Reads Raw Data_Global Equipment-Based Productivity.xlsx and generates both
index.html and dashboard_template.html.

UI target:
- No "Factory for Trend" top filter
- View: Weekly / Monthly
- Year / Week or Month selector
- Equipment table, production table, bar charts, and recent 10-week trend charts
"""

import json
import math
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
EXCEL_FILE = ROOT / "Raw Data_Global Equipment-Based Productivity.xlsx"
INDEX_FILE = ROOT / "index.html"
TEMPLATE_FILE = ROOT / "dashboard_template.html"

FACTORY_ORDER = ["VN #1", "VN #2", "TW", "CN", "TR #1", "TR #2", "AM", "PT On", "PT Off"]
MACHINE_CATEGORIES = ["Roll Bending Machine", "L/W Machine", "C/W Machine", "Growing Line", "Paint Booth", "Paint Line"]
PERFORMANCE_CATEGORIES = ["Bending", "L/W", "C/W", "BT GT", "WT GT"]
METRICS = [
    {"key": "bending", "title": "Bending / Machine", "unit": "Can / Machine", "num": "Bending", "den": "Roll Bending Machine"},
    {"key": "lw", "title": "L/W / Machine", "unit": "Can / Machine", "num": "L/W", "den": "L/W Machine"},
    {"key": "cw", "title": "C/W / Machine", "unit": "CS Joint / Machine", "num": "C/W", "den": "C/W Machine"},
    {"key": "bt", "title": "BT GT / Growing Line", "unit": "BT Sec / Line", "num": "BT GT", "den": "Growing Line"},
    {"key": "wt_booth", "title": "WT GT / Paint Booth", "unit": "WT Sec / Booth", "num": "WT GT", "den": "Paint Booth"},
    {"key": "wt_line", "title": "WT GT / Paint Line", "unit": "WT Sec / Line", "num": "WT GT", "den": "Paint Line"},
]
MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_num(value):
    if pd.isna(value):
        return 0.0
    try:
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if value in ("", "-", "#N/A"):
                return 0.0
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return 0.0
        return number
    except Exception:
        return 0.0


def read_source():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    errors = []
    for sheet in ["GitHub_Export", "Raw(1)"]:
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, engine="openpyxl")
            df.columns = [clean_text(c) for c in df.columns]
            required = ["Year", "Week", "Factory", "Type", "Category", "Q'ty"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                raise ValueError(f"missing columns {missing}")
            print(f"Using sheet: {sheet}, rows={len(df)}")
            return prepare_df(df)
        except Exception as exc:
            errors.append(f"{sheet}: {exc}")

    raise RuntimeError("Unable to read source Excel. " + " | ".join(errors))


def prepare_df(df):
    keep = ["Year", "Week", "Factory", "Type", "Category", "Q'ty"]
    for col in ["Month", "YearMonth", "Month_Sort"]:
        if col in df.columns:
            keep.append(col)
    df = df[keep].copy()
    df = df.dropna(subset=["Year", "Week", "Factory", "Type", "Category"], how="any")
    df["Year"] = df["Year"].apply(lambda x: int(to_num(x)) if to_num(x) else 0)
    df["Week"] = df["Week"].apply(lambda x: int(to_num(x)) if to_num(x) else 0)
    df["Factory"] = df["Factory"].apply(clean_text)
    df["Type"] = df["Type"].apply(clean_text)
    df["Category"] = df["Category"].apply(clean_text)
    df["Q'ty"] = df["Q'ty"].apply(to_num)
    df = df[(df["Year"] > 0) & (df["Week"] > 0) & (df["Factory"] != "")]

    if "Month" not in df.columns or df["Month"].isna().all():
        df["Month"] = df["Week"].apply(lambda w: MONTH_ORDER[min(max(int((w - 1) / 4), 0), 11)])
    else:
        df["Month"] = df["Month"].apply(clean_text)

    if "YearMonth" not in df.columns or df["YearMonth"].isna().all():
        df["YearMonth"] = df.apply(lambda r: f"{r['Year']}-{r['Month']}", axis=1)
    else:
        df["YearMonth"] = df["YearMonth"].apply(clean_text)

    if "Month_Sort" not in df.columns or df["Month_Sort"].isna().all():
        month_index = {m: i + 1 for i, m in enumerate(MONTH_ORDER)}
        df["Month_Sort"] = df.apply(lambda r: r["Year"] * 100 + month_index.get(r["Month"], 0), axis=1)
    else:
        df["Month_Sort"] = df["Month_Sort"].apply(to_num)
    return df


def get_factories(df):
    found = list(dict.fromkeys(df["Factory"].tolist()))
    ordered = [f for f in FACTORY_ORDER if f in found]
    ordered.extend(sorted([f for f in found if f not in ordered]))
    return ordered


def pivot(df, type_name, categories, factories):
    result = {f: {c: 0.0 for c in categories} for f in factories}
    sub = df[df["Type"].str.lower() == type_name.lower()]
    grouped = sub.groupby(["Factory", "Category"], dropna=False)["Q'ty"].sum()
    for (factory, category), value in grouped.items():
        factory = clean_text(factory)
        category = clean_text(category)
        if factory in result and category in categories:
            result[factory][category] = round(float(value), 4)
    return result


def productivity(machine, perf, factories):
    result = {f: {} for f in factories}
    for f in factories:
        for m in METRICS:
            den = machine.get(f, {}).get(m["den"], 0.0)
            num = perf.get(f, {}).get(m["num"], 0.0)
            result[f][m["key"]] = round(num / den, 4) if den else 0.0
    return result


def build_period(df, factories):
    machine = pivot(df, "Machine", MACHINE_CATEGORIES, factories)
    perf = pivot(df, "Performance", PERFORMANCE_CATEGORIES, factories)
    prod = productivity(machine, perf, factories)
    return {"machine": machine, "performance": perf, "productivity": prod}


def build_data(df):
    factories = get_factories(df)
    years = sorted(df["Year"].unique().tolist())
    weekly = {}
    monthly = {}
    weeks_by_year = {}
    months_by_year = {}

    for year in years:
        ykey = f"{int(year)}Y"
        ydf = df[df["Year"] == year]
        weekly[ykey] = {}
        for week in sorted(ydf["Week"].unique().tolist()):
            weekly[ykey][f"WK{int(week):02d}"] = build_period(ydf[ydf["Week"] == week], factories)
        weeks_by_year[ykey] = list(weekly[ykey].keys())

        monthly[ykey] = {}
        months = ydf[["YearMonth", "Month_Sort"]].drop_duplicates().sort_values("Month_Sort")
        for ym in months["YearMonth"].tolist():
            monthly[ykey][str(ym)] = build_period(ydf[ydf["YearMonth"] == ym], factories)
        months_by_year[ykey] = list(monthly[ykey].keys())

    default_year = f"{int(years[-1])}Y" if years else ""

    def last_nonzero(periods):
        last = ""
        for key, value in periods.items():
            total = 0.0
            for fdata in value["performance"].values():
                total += sum(float(v or 0) for v in fdata.values())
            if total > 0:
                last = key
        return last or (list(periods.keys())[-1] if periods else "")

    return {
        "factories": factories,
        "years": [f"{int(y)}Y" for y in years],
        "weeksByYear": weeks_by_year,
        "monthsByYear": months_by_year,
        "weekly": weekly,
        "monthly": monthly,
        "defaultYear": default_year,
        "defaultWeek": last_nonzero(weekly.get(default_year, {})),
        "defaultMonth": last_nonzero(monthly.get(default_year, {})),
        "metrics": METRICS,
        "machineCategories": MACHINE_CATEGORIES,
        "performanceCategories": PERFORMANCE_CATEGORIES,
    }


def make_html(data):
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CS WIND Global Equipment-Based Productivity Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{--navy:#1f4e78;--bg:#eef2f6;--line:#d7e0ea;--card:#fff;--text:#111827;}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font-family:Arial,'Malgun Gothic',sans-serif}.wrap{max-width:1480px;margin:0 auto;background:#fff;min-height:100vh}.header{background:var(--navy);color:#fff;padding:28px;text-align:center}.header h1{margin:0;font-size:28px;letter-spacing:.5px;font-weight:800}.controls{display:flex;gap:14px;align-items:center;padding:16px 24px;border-bottom:1px solid var(--line);background:#f8fafc;flex-wrap:wrap}.controls label{font-weight:700;font-size:13px}.btn-group{display:inline-flex;border:1px solid #b9c7d6;border-radius:4px;overflow:hidden;background:#fff}.btn-group button{min-width:72px;padding:8px 16px;border:0;background:#fff;cursor:pointer;font-weight:700}.btn-group button.active{background:var(--navy);color:#fff}select{min-width:96px;padding:8px 12px;border:1px solid #b9c7d6;border-radius:4px;background:#fff;font-weight:700}.content{padding:18px 24px 40px}.section-title{background:var(--navy);color:#fff;padding:12px 16px;border-radius:5px;font-weight:800;margin:12px 0 16px;font-size:17px}table{width:100%;border-collapse:collapse;margin-bottom:22px;font-size:12px;background:#fff}th{background:#315d84;color:#fff;padding:10px 8px;border:1px solid #d9e2ec;text-align:center}td{padding:8px;border:1px solid #d9e2ec;text-align:center}td:first-child,th:first-child{text-align:left;font-weight:700}tfoot td{background:#fff1b8;font-weight:800}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-bottom:28px}.chart-card{background:var(--card);border:1px solid #cfd9e5;border-radius:7px;padding:12px 14px;min-height:280px;box-shadow:0 1px 2px rgba(0,0,0,.04)}.chart-title{text-align:center;font-weight:800;margin:4px 0 0;font-size:13px}.chart-unit{text-align:right;font-size:10px;color:#475569;margin-bottom:4px}.chart-card canvas{width:100%!important;height:220px!important}.trend-title{background:var(--navy);color:#fff;padding:10px 14px;border-radius:5px;font-weight:800;margin-top:18px}.factory-box{border:1px solid #cfd9e5;background:#f8fafc;border-radius:7px;padding:12px;margin:12px 0 18px}.factory-grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:8px}.factory-grid label{border:1px solid #89a7c5;border-radius:4px;padding:5px 8px;font-size:12px;background:#fff;font-weight:700}.small-note{color:#64748b;font-size:11px;margin-top:6px}@media(max-width:900px){.grid{grid-template-columns:1fr}.factory-grid{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="wrap">
<div class="header"><h1>CS WIND Global Equipment-Based Productivity Dashboard</h1></div>
<div class="controls"><label>View:</label><div class="btn-group"><button id="weeklyBtn" onclick="setView('weekly')">Weekly</button><button id="monthlyBtn" onclick="setView('monthly')">Monthly</button></div><label>Year:</label><select id="yearSelect" onchange="onYearChange()"></select><label id="periodLabel">Week:</label><select id="periodSelect" onchange="renderAll()"></select></div>
<div class="content"><div class="section-title">1. 공장별 장비 현황 (Equipment Inventory)</div><div id="machineTable"></div><div class="section-title">2. 공장별 생산 실적 (Production Performance)</div><div id="performanceTable"></div><div class="section-title">3. 장비당 생산 효율성 (Production per Equipment)</div><div class="grid" id="barCharts"></div><div class="trend-title">Trend of Equipment-Based Productivity (최근 10주)</div><div class="factory-box"><b>공장 선택 (Factory Selection)</b><div class="factory-grid" id="factoryChecks"></div><div class="small-note">체크한 공장만 하단 Trend 라인그래프에 표시됩니다.</div></div><div class="grid" id="trendCharts"></div></div>
</div>
<script>
const DASHBOARD_DATA = __DATA_JSON__;
let currentView='weekly';let chartObjects=[];const palette=['#2563eb','#38bdf8','#0f766e','#dc2626','#f97316','#8b5cf6','#0891b2','#64748b','#111827','#84cc16'];
function fmt(n,d=2){const v=Number(n||0);return v===0?'0':v.toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d});}
function destroyCharts(){chartObjects.forEach(c=>c.destroy());chartObjects=[];}
function getYear(){return document.getElementById('yearSelect').value;}function getPeriod(){return document.getElementById('periodSelect').value;}
function getPeriodData(){const y=getYear(),p=getPeriod();return (DASHBOARD_DATA[currentView]&&DASHBOARD_DATA[currentView][y]&&DASHBOARD_DATA[currentView][y][p])||null;}
function setView(view){currentView=view;document.getElementById('weeklyBtn').classList.toggle('active',view==='weekly');document.getElementById('monthlyBtn').classList.toggle('active',view==='monthly');document.getElementById('periodLabel').textContent=view==='weekly'?'Week:':'Month:';populatePeriods();renderAll();}
function populateYears(){const sel=document.getElementById('yearSelect');sel.innerHTML='';DASHBOARD_DATA.years.forEach(y=>{const o=document.createElement('option');o.value=y;o.textContent=y;sel.appendChild(o);});if(DASHBOARD_DATA.defaultYear)sel.value=DASHBOARD_DATA.defaultYear;}
function populatePeriods(){const y=getYear();const list=currentView==='weekly'?(DASHBOARD_DATA.weeksByYear[y]||[]):(DASHBOARD_DATA.monthsByYear[y]||[]);const sel=document.getElementById('periodSelect');sel.innerHTML='';list.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o);});const def=currentView==='weekly'?DASHBOARD_DATA.defaultWeek:DASHBOARD_DATA.defaultMonth;if(list.includes(def))sel.value=def;}
function onYearChange(){populatePeriods();renderAll();}
function tableHtml(cols,rows,total){let html='<table><thead><tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr></thead><tbody>';rows.forEach(r=>{html+='<tr>'+r.map(c=>`<td>${c}</td>`).join('')+'</tr>';});html+='</tbody>';if(total&&rows.length){const totals=['CS Total'];for(let i=1;i<cols.length-1;i++){let s=0;rows.forEach(r=>{s+=Number(String(r[i]).replace(/,/g,''))||0;});totals.push(fmt(s,2));}totals.push('');html+='<tfoot><tr>'+totals.map(c=>`<td>${c}</td>`).join('')+'</tr></tfoot>';}return html+'</table>';}
function renderTables(data){const f=DASHBOARD_DATA.factories;const mcols=['Factory'].concat(DASHBOARD_DATA.machineCategories).concat(['Remark']);const mrows=f.map(x=>[x].concat(DASHBOARD_DATA.machineCategories.map(c=>fmt(data.machine[x]?.[c],2))).concat(['']));document.getElementById('machineTable').innerHTML=tableHtml(mcols,mrows,true);const pcols=['Factory'].concat(DASHBOARD_DATA.performanceCategories).concat(['Remark']);const prows=f.map(x=>[x].concat(DASHBOARD_DATA.performanceCategories.map(c=>fmt(data.performance[x]?.[c],2))).concat(['']));document.getElementById('performanceTable').innerHTML=tableHtml(pcols,prows,false);}
function makeChart(canvas,config){const c=new Chart(canvas,config);chartObjects.push(c);return c;}
function renderBarCharts(data){const box=document.getElementById('barCharts');box.innerHTML='';const labels=DASHBOARD_DATA.factories.concat(['CS Total']);DASHBOARD_DATA.metrics.forEach(m=>{const card=document.createElement('div');card.className='chart-card';card.innerHTML=`<div class="chart-title">${m.title}</div><div class="chart-unit">(${m.unit})</div><canvas></canvas>`;box.appendChild(card);const vals=DASHBOARD_DATA.factories.map(f=>Number(data.productivity[f]?.[m.key]||0));const nz=vals.filter(v=>v);const total=nz.length?nz.reduce((a,b)=>a+b,0)/nz.length:0;makeChart(card.querySelector('canvas'),{type:'bar',data:{labels:labels,datasets:[{data:vals.concat([Number(total.toFixed(2))]),backgroundColor:labels.map((_,i)=>palette[i%palette.length]),borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'#e5e7eb'}},x:{grid:{display:false},ticks:{font:{size:10}}}}}});});}
function renderFactoryChecks(){const box=document.getElementById('factoryChecks');if(box.children.length)return;DASHBOARD_DATA.factories.concat(['CS Total']).forEach(f=>{const label=document.createElement('label');label.innerHTML=`<input type="checkbox" value="${f}" checked onchange="renderTrendCharts()"> ${f}`;box.appendChild(label);});}
function selectedFactories(){return Array.from(document.querySelectorAll('#factoryChecks input:checked')).map(i=>i.value);}
function recentWeeklyKeys(){const y=getYear();const list=DASHBOARD_DATA.weeksByYear[y]||[];const current=getPeriod();let idx=list.indexOf(current);if(idx<0)idx=list.length-1;return list.slice(Math.max(0,idx-9),idx+1);}
function renderTrendCharts(){const box=document.getElementById('trendCharts');box.innerHTML='';const y=getYear();const weeks=recentWeeklyKeys();const selected=selectedFactories();DASHBOARD_DATA.metrics.forEach(m=>{const card=document.createElement('div');card.className='chart-card';card.innerHTML=`<div class="chart-title">${m.title}</div><div class="chart-unit">(${m.unit})</div><canvas></canvas>`;box.appendChild(card);const datasets=selected.map((f,i)=>({label:f,data:weeks.map(w=>{const d=DASHBOARD_DATA.weekly[y]?.[w];if(!d)return 0;if(f==='CS Total'){const vals=DASHBOARD_DATA.factories.map(ff=>Number(d.productivity[ff]?.[m.key]||0)).filter(v=>v);return vals.length?Number((vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(2)):0;}return Number(d.productivity[f]?.[m.key]||0);}),borderColor:palette[i%palette.length],backgroundColor:palette[i%palette.length],tension:.25,pointRadius:2,fill:false}));makeChart(card.querySelector('canvas'),{type:'line',data:{labels:weeks,datasets:datasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,position:'bottom',labels:{boxWidth:10,font:{size:10}}}},scales:{y:{beginAtZero:true,grid:{color:'#e5e7eb'}},x:{grid:{display:false},ticks:{font:{size:10}}}}}});});}
function renderAll(){destroyCharts();const data=getPeriodData();if(!data){document.getElementById('machineTable').innerHTML='<p>No data.</p>';document.getElementById('performanceTable').innerHTML='<p>No data.</p>';document.getElementById('barCharts').innerHTML='';document.getElementById('trendCharts').innerHTML='';return;}renderTables(data);renderBarCharts(data);renderFactoryChecks();renderTrendCharts();}
populateYears();setView('weekly');
</script>
</body>
</html>
'''.replace("__DATA_JSON__", data_json)


def main():
    df = read_source()
    data = build_data(df)
    html = make_html(data)
    INDEX_FILE.write_text(html, encoding="utf-8")
    TEMPLATE_FILE.write_text(html, encoding="utf-8")
    print(f"Generated {INDEX_FILE.name} and {TEMPLATE_FILE.name}")
    print("Years:", ", ".join(data.get("years", [])))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
