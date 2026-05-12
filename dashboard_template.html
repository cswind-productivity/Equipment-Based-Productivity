<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>CS WIND Global Equipment-Based Productivity Dashboard</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root { --bg:#f4f7fb; --card:#ffffff; --ink:#172033; --muted:#667085; --line:#d8e0ea; --blue:#2563eb; --green:#16a34a; --orange:#f59e0b; }
    * { box-sizing: border-box; }
    body { margin:0; font-family: Arial, 'Malgun Gothic', sans-serif; background:var(--bg); color:var(--ink); }
    .wrap { max-width: 1480px; margin: 0 auto; padding: 24px; }
    .hero { background: linear-gradient(135deg, #102a43, #1d4ed8); color:white; border-radius: 18px; padding: 24px 28px; box-shadow:0 12px 28px rgba(16,42,67,.18); }
    h1 { margin:0; font-size: 28px; letter-spacing:-.3px; }
    .subtitle { margin-top:8px; color:#dbeafe; font-size:14px; }
    .controls { display:grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap:14px; margin-top:18px; }
    .control { background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.2); border-radius:14px; padding:10px 12px; }
    .control label { display:block; font-size:12px; color:#dbeafe; margin-bottom:6px; }
    select { width:100%; padding:9px 10px; border:1px solid #cbd5e1; border-radius:10px; background:white; color:#111827; font-size:14px; }
    .grid { display:grid; grid-template-columns: 1fr; gap:18px; margin-top:20px; }
    .card { background:var(--card); border:1px solid var(--line); border-radius:18px; padding:18px; box-shadow:0 8px 18px rgba(15,23,42,.06); }
    .card h2 { margin:0 0 14px; font-size:18px; }
    .kpis { display:grid; grid-template-columns: repeat(4, 1fr); gap:14px; margin-top:20px; }
    .kpi { background:var(--card); border:1px solid var(--line); border-radius:18px; padding:18px; box-shadow:0 8px 18px rgba(15,23,42,.06); }
    .kpi .label { color:var(--muted); font-size:13px; }
    .kpi .value { margin-top:7px; font-size:27px; font-weight:700; }
    table { width:100%; border-collapse:collapse; font-size:13px; }
    th, td { border-bottom:1px solid #eef2f7; padding:9px 10px; text-align:right; white-space:nowrap; }
    th:first-child, td:first-child { text-align:left; font-weight:600; }
    th { background:#f8fafc; color:#344054; font-size:12px; }
    tbody tr:hover { background:#f8fbff; }
    .two { display:grid; grid-template-columns: 1fr 1fr; gap:18px; }
    .chartBox { height:330px; }
    .foot { text-align:center; color:var(--muted); font-size:12px; padding:22px 0 6px; }
    @media(max-width: 1000px){ .controls,.kpis,.two{grid-template-columns:1fr;} .wrap{padding:14px;} }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <h1>CS WIND Global Equipment-Based Productivity Dashboard</h1>
      <div class="subtitle">Weekly / Monthly view generated automatically from Excel Raw(1) data</div>
      <div class="controls">
        <div class="control"><label>View Mode</label><select id="modeSelect"><option value="weekly">Weekly</option><option value="monthly">Monthly</option></select></div>
        <div class="control"><label>Year</label><select id="yearSelect"></select></div>
        <div class="control"><label>Period</label><select id="periodSelect"></select></div>
        <div class="control"><label>Factory for Trend</label><select id="factorySelect"></select></div>
      </div>
    </div>

    <div class="kpis">
      <div class="kpi"><div class="label">Selected Period</div><div class="value" id="kpiPeriod">-</div></div>
      <div class="kpi"><div class="label">Total Equipment</div><div class="value" id="kpiEquipment">-</div></div>
      <div class="kpi"><div class="label">Total Production</div><div class="value" id="kpiProduction">-</div></div>
      <div class="kpi"><div class="label">Avg. Productivity</div><div class="value" id="kpiProductivity">-</div></div>
    </div>

    <div class="grid">
      <div class="card"><h2>1. 공장별 장비 현황 (Equipment Inventory)</h2><div id="equipmentTable"></div></div>
      <div class="card"><h2>2. 공장별 생산 실적 (Production Performance)</h2><div id="productionTable"></div></div>
      <div class="card"><h2>3. 장비당 생산 효율성 (Production per Equipment)</h2><div id="efficiencyTable"></div></div>
      <div class="two">
        <div class="card"><h2>Factory Productivity Comparison</h2><div class="chartBox"><canvas id="barChart"></canvas></div></div>
        <div class="card"><h2>Trend of Equipment-Based Productivity</h2><div class="chartBox"><canvas id="trendChart"></canvas></div></div>
      </div>
    </div>
    <div class="foot">Made by SJ Lee · Monthly view added for CS WIND Productivity Dashboard</div>
  </div>

<script>
%%RAWDATA_PLACEHOLDER%%

const factoryOrder = ['VN #1', 'VN #2', 'TW', 'CN', 'TR #1', 'TR #2', 'AM', 'PT On', 'PT Off'];
const metricLabels = {
  bending: 'Bending / Machine', lw: 'L/W / Machine', cw: 'C/W / Machine',
  btgt: 'BT GT / Growing Line', wtgtBooth: 'WT GT / Paint Booth', wtgtLine: 'WT GT / Paint Line'
};
let barChart, trendChart;

function fmt(n, digits=0){ return Number(n || 0).toLocaleString(undefined, {maximumFractionDigits:digits, minimumFractionDigits:digits}); }
function activeData(){ return document.getElementById('modeSelect').value === 'monthly' ? monthlyData : rawData; }
function getYears(){ return Object.keys(activeData()).sort(); }
function periodSortKey(p){
  if(p.startsWith('WK')) return Number(p.replace('WK',''));
  const [y,m] = p.split('-'); const months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
  return Number(y)*100 + (months[m] || 0);
}
function getPeriods(year){ return Object.keys(activeData()[year] || {}).sort((a,b)=>periodSortKey(a)-periodSortKey(b)); }
function selectedPayload(){ const y=yearSelect.value, p=periodSelect.value; return (activeData()[y] || {})[p] || {equipment:{}, production:{}}; }
function factoriesIn(payload){ return factoryOrder.filter(f => payload.equipment[f] || payload.production[f]).concat(Object.keys(payload.equipment).filter(f => !factoryOrder.includes(f))); }
function sumObj(obj){ return Object.values(obj || {}).reduce((a,b)=>a+Number(b||0),0); }
function calcEff(eq, prod){
  return {
    bending: (prod.bending||0)/(eq.bending||0), lw: (prod.lw||0)/(eq.lw||0), cw: (prod.cw||0)/(eq.cw||0),
    btgt: (prod.btgt||0)/(eq.growing||0), wtgtBooth: (prod.wtgt||0)/(eq.paintBooth||0), wtgtLine: (prod.wtgt||0)/(eq.paintLine||0)
  };
}
function makeTable(headers, rows){
  let html = '<table><thead><tr>' + headers.map(h=>`<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  html += rows.map(r=>'<tr>'+r.map(c=>`<td>${c}</td>`).join('')+'</tr>').join('');
  return html + '</tbody></table>';
}
function fillControls(){
  const years = getYears();
  yearSelect.innerHTML = years.map(y=>`<option value="${y}">${y}</option>`).join('');
  yearSelect.value = years[years.length-1] || '';
  fillPeriods();
}
function fillPeriods(){
  const periods = getPeriods(yearSelect.value);
  periodSelect.innerHTML = periods.map(p=>`<option value="${p}">${p}</option>`).join('');
  periodSelect.value = periods[periods.length-1] || '';
  fillFactories();
}
function fillFactories(){
  const payload = selectedPayload(); const factories = factoriesIn(payload);
  const current = factorySelect.value;
  factorySelect.innerHTML = factories.map(f=>`<option value="${f}">${f}</option>`).join('');
  factorySelect.value = factories.includes(current) ? current : (factories[0] || '');
}
function renderTables(){
  const payload = selectedPayload(); const factories = factoriesIn(payload);
  const eqHeaders = ['Factory','Roll Bending Machine','L/W Machine','C/W Machine','Growing Line','Paint Booth','Paint Line'];
  const prodHeaders = ['Factory','Bending','L/W','C/W','BT GT','WT GT'];
  const effHeaders = ['Factory','Bending / Machine','L/W / Machine','C/W / Machine','BT GT / Growing Line','WT GT / Paint Booth','WT GT / Paint Line'];
  let totalEq=0, totalProd=0, effVals=[];
  const eqRows=[], prodRows=[], effRows=[];
  factories.forEach(f=>{
    const e=payload.equipment[f]||{}, p=payload.production[f]||{}, eff=calcEff(e,p);
    totalEq += sumObj(e); totalProd += sumObj(p);
    Object.values(eff).forEach(v=>{ if(isFinite(v) && v>0) effVals.push(v); });
    eqRows.push([f, fmt(e.bending), fmt(e.lw), fmt(e.cw), fmt(e.growing), fmt(e.paintBooth), fmt(e.paintLine)]);
    prodRows.push([f, fmt(p.bending), fmt(p.lw), fmt(p.cw), fmt(p.btgt), fmt(p.wtgt)]);
    effRows.push([f, fmt(eff.bending,2), fmt(eff.lw,2), fmt(eff.cw,2), fmt(eff.btgt,2), fmt(eff.wtgtBooth,2), fmt(eff.wtgtLine,2)]);
  });
  equipmentTable.innerHTML = makeTable(eqHeaders, eqRows);
  productionTable.innerHTML = makeTable(prodHeaders, prodRows);
  efficiencyTable.innerHTML = makeTable(effHeaders, effRows);
  kpiPeriod.textContent = `${yearSelect.value} ${periodSelect.value}`;
  kpiEquipment.textContent = fmt(totalEq);
  kpiProduction.textContent = fmt(totalProd);
  kpiProductivity.textContent = effVals.length ? fmt(effVals.reduce((a,b)=>a+b,0)/effVals.length,2) : '-';
}
function renderCharts(){
  const payload=selectedPayload(), factories=factoriesIn(payload);
  const barData = factories.map(f=>{
    const e=payload.equipment[f]||{}, p=payload.production[f]||{};
    const v=(p.bending||0)/(e.bending||0);
    return isFinite(v) ? v : 0;
  });
  if(barChart) barChart.destroy();
  barChart = new Chart(document.getElementById('barChart'), {type:'bar', data:{labels:factories, datasets:[{label:'Bending / Machine', data:barData, borderWidth:1}]}, options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}}}});

  const mode = modeSelect.value; const data = activeData(); const y=yearSelect.value; const periods=getPeriods(y); const idx=periods.indexOf(periodSelect.value); const selectedFactory=factorySelect.value;
  const recent = periods.slice(Math.max(0, idx-9), idx+1);
  const trend = recent.map(p=>{ const pl=(data[y]||{})[p]||{equipment:{},production:{}}; const e=(pl.equipment||{})[selectedFactory]||{}, pr=(pl.production||{})[selectedFactory]||{}; const v=(pr.bending||0)/(e.bending||0); return isFinite(v)?v:0; });
  if(trendChart) trendChart.destroy();
  trendChart = new Chart(document.getElementById('trendChart'), {type:'line', data:{labels:recent, datasets:[{label:`${selectedFactory} Bending / Machine`, data:trend, tension:.25, fill:false}]}, options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:true}}, scales:{y:{beginAtZero:true}}}});
}
function refresh(){ fillFactories(); renderTables(); renderCharts(); }
modeSelect.addEventListener('change', ()=>{ fillControls(); refresh(); });
yearSelect.addEventListener('change', ()=>{ fillPeriods(); refresh(); });
periodSelect.addEventListener('change', refresh);
factorySelect.addEventListener('change', renderCharts);
fillControls(); refresh();
</script>
</body>
</html>
