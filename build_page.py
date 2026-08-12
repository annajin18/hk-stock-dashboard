#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""港股基石解禁提醒 - 网页生成脚本"""

import openpyxl
import json
from datetime import datetime, date
from collections import defaultdict

EXCEL_PATH = "D:/港股交易app/基石投资者0810.xlsx"
OUTPUT_PATH = "D:/港股交易app/index.html"
TODAY = date(2026, 8, 10)
USD_TO_HKD = 7.8


def process_data():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["基石投资者"]

    records = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue

        lift_date = row[15]
        if isinstance(lift_date, str):
            try:
                d = datetime.strptime(lift_date[:10], "%Y-%m-%d").date()
            except ValueError:
                continue
        elif lift_date:
            d = lift_date
        else:
            continue

        listing_date = row[2]
        if isinstance(listing_date, str):
            ld = listing_date[:10]
        elif listing_date:
            ld = listing_date.strftime("%Y-%m-%d")
        else:
            ld = ""

        days_to_lift = (d - TODAY).days
        if days_to_lift < 0:
            urgency = "past"
        elif days_to_lift <= 30:
            urgency = "soon30"
        elif days_to_lift <= 90:
            urgency = "soon90"
        else:
            urgency = "future"

        amount = row[11] or 0
        currency = row[12] or "HKD"
        amount_hkd = amount if currency == "HKD" else amount * USD_TO_HKD

        records.append({
            "code": row[0] or "",
            "name": row[1] or "",
            "listingDate": ld,
            "liftDate": d.strftime("%Y-%m-%d"),
            "investor": row[6] or "",
            "shares": row[10] or 0,
            "amount": amount,
            "currency": currency,
            "amountHKD": round(amount_hkd),
            "percentage": row[13] or 0,
            "lockupMonths": row[14] or 0,
            "industry": row[16] or "",
            "thsIndustry": row[17] or "",
            "daysToLift": days_to_lift,
            "urgency": urgency,
        })

    # Summary
    within30 = [r for r in records if r["urgency"] == "soon30"]
    within90 = [r for r in records if r["urgency"] in ("soon30", "soon90")]
    within30_stocks = len(set(r["code"] for r in within30))
    within90_stocks = len(set(r["code"] for r in within90))

    # Monthly aggregation
    monthly_map = defaultdict(lambda: {"count": 0, "stocks": set(), "amountHKD": 0})
    for r in records:
        month_key = r["liftDate"][:7]
        monthly_map[month_key]["count"] += 1
        monthly_map[month_key]["stocks"].add(r["code"])
        monthly_map[month_key]["amountHKD"] += r["amountHKD"]

    monthly = []
    for k in sorted(monthly_map.keys()):
        v = monthly_map[k]
        month_start = datetime.strptime(k + "-01", "%Y-%m-%d").date()
        days = (month_start - TODAY).days
        if days < -31:
            status = "past"
        elif days < 0:
            status = "current"
        elif days <= 30:
            status = "soon"
        else:
            status = "future"
        monthly.append({
            "month": k,
            "count": v["count"],
            "stocks": len(v["stocks"]),
            "amountHKD": round(v["amountHKD"]),
            "amountYi": round(v["amountHKD"] / 1e8, 2),
            "status": status,
        })

    # Stock aggregation
    stock_map = defaultdict(list)
    for r in records:
        stock_map[r["code"]].append(r)

    stocks = []
    for code, recs in stock_map.items():
        first = recs[0]
        total_amount_hkd = sum(r["amountHKD"] for r in recs)
        total_percentage = sum(r["percentage"] for r in recs)
        min_days = min(r["daysToLift"] for r in recs)
        if min_days < 0:
            urgency = "past"
        elif min_days <= 30:
            urgency = "soon30"
        elif min_days <= 90:
            urgency = "soon90"
        else:
            urgency = "future"

        sorted_recs = sorted(recs, key=lambda x: x["amountHKD"], reverse=True)
        stocks.append({
            "code": code,
            "name": first["name"],
            "listingDate": first["listingDate"],
            "liftDate": min(r["liftDate"] for r in recs),
            "industry": first["industry"],
            "thsIndustry": first["thsIndustry"],
            "investorCount": len(recs),
            "totalAmountHKD": round(total_amount_hkd),
            "totalAmountYi": round(total_amount_hkd / 1e8, 2),
            "totalPercentage": round(total_percentage, 2),
            "urgency": urgency,
            "daysToLift": min_days,
            "investors": sorted_recs,
        })

    stocks.sort(key=lambda x: x["liftDate"])
    industries = sorted(set(r["industry"] for r in records if r["industry"]))

    return {
        "updateDate": TODAY.strftime("%Y-%m-%d"),
        "totalStocks": len(stocks),
        "totalRecords": len(records),
        "summary": {
            "within30": {"count": len(within30), "stocks": within30_stocks},
            "within90": {"count": len(within90), "stocks": within90_stocks},
        },
        "monthly": monthly,
        "stocks": stocks,
        "industries": industries,
    }


HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>港股提醒</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#141413;color:#faf9f5;font-family:'Lora',Georgia,'Noto Serif SC',serif;min-height:100vh}
h1,h2,h3,.tab,.card-value,.card-label,.filter-bar,.data-table th,.coming-soon{font-family:'Poppins',Arial,'Noto Sans SC',sans-serif}

/* Tab Nav */
.tab-nav{display:flex;background:#1f1e1c;border-bottom:1px solid #2a2926;padding:0 24px;position:sticky;top:0;z-index:100;overflow-x:auto}
.tab{padding:16px 22px;background:transparent;border:none;color:#b0aea5;font-size:14px;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap}
.tab:hover{color:#faf9f5}
.tab.active{color:#faf9f5;border-bottom-color:#d97757;font-weight:600}

/* Tab Content */
.tab-content{display:none;padding:24px;max-width:1440px;margin:0 auto}
.tab-content.active{display:block}

/* Header */
.page-header{margin-bottom:20px}
.page-header h1{font-size:26px;font-weight:600;color:#faf9f5;margin-bottom:4px}
.page-header p{color:#b0aea5;font-size:14px}

/* Cards */
.cards-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}
.card{background:#1f1e1c;border:1px solid #2a2926;border-radius:12px;padding:20px;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:4px}
.card.urgent::before{background:#d97757}
.card.info::before{background:#6a9bcc}
.card.neutral::before{background:#b0aea5}
.card-label{font-size:13px;color:#b0aea5;margin-bottom:8px}
.card-value{font-size:30px;font-weight:700;color:#faf9f5}
.card-sub{font-size:12px;color:#b0aea5;margin-top:4px}

/* Chart */
.chart-wrap{background:#1f1e1c;border:1px solid #2a2926;border-radius:12px;padding:24px;margin-bottom:24px}
.chart-wrap h3{font-size:16px;color:#faf9f5;margin-bottom:16px}
.chart-canvas-box{height:360px;position:relative}

/* Filter Bar */
.filter-bar{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.filter-bar input,.filter-bar select{background:#1f1e1c;border:1px solid #2a2926;border-radius:8px;padding:9px 14px;color:#faf9f5;font-size:14px;font-family:'Poppins',Arial,sans-serif}
.filter-bar input{flex:1;min-width:200px}
.filter-bar input:focus,.filter-bar select:focus{outline:none;border-color:#d97757}
.filter-bar select{cursor:pointer}
.view-toggle{display:flex;background:#1f1e1c;border:1px solid #2a2926;border-radius:8px;overflow:hidden}
.view-toggle button{padding:9px 16px;background:transparent;border:none;color:#b0aea5;font-size:13px;cursor:pointer;font-family:'Poppins',Arial,sans-serif;transition:all .2s}
.view-toggle button.active{background:#d97757;color:#faf9f5;font-weight:600}
.filter-info{font-size:13px;color:#b0aea5;margin-left:auto}
.clear-filter{color:#d97757;cursor:pointer;font-size:13px;text-decoration:underline;border:none;background:none;font-family:inherit}

/* Table */
.table-wrap{background:#1f1e1c;border:1px solid #2a2926;border-radius:12px;overflow:auto}
.data-table{width:100%;border-collapse:collapse}
.data-table th{background:#25241f;padding:12px 16px;text-align:left;font-size:13px;color:#b0aea5;font-weight:500;cursor:pointer;position:sticky;top:0;white-space:nowrap;user-select:none;border-bottom:1px solid #2a2926}
.data-table th:hover{color:#faf9f5}
.data-table th .sort-arrow{font-size:10px;opacity:.5}
.data-table td{padding:11px 16px;border-bottom:1px solid #2a2926;font-size:14px;color:#faf9f5}
.data-table tr:hover{background:#25241f}
.data-table tr.expanded{background:#25241f}

/* Urgency row borders */
tr.row-soon30 td:first-child{border-left:3px solid #d97757}
tr.row-soon90 td:first-child{border-left:3px solid #6a9bcc}
tr.row-past td:first-child{border-left:3px solid #4a4944}
tr.row-future td:first-child{border-left:3px solid transparent}

/* Urgency badges */
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;font-family:'Poppins',Arial,sans-serif}
.badge-soon30{background:#d9775722;color:#d97757}
.badge-soon90{background:#6a9bcc22;color:#6a9bcc}
.badge-past{background:#4a4944;color:#b0aea5}
.badge-future{background:#788c5d22;color:#788c5d}

/* Expand row */
.expand-row td{padding:0;background:#141413}
.expand-content{padding:16px 24px}
.expand-content table{width:100%;border-collapse:collapse}
.expand-content th{font-size:12px;color:#b0aea5;padding:8px 12px;text-align:left;font-family:'Poppins',Arial,sans-serif;border-bottom:1px solid #2a2926}
.expand-content td{font-size:13px;color:#faf9f5;padding:8px 12px;border-bottom:1px solid #1f1e1c}
.expand-toggle{cursor:pointer;color:#d97757;font-size:14px;user-select:none}
.expand-toggle:hover{text-decoration:underline}

/* Coming Soon */
.coming-soon{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:55vh;color:#b0aea5}
.coming-soon .icon{font-size:48px;margin-bottom:16px;opacity:.5}
.coming-soon h2{font-size:22px;font-weight:500;margin-bottom:8px;color:#b0aea5}
.coming-soon p{font-size:14px}

/* Scrollbar */
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:#1f1e1c}
::-webkit-scrollbar-thumb{background:#2a2926;border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:#3a3935}

/* Responsive */
@media(max-width:900px){
  .cards-row{grid-template-columns:repeat(2,1fr)}
  .tab{padding:14px 14px;font-size:13px}
}
@media(max-width:600px){
  .cards-row{grid-template-columns:1fr}
  .filter-bar{flex-direction:column}
  .filter-bar input{width:100%}
}
</style>
</head>
<body>

<nav class="tab-nav">
  <button class="tab active" onclick="switchTab(1)">港股基石解禁提醒</button>
  <button class="tab" onclick="switchTab(2)">港股打新提醒</button>
  <button class="tab" onclick="switchTab(3)">港股绿鞋使用情况</button>
  <button class="tab" onclick="switchTab(4)">港股通冲刺</button>
  <button class="tab" onclick="switchTab(5)">CRS税务</button>
</nav>

<!-- Tab 1: 港股基石解禁提醒 -->
<div id="tab-1" class="tab-content active">
  <div class="page-header">
    <h1>港股基石解禁提醒</h1>
    <p id="header-sub"></p>
  </div>

  <div class="cards-row" id="cards-row"></div>

  <div class="chart-wrap">
    <h3>月度解禁分布</h3>
    <div class="chart-canvas-box"><canvas id="monthlyChart"></canvas></div>
  </div>

  <div class="filter-bar">
    <input type="text" id="searchInput" placeholder="搜索股票代码 / 名称 / 投资者..." oninput="onFilterChange()">
    <select id="timeRange" onchange="onFilterChange()">
      <option value="all">全部时间</option>
      <option value="30" selected>30天内</option>
      <option value="90">90天内</option>
      <option value="future">未解禁</option>
      <option value="past">已解禁</option>
    </select>
    <select id="industryFilter" onchange="onFilterChange()">
      <option value="">全部行业</option>
    </select>
    <div class="view-toggle">
      <button id="view-detail" class="active" onclick="setView('detail')">明细视图</button>
      <button id="view-summary" onclick="setView('summary')">股票汇总</button>
    </div>
    <span class="filter-info" id="filter-info"></span>
  </div>

  <div class="table-wrap">
    <table class="data-table" id="dataTable">
      <thead id="table-head"></thead>
      <tbody id="table-body"></tbody>
    </table>
  </div>
</div>

<!-- Tab 2-5: Coming Soon -->
<div id="tab-2" class="tab-content">
  <div class="coming-soon"><div class="icon">&#128640;</div><h2>港股打新提醒</h2><p>即将上线</p></div>
</div>
<div id="tab-3" class="tab-content">
  <div class="coming-soon"><div class="icon">&#128640;</div><h2>港股绿鞋使用情况</h2><p>即将上线</p></div>
</div>
<div id="tab-4" class="tab-content">
  <div class="coming-soon"><div class="icon">&#128640;</div><h2>港股通冲刺</h2><p>即将上线</p></div>
</div>
<div id="tab-5" class="tab-content">
  <div class="coming-soon"><div class="icon">&#128640;</div><h2>CRS税务</h2><p>即将上线</p></div>
</div>

<script>
const DATA = __DATA_JSON__;

// State
let currentView = 'detail';
let sortColumn = 'liftDate';
let sortDir = 'asc';
let selectedMonth = null;
let chart = null;
let expandedStocks = new Set();

// Utils
function fmtNum(n) {
  if (n >= 1e8) return (n / 1e8).toFixed(2) + '亿';
  if (n >= 1e4) return (n / 1e4).toFixed(2) + '万';
  return n.toLocaleString();
}
function fmtAmount(r) {
  if (r.currency === 'USD') return '$' + (r.amount / 1e6).toFixed(2) + 'M';
  return (r.amount / 1e8).toFixed(2) + '亿港元';
}
function fmtShares(n) { return n.toLocaleString(); }
function fmtPct(n) { return n.toFixed(2) + '%'; }
function urgencyBadge(u) {
  const map = {
    'soon30': '<span class="badge badge-soon30">30天内</span>',
    'soon90': '<span class="badge badge-soon90">90天内</span>',
    'past': '<span class="badge badge-past">已解禁</span>',
    'future': '<span class="badge badge-future">未到期</span>'
  };
  return map[u] || '';
}
function rowClass(u) { return 'row-' + u; }

// Tab switching
function switchTab(idx) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    t.classList.toggle('active', i === idx - 1);
  });
  document.querySelectorAll('.tab-content').forEach((c, i) => {
    c.classList.toggle('active', i === idx - 1);
  });
}

// Render header
function renderHeader() {
  document.getElementById('header-sub').textContent =
    '数据截至 ' + DATA.updateDate + ' · 共监控' + DATA.totalStocks + '只港股 · ' + DATA.totalRecords + '条基石投资者记录';
}

// Render summary cards
function renderCards() {
  const s = DATA.summary;
  const cards = [
    {cls:'urgent', label:'30天内解禁', value:s.within30.count+'笔', sub:'涉及'+s.within30.stocks+'只股票'},
    {cls:'info', label:'90天内解禁', value:s.within90.count+'笔', sub:'涉及'+s.within90.stocks+'只股票'},
    {cls:'neutral', label:'监控股票', value:DATA.totalStocks+'只', sub:''},
    {cls:'neutral', label:'基石投资者记录', value:DATA.totalRecords+'条', sub:''}
  ];
  document.getElementById('cards-row').innerHTML = cards.map(c =>
    '<div class="card '+c.cls+'"><div class="card-label">'+c.label+'</div><div class="card-value">'+c.value+'</div>' +
    (c.sub?'<div class="card-sub">'+c.sub+'</div>':'') + '</div>'
  ).join('');
}

// Render chart
function renderChart() {
  const ctx = document.getElementById('monthlyChart').getContext('2d');
  const labels = DATA.monthly.map(m => m.month);
  const counts = DATA.monthly.map(m => m.count);
  const amounts = DATA.monthly.map(m => m.amountYi);
  const colors = DATA.monthly.map(m => {
    if (m.status === 'past') return '#4a4944';
    if (m.status === 'current') return '#b0aea5';
    if (m.status === 'soon') return '#d97757';
    return '#6a9bcc';
  });
  const borderColors = DATA.monthly.map(m => {
    if (selectedMonth === m.month) return '#faf9f5';
    return colors[m.status === 'past' ? 0 : 1];
  });

  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: '解禁笔数',
          data: counts,
          backgroundColor: colors,
          borderColor: colors.map(c => c),
          borderWidth: 1,
          borderRadius: 4,
          yAxisID: 'y',
        },
        {
          label: '解禁金额(亿港元)',
          data: amounts,
          type: 'line',
          borderColor: '#788c5d',
          backgroundColor: '#788c5d22',
          borderWidth: 2,
          pointRadius: 4,
          pointBackgroundColor: '#788c5d',
          yAxisID: 'y1',
          tension: 0.3,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      onClick: (e, els) => {
        if (els.length > 0) {
          const idx = els[0].index;
          const m = DATA.monthly[idx];
          if (selectedMonth === m.month) {
            selectedMonth = null;
          } else {
            selectedMonth = m.month;
          }
          renderChart();
          renderTable();
        }
      },
      plugins: {
        legend: { labels: { color: '#b0aea5', font: { family: 'Poppins, Arial, sans-serif', size: 12 } } },
        tooltip: {
          backgroundColor: '#1f1e1c',
          titleColor: '#faf9f5',
          bodyColor: '#faf9f5',
          borderColor: '#2a2926',
          borderWidth: 1,
          callbacks: {
            afterLabel: function(ctx) {
              const m = DATA.monthly[ctx.dataIndex];
              return m.stocks + '只股票';
            }
          }
        }
      },
      scales: {
        x: { ticks: { color: '#b0aea5', font: { size: 11 } }, grid: { color: '#2a2926' } },
        y: { beginAtZero: true, ticks: { color: '#b0aea5' }, grid: { color: '#2a2926' }, title: { display: true, text: '笔数', color: '#b0aea5' } },
        y1: { beginAtZero: true, position: 'right', ticks: { color: '#788c5d', callback: v => v + '亿' }, grid: { drawOnChartArea: false }, title: { display: true, text: '亿港元', color: '#788c5d' } }
      }
    }
  });
}

// Render industry filter
function renderIndustryFilter() {
  const sel = document.getElementById('industryFilter');
  DATA.industries.forEach(ind => {
    const opt = document.createElement('option');
    opt.value = ind;
    opt.textContent = ind;
    sel.appendChild(opt);
  });
}

// Get filtered records (flat list)
function getFilteredRecords() {
  const search = document.getElementById('searchInput').value.toLowerCase().trim();
  const timeRange = document.getElementById('timeRange').value;
  const industry = document.getElementById('industryFilter').value;

  let recs = [];
  DATA.stocks.forEach(s => {
    s.investors.forEach(inv => {
      recs.push({...inv, stockName: s.name, stockIndustry: s.industry});
    });
  });

  recs = recs.filter(r => {
    if (search && !r.code.toLowerCase().includes(search) && !r.stockName.toLowerCase().includes(search) && !r.investor.toLowerCase().includes(search)) return false;
    if (industry && r.stockIndustry !== industry) return false;
    if (selectedMonth && r.liftDate.substring(0, 7) !== selectedMonth) return false;
    if (timeRange === '30' && r.urgency !== 'soon30') return false;
    if (timeRange === '90' && r.urgency !== 'soon30' && r.urgency !== 'soon90') return false;
    if (timeRange === 'future' && r.urgency !== 'future') return false;
    if (timeRange === 'past' && r.urgency !== 'past') return false;
    return true;
  });

  recs.sort((a, b) => {
    let va = a[sortColumn], vb = b[sortColumn];
    if (typeof va === 'string') {
      return sortDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    }
    return sortDir === 'asc' ? va - vb : vb - va;
  });

  return recs;
}

// Get filtered stocks (aggregated)
function getFilteredStocks() {
  const search = document.getElementById('searchInput').value.toLowerCase().trim();
  const timeRange = document.getElementById('timeRange').value;
  const industry = document.getElementById('industryFilter').value;

  let stocks = DATA.stocks.filter(s => {
    if (search && !s.code.toLowerCase().includes(search) && !s.name.toLowerCase().includes(search)) return false;
    if (industry && s.industry !== industry) return false;
    if (selectedMonth && s.liftDate.substring(0, 7) !== selectedMonth) return false;
    if (timeRange === '30' && s.urgency !== 'soon30') return false;
    if (timeRange === '90' && s.urgency !== 'soon30' && s.urgency !== 'soon90') return false;
    if (timeRange === 'future' && s.urgency !== 'future') return false;
    if (timeRange === 'past' && s.urgency !== 'past') return false;
    return true;
  });

  stocks.sort((a, b) => {
    let va = a[sortColumn === 'investor' ? 'investorCount' : sortColumn === 'amount' ? 'totalAmountHKD' : sortColumn === 'percentage' ? 'totalPercentage' : sortColumn], vb = b[sortColumn === 'investor' ? 'investorCount' : sortColumn === 'amount' ? 'totalAmountHKD' : sortColumn === 'percentage' ? 'totalPercentage' : sortColumn];
    if (typeof va === 'string') return sortDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    return sortDir === 'asc' ? va - vb : vb - va;
  });

  return stocks;
}

// Render table
function renderTable() {
  const head = document.getElementById('table-head');
  const body = document.getElementById('table-body');

  if (currentView === 'detail') {
    const recs = getFilteredRecords();
    head.innerHTML = '<tr>' +
      '<th onclick="sortTable(\'liftDate\')">解禁日 <span class="sort-arrow">'+(sortColumn==='liftDate'?(sortDir==='asc'?'&#9650;':'&#9660;'):'')+'</span></th>' +
      '<th onclick="sortTable(\'code\')">代码 <span class="sort-arrow">'+(sortColumn==='code'?(sortDir==='asc'?'&#9650;':'&#9660;'):'')+'</span></th>' +
      '<th onclick="sortTable(\'name\')">名称</th>' +
      '<th>基石投资者</th>' +
      '<th onclick="sortTable(\'shares\')">认购股数</th>' +
      '<th onclick="sortTable(\'amount\')">认购金额</th>' +
      '<th>币种</th>' +
      '<th onclick="sortTable(\'percentage\')">认购占比</th>' +
      '<th>限售期</th>' +
      '<th>行业</th>' +
      '</tr>';

    if (recs.length === 0) {
      body.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#b0aea5">无符合条件的记录</td></tr>';
    } else {
      body.innerHTML = recs.map(r =>
        '<tr class="'+rowClass(r.urgency)+'">' +
        '<td><span style="white-space:nowrap">'+r.liftDate+'</span> '+urgencyBadge(r.urgency)+'</td>' +
        '<td>'+r.code+'</td>' +
        '<td>'+r.stockName+'</td>' +
        '<td style="max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+r.investor.replace(/"/g,'&quot;')+'">'+r.investor+'</td>' +
        '<td>'+fmtShares(r.shares)+'</td>' +
        '<td>'+fmtAmount(r)+'</td>' +
        '<td>'+r.currency+'</td>' +
        '<td>'+fmtPct(r.percentage)+'</td>' +
        '<td>'+r.lockupMonths+'月</td>' +
        '<td style="font-size:13px;color:#b0aea5">'+(r.stockIndustry||'')+'</td>' +
        '</tr>'
      ).join('');
    }
    document.getElementById('filter-info').textContent = '共 ' + recs.length + ' 条记录';
  } else {
    const stocks = getFilteredStocks();
    head.innerHTML = '<tr>' +
      '<th onclick="sortTable(\'liftDate\')">解禁日 <span class="sort-arrow">'+(sortColumn==='liftDate'?(sortDir==='asc'?'&#9650;':'&#9660;'):'')+'</span></th>' +
      '<th onclick="sortTable(\'code\')">代码</th>' +
      '<th onclick="sortTable(\'name\')">名称</th>' +
      '<th onclick="sortTable(\'investor\')">基石投资者数</th>' +
      '<th onclick="sortTable(\'amount\')">合计金额</th>' +
      '<th onclick="sortTable(\'percentage\')">合计占比</th>' +
      '<th>行业</th>' +
      '<th>详情</th>' +
      '</tr>';

    if (stocks.length === 0) {
      body.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:40px;color:#b0aea5">无符合条件的记录</td></tr>';
    } else {
      body.innerHTML = stocks.map(s => {
        const isExpanded = expandedStocks.has(s.code);
        let html = '<tr class="'+rowClass(s.urgency)+'" onclick="toggleExpand(\''+s.code+'\')" style="cursor:pointer">' +
          '<td><span style="white-space:nowrap">'+s.liftDate+'</span> '+urgencyBadge(s.urgency)+'</td>' +
          '<td>'+s.code+'</td>' +
          '<td>'+s.name+'</td>' +
          '<td>'+s.investorCount+'</td>' +
          '<td>'+s.totalAmountYi+'亿港元</td>' +
          '<td>'+fmtPct(s.totalPercentage)+'</td>' +
          '<td style="font-size:13px;color:#b0aea5">'+(s.industry||'')+'</td>' +
          '<td><span class="expand-toggle">'+(isExpanded?'&#9660; 收起':'&#9654; 展开')+'</span></td>' +
          '</tr>';
        if (isExpanded) {
          html += '<tr class="expand-row"><td colspan="8"><div class="expand-content">' +
            '<table><thead><tr><th>基石投资者</th><th>认购股数</th><th>认购金额</th><th>币种</th><th>认购占比</th><th>限售期</th></tr></thead><tbody>' +
            s.investors.map(inv =>
              '<tr><td style="max-width:360px">'+inv.investor+'</td><td>'+fmtShares(inv.shares)+'</td><td>'+fmtAmount(inv)+'</td><td>'+inv.currency+'</td><td>'+fmtPct(inv.percentage)+'</td><td>'+inv.lockupMonths+'月</td></tr>'
            ).join('') +
            '</tbody></table></div></td></tr>';
        }
        return html;
      }).join('');
    }
    document.getElementById('filter-info').textContent = '共 ' + stocks.length + ' 只股票';
  }
}

// Toggle stock expansion
function toggleExpand(code) {
  if (expandedStocks.has(code)) {
    expandedStocks.delete(code);
  } else {
    expandedStocks.add(code);
  }
  renderTable();
}

// Sort
function sortTable(col) {
  if (sortColumn === col) {
    sortDir = sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    sortColumn = col;
    sortDir = 'asc';
  }
  renderTable();
}

// View toggle
function setView(view) {
  currentView = view;
  document.getElementById('view-detail').classList.toggle('active', view === 'detail');
  document.getElementById('view-summary').classList.toggle('active', view === 'summary');
  expandedStocks.clear();
  sortColumn = 'liftDate';
  sortDir = 'asc';
  renderTable();
}

// Filter change
function onFilterChange() {
  renderTable();
}

// Init
renderHeader();
renderCards();
renderIndustryFilter();
renderChart();
renderTable();
</script>
</body>
</html>'''


def main():
    data = process_data()
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    html = HTML_TEMPLATE.replace("__DATA_JSON__", json_str)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Generated: {OUTPUT_PATH}")
    print(f"Stocks: {data['totalStocks']}, Records: {data['totalRecords']}")
    print(f"Within 30 days: {data['summary']['within30']['count']} ({data['summary']['within30']['stocks']} stocks)")
    print(f"Within 90 days: {data['summary']['within90']['count']} ({data['summary']['within90']['stocks']} stocks)")
    print(f"Monthly buckets: {len(data['monthly'])}")


if __name__ == "__main__":
    main()
