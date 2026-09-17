from flask import Flask, jsonify, Response
import math
from scipy.optimize import brentq
from scipy.stats import norm

app = Flask(name)

def bs_price(S, K, T, r, sigma, kind):
    if T <= 0 or sigma <= 0:
        return max(S - K, 0.0) if kind == "C" else max(K - S, 0.0)
    d1 = (math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    if kind == "C":
        return S * norm.cdf(d1) - K * math.exp(-r * T) * norm.cdf(d2)
    return K * math.exp(-r * T) * norm.cdf(-d2) - S * norm.cdf(-d1)

@app.route("/")
def index():
    return """<!DOCTYPE html><html><head><meta charset="utf-8"><title>BTC IV</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>body{font:15px sans-serif;margin:0;background:#fbfbfa}
.wrap{max-width:1000px;margin:0 auto;padding:30px}h1{margin:0}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin:20px 0}
.kpi{background:#fff;border:1px solid #ddd;padding:15px;border-radius:8px}
.kpi .v{font-size:24px;font-weight:600}.kpi .l{font-size:11px;color:#888}
.card{background:#fff;border:1px solid #ddd;padding:18px;margin:20px 0;border-radius:8px}
.chart{position:relative;height:300px}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:right;color:#888;border-bottom:1px solid #ddd;padding:8px}
td{padding:8px;border-bottom:1px solid #eee}</style></head><body>
<div class="wrap"><h1>BTC options IV</h1><div class="sub" id="sub">loading...</div>
<div class="grid" id="kpis"></div>
<div class="card"><h2>Smile</h2><div class="chart"><canvas id="c"></canvas></div></div>
<div class="card"><h2>Options</h2><div id="tbl"></div></div></div>
<script>
fetch("/api/data").then(r=>r.json()).then(d=>{
  document.getElementById("sub").textContent = "$"+Math.round(d.spot).toLocaleString()+" • "+d.count+" options";
  document.getElementById("kpis").innerHTML = 
    '<div class="kpi"><div class="l">Spot</div><div class="v">$'+Math.round(d.spot).toLocaleString()+'</div></div>'+
    '<div class="kpi"><div class="l">Mean IV</div><div class="v">'+(d.mean_iv*100).toFixed(1)+'%</div></div>'+
    '<div class="kpi"><div class="l">Range</div><div class="v">'+(d.min_iv*100).toFixed(0)+'-'+(d.max_iv*100).toFixed(0)+'%</div></div>'+
    '<div class="kpi"><div class="l">Options</div><div class="v">'+d.count+'</div></div>';
  
  const C=d.rows.filter(r=>r.kind==='C'), P=d.rows.filter(r=>r.kind==='P');
  new Chart(document.getElementById("c"),{type:'scatter',data:{datasets:[
    {label:'Calls',data:C.map(r=>({x:r.moneyness,y:r.iv*100})),borderColor:'#2a78d6',backgroundColor:'#2a78d6',showLine:true,pointRadius:3},
    {label:'Puts',data:P.map(r=>({x:r.moneyness,y:r.iv*100})),borderColor:'#d95926',backgroundColor:'#d95926',showLine:true,pointRadius:3}
  ]},options:{responsive:true,maintainAspectRatio:false}});
  
  document.getElementById("tbl").innerHTML = '<table><tr><th>Strike</th><th>Kind</th><th>Bid</th><th>Ask</th><th>IV</th></tr>'+
    d.rows.slice(0,25).map(r=>'<tr><td>$'+Math.round(r.strike)+'</td><td>'+r.kind+'</td><td>$'+r.bid.toFixed(0)+'</td><td>$'+r.ask.toFixed(0)+'</td><td>'+(r.iv*100).toFixed(1)+'%</td></tr>').join('')+'</table>';
});
</script></body></html>"""

@app.route("/api/data")
def data():
    spot, T = 95000.0, 30/365
    rows = []
    for K in range(70000, 130001, 2500):
        iv = 0.55 + 0.15 * abs(1 - K/spot) if K != 92500 else 0.78
        for kind in "CP":
            px = bs_price(spot, K, T, 0.05, iv, kind) / spot
            rows.append({"instrument":f"BTC-17OCT26-{K}-{kind}","strike":K,"kind":kind,"spot":spot,
                        "bid":px*spot*0.985,"ask":px*spot*1.015,"iv":iv,"moneyness":K/spot})
    ivs=[r["iv"]for r in rows]
    return jsonify({"spot":spot,"count":len(rows),"mean_iv":sum(ivs)/len(ivs),
                   "min_iv":min(ivs),"max_iv":max(ivs),"rows":rows})

if name=="main":
    app.run(host="0.0.0.0",port=int(import('os').environ.get('PORT',5000)))