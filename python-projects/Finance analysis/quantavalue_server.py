#!/usr/bin/env python3
"""
QuantaValue Investment Dashboard - Local Server
Run: python server.py
Then open: http://localhost:8765
"""

import http.server
import socketserver
import json
import urllib.request
import urllib.parse
import threading
import webbrowser
import time
import os
from http.server import BaseHTTPRequestHandler

PORT = 8765
AV_BASE = "https://www.alphavantage.co/query"

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>QuantaValue — Live Investment Analysis</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400&family=IBM+Plex+Mono:wght@300;400;500;600&family=Outfit:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root {
  --bg:#05080f; --surface:#0a0f1a; --surface2:#0f1624; --surface3:#151e30;
  --border:rgba(148,200,255,0.1); --border2:rgba(148,200,255,0.05);
  --gold:#e8b84b; --gold2:#f5d07a; --blue:#5ba4f5; --blue2:#94c4ff;
  --green:#4ade9a; --red:#f87171; --orange:#fb923c;
  --text:#dce8f5; --text2:#8ba5c4; --text3:#3a5070;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--text);font-family:'Outfit',sans-serif;font-size:14px;line-height:1.6;min-height:100vh;overflow-x:hidden}
body::after{content:'';position:fixed;inset:0;background-image:linear-gradient(rgba(91,164,245,0.025) 1px,transparent 1px),linear-gradient(90deg,rgba(91,164,245,0.025) 1px,transparent 1px);background-size:60px 60px;pointer-events:none;z-index:0}
.page{position:relative;z-index:1;max-width:1280px;margin:0 auto;padding:0 28px 80px}
nav{display:flex;align-items:center;justify-content:space-between;padding:28px 0 24px;border-bottom:1px solid var(--border2);margin-bottom:36px}
.brand{display:flex;align-items:center;gap:14px}
.brand-icon{width:40px;height:40px;background:linear-gradient(135deg,#e8b84b,#5ba4f5);border-radius:10px;display:flex;align-items:center;justify-content:center}
.brand-name{font-family:'Playfair Display',serif;font-size:20px;letter-spacing:-0.3px}
.brand-name em{color:var(--gold);font-style:italic}
.live-badge{display:flex;align-items:center;gap:7px;font-family:'IBM Plex Mono',monospace;font-size:10px;font-weight:500;color:var(--green);background:rgba(74,222,154,0.07);border:1px solid rgba(74,222,154,0.2);padding:6px 14px;border-radius:100px;letter-spacing:.06em;text-transform:uppercase}
.live-dot{width:6px;height:6px;border-radius:50%;background:var(--green);box-shadow:0 0 8px var(--green);animation:blink 2s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
.input-section{background:var(--surface);border:1px solid var(--border);border-radius:20px;padding:32px;margin-bottom:28px;position:relative;overflow:hidden}
.input-section::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,var(--gold) 40%,var(--blue) 60%,transparent)}
.input-label{font-family:'IBM Plex Mono',monospace;font-size:10px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--text3);margin-bottom:20px}
.input-row{display:grid;grid-template-columns:160px 1fr 140px 140px auto;gap:14px;align-items:end}
.field{display:flex;flex-direction:column;gap:7px}
.field label{font-size:11px;font-weight:500;color:var(--text2);letter-spacing:.04em;text-transform:uppercase}
.field input{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;color:var(--text);font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:500;outline:none;transition:all .2s;width:100%}
.field input:focus{border-color:rgba(91,164,245,.45);box-shadow:0 0 0 3px rgba(91,164,245,.07)}
.field input::placeholder{color:var(--text3)}
.go-btn{background:linear-gradient(135deg,var(--gold),#d4921e);color:#0a0800;border:none;border-radius:10px;padding:13px 28px;font-family:'Outfit',sans-serif;font-weight:600;font-size:14px;cursor:pointer;transition:all .2s;white-space:nowrap;height:46px}
.go-btn:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(232,184,75,.28)}
.go-btn:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
.api-hint{margin-top:14px;font-size:11px;color:var(--text3)}
.api-hint a{color:var(--blue2);text-decoration:none}
#progress{display:none;background:var(--surface);border:1px solid var(--border2);border-radius:14px;padding:20px 24px;margin-bottom:24px;animation:fadeUp .3s ease}
.progress-header{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.spinner{width:16px;height:16px;border:2px solid var(--border);border-top-color:var(--gold);border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
.progress-label{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--text2);flex:1}
.progress-steps{display:flex;gap:6px;flex-wrap:wrap}
.pstep{font-family:'IBM Plex Mono',monospace;font-size:9px;padding:3px 10px;border-radius:100px;border:1px solid;transition:all .3s;letter-spacing:.05em;text-transform:uppercase}
.pstep.done{background:rgba(74,222,154,.08);border-color:rgba(74,222,154,.3);color:var(--green)}
.pstep.active{background:rgba(232,184,75,.1);border-color:rgba(232,184,75,.35);color:var(--gold)}
.pstep.wait{background:transparent;border-color:var(--border2);color:var(--text3)}
#errorBox{display:none;background:rgba(248,113,113,.07);border:1px solid rgba(248,113,113,.25);border-radius:12px;padding:18px 22px;margin-bottom:20px;font-size:13px;color:var(--red);animation:fadeUp .3s ease}
#dash{display:none;animation:fadeUp .5s ease}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.co-header{display:grid;grid-template-columns:1fr auto;gap:20px;align-items:start;margin-bottom:24px}
.co-name{font-family:'Playfair Display',serif;font-size:46px;line-height:1.05;letter-spacing:-1.5px;margin-bottom:6px}
.co-name .sym{color:var(--gold);font-style:italic}
.co-meta{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--text2);flex-wrap:wrap}
.meta-sep{color:var(--text3)}
.meta-chip{font-family:'IBM Plex Mono',monospace;font-size:10px;padding:3px 10px;border-radius:5px;background:rgba(91,164,245,.1);border:1px solid rgba(91,164,245,.2);color:var(--blue2)}
.fair-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:22px 26px;text-align:right;min-width:200px}
.fair-label{font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--text3);margin-bottom:6px}
.fair-price{font-family:'Playfair Display',serif;font-size:36px;line-height:1;margin-bottom:4px}
.fair-upside{font-size:12px;font-weight:600;font-family:'IBM Plex Mono',monospace}
.up{color:var(--green)}.down{color:var(--red)}
.kpi-row{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:20px}
.kpi{background:var(--surface);border:1px solid var(--border2);border-radius:12px;padding:15px;position:relative;overflow:hidden;transition:border-color .2s}
.kpi:hover{border-color:var(--border)}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:var(--kc,var(--blue));opacity:.4}
.kpi-lbl{font-family:'IBM Plex Mono',monospace;font-size:8.5px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);margin-bottom:7px}
.kpi-val{font-family:'Playfair Display',serif;font-size:21px;line-height:1;color:var(--kc,var(--blue2));margin-bottom:2px}
.kpi-sub{font-size:10px;color:var(--text3)}
.tabs{display:flex;gap:2px;background:var(--surface2);border-radius:12px;padding:4px;margin-bottom:20px}
.tab{flex:1;padding:9px 14px;border-radius:9px;text-align:center;font-size:12px;font-weight:500;cursor:pointer;color:var(--text3);transition:all .2s;white-space:nowrap}
.tab.on{background:var(--surface);color:var(--text);box-shadow:0 1px 5px rgba(0,0,0,.4)}
.tc{display:none}.tc.on{display:block;animation:fadeUp .3s ease}
.panel{background:var(--surface);border:1px solid var(--border2);border-radius:16px;padding:24px;margin-bottom:14px;position:relative;overflow:hidden}
.panel-title{font-family:'IBM Plex Mono',monospace;font-size:9.5px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--text3);margin-bottom:20px;display:flex;align-items:center;gap:8px}
.panel-title::before{content:'';width:3px;height:14px;background:var(--gold);border-radius:2px;flex-shrink:0}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.vbars{display:flex;flex-direction:column;gap:18px}
.vrow-head{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px}
.vmethod{font-size:12px;font-weight:500;color:var(--text2)}
.vprice{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:var(--text)}
.vdelta{font-family:'IBM Plex Mono',monospace;font-size:10px;margin-right:8px}
.vtrack{height:5px;background:var(--surface3);border-radius:3px;position:relative;overflow:visible}
.vfill{height:100%;border-radius:3px;transition:width 1.1s cubic-bezier(.4,0,.2,1)}
.vmarker{position:absolute;top:-5px;width:2px;height:15px;background:var(--red);border-radius:1px}
.vmarker-lbl{position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:8px;color:var(--red);font-family:'IBM Plex Mono',monospace;white-space:nowrap}
.scenarios{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
.sc-card{border-radius:12px;padding:18px;border:1px solid;background:var(--surface2)}
.sc-card.bear{border-color:rgba(248,113,113,.18)}.sc-card.base{border-color:rgba(232,184,75,.22)}.sc-card.bull{border-color:rgba(74,222,154,.18)}
.sc-lbl{font-family:'IBM Plex Mono',monospace;font-size:8.5px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;margin-bottom:8px}
.sc-card.bear .sc-lbl{color:var(--red)}.sc-card.base .sc-lbl{color:var(--gold)}.sc-card.bull .sc-lbl{color:var(--green)}
.sc-price{font-family:'Playfair Display',serif;font-size:26px;line-height:1;margin-bottom:10px}
.sc-card.bear .sc-price{color:var(--red)}.sc-card.base .sc-price{color:var(--gold2)}.sc-card.bull .sc-price{color:var(--green)}
.sc-row{display:flex;justify-content:space-between;font-size:10px;color:var(--text3);margin-top:3px}
.sc-row span:last-child{color:var(--text2);font-family:'IBM Plex Mono',monospace}
.stbl{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:11px}
.stbl th,.stbl td{padding:10px 16px;text-align:center;border:1px solid var(--border2)}
.stbl th{background:var(--surface2);color:var(--text3);font-size:9px;letter-spacing:.08em;text-transform:uppercase}
.cell-hi{background:rgba(74,222,154,.1);color:var(--green);font-weight:600}
.cell-md{background:rgba(232,184,75,.07);color:var(--gold)}
.cell-lo{background:rgba(248,113,113,.09);color:var(--red);font-weight:600}
.cell-cu{background:rgba(91,164,245,.15);color:var(--blue2);font-weight:700}
.dtbl{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:11px}
.dtbl th{padding:8px 12px;text-align:right;color:var(--text3);font-size:8.5px;letter-spacing:.09em;text-transform:uppercase;border-bottom:1px solid var(--border2)}
.dtbl th:first-child{text-align:left}
.dtbl td{padding:7px 12px;text-align:right;border-bottom:1px solid var(--border2);color:var(--text2)}
.dtbl td:first-child{text-align:left;color:var(--text3);font-size:10px}
.dtbl tr:hover td{background:rgba(91,164,245,.04)}
.dtbl tr.total-row td{border-bottom:none;font-weight:700;color:var(--text);border-top:1px solid var(--border)}
.dtbl tr.tv-row td{color:var(--gold);font-weight:600;border-bottom:none}
.g-box{background:var(--surface2);border:1px solid var(--border2);border-radius:12px;padding:20px}
.g-box-label{font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);margin-bottom:6px}
.g-box-val{font-family:'Playfair Display',serif;font-size:34px;color:var(--gold);line-height:1;margin-bottom:4px}
.g-box-formula{font-family:'IBM Plex Mono',monospace;font-size:9.5px;color:var(--text3);line-height:1.7}
.checklist{display:flex;flex-direction:column;gap:8px}
.cl-item{display:flex;align-items:center;gap:10px;padding:9px 12px;background:var(--surface2);border:1px solid var(--border2);border-radius:8px}
.cl-icon{font-size:13px;flex-shrink:0}.cl-text{font-size:11px;color:var(--text2);flex:1}
.cl-val{font-family:'IBM Plex Mono',monospace;font-size:10px;flex-shrink:0}
.moat-list{display:flex;flex-direction:column;gap:13px}
.moat-row{display:flex;align-items:center;gap:12px}
.moat-name{font-size:11px;color:var(--text2);width:130px;flex-shrink:0}
.moat-track{flex:1;height:4px;background:var(--surface3);border-radius:2px;overflow:hidden}
.moat-fill{height:100%;border-radius:2px;transition:width 1.2s cubic-bezier(.4,0,.2,1)}
.moat-rating{font-family:'IBM Plex Mono',monospace;font-size:10px;width:55px;text-align:right;flex-shrink:0}
.risk-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.risk-item{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;background:var(--surface2);border:1px solid var(--border2);border-radius:8px}
.risk-lbl{font-size:11px;color:var(--text2)}
.risk-badge{font-family:'IBM Plex Mono',monospace;font-size:8.5px;font-weight:700;padding:3px 8px;border-radius:4px;text-transform:uppercase;letter-spacing:.06em}
.rh{background:rgba(248,113,113,.14);color:var(--red)}.rm{background:rgba(232,184,75,.11);color:var(--gold)}.rl{background:rgba(74,222,154,.09);color:var(--green)}
.sc-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.sc-item{background:var(--surface2);border:1px solid var(--border2);border-radius:10px;padding:14px;text-align:center}
.sc-cat{font-size:10px;color:var(--text3);margin-bottom:8px;line-height:1.3}
.ring{width:52px;height:52px;margin:0 auto 5px;position:relative}
.ring svg{transform:rotate(-90deg)}
.ring-num{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-family:'Playfair Display',serif;font-size:17px}
.ftbl{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:11px}
.ftbl th{padding:8px 14px;text-align:right;color:var(--text3);font-size:8.5px;letter-spacing:.09em;text-transform:uppercase;border-bottom:1px solid var(--border2)}
.ftbl th:first-child{text-align:left}
.ftbl td{padding:8px 14px;text-align:right;border-bottom:1px solid var(--border2);color:var(--text2)}
.ftbl td:first-child{text-align:left;color:var(--text3)}
.ftbl tr:last-child td{border-bottom:none;font-weight:600;color:var(--text)}
.stat-card{background:var(--surface2);border:1px solid var(--border2);border-radius:10px;padding:14px 16px}
.stat-lbl{font-family:'IBM Plex Mono',monospace;font-size:8.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);margin-bottom:5px}
.stat-val{font-family:'Playfair Display',serif;font-size:20px;color:var(--blue2)}
.thesis-box{background:var(--surface2);border:1px solid var(--border2);border-radius:12px;padding:20px;margin-top:16px}
.thesis-txt{font-size:13px;color:var(--text2);line-height:1.85}
.disclaimer{margin-top:24px;padding:15px 20px;background:rgba(248,113,113,.04);border:1px solid rgba(248,113,113,.1);border-radius:10px;font-size:11px;color:var(--text3);line-height:1.7}
.disclaimer strong{color:var(--red)}
::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(3,1fr)}.input-row{grid-template-columns:1fr 1fr}.go-btn{grid-column:1/-1}.g2,.g3{grid-template-columns:1fr}.scenarios{grid-template-columns:1fr}.co-header{grid-template-columns:1fr}.sc-grid{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="page">
<nav>
  <div class="brand">
    <div class="brand-icon">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="1.8"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg>
    </div>
    <div class="brand-name">Quanta<em>Value</em></div>
  </div>
  <div class="live-badge"><div class="live-dot"></div>Alpha Vantage · Live Data</div>
</nav>

<div class="input-section">
  <div class="input-label">▸ Live Analysis — Data fetched server-side (no CORS issues)</div>
  <div class="input-row">
    <div class="field">
      <label>Ticker</label>
      <input id="ticker" type="text" placeholder="AAPL" maxlength="6" style="text-transform:uppercase;font-size:18px;font-weight:600;letter-spacing:.05em">
    </div>
    <div class="field">
      <label>Alpha Vantage API Key</label>
      <input id="apikey" type="text" placeholder="Your free API key from alphavantage.co">
    </div>
    <div class="field">
      <label>Phase 1 Growth % (Yr 1-5)</label>
      <input id="g1" type="number" placeholder="8" value="8" step="0.5">
    </div>
    <div class="field">
      <label>Terminal Growth %</label>
      <input id="tgr" type="number" placeholder="2.5" value="2.5" step="0.5">
    </div>
    <button class="go-btn" id="goBtn" onclick="runAnalysis()">Analyze →</button>
  </div>
  <div class="api-hint">
    Free key: <a href="https://www.alphavantage.co/support/#api-key" target="_blank">alphavantage.co</a>
    &nbsp;·&nbsp; Free tier: 25 req/day, 5/min &nbsp;·&nbsp; Calls are made server-side — no CORS &nbsp;·&nbsp; Running on <strong style="color:var(--green)">localhost:""" + str(PORT) + r"""</strong>
  </div>
</div>

<div id="progress">
  <div class="progress-header">
    <div class="spinner"></div>
    <div class="progress-label" id="pLabel">Initializing...</div>
  </div>
  <div class="progress-steps" id="pSteps"></div>
</div>

<div id="errorBox"></div>
<div id="dash">
  <div class="co-header">
    <div>
      <div class="co-name"><span class="sym" id="dTicker">—</span> <span id="dName">—</span></div>
      <div class="co-meta">
        <span id="dSector">—</span><span class="meta-sep">·</span>
        <span id="dIndustry">—</span><span class="meta-sep">·</span>
        <span class="meta-chip" id="dExchange">—</span><span class="meta-sep">·</span>
        <span id="dMarketCap" style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--text2)">—</span>
      </div>
    </div>
    <div class="fair-card">
      <div class="fair-label">Composite Fair Value</div>
      <div class="fair-price" id="dFairValue">—</div>
      <div class="fair-upside" id="dUpside">—</div>
    </div>
  </div>
  <div class="kpi-row" id="kpiRow"></div>
  <div class="tabs">
    <div class="tab on" onclick="tab('valuation',this)">Valuation</div>
    <div class="tab" onclick="tab('dcf',this)">DCF Model</div>
    <div class="tab" onclick="tab('graham',this)">Graham</div>
    <div class="tab" onclick="tab('fundamentals',this)">Fundamentals</div>
    <div class="tab" onclick="tab('scorecard',this)">Scorecard</div>
  </div>
  <!-- VALUATION TAB -->
  <div class="tc on" id="tc-valuation">
    <div class="g2">
      <div class="panel"><div class="panel-title">Method Comparison vs Current Price</div><div class="vbars" id="vbars"></div></div>
      <div class="panel"><div class="panel-title">Bear / Base / Bull Scenarios</div><div class="scenarios" id="scenarios"></div></div>
    </div>
    <div class="panel"><div class="panel-title">Sensitivity — Intrinsic Value/Share (WACC × Terminal Growth Rate)</div><div id="sensTable"></div>
      <div style="margin-top:10px;font-size:10px;color:var(--text3);font-family:'IBM Plex Mono',monospace">🟢 &gt;15% upside · 🟡 -5% to +15% · 🔴 &gt;5% downside · <span style="color:var(--blue2)">■</span> base case</div>
    </div>
  </div>
  <!-- DCF TAB -->
  <div class="tc" id="tc-dcf">
    <div class="panel">
      <div class="panel-title">10-Year DCF Model — Built on Real TTM Free Cash Flow</div>
      <div style="overflow-x:auto"><table class="dtbl" id="dcfTbl"></table></div>
      <div class="g4" style="margin-top:16px" id="dcfStats"></div>
      <div style="margin-top:12px;font-size:10px;color:var(--text3);font-family:'IBM Plex Mono',monospace;line-height:1.7">
        WACC = Cost of Equity (CAPM: Rf + β×ERP) weighted with after-tax cost of debt · Phase 2 growth = Phase 1 × 0.55 (industry maturation)
      </div>
    </div>
  </div>
  <!-- GRAHAM TAB -->
  <div class="tc" id="tc-graham">
    <div class="g2">
      <div class="panel">
        <div class="panel-title">Graham Valuation — Real EPS & Book Value from AV</div>
        <div class="g2" style="margin-bottom:16px" id="grahamBoxes"></div>
        <div style="padding:14px;background:var(--surface2);border-radius:10px;border:1px solid var(--border2);font-size:11px;color:var(--text2);line-height:1.75">
          <strong style="color:var(--text)">Note:</strong> Graham criteria suit asset-heavy 1930–70s industrials. Modern capital-light businesses (SaaS, platforms) often fail P/E and P/B tests — this alone does not signal overvaluation. Weight DCF and Growth Formula more heavily for such companies.
        </div>
      </div>
      <div class="panel">
        <div class="panel-title">Graham 8-Criteria Checklist</div>
        <div class="checklist" id="grahamCl"></div>
        <div style="margin-top:14px;padding:10px 14px;background:var(--surface2);border-radius:8px;display:flex;justify-content:space-between;align-items:center">
          <span style="font-size:12px;color:var(--text2)">Graham Score</span>
          <span style="font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:600;color:var(--gold)" id="grahamScore">— / 8</span>
        </div>
      </div>
    </div>
  </div>
  <!-- FUNDAMENTALS TAB -->
  <div class="tc" id="tc-fundamentals">
    <div class="g2">
      <div class="panel"><div class="panel-title">Moat Proxy Indicators</div><div class="moat-list" id="moatList"></div></div>
      <div class="panel"><div class="panel-title">Risk Assessment</div><div class="risk-grid" id="riskGrid"></div></div>
    </div>
    <div class="panel"><div class="panel-title">5-Year Financial Snapshot — Alpha Vantage Annual Data</div><div style="overflow-x:auto"><table class="ftbl" id="finTbl"></table></div></div>
    <div class="panel"><div class="panel-title">Earnings Quality Indicators</div><div class="g3" id="eqIndicators"></div></div>
  </div>
  <!-- SCORECARD TAB -->
  <div class="tc" id="tc-scorecard">
    <div class="panel">
      <div class="panel-title">Investment Quality Scorecard</div>
      <div class="sc-grid" id="scorecardGrid"></div>
      <div class="thesis-box"><div class="panel-title" style="margin-bottom:12px">Auto-Generated Investment Thesis</div><div class="thesis-txt" id="thesisTxt"></div></div>
    </div>
  </div>
</div>

<div class="disclaimer">
  <strong>⚠ Disclaimer:</strong> For <strong>educational purposes only</strong>. Not financial advice. Data from Alpha Vantage may have delays. All projections are model estimates — not predictions. Consult a licensed advisor before investing.
</div>
</div>

<script>
const STEPS = ['Overview & Quote','Income Statement','Balance Sheet','Cash Flow','Computing Models'];
function tab(n,el){document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));document.querySelectorAll('.tc').forEach(t=>t.classList.remove('on'));el.classList.add('on');document.getElementById('tc-'+n).classList.add('on')}
function setProgress(lbl,step){document.getElementById('pLabel').textContent=lbl;document.getElementById('pSteps').innerHTML=STEPS.map((s,i)=>`<div class="pstep ${i<step?'done':i===step?'active':'wait'}">${s}</div>`).join('')}
function showError(msg){const el=document.getElementById('errorBox');el.innerHTML='⚠ '+msg;el.style.display='block'}
function fmt(v,d=2){return(isNaN(v)||!isFinite(v))?'N/A':v.toFixed(d)}
function fmtBig(v){if(!v||isNaN(v))return'N/A';const a=Math.abs(v);if(a>=1e12)return(v/1e12).toFixed(2)+'T';if(a>=1e9)return(v/1e9).toFixed(2)+'B';if(a>=1e6)return(v/1e6).toFixed(2)+'M';return v.toFixed(0)}
function fd(v){return(!v||isNaN(v)||!isFinite(v))?'N/A':'$'+v.toFixed(2)}
function n(v){return parseFloat(v)||0}

async function runAnalysis(){
  const ticker=document.getElementById('ticker').value.trim().toUpperCase();
  const apikey=document.getElementById('apikey').value.trim();
  const g1pct=parseFloat(document.getElementById('g1').value)||8;
  const tgrPct=parseFloat(document.getElementById('tgr').value)||2.5;
  if(!ticker){showError('Enter a ticker symbol.');return}
  if(!apikey){showError('Enter your Alpha Vantage API key.');return}

  document.getElementById('errorBox').style.display='none';
  document.getElementById('dash').style.display='none';
  document.getElementById('progress').style.display='block';
  document.getElementById('goBtn').disabled=true;
  document.getElementById('goBtn').textContent='Fetching...';
  setProgress('Connecting to Alpha Vantage...',0);

  try {
    const res = await fetch(`/api/analyze?ticker=${encodeURIComponent(ticker)}&apikey=${encodeURIComponent(apikey)}`);
    const raw = await res.json();
    if(raw.error) throw new Error(raw.error);

    setProgress('Computing models...',4);
    const d = compute(raw, g1pct, tgrPct);
    render(d, ticker);
    document.getElementById('progress').style.display='none';
    document.getElementById('dash').style.display='block';
  } catch(e) {
    document.getElementById('progress').style.display='none';
    showError(e.message||'Unknown error.');
  } finally {
    document.getElementById('goBtn').disabled=false;
    document.getElementById('goBtn').textContent='Analyze →';
  }
}

function compute(raw, g1pct, tgrPct){
  const {overview:ov, quote:gq, income, balance, cashflow} = raw;
  const price=n(gq['05. price']);
  const changePct=parseFloat((gq['10. change percent']||'0').replace('%',''));
  const sharesOut=n(ov.SharesOutstanding);
  const marketCap=price*sharesOut;
  const beta=n(ov.Beta)||1.0;
  const divYield=n(ov.DividendYield)*100;
  const peRatio=n(ov.PERatio);
  const fwdPE=n(ov.ForwardPE);
  const pbRatio=n(ov.PriceToBookRatio);
  const evEbitda=n(ov.EVToEBITDA);
  const eps=n(ov.EPS);
  const analystTarget=n(ov.AnalystTargetPrice);

  const annIncome=(income.annualReports||[]).slice(0,5);
  const annBalance=(balance.annualReports||[]).slice(0,5);
  const annCashflow=(cashflow.annualReports||[]).slice(0,5);
  const inc0=annIncome[0]||{};
  const bal0=annBalance[0]||{};
  const cf0=annCashflow[0]||{};

  const revenue=n(inc0.totalRevenue);
  const grossProfit=n(inc0.grossProfit);
  const ebitda=n(inc0.ebitda);
  const ebit=n(inc0.ebit);
  const netIncome=n(inc0.netIncome);
  const ocf=n(cf0.operatingCashflow);
  const capex=Math.abs(n(cf0.capitalExpenditures));
  const fcf=ocf-capex;
  const fcfMarginPct=revenue>0?fcf/revenue*100:0;
  const grossMargin=revenue>0?grossProfit/revenue*100:0;
  const ebitdaMargin=revenue>0?ebitda/revenue*100:0;
  const netMargin=revenue>0?netIncome/revenue*100:0;

  const cash=n(bal0.cashAndCashEquivalentsAtCarryingValue)||n(bal0.cashAndShortTermInvestments);
  const longDebt=n(bal0.longTermDebtNoncurrent)||n(bal0.longTermDebt)||0;
  const shortDebt=n(bal0.shortLongTermDebtTotal)||n(bal0.shortTermDebt)||0;
  const totalDebt=longDebt+shortDebt;
  const netDebt=totalDebt-cash;
  const equity=n(bal0.totalShareholderEquity);
  const bvps=sharesOut>0?equity/sharesOut:0;
  const currentAssets=n(bal0.totalCurrentAssets);
  const currentLiab=n(bal0.totalCurrentLiabilities);
  const currentRatio=currentLiab>0?currentAssets/currentLiab:0;
  const inventory=n(bal0.inventory)||0;
  const quickRatio=currentLiab>0?(currentAssets-inventory)/currentLiab:0;
  const investedCapital=totalDebt+equity-cash;
  const interestExp=Math.abs(n(inc0.interestExpense))||1;
  const taxRate=n(inc0.incomeTaxExpense)/(n(inc0.incomeBeforeTax)||1);
  const afterTaxCostOfDebt=totalDebt>0?(interestExp/totalDebt)*(1-Math.max(taxRate,0)):0.03;
  const Rf=0.043;const ERP=0.055;
  const costOfEquity=Rf+beta*ERP;
  const eW=marketCap/(marketCap+Math.max(netDebt,0));
  const wacc=eW*costOfEquity+(1-eW)*afterTaxCostOfDebt;
  const roic=investedCapital>0?(ebit*(1-Math.max(taxRate,0.2)))/investedCapital*100:0;
  const fcfToNI=netIncome>0?fcf/netIncome:null;
  const accruals=n(bal0.totalAssets)>0?(netIncome-ocf)/n(bal0.totalAssets):null;
  const fcfPerShare=sharesOut>0?fcf/sharesOut:0;
  const pfcf=fcf>0?marketCap/fcf:null;
  const netDebtEbitda=ebitda>0?netDebt/ebitda:null;
  const intCoverage=interestExp>0?ebit/interestExp:null;

  // DCF
  const g1=g1pct/100;const g2=g1*0.55;const tgr=tgrPct/100;
  const baseFCF=Math.max(fcf,1);
  const dcfRows=[];let cumPV=0;let fcfY=baseFCF;
  for(let y=1;y<=10;y++){
    const gr=y<=5?g1:g2;
    fcfY*=(1+gr);
    const df=1/Math.pow(1+wacc,y);
    const pv=fcfY*df;
    cumPV+=pv;
    dcfRows.push({yr:'Year '+y,rev:'—',fcfM:fmt(fcfY/Math.max(revenue*(1+(y<=5?g1:g2)*y),1)*100,1)+'%',fcf:fmtBig(fcfY),df:df.toFixed(4),pv:fmtBig(pv)});
  }
  const terminalFCF=fcfY*(1+tgr);
  const terminalValue=wacc>tgr?terminalFCF/(wacc-tgr):0;
  const pvTerminal=terminalValue/Math.pow(1+wacc,10);
  const ev2=cumPV+pvTerminal;
  const equityValueDCF=ev2-Math.max(netDebt,0)+cash;
  const intrinsicDCF=sharesOut>0?equityValueDCF/sharesOut:0;

  function scDCF(gP1,fcfAdj,wAdj){
    let fv=baseFCF,sum=0;
    for(let y=1;y<=10;y++){const gr=y<=5?gP1:gP1*0.55;fv*=(1+gr*fcfAdj);sum+=fv/Math.pow(1+wAdj,y);}
    const tv=wacc>tgr?fv*(1+tgr)/(wAdj-tgr):0;
    const eq=(sum+tv/Math.pow(1+wAdj,10))-Math.max(netDebt,0)+cash;
    return sharesOut>0?eq/sharesOut:0;
  }
  const bearVal=scDCF(g1*0.4,0.85,wacc*1.12);
  const baseVal=intrinsicDCF;
  const bullVal=scDCF(g1*1.6,1.1,wacc*0.9);

  // EPS history
  const epsArr=annIncome.map(r=>n(r.netIncome)/(sharesOut||1));
  const eps10yrGrowth=epsArr.length>=2?((epsArr[0]/Math.max(epsArr[epsArr.length-1],0.01))**(1/(epsArr.length-1))-1)*100:8;
  const epsConsistent=epsArr.every(e=>e>0);
  const eps33growth=epsArr.length>=2?((epsArr[0]-epsArr[epsArr.length-1])/Math.max(Math.abs(epsArr[epsArr.length-1]),0.01)*100)>=33:false;

  // Graham
  const aaaYield=5.4;
  const grahamNum=eps>0&&bvps>0?Math.sqrt(22.5*eps*bvps):0;
  const grahamGrowth=eps>0?eps*(8.5+2*Math.max(eps10yrGrowth,0))*(4.4/aaaYield):0;

  const grahamChecks=[
    {icon:peRatio<15&&peRatio>0?'🟢':'🔴',text:'P/E < 15',val:peRatio>0?fmt(peRatio,1)+'×':'N/A',pass:peRatio<15&&peRatio>0},
    {icon:pbRatio<1.5&&pbRatio>0?'🟢':pbRatio<3?'🟡':'🔴',text:'P/B < 1.5',val:pbRatio>0?fmt(pbRatio,1)+'×':'N/A',pass:pbRatio<1.5&&pbRatio>0},
    {icon:currentRatio>=2?'🟢':currentRatio>=1.5?'🟡':'🔴',text:'Current Ratio ≥ 2',val:fmt(currentRatio,2)+'×',pass:currentRatio>=2},
    {icon:longDebt<=(currentAssets-currentLiab)?'🟢':'🔴',text:'LT Debt ≤ Net Working Capital',val:longDebt<=(currentAssets-currentLiab)?'PASS':'FAIL',pass:longDebt<=(currentAssets-currentLiab)},
    {icon:epsConsistent?'🟢':'🔴',text:'Consistent EPS (no losses)',val:epsConsistent?'PASS':'FAIL',pass:epsConsistent},
    {icon:divYield>0?'🟢':'🔴',text:'Pays a Dividend',val:divYield>0?fmt(divYield,2)+'%':'None',pass:divYield>0},
    {icon:eps33growth?'🟢':'🔴',text:'EPS growth ≥ 33% (historical)',val:eps33growth?'PASS':'FAIL',pass:eps33growth},
    {icon:grahamNum>price&&grahamNum>0?'🟢':'🔴',text:'Graham Number > Market Price',val:grahamNum>0?'$'+fmt(grahamNum,2):'N/A',pass:grahamNum>price&&grahamNum>0},
  ];
  const grahamScore=grahamChecks.filter(c=>c.pass).length;

  const analystW=analystTarget>0?analystTarget:intrinsicDCF;
  const compositeValue=intrinsicDCF*0.40+(grahamNum>0?grahamNum*0.10:intrinsicDCF*0.10)+(grahamGrowth>0?grahamGrowth*0.15:intrinsicDCF*0.15)+analystW*0.35;
  const upsidePct=price>0?(compositeValue-price)/price*100:0;

  // Sensitivity
  const waccVals=[wacc-0.01,wacc,wacc+0.01];
  const tgrVals=[tgr-0.01,tgr,tgr+0.01];
  const sensData=tgrVals.map(tg=>waccVals.map(wc=>{
    if(wc<=tg)return 0;
    let fv2=baseFCF,sum2=0;
    for(let y=1;y<=10;y++){const gr=y<=5?g1:g2;fv2*=(1+gr);sum2+=fv2/Math.pow(1+wc,y);}
    const tv2=fv2*(1+tg)/(wc-tg);
    const eq2=(sum2+tv2/Math.pow(1+wc,10))-Math.max(netDebt,0)+cash;
    return sharesOut>0?eq2/sharesOut:0;
  }));

  const finRows=annIncome.map((inc,i)=>{
    const cf=annCashflow[i]||{};const rv=n(inc.totalRevenue);const gp=n(inc.grossProfit);const ni=n(inc.netIncome);
    const oo=n(cf.operatingCashflow);const cx=Math.abs(n(cf.capitalExpenditures));
    return{yr:(inc.fiscalDateEnding||'').slice(0,4),rev:fmtBig(rv),gm:rv>0?fmt(gp/rv*100,1)+'%':'N/A',ni:fmtBig(ni),fcf:fmtBig(oo-cx),eps:sharesOut>0?'$'+fmt(ni/sharesOut,2):'N/A'};
  });

  return {
    name:ov.Name||'—',sector:ov.Sector||'—',industry:ov.Industry||'—',exchange:ov.Exchange||'—',
    price,changePct,marketCap,sharesOut,beta,
    peRatio,fwdPE,pbRatio,evEbitda,pfcf,divYield,
    roe:n(ov.ReturnOnEquityTTM)*100,roa:n(ov.ReturnOnAssetsTTM)*100,roic,
    currentRatio,quickRatio,netDebtEbitda,intCoverage,
    grossMargin,ebitdaMargin,netMargin,fcfMarginPct,
    revenue,grossProfit,ebitda,netIncome,eps,bvps,
    ocf,capex,fcf,fcfPerShare,cash,totalDebt,netDebt,equity,
    wacc,costOfEquity,afterTaxCostOfDebt,Rf,ERP,g1,g2,tgr,
    dcfRows,cumPV,pvTerminal,ev2,equityValueDCF,intrinsicDCF,
    bearVal,baseVal,bullVal,
    grahamNum,grahamGrowth,grahamChecks,grahamScore,aaaYield,eps10yrGrowth,
    compositeValue,upsidePct,analystTarget,
    sensData,waccVals,tgrVals,
    finRows,fcfToNI,accruals,
    epsConsistent,eps33growth,
  };
}

function render(d, ticker){
  document.getElementById('dTicker').textContent=ticker;
  document.getElementById('dName').textContent=d.name;
  document.getElementById('dSector').textContent=d.sector;
  document.getElementById('dIndustry').textContent=d.industry;
  document.getElementById('dExchange').textContent=d.exchange;
  document.getElementById('dMarketCap').textContent='Mkt Cap: $'+fmtBig(d.marketCap);
  document.getElementById('dFairValue').textContent=fd(d.compositeValue);
  document.getElementById('dFairValue').className='fair-price '+(d.upsidePct>=0?'up':'down');
  document.getElementById('dUpside').textContent=(d.upsidePct>=0?'▲ +':'▼ ')+fmt(d.upsidePct,1)+'% '+(d.upsidePct>=0?'upside':'downside')+' · Current: $'+fmt(d.price,2);
  document.getElementById('dUpside').className='fair-upside '+(d.upsidePct>=0?'up':'down');

  const kpis=[
    {l:'Price',v:'$'+fmt(d.price,2),s:(d.changePct>=0?'▲ +':'▼ ')+fmt(d.changePct,2)+'%',c:'#5ba4f5'},
    {l:'P/E (TTM)',v:d.peRatio>0?fmt(d.peRatio,1)+'×':'N/A',s:d.fwdPE>0?'Fwd: '+fmt(d.fwdPE,1)+'×':'—',c:'#e8b84b'},
    {l:'EV/EBITDA',v:d.evEbitda>0?fmt(d.evEbitda,1)+'×':'N/A',s:'Enterprise mult.',c:'#fb923c'},
    {l:'FCF Yield',v:d.marketCap>0&&d.fcf>0?fmt(d.fcf/d.marketCap*100,2)+'%':'N/A',s:'TTM FCF / Mkt Cap',c:'#4ade9a'},
    {l:'ROIC',v:fmt(d.roic,1)+'%',s:'vs WACC '+fmt(d.wacc*100,2)+'%',c:'#4ade9a'},
    {l:'Beta',v:fmt(d.beta,2),s:'vs S&P 500',c:'#94c4ff'},
  ];
  document.getElementById('kpiRow').innerHTML=kpis.map(k=>`<div class="kpi" style="--kc:${k.c}"><div class="kpi-lbl">${k.l}</div><div class="kpi-val">${k.v}</div><div class="kpi-sub">${k.s}</div></div>`).join('');

  const methods=[
    {m:'DCF — Base Case',p:d.intrinsicDCF,c:'#5ba4f5'},
    {m:'Graham Number',p:d.grahamNum,c:'#e8b84b'},
    {m:'Graham Growth Formula',p:d.grahamGrowth,c:'#f5d07a'},
    {m:'Analyst Consensus',p:d.analystTarget>0?d.analystTarget:null,c:'#4ade9a'},
    {m:'Composite Fair Value',p:d.compositeValue,c:'#fb923c'},
  ].filter(m=>m.p>0);
  const maxP=Math.max(...methods.map(m=>m.p),d.price)*1.12;
  document.getElementById('vbars').innerHTML=methods.map(m=>{
    const diff=d.price>0?(m.p-d.price)/d.price*100:0;
    const dc=diff>=0?'#4ade9a':'#f87171';
    return`<div class="vrow"><div class="vrow-head"><div class="vmethod">${m.m}</div><div style="display:flex;align-items:center;gap:6px"><span class="vdelta" style="color:${dc}">${diff>=0?'+':''}${fmt(diff,1)}%</span><div class="vprice">${fd(m.p)}</div></div></div><div class="vtrack"><div class="vfill" style="width:${(m.p/maxP*100).toFixed(1)}%;background:${m.c};opacity:.75"></div><div class="vmarker" style="left:${(d.price/maxP*100).toFixed(1)}%"><div class="vmarker-lbl">$${fmt(d.price,2)}</div></div></div></div>`;
  }).join('');

  document.getElementById('scenarios').innerHTML=[
    {t:'bear',l:'Bear Case',p:d.bearVal,cagr:fmt(d.g1*0.4*100,1)+'%',wacc:fmt(d.wacc*1.12*100,1)+'%'},
    {t:'base',l:'Base Case',p:d.baseVal,cagr:fmt(d.g1*100,1)+'%',wacc:fmt(d.wacc*100,1)+'%'},
    {t:'bull',l:'Bull Case',p:d.bullVal,cagr:fmt(d.g1*1.6*100,1)+'%',wacc:fmt(d.wacc*0.9*100,1)+'%'},
  ].map(s=>`<div class="sc-card ${s.t}"><div class="sc-lbl">${s.l}</div><div class="sc-price">${fd(s.p)}</div><div class="sc-row"><span>Rev CAGR</span><span>${s.cagr}</span></div><div class="sc-row"><span>WACC</span><span>${s.wacc}</span></div><div class="sc-row"><span>Implied ∆</span><span style="color:${s.p>=d.price?'#4ade9a':'#f87171'}">${s.p>=d.price?'+':''}${fmt((s.p-d.price)/d.price*100,1)}%</span></div></div>`).join('');

  let sh=`<table class="stbl"><tr><th style="text-align:left">TGR \\ WACC</th>`;
  d.waccVals.forEach(w=>sh+=`<th>${fmt(w*100,1)}%</th>`);sh+='</tr>';
  d.tgrVals.forEach((tg,i)=>{
    sh+=`<tr><td style="text-align:left;color:var(--text3);font-family:'IBM Plex Mono',monospace;font-size:9px">${fmt(tg*100,1)}%</td>`;
    d.sensData[i].forEach((v,j)=>{const diff=(v-d.price)/d.price;const cls=(i===1&&j===1)?'cell-cu':diff>0.15?'cell-hi':diff<-0.05?'cell-lo':'cell-md';sh+=`<td class="${cls}">${v>0?'$'+fmt(v,2):'N/A'}</td>`;});
    sh+='</tr>';
  });
  sh+='</table>';
  document.getElementById('sensTable').innerHTML=sh;

  let dt=`<thead><tr><th>Year</th><th>FCF</th><th>Discount Factor</th><th>PV of FCF</th></tr></thead><tbody>`;
  d.dcfRows.forEach(r=>dt+=`<tr><td>${r.yr}</td><td>${r.fcf}</td><td>${r.df}</td><td>${r.pv}</td></tr>`);
  dt+=`<tr class="tv-row"><td>Terminal Value (PV)</td><td colspan="3" style="text-align:right;color:var(--gold)">${fmtBig(d.pvTerminal)}</td></tr>`;
  dt+=`<tr class="total-row"><td>Enterprise Value</td><td colspan="3" style="text-align:right">${fmtBig(d.ev2)}</td></tr>`;
  dt+='</tbody>';
  document.getElementById('dcfTbl').innerHTML=dt;

  document.getElementById('dcfStats').innerHTML=[
    {l:'TTM FCF',v:fmtBig(d.fcf)},{l:'WACC',v:fmt(d.wacc*100,2)+'%'},
    {l:'Cost of Equity',v:fmt(d.costOfEquity*100,2)+'%'},{l:'Beta',v:fmt(d.beta,2)},
    {l:'Phase 1 Growth',v:fmt(d.g1*100,1)+'%'},{l:'Phase 2 Growth',v:fmt(d.g2*100,1)+'%'},
    {l:'Terminal Growth',v:fmt(d.tgr*100,1)+'%'},{l:'Intrinsic / Share',v:fd(d.intrinsicDCF)},
    {l:'Equity Value',v:fmtBig(d.equityValueDCF)},{l:'Net Debt',v:fmtBig(d.netDebt)},
    {l:'Shares Out.',v:fmtBig(d.sharesOut)},{l:'Rf (Treasury)',v:fmt(d.Rf*100,1)+'%'},
  ].map(s=>`<div class="stat-card"><div class="stat-lbl">${s.l}</div><div class="stat-val">${s.v}</div></div>`).join('');

  document.getElementById('grahamBoxes').innerHTML=[
    {l:'Graham Number',v:d.grahamNum>0?fd(d.grahamNum):'N/A',f:`√(22.5 × EPS × BVPS)\nEPS: $${fmt(d.eps,2)}  BVPS: $${fmt(d.bvps,2)}`,delta:d.grahamNum>0?((d.grahamNum-d.price)/d.price*100):null},
    {l:'Graham Growth Formula',v:d.grahamGrowth>0?fd(d.grahamGrowth):'N/A',f:`EPS × (8.5 + 2g) × (4.4/AAA)\ng: ${fmt(d.eps10yrGrowth,1)}%  AAA: ${d.aaaYield}%`,delta:d.grahamGrowth>0?((d.grahamGrowth-d.price)/d.price*100):null},
  ].map(g=>`<div class="g-box"><div class="g-box-label">${g.l}</div><div class="g-box-val">${g.v}</div>${g.delta!==null?`<div style="font-size:11px;color:${g.delta>=0?'#4ade9a':'#f87171'};font-family:'IBM Plex Mono',monospace;margin-bottom:8px">${g.delta>=0?'+':''}${fmt(g.delta,1)}% vs price</div>`:''}<div class="g-box-formula">${g.f}</div></div>`).join('');

  document.getElementById('grahamCl').innerHTML=d.grahamChecks.map(c=>`<div class="cl-item"><div class="cl-icon">${c.icon}</div><div class="cl-text">${c.text}</div><div class="cl-val" style="color:${c.icon==='🟢'?'#4ade9a':c.icon==='🟡'?'#e8b84b':'#f87171'}">${c.val}</div></div>`).join('');
  document.getElementById('grahamScore').textContent=`${d.grahamScore} / 8`;

  const moatItems=[
    {n:'Gross Margin Quality',s:Math.min(d.grossMargin*1.3,100)},
    {n:'FCF Conversion',s:d.fcfToNI?Math.min(d.fcfToNI*80,100):40},
    {n:'ROIC vs WACC',s:Math.min(Math.max((d.roic-d.wacc*100+5)*4,5),95)},
    {n:'Net Margin Strength',s:Math.min(d.netMargin*3.5,95)},
    {n:'EBITDA Margin',s:Math.min(d.ebitdaMargin*2,95)},
  ];
  document.getElementById('moatList').innerHTML=moatItems.map(m=>{
    const c=m.s>=75?'#4ade9a':m.s>=50?'#5ba4f5':m.s>=30?'#e8b84b':'#f87171';
    const r=m.s>=75?'Strong':m.s>=50?'Moderate':m.s>=30?'Weak':'Poor';
    return`<div class="moat-row"><div class="moat-name">${m.n}</div><div class="moat-track"><div class="moat-fill" style="width:${m.s}%;background:${c}"></div></div><div class="moat-rating" style="color:${c}">${r}</div></div>`;
  }).join('');

  const riskClass={High:'rh',Medium:'rm',Low:'rl'};
  document.getElementById('riskGrid').innerHTML=[
    {l:'Valuation Risk',lv:d.upsidePct<-10?'High':d.upsidePct<10?'Medium':'Low'},
    {l:'Leverage Risk',lv:(d.netDebtEbitda||0)>3?'High':(d.netDebtEbitda||0)>1.5?'Medium':'Low'},
    {l:'Earnings Quality',lv:d.fcfToNI>0.9?'Low':d.fcfToNI>0.6?'Medium':'High'},
    {l:'Interest Rate Sensitivity',lv:d.beta>1.3?'High':d.beta>0.9?'Medium':'Low'},
    {l:'FCF Consistency',lv:d.fcfToNI>1?'Low':d.fcfToNI>0.7?'Medium':'High'},
    {l:'Dividend Safety',lv:d.divYield>0&&d.fcfToNI>0.8?'Low':d.divYield>0?'Medium':'Low'},
    {l:'Growth Sustainability',lv:d.g1*100>20?'Medium':'Low'},
    {l:'Liquidity',lv:d.currentRatio<1?'High':d.currentRatio<1.5?'Medium':'Low'},
  ].map(r=>`<div class="risk-item"><div class="risk-lbl">${r.l}</div><div class="risk-badge ${riskClass[r.lv]}">${r.lv}</div></div>`).join('');

  let ft=`<thead><tr><th style="text-align:left">Year</th><th>Revenue</th><th>Gross Margin</th><th>Net Income</th><th>FCF</th><th>EPS</th></tr></thead><tbody>`;
  d.finRows.forEach(r=>ft+=`<tr><td>${r.yr}</td><td>${r.rev}</td><td>${r.gm}</td><td>${r.ni}</td><td>${r.fcf}</td><td>${r.eps}</td></tr>`);
  ft+='</tbody>';
  document.getElementById('finTbl').innerHTML=ft;

  document.getElementById('eqIndicators').innerHTML=[
    {l:'FCF / Net Income',v:d.fcfToNI?fmt(d.fcfToNI,2)+'×':'N/A',note:d.fcfToNI>1?'🟢 Excellent':d.fcfToNI>0.7?'🟡 Good':'🔴 Weak'},
    {l:'Accruals Ratio',v:d.accruals?fmt(d.accruals*100,2)+'%':'N/A',note:d.accruals&&d.accruals<0.05?'🟢 Low':d.accruals&&d.accruals<0.1?'🟡 Moderate':'🔴 High'},
    {l:'Gross Margin',v:fmt(d.grossMargin,1)+'%',note:d.grossMargin>50?'🟢 Premium':d.grossMargin>30?'🟡 OK':'🔴 Thin'},
    {l:'Net Margin',v:fmt(d.netMargin,1)+'%',note:d.netMargin>15?'🟢 Strong':d.netMargin>5?'🟡 Average':'🔴 Thin'},
    {l:'CapEx Intensity',v:d.revenue>0?fmt(d.capex/d.revenue*100,1)+'%':'N/A',note:d.capex/d.revenue<0.05?'🟢 Asset-light':d.capex/d.revenue<0.15?'🟡 Moderate':'🔴 Heavy'},
    {l:'OCF Margin',v:d.revenue>0?fmt(d.ocf/d.revenue*100,1)+'%':'N/A',note:d.ocf/d.revenue>0.2?'🟢 Strong':d.ocf/d.revenue>0.1?'🟡 Moderate':'🔴 Weak'},
  ].map(e=>`<div class="stat-card"><div class="stat-lbl">${e.l}</div><div class="stat-val" style="font-size:18px">${e.v}</div><div style="font-size:10px;color:var(--text3);margin-top:4px">${e.note}</div></div>`).join('');

  const s1=Math.min(d.grossMargin/60*10,10);
  const s2=d.fcfToNI?Math.min(d.fcfToNI*8,10):5;
  const s3=Math.min(d.roic/3,10);
  const s4=Math.min(d.g1*100/3,10);
  const s5=Math.max(0,Math.min(10,5+d.upsidePct/10));
  const s6=d.netDebtEbitda!==null?Math.max(0,10-d.netDebtEbitda*2):7;
  const s7=Math.min(d.currentRatio/2*8,10);
  const s8=d.divYield>0?Math.min(d.divYield+5,10):4;
  const overall=((s1+s2+s3+s4+s5+s6+s7+s8)/8);
  const scItems=[
    {c:'Profitability',s:Math.round(s1*10)/10},{c:'Earnings Quality',s:Math.round(s2*10)/10},
    {c:'Capital Efficiency',s:Math.round(s3*10)/10},{c:'Growth Outlook',s:Math.round(s4*10)/10},
    {c:'Valuation',s:Math.round(s5*10)/10},{c:'Leverage',s:Math.round(s6*10)/10},
    {c:'Liquidity',s:Math.round(s7*10)/10},{c:'Overall Score',s:Math.round(overall*10)/10},
  ];
  document.getElementById('scorecardGrid').innerHTML=scItems.map(s=>{
    const pct=s.s/10;const r=22;const circ=2*Math.PI*r;const dash=circ*pct;
    const col=s.s>=7.5?'#4ade9a':s.s>=5?'#e8b84b':'#f87171';
    return`<div class="sc-item"><div class="sc-cat">${s.c}</div><div class="ring"><svg width="52" height="52" viewBox="0 0 52 52"><circle cx="26" cy="26" r="${r}" fill="none" stroke="var(--surface3)" stroke-width="4"/><circle cx="26" cy="26" r="${r}" fill="none" stroke="${col}" stroke-width="4" stroke-dasharray="${dash.toFixed(1)} ${circ.toFixed(1)}" stroke-linecap="round"/></svg><div class="ring-num" style="color:${col}">${s.s.toFixed(1)}</div></div></div>`;
  }).join('');

  const updownStr=d.upsidePct>=0?`~${fmt(d.upsidePct,1)}% upside`:`~${fmt(Math.abs(d.upsidePct),1)}% downside`;
  document.getElementById('thesisTxt').textContent=
    `${d.name} (${ticker}) generates $${fmtBig(d.revenue)} in annual revenue with a ${fmt(d.grossMargin,1)}% gross margin and ${fmt(d.fcfMarginPct,1)}% FCF margin. TTM free cash flow of ${fmtBig(d.fcf)} ($${fmt(d.fcfPerShare,2)}/share) is the anchor for the DCF model. `+
    `Using a ${fmt(d.wacc*100,2)}% WACC (β=${fmt(d.beta,2)}, Rf=${fmt(d.Rf*100,1)}%, ERP=5.5%) and ${fmt(d.g1*100,1)}% Phase 1 growth rate, the DCF implies ${fd(d.intrinsicDCF)} per share. `+
    `The Graham Number (${d.grahamNum>0?fd(d.grahamNum):'N/A'}) and Growth Formula (${d.grahamGrowth>0?fd(d.grahamGrowth):'N/A'}) reflect classical value discipline — the company scores ${d.grahamScore}/8 on Graham's criteria. `+
    `The composite fair value across all methods is ${fd(d.compositeValue)}, implying ${updownStr} from the current price of ${fd(d.price)}. `+
    `Key metrics to monitor: FCF/NI ratio (${d.fcfToNI?fmt(d.fcfToNI,2)+'×':'N/A'}), ROIC (${fmt(d.roic,1)}%) vs WACC (${fmt(d.wacc*100,2)+'%'}), and whether revenue growth tracks the ${fmt(d.g1*100,1)}% base assumption.`;
}

document.getElementById('ticker').addEventListener('keydown',e=>{if(e.key==='Enter')runAnalysis()});
document.getElementById('apikey').addEventListener('keydown',e=>{if(e.key==='Enter')runAnalysis()});
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        # Clean logging
        print(f"  {args[0]} {args[1]}", flush=True)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode('utf-8'))

        elif parsed.path == '/api/analyze':
            params = urllib.parse.parse_qs(parsed.query)
            ticker = params.get('ticker', [''])[0].upper()
            apikey = params.get('apikey', [''])[0]

            if not ticker or not apikey:
                self._json_error('Missing ticker or apikey', 400)
                return

            try:
                result = self._fetch_all(ticker, apikey)
                self._json_ok(result)
            except Exception as e:
                self._json_error(str(e), 500)

        else:
            self.send_response(404)
            self.end_headers()

    def _json_ok(self, data):
        body = json.dumps(data).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json_error(self, msg, code=500):
        body = json.dumps({'error': msg}).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _av_fetch(self, params):
        """Fetch one Alpha Vantage endpoint"""
        url = AV_BASE + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={'User-Agent': 'QuantaValue/1.0'})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read().decode('utf-8'))
        if 'Note' in data:
            raise Exception('Alpha Vantage rate limit hit (5 req/min on free tier). Wait 60 seconds and retry.')
        if 'Information' in data:
            raise Exception('API limit reached: ' + data['Information'])
        return data

    def _fetch_all(self, ticker, apikey):
        """Fetch all 4 endpoints with 1.2s delays for rate limiting"""
        import time

        print(f"\n  📊 Fetching data for {ticker}...", flush=True)

        print(f"  [1/4] Overview + Quote", flush=True)
        overview = self._av_fetch({'function': 'OVERVIEW', 'symbol': ticker, 'apikey': apikey})
        if not overview.get('Symbol'):
            raise Exception(f'No data found for "{ticker}". Check the symbol — only US-listed stocks are supported on the free tier.')

        time.sleep(1.2)
        quote_data = self._av_fetch({'function': 'GLOBAL_QUOTE', 'symbol': ticker, 'apikey': apikey})
        quote = quote_data.get('Global Quote', {})

        time.sleep(1.2)
        print(f"  [2/4] Income Statement", flush=True)
        income = self._av_fetch({'function': 'INCOME_STATEMENT', 'symbol': ticker, 'apikey': apikey})
        if not (income.get('annualReports') or []):
            raise Exception(f'No financial statements found for "{ticker}". This symbol may not be supported on the Alpha Vantage free tier.')

        time.sleep(1.2)
        print(f"  [3/4] Balance Sheet", flush=True)
        balance = self._av_fetch({'function': 'BALANCE_SHEET', 'symbol': ticker, 'apikey': apikey})

        time.sleep(1.2)
        print(f"  [4/4] Cash Flow", flush=True)
        cashflow = self._av_fetch({'function': 'CASH_FLOW', 'symbol': ticker, 'apikey': apikey})

        print(f"  ✅ All data fetched for {ticker}\n", flush=True)

        return {
            'overview': overview,
            'quote': quote,
            'income': income,
            'balance': balance,
            'cashflow': cashflow,
        }


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    print("\n" + "═" * 55)
    print("  QuantaValue — Investment Analysis Dashboard")
    print("═" * 55)
    print(f"  Server: http://localhost:{PORT}")
    print(f"  Data:   Alpha Vantage (server-side, no CORS)")
    print(f"  Free key: https://www.alphavantage.co/support/#api-key")
    print("═" * 55)
    print("  Press Ctrl+C to stop\n")

    # Open browser after short delay
    def open_browser():
        time.sleep(1.2)
        webbrowser.open(f'http://localhost:{PORT}')

    threading.Thread(target=open_browser, daemon=True).start()

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        httpd.allow_reuse_address = True
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n  Server stopped. Goodbye!\n")


if __name__ == '__main__':
    main()
