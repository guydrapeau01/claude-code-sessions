"""
DCF Valuation Dashboard
Growth calculated from historical FCF CAGR (same logic as GAS script)
"""

from flask import Flask, render_template_string, request, jsonify
import requests
import math
import os
from datetime import datetime

app = Flask(__name__)

ALPHAVANTAGE_API_KEY = os.environ.get('AV_API_KEY', 'X7L7XAQ31K35QFG3')
DEFAULT_GROWTH = 0.05  # 5% fallback only when FCF growth cannot be computed

# ─── Helpers ────────────────────────────────────────────────────────────────

def to_num(v):
    try:
        n = float(v)
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None

def is_pos(n):
    return isinstance(n, (int, float)) and math.isfinite(n) and n > 0

def strip_exchange_prefix(s):
    s = str(s or '').strip()
    parts = s.split(':')
    return parts[1].strip() if len(parts) == 2 else s

# ─── Alpha Vantage fetchers ──────────────────────────────────────────────────

def av_fetch(url):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code != 200:
            return None
        js = r.json()
        if js.get('Note') or js.get('Information'):
            return {'_throttled': True}
        return js
    except Exception:
        return None

def av_fetch_price(symbol):
    url = f'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={ALPHAVANTAGE_API_KEY}'
    js = av_fetch(url)
    if js and 'Global Quote' in js:
        p = to_num(js['Global Quote'].get('05. price'))
        if is_pos(p):
            return p
    return None

def av_fetch_overview(symbol):
    url = f'https://www.alphavantage.co/query?function=OVERVIEW&symbol={symbol}&apikey={ALPHAVANTAGE_API_KEY}'
    js = av_fetch(url)
    if not js:
        return None
    so  = to_num(js.get('SharesOutstanding'))
    mkt = to_num(js.get('MarketCapitalization'))
    eps = to_num(js.get('DilutedEPSTTM')) or to_num(js.get('EPS'))
    pe  = to_num(js.get('PERatio')) or to_num(js.get('TrailingPE'))
    return {
        'shares':    so  if is_pos(so)  else None,
        'marketCap': mkt if is_pos(mkt) else None,
        'eps': eps, 'pe': pe
    }

def _extract_fcf_from_report(report):
    """Extract FCF = OCF - |CapEx|, fallback to freeCashFlow field, then OCF proxy."""
    ocf = (to_num(report.get('netCashProvidedByOperatingActivities'))
        or to_num(report.get('operatingCashflow'))
        or to_num(report.get('operatingCashFlow')))
    capex = (to_num(report.get('capitalExpenditures'))
          or to_num(report.get('investmentsInPropertyPlantAndEquipment')))
    if ocf is not None and capex is not None:
        return ocf - abs(capex)
    fcf = to_num(report.get('freeCashFlow'))
    if fcf is not None:
        return fcf
    if ocf is not None:
        return ocf   # last-resort proxy
    return None

def av_fetch_cash_flow(symbol):
    """Fetch raw cash flow JSON (cached across calls via caller)."""
    url = f'https://www.alphavantage.co/query?function=CASH_FLOW&symbol={symbol}&apikey={ALPHAVANTAGE_API_KEY}'
    return av_fetch(url)

def av_fetch_fcf(symbol, cf_json=None):
    """Most-recent annual FCF → TTM fallback. Same logic as GAS avFetchFCF_."""
    js = cf_json or av_fetch_cash_flow(symbol)
    if not js:
        return None

    annual = js.get('annualReports', [])
    if annual:
        fcf = _extract_fcf_from_report(annual[0])
        if fcf is not None and math.isfinite(fcf):
            return fcf

    # TTM: sum of last 4 quarters
    quarterly = js.get('quarterlyReports', [])
    if len(quarterly) >= 4:
        ttm = 0
        for q in quarterly[:4]:
            v = _extract_fcf_from_report(q)
            if v is None:
                ttm = None
                break
            ttm += v
        if ttm is not None and math.isfinite(ttm):
            return ttm

    return None

def av_fetch_growth_from_fcf(symbol, cf_json=None):
    """
    Exact same logic as GAS avFetchGrowthFromFCF_:
      1. Annual CAGR over up to 5 most-recent years (requires both endpoints positive)
      2. Fallback: TTM YoY (last 4 quarters vs prior 4 quarters)
    Returns growth as a decimal (e.g. 0.12 = 12%).
    Also returns the full FCF series for display.
    """
    js = cf_json or av_fetch_cash_flow(symbol)
    if not js:
        return None, []

    # ── Step 1: Annual CAGR ──────────────────────────────────────────────────
    annual = js.get('annualReports', [])
    series = []   # list of (year_str, fcf_value) newest→oldest
    for a in annual:
        if len(series) >= 5:
            break
        fcf = _extract_fcf_from_report(a)
        if fcf is not None and math.isfinite(fcf):
            series.append((a.get('fiscalDateEnding', '')[:4], fcf))

    if len(series) >= 2:
        end_val   = series[0][1]   # most recent
        start_val = series[-1][1]  # oldest
        n         = len(series) - 1
        # CAGR only valid when both endpoints are positive (same guard as GAS)
        if is_pos(end_val) and is_pos(start_val):
            cagr = math.pow(end_val / start_val, 1.0 / n) - 1
            return cagr, series

    # ── Step 2: TTM YoY fallback ─────────────────────────────────────────────
    quarterly = js.get('quarterlyReports', [])
    if len(quarterly) >= 8:
        ttm_curr, ttm_prev = 0, 0
        ok_curr = ok_prev = True
        for q in quarterly[:4]:
            v = _extract_fcf_from_report(q)
            if v is None:
                ok_curr = False
                break
            ttm_curr += v
        for q in quarterly[4:8]:
            v = _extract_fcf_from_report(q)
            if v is None:
                ok_prev = False
                break
            ttm_prev += v
        if ok_curr and ok_prev and is_pos(ttm_prev):
            yoy = (ttm_curr / ttm_prev) - 1
            return yoy, series

    return None, series

# ─── DCF Calculation (exact same equation as GAS) ───────────────────────────

def run_dcf(fcf, growth, discount_rate, years, terminal_growth_rate, shares):
    """
    Mirrors GAS calculateIntrinsicValueAlphaVantage DCF block exactly:

      projectedFCF *= (1 + growth)   ← applied each year
      intrinsicValue += projectedFCF / (1 + discountRate)^i

      terminalValue = projectedFCF * (1 + terminalGrowthRate)
                      / (discountRate - terminalGrowthRate)
      intrinsicValue += terminalValue / (1 + discountRate)^years

      ivPerShare = intrinsicValue / shares
      MOS = (ivPerShare - price) / ivPerShare * 100
    """
    intrinsic_value = 0.0
    projected_fcf   = fcf
    yearly_rows     = []

    for i in range(1, years + 1):
        projected_fcf  *= (1 + growth)
        discount_factor = math.pow(1 + discount_rate, i)
        pv              = projected_fcf / discount_factor
        intrinsic_value += pv
        yearly_rows.append({
            'year':    i,
            'fcf':     projected_fcf,
            'df':      1.0 / discount_factor,
            'pv':      pv,
        })

    terminal_value    = projected_fcf * (1 + terminal_growth_rate) / (discount_rate - terminal_growth_rate)
    pv_terminal       = terminal_value / math.pow(1 + discount_rate, years)
    intrinsic_value  += pv_terminal

    iv_per_share = intrinsic_value / shares if is_pos(shares) else None

    return {
        'intrinsic_value': intrinsic_value,
        'terminal_value':  terminal_value,
        'pv_terminal':     pv_terminal,
        'iv_per_share':    iv_per_share,
        'yearly_rows':     yearly_rows,
    }

# ─── Routes ─────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    raw_ticker    = str(data.get('ticker', '')).strip()
    discount_rate = to_num(data.get('discount_rate'))   # already decimal
    years         = int(data.get('years', 10))
    term_growth   = to_num(data.get('terminal_growth'))  # already decimal

    if not raw_ticker:
        return jsonify({'error': 'Please enter a ticker symbol.'}), 400
    if not is_pos(discount_rate):
        return jsonify({'error': 'Discount rate must be a positive number.'}), 400
    if discount_rate <= term_growth:
        return jsonify({'error': 'Discount rate must be greater than terminal growth rate.'}), 400

    symbol = strip_exchange_prefix(raw_ticker).upper()
    status = []
    result = {'symbol': symbol, 'timestamp': datetime.utcnow().isoformat() + 'Z'}

    # Price
    price = av_fetch_price(symbol)
    status.append('Price OK (AV)' if price else 'Price missing')

    # Overview → shares / EPS / PE
    overview = av_fetch_overview(symbol)
    shares = overview['shares'] if overview else None
    if not is_pos(shares) and overview and is_pos(overview.get('marketCap')) and is_pos(price):
        shares = overview['marketCap'] / price
        status.append('Shares computed from MarketCap/Price')
    else:
        status.append('Shares OK (AV OVERVIEW)' if is_pos(shares) else 'Shares missing')

    eps = overview['eps'] if overview else None
    pe  = overview['pe']  if overview else None
    if pe is None and is_pos(price) and eps:
        try: pe = price / eps
        except: pe = None

    # Cash flow (one API call reused for both FCF and growth)
    cf_json = av_fetch_cash_flow(symbol)

    # FCF
    fcf = av_fetch_fcf(symbol, cf_json)
    status.append('FCF OK (AV)' if fcf is not None else 'FCF missing – enter manually')

    # Growth from FCF (exact same GAS logic)
    growth, fcf_series = av_fetch_growth_from_fcf(symbol, cf_json)
    if growth is not None and math.isfinite(growth):
        status.append(f'Growth OK (AV FCF CAGR: {growth*100:.1f}%)')
    else:
        growth = DEFAULT_GROWTH
        status.append(f'Growth DEFAULT {DEFAULT_GROWTH*100:.0f}% (FCF CAGR unavailable)')

    # Allow manual overrides from the form
    if data.get('manual_fcf') not in (None, ''):
        fcf = to_num(data['manual_fcf'])
        status.append('FCF from manual input')
    if data.get('manual_growth') not in (None, ''):
        growth = to_num(data['manual_growth']) / 100.0
        status.append(f'Growth from manual input: {growth*100:.1f}%')
    if data.get('manual_shares') not in (None, ''):
        shares = to_num(data['manual_shares'])
        status.append('Shares from manual input')

    result.update({
        'price':      price,
        'shares':     shares,
        'eps':        eps,
        'pe':         pe,
        'fcf':        fcf,
        'growth':     growth,
        'fcf_series': fcf_series,
        'status':     status,
    })

    # DCF (only if we have shares + FCF)
    if is_pos(shares) and fcf is not None and math.isfinite(fcf):
        dcf = run_dcf(fcf, growth, discount_rate, years, term_growth, shares)
        iv  = dcf['iv_per_share']
        mos = ((iv - price) / iv * 100) if (is_pos(iv) and is_pos(price)) else None
        result.update({
            'dcf':        dcf,
            'iv_per_share': iv,
            'mos':         mos,
        })
    else:
        result['error_dcf'] = 'Missing FCF or Shares — enter manually and retry.'

    return jsonify(result)

# ─── HTML Template ───────────────────────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DCF Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=Inconsolata:wght@300;400;500&display=swap" rel="stylesheet">
<style>
:root {
  --bg: #07090f;
  --s1: #0d1117;
  --s2: #131b27;
  --border: #1c2a3a;
  --a1: #e8c547;
  --a2: #3b9eff;
  --green: #2dd4a0;
  --red: #ff5e5e;
  --text: #d4dce8;
  --muted: #4a5a6e;
  --mono: 'Inconsolata', monospace;
  --sans: 'Syne', sans-serif;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--text);font-family:var(--mono);min-height:100vh;overflow-x:hidden}

/* ── noise texture overlay ── */
body::after{
  content:'';position:fixed;inset:0;
  background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.04'/%3E%3C/svg%3E");
  pointer-events:none;z-index:9999;opacity:.5
}

.wrap{max-width:1200px;margin:0 auto;padding:40px 24px}

/* ── header ── */
header{margin-bottom:48px;border-bottom:1px solid var(--border);padding-bottom:32px;display:flex;align-items:flex-end;justify-content:space-between;flex-wrap:wrap;gap:16px}
.hd-left .eyebrow{font-size:11px;letter-spacing:.25em;text-transform:uppercase;color:var(--a1);margin-bottom:10px}
h1{font-family:var(--sans);font-size:clamp(28px,4vw,48px);font-weight:800;color:#fff;line-height:1}
h1 em{color:var(--a1);font-style:normal}
.hd-right{font-size:11px;color:var(--muted);line-height:1.8;text-align:right}

/* ── input panel ── */
.input-panel{
  background:var(--s1);border:1px solid var(--border);border-radius:8px;
  padding:28px;margin-bottom:24px;
  display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:20px;align-items:end
}
.field label{display:block;font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);margin-bottom:6px}
.field input{
  background:var(--s2);border:1px solid var(--border);border-radius:4px;
  color:var(--text);font-family:var(--mono);font-size:13px;
  padding:9px 12px;width:100%;outline:none;transition:border-color .15s
}
.field input:focus{border-color:var(--a1)}
.field input::placeholder{color:var(--muted)}

.run-btn{
  background:var(--a1);border:none;border-radius:4px;
  color:#000;font-family:var(--sans);font-size:12px;font-weight:700;
  letter-spacing:.12em;text-transform:uppercase;
  padding:11px 24px;cursor:pointer;white-space:nowrap;
  transition:opacity .15s,transform .1s
}
.run-btn:hover{opacity:.88;transform:translateY(-1px)}
.run-btn:active{transform:translateY(0)}

/* ── spinner ── */
#spinner{display:none;text-align:center;padding:60px;color:var(--muted);font-size:13px;letter-spacing:.1em}
.dot-pulse{display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--a1);margin:0 3px;animation:dp 1.2s infinite}
.dot-pulse:nth-child(2){animation-delay:.2s}
.dot-pulse:nth-child(3){animation-delay:.4s}
@keyframes dp{0%,80%,100%{opacity:.2;transform:scale(.8)}40%{opacity:1;transform:scale(1)}}

/* ── error ── */
#error-box{display:none;background:rgba(255,94,94,.08);border:1px solid rgba(255,94,94,.3);border-radius:6px;padding:16px 20px;color:var(--red);font-size:13px;margin-bottom:20px}

/* ── results ── */
#results{display:none;animation:fadeIn .4s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}

/* hero strip */
.hero{
  background:linear-gradient(135deg,#0c1a2e 0%,#091420 100%);
  border:1px solid var(--border);border-radius:8px;padding:36px;
  margin-bottom:20px;position:relative;overflow:hidden;
  display:grid;grid-template-columns:1fr auto;gap:24px;align-items:center
}
.hero::before{
  content:'';position:absolute;right:-80px;top:-80px;
  width:280px;height:280px;
  background:radial-gradient(circle,rgba(232,197,71,.07) 0%,transparent 70%)
}
.hero-ticker{font-family:var(--sans);font-size:13px;font-weight:700;letter-spacing:.2em;color:var(--a1);margin-bottom:8px}
.hero-iv{font-family:var(--sans);font-size:clamp(44px,7vw,72px);font-weight:800;color:#fff;line-height:1;margin-bottom:4px}
.hero-label{font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--muted)}
.hero-price{font-size:18px;color:var(--text);margin-top:12px}
.hero-price span{color:var(--muted);font-size:12px;margin-left:6px}

.mos-badge{
  border-radius:6px;padding:20px 28px;text-align:center;min-width:140px
}
.mos-badge.under{background:rgba(45,212,160,.08);border:1px solid rgba(45,212,160,.25)}
.mos-badge.over {background:rgba(255,94,94,.08); border:1px solid rgba(255,94,94,.25)}
.mos-badge.fair {background:rgba(59,158,255,.08); border:1px solid rgba(59,158,255,.25)}
.mos-num{font-family:var(--sans);font-size:32px;font-weight:800;line-height:1}
.mos-badge.under .mos-num{color:var(--green)}
.mos-badge.over  .mos-num{color:var(--red)}
.mos-badge.fair  .mos-num{color:var(--a2)}
.mos-lbl{font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);margin-top:6px}
.mos-verdict{font-size:12px;font-weight:700;margin-top:8px}
.mos-badge.under .mos-verdict{color:var(--green)}
.mos-badge.over  .mos-verdict{color:var(--red)}
.mos-badge.fair  .mos-verdict{color:var(--a2)}

/* metrics row */
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px}
.metric{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:16px}
.metric-label{font-size:9px;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);margin-bottom:6px}
.metric-value{font-size:20px;font-family:var(--sans);font-weight:700;color:var(--text)}
.metric-sub{font-size:10px;color:var(--muted);margin-top:2px}

/* two-col layout */
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
@media(max-width:700px){.two-col{grid-template-columns:1fr}.hero{grid-template-columns:1fr}}

/* cards */
.card{background:var(--s1);border:1px solid var(--border);border-radius:8px;padding:24px}
.card-title{font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:var(--a1);margin-bottom:18px;display:flex;align-items:center;gap:8px}
.card-title::before{content:'';width:5px;height:5px;border-radius:50%;background:var(--a1)}

/* FCF series */
.fcf-row{display:grid;grid-template-columns:60px 1fr 80px;align-items:center;gap:12px;padding:7px 0;border-bottom:1px solid rgba(28,42,58,.8)}
.fcf-row:last-child{border-bottom:none}
.fcf-yr{font-size:12px;color:var(--muted)}
.fcf-bar-wrap{height:4px;background:var(--s2);border-radius:2px;overflow:hidden}
.fcf-bar{height:100%;border-radius:2px;transition:width .5s ease}
.fcf-val{font-size:12px;text-align:right}

/* projection table */
.proj-table{width:100%;border-collapse:collapse;font-size:12px}
.proj-table th{font-size:9px;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);text-align:right;padding:0 10px 10px;border-bottom:1px solid var(--border)}
.proj-table th:first-child{text-align:left}
.proj-table td{padding:8px 10px;text-align:right;border-bottom:1px solid rgba(28,42,58,.5);color:var(--text)}
.proj-table td:first-child{text-align:left;color:var(--muted)}
.proj-table tr.terminal td{color:var(--a1)}
.proj-table tr:last-child td{border-bottom:none}

/* status log */
.status-log{font-size:11px;color:var(--muted);line-height:2;margin-top:16px}
.status-log span{display:inline-block;margin-right:8px;padding:2px 8px;border-radius:3px;background:var(--s2);border:1px solid var(--border)}
</style>
</head>
<body>
<div class="wrap">

  <header>
    <div class="hd-left">
      <div class="eyebrow">Intrinsic Valuation Tool</div>
      <h1>DCF <em>Dashboard</em></h1>
    </div>
    <div class="hd-right">
      Growth derived from<br>historical Free Cash Flow CAGR<br>
      <span style="color:var(--a1)">Alpha Vantage</span> data source
    </div>
  </header>

  <!-- Input Panel -->
  <div class="input-panel">
    <div class="field" style="min-width:100px">
      <label>Ticker</label>
      <input id="ticker" type="text" placeholder="AAPL" value="">
    </div>
    <div class="field">
      <label>Discount Rate (WACC %)</label>
      <input id="discount" type="number" value="10" step="0.1">
    </div>
    <div class="field">
      <label>Terminal Growth %</label>
      <input id="term_growth" type="number" value="3" step="0.1">
    </div>
    <div class="field">
      <label>Projection Years</label>
      <input id="years" type="number" value="10" min="3" max="20">
    </div>
    <div class="field">
      <label>Manual FCF ($M) — optional</label>
      <input id="manual_fcf" type="number" placeholder="auto">
    </div>
    <div class="field">
      <label>Manual Growth % — optional</label>
      <input id="manual_growth" type="number" placeholder="auto">
    </div>
    <div class="field">
      <label>Manual Shares (M) — optional</label>
      <input id="manual_shares" type="number" placeholder="auto">
    </div>
    <div class="field" style="align-self:end">
      <button class="run-btn" onclick="runAnalysis()">▶ Analyze</button>
    </div>
  </div>

  <div id="error-box"></div>

  <div id="spinner">
    <div class="dot-pulse"></div><div class="dot-pulse"></div><div class="dot-pulse"></div>
    <p style="margin-top:16px">Fetching data from Alpha Vantage…</p>
  </div>

  <div id="results">
    <!-- Hero -->
    <div class="hero" id="hero-strip"></div>

    <!-- Key Metrics -->
    <div class="metrics" id="metrics-row"></div>

    <!-- FCF series + Projection table -->
    <div class="two-col">
      <div class="card">
        <div class="card-title">Historical FCF Series</div>
        <div id="fcf-series"></div>
      </div>
      <div class="card">
        <div class="card-title">DCF Projection</div>
        <div style="overflow-x:auto"><table class="proj-table" id="proj-table"></table></div>
      </div>
    </div>

    <!-- Status -->
    <div class="card">
      <div class="card-title">Data Source Log</div>
      <div class="status-log" id="status-log"></div>
    </div>
  </div>

</div>
<script>
const fmt = (n, decimals=0) => {
  if (n == null) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e12) return '$' + (n/1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return '$' + (n/1e9).toFixed(2)  + 'B';
  if (abs >= 1e6)  return '$' + (n/1e6).toFixed(2)  + 'M';
  return '$' + n.toFixed(decimals);
};
const pct = n => n == null ? '—' : (n >= 0 ? '+' : '') + n.toFixed(1) + '%';
const fmtFCF = n => {
  if (n == null) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e9)  return (n/1e9).toFixed(2)  + 'B';
  if (abs >= 1e6)  return (n/1e6).toFixed(2)  + 'M';
  return n.toFixed(0);
};

async function runAnalysis() {
  const ticker = document.getElementById('ticker').value.trim();
  if (!ticker) { showError('Please enter a ticker symbol.'); return; }

  showError('');
  document.getElementById('results').style.display = 'none';
  document.getElementById('spinner').style.display = 'block';

  const payload = {
    ticker,
    discount_rate:  parseFloat(document.getElementById('discount').value) / 100,
    terminal_growth: parseFloat(document.getElementById('term_growth').value) / 100,
    years:          parseInt(document.getElementById('years').value),
    manual_fcf:     document.getElementById('manual_fcf').value,
    manual_growth:  document.getElementById('manual_growth').value,
    manual_shares:  document.getElementById('manual_shares').value,
  };

  try {
    const res  = await fetch('/api/analyze', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) });
    const data = await res.json();
    document.getElementById('spinner').style.display = 'none';

    if (!res.ok) { showError(data.error || 'Unknown error'); return; }
    renderResults(data, payload);
  } catch(e) {
    document.getElementById('spinner').style.display = 'none';
    showError('Network error: ' + e.message);
  }
}

function showError(msg) {
  const el = document.getElementById('error-box');
  el.style.display = msg ? 'block' : 'none';
  el.textContent = msg;
}

function renderResults(d, params) {
  // ── Hero ──────────────────────────────────────────────────────────────────
  const iv = d.iv_per_share;
  const mos = d.mos;
  let mosClass = 'fair', mosVerdict = 'FAIRLY VALUED';
  if (mos != null) {
    if (mos > 15)       { mosClass = 'under'; mosVerdict = 'UNDERVALUED'; }
    else if (mos < -15) { mosClass = 'over';  mosVerdict = 'OVERVALUED';  }
  }

  document.getElementById('hero-strip').innerHTML = `
    <div>
      <div class="hero-ticker">${d.symbol}</div>
      <div class="hero-iv">${iv != null ? '$' + iv.toFixed(2) : '—'}</div>
      <div class="hero-label">Intrinsic Value Per Share</div>
      ${d.price != null ? `<div class="hero-price">Market Price: <b>$${d.price.toFixed(2)}</b><span>current</span></div>` : ''}
    </div>
    ${mos != null ? `
    <div class="mos-badge ${mosClass}">
      <div class="mos-num">${pct(mos)}</div>
      <div class="mos-lbl">Margin of Safety</div>
      <div class="mos-verdict">${mosVerdict}</div>
    </div>` : ''}
  `;

  // ── Metrics ───────────────────────────────────────────────────────────────
  const growthPct = d.growth != null ? (d.growth * 100).toFixed(1) + '%' : '—';
  document.getElementById('metrics-row').innerHTML = `
    <div class="metric"><div class="metric-label">FCF (Base)</div><div class="metric-value">${d.fcf != null ? fmtFCF(d.fcf) : '—'}</div><div class="metric-sub">used in projection</div></div>
    <div class="metric"><div class="metric-label">FCF Growth (CAGR)</div><div class="metric-value" style="color:${d.growth>=0?'var(--green)':'var(--red)'}">${growthPct}</div><div class="metric-sub">from historical FCF</div></div>
    <div class="metric"><div class="metric-label">WACC</div><div class="metric-value">${(params.discount_rate*100).toFixed(1)}%</div><div class="metric-sub">discount rate</div></div>
    <div class="metric"><div class="metric-label">Terminal Growth</div><div class="metric-value">${(params.terminal_growth*100).toFixed(1)}%</div></div>
    <div class="metric"><div class="metric-label">PV Terminal Value</div><div class="metric-value">${d.dcf ? fmtFCF(d.dcf.pv_terminal) : '—'}</div></div>
    <div class="metric"><div class="metric-label">Shares Outstanding</div><div class="metric-value">${d.shares != null ? (d.shares/1e6).toFixed(0)+'M' : '—'}</div></div>
    ${d.eps  != null ? `<div class="metric"><div class="metric-label">EPS (TTM)</div><div class="metric-value">$${d.eps.toFixed(2)}</div></div>` : ''}
    ${d.pe   != null ? `<div class="metric"><div class="metric-label">P/E Ratio</div><div class="metric-value">${d.pe.toFixed(1)}x</div></div>` : ''}
  `;

  // ── FCF Series ────────────────────────────────────────────────────────────
  const series = d.fcf_series || [];
  if (series.length) {
    const maxAbs = Math.max(...series.map(([,v]) => Math.abs(v)));
    document.getElementById('fcf-series').innerHTML = series.map(([yr, val], idx) => {
      const barW  = maxAbs > 0 ? Math.abs(val) / maxAbs * 100 : 0;
      const color = val >= 0 ? 'var(--green)' : 'var(--red)';
      const growthStr = idx < series.length - 1
        ? (() => {
            const prev = series[idx+1][1];
            if (!prev) return '';
            const g = ((val - prev) / Math.abs(prev)) * 100;
            return `<span style="color:${g>=0?'var(--green)':'var(--red)'};font-size:10px">${g>=0?'+':''}${g.toFixed(1)}%</span>`;
          })()
        : '';
      return `<div class="fcf-row">
        <span class="fcf-yr">${yr}</span>
        <div class="fcf-bar-wrap"><div class="fcf-bar" style="width:${barW}%;background:${color}"></div></div>
        <span class="fcf-val" style="color:${color}">${fmtFCF(val)} ${growthStr}</span>
      </div>`;
    }).join('');
  } else {
    document.getElementById('fcf-series').innerHTML = '<p style="color:var(--muted);font-size:12px">No annual FCF series available.</p>';
  }

  // ── Projection Table ──────────────────────────────────────────────────────
  if (d.dcf) {
    const rows = d.dcf.yearly_rows.map(r => `
      <tr>
        <td>Year ${r.year}</td>
        <td>${fmtFCF(r.fcf)}</td>
        <td>${r.df.toFixed(4)}</td>
        <td>${fmtFCF(r.pv)}</td>
      </tr>`).join('');
    const tvRow = `<tr class="terminal">
      <td>Terminal</td>
      <td>${fmtFCF(d.dcf.terminal_value)}</td>
      <td>${(1/Math.pow(1+params.discount_rate, params.years)).toFixed(4)}</td>
      <td>${fmtFCF(d.dcf.pv_terminal)}</td>
    </tr>`;
    document.getElementById('proj-table').innerHTML = `
      <thead><tr>
        <th style="text-align:left">Period</th>
        <th>Proj. FCF</th><th>Disc. Factor</th><th>PV of FCF</th>
      </tr></thead>
      <tbody>${rows}${tvRow}</tbody>`;
  }

  // ── Status Log ────────────────────────────────────────────────────────────
  document.getElementById('status-log').innerHTML =
    (d.status || []).map(s => `<span>${s}</span>`).join(' ');

  document.getElementById('results').style.display = 'block';
  document.getElementById('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Allow Enter key
document.addEventListener('keydown', e => { if (e.key === 'Enter') runAnalysis(); });
</script>
</body>
</html>"""

if __name__ == '__main__':
    app.run(debug=True, port=5050)
