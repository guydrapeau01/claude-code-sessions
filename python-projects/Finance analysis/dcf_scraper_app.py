"""
DCF + Graham Dashboard with AI Analyst Summary
"""
from flask import Flask, request, jsonify
import math, re, requests as req_lib
from datetime import datetime

app = Flask(__name__)
DEFAULT_GROWTH = 0.05

# ── helpers ───────────────────────────────────────────────────────────────────

def to_num(v):
    try:
        n = float(str(v).replace(",","").replace("%","").strip())
        return n if math.isfinite(n) else None
    except:
        return None

def is_pos(n):
    return isinstance(n, (int,float)) and math.isfinite(n) and n > 0

def parse_shorthand(s):
    s = str(s or "").strip().replace(",","")
    mult = {"T":1e12,"B":1e9,"M":1e6,"K":1e3}
    if s and s[-1].upper() in mult:
        n = to_num(s[:-1])
        return n * mult[s[-1].upper()] if n is not None else None
    return to_num(s)

# ── scraper ───────────────────────────────────────────────────────────────────

def scrape_all(symbol):
    from playwright.sync_api import sync_playwright
    import urllib.parse, json

    UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    status = []
    result = {"price":None,"shares":None,"fcf_series":[],"graham":{}}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, args=["--no-sandbox","--disable-dev-shm-usage","--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(user_agent=UA, locale="en-CA", timezone_id="America/Toronto",
                                   extra_http_headers={"Accept-Language":"en-CA,en;q=0.9"})
        ctx.route("**/*", lambda r: r.abort() if "guce.yahoo.com" in r.request.url else r.continue_())
        page = ctx.new_page()

        # ── 1. Get cookies via Playwright, then do API calls via requests ─────
        try:
            # Navigate home to get cookies set
            page.goto("https://ca.finance.yahoo.com/", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1500)

            # Use Playwright's built-in request method to fetch crumb with browser cookies
            pw_cookies = ctx.cookies(["https://ca.finance.yahoo.com", "https://query1.finance.yahoo.com"])
            cookie_str = "; ".join(f"{c['name']}={c['value']}" for c in pw_cookies)

            sess = req_lib.Session()
            sess.headers.update({
                "User-Agent": UA,
                "Accept-Language": "en-CA,en;q=0.9",
                "Referer": "https://ca.finance.yahoo.com/",
                "Cookie": cookie_str,
            })

            crumb_resp = sess.get("https://query1.finance.yahoo.com/v1/test/getcrumb", timeout=10)
            crumb = crumb_resp.text.strip()
            if not crumb or "<" in crumb or len(crumb) > 50:
                raise Exception(f"Bad crumb: {repr(crumb[:40])}")
            status.append(f"✓ Crumb OK")

            # Price
            chart = sess.get(
                f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?interval=1d&range=1d&crumb={urllib.parse.quote(crumb)}",
                timeout=10).json()
            price = chart["chart"]["result"][0]["meta"]["regularMarketPrice"]
            result["price"] = float(price)
            status.append(f"✓ Price: {price}")

            # Fundamentals
            qs_raw = sess.get(
                f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{symbol}"
                f"?modules=defaultKeyStatistics%2CfinancialData%2CsummaryDetail%2CincomeStatementHistory"
                f"&crumb={urllib.parse.quote(crumb)}",
                timeout=15).json()
            qs_list = (qs_raw.get("quoteSummary") or {}).get("result") or []
            if not qs_list:
                raise Exception(f"quoteSummary empty: {str((qs_raw.get('quoteSummary') or {}).get('error'))[:80]}")
            r0 = qs_list[0]
            ks  = r0.get("defaultKeyStatistics", {})
            fd  = r0.get("financialData", {})
            sd  = r0.get("summaryDetail", {})
            inc = r0.get("incomeStatementHistory", {}).get("incomeStatementHistory", [])

            def raw(d, k): return d.get(k, {}).get("raw") if isinstance(d.get(k), dict) else None

            shares = raw(ks,"impliedSharesOutstanding") or raw(ks,"sharesOutstanding")
            result["shares"] = float(shares) if shares else None
            status.append(f"✓ Shares: {result['shares']}" if result["shares"] else "✗ Shares MISSING")

            result["graham"] = {
                "price": result["price"],
                "eps":              raw(ks,"trailingEps"),
                "bvps":             raw(ks,"bookValue"),
                "pe":               raw(sd,"trailingPE"),
                "pb":               raw(ks,"priceToBook"),
                "currentRatio":     raw(fd,"currentRatio"),
                "debtToEquity":     raw(fd,"debtToEquity"),
                "totalDebt":        raw(fd,"totalDebt"),
                "totalCash":        raw(fd,"totalCash"),
                "revenue":          raw(fd,"totalRevenue"),
                "grossMargins":     raw(fd,"grossMargins"),
                "operatingMargins": raw(fd,"operatingMargins"),
                "profitMargins":    raw(fd,"profitMargins"),
                "ebitdaMargins":    raw(fd,"ebitdaMargins"),
                "earningsGrowth":   raw(fd,"earningsGrowth"),
                "revenueGrowth":    raw(fd,"revenueGrowth"),
                "returnOnEquity":   raw(fd,"returnOnEquity"),
                "returnOnAssets":   raw(fd,"returnOnAssets"),
                "freeCashflow":     raw(fd,"freeCashflow"),
                "dividendYield":    raw(sd,"dividendYield"),
                "payoutRatio":      raw(sd,"payoutRatio"),
                "analystTarget":    raw(fd,"targetMeanPrice"),
                "recommendationKey": fd.get("recommendationKey"),
                "numberOfAnalystOpinions": raw(fd,"numberOfAnalystOpinions"),
                "netIncome": [s.get("netIncome",{}).get("raw") for s in inc],
                "incDates":  [s.get("endDate",{}).get("fmt") for s in inc],
            }
            status.append(f"✓ Fundamentals: EPS={result['graham']['eps']} BVPS={result['graham']['bvps']}")
        except Exception as e:
            status.append(f"✗ API FAILED: {type(e).__name__}: {str(e)[:120]}")

        # ── 2. FCF from HTML ──────────────────────────────────────────────────
        try:
            page.goto(f"https://ca.finance.yahoo.com/quote/{symbol}/cash-flow", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(3000)
            data = page.evaluate("""
                () => {
                    const hRow = document.querySelector('[class*="tableHeader"] [class*="row"]');
                    const headers = hRow ? Array.from(hRow.querySelectorAll('[class*="column"]')).map(c=>c.textContent.trim()) : [];
                    const fcfRow = Array.from(document.querySelectorAll('[class*="row"]'))
                        .find(r => r.textContent.trim().startsWith('Free Cash Flow') && r.textContent.trim().length < 400);
                    const values = fcfRow ? Array.from(fcfRow.querySelectorAll('[class*="column"]')).map(c=>c.textContent.trim()) : [];
                    return {headers, values};
                }
            """)
            headers, values = data["headers"], data["values"]
            fcf_series = []
            for i, val in enumerate(values[1:], start=1):
                if i < len(headers) and re.match(r'\d{4}-\d{2}-\d{2}', headers[i]):
                    num = to_num(val.replace(",",""))
                    if num is not None:
                        fcf_series.append((headers[i][:4], num * 1000))
            fcf_series = sorted(fcf_series, key=lambda x: x[0], reverse=True)
            result["fcf_series"] = fcf_series
            if fcf_series:
                status.append("✓ FCF: " + ", ".join(f"{y}:{v/1e9:.1f}B" for y,v in fcf_series))
            else:
                status.append(f"✗ FCF MISSING — headers={headers[:4]} values={values[:4]}")
        except Exception as e:
            status.append(f"✗ FCF MISSING: {type(e).__name__}: {str(e)[:120]}")

        browser.close()
    return result, status

# ── Graham valuation ──────────────────────────────────────────────────────────

def calc_graham(g, price):
    eps  = g.get("eps")
    bvps = g.get("bvps")
    pe   = g.get("pe")
    pb   = g.get("pb")
    cr   = g.get("currentRatio")
    dte  = g.get("debtToEquity")
    div  = g.get("dividendYield")

    # Graham Number
    graham_number = None
    if eps and bvps and eps > 0 and bvps > 0:
        graham_number = math.sqrt(22.5 * eps * bvps)

    # Graham criteria (7 classic tests)
    criteria = []
    def chk(label, value, passed, note=""):
        criteria.append({"label":label,"value":value,"passed":passed,"note":note})

    chk("Adequate Size",        None,  True,  "Manual check recommended")
    chk("Strong Fin. Condition", f"CR: {cr:.2f}" if cr else "N/A", cr >= 2.0 if cr else False, "Current Ratio ≥ 2")
    chk("Earnings Stability",   None,  True,  "Check 10yr EPS history")
    chk("Dividend Record",      f"{div*100:.2f}%" if div else "None", bool(div and div > 0), "Pays dividend")
    chk("Earnings Growth",      f"{g.get('earningsGrowth',0)*100:.1f}%" if g.get('earningsGrowth') else "N/A",
        bool(g.get('earningsGrowth') and g['earningsGrowth'] > 0), "Positive EPS growth")
    chk("Moderate P/E",         f"{pe:.1f}x" if pe else "N/A", pe <= 15 if pe else False, "P/E ≤ 15")
    chk("Moderate P/B",         f"{pb:.2f}x" if pb else "N/A", pb <= 1.5 if pb else False, "P/B ≤ 1.5")

    # P/E × P/B composite
    pepb = (pe * pb) if (pe and pb) else None
    chk("P/E × P/B ≤ 22.5",    f"{pepb:.1f}" if pepb else "N/A", pepb <= 22.5 if pepb else False, "Graham composite")

    mos = ((graham_number - price) / graham_number * 100) if (graham_number and price) else None

    return {
        "graham_number": graham_number,
        "mos": mos,
        "criteria": criteria,
        "passed": sum(1 for c in criteria if c["passed"]),
        "total": len(criteria),
    }

# ── DCF ───────────────────────────────────────────────────────────────────────

def calc_growth_rate(series):
    vals = [v for _,v in series]
    if len(vals) < 2: return None
    e, s, n = vals[0], vals[-1], len(vals)-1
    if is_pos(e) and is_pos(s): return math.pow(e/s, 1.0/n) - 1
    if len(vals) >= 2 and is_pos(vals[1]): return (vals[0]/vals[1]) - 1
    return None

def run_dcf(fcf, growth, discount, years, tgr, shares):
    intrinsic, projected = 0.0, fcf
    rows = []
    for i in range(1, years+1):
        projected *= (1+growth)
        df = math.pow(1+discount, i)
        pv = projected/df
        intrinsic += pv
        rows.append({"year":i,"fcf":projected,"df":round(1/df,5),"pv":pv})
    tv    = projected*(1+tgr)/(discount-tgr)
    pv_tv = tv/math.pow(1+discount, years)
    intrinsic += pv_tv
    iv = intrinsic/shares if is_pos(shares) else None
    return {"intrinsic":intrinsic,"tv":tv,"pv_tv":pv_tv,"iv_per_share":iv,"rows":rows}

# ── API ───────────────────────────────────────────────────────────────────────

@app.route("/api/analyze", methods=["POST"])
def analyze():
    d        = request.json
    symbol   = str(d.get("ticker","")).strip().upper()
    discount = float(d.get("discount",10))/100
    tgr      = float(d.get("tgr",3))/100
    years    = int(d.get("years",10))
    net_debt = float(d.get("net_debt",0))*1e6

    if not symbol:      return jsonify({"error":"Enter a ticker."}), 400
    if discount <= tgr: return jsonify({"error":"WACC must be > Terminal Growth Rate."}), 400

    scraped, status = scrape_all(symbol)

    # Auto-retry once if FCF missing (consent wall timing issue)
    if not scraped["fcf_series"]:
        status.append("⚠ Retrying...")
        scraped2, status2 = scrape_all(symbol)
        if scraped2["fcf_series"]:
            scraped = scraped2
            status = status2

    price  = to_num(d.get("manual_price")) or scraped["price"]
    shares = scraped["shares"]
    if d.get("manual_shares"): shares = to_num(d["manual_shares"])*1e6

    fcf_series = scraped["fcf_series"]
    if d.get("manual_fcf"):
        fcf_series = [(str(datetime.now().year-1), to_num(d["manual_fcf"])*1e6)] + fcf_series

    if not fcf_series:
        return jsonify({"error":"Could not get FCF data. Check scrape log.", "status":status}), 400
    if not is_pos(shares):
        return jsonify({"error":"Could not get Shares. Use manual override.", "status":status}), 400

    if d.get("manual_growth"):
        growth = to_num(d["manual_growth"])/100
        status.append(f"Growth: manual {growth*100:.1f}%")
    else:
        growth = calc_growth_rate(fcf_series)
        if growth is None:
            growth = DEFAULT_GROWTH
            status.append("Growth: DEFAULT 5%")
        else:
            status.append(f"Growth: FCF CAGR {growth*100:.1f}% over {len(fcf_series)-1} yr(s)")

    growth   = max(min(growth,0.60),-0.50)
    base_fcf = fcf_series[0][1]
    dcf      = run_dcf(base_fcf, growth, discount, years, tgr, shares)
    iv       = dcf["iv_per_share"]
    if iv and net_debt: iv -= net_debt/shares
    mos_dcf  = ((iv-price)/iv*100) if (iv and is_pos(price) and is_pos(iv)) else None

    graham = calc_graham(scraped["graham"], price) if price else {}

    return jsonify({
        "symbol":symbol, "price":price, "shares":shares,
        "fcf_series":fcf_series, "base_fcf":base_fcf, "growth":growth,
        "dcf":dcf, "iv_per_share":iv, "mos":mos_dcf,
        "graham":graham, "fundamentals": scraped["graham"],
        "status":status, "timestamp":datetime.utcnow().isoformat()+"Z",
    })

# ── Frontend ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return HTML

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>DCF + Graham Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=Inconsolata:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root{--bg:#07090f;--s1:#0d1117;--s2:#131b27;--border:#1c2a3a;--a1:#e8c547;--a2:#3b9eff;--green:#2dd4a0;--red:#ff5e5e;--text:#d4dce8;--muted:#4a5a6e;--mono:'Inconsolata',monospace;--sans:'Syne',sans-serif}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--text);font-family:var(--mono);min-height:100vh}
.wrap{max-width:1200px;margin:0 auto;padding:40px 24px}
header{margin-bottom:40px;padding-bottom:28px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:16px}
.eyebrow{font-size:11px;letter-spacing:.25em;text-transform:uppercase;color:var(--a1);margin-bottom:8px}
h1{font-family:var(--sans);font-size:clamp(24px,4vw,42px);font-weight:800;color:#fff;line-height:1.05}
h1 em{color:var(--a1);font-style:normal}
.hdr{font-size:11px;color:var(--muted);line-height:2;text-align:right}
.card{background:var(--s1);border:1px solid var(--border);border-radius:10px;padding:26px;margin-bottom:16px}
.sec-lbl{font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:var(--a1);margin-bottom:14px}
.frow{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;align-items:end}
.field label{display:block;font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}
input{background:var(--s2);border:1px solid var(--border);border-radius:5px;color:var(--text);font-family:var(--mono);font-size:13px;padding:9px 12px;width:100%;outline:none;transition:border-color .15s;-moz-appearance:textfield}
input:focus{border-color:var(--a1)}
input::placeholder{color:var(--muted)}
input::-webkit-outer-spin-button,input::-webkit-inner-spin-button{-webkit-appearance:none}
hr.div{border:none;border-top:1px solid var(--border);margin:18px 0}
.btn{background:var(--a1);border:none;border-radius:5px;color:#000;font-family:var(--sans);font-size:12px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;padding:12px 28px;cursor:pointer;width:100%;transition:opacity .15s,transform .1s}
.btn:hover{opacity:.85;transform:translateY(-1px)}
.btn-ai{background:linear-gradient(135deg,#3b9eff,#7c3aed);color:#fff;margin-top:10px}
#spinner{display:none;text-align:center;padding:60px;color:var(--muted);font-size:12px;letter-spacing:.12em}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--a1);margin:0 3px;animation:dp 1.3s infinite ease-in-out}
.dot:nth-child(2){animation-delay:.18s}.dot:nth-child(3){animation-delay:.36s}
@keyframes dp{0%,80%,100%{opacity:.15;transform:scale(.7)}40%{opacity:1;transform:scale(1)}}
#err{display:none;background:rgba(255,94,94,.07);border:1px solid rgba(255,94,94,.3);border-radius:6px;padding:14px 18px;color:var(--red);font-size:13px;margin-bottom:16px}
#results{display:none;animation:fi .4s ease}
@keyframes fi{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.three-col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
@media(max-width:900px){.three-col{grid-template-columns:1fr 1fr}}
@media(max-width:600px){.three-col,.two-col{grid-template-columns:1fr}}
.val-card{background:linear-gradient(140deg,#0d1e34,#081320);border:1px solid rgba(232,197,71,.2);border-radius:10px;padding:28px;position:relative;overflow:hidden}
.val-card::after{content:'';position:absolute;right:-40px;top:-40px;width:160px;height:160px;background:radial-gradient(circle,rgba(232,197,71,.07),transparent 70%);pointer-events:none}
.val-card.blue{border-color:rgba(59,158,255,.25)}
.val-card.blue::after{background:radial-gradient(circle,rgba(59,158,255,.07),transparent 70%)}
.val-sym{font-family:var(--sans);font-size:11px;font-weight:700;letter-spacing:.2em;color:var(--a1);margin-bottom:6px}
.val-card.blue .val-sym{color:var(--a2)}
.val-iv{font-family:var(--sans);font-size:clamp(32px,5vw,52px);font-weight:800;color:#fff;line-height:1}
.val-lbl{font-size:9px;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);margin-bottom:10px}
.val-price{font-size:14px;color:var(--text);margin-bottom:12px}
.mos{border-radius:6px;padding:10px 16px;display:inline-block;margin-top:4px}
.mos.under{background:rgba(45,212,160,.1);border:1px solid rgba(45,212,160,.3)}
.mos.over{background:rgba(255,94,94,.1);border:1px solid rgba(255,94,94,.3)}
.mos.fair{background:rgba(59,158,255,.1);border:1px solid rgba(59,158,255,.3)}
.mos-num{font-family:var(--sans);font-size:20px;font-weight:800}
.under .mos-num{color:var(--green)}.over .mos-num{color:var(--red)}.fair .mos-num{color:var(--a2)}
.mos-vrd{font-size:9px;letter-spacing:.12em;text-transform:uppercase;margin-top:2px}
.under .mos-vrd{color:var(--green)}.over .mos-vrd{color:var(--red)}.fair .mos-vrd{color:var(--a2)}
.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:16px}
.met{background:var(--s1);border:1px solid var(--border);border-radius:7px;padding:14px}
.met-lbl{font-size:9px;letter-spacing:.16em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}
.met-val{font-family:var(--sans);font-size:17px;font-weight:700}
.met-sub{font-size:10px;color:var(--muted);margin-top:2px}
.card-title{font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:var(--a1);margin-bottom:16px;display:flex;align-items:center;gap:7px}
.card-title::before{content:'';width:5px;height:5px;border-radius:50%;background:var(--a1);flex-shrink:0}
.card-title.blue-title{color:var(--a2)}
.card-title.blue-title::before{background:var(--a2)}
.fcf-row{display:grid;grid-template-columns:52px 1fr 100px;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid rgba(28,42,58,.7)}
.fcf-row:last-child{border-bottom:none}
.fcf-yr{font-size:12px;color:var(--muted)}
.bar-bg{height:5px;background:var(--s2);border-radius:3px;overflow:hidden}
.bar-fill{height:100%;border-radius:3px}
.fcf-val{font-size:11px;text-align:right;line-height:1.5}
.tbl{width:100%;border-collapse:collapse;font-size:12px}
.tbl th{font-size:9px;letter-spacing:.13em;text-transform:uppercase;color:var(--muted);text-align:right;padding:0 8px 10px;border-bottom:1px solid var(--border)}
.tbl th:first-child{text-align:left}
.tbl td{padding:7px 8px;text-align:right;border-bottom:1px solid rgba(28,42,58,.5);color:var(--text)}
.tbl td:first-child{text-align:left;color:var(--muted)}
.tbl tr.tv td{color:var(--a1);font-weight:600}
.criteria-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}
.crit{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:6px;background:var(--s2);border:1px solid var(--border)}
.crit.pass{border-color:rgba(45,212,160,.25)}
.crit.fail{border-color:rgba(255,94,94,.2)}
.crit-icon{font-size:14px;flex-shrink:0}
.crit-info{flex:1}
.crit-label{font-size:11px;color:var(--text)}
.crit-val{font-size:10px;color:var(--muted);margin-top:2px}
.fund-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:8px}
.fund-item{padding:10px 14px;background:var(--s2);border-radius:6px;border:1px solid var(--border)}
.fund-lbl{font-size:9px;letter-spacing:.13em;text-transform:uppercase;color:var(--muted);margin-bottom:4px}
.fund-val{font-size:13px;font-weight:600}
.stags{display:flex;flex-wrap:wrap;gap:6px}
.stag{font-size:10px;padding:3px 9px;border-radius:3px;background:var(--s2);border:1px solid var(--border);color:var(--muted)}
.stag.ok{border-color:rgba(45,212,160,.3);color:var(--green)}
.stag.warn{border-color:rgba(232,197,71,.3);color:var(--a1)}
.stag.err{border-color:rgba(255,94,94,.3);color:var(--red)}
#ai-section{display:none}
#ai-output{font-size:13px;line-height:1.8;color:var(--text);white-space:pre-wrap}
#ai-spinner{display:none;color:var(--muted);font-size:12px;padding:20px 0}
.rec-badge{display:inline-block;padding:4px 12px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;margin-left:10px}
.rec-buy{background:rgba(45,212,160,.15);color:var(--green);border:1px solid rgba(45,212,160,.3)}
.rec-hold{background:rgba(232,197,71,.15);color:var(--a1);border:1px solid rgba(232,197,71,.3)}
.rec-sell{background:rgba(255,94,94,.15);color:var(--red);border:1px solid rgba(255,94,94,.3)}
</style>
</head>
<body>
<div class="wrap">
<header>
  <div><div class="eyebrow">Yahoo Finance CA · DCF + Graham + AI</div><h1>Valuation <em>Dashboard</em></h1></div>
  <div class="hdr">Auto-scrapes <b>price · shares · FCF · fundamentals</b><br>DCF · Graham Number · AI analyst summary</div>
</header>
<div class="card">
  <div class="sec-lbl">Ticker &amp; DCF Parameters</div>
  <div class="frow">
    <div class="field"><label>Ticker</label><input id="ticker" type="text" placeholder="AAPL or SHOP.TO" autofocus></div>
    <div class="field"><label>WACC %</label><input id="wacc" type="number" value="10" step="0.1"></div>
    <div class="field"><label>Terminal Growth %</label><input id="tgr" type="number" value="3" step="0.1"></div>
    <div class="field"><label>Projection Years</label><input id="years" type="number" value="10" min="3" max="25"></div>
    <div class="field"><label>Net Debt ($M)</label><input id="net_debt" type="number" placeholder="0" value="0"></div>
  </div>
  <hr class="div">
  <div class="sec-lbl">Manual Overrides</div>
  <div class="frow">
    <div class="field"><label>Price ($)</label><input id="m_price" type="number" placeholder="auto"></div>
    <div class="field"><label>Shares (M)</label><input id="m_shares" type="number" placeholder="auto"></div>
    <div class="field"><label>Base FCF ($M)</label><input id="m_fcf" type="number" placeholder="auto"></div>
    <div class="field"><label>Growth %</label><input id="m_growth" type="number" placeholder="auto"></div>
    <div class="field" style="align-self:end"><button class="btn" onclick="run()">▶ Analyze</button></div>
  </div>
</div>
<div id="err"></div>
<div id="spinner"><div class="dot"></div><div class="dot"></div><div class="dot"></div><p style="margin-top:16px">Scraping Yahoo Finance… (~20s)</p></div>

<div id="results">
  <!-- Valuation cards -->
  <div class="three-col" id="val-cards"></div>

  <!-- Key metrics -->
  <div class="metrics" id="metrics"></div>

  <!-- FCF + DCF table -->
  <div class="two-col">
    <div class="card"><div class="card-title">Historical FCF</div><div id="fcf-chart"></div></div>
    <div class="card"><div class="card-title">DCF Projection</div><div style="overflow-x:auto"><table class="tbl" id="proj-tbl"></table></div></div>
  </div>

  <!-- Graham criteria -->
  <div class="card">
    <div class="card-title blue-title">Graham Criteria</div>
    <div class="criteria-grid" id="graham-criteria"></div>
  </div>

  <!-- Fundamentals -->
  <div class="card">
    <div class="card-title">Key Fundamentals</div>
    <div class="fund-grid" id="fundamentals"></div>
  </div>

  <!-- AI Analyst -->
  <div class="card" id="ai-section">
    <div class="card-title" style="color:var(--a2)">&#x2728; AI Analyst Summary</div>
    <div id="ai-spinner"><div class="dot" style="background:var(--a2)"></div><div class="dot" style="background:var(--a2)"></div><div class="dot" style="background:var(--a2)"></div> Generating expert analysis…</div>
    <div id="ai-output"></div>
  </div>

  <!-- Log -->
  <div class="card"><div class="card-title">Scrape Log</div><div class="stags" id="log"></div></div>
</div>
</div>

<script>
const fmtM=n=>{if(n==null)return'—';const a=Math.abs(n);if(a>=1e12)return'$'+(n/1e12).toFixed(2)+'T';if(a>=1e9)return'$'+(n/1e9).toFixed(2)+'B';if(a>=1e6)return'$'+(n/1e6).toFixed(2)+'M';return'$'+n.toFixed(2)};
const pct=n=>n==null?'—':(n>=0?'+':'')+n.toFixed(1)+'%';
const fmt2=n=>n==null?'—':n.toFixed(2);
const fmtPct=n=>n==null?'—':(n*100).toFixed(1)+'%';

let _lastData = null;

async function run(){
  const ticker=document.getElementById('ticker').value.trim();
  if(!ticker){showErr('Enter a ticker.');return;}
  showErr('');
  document.getElementById('results').style.display='none';
  document.getElementById('ai-section').style.display='none';
  document.getElementById('spinner').style.display='block';
  const body={ticker,discount:+document.getElementById('wacc').value,tgr:+document.getElementById('tgr').value,years:+document.getElementById('years').value,net_debt:+document.getElementById('net_debt').value||0,manual_price:document.getElementById('m_price').value,manual_shares:document.getElementById('m_shares').value,manual_fcf:document.getElementById('m_fcf').value,manual_growth:document.getElementById('m_growth').value};
  try{
    const res=await fetch('/api/analyze',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    const data=await res.json();
    document.getElementById('spinner').style.display='none';
    if(!res.ok){showErr(data.error||'Error');renderLog(data.status||[]);document.getElementById('results').style.display='block';return;}
    _lastData=data;
    render(data,body);
    runAI(data);
  }catch(e){document.getElementById('spinner').style.display='none';showErr('Error: '+e.message);}
}

function showErr(m){const e=document.getElementById('err');e.style.display=m?'block':'none';e.textContent=m;}

function mosCls(mos){return mos==null?'fair':mos>15?'under':mos<-15?'over':'fair';}
function mosVrd(mos){return mos==null?'N/A':mos>15?'UNDERVALUED':mos<-15?'OVERVALUED':'FAIRLY VALUED';}

function render(d,p){
  const iv=d.iv_per_share, gn=d.graham?.graham_number, price=d.price;
  const dcfMos=d.mos, gnMos=d.graham?.mos;

  // Valuation cards
  document.getElementById('val-cards').innerHTML = `
    <div class="val-card">
      <div class="val-sym">${d.symbol} · DCF VALUE</div>
      <div class="val-iv">${iv!=null?'$'+iv.toFixed(2):'—'}</div>
      <div class="val-lbl">Intrinsic Value / Share</div>
      ${price!=null?`<div class="val-price">Market: <b>$${price.toFixed(2)}</b></div>`:''}
      ${dcfMos!=null?`<div class="mos ${mosCls(dcfMos)}"><div class="mos-num">${pct(dcfMos)}</div><div class="mos-vrd">${mosVrd(dcfMos)}</div></div>`:''}
    </div>
    <div class="val-card blue">
      <div class="val-sym" style="color:var(--a2)">${d.symbol} · GRAHAM NUMBER</div>
      <div class="val-iv">${gn!=null?'$'+gn.toFixed(2):'—'}</div>
      <div class="val-lbl">√(22.5 × EPS × BVPS)</div>
      ${price!=null?`<div class="val-price">Market: <b>$${price.toFixed(2)}</b></div>`:''}
      ${gnMos!=null?`<div class="mos ${mosCls(gnMos)}"><div class="mos-num">${pct(gnMos)}</div><div class="mos-vrd">${mosVrd(gnMos)}</div></div>`:''}
    </div>
    <div class="val-card" style="border-color:rgba(45,212,160,.2)">
      <div class="val-sym" style="color:var(--green)">${d.symbol} · ANALYST TARGET</div>
      <div class="val-iv">${d.fundamentals?.analystTarget!=null?'$'+d.fundamentals.analystTarget.toFixed(2):'—'}</div>
      <div class="val-lbl">Mean analyst price target</div>
      ${price!=null?`<div class="val-price">Market: <b>$${price.toFixed(2)}</b></div>`:''}
      ${d.fundamentals?.recommendationKey?`<div class="mos fair"><div class="mos-num" style="font-size:14px;text-transform:uppercase">${d.fundamentals.recommendationKey}</div></div>`:''}
    </div>`;

  // Key metrics
  const f=d.fundamentals||{}, g=d.growth;
  document.getElementById('metrics').innerHTML=[
    {l:'Base FCF',v:fmtM(d.base_fcf),s:'latest annual'},
    {l:'FCF Growth',v:pct(g*100),s:'CAGR',c:g>=0?'var(--green)':'var(--red)'},
    {l:'EPS (TTM)',v:f.eps?'$'+f.eps.toFixed(2):'—',s:'trailing'},
    {l:'P/E Ratio',v:f.pe?f.pe.toFixed(1)+'x':'—',s:f.pe<=15?'✓ Graham':'⚠ >15'},
    {l:'P/B Ratio',v:f.pb?f.pb.toFixed(2)+'x':'—',s:f.pb<=1.5?'✓ Graham':'⚠ >1.5'},
    {l:'Book Value/Sh',v:f.bvps?'$'+f.bvps.toFixed(2):'—',s:'per share'},
    {l:'Current Ratio',v:f.currentRatio?f.currentRatio.toFixed(2):'—',s:f.currentRatio>=2?'✓ Graham':'⚠ <2'},
    {l:'Debt/Equity',v:f.debtToEquity?f.debtToEquity.toFixed(1):'—',s:'%'},
    {l:'ROE',v:fmtPct(f.returnOnEquity),s:'return on equity'},
    {l:'Gross Margin',v:fmtPct(f.grossMargins),s:''},
    {l:'Net Margin',v:fmtPct(f.profitMargins),s:''},
    {l:'Shares',v:d.shares?(d.shares/1e9).toFixed(2)+'B':'—',s:'implied outstanding'},
  ].map(m=>`<div class="met"><div class="met-lbl">${m.l}</div><div class="met-val"${m.c?` style="color:${m.c}"`:''}>
    ${m.v}</div>${m.s?`<div class="met-sub">${m.s}</div>`:''}</div>`).join('');

  // FCF chart
  const series=d.fcf_series||[], maxA=series.length?Math.max(...series.map(([,v])=>Math.abs(v))):1;
  document.getElementById('fcf-chart').innerHTML=series.length?series.map(([yr,v],i)=>{
    const bw=maxA?Math.abs(v)/maxA*100:0,col=v>=0?'var(--green)':'var(--red)';
    const prev=series[i+1]?.[1];
    const gs=prev&&Math.abs(prev)>0?`<span style="color:${(v-prev)/Math.abs(prev)>=0?'var(--green)':'var(--red)'}">${((v-prev)/Math.abs(prev)*100>=0?'+':'')+((v-prev)/Math.abs(prev)*100).toFixed(1)}%</span>`:'';
    return`<div class="fcf-row"><span class="fcf-yr">${yr}</span><div class="bar-bg"><div class="bar-fill" style="width:${bw}%;background:${col}"></div></div><span class="fcf-val" style="color:${col}">${fmtM(v)}<br>${gs}</span></div>`;
  }).join(''):'<p style="color:var(--muted);font-size:12px">No FCF data.</p>';

  // DCF table
  const rows=(d.dcf.rows||[]).map(r=>`<tr><td>Year ${r.year}</td><td>${fmtM(r.fcf)}</td><td>${r.df.toFixed(4)}</td><td>${fmtM(r.pv)}</td></tr>`).join('');
  const tvDF=(1/Math.pow(1+p.discount/100,p.years)).toFixed(4);
  document.getElementById('proj-tbl').innerHTML=`<thead><tr><th style="text-align:left">Period</th><th>Proj FCF</th><th>Disc Factor</th><th>PV</th></tr></thead><tbody>${rows}<tr class="tv"><td>Terminal</td><td>${fmtM(d.dcf.tv)}</td><td>${tvDF}</td><td>${fmtM(d.dcf.pv_tv)}</td></tr></tbody>`;

  // Graham criteria
  const crit=d.graham?.criteria||[];
  document.getElementById('graham-criteria').innerHTML=crit.map(c=>`
    <div class="crit ${c.passed?'pass':'fail'}">
      <span class="crit-icon">${c.passed?'✅':'❌'}</span>
      <div class="crit-info"><div class="crit-label">${c.label}</div>
      <div class="crit-val">${c.value||''} ${c.note?'· '+c.note:''}</div></div>
    </div>`).join('');
  const passed=d.graham?.passed||0, total=d.graham?.total||0;
  document.querySelector('#graham-criteria').insertAdjacentHTML('afterend',
    `<p style="margin-top:12px;font-size:11px;color:var(--muted)">Passed <b style="color:${passed>=5?'var(--green)':'var(--a1)'}">${passed}/${total}</b> Graham criteria</p>`);

  // Fundamentals
  document.getElementById('fundamentals').innerHTML=[
    {l:'Revenue',v:fmtM(f.revenue)},
    {l:'Free Cash Flow',v:fmtM(f.freeCashflow)},
    {l:'Total Debt',v:fmtM(f.totalDebt)},
    {l:'Total Cash',v:fmtM(f.totalCash)},
    {l:'EBITDA Margin',v:fmtPct(f.ebitdaMargins)},
    {l:'Op. Margin',v:fmtPct(f.operatingMargins)},
    {l:'Rev. Growth',v:fmtPct(f.revenueGrowth)},
    {l:'Earnings Growth',v:fmtPct(f.earningsGrowth)},
    {l:'ROA',v:fmtPct(f.returnOnAssets)},
    {l:'Dividend Yield',v:f.dividendYield?fmtPct(f.dividendYield):'None'},
    {l:'Payout Ratio',v:fmtPct(f.payoutRatio)},
    {l:'Analyst Opinions',v:f.numberOfAnalystOpinions||'—'},
  ].map(m=>`<div class="fund-item"><div class="fund-lbl">${m.l}</div><div class="fund-val">${m.v}</div></div>`).join('');

  renderLog(d.status||[]);
  document.getElementById('results').style.display='block';
  document.getElementById('results').scrollIntoView({behavior:'smooth',block:'start'});
}

async function runAI(d){
  document.getElementById('ai-section').style.display='block';
  document.getElementById('ai-spinner').style.display='block';
  document.getElementById('ai-output').textContent='';
  const f=d.fundamentals||{};
  const prompt=`You are a senior equity research analyst. Provide a concise but thorough fundamental analysis of ${d.symbol} based on the following data:

VALUATION
- Market Price: $${d.price?.toFixed(2)}
- DCF Intrinsic Value: $${d.iv_per_share?.toFixed(2)} (MoS: ${d.mos?.toFixed(1)}%)
- Graham Number: $${d.graham?.graham_number?.toFixed(2)} (MoS: ${d.graham?.mos?.toFixed(1)}%)
- Analyst Mean Target: $${f.analystTarget?.toFixed(2)} (${f.recommendationKey})
- P/E: ${f.pe?.toFixed(1)}x | P/B: ${f.pb?.toFixed(2)}x | EPS: $${f.eps?.toFixed(2)}

FINANCIAL HEALTH
- Revenue: ${(f.revenue/1e9)?.toFixed(1)}B | Gross Margin: ${(f.grossMargins*100)?.toFixed(1)}% | Net Margin: ${(f.profitMargins*100)?.toFixed(1)}%
- FCF (TTM): ${(f.freeCashflow/1e9)?.toFixed(1)}B | FCF CAGR: ${(d.growth*100)?.toFixed(1)}%
- Total Debt: ${(f.totalDebt/1e9)?.toFixed(1)}B | Cash: ${(f.totalCash/1e9)?.toFixed(1)}B | D/E: ${f.debtToEquity?.toFixed(1)}
- Current Ratio: ${f.currentRatio?.toFixed(2)} | ROE: ${(f.returnOnEquity*100)?.toFixed(1)}%
- Earnings Growth: ${(f.earningsGrowth*100)?.toFixed(1)}% | Revenue Growth: ${(f.revenueGrowth*100)?.toFixed(1)}%

GRAHAM CRITERIA: ${d.graham?.passed}/${d.graham?.total} passed

Write 3-4 paragraphs covering: (1) valuation verdict comparing DCF vs Graham vs analyst target, (2) financial health and quality of earnings, (3) key risks and red flags, (4) overall investment thesis and recommendation. Be direct, specific, and data-driven. Do not use markdown headers or bullet points — write in flowing paragraphs like a real research note.`;

  try {
    const res = await fetch('/api/ai', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({prompt})
    });
    const data = await res.json();
    document.getElementById('ai-spinner').style.display='none';
    document.getElementById('ai-output').textContent = data.text || 'Analysis unavailable.';
  } catch(e) {
    document.getElementById('ai-spinner').style.display='none';
    document.getElementById('ai-output').textContent = 'AI analysis unavailable: ' + e.message;
  }
}

function renderLog(s){document.getElementById('log').innerHTML=s.map(t=>{const c=t.startsWith('✓')?'ok':t.startsWith('✗')||t.includes('MISSING')?'err':t.includes('DEFAULT')?'warn':'ok';return`<span class="stag ${c}">${t}</span>`}).join('');}
document.addEventListener('keydown',e=>{if(e.key==='Enter')run();});
</script>
</body>
</html>"""

def get_api_key():
    """Read API key from config.txt in same folder as this script."""
    import os
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.txt")
    if os.path.exists(config_path):
        for line in open(config_path):
            line = line.strip()
            if line and not line.startswith("#"):
                return line
    return os.environ.get("ANTHROPIC_API_KEY", "")

@app.route("/api/ai", methods=["POST"])
def ai_analyze():
    try:
        api_key = get_api_key()
        if not api_key:
            return jsonify({"text": "⚠ No API key found. Create a file called config.txt in the same folder as dcf_scraper_app.py and paste your Anthropic API key on the first line. Get one free at console.anthropic.com"}), 200

        prompt = request.json.get("prompt","")
        resp = req_lib.post(
            "https://api.anthropic.com/v1/messages",
            json={"model":"claude-sonnet-4-20250514","max_tokens":1000,
                  "messages":[{"role":"user","content":prompt}]},
            headers={
                "Content-Type":"application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            timeout=60
        )
        data = resp.json()
        if "error" in data:
            return jsonify({"text": f"⚠ API error: {data['error'].get('message','Unknown error')}"}), 200
        text = data.get("content",[{}])[0].get("text","Analysis unavailable.")
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"text": f"AI analysis unavailable: {e}"}), 500

if __name__ == "__main__":
    print("\n  DCF + Graham + AI Dashboard")
    print("  Open: http://localhost:5050\n")
    app.run(debug=False, port=5050)
