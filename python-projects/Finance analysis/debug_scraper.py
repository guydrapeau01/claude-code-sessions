import sys
import traceback

# Keep window open no matter what
try:
    symbol = sys.argv[1] if len(sys.argv) > 1 else "AAPL"
    print(f"Testing symbol: {symbol}\n")

    print("[1] Testing playwright import...")
    from playwright.sync_api import sync_playwright
    print("    OK")

    print("[2] Launching browser...")
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=True, args=["--no-sandbox"])
    print("    OK")

    print("[3] Creating context + page...")
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        locale="en-CA",
    )
    page = ctx.new_page()
    print("    OK")

    print("[4] Loading Yahoo home...")
    page.goto("https://ca.finance.yahoo.com/", wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)
    print(f"    Title: {page.title()}")
    print(f"    URL:   {page.url}")

    print(f"\n[5] Loading cash-flow page...")
    page.goto(f"https://ca.finance.yahoo.com/quote/{symbol}/cash-flow", wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(3000)
    print(f"    Title: {page.title()}")
    print(f"    URL:   {page.url}")

    print("\n[6] Checking page content...")
    result = page.evaluate("""
        () => {
            const rows = Array.from(document.querySelectorAll('[class*="row"]'));
            const fcf = rows.find(r => r.textContent.trim().startsWith('Free Cash Flow') && r.textContent.trim().length < 400);
            return {
                totalRows: rows.length,
                fcfFound: !!fcf,
                fcfText: fcf ? fcf.textContent.trim().substring(0, 150) : null,
                bodySnippet: document.body.innerText.substring(0, 200)
            };
        }
    """)
    print(f"    Total rows: {result['totalRows']}")
    print(f"    FCF found:  {result['fcfFound']}")
    print(f"    FCF text:   {result['fcfText']}")
    print(f"    Body:       {result['bodySnippet'][:150]}")

    page.screenshot(path="debug_step.png")
    print("\n    Screenshot saved: debug_step.png")

    browser.close()
    pw.stop()
    print("\nDone.")

except Exception as e:
    print(f"\nERROR: {e}")
    traceback.print_exc()

input("\nPress Enter to close...")
