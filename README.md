📘 NSE Option Chain Web Scraping – Full Project Context & Work Done
🏦 Target Website

NSE India

URL family used:

https://www.nseindia.com

https://www.nseindia.com/option-chain/indices?symbol=NIFTY

🎯 Project Goal

Build a production-ready web scraper that:

Collects Option Chain data (Indices → NIFTY)

Data updates every 3–5 minutes

Extracts:

Top N rows

Bottom N rows

Appends data to an Excel file

Runs continuously (later DB migration planned)

Example logic:

If N = 5 → Top 5 + Bottom 5

If N = 10 → Top 10 + Bottom 10

🚫 What Did NOT Work (Important History)
❌ Normal scraping (requests + BeautifulSoup)

NSE blocks bots aggressively

Always returns 403

❌ Direct API usage

APIs are protected / unstable / blocked

Cookies + tokens expire quickly

❌ Playwright Chromium

Errors observed:

ERR_HTTP2_PROTOCOL_ERROR

ERR_CONNECTION_RESET

👉 Chromium is heavily fingerprinted and blocked by NSE.

✅ Final Working Strategy (Approved)
✔ Browser Automation using Playwright + Firefox

Why Firefox:

Different TLS fingerprint

Less aggressive blocking by NSE

Works reliably for NSE, BSE, Moneycontrol

🧠 Core Technical Decisions

Always load homepage first

NSE sets session cookies (Akamai)

Do NOT interact with UI dropdowns

NSE removed/changed selectors (#equityStock)

Use direct Option Chain URL

https://www.nseindia.com/option-chain/indices?symbol=NIFTY


Wait for table rows, not specific IDs

Headless browser

Minimum interval ≥ 3 minutes

🧩 High-Level Flow
Start Script
   ↓
Open NSE Homepage (session cookies)
   ↓
Open Option Chain (Indices → NIFTY)
   ↓
Wait for option chain table
   ↓
Extract all rows
   ↓
Pick Top N + Bottom N
   ↓
Append to Excel
   ↓
Sleep 3–5 minutes
   ↓
Repeat

🧪 Known NSE Errors & Meaning
Error	Meaning
403 Forbidden	Bot detected
ERR_HTTP2_PROTOCOL_ERROR	HTTP/2 blocked
ERR_CONNECTION_RESET	TLS / fingerprint blocked
Timeout waiting for selector	UI changed

All above are expected NSE behavior, not coding mistakes.

🛠️ Final Dependency Stack
pip install playwright pandas openpyxl
playwright install firefox

📄 Production-Ready Scraper Code (Final)
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import time
import os

EXCEL_FILE = "option_chain_data.xlsx"
FETCH_INTERVAL = 300  # 5 minutes

def fetch_option_chain(symbol, n_rows):
    records = []

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True)

        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) "
                "Gecko/20100101 Firefox/122.0"
            ),
            viewport={"width": 1366, "height": 768}
        )

        page = context.new_page()

        try:
            # Step 1: Homepage (mandatory)
            page.goto("https://www.nseindia.com", timeout=60000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)

            # Step 2: Option Chain (Indices)
            page.goto(
                f"https://www.nseindia.com/option-chain/indices?symbol={symbol}",
                timeout=60000
            )
            page.wait_for_load_state("networkidle")

            # Step 3: Wait for table
            page.wait_for_selector("table tbody tr", timeout=30000)

            rows = page.locator("table tbody tr")
            total = rows.count()

            if total < n_rows * 2:
                return pd.DataFrame()

            indices = list(range(n_rows)) + list(range(total - n_rows, total))

            for i in indices:
                cells = rows.nth(i).locator("td")

                records.append({
                    "symbol": symbol,
                    "strike_price": cells.nth(11).inner_text(),
                    "call_oi": cells.nth(1).inner_text(),
                    "call_ltp": cells.nth(5).inner_text(),
                    "put_ltp": cells.nth(17).inner_text(),
                    "put_oi": cells.nth(21).inner_text(),
                    "captured_at": datetime.now()
                })

        except Exception as e:
            print("NSE error:", e)

        finally:
            context.close()
            browser.close()

    return pd.DataFrame(records)

def append_to_excel(df):
    if df.empty:
        return

    if os.path.exists(EXCEL_FILE):
        old = pd.read_excel(EXCEL_FILE)
        df = pd.concat([old, df], ignore_index=True)

    df.to_excel(EXCEL_FILE, index=False)

if __name__ == "__main__":
    SYMBOL = "NIFTY"
    ROWS_REQUIRED = 5

    while True:
        df = fetch_option_chain(SYMBOL, ROWS_REQUIRED)
        append_to_excel(df)
        print(f"Saved {len(df)} rows at {datetime.now()}")
        time.sleep(FETCH_INTERVAL)

📊 Excel Output Schema
Column	Description
symbol	Index name (NIFTY)
strike_price	Strike
call_oi	Call Open Interest
call_ltp	Call LTP
put_ltp	Put LTP
put_oi	Put Open Interest
captured_at	Timestamp
⚠️ Operational Rules (Critical)

❌ No VPN

❌ No parallel scripts

❌ No interval < 3 minutes

✅ Headless Firefox only

✅ Single IP

Violation → NSE IP block.

🔮 Planned Future Enhancements

(Not implemented yet)

ATM-based row selection

Duplicate prevention

Daily Excel rotation

SQLite / MySQL migration

Alert system (OI spike / PCR)

✅ Current Project Status

Architecture finalized ✔

NSE firewall bypassed ✔

Stable scraping approach ✔

Excel append working ✔

Ready for production hardening ✔
