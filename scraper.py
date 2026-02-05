from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
from datetime import datetime
import time
import os

URL = "https://www.nseindia.com/option-chain"

EXCEL_FILE = "option_chain_data.xlsx"
FETCH_INTERVAL = 300   # 5 minutes (keep >= 180)

# =========================
# CORE SCRAPER
# =========================
def fetch_option_chain(symbol, expiry, n_rows):
    records = []

    with sync_playwright() as p:
        browser = p.firefox.launch(
            headless=True,
            args=[
                "--disable-http2",
                "--disable-blink-features=AutomationControlled"
            ]
        )

        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1366, "height": 768},
            java_script_enabled=True
        )

        page = context.new_page()

        try:
            # 🔹 STEP 1: Open NSE homepage (MANDATORY)
            page.goto("https://www.nseindia.com", timeout=60000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(4000)

            # 🔹 STEP 2: Navigate internally to Option Chain
            page.goto("https://www.nseindia.com/option-chain", timeout=60000)
            page.wait_for_load_state("networkidle")

            # 🔹 STEP 3: Select symbol
            page.wait_for_selector("#equityStock", timeout=20000)
            page.select_option("#equityStock", symbol)

            # 🔹 STEP 4: Select expiry
            page.wait_for_selector("#expirySelect", timeout=20000)
            page.select_option("#expirySelect", expiry)

            # 🔹 STEP 5: Wait for table
            page.wait_for_selector("table tbody tr", timeout=20000)

            rows = page.locator("table tbody tr")
            total = rows.count()

            if total < n_rows * 2:
                print("⚠ Not enough rows")
                return pd.DataFrame()

            indices = list(range(n_rows)) + list(range(total - n_rows, total))

            for i in indices:
                row = rows.nth(i)
                cells = row.locator("td")

                records.append({
                    "symbol": symbol,
                    "expiry": expiry,
                    "strike_price": cells.nth(11).inner_text(),
                    "call_oi": cells.nth(1).inner_text(),
                    "call_ltp": cells.nth(5).inner_text(),
                    "put_ltp": cells.nth(17).inner_text(),
                    "put_oi": cells.nth(21).inner_text(),
                    "captured_at": datetime.now()
                })

        except Exception as e:
            print("❌ NSE error:", e)

        finally:
            context.close()
            browser.close()

    return pd.DataFrame(records)

# =========================
# EXCEL APPEND (SAFE)
# =========================
def append_to_excel(df):
    if df.empty:
        return

    if os.path.exists(EXCEL_FILE):
        existing = pd.read_excel(EXCEL_FILE)
        df = pd.concat([existing, df], ignore_index=True)

    df.to_excel(EXCEL_FILE, index=False)

# =========================
# MAIN LOOP
# =========================
if __name__ == "__main__":

    # ---- USER INPUTS ----
    SYMBOL = "NIFTY"
    EXPIRY = "29-Feb-2026"   # must match dropdown exactly
    ROWS_REQUIRED = 5       # top N + bottom N

    print("🚀 NSE Option Chain Scraper Started")

    while True:
        print(f"⏳ Fetching at {datetime.now().strftime('%H:%M:%S')}")

        df = fetch_option_chain(
            symbol=SYMBOL,
            expiry=EXPIRY,
            n_rows=ROWS_REQUIRED
        )

        append_to_excel(df)

        print(f"✅ Saved {len(df)} rows")
        print("🛑 Waiting...\n")

        time.sleep(FETCH_INTERVAL)
