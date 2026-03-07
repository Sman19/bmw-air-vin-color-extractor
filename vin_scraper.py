import csv
import os
import random
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# =========================================================
# CONFIG
# =========================================================

INPUT_CSV = "vin_input.csv"
OUTPUT_CSV = "vin_output_with_colors.csv"
BMW_AIR_URL = "https://your-bmw-air-url-here.com"

DAILY_LIMIT = 50          # change to 100 if you want
DELAY_MIN = 7             # random wait min seconds
DELAY_MAX = 18            # random wait max seconds
SEARCH_LOAD_WAIT = 3      # base wait after pressing Enter
HEADLESS = False
SLOW_MO_MS = 150
DEFAULT_TIMEOUT_MS = 15000

# =========================================================
# SELECTORS - THESE WILL PROBABLY NEED ADJUSTMENT
# =========================================================

SEARCH_INPUT_SELECTOR = 'input[type="text"]'

# Try a few possible label spellings / layouts
EXTERIOR_COLOR_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Exterior Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Ext. Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Paint")]/following::*[1]',
]

INTERIOR_COLOR_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Interior Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Upholstery")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Int. Color")]/following::*[1]',
]

# =========================================================
# HELPERS
# =========================================================

def clean_text(value: str) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())

def random_delay():
    delay = random.uniform(DELAY_MIN, DELAY_MAX)
    print(f"Waiting {delay:.1f} seconds...")
    time.sleep(delay)

def ensure_output_exists(output_path: str):
    if not Path(output_path).exists():
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["vin", "last7", "exterior_color", "interior_color", "status"]
            )
            writer.writeheader()

def load_already_processed(output_path: str) -> set:
    processed = set()
    if not Path(output_path).exists():
        return processed

    with open(output_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            vin = clean_text(row.get("vin", ""))
            last7 = clean_text(row.get("last7", ""))
            if vin and last7:
                processed.add((vin, last7))
    return processed

def append_result(output_path: str, row: dict):
    with open(output_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["vin", "last7", "exterior_color", "interior_color", "status"]
        )
        writer.writerow(row)

def read_input_csv(input_path: str) -> list:
    rows = []
    with open(input_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        required = {"vin", "last7"}
        if not reader.fieldnames:
            raise ValueError("Input CSV is missing headers.")
        missing = required - set(h.strip() for h in reader.fieldnames if h)
        if missing:
            raise ValueError(f"Input CSV must contain headers: vin,last7 | Missing: {', '.join(sorted(missing))}")

        for line in reader:
            vin = clean_text(line.get("vin", ""))
            last7 = clean_text(line.get("last7", "")).upper()

            if not vin or not last7:
                continue

            if len(last7) != 7:
                rows.append({
                    "vin": vin,
                    "last7": last7,
                    "skip_reason": "last7 is not 7 characters"
                })
            else:
                rows.append({
                    "vin": vin,
                    "last7": last7,
                    "skip_reason": ""
                })
    return rows

def try_extract_from_xpaths(page, xpaths: list[str]) -> str:
    for xp in xpaths:
        try:
            locator = page.locator(xp).first
            locator.wait_for(state="visible", timeout=4000)
            txt = clean_text(locator.inner_text())
            if txt:
                return txt
        except Exception:
            pass
    return ""

def search_last7(page, last7: str):
    search_box = page.locator(SEARCH_INPUT_SELECTOR).first
    search_box.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
    search_box.click()
    search_box.fill("")
    search_box.type(last7, delay=70)
    search_box.press("Enter")
    time.sleep(SEARCH_LOAD_WAIT)

def extract_colors(page) -> tuple[str, str]:
    exterior = try_extract_from_xpaths(page, EXTERIOR_COLOR_XPATHS)
    interior = try_extract_from_xpaths(page, INTERIOR_COLOR_XPATHS)
    return exterior, interior

# =========================================================
# MAIN
# =========================================================

def main():
    if not Path(INPUT_CSV).exists():
        print(f"Input CSV not found: {INPUT_CSV}")
        return

    ensure_output_exists(OUTPUT_CSV)

    input_rows = read_input_csv(INPUT_CSV)
    already_processed = load_already_processed(OUTPUT_CSV)

    valid_rows = []
    for row in input_rows:
        vin = row["vin"]
        last7 = row["last7"]

        if row["skip_reason"]:
            print(f"Skipping {vin} / {last7}: {row['skip_reason']}")
            continue

        if (vin, last7) in already_processed:
            continue

        valid_rows.append(row)

    if not valid_rows:
        print("No new VINs to process.")
        return

    to_process = valid_rows[:DAILY_LIMIT]

    print(f"Found {len(valid_rows)} unprocessed VINs.")
    print(f"Will process {len(to_process)} today (DAILY_LIMIT={DAILY_LIMIT}).")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)

        print("Opening BMW AIR...")
        page.goto(BMW_AIR_URL)

        input("Log into BMW AIR manually, navigate to the VIN search page, then press ENTER here...")

        for i, row in enumerate(to_process, start=1):
            vin = row["vin"]
            last7 = row["last7"]

            exterior = ""
            interior = ""
            status = "OK"

            print(f"\n[{i}/{len(to_process)}] Processing VIN={vin} | LAST7={last7}")

            try:
                search_last7(page, last7)
                exterior, interior = extract_colors(page)

                if not exterior and not interior:
                    status = "No data found / selector mismatch"

            except PlaywrightTimeoutError:
                status = "Timeout"
            except Exception as e:
                status = f"Error: {str(e)}"

            result_row = {
                "vin": vin,
                "last7": last7,
                "exterior_color": exterior,
                "interior_color": interior,
                "status": status
            }

            append_result(OUTPUT_CSV, result_row)

            print(f"Exterior: {exterior or '[blank]'}")
            print(f"Interior: {interior or '[blank]'}")
            print(f"Status:   {status}")

            if i < len(to_process):
                random_delay()

        browser.close()

    print(f"\nDone. Results saved to: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
