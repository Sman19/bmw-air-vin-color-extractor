import csv
import random
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# =========================================================
# CONFIG
# =========================================================

INPUT_CSV = "vin_input.csv"
OUTPUT_CSV = "vin_output_filled.csv"
BMW_AIR_URL = "https://your-bmw-air-url-here.com"

DAILY_LIMIT = 50
DELAY_MIN = 7
DELAY_MAX = 18
SEARCH_LOAD_WAIT = 3
HEADLESS = False
SLOW_MO_MS = 150
DEFAULT_TIMEOUT_MS = 15000
HAS_HEADER = True

# =========================================================
# FILTER MODE
# Leave blank "" to run all eligible rows
#
# Examples:
# RUN_FILTER = "E202200-E202257"
# =========================================================

RUN_FILTER = "E202200-E202257"

# =========================================================
# COLUMN INDEXES (0-based)
# Excel E = 4, G = 6, I = 8, K = 10
# =========================================================

COL_EXTERIOR = 4   # E
COL_INTERIOR = 6   # G
COL_FULL_VIN = 8   # I
COL_LAST7 = 10     # K

# =========================================================
# SELECTORS - THESE WILL PROBABLY NEED ADJUSTMENT
# =========================================================

SEARCH_INPUT_SELECTOR = 'input[type="text"]'

FULL_VIN_XPATHS = [
    'xpath=//*[contains(normalize-space(),"VIN")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Vehicle Identification Number")]/following::*[1]',
]

EXTERIOR_COLOR_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Exterior Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Ext. Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Paint")]/following::*[1]',
]

INTERIOR_COLOR_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Interior Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Int. Color")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Upholstery")]/following::*[1]',
    'xpath=//*[contains(normalize-space(),"Leather")]/following::*[1]',
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

def ensure_row_length(row, min_len):
    while len(row) < min_len:
        row.append("")
    return row

def try_extract_from_xpaths(page, xpaths):
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

def extract_vehicle_data(page):
    full_vin = try_extract_from_xpaths(page, FULL_VIN_XPATHS)
    exterior = try_extract_from_xpaths(page, EXTERIOR_COLOR_XPATHS)
    interior = try_extract_from_xpaths(page, INTERIOR_COLOR_XPATHS)
    return full_vin, exterior, interior

def load_csv_rows(path):
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        return list(reader)

def save_csv_rows(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

# =========================================================
# RUN FILTER PARSER
# =========================================================

def expand_last7_range(start_code: str, end_code: str) -> set:
    """
    Example:
    E202200-E202257
    -> E202200, E202201, ... E202257
    """
    start_code = clean_text(start_code).upper()
    end_code = clean_text(end_code).upper()

    if len(start_code) != 7 or len(end_code) != 7:
        raise ValueError(f"Invalid range values: {start_code}-{end_code}")

    start_prefix = start_code[0]
    end_prefix = end_code[0]

    if start_prefix != end_prefix:
        raise ValueError(f"Range prefixes do not match: {start_code}-{end_code}")

    start_num = int(start_code[1:])
    end_num = int(end_code[1:])

    if start_num > end_num:
        start_num, end_num = end_num, start_num

    return {f"{start_prefix}{num:06d}" for num in range(start_num, end_num + 1)}

def parse_run_filter(filter_text: str) -> set:
    """
    Supports:
    - single values: E202244
    - ranges: E202200-E202257
    - mixed: E202200-E202210,E202244,E202250-E202255
    """
    filter_text = clean_text(filter_text)
    if not filter_text:
        return set()

    selected = set()
    parts = [p.strip().upper() for p in filter_text.split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            left, right = part.split("-", 1)
            selected.update(expand_last7_range(left, right))
        else:
            code = clean_text(part).upper()
            if len(code) != 7:
                raise ValueError(f"Invalid last7 in RUN_FILTER: {code}")
            selected.add(code)

    return selected

# =========================================================
# MAIN
# =========================================================

def main():
    if not Path(INPUT_CSV).exists():
        print(f"Input CSV not found: {INPUT_CSV}")
        return

    rows = load_csv_rows(INPUT_CSV)
    if not rows:
        print("CSV is empty.")
        return

    try:
        selected_last7 = parse_run_filter(RUN_FILTER)
    except Exception as e:
        print(f"RUN_FILTER error: {e}")
        return

    if selected_last7:
        print(f"RUN_FILTER active. {len(selected_last7)} last7 values selected.")
    else:
        print("RUN_FILTER blank. Running all eligible rows.")

    start_index = 1 if HAS_HEADER else 0
    max_needed_cols = max(COL_EXTERIOR, COL_INTERIOR, COL_FULL_VIN, COL_LAST7) + 1
    processed_today = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)

        print("Opening BMW AIR...")
        page.goto(BMW_AIR_URL)

        input("Log into BMW AIR manually, navigate to the VIN search page, then press ENTER here...")

        for row_num in range(start_index, len(rows)):
            if processed_today >= DAILY_LIMIT:
                print(f"Reached DAILY_LIMIT={DAILY_LIMIT}. Stopping for today.")
                break

            row = ensure_row_length(rows[row_num], max_needed_cols)

            last7 = clean_text(row[COL_LAST7]).upper()
            existing_vin = clean_text(row[COL_FULL_VIN])
            existing_ext = clean_text(row[COL_EXTERIOR])
            existing_int = clean_text(row[COL_INTERIOR])

            if not last7:
                print(f"Row {row_num + 1}: blank last7 in column K, skipping.")
                continue

            if len(last7) != 7:
                print(f"Row {row_num + 1}: invalid last7 '{last7}', skipping.")
                continue

            if selected_last7 and last7 not in selected_last7:
                continue

            if existing_vin and existing_ext and existing_int:
                print(f"Row {row_num + 1}: already filled, skipping.")
                continue

            full_vin = ""
            exterior = ""
            interior = ""

            print(f"\nProcessing row {row_num + 1} | LAST7={last7}")

            try:
                search_last7(page, last7)
                full_vin, exterior, interior = extract_vehicle_data(page)

                if full_vin:
                    row[COL_FULL_VIN] = full_vin
                if exterior:
                    row[COL_EXTERIOR] = exterior
                if interior:
                    row[COL_INTERIOR] = interior

                print(f"Full VIN:  {full_vin or '[blank]'}")
                print(f"Exterior:  {exterior or '[blank]'}")
                print(f"Interior:  {interior or '[blank]'}")

            except PlaywrightTimeoutError:
                print(f"Row {row_num + 1}: Timeout")
            except Exception as e:
                print(f"Row {row_num + 1}: Error: {e}")

            processed_today += 1

            if processed_today < DAILY_LIMIT:
                random_delay()

        browser.close()

    save_csv_rows(OUTPUT_CSV, rows)
    print(f"\nDone. Saved updated file to: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
