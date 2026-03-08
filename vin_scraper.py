import csv
import random
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# =========================================================
# CONFIG
# =========================================================

INPUT_CSV = "2010_E90_M3_master_cleaned.csv"
OUTPUT_CSV = "2010_E90_M3_air_filled.csv"
BMW_AIR_URL = "https://bmwtechinfo.bmwgroup.com/tisUI/#/login"

DAILY_LIMIT = 50
DELAY_MIN = 7
DELAY_MAX = 18
SEARCH_LOAD_WAIT = 5
HEADLESS = False
SLOW_MO_MS = 150
DEFAULT_TIMEOUT_MS = 20000
HAS_HEADER = False

# Optional filter:
# Leave blank "" to run all eligible rows
# Examples:
# RUN_FILTER = "E202200-E202257"
# RUN_FILTER = "E202200,E202205,E202244"
RUN_FILTER = "E202200-E202257"

# =========================================================
# COLUMN INDEXES (0-based)
# Excel E = 4, G = 6, I = 8, K = 10
# E = paint code
# G = upholstery code
# I = full VIN
# K = last 7
# =========================================================

COL_PAINT_CODE = 4   # E
COL_UPHOLSTERY = 6   # G
COL_FULL_VIN = 8     # I
COL_LAST7 = 10       # K

# =========================================================
# SELECTORS
# =========================================================

# Top-left VIN search box on AIR search screen
SEARCH_INPUT_SELECTOR = 'input[placeholder="Start new search"]'

PAINT_CODE_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Paint code")]/following-sibling::*[1]',
    'xpath=//*[contains(normalize-space(),"Paint code")]/following::*[1]',
]

UPHOLSTERY_CODE_XPATHS = [
    'xpath=//*[contains(normalize-space(),"Upholstery code")]/following-sibling::*[1]',
    'xpath=//*[contains(normalize-space(),"Upholstery code")]/following::*[1]',
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
            locator.wait_for(state="visible", timeout=5000)
            txt = clean_text(locator.inner_text())
            if txt:
                return txt
        except Exception:
            pass
    return ""

def extract_full_vin(page):
    """
    On the result page, the VIN appears directly under the vehicle title,
    and it includes the current last7. We look for the first visible text
    that starts with WBS/WBA/WBZ/WBX etc. and is 17 chars.
    """
    prefixes = ("WBS", "WBA", "WBX", "WBY", "WBM", "5YM", "4US")
    try:
        texts = page.locator("body *").all_inner_texts()
        for t in texts:
            txt = clean_text(t)
            if len(txt) == 17 and txt[:3].upper() in prefixes:
                return txt
    except Exception:
        pass
    return ""

def wait_for_results_page(page):
    """
    Wait until the result/details page shows at least one of the expected fields.
    """
    page.wait_for_load_state("domcontentloaded")
    deadline = time.time() + 15

    while time.time() < deadline:
        try:
            paint = try_extract_from_xpaths(page, PAINT_CODE_XPATHS)
            upholstery = try_extract_from_xpaths(page, UPHOLSTERY_CODE_XPATHS)
            vin = extract_full_vin(page)

            if paint or upholstery or vin:
                return True
        except Exception:
            pass

        time.sleep(0.75)

    return False

def search_last7(page, last7: str):
    # There are multiple "Start new search" fields.
    # The VIN field is the first visible text input on the left.
    search_box = page.locator('input[type="text"]').nth(0)
    search_box.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
    search_box.click()
    search_box.fill("")
    search_box.type(last7, delay=70)
    search_box.press("Enter")
    time.sleep(SEARCH_LOAD_WAIT)

def extract_vehicle_data(page):
    full_vin = extract_full_vin(page)
    paint_code = try_extract_from_xpaths(page, PAINT_CODE_XPATHS)
    upholstery_code = try_extract_from_xpaths(page, UPHOLSTERY_CODE_XPATHS)
    return full_vin, paint_code, upholstery_code

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
    max_needed_cols = max(COL_PAINT_CODE, COL_UPHOLSTERY, COL_FULL_VIN, COL_LAST7) + 1
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
            existing_paint = clean_text(row[COL_PAINT_CODE])
            existing_upholstery = clean_text(row[COL_UPHOLSTERY])

            if not last7:
                print(f"Row {row_num + 1}: blank last7 in column K, skipping.")
                continue

            # Normalize values like 202423 -> E202423
            if len(last7) == 6 and last7.isdigit():
                last7 = "E" + last7
            elif len(last7) == 7 and last7[0].isalpha():
                pass
            else:
                print(f"Row {row_num + 1}: invalid last7 '{last7}', skipping.")
                continue

            if selected_last7 and last7 not in selected_last7:
                continue

            # Skip if all three target fields already exist
            if existing_vin and existing_paint and existing_upholstery:
                print(f"Row {row_num + 1}: already filled, skipping.")
                continue

            full_vin = ""
            paint_code = ""
            upholstery_code = ""

            print(f"\nProcessing row {row_num + 1} | LAST7={last7}")

            try:
                search_last7(page, last7)

                if not wait_for_results_page(page):
                    print(f"Row {row_num + 1}: Timeout waiting for result details page")
                else:
                    full_vin, paint_code, upholstery_code = extract_vehicle_data(page)

                    if full_vin:
                        row[COL_FULL_VIN] = full_vin
                    if paint_code:
                        row[COL_PAINT_CODE] = paint_code
                    if upholstery_code:
                        row[COL_UPHOLSTERY] = upholstery_code

                    print(f"Full VIN:         {full_vin or '[blank]'}")
                    print(f"Paint code:       {paint_code or '[blank]'}")
                    print(f"Upholstery code:  {upholstery_code or '[blank]'}")

                    if not full_vin and not paint_code and not upholstery_code:
                        print("Status: selector mismatch or no data found")

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
