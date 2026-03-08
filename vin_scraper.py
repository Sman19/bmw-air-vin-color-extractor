import csv
import random
import re
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
SEARCH_LOAD_WAIT = 8
HEADLESS = False
SLOW_MO_MS = 150
DEFAULT_TIMEOUT_MS = 25000
HAS_HEADER = False

# Test one VIN first. Change later if it works.
RUN_FILTER = "E202200"

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

def load_csv_rows(path):
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))

def save_csv_rows(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

def expand_last7_range(start_code: str, end_code: str) -> set:
    start_code = clean_text(start_code).upper()
    end_code = clean_text(end_code).upper()

    if len(start_code) != 7 or len(end_code) != 7:
        raise ValueError(f"Invalid range values: {start_code}-{end_code}")

    if start_code[0] != end_code[0]:
        raise ValueError(f"Range prefixes do not match: {start_code}-{end_code}")

    start_num = int(start_code[1:])
    end_num = int(end_code[1:])

    if start_num > end_num:
        start_num, end_num = end_num, start_num

    return {f"{start_code[0]}{num:06d}" for num in range(start_num, end_num + 1)}

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
            if len(part) != 7:
                raise ValueError(f"Invalid last7 in RUN_FILTER: {part}")
            selected.add(part)

    return selected

# =========================================================
# BMW AIR ACTIONS
# =========================================================

def search_last7(page, last7: str):
    # Top-left VIN search field
    search_box = page.locator('input[type="text"]').nth(0)
    search_box.wait_for(state="visible", timeout=DEFAULT_TIMEOUT_MS)
    search_box.click()
    search_box.fill("")
    search_box.type(last7, delay=70)
    search_box.press("Enter")
    time.sleep(SEARCH_LOAD_WAIT)

def extract_vehicle_data(page):
    full_vin = ""
    paint_code = ""
    upholstery_code = ""

    texts = []

    # Main page text
    try:
        txt = page.evaluate("() => document.body ? document.body.innerText : ''")
        if txt:
            texts.append(txt)
    except Exception:
        pass

    # All frame texts
    for frame in page.frames:
        try:
            txt = frame.evaluate("() => document.body ? document.body.innerText : ''")
            if txt:
                texts.append(txt)
        except Exception:
            pass

    all_text = "\n".join(texts)
    all_text = clean_text(all_text)

    vin_match = re.search(r"\b(WBS[A-Z0-9]{14})\b", all_text)
    if vin_match:
        full_vin = vin_match.group(1)

    paint_match = re.search(r"Paint code\s+([0-9]{3})\b", all_text, re.IGNORECASE)
    if paint_match:
        paint_code = paint_match.group(1)

    upholstery_match = re.search(r"Upholstery code\s+([A-Z0-9]{4})\b", all_text, re.IGNORECASE)
    if upholstery_match:
        upholstery_code = upholstery_match.group(1)

    return full_vin, paint_code, upholstery_code

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

            if existing_vin and existing_paint and existing_upholstery:
                print(f"Row {row_num + 1}: already filled, skipping.")
                continue

            print(f"\nProcessing row {row_num + 1} | LAST7={last7}")

            try:
                search_last7(page, last7)

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
                    print("Status: extraction failed")

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
