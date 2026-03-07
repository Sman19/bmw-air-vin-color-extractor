import csv
import time
import random
from playwright.sync_api import sync_playwright

INPUT_CSV = "vin_input.csv"
OUTPUT_CSV = "vin_output_with_colors.csv"
BMW_AIR_URL = "https://your-bmw-air-url-here.com"

# ===== USER SETTINGS =====

DAILY_LIMIT = 100          # change to 50, 100 etc
DELAY_MIN = 8              # minimum seconds between VIN searches
DELAY_MAX = 20             # maximum seconds between VIN searches

# =========================

SEARCH_INPUT_SELECTOR = 'input[type="text"]'
EXTERIOR_COLOR_SELECTOR = 'xpath=//*[contains(text(),"Exterior Color")]/following::*[1]'
INTERIOR_COLOR_SELECTOR = 'xpath=//*[contains(text(),"Interior Color")]/following::*[1]'


def read_csv(file):

    vins = []

    with open(file, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)

        for row in reader:
            vins.append({
                "vin": row["vin"].strip(),
                "last7": row["last7"].strip()
            })

    return vins


def random_delay():

    delay = random.uniform(DELAY_MIN, DELAY_MAX)

    print(f"Waiting {round(delay,1)} seconds...")
    time.sleep(delay)


def main():

    vin_list = read_csv(INPUT_CSV)

    vin_list = vin_list[:DAILY_LIMIT]

    results = []

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False, slow_mo=200)

        context = browser.new_context()

        page = context.new_page()

        print("Opening BMW AIR...")
        page.goto(BMW_AIR_URL)

        input("Log into BMW AIR manually then press ENTER here...")

        for i, item in enumerate(vin_list):

            vin = item["vin"]
            last7 = item["last7"]

            print(f"\nProcessing {i+1}/{len(vin_list)}  |  {last7}")

            page.fill(SEARCH_INPUT_SELECTOR, "")
            page.type(SEARCH_INPUT_SELECTOR, last7)
            page.press(SEARCH_INPUT_SELECTOR, "Enter")

            time.sleep(3)

            try:
                exterior = page.locator(EXTERIOR_COLOR_SELECTOR).inner_text().strip()
            except:
                exterior = ""

            try:
                interior = page.locator(INTERIOR_COLOR_SELECTOR).inner_text().strip()
            except:
                interior = ""

            results.append({
                "vin": vin,
                "last7": last7,
                "exterior_color": exterior,
                "interior_color": interior
            })

            print("Exterior:", exterior)
            print("Interior:", interior)

            random_delay()

        browser.close()

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:

        writer = csv.DictWriter(
            f,
            fieldnames=["vin","last7","exterior_color","interior_color"]
        )

        writer.writeheader()
        writer.writerows(results)

    print("\nFinished! CSV saved.")


if __name__ == "__main__":
    main()
