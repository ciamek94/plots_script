import importlib.util
import subprocess
import sys
import os

# 📦 Auto-install required libraries if missing
required_packages = ['selenium', 'pandas', 'geopy', 'openpyxl']
for package in required_packages:
    if importlib.util.find_spec(package) is None:
        print(f"📦 Installing: {package}")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

# 📚 Imports
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import pandas as pd
import time
import datetime
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut

from openpyxl.utils import get_column_letter
from openpyxl import Workbook, load_workbook
from collections import OrderedDict

# 📂 Relative paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GECKODRIVER_PATH = os.path.join(SCRIPT_DIR, 'geckodriver.exe')
FIREFOX_BINARY_PATH = os.path.join(SCRIPT_DIR, 'firefox.exe')

# 🔄 Fallback to system Firefox if bundled one not found
if not os.path.exists(FIREFOX_BINARY_PATH):
    FIREFOX_BINARY_PATH = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    if not os.path.exists(FIREFOX_BINARY_PATH):
        FIREFOX_BINARY_PATH = '/usr/bin/firefox'  # Linux

# 🌐 Configure headless Firefox
options = Options()
options.binary_location = FIREFOX_BINARY_PATH
options.add_argument('--headless')

service = FirefoxService(executable_path=GECKODRIVER_PATH)
driver = webdriver.Firefox(service=service, options=options)

# 🌍 Geocoder
geolocator = Nominatim(user_agent="dzialki_skrypt", timeout=5)

# 📊 Excel output file
EXCEL_FILE = os.path.join(SCRIPT_DIR, 'wyniki_ofert_z_filtra.xlsx')

# 🔗 Scraping URLs
BASE_LINK_KRAKOW = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/krakowski?limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC&mapBounds=19.88207268057927%2C50.13765811720768%2C19.522553426522684%2C49.996078810825225'
BASE_LINK_WIELICKI = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/wielicki?distanceRadius=5&limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC&mapBounds=20.384339742267844%2C50.01972870299939%2C20.25464958097848%2C49.96857906158435'

# 📍 Krakow coordinates
KRAKOW_COORDS = (50.0647, 19.9450)

# 🧾 Excel headers
HEADERS = [
    'Tytuł', 'Lokalizacja', 'Cena pierwszego znalezienia', 'Data pierwszego znalezienia',
    'Data ostatniej aktualizacji', 'Cena ostatniej aktualizacji', 'Odległość od Krakowa (km)',
    'Aktywne', 'Link'
]

# Create Excel file with sheets if not exists
def create_excel_with_sheets():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        for sheet_name in ['powiat krakowski', 'powiat wielicki']:
            ws = wb.create_sheet(sheet_name)
            for col_idx, header in enumerate(HEADERS, start=1):
                ws.cell(row=1, column=col_idx, value=header)
            for col_idx in range(1, len(HEADERS) + 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = max(15, len(HEADERS[col_idx - 1]) + 2)
        wb.save(EXCEL_FILE)
        print(f"📁 Created Excel file: {EXCEL_FILE}")
    else:
        print(f"📁 Excel file already exists: {EXCEL_FILE}")

# Convert price string to integer
def parse_price(price_str):
    return int(price_str.replace(' ', '').replace('zł', '').replace('PLN', '').strip())

# Try to extract town name from location string
def extract_town_from_location(location):
    parts = [p.strip() for p in location.split(',')]
    for i, part in enumerate(parts):
        if 'ul.' in part.lower():
            return parts[i + 1] if i + 1 < len(parts) else parts[i]
    return parts[-1]

# Retry-safe geocoding
def safe_geocode(location, max_retries=3):
    for _ in range(max_retries):
        try:
            return geolocator.geocode(location, exactly_one=False)
        except GeocoderTimedOut:
            time.sleep(1)
    return None

# Calculate distance to Krakow from given town name
def get_distance_to_krakow(town_name):
    places = safe_geocode(f"{town_name}, Polska")
    if not places:
        print(f"⚠️ Geocoding error: {town_name}")
        return None
    min_dist = float('inf')
    for place in places:
        if place and hasattr(place, 'point'):
            distance = geodesic(KRAKOW_COORDS, (place.latitude, place.longitude)).km
            min_dist = min(min_dist, distance)
    return round(min_dist, 2) if min_dist < float('inf') else None

# Scrape offers from Otodom for a given district
def scrape_offers(base_link, district_name):
    results = []
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    print(f"\n🔍 Scraping: {district_name}")
    try:
        driver.get(base_link)
        time.sleep(5)
        selector = 'a[href*="/pl/oferta/"]' if district_name == 'powiat wielicki' else 'a[data-cy="listing-item-link"]'
        links_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector)))
        listing_links = list(OrderedDict.fromkeys([e.get_attribute('href') for e in links_elements]))

        print(f"🔗 Found {len(listing_links)} links.")

        for idx, url in enumerate(listing_links, 1):
            print(f"➡️ [{idx}/{len(listing_links)}] {url}")
            driver.get(url)
            time.sleep(2)

            title = 'No title'
            try:
                title = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'h1'))).text
            except TimeoutException:
                pass

            try:
                price = parse_price(driver.find_element(By.CSS_SELECTOR, 'strong[data-cy="adPageHeaderPrice"]').text)
            except NoSuchElementException:
                price = None

            try:
                location = driver.find_element(By.CSS_SELECTOR, 'div[data-sentry-component="MapLink"] a').text
            except NoSuchElementException:
                location = 'No location'

            town_name = extract_town_from_location(location)
            distance_km = get_distance_to_krakow(town_name)

            results.append({
                'Tytuł': title, 'Lokalizacja': location, 'Cena pierwszego znalezienia': price,
                'Data pierwszego znalezienia': today_str, 'Data ostatniej aktualizacji': today_str,
                'Cena ostatniej aktualizacji': price, 'Odległość od Krakowa (km)': distance_km,
                'Aktywne': True, 'Link': url
            })
    except Exception as e:
        print(f"❌ Error scraping {district_name}: {e}")
    return results

# Update the sheet in Excel file
def update_sheet(results, sheet_name):
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        existing_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame(columns=HEADERS)
    else:
        existing_df = pd.DataFrame(columns=HEADERS)

    new_rows = []
    for r in results:
        match = existing_df[(existing_df['Tytuł'] == r['Tytuł']) & (existing_df['Cena ostatniej aktualizacji'] == r['Cena ostatniej aktualizacji'])]
        if not match.empty:
            idx = match.index[0]
            existing_df.at[idx, 'Data ostatniej aktualizacji'] = today_str
            existing_df.at[idx, 'Aktywne'] = True
        else:
            new_rows.append(r)

    if new_rows:
        existing_df = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)

    existing_links = [r['Link'] for r in results]
    existing_df['Aktywne'] = existing_df['Link'].apply(lambda l: l in existing_links)

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        wb.save(EXCEL_FILE)
        wb.close()

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        existing_df.to_excel(writer, sheet_name=sheet_name, index=False)
        for idx, col in enumerate(existing_df.columns, 1):
            width = max([len(str(val)) for val in existing_df[col].astype(str)] + [len(col)])
            writer.sheets[sheet_name].column_dimensions[get_column_letter(idx)].width = width + 2

    print(f"✅ Saved {len(existing_df)} offers → sheet: {sheet_name}")

# ▶️ Main program
if __name__ == "__main__":
    create_excel_with_sheets()
    try:
        res_krk = scrape_offers(BASE_LINK_KRAKOW, 'powiat krakowski')
        update_sheet(res_krk, 'powiat krakowski')
        res_wlk = scrape_offers(BASE_LINK_WIELICKI, 'powiat wielicki')
        update_sheet(res_wlk, 'powiat wielicki')
        print(f"\n📁 Saved data to file: {EXCEL_FILE}")
    finally:
        driver.quit()