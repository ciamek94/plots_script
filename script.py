import platform
import sys
import os
import time
import datetime
import pandas as pd
from collections import OrderedDict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
from openpyxl.utils import get_column_letter
from openpyxl import Workbook, load_workbook
from webdriver_manager.firefox import GeckoDriverManager

# =======================
# Firefox & Geckodriver
# =======================
def get_firefox_driver():
    """
    Configure Firefox WebDriver with geckodriver v0.34.0.
    Works on both Windows and Linux (GitHub Actions).
    """
    options = Options()
    options.add_argument('--headless')  
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    # Windows: specify Firefox binary if needed
    if platform.system().lower() == "windows":
        firefox_binary = r"C:\Program Files\Mozilla Firefox\firefox.exe"
        options.binary_location = firefox_binary

    # Use a fixed geckodriver version for stability
    service = Service(GeckoDriverManager(version="v0.34.0").install(), timeout=120)
    return webdriver.Firefox(service=service, options=options)

driver = get_firefox_driver()

# =======================
# CONFIGURATION
# =======================
BASE_LINK_KRAKOW = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/krakowski?limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC&mapBounds=19.88207268057927%2C50.13765811720768%2C19.522553426522684%2C49.996078810825225'
BASE_LINK_WIELICKI = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/wielicki?distanceRadius=5&limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC&mapBounds=20.384339742267844%2C50.01972870299939%2C20.25464958097848%2C49.96857906158435'

KRAKOW_COORDS = (50.0647, 19.9450)
EXCEL_FILE = 'wyniki_ofert_z_filtra.xlsx'
HEADERS = [
    'Tytuł',
    'Lokalizacja',
    'Cena pierwszego znalezienia',
    'Data pierwszego znalezienia',
    'Data ostatniej aktualizacji',
    'Cena ostatniej aktualizacji',
    'Odległość od Krakowa (km)',
    'Aktywne',
    'Link'
]

geolocator = Nominatim(user_agent="dzialki_skrypt")

# =======================
# EXCEL HANDLING
# =======================
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
                ws.column_dimensions[get_column_letter(col_idx)].width = max(15, len(HEADERS[col_idx-1]) + 2)
        wb.save(EXCEL_FILE)
        print(f"Created new Excel file: '{EXCEL_FILE}' with sheets and headers.")
    else:
        print(f"Excel file '{EXCEL_FILE}' already exists.")

# =======================
# UTILITY FUNCTIONS
# =======================
def parse_price(price_str):
    return int(price_str.replace(' ', '').replace('zł', '').replace('PLN', '').strip())

def extract_town_from_location(location):
    parts = [p.strip() for p in location.split(',')]
    if len(parts) == 1:
        return parts[0]
    else:
        for i, part in enumerate(parts):
            if 'ul.' in part.lower():
                return parts[1] if i == 0 and len(parts) > 1 else parts[0]
        return parts[0]

def safe_geocode(location, max_retries=3):
    for _ in range(max_retries):
        try:
            return geolocator.geocode(location, exactly_one=False)
        except GeocoderTimedOut:
            time.sleep(1)
    return None

def get_distance_to_krakow(town_name):
    places = safe_geocode(f"{town_name}, Polska")
    if not places:
        print(f"Geocoding error for '{town_name}'")
        return None
    min_dist = float('inf')
    for place in places:
        if place and hasattr(place, 'point'):
            distance = geodesic(KRAKOW_COORDS, (place.latitude, place.longitude)).km
            min_dist = min(min_dist, distance)
    return round(min_dist, 2) if min_dist != float('inf') else None

# =======================
# SCRAPING OFFERS
# =======================
def scrape_offers(base_link, district_name):
    results = []
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    print(f"Scraping offers for: {district_name}")
    try:
        driver.get(base_link)
        time.sleep(5)
        selector = 'a[data-cy="listing-item-link"]' if district_name != 'powiat wielicki' else 'a[href*="/pl/oferta/"]'
        links_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
        )
        listing_links = list(OrderedDict.fromkeys(elem.get_attribute('href') for elem in links_elements))
        if not listing_links:
            print(f"[{district_name}] No data found.")
            return results
        print(f"Found {len(listing_links)} unique offers on the page for {district_name}.")
        for idx, url in enumerate(listing_links, start=1):
            print(f"Fetching offer {idx}/{len(listing_links)}: {url}")
            driver.get(url)
            time.sleep(3)
            try:
                title = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'h1'))).text
            except TimeoutException:
                title = 'No title'
            try:
                price_raw = driver.find_element(By.CSS_SELECTOR, 'strong[data-cy="adPageHeaderPrice"]').text
                price = parse_price(price_raw)
            except NoSuchElementException:
                price = None
            try:
                location = driver.find_element(By.CSS_SELECTOR, 'div[data-sentry-component="MapLink"] a').text
            except NoSuchElementException:
                location = 'No location'
            town_name = extract_town_from_location(location)
            distance_km = get_distance_to_krakow(town_name) or 0.0
            results.append({
                'Tytuł': title,
                'Lokalizacja': location,
                'Cena pierwszego znalezienia': price,
                'Data pierwszego znalezienia': today_str,
                'Data ostatniej aktualizacji': today_str,
                'Cena ostatniej aktualizacji': price,
                'Odległość od Krakowa (km)': distance_km,
                'Aktywne': True,
                'Link': url
            })
    except TimeoutException:
        print(f"[{district_name}] Timeout: No offers found.")
    except Exception as e:
        print(f"[{district_name}] Error: {e}")
    return results

# =======================
# UPDATE EXCEL SHEET
# =======================
def update_sheet(results, sheet_name):
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        existing_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame(columns=HEADERS)
    else:
        existing_df = pd.DataFrame(columns=HEADERS)

    new_rows = []
    for result in results:
        matched_rows = existing_df[(existing_df['Tytuł'] == result['Tytuł']) &
                                   (existing_df['Cena ostatniej aktualizacji'] == result['Cena ostatniej aktualizacji'])]
        if not matched_rows.empty:
            idx = matched_rows.index[0]
            existing_df.at[idx, 'Data ostatniej aktualizacji'] = today_str
            existing_df.at[idx, 'Aktywne'] = True
        else:
            new_rows.append(result)

    if new_rows:
        existing_df = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)

    existing_links = [r['Link'] for r in results]
    if not existing_df.empty and 'Link' in existing_df.columns:
        existing_df['Aktywne'] = existing_df['Link'].apply(lambda l: l in existing_links)

    # Save to Excel
    if os.path.exists(EXCEL_FILE):
        book = load_workbook(EXCEL_FILE)
        if sheet_name in book.sheetnames:
            std = book[sheet_name]
            book.remove(std)
        book.save(EXCEL_FILE)
        book.close()

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        existing_df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(existing_df.columns, 1):
            max_length = max([len(str(cell)) for cell in existing_df[col].astype(str).values] + [len(col)])
            worksheet.column_dimensions[get_column_letter(idx)].width = max_length + 2

    print(f"Saved {len(existing_df)} offers to sheet '{sheet_name}'")

# =======================
# MAIN EXECUTION
# =======================
if __name__ == "__main__":
    create_excel_with_sheets()
    try:
        results_krakow = scrape_offers(BASE_LINK_KRAKOW, 'powiat krakowski')
        update_sheet(results_krakow, 'powiat krakowski')

        results_wielicki = scrape_offers(BASE_LINK_WIELICKI, 'powiat wielicki')
        update_sheet(results_wielicki, 'powiat wielicki')

        print(f"✅ Data saved to '{EXCEL_FILE}' with sheets: powiat krakowski, powiat wielicki")
    finally:
        driver.quit()