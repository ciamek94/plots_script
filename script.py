import requests
import time
import datetime
import os
import pandas as pd
from collections import OrderedDict
from bs4 import BeautifulSoup
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# =======================
# CONFIGURATION
# =======================
BASE_LINK_KRAKOW = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/krakowski?limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'
BASE_LINK_WIELICKI = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/wielicki?distanceRadius=5&limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'

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
    return int(price_str.replace(' ', '').replace('zł', '').replace('PLN', '').replace(',', '').strip())

def extract_town_from_location(location):
    """
    Wyciąga miejscowość z pola lokalizacja (ostatni fragment).
    """
    parts = [p.strip() for p in location.split(',')]
    return parts[-1] if parts else location

def safe_geocode(location, max_retries=3):
    for _ in range(max_retries):
        try:
            return geolocator.geocode(location, exactly_one=False)
        except GeocoderTimedOut:
            time.sleep(1)
    return None

def get_distance_to_krakow(town_name):
    """
    Liczy odległość w km od Krakowa do miejscowości.
    """
    query = f"{town_name}, małopolskie, Polska"
    places = safe_geocode(query)
    if not places:
        print(f"Geocoding error for '{query}'")
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
HEADERS_HTTP = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8"
}

def scrape_offers(base_link, district_name):
    results = []
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    print(f"Scraping offers for: {district_name}")

    try:
        response = requests.get(base_link, headers=HEADERS_HTTP, timeout=30)
        if response.status_code != 200:
            print(f"[{district_name}] HTTP Error {response.status_code}")
            return results

        soup = BeautifulSoup(response.text, "html.parser")
        offer_links = []
        for a in soup.select('a[data-cy="listing-item-link"]'):
            href = a.get('href')
            if href:
                if href.startswith('/'):
                    href = 'https://www.otodom.pl' + href
                offer_links.append(href)

        offer_links = list(OrderedDict.fromkeys(offer_links))  # unique
        print(f"Found {len(offer_links)} unique offers on the page for {district_name}.")

        for idx, url in enumerate(offer_links, start=1):
            print(f"Fetching offer {idx}/{len(offer_links)}: {url}")
            try:
                offer_resp = requests.get(url, headers=HEADERS_HTTP, timeout=30)
                if offer_resp.status_code != 200:
                    print(f"Error {offer_resp.status_code} for {url}")
                    continue

                offer_soup = BeautifulSoup(offer_resp.text, "html.parser")
                title = offer_soup.find('h1').text.strip() if offer_soup.find('h1') else "No title"

                price_tag = offer_soup.select_one('strong[data-cy="adPageHeaderPrice"]')
                price = parse_price(price_tag.text) if price_tag else None

                location_tag = offer_soup.select_one('div[data-sentry-component="MapLink"] a')
                location = location_tag.text.strip() if location_tag else "No location"

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

                time.sleep(1)  # małe opóźnienie dla serwera

            except Exception as e:
                print(f"Error scraping {url}: {e}")

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
        new_df = pd.DataFrame(new_rows)
        if not new_df.empty:
            # Filtrowanie pustych wierszy, aby nie było warningów Pandas
            new_df = new_df.dropna(how='all')
            existing_df = pd.concat([existing_df, new_df], ignore_index=True)

    existing_links = [r['Link'] for r in results]
    if not existing_df.empty and 'Link' in existing_df.columns:
        existing_df['Aktywne'] = existing_df['Link'].apply(lambda l: l in existing_links)

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
    results_krakow = scrape_offers(BASE_LINK_KRAKOW, 'powiat krakowski')
    update_sheet(results_krakow, 'powiat krakowski')

    results_wielicki = scrape_offers(BASE_LINK_WIELICKI, 'powiat wielicki')
    update_sheet(results_wielicki, 'powiat wielicki')

    print(f"✅ Data saved to '{EXCEL_FILE}' with sheets: powiat krakowski, powiat wielicki")