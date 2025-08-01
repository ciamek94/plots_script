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
# -------------------------------
# 🔧 Uncomment the following 2 lines locally to enable loading variables from the .env file
from dotenv import load_dotenv
load_dotenv()

# -------------------------------
# 🔐 Environment variables required in .env:
# ONEDRIVE_CLIENT_ID=your_client_id
# ONEDRIVE_REFRESH_TOKEN=your_refresh_token
# 
# NOTE: The .env file is not pushed to GitHub — add these values to GitHub Secrets as well
# if you want to run the script automatically via GitHub Actions.
# -------------------------------

# -------------------------------
# OneDrive Configuration
# -------------------------------
CLIENT_ID = os.environ['ONEDRIVE_CLIENT_ID']
REFRESH_TOKEN = os.environ['ONEDRIVE_REFRESH_TOKEN']
SCOPES = ['offline_access', 'Files.ReadWrite.All']
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

# -------------------------------
# App Configuration
# -------------------------------
BASE_LINK_KRAKOW = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/krakowski?limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'
BASE_LINK_WIELICKI = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/wielicki?distanceRadius=5&limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'
KRAKOW_COORDS = (50.0647, 19.9450)
EXCEL_FILE = 'wyniki_ofert_z_filtra.xlsx'
SHEET_NAMES = ['powiat krakowski', 'powiat wielicki']
HEADERS = [
    'Tytuł', 'Lokalizacja', 'Cena pierwszego znalezienia',
    'Data pierwszego znalezienia', 'Data ostatniej aktualizacji',
    'Cena ostatniej aktualizacji', 'Odległość od Krakowa (km)',
    'Aktywne', 'Link'
]
geolocator = Nominatim(user_agent="dzialki_skrypt")
max_distance_from_Krakow = 40 #km

# -------------------------------
# OneDrive Auth (Refresh Token Flow)
# -------------------------------
def authenticate():
    data = {
        'client_id': CLIENT_ID,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'offline_access Files.ReadWrite.All',
    }
    resp = requests.post(TOKEN_URL, data=data)
    if resp.status_code != 200:
        raise Exception(f"❌ Failed to authenticate: {resp.text}")
    return resp.json()

# -------------------------------
# OneDrive Upload
# -------------------------------
def upload_to_onedrive(file_path, token):
    headers = {
        'Authorization': f"Bearer {token['access_token']}",
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    with open(file_path, 'rb') as f:
        file_data = f.read()

    upload_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content'
    r = requests.put(upload_url, headers=headers, data=file_data)
    if r.status_code in (200, 201):
        print(f"✅ File uploaded to OneDrive: {file_path}")
    else:
        print(f"❌ Upload failed: {r.status_code} {r.text}")

# -------------------------------
# Excel Functions
# -------------------------------
def create_excel_with_sheets():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.remove(wb.active)
        for name in SHEET_NAMES:
            ws = wb.create_sheet(name)
            for i, col in enumerate(HEADERS, 1):
                ws.cell(row=1, column=i, value=col)
                ws.column_dimensions[get_column_letter(i)].width = max(15, len(col) + 2)
        wb.save(EXCEL_FILE)
        print(f"📄 Created Excel: {EXCEL_FILE}")

# -------------------------------
# Scraping & Utils
# -------------------------------
def parse_price(price_str):
    return int(price_str.replace(' ', '').replace('zł', '').replace('PLN', '').replace(',', '').strip())

def extract_relevant_town(location):
    parts = [part.strip() for part in location.split(',')]
    if parts[0].lower().startswith("ul.") and len(parts) > 1:
        return parts[1]
    else:
        return parts[0]

def safe_geocode(loc):
    for _ in range(3):
        try:
            return geolocator.geocode(loc, exactly_one=False)
        except GeocoderTimedOut:
            time.sleep(1)
    return None

def get_distance_to_krakow(location_string):
    town = extract_relevant_town(location_string)
    query = f"{town}, małopolskie, Polska"
    places = safe_geocode(query)
    if not places:
        return None
    distances = [
        geodesic(KRAKOW_COORDS, (p.latitude, p.longitude)).km
        for p in places if hasattr(p, 'latitude') and hasattr(p, 'longitude')
    ]
    if not distances:
        return None
    return round(min([d for d in distances if d < max_distance_from_Krakow] or [distances[0]]), 2)

HEADERS_HTTP = {
    "User-Agent": "Mozilla/5.0",
    "Accept-Language": "pl-PL,pl;q=0.9"
}

def scrape_offers(base_link, name):
    results = []
    today = datetime.date.today().strftime('%Y-%m-%d')
    try:
        res = requests.get(base_link, headers=HEADERS_HTTP, timeout=30)
        soup = BeautifulSoup(res.text, "html.parser")
        links = list(OrderedDict.fromkeys([
            'https://www.otodom.pl' + a['href'] if a['href'].startswith('/') else a['href']
            for a in soup.select('a[data-cy="listing-item-link"]') if a.get('href')
        ]))
        print(f"🔍 {name}: {len(links)} offers")

        for idx, url in enumerate(links, 1):
            print(f"➡️ {idx}/{len(links)}: {url}")
            try:
                o = requests.get(url, headers=HEADERS_HTTP, timeout=30)
                s = BeautifulSoup(o.text, "html.parser")
                title = s.find('h1').text.strip() if s.find('h1') else 'No title'
                price = parse_price(s.select_one('strong[data-cy="adPageHeaderPrice"]').text)
                location = s.select_one('div[data-sentry-component="MapLink"] a').text.strip()
                town = extract_relevant_town(location)
                distance = get_distance_to_krakow(town) or 0.0

                results.append({
                    'Tytuł': title,
                    'Lokalizacja': location,
                    'Cena pierwszego znalezienia': price,
                    'Data pierwszego znalezienia': today,
                    'Data ostatniej aktualizacji': today,
                    'Cena ostatniej aktualizacji': price,
                    'Odległość od Krakowa (km)': distance,
                    'Aktywne': True,
                    'Link': url
                })
                time.sleep(2)
            except Exception as e:
                print(f"❌ Skipping offer: {e}")
    except Exception as e:
        print(f"❌ Scrape error: {e}")
    return results

def update_sheet(results, sheet_name):
    today = datetime.date.today().strftime('%Y-%m-%d')
    df_new = pd.DataFrame(results).dropna(how='all')

    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        df_old = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame(columns=HEADERS)
    else:
        df_old = pd.DataFrame(columns=HEADERS)

    for r in results:
        match = df_old[(df_old['Tytuł'] == r['Tytuł']) & (df_old['Cena ostatniej aktualizacji'] == r['Cena ostatniej aktualizacji'])]
        if not match.empty:
            i = match.index[0]
            df_old.at[i, 'Data ostatniej aktualizacji'] = today
            df_old.at[i, 'Aktywne'] = True
        else:
            df_old = pd.concat([df_old, pd.DataFrame([r])], ignore_index=True)

    existing_links = [r['Link'] for r in results]
    if 'Link' in df_old.columns:
        df_old['Aktywne'] = df_old['Link'].apply(lambda x: x in existing_links)

    wb = load_workbook(EXCEL_FILE)
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    wb.save(EXCEL_FILE)
    wb.close()

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_old.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        # Dopasowanie szerokości kolumn do zawartości
        for i, col in enumerate(df_old.columns, 1):
            max_length = max(
                df_old[col].astype(str).map(len).max(),
                len(col)
            )
            adjusted_width = max_length + 2  # margines
            ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    print(f"✅ Saved {len(df_old)} offers to '{sheet_name}'")

# -------------------------------
# MAIN
# -------------------------------
if __name__ == "__main__":
    create_excel_with_sheets()
    update_sheet(scrape_offers(BASE_LINK_KRAKOW, 'powiat krakowski'), 'powiat krakowski')
    update_sheet(scrape_offers(BASE_LINK_WIELICKI, 'powiat wielicki'), 'powiat wielicki')
    print(f"📦 Done. Uploading {EXCEL_FILE} to OneDrive...")
    token = authenticate()
    upload_to_onedrive(EXCEL_FILE, token)