import requests
import time
import datetime
import os
import pandas as pd
from collections import OrderedDict
from bs4 import BeautifulSoup
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError
from requests.exceptions import ConnectionError, ReadTimeout
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import folium
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
# 🔐 Environment variables for OneDrive auth (.env file required locally)
CLIENT_ID = os.environ['ONEDRIVE_CLIENT_ID']
REFRESH_TOKEN = os.environ['ONEDRIVE_REFRESH_TOKEN']
SCOPES = ['offline_access', 'Files.ReadWrite.All']
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

# -------------------------------
# 📦 App Configuration
BASE_LINK_KRAKOW = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/krakowski?limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'
BASE_LINK_WIELICKI = 'https://www.otodom.pl/pl/wyniki/sprzedaz/dzialka/malopolskie/wielicki?distanceRadius=5&limit=72&priceMax=250000&areaMin=1300&plotType=%5BBUILDING%2CAGRICULTURAL_BUILDING%5D&by=DEFAULT&direction=DESC'
KRAKOW_COORDS = (50.0647, 19.9450)
EXCEL_FOLDER = 'dzialki'
EXCEL_FILENAME = 'wyniki_ofert_z_filtra.xlsx'
EXCEL_FILE = os.path.join(EXCEL_FOLDER, EXCEL_FILENAME)
SHEET_NAMES = ['powiat krakowski', 'powiat wielicki']
HEADERS = [
    'Title', 'Location', 'Price at first find',
    'Date first found', 'Date last updated',
    'Price last updated', 'Distance from Krakow (km)',
    'Active', 'Link'
]

# Allowed counties around Krakow for better geocoding accuracy
ALLOWED_COUNTIES = ['krakowski', 'wielicki', 'wadowicki', 'chrzanowski', 'olkuski', 'myślenicki']

geolocator = Nominatim(user_agent="plot_script")
max_distance_from_Krakow = 50


# -------------------------------
# 🔐 OneDrive token refresh function

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
# ☁️ Upload file to OneDrive root directory

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
# 📊 Excel creation with headers and sheets

def create_excel_with_sheets():
    if not os.path.exists(EXCEL_FOLDER):
        os.makedirs(EXCEL_FOLDER)

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.remove(wb.active)
        for name in SHEET_NAMES:
            ws = wb.create_sheet(name)
            for i, col in enumerate(HEADERS, 1):
                ws.cell(row=1, column=i, value=col)
                ws.column_dimensions[get_column_letter(i)].width = max(15, len(col) + 2)
        wb.save(EXCEL_FILE)
        print(f"📄 Created Excel file: {EXCEL_FILE}")

# -------------------------------
# 🔍 Utility functions

def parse_price(price_str):
    return int(price_str.replace(' ', '').replace('zł', '').replace('PLN', '').replace(',', '').strip())

def extract_relevant_town(location):
    parts = [part.strip() for part in location.split(',')]
    if parts[0].lower().startswith("ul.") and len(parts) > 1:
        return parts[1]
    else:
        return parts[0]

def safe_geocode(loc: str, max_retries: int = 2, timeout: int = 5):
    for attempt in range(max_retries):
        try:
            return geolocator.geocode(loc, exactly_one=False, timeout=timeout)
        except (GeocoderUnavailable, GeocoderServiceError, ConnectionError, ReadTimeout) as e:
            print(f"⏳ Geocoding failed ({attempt + 1}/{max_retries}): {loc} -> {e}")
            time.sleep(1)
    return None

def get_distance_to_krakow(town, county=""):
    allowed_counties = ALLOWED_COUNTIES
    queries = []

    # Step 1: Try with county if it's in allowed list
    if county and county.lower() in allowed_counties:
        queries.append(f"{town}, powiat {county}, małopolskie, Polska")

    # Step 2: Try without county
    queries.append(f"{town}, małopolskie, Polska")

    for query in queries:
        places = safe_geocode(query)
        if places:
            distances = [
                geodesic(KRAKOW_COORDS, (p.latitude, p.longitude)).km
                for p in places if hasattr(p, 'latitude') and hasattr(p, 'longitude')
            ]
            if distances:
                valid_distances = [d for d in distances if d < max_distance_from_Krakow]
                return round(min(valid_distances or distances), 2)

    print(f"⚠️ Nie znaleziono lokalizacji: {town} ({county}) – ustawiam jako -1 km")
    return -1.0  # Use -1.0 to mark it clearly as 'not found'

# -------------------------------
# 🧲 Scrape offers from Otodom
HEADERS_HTTP = {
    "User-Agent": "Mozilla/5.0",
    "Accept-Language": "pl-PL,pl;q=0.9"
}

def scrape_offers(base_link, name):
    results = []
    today = datetime.date.today().strftime('%Y-%m-%d')
    county = name.replace("powiat ", "")
    try:
        res = requests.get(base_link, headers=HEADERS_HTTP, timeout=30)
        soup = BeautifulSoup(res.text, "html.parser")
        links = list(OrderedDict.fromkeys([
            'https://www.otodom.pl' + a['href'] if a['href'].startswith('/') else a['href']
            for a in soup.select('a[data-cy="listing-item-link"]') if a.get('href')
        ]))
        print(f"🔍 {name}: {len(links)} offers found")

        for idx, url in enumerate(links, 1):
            print(f"➡️ Processing {idx}/{len(links)}: {url}")
            try:
                o = requests.get(url, headers=HEADERS_HTTP, timeout=30)
                s = BeautifulSoup(o.text, "html.parser")
                title = s.find('h1').text.strip() if s.find('h1') else 'No title'
                price = parse_price(s.select_one('strong[data-cy="adPageHeaderPrice"]').text)
                location = s.select_one('div[data-sentry-component="MapLink"] a').text.strip()
                town = extract_relevant_town(location)
                distance = get_distance_to_krakow(town, county)
                if distance is None:
                    print(f"⚠️ Nie znaleziono lokalizacji: {town} ({county}) – ustawiam jako -1 km")
                    distance = -1.0

                results.append({
                    'Title': title,
                    'Location': location,
                    'Price at first find': price,
                    'Date first found': today,
                    'Date last updated': today,
                    'Price last updated': price,
                    'Distance from Krakow (km)': distance,
                    'Active': True,
                    'Link': url
                })
                time.sleep(2)
            except Exception as e:
                print(f"❌ Skipping offer due to error: {e}")
    except Exception as e:
        print(f"❌ Scraping error: {e}")
    return results

# -------------------------------
# 🧾 Update Excel sheet with offers

def update_sheet(results, sheet_name):
    today = datetime.date.today().strftime('%Y-%m-%d')
    df_new = pd.DataFrame(results).dropna(how='all')

    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        df_old = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name) if sheet_name in xls.sheet_names else pd.DataFrame(columns=HEADERS)
    else:
        df_old = pd.DataFrame(columns=HEADERS)

    for r in results:
        match = df_old[(df_old['Title'] == r['Title']) & (df_old['Price last updated'] == r['Price last updated'])]
        if not match.empty:
            i = match.index[0]
            df_old.at[i, 'Date last updated'] = today
            df_old.at[i, 'Active'] = True
        else:
            df_old = pd.concat([df_old, pd.DataFrame([r])], ignore_index=True)

    existing_links = [r['Link'] for r in results]
    if 'Link' in df_old.columns:
        df_old['Active'] = df_old['Link'].apply(lambda x: x in existing_links)

    wb = load_workbook(EXCEL_FILE)
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    wb.save(EXCEL_FILE)
    wb.close()

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_old.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        for i, col in enumerate(df_old.columns, 1):
            max_length = max(
                df_old[col].astype(str).map(len).max(),
                len(col)
            )
            adjusted_width = max_length + 2
            ws.column_dimensions[get_column_letter(i)].width = adjusted_width

    print(f"✅ Saved {len(df_old)} offers to sheet '{sheet_name}'")

# -------------------------------
# 🗺️ Generate map from Excel data
def generate_map():
    import csv

    # Initialize the map centered around Krakow
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Add reference marker for Krakow
    folium.Marker(
        location=KRAKOW_COORDS,
        popup=folium.Popup("<b>Kraków</b><br>Punkt odniesienia", max_width=200),
        tooltip="Kraków (punkt odniesienia)",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    df_combined = pd.DataFrame()

    # Read data from all Excel sheets
    for sheet_name in SHEET_NAMES:
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
            df['sheet_name'] = sheet_name
            df_combined = pd.concat([df_combined, df], ignore_index=True)

    marker_count = 0
    location_to_listings = {}
    valid_counties = [name.replace("powiat ", "") for name in SHEET_NAMES]

    # Iterate through each row (offer) in the Excel sheets
    for index, row in df_combined.iterrows():
        if not row.get('Active', False):
            continue

        location_string = row.get('Location', '')
        if not isinstance(location_string, str) or location_string.strip() == '':
            print(f"⚠️ Skipped: missing location (row {index})")
            continue

        town = extract_relevant_town(location_string)
        county = row['sheet_name'].replace("powiat ", "")
        geocode_attempts = []

        # First, try with the county if it's valid
        if county in valid_counties:
            geocode_attempts.append(f"{town}, powiat {county}, małopolskie, Polska")

        # Fallback: try without the county
        geocode_attempts.append(f"{town}, małopolskie, Polska")

        coordinates_found = False
        for query in geocode_attempts:
            geocode_result = safe_geocode(query)
            if not geocode_result:
                continue

            # Calculate distances from Krakow for each geocode result
            places_with_distance = []
            for place in geocode_result:
                lat, lon = place.latitude, place.longitude
                distance = geodesic(KRAKOW_COORDS, (lat, lon)).km
                places_with_distance.append((distance, lat, lon))

            # Sort by distance
            places_with_distance.sort(key=lambda x: x[0])

            # Try placing a marker if within range
            for distance, lat, lon in places_with_distance:
                if distance < max_distance_from_Krakow:
                    coord_key = (round(lat, 5), round(lon, 5))

                    if coord_key not in location_to_listings:
                        location_to_listings[coord_key] = []

                    location_to_listings[coord_key].append({
                        "Title": row['Title'],
                        "Location": location_string,
                        "Price": row['Price last updated'],
                        "Link": row['Link'],
                        "Distance": round(distance, 2)
                    })

                    coordinates_found = True
                    break

            if coordinates_found:
                break

        if not coordinates_found:
            print(f"⚠️ No coordinates found for: {town} ({county}) – offer NOT added to map")
        else:
            print(f"✔️ Geolocation successful: {row['Title']} | {location_string}")

    # Add markers to the map from grouped listings
    added_markers = []
    for (lat, lon), listings in location_to_listings.items():
        popup_html = ""
        for offer in listings:
            popup_html += f"<b>{offer['Title']}</b><br>{offer['Location']}<br>{offer['Price']} PLN<br><a href='{offer['Link']}' target='_blank'>Zobacz ogłoszenie</a><hr>"

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{len(listings)} ogłoszeń",
            icon=folium.Icon(color="green", icon="home", prefix="fa")
        ).add_to(m)

        added_markers.append({
            "Titles": "; ".join([o["Title"] for o in listings]),
            "Location": listings[0]['Location'],
            "Lat": lat,
            "Lon": lon,
            "Distance (km)": listings[0]['Distance'],
            "Count": len(listings)
        })

        print(f"✅ Marker added: {len(listings)} offer(s) at {listings[0]['Location']} → ({lat}, {lon})")
        marker_count += 1

    print(f"📍 Total markers added: {marker_count}")

    # Save list of added markers to CSV file
    if added_markers:
        csv_path = os.path.join(EXCEL_FOLDER, "added_markers.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=added_markers[0].keys())
            writer.writeheader()
            writer.writerows(added_markers)
        print(f"✅ CSV saved: {csv_path}")

    # Save final map to HTML file
    map_file = os.path.join(EXCEL_FOLDER, "map_of_offers.html")
    m.save(map_file)
    print(f"🗺️ Map saved to file: {map_file}")

    return map_file

# -------------------------------
# 🚀 MAIN SCRIPT ENTRY POINT

if __name__ == "__main__":
    create_excel_with_sheets()
    update_sheet(scrape_offers(BASE_LINK_KRAKOW, 'powiat krakowski'), 'powiat krakowski')
    update_sheet(scrape_offers(BASE_LINK_WIELICKI, 'powiat wielicki'), 'powiat wielicki')
    map_path = generate_map()
    
    print(f"📦 Done. Uploading {EXCEL_FILE} and map to OneDrive...")
    token = authenticate()
    upload_to_onedrive(EXCEL_FILE, token)
    upload_to_onedrive(map_path, token)