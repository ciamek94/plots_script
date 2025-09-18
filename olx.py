import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import os
import time
import random
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import folium
from folium.plugins import MarkerCluster
from collections import defaultdict
from datetime import datetime, date, timedelta
import re

# =========================================
# Constants & output paths
# =========================================

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
# 🔐 OneDrive authentication variables (stored in .env)
CLIENT_ID = os.environ['ONEDRIVE_CLIENT_ID']
REFRESH_TOKEN = os.environ['ONEDRIVE_REFRESH_TOKEN']
SCOPES = ['offline_access', 'Files.ReadWrite.All']
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

KRAKOW_COORDS = (50.0647, 19.9450)
MAX_DISTANCE_KM = 50
EXCEL_FOLDER = 'dzialki'
EXCEL_FILENAME = 'olx_dzialki.xlsx'
EXCEL_FILE = os.path.join(EXCEL_FOLDER, EXCEL_FILENAME)
MAP_FILE = os.path.join(EXCEL_FOLDER, 'olx_map_listings.html')

os.makedirs(EXCEL_FOLDER, exist_ok=True)

# =========================================
# Helpers
# =========================================

# -------------------------------
# 🔐 OneDrive token refresh
def authenticate():
    """Authenticate with OneDrive API using refresh token"""
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
# ☁️ Upload file to OneDrive
def upload_to_onedrive(file_path, token):
    """Upload a file to OneDrive root directory"""
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
# ☁️ Download file from OneDrive
def download_from_onedrive(file_path, token):
    """Download a file from OneDrive and save it locally"""
    headers = {
        'Authorization': f"Bearer {token['access_token']}",
    }
    download_url = f'https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content'
    r = requests.get(download_url, headers=headers)
    if r.status_code == 200:
        with open(file_path, 'wb') as f:
            f.write(r.content)
        print(f"✅ File downloaded from OneDrive: {file_path}")
    else:
        print(f"⚠️ Failed to download file from OneDrive: {r.status_code} {r.text}")


# -------------------------------
def clean_price(price_str: str) -> str:
    if not price_str:
        return ""
    return (price_str.replace("zł", "")
                     .replace("do negocjacji", "")
                     .replace(" ", "")
                     .strip())

def parse_location_date(loc_date_str: str):
    parts = loc_date_str.split(" - ")
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return loc_date_str.strip(), ""

# -------------------------------
def autosize_columns(filename: str) -> None:
    """Autosize Excel columns based on content length"""
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2
    wb.save(filename)

# -------------------------------
def load_towns(file_path="town_list.txt"):
    """Load towns and their coordinates from file"""
    towns = defaultdict(list)
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split("|")
                if len(parts) == 3:
                    town, lat, lon = parts
                    towns[town.lower()].append((float(lat), float(lon)))
    print(f"ℹ️ Loaded {sum(len(v) for v in towns.values())} town coordinates from {file_path}")
    return towns

TOWN_COORDS = load_towns("town_list.txt")

# -------------------------------
def get_distance_from_krakow(location: str):
    """Return a list of (distance, lat, lon) for all coordinates of a town within MAX_DISTANCE_KM"""
    town = location.split("(")[0].strip().lower()
    results = []

    # Use coordinates from town_list.txt
    if town in TOWN_COORDS:
        for lat, lon in TOWN_COORDS[town]:
            distance = round(geodesic(KRAKOW_COORDS, (lat, lon)).km, 2)
            if distance <= MAX_DISTANCE_KM:
                results.append((distance, lat, lon))

    # Fallback to geopy if no town coordinates found
    if not results:
        geolocator = Nominatim(user_agent="dzialki_app")
        try:
            geo = geolocator.geocode(f"{town}, Malopolskie, Poland", timeout=8)
            if geo:
                distance = round(geodesic(KRAKOW_COORDS, (geo.latitude, geo.longitude)).km, 2)
                if distance <= MAX_DISTANCE_KM:
                    results.append((distance, geo.latitude, geo.longitude))
        except Exception as e:
            print(f"⚠️ Geopy failed for {town}: {e}")

    return results  # always return a list (can be empty)

# -------------------------------
def check_if_active(url: str, headers: dict) -> bool:
    """Check if a listing URL is still active"""
    try:
        r = requests.head(url, headers=headers, allow_redirects=True, timeout=5)
        if r.status_code == 200:
            return True
        r = requests.get(url, headers=headers, timeout=8)
        return r.status_code == 200
    except Exception:
        return True

# -------------------------------
def get_with_retry(url: str, headers: dict, retries: int = 5):
    """Retry GET request up to N times"""
    for _ in range(retries):
        try:
            r = requests.get(url, headers=headers, timeout=12)
            if r.status_code == 200:
                return r
        except Exception:
            pass
        time.sleep(random.uniform(2.5, 5.0))
    return None

# -------------------------------
def generate_map(df: pd.DataFrame) -> None:
    """Generate HTML map with markers for listings"""
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Marker for Krakow
    folium.Marker(
        location=KRAKOW_COORDS,
        popup=folium.Popup("<b>Kraków</b><br>Reference point", max_width=200),
        tooltip="Kraków",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    plotted = 0
    grouped = df[df["Active"]].groupby(["Latitude", "Longitude"])

    for (lat, lon), group in grouped:
        if pd.isna(lat) or pd.isna(lon):
            continue

        popup_html = ""
        for _, row in group.iterrows():
            popup_html += f"""
            <b>{row['Title']}</b><br>
            {row['Location']}<br>
            {row.get('Price last updated', '')} PLN<br>
            <a href='{row['Link']}' target='_blank'>View listing</a>
            <hr>
            """

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{len(group)} listings at this location",
            icon=folium.Icon(color="blue", icon="home", prefix="fa")
        ).add_to(m)
        plotted += 1

    m.save(MAP_FILE)
    print(f"🗺️ Map saved to: {MAP_FILE} (markers plotted: {plotted})")

# -------------------------------
# 📅 Convert OLX date strings to dd.mm.yyyy format
def parse_olx_date(date_str: str):
    """Convert OLX date string to yyyy-mm-dd format"""
    if not date_str or not isinstance(date_str, str):
        return None

    date_str = date_str.lower().strip()
    today = date.today()
    
    # Handle "today" and "yesterday"
    if "dzisiaj" in date_str or "odświeżono dzisiaj" in date_str:
        return today.strftime("%Y-%m-%d")
    if "wczoraj" in date_str:
        return (today - timedelta(days=1)).strftime("%Y-%m-%d")

    # Handle dd.mm.yyyy format (e.g. 17.09.2025)
    try:
        return datetime.strptime(date_str, "%d.%m.%Y").strftime("%Y-%m-%d")
    except ValueError:
        pass

    # Handle day month_name year (e.g. '17 września 2025')
    months = {
        "stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5,
        "czerwca": 6, "lipca": 7, "sierpnia": 8, "września": 9,
        "października": 10, "listopada": 11, "grudnia": 12
    }

    match = re.search(r'(\d{1,2}) (\w+) (\d{4})', date_str)
    if match:
        day, month_str, year = match.groups()
        month = months.get(month_str, 0)
        if month:
            return f"{year}-{month:02d}-{int(day):02d}"  # yyyy-mm-dd format
    
    # Fallback: return as-is
    return date_str


# =========================================
# Main scraping & export
# =========================================
def main():
    base_url = (
    "https://www.olx.pl/nieruchomosci/dzialki/sprzedaz/krakow/"
    "?search%5Bdist%5D=30"
    "&search%5Bfilter_float_price:to%5D=250000"
    "&search%5Bfilter_enum_type%5D%5B0%5D=dzialki-budowlane"
    "&search%5Bfilter_enum_type%5D%5B1%5D=dzialki-rolno-budowlane"
    "&search%5Bfilter_float_m:from%5D=1150"
    "&page={page}"
    )
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "pl-PL,pl;q=0.9"
    }

    # If OneDrive credentials are set, download the latest Excel
    if CLIENT_ID and REFRESH_TOKEN:
        print("☁️ OneDrive credentials found. Downloading latest Excel copy...")
        token = authenticate()
        download_from_onedrive(EXCEL_FILE, token)
    else:
        print("⚠️ OneDrive credentials not found. Using local Excel copy.")

    try:
        df_existing = pd.read_excel(EXCEL_FILE)
        if not df_existing.empty and "Link" in df_existing.columns:
            df_existing = df_existing.drop_duplicates(subset="Link", keep="first")
        print(f"📄 Existing Excel file found: {EXCEL_FILE}")
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=[
            "Title", "Location", "Price at first find", "Date first found",
            "Date last updated", "Price last updated", "Distance from Krakow (km)",
            "Active", "Link", "Latitude", "Longitude"
        ])
        print("📄 No existing Excel file found. A new one will be created.")

    # -------------------------------
    # 📅 Normalize existing dates to dd.mm.yyyy format
    if not df_existing.empty:
        df_existing["Date first found"] = df_existing["Date first found"].apply(parse_olx_date)
        df_existing["Date last updated"] = df_existing["Date last updated"].apply(parse_olx_date)

    all_listings = []
    page = 1
    empty_pages = 0
    max_empty_pages = 3
    print("🔍 Starting search for listings...")

    while empty_pages < max_empty_pages:
        url = base_url.format(page=page)
        print(f"🌐 Fetching page {page}: {url}")
        response = get_with_retry(url, headers)
        if response is None:
            empty_pages += 1
            page += 1
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        cards = soup.find_all("div", {"data-cy": "l-card"})
        if not cards:
            empty_pages += 1
            page += 1
            continue

        print(f"✅ {len(cards)} listings found on page {page}")
        empty_pages = 0

        for card in cards:
            title_elem = card.select_one('div[data-cy="ad-card-title"] h4')
            title = title_elem.get_text(strip=True) if title_elem else ""

            link_elem = card.find("a", class_="css-1tqlkj0")
            link = link_elem["href"] if link_elem and link_elem.has_attr("href") else ""
            if link and not link.startswith("http"):
                link = "https://www.olx.pl" + link
            if not link:
                continue

            price_elem = card.find("p", {"data-testid": "ad-price"})
            price = clean_price(price_elem.get_text(strip=True)) if price_elem else ""

            loc_date_elem = card.find("p", {"data-testid": "location-date"})
            loc_date = loc_date_elem.get_text(strip=True) if loc_date_elem else ""
            location, date_added_raw = parse_location_date(loc_date)
            date_added = parse_olx_date(date_added_raw)

            # Get all coordinates for this location
            coords_list = get_distance_from_krakow(location)
            if not coords_list:
                continue

            # Create a separate row for each coordinate
            for distance, lat, lon in coords_list:
                all_listings.append({
                    "Title": title,
                    "Location": location,
                    "Price at first find": price,
                    "Date first found": date_added,
                    "Date last updated": date_added,
                    "Price last updated": price,
                    "Distance from Krakow (km)": distance,
                    "Active": True,
                    "Link": link,
                    "Latitude": lat,
                    "Longitude": lon
                })

        page += 1
        time.sleep(random.uniform(2, 4))

    if not all_listings:
        print("🚫 No listings found.")
        return

    df_new = pd.DataFrame(all_listings)
    if not df_existing.empty:
        df_merged = df_new.merge(
            df_existing[["Link", "Date first found", "Price at first find"]],
            on="Link",
            how="left",
            suffixes=("", "_old")
        )
        df_merged["Date first found"] = df_merged["Date first found_old"].fillna(df_merged["Date first found"])
        df_merged["Price at first find"] = df_merged["Price at first find_old"].fillna(df_merged["Price at first find"])
        df_merged = df_merged.drop(columns=["Date first found_old", "Price at first find_old"])
    else:
        df_merged = df_new.copy()

    df_merged["Active"] = df_merged["Link"].apply(lambda u: check_if_active(u, headers))

    columns_final = [
        "Title", "Location", "Price at first find", "Date first found",
        "Date last updated", "Price last updated", "Distance from Krakow (km)",
        "Active", "Link", "Latitude", "Longitude"
    ]

    df_updated = df_merged[columns_final].copy()
    df_updated.to_excel(EXCEL_FILE, index=False)
    autosize_columns(EXCEL_FILE)  # Fixed: function is now defined before main
    print(f"📂 Listings saved to Excel file: {EXCEL_FILE}")
    print(f"🧮 Rows in Excel: {len(df_updated)} (active: {int(df_updated['Active'].sum())})")

    generate_map(df_updated)

    if CLIENT_ID and REFRESH_TOKEN:
        print(f"📦 Done. Uploading {EXCEL_FILE} and map to OneDrive...")
        token = authenticate()
        upload_to_onedrive(EXCEL_FILE, token)
        upload_to_onedrive(MAP_FILE, token)
    else:
        print("⚠️ OneDrive credentials not found. Skipping upload.")

if __name__ == "__main__":
    main()
