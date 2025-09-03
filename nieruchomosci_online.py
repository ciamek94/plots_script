import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os
import openpyxl
from openpyxl.utils import get_column_letter
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import folium
from collections import defaultdict

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

# Constants
KRAKOW_COORDS = (50.0647, 19.9450)
MAX_DISTANCE_KM = 50
MAX_PAGES = 20
EXCEL_FOLDER = 'dzialki'
EXCEL_FILENAME = 'nieruchomosci_online_dzialki.xlsx'
EXCEL_FILE = os.path.join(EXCEL_FOLDER, EXCEL_FILENAME)
MAP_FILE = os.path.join(EXCEL_FOLDER, 'nieruchomosci_online_map_listings.html')
HEADERS = {"User-Agent": "Mozilla/5.0"}
BASE_URL = "https://www.nieruchomosci-online.pl/szukaj.html?3,dzialka,sprzedaz,,Krak%C3%B3w:5600,,,25,-250000,1150,,,,,,,,,,,,,1"

# -------------------------------
# 🔐 OneDrive authentication variables (stored in .env)
CLIENT_ID = os.environ['ONEDRIVE_CLIENT_ID']
REFRESH_TOKEN = os.environ['ONEDRIVE_REFRESH_TOKEN']
SCOPES = ['offline_access', 'Files.ReadWrite.All']
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

results = []
geolocator = Nominatim(user_agent="dzialki_locator")

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


# Get coordinates and distance from Kraków using only city/town part
def get_distance_from_krakow(location):
    try:
        town = location.split("(")[0].strip()
        geo = geolocator.geocode(f"{town}, Poland")
        if geo:
            coords = (geo.latitude, geo.longitude)
            distance = round(geodesic(KRAKOW_COORDS, coords).km, 2)
            if distance <= MAX_DISTANCE_KM:
                return distance, geo.latitude, geo.longitude
    except:
        pass
    return None, None, None

def main():
    # Ensure output folder exists
    os.makedirs(EXCEL_FOLDER, exist_ok=True)

    # Loop through listing pages
    for page in range(1, MAX_PAGES + 1):
        url = BASE_URL if page == 1 else f"{BASE_URL}&p={page}"
        print(f"Fetching page {page}: {url}")
        response = requests.get(url, headers=HEADERS)
        if response.status_code != 200:
            print(f"Failed to load page {page}")
            break

        soup = BeautifulSoup(response.text, "html.parser")
        listings = soup.find_all("div", class_="tile-inner") + soup.find_all("div", class_="tertiary")

        for listing in listings:
            try:
                title_tag = listing.find("h2", class_="name")
                title = title_tag.text.strip() if title_tag else "No title"
                link = title_tag.find("a")["href"] if title_tag and title_tag.find("a") else "No link"

                price_tag = listing.find("p", class_="title-a")
                price_span = price_tag.find("span") if price_tag else None
                price = price_span.text.strip() if price_span else "No price"

                location_p = listing.find("p", class_="province")
                location = " ".join([el.text.strip() for el in location_p.find_all(['a', 'span'])]) if location_p else "No location"

                if link != "No link":
                    distance, lat, lon = get_distance_from_krakow(location)
                else:
                    distance, lat, lon = None, None, None

                results.append({
                    "Title": title,
                    "Location": location,
                    "Price": price,
                    "Link": link,
                    "Distance from Krakow (km)": distance,
                    "Latitude": lat,
                    "Longitude": lon
                })

            except Exception as e:
                print(f"❗ Error parsing listing: {e}")
                continue

        time.sleep(1)

    # Create DataFrame from scraped results
    df = pd.DataFrame(results).drop_duplicates()

    # Filter out archived or invalid offers
    df = df[df["Link"] != "No link"]
    df = df[pd.notna(df["Latitude"]) & pd.notna(df["Longitude"])]

    # Save to Excel
    df.to_excel(EXCEL_FILE, index=False)

    # Auto-fit Excel column widths
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    for column_cells in ws.columns:
        max_length = 0
        column = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2
    wb.save(EXCEL_FILE)

    # Create interactive map centered on Kraków
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Add Kraków marker
    folium.Marker(
        location=KRAKOW_COORDS,
        popup="Kraków - Reference Point",
        tooltip="Kraków",
        icon=folium.Icon(color="purple")
    ).add_to(m)

    # Group listings by exact coordinates
    marker_groups = defaultdict(list)
    for _, row in df.iterrows():
        coord = (row["Latitude"], row["Longitude"])
        marker_groups[coord].append(row)

    # Add grouped markers to the map
    for coord, listings in marker_groups.items():
        if len(listings) == 1:
            l = listings[0]
            popup_html = f"""
            <b>{l['Title']}</b><br>
            {l['Location']}<br>
            {l['Price']}<br>
            <a href="{l['Link']}" target="_blank">View Listing</a>
            """
            tooltip = l["Title"]
        else:
            popup_html = f"<b>{len(listings)} listings</b><br><ul>"
            for l in listings:
                popup_html += f"<li><a href='{l['Link']}' target='_blank'>{l['Title']}</a> – {l['Price']}</li>"
            popup_html += "</ul>"
            tooltip = f"{len(listings)} listings at same location"

        folium.Marker(
            location=coord,
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=tooltip,
            icon=folium.Icon(color="orange", icon="home")
        ).add_to(m)

    # Save map to file
    m.save(MAP_FILE)

    print(f"\n✅ Data saved to: {EXCEL_FILE}")
    print(f"🗺️ Map saved to: {MAP_FILE}")

    print(f"📦 Done. Uploading {EXCEL_FILE} and map to OneDrive...")
    token = authenticate()
    upload_to_onedrive(EXCEL_FILE, token)
    upload_to_onedrive(MAP_FILE, token)


if __name__ == "__main__":
    main()
