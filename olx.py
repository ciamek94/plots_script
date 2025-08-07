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

# Constants
KRAKOW_COORDS = (50.0647, 19.9450)
MAX_DISTANCE_KM = 50
EXCEL_FOLDER = 'dzialki'
EXCEL_FILENAME = 'olx_dzialki.xlsx'
EXCEL_FILE = os.path.join(EXCEL_FOLDER, EXCEL_FILENAME)
MAP_FILE = os.path.join(EXCEL_FOLDER, 'olx_map_listings.html')

# Ensure the output folder exists
os.makedirs(EXCEL_FOLDER, exist_ok=True)

def clean_price(price_str):
    if not price_str:
        return ""
    return price_str.replace("zł", "").replace("do negocjacji", "").replace(" ", "").strip()

def parse_location_date(loc_date_str):
    parts = loc_date_str.split(" - ")
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return loc_date_str.strip(), ""

def autosize_columns(filename):
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

def get_distance_from_krakow(location):
    geolocator = Nominatim(user_agent="dzialki_app")

    def try_geocode(query):
        try:
            return geolocator.geocode(query)
        except:
            return None

    queries = [f"{location}, Poland", f"{location}, Malopolskie, Poland"]
    for query in queries:
        loc = try_geocode(query)
        if loc:
            distance = round(geodesic(KRAKOW_COORDS, (loc.latitude, loc.longitude)).km, 2)
            if distance <= MAX_DISTANCE_KM:
                return distance, loc.latitude, loc.longitude
    return None, None, None

def check_if_active(url, headers):
    try:
        r = requests.head(url, headers=headers, allow_redirects=True, timeout=5)
        return r.status_code == 200
    except:
        return False

def get_with_retry(url, headers, retries=5):
    for i in range(retries):
        try:
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code == 200:
                return r
        except:
            pass
        time.sleep(random.uniform(3, 6))
    return None

def generate_map(df):
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)
    folium.Marker(
        location=KRAKOW_COORDS,
        popup=folium.Popup("<b>Kraków</b><br>Reference point", max_width=200),
        tooltip="Kraków",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    for _, row in df.iterrows():
        if not row.get("Active", False):
            continue
        location = row.get("Location")
        distance, lat, lon = get_distance_from_krakow(location)
        if lat is None or lon is None:
            continue

        popup_html = f"""
        <b>{row['Title']}</b><br>
        {location}<br>
        {row['Price last updated']} PLN<br>
        <a href='{row['Link']}' target='_blank'>Zobacz ogłoszenie</a>
        """

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=row['Title'],
            icon=folium.Icon(color="blue", icon="home", prefix="fa")
        ).add_to(m)

    m.save(MAP_FILE)
    print(f"🗺️ Map saved to: {MAP_FILE}")

def main():
    base_url = (
        "https://www.olx.pl/nieruchomosci/dzialki/sprzedaz/krakow/"
        "q-dzialka-budowlana/?search[dist]=30"
        "&search[filter_enum_type][0]=dzialki-budowlane"
        "&search[filter_float_m:from]=1150"
        "&search[filter_float_price:to]=250000"
        "&page={page}"
    )
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "pl-PL,pl;q=0.9"
    }

    try:
        df_existing = pd.read_excel(EXCEL_FILE)
        print(f"📄 Existing Excel file found: {EXCEL_FILE}")
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=[
            "Title", "Location", "Price at first find", "Date first found",
            "Date last updated", "Price last updated", "Distance from Krakow (km)",
            "Active", "Link"
        ])
        print("📄 No existing Excel file found. A new one will be created.")

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
            print(f"❌ Failed to fetch page {page}")
            empty_pages += 1
            page += 1
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        cards = soup.find_all("div", {"data-cy": "l-card"})
        if not cards:
            print(f"📭 No listings found on page {page}")
            empty_pages += 1
            page += 1
            continue

        print(f"✅ {len(cards)} listings found on page {page}")
        empty_pages = 0

        for card in cards:
            title_elem = card.find("h4", class_="css-1g61gc2")
            price_elem = card.find("p", {"data-testid": "ad-price"})
            loc_date_elem = card.find("p", {"data-testid": "location-date"})
            link_elem = card.find("a", class_="css-1tqlkj0")

            title = title_elem.text.strip() if title_elem else ""
            price = clean_price(price_elem.text.strip()) if price_elem else ""
            loc_date = loc_date_elem.text.strip() if loc_date_elem else ""
            location, date_added = parse_location_date(loc_date)
            link = link_elem["href"] if link_elem else ""
            if not link.startswith("http"):
                link = "https://www.olx.pl" + link

            distance, lat, lon = get_distance_from_krakow(location)
            if distance is None:
                continue

            all_listings.append({
                "Title": title,
                "Location": location,
                "Price at first find": price,
                "Date first found": date_added,
                "Date last updated": date_added,
                "Price last updated": price,
                "Distance from Krakow (km)": distance,
                "Active": True,
                "Link": link
            })

        page += 1
        time.sleep(random.uniform(2, 4))

    if not all_listings:
        print("🚫 No listings found.")
        return

    print(f"📊 Total listings collected: {len(all_listings)}")

    df_new = pd.DataFrame(all_listings)
    df_existing.set_index("Link", inplace=True)
    df_new.set_index("Link", inplace=True)

    for link in df_new.index:
        if link in df_existing.index:
            df_new.at[link, "Date first found"] = df_existing.at[link, "Date first found"]
            df_new.at[link, "Price at first find"] = df_existing.at[link, "Price at first find"]

    df_updated = df_new.copy()

    for link in df_updated.index:
        df_updated.at[link, "Active"] = check_if_active(link, headers)

    df_updated.reset_index(inplace=True)

    columns_final = [
        "Title",
        "Location",
        "Price at first find",
        "Date first found",
        "Date last updated",
        "Price last updated",
        "Distance from Krakow (km)",
        "Active",
        "Link"
    ]

    df_updated = df_updated[columns_final]

    df_updated.to_excel(EXCEL_FILE, index=False)
    autosize_columns(EXCEL_FILE)
    print(f"💾 Listings saved to Excel file: {EXCEL_FILE}")
    print(f"🆕 New rows added: {len(df_updated)}")

    generate_map(df_updated)

if __name__ == "__main__":
    main()
