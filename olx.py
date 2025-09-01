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

# =========================================
# Constants & output paths
# =========================================
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

def autosize_columns(filename: str) -> None:
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

def get_distance_from_krakow(location: str):
    geolocator = Nominatim(user_agent="dzialki_app")
    def try_geocode(query):
        try:
            return geolocator.geocode(query, timeout=8)
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

def check_if_active(url: str, headers: dict) -> bool:
    try:
        r = requests.head(url, headers=headers, allow_redirects=True, timeout=5)
        if r.status_code == 200:
            return True
        r = requests.get(url, headers=headers, timeout=8)
        return r.status_code == 200
    except Exception:
        return True

def get_with_retry(url: str, headers: dict, retries: int = 5):
    for _ in range(retries):
        try:
            r = requests.get(url, headers=headers, timeout=12)
            if r.status_code == 200:
                return r
        except Exception:
            pass
        time.sleep(random.uniform(2.5, 5.0))
    return None

# =========================================
# Map generation
# =========================================
def generate_map(df: pd.DataFrame) -> None:
    """Tworzy mapę z jednym markerem na punkt dla wielu ogłoszeń w tym samym miejscu."""
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Marker Krakowa
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

        # Tworzymy HTML popup z listą ogłoszeń
        popup_html = ""
        for _, row in group.iterrows():
            popup_html += f"""
            <b>{row['Title']}</b><br>
            {row['Location']}<br>
            {row.get('Price last updated', '')} PLN<br>
            <a href='{row['Link']}' target='_blank'>Zobacz ogłoszenie</a>
            <hr>
            """

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{len(group)} ogłoszeń w tym miejscu",
            icon=folium.Icon(color="blue", icon="home", prefix="fa")
        ).add_to(m)
        plotted += 1

    m.save(MAP_FILE)
    print(f"🗺️ Map saved to: {MAP_FILE} (markers plotted: {plotted})")

# =========================================
# Main scraping & export
# =========================================
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
            location, date_added = parse_location_date(loc_date)

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
    autosize_columns(EXCEL_FILE)
    print(f"📂 Listings saved to Excel file: {EXCEL_FILE}")
    print(f"🧮 Rows in Excel: {len(df_updated)} (active: {int(df_updated['Active'].sum())})")

    generate_map(df_updated)

if __name__ == "__main__":
    main()
