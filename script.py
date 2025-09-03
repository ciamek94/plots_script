import os
import pandas as pd
import folium
import openpyxl
from openpyxl.utils import get_column_letter
from collections import defaultdict
import requests

# Importy Twoich skryptów (zakładam, że mają main() albo analogiczne funkcje uruchamiające)
from otodom import main as main_script1
from olx import main as main_script2
from nieruchomosci_online import main as main_script3  # Twój nowy skrypt nieruchomosci-online

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

# Ścieżki do plików Excel generowanych przez poszczególne skrypty
EXCEL_FILE_1 = os.path.join('dzialki', 'otodom_dzialki.xlsx')
EXCEL_FILE_2 = os.path.join('dzialki', 'olx_dzialki.xlsx')
EXCEL_FILE_3 = os.path.join('dzialki', 'nieruchomosci_online_dzialki.xlsx')  # musi być tak zapisany w main_script3

# Ścieżki do scalonych plików
EXCEL_MERGED = os.path.join('dzialki_merged', 'dzialki_merged.xlsx')
MAP_MERGED = os.path.join('dzialki_merged', 'map_merged.html')

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

def merge_excels(files_list, output_file):
    dfs = []
    for idx, file in enumerate(files_list):
        if not os.path.exists(file):
            print(f"❗ File not found: {file}")
            continue

        # Dla pliku otodom – dwie zakładki, dla pozostałych pojedyncze arkusze
        if idx == 0:
            xls = pd.ExcelFile(file)
            try:
                df_krakow = pd.read_excel(xls, sheet_name='powiat krakowski')
                df_wielicki = pd.read_excel(xls, sheet_name='powiat wielicki')
                df = pd.concat([df_krakow, df_wielicki], ignore_index=True)
            except Exception as e:
                print(f"Error reading sheets in {file}: {e}")
                df = pd.read_excel(file)  # fallback na cały plik
        else:
            df = pd.read_excel(file)

        # Dodajemy kolumnę Source wg indexu
        source_name = ['otodom', 'olx', 'nieruchomosci-online'][idx]
        df['Source'] = source_name

        dfs.append(df)

    if not dfs:
        print("❗ Brak plików do scalania!")
        return None

    df_combined = pd.concat(dfs, ignore_index=True)

    # Normalizacja linków i usuwanie duplikatów po linku (jeśli istnieje)
    if 'Link' in df_combined.columns:
        df_combined['Link'] = df_combined['Link'].astype(str).str.strip().str.lower()
        df_unique = df_combined.drop_duplicates(subset=['Link'], keep='first').reset_index(drop=True)
    else:
        df_unique = df_combined.drop_duplicates().reset_index(drop=True)

    # Możesz dostosować kolumny do zapisu, tutaj prosto zapisujemy wszystko
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_unique.to_excel(output_file, index=False)

    # Auto szerokość kolumn w Excelu
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active

    for col_cells in ws.columns:
        max_length = 0
        col = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col].width = max_length + 2

    wb.save(output_file)
    print(f"💾 Merged Excel saved: {output_file}")
    return df_unique

def generate_merged_map(df, map_path):
    KRAKOW_COORDS = (50.0647, 19.9450)
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    folium.Marker(
        location=KRAKOW_COORDS,
        popup="<b>Kraków</b><br>Reference point",
        tooltip="Kraków",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    location_to_listings = defaultdict(list)

    for _, row in df.iterrows():
        if 'Active' in df.columns and not row.get("Active", True):
            continue
        lat, lon = row.get("Latitude"), row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            continue
        coord_key = (round(lat, 5), round(lon, 5))
        location_to_listings[coord_key].append(row)

    for (lat, lon), listings in location_to_listings.items():
        if len(listings) == 1:
            l = listings[0]
            popup_html = f"""
                <b>{l.get('Title', 'Brak tytułu')}</b><br>
                {l.get('Location', '')}<br>
                {l.get('Price last updated', l.get('Price at first find', l.get('Price', '?')))}<br>
                <a href="{l.get('Link', '#')}" target="_blank">View listing</a>
            """
            tooltip = l.get("Title", "Listing")
            source = l.get("Source", "").lower()
        else:
            popup_html = f"<b>{len(listings)} ogłoszenia</b><br><ul>"
            for l in listings:
                price = l.get('Price last updated', l.get('Price at first find', l.get('Price', '?')))
                popup_html += f"<li><a href='{l.get('Link', '#')}' target='_blank'>{l.get('Title', 'Brak tytułu')}</a> – {price}</li>"
            popup_html += "</ul>"
            tooltip = f"{len(listings)} ogłoszeń"
            source = listings[0].get("Source", "").lower()

        # Kolor markera wg źródła
        if source == 'otodom':
            color = "green"
        elif source == 'olx':
            color = "blue"
        elif source == 'nieruchomosci-online':
            color = "orange"
        else:
            color = "gray"

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=tooltip,
            icon=folium.Icon(color=color, icon="home", prefix="fa")
        ).add_to(m)

    m.save(map_path)
    print(f"🗺️ Merged map saved: {map_path}")

def main():
    print("🚀 Uruchamiam skrypt 1 (Otodom)...")
    main_script1()

    print("🚀 Uruchamiam skrypt 2 (OLX)...")
    main_script2()

    print("🚀 Uruchamiam skrypt 3 (Nieruchomosci-online)...")
    main_script3()

    print("🔄 Scalanie wszystkich plików Excel...")
    df_merged = merge_excels([EXCEL_FILE_1, EXCEL_FILE_2, EXCEL_FILE_3], EXCEL_MERGED)
    if df_merged is not None:
        print("🗺️ Tworzenie mapy z danych scalonych...")
        generate_merged_map(df_merged, MAP_MERGED)

    print(f"📦 Done. Uploading {EXCEL_MERGED} and map to OneDrive...")
    token = authenticate()
    upload_to_onedrive(EXCEL_MERGED, token)
    upload_to_onedrive(MAP_MERGED, token)

if __name__ == "__main__":
    main()
