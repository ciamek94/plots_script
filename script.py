import os
import pandas as pd
from otodom import main as main_script1  # zakładam, że skrypt1 ma funkcję main()
from olx import main as main_script2  # podobnie dla skryptu2
import folium

# Ścieżki do plików Excel ze skryptów 1 i 2
EXCEL_FILE_1 = os.path.join('dzialki', 'otodom_dzialki.xlsx')  # przykład
EXCEL_FILE_2 = os.path.join('dzialki', 'olx_dzialki.xlsx')    # przykład, z Twojego skryptu 2
EXCEL_MERGED = os.path.join('dzialki_merged', 'dzialki_merged.xlsx')
MAP_MERGED = os.path.join('dzialki_merged', 'map_merged.html')

def merge_excels(file1, file2, output_file):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Scal dane na podstawie unikalnego klucza - tutaj załóżmy, że unikalnym jest 'Link'
    df1.set_index('Link', inplace=True)
    df2.set_index('Link', inplace=True)

    # Połącz i zachowaj unikalne wiersze z obu
    df_merged = pd.concat([df1, df2[~df2.index.isin(df1.index)]])
    
    df_merged.reset_index(inplace=True)
    
    # Zapisz scalony Excel
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_merged.to_excel(output_file, index=False)
    print(f"💾 Scalony plik Excel zapisany: {output_file}")
    return df_merged

def generate_merged_map(df, map_path):
    KRAKOW_COORDS = (50.0647, 19.9450)
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Marker Krakowa
    folium.Marker(
        location=KRAKOW_COORDS,
        popup=folium.Popup("<b>Kraków</b><br>Reference point", max_width=200),
        tooltip="Kraków",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    location_to_listings = {}

    for _, row in df.iterrows():
        if not row.get("Active", True):
            continue

        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            continue

        coord_key = (round(lat, 5), round(lon, 5))
        if coord_key not in location_to_listings:
            location_to_listings[coord_key] = []

        location_to_listings[coord_key].append({
            "Title": row['Title'],
            "Location": row['Location'],
            "Price": row.get('Price last updated', row.get('Price at first find', '?')),
            "Link": row['Link'],
            "Distance": row.get("Distance from Krakow (km)", "?")
        })

    for (lat, lon), listings in location_to_listings.items():
        popup_html = ""
        for offer in listings:
            popup_html += f"""
            <b>{offer['Title']}</b><br>
            {offer['Location']}<br>
            {offer['Price']} PLN<br>
            <a href='{offer['Link']}' target='_blank'>Zobacz ogłoszenie</a>
            <hr>
            """

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{len(listings)} ogłoszeń",
            icon=folium.Icon(color="blue", icon="home", prefix="fa")
        ).add_to(m)

    m.save(map_path)
    print(f"🗺️ Scalona mapa zapisana: {map_path}")

def main():
    # 1. Uruchom skrypt 1
    print("🚀 Uruchamiam skrypt 1...")
    main_script1()

    # 2. Uruchom skrypt 2
    print("🚀 Uruchamiam skrypt 2...")
    main_script2()

    # 3. Scal Excels
    df_merged = merge_excels(EXCEL_FILE_1, EXCEL_FILE_2, EXCEL_MERGED)

    # 4. Wygeneruj scaloną mapę
    generate_merged_map(df_merged, MAP_MERGED)

if __name__ == "__main__":
    main()