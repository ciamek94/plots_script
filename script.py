import os
import pandas as pd
from otodom import main as main_script1  # Assuming script1 has a main() function
from olx import main as main_script2      # Assuming script2 has a main() function
import folium
import openpyxl
from openpyxl.utils import get_column_letter

# Paths to Excel files produced by script 1 and script 2
EXCEL_FILE_1 = os.path.join('dzialki', 'otodom_dzialki.xlsx')
EXCEL_FILE_2 = os.path.join('dzialki', 'olx_dzialki.xlsx')
EXCEL_MERGED = os.path.join('dzialki_merged', 'dzialki_merged.xlsx')
MAP_MERGED = os.path.join('dzialki_merged', 'map_merged.html')

def merge_excels(file1, file2, output_file):
    # Read two sheets from the first Excel file (for Krakow and Wielicki counties)
    xls1 = pd.ExcelFile(file1)
    df1_krakow = pd.read_excel(xls1, sheet_name='powiat krakowski')
    df1_wielicki = pd.read_excel(xls1, sheet_name='powiat wielicki')
    df1 = pd.concat([df1_krakow, df1_wielicki], ignore_index=True)

    # Add a 'Source' column to identify origin as 'otodom'
    df1['Source'] = 'otodom'

    # Read the second Excel file fully (assuming single sheet)
    df2 = pd.read_excel(file2)
    # Add a 'Source' column to identify origin as 'olx'
    df2['Source'] = 'olx'

    # Combine both dataframes vertically
    df_combined = pd.concat([df1, df2], ignore_index=True)

    # Normalize 'Link' column by stripping spaces and lowering case for proper deduplication
    df_combined['Link'] = df_combined['Link'].astype(str).str.strip().str.lower()

    # Remove duplicate listings by 'Link', keeping only the first occurrence
    df_unique = df_combined.drop_duplicates(subset=['Link'], keep='first').reset_index(drop=True)

    # Define desired column order (adjust as needed)
    desired_columns = [
        'Title',
        'Location',
        'Latitude',
        'Longitude',
        'Price last updated',
        'Price at first find',
        'Distance from Krakow (km)',
        'Link',
        'Source',
        'Active'
    ]
    columns_to_save = [col for col in desired_columns if col in df_unique.columns]

    # Ensure output directory exists and save merged data to Excel with correct column order
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_unique.to_excel(output_file, index=False, columns=columns_to_save)

    # Adjust Excel column widths automatically based on max length of content
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active

    for col_idx, col in enumerate(columns_to_save, 1):
        max_length = len(col)
        for cell in ws[get_column_letter(col_idx)]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2  # Add some padding for readability
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    wb.save(output_file)
    print(f"💾 Merged Excel file saved with adjusted column widths: {output_file}")

    return df_unique

def generate_merged_map(df, map_path):
    # Coordinates of Krakow city center (reference point)
    KRAKOW_COORDS = (50.0647, 19.9450)
    # Initialize folium map centered on Krakow
    m = folium.Map(location=KRAKOW_COORDS, zoom_start=10)

    # Add marker for Krakow city center
    folium.Marker(
        location=KRAKOW_COORDS,
        popup=folium.Popup("<b>Kraków</b><br>Reference point", max_width=200),
        tooltip="Kraków",
        icon=folium.Icon(color="purple", icon="star", prefix="fa")
    ).add_to(m)

    # Group listings by rounded coordinates (to cluster offers at same location)
    location_to_listings = {}

    for _, row in df.iterrows():
        # Skip inactive listings
        if not row.get("Active", True):
            continue

        lat = row.get("Latitude")
        lon = row.get("Longitude")
        # Skip listings with missing coordinates
        if pd.isna(lat) or pd.isna(lon):
            continue

        # Round coordinates to 5 decimal places for grouping
        coord_key = (round(lat, 5), round(lon, 5))
        if coord_key not in location_to_listings:
            location_to_listings[coord_key] = []

        # Append listing info along with its source (otodom or olx)
        location_to_listings[coord_key].append({
            "Title": row['Title'],
            "Location": row['Location'],
            "Price": row.get('Price last updated', row.get('Price at first find', '?')),
            "Link": row['Link'],
            "Distance": row.get("Distance from Krakow (km)", "?"),
            "Source": row.get("Source", "unknown")
        })

    # Create markers for each location with a popup listing all offers there
    for (lat, lon), listings in location_to_listings.items():
        popup_html = ""
        for offer in listings:
            popup_html += f"""
            <b>{offer['Title']}</b><br>
            {offer['Location']}<br>
            {offer['Price']} PLN<br>
            <a href='{offer['Link']}' target='_blank'>View listing</a>
            <hr>
            """

        # Decide marker color based on the source of the first listing at this location
        first_source = listings[0]['Source'].lower()
        if first_source == 'otodom':
            marker_color = "green"
        elif first_source == 'olx':
            marker_color = "blue"
        else:
            marker_color = "gray"

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"{len(listings)} listings",
            icon=folium.Icon(color=marker_color, icon="home", prefix="fa")
        ).add_to(m)

    # Save the map as an HTML file
    m.save(map_path)
    print(f"🗺️ Merged map saved: {map_path}")

def main():
    # 1. Run script 1 (Otodom scraper)
    print("🚀 Running script 1 (Otodom)...")
    main_script1()

    # 2. Run script 2 (OLX scraper)
    print("🚀 Running script 2 (OLX)...")
    main_script2()

    # 3. Merge Excel files and remove duplicate listings
    df_merged = merge_excels(EXCEL_FILE_1, EXCEL_FILE_2, EXCEL_MERGED)

    # 4. Generate merged map with color-coded markers
    generate_merged_map(df_merged, MAP_MERGED)

if __name__ == "__main__":
    main()
