import osmnx as ox
import pandas as pd

# -------------------------------
# CONFIGURATION
# -------------------------------

center_lat, center_lon = 50.0647, 19.9450  
radius_m = 40000  # 40 km radius
tags = {'place': ['city', 'town', 'village', 'hamlet']}
output_file = "town_list.txt"

# -------------------------------
# FETCH DATA
# -------------------------------

gdf = ox.features_from_point((center_lat, center_lon), tags=tags, dist=radius_m)
gdf = gdf[gdf['name'].notna()]

# -------------------------------
# PROCESS DATA
# -------------------------------

# 1. Reproject to metric CRS for accurate centroid calculation
gdf = gdf.to_crs(epsg=2180)
gdf['centroid'] = gdf.geometry.centroid

# 2. Convert centroid to WGS84 and extract lat/lon
gdf = gdf.set_geometry('centroid').to_crs(epsg=4326)
gdf['lat'] = gdf.geometry.y
gdf['lon'] = gdf.geometry.x

# Keep only name, lat, lon
places_df = gdf[['name', 'lat', 'lon']]
places_df.loc[:, 'name'] = places_df['name'].str.lower()

# -------------------------------
# SAVE TO TXT
# -------------------------------

places_df.to_csv(output_file, sep='|', index=False, header=False)
print(f"Saved {len(places_df)} places to '{output_file}'")
