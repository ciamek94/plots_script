import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

BASE_URL = "https://www.nieruchomosci-online.pl/szukaj.html?3,dzialka,sprzedaz,,Krak%C3%B3w:5600,,,25,-250000,1150,,,,,,,,,,,,,1"
MAX_PAGES = 20

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

results = []

for page in range(1, MAX_PAGES + 1):
    url = BASE_URL if page == 1 else f"{BASE_URL}&p={page}"
    print(f"Pobieram stronę {page}: {url}")
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"❌ Błąd pobierania strony {page}")
        break

    soup = BeautifulSoup(response.text, "html.parser")

    # Kafelki typu "tile"
    tiles = soup.find_all("div", class_="tile-inner")
    for tile in tiles:
        try:
            title_tag = tile.find("h2", class_="name")
            title = title_tag.text.strip() if title_tag else "Brak tytułu"
            link = title_tag.find("a")["href"] if title_tag and title_tag.find("a") else "Brak linku"

            price_tag = tile.find("p", class_="title-a")
            price_span = price_tag.find("span") if price_tag else None
            price = price_span.text.strip() if price_span else "Brak ceny"

            location_p = tile.find("p", class_="province")
            location = " ".join([el.text.strip() for el in location_p.find_all(['a', 'span'])]) if location_p else "Brak lokalizacji"

            results.append({
                "Title": title,
                "Location": location,
                "Price": price,
                "Link": link
            })
        except Exception as e:
            print(f"❗ Błąd parsowania kafelka tile: {e}")
            continue

    # Kafelki typu "tertiary"
    tertiaries = soup.find_all("div", class_="tertiary")
    for tertiary in tertiaries:
        try:
            title_tag = tertiary.find("h2", class_="name")
            title = title_tag.text.strip() if title_tag else "Brak tytułu"
            link = title_tag.find("a")["href"] if title_tag and title_tag.find("a") else "Brak linku"

            price_tag = tertiary.find("p", class_="title-a")
            price_span = price_tag.find("span") if price_tag else None
            price = price_span.text.strip() if price_span else "Brak ceny"

            location_p = tertiary.find("p", class_="province")
            location = " ".join([el.text.strip() for el in location_p.find_all(['a', 'span'])]) if location_p else "Brak lokalizacji"

            results.append({
                "Title": title,
                "Location": location,
                "Price": price,
                "Link": link
            })
        except Exception as e:
            print(f"❗ Błąd parsowania kafelka tertiary: {e}")
            continue

    time.sleep(1)  # nie przeciążajmy serwera

# Usunięcie duplikatów i zapis
df = pd.DataFrame(results).drop_duplicates()
df.to_excel("dzialki_krakow.xlsx", index=False)
print(f"✅ Zapisano {len(df)} ogłoszeń do pliku 'dzialki_krakow.xlsx'")
