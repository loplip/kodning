import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pytz
import pandas as pd
import time
import random

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/132.0.0.0 Safari/537.36"
    )
}

def fetch_html(url):
    # Lägg in en liten slumpmässig paus för att efterlikna mänskligt beteende
    time.sleep(random.uniform(2, 5))
    resp = requests.get(url, headers=HEADERS)
    resp.raise_for_status()
    return resp.text

def get_rank(search_url, product_name):
    html = fetch_html(search_url)
    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select("div.item-cell")
    for idx, card in enumerate(cards, start=1):
        title = card.select_one(".item-title")
        if title and product_name.lower() in title.get_text(strip=True).lower():
            return idx
    return None

def append_to_excel(file_name, sheet_name, row):
    try:
        df_existing = pd.read_excel(file_name, sheet_name=sheet_name)
        df = pd.concat([df_existing, pd.DataFrame([row])], ignore_index=True)
    except (FileNotFoundError, ValueError):
        df = pd.DataFrame([row], columns=["Datum", "Scape Dark rank", "Scape Light rank"])
    df.to_excel(file_name, sheet_name=sheet_name, index=False)

def main():
    tz = pytz.timezone('Europe/Stockholm')
    now = datetime.now(tz).strftime('%Y-%m-%d %H:%M')
    dark_rank = get_rank("https://www.newegg.com/p/pl?d=fractal+scape+dark", "Scape Dark")
    light_rank = get_rank("https://www.newegg.com/p/pl?d=fractal+scape+light", "Scape Light")

    append_to_excel("data.xlsx", "FRACTL_inet", {"Datum": now,
                                                 "Scape Dark rank": dark_rank,
                                                 "Scape Light rank": light_rank})
    print(f"La till: {dark_rank} {light_rank}")

if __name__ == "__main__":
    main()
