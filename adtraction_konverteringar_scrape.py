import os
from datetime import datetime

import requests
from bs4 import BeautifulSoup


URL = "https://adtraction.com/se/om-adtraction"
OUTFILE = "adtraction_statistics.csv"


def fetch_statistics():
    """Hämta statistik-rutor från Adtractions "om"-sida.

    Returnerar en lista av tuple (titel, värde) i samma ordning som de
    förekommer på sidan.
    """

    response = requests.get(URL)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    boxes = soup.find_all("div", class_="statistics_card_container")
    if not boxes:
        raise Exception("Kunde inte hitta statistikrutor på sidan")

    stats = []
    for box in boxes:
        title_tag = box.find(class_="statistics_card_title")
        value_tag = box.find(class_="statistics_card_text")
        if not title_tag or not value_tag:
            continue

        # Ta bort mellanslag i siffran för att kunna konvertera till int
        value = int(value_tag.text.strip().replace(" ", ""))
        stats.append((title_tag.text.strip(), value))

    return stats


def log_statistics():
    stats = fetch_statistics()
    datum = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Skapa filen med rubriker om den inte finns
    file_exists = os.path.exists(OUTFILE)
    with open(OUTFILE, "a", encoding="utf-8") as f:
        if not file_exists:
            headers = ["timestamp"] + [title for title, _ in stats]
            f.write(",".join(headers) + "\n")

        values = [datum] + [str(value) for _, value in stats]
        f.write(",".join(values) + "\n")

    print(datum, dict(stats))


if __name__ == "__main__":
    log_statistics()

