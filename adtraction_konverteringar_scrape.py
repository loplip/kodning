import requests
from bs4 import BeautifulSoup
from datetime import datetime

URL = "https://adtraction.com/se/om-adtraction"
OUTFILE = "konverteringar_log.txt"

def fetch_konverteringar():
    response = requests.get(URL)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    # Leta efter rätt container: "Konverteringar" och ta nästa element med siffran
    box = soup.find(lambda tag: tag.name == "div" and "Konverteringar" in tag.text)
    if not box:
        raise Exception("Kunde inte hitta 'Konverteringar'")
    # Siffran finns som ett <span> eller <div> med siffran, hitta första siffersträng
    import re
    match = re.search(r"(\d[\d ]+\d)", box.text)
    if not match:
        raise Exception("Kunde inte hitta siffersträng i 'Konverteringar'-rutan")
    siffra = int(match.group(1).replace(" ", ""))
    return siffra

def log_konverteringar():
    siffra = fetch_konverteringar()
    datum = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(OUTFILE, "a", encoding="utf-8") as f:
        f.write(f"{datum} {siffra}\n")
    print(f"{datum} {siffra}")

if __name__ == "__main__":
    log_konverteringar()