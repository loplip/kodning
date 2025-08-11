import os
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from zoneinfo import ZoneInfo

URL = "https://adtraction.com/se/om-adtraction"
OUTFILE = "Adtraction_conversions.txt"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; AdtractionLogger/1.0; +https://example.com)"
}

# Tillåt vanliga mellanrum, icke-brytande (nbsp) och smala mellanrum
SEP_CHARS = "\u00A0\u202F "  # nbsp, thin space, space

def _to_int(s: str) -> int:
    s = s.strip().rstrip("+")
    s = re.sub(f"[{SEP_CHARS}]", "", s)  # ta bort tusentalsmellanrum
    s = re.sub(r"[^\d]", "", s)          # ta bort annat skräp
    return int(s)

def fetch_konverteringar() -> int:
    resp = requests.get(URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # 1) Hitta textnoden som exakt är "Konverteringar"
    label = soup.find(string=re.compile(r"^\s*Konverteringar\s*$", re.I))
    if not label:
        raise RuntimeError("Kunde inte hitta etiketten 'Konverteringar'")

    # 2) Försök hitta närmsta siffersträng i samma kort/container
    cand_re = re.compile(
        r"^\s*\d{1,3}(?:[ \u00A0\u202F]\d{3})+(?:\+)?\s*$|^\s*\d{5,}\s*$"
    )

    parent = label.parent
    for t in parent.find_all(string=cand_re):
        return _to_int(t)

    # 3) Annars leta framåt bland syskon/element ett litet tag
    for el in label.parent.next_elements:
        if isinstance(el, str) and cand_re.match(el):
            return _to_int(el)
        if getattr(el, "name", "") in {"section", "footer", "header"}:
            break

    # 4) Fallback: ta största talet på sidan (om layouten ändrats)
    nums = []
    for t in soup.find_all(string=re.compile(r"\d")):
        txt = t.strip()
        if cand_re.match(txt):
            try:
                nums.append(_to_int(txt))
            except ValueError:
                pass
    if nums:
        return max(nums)

    raise RuntimeError("Kunde inte extrahera siffra för 'Konverteringar'")

def ensure_outfile():
    # Skapa filen om den inte finns, så att git kan adda den
    if not os.path.exists(OUTFILE):
        with open(OUTFILE, "w", encoding="utf-8") as f:
            f.write("Datum Tid (Europe/Stockholm) Konverteringar\n")

def log_konverteringar():
    ensure_outfile()
    siffra = fetch_konverteringar()
    # Logga med svensk tid (CET/CEST)
    datum = datetime.now(ZoneInfo("Europe/Stockholm")).strftime("%Y-%m-%d %H:%M")
    with open(OUTFILE, "a", encoding="utf-8") as f:
        f.write(f"{datum} {siffra}\n")
    print(f"{datum} {siffra}")

if __name__ == "__main__":
    log_konverteringar()
