import re
import requests
from bs4 import BeautifulSoup

URL = "https://adtraction.com/se/om-adtraction"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; AdtractionLogger/1.0)"}

SEP_CHARS = "\u00A0\u202F "  # nbsp, thin space, vanligt mellanslag

def _to_int(s: str) -> int:
    # behåll bara siffror; tillåt tusentalsavskiljare (mellanrum) och + i slutet
    s = s.strip().rstrip("+")
    s = re.sub(f"[{SEP_CHARS}]", "", s)
    s = re.sub(r"[^\d]", "", s)
    return int(s)

def fetch_konverteringar():
    resp = requests.get(URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # 1) Hitta textnoden som ÄR "Konverteringar" (exakt match, skiftlägesokänslig)
    label = soup.find(string=re.compile(r"^\s*Konverteringar\s*$", re.I))
    if not label:
        raise RuntimeError("Kunde inte hitta etiketten 'Konverteringar'")

    # 2) Leta efter första siffersträng i nära närhet: först i samma kort/parent,
    #    sen i följande syskon/element.
    CAND_RE = re.compile(r"^\s*\d{1,3}(?:[ \u00A0\u202F]\d{3})+(?:\+)?\s*$|^\s*\d{5,}\s*$")
    # a) inom samma kort/parent
    parent = label.parent
    for t in parent.find_all(string=CAND_RE):
        return _to_int(t)

    # b) syskon/”nästa element” – gå framåt en bit och ta första som matchar
    for el in label.parent.next_elements:
        if isinstance(el, str) and CAND_RE.match(el):
            return _to_int(el)
        # stoppa om vi vandrat för långt bort från kortet (heuristik)
        if getattr(el, "name", "") in {"section", "footer", "header"}:
            break

    # 3) Fallback: ta största talet på hela sidan (om layouten ändras)
    nums = []
    for t in soup.find_all(string=re.compile(r"\d")):
        txt = t.strip()
        if CAND_RE.match(txt):
            try:
                nums.append(_to_int(txt))
            except ValueError:
                pass
    if nums:
        return max(nums)

    raise RuntimeError("Kunde inte extrahera siffra för 'Konverteringar'")
