import re, time, random, datetime, os
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

# valfri fallback (pip install cloudscraper)
try:
    import cloudscraper  # type: ignore
except Exception:
    cloudscraper = None

BASE_URL   = "https://www.newegg.com/Gaming-Headsets/SubCategory/ID-3767"
PAGE_SIZE  = 96  # motsvarar "View: 96"
PAGES      = 1   # vi kör bara första sidan med 96 items
SKIP_SPONSORED = True
XLSX_FILE  = "data.xlsx"
SHEET_NAME = "FRCTL_headset"

PATTERN_DARK  = re.compile(r"fractal design scape dark", re.I)
PATTERN_LIGHT = re.compile(r"fractal design scape light", re.I)

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Upgrade-Insecure-Requests": "1",
    "Connection": "keep-alive",
    "Referer": "https://www.newegg.com/",
}

def looks_sponsored(tile):
    if not SKIP_SPONSORED:
        return False
    txt = tile.get_text(" ", strip=True).lower()
    return "sponsored" in txt or "advertisement" in txt

def parse_html(html):
    soup = BeautifulSoup(html, "html.parser")
    tiles = soup.select("div.item-cell") or soup.select("div.item-container, div.item-grid > div")
    items = []
    for idx_on_page, tile in enumerate(tiles, start=1):
        if looks_sponsored(tile):
            continue
        a = tile.select_one("a.item-title")
        title = a.get_text(strip=True) if a else ""
        items.append({"idx_on_page": idx_on_page, "title": title})
    return items

def fetch_html_with_requests():
    sess = requests.Session()
    sess.headers.update(HEADERS)
    # få cookies först
    try:
        sess.get("https://www.newegg.com/", timeout=30)
        time.sleep(0.5)
    except Exception:
        pass

    params = {"Order": "3", "Page": "1", "PageSize": str(PAGE_SIZE)}
    for attempt in range(3):
        r = sess.get(BASE_URL, params=params, timeout=45)
        if r.status_code == 200:
            return r.text
        if r.status_code == 403:
            break
        time.sleep(1.2 * (attempt + 1))
    # valfri fallback via cloudscraper om tillgängligt
    if cloudscraper is not None:
        scraper = cloudscraper.create_scraper(browser={"browser":"chrome","platform":"windows","mobile":False})
        r = scraper.get(BASE_URL, params=params, headers=HEADERS, timeout=60)
        r.raise_for_status()
        return r.text
    r.raise_for_status()  # kastar 403 om inget funkar

def find_positions_requests():
    positions = {"dark": None, "light": None}
    html = fetch_html_with_requests()
    items = parse_html(html)
    for visible_idx, it in enumerate(items, start=1):
        global_pos = (1 - 1) * PAGE_SIZE + visible_idx
        t = it["title"]
        if positions["dark"] is None and PATTERN_DARK.search(t):
            positions["dark"] = global_pos
        if positions["light"] is None and PATTERN_LIGHT.search(t):
            positions["light"] = global_pos
    return positions

def save_to_excel(date_str, store, dark_pos, light_pos):
    from openpyxl.utils import get_column_letter
    if os.path.exists(XLSX_FILE):
        wb = load_workbook(XLSX_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(["Datum", "Butik", "Fractal Design Scape Dark", "Fractal Design Scape Light"])
    else:
        ws = wb[SHEET_NAME]
    ws.append([date_str, store, dark_pos, light_pos])
    wb.save(XLSX_FILE)

if __name__ == "__main__":
    today = datetime.date.today().isoformat()
    store = "Newegg"
    pos = find_positions_requests()
    save_to_excel(today, store, pos["dark"], pos["light"])

    # Utskrift enligt mallen
    d = pos["dark"] if pos["dark"] is not None else "-"
    l = pos["light"] if pos["light"] is not None else "-"
    print('"Datum" "Butik" "Fractal Design Scape Dark" "Fractal Design Scape Light"')
    print(f"{today} {store} {d} {l}")
