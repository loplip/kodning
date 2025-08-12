import re, time, random, datetime, os, unicodedata
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

try:
    import cloudscraper
except Exception:
    cloudscraper = None

BASE_URL   = "https://www.newegg.com/Gaming-Chairs/SubCategory/ID-3628"
PAGE_SIZE  = 96
PAGES      = 1
SKIP_SPONSORED = True
XLSX_FILE  = "data.xlsx"
SHEET_NAME = "FRCTL_chair"

PATTERNS = {
    "fabric_dark": re.compile(
        r"fractal\s+(?:design\s+)?refine.*?(fabric|woven).*?(dark|black|charcoal|noir|graphite|charcoal\s*gray|grey)",
        re.I | re.S),
    "fabric_light": re.compile(
        r"fractal\s+(?:design\s+)?refine.*?(fabric|woven).*?(light|white|silver|light\s*gray|light\s*grey|grey|gray)",
        re.I | re.S),
    "mesh_dark": re.compile(
        r"fractal\s+(?:design\s+)?refine.*?mesh.*?(dark|black|charcoal|noir|graphite)",
        re.I | re.S),
    "mesh_light": re.compile(
        r"fractal\s+(?:design\s+)?refine.*?mesh.*?(light|white|silver|light\s*gray|light\s*grey|grey|gray)",
        re.I | re.S),
    "alcantara": re.compile(r"fractal\s+(?:design\s+)?refine.*?alcantara", re.I | re.S),
}

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

def normalize_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    return s.replace("®", " ").replace("™", " ")

def parse_html(html):
    soup = BeautifulSoup(html, "html.parser")
    tiles = soup.select("div.item-cell") or soup.select("div.item-container, div.item-grid > div")
    items = []
    for idx_on_page, tile in enumerate(tiles, start=1):
        if looks_sponsored(tile):
            continue
        a = tile.select_one("a.item-title")
        title = a.get_text(strip=True) if a else ""
        spec = tile.get_text(" ", strip=True)
        fulltext = normalize_text((title + " " + spec)).lower()
        items.append({"idx_on_page": idx_on_page, "title": title, "fulltext": fulltext})
    return items

def fetch_html_with_requests(page=1):
    sess = requests.Session()
    sess.headers.update(HEADERS)
    try:
        sess.get("https://www.newegg.com/", timeout=30)
        time.sleep(0.4 + random.random()*0.6)
    except Exception:
        pass
    params = {"Order": "3", "Page": str(page), "PageSize": str(PAGE_SIZE)}
    for attempt in range(3):
        r = sess.get(BASE_URL, params=params, timeout=45)
        if r.status_code == 200 and "captcha" not in r.text.lower():
            return r.text
        time.sleep(1.0 + attempt*0.8)
    if cloudscraper is not None:
        scraper = cloudscraper.create_scraper(browser={"browser": "chrome", "platform": "windows", "mobile": False})
        r = scraper.get(BASE_URL, params=params, headers=HEADERS, timeout=60)
        r.raise_for_status()
        return r.text
    r.raise_for_status()
    return ""

def find_positions():
    positions = {k: None for k in PATTERNS}
    for page in range(1, PAGES + 1):
        html = fetch_html_with_requests(page=page)
        items = parse_html(html)
        for visible_idx, it in enumerate(items, start=1):
            global_pos = (page - 1) * PAGE_SIZE + visible_idx
            txt = it["fulltext"]
            if "fractal" not in txt or "refine" not in txt:
                continue
            for key, pattern in PATTERNS.items():
                if positions[key] is None and pattern.search(txt):
                    positions[key] = global_pos
    return positions

def save_to_excel(date_time_str, store, pos):
    if os.path.exists(XLSX_FILE):
        wb = load_workbook(XLSX_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append([
            "Datum", "Butik",
            "Refine Fabric Dark", "Refine Fabric Light",
            "Refine Mesh Dark",   "Refine Mesh Light",
            "Refine Alcantara"
        ])
    else:
        ws = wb[SHEET_NAME]
    ws.append([
        date_time_str, store,
        pos["fabric_dark"], pos["fabric_light"],
        pos["mesh_dark"], pos["mesh_light"],
        pos["alcantara"]
    ])
    wb.save(XLSX_FILE)

if __name__ == "__main__":
    # YYYY-MM-DD HH:MM
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    store = "Newegg"
    pos = find_positions()
    save_to_excel(now_str, store, pos)

    x = pos["fabric_dark"]  if pos["fabric_dark"]  is not None else "-"
    y = pos["fabric_light"] if pos["fabric_light"] is not None else "-"
    z = pos["mesh_dark"]    if pos["mesh_dark"]    is not None else "-"
    v = pos["mesh_light"]   if pos["mesh_light"]   is not None else "-"
    w = pos["alcantara"]    if pos["alcantara"]    is not None else "-"
    print(f"La till positionerna {x}, {y}, {z}, {v}, {w} i FRCTL_chair")
