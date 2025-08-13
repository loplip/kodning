import re, datetime, os, time
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

BASE_URL  = "https://www.newegg.com/Gaming-Chairs/SubCategory/ID-3628"
PAGE_SIZE = 96
XLSX_FILE = "data.xlsx"
SHEET_NAME = "FRCTL_chair"

# Regex för varianter (titel + specs/primary color). Case-insensitive.
PATTERNS = {
    "fabric_dark":  re.compile(r"\bfractal(?:\s+design)?\b.*\brefine\b.*?(fabric|woven).*?(dark|black|charcoal|noir|graphite|charcoal\s*gray|grey)", re.I|re.S),
    "fabric_light": re.compile(r"\bfractal(?:\s+design)?\b.*\brefine\b.*?(fabric|woven).*?(light|white|silver|light\s*gray|light\s*grey|grey|gray)", re.I|re.S),
    "mesh_dark":    re.compile(r"\bfractal(?:\s+design)?\b.*\brefine\b.*?mesh.*?(dark|black|charcoal|noir|graphite)", re.I|re.S),
    "mesh_light":   re.compile(r"\bfractal(?:\s+design)?\b.*\brefine\b.*?mesh.*?(light|white|silver|light\s*gray|light\s*grey|grey|gray)", re.I|re.S),
    "alcantara":    re.compile(r"\bfractal(?:\s+design)?\b.*\brefine\b.*?alcantara", re.I|re.S),
}

def looks_sponsored_text(txt: str) -> bool:
    t = txt.lower()
    return ("sponsored" in t) or ("advertisement" in t)

def fetch_items_with_playwright():
    """Returnerar lista [{'title':..., 'fulltext':...}] i visningsordning (skippar Sponsored)."""
    params = f"?Order=3&Page=1&PageSize={PAGE_SIZE}"
    url = BASE_URL + params
    items = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/126.0.0.0 Safari/537.36"),
            locale="en-US",
        )
        page = context.new_page()

        # Cookie warm-up
        page.goto("https://www.newegg.com/", wait_until="domcontentloaded", timeout=60000)
        time.sleep(0.4)

        # Gå till listan
        page.goto(url, wait_until="networkidle", timeout=90000)

        # Vänta in kort
        for _ in range(8):
            if page.locator("div.item-cell a.item-title").first.is_visible():
                break
            time.sleep(0.6)

        tiles = page.locator("div.item-cell")
        if tiles.count() == 0:
            tiles = page.locator("div.item-container, div.item-grid > div")

        count = tiles.count()
        for i in range(count):
            t = tiles.nth(i)
            # Hämta all text i kortet (titel + specs, inkl. 'Primary Color')
            try:
                fulltext = t.inner_text().replace("®", " ").replace("™", " ")
            except Exception:
                continue
            if looks_sponsored_text(fulltext):
                continue

            title = ""
            try:
                if t.locator("a.item-title").count():
                    title = t.locator("a.item-title").inner_text().strip()
            except Exception:
                pass

            if title or fulltext:
                items.append({
                    "title": title,
                    "fulltext": (title + " " + fulltext).lower()
                })

        context.close()
        browser.close()

    return items

def find_positions(items):
    """Returnerar positions-dict för våra 5 varianter (1-baserat index i visningsordning)."""
    positions = {k: None for k in PATTERNS.keys()}
    for idx, it in enumerate(items, start=1):
        txt = it["fulltext"]
        if ("fractal" not in txt) or ("refine" not in txt):
            continue
        for key, pat in PATTERNS.items():
            if positions[key] is None and pat.search(txt):
                positions[key] = idx
        if all(v is not None for v in positions.values()):
            break
    return positions

def save_to_excel(datetime_str, store, pos):
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
        datetime_str, store,
        pos["fabric_dark"], pos["fabric_light"],
        pos["mesh_dark"], pos["mesh_light"],
        pos["alcantara"]
    ])
    wb.save(XLSX_FILE)

if __name__ == "__main__":
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    store = "Newegg"
    items = fetch_items_with_playwright()
    pos = find_positions(items)
    save_to_excel(now_str, store, pos)

    x = pos["fabric_dark"]  if pos["fabric_dark"]  is not None else "-"
    y = pos["fabric_light"] if pos["fabric_light"] is not None else "-"
    z = pos["mesh_dark"]    if pos["mesh_dark"]    is not None else "-"
    v = pos["mesh_light"]   if pos["mesh_light"]   is not None else "-"
    w = pos["alcantara"]    if pos["alcantara"]    is not None else "-"
    print(f"La till positionerna {x}, {y}, {z}, {v}, {w} i FRCTL_chair")
