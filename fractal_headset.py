import re, datetime, os, time, random
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

BASE_URL   = "https://www.newegg.com/Gaming-Headsets/SubCategory/ID-3767"
PAGE_SIZE  = 96  # View: 96 på en sida
XLSX_FILE  = "data.xlsx"
SHEET_NAME = "FRCTL_headset"

# Matcha varianter: "fractal ... scape ... dark/light"
PATTERN_DARK  = re.compile(r"\bfractal\b.*\bscape\b.*\bdark\b", re.I)
PATTERN_LIGHT = re.compile(r"\bfractal\b.*\bscape\b.*\blight\b", re.I)

def fetch_items_with_playwright():
    """Returnerar lista av dicts: [{'title':..., 'brand':...}, ...] i visningsordning."""
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
        page.goto("https://www.newegg.com/", wait_until="domcontentloaded", timeout=60000)
        time.sleep(0.5)  # få cookies

        # Gå till kategori med sortering & PageSize=96
        page.goto(url, wait_until="networkidle", timeout=90000)

        # Vänta in produktkort (hantera ev. Cloudflare med några försök)
        for _ in range(6):
            if page.locator("div.item-cell a.item-title").first.is_visible():
                break
            time.sleep(1.0)

        tiles = page.locator("div.item-cell")
        count = tiles.count()
        # Fallback om gridmark-up varierar
        if count == 0:
            tiles = page.locator("div.item-container, div.item-grid > div")
            count = tiles.count()

        for i in range(count):
            t = tiles.nth(i)
            title = ""
            brand = ""
            try:
                if t.locator("a.item-title").count():
                    title = t.locator("a.item-title").inner_text().strip()
                # brand kan saknas
                if t.locator("a.item-brand").count():
                    brand = t.locator("a.item-brand").inner_text().strip()
                elif t.locator("div.item-branding a").count():
                    brand = t.locator("div.item-branding a").inner_text().strip()
            except Exception:
                pass
            if title:
                items.append({"title": title, "brand": brand})
        context.close()
        browser.close()
    return items

def find_positions(items):
    pos_dark = None
    pos_light = None
    for idx, it in enumerate(items, start=1):
        combo = f"{it.get('brand','')} {it.get('title','')}"
        if pos_dark is None and PATTERN_DARK.search(combo):
            pos_dark = idx
        if pos_light is None and PATTERN_LIGHT.search(combo):
            pos_light = idx
        if pos_dark is not None and pos_light is not None:
            break
    return pos_dark, pos_light

def save_to_excel(datetime_str, store, dark_pos, light_pos):
    if os.path.exists(XLSX_FILE):
        wb = load_workbook(XLSX_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(["Datum", "Butik", "Scape Dark", "Scape Light"])
    else:
        ws = wb[SHEET_NAME]
        if ws.max_row == 0:
            ws.append(["Datum", "Butik", "Scape Dark", "Scape Light"])
    ws.append([datetime_str, "Newegg", dark_pos, light_pos])
    wb.save(XLSX_FILE)

if __name__ == "__main__":
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    items = fetch_items_with_playwright()
    dark_pos, light_pos = find_positions(items)
    save_to_excel(now_str, "Newegg", dark_pos, light_pos)

    d = dark_pos if dark_pos is not None else "-"
    l = light_pos if light_pos is not None else "-"
    print(f"La till positionerna {d}, {l} i {SHEET_NAME}")
