import os
import re
import time
import random
import datetime
from typing import List, Dict, Optional, Tuple

from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# Optionalt stealth-plugin
try:
    from playwright_stealth import stealth_sync  # type: ignore
except ImportError:
    stealth_sync = None  # type: ignore

# ========= Konfiguration =========
XLSX_FILE = "data.xlsx"
SHEET_NAME = "FRACTL_inet"

# Sök-URL:er – vi trycker in Order=3 och PageSize=96 för max produkter utan paginering
SEARCHES = {
    "Scape Dark rank": "https://www.newegg.com/p/pl?d=fractal+scape+dark&Order=3&PageSize=96",
    "Scape Light rank": "https://www.newegg.com/p/pl?d=fractal+scape+light&Order=3&PageSize=96",
}

# Mönster för att matcha rätt variant i korttexten (robust mot färgord/synonymer)
PATTERNS = {
    "Scape Dark rank": re.compile(
        r"\bfractal(?:\s+design)?\b.*\bscape\b.*\b(dark|black|charcoal|noir|graphite|midnight)\b",
        re.I | re.S,
    ),
    "Scape Light rank": re.compile(
        r"\bfractal(?:\s+design)?\b.*\bscape\b.*\b(light|white|silver|pearl|snow|light\s*gray|light\s*grey|grey|gray)\b",
        re.I | re.S,
    ),
}


# ========= Hjälpfunktioner (Playwright) =========
def looks_sponsored_text(txt: str) -> bool:
    t = txt.lower()
    return ("sponsored" in t) or ("advertisement" in t)


def dismiss_popups(page) -> None:
    selectors = [
        'button:has-text("Accept All")',
        'button:has-text("Accept")',
        'button:has-text("Continue")',
        'button[class*="close"]',
        'div[class*="close"]',
        'button[aria-label="Close"]',
        '#truste-consent-button',
    ]
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.is_visible():
                loc.click(timeout=500)
                time.sleep(0.2 + random.uniform(0, 0.3))
        except Exception:
            pass


def goto_with_retries(page, url: str, attempts: int = 4, timeout: int = 70000) -> None:
    last_err = None
    for i in range(attempts):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            dismiss_popups(page)
            # vänta tills nätverket är lugnt
            try:
                page.wait_for_load_state("networkidle", timeout=timeout)
            except Exception:
                pass

            # trigga lazy-load
            for _ in range(3):
                page.mouse.wheel(0, 2500)
                time.sleep(0.5)
            page.evaluate("window.scrollTo(0, 0)")

            # acceptera fler selektorer (markup varierar)
            ANY_LIST_SEL = (
                "div.item-cell a.item-title, "
                "div.item-container a.item-title, "
                "a.item-title, "
                "div.item-cell, "
                "div.item-container, "
                "div.item-grid > div"
            )
            page.wait_for_selector(ANY_LIST_SEL, state="attached", timeout=timeout)
            return
        except Exception as e:
            last_err = e
            time.sleep(1.2 * (i + 1))
    raise last_err if last_err else RuntimeError("Navigation failed")



def build_browser_context(p, headless: bool = True):
    USER_AGENTS = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) Gecko/20100101 Firefox/124.0",
    ]
    ua = random.choice(USER_AGENTS)
    viewport = {"width": random.randint(1280, 1920), "height": random.randint(720, 1080)}

    proxy_settings = None
    proxy_server = os.getenv("PROXY_SERVER")
    if proxy_server:
        proxy_settings = {"server": proxy_server}
        username = os.getenv("PROXY_USERNAME")
        password = os.getenv("PROXY_PASSWORD")
        if username and password:
            proxy_settings.update({"username": username, "password": password})

    browser = p.chromium.launch(
        headless=headless,
        proxy=proxy_settings,
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = browser.new_context(viewport=viewport, user_agent=ua, locale="en-US")
    context.add_init_script('Object.defineProperty(navigator, "webdriver", {get: () => undefined})')
    return browser, context


def collect_items_on_page(page) -> List[Dict[str, str]]:
    # Hämta kort; fallback för alternativa layouter
    tiles = page.locator("div.item-cell")
    if tiles.count() == 0:
        tiles = page.locator("div.item-container, div.item-grid > div")

    items: List[Dict[str, str]] = []
    count = tiles.count()
    for i in range(count):
        t = tiles.nth(i)
        try:
            fulltext = t.inner_text(timeout=3000).replace("®", " ").replace("™", " ")
        except Exception:
            continue
        if looks_sponsored_text(fulltext):
            continue

        title = ""
        try:
            if t.locator("a.item-title").count():
                title = t.locator("a.item-title").inner_text(timeout=2000).strip()
        except Exception:
            pass

        if title or fulltext:
            items.append({"title": title, "fulltext": (title + " " + fulltext).lower()})
    return items


def first_position_for_variant(items: List[Dict[str, str]], pattern: re.Pattern) -> Optional[int]:
    for idx, it in enumerate(items, start=1):
        txt = it["fulltext"]
        # Säkerställ att det verkligen rör Fractal Scape
        if ("fractal" not in txt) or ("scape" not in txt):
            continue
        if pattern.search(txt):
            return idx
    return None


# ========= Excel-hjälp =========
def save_to_excel(datetime_str: str, dark_pos: Optional[int], light_pos: Optional[int]) -> None:
    if os.path.exists(XLSX_FILE):
        wb = load_workbook(XLSX_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(["Datum", "Scape Dark rank", "Scape Light rank"])
    else:
        ws = wb[SHEET_NAME]

    ws.append([datetime_str, dark_pos if dark_pos is not None else "-", light_pos if light_pos is not None else "-"])
    wb.save(XLSX_FILE)


# ========= Huvudlogik =========
def main() -> None:
    # Tidsstämpel i Europe/Stockholm
    tz = datetime.timezone(datetime.timedelta(hours=2))  # Sommartid; byt till +1 vid vintertid om du vill vara exakt
    now_str = datetime.datetime.now(tz=tz).strftime("%Y-%m-%d %H:%M")

    with sync_playwright() as p:
        browser, context = build_browser_context(p, headless=True)
        page = context.new_page()

        if stealth_sync is not None:
            try:
                stealth_sync(page)
            except Exception:
                pass

        # Värm upp cookies
        try:
            page.goto("https://www.newegg.com/", wait_until="domcontentloaded", timeout=45000)
            dismiss_popups(page)
            time.sleep(0.3 + random.uniform(0, 0.5))
        except Exception:
            pass

        results: Dict[str, Optional[int]] = {"Scape Dark rank": None, "Scape Light rank": None}

        for key, url in SEARCHES.items():
            goto_with_retries(page, url, attempts=3)
            items = collect_items_on_page(page)
            results[key] = first_position_for_variant(items, PATTERNS[key])
            # Små pauser mellan sidorna
            time.sleep(0.4 + random.uniform(0, 0.6))

        context.close()
        browser.close()

    save_to_excel(now_str, results["Scape Dark rank"], results["Scape Light rank"])
    d = results["Scape Dark rank"] if results["Scape Dark rank"] is not None else "-"
    l = results["Scape Light rank"] if results["Scape Light rank"] is not None else "-"
    print(f"La till: {d} {l} i {SHEET_NAME}")


if __name__ == "__main__":
    main()
