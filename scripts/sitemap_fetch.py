#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_sitemaps.py (v3, strukturerad)
- Läser sources/sitemaps_bolag.xlsx
- Hämtar & parser sitemaps (XML + fallback på text/Yoast-tabell)
- Tar nya poster (datum+bolag+typ+länk) och lägger till i data/data_sitemap.xlsx
- Sorterar 'databas' efter senaste datum
- Skriver ENDAST nya rader i terminalen: "YYYY-MM-DD Bolag: Länk"
- Ingen loggfil, inga "Hämtar ..." och inget "Klar" i slutet
- Lägger kolumnen "Sitemap" som källa till varje Länk
"""

import sys
import re
import os
from datetime import datetime
from pathlib import Path

import requests
import pandas as pd
from xml.etree import ElementTree as ET

# --- Repo-paths --------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]  # <repo>/  (filen ligger i /scripts)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.common.paths import DATA_DIR, SOURCES_DIR  # kräver uppdatering enligt steg 1

INPUT_XLSX = SOURCES_DIR / "sitemaps_bolag.xlsx"
OUTPUT_XLSX = DATA_DIR / "data_sitemap.xlsx"
SHEET = "databas"

DATE_PAT = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")

# --- Helpers -----------------------------------------------------------------
def norm_date(s: str | None) -> str | None:
    if not s:
        return None
    m = DATE_PAT.search(s)
    if m:
        return m.group(1).replace("/", "-")
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%d %H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return None

def fetch(url: str, timeout=25):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "sv-SE,sv;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "keep-alive",
    }
    with requests.Session() as s:
        s.headers.update(headers)
        r = s.get(url, timeout=timeout, allow_redirects=True)
        r.raise_for_status()
        return r.text, r.headers.get("Content-Type", "")


def looks_like_xml(text: str) -> bool:
    t = text.lstrip().lower()
    return t.startswith("<?xml") or "<urlset" in t or "<sitemapindex" in t

# --- Parsers -----------------------------------------------------------------
def parse_xml(xml_text: str) -> list[tuple[str, str | None]]:
    """Return list of (url, date). Följer ev. <sitemapindex> rekursivt."""
    out: list[tuple[str, str | None]] = []
    root = ET.fromstring(xml_text.encode("utf-8"))
    ns = {"ns": root.tag.split('}')[0].strip('{')} if root.tag.startswith("{") else {}

    # sitemapindex -> följ undersitemaps
    sitems = root.findall(".//ns:sitemap", ns) if ns else root.findall(".//sitemap")
    for sm in sitems:
        loc_el = sm.find("ns:loc", ns) if ns else sm.find("loc")
        if loc_el is not None and loc_el.text:
            try:
                text, _ = fetch(loc_el.text.strip())
                out.extend(parse(text))
            except Exception as e:
                print(f"Fel vid hämtning av undersitemap: {loc_el.text} ({e})")

    # urlset -> samla länkar
    url_elems = root.findall(".//ns:url", ns) if ns else root.findall(".//url")
    date_field_names = {
        "lastmod", "last-mod", "last-modified", "last_modified",
        "modified", "updated", "pubdate", "publication_date", "publication-date"
    }
    for u in url_elems:
        loc_el = u.find("ns:loc", ns) if ns else u.find("loc")
        loc = loc_el.text.strip() if (loc_el is not None and loc_el.text) else None
        date_val = None
        for tag in date_field_names:
            el = u.find(f"ns:{tag}", ns) if ns else u.find(tag)
            if el is not None and el.text:
                date_val = norm_date(el.text)
                if date_val:
                    break
        if not date_val:
            for el in u.iter():
                if el.text:
                    maybe = norm_date(el.text)
                    if maybe:
                        date_val = maybe
                        break
        if loc:
            out.append((loc, date_val))
    return out

def parse_table_or_text(text: str) -> list[tuple[str, str | None]]:
    out: list[tuple[str, str | None]] = []
    for line in text.splitlines():
        m_url = re.search(r"https?://\S+", line)
        if not m_url:
            continue
        url = m_url.group(0)
        d = norm_date(line)
        out.append((url, d))
    if not out:
        urls = re.findall(r"https?://[\w\-\./%?=&#]+", text)
        dates = re.findall(DATE_PAT, text)
        if urls:
            if len(urls) == len(dates):
                out = list(zip(urls, [d[0] for d in dates]))
            else:
                out = [(u, None) for u in urls]
    return out

def parse(text: str) -> list[tuple[str, str | None]]:
    if looks_like_xml(text):
        try:
            return parse_xml(text)
        except Exception:
            return parse_table_or_text(text)
    return parse_table_or_text(text)

# --- Main --------------------------------------------------------------------
def main():
    if not INPUT_XLSX.exists():
        print(f"Hittar inte {INPUT_XLSX}")
        sys.exit(1)

    src = pd.read_excel(INPUT_XLSX)
    src = src.rename(columns={c: c.strip() for c in src.columns})
    required = ["Bolag", "Typ av sajt", "Länk"]
    for r in required:
        if r not in src.columns:
            raise SystemExit(f"Saknar kolumn '{r}' i {INPUT_XLSX}")

    # Läs befintlig databas (nu med Sitemap-kolumn)
    if OUTPUT_XLSX.exists():
        try:
            db = pd.read_excel(OUTPUT_XLSX, sheet_name=SHEET)
        except Exception:
            db = pd.DataFrame(columns=["Senast modifierad", "Bolag", "Typ av sajt", "Länk", "Sitemap"])
    else:
        db = pd.DataFrame(columns=["Senast modifierad", "Bolag", "Typ av sajt", "Länk", "Sitemap"])

    # Nyckel för att undvika dubletter i historiken
    def keyify(row):
        return (
            str(row.get("Senast modifierad", "")),
            str(row.get("Bolag", "")),
            str(row.get("Typ av sajt", "")),
            str(row.get("Länk", "")),
        )

    existing = set(db.apply(keyify, axis=1).tolist()) if not db.empty else set()
    new_rows = []

    for _, r in src.iterrows():
        company = str(r["Bolag"]).strip()
        site_type = str(r["Typ av sajt"]).strip()
        sm_url = str(r["Länk"]).strip()
        if not sm_url.startswith(("http://", "https://")):
            print(f"Ogiltig sitemap-URL: {sm_url}")
            continue
        try:
            text, _ = fetch(sm_url)
        except Exception as e:
            print(f"Kunde inte läsa sitemap: {sm_url} ({e})")
            continue

        entries = parse(text)
        for url, date in entries:
            d = norm_date(date) if date else None
            if not d:
                continue
            tup = (d, company, site_type, url)
            if tup not in existing:
                existing.add(tup)
                new_rows.append({
                    "Senast modifierad": d,
                    "Bolag": company,
                    "Typ av sajt": site_type,
                    "Länk": url,
                    "Sitemap": sm_url,
                })

    # Lägg till och sortera
    if new_rows:
        db = pd.concat([db, pd.DataFrame(new_rows)], ignore_index=True)

    if not db.empty:
        db["Senast modifierad"] = pd.to_datetime(db["Senast modifierad"], errors="coerce")
        db = db.sort_values("Senast modifierad", ascending=False)
        db["Senast modifierad"] = db["Senast modifierad"].dt.strftime("%Y-%m-%d")

    # Säkra mapparna finns
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="w") as xw:
        db.to_excel(xw, index=False, sheet_name=SHEET)

    # Skriv endast nya rader i terminalen (som i ditt exempel)
    if new_rows:
        for row in sorted(new_rows, key=lambda x: x["Senast modifierad"], reverse=True):
            print(f"{row['Senast modifierad']} {row['Bolag']}: {row['Länk']}")
    else:
        print("Inga nya ändringar funna denna körning.")

if __name__ == "__main__":
    main()
