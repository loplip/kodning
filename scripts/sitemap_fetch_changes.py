#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sitemap_changes.py
--------------------------------
- Hårdkodar källfil: sources/sitemaps_bolag.xlsx
- Lagrar last_fetch_date (YYYY-MM-DD) per sitemap (dag-nivå)
- Processar endast URL:er vars sitemap-lastmod > last_fetch_date och >= 2023-01-01
- Öppnar/uppdaterar SQLite endast för hosts som faktiskt har kandidater
- Skriver/uppdaterar Excel: data/data_sitemap.xlsx, blad: 'databas_incl_changes'
  Kolumner: Senast modifierad, Bolag, Typ av sajt, Länk, Sitemap, Status, Ändringar
"""

from __future__ import annotations
import sys
import re
import difflib
import hashlib
import sqlite3
import argparse
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Tuple, Dict
from urllib.parse import urlparse

import requests
import pandas as pd
from xml.etree import ElementTree as ET

try:
    from bs4 import BeautifulSoup  # för textdiff
except Exception:
    BeautifulSoup = None  # fall tillbaka till rå HTML

# --- Repo-paths --------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.common.paths import DATA_DIR, SOURCES_DIR  # noqa: E402

INPUT_XLSX = SOURCES_DIR / "sitemaps_bolag.xlsx"         # hårdkodat
OUTPUT_XLSX = DATA_DIR / "data_sitemap.xlsx"
OUTPUT_SHEET = "databas_incl_changes"
SITES_DIR = SOURCES_DIR / "sites"                        # här läggs SQLite
SITES_DIR.mkdir(parents=True, exist_ok=True)

DATE_PAT = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")
UA = "Mozilla/5.0 (compatible; SitemapWatcher/6.1)"

# Global cutoff – URL:er äldre än detta tas inte med i SQLite alls
MIN_LASTMOD_DATE = "2023-01-01"

def utc_today_str() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")

def utc_now_str() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

# --- Hjälpare ----------------------------------------------------------------
def norm_date(s: Optional[str]) -> Optional[str]:
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

def fetch(url: str, timeout=15) -> requests.Response:
    r = requests.get(url, headers={"User-Agent": UA}, timeout=timeout, allow_redirects=True)
    r.raise_for_status()
    return r

def looks_like_xml(text: str) -> bool:
    t = text.lstrip().lower()
    return t.startswith("<?xml") or "<urlset" in t or "<sitemapindex" in t

def text_fingerprint(html: str) -> Tuple[str, str]:
    """Returnera (hash, text_for_diff). Använder bs4 om tillgänglig, annars rå html."""
    if BeautifulSoup is not None:
        try:
            soup = BeautifulSoup(html, "html.parser")
            for tag in soup(["script", "style", "noscript"]):
                tag.decompose()
            text = soup.get_text(separator="\n")
        except Exception:
            text = html
    else:
        text = html
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    normalized = "\n".join(lines)
    sha = hashlib.sha256(normalized.encode("utf-8", errors="ignore")).hexdigest()
    return sha, normalized

def compute_diff(old_text: str, new_text: str, max_lines: int = 80) -> str:
    diff_iter = difflib.unified_diff(
        old_text.splitlines(), new_text.splitlines(),
        fromfile="old", tofile="new", lineterm=""
    )
    diff_lines = list(diff_iter)
    if max_lines and len(diff_lines) > max_lines:
        head = diff_lines[:max_lines]
        head.append(f"... (trunkerad, totalt {len(diff_lines)} diff-rader)")
        return "\n".join(head)
    return "\n".join(diff_lines)

# --- Parsers -----------------------------------------------------------------
def parse_xml(xml_text: str) -> list[tuple[str, Optional[str]]]:
    """Returnerar list[(url, date)] och följer ev. <sitemapindex> rekursivt."""
    out: list[tuple[str, Optional[str]]] = []
    root = ET.fromstring(xml_text.encode("utf-8"))
    ns = {"ns": root.tag.split('}')[0].strip('{')} if root.tag.startswith("{") else {}

    # sitemapindex -> följ undersitemaps
    sitems = root.findall(".//ns:sitemap", ns) if ns else root.findall(".//sitemap")
    for sm in sitems:
        loc_el = sm.find("ns:loc", ns) if ns else sm.find("loc")
        if loc_el is not None and loc_el.text:
            loc = loc_el.text.strip()
            try:
                r = fetch(loc)
                out.extend(parse(r.text))
            except Exception:
                pass

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
            # fallback: sök datum i alla texter
            for el in u.iter():
                if el.text:
                    maybe = norm_date(el.text)
                    if maybe:
                        date_val = maybe
                        break
        if loc:
            out.append((loc, date_val))
    return out

def parse_table_or_text(text: str) -> list[tuple[str, Optional[str]]]:
    out: list[tuple[str, Optional[str]]] = []
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

def parse(text: str) -> list[tuple[str, Optional[str]]]:
    if looks_like_xml(text):
        try:
            return parse_xml(text)
        except Exception:
            return parse_table_or_text(text)
    return parse_table_or_text(text)

# --- DB (per host) -----------------------------------------------------------
def ensure_db(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(path)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pages (
            url TEXT PRIMARY KEY,
            last_hash TEXT,
            last_content TEXT,
            last_fetched TEXT,
            last_modified TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            url TEXT,
            fetched_at TEXT,
            status TEXT,
            diff_text TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sitemaps (
            sm_url TEXT PRIMARY KEY,
            last_fetch_date TEXT,     -- YYYY-MM-DD (dag-nivå)
            sm_hash TEXT,
            last_checked TEXT
        )
    """)
    # säkerställ ev. nya kolumner
    def ensure_col(tbl: str, col: str):
        cur.execute(f"PRAGMA table_info({tbl})")
        cols = {r[1] for r in cur.fetchall()}
        if col not in cols:
            cur.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} TEXT")
    for col in ("last_fetch_date", "sm_hash", "last_checked"):
        ensure_col("sitemaps", col)
    con.commit()
    return con

def db_get_sitemap(con: sqlite3.Connection, sm_url: str):
    cur = con.cursor()
    cur.execute("SELECT sm_url,last_fetch_date,sm_hash,last_checked FROM sitemaps WHERE sm_url=?", (sm_url,))
    return cur.fetchone()

def db_set_sitemap_fetch_date(con: sqlite3.Connection, sm_url: str, day: str, sm_hash: Optional[str]) -> None:
    cur = con.cursor()
    now = utc_now_str()
    cur.execute("""
        INSERT INTO sitemaps(sm_url,last_fetch_date,sm_hash,last_checked)
        VALUES(?,?,?,?)
        ON CONFLICT(sm_url) DO UPDATE SET
            last_fetch_date=excluded.last_fetch_date,
            sm_hash=COALESCE(excluded.sm_hash, sitemaps.sm_hash),
            last_checked=excluded.last_checked
    """, (sm_url, day, sm_hash, now))
    con.commit()

def db_get_page(con: sqlite3.Connection, url: str):
    cur = con.cursor()
    cur.execute("SELECT url,last_hash,last_content,last_fetched,last_modified FROM pages WHERE url=?", (url,))
    return cur.fetchone()

def db_upsert_page(con: sqlite3.Connection, url: str, h: str, content: str, modified: Optional[str]) -> None:
    cur = con.cursor()
    now = utc_now_str()
    cur.execute("""
        INSERT INTO pages(url,last_hash,last_content,last_fetched,last_modified)
        VALUES(?,?,?,?,?)
        ON CONFLICT(url) DO UPDATE SET
            last_hash=excluded.last_hash,
            last_content=excluded.last_content,
            last_fetched=excluded.last_fetched,
            last_modified=COALESCE(excluded.last_modified, pages.last_modified)
    """, (url, h, content, now, modified))
    con.commit()

def db_add_change(con: sqlite3.Connection, url: str, status: str, diff_text: str) -> None:
    cur = con.cursor()
    cur.execute("INSERT INTO changes(url,fetched_at,status,diff_text) VALUES(?,?,?,?)", (url, utc_now_str(), status, diff_text))
    con.commit()

# Enkel connection-cache per host (öppna endast när vi måste)
_conns: Dict[str, sqlite3.Connection] = {}
def conn_for_host(host: str) -> sqlite3.Connection:
    if host not in _conns:
        _conns[host] = ensure_db(SITES_DIR / f"{host.replace(':','_')}.sqlite")
    return _conns[host]

# --- Main --------------------------------------------------------------------
def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--show-progress", dest="show_progress", action="store_true", help="Visa progress prints")
    ap.add_argument("--no-show-progress", dest="show_progress", action="store_false", help="Dölj progress prints")
    ap.set_defaults(show_progress=True)
    args = ap.parse_args()
    SHOW = args.show_progress

    if not INPUT_XLSX.exists():
        raise SystemExit(f"Hittar inte {INPUT_XLSX}")

    if SHOW:
        print(f"[INIT] Läser källor från {INPUT_XLSX}")

    # Läs källor
    src = pd.read_excel(INPUT_XLSX)
    src = src.rename(columns={c: c.strip() for c in src.columns})
    required = ["Bolag", "Typ av sajt", "Länk"]
    for r in required:
        if r not in src.columns:
            raise SystemExit(f"Saknar kolumn '{r}' i {INPUT_XLSX}")

    # Läs befintlig output (för historik)
    if OUTPUT_XLSX.exists():
        try:
            db_out = pd.read_excel(OUTPUT_XLSX, sheet_name=OUTPUT_SHEET)
        except Exception:
            db_out = pd.DataFrame(columns=["Senast modifierad","Bolag","Typ av sajt","Länk","Sitemap","Status","Ändringar"])
    else:
        db_out = pd.DataFrame(columns=["Senast modifierad","Bolag","Typ av sajt","Länk","Sitemap","Status","Ändringar"])

    def keyify(row):
        return (
            str(row.get("Senast modifierad","")),
            str(row.get("Bolag","")),
            str(row.get("Typ av sajt","")),
            str(row.get("Länk","")),
            str(row.get("Status","")),
        )
    existing = set(db_out.apply(keyify, axis=1).tolist()) if not db_out.empty else set()

    new_rows = []
    today = utc_today_str()

    for _, r in src.iterrows():
        company = str(r["Bolag"]).strip()
        site_type = str(r["Typ av sajt"]).strip()
        sm_url = str(r["Länk"]).strip()
        if not sm_url.startswith(("http://","https://")):
            if SHOW: print(f"[SKIP] Ogiltig sitemap-URL: {sm_url}")
            continue

        sm_host = urlparse(sm_url).netloc
        sm_con = conn_for_host(sm_host)
        sm_prev = db_get_sitemap(sm_con, sm_url)
        last_fetch_date = sm_prev[1] if sm_prev else None

        if last_fetch_date == today:
            if SHOW: print(f"[SKIP] {sm_url} redan fetchad idag ({today})")
            continue

        if SHOW: print(f"[SITEMAP] Hämtar {sm_url}")
        try:
            sm_resp = fetch(sm_url, timeout=15)
        except Exception as e:
            print(f"[FEL] Kunde inte läsa sitemap: {sm_url} ({e})")
            continue

        sm_text = sm_resp.text
        sm_hash = hashlib.sha256(sm_text.encode("utf-8", errors="ignore")).hexdigest()
        entries = parse(sm_text)
        total = len(entries)

        # Filtrera: kräver lastmod och ska vara >= MIN_LASTMOD_DATE och > last_fetch_date (om finns)
        candidates: list[tuple[str, str]] = []
        for page_url, d in entries:
            if not d:
                continue
            if d < MIN_LASTMOD_DATE:
                continue
            if (last_fetch_date is None) or (d > last_fetch_date):
                candidates.append((page_url, d))

        if SHOW:
            print(f"[FILTER] {sm_url}: {total} poster → kandidater efter datumfilter: {len(candidates)} "
                  f"(cutoff {MIN_LASTMOD_DATE}, efter {last_fetch_date or '—'})")

        if not candidates:
            db_set_sitemap_fetch_date(sm_con, sm_url, today, sm_hash)
            if SHOW: print(f"[DONE] Inga kandidater. Markerar {sm_url} som fetchad {today}.")
            continue

        # Bearbeta endast kandidaterna
        per_host_counts: Dict[str,int] = {}
        for page_url, d in candidates:
            page_host = urlparse(page_url).netloc
            con = conn_for_host(page_host)
            per_host_counts[page_host] = per_host_counts.get(page_host, 0) + 1

            try:
                resp = fetch(page_url, timeout=15)
            except Exception as e:
                status = "Otillgänglig"
                change_note = f"Kunde inte hämta sida: {e}"
                row = {
                    "Senast modifierad": d,
                    "Bolag": company,
                    "Typ av sajt": site_type,
                    "Länk": page_url,
                    "Sitemap": sm_url,
                    "Status": status,
                    "Ändringar": change_note,
                }
                if keyify(row) not in existing:
                    new_rows.append(row)
                continue

            html = resp.text
            new_hash, new_text = text_fingerprint(html)
            prev = db_get_page(con, page_url)

            if prev is None:
                # Ny sida (men fortfarande efter cutoff)
                db_upsert_page(con, page_url, new_hash, new_text, d)
                db_add_change(con, page_url, "Ny", "")
                row = {
                    "Senast modifierad": d,
                    "Bolag": company,
                    "Typ av sajt": site_type,
                    "Länk": page_url,
                    "Sitemap": sm_url,
                    "Status": "Ny",
                    "Ändringar": "",
                }
                if keyify(row) not in existing:
                    new_rows.append(row)
            else:
                _, old_hash, old_text, _, _ = prev
                if old_hash != new_hash:
                    diff_text = compute_diff(old_text or "", new_text or "", max_lines=80)
                    db_upsert_page(con, page_url, new_hash, new_text, d)
                    db_add_change(con, page_url, "Modifierad", diff_text)
                    row = {
                        "Senast modifierad": d,
                        "Bolag": company,
                        "Typ av sajt": site_type,
                        "Länk": page_url,
                        "Sitemap": sm_url,
                        "Status": "Modifierad",
                        "Ändringar": diff_text,
                    }
                    if keyify(row) not in existing:
                        new_rows.append(row)

        # Markera sitemapen som fetchad idag
        db_set_sitemap_fetch_date(sm_con, sm_url, today, sm_hash)
        if SHOW:
            tot = sum(per_host_counts.values())
            details = ", ".join(f"{h}:{n}" for h,n in per_host_counts.items())
            print(f"[DONE] {sm_url} — kandidater: {tot} ({details})")

    # Skriv Excel endast om något nytt kom in
    if new_rows:
        df_append = pd.DataFrame(new_rows)

        # sortera och normalisera datumformat
        df_append["Senast modifierad"] = pd.to_datetime(df_append["Senast modifierad"], errors="coerce")
        df_append = df_append.sort_values("Senast modifierad", ascending=False)
        df_append["Senast modifierad"] = df_append["Senast modifierad"].dt.strftime("%Y-%m-%d")

        OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

        if OUTPUT_XLSX.exists():
            # APPEND till befintligt blad
            try:
                # Pandas >= 1.4: använd if_sheet_exists="overlay"
                with pd.ExcelWriter(
                    OUTPUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="overlay"
                ) as xw:
                    try:
                        ws = xw.book[OUTPUT_SHEET]
                        startrow = ws.max_row if ws.max_row else 0
                        header = False if startrow and startrow > 1 else True
                    except KeyError:
                        # Bladet finns inte -> skapa från rad 0 med header
                        startrow, header = 0, True

                    df_append.to_excel(
                        xw, index=False, sheet_name=OUTPUT_SHEET, startrow=startrow, header=header
                    )

            except TypeError:
                # Äldre Pandas utan if_sheet_exists -> fallback med openpyxl
                import openpyxl
                from openpyxl.utils.dataframe import dataframe_to_rows

                wb = openpyxl.load_workbook(OUTPUT_XLSX)
                ws = wb[OUTPUT_SHEET] if OUTPUT_SHEET in wb.sheetnames else wb.create_sheet(OUTPUT_SHEET)

                write_header = ws.max_row <= 1  # skriv header om bladet i princip är tomt
                for i, row in enumerate(dataframe_to_rows(df_append, index=False, header=write_header)):
                    # openpyxl returnerar tomma rader ibland – hoppa över dem
                    if row is None or (isinstance(row, list) and all(c is None for c in row)):
                        continue
                    ws.append(row)
                wb.save(OUTPUT_XLSX)

        else:
            # Filen finns inte – skapa ny och skriv hela df_append
            with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="w") as xw:
                df_append.to_excel(xw, index=False, sheet_name=OUTPUT_SHEET)

        # Liten konsolrapport
        for row in sorted(new_rows, key=lambda x: x["Senast modifierad"], reverse=True):
            print(f"{row['Senast modifierad']} {row['Status']} {row['Bolag']}: {row['Länk']}")
    else:
        print("Inga nya ändringar funna – skippade Excel-skrivning.")

if __name__ == "__main__":
    main()
