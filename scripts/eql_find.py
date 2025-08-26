import io
import os
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Optional, Dict, List

import pandas as pd
import numpy as np
import requests
import warnings

# ---- lägg till repo-rot och data-sökväg ----
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from scripts.common.paths import DATA_DIR  # kräver att denna finns i ditt repo
OUT_PATH = DATA_DIR / "eql_data.xlsx"

# Tysta harmlös openpyxl-varning
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet"
)

SHEET_NAME = "Preparat"
KEY_COLS = ['Country', 'Product Name', 'Strength', 'Active Substances']

# --------------------------
# Hjälpare
# --------------------------
def _coerce_date(series: pd.Series) -> pd.Series:
    """Försök konvertera olika datumformat (inkl. Excel-serie) till date-objekt."""
    s = pd.to_datetime(series, errors="coerce", format="%Y-%m-%d")  # ISO
    mask = s.isna()
    if mask.any():
        s.loc[mask] = pd.to_datetime(series[mask], errors="coerce", dayfirst=True)
    num = pd.to_numeric(series, errors="coerce")
    mask = s.isna() & num.notna()
    if mask.any():
        s.loc[mask] = pd.to_datetime(num[mask], origin="1899-12-30", unit="D", errors="coerce")
    return s.dt.date

def _finalize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Standardisera kolumner, sortering och tomma celler."""
    # Säkerställ kolumnordning
    cols = ['Country', 'Product Name', 'Strength', 'Active Substances',
            'Approval Date', 'Marketing Holder', 'Distributor']
    df = df.reindex(columns=cols)

    # Sortera alltid senaste Approval Date överst
    dt = pd.to_datetime(df['Approval Date'], errors='coerce')
    df = df.assign(_sort_dt=dt).sort_values(
        by=['_sort_dt', 'Country', 'Product Name'],
        ascending=[False, True, True]
    ).drop(columns=['_sort_dt'])

    # Format: YYYY-MM-DD som text
    df['Approval Date'] = pd.to_datetime(df['Approval Date'], errors='coerce').dt.strftime('%Y-%m-%d')

    # Tomt istället för "nan" (NaN -> None gör att Excel-celler lämnas blanka)
    df = df.replace({np.nan: None, pd.NA: None, 'nan': None})

    return df.reset_index(drop=True)

# --------- Hämt- och parselogik (oförändrad i sak) ----------
# Sverige
def fetch_sweden_eql() -> pd.DataFrame:
    url = "https://www.lakemedelsverket.se/globalassets/excel/lakemedelsprodukter.xlsx"
    headers = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/112.0.0.0 Safari/537.36'
        )
    }
    buffer: Optional[io.BytesIO] = None
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        resp.raise_for_status()
        buffer = io.BytesIO(resp.content)
    except Exception:
        for fname in ('Lakemedelsprodukter.xlsx', 'lakemedelsprodukter.xlsx'):
            if os.path.exists(fname):
                with open(fname, 'rb') as f:
                    buffer = io.BytesIO(f.read())
                break
        if buffer is None:
            raise RuntimeError("Ladda ner 'Lakemedelsprodukter.xlsx' manuellt och placera i arbetskatalogen.")

    xls = pd.ExcelFile(buffer, engine="openpyxl")
    df = xls.parse(xls.sheet_names[0])

    for col in ('Innehavare', 'Ombud', 'Namn'):
        if col in df.columns:
            df[col] = df[col].astype(str)

    has_eql_holder = df['Innehavare'].str.contains(r'\bEQL\b', case=False, na=False)
    has_eql_agent  = df.get('Ombud', pd.Series(False, index=df.index)).str.contains(r'\bEQL\b', case=False, na=False)
    mask = has_eql_holder | has_eql_agent

    want_cols = ['Namn', 'Styrka', 'Aktiv substans', 'Godkännande-datum', 'Innehavare']
    if 'Ombud' in df.columns:
        want_cols.append('Ombud')

    subset = df.loc[mask, want_cols].copy()
    subset = subset.rename(columns={
        'Namn': 'Product Name',
        'Styrka': 'Strength',
        'Aktiv substans': 'Active Substances',
        'Godkännande-datum': 'Approval Date',
        'Innehavare': 'Marketing Holder',
    })
    if 'Ombud' in subset.columns:
        subset['Distributor'] = subset['Ombud'].replace('', pd.NA)
        subset = subset.drop(columns=['Ombud'])
    else:
        subset['Distributor'] = None

    subset['Approval Date'] = _coerce_date(subset['Approval Date'])
    subset['Country'] = 'Sweden'
    return subset[['Country', 'Product Name', 'Strength', 'Active Substances',
                   'Approval Date', 'Marketing Holder', 'Distributor']].reset_index(drop=True)

# Danmark
def fetch_denmark_eql() -> pd.DataFrame:
    url = "https://laegemiddelstyrelsen.dk/ftp-upload/ListeOverGodkendteLaegemidler.xlsx"
    buffer: Optional[io.BytesIO] = None
    headers = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/112.0.0.0 Safari/537.36'
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=60)
        resp.raise_for_status()
        buffer = io.BytesIO(resp.content)
    except Exception:
        for fname in ('ListeOverGodkendteLaegemidler.xlsx', 'listeovergodkentedlaegemidler.xlsx'):
            if os.path.exists(fname):
                with open(fname, 'rb') as f:
                    buffer = io.BytesIO(f.read())
                break
        if buffer is None:
            raise RuntimeError("Ladda ner 'ListeOverGodkendteLaegemidler.xlsx' manuellt och placera i arbetskatalogen.")

    xls = pd.ExcelFile(buffer, engine="openpyxl")
    df = xls.parse('Godkendte Lægemidler')

    df['MftIndehaver'] = df['MftIndehaver'].astype(str)
    df['Navn'] = df['Navn'].astype(str)

    mask = df['MftIndehaver'].str.contains(r'\bEQL\b', case=False, na=False)

    strength_col = 'Styrketekst' if 'Styrketekst' in df.columns else ('Styrke' if 'Styrke' in df.columns else None)

    want_cols = ['Navn', 'AktiveSubstanser', 'Registreringsdato', 'MftIndehaver']
    if strength_col:
        want_cols.insert(1, strength_col)

    subset = df.loc[mask, want_cols].copy()

    rename_map = {
        'Navn': 'Product Name',
        'AktiveSubstanser': 'Active Substances',
        'Registreringsdato': 'Approval Date',
        'MftIndehaver': 'Marketing Holder',
    }
    if strength_col:
        rename_map[strength_col] = 'Strength'

    subset = subset.rename(columns=rename_map)

    subset['Approval Date'] = _coerce_date(subset['Approval Date'])
    subset['Country'] = 'Denmark'
    subset['Distributor'] = None

    return subset[['Country', 'Product Name', 'Strength', 'Active Substances',
                   'Approval Date', 'Marketing Holder', 'Distributor']].reset_index(drop=True)

# Finland
def _parse_fimea_xml_from_bytes(data: bytes) -> pd.DataFrame:
    product_substances: Dict[str, set[str]] = defaultdict(set)
    for _, elem in ET.iterparse(io.BytesIO(data), events=('end',)):
        if elem.tag == 'Pakkaus':
            prod_ref = elem.attrib.get('Laakevalmiste-ref')
            if prod_ref:
                for pkg_sub in elem.findall('Pakkaus_Laakeaine'):
                    sub_ref = pkg_sub.attrib.get('Laakeaine-ref')
                    if sub_ref:
                        product_substances[prod_ref].add(sub_ref)
            elem.clear()

    substances: Dict[str, str] = {}
    for _, elem in ET.iterparse(io.BytesIO(data), events=('end',)):
        if elem.tag == 'Laakeaine':
            sid = elem.attrib.get('id')
            va = elem.find('VaikuttavaAine')
            name: Optional[str] = None
            if va is not None:
                for aine in va.findall('Aine'):
                    value = aine.attrib.get('value')
                    if value:
                        name = value.strip()
                        break
            if sid and name:
                substances[sid] = name
            elem.clear()

    records: List[Dict[str, Optional[str]]] = []
    for _, elem in ET.iterparse(io.BytesIO(data), events=('end',)):
        if elem.tag == 'Laakevalmiste':
            prod_id = elem.attrib.get('id')
            name_elem = elem.find('Kauppanimi')
            prod_name = name_elem.text.strip() if (name_elem is not None and name_elem.text) else None
            strength_elem = elem.find('Vahvuus')
            strength = strength_elem.text.strip() if (strength_elem is not None and strength_elem.text) else None
            holders: List[str] = []
            dates_raw: List[str] = []
            for ml_elem in elem.findall('.//Myyntilupa'):
                h_txt = (ml_elem.findtext('Haltija') or '').strip()
                if h_txt:
                    holders.append(h_txt)
                d_txt = (ml_elem.findtext('Myontamispaiva') or '').strip()
                if d_txt:
                    dates_raw.append(d_txt)
            approval_date: Optional[str] = None
            if dates_raw:
                dts = pd.to_datetime(pd.Series(dates_raw), errors='coerce')
                if dts.notna().any():
                    approval_date = dts.min().date().isoformat()
            distributors: List[str] = []
            for dist_elem in elem.findall('.//Jakelija'):
                if dist_elem.text:
                    distributors.append(dist_elem.text.strip())

            def is_eql(x: Optional[str]) -> bool:
                return bool(x and 'EQL' in x.upper())

            if any(is_eql(h) for h in holders + distributors):
                sub_ids = product_substances.get(prod_id, set())
                sub_names = [substances.get(sid) for sid in sub_ids if substances.get(sid)]
                sub_names_uniq = sorted(set(sub_names)) if sub_names else []
                records.append({
                    'Country': 'Finland',
                    'Product Name': prod_name,
                    'Strength': strength,
                    'Active Substances': ', '.join(sub_names_uniq) if sub_names_uniq else None,
                    'Approval Date': approval_date,
                    'Marketing Holder': '; '.join(holders) if holders else None,
                    'Distributor': '; '.join(distributors) if distributors else None
                })
            elem.clear()
    return pd.DataFrame(records)

def fetch_finland_eql() -> pd.DataFrame:
    url = 'https://data.pilvi.fimea.fi/avoin-data/Perusrekisteri.xml'
    r = requests.get(url, timeout=120)
    r.raise_for_status()
    df = _parse_fimea_xml_from_bytes(r.content)
    if not df.empty:
        df['Approval Date'] = _coerce_date(df['Approval Date'])
    return df

# --------------------------
# Huvudflöde (utan konkurrenträkning) + skriv alltid sorterat
# --------------------------
def _collect_current() -> pd.DataFrame:
    sweden_eql = fetch_sweden_eql()
    denmark_eql = fetch_denmark_eql()
    finland_eql = fetch_finland_eql()
    combined = pd.concat([sweden_eql, denmark_eql, finland_eql], ignore_index=True)
    return _finalize_df(combined)

def _dedupe_on_key(df: pd.DataFrame) -> pd.DataFrame:
    k = df[KEY_COLS].astype(str).agg('||'.join, axis=1)
    df = df.assign(_k=k).drop_duplicates('_k').drop(columns=['_k'])
    return df

def _read_existing(path: Path) -> Optional[pd.DataFrame]:
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=str)
    except Exception:
        return None
    # Konvertera datum och städa tomma celler
    if 'Approval Date' in df.columns:
        df['Approval Date'] = _coerce_date(df['Approval Date'])
    # Tomma strängar -> NaN (så att finalize gör blanka)
    df = df.replace({'': np.nan, 'nan': np.nan})
    return df

def _write_sorted(path: Path, df: pd.DataFrame) -> None:
    """Skriv HELA bladet sorterat. Pandas hanterar NaN -> tom cell i Excel."""
    df = _finalize_df(df)
    with pd.ExcelWriter(path, engine='openpyxl', mode='a' if path.exists() else 'w') as w:
        # Om bladet finns, ta bort det först så att vi ersätter med sorterad version
        if path.exists():
            try:
                book = w.book
                if SHEET_NAME in book.sheetnames:
                    std = book[SHEET_NAME]
                    book.remove(std)
                    book.save(path)
            except Exception:
                pass
        df.to_excel(w, index=False, sheet_name=SHEET_NAME)

def main(output_path: Optional[Path] = None) -> None:
    out_path = Path(output_path) if output_path else OUT_PATH
    out_path.parent.mkdir(parents=True, exist_ok=True)

    current = _collect_current()
    existing = _read_existing(out_path)

    if existing is None or existing.empty:
        merged = current
        new_rows = current
    else:
        merged = pd.concat([existing, current], ignore_index=True)
        merged = _dedupe_on_key(merged)
        # Identifiera vilka som är nya (finns i current men saknas i existing)
        existing_keys = existing[KEY_COLS].astype(str).agg('||'.join, axis=1)
        current_keys = current[KEY_COLS].astype(str).agg('||'.join, axis=1)
        new_mask = ~current_keys.isin(existing_keys)
        new_rows = current.loc[new_mask]

    # Skriv alltid HELA bladet sorterat (ny + gammal data)
    _write_sorted(out_path, merged)

    if new_rows.empty:
        print("Inga nya preparat.")
        return

    for _, r in new_rows.iterrows():
        name = r.get('Product Name', '')
        country = r.get('Country', '')
        print(f"EQL: {name} har lagts till för {country}.")

if __name__ == '__main__':
    main()
