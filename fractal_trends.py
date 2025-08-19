#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Trends månadsvis (2016-01-01 -> idag) i batchar (4 termer + referens).
Normaliserar batchar via referensen.

Sparar till data_monthly.xlsx -> flik "Fractal_trends" (ersätter innehåll).
Excel heltal. Extra kolumner (SUMMA + YoY %):
Total | Total Chassin | Total Övrigt | % Total | % Chassin | % Övrigt

Skriver även ut de tre senaste månaderna för Fractal North i terminalen.
"""

from __future__ import annotations
import time
from pathlib import Path
from datetime import date
from typing import List

import numpy as np
import pandas as pd
pd.set_option('future.no_silent_downcasting', True)  # undvik FutureWarning i pytrends/pandas

from pytrends.request import TrendReq

# -------------------------- Konfiguration --------------------------

REFERENCE = "Fractal North"

ALL_TERMS = [
    "Fractal Mood",
    "Fractal Epoch",
    "Fractal Torrent",
    "Fractal Pop",
    "Fractal Ridge",
    "Fractal Terra",
    "Fractal Meshify",
    "Fractal Focus",
    "Fractal Vector",
    "Fractal Era",
    "Fractal Define",
    "Fractal Core",
    "Fractal Arc",
    "Fractal Node",
    "Fractal Refine",
    "Fractal Scape",
]

START_DATE = "2016-01-01"
END_DATE = date.today().strftime("%Y-%m-%d")
TIMEFRAME = f"{START_DATE} {END_DATE}"

HL = "sv-SE"
TZ = 0     # UTC
GEO = ""   # worldwide
CAT = 0
GPROP = "" # web search

# Spara alltid till samma fil bredvid scriptet
OUT_FILE = Path(__file__).with_name("data_monthly.xlsx")
SHEET_NAME = "Fractal_trends"

# -------------------------- Hjälpfunktioner --------------------------

def chunk_list(items: List[str], size: int) -> List[List[str]]:
    return [items[i:i+size] for i in range(0, len(items), size)]

def safe_mean(s: pd.Series) -> float:
    s = s.dropna()
    return float(s.mean()) if len(s) else 0.0

def normalize_to_reference(base_ref: pd.Series, other_ref: pd.Series) -> float:
    joined = pd.concat([base_ref.rename("base"), other_ref.rename("other")], axis=1, join="inner").dropna()
    if joined.empty:
        return 1.0
    base_mean = safe_mean(joined["base"])
    other_mean = safe_mean(joined["other"])
    return base_mean / other_mean if other_mean != 0 else 1.0

def monthly_resample(df: pd.DataFrame) -> pd.DataFrame:
    # veckodata -> månadsmedel (månadens start)
    return df.resample("MS").mean()

def build_payload(pytrends: TrendReq, kw_list: List[str]):
    pytrends.build_payload(kw_list=kw_list, timeframe=TIMEFRAME, geo=GEO, cat=CAT, gprop=GPROP)
    time.sleep(1.0)  # enkel backoff

def fetch_group(pytrends: TrendReq, kw_list: List[str]) -> pd.DataFrame:
    for attempt in range(5):
        try:
            build_payload(pytrends, kw_list)
            df = pytrends.interest_over_time()
            if df.empty:
                raise RuntimeError("Tomt svar från Google Trends.")
            if "isPartial" in df.columns:
                df = df.drop(columns=["isPartial"])
            return monthly_resample(df)
        except Exception:
            if attempt == 4:
                raise
            time.sleep(2 ** attempt)

def remove_duplicate_column(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    seen, cols = set(), []
    for c in df.columns:
        cname = str(c)
        if cname == col_name:
            if cname in seen:
                continue
            seen.add(cname)
        cols.append(c)
    return df.loc[:, cols]

def yoy_percent(series: pd.Series) -> pd.Series:
    """% YoY, NaN om jämförelsemånad saknas eller = 0."""
    prev = series.shift(12)
    return pd.Series(
        np.where((prev.isna()) | (prev == 0), np.nan, (series - prev) / prev * 100.0),
        index=series.index, dtype="float64"
    )

# -------------------------- Huvudflöde --------------------------

def main():
    # Exkludera referensen från batchningen (lägger till den i varje request)
    keywords = [t for t in ALL_TERMS if t != REFERENCE]
    batches = chunk_list(keywords, 4)  # 4 + ref = 5 per request

    pytrends = TrendReq(hl=HL, tz=TZ)

    combined = None
    base_ref_series = None

    for i, batch in enumerate(batches, start=1):
        kw_list = [REFERENCE] + batch
        print(f"Hämtar batch {i}: {kw_list}")
        gdf = fetch_group(pytrends, kw_list)

        if combined is None:
            combined = gdf.copy()
            base_ref_series = combined[REFERENCE].copy()
        else:
            factor = normalize_to_reference(base_ref_series, gdf[REFERENCE])
            scaled = gdf.copy()
            for col in scaled.columns:
                if col != REFERENCE:
                    scaled[col] = scaled[col] * factor
            combined = combined.join(scaled.drop(columns=[REFERENCE]), how="outer")

    if combined is None or combined.empty:
        raise SystemExit("Ingen data hämtades.")

    # ta bort ev. dublett av referenskolumn
    combined = remove_duplicate_column(combined, REFERENCE)

    # Kolumnordning: referensen först, sedan ALL_TERMS i given ordning (de som finns)
    ordered_cols = [REFERENCE] + [t for t in ALL_TERMS if t in combined.columns]
    combined = combined[ordered_cols].sort_index()

    # -------- Extra kolumner (SUMMA + YoY %) --------
    # Total = summa alla kolumner (inkl. referensen)
    total_all = combined.sum(axis=1)

    # Total Övrigt = ENDAST Refine + Scape
    ovr_cols = [c for c in ["Fractal Refine", "Fractal Scape"] if c in combined.columns]
    total_ovrigt = combined[ovr_cols].sum(axis=1) if len(ovr_cols) >= 1 else pd.Series(index=combined.index, dtype=float)

    # Total Chassin = ALLA övriga (exkl. Refine & Scape)
    chassin_cols = [c for c in combined.columns if c not in ["Fractal Refine", "Fractal Scape"]]
    total_chassin = combined[chassin_cols].sum(axis=1) if chassin_cols else pd.Series(index=combined.index, dtype=float)

    # YoY på respektive summa (%)
    yoy_total    = yoy_percent(total_all).replace([np.inf, -np.inf], np.nan)
    yoy_chassin  = yoy_percent(total_chassin).replace([np.inf, -np.inf], np.nan)
    yoy_ovrigt   = yoy_percent(total_ovrigt).replace([np.inf, -np.inf], np.nan)

    # -------- Heltal utan decimaler --------
    combined_int = combined.round(0).astype("Int64")
    extras = pd.DataFrame({
        "Total":          total_all.round(0).astype("Int64"),
        "Total Chassin":  total_chassin.round(0).astype("Int64"),
        "Total Övrigt":   total_ovrigt.round(0).astype("Int64"),
        "% Total":        yoy_total.round(0).astype("Int64"),
        "% Chassin":      yoy_chassin.round(0).astype("Int64"),
        "% Övrigt":       yoy_ovrigt.round(0).astype("Int64"),
    }, index=combined.index)

    # Bygg utdata: Månad (YYYY-MM), sedan värden
    df_out = pd.concat([combined_int, extras], axis=1)
    df_out.index = df_out.index.strftime("%Y-%m")
    df_out.insert(0, "Månad", df_out.index)
    df_out = df_out.reset_index(drop=True)

    # -------- Spara ENDAST till samma xlsx & samma flik --------
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)  # om katalog saknas
    writer_kwargs = dict(engine="openpyxl")
    if OUT_FILE.exists():
        writer_kwargs.update(mode="a", if_sheet_exists="replace")
    else:
        writer_kwargs.update(mode="w")  # if_sheet_exists får EJ sättas här

    with pd.ExcelWriter(OUT_FILE, **writer_kwargs) as writer:
        df_out.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    # Skriv ut de tre senaste månaderna för referensen
    if REFERENCE in combined_int.columns:
        last3 = combined_int.tail(3)[[REFERENCE]].copy()
        last3.index = last3.index.strftime("%Y-%m")
        print(f"\nSenaste 3 månaderna – {REFERENCE}:")
        for idx, val in last3[REFERENCE].items():
            print(f"{idx}: {val}")
    else:
        print(f"Kolumnen '{REFERENCE}' hittades inte i resultatet.")

    print(f"\nKlart! Skrev {len(df_out)} rader och {len(df_out.columns)-1} kolumner till {OUT_FILE.name} ({SHEET_NAME}).")

if __name__ == "__main__":
    main()
