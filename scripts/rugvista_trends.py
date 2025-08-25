#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Trends månadsvis (2016-01-01 -> idag) för Rugvista.

Söker på: "rugvista"
Sparar till data_monthly.xlsx -> flik "rugvista_trends" (ersätter innehåll).
Alla värden heltal. Extra kolumn: % YoY rugvista
"""

from __future__ import annotations
import time
from pathlib import Path
from datetime import date
import numpy as np
import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

from pytrends.request import TrendReq

# ------------------ Konfiguration ------------------
TERM = "rugvista"

START_DATE = "2016-01-01"
END_DATE = date.today().strftime("%Y-%m-%d")
TIMEFRAME = f"{START_DATE} {END_DATE}"

HL = "sv-SE"
TZ = 0
GEO = ""   # worldwide
CAT = 0
GPROP = "" # web search

OUT_FILE = Path(__file__).with_name("data_monthly.xlsx")
SHEET_NAME = "rugvista_trends"

# ------------------ Hjälpfunktioner ------------------
def monthly_resample(df: pd.DataFrame) -> pd.DataFrame:
    return df.resample("MS").mean()

def yoy_percent(series: pd.Series) -> pd.Series:
    prev = series.shift(12)
    return pd.Series(
        np.where((prev.isna()) | (prev == 0), np.nan, (series - prev) / prev * 100.0),
        index=series.index, dtype="float64"
    )

def fetch_trends(term: str) -> pd.DataFrame:
    py = TrendReq(hl=HL, tz=TZ)
    for attempt in range(5):
        try:
            py.build_payload(kw_list=[term], timeframe=TIMEFRAME, geo=GEO, cat=CAT, gprop=GPROP)
            df = py.interest_over_time()
            if df.empty:
                raise RuntimeError("Tomt svar från Google Trends.")
            if "isPartial" in df.columns:
                df = df.drop(columns=["isPartial"])
            return monthly_resample(df)
        except Exception:
            if attempt == 4:
                raise
            time.sleep(2 ** attempt)

# ------------------ Main ------------------
def main():
    monthly = fetch_trends(TERM)

    # Heltal (ingen decimal)
    base = monthly[[TERM]].round(0).astype("Int64")

    # YoY % (avrundade till heltal)
    yoy = yoy_percent(monthly[TERM]).round(0).astype("Int64")
    yoy_df = pd.DataFrame({f"% YoY {TERM}": yoy}, index=monthly.index)

    df_out = pd.concat([base, yoy_df], axis=1)
    df_out.index = df_out.index.strftime("%Y-%m")
    df_out.insert(0, "Månad", df_out.index)
    df_out = df_out.reset_index(drop=True)

    # Skriv till Excel, skapa fil om den inte finns
    mode = "a" if OUT_FILE.exists() else "w"
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
        df_out.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print(f"Klart! Skrev {len(df_out)} rader till {OUT_FILE.name} ({SHEET_NAME}).")

if __name__ == "__main__":
    main()
