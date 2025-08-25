#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations
import time, sys
from datetime import date

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from scripts.common.paths import DATA_DIR


import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

from pytrends.request import TrendReq

# ---------------------- Konfiguration ----------------------
START_DATE = "2016-01-01"
END_DATE = date.today().strftime("%Y-%m-%d")
TIMEFRAME = f"{START_DATE} {END_DATE}"

HL = "sv-SE"
TZ = 0     # UTC
CAT = 0
GPROP = ""  # web search

OUT_FILE = DATA_DIR / "data_monthly.xlsx"
SHEET_NAME = "soder_trends"

# ---------------------- Hjälpfunktioner ----------------------

def monthly_resample(df: pd.DataFrame) -> pd.DataFrame:
    return df.resample("MS").mean()

def build_payload(pytrends: TrendReq, kw_list, timeframe, geo):
    pytrends.build_payload(kw_list=kw_list, timeframe=timeframe, geo=geo, cat=CAT, gprop=GPROP)
    time.sleep(1.0)

def fetch_series(pytrends: TrendReq, term: str, geo: str) -> pd.Series:
    for attempt in range(5):
        try:
            build_payload(pytrends, [term], TIMEFRAME, geo)
            df = pytrends.interest_over_time()
            if df.empty:
                raise RuntimeError("Tomt svar från Google Trends.")
            if "isPartial" in df.columns:
                df = df.drop(columns=["isPartial"])
            return monthly_resample(df)[term]
        except Exception:
            if attempt == 4:
                raise
            time.sleep(2 ** attempt)

def yoy_percent(series: pd.Series) -> pd.Series:
    prev = series.shift(12)
    yoy = (series - prev) / prev * 100.0
    return yoy.where(~(prev.isna() | (prev == 0)))

# ---------------------- Huvudflöde ----------------------

def main() -> None:
    pytrends = TrendReq(hl=HL, tz=TZ)

    se_term = "sportfiskeprylar"
    ww_term = "sportfishtackle"

    se_series = fetch_series(pytrends, se_term, geo="SE")
    ww_series = fetch_series(pytrends, ww_term, geo="")

    se_vals = se_series.round(0).astype("Int64")
    ww_vals = ww_series.round(0).astype("Int64")

    se_yoy = yoy_percent(se_series).round(0).astype("Int64")
    ww_yoy = yoy_percent(ww_series).round(0).astype("Int64")

    df = pd.concat({
        se_term: se_vals,
        f"% YoY {se_term}": se_yoy,
        ww_term: ww_vals,
        f"% YoY {ww_term}": ww_yoy,
    }, axis=1)

    df = df.sort_index()
    df.index = df.index.strftime("%Y-%m")
    df.insert(0, "Månad", df.index)
    df = df.reset_index(drop=True)

    if OUT_FILE.exists():
        mode = "a"
        with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode=mode, if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    else:
        with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    print(f"Klart! Skrev {len(df)} rader och {len(df.columns)-1} kolumner till {OUT_FILE.name} ({SHEET_NAME}).")

if __name__ == "__main__":
    main()
