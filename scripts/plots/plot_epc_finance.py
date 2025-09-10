#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Plotta alla flikars tidsserier i EN figur (subplots i rutnät).
- Automatisk filupptäckt: letar efter data_epc_finance.xlsx i DATA_DIR m.m.
- Titel per subplot = <fliknamn> – <värdekolumnens rubrik>
- Hanterar både ',' och '.' som decimal + tusentalsavgränsare
- SHOW_PROGRESS = True
- X-axel: YYYY-MM-DD (skalenlig)
- Headless fallback: sparar PNG till DATA_DIR/plots om ingen display finns.

Kör:  py plot_epc_finance.py   (från scripts/plots/)
"""

from __future__ import annotations
from datetime import datetime
from pathlib import Path
import sys, os, re, math
from typing import Tuple, Iterable, List

import pandas as pd

# --- Matplotlib för pop-up ---
import matplotlib
# Välj interaktiv backend om möjligt, annars Agg i headless-miljöer
if not (os.environ.get("DISPLAY") or sys.platform.startswith("win") or sys.platform == "darwin"):
    matplotlib.use("Agg")

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter

# Repo-vägar (DATA_DIR, HISTORY_DIR, SOURCES_DIR)
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
try:
    from scripts.common.paths import DATA_DIR  # type: ignore
except Exception:
    print("❌ Kunde inte importera DATA_DIR från scripts/common/paths.py", file=sys.stderr)
    raise

SHOW_PROGRESS: bool = True  # alltid på

def log(msg: str) -> None:
    if SHOW_PROGRESS:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        print(f"[{now}] {msg}")

# ---------- Numerik: ,/. decimal + tusental ----------
def _normalize_number_str(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().replace("\u00A0", " ").replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return ""
    has_comma = "," in s
    has_dot = "." in s
    if has_comma and has_dot:
        # sista separatorn blir decimal
        last_comma, last_dot = s.rfind(","), s.rfind(".")
        if last_comma > last_dot:
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        s = s.replace(",", ".")
    return s

def coerce_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.apply(_normalize_number_str), errors="coerce")

# ---------- Datum: robust tolkning ----------
def coerce_datetime(series: pd.Series) -> pd.Series:
    """Hanterar Excel-serietal, dolda tecken, och 'YYYY-MM-DD[ HH:MM]'."""
    s = series.copy()

    # Excel-serietal (tal > 40000 ≈ datum år 2009+)
    as_num = pd.to_numeric(s, errors="coerce")
    excel_mask = as_num.notna() & (as_num > 40000)

    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    if excel_mask.any():
        out.loc[excel_mask] = pd.to_datetime(
            as_num.loc[excel_mask],
            unit="d",
            origin="1899-12-30",
            errors="coerce",
        )

    txt_mask = ~excel_mask
    if txt_mask.any():
        t = s.loc[txt_mask].astype(str)
        t = (
            t.str.replace("\u00A0", " ", regex=False)  # NBSP
             .str.replace("\u200B", "", regex=False)   # ZWSP
             .str.replace("\u200C", "", regex=False)
             .str.replace("\u200D", "", regex=False)
             .str.replace("\ufeff", "", regex=False)   # BOM
             .str.strip()
        )
        # behåll rimliga datumtecken
        t = t.apply(lambda x: re.sub(r"[^0-9T:\- ]", "", x))
        p1 = pd.to_datetime(t, errors="coerce", utc=False, infer_datetime_format=True)
        need = p1.isna()
        if need.any():
            p2 = pd.to_datetime(t.loc[need].str.split().str[0], errors="coerce", utc=False)
            p1.loc[need] = p2
        out.loc[txt_mask] = p1

    return out

# ---------- Excel-filupptäckt ----------
CANDIDATE_NAMES = ["data_epc_finance.xlsx", "data_epc_alaal.xlsx", "data-epc-finance.xlsx"]

def candidate_paths() -> Iterable[Path]:
    for name in CANDIDATE_NAMES:
        yield Path(DATA_DIR) / name
    for p in Path(DATA_DIR).glob("*.xlsx"):
        n = p.name.lower()
        if "epc" in n and "finance" in n:
            yield p
    for name in CANDIDATE_NAMES:
        yield ROOT / name
    for p in ROOT.glob("*.xlsx"):
        n = p.name.lower()
        if "epc" in n and "finance" in n:
            yield p
    cwd = Path.cwd()
    for name in CANDIDATE_NAMES:
        yield cwd / name
    for p in cwd.glob("*.xlsx"):
        n = p.name.lower()
        if "epc" in n and "finance" in n:
            yield p

def find_excel_file() -> Path | None:
    for p in candidate_paths():
        if p.exists():
            return p
    return None

# ---------- Läs en flik -> (x, y, header) ----------
def read_sheet_to_series(df: pd.DataFrame) -> Tuple[pd.Series, pd.Series, str]:
    if df.shape[1] < 2:
        raise ValueError("Minst två kolumner krävs (datum + värde).")

    date_col = df.columns[0]
    value_col = df.columns[1]
    log(f"Läser kolumner: datum='{date_col}', värde='{value_col}'")

    x_raw = df[date_col]
    y_raw = df[value_col]

    x = coerce_datetime(x_raw)
    y = coerce_numeric(y_raw)

    bad_date = x.isna()
    bad_val = y.isna()
    if bad_date.any() or bad_val.any():
        log(f"⚠️ Ogiltiga datumrader: {int(bad_date.sum())}, ogiltiga värderader: {int(bad_val.sum())}")
        ex_idx = df.index[bad_date | bad_val][:5]
        for i in ex_idx:
            log(f"  - Rad {i+2}: datum='{x_raw.iloc[i]}', värde='{y_raw.iloc[i]}' (droppas)")

    mask = x.notna() & y.notna()
    x, y = x[mask], y[mask]

    order = x.argsort(kind="mergesort")
    x, y = x.iloc[order], y.iloc[order]

    return x, y, str(value_col)

# ---------- Formatterare ----------
def _space_thousands(x: float, pos=None) -> str:
    s = f"{x:,.12g}".replace(",", " ").replace("\xa0", " ")
    return s

# ---------- Rita alla subplots i en stor figur ----------
def plot_all_subplots(series_list: List[Tuple[str, str, pd.Series, pd.Series]]) -> None:
    """series_list: [(sheet_name, value_header, x_series, y_series), ...]"""
    n = len(series_list)
    if n == 0:
        log("Inget att plotta.")
        return

    # Välj layout: 1–3 kolumner beroende på antal
    if n == 1:
        ncols = 1
    elif n <= 4:
        ncols = 2
    else:
        ncols = 3
    nrows = math.ceil(n / ncols)

    # Rimlig storlek per subplot
    fig_w = 6 * ncols
    fig_h = 4 * nrows
    fig, axes = plt.subplots(nrows=nrows, ncols=ncols, figsize=(fig_w, fig_h), squeeze=False, constrained_layout=True)

    for idx, (sheet, header, x, y) in enumerate(series_list):
        r, c = divmod(idx, ncols)
        ax = axes[r][c]
        ax.plot(x, y, linewidth=2, marker="o", markersize=3)
        ax.set_title(f"{sheet} – {header}", fontsize=11, pad=6)
        ax.set_xlabel("Datum")
        ax.set_ylabel("Värde")
        ax.grid(True, alpha=0.25)

        # X-axel: YYYY-MM-DD
        ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=4, maxticks=10))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
        for label in ax.get_xticklabels():
            label.set_rotation(20)
            label.set_ha("right")

        # Y-axel format
        ax.yaxis.set_major_formatter(FuncFormatter(_space_thousands))

    # Dölj ev. tomma axlar om n inte fyller rutnätet
    for j in range(n, nrows * ncols):
        r, c = divmod(j, ncols)
        axes[r][c].set_visible(False)

    backend = matplotlib.get_backend().lower()
    headless = backend in ("agg", "cairoagg", "svg", "pdf", "ps")
    if headless:
        out_dir = Path(DATA_DIR) / "plots"
        out_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        out_path = out_dir / f"epc_finance_all_sheets__{timestamp}.png"
        fig.savefig(out_path, dpi=150)
        log(f"💾 Ingen display – sparade figur till: {out_path}")
        plt.close(fig)
    else:
        plt.show()

# ---------- Huvudflöde ----------
def main() -> int:
    xls_path = find_excel_file()
    if not xls_path:
        print("❌ Hittar ingen Excel-fil. Lägg 'data_epc_finance.xlsx' i DATA_DIR eller repo-rot.", file=sys.stderr)
        return 1

    log(f"Läser arbetsbok: {xls_path}")

    xfile = pd.ExcelFile(xls_path, engine="openpyxl")
    sheet_names = xfile.sheet_names
    if not sheet_names:
        print("❌ Inga flikar hittades i Excel-filen.", file=sys.stderr)
        return 1

    series_list: List[Tuple[str, str, pd.Series, pd.Series]] = []
    for sheet in sheet_names:
        log(f"Läser flik: {sheet}")
        df = xfile.parse(sheet_name=sheet, dtype=object)

        if df.empty or df.shape[1] < 2:
            print(f"⚠️ Fliken '{sheet}' saknar användbara kolumner – hoppar över.", file=sys.stderr)
            continue

        try:
            x, y, value_header = read_sheet_to_series(df)
        except Exception as e:
            print(f"⚠️ Fliken '{sheet}' kunde inte tolkas ({e}) – hoppar över.", file=sys.stderr)
            continue

        if len(x) == 0:
            print(f"⚠️ Fliken '{sheet}' saknar användbara datapunkter – hoppar över.", file=sys.stderr)
            continue

        log(f"✔️ Tar med {len(x)} punkter från '{sheet}'")
        series_list.append((sheet, value_header, x, y))

    plot_all_subplots(series_list)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
