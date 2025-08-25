# scripts/common/paths.py
from pathlib import Path

# repo-root = två nivåer upp från den här filen
ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "data"
HISTORY_DIR = ROOT / "history"

# se till att mapparna finns vid körning
DATA_DIR.mkdir(exist_ok=True)
HISTORY_DIR.mkdir(exist_ok=True)
