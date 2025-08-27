# scripts/common/paths.py
from pathlib import Path

# repo-roten (…/scripts/common -> …/scripts -> …/root)
ROOT = Path(__file__).resolve().parents[2]

DATA_DIR = ROOT / "data"
HISTORY_DIR = ROOT / "history"
SOURCES_DIR = ROOT / "sources"

DATA_DIR.mkdir(exist_ok=True)
HISTORY_DIR.mkdir(exist_ok=True)
SOURCES_DIR.mkdir(exist_ok=True)
