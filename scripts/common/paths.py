# scripts/common/paths.py
from pathlib import Path

# repo-roten (…/scripts/common -> …/scripts -> …/root)
ROOT = Path(__file__).resolve().parents[2]

DATA_DIR = ROOT / "data"
HISTORY_DIR = ROOT / "history"

DATA_DIR.mkdir(exist_ok=True)
HISTORY_DIR.mkdir(exist_ok=True)
