#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kör månads-scripten för Google Trends:

- fractal_trends.py
- soder_trends.py
- rugvista_trends.py

Skriver ut status om körningen lyckades eller misslyckades.
"""

import subprocess
import sys
from pathlib import Path

SCRIPTS = [
    "fractal_trends.py",
    "soder_trends.py",
    "rugvista_trends.py",
]

def run_script(script: str) -> bool:
    path = Path(__file__).with_name(script)
    if not path.exists():
        print(f"❌ {script} saknas!")
        return False
    try:
        subprocess.run([sys.executable, str(path)], check=True)
        print(f"✅ {script} kördes klart.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {script} misslyckades (exit {e.returncode}).")
        return False

def main():
    print("=== Kör månads-scripten ===")
    all_ok = True
    for script in SCRIPTS:
        ok = run_script(script)
        all_ok = all_ok and ok
    print("===========================")
    if all_ok:
        print("✅ Alla scripts kördes utan fel!")
    else:
        print("⚠️ Något script misslyckades, se ovan.")

if __name__ == "__main__":
    main()
