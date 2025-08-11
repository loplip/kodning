#!/usr/bin/env python3
import sys
import subprocess
from pathlib import Path

DEFAULT_JOB = "adtraction_stats.py"

def main():
    # Om inget argument: kör default-jobbet i roten
    if len(sys.argv) < 2:
        job_file = Path(DEFAULT_JOB)
    else:
        name = sys.argv[1]
        job_file = Path(f"{name}.py") if not name.endswith(".py") else Path(name)

    if not job_file.exists():
        print(f"Fel: {job_file} finns inte i roten.")
        sys.exit(1)

    try:
        subprocess.run([sys.executable, str(job_file)], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Scriptet misslyckades med kod {e.returncode}.")
        sys.exit(e.returncode)

if __name__ == "__main__":
    main()
