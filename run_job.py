#!/usr/bin/env python3
import sys
import subprocess
from pathlib import Path

def main():
    if len(sys.argv) < 2:
        print("Användning: python run_job.py <jobbnamn>")
        print("Exempel:   python run_job.py adtraction_stats")
        sys.exit(1)

    job_name = sys.argv[1]
    job_file = Path("jobs") / f"{job_name}.py"

    if not job_file.exists():
        print(f"Fel: {job_file} finns inte.")
        sys.exit(1)

    # Kör scriptet som separat process
    try:
        subprocess.run([sys.executable, str(job_file)], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Scriptet misslyckades med kod {e.returncode}.")
        sys.exit(e.returncode)

if __name__ == "__main__":
    main()
