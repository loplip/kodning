# 📘 Instructions – Style Guide för Projektet

Detta dokument beskriver standarder och förutsättningar som gäller för projektet.  
Det används både som **manual för utvecklare** och som **instruktioner för AI-assistans**.  

---

## 📝 Allmänt
- Alla svar och dokumentation ska vara på **svenska**.  
- Fokus ligger på:
  - **Kodning** i Python och GitHub Actions.  
  - **Förslag på förbättringar** och smartare lösningar.  
  - **Strukturering** och bästa praxis.  
- Om något är oklart: **ställ frågor först** innan kod skrivs.  

---

## ⚙️ Kodningskrav

### 1. Excel-hantering
- Resultat ska alltid skrivas till en `.xlsx`-fil.  
- Om filen **inte finns**: skapa en ny.  
- Om filen **finns**:  
  - Om fliken inte finns: skapa ny flik.  
  - Om fliken finns: lägg till en **ny rad** med de nya datavärdena.  
- **Datumformat:**  
  - `"YYYY-MM-DD HH:MM"`.  
  - Tänk på att jag har svensk tidszon.
- **Talformat:**  
  - Mellanslag som tusentalsavgränsare, exempel:  
    - `1 300`  
    - `23 000 540`.  

### 2. Körning & progress
Alla skript ska ha en toggle för att styra utskrifter:  

```python
SHOW_PROGRESS = True  # eller False
```

- När `SHOW_PROGRESS = True`: skriv ut loggar/progress löpande.  
- När `SHOW_PROGRESS = False`: skriv endast en **sammanfattning** i slutet (t.ex. antal bearbetade rader, antal fel, antal sparade rader).  

### 3. GitHub Actions
- Koden ska kunna köras **både lokalt** och via **GitHub Actions**.  
- Workflow-filen heter `scrape.yml` och ligger i:  
  ```
  ./.github/workflows/scrape.yml
  ```  

---

## 📂 Repo-struktur
```bash
/Kodning
  ./data        # xlsx-filer som skripts exporterar
  ./history     
  ./scripts     # här ligger alla skripts
    ./__pycache__
    ./common
      ./__pycache__
      paths.py
  ./sources
    ./sites
```

---

## 🛤 Paths-hantering
Alla skript använder `./common/paths.py` för att hitta rätt kataloger.  
Standardimporten ser ut så här:  

```python
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.common.paths import DATA_DIR, HISTORY_DIR, SOURCES_DIR
```

- `DATA_DIR` → `./data`  
- `HISTORY_DIR` → `./history`  
- `SOURCES_DIR` → `./sources`  

---

## 📦 Övriga projektfiler
- **`requirements.txt`** finns i root och listar Python-beroenden.  
- **`scrape.yml`** används för GitHub Actions.  

---

## 🔑 Sammanfattning
- Skrivsvar och dokumentation på svenska.  
- Excel: skapa ny fil/flik/rad enligt regler ovan.  
- Datumformat: `YYYY-MM-DD HH:MM`. Tänk på att jag har svensk tidszon.
- Tal: tusentalsavgränsare med mellanslag.  
- Toggle för progress: `SHOW_PROGRESS = True/False`.  
- Körbart både lokalt och via GitHub Actions.  
- Repo-struktur och paths ska alltid följas.  
- Vid osäkerhet: **fråga först**.  

---
