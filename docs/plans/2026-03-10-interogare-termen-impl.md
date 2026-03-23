# Interogare Termen — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Aplicație separată `interogare_termen_processor.py` care procesează verificarea trimestrială legală (30.03 / 30.09) a tuturor beneficiarilor din Raport Stare față de S1–S6 DGITL, generând Suspendare, Repunere, Import Observatii și Datornici termen, cu memorie persistentă a contingentului activ.

**Architecture:** Modul nou `termen_logic.py` cu matricea de decizie; `contingent_tracker.py` pentru persistența JSON; `termen_processor.py` orchestrează pipeline-ul (refolosind `SectorLoader`, `Merger`, `RaportStariLoader`); UI în `interogare_termen_processor.py`. Modulele existente din Reinterogare Processor sunt refolosite nemodificate.

**Tech Stack:** Python 3.10+, pandas, openpyxl, CustomTkinter

**APP_DIR**: `C:\EXCELURI\Reinterogari\Reinterogare Processor\`

**Design doc**: `docs/plans/2026-03-10-interogare-termen-design.md`

---

## Referințe critice

**Coloane Raport Stare** (din `RaportStariLoader`):
- `Numar Cerere`, `Cnp Adult Handicap`, `Stare Cerere`, `Motiv`, `Data Inceput Stare`
- `Nume Adult cu Handicap`, `Prenume Adult cu Handicap`, `Data Cerere`

**Coloane sector după merge** (prefixate): `S1.AreNu / AreDatorii`, `S2.AreNu / AreDatorii` etc.
- Valori fără datorii: `"NU"`, `"Nu există"`, `"Nu exista"`, `""`, `NaN`
- Valori cu datorii: orice altceva (ex. `"DA"`, `"Are"`)

**Formate date output**:
- Data suspendare: prima zi a lunii DUPĂ termen (30.09.2025 → 01.10.2025)
- Data repunere: prima zi a lunii termenului (30.09.2025 → 01.09.2025)

**Motive**:
- Suspendare: `"Interogare termen {DD.MM.YYYY} / datorii la S1 S3"`
- Repunere: `"Interogare termen {DD.MM.YYYY} / Fara datorii"`
- Import Obs (cu datorii): `"Datorii interogare termen {DD.MM.YYYY} S2 S4"`
- Import Obs (fără): `"Fara datorii interogare termen {DD.MM.YYYY}"`

**Datornici termen** = Suspendare ∪ (Suspendat + cu datorii la termen)

---

## Task 1: `modules/termen_logic.py` — matricea de decizie

**Files:**
- Create: `APP_DIR/modules/termen_logic.py`
- Create: `APP_DIR/tests/test_termen_logic.py`

**Step 1: Scrie testele (6 scenarii din matricea de decizie)**

```python
# tests/test_termen_logic.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
import pytest
from datetime import date
from modules.termen_logic import classify_rows, build_suspendare, build_repunere, build_import_obs, build_datornici

TERM_DATE = date(2025, 9, 30)
SECTORS = ["S1", "S2"]

def _row(stare, motiv, s1="NU", s2="NU"):
    """Construiește un rând de test din merged DataFrame."""
    return {
        "Numar Cerere": "123/456",
        "Data Cerere": "2020-01-01",
        "Cnp Adult Handicap": "1234567890123",
        "Nume Adult cu Handicap": "Popescu",
        "Prenume Adult cu Handicap": "Ion",
        "Stare Cerere": stare,
        "Motiv": motiv,
        "Data Inceput Stare": "2025-01-01",
        "S1.AreNu / AreDatorii": s1,
        "S2.AreNu / AreDatorii": s2,
    }

def _df(*rows):
    return pd.DataFrame(list(rows))

def test_aprobat_fara_datorii_import_obs():
    df = _df(_row("Aprobat", ""))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Import Observatii"

def test_aprobat_cu_datorii_suspendare():
    df = _df(_row("Aprobat", "", s1="DA"))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Suspendare"

def test_suspendat_datorii_motiv_fara_deb_termen_repunere():
    df = _df(_row("Suspendat", "Datorii interogare termen 30.03.2025 S2"))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Repunere"

def test_suspendat_datorii_motiv_cu_deb_termen_import_obs():
    df = _df(_row("Suspendat", "Datorii interogare termen 30.03.2025 S2", s2="DA"))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Import Observatii"

def test_suspendat_alt_motiv_import_obs():
    df = _df(_row("Suspendat", "Cerere incompleta"))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Import Observatii"

def test_incetat_import_obs():
    df = _df(_row("Încetat", "Orice"))
    result = classify_rows(df, SECTORS, TERM_DATE)
    assert result["_actiune"].iloc[0] == "Import Observatii"

def test_build_suspendare_columns():
    df = _df(_row("Aprobat", "", s1="DA"))
    classified = classify_rows(df, SECTORS, TERM_DATE)
    susp = build_suspendare(classified, TERM_DATE)
    assert "Motiv suspendare" in susp.columns
    assert "Data suspendare" in susp.columns
    assert susp["Motiv suspendare"].iloc[0] == "Interogare termen 30.09.2025 / datorii la S1"

def test_build_repunere_columns():
    df = _df(_row("Suspendat", "Datorii interogare termen 30.03.2025 S2"))
    classified = classify_rows(df, SECTORS, TERM_DATE)
    rep = build_repunere(classified, TERM_DATE)
    assert rep["Motiv repunere"].iloc[0] == "Interogare termen 30.09.2025 / Fara datorii"
    assert rep["Data repunere"].iloc[0] == "01.09.2025"

def test_build_import_obs_fara_datorii():
    df = _df(_row("Aprobat", ""))
    classified = classify_rows(df, SECTORS, TERM_DATE)
    imp = build_import_obs(classified, TERM_DATE)
    assert "Fara datorii interogare termen 30.09.2025" in imp["Observatii"].iloc[0]

def test_build_datornici():
    # Aprobat cu datorii (Suspendare) + Suspendat cu datorii → ambii în datornici
    df = _df(
        _row("Aprobat", "", s1="DA"),
        _row("Suspendat", "Datorii interogare termen 30.03.2025", s2="DA"),
    )
    classified = classify_rows(df, SECTORS, TERM_DATE)
    dat = build_datornici(classified, TERM_DATE)
    assert len(dat) == 2
```

**Step 2: Rulează să confirmi FAIL**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -m pytest tests/test_termen_logic.py -v
```

Expected: `ModuleNotFoundError`

**Step 3: Creează `modules/termen_logic.py`**

```python
"""Decision matrix logic for Interogare Termen (quarterly DGITL check)."""
import re
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

# Valori care înseamnă "fără datorii" în coloana AreNu/AreDatorii
_NO_DEBT_VALUES = {"nu", "nu există", "nu exista", "nu există", "", "nan", "none"}

_DEBT_MOTIV_RE = re.compile(
    r'datorii|debit|interogare|termen', re.IGNORECASE
)


def _is_debt_motiv(motiv: str) -> bool:
    """True dacă motivul de suspendare e legat de datorii (orice tip)."""
    return bool(_DEBT_MOTIV_RE.search(str(motiv or "")))


def _debt_cols_for_sectors(sectors: list, all_cols: list) -> list:
    """Returnează coloanele de datorii pentru sectoarele date."""
    result = []
    for s in sectors:
        for col in all_cols:
            if col.startswith(s + "."):
                result.append(col)
    return result


def _has_debt(row: pd.Series, debt_cols: list) -> bool:
    """True dacă cel puțin un sector are datorii."""
    for col in debt_cols:
        val = str(row.get(col, "") or "").strip().lower()
        if val not in _NO_DEBT_VALUES:
            return True
    return False


def _sectors_with_debt(row: pd.Series, sectors: list, all_cols: list) -> list:
    """Returnează lista sectorelor (ex. ['S1', 'S3']) cu datorii confirmate."""
    result = []
    for s in sectors:
        for col in all_cols:
            if col.startswith(s + "."):
                val = str(row.get(col, "") or "").strip().lower()
                if val not in _NO_DEBT_VALUES:
                    result.append(s)
                    break
    return result


def classify_rows(
    merged_df: pd.DataFrame,
    sectors: list,
    term_date: date,
) -> pd.DataFrame:
    """
    Aplică matricea de decizie pe merged_df (Raport Stare + sectoare joined).
    Adaugă coloana '_actiune': "Suspendare" | "Repunere" | "Import Observatii"
    și '_sectoare_datorii': lista sectorelor cu datorii.

    sectors: ["S1", "S2", ..., "S6"] — doar cele încărcate
    """
    df = merged_df.copy()
    all_cols = list(df.columns)
    debt_cols = _debt_cols_for_sectors(sectors, all_cols)

    actiuni = []
    sectoare_datorii = []

    for _, row in df.iterrows():
        stare = str(row.get("Stare Cerere", "") or "").strip().lower()
        motiv = str(row.get("Motiv", "") or "").strip()
        has_debt = _has_debt(row, debt_cols)
        sec_deb = _sectors_with_debt(row, sectors, all_cols) if has_debt else []

        if "încetat" in stare or "incetat" in stare:
            actiuni.append("Import Observatii")
        elif "aprobat" in stare or "aprobată" in stare:
            actiuni.append("Suspendare" if has_debt else "Import Observatii")
        elif "suspendat" in stare or "suspendată" in stare or "suspendata" in stare:
            if _is_debt_motiv(motiv):
                actiuni.append("Repunere" if not has_debt else "Import Observatii")
            else:
                actiuni.append("Import Observatii")
        else:
            actiuni.append("Import Observatii")

        sectoare_datorii.append(sec_deb)

    df["_actiune"] = actiuni
    df["_sectoare_datorii"] = sectoare_datorii
    return df


def _term_date_str(term_date: date) -> str:
    return term_date.strftime("%d.%m.%Y")


def _data_suspendare(term_date: date) -> str:
    """Prima zi a lunii DUPĂ termen. Ex: 30.09.2025 → 01.10.2025"""
    d = (term_date.replace(day=1) + relativedelta(months=1))
    return d.strftime("%d.%m.%Y")


def _data_repunere(term_date: date) -> str:
    """Prima zi a lunii termenului. Ex: 30.09.2025 → 01.09.2025"""
    return term_date.replace(day=1).strftime("%d.%m.%Y")


SUSPENDARE_COLUMNS = [
    "Numar Cerere", "Data Cerere",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap",
    "Motiv suspendare", "Data suspendare",
]

REPUNERE_COLUMNS = [
    "Numar Cerere", "Data Cerere",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap",
    "Motiv repunere", "Data repunere",
]

IMPORT_OBS_COLUMNS = [
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap", "Numar Cerere", "Observatii",
]

DATORNICI_COLUMNS = [
    "Numar Cerere", "Cnp Adult Handicap",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Sectoare datorii", "Stare Cerere", "Motiv",
]


def build_suspendare(df: pd.DataFrame, term_date: date) -> pd.DataFrame:
    """Subset Suspendare cu coloane finale."""
    sub = df[df["_actiune"] == "Suspendare"].copy()
    tds = _term_date_str(term_date)
    sub["Motiv suspendare"] = sub["_sectoare_datorii"].apply(
        lambda sl: f"Interogare termen {tds} / datorii la {' '.join(sl)}"
    )
    sub["Data suspendare"] = _data_suspendare(term_date)
    cols = [c for c in SUSPENDARE_COLUMNS if c in sub.columns]
    return sub[cols].reset_index(drop=True)


def build_repunere(df: pd.DataFrame, term_date: date) -> pd.DataFrame:
    """Subset Repunere cu coloane finale."""
    sub = df[df["_actiune"] == "Repunere"].copy()
    tds = _term_date_str(term_date)
    sub["Motiv repunere"] = f"Interogare termen {tds} / Fara datorii"
    sub["Data repunere"] = _data_repunere(term_date)
    cols = [c for c in REPUNERE_COLUMNS if c in sub.columns]
    return sub[cols].reset_index(drop=True)


def build_import_obs(df: pd.DataFrame, term_date: date) -> pd.DataFrame:
    """Subset Import Observatii cu coloane finale."""
    sub = df[df["_actiune"] == "Import Observatii"].copy()
    tds = _term_date_str(term_date)

    def _obs(row):
        sl = row["_sectoare_datorii"]
        if sl:
            return f"Datorii interogare termen {tds} {' '.join(sl)}"
        return f"Fara datorii interogare termen {tds}"

    sub["Observatii"] = sub.apply(_obs, axis=1)
    cols = [c for c in IMPORT_OBS_COLUMNS if c in sub.columns]
    return sub[cols].reset_index(drop=True)


def build_datornici(df: pd.DataFrame, term_date: date) -> pd.DataFrame:
    """
    Datornici termen = Suspendare ∪ Suspendat-cu-datorii.
    Feed pentru Reinterogarea lunară.
    """
    mask = (df["_actiune"] == "Suspendare") | (
        (df["_actiune"] == "Import Observatii") &
        (df["Stare Cerere"].str.lower().str.contains("suspendat|suspendata|suspendată", na=False)) &
        (df["_sectoare_datorii"].apply(len) > 0)
    )
    sub = df[mask].copy()
    sub["Sectoare datorii"] = sub["_sectoare_datorii"].apply(lambda sl: " ".join(sl))
    cols = [c for c in DATORNICI_COLUMNS if c in sub.columns]
    return sub[cols].reset_index(drop=True)
```

**Step 4: Rulează testele**

```bash
python -m pytest tests/test_termen_logic.py -v
```

Expected: 10 PASS

**Step 5: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 6: Commit**

```bash
git add modules/termen_logic.py tests/test_termen_logic.py
git commit -m "feat: add termen_logic with decision matrix and output builders"
```

---

## Task 2: `modules/contingent_tracker.py`

**Files:**
- Create: `APP_DIR/modules/contingent_tracker.py`
- Create: `APP_DIR/tests/test_contingent_tracker.py`

**Step 1: Scrie testele**

```python
# tests/test_contingent_tracker.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
import pytest
from datetime import date
from modules.contingent_tracker import ContingentTracker


def test_init_from_datornici(tmp_path):
    tracker = ContingentTracker(tmp_path / "contingent.json")
    datornici = pd.DataFrame([
        {"Numar Cerere": "1/1", "Cnp Adult Handicap": "111", "Sectoare datorii": "S2"},
        {"Numar Cerere": "2/2", "Cnp Adult Handicap": "222", "Sectoare datorii": "S1 S3"},
    ])
    tracker.init_from_termen(datornici, date(2025, 9, 30))

    state = tracker.load()
    assert state["count"] == 2
    assert len(state["entries"]) == 2
    assert state["source"] == "Interogare termen 30.09.2025"


def test_clean_against_raport_stare(tmp_path):
    tracker = ContingentTracker(tmp_path / "contingent.json")
    datornici = pd.DataFrame([
        {"Numar Cerere": "1/1", "Cnp Adult Handicap": "111", "Sectoare datorii": "S2"},
        {"Numar Cerere": "2/2", "Cnp Adult Handicap": "222", "Sectoare datorii": "S1"},
    ])
    tracker.init_from_termen(datornici, date(2025, 9, 30))

    # CNP 111 devine Aprobat → iese din contingent
    raport_stare = pd.DataFrame([
        {"Cnp Adult Handicap": "111", "Stare Cerere": "Aprobat"},
        {"Cnp Adult Handicap": "222", "Stare Cerere": "Suspendat"},
    ])
    removed = tracker.clean_against_raport(raport_stare, date(2025, 10, 31))

    state = tracker.load()
    assert state["count"] == 1
    assert removed == 1
    assert state["entries"][0]["cnp"] == "222"


def test_contingent_empty_after_all_approved(tmp_path):
    tracker = ContingentTracker(tmp_path / "contingent.json")
    datornici = pd.DataFrame([
        {"Numar Cerere": "1/1", "Cnp Adult Handicap": "111", "Sectoare datorii": "S2"},
    ])
    tracker.init_from_termen(datornici, date(2025, 9, 30))

    raport = pd.DataFrame([{"Cnp Adult Handicap": "111", "Stare Cerere": "Aprobat"}])
    tracker.clean_against_raport(raport, date(2025, 10, 31))

    state = tracker.load()
    assert state["count"] == 0
```

**Step 2: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_contingent_tracker.py -v
```

Expected: `ModuleNotFoundError`

**Step 3: Creează `modules/contingent_tracker.py`**

```python
"""Persistent tracker for the active datornici contingent."""
import json
from datetime import date
from pathlib import Path
import pandas as pd


class ContingentTracker:
    def __init__(self, path):
        self.path = Path(path)

    def load(self) -> dict:
        if not self.path.exists():
            return {"count": 0, "entries": [], "source": "", "updated_at": ""}
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {"count": 0, "entries": [], "source": "", "updated_at": ""}

    def _save(self, state: dict):
        self.path.write_text(
            json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def init_from_termen(self, datornici_df: pd.DataFrame, term_date: date):
        """Inițializează contingentul din Datornici termen (la fiecare Interogare Termen)."""
        entries = []
        for _, row in datornici_df.iterrows():
            entries.append({
                "cnp": str(row.get("Cnp Adult Handicap", "") or "").strip(),
                "numar_cerere": str(row.get("Numar Cerere", "") or "").strip(),
                "sectoare_datorii": str(row.get("Sectoare datorii", "") or "").strip(),
                "data_intrare": term_date.strftime("%Y-%m-%d"),
                "stare_initiala": str(row.get("Stare Cerere", "") or "").strip(),
            })
        state = {
            "source": f"Interogare termen {term_date.strftime('%d.%m.%Y')}",
            "updated_at": term_date.strftime("%Y-%m-%d"),
            "count": len(entries),
            "entries": entries,
        }
        self._save(state)

    def clean_against_raport(self, raport_df: pd.DataFrame, check_date: date) -> int:
        """
        Elimină din contingent CNP-urile care apar cu Stare = Aprobat în Raport Stare.
        Returnează numărul de intrări eliminate.
        """
        state = self.load()

        # Construiește set CNP-uri Aprobate
        approved_cnps = set()
        stare_col = "Stare Cerere" if "Stare Cerere" in raport_df.columns else None
        cnp_col = "Cnp Adult Handicap" if "Cnp Adult Handicap" in raport_df.columns else None

        if stare_col and cnp_col:
            mask = raport_df[stare_col].str.lower().str.contains(
                "aprobat|aprobată", na=False
            )
            approved_cnps = set(raport_df.loc[mask, cnp_col].str.strip().dropna())

        before = len(state["entries"])
        state["entries"] = [
            e for e in state["entries"]
            if e.get("cnp", "") not in approved_cnps
        ]
        removed = before - len(state["entries"])
        state["count"] = len(state["entries"])
        state["updated_at"] = check_date.strftime("%Y-%m-%d")
        self._save(state)
        return removed
```

**Step 4: Rulează testele**

```bash
python -m pytest tests/test_contingent_tracker.py -v
```

Expected: 3 PASS

**Step 5: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 6: Commit**

```bash
git add modules/contingent_tracker.py tests/test_contingent_tracker.py
git commit -m "feat: add ContingentTracker for active datornici state"
```

---

## Task 3: `modules/termen_processor.py` + output în `output_builder.py`

**Files:**
- Create: `APP_DIR/modules/termen_processor.py`
- Modify: `APP_DIR/modules/output_builder.py` — adaugă `write_termen_outputs()`
- Create: `APP_DIR/tests/test_termen_processor.py`

**Step 1: Adaugă `write_termen_outputs` în `output_builder.py`**

Adaugă constantele și metoda în clasa `OutputBuilder` (după `write_verification_file`):

```python
# Constante coloane output termen (în output_builder.py, la nivel de modul)
TERMEN_SUSPENDARE_COLUMNS = [
    "Numar Cerere", "Data Cerere",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap", "Motiv suspendare", "Data suspendare",
]
TERMEN_REPUNERE_COLUMNS = [
    "Numar Cerere", "Data Cerere",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap", "Motiv repunere", "Data repunere",
]
TERMEN_IMPORT_OBS_COLUMNS = [
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Cnp Adult Handicap", "Numar Cerere", "Observatii",
]
TERMEN_DATORNICI_COLUMNS = [
    "Numar Cerere", "Cnp Adult Handicap",
    "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
    "Sectoare datorii", "Stare Cerere", "Motiv",
]

# Metodă în OutputBuilder:
def write_termen_outputs(
    self,
    output_dir,
    term_date_str: str,
    suspendare_df: pd.DataFrame,
    repunere_df: pd.DataFrame,
    import_obs_df: pd.DataFrame,
    datornici_df: pd.DataFrame,
):
    """Scrie cele 4 fișiere output Interogare Termen în output_dir."""
    from datetime import datetime
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    files = [
        ("Suspendare.xlsx", suspendare_df, TERMEN_SUSPENDARE_COLUMNS),
        ("Repunere.xlsx", repunere_df, TERMEN_REPUNERE_COLUMNS),
        ("Import Observatii.xlsx", import_obs_df, TERMEN_IMPORT_OBS_COLUMNS),
        (f"Datornici termen {term_date_str}.xlsx", datornici_df, TERMEN_DATORNICI_COLUMNS),
    ]
    for filename, df, cols in files:
        try:
            out_cols = [c for c in cols if c in df.columns]
            out = df[out_cols].copy() if out_cols else pd.DataFrame(columns=cols)
            out.to_excel(output_dir / filename, index=False)
        except Exception as e:
            print(f"[warn] Nu s-a putut scrie {filename}: {e}")
```

**Step 2: Scrie testul**

```python
# tests/test_termen_processor.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from datetime import date
from modules.output_builder import OutputBuilder


def test_write_termen_outputs_creates_files(tmp_path):
    term_date_str = "30.09.2025"
    susp = pd.DataFrame([{
        "Numar Cerere": "1", "Cnp Adult Handicap": "111",
        "Motiv suspendare": "Interogare termen 30.09.2025 / datorii la S1",
        "Data suspendare": "01.10.2025",
        "Nume Adult cu Handicap": "Pop", "Prenume Adult cu Handicap": "Ion",
        "Data Cerere": "2020-01-01",
    }])
    rep = pd.DataFrame(columns=["Numar Cerere", "Cnp Adult Handicap",
                                  "Motiv repunere", "Data repunere",
                                  "Nume Adult cu Handicap", "Prenume Adult cu Handicap",
                                  "Data Cerere"])
    imp = pd.DataFrame([{
        "Numar Cerere": "2", "Cnp Adult Handicap": "222",
        "Observatii": "Fara datorii interogare termen 30.09.2025",
        "Nume Adult cu Handicap": "Ion", "Prenume Adult cu Handicap": "Maria",
    }])
    dat = pd.DataFrame([{
        "Numar Cerere": "1", "Cnp Adult Handicap": "111",
        "Sectoare datorii": "S1", "Stare Cerere": "Aprobat", "Motiv": "",
        "Nume Adult cu Handicap": "Pop", "Prenume Adult cu Handicap": "Ion",
    }])

    OutputBuilder().write_termen_outputs(tmp_path, term_date_str, susp, rep, imp, dat)

    assert (tmp_path / "Suspendare.xlsx").exists()
    assert (tmp_path / "Repunere.xlsx").exists()
    assert (tmp_path / "Import Observatii.xlsx").exists()
    assert (tmp_path / "Datornici termen 30.09.2025.xlsx").exists()
```

**Step 3: Rulează să confirmi FAIL**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -m pytest tests/test_termen_processor.py::test_write_termen_outputs_creates_files -v
```

Expected: `AttributeError`

**Step 4: Creează `modules/termen_processor.py`**

```python
"""Orchestrates the Interogare Termen processing pipeline."""
import pandas as pd
from pathlib import Path
from datetime import date

from modules.config_manager import ConfigManager
from modules.column_detector import ColumnDetector
from modules.sector_loader import SectorLoader
from modules.raport_stari_loader import RaportStariLoader
from modules.merger import Merger
from modules.output_builder import OutputBuilder
from modules.termen_logic import (
    classify_rows, build_suspendare, build_repunere,
    build_import_obs, build_datornici,
)
from modules.contingent_tracker import ContingentTracker


class TermenProcessor:
    def __init__(self, config_manager: ConfigManager, progress_callback=None,
                 contingent_path: Path = None):
        self.cm = config_manager
        self.progress_cb = progress_callback or (lambda msg: None)
        self.contingent_path = contingent_path

    def process(
        self,
        sector_paths: dict,
        term_date: date,
        output_dir: Path,
        raport_stari_path: Path = None,
        debt_col_overrides: dict = None,
    ) -> dict:
        aliases = self.cm.load_column_aliases()
        detector = ColumnDetector(aliases)
        loader = SectorLoader(detector)
        merger = Merger()
        builder = OutputBuilder()

        sector_dfs_narrow = {}
        detection_log = {}
        warnings = []

        # --- Load sectors ---
        loaded_sectors = []
        for sector_name, filepath in sector_paths.items():
            self.progress_cb(f"Loading {sector_name}...")
            if not filepath or not Path(filepath).exists():
                warnings.append(f"{sector_name}: file not found, skipped")
                continue
            try:
                forced = (debt_col_overrides or {}).get(sector_name)
                df_narrow, detected, missing = loader.load(
                    Path(filepath), sector_name,
                    run_date=term_date, forced_debt_col=forced,
                )
            except Exception as e:
                warnings.append(f"{sector_name}: load error — {e}")
                continue

            detection_log[sector_name] = {
                k: str(v) for k, v in detected.items()
                if v and not k.startswith("_")
            }
            if missing:
                warnings.append(f"{sector_name}: critical columns not detected: {missing}")

            sector_dfs_narrow[sector_name] = df_narrow
            loaded_sectors.append(sector_name)

        self.cm.save_column_aliases(detector.config)

        if not raport_stari_path or not Path(raport_stari_path).exists():
            return {
                "success": False,
                "error": "Raport Stare lipsă — necesar pentru Interogare Termen.",
                "warnings": warnings,
            }

        # --- Load Raport Stare + merge ---
        self.progress_cb("Loading Raport Stare...")
        try:
            base_df = RaportStariLoader().load(Path(raport_stari_path))
            self.progress_cb("Merging sectors...")
            merged_df = merger.merge(base_df, sector_dfs_narrow)
        except Exception as e:
            return {"success": False, "error": str(e), "warnings": warnings}

        # --- Classify ---
        self.progress_cb("Applying decision matrix...")
        classified = classify_rows(merged_df, loaded_sectors, term_date)

        suspendare_df = build_suspendare(classified, term_date)
        repunere_df   = build_repunere(classified, term_date)
        import_obs_df = build_import_obs(classified, term_date)
        datornici_df  = build_datornici(classified, term_date)

        # --- Write outputs ---
        self.progress_cb("Writing output files...")
        term_date_str = term_date.strftime("%d.%m.%Y")
        try:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            builder.write_termen_outputs(
                output_dir, term_date_str,
                suspendare_df, repunere_df, import_obs_df, datornici_df,
            )
        except PermissionError as e:
            return {"success": False, "error": str(e), "warnings": warnings}

        # --- Update contingent ---
        if self.contingent_path:
            try:
                tracker = ContingentTracker(self.contingent_path)
                tracker.init_from_termen(datornici_df, term_date)
            except Exception as e:
                warnings.append(f"Contingent update failed: {e}")

        self.progress_cb("Done.")
        return {
            "success": True,
            "n_suspendare": len(suspendare_df),
            "n_repunere": len(repunere_df),
            "n_import_obs": len(import_obs_df),
            "n_datornici": len(datornici_df),
            "warnings": warnings,
            "detection_log": detection_log,
        }
```

**Step 5: Rulează testele**

```bash
python -m pytest tests/test_termen_processor.py -v
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 6: Commit**

```bash
git add modules/termen_processor.py modules/output_builder.py tests/test_termen_processor.py
git commit -m "feat: add TermenProcessor pipeline and write_termen_outputs"
```

---

## Task 4: `interogare_termen_processor.py` — UI

**Files:**
- Create: `APP_DIR/interogare_termen_processor.py`

**Step 1: Verifică sintaxa după creare**

```bash
python -c "
import ast
with open('interogare_termen_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

**Step 2: Creează `interogare_termen_processor.py`**

```python
"""Interogare Termen Processor — UI CustomTkinter."""
import os
import threading
from pathlib import Path
from datetime import date, datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox

from modules.config_manager import ConfigManager
from modules.termen_processor import TermenProcessor
from modules.contingent_tracker import ContingentTracker
from modules.ui_components import FilePicker, LogBox

APP_DIR = Path(__file__).parent
CONFIG_DIR = APP_DIR / "config"
CONTINGENT_PATH = APP_DIR / "contingent.json"

SECTORS = ["S1", "S2", "S3", "S4", "S5", "S6"]

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class InterogareTernenApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Interogare Termen Processor")
        self.geometry("720x700")
        self.resizable(True, True)

        self.cm = ConfigManager(CONFIG_DIR)
        self.debt_col_overrides = {}

        self._build_ui()
        self._load_settings()
        self._refresh_contingent_label()

    def _build_ui(self):
        # --- Titlu ---
        ctk.CTkLabel(self, text="Interogare Termen",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(16, 4))

        # --- Data termen ---
        date_frame = ctk.CTkFrame(self, fg_color="transparent")
        date_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(date_frame, text="Data termen:", width=110, anchor="w").pack(side="left")
        self.date_entry = ctk.CTkEntry(date_frame, width=120, placeholder_text="30.09.2025")
        self.date_entry.pack(side="left", padx=(0, 8))
        today = datetime.now()
        default_date = f"30.09.{today.year}" if today.month <= 9 else f"30.03.{today.year + 1}"
        self.date_entry.insert(0, default_date)

        # --- Raport Stare ---
        rs_frame = ctk.CTkFrame(self)
        rs_frame.pack(fill="x", padx=16, pady=4)
        self.raport_stari_picker = FilePicker(rs_frame, label="Raport Stare")
        self.raport_stari_picker.pack(fill="x", padx=10, pady=6)

        # --- Sectoare S1-S6 ---
        sectors_frame = ctk.CTkFrame(self)
        sectors_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(sectors_frame, text="Fișiere sector (S1–S6):",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(6, 2))

        self.sector_pickers = {}
        for s in SECTORS:
            picker = FilePicker(
                sectors_frame, label=s, clearable=True,
                on_change=lambda p, sec=s: self._on_sector_changed(sec, p),
                on_clear=lambda sec=s: self._on_sector_cleared(sec),
                filetypes=[("Excel/XLS files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
            )
            picker.pack(fill="x", padx=10, pady=2)
            self.sector_pickers[s] = picker

        # --- Output dir ---
        out_frame = ctk.CTkFrame(self)
        out_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(out_frame, text="Output dir:", width=110, anchor="w").pack(side="left", padx=(10, 6))
        self.output_entry = ctk.CTkEntry(out_frame, width=340)
        self.output_entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(out_frame, text="...", width=32,
                      command=self._pick_output_dir).pack(side="left")

        # --- Buton Procesează ---
        self.run_btn = ctk.CTkButton(
            self, text="▶  Procesează", width=200, height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run,
        )
        self.run_btn.pack(pady=(12, 4))

        # --- Progress + Contingent ---
        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(fill="x", padx=16, pady=4)
        self.progress.set(0)

        self.contingent_label = ctk.CTkLabel(
            self, text="Contingent activ: —", anchor="w",
            font=ctk.CTkFont(size=12),
        )
        self.contingent_label.pack(anchor="w", padx=16, pady=2)

        # --- Log ---
        self.log = LogBox(self, height=160)
        self.log.pack(fill="both", expand=True, padx=16, pady=(4, 16))

    def _load_settings(self):
        settings = self.cm.load_settings()
        key = "termen_sector_paths"
        saved = settings.get(key, {})
        for s, picker in self.sector_pickers.items():
            if saved.get(s):
                picker.set(saved[s])
        if settings.get("termen_raport_stari_path"):
            self.raport_stari_picker.set(settings["termen_raport_stari_path"])
        if settings.get("termen_output_dir"):
            self.output_entry.insert(0, settings["termen_output_dir"])

    def _refresh_contingent_label(self):
        try:
            tracker = ContingentTracker(CONTINGENT_PATH)
            state = tracker.load()
            count = state.get("count", 0)
            updated = state.get("updated_at", "")
            source = state.get("source", "")
            self.contingent_label.configure(
                text=f"Contingent activ: {count} beneficiari  |  {source}  ({updated})"
            )
        except Exception:
            self.contingent_label.configure(text="Contingent activ: —")

    def _on_sector_changed(self, sector: str, path: str):
        pass  # extensibil: auto-inspect la nevoie

    def _on_sector_cleared(self, sector: str):
        self.debt_col_overrides.pop(sector, None)

    def _pick_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, d)

    def _on_progress(self, msg: str):
        self.after(0, lambda: self.log.append(msg))
        self.after(0, lambda: self.progress.set(self.progress.get() + 0.1))

    def _run(self):
        sector_paths = {}
        for s, picker in self.sector_pickers.items():
            p = picker.get()
            if p:
                sector_paths[s] = p

        if not sector_paths:
            messagebox.showwarning("Atenție", "Selectează cel puțin un fișier sector.")
            return

        output_dir = self.output_entry.get().strip()
        if not output_dir:
            messagebox.showwarning("Atenție", "Selectează directorul de output.")
            return

        date_str = self.date_entry.get().strip()
        try:
            term_date = datetime.strptime(date_str, "%d.%m.%Y").date()
        except ValueError:
            messagebox.showwarning("Atenție", f"Dată invalidă: '{date_str}'\nFormat: ZZ.LL.AAAA")
            return

        raport_path = self.raport_stari_picker.get()
        if not raport_path:
            messagebox.showwarning("Atenție", "Selectează Raportul de Stare.")
            return

        # Save settings
        self.cm.save_setting("termen_sector_paths",
                             {s: picker.get() for s, picker in self.sector_pickers.items()})
        self.cm.save_setting("termen_raport_stari_path", raport_path)
        self.cm.save_setting("termen_output_dir", output_dir)

        self.run_btn.configure(state="disabled")
        self.log.clear()
        self.progress.set(0.05)

        def _worker():
            processor = TermenProcessor(
                self.cm,
                progress_callback=self._on_progress,
                contingent_path=CONTINGENT_PATH,
            )
            result = processor.process(
                sector_paths=sector_paths,
                term_date=term_date,
                output_dir=Path(output_dir),
                raport_stari_path=Path(raport_path),
                debt_col_overrides=self.debt_col_overrides,
            )
            self.after(0, lambda: self._on_done(result, output_dir))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self, result: dict, output_dir: str):
        self.run_btn.configure(state="normal")
        self.progress.set(1.0)

        if not result.get("success"):
            err = result.get("error", "Eroare necunoscută")
            messagebox.showerror("Eroare", err)
            self.log.append(f"EROARE: {err}")
            return

        n_s = result.get("n_suspendare", 0)
        n_r = result.get("n_repunere", 0)
        n_i = result.get("n_import_obs", 0)
        n_d = result.get("n_datornici", 0)

        self.log.append(
            f"\nRezultat:\n"
            f"  Suspendare:       {n_s}\n"
            f"  Repunere:         {n_r}\n"
            f"  Import Observatii:{n_i}\n"
            f"  Datornici termen: {n_d}\n"
            f"\nOutput: {output_dir}"
        )

        for w in result.get("warnings", []):
            self.log.append(f"  [warn] {w}")

        self._refresh_contingent_label()

        if messagebox.askyesno("Gata!", f"Procesare completă.\nDeschizi folderul output?"):
            os.startfile(output_dir)


if __name__ == "__main__":
    app = InterogareTernenApp()
    app.mainloop()
```

**Step 3: Verifică sintaxa**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -c "
import ast
with open('interogare_termen_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

Expected: `Syntax OK`

**Step 4: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 5: Commit**

```bash
git add interogare_termen_processor.py
git commit -m "feat: add InterogareTermenApp UI"
```

---

## Task 5: Wire contingent cleanup în Reinterogare Processor

**Files:**
- Modify: `APP_DIR/reinterogare_processor.py` — adaugă curățare contingent după procesare

**Step 1: Găsește `_on_done` în `reinterogare_processor.py`**

Metoda `_on_done` (linia ~599) este apelată după procesare. Adaugă la final (după pre-fill verif_output_entry):

```python
# Curăță contingentul față de Raport Stare curent
_raport_path = self.raport_stari_picker.get()
if _raport_path and Path(_raport_path).exists():
    try:
        from modules.contingent_tracker import ContingentTracker
        import pandas as pd as _pd
        _rs = _pd.read_excel(_raport_path, dtype=str)
        _tracker = ContingentTracker(Path(__file__).parent / "contingent.json")
        _removed = _tracker.clean_against_raport(_rs, __import__('datetime').date.today())
        if _removed > 0:
            self.log.append(f"  [contingent] {_removed} beneficiari ieșiți din contingent (Aprobat)")
    except Exception as _e:
        pass  # non-critical
```

**ATENȚIE**: Importul `import pandas as pd as _pd` nu e valid Python. Folosește:
```python
import pandas as _pd
```

**Step 2: Implementează corect**

În `reinterogare_processor.py`, în metoda `_on_done`, după blocul de pre-fill (`self.verif_output_entry.insert`), adaugă:

```python
# Curăță contingentul față de Raport Stare curent (non-critical)
_raport_path = self.raport_stari_picker.get()
if _raport_path and Path(_raport_path).exists():
    try:
        import pandas as _pd_cont
        from modules.contingent_tracker import ContingentTracker as _CT
        _rs = _pd_cont.read_excel(_raport_path, dtype=str)
        _tracker = _CT(Path(__file__).parent / "contingent.json")
        _removed = _tracker.clean_against_raport(
            _rs, __import__("datetime").date.today()
        )
        if _removed > 0:
            self.log.append(
                f"  [contingent] {_removed} beneficiari ieșiți din contingent (Aprobat)"
            )
    except Exception:
        pass  # non-critical
```

**Step 3: Verifică sintaxa**

```bash
python -c "
import ast
with open('reinterogare_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

**Step 4: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 5: Commit**

```bash
git add reinterogare_processor.py
git commit -m "feat: wire contingent cleanup in Reinterogare Processor after processing"
```

---

## Checklist final

- [ ] `classify_rows` aplică corect matricea de decizie pentru toate 6 scenarii
- [ ] `build_suspendare` generează motivul și data corect
- [ ] `build_repunere` generează motivul și data corect
- [ ] `build_import_obs` generează observația corect (cu/fără datorii)
- [ ] `build_datornici` = Suspendare ∪ Suspendat-cu-datorii
- [ ] `ContingentTracker.init_from_termen` → salvează JSON
- [ ] `ContingentTracker.clean_against_raport` → elimină Aprobații
- [ ] `write_termen_outputs` → 4 fișiere create
- [ ] UI: pickers S1–S6 + Raport Stare + dată termen + output dir + buton
- [ ] UI: label contingent activ actualizat după procesare
- [ ] Reinterogare Processor curăță automat contingentul după run
- [ ] Toate testele pytest trec
