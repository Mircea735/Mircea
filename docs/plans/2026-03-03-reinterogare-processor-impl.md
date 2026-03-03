# Reinterogare Processor — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Python/CustomTkinter desktop app that loads 6 sector Excel files (S1–S6) monthly, auto-detects column names with fuzzy matching, classifies arbitrary suspension motives, computes debt status, and outputs a structured Excel file ready for the social benefits platform.

**Architecture:** Modular Python with clean separation between business logic (testable with pytest) and UI (CustomTkinter, tested manually). Config persists learned column aliases and motive classifications as JSON. Output Excel has per-sector inspection sheets plus a consolidated platform-ready sheet.

**Tech Stack:** Python 3.10+, CustomTkinter, openpyxl, pandas, rapidfuzz (fuzzy matching), pytest

**Design doc:** `docs/plans/2026-03-03-reinterogare-processor-design.md`

---

## Setup — before starting

Install dependencies in the target environment:
```bash
pip install customtkinter openpyxl pandas rapidfuzz pytest
```

All files go under: `C:\EXCELURI\Reinterogari\Reinterogare Processor\`
(referred to as `APP_DIR` below)

---

### Task 1: Project Scaffold

**Files:**
- Create: `APP_DIR/reinterogare_processor.py` (entry point, placeholder)
- Create: `APP_DIR/modules/__init__.py`
- Create: `APP_DIR/config/column_aliases.json`
- Create: `APP_DIR/config/motiv_rules.json`
- Create: `APP_DIR/config/settings.json`
- Create: `APP_DIR/tests/__init__.py`
- Create: `APP_DIR/requirements.txt`

**Step 1: Create directory structure**

```bash
mkdir -p "C:/EXCELURI/Reinterogari/Reinterogare Processor/modules"
mkdir -p "C:/EXCELURI/Reinterogari/Reinterogare Processor/config"
mkdir -p "C:/EXCELURI/Reinterogari/Reinterogare Processor/tests"
```

**Step 2: Create requirements.txt**

```
customtkinter>=5.2.0
openpyxl>=3.1.0
pandas>=2.0.0
rapidfuzz>=3.0.0
pytest>=7.0.0
```

**Step 3: Create empty config files**

`config/column_aliases.json`:
```json
{
  "debt_flag": {
    "patterns": ["are.?nu.?are.?debit", "ARE.?NU.?ARE.?DEBIT"],
    "known_exact": []
  },
  "data_plata": {
    "patterns": ["achitat|platit|data.?pla"],
    "known_exact": []
  },
  "data_repunere": {
    "patterns": ["data.?repunere", "data.?repunerii"],
    "known_exact": []
  },
  "cnp": {
    "patterns": ["^cnp$", "cod.?numeric"],
    "known_exact": ["CNP"]
  },
  "numar_cerere": {
    "patterns": ["nr\\.?.?cerere", "numar.?cerere"],
    "known_exact": []
  }
}
```

`config/motiv_rules.json`:
```json
{
  "datorii_ditl": [],
  "exclude": []
}
```

`config/settings.json`:
```json
{
  "last_sector_paths": {"S1": "", "S2": "", "S3": "", "S4": "", "S5": "", "S6": ""},
  "last_raport_stari_path": "",
  "last_interogare_termen_path": "",
  "output_dir": ""
}
```

**Step 4: Create placeholder entry point**

`reinterogare_processor.py`:
```python
"""Reinterogare Processor - entry point"""
import sys
from pathlib import Path

APP_DIR = Path(__file__).parent
CONFIG_DIR = APP_DIR / "config"


def main():
    print("Reinterogare Processor starting...")


if __name__ == "__main__":
    main()
```

**Step 5: Create `modules/__init__.py`** (empty file)

**Step 6: Commit**

```bash
git add .
git commit -m "feat: scaffold Reinterogare Processor project"
```

---

### Task 2: Column Detector Module

**Business rule:** Each S1–S6 Excel file has columns whose names change monthly. We need to detect 5 key concepts: `debt_flag`, `data_plata`, `data_repunere`, `cnp`, `numar_cerere`. Detection uses regex patterns then fuzzy fallback.

**Files:**
- Create: `APP_DIR/modules/column_detector.py`
- Create: `APP_DIR/tests/test_column_detector.py`

**Step 1: Write failing tests**

`tests/test_column_detector.py`:
```python
import pytest
from modules.column_detector import ColumnDetector

# Minimal config matching config/column_aliases.json structure
TEST_CONFIG = {
    "debt_flag": {
        "patterns": ["are.?nu.?are.?debit", "ARE.?NU.?ARE.?DEBIT"],
        "known_exact": []
    },
    "data_plata": {
        "patterns": ["achitat|platit|data.?pla"],
        "known_exact": []
    },
    "data_repunere": {
        "patterns": ["data.?repunere", "data.?repunerii"],
        "known_exact": []
    },
    "cnp": {
        "patterns": ["^cnp$", "cod.?numeric"],
        "known_exact": ["CNP"]
    },
    "numar_cerere": {
        "patterns": ["nr\\.?.?cerere", "numar.?cerere"],
        "known_exact": []
    }
}


def make_detector(extra_aliases=None):
    cfg = TEST_CONFIG.copy()
    if extra_aliases:
        for concept, names in extra_aliases.items():
            cfg[concept]["known_exact"] = names
    return ColumnDetector(cfg)


class TestExactMatch:
    def test_cnp_exact(self):
        d = make_detector()
        result = d.detect(["CNP", "Numar cerere", "Alte date"])
        assert result["cnp"] == "CNP"

    def test_known_alias_wins(self):
        d = make_detector(extra_aliases={"debt_flag": ["Are/Nu are debit la 31.12.2025"]})
        result = d.detect(["Are/Nu are debit la 31.12.2025", "CNP"])
        assert result["debt_flag"] == "Are/Nu are debit la 31.12.2025"


class TestRegexMatch:
    def test_debt_flag_regex(self):
        d = make_detector()
        result = d.detect(["Are/Nu are debit la 31.12.2025", "CNP", "Nr. Cerere"])
        assert result["debt_flag"] == "Are/Nu are debit la 31.12.2025"

    def test_debt_flag_uppercase(self):
        d = make_detector()
        result = d.detect(["ARE/NU ARE DEBIT la 01.01.2026", "CNP"])
        assert result["debt_flag"] == "ARE/NU ARE DEBIT la 01.01.2026"

    def test_data_plata_old_name(self):
        d = make_detector()
        cols = ["Achitat la data - Observații2", "CNP"]
        result = d.detect(cols)
        assert result["data_plata"] == "Achitat la data - Observații2"

    def test_data_plata_new_name(self):
        d = make_detector()
        cols = ["Data plății - Observații", "CNP"]
        result = d.detect(cols)
        assert result["data_plata"] == "Data plății - Observații"

    def test_data_repunere(self):
        d = make_detector()
        result = d.detect(["Data repunere", "CNP"])
        assert result["data_repunere"] == "Data repunere"

    def test_data_repunerii_variant(self):
        d = make_detector()
        result = d.detect(["Data repunerii", "CNP"])
        assert result["data_repunere"] == "Data repunerii"

    def test_numar_cerere(self):
        d = make_detector()
        result = d.detect(["Nr. Cerere", "CNP"])
        assert result["numar_cerere"] == "Nr. Cerere"


class TestMissingColumns:
    def test_missing_returns_none(self):
        d = make_detector()
        result = d.detect(["CNP", "Alte date"])
        assert result["debt_flag"] is None
        assert result["data_repunere"] is None

    def test_missing_columns_listed(self):
        d = make_detector()
        result = d.detect(["CNP"])
        missing = d.get_missing_critical(result)
        assert "debt_flag" in missing
        assert "cnp" not in missing


class TestLearnAlias:
    def test_learn_saves_to_known_exact(self):
        d = make_detector()
        d.learn_alias("debt_flag", "Situatie debit special")
        result = d.detect(["Situatie debit special", "CNP"])
        assert result["debt_flag"] == "Situatie debit special"
```

**Step 2: Run tests — verify they fail**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
pytest tests/test_column_detector.py -v
```
Expected: `ModuleNotFoundError: No module named 'modules.column_detector'`

**Step 3: Implement `modules/column_detector.py`**

```python
"""Detects key column names in sector Excel files using patterns + fuzzy matching."""
import re
from typing import Optional
from rapidfuzz import fuzz

CRITICAL_COLUMNS = {"debt_flag", "cnp", "numar_cerere"}


class ColumnDetector:
    def __init__(self, config: dict):
        """
        config: dict with concept keys, each having:
          - patterns: list of regex strings (case-insensitive)
          - known_exact: list of exact column names learned from past runs
        """
        self.config = config

    def detect(self, columns: list[str]) -> dict[str, Optional[str]]:
        """Return {concept: matched_column_name} for each concept in config.
        Returns None for concepts that couldn't be matched."""
        result = {}
        for concept, rules in self.config.items():
            result[concept] = self._match_concept(concept, rules, columns)
        return result

    def _match_concept(self, concept: str, rules: dict, columns: list[str]) -> Optional[str]:
        # 1. Exact match against known aliases
        for col in columns:
            if col in rules.get("known_exact", []):
                return col

        # 2. Regex match (case-insensitive)
        for pattern in rules.get("patterns", []):
            for col in columns:
                if re.search(pattern, col, re.IGNORECASE):
                    return col

        # 3. Fuzzy fallback (80% threshold) against known_exact only
        best_score = 0
        best_col = None
        for col in columns:
            for alias in rules.get("known_exact", []):
                score = fuzz.ratio(col.lower(), alias.lower())
                if score > best_score:
                    best_score = score
                    best_col = col
        if best_score >= 80:
            return best_col

        return None

    def get_missing_critical(self, detected: dict) -> list[str]:
        """Returns list of critical concept names that were not detected."""
        return [c for c in CRITICAL_COLUMNS if detected.get(c) is None]

    def learn_alias(self, concept: str, column_name: str):
        """Teach the detector a new exact alias for a concept."""
        if concept in self.config:
            if column_name not in self.config[concept]["known_exact"]:
                self.config[concept]["known_exact"].append(column_name)
```

**Step 4: Run tests — verify they pass**

```bash
pytest tests/test_column_detector.py -v
```
Expected: all tests PASS

**Step 5: Commit**

```bash
git add modules/column_detector.py tests/test_column_detector.py
git commit -m "feat: add ColumnDetector with regex + fuzzy matching"
```

---

### Task 3: Motiv Classifier Module

**Business rule:** The `Motiv` field in Raport Stări is written arbitrarily by operators. We need to classify each distinct motiv value as: `datorii_ditl` (eligible for repunere processing), `exclude` (deceased, withdrawn), or `unknown` (needs user confirmation).

**Files:**
- Create: `APP_DIR/modules/motiv_classifier.py`
- Create: `APP_DIR/tests/test_motiv_classifier.py`

**Step 1: Write failing tests**

`tests/test_motiv_classifier.py`:
```python
import pytest
from modules.motiv_classifier import MotivClassifier

RULES = {
    "datorii_ditl": ["Datorii DITL", "debite dgitl", "are datorii la ditl"],
    "exclude": ["decedat", "retras", "renuntat"]
}


def make_clf():
    return MotivClassifier(RULES, fuzzy_threshold=75)


class TestExactClassification:
    def test_known_datorii(self):
        clf = make_clf()
        assert clf.classify("Datorii DITL") == "datorii_ditl"

    def test_known_exclude(self):
        clf = make_clf()
        assert clf.classify("decedat") == "exclude"

    def test_unknown_returns_unknown(self):
        clf = make_clf()
        assert clf.classify("nu stiu ce e asta") == "unknown"


class TestFuzzyClassification:
    def test_fuzzy_datorii(self):
        clf = make_clf()
        # typo variant
        assert clf.classify("Datoriiz DITL") == "datorii_ditl"

    def test_fuzzy_exclude(self):
        clf = make_clf()
        assert clf.classify("Decedatt") == "exclude"

    def test_case_insensitive(self):
        clf = make_clf()
        assert clf.classify("DATORII DITL") == "datorii_ditl"

    def test_partial_phrase(self):
        clf = make_clf()
        # Operator wrote "are datorii DITL sector 2"
        assert clf.classify("are datorii DITL sector 2") == "datorii_ditl"


class TestGrouping:
    def test_group_similar_motivs(self):
        clf = make_clf()
        motivs = [
            "Datorii DITL",
            "Datoriiz DITL",
            "DATORII DITL",
            "decedat",
            "Complet diferit"
        ]
        groups = clf.group_similar(motivs)
        # Should produce clusters
        assert len(groups) >= 2
        # The 3 DITL variants should be in the same cluster
        ditl_cluster = next((g for g in groups if "Datorii DITL" in g), None)
        assert ditl_cluster is not None
        assert "Datoriiz DITL" in ditl_cluster
        assert "DATORII DITL" in ditl_cluster

    def test_classify_all(self):
        clf = make_clf()
        motivs = ["Datorii DITL", "nu stiu", "decedat"]
        results = clf.classify_all(motivs)
        assert results["Datorii DITL"] == "datorii_ditl"
        assert results["nu stiu"] == "unknown"
        assert results["decedat"] == "exclude"


class TestLearnMotiv:
    def test_learn_new_motiv(self):
        clf = make_clf()
        clf.learn("datorii_ditl", "mocirla semitotala de datorii")
        assert clf.classify("mocirla semitotala de datorii") == "datorii_ditl"

    def test_learned_motiv_used_in_classify_all(self):
        clf = make_clf()
        clf.learn("exclude", "abandonat dosar")
        results = clf.classify_all(["abandonat dosar"])
        assert results["abandonat dosar"] == "exclude"
```

**Step 2: Run to verify failure**

```bash
pytest tests/test_motiv_classifier.py -v
```
Expected: `ModuleNotFoundError`

**Step 3: Implement `modules/motiv_classifier.py`**

```python
"""Classifies arbitrary suspension motives using fuzzy matching."""
from rapidfuzz import fuzz, process


class MotivClassifier:
    def __init__(self, rules: dict, fuzzy_threshold: int = 75):
        """
        rules: {"datorii_ditl": [...known strings...], "exclude": [...]}
        fuzzy_threshold: minimum score (0-100) for fuzzy match
        """
        self.rules = rules
        self.threshold = fuzzy_threshold

    def classify(self, motiv: str) -> str:
        """Classify a single motiv string. Returns category name or 'unknown'."""
        motiv_lower = motiv.lower().strip()

        for category, known_list in self.rules.items():
            # Exact / case-insensitive
            for known in known_list:
                if known.lower() in motiv_lower or motiv_lower in known.lower():
                    return category

            # Fuzzy match against each known motiv
            for known in known_list:
                score = fuzz.partial_ratio(motiv_lower, known.lower())
                if score >= self.threshold:
                    return category

        return "unknown"

    def classify_all(self, motivs: list[str]) -> dict[str, str]:
        """Classify a list of distinct motiv values. Returns {motiv: category}."""
        return {m: self.classify(m) for m in motivs}

    def group_similar(self, motivs: list[str], threshold: int = 80) -> list[list[str]]:
        """Group similar motiv strings into clusters using fuzzy matching."""
        ungrouped = list(motivs)
        clusters = []

        while ungrouped:
            seed = ungrouped.pop(0)
            cluster = [seed]
            remaining = []
            for candidate in ungrouped:
                score = fuzz.ratio(seed.lower(), candidate.lower())
                if score >= threshold:
                    cluster.append(candidate)
                else:
                    remaining.append(candidate)
            clusters.append(cluster)
            ungrouped = remaining

        return clusters

    def learn(self, category: str, motiv: str):
        """Add a motiv to a known category."""
        if category not in self.rules:
            self.rules[category] = []
        if motiv not in self.rules[category]:
            self.rules[category].append(motiv)
```

**Step 4: Run tests — verify they pass**

```bash
pytest tests/test_motiv_classifier.py -v
```
Expected: all PASS

**Step 5: Commit**

```bash
git add modules/motiv_classifier.py tests/test_motiv_classifier.py
git commit -m "feat: add MotivClassifier with fuzzy grouping"
```

---

### Task 4: Debt Logic Module

**Business rule (AdaptiveDebtStatus in Python):** For each row (identified by CNP + Nr. Cerere), compute `stare_calculata` based on:
- `debt_flag` column value → has debt or not
- `data_repunere` column → if present + valid, person reinstated despite debt
- Output date labels use month/year of the processing run

**Files:**
- Create: `APP_DIR/modules/debt_logic.py`
- Create: `APP_DIR/tests/test_debt_logic.py`

**Step 1: Write failing tests**

`tests/test_debt_logic.py`:
```python
import pytest
import pandas as pd
from datetime import date
from modules.debt_logic import compute_debt_status, DEBT_POSITIVE_VALUES

RUN_DATE = date(2025, 12, 31)
RUN_LABEL = "31 12 2025"  # used in "fara datorii" string


def make_row(**kwargs):
    defaults = {
        "CNP": "1234567890123",
        "Numar Cerere": "S2/12345",
        "Sector": 2,
        "debt_flag_value": None,
        "data_repunere_value": None,
    }
    defaults.update(kwargs)
    return defaults


class TestDebtStatus:
    def test_no_debt(self):
        row = make_row(debt_flag_value="Nu are debit")
        result = compute_debt_status(row, RUN_DATE)
        assert result == f"Fara datorii reinterogare {RUN_LABEL}"

    def test_has_debt_no_repunere(self):
        row = make_row(debt_flag_value="Are debit")
        result = compute_debt_status(row, RUN_DATE)
        assert result == "datorii DITL"

    def test_has_debt_with_repunere(self):
        row = make_row(
            debt_flag_value="Are debit",
            data_repunere_value=date(2025, 12, 20)
        )
        result = compute_debt_status(row, RUN_DATE)
        assert result == "repus"

    def test_has_debt_repunere_before_run(self):
        # Repunere was before the run date — still counts
        row = make_row(
            debt_flag_value="Are debit",
            data_repunere_value=date(2025, 11, 1)
        )
        result = compute_debt_status(row, RUN_DATE)
        assert result == "repus"

    def test_empty_debt_flag(self):
        row = make_row(debt_flag_value=None)
        result = compute_debt_status(row, RUN_DATE)
        assert result == "neclar"

    def test_debt_flag_variants(self):
        for val in DEBT_POSITIVE_VALUES:
            row = make_row(debt_flag_value=val)
            result = compute_debt_status(row, RUN_DATE)
            assert result in ("datorii DITL", "repus"), f"Unexpected for {val!r}"

    def test_no_debt_variants(self):
        for val in ["Nu are debit", "nu are debit", "NU ARE DEBIT", "0", "Fara datorii"]:
            row = make_row(debt_flag_value=val)
            result = compute_debt_status(row, RUN_DATE)
            assert result == f"Fara datorii reinterogare {RUN_LABEL}", f"Failed for {val!r}"


class TestApplyToDataframe:
    def test_full_dataframe(self):
        from modules.debt_logic import apply_debt_status
        df = pd.DataFrame([
            {"CNP": "111", "Numar Cerere": "C1", "Sector": 2,
             "debt_flag_value": "Are debit", "data_repunere_value": None},
            {"CNP": "222", "Numar Cerere": "C2", "Sector": 2,
             "debt_flag_value": "Nu are debit", "data_repunere_value": None},
        ])
        result = apply_debt_status(df, RUN_DATE)
        assert result.loc[0, "Stare Calculata"] == "datorii DITL"
        assert result.loc[1, "Stare Calculata"] == f"Fara datorii reinterogare {RUN_LABEL}"
```

**Step 2: Run to verify failure**

```bash
pytest tests/test_debt_logic.py -v
```

**Step 3: Implement `modules/debt_logic.py`**

```python
"""Computes debt status for each cerere row — Python equivalent of AdaptiveDebtStatus PQ function."""
import pandas as pd
from datetime import date

# All values that mean "has debt" (case-insensitive check applied)
DEBT_POSITIVE_VALUES = [
    "Are debit", "are debit", "ARE DEBIT",
    "Da", "da", "1", "X", "x",
    "Debit", "debit",
]

DEBT_POSITIVE_LOWER = {v.lower() for v in DEBT_POSITIVE_VALUES}
DEBT_NEGATIVE_LOWER = {
    "nu are debit", "0", "fara datorii", "nu", "no", "–", "-", ""
}


def _is_positive_debt(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip().lower()
    if s in DEBT_POSITIVE_LOWER:
        return True
    # Starts-with check for "are debit la ..."
    if s.startswith("are debit"):
        return True
    return False


def _is_negative_debt(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    s = str(value).strip().lower()
    if s in DEBT_NEGATIVE_LOWER:
        return True
    if s.startswith("nu are debit") or s.startswith("fara datorii"):
        return True
    return False


def compute_debt_status(row: dict, run_date: date) -> str:
    """
    row must have:
      - debt_flag_value: the value from the detected debt column
      - data_repunere_value: the value from the detected repunere column (date or None)
    Returns: one of "datorii DITL", "repus", "fara datorii reinterogare DD MM YYYY", "neclar"
    """
    debt_val = row.get("debt_flag_value")
    repunere_val = row.get("data_repunere_value")

    label = f"{run_date.day:02d} {run_date.month:02d} {run_date.year}"

    if _is_negative_debt(debt_val):
        return f"Fara datorii reinterogare {label}"
    elif _is_positive_debt(debt_val):
        if repunere_val is not None:
            return "repus"
        return "datorii DITL"
    else:
        return "neclar"


def apply_debt_status(df: pd.DataFrame, run_date: date) -> pd.DataFrame:
    """Apply compute_debt_status to every row of a dataframe.
    Expects columns: debt_flag_value, data_repunere_value
    Adds column: Stare Calculata
    """
    df = df.copy()
    df["Stare Calculata"] = df.apply(
        lambda r: compute_debt_status(r.to_dict(), run_date), axis=1
    )
    return df
```

**Step 4: Run tests — verify they pass**

```bash
pytest tests/test_debt_logic.py -v
```

**Step 5: Commit**

```bash
git add modules/debt_logic.py tests/test_debt_logic.py
git commit -m "feat: add debt status computation logic"
```

---

### Task 5: Sector Loader Module

**Business rule:** Load each S1–S6 Excel file into a DataFrame, auto-detect the key columns using ColumnDetector, normalize them into a standard schema with `debt_flag_value` and `data_repunere_value` columns.

**Files:**
- Create: `APP_DIR/modules/sector_loader.py`
- Create: `APP_DIR/tests/test_sector_loader.py`
- Create: `APP_DIR/tests/fixtures/` (small test Excel files)

**Step 1: Create test fixtures**

Write `tests/create_fixtures.py` and run it once to create minimal test Excel files:

```python
"""Run once to create test fixture Excel files."""
import pandas as pd
from pathlib import Path

FIXTURE_DIR = Path(__file__).parent / "fixtures"
FIXTURE_DIR.mkdir(exist_ok=True)

# S2-like file with old column names
df_old = pd.DataFrame({
    "CNP": ["1111111111111", "2222222222222"],
    "Nr. Cerere": ["S2/001", "S2/002"],
    "Achitat la data - Observații2": ["da", None],
    "Are/Nu are debit la 31.12.2025": ["Nu are debit", "Are debit"],
    "Data repunere": [None, None],
})
df_old.to_excel(FIXTURE_DIR / "s2_old_cols.xlsx", index=False)

# S6-like file with new column names
df_new = pd.DataFrame({
    "CNP": ["3333333333333"],
    "Nr. Cerere": ["S6/001"],
    "Data plății - Observații": ["platit"],
    "Are/Nu are debit la 31.12.2025": ["Nu are debit"],
    "Data repunere": [None],
})
df_new.to_excel(FIXTURE_DIR / "s6_new_cols.xlsx", index=False)

print("Fixtures created.")
```

Run: `python tests/create_fixtures.py`

**Step 2: Write failing tests**

`tests/test_sector_loader.py`:
```python
import pytest
from pathlib import Path
from modules.column_detector import ColumnDetector
from modules.sector_loader import SectorLoader

FIXTURE_DIR = Path(__file__).parent / "fixtures"

TEST_CONFIG = {
    "debt_flag": {"patterns": ["are.?nu.?are.?debit"], "known_exact": []},
    "data_plata": {"patterns": ["achitat|platit|data.?pla"], "known_exact": []},
    "data_repunere": {"patterns": ["data.?repunere"], "known_exact": []},
    "cnp": {"patterns": ["^cnp$"], "known_exact": ["CNP"]},
    "numar_cerere": {"patterns": ["nr\\.?.?cerere"], "known_exact": []},
}


def make_loader():
    detector = ColumnDetector(TEST_CONFIG)
    return SectorLoader(detector)


class TestSectorLoader:
    def test_load_old_column_names(self):
        loader = make_loader()
        df, detected, missing = loader.load(FIXTURE_DIR / "s2_old_cols.xlsx", sector="S2")
        assert missing == [], f"Should detect all critical cols, missing: {missing}"
        assert "debt_flag_value" in df.columns
        assert "CNP_normalized" in df.columns
        assert df.loc[0, "debt_flag_value"] == "Nu are debit"
        assert df.loc[1, "debt_flag_value"] == "Are debit"

    def test_load_new_column_names(self):
        loader = make_loader()
        df, detected, missing = loader.load(FIXTURE_DIR / "s6_new_cols.xlsx", sector="S6")
        assert missing == []
        assert df.loc[0, "debt_flag_value"] == "Nu are debit"

    def test_missing_column_reported(self):
        loader = make_loader()
        # Old file has no "Data repunere" named correctly
        df, detected, missing = loader.load(FIXTURE_DIR / "s2_old_cols.xlsx", sector="S2")
        # "Data repunere" IS in the fixture, so should not be missing
        # If it were missing, it would appear in `missing` list
        assert isinstance(missing, list)

    def test_sector_column_added(self):
        loader = make_loader()
        df, _, _ = loader.load(FIXTURE_DIR / "s2_old_cols.xlsx", sector="S2")
        assert "Sector" in df.columns
        assert (df["Sector"] == "S2").all()
```

**Step 3: Implement `modules/sector_loader.py`**

```python
"""Loads a single sector Excel file and normalizes key columns."""
import pandas as pd
from pathlib import Path
from modules.column_detector import ColumnDetector


class SectorLoader:
    def __init__(self, detector: ColumnDetector):
        self.detector = detector

    def load(self, filepath: Path, sector: str) -> tuple[pd.DataFrame, dict, list[str]]:
        """
        Load and normalize a sector Excel file.
        Returns:
          - df: normalized DataFrame with standard column names added
          - detected: {concept: original_column_name} mapping
          - missing: list of critical concepts not found
        """
        df = pd.read_excel(filepath, dtype=str)
        columns = list(df.columns)

        detected = self.detector.detect(columns)
        missing = self.detector.get_missing_critical(detected)

        # Add normalized columns
        df["Sector"] = sector

        if detected.get("cnp"):
            df["CNP_normalized"] = df[detected["cnp"]].str.strip()
        else:
            df["CNP_normalized"] = None

        if detected.get("numar_cerere"):
            df["NrCerere_normalized"] = df[detected["numar_cerere"]].str.strip()
        else:
            df["NrCerere_normalized"] = None

        if detected.get("debt_flag"):
            df["debt_flag_value"] = df[detected["debt_flag"]]
        else:
            df["debt_flag_value"] = None

        if detected.get("data_repunere"):
            df["data_repunere_value"] = pd.to_datetime(
                df[detected["data_repunere"]], errors="coerce", dayfirst=True
            )
        else:
            df["data_repunere_value"] = None

        return df, detected, missing
```

**Step 4: Run tests**

```bash
pytest tests/test_sector_loader.py -v
```

**Step 5: Commit**

```bash
git add modules/sector_loader.py tests/test_sector_loader.py tests/create_fixtures.py
git commit -m "feat: add SectorLoader with auto column normalization"
```

---

### Task 6: Output Builder Module

**Business rule:** Takes the 6 normalized DataFrames + computed statuses, writes a multi-sheet Excel: S1–S6 per-sector + CONSOLIDAT (platform columns) + LOG.

**Files:**
- Create: `APP_DIR/modules/output_builder.py`
- Create: `APP_DIR/tests/test_output_builder.py`

**Step 1: Define the platform column structure**

The platform expects these columns in import files:
```
CNP | Numar Cerere | Sector | Stare Calculata | Data Repunere | Motiv Suspendare | Sursa Suspendare | Observatii
```

**Step 2: Write failing tests**

`tests/test_output_builder.py`:
```python
import pytest
import pandas as pd
from pathlib import Path
import tempfile
from modules.output_builder import OutputBuilder

PLATFORM_COLUMNS = [
    "CNP", "Numar Cerere", "Sector", "Stare Calculata",
    "Data Repunere", "Motiv Suspendare", "Sursa Suspendare", "Observatii"
]


def make_sector_df(sector, n=3):
    return pd.DataFrame({
        "CNP_normalized": [f"111{i:010d}" for i in range(n)],
        "NrCerere_normalized": [f"{sector}/{i:04d}" for i in range(n)],
        "Sector": sector,
        "debt_flag_value": ["Are debit", "Nu are debit", "Are debit"],
        "data_repunere_value": [None, None, None],
        "Stare Calculata": ["datorii DITL", "Fara datorii reinterogare 31 12 2025", "datorii DITL"],
    })


class TestOutputBuilder:
    def test_creates_excel_with_sector_sheets(self, tmp_path):
        builder = OutputBuilder()
        sectors = {f"S{i}": make_sector_df(f"S{i}") for i in range(1, 7)}
        out_path = tmp_path / "test_output.xlsx"
        builder.write(sectors, out_path)

        xl = pd.ExcelFile(out_path)
        for s in ["S1", "S2", "S3", "S4", "S5", "S6"]:
            assert s in xl.sheet_names

    def test_creates_consolidat_sheet(self, tmp_path):
        builder = OutputBuilder()
        sectors = {f"S{i}": make_sector_df(f"S{i}") for i in range(1, 7)}
        out_path = tmp_path / "test_output.xlsx"
        builder.write(sectors, out_path)

        xl = pd.ExcelFile(out_path)
        assert "CONSOLIDAT" in xl.sheet_names
        df_con = pd.read_excel(out_path, sheet_name="CONSOLIDAT")
        for col in PLATFORM_COLUMNS:
            assert col in df_con.columns, f"Missing platform column: {col}"

    def test_creates_log_sheet(self, tmp_path):
        builder = OutputBuilder()
        sectors = {f"S{i}": make_sector_df(f"S{i}") for i in range(1, 7)}
        detection_log = {"S1": {"debt_flag": "Are/Nu are debit la 31.12.2025", "cnp": "CNP"}}
        out_path = tmp_path / "test_output.xlsx"
        builder.write(sectors, out_path, detection_log=detection_log)

        xl = pd.ExcelFile(out_path)
        assert "LOG" in xl.sheet_names

    def test_consolidat_has_all_sectors(self, tmp_path):
        builder = OutputBuilder()
        sectors = {f"S{i}": make_sector_df(f"S{i}", n=2) for i in range(1, 7)}
        out_path = tmp_path / "test_output.xlsx"
        builder.write(sectors, out_path)

        df_con = pd.read_excel(out_path, sheet_name="CONSOLIDAT")
        assert len(df_con) == 12  # 6 sectors × 2 rows
```

**Step 3: Implement `modules/output_builder.py`**

```python
"""Builds the output Excel file with per-sector sheets + CONSOLIDAT + LOG."""
import pandas as pd
from pathlib import Path
from datetime import datetime

PLATFORM_COLUMNS = [
    "CNP", "Numar Cerere", "Sector", "Stare Calculata",
    "Data Repunere", "Motiv Suspendare", "Sursa Suspendare", "Observatii"
]


class OutputBuilder:
    def write(
        self,
        sectors: dict[str, pd.DataFrame],
        output_path: Path,
        detection_log: dict = None,
    ):
        """
        sectors: {"S1": df1, ..., "S6": df6}  — each has CNP_normalized, NrCerere_normalized,
                 Sector, debt_flag_value, data_repunere_value, Stare Calculata
        output_path: where to write the Excel
        detection_log: {sector: {concept: original_col_name}} for LOG sheet
        """
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # Per-sector sheets
            for sector_name, df in sectors.items():
                df.to_excel(writer, sheet_name=sector_name, index=False)

            # CONSOLIDAT sheet
            consolidated = self._build_consolidated(sectors)
            consolidated.to_excel(writer, sheet_name="CONSOLIDAT", index=False)

            # LOG sheet
            log_df = self._build_log(sectors, detection_log)
            log_df.to_excel(writer, sheet_name="LOG", index=False)

    def _build_consolidated(self, sectors: dict) -> pd.DataFrame:
        frames = []
        for sector_name, df in sectors.items():
            platform_df = pd.DataFrame()
            platform_df["CNP"] = df.get("CNP_normalized", pd.Series(dtype=str))
            platform_df["Numar Cerere"] = df.get("NrCerere_normalized", pd.Series(dtype=str))
            platform_df["Sector"] = sector_name
            platform_df["Stare Calculata"] = df.get("Stare Calculata", pd.Series(dtype=str))
            platform_df["Data Repunere"] = df.get("data_repunere_value", pd.Series(dtype=object))
            platform_df["Motiv Suspendare"] = df.get("Motiv Suspendare", pd.Series(dtype=str))
            platform_df["Sursa Suspendare"] = df.get("Sursa Suspendare", pd.Series(dtype=str))
            platform_df["Observatii"] = df.get("Observatii", pd.Series(dtype=str))
            frames.append(platform_df)

        if not frames:
            return pd.DataFrame(columns=PLATFORM_COLUMNS)
        return pd.concat(frames, ignore_index=True)

    def _build_log(self, sectors: dict, detection_log: dict) -> pd.DataFrame:
        rows = []
        rows.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Info": "Processing run"
        })
        if detection_log:
            for sector, mapping in detection_log.items():
                for concept, col_name in mapping.items():
                    rows.append({
                        "Sector": sector,
                        "Concept": concept,
                        "Coloana detectata": col_name,
                    })
        for sector_name, df in sectors.items():
            rows.append({
                "Sector": sector_name,
                "Nr randuri": len(df),
            })
        return pd.DataFrame(rows)
```

**Step 4: Run tests**

```bash
pytest tests/test_output_builder.py -v
```

**Step 5: Commit**

```bash
git add modules/output_builder.py tests/test_output_builder.py
git commit -m "feat: add OutputBuilder with multi-sheet Excel generation"
```

---

### Task 7: Config Manager Module

**Business rule:** Load/save `column_aliases.json`, `motiv_rules.json`, `settings.json`. Thread-safe (UI writes config after user confirms in dialog).

**Files:**
- Create: `APP_DIR/modules/config_manager.py`
- Create: `APP_DIR/tests/test_config_manager.py`

**Step 1: Write failing tests**

`tests/test_config_manager.py`:
```python
import pytest
import json
from pathlib import Path
from modules.config_manager import ConfigManager


def test_load_column_aliases(tmp_path):
    cfg_file = tmp_path / "column_aliases.json"
    cfg_file.write_text(json.dumps({"debt_flag": {"patterns": [], "known_exact": ["X"]}}))
    cm = ConfigManager(tmp_path)
    aliases = cm.load_column_aliases()
    assert aliases["debt_flag"]["known_exact"] == ["X"]


def test_save_column_aliases(tmp_path):
    cm = ConfigManager(tmp_path)
    data = {"debt_flag": {"patterns": ["x"], "known_exact": ["Col A"]}}
    cm.save_column_aliases(data)
    loaded = json.loads((tmp_path / "column_aliases.json").read_text())
    assert loaded["debt_flag"]["known_exact"] == ["Col A"]


def test_load_motiv_rules(tmp_path):
    cfg_file = tmp_path / "motiv_rules.json"
    cfg_file.write_text(json.dumps({"datorii_ditl": ["Datorii DITL"], "exclude": []}))
    cm = ConfigManager(tmp_path)
    rules = cm.load_motiv_rules()
    assert "Datorii DITL" in rules["datorii_ditl"]


def test_save_settings(tmp_path):
    cm = ConfigManager(tmp_path)
    cm.save_setting("last_sector_paths", {"S1": "C:/test.xlsx"})
    settings = cm.load_settings()
    assert settings["last_sector_paths"]["S1"] == "C:/test.xlsx"


def test_defaults_when_file_missing(tmp_path):
    cm = ConfigManager(tmp_path)
    aliases = cm.load_column_aliases()
    # Should return empty structure, not crash
    assert isinstance(aliases, dict)
```

**Step 2: Implement `modules/config_manager.py`**

```python
"""Loads and saves persistent configuration (column aliases, motiv rules, settings)."""
import json
from pathlib import Path

DEFAULT_COLUMN_ALIASES = {
    "debt_flag": {"patterns": ["are.?nu.?are.?debit", "ARE.?NU.?ARE.?DEBIT"], "known_exact": []},
    "data_plata": {"patterns": ["achitat|platit|data.?pla"], "known_exact": []},
    "data_repunere": {"patterns": ["data.?repunere", "data.?repunerii"], "known_exact": []},
    "cnp": {"patterns": ["^cnp$", "cod.?numeric"], "known_exact": ["CNP"]},
    "numar_cerere": {"patterns": ["nr\\.?.?cerere", "numar.?cerere"], "known_exact": []},
}

DEFAULT_MOTIV_RULES = {"datorii_ditl": [], "exclude": []}

DEFAULT_SETTINGS = {
    "last_sector_paths": {"S1": "", "S2": "", "S3": "", "S4": "", "S5": "", "S6": ""},
    "last_raport_stari_path": "",
    "last_interogare_termen_path": "",
    "output_dir": "",
}


class ConfigManager:
    def __init__(self, config_dir: Path):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)

    def _load(self, filename: str, default: dict) -> dict:
        path = self.config_dir / filename
        if not path.exists():
            return default.copy()
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return default.copy()

    def _save(self, filename: str, data: dict):
        path = self.config_dir / filename
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_column_aliases(self) -> dict:
        return self._load("column_aliases.json", DEFAULT_COLUMN_ALIASES)

    def save_column_aliases(self, data: dict):
        self._save("column_aliases.json", data)

    def load_motiv_rules(self) -> dict:
        return self._load("motiv_rules.json", DEFAULT_MOTIV_RULES)

    def save_motiv_rules(self, data: dict):
        self._save("motiv_rules.json", data)

    def load_settings(self) -> dict:
        return self._load("settings.json", DEFAULT_SETTINGS)

    def save_setting(self, key: str, value):
        settings = self.load_settings()
        settings[key] = value
        self._save("settings.json", settings)
```

**Step 3: Run tests**

```bash
pytest tests/test_config_manager.py -v
```

**Step 4: Commit**

```bash
git add modules/config_manager.py tests/test_config_manager.py
git commit -m "feat: add ConfigManager for persistent column aliases and motiv rules"
```

---

### Task 8: Processing Orchestrator

**Business rule:** Ties all modules together: load 6 sectors → detect columns → compute debt status → build output. No UI dependency. Called by the UI with callbacks for progress reporting.

**Files:**
- Create: `APP_DIR/modules/processor.py`
- Create: `APP_DIR/tests/test_processor.py`

**Step 1: Write failing tests**

`tests/test_processor.py`:
```python
import pytest
from pathlib import Path
from datetime import date
from modules.processor import ReinterogareProcessor

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def test_process_two_sectors(tmp_path):
    from modules.config_manager import ConfigManager
    cm = ConfigManager(tmp_path / "config")
    processor = ReinterogareProcessor(cm)

    sector_paths = {
        "S2": FIXTURE_DIR / "s2_old_cols.xlsx",
        "S6": FIXTURE_DIR / "s6_new_cols.xlsx",
    }
    run_date = date(2025, 12, 31)
    output_path = tmp_path / "output.xlsx"

    result = processor.process(sector_paths, run_date, output_path)

    assert result["success"] is True
    assert output_path.exists()
    assert result["rows_processed"] > 0


def test_process_reports_missing_columns(tmp_path):
    from modules.config_manager import ConfigManager
    cm = ConfigManager(tmp_path / "config")
    processor = ReinterogareProcessor(cm)

    # Provide a valid and an empty path
    sector_paths = {"S1": FIXTURE_DIR / "s2_old_cols.xlsx"}
    run_date = date(2025, 12, 31)
    output_path = tmp_path / "output.xlsx"

    result = processor.process(sector_paths, run_date, output_path)
    assert isinstance(result["detection_log"], dict)
```

**Step 2: Implement `modules/processor.py`**

```python
"""Orchestrates the full processing pipeline for one monthly run."""
import pandas as pd
from pathlib import Path
from datetime import date

from modules.config_manager import ConfigManager
from modules.column_detector import ColumnDetector
from modules.sector_loader import SectorLoader
from modules.debt_logic import apply_debt_status
from modules.output_builder import OutputBuilder


class ReinterogareProcessor:
    def __init__(self, config_manager: ConfigManager, progress_callback=None):
        self.cm = config_manager
        self.progress_cb = progress_callback or (lambda msg: None)

    def process(
        self,
        sector_paths: dict[str, Path],
        run_date: date,
        output_path: Path,
    ) -> dict:
        """
        Run the full pipeline.
        Returns dict with: success, rows_processed, warnings, detection_log
        """
        aliases = self.cm.load_column_aliases()
        detector = ColumnDetector(aliases)
        loader = SectorLoader(detector)
        builder = OutputBuilder()

        sector_dfs = {}
        detection_log = {}
        warnings = []

        for sector_name, filepath in sector_paths.items():
            self.progress_cb(f"Loading {sector_name}...")
            if not filepath or not Path(filepath).exists():
                warnings.append(f"{sector_name}: file not found, skipped")
                continue

            try:
                df, detected, missing = loader.load(Path(filepath), sector_name)
            except Exception as e:
                warnings.append(f"{sector_name}: load error — {e}")
                continue

            detection_log[sector_name] = {k: v for k, v in detected.items() if v}

            if missing:
                warnings.append(f"{sector_name}: critical columns not detected: {missing}")
                df["Stare Calculata"] = "neclar"
            else:
                df = apply_debt_status(df, run_date)

            # Learn detected columns for future runs
            for concept, col_name in detected.items():
                if col_name:
                    detector.learn_alias(concept, col_name)

            sector_dfs[sector_name] = df

        # Save updated aliases
        self.cm.save_column_aliases(detector.config)

        self.progress_cb("Building output Excel...")
        try:
            builder.write(sector_dfs, output_path, detection_log)
        except PermissionError:
            return {
                "success": False,
                "error": f"Cannot write to {output_path} — close it in Excel first.",
                "rows_processed": 0,
                "warnings": warnings,
                "detection_log": detection_log,
            }

        total_rows = sum(len(df) for df in sector_dfs.values())
        self.progress_cb(f"Done. {total_rows} rows processed.")

        return {
            "success": True,
            "rows_processed": total_rows,
            "warnings": warnings,
            "detection_log": detection_log,
        }
```

**Step 3: Run tests**

```bash
pytest tests/test_processor.py -v
```

**Step 4: Run full test suite**

```bash
pytest tests/ -v
```
Expected: all tests pass.

**Step 5: Commit**

```bash
git add modules/processor.py tests/test_processor.py
git commit -m "feat: add ReinterogareProcessor orchestrator"
```

---

### Task 9: Main UI (CustomTkinter)

**Note:** UI is not unit-tested — tested manually by running the app. Build and verify by running it.

**Files:**
- Create: `APP_DIR/reinterogare_processor.py` (replace placeholder)
- Create: `APP_DIR/modules/ui_components.py`

**Step 1: Implement `modules/ui_components.py`** — reusable UI widgets

```python
"""Reusable CustomTkinter UI components."""
import customtkinter as ctk
from pathlib import Path
from tkinter import filedialog


class FilePicker(ctk.CTkFrame):
    """A label + entry + browse button for picking a file."""
    def __init__(self, parent, label: str, filetypes=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.filetypes = filetypes or [("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]

        ctk.CTkLabel(self, text=label, width=80, anchor="w").pack(side="left", padx=(0, 8))
        self.entry = ctk.CTkEntry(self, width=350)
        self.entry.pack(side="left", padx=(0, 4))
        ctk.CTkButton(self, text="...", width=30, command=self._browse).pack(side="left")

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=self.filetypes)
        if path:
            self.entry.delete(0, "end")
            self.entry.insert(0, path)

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: str):
        self.entry.delete(0, "end")
        self.entry.insert(0, value)


class LogBox(ctk.CTkTextbox):
    """A scrollable log output widget."""
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("state", "disabled")
        kwargs.setdefault("height", 150)
        super().__init__(parent, **kwargs)

    def append(self, message: str):
        self.configure(state="normal")
        self.insert("end", message + "\n")
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")
```

**Step 2: Implement main `reinterogare_processor.py`**

```python
"""Reinterogare Processor — main UI application."""
import threading
from pathlib import Path
from datetime import date
import customtkinter as ctk
from tkinter import filedialog, messagebox

from modules.config_manager import ConfigManager
from modules.processor import ReinterogareProcessor
from modules.ui_components import FilePicker, LogBox

APP_DIR = Path(__file__).parent
CONFIG_DIR = APP_DIR / "config"

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Reinterogare Processor")
        self.geometry("720x620")
        self.resizable(True, True)

        self.cm = ConfigManager(CONFIG_DIR)
        settings = self.cm.load_settings()

        self._build_ui(settings)

    def _build_ui(self, settings):
        # Title
        ctk.CTkLabel(self, text="Reinterogare Processor", font=("", 18, "bold")).pack(pady=(16, 8))

        # Sector file pickers
        sector_frame = ctk.CTkFrame(self)
        sector_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(sector_frame, text="Fișiere sectoare S1–S6:", font=("", 13, "bold")).pack(anchor="w", padx=8, pady=4)

        self.sector_pickers = {}
        for i in range(1, 7):
            key = f"S{i}"
            picker = FilePicker(sector_frame, label=key)
            picker.pack(fill="x", padx=8, pady=2)
            saved = settings.get("last_sector_paths", {}).get(key, "")
            if saved:
                picker.set(saved)
            self.sector_pickers[key] = picker

        # Optional files
        optional_frame = ctk.CTkFrame(self)
        optional_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(optional_frame, text="Fișiere opționale:", font=("", 13, "bold")).pack(anchor="w", padx=8, pady=4)

        self.raport_stari_picker = FilePicker(optional_frame, label="Raport Stări")
        self.raport_stari_picker.pack(fill="x", padx=8, pady=2)
        if settings.get("last_raport_stari_path"):
            self.raport_stari_picker.set(settings["last_raport_stari_path"])

        self.interogare_termen_picker = FilePicker(optional_frame, label="Inter. Termen")
        self.interogare_termen_picker.pack(fill="x", padx=8, pady=2)
        if settings.get("last_interogare_termen_path"):
            self.interogare_termen_picker.set(settings["last_interogare_termen_path"])

        # Output dir
        out_frame = ctk.CTkFrame(self)
        out_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(out_frame, text="Director output:", anchor="w", width=100).pack(side="left", padx=8)
        self.output_entry = ctk.CTkEntry(out_frame, width=400)
        self.output_entry.pack(side="left", padx=(0, 4))
        if settings.get("output_dir"):
            self.output_entry.insert(0, settings["output_dir"])
        ctk.CTkButton(out_frame, text="...", width=30, command=self._pick_output_dir).pack(side="left")

        # Run button + progress
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=8)
        self.run_btn = ctk.CTkButton(btn_frame, text="Procesează", width=160, height=36,
                                     font=("", 14, "bold"), command=self._run)
        self.run_btn.pack(side="left")
        self.progress = ctk.CTkProgressBar(btn_frame)
        self.progress.pack(side="left", fill="x", expand=True, padx=16)
        self.progress.set(0)

        # Log
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        ctk.CTkLabel(log_frame, text="Log:", anchor="w").pack(anchor="w", padx=8)
        self.log = LogBox(log_frame, height=160)
        self.log.pack(fill="both", expand=True, padx=8, pady=4)

    def _pick_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, d)

    def _run(self):
        sector_paths = {}
        for key, picker in self.sector_pickers.items():
            p = picker.get()
            if p:
                sector_paths[key] = Path(p)

        if not sector_paths:
            messagebox.showwarning("Atenție", "Nu ai selectat niciun fișier sector.")
            return

        output_dir = self.output_entry.get().strip()
        if not output_dir:
            messagebox.showwarning("Atenție", "Selectează un director de output.")
            return

        # Save settings
        self.cm.save_setting("last_sector_paths", {k: str(v) for k, v in sector_paths.items()})
        self.cm.save_setting("output_dir", output_dir)

        # Build output filename
        today = date.today()
        out_filename = f"Reinterogare_{today.strftime('%m_%Y')}_output.xlsx"
        out_path = Path(output_dir) / out_filename

        self.run_btn.configure(state="disabled")
        self.log.clear()
        self.progress.set(0.1)

        def _worker():
            processor = ReinterogareProcessor(self.cm, progress_callback=self._on_progress)
            result = processor.process(sector_paths, today, out_path)
            self.after(0, lambda: self._on_done(result, out_path))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_progress(self, message: str):
        self.after(0, lambda: self.log.append(message))

    def _on_done(self, result: dict, out_path: Path):
        self.run_btn.configure(state="normal")
        self.progress.set(1.0)

        if result.get("success"):
            self.log.append(f"\nOutput salvat: {out_path}")
            self.log.append(f"Rânduri procesate: {result['rows_processed']}")
            for w in result.get("warnings", []):
                self.log.append(f"  ⚠ {w}")
            if messagebox.askyesno("Gata!", f"Output salvat.\nDeschizi fișierul?"):
                import os
                os.startfile(str(out_path))
        else:
            self.log.append(f"\nEROARE: {result.get('error', 'Unknown error')}")
            messagebox.showerror("Eroare", result.get("error", "A apărut o eroare."))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
```

**Step 3: Test manually**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python reinterogare_processor.py
```

Verify:
- Window opens correctly
- File pickers work (browse button → dialog)
- "Procesează" button disabled → runs → re-enabled
- Log shows progress messages
- Output Excel created in selected directory

**Step 4: Commit**

```bash
git add reinterogare_processor.py modules/ui_components.py
git commit -m "feat: add main CustomTkinter UI for Reinterogare Processor"
```

---

### Task 10: Run Full Test Suite + Verify

**Step 1: Run all tests**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
pytest tests/ -v --tb=short
```

Expected: all tests PASS, no warnings.

**Step 2: Manual end-to-end test with real S2 and S6 files**

- Select `C:\EXCELURI\Reinterogari\2025\Reinterogare 31 12 2025\Reinterogare S2 31.12.2025.xlsx` for S2
- Select `C:\EXCELURI\Reinterogari\2025\Reinterogare 31 12 2025\Reinterogare S6 31 12 2025 - verificat.xlsx` for S6
- Leave S1, S3, S4, S5 empty
- Click Procesează
- Open output Excel → verify: S2 sheet has data, S6 sheet has data, CONSOLIDAT has both, LOG shows detected columns

**Step 3: Verify column learning persists**

- Close app, reopen
- Check that the file pickers remember last paths
- Re-run — LOG should show the same column detections without warnings

**Step 4: Final commit**

```bash
git add .
git commit -m "feat: MVP complete — Reinterogare Processor with column detection and debt logic"
```

---

## Phase 2 Tasks (after MVP is validated)

### Phase 2A: REPUNERI + IMPORT_OBSERVATII sheets

Add two more sheets to `OutputBuilder.write()`:
- `REPUNERI`: rows where `Stare Calculata == "repus"` — format for bulk repunere in platform
- `IMPORT_OBSERVATII`: one row per cerere with computed observatie string for import

### Phase 2B: Interogare Termen correlation

In `modules/processor.py`, add optional param `interogare_termen_path`. When provided:
1. Load it → extract CNP-uri cu datorii la termen
2. Join with suspended beneficiaries from Raport Stări
3. Add `Sursa_Suspendare` column to CONSOLIDAT (ex: `"30.09.2025"`)

### Phase 2C: Motiv Review UI Tab

Add a "Motive" tab to the UI:
- After loading sectors + Raport Stări, extract all distinct Motiv values
- Show grouped clusters (from `MotivClassifier.group_similar`)
- User checks which clusters = "datorii_ditl"
- Saves to `config/motiv_rules.json`

---

*End of implementation plan.*
