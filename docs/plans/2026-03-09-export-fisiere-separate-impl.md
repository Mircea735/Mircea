# Export Fișiere Separate — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** La fiecare procesare, salvează automat `Repunere.xlsx` și `Import Observatii.xlsx` în output_dir; la verificare, salvează `Verificare Reinterogare {DD.MM.YYYY}.xlsx` în același folder.

**Architecture:** Metodă nouă `write_separate_files` în `OutputBuilder`; `write()` returnează DataFrames pentru a evita recalculul; `processor.process()` apelează `write_separate_files`; metodă nouă `write_verification_file` apelată din `_run_verificare`.

**Tech Stack:** Python 3.10+, pandas, openpyxl

**APP_DIR**: `C:\EXCELURI\Reinterogari\Reinterogare Processor\`

---

## Referințe critice

- `modules/output_builder.py` — `OutputBuilder.write()`, `append_verification()`
- `modules/processor.py:162` — apelul `builder.write(...)`, return dict
- `reinterogare_processor.py:628` — `_run_verificare()`
- Output filename pattern: `Reinterogare_31_01_2026_143022_output.xlsx` → dată = `\d{2}_\d{2}_\d{4}`

---

## Task 1: `write_separate_files` în `output_builder.py`

**Files:**
- Modify: `APP_DIR/modules/output_builder.py`
- Modify: `APP_DIR/tests/test_output_builder.py`

**Step 1: Scrie testul**

```python
# Adaugă în tests/test_output_builder.py
import re
from datetime import date

def test_write_separate_files_creates_xlsx(tmp_path):
    import pandas as pd
    from modules.output_builder import OutputBuilder

    rep_df = pd.DataFrame([{
        "Nr. cerere": "1",
        "Nume adult cu handicap": "Popescu",
        "Prenume adult cu handicap": "Ion",
        "CNP adult cu handicap": "123",
        "Motiv repunere": "Fara datorii",
        "Data repunere": "2026-01-01",
        "Data cererii": "2024-01-01",
    }])
    imp_df = pd.DataFrame([{
        "Nr. cerere": "2",
        "Nume adult cu handicap": "Ionescu",
        "Prenume adult cu handicap": "Maria",
        "CNP adult cu handicap": "456",
        "Observatii": "Datorii reinterogare S2",
    }])

    run_date = date(2026, 1, 31)
    OutputBuilder().write_separate_files(tmp_path, run_date, rep_df, imp_df)

    assert (tmp_path / "Repunere.xlsx").exists()
    assert (tmp_path / "Import Observatii.xlsx").exists()

    rep_result = pd.read_excel(tmp_path / "Repunere.xlsx")
    assert len(rep_result) == 1
    assert "Nr. cerere" in rep_result.columns

    imp_result = pd.read_excel(tmp_path / "Import Observatii.xlsx")
    assert len(imp_result) == 1
    assert "Observatii" in imp_result.columns
```

**Step 2: Rulează să confirmi FAIL**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -m pytest tests/test_output_builder.py::test_write_separate_files_creates_xlsx -v
```

Expected: `AttributeError: 'OutputBuilder' object has no attribute 'write_separate_files'`

**Step 3: Adaugă `write_separate_files` în `output_builder.py`**

Adaugă la sfârșitul clasei `OutputBuilder` (după `_build_log`):

```python
def write_separate_files(
    self,
    output_dir,
    run_date,
    repunere_df: pd.DataFrame,
    import_obs_df: pd.DataFrame,
):
    """Scrie Repunere.xlsx și Import Observatii.xlsx în output_dir."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        rep_cols = [c for c in REPUNERE_COLUMNS if c in repunere_df.columns]
        rep_out = repunere_df[rep_cols].copy() if rep_cols else pd.DataFrame(columns=REPUNERE_COLUMNS)
        rep_out.to_excel(output_dir / "Repunere.xlsx", index=False)
    except Exception as e:
        print(f"[warn] Nu s-a putut scrie Repunere.xlsx: {e}")

    try:
        imp_cols = [c for c in IMPORT_OBS_COLUMNS if c in import_obs_df.columns]
        imp_out = import_obs_df[imp_cols].copy() if imp_cols else pd.DataFrame(columns=IMPORT_OBS_COLUMNS)
        imp_out.to_excel(output_dir / "Import Observatii.xlsx", index=False)
    except Exception as e:
        print(f"[warn] Nu s-a putut scrie Import Observatii.xlsx: {e}")
```

**Step 4: Rulează testul**

```bash
python -m pytest tests/test_output_builder.py::test_write_separate_files_creates_xlsx -v
```

Expected: PASS

**Step 5: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 6: Commit**

```bash
git add modules/output_builder.py tests/test_output_builder.py
git commit -m "feat: add write_separate_files to OutputBuilder"
```

---

## Task 2: Wire `write_separate_files` în pipeline

**Files:**
- Modify: `APP_DIR/modules/output_builder.py` — `write()` returnează DataFrames
- Modify: `APP_DIR/modules/processor.py` — apelează `write_separate_files` după `write()`

**Step 1: Modifică `write()` să returneze DataFrames**

În `output_builder.py`, în metoda `write()`, după ce calculezi `import_df` și `repunere_df` (liniile cu `consolidat_df.rename` și filtrele de Status), adaugă variabile locale și returnează-le la final:

```python
# La ÎNCEPUTUL blocului if consolidat_df is not None:
_repunere_out = pd.DataFrame(columns=REPUNERE_COLUMNS)
_import_out   = pd.DataFrame(columns=IMPORT_OBS_COLUMNS)
```

Înlocuiește liniile existente de calcul cu:
```python
df = consolidat_df.rename(columns=_RENAME)
import_df = df[
    (df["Status"] == STATUS_IMPORT_OBSERVATII) &
    (df["Observatii"].str.startswith("Datorii", na=False))
]
repunere_df = df[df["Status"] == STATUS_REPUNERE]
_import_out   = import_df
_repunere_out = repunere_df
```

La finalul metodei `write()` (după `log_df.to_excel`), adaugă:
```python
return _repunere_out, _import_out
```

Dacă consolidat_df lipsește, în ramura `else` inițializează `_repunere_out` și `_import_out` ca DataFrames goale și returnează-le la fel.

**Step 2: Scrie testul pentru return**

```python
# În test_output_builder.py
def test_write_returns_dataframes(tmp_path):
    import pandas as pd
    from modules.output_builder import OutputBuilder

    result = OutputBuilder().write(
        sector_dfs={},
        output_path=tmp_path / "out.xlsx",
    )
    assert result is not None
    assert len(result) == 2
    rep, imp = result
    assert isinstance(rep, pd.DataFrame)
    assert isinstance(imp, pd.DataFrame)
```

**Step 3: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_output_builder.py::test_write_returns_dataframes -v
```

Expected: FAIL (write() returnează None)

**Step 4: Implementează modificarea în `write()`**

Modifică `write()` conform Step 1. Structura finală `write()`:

```python
def write(self, sector_dfs, output_path, detection_log=None, motiv_data=None, consolidat_df=None):
    _repunere_out = pd.DataFrame(columns=REPUNERE_COLUMNS)
    _import_out   = pd.DataFrame(columns=IMPORT_OBS_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if consolidat_df is not None and len(consolidat_df) > 0:
            df = consolidat_df.rename(columns=_RENAME)
            import_df = df[
                (df["Status"] == STATUS_IMPORT_OBSERVATII) &
                (df["Observatii"].str.startswith("Datorii", na=False))
            ]
            repunere_df = df[df["Status"] == STATUS_REPUNERE]
            _import_out   = import_df
            _repunere_out = repunere_df
            self._write_import_obs(writer, import_df)
            self._write_repunere(writer, repunere_df)
            self._write_motive_repunere(writer, repunere_df)
            self._write_consolidat(writer, df)
        else:
            pd.DataFrame(columns=IMPORT_OBS_COLUMNS).to_excel(
                writer, sheet_name="IMPORT OBSERVATII", index=False)
            pd.DataFrame(columns=REPUNERE_COLUMNS).to_excel(
                writer, sheet_name="REPUNERE", index=False)

        motiv_data = motiv_data or {}
        self._write_motiv_sheet(writer, "MOTIVE INCLUSE", motiv_data.get("datorii_ditl", []))
        self._write_motiv_sheet(writer, "MOTIVE EXCLUSE", motiv_data.get("exclude", []))
        self._write_motiv_sheet(writer, "MOTIVE NECUNOSCUTE", motiv_data.get("unknown", []))

        consolidat_cnps = set()
        if consolidat_df is not None and "Cnp Adult Handicap" in consolidat_df.columns:
            consolidat_cnps = {
                str(v).strip()
                for v in consolidat_df["Cnp Adult Handicap"].dropna()
            }

        for sector_name, df_sec in (sector_dfs or {}).items():
            df_sec.to_excel(writer, sheet_name=sector_name, index=False)
            cnp_col = (detection_log or {}).get(sector_name, {}).get("cnp")
            color = SECTOR_FILL_COLORS.get(sector_name)
            if cnp_col and color and consolidat_cnps:
                self._color_sector_cnps(
                    writer, sector_name, df_sec, cnp_col, consolidat_cnps, color
                )

        log_df = self._build_log(sector_dfs or {}, detection_log)
        log_df.to_excel(writer, sheet_name="LOG", index=False)

    return _repunere_out, _import_out
```

**Step 5: Rulează testul**

```bash
python -m pytest tests/test_output_builder.py::test_write_returns_dataframes -v
```

Expected: PASS

**Step 6: Modifică `processor.py` să apeleze `write_separate_files`**

În `modules/processor.py`, după apelul `builder.write(...)` (linia ~162), captează return-ul și apelează `write_separate_files`:

```python
rep_df_out, imp_df_out = builder.write(
    sector_dfs_full,
    output_path,
    detection_log=detection_log,
    motiv_data=motiv_data,
    consolidat_df=consolidat_df,
)
# Export fișiere separate
try:
    builder.write_separate_files(
        output_path.parent,
        run_date,
        rep_df_out,
        imp_df_out,
    )
except Exception as e:
    warnings.append(f"Export fișiere separate eșuat: {e}")
```

**Step 7: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 8: Commit**

```bash
git add modules/output_builder.py modules/processor.py
git commit -m "feat: wire write_separate_files into processing pipeline"
```

---

## Task 3: `write_verification_file` + wire în `_run_verificare`

**Files:**
- Modify: `APP_DIR/modules/output_builder.py`
- Modify: `APP_DIR/reinterogare_processor.py`
- Modify: `APP_DIR/tests/test_output_builder.py`

**Step 1: Scrie testul**

```python
# În test_output_builder.py
def test_write_verification_file(tmp_path):
    import pandas as pd
    from modules.output_builder import OutputBuilder
    from openpyxl import load_workbook

    verif_rep = pd.DataFrame([{"Nr. cerere": "1", "Stare OK": "✓", "Motiv OK": "✓", "Data OK": "✓"}])
    verif_imp = pd.DataFrame([{"Nr. cerere": "2", "Stare OK": "✗", "Obs OK": "✗"}])

    OutputBuilder().write_verification_file(tmp_path, "31.01.2026", verif_rep, verif_imp)

    verif_path = tmp_path / "Verificare Reinterogare 31.01.2026.xlsx"
    assert verif_path.exists()

    wb = load_workbook(str(verif_path))
    assert "VERIF REPUNERE" in wb.sheetnames
    assert "VERIF IMPORT OBS" in wb.sheetnames
```

**Step 2: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_output_builder.py::test_write_verification_file -v
```

Expected: `AttributeError`

**Step 3: Adaugă `write_verification_file` în `output_builder.py`**

```python
def write_verification_file(
    self,
    output_dir,
    date_str: str,
    verif_repunere_df: pd.DataFrame,
    verif_import_obs_df: pd.DataFrame,
):
    """Scrie 'Verificare Reinterogare {date_str}.xlsx' în output_dir cu 2 sheet-uri."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    verif_path = output_dir / f"Verificare Reinterogare {date_str}.xlsx"

    try:
        with pd.ExcelWriter(verif_path, engine="openpyxl") as writer:
            verif_repunere_df.to_excel(writer, sheet_name="VERIF REPUNERE", index=False)
            verif_import_obs_df.to_excel(writer, sheet_name="VERIF IMPORT OBS", index=False)
            self._color_verif_sheet(
                writer, "VERIF REPUNERE", verif_repunere_df,
                ok_cols=["Stare OK", "Motiv OK", "Data OK"]
            )
            self._color_verif_sheet(
                writer, "VERIF IMPORT OBS", verif_import_obs_df,
                ok_cols=["Stare OK", "Obs OK"]
            )
    except Exception as e:
        print(f"[warn] Nu s-a putut scrie fișierul de verificare: {e}")
```

**Step 4: Rulează testul**

```bash
python -m pytest tests/test_output_builder.py::test_write_verification_file -v
```

Expected: PASS

**Step 5: Wire în `_run_verificare` din `reinterogare_processor.py`**

În `_run_verificare`, după apelul `OutputBuilder().append_verification(Path(output_path), verif_rep, verif_imp)`, adaugă:

```python
# Scrie fișier de verificare separat
import re as _re
m = _re.search(r'(\d{2})_(\d{2})_(\d{4})', Path(output_path).name)
if m:
    date_str = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    try:
        OutputBuilder().write_verification_file(
            Path(output_path).parent,
            date_str,
            verif_rep,
            verif_imp,
        )
    except Exception as e:
        pass  # non-critical
```

**Step 6: Verifică sintaxa**

```bash
python -c "
import ast
with open('reinterogare_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

Expected: `Syntax OK`

**Step 7: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 8: Commit**

```bash
git add modules/output_builder.py reinterogare_processor.py tests/test_output_builder.py
git commit -m "feat: add write_verification_file and wire into verificare UI"
```

---

## Checklist final

- [ ] `Repunere.xlsx` generat automat la procesare în output_dir
- [ ] `Import Observatii.xlsx` generat automat la procesare în output_dir
- [ ] `Verificare Reinterogare {DD.MM.YYYY}.xlsx` generat la "Genereaza verificare" cu 2 sheet-uri colorate
- [ ] PermissionError pe fișierele separate → avertisment în log, procesarea nu e afectată
- [ ] Re-run suprascrie fișierele (fără duplicare)
- [ ] Toate testele pytest trec
