# Verificare Post-Reinterogare — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Adaugă un modul de verificare post-reinterogare care compară sheet-urile REPUNERE și IMPORT OBSERVATII din output-ul existent cu un Raport Stare ulterior, generând două sheet-uri de detaliu (`VERIF REPUNERE`, `VERIF IMPORT OBS`) adăugate în același fișier Excel.

**Architecture:** Modul nou `modules/verifier.py` cu logica pură (pandas, testabilă); extensie `output_builder.py` cu metoda `append_verification()`; UI în `reinterogare_processor.py` cu picker + buton dedicat. Nicio modificare la pipeline-ul de procesare existent.

**Tech Stack:** Python 3.10+, pandas, openpyxl, CustomTkinter

**Design doc:** `docs/plans/2026-03-09-verificare-post-reinterogare-design.md`

---

## Referințe critice

- `APP_DIR` = `C:\EXCELURI\Reinterogari\Reinterogare Processor\`
- Coloane Raport Stare: `Numar Cerere`, `Cnp Adult Handicap`, `Stare Cerere`, `Motiv`, `Data Inceput Stare`, `Observatii Cerere`
- Coloane REPUNERE sheet: `Nr. cerere`, `CNP adult cu handicap`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `Motiv repunere`, `Data repunere`
- Coloane IMPORT OBS sheet: `Nr. cerere`, `CNP adult cu handicap`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `Observatii`
- Join key: `CNP adult cu handicap` ↔ `Cnp Adult Handicap`
- Colorare: OK = `#C6EFCE` · NOK = `#FFC7CE` · NEGASIT = `#FFEB9C`

---

## Task 1: Modul `verifier.py` — `verify_repunere`

**Files:**
- Create: `APP_DIR/modules/verifier.py`
- Create: `APP_DIR/tests/test_verifier.py`

**Step 1: Scrie testul pentru cazul OK**

```python
# tests/test_verifier.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
import pytest
from modules.verifier import verify_repunere

def _make_repunere():
    return pd.DataFrame([{
        "Nr. cerere": "123",
        "CNP adult cu handicap": "1234567890123",
        "Nume adult cu handicap": "Popescu",
        "Prenume adult cu handicap": "Ion",
        "Motiv repunere": "Fara datorii reinterogare 31 01 2026",
        "Data repunere": "2026-01-01",
    }])

def _make_raport(stare, motiv, data):
    return pd.DataFrame([{
        "Numar Cerere": "123",
        "Cnp Adult Handicap": "1234567890123",
        "Stare Cerere": stare,
        "Motiv": motiv,
        "Data Inceput Stare": data,
        "Observatii Cerere": "",
    }])

def test_verify_repunere_ok():
    rep = _make_repunere()
    rs  = _make_raport("Aprobat",
                       "Fara datorii reinterogare 31 01 2026",
                       "2026-01-01")
    result = verify_repunere(rep, rs)
    assert result["Stare OK"].iloc[0] == "✓"
    assert result["Motiv OK"].iloc[0] == "✓"
    assert result["Data OK"].iloc[0]  == "✓"
```

**Step 2: Rulează să confirmi că FAIL**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -m pytest tests/test_verifier.py::test_verify_repunere_ok -v
```

Expected: `ImportError` sau `ModuleNotFoundError`

**Step 3: Creează `modules/verifier.py` cu `verify_repunere`**

```python
"""Post-reinterogare verification logic."""
import pandas as pd


def _norm_cnp(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def _norm_date(val) -> str:
    """Normalizează data la string YYYY-MM-DD pentru comparare."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    # Încearcă parse + reformat
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(ts):
            ts = pd.to_datetime(s, errors="coerce")
        if not pd.isna(ts):
            return ts.strftime("%Y-%m-%d")
    except Exception:
        pass
    return s


def _norm_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def verify_repunere(
    repunere_df: pd.DataFrame,
    raport_stare_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Verifică dacă beneficiarii din REPUNERE sheet apar ca Aprobat
    în Raport Stare cu același Motiv și Data Inceput Stare.

    Returns DataFrame cu coloanele originale + Stare Actuală,
    Motiv Actual, Data Inceput Stare Actuală, Stare OK, Motiv OK, Data OK.
    """
    rs = raport_stare_df.copy()
    rs["_cnp"] = _norm_cnp(rs["Cnp Adult Handicap"])

    # Index pe CNP pentru lookup rapid
    rs_idx = rs.set_index("_cnp")

    rows = []
    for _, r in repunere_df.iterrows():
        cnp = _norm_cnp(pd.Series([r["CNP adult cu handicap"]])).iloc[0]
        base = {
            "Nr. cerere":              r.get("Nr. cerere", ""),
            "CNP adult cu handicap":   r.get("CNP adult cu handicap", ""),
            "Nume adult cu handicap":  r.get("Nume adult cu handicap", ""),
            "Prenume adult cu handicap": r.get("Prenume adult cu handicap", ""),
            "Motiv repunere (asteptat)": r.get("Motiv repunere", ""),
            "Data repunere (asteptata)": r.get("Data repunere", ""),
        }

        if cnp not in rs_idx.index:
            base.update({
                "Stare Actuala": "NEGASIT",
                "Motiv Actual":  "NEGASIT",
                "Data Inceput Stare Actuala": "NEGASIT",
                "Stare OK": "NEGASIT",
                "Motiv OK": "NEGASIT",
                "Data OK":  "NEGASIT",
            })
        else:
            match = rs_idx.loc[cnp]
            # Dacă CNP apare de mai multe ori → iau primul rând
            if isinstance(match, pd.DataFrame):
                match = match.iloc[0]

            stare_act  = _norm_str(match.get("Stare Cerere"))
            motiv_act  = _norm_str(match.get("Motiv"))
            data_act   = _norm_date(match.get("Data Inceput Stare"))
            motiv_exp  = _norm_str(base["Motiv repunere (asteptat)"])
            data_exp   = _norm_date(base["Data repunere (asteptata)"])

            base.update({
                "Stare Actuala": stare_act,
                "Motiv Actual":  motiv_act,
                "Data Inceput Stare Actuala": data_act,
                "Stare OK": "✓" if stare_act.lower() == "aprobat" else "✗",
                "Motiv OK": "✓" if motiv_act == motiv_exp else "✗",
                "Data OK":  "✓" if data_act  == data_exp  else "✗",
            })
        rows.append(base)

    return pd.DataFrame(rows)
```

**Step 4: Rulează testul**

```bash
python -m pytest tests/test_verifier.py::test_verify_repunere_ok -v
```

Expected: PASS

**Step 5: Adaugă teste pentru cazurile NOK și NEGASIT**

```python
def test_verify_repunere_nerepus():
    rep = _make_repunere()
    rs  = _make_raport("Suspendat", "Datorii interogare", "2026-01-01")
    result = verify_repunere(rep, rs)
    assert result["Stare OK"].iloc[0] == "✗"

def test_verify_repunere_negasit():
    rep = _make_repunere()
    rs  = pd.DataFrame(columns=["Numar Cerere", "Cnp Adult Handicap",
                                 "Stare Cerere", "Motiv",
                                 "Data Inceput Stare", "Observatii Cerere"])
    result = verify_repunere(rep, rs)
    assert result["Stare OK"].iloc[0] == "NEGASIT"

def test_verify_repunere_motiv_diferit():
    rep = _make_repunere()
    rs  = _make_raport("Aprobat", "Alt motiv", "2026-01-01")
    result = verify_repunere(rep, rs)
    assert result["Stare OK"].iloc[0] == "✓"
    assert result["Motiv OK"].iloc[0] == "✗"

def test_verify_repunere_data_diferita():
    rep = _make_repunere()
    rs  = _make_raport("Aprobat",
                       "Fara datorii reinterogare 31 01 2026",
                       "2026-02-01")
    result = verify_repunere(rep, rs)
    assert result["Data OK"].iloc[0] == "✗"
```

**Step 6: Rulează toate testele**

```bash
python -m pytest tests/test_verifier.py -v
```

Expected: 5 PASS

**Step 7: Commit**

```bash
git add modules/verifier.py tests/test_verifier.py
git commit -m "feat: add verify_repunere logic with tests"
```

---

## Task 2: Modul `verifier.py` — `verify_import_obs`

**Files:**
- Modify: `APP_DIR/modules/verifier.py`
- Modify: `APP_DIR/tests/test_verifier.py`

**Step 1: Scrie testele**

```python
from modules.verifier import verify_import_obs

def _make_import_obs():
    return pd.DataFrame([{
        "Nr. cerere":              "456",
        "CNP adult cu handicap":   "9876543210987",
        "Nume adult cu handicap":  "Ionescu",
        "Prenume adult cu handicap": "Maria",
        "Observatii":              "Datorii reinterogare 31 01 2026 S2",
    }])

def _make_raport_imp(stare, obs_platforma=""):
    return pd.DataFrame([{
        "Numar Cerere":        "456",
        "Cnp Adult Handicap":  "9876543210987",
        "Stare Cerere":        stare,
        "Motiv":               "Datorii interogare",
        "Data Inceput Stare":  "2025-10-01",
        "Observatii Cerere":   obs_platforma,
    }])

def test_verify_import_obs_ok():
    imp = _make_import_obs()
    rs  = _make_raport_imp("Suspendat",
                           "Datorii reinterogare 31 01 2026 S2")
    result = verify_import_obs(imp, rs)
    assert result["Stare OK"].iloc[0] == "✓"
    assert result["Obs OK"].iloc[0]   == "✓"

def test_verify_import_obs_stare_schimbata():
    imp = _make_import_obs()
    rs  = _make_raport_imp("Aprobat", "Datorii reinterogare 31 01 2026 S2")
    result = verify_import_obs(imp, rs)
    assert result["Stare OK"].iloc[0] == "✗"

def test_verify_import_obs_fara_observatii_camp():
    """Raport Stare fără câmp Observatii Cerere → Obs OK = '(fara camp)'."""
    imp = _make_import_obs()
    rs  = _make_raport_imp("Suspendat", "")
    # Șterge coloana Observatii Cerere
    rs = rs.drop(columns=["Observatii Cerere"])
    result = verify_import_obs(imp, rs)
    assert result["Obs OK"].iloc[0] == "(fara camp)"

def test_verify_import_obs_negasit():
    imp = _make_import_obs()
    rs  = pd.DataFrame(columns=["Numar Cerere", "Cnp Adult Handicap",
                                 "Stare Cerere", "Motiv",
                                 "Data Inceput Stare", "Observatii Cerere"])
    result = verify_import_obs(imp, rs)
    assert result["Stare OK"].iloc[0] == "NEGASIT"
```

**Step 2: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_verifier.py::test_verify_import_obs_ok -v
```

Expected: `ImportError`

**Step 3: Adaugă `verify_import_obs` în `verifier.py`**

```python
def verify_import_obs(
    import_obs_df: pd.DataFrame,
    raport_stare_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Verifică dacă beneficiarii din IMPORT OBSERVATII sunt încă Suspendați
    și dacă observația importată apare în câmpul Observatii Cerere din platformă.
    """
    rs = raport_stare_df.copy()
    rs["_cnp"] = _norm_cnp(rs["Cnp Adult Handicap"])
    has_obs_col = "Observatii Cerere" in rs.columns
    rs_idx = rs.set_index("_cnp")

    rows = []
    for _, r in import_obs_df.iterrows():
        cnp = _norm_cnp(pd.Series([r["CNP adult cu handicap"]])).iloc[0]
        obs_exp = _norm_str(r.get("Observatii", ""))

        base = {
            "Nr. cerere":              r.get("Nr. cerere", ""),
            "CNP adult cu handicap":   r.get("CNP adult cu handicap", ""),
            "Nume adult cu handicap":  r.get("Nume adult cu handicap", ""),
            "Prenume adult cu handicap": r.get("Prenume adult cu handicap", ""),
            "Observatie importata (asteptata)": obs_exp,
        }

        if cnp not in rs_idx.index:
            base.update({
                "Stare Actuala":    "NEGASIT",
                "Observatii platforma": "NEGASIT",
                "Stare OK": "NEGASIT",
                "Obs OK":   "NEGASIT",
            })
        else:
            match = rs_idx.loc[cnp]
            if isinstance(match, pd.DataFrame):
                match = match.iloc[0]

            stare_act = _norm_str(match.get("Stare Cerere"))
            if has_obs_col:
                obs_platf = _norm_str(match.get("Observatii Cerere", ""))
                obs_ok = "✓" if obs_exp and obs_exp in obs_platf else "✗"
            else:
                obs_platf = ""
                obs_ok = "(fara camp)"

            base.update({
                "Stare Actuala":    stare_act,
                "Observatii platforma": obs_platf,
                "Stare OK": "✓" if stare_act.lower() == "suspendat" else "✗",
                "Obs OK":   obs_ok,
            })
        rows.append(base)

    return pd.DataFrame(rows)
```

**Step 4: Rulează toate testele**

```bash
python -m pytest tests/test_verifier.py -v
```

Expected: 9 PASS

**Step 5: Commit**

```bash
git add modules/verifier.py tests/test_verifier.py
git commit -m "feat: add verify_import_obs logic with tests"
```

---

## Task 3: `output_builder.py` — `append_verification`

**Files:**
- Modify: `APP_DIR/modules/output_builder.py`
- Modify: `APP_DIR/tests/test_output_builder.py`

**Step 1: Scrie testul**

```python
# În test_output_builder.py — adaugă la final:
from modules.verifier import verify_repunere, verify_import_obs
from modules.output_builder import OutputBuilder
from openpyxl import load_workbook

def test_append_verification_adds_sheets(tmp_path):
    # Creează un Excel de bază cu sheet REPUNERE și IMPORT OBSERVATII
    base = tmp_path / "output.xlsx"
    with pd.ExcelWriter(base, engine="openpyxl") as w:
        pd.DataFrame([{"A": 1}]).to_excel(w, sheet_name="CONSOLIDAT", index=False)

    verif_rep = pd.DataFrame([{
        "Nr. cerere": "1", "CNP adult cu handicap": "111",
        "Stare OK": "✓", "Motiv OK": "✓", "Data OK": "✓",
    }])
    verif_imp = pd.DataFrame([{
        "Nr. cerere": "2", "CNP adult cu handicap": "222",
        "Stare OK": "✗", "Obs OK": "✗",
    }])

    OutputBuilder().append_verification(base, verif_rep, verif_imp)

    wb = load_workbook(str(base))
    assert "VERIF REPUNERE" in wb.sheetnames
    assert "VERIF IMPORT OBS" in wb.sheetnames
```

**Step 2: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_output_builder.py::test_append_verification_adds_sheets -v
```

Expected: `AttributeError: 'OutputBuilder' object has no attribute 'append_verification'`

**Step 3: Adaugă `append_verification` în `output_builder.py`**

```python
# Adaugă constantele de culoare în output_builder.py (după SECTOR_FILL_COLORS):
VERIF_FILL_OK    = "C6EFCE"   # verde deschis
VERIF_FILL_NOK   = "FFC7CE"   # rosu deschis
VERIF_FILL_MISS  = "FFEB9C"   # galben

# Adaugă metoda în clasa OutputBuilder:
def append_verification(
    self,
    output_path: Path,
    verif_repunere_df: pd.DataFrame,
    verif_import_obs_df: pd.DataFrame,
):
    """Adaugă sheet-urile VERIF REPUNERE și VERIF IMPORT OBS în fișierul Excel existent."""
    from openpyxl import load_workbook

    # Șterge sheet-urile vechi dacă există (re-run)
    wb = load_workbook(str(output_path))
    for name in ("VERIF REPUNERE", "VERIF IMPORT OBS"):
        if name in wb.sheetnames:
            del wb[name]
    wb.save(str(output_path))

    with pd.ExcelWriter(
        output_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        verif_repunere_df.to_excel(
            writer, sheet_name="VERIF REPUNERE", index=False
        )
        verif_import_obs_df.to_excel(
            writer, sheet_name="VERIF IMPORT OBS", index=False
        )
        self._color_verif_sheet(
            writer, "VERIF REPUNERE", verif_repunere_df,
            ok_cols=["Stare OK", "Motiv OK", "Data OK"]
        )
        self._color_verif_sheet(
            writer, "VERIF IMPORT OBS", verif_import_obs_df,
            ok_cols=["Stare OK", "Obs OK"]
        )

def _color_verif_sheet(
    self, writer, sheet_name: str, df: pd.DataFrame, ok_cols: list
):
    """Colorează rânduri: verde dacă toate ok_cols=✓, galben dacă NEGASIT, roșu altfel."""
    ws = writer.sheets.get(sheet_name)
    if ws is None:
        return

    fill_ok   = PatternFill(start_color=VERIF_FILL_OK,   end_color=VERIF_FILL_OK,   fill_type="solid")
    fill_nok  = PatternFill(start_color=VERIF_FILL_NOK,  end_color=VERIF_FILL_NOK,  fill_type="solid")
    fill_miss = PatternFill(start_color=VERIF_FILL_MISS,  end_color=VERIF_FILL_MISS, fill_type="solid")

    n_cols = len(df.columns)
    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        vals = [str(row.get(c, "")) for c in ok_cols if c in df.columns]
        if all(v == "NEGASIT" for v in vals):
            fill = fill_miss
        elif all(v == "✓" for v in vals):
            fill = fill_ok
        else:
            fill = fill_nok
        for col_idx in range(1, n_cols + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill
```

**Step 4: Rulează testul**

```bash
python -m pytest tests/test_output_builder.py::test_append_verification_adds_sheets -v
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
git commit -m "feat: add append_verification to OutputBuilder"
```

---

## Task 4: UI — picker + buton Verificare

**Files:**
- Modify: `APP_DIR/reinterogare_processor.py`

**Step 1: Adaugă secțiunea UI în `_build_ui`**

Imediat după secțiunea `# --- Optional files ---`, adaugă:

```python
# --- Verificare post-Reinterogare ---
verif_frame = ctk.CTkFrame(self)
verif_frame.pack(fill="x", padx=16, pady=4)
ctk.CTkLabel(verif_frame, text="Verificare post-Reinterogare:",
             font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(6, 2))

self.verif_raport_picker = FilePicker(
    verif_frame, label="Raport Stare verif.", clearable=True,
)
self.verif_raport_picker.pack(fill="x", padx=10, pady=2)

out_row = ctk.CTkFrame(verif_frame, fg_color="transparent")
out_row.pack(fill="x", padx=10, pady=2)
ctk.CTkLabel(out_row, text="Output reinterogare", width=90, anchor="w").pack(side="left", padx=(0, 6))
self.verif_output_entry = ctk.CTkEntry(out_row, width=340)
self.verif_output_entry.pack(side="left", padx=(0, 4))
ctk.CTkButton(out_row, text="...", width=32,
              command=self._pick_verif_output).pack(side="left")

self.verif_btn = ctk.CTkButton(
    verif_frame, text="Genereaza verificare", width=180, height=32,
    fg_color="#1b5e20", hover_color="#2e7d32",
    command=self._run_verificare,
)
self.verif_btn.pack(anchor="w", padx=10, pady=(4, 8))
```

**Step 2: Adaugă `_pick_verif_output` și pre-completare automată**

La sfârșitul metodei `_on_done` (după `self.log.append(f"\nOutput salvat: {out_path}")`), adaugă:

```python
# Pre-completează câmpul output pentru verificare
self.verif_output_entry.delete(0, "end")
self.verif_output_entry.insert(0, str(out_path))
```

Și adaugă metoda:

```python
def _pick_verif_output(self):
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if path:
        self.verif_output_entry.delete(0, "end")
        self.verif_output_entry.insert(0, path)
```

**Step 3: Adaugă `_run_verificare`**

```python
def _run_verificare(self):
    raport_path = self.verif_raport_picker.get()
    output_path = self.verif_output_entry.get().strip()

    if not raport_path or not Path(raport_path).exists():
        messagebox.showwarning("Atentie", "Selecteaza Raportul de Stare pentru verificare.")
        return
    if not output_path or not Path(output_path).exists():
        messagebox.showwarning("Atentie", "Selecteaza fisierul output reinterogare.")
        return

    self.verif_btn.configure(state="disabled")
    self.log.clear()
    self.log.append("Verificare post-reinterogare...")

    def _worker():
        try:
            import pandas as pd
            from modules.verifier import verify_repunere, verify_import_obs
            from modules.raport_stari_loader import RaportStariLoader
            from modules.output_builder import OutputBuilder

            # Citeste Raportul de Stare verificare
            rs_df = RaportStariLoader().load(Path(raport_path))

            # Citeste sheet-urile din output reinterogare
            try:
                rep_df = pd.read_excel(output_path, sheet_name="REPUNERE", dtype=str)
            except Exception:
                rep_df = pd.DataFrame()

            try:
                imp_df = pd.read_excel(output_path, sheet_name="IMPORT OBSERVATII", dtype=str)
            except Exception:
                imp_df = pd.DataFrame()

            if rep_df.empty and imp_df.empty:
                self.after(0, lambda: messagebox.showwarning(
                    "Atentie",
                    "Nu s-au gasit sheet-urile REPUNERE sau IMPORT OBSERVATII in fisierul selectat."
                ))
                return

            verif_rep = verify_repunere(rep_df, rs_df) if not rep_df.empty else pd.DataFrame()
            verif_imp = verify_import_obs(imp_df, rs_df) if not imp_df.empty else pd.DataFrame()

            OutputBuilder().append_verification(Path(output_path), verif_rep, verif_imp)

            n_rep = len(verif_rep)
            n_ok_rep = (verif_rep[["Stare OK","Motiv OK","Data OK"]] == "✓").all(axis=1).sum() if not verif_rep.empty else 0
            n_imp = len(verif_imp)
            n_ok_imp = (verif_imp["Stare OK"] == "✓").sum() if not verif_imp.empty else 0

            self.after(0, lambda: self.log.append(
                f"REPUNERE: {n_ok_rep}/{n_rep} OK | IMPORT OBS: {n_ok_imp}/{n_imp} OK\n"
                f"Sheet-uri adaugate in: {output_path}"
            ))
        except PermissionError:
            self.after(0, lambda: messagebox.showerror(
                "Eroare", f"Inchide fisierul Excel inainte de verificare:\n{output_path}"
            ))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Eroare", str(e)))
        finally:
            self.after(0, lambda: self.verif_btn.configure(state="normal"))

    threading.Thread(target=_worker, daemon=True).start()
```

**Step 4: Verifică sintaxa**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python -c "
import ast
with open('reinterogare_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

Expected: `Syntax OK`

**Step 5: Commit**

```bash
git add reinterogare_processor.py
git commit -m "feat: add verificare post-reinterogare UI (picker + button)"
```

---

## Task 5: Test manual end-to-end

**Step 1: Pornește aplicația**

```bash
cd "C:/EXCELURI/Reinterogari/Reinterogare Processor"
python reinterogare_processor.py
```

**Step 2: Verifică UI**
- Secțiunea "Verificare post-Reinterogare" apare sub "Fișiere opționale"
- Picker "Raport Stare verif." funcționează (buton `...` și `X`)
- Câmpul "Output reinterogare" e editabil + picker `...`

**Step 3: Test după o procesare**
- Rulează o procesare normală (▶ Procesează)
- Verifică că "Output reinterogare" se pre-completează automat cu calea fișierului generat

**Step 4: Test verificare**
- Selectează un Raport Stare mai recent
- Click "Generează verificare"
- Verifică că în fișierul Excel apar sheet-urile `VERIF REPUNERE` și `VERIF IMPORT OBS`
- Verifică colorarea: verde OK, roșu NOK, galben NEGASIT

**Step 5: Test PermissionError**
- Deschide fișierul output în Excel
- Click "Generează verificare"
- Expected: mesaj de eroare "Inchide fisierul Excel..."

---

## Checklist final

- [ ] `verify_repunere` returnează ✓/✗/NEGASIT per condiție
- [ ] `verify_import_obs` returnează ✓/✗/NEGASIT/(fara camp) per condiție
- [ ] `append_verification` adaugă sheet-urile în fișierul existent fără să modifice restul
- [ ] Re-rulare suprascrie sheet-urile vechi (nu duplică)
- [ ] UI: picker Raport Stare verif. + câmp output + buton
- [ ] Pre-completare automată a câmpului output după procesare
- [ ] Colorare corectă în ambele sheet-uri de verificare
- [ ] PermissionError gestionat cu mesaj clar
- [ ] Toate testele pytest trec
