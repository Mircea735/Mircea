# Loop Processor — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Aplicație wizard CustomTkinter dark-mode care integrează Interogare Termen + Reinterogare Lunară într-un singur UI cu timeline vizual, contoare colorate și memorie persistentă a ciclului semestriar.

**Architecture:** `loop_processor.py` este app-ul principal cu layout două coloane: sidebar timeline stânga + panel dinamic dreapta. `cycle_state.py` gestionează persistența JSON a ciclului. Modulele procesoare sunt copiate din proiectele existente. Filtrul global de sector adresă e aplicat în `RaportStariLoader` înainte de merge.

**Tech Stack:** Python 3.10+, CustomTkinter, pandas, openpyxl

**APP_DIR**: `C:\EXCELURI\Interogari termen\Loop Processor\`

**Design doc**: `docs/plans/2026-03-10-loop-processor-design.md`

---

## Referințe critice

**Surse module:**
- `C:\EXCELURI\Interogari termen\Interogare Termen Processor\modules\` — termen_logic, termen_processor, contingent_tracker + shared
- `C:\EXCELURI\Reinterogari\Reinterogare Processor\modules\` — processor (ReinterogareProcessor), debt_logic

**Paleta culori:**
```python
BG_MAIN    = "#0D1117"
BG_SIDEBAR = "#161B22"
BG_CARD    = "#1C2128"
ACCENT     = "#2DD4BF"   # teal — activ
DONE       = "#6366F1"   # indigo — completat
WARN       = "#F59E0B"   # amber — warning
C_SUSP     = "#EF4444"   # roșu — Suspendare
C_REP      = "#22C55E"   # verde — Repunere
C_IMP      = "#3B82F6"   # albastru — Import Obs
C_DAT      = "#F59E0B"   # galben — Datornici
BORDER     = "#30363D"
TEXT_DIM   = "#8B949E"
TEXT       = "#E6EDF3"
```

**Structura unui pas în cycle_state.json:**
```json
{
  "id": "termen_30.09.2025",
  "type": "termen",
  "label": "IT 30.09.2025",
  "status": "pending",
  "n_suspendare": 0, "n_repunere": 0,
  "n_import_obs": 0, "n_datornici": 0,
  "output_dir": ""
}
```

**Coloane Raport Stare pentru filtrul de sector:**
- Sector extras din `Adresa Adult Handicap` cu regex `[Ss]ect(?:or)?\.?\s*(\d)` → `"Sector N"`
- Filtrul compară cu lista selectată ex. `["Sector 1", "Sector 2"]`

---

## Task 1: Setup proiect + copiere module

**Files:**
- Create: `APP_DIR/` (folder structure)
- Create: `APP_DIR/modules/__init__.py`
- Create: `APP_DIR/tests/__init__.py`
- Create: `APP_DIR/config/settings.json`
- Create: `APP_DIR/config/column_aliases.json` (copiat)

**Step 1: Creează structura de foldere**

```bash
mkdir -p "C:/EXCELURI/Interogari termen/Loop Processor/modules"
mkdir -p "C:/EXCELURI/Interogari termen/Loop Processor/tests"
mkdir -p "C:/EXCELURI/Interogari termen/Loop Processor/config"
```

**Step 2: Copiază modulele shared**

```bash
SRC_IT="C:/EXCELURI/Interogari termen/Interogare Termen Processor/modules"
SRC_RP="C:/EXCELURI/Reinterogari/Reinterogare Processor/modules"
DST="C:/EXCELURI/Interogari termen/Loop Processor/modules"

cp "$SRC_IT/config_manager.py" "$DST/"
cp "$SRC_IT/column_detector.py" "$DST/"
cp "$SRC_IT/sector_loader.py" "$DST/"
cp "$SRC_IT/raport_stari_loader.py" "$DST/"
cp "$SRC_IT/merger.py" "$DST/"
cp "$SRC_IT/output_builder.py" "$DST/"
cp "$SRC_IT/ui_components.py" "$DST/"
cp "$SRC_IT/termen_logic.py" "$DST/"
cp "$SRC_IT/termen_processor.py" "$DST/"
cp "$SRC_IT/contingent_tracker.py" "$DST/"
cp "$SRC_RP/processor.py" "$DST/"
cp "$SRC_RP/debt_logic.py" "$DST/"
cp "$SRC_IT/../config/column_aliases.json" \
   "C:/EXCELURI/Interogari termen/Loop Processor/config/"
```

**Step 3: Creează fișierele init**

`APP_DIR/modules/__init__.py` — gol
`APP_DIR/tests/__init__.py` — gol
`APP_DIR/config/settings.json` — conținut: `{}`

**Step 4: Init git**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
git init
git add .
git commit -m "chore: initial project setup — copy shared modules"
```

---

## Task 2: `modules/cycle_state.py` — logica ciclului

**Files:**
- Create: `APP_DIR/modules/cycle_state.py`
- Create: `APP_DIR/tests/test_cycle_state.py`

**Step 1: Scrie testele**

```python
# tests/test_cycle_state.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest
from datetime import date
from modules.cycle_state import CycleState

def test_init_generates_steps(tmp_path):
    cs = CycleState(tmp_path / "cycle_state.json")
    cs.init_cycle(date(2025, 9, 30))
    state = cs.load()
    assert len(state["steps"]) == 7  # IT + 5 luni + IT
    assert state["steps"][0]["type"] == "termen"
    assert state["steps"][0]["id"] == "termen_30.09.2025"
    assert state["steps"][1]["type"] == "reinterogare"
    assert state["steps"][1]["id"] == "reinterogare_Oct_2025"
    assert state["steps"][-1]["type"] == "termen"
    assert state["steps"][-1]["id"] == "termen_30.03.2026"

def test_first_step_active_after_init(tmp_path):
    cs = CycleState(tmp_path / "cycle_state.json")
    cs.init_cycle(date(2025, 9, 30))
    active = cs.get_active_step()
    assert active["id"] == "termen_30.09.2025"
    assert active["status"] == "active"

def test_mark_completed_unlocks_next(tmp_path):
    cs = CycleState(tmp_path / "cycle_state.json")
    cs.init_cycle(date(2025, 9, 30))
    cs.mark_completed("termen_30.09.2025", {"n_suspendare": 100, "n_datornici": 80})
    state = cs.load()
    assert state["steps"][0]["status"] == "completed"
    assert state["steps"][0]["n_suspendare"] == 100
    assert state["steps"][1]["status"] == "active"

def test_reset_clears_state(tmp_path):
    cs = CycleState(tmp_path / "cycle_state.json")
    cs.init_cycle(date(2025, 9, 30))
    cs.mark_completed("termen_30.09.2025", {})
    cs.reset()
    state = cs.load()
    assert state == {}

def test_get_step_stats(tmp_path):
    cs = CycleState(tmp_path / "cycle_state.json")
    cs.init_cycle(date(2025, 9, 30))
    cs.mark_completed("termen_30.09.2025", {"n_repunere": 50, "output_dir": "/out"})
    step = cs.get_step("termen_30.09.2025")
    assert step["n_repunere"] == 50
    assert step["output_dir"] == "/out"
```

**Step 2: Rulează să confirmi FAIL**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -m pytest tests/test_cycle_state.py -v 2>&1 | head -10
```

Expected: `ModuleNotFoundError`

**Step 3: Creează `modules/cycle_state.py`**

```python
"""Persistent cycle state for the Loop Processor wizard."""
import json
from datetime import date
from pathlib import Path
from dateutil.relativedelta import relativedelta


def _month_label(d: date) -> str:
    LUNI = {1:"Ian",2:"Feb",3:"Mar",4:"Apr",5:"Mai",6:"Iun",
            7:"Iul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{LUNI[d.month]} {d.year}"


def _generate_steps(term_start: date) -> list:
    """
    Generează pașii unui ciclu semestriar:
    IT(term_start) → N luni reinterogare → IT(term_end)
    term_start: 30.09 → 5 luni (Oct-Feb) → 30.03 următor
    term_start: 30.03 → 5 luni (Apr-Aug) → 30.09 următor
    """
    steps = []
    tds = term_start.strftime("%d.%m.%Y")
    steps.append({
        "id": f"termen_{tds}",
        "type": "termen",
        "label": f"IT {tds}",
        "status": "active",
        "n_suspendare": 0, "n_repunere": 0,
        "n_import_obs": 0, "n_datornici": 0,
        "output_dir": "",
    })

    # Lunile de reinterogare: prima lună după termen până la luna înainte de term_end
    month_cursor = term_start.replace(day=1) + relativedelta(months=1)
    term_end = term_start.replace(day=1) + relativedelta(months=6)
    # term_end e prima zi a lunii termenului următor
    # Reinterogările sunt în lunile [month_cursor, term_end) exclusive
    while month_cursor < term_end:
        label = _month_label(month_cursor)
        steps.append({
            "id": f"reinterogare_{label.replace(' ', '_')}",
            "type": "reinterogare",
            "label": f"Reintr. {label}",
            "status": "pending",
            "n_repunere": 0, "n_import_obs": 0,
            "output_dir": "",
            "contingent_source": "cycle",
        })
        month_cursor += relativedelta(months=1)

    # Termenul următor: ultima zi a lunii term_end - 1 zi
    term_next = term_end - relativedelta(days=1)
    tdn = term_next.strftime("%d.%m.%Y")
    steps.append({
        "id": f"termen_{tdn}",
        "type": "termen",
        "label": f"IT {tdn}",
        "status": "pending",
        "n_suspendare": 0, "n_repunere": 0,
        "n_import_obs": 0, "n_datornici": 0,
        "output_dir": "",
    })
    return steps


class CycleState:
    def __init__(self, path):
        self.path = Path(path)

    def load(self) -> dict:
        if not self.path.exists():
            return {}
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}

    def _save(self, state: dict):
        self.path.write_text(
            json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def init_cycle(self, term_start: date, sector_filter: list = None):
        state = {
            "term_start": term_start.strftime("%d.%m.%Y"),
            "sector_filter": sector_filter or [],
            "steps": _generate_steps(term_start),
        }
        self._save(state)

    def get_active_step(self) -> dict | None:
        state = self.load()
        for step in state.get("steps", []):
            if step["status"] == "active":
                return step
        return None

    def get_step(self, step_id: str) -> dict | None:
        state = self.load()
        for step in state.get("steps", []):
            if step["id"] == step_id:
                return step
        return None

    def mark_completed(self, step_id: str, stats: dict):
        state = self.load()
        steps = state.get("steps", [])
        for i, step in enumerate(steps):
            if step["id"] == step_id:
                step["status"] = "completed"
                step.update(stats)
                # Activează pasul următor
                if i + 1 < len(steps) and steps[i + 1]["status"] == "pending":
                    steps[i + 1]["status"] = "active"
                break
        self._save(state)

    def reset(self):
        if self.path.exists():
            self.path.write_text("{}", encoding="utf-8")

    def update_sector_filter(self, sectors: list):
        state = self.load()
        state["sector_filter"] = sectors
        self._save(state)
```

**Step 4: Rulează testele**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -m pytest tests/test_cycle_state.py -v
```

Expected: 5 PASS

**Step 5: Commit**

```bash
git add modules/cycle_state.py tests/test_cycle_state.py
git commit -m "feat: add CycleState for wizard step management"
```

---

## Task 3: Filtrul de sector adresă

**Files:**
- Create: `APP_DIR/modules/sector_address_filter.py`
- Create: `APP_DIR/tests/test_sector_address_filter.py`

**Step 1: Scrie testele**

```python
# tests/test_sector_address_filter.py
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from modules.sector_address_filter import filter_by_sector_address

def _df(addresses):
    return pd.DataFrame({"Adresa Adult Handicap": addresses, "Cnp Adult Handicap": range(len(addresses))})

def test_filter_keeps_matching_sector():
    df = _df(["Str. X, Sector 2", "Str. Y, Sector 3", "Sector 1, nr. 5"])
    result = filter_by_sector_address(df, ["Sector 1", "Sector 2"])
    assert len(result) == 2
    assert "Sector 3" not in result["Adresa Adult Handicap"].values[0]

def test_empty_filter_returns_all():
    df = _df(["Sector 1", "Sector 2", "Sector 3"])
    result = filter_by_sector_address(df, [])
    assert len(result) == 3

def test_no_address_col_returns_all():
    df = pd.DataFrame({"Cnp Adult Handicap": [1, 2, 3]})
    result = filter_by_sector_address(df, ["Sector 1"])
    assert len(result) == 3

def test_undetected_address_excluded():
    df = _df(["Sector 1", "adresa fara sector"])
    result = filter_by_sector_address(df, ["Sector 1"])
    assert len(result) == 1
```

**Step 2: Rulează să confirmi FAIL**

```bash
python -m pytest tests/test_sector_address_filter.py -v 2>&1 | head -10
```

**Step 3: Creează `modules/sector_address_filter.py`**

```python
"""Filter Raport Stare rows by beneficiary's home address sector."""
import re
import pandas as pd

_SECTOR_RE = re.compile(r'[Ss]ect(?:or)?\.?\s*(\d)', re.IGNORECASE)
_ADDR_COL = "Adresa Adult Handicap"


def _extract_sector(addr) -> str | None:
    if not addr or (isinstance(addr, float)):
        return None
    m = _SECTOR_RE.search(str(addr))
    return f"Sector {m.group(1)}" if m else None


def filter_by_sector_address(df: pd.DataFrame, sectors: list) -> pd.DataFrame:
    """
    Păstrează doar rândurile unde sectorul din adresă ∈ sectors.
    Dacă sectors e gol sau coloana adresă lipsește → returnează df nemodificat.
    """
    if not sectors or _ADDR_COL not in df.columns:
        return df
    detected = df[_ADDR_COL].apply(_extract_sector)
    mask = detected.isin(sectors)
    return df[mask].reset_index(drop=True)
```

**Step 4: Rulează testele**

```bash
python -m pytest tests/test_sector_address_filter.py -v
```

Expected: 4 PASS

**Step 5: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS

**Step 6: Commit**

```bash
git add modules/sector_address_filter.py tests/test_sector_address_filter.py
git commit -m "feat: add sector address filter for global beneficiary filtering"
```

---

## Task 4: `loop_processor.py` — UI principal (structură + timeline)

**Files:**
- Create: `APP_DIR/loop_processor.py`

Aceasta este cea mai complexă componentă. Creăm UI-ul în pași: mai întâi structura, apoi panelurile.

**Step 1: Creează `loop_processor.py` cu structura de bază**

```python
"""Loop Processor — Integrated Interogare Termen + Reinterogare wizard."""
import os
import threading
from pathlib import Path
from datetime import date, datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox

from modules.config_manager import ConfigManager
from modules.cycle_state import CycleState
from modules.contingent_tracker import ContingentTracker

APP_DIR = Path(__file__).parent
CONFIG_DIR = APP_DIR / "config"
CYCLE_STATE_PATH = APP_DIR / "cycle_state.json"
CONTINGENT_PATH = APP_DIR / "contingent.json"

# ── Culori ────────────────────────────────────────────────────────────────────
BG_MAIN    = "#0D1117"
BG_SIDEBAR = "#161B22"
BG_CARD    = "#1C2128"
ACCENT     = "#2DD4BF"
DONE       = "#6366F1"
WARN       = "#F59E0B"
C_SUSP     = "#EF4444"
C_REP      = "#22C55E"
C_IMP      = "#3B82F6"
C_DAT      = "#F59E0B"
BORDER     = "#30363D"
TEXT_DIM   = "#8B949E"
TEXT       = "#E6EDF3"

SECTORS = ["S1", "S2", "S3", "S4", "S5", "S6"]
ADDR_SECTORS = ["Sector 1", "Sector 2", "Sector 3",
                "Sector 4", "Sector 5", "Sector 6"]

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class StatCard(ctk.CTkFrame):
    """Card vizual pentru un contor (Suspendare / Repunere etc.)."""
    def __init__(self, master, label: str, color: str, **kwargs):
        super().__init__(master, fg_color=BG_CARD, corner_radius=10,
                         border_width=1, border_color=BORDER, **kwargs)
        self._color = color
        ctk.CTkLabel(self, text=label, font=ctk.CTkFont(size=10),
                     text_color=TEXT_DIM).pack(pady=(8, 0))
        self._val = ctk.CTkLabel(self, text="—",
                                  font=ctk.CTkFont(size=22, weight="bold"),
                                  text_color=color)
        self._val.pack(pady=(0, 8))

    def set_value(self, v: int):
        self._val.configure(text=str(v))


class TimelineStep(ctk.CTkFrame):
    """Widget pentru un pas din timeline sidebar."""
    def __init__(self, master, step: dict, on_click, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.step_id = step["id"]
        self.on_click = on_click

        status = step.get("status", "pending")
        if status == "completed":
            dot_color, label_color = DONE, TEXT
            dot_text = "✓"
        elif status == "active":
            dot_color, label_color = ACCENT, TEXT
            dot_text = "●"
        else:
            dot_color, label_color = "#374151", TEXT_DIM
            dot_text = "○"

        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=2)

        dot = ctk.CTkLabel(row, text=dot_text, width=20,
                            text_color=dot_color,
                            font=ctk.CTkFont(size=14, weight="bold"))
        dot.pack(side="left")

        lbl = ctk.CTkLabel(row, text=step.get("label", step["id"]),
                            text_color=label_color,
                            font=ctk.CTkFont(size=12),
                            anchor="w")
        lbl.pack(side="left", fill="x", expand=True)

        if status in ("completed", "active"):
            for w in (row, dot, lbl):
                w.bind("<Button-1>", lambda e: self.on_click(self.step_id))
            lbl.configure(cursor="hand2")


class LoopProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Loop Processor — Ciclu Datornici DGITL")
        self.geometry("980x720")
        self.resizable(True, True)
        self.configure(fg_color=BG_MAIN)

        self.cm = ConfigManager(CONFIG_DIR)
        self.cs = CycleState(CYCLE_STATE_PATH)
        self._sector_vars = {}   # {"Sector 1": BooleanVar, ...}
        self._current_panel = None

        self._build_header()
        self._build_body()
        self._refresh_ui()

    # ── Header ────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_CARD,
                           corner_radius=0, border_width=0)
        hdr.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(hdr, text="Ciclu Datornici DGITL",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=ACCENT).pack(side="left", padx=16, pady=10)

        # Sector address filter
        filter_frame = ctk.CTkFrame(hdr, fg_color="transparent")
        filter_frame.pack(side="left", padx=16)
        ctk.CTkLabel(filter_frame, text="Sector adresă:",
                     text_color=TEXT_DIM,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0,6))
        for s in ADDR_SECTORS:
            var = ctk.BooleanVar(value=False)
            self._sector_vars[s] = var
            cb = ctk.CTkCheckBox(filter_frame, text=s.replace("Sector ", "S"),
                                  variable=var, width=44,
                                  fg_color=ACCENT, hover_color=DONE,
                                  text_color=TEXT,
                                  font=ctk.CTkFont(size=11),
                                  command=self._on_sector_filter_changed)
            cb.pack(side="left", padx=2)

        ctk.CTkButton(hdr, text="Tot",
                      width=40, height=26,
                      fg_color=BG_MAIN, hover_color=BORDER,
                      text_color=TEXT_DIM, font=ctk.CTkFont(size=11),
                      command=self._select_all_sectors).pack(side="left", padx=4)

        # Reset + New cycle buttons
        ctk.CTkButton(hdr, text="Reset ciclu",
                      width=100, height=28,
                      fg_color="#7F1D1D", hover_color="#991B1B",
                      text_color=TEXT, font=ctk.CTkFont(size=11),
                      command=self._reset_cycle).pack(side="right", padx=12, pady=8)

        ctk.CTkButton(hdr, text="+ Ciclu nou",
                      width=100, height=28,
                      fg_color=BG_MAIN, hover_color=BORDER,
                      border_width=1, border_color=ACCENT,
                      text_color=ACCENT, font=ctk.CTkFont(size=11),
                      command=self._new_cycle_dialog).pack(side="right", padx=4, pady=8)

    # ── Body (sidebar + content) ───────────────────────────────────────────────

    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=0, pady=0)

        # Sidebar
        self._sidebar = ctk.CTkScrollableFrame(
            body, width=190, fg_color=BG_SIDEBAR,
            corner_radius=0, scrollbar_button_color=BORDER,
        )
        self._sidebar.pack(side="left", fill="y", padx=0, pady=0)

        # Contingent label în sidebar
        self._contingent_lbl = ctk.CTkLabel(
            self._sidebar, text="Contingent:\n—",
            font=ctk.CTkFont(size=11), text_color=TEXT_DIM,
            justify="center",
        )
        self._contingent_lbl.pack(pady=(8, 4))

        ctk.CTkFrame(self._sidebar, height=1,
                     fg_color=BORDER).pack(fill="x", padx=8, pady=4)

        # Content panel frame
        self._content_frame = ctk.CTkFrame(
            body, fg_color=BG_MAIN, corner_radius=0,
        )
        self._content_frame.pack(side="left", fill="both", expand=True)

    # ── Timeline ──────────────────────────────────────────────────────────────

    def _rebuild_timeline(self, steps: list):
        for w in self._sidebar.winfo_children():
            if isinstance(w, TimelineStep):
                w.destroy()
        for step in steps:
            w = TimelineStep(self._sidebar, step,
                              on_click=self._on_step_clicked)
            w.pack(fill="x", pady=1)

    def _on_step_clicked(self, step_id: str):
        step = self.cs.get_step(step_id)
        if step:
            self._show_panel(step)

    # ── Panels ────────────────────────────────────────────────────────────────

    def _show_panel(self, step: dict):
        if self._current_panel:
            self._current_panel.destroy()
        if step["type"] == "termen":
            self._current_panel = TermenPanel(
                self._content_frame, step, self.cm,
                sector_filter=self._get_selected_sectors(),
                contingent_path=CONTINGENT_PATH,
                on_done=self._on_step_done,
            )
        else:
            self._current_panel = ReinterogarePanel(
                self._content_frame, step, self.cm,
                sector_filter=self._get_selected_sectors(),
                contingent_path=CONTINGENT_PATH,
                on_done=self._on_step_done,
            )
        self._current_panel.pack(fill="both", expand=True)

    # ── Events ────────────────────────────────────────────────────────────────

    def _on_step_done(self, step_id: str, stats: dict):
        self.cs.mark_completed(step_id, stats)
        self._refresh_ui()

        # Verifică contingent epuizat
        if stats.get("n_datornici", -1) == 0 or self._contingent_count() == 0:
            active = self.cs.get_active_step()
            if active and active["type"] == "reinterogare":
                messagebox.showwarning(
                    "Contingent epuizat",
                    "Toți datornicii au fost repuși. "
                    "Contingentul activ este gol."
                )

    def _on_sector_filter_changed(self):
        sectors = self._get_selected_sectors()
        self.cs.update_sector_filter(sectors)

    def _select_all_sectors(self):
        all_selected = all(v.get() for v in self._sector_vars.values())
        for v in self._sector_vars.values():
            v.set(not all_selected)
        self._on_sector_filter_changed()

    def _get_selected_sectors(self) -> list:
        return [s for s, v in self._sector_vars.items() if v.get()]

    def _contingent_count(self) -> int:
        try:
            return ContingentTracker(CONTINGENT_PATH).load().get("count", 0)
        except Exception:
            return 0

    def _refresh_contingent_label(self):
        try:
            state = ContingentTracker(CONTINGENT_PATH).load()
            count = state.get("count", 0)
            src = state.get("source", "")
            self._contingent_lbl.configure(
                text=f"Contingent:\n{count} dosare\n{src}"
            )
        except Exception:
            self._contingent_lbl.configure(text="Contingent:\n—")

    def _refresh_ui(self):
        state = self.cs.load()
        steps = state.get("steps", [])
        self._rebuild_timeline(steps)
        self._refresh_contingent_label()
        # Restaurează filtru sector din state
        saved_filter = state.get("sector_filter", [])
        for s, v in self._sector_vars.items():
            v.set(s in saved_filter)
        # Arată panelul pasului activ
        active = self.cs.get_active_step()
        if active:
            self._show_panel(active)
        elif not steps:
            self._show_no_cycle()

    def _show_no_cycle(self):
        if self._current_panel:
            self._current_panel.destroy()
        self._current_panel = ctk.CTkFrame(
            self._content_frame, fg_color="transparent"
        )
        ctk.CTkLabel(
            self._current_panel,
            text="Niciun ciclu activ.\nApasă „+ Ciclu nou" pentru a începe.",
            font=ctk.CTkFont(size=14), text_color=TEXT_DIM,
        ).pack(expand=True)
        self._current_panel.pack(fill="both", expand=True)

    def _new_cycle_dialog(self):
        dialog = ctk.CTkInputDialog(
            text="Data termenului de start (ex. 30.09.2025):",
            title="Ciclu nou",
        )
        val = dialog.get_input()
        if not val:
            return
        try:
            term_date = datetime.strptime(val.strip(), "%d.%m.%Y").date()
        except ValueError:
            messagebox.showerror("Eroare", f"Dată invalidă: '{val}'\nFormat: ZZ.LL.AAAA")
            return
        self.cs.init_cycle(term_date, self._get_selected_sectors())
        self._refresh_ui()

    def _reset_cycle(self):
        if messagebox.askyesno(
            "Reset ciclu",
            "Resetezi ciclul curent?\nStatisticile salvate vor fi șterse.",
        ):
            self.cs.reset()
            self._refresh_ui()


if __name__ == "__main__":
    app = LoopProcessorApp()
    app.mainloop()
```

**Step 2: Verifică sintaxa**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -c "
import ast
with open('loop_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

Expected: `Syntax OK`

**Step 3: Commit (fără paneluri încă)**

```bash
git add loop_processor.py
git commit -m "feat: add LoopProcessorApp shell with timeline sidebar and header"
```

---

## Task 5: `TermenPanel` — panelul Interogare Termen

**Files:**
- Modify: `APP_DIR/loop_processor.py` — adaugă clasa `TermenPanel` înainte de `LoopProcessorApp`

**Step 1: Adaugă `TermenPanel` în `loop_processor.py`** (înainte de `class LoopProcessorApp`)

```python
class TermenPanel(ctk.CTkFrame):
    """Panelul pentru pasul Interogare Termen."""

    def __init__(self, master, step: dict, cm: ConfigManager,
                 sector_filter: list, contingent_path: Path,
                 on_done, **kwargs):
        super().__init__(master, fg_color=BG_MAIN, **kwargs)
        self.step = step
        self.cm = cm
        self.sector_filter = sector_filter
        self.contingent_path = contingent_path
        self.on_done = on_done
        self._sector_pickers = {}
        self._debt_overrides = {}

        self._build()
        self._load_settings()

    def _build(self):
        from modules.ui_components import FilePicker, LogBox

        # Titlu pas
        ctk.CTkLabel(self, text=f"Interogare Termen — {self.step.get('label','')}",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=ACCENT).pack(anchor="w", padx=20, pady=(16,4))

        # Read-only badge dacă e completat
        if self.step.get("status") == "completed":
            ctk.CTkLabel(self, text="✓ Completat — statistici:",
                         text_color=DONE, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=20)
            self._build_stat_cards(self.step)
            return

        # Raport Stare
        rs = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8)
        rs.pack(fill="x", padx=16, pady=4)
        self.raport_picker = FilePicker(rs, label="Raport Stare")
        self.raport_picker.pack(fill="x", padx=10, pady=6)

        # Sectoare S1-S6
        sec_frame = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8)
        sec_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(sec_frame, text="Fișiere sector S1–S6",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT).pack(anchor="w", padx=10, pady=(6,2))
        for s in SECTORS:
            picker = FilePicker(
                sec_frame, label=s, clearable=True,
                filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All", "*.*")],
            )
            picker.pack(fill="x", padx=10, pady=1)
            self._sector_pickers[s] = picker

        # Output dir
        out = ctk.CTkFrame(self, fg_color="transparent")
        out.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(out, text="Output dir:", text_color=TEXT_DIM,
                     width=90, anchor="w").pack(side="left")
        self._output_entry = ctk.CTkEntry(out, width=340)
        self._output_entry.pack(side="left", padx=(0,4))
        ctk.CTkButton(out, text="...", width=32,
                      command=self._pick_output).pack(side="left")

        # Buton + progress
        self._run_btn = ctk.CTkButton(
            self, text="▶  Procesează Interogare Termen",
            height=40, fg_color=ACCENT, hover_color=DONE,
            text_color=BG_MAIN, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._run,
        )
        self._run_btn.pack(pady=(12,4))

        self._progress = ctk.CTkProgressBar(self, progress_color=ACCENT)
        self._progress.pack(fill="x", padx=16, pady=4)
        self._progress.set(0)

        # Stat cards (inițial gole)
        self._cards_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._cards_frame.pack(fill="x", padx=16, pady=4)
        self._card_susp = StatCard(self._cards_frame, "SUSPENDARE", C_SUSP, width=120)
        self._card_susp.pack(side="left", padx=4, expand=True, fill="x")
        self._card_rep  = StatCard(self._cards_frame, "REPUNERE", C_REP, width=120)
        self._card_rep.pack(side="left", padx=4, expand=True, fill="x")
        self._card_imp  = StatCard(self._cards_frame, "IMP. OBS.", C_IMP, width=120)
        self._card_imp.pack(side="left", padx=4, expand=True, fill="x")
        self._card_dat  = StatCard(self._cards_frame, "DATORNICI", C_DAT, width=120)
        self._card_dat.pack(side="left", padx=4, expand=True, fill="x")

        # Log
        self._log = LogBox(self, height=140)
        self._log.pack(fill="both", expand=True, padx=16, pady=(4,16))

    def _build_stat_cards(self, step: dict):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.pack(fill="x", padx=16, pady=8)
        for label, key, color in [
            ("SUSPENDARE","n_suspendare",C_SUSP),
            ("REPUNERE","n_repunere",C_REP),
            ("IMP. OBS.","n_import_obs",C_IMP),
            ("DATORNICI","n_datornici",C_DAT),
        ]:
            c = StatCard(f, label, color, width=120)
            c.pack(side="left", padx=4, expand=True, fill="x")
            c.set_value(step.get(key, 0))

    def _load_settings(self):
        s = self.cm.load_settings()
        for sec, picker in self._sector_pickers.items():
            v = s.get(f"loop_termen_{sec}")
            if v:
                picker.set(v)
        if s.get("loop_termen_raport"):
            self.raport_picker.set(s["loop_termen_raport"])
        if s.get("loop_termen_output") and hasattr(self, "_output_entry"):
            self._output_entry.insert(0, s["loop_termen_output"])

    def _pick_output(self):
        d = filedialog.askdirectory()
        if d:
            self._output_entry.delete(0, "end")
            self._output_entry.insert(0, d)

    def _run(self):
        from modules.termen_processor import TermenProcessor
        from modules.sector_address_filter import filter_by_sector_address

        sector_paths = {s: p.get() for s, p in self._sector_pickers.items() if p.get()}
        output_dir = self._output_entry.get().strip()
        raport_path = self.raport_picker.get()
        date_str = self.step["id"].replace("termen_", "")

        if not sector_paths:
            messagebox.showwarning("Atenție", "Selectează cel puțin un fișier sector.")
            return
        if not output_dir:
            messagebox.showwarning("Atenție", "Selectează directorul de output.")
            return
        if not raport_path:
            messagebox.showwarning("Atenție", "Selectează Raportul de Stare.")
            return

        try:
            term_date = datetime.strptime(date_str, "%d.%m.%Y").date()
        except ValueError:
            messagebox.showerror("Eroare", f"Dată invalidă în pasul ciclului: {date_str}")
            return

        # Save settings
        self.cm.save_setting("loop_termen_raport", raport_path)
        self.cm.save_setting("loop_termen_output", output_dir)
        for s, p in self._sector_pickers.items():
            self.cm.save_setting(f"loop_termen_{s}", p.get())

        self._run_btn.configure(state="disabled")
        self._log.clear()
        self._progress.set(0.05)

        def _worker():
            # Aplică filtru sector adresă
            _sector_filter = self.sector_filter

            def _progress_cb(msg):
                self.after(0, lambda: self._log.append(msg))
                self.after(0, lambda: self._progress.set(
                    min(self._progress.get() + 0.1, 0.9)))

            processor = TermenProcessor(
                self.cm,
                progress_callback=_progress_cb,
                contingent_path=self.contingent_path,
            )
            # Monkey-patch raport loader to apply sector filter
            processor._sector_filter = _sector_filter

            result = processor.process(
                sector_paths=sector_paths,
                term_date=term_date,
                output_dir=Path(output_dir),
                raport_stari_path=Path(raport_path),
            )

            # Mută Datornici în subfolder "Contingent datornici"
            if result.get("success"):
                try:
                    _move_datornici_to_subfolder(Path(output_dir), date_str)
                except Exception as e:
                    result.setdefault("warnings", []).append(
                        f"Contingent subfolder: {e}"
                    )

            self.after(0, lambda: self._on_done(result, output_dir))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self, result: dict, output_dir: str):
        self._run_btn.configure(state="normal")
        self._progress.set(1.0)

        if not result.get("success"):
            err = result.get("error", "Eroare necunoscută")
            messagebox.showerror("Eroare", err)
            self._log.append(f"EROARE: {err}")
            return

        n_s = result.get("n_suspendare", 0)
        n_r = result.get("n_repunere", 0)
        n_i = result.get("n_import_obs", 0)
        n_d = result.get("n_datornici", 0)

        self._card_susp.set_value(n_s)
        self._card_rep.set_value(n_r)
        self._card_imp.set_value(n_i)
        self._card_dat.set_value(n_d)

        for w in result.get("warnings", []):
            self._log.append(f"  [warn] {w}")
        self._log.append(f"\nOutput: {output_dir}")

        stats = {
            "n_suspendare": n_s, "n_repunere": n_r,
            "n_import_obs": n_i, "n_datornici": n_d,
            "output_dir": output_dir,
        }
        self.on_done(self.step["id"], stats)

        if messagebox.askyesno("Gata!", "Deschizi folderul output?"):
            os.startfile(output_dir)
```

Adaugă și funcția helper `_move_datornici_to_subfolder` la nivel de modul (înainte de clase):

```python
def _move_datornici_to_subfolder(output_dir: Path, date_str: str):
    """Mută Datornici termen {date_str}.xlsx în subfolder 'Contingent datornici'."""
    import shutil
    src = output_dir / f"Datornici termen {date_str}.xlsx"
    if src.exists():
        dest_dir = output_dir / "Contingent datornici"
        dest_dir.mkdir(exist_ok=True)
        shutil.move(str(src), str(dest_dir / src.name))
```

**Step 2: Verifică sintaxa**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -c "
import ast
with open('loop_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

**Step 3: Commit**

```bash
git add loop_processor.py
git commit -m "feat: add TermenPanel with stat cards and contingent subfolder"
```

---

## Task 6: `ReinterogarePanel` — panelul Reinterogare Lunară

**Files:**
- Modify: `APP_DIR/loop_processor.py` — adaugă `ReinterogarePanel` înainte de `LoopProcessorApp`

**Step 1: Adaugă `ReinterogarePanel`**

```python
class ReinterogarePanel(ctk.CTkFrame):
    """Panelul pentru pasul Reinterogare Lunară."""

    def __init__(self, master, step: dict, cm: ConfigManager,
                 sector_filter: list, contingent_path: Path,
                 on_done, **kwargs):
        super().__init__(master, fg_color=BG_MAIN, **kwargs)
        self.step = step
        self.cm = cm
        self.sector_filter = sector_filter
        self.contingent_path = contingent_path
        self.on_done = on_done
        self._sector_pickers = {}

        self._build()
        self._load_settings()

    def _build(self):
        from modules.ui_components import FilePicker, LogBox

        ctk.CTkLabel(self, text=f"Reinterogare — {self.step.get('label','')}",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=ACCENT).pack(anchor="w", padx=20, pady=(16,4))

        if self.step.get("status") == "completed":
            ctk.CTkLabel(self, text="✓ Completat — statistici:",
                         text_color=DONE, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=20)
            self._build_stat_cards(self.step)
            return

        # Contingent source
        cont_frame = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8,
                                   border_width=1, border_color=BORDER)
        cont_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(cont_frame, text="Contingent datornici",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT).pack(anchor="w", padx=10, pady=(8,2))

        self._cont_source = ctk.StringVar(value="cycle")

        # Check dacă există contingent din ciclu
        try:
            from modules.contingent_tracker import ContingentTracker
            state = ContingentTracker(self.contingent_path).load()
            count = state.get("count", 0)
            source_lbl = f"Din ciclu curent ({count} datornici — {state.get('source','')})"
            cycle_available = count > 0
        except Exception:
            source_lbl = "Din ciclu curent (indisponibil)"
            cycle_available = False

        rb_cycle = ctk.CTkRadioButton(
            cont_frame, text=source_lbl,
            variable=self._cont_source, value="cycle",
            text_color=TEXT if cycle_available else TEXT_DIM,
            state="normal" if cycle_available else "disabled",
            fg_color=ACCENT, hover_color=DONE,
            command=self._on_source_changed,
        )
        rb_cycle.pack(anchor="w", padx=16, pady=2)

        extern_row = ctk.CTkFrame(cont_frame, fg_color="transparent")
        extern_row.pack(fill="x", padx=10, pady=(2,8))
        rb_ext = ctk.CTkRadioButton(
            extern_row, text="Fișier extern",
            variable=self._cont_source, value="external",
            text_color=TEXT, fg_color=ACCENT, hover_color=DONE,
            command=self._on_source_changed,
        )
        rb_ext.pack(side="left")
        self._ext_picker = FilePicker(
            extern_row, label="",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All","*.*")],
        )
        self._ext_picker.pack(side="left", fill="x", expand=True, padx=(8,0))

        if not cycle_available:
            self._cont_source.set("external")
        self._on_source_changed()

        # Raport Stare
        rs = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8)
        rs.pack(fill="x", padx=16, pady=4)
        self.raport_picker = FilePicker(rs, label="Raport Stare")
        self.raport_picker.pack(fill="x", padx=10, pady=6)

        # Sectoare S1-S6
        sec_frame = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=8)
        sec_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(sec_frame, text="Fișiere sector S1–S6",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=TEXT).pack(anchor="w", padx=10, pady=(6,2))
        for s in SECTORS:
            picker = FilePicker(
                sec_frame, label=s, clearable=True,
                filetypes=[("Excel files","*.xlsx *.xls *.xlsm"),("All","*.*")],
            )
            picker.pack(fill="x", padx=10, pady=1)
            self._sector_pickers[s] = picker

        # Output dir
        out = ctk.CTkFrame(self, fg_color="transparent")
        out.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(out, text="Output dir:", text_color=TEXT_DIM,
                     width=90, anchor="w").pack(side="left")
        self._output_entry = ctk.CTkEntry(out, width=340)
        self._output_entry.pack(side="left", padx=(0,4))
        ctk.CTkButton(out, text="...", width=32,
                      command=self._pick_output).pack(side="left")

        # Buton + progress
        self._run_btn = ctk.CTkButton(
            self, text="▶  Procesează Reinterogare",
            height=40, fg_color=ACCENT, hover_color=DONE,
            text_color=BG_MAIN, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._run,
        )
        self._run_btn.pack(pady=(12,4))

        self._progress = ctk.CTkProgressBar(self, progress_color=ACCENT)
        self._progress.pack(fill="x", padx=16, pady=4)
        self._progress.set(0)

        # Stat cards
        self._cards_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._cards_frame.pack(fill="x", padx=16, pady=4)
        self._card_rep = StatCard(self._cards_frame, "REPUNERE", C_REP, width=160)
        self._card_rep.pack(side="left", padx=4, expand=True, fill="x")
        self._card_imp = StatCard(self._cards_frame, "IMP. OBS.", C_IMP, width=160)
        self._card_imp.pack(side="left", padx=4, expand=True, fill="x")
        self._card_cont = StatCard(self._cards_frame, "CONTINGENT RĂMAS", C_DAT, width=160)
        self._card_cont.pack(side="left", padx=4, expand=True, fill="x")

        # Log
        self._log = LogBox(self, height=120)
        self._log.pack(fill="both", expand=True, padx=16, pady=(4,16))

    def _build_stat_cards(self, step: dict):
        f = ctk.CTkFrame(self, fg_color="transparent")
        f.pack(fill="x", padx=16, pady=8)
        for label, key, color in [
            ("REPUNERE","n_repunere",C_REP),
            ("IMP. OBS.","n_import_obs",C_IMP),
        ]:
            c = StatCard(f, label, color, width=160)
            c.pack(side="left", padx=4, expand=True, fill="x")
            c.set_value(step.get(key, 0))

    def _on_source_changed(self):
        is_ext = self._cont_source.get() == "external"
        # Activează/dezactivează picker extern
        state = "normal" if is_ext else "disabled"
        for child in self._ext_picker.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass

    def _load_settings(self):
        s = self.cm.load_settings()
        for sec, picker in self._sector_pickers.items():
            v = s.get(f"loop_reintr_{sec}")
            if v:
                picker.set(v)
        if s.get("loop_reintr_raport"):
            self.raport_picker.set(s["loop_reintr_raport"])
        if s.get("loop_reintr_output") and hasattr(self, "_output_entry"):
            self._output_entry.insert(0, s["loop_reintr_output"])

    def _pick_output(self):
        d = filedialog.askdirectory()
        if d:
            self._output_entry.delete(0, "end")
            self._output_entry.insert(0, d)

    def _run(self):
        from modules.processor import ReinterogareProcessor
        from modules.contingent_tracker import ContingentTracker

        sector_paths = {s: p.get() for s, p in self._sector_pickers.items() if p.get()}
        output_dir = self._output_entry.get().strip()
        raport_path = self.raport_picker.get()

        if not sector_paths:
            messagebox.showwarning("Atenție", "Selectează cel puțin un fișier sector.")
            return
        if not output_dir:
            messagebox.showwarning("Atenție", "Selectează directorul de output.")
            return
        if not raport_path:
            messagebox.showwarning("Atenție", "Selectează Raportul de Stare.")
            return

        # Contingent extern dacă e selectat
        if self._cont_source.get() == "external":
            ext_path = self._ext_picker.get()
            if not ext_path:
                messagebox.showwarning("Atenție", "Selectează fișierul contingent extern.")
                return
            try:
                import pandas as pd
                dat_df = pd.read_excel(ext_path, dtype=str)
                tracker = ContingentTracker(self.contingent_path)
                # Detectează data din nume fișier sau folosește azi
                from datetime import date as _date
                tracker.init_from_termen(dat_df, _date.today())
            except Exception as e:
                messagebox.showerror("Eroare contingent extern", str(e))
                return

        # Determină run_date din label pasului
        label = self.step.get("label", "")
        try:
            from datetime import datetime as _dt
            # Label: "Reintr. Oct 2025" → prima zi a lunii
            parts = label.replace("Reintr. ", "").split()
            LUNI_EN = {"Ian":1,"Feb":2,"Mar":3,"Apr":4,"Mai":5,"Iun":6,
                       "Iul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
            run_date = _date(int(parts[1]), LUNI_EN[parts[0]], 1)
        except Exception:
            run_date = _date.today()

        self.cm.save_setting("loop_reintr_raport", raport_path)
        self.cm.save_setting("loop_reintr_output", output_dir)
        for s, p in self._sector_pickers.items():
            self.cm.save_setting(f"loop_reintr_{s}", p.get())

        self._run_btn.configure(state="disabled")
        self._log.clear()
        self._progress.set(0.05)

        def _worker():
            def _progress_cb(msg):
                self.after(0, lambda: self._log.append(msg))
                self.after(0, lambda: self._progress.set(
                    min(self._progress.get() + 0.1, 0.9)))

            processor = ReinterogareProcessor(self.cm, progress_callback=_progress_cb)
            result = processor.process(
                sector_paths=sector_paths,
                run_date=run_date,
                output_path=Path(output_dir) / f"Reinterogare_{run_date.strftime('%m_%Y')}_output.xlsx",
                raport_stari_path=Path(raport_path),
            )

            # Curăță contingentul
            if result.get("success") and Path(raport_path).exists():
                try:
                    import pandas as _pd
                    _rs = _pd.read_excel(raport_path, dtype=str)
                    ContingentTracker(self.contingent_path).clean_against_raport(_rs, run_date)
                except Exception:
                    pass

            self.after(0, lambda: self._on_done(result, output_dir))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self, result: dict, output_dir: str):
        self._run_btn.configure(state="normal")
        self._progress.set(1.0)

        if not result.get("success"):
            err = result.get("error", "Eroare necunoscută")
            messagebox.showerror("Eroare", err)
            self._log.append(f"EROARE: {err}")
            return

        n_r = result.get("consolidat_rows", 0)
        for w in result.get("warnings", []):
            self._log.append(f"  [warn] {w}")
        self._log.append(f"\nOutput: {output_dir}")

        # Contingent rămas
        try:
            from modules.contingent_tracker import ContingentTracker
            cont_count = ContingentTracker(self.contingent_path).load().get("count", 0)
        except Exception:
            cont_count = 0

        self._card_rep.set_value(result.get("consolidat_rows", 0))
        self._card_cont.set_value(cont_count)

        stats = {"n_repunere": n_r, "n_import_obs": 0, "output_dir": output_dir}
        self.on_done(self.step["id"], stats)

        if messagebox.askyesno("Gata!", "Deschizi folderul output?"):
            os.startfile(output_dir)
```

**Step 2: Verifică sintaxa**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -c "
import ast
with open('loop_processor.py', encoding='utf-8') as f:
    ast.parse(f.read())
print('Syntax OK')
"
```

Expected: `Syntax OK`

**Step 3: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate PASS (9 teste: 5 cycle_state + 4 sector_filter)

**Step 4: Commit**

```bash
git add loop_processor.py
git commit -m "feat: add ReinterogarePanel with contingent source selector"
```

---

## Task 7: Integrare filtru sector adresă în TermenProcessor și ReinterogareProcessor

**Files:**
- Modify: `APP_DIR/modules/termen_processor.py` — aplică filtrul după load Raport Stare
- Modify: `APP_DIR/modules/processor.py` — aplică filtrul după load Raport Stare

**Step 1: Modifică `termen_processor.py`**

În metoda `process`, după `base_df = RaportStariLoader().load(...)`, adaugă:

```python
from modules.sector_address_filter import filter_by_sector_address
if hasattr(self, '_sector_filter') and self._sector_filter:
    base_df = filter_by_sector_address(base_df, self._sector_filter)
```

**Step 2: Modifică `processor.py` (ReinterogareProcessor)**

Același pattern — în `process`, după `base_df = RaportStariLoader().load(...)`:

```python
from modules.sector_address_filter import filter_by_sector_address
if hasattr(self, '_sector_filter') and self._sector_filter:
    base_df = filter_by_sector_address(base_df, self._sector_filter)
```

**Step 3: În `loop_processor.py`, setează filtrul pe processor înainte de `process()`**

În `TermenPanel._run._worker` (deja schelet):
```python
processor._sector_filter = _sector_filter
```
În `ReinterogarePanel._run._worker`:
```python
processor._sector_filter = self.sector_filter
```

**Step 4: Verifică sintaxa ambelor module**

```bash
cd "C:/EXCELURI/Interogari termen/Loop Processor"
python -c "import modules.termen_processor; import modules.processor; print('OK')"
```

**Step 5: Rulează testele**

```bash
python -m pytest tests/ -v
```

**Step 6: Commit**

```bash
git add modules/termen_processor.py modules/processor.py loop_processor.py
git commit -m "feat: wire global sector address filter into both processors"
```

---

## Checklist final

- [ ] `CycleState.init_cycle` generează 7 pași pentru un termen 30.09
- [ ] `CycleState.mark_completed` activează pasul următor
- [ ] `filter_by_sector_address` filtrează corect după sector adresă
- [ ] Timeline sidebar afișează pași colorați (○ / ● / ✓)
- [ ] Click pe pas completat → statistici read-only în panel dreapta
- [ ] `TermenPanel`: file pickers + stat cards + progress
- [ ] `ReinterogarePanel`: radio buttons contingent (ciclu / extern) + stat cards
- [ ] Subfolder „Contingent datornici\" creat automat după Interogare Termen
- [ ] Filtrul global sector adresă aplicat la ambii procesori
- [ ] Reset ciclu curăță `cycle_state.json`
- [ ] Toate testele pytest trec
