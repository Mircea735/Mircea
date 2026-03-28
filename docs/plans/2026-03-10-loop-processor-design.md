# Loop Processor — Design Document

**Data**: 2026-03-10
**Status**: Aprobat

## Scop

Aplicație Python/CustomTkinter care integrează Interogare Termen și Reinterogare Lunară într-un singur wizard secvențial cu UI vizual, dark mode, și memorie persistentă a ciclului semestriar.

---

## Contextul problemei

Utilizatorul lucrează cu două aplicații separate (Interogare Termen Processor + Reinterogare Processor) care trebuie rulate în ordine specifică pe parcursul unui ciclu de 6 luni. App-ul integrat ghidează utilizatorul prin fiecare pas, previne erori de ordine, și oferă vizualizare unificată a statisticilor per run și per ciclu.

---

## Locație

```
C:\EXCELURI\Interogari termen\Loop Processor\
```

---

## UI — Layout general

```
┌──────────────────────────────────────────────────────────┐
│  Ciclu Datornici DGITL                    [Reset ciclu]  │
│  Sector adresă: [✓S1][✓S2][□S3][□S4][□S5][□S6]  [Tot]  │
├──────────────┬───────────────────────────────────────────┤
│  TIMELINE    │  PANEL ACTIV                              │
│              │                                           │
│ ●─ IT 30.09  │  [conținut dinamic per pas]               │
│ │  ✓ gata    │                                           │
│ ●─ Reintr.   │  [file pickers]                           │
│ │  Oct 2025  │  [▶ Procesează]                          │
│ ○  Nov 2025  │  ████████░░░░  62%                       │
│ ○  Dec 2025  │                                           │
│ ○  Ian 2026  │  ┌────────┐┌────────┐┌────────┐┌───────┐ │
│ ○  Feb 2026  │  │  842   ││  112   ││ 8.293  ││  954  │ │
│ ●─ IT 30.03  │  │SUSPEND.││REPUNERE││IMP.OBS.││DATORN.│ │
│              │  └────────┘└────────┘└────────┘└───────┘ │
│ Contingent:  │                                           │
│ 954 dosare   │  [log text box]                           │
└──────────────┴───────────────────────────────────────────┘
```

---

## Paleta vizuală (dark mode)

| Element | Culoare | Hex |
|---|---|---|
| Fundal principal | Navy închis | `#0D1117` |
| Sidebar timeline | Navy mediu | `#161B22` |
| Carduri/paneluri | Navy deschis | `#1C2128` |
| Accent primar (activ) | Teal | `#2DD4BF` |
| Accent completat | Indigo | `#6366F1` |
| Warning | Amber | `#F59E0B` |
| Contor Suspendare | Roșu | `#EF4444` |
| Contor Repunere | Verde | `#22C55E` |
| Contor Import Obs. | Albastru | `#3B82F6` |
| Contor Datornici | Galben | `#F59E0B` |
| Borduri | Gri subtil | `#30363D` |

**Timeline indicatori:**
- `○` pending → `#374151`
- `●` activ → `#2DD4BF` (pulsând)
- `✓` completat → `#6366F1`

---

## Arhitectură

### Fișiere proiect

```
Loop Processor\
├── loop_processor.py          ← app principal (UI + wizard logic)
├── modules\
│   ├── termen_logic.py        ← copiat din Interogare Termen Processor
│   ├── termen_processor.py    ← copiat
│   ├── contingent_tracker.py  ← copiat
│   ├── processor.py           ← copiat din Reinterogare Processor
│   ├── debt_logic.py          ← copiat
│   ├── sector_loader.py       ← shared
│   ├── raport_stari_loader.py ← shared (cu filtru sector adresă)
│   ├── merger.py              ← shared
│   ├── output_builder.py      ← shared
│   ├── config_manager.py      ← shared
│   ├── column_detector.py     ← shared
│   └── ui_components.py       ← shared
├── cycle_state.json           ← stare ciclu activ
├── contingent.json            ← contingent activ
└── config\
    ├── column_aliases.json
    └── settings.json
```

### `cycle_state.json`

```json
{
  "term_start": "30.09.2025",
  "term_end": "30.03.2026",
  "sector_filter": ["S1", "S2"],
  "steps": [
    {
      "id": "termen_30.09.2025",
      "type": "termen",
      "status": "completed",
      "n_suspendare": 842,
      "n_repunere": 112,
      "n_import_obs": 8293,
      "n_datornici": 954,
      "output_dir": "C:\\..."
    },
    {
      "id": "reinterogare_Oct_2025",
      "type": "reinterogare",
      "status": "completed",
      "n_repunere": 67,
      "n_import_obs": 887,
      "contingent_source": "cycle"
    },
    {
      "id": "reinterogare_Nov_2025",
      "type": "reinterogare",
      "status": "pending",
      "contingent_source": null
    }
  ]
}
```

---

## Filtru sector adresă (global)

- Checkboxes S1–S6 în header, persistent în `settings.json`
- Aplicat ca pre-filter în `RaportStariLoader`: înainte de merge, păstrează doar rândurile unde `Sector` (extras din coloana adresă) ∈ sectoare selectate
- Dacă niciun sector selectat → procesează toți beneficiarii (comportament implicit)
- Filtrul se aplică la **ambii** procesori (Interogare Termen + Reinterogare)

---

## Output Contingent datornici

La finalizarea Interogare Termen, subfolder automat în output dir:

```
Output dir\
├── Suspendare.xlsx
├── Repunere.xlsx
├── Import Observatii.xlsx
└── Contingent datornici\
    └── Datornici termen 30.09.2025.xlsx
```

---

## Comportament Wizard

### Inițializare ciclu

1. Utilizatorul setează data termenului de start (ex. `30.09.2025`)
2. App generează automat pașii: `IT 30.09` → `Oct` → `Nov` → `Dec` → `Ian` → `Feb` → `IT 30.03`
3. Pasul 1 (Interogare Termen) devine activ, restul locked

### Tranziții între pași

**Interogare Termen completată:**
- `contingent.json` inițializat
- Subfolder „Contingent datornici\" creat cu `Datornici termen.xlsx`
- Pasul următor devine activ
- Mesaj: „Contingent inițializat: 954 datornici. Poți începe Reinterogarea Oct."

**Reinterogare lunară completată:**
- `contingent.json` curățat automat (elimină Aprobații)
- Contoare actualizate în timeline
- Pasul următor devine activ
- Dacă contingent = 0 → alertă amber: „Contingent epuizat — toți datornicii repuși"

**Click pe pas completat:**
- Panel dreapta arată statisticile acelui run (read-only)
- Buton „Re-rulează" disponibil

### Reset ciclu

- Buton discret în header
- Confirmă cu dialog → `cycle_state.json` resetat, timeline regenerat pentru noul termen
- `contingent.json` NU se șterge automat

---

## Contingent la Reinterogare — sursă flexibilă

Panelul Reinterogare Lunară permite alegerea sursei contingentului:

```
┌─ Contingent datornici ──────────────────────────────────┐
│  ◉ Din ciclu curent  (954 datornici — IT 30.09.2025)    │
│  ○ Fișier extern     [_______________________] [...]    │
└─────────────────────────────────────────────────────────┘
```

- **Din ciclu curent**: folosește `contingent.json` local
- **Fișier extern**: utilizatorul selectează `Datornici termen {DD.MM.YYYY}.xlsx` de oriunde; app-ul îl importă și inițializează contingentul
- Dacă nu există ciclu activ → opțiunea „Din ciclu curent" e disabled, fișier extern obligatoriu
- Permite pornirea din mijlocul ciclului fără a fi rulat app-ul anterior

---

## Edge cases

| Situație | Comportament |
|---|---|
| Reinterogare fără Interogare Termen anterioară | Warning amber, fișier extern contingent obligatoriu |
| Sector DGITL lipsă la un pas | Warning per sector, procesare continuă cu cele disponibile |
| Contingent epuizat după Reinterogare | Alertă amber, pasul următor rămâne disponibil |
| Output dir blocat (Excel deschis) | Eroare clară, procesare nu continuă |

---

## Testing

- `tests/test_cycle_state.py` — inițializare, tranziții, reset
- `tests/test_sector_filter.py` — filtrul de sector adresă aplicat corect
- `tests/test_contingent_source.py` — sursă ciclu vs. fișier extern
