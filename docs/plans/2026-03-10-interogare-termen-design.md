# Interogare Termen — Design Document

**Data**: 2026-03-10
**Status**: Aprobat

## Scop

Aplicație Python/CustomTkinter separată pentru procesarea interogărilor trimestriale legale (30.03 / 30.09) la toate sectoarele DGITL. Reutilizează modulele din Reinterogare Processor și menține un contingent activ de datornici, care constituie feedul pentru Reinterogările lunare ulterioare.

---

## Contextul problemei

Interogarea la termen (legală, de 2 ori/an) verifică **toți** beneficiarii din platformă față de registrele DGITL ale celor 6 sectoare. Rezultatele determină:
- Suspendarea celor cu datorii nou-descoperite
- Repunerea celor suspendați care nu mai au datorii
- Importul observațiilor pentru toți ceilalți

Datornicii confirmați la termen → **contingent activ** → feed pentru Reinterogările lunare (Oct–Feb / Apr–Aug), care îl restrâng progresiv prin repuneri lunare sau achitări directe.

---

## Matricea de decizie

| Stare Cerere | Motiv suspendare | Datorii la termen | Acțiune |
|---|---|---|---|
| Aprobat | — | fără datorii | → **Import Observatii** |
| Aprobat | — | cu datorii | → **Suspendare** |
| Suspendat | datorii (orice tip: termen SAU lunar) | fără datorii | → **Repunere** |
| Suspendat | datorii (orice tip) | cu datorii | → **Import Observatii** |
| Suspendat | alt motiv (non-datorii) | orice | → **Import Observatii** |
| Încetat | orice | orice | → **Import Observatii** |

**Acțiunile sunt mutual exclusive**: dacă o cerere intră în Suspendare sau Repunere, nu apare în Import Observatii.

---

## Fluxul de date

```
Interogare Termen 30.09.2025
    ├── Suspendare.xlsx          → import în platformă
    ├── Repunere.xlsx            → import în platformă
    ├── Import Observatii.xlsx   → import în platformă
    └── Datornici termen 30.09.2025.xlsx
              │  (Suspendare ∪ Suspendat-cu-datorii)
              ▼
    contingent.json  ← init
              │
    Reinterogare Oct 2025
    ├── Repuși → ies din contingent
    ├── Achitări directe → ies din contingent (detectate prin Raport Stare = Aprobat)
    └── contingent.json  ← actualizat
              │
    Reinterogare Nov 2025 ... Dec 2025
              │
    Interogare Termen 30.03.2026
    └── contingent.json  ← reinit
```

---

## Fișiere sector S1–S6

- Format: XLS exportat din DGITL, sheet `RptRaportExportFisierPersoane`
- Coloana cheie: `AreNu / AreDatorii` (redenumită S1, S2... în Power Query)
- Join pe: CNP (detectat fuzzy, ca în Reinterogare)
- Detecție coloană debt: alias nou `"are.?nu.*are.*datorii"` adăugat în `column_aliases.json`

---

## Output files

### Suspendare.xlsx
Coloane: `Numar cerere`, `Data Cerere`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `CNP adult cu handicap`, `Motiv suspendare`, `Data suspendare`
- Motiv: `"Interogare termen {DD.MM.YYYY} / datorii la S{n} S{m}"`
- Data suspendare: prima zi a lunii următoare termenului (ex. 01.10.2025 pentru termen 30.09.2025)

### Repunere.xlsx
Coloane: `Numar cerere`, `Data Cerere`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `CNP adult cu handicap`, `Motiv repunere`, `Data repunere`
- Motiv: `"Interogare termen {DD.MM.YYYY} / Fara datorii"`
- Data repunere: prima zi a lunii termenului (ex. 01.09.2025 pentru termen 30.09.2025)

### Import Observatii.xlsx
Coloane: `Nume adult cu handicap`, `Prenume adult cu handicap`, `CNP adult cu handicap`, `Nr. cerere`, `Observatii`
- Observatie (cu datorii): `"Datorii interogare termen {DD.MM.YYYY} S{n}"`
- Observatie (fără datorii): `"Fara datorii interogare termen {DD.MM.YYYY}"`

### Datornici termen {DD.MM.YYYY}.xlsx
Coloane: `Numar cerere`, `CNP adult cu handicap`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `Sectoare datorii`, `Stare Cerere`, `Motiv`
- Conținut: Suspendare ∪ Suspendat-cu-datorii-confirmate
- Feed direct pentru Reinterogare lunară

---

## Structura fișiere proiect

```
Reinterogare Processor\
├── interogare_termen_processor.py     NOU — UI principal
├── contingent.json                    NOU — stare contingent activ
├── modules\
│   ├── termen_logic.py                NOU — matricea de decizie
│   ├── contingent_tracker.py          NOU — read/write contingent
│   ├── sector_loader.py               REFOLOSIT (cu alias nou pentru AreNu/AreDatorii)
│   ├── output_builder.py              REFOLOSIT / extins cu write_termen_outputs()
│   ├── config_manager.py              REFOLOSIT
│   └── ui_components.py              REFOLOSIT
└── config\
    └── column_aliases.json            MODIFICAT — alias AreNu/AreDatorii
```

---

## UI — interogare_termen_processor.py

```
┌─────────────────────────────────────────────────┐
│  Interogare Termen                               │
│  Data termen  [30.09.2025    ]                   │
│                                                  │
│  Raport Stare  [___________] [...]               │
│                                                  │
│  S1 [___________] [...] [X]                      │
│  S2 [___________] [...] [X]                      │
│  S3 [___________] [...] [X]                      │
│  S4 [___________] [...] [X]                      │
│  S5 [___________] [...] [X]                      │
│  S6 [___________] [...] [X]                      │
│                                                  │
│  Output dir   [___________] [...]                │
│  [ ▶ Procesează ]                                │
│                                                  │
│  Contingent activ: 1.247 beneficiari             │
│  (ultima actualizare: 30.09.2025)                │
│                                                  │
│  [log box]                                       │
└─────────────────────────────────────────────────┘
```

---

## Contingent tracker

`contingent.json`:
```json
{
  "updated_at": "2025-10-01",
  "source": "Interogare termen 30.09.2025",
  "count": 1247,
  "entries": [
    {
      "cnp": "1234567890123",
      "numar_cerere": "123/456",
      "sectoare_datorii": ["S2", "S4"],
      "data_intrare": "2025-10-01",
      "stare_initiala": "Aprobat"
    }
  ]
}
```

**Actualizare din Reinterogare lunară**: la fiecare rulare, Reinterogare compară CNP-urile din contingent cu Raport Stare curent și elimină automat pe cei cu `Stare = Aprobat`.

---

## Integrare cu Reinterogare Processor

- `column_aliases.json`: alias nou `"are.?nu.*are.*datorii"` pentru detecția coloanei S1-S6
- `reinterogare_processor.py`: la finalizarea procesării, dacă contingent.json există, curăță automat contingentul față de Raport Stare curent
- Picker „Inter. Termen" din Reinterogare: acceptă `Datornici termen {DD.MM.YYYY}.xlsx`

---

## Testing

- `tests/test_termen_logic.py` — matricea de decizie (6 scenarii)
- `tests/test_contingent_tracker.py` — init, update, curățare automată
- `tests/test_output_builder_termen.py` — output files format corect
