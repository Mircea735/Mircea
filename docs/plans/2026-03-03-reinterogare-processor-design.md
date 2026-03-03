# Reinterogare Processor — Design Document
**Data**: 2026-03-03
**Status**: Aprobat de utilizator
**Scop**: Aplicație Python/CustomTkinter care înlocuiește/paralelizează procesul Power Query pentru reinterogările lunare DGITL sector 2

---

## 1. Contextul problemei

**Procesul actual** (brittle):
- 6 fișiere Excel lunare (S1–S6), furnizate de DGITL, cu denumiri de coloane schimbate aleator de operatori
- Import manual în Power Query → Merge1 → funcție AdaptiveDebtStatus → output
- La fiecare lună nouă: coloane redenumite → query-urile cad → fix manual

**Probleme principale**:
1. Denumiri de coloane inconsistente lunar (ex: `Achitat la data - Observații2` → `Data plății - Observații`)
2. Motive de suspendare scrise arbitrar de operatori ("mocirla semitotala")
3. Fișierele S1–S6 se deschid manual, se curăță în Power Query, se salvează — flux lent și greșit-prone
4. Lipsă de vizibilitate pe datele brute înainte de output

**Nevoile utilizatorului**:
- Singur utilizator, output în Excel (nu web)
- Output trebuie să mapeze pe structura de coloane a **platformei** (repunere + import observații)
- Vizibilitate: cele 6 sectoare afișate separat + consolidat (pentru inspecție vizuală)
- Coloana consolidată finală cu toate cele 6 sectoare
- Corelare cu **interogarea termen** (30.03/30.09) care a generat suspendarea

---

## 2. Arhitectura generală

```
┌─────────────────────────────────────────────┐
│           Reinterogare Processor UI         │
│           (CustomTkinter, single window)    │
└─────────────────┬───────────────────────────┘
                  │
         ┌────────▼────────┐
         │   Input Layer   │
         │  6 × S1–S6.xlsx │ ← drag & drop sau file picker
         │  Raport Stări   │ ← pentru corelarea suspendărilor
         │  Interogare     │ ← pentru corelare interogare termen
         │  Termen (opt.)  │
         └────────┬────────┘
                  │
         ┌────────▼────────┐
         │  Detection Layer│
         │ fuzzy col detect│ ← detectează coloanele cheie indiferent de denumire
         │ motiv classifier│ ← grupare fuzzy motive suspendare
         └────────┬────────┘
                  │
         ┌────────▼────────┐
         │  Logic Layer    │
         │ debt status calc│ ← logica AdaptiveDebtStatus în Python
         │ repunere detect │ ← data repunere vs. data datorii
         │ suspendare corr.│ ← corelat cu interogare termen
         └────────┬────────┘
                  │
         ┌────────▼────────┐
         │  Output Layer   │
         │ Excel: S1..S6   │ ← sheet-uri separate pentru inspecție
         │ + CONSOLIDAT    │ ← toate sectoarele, structura platformei
         │ + REPUNERI      │ ← candidați repunere
         │ + IMPORT OBS    │ ← format import observații platformă
         └─────────────────┘
```

---

## 3. Componente principale

### 3.1 UI (CustomTkinter)
- **Tab 1 — Input**: 6 file picker-e (S1–S6) + Raport Stări + Interogare Termen (opțional)
- **Tab 2 — Coloane detectate**: tabel cu coloana cheie detectată vs. denumirea originală din fișier (permite corecție manuală)
- **Tab 3 — Motive**: lista de motive detectate + grupare fuzzy + confirmare utilizator ce înseamnă "datorii DITL"
- **Tab 4 — Output**: progress bar + log + buton deschidere fișier output

### 3.2 Detector de coloane (fuzzy)
Fiecare sector are **coloane cheie** care variază ca denumire. Detectorul le caută prin:
1. **Match exact** în dicționar de alias-uri știute (din runs anterioare)
2. **Regex** pe pattern-uri (ex: `are.?nu.?are.?debit`, `achitat|platit`, `data.?repunere`)
3. **Fuzzy** (difflib SequenceMatcher) pentru denumiri apropiate

Coloane cheie detectate:
| Concept | Exemple denumiri văzute |
|---------|------------------------|
| Are/Nu are debit | `Are/Nu are debit la 31.12.2025`, `ARE/NU ARE DEBIT la 01.01.2026` |
| Data plată | `Achitat la data - Observații2`, `Data plății - Observații` |
| Data repunere | `Data repunere`, `Data repunerii` |
| CNP | `CNP`, `Cod Numeric Personal` |
| Număr cerere | `Nr. Cerere`, `Numar cerere` |

Dacă detectorul e nesigur → UI arată coloana cu semn de întrebare + dropdown pentru selecție manuală.
Rezultatele se salvează în `config/column_aliases.json` pentru viitoare rulări.

### 3.3 Clasificator motive suspendare
Motive de suspendare provin din **Raport Stări** (câmp "Motiv"). Sunt scrise haotic.

**Flux**:
1. Extrage toate valorile distincte ale câmpului Motiv din Raport Stări
2. Grupează fuzzy (threshold 80% similitudine) → clustere
3. UI afișează clusterele: utilizatorul marchează ce înseamnă "datorii DITL → suspendare eligibilă"
4. Config salvat în `config/motiv_rules.json`:
   ```json
   {
     "datorii_ditl": ["Datorii DITL", "Datorii dgitl", "debite dgitl", "are datorii la ditl"],
     "exclude": ["decedat", "retras", "renuntat"]
   }
   ```
5. La rulări viitoare: motivele necunoscute → pop-up confirmare; cele știute → automat

### 3.4 Logica AdaptiveDebtStatus (Python)
Replică logica din Power Query, îmbunătățită:

```
Pentru fiecare cerere (CNP + Nr. Cerere):
  debt_flag = valoarea din coloana "Are/Nu are debit" a sectorului
  data_repunere = valoarea din coloana "Data repunere" (dacă există)

  IF debt_flag == "Are debit":
    IF data_repunere IS NOT NULL AND data_repunere >= data_interogare:
      status = "repus"  # are datorii dar s-a repus
    ELSE:
      status = "datorii DITL"
  ELIF debt_flag == "Nu are debit":
    status = "fara datorii reinterogare {luna} {an}"
  ELSE:
    status = "neclar"  # coloana goală / valoare nerecunoscută
```

### 3.5 Corelarea cu interogarea termen
**Scop**: Identifică cu certitudine cine e suspendat *din cauza datoriilor de la interogarea termen* (30.03 sau 30.09), nu din altă cauză.

**Flux** (când fișierul Interogare Termen e furnizat):
1. Load Raport Stări → filtrare pe `Stare = Suspendat` + motiv clasificat ca "datorii DITL"
2. Intersectare cu interogarea termen: CNP-uri care *au apărut cu datorii* la 30.03 sau 30.09
3. Rezultat: listă confirmată de beneficiari suspendați pentru datorii la termen, eligibili pentru repunere
4. Coloană suplimentară în output: `Sursa_Suspendare` = data interogării termen care a generat suspendarea (ex: `30.09.2025`)

**Dacă fișierul Interogare Termen nu e disponibil**: se folosesc doar motivele din Raport Stări (mai puțin precis, avertisment în UI).

---

## 4. Structura output Excel

**Fișier**: `Reinterogare_LUNA_AN_output.xlsx` (ex: `Reinterogare_Dec_2025_output.xlsx`)

| Sheet | Conținut |
|-------|---------|
| `S1` ... `S6` | Date brute detectate per sector (inspecție vizuală) |
| `CONSOLIDAT` | Toate cele 6 sectoare unite, coloane standardizate (structura platformei) |
| `REPUNERI` | Beneficiari cu datorii dar cu dată repunere → candidați repunere în platformă |
| `IMPORT_OBSERVATII` | Format gata de import în platformă (observații per cerere) |
| `NECLAR` | Rânduri unde detectarea a eșuat → necesită revizie manuală |
| `LOG` | Coloane detectate, aliases folosite, motive clasificate — trasabilitate |

**Structura CONSOLIDAT** (mapată pe coloanele platformei):
- `CNP`, `Număr Cerere`, `Sector`, `Stare Calculată`, `Data Repunere`, `Motiv Suspendare`, `Sursa Suspendare`, `Observații`, `[coloane originale S1..S6 relevante]`

---

## 5. Configurație persistentă

```
reinterogare_processor/
├── config/
│   ├── column_aliases.json   # denumiri coloane cunoscute per sector
│   ├── motiv_rules.json      # clasificare motive confirmate
│   └── settings.json         # paths ultimele fișiere, output dir, etc.
├── reinterogare_processor.py # app principal
└── modules/
    ├── column_detector.py
    ├── motiv_classifier.py
    ├── debt_logic.py
    ├── output_builder.py
    └── ui_components.py
```

---

## 6. Gestionarea erorilor

- **Coloană critică nedetectată**: UI blochează procesarea + arată ce lipsește, nu produce output parțial fals
- **Sector lipsă**: procesare parțială acceptată, sheet-ul sectorului lipsă rămâne gol cu notă
- **Motiv necunoscut nou**: pop-up interactiv (nu ignorat silențios)
- **Fișier deschis în Excel**: detectare PermissionError → mesaj clar "Închideți fișierul X în Excel"
- **Sheet NECLAR**: orice rând cu date insuficiente pentru decizie → izolat pentru revizie manuală

---

## 7. Prioritate implementare

**Faza 1 (MVP)** — fără Interogare Termen:
- Detector coloane + UI confirmare
- Clasificator motive + config persistent
- Logica AdaptiveDebtStatus
- Output: S1–S6 + CONSOLIDAT + LOG

**Faza 2**:
- Sheet REPUNERI + IMPORT_OBSERVATII (format platformă)
- Corelarea cu Interogare Termen
- Sheet NECLAR

**Faza 3**:
- Drag & drop fișiere
- Auto-detectare lună/an din numele fișierelor
- Comparație cu luna precedentă (cine a apărut nou cu datorii)
