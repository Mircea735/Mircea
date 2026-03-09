# Verificare Post-Reinterogare — Design Document

**Data**: 2026-03-09
**Status**: Aprobat
**Scop**: Adaugă în Reinterogare Processor un modul de verificare post-procesare care compară rezultatele reinterogării (REPUNERE + IMPORT OBSERVATII) cu un Raport Stare ulterior, pentru a confirma că repunerile și observațiile au fost aplicate corect în platformă.

---

## 1. Contextul problemei

După rularea lunară a reinterogării:
- Sheet-ul **REPUNERE** conține beneficiarii care trebuie repuși (reactivați) în platformă
- Sheet-ul **IMPORT OBSERVATII** conține beneficiarii cu datorii confirmate

Operatorul efectuează repunerile și importul observațiilor manual în platformă. Verificarea post-reinterogare confirmă că aceste acțiuni s-au realizat corect, prin compararea cu un Raport Stare generat ulterior.

---

## 2. UI — Adăugiri

Secțiune nouă în fereastra principală, sub fișierele opționale existente:

```
┌─────────────────────────────────────────────────────┐
│  Verificare post-Reinterogare                       │
│  Raport Stare verificare  [___________________] [...] [X]  │
│  Fișier output reinterogare [________________] [...]       │
│  [ Generează verificare ]                           │
└─────────────────────────────────────────────────────┘
```

- **Raport Stare verificare** — Raportul de stare din luna ulterioară reinterogării (picker nou)
- **Fișier output reinterogare** — pre-completat automat după procesare; selectabil manual pentru rulări anterioare
- **Buton "Generează verificare"** — activ doar când ambele fișiere sunt selectate; rulează în thread background

---

## 3. Logica de verificare

### 3.1 VERIF REPUNERE

Sursa: sheet-ul `REPUNERE` din fișierul output reinterogare.
Join pe: `CNP adult cu handicap` + `Nr. cerere` → Raport Stare verificare.

**Condiții verificate** (3 coloane separate de detaliu):

| Condiție | Coloană detaliu | Valoare așteptată |
|----------|----------------|-------------------|
| Stare | `Stare OK` | `Stare Cerere = Aprobat` |
| Motiv | `Motiv OK` | `Motiv` = `Motiv repunere` din REPUNERE |
| Data | `Data OK` | `Data Inceput Stare` = `Data repunere` din REPUNERE |

**Coloane output VERIF REPUNERE:**

| Coloană | Sursă |
|---------|-------|
| Nr. cerere | REPUNERE |
| CNP adult cu handicap | REPUNERE |
| Nume / Prenume | REPUNERE |
| Motiv repunere (așteptat) | REPUNERE |
| Data repunere (așteptată) | REPUNERE |
| Stare Actuală | Raport Stare verificare |
| Motiv Actual | Raport Stare verificare |
| Data Inceput Stare Actuală | Raport Stare verificare |
| Stare OK | `✓` / `✗` / `NEGASIT` |
| Motiv OK | `✓` / `✗` / `NEGASIT` |
| Data OK | `✓` / `✗` / `NEGASIT` |

**Colorare rânduri:**
- Toate 3 coloane `✓` → fundal verde deschis (`#C6EFCE`)
- Cel puțin o coloană `✗` → fundal roșu deschis (`#FFC7CE`)
- `NEGASIT` → fundal galben (`#FFEB9C`)

### 3.2 VERIF IMPORT OBS

Sursa: sheet-ul `IMPORT OBSERVATII` din fișierul output reinterogare.
Join pe: `CNP adult cu handicap` + `Nr. cerere` → Raport Stare verificare.

**Condiții verificate:**

| Condiție | Coloană detaliu | Valoare așteptată |
|----------|----------------|-------------------|
| Stare | `Stare OK` | `Stare Cerere = Suspendat` (datorii confirm |
| Observatii | `Obs OK` | câmpul Observatii din platformă conține textul importat |

**Coloane output VERIF IMPORT OBS:**

| Coloană | Sursă |
|---------|-------|
| Nr. cerere | IMPORT OBSERVATII |
| CNP adult cu handicap | IMPORT OBSERVATII |
| Nume / Prenume | IMPORT OBSERVATII |
| Observație importată (așteptată) | IMPORT OBSERVATII (col Observatii) |
| Stare Actuală | Raport Stare verificare |
| Observații platformă | Raport Stare verificare |
| Stare OK | `✓` / `✗` / `NEGASIT` |
| Obs OK | `✓` / `✗` / `NEGASIT` / `(fara camp)` |

**Colorare rânduri:** identic cu VERIF REPUNERE.

---

## 4. Structura fișiere

Cele două sheet-uri (`VERIF REPUNERE`, `VERIF IMPORT OBS`) se adaugă în **același fișier Excel** output al reinterogării, fără a modifica sheet-urile existente.

---

## 5. Module noi

```
modules/
  verifier.py          # logica de verificare (PostReinterogareVerifier)
```

Funcții principale în `verifier.py`:
- `verify_repunere(repunere_df, raport_stare_df) -> pd.DataFrame`
- `verify_import_obs(import_obs_df, raport_stare_df) -> pd.DataFrame`

Adăugiri în `output_builder.py`:
- `append_verification(output_path, verif_repunere_df, verif_import_obs_df)` — deschide fișierul existent și adaugă cele 2 sheet-uri

Adăugiri în `reinterogare_processor.py` (UI):
- Picker `Raport Stare verificare`
- Entry pre-completat `Fișier output reinterogare`
- Buton `Generează verificare` → apelează `PostReinterogareVerifier`

---

## 6. Coloane Raport Stare (referință join)

Join pe: `Cnp Adult Handicap` + `Numar Cerere` (sau doar CNP dacă nr. cerere lipsește).

Câmpuri relevante din Raport Stare:
- `Stare Cerere`
- `Motiv`
- `Data Inceput Stare`
- `Observatii` (dacă există — pentru verificarea import obs)
