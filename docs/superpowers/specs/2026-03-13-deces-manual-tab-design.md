# Design Spec: Tab "Omise / Deces manual" — Adeverință Notar

**Data:** 2026-03-13
**Fișier vizat:** `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py`

---

## Problemă

Raportul cu Istoric Plăți nu conține întotdeauna marcajul `"Beneficiar decedat"` în coloana Istoric Plăți. În aceste cazuri, `data_deces_raw` rămâne `None` după `_load_excel_istoric()`, `compute_outstanding()` returnează `None`, iar adeverința nu se generează — CNP-ul este omis silențios.

---

## Soluție

Un nou tab **"Omise / Deces manual"** în `CTkTabview`-ul existent. Tab-ul se populează automat după o rulare de generare cu lista CNP-urilor omise, permite introducerea manuală a datei decesului în format `MM.YYYY` și declanșează generarea individuală per CNP completat.

---

## Arhitectură

### 1. Colectarea CNP-urilor omise

În metoda de generare din UI (după `generate_all()`), se parcurge `engine.raw_data` și se colectează CNP-urile unde `data_deces_raw is None`. Acestea se stochează într-o variabilă de instanță a clasei UI (`self._cnps_omise: list[str]`).

### 2. Tab UI — componente

- **Label:** `"X CNP-uri fără deces detectat automat"`
- **Notă informativă** (CTkLabel stilizat, albastru): explică că marcajul lipsește din Istoric Plăți
- **CTkScrollableFrame** cu rânduri per CNP:
  - CNP (label monospace)
  - Nume Prenume (label)
  - CTkEntry placeholder `"MM.YYYY"`
- **Buton** `"Generează pentru completați"` — activ doar dacă ≥1 câmp valid completat

### 3. Badge pe tab

`CTkTabview` nu suportă nativ badge-uri. Se simulează cu un `CTkLabel` suprapus peste butonul de tab, actualizat după fiecare generare.
Alternativ (mai simplu): textul tab-ului include numărul: `f"Omise ({n})"` — actualizat dinamic.

### 4. Validare dată

La apăsarea butonului "Generează":
- Regex: `re.match(r'^(\d{1,2})\.(\d{4})$', val.strip())`
- Invalid → entry bordură roșie + mesaj eroare sub tabel
- Valid → se construiește `data_deces_raw = f"01.{mm}.{yyyy}"` (ziua 1 — suficient pentru `parse_data_deces` care extrage doar luna/anul)

### 5. Generare

Per CNP cu dată validă:
```python
engine.raw_data[cnp]['data_deces_raw'] = f"01.{mm}.{yyyy}"
engine.generate_all(
    template_path, out_dir,
    indicativ=..., nr_inreg=..., cod_mt=...,
    filter_cnp=cnp,
    callback=...
)
```

Rezultatele se afișează în log-ul existent (tab Principal sau un mini-log în tab Omise).

### 6. Stare inițială tab

Dacă nu s-a rulat nicio generare: mesaj `"Tab-ul se populează după o rulare din tab-ul Principal."`. Tab-ul este vizibil de la start (nu apare dinamic) pentru a nu complica `CTkTabview`.

---

## Format dată acceptat

`MM.YYYY` — ex: `06.2024`
Intern se stochează ca `"01.MM.YYYY"` pentru compatibilitate cu `parse_data_deces()` și `format_data_deces()` existente.

---

## Ce NU se modifică

- `AdeverintaEngine` — nicio modificare la engine
- `generate_all()`, `compute_outstanding()` — nicio modificare
- Formatul fișierelor output — nicio modificare
- Tab-ul Principal — nicio modificare funcțională

---

## Fișiere afectate

| Fișier | Modificare |
|--------|-----------|
| `adeverinta_notar_app.py` | Adăugare tab + logică UI |

---

## Criterii de acceptare

1. După generare, tab-ul "Omise" listează toate CNP-urile cu `data_deces_raw = None`
2. Introducerea `MM.YYYY` valid → butonul devine activ
3. Format invalid → bordură roșie + mesaj, generarea blocată
4. Generarea produce același output ca fluxul normal pentru CNP-ul respectiv
5. Tab-ul Principal nu este modificat vizual sau funcțional
