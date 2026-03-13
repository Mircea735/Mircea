# Design Spec: Tab "Omise / Deces manual" — Adeverință Notar

**Data:** 2026-03-13
**Fișier vizat:** `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py`

---

## Problemă

Raportul cu Istoric Plăți nu conține întotdeauna marcajul `"Beneficiar decedat"` în coloana Istoric Plăți. În aceste cazuri, `data_deces_raw` rămâne `None` după `_load_excel_istoric()`, `compute_outstanding()` returnează `None`, iar adeverința nu se generează — CNP-ul este omis silențios.

---

## Soluție

Un nou tab **"Omise"** în `CTkTabview`-ul existent. Tab-ul se populează automat după o rulare de generare cu lista CNP-urilor omise, permite introducerea manuală a datei decesului în format `MM.YYYY` și declanșează generarea individuală per CNP completat.

---

## Confirmare parametru existent

`generate_all()` acceptă deja parametrul `filter_cnp` (linia ~724 din cod) — nicio modificare la engine nu este necesară.

---

## Arhitectură

### 1. Colectarea CNP-urilor omise

**Când:** imediat după finalizarea `generate_all()` din thread-ul de generare din tab-ul Principal, în același callback `callback(100, ...)`.

**Cum:** se parcurge `engine.raw_data` și se colectează CNP-urile unde `data_deces_raw is None`. Dacă `engine.raw_data` e gol, lista e vidă.

```python
self._cnps_omise = [
    cnp for cnp, p in self.engine.raw_data.items()
    if p['data_deces_raw'] is None
]
```

**Re-run:** la fiecare nouă generare din Principal, lista se recalculează de la zero, tab-ul Omise se repopulează complet (câmpurile introduse anterior se pierd).

### 2. Tab UI — componente

Tab-ul se numește fix `"Omise"` — CTkTabview nu suportă redenumire dinamică după creare. Numărul CNP-urilor omise se afișează într-un **CTkLabel** de tip header în interiorul tab-ului:

```
● 3 CNP-uri fără deces detectat automat     [actualizat după fiecare generare]
```

Când lista e vidă sau nu s-a rulat niciodată o generare: se afișează mesajul `"Tab-ul se populează după o rulare din tab-ul Principal."` în locul tabelului.

**Componente în ordine:**
1. `CTkLabel` header cu numărul CNP-urilor (sau mesaj stare inițială)
2. `CTkLabel` notă informativă: `"Aceste CNP-uri nu au marcajul «Beneficiar decedat» în Istoric Plăți. Introduceți luna și anul decesului."`
3. `CTkScrollableFrame` cu rânduri per CNP (detalii mai jos)
4. `CTkLabel` mesaj validare (inițial ascuns)
5. `CTkButton` `"Generează pentru completați"` — inițial `state="disabled"`
6. `CTkTextbox` mini-log (4 linii, read-only) — afișează rezultate generare

### 3. Rânduri în ScrollableFrame

Structura `self._omise_entries: dict[str, CTkEntry]` se construiește în metoda `_populate_omise_tab()`, apelată din thread-ul principal (via `self.after(0, self._populate_omise_tab)`) după actualizarea `self._cnps_omise`.

Per CNP, un rând orizontal:
- `CTkLabel` CNP monospace (font `("Courier", 12)`)
- `CTkLabel` `f"{p['prenume']} {p['nume']}"` unde `p = engine.raw_data[cnp]`
- `CTkEntry` placeholder `"MM.YYYY"`, width=100, cu binding `<KeyRelease>` → `_on_omise_entry_change()`

La fiecare keystroke, `_on_omise_entry_change()` verifică dacă cel puțin un câmp are text nevid → activează/dezactivează butonul. Nu face validare regex (validarea se face doar la submit).

### 4. Validare dată (la submit)

La apăsarea butonului "Generează pentru completați":
- Per câmp nevid: `re.match(r'^(\d{1,2})\.(\d{4})$', val.strip())`
- Câmp gol → ignorat
- Orice câmp nevid cu format invalid → `border_color="red"` pe acel entry + mesaj vizibil `"Format invalid — folosiți MM.YYYY (ex: 06.2024)"` + generarea NU pornește
- Toate câmpurile nevide trec validarea → generare pornește, mesaj de eroare ascuns

### 5. Generare

Se rulează în **thread separat** (ca generarea din Principal). Pe durata rulării:
- Butonul "Generează pentru completați" → `state="disabled"`
- Butonul "Generează" din tab-ul Principal → `state="disabled"` (aceeași metodă `_set_busy()` existentă)
- La finalizare (succes sau eroare) → ambele butoane reactivate

**Guard pre-generare:** dacă `template_path` sau `out_dir` nu sunt setate (utilizatorul n-a rulat niciodată generarea principală), se afișează mesaj de eroare în mini-log: `"Configurați mai întâi generarea din tab-ul Principal (template + folder output)."` și generarea nu pornește.

```python
def _run_omise_generation(self):
    if not self.template_path or not self.out_dir:
        self._omise_minilog("✘ Configurați mai întâi generarea din tab-ul Principal.")
        return

    for cnp, entry_widget in self._omise_entries.items():
        val = entry_widget.get().strip()
        if not val:
            continue
        m = re.match(r'^(\d{1,2})\.(\d{4})$', val)
        if not m:
            continue  # validate step already ran; skip invalid
        mm, yyyy = m.group(1).zfill(2), m.group(2)
        self.engine.raw_data[cnp]['data_deces_raw'] = f"01.{mm}.{yyyy}"
        n_gen, n_fara, erori = self.engine.generate_all(
            self.template_path, self.out_dir,
            indicativ=self.entry_indicativ.get(),
            nr_inreg=self.entry_nr_inreg.get(),
            cod_mt=self.entry_cod_mt.get(),
            filter_cnp=cnp,
            callback=lambda pct, msg: self._omise_minilog(msg)
        )
        if erori:
            self._omise_minilog(f"✘ {cnp}: {erori[0]}")
        elif n_gen == 0:
            self._omise_minilog(f"⚠ {cnp}: fără restanțe — adeverință negenerată")
        else:
            self._omise_minilog(f"✔ {cnp}: adeverință generată")
            self.after(0, lambda w=entry_widget: w.configure(state="disabled", text_color="gray"))
```

**`_omise_minilog(msg)`**: apelează `self.after(0, lambda: ...)` pentru a scrie în CTkTextbox din thread-ul secundar (thread-safe).

**`template_path`** și **`out_dir`** sunt variabile de instanță deja existente în UI, setate când utilizatorul selectează fișierele în tab-ul Principal.

### 6. Comportament după generare parțială

- CNP generat cu succes → entry `state="disabled"`, text gri — nu mai poate fi reeditat în sesiunea curentă
- CNP fără restanțe sau cu eroare → entry rămâne activ, utilizatorul poate corecta și reîncerca
- La o nouă generare din Principal → tab-ul Omise se resetează complet (inclusiv entrys dezactivate)

---

## Format dată acceptat

`MM.YYYY` — ex: `06.2024`
Intern: `f"01.{mm}.{yyyy}"` → compatibil cu `parse_data_deces()` și `format_data_deces()`.

---

## Ce NU se modifică

- `AdeverintaEngine` și toate metodele sale
- `generate_all()`, `compute_outstanding()`, `_load_excel_*`
- Formatul fișierelor output
- Tab-ul Principal — nicio modificare vizuală sau funcțională

---

## Fișiere afectate

| Fișier | Modificare |
|--------|-----------|
| `adeverinta_notar_app.py` | Adăugare tab Omise + logică UI (nicio modificare la engine) |

**Metode noi adăugate în clasa UI:**
- `_populate_omise_tab()` — reconstruiește conținutul tab-ului Omise
- `_on_omise_entry_change()` — activează/dezactivează butonul la keystroke
- `_run_omise_generation()` — rulează în thread, generează per CNP
- `_omise_minilog(msg)` — scrie în mini-log thread-safe

---

## Criterii de acceptare

1. După generare din Principal, tab-ul "Omise" afișează header cu numărul și lista CNP-urilor cu `data_deces_raw = None`
2. La o nouă generare din Principal, tab-ul Omise se resetează complet cu lista actualizată
3. Butonul "Generează pentru completați" este dezactivat inițial; se activează când cel puțin un câmp CTkEntry conține text nevid
4. La submit cu cel puțin un câmp nevid cu format invalid: bordură roșie + mesaj eroare, generarea nu pornește
5. Câmpuri goale sunt ignorate; nu blochează generarea câmpurilor completate valid
6. Generarea produce același output Word ca fluxul normal pentru CNP-ul respectiv (via `generate_all(filter_cnp=cnp)`)
7. CNP generat cu succes → entry dezactivat vizual (gri)
8. CNP fără restanțe → mesaj `"⚠ fără restanțe"` în mini-log, entry rămâne editabil
9. Pe durata generării din Omise, butonul "Generează" din Principal este dezactivat (și invers)
10. Dacă `template_path`/`out_dir` nu sunt setate: mesaj de eroare în mini-log, generarea nu pornește
11. Tab-ul Principal nu este modificat vizual sau funcțional
12. Stare inițială (nicio generare anterioară): tab-ul Omise afișează mesajul `"Tab-ul se populează după o rulare din tab-ul Principal."`
