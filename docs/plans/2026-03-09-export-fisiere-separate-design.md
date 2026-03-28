# Export Fișiere Separate — Design Document

**Data**: 2026-03-09
**Status**: Aprobat

## Scop

La fiecare procesare reinterogare, pe lângă fișierul principal output, se salvează automat fișiere Excel separate în același folder output, pentru import direct în platformă.

---

## Structura output

```
Output dir\
├── Reinterogare_31_01_2026_143022_output.xlsx   ← fișier principal (neschimbat)
├── Repunere.xlsx                                 ← export automat la procesare
├── Import Observatii.xlsx                        ← export automat la procesare
└── Verificare Reinterogare 31.01.2026.xlsx       ← generat la "Genereaza verificare"
```

---

## Fișiere generate automat la procesare

### `Repunere.xlsx`
- Conținut identic cu sheet-ul `REPUNERE` din output-ul principal
- Coloane: `Nr. cerere`, `Data cererii`, `Nume adult cu handicap`, `Prenume adult cu handicap`, `CNP adult cu handicap`, `Motiv repunere`, `Data repunere`

### `Import Observatii.xlsx`
- Conținut identic cu sheet-ul `IMPORT OBSERVATII` din output-ul principal
- Coloane: `Nume adult cu handicap`, `Prenume adult cu handicap`, `CNP adult cu handicap`, `Nr. cerere`, `Observatii`

Ambele fișiere se suprascriu la re-run fără eroare.

---

## Fișier generat la "Genereaza verificare"

### `Verificare Reinterogare {DD.MM.YYYY}.xlsx`
- Conține două sheet-uri: `VERIF REPUNERE` și `VERIF IMPORT OBS`
- Identic ca și conținut cu sheet-urile adăugate în output-ul principal
- Data `{DD.MM.YYYY}` extrasă din denumirea fișierului output selectat (pattern `\d{2}_\d{2}_\d{4}` → reformatat cu `.`)
- Salvat în același `output_dir` ca fișierul output selectat
- Suprascris la re-run

---

## Error handling

- Eșecul scrierii fișierelor separate nu întrerupe procesarea principală — se loghează un avertisment
- PermissionError (fișier deschis în Excel) → mesaj clar în log

---

## Module modificate

- `modules/output_builder.py` — metodă nouă `write_separate_files(output_dir, repunere_df, import_obs_df)`
- `modules/output_builder.py` — `append_verification` extins: scrie și `Verificare Reinterogare {DD.MM.YYYY}.xlsx` în `output_dir`
- `reinterogare_processor.py` — apel `write_separate_files` după `write()` în pipeline
- `reinterogare_processor.py` — `_run_verificare` extins: derivă `output_dir` și data din calea fișierului output, scrie fișierul de verificare separat
