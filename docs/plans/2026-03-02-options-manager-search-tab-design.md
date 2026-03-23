# Design: Tab "Caută Beneficiar" în Options Manager Sidebar

**Data**: 2026-03-02
**Proiect**: Registru DGASMB — Google Sheets Options Manager
**Spreadsheet**: `1xFTA80PFIlByy29dBTNfcgeJn9FOXmeEOLCi18Gy6hs`
**Script ID**: `18BkQLG-iV98F69WgAJ-KQXlGAXTvoL3Qc9EmzoQy3FPyZSnd3ndgaPKq`

---

## Scop

Adăugarea unui tab "Caută" în sidebar-ul existent (SidebarV2.html) care permite căutarea documentelor depuse de un beneficiar după CNP sau adresă de email, cu vizualizare inline și export pe un sheet dedicat.

---

## UI Layout

Tab bar adăugat în partea de sus a SidebarV2:
```
[ Opțiuni ]  [ Caută ]
```

Conținut tab "Caută":
```
┌─────────────────────────────────┐
│  🔍 CNP sau Email               │  ← input unic
│  ─────────────────────────────  │
│  Caută în:                      │
│  ☑ Registru E mail 43 2026      │  ← checkboxes încărcate dinamic
│  ☑ Registru ghiseu 41 2026      │    (toate sheet-urile cu "Registru")
│  ☐ Registru Email 96 2025       │
│  [ Selectează toate ]           │
│  ─────────────────────────────  │
│  [ 🔍 Caută ]                   │
│  ─────────────────────────────  │
│  3 documente găsite             │  ← status
│  ┌──────────┬──────────┬──────┐ │
│  │ Data     │ Conținut │ Reg. │ │  ← tabel rezultate
│  │ 5.1.2026 │ Adeverință│ E43 │ │
│  │ 12.2.26  │ C Deces  │ E43 │ │
│  └──────────┴──────────┴──────┘ │
│  [ 📋 Export în sheet ]         │  ← vizibil doar cu rezultate
└─────────────────────────────────┘
```

---

## Auto-detectare tip query

| Input | Logică | Coloană căutată |
|-------|--------|-----------------|
| 13 cifre | CNP | col M (idx 12) |
| conține `@` | Email | col Q (idx 16) |
| altceva | Ambele | col M + col Q |

---

## Funcții noi în Code.gs

### `getRegistrySheets()`
- Returnează array cu numele tuturor sheet-urilor care conțin "Registru" în titlu
- Folosit la inițializarea tab-ului "Caută" pentru a popula checkboxes

### `searchBeneficiar(query, sheetNames)`
- Parametri: `query` (string), `sheetNames` (array de string-uri)
- Citește `getRange(2, 1, lastRow-1, 17)` din fiecare sheet selectat (17 col → include Email col Q)
- Returnează array de obiecte: `[{ data, continut, registru }]`
  - `data` = col E (idx 4)
  - `continut` = col L (idx 11)
  - `registru` = numele sheet-ului

### `exportSumarBeneficiar(query, results)`
- Creează sheet "Sumar Beneficiar" dacă nu există
- Append (nu overwrite) — fiecare export adaugă un bloc nou
- Format bloc:

```
CNP/Email căutat: <query>   | Generat: ZZ.LL.AAAA
Data          | Conținut    | Registru
5.1.2026      | Adeverință  | Registru E mail 43 2026
...
(rând gol separator)
```

---

## Structura fișiere modificate

| Fișier | Modificare |
|--------|------------|
| `SidebarV2.html` | Adăugare tab bar + secțiunea tab "Caută" cu JS |
| `Code.gs` | 3 funcții noi: `getRegistrySheets`, `searchBeneficiar`, `exportSumarBeneficiar` |

Deploy prin `update_options_manager_v2.py` (extins) sau direct via Apps Script editor.

---

## Coloane relevante din registru (referință)

| Col | Idx | Conținut |
|-----|-----|----------|
| E | 4 | Data înregistrării (formulă concatenare din B+C+D) |
| L | 11 | Conținut pe scurt (91 opțiuni dropdown) |
| M | 12 | CNP Adult cu Handicap |
| N | 13 | Sector (1-6) |
| Q | 16 | Email expeditor (trigger auto-reply) |

`getRange(2, 1, lastRow-1, 17)` — 17 coloane necesare pentru a include col Q.

---

## Decizii de design

- **Tab în SidebarV2 existent** (nu sidebar nou) — UX fluent, fără fricțiune
- **Auto-detectare CNP/email** — un singur input în loc de două câmpuri separate
- **Multi-select sheet-uri** — utilizatorul alege ce registre caută
- **Export append** — istoricul exporturilor se păstrează, nu se suprascrie
- **Sheet "Sumar Beneficiar" creat automat** — fără intervenție manuală
