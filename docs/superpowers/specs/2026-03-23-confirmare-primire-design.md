# Design Spec: Confirmare de Primire App

**Data:** 2026-03-23
**Status:** Aprobat de utilizator

---

## Context

Aplicație desktop pentru generarea automată a formularelor "Confirmare de Primire (A.R.)" (CP14 / AR) ale Poștei Române, folosite de SSFAH/DGASMB pentru trimiteri recomandate către beneficiari.

Formularul curent este completat manual. Scopul aplicației este să automatizeze completarea din listă Excel.

**Template sursă:** `C:\EXCELURI\Tipizate\Model completare Confirmare de primire 2024 (1).docx`
- Format: imagini (layout grafic) + 80 text box-uri goale suprapuse
- 3 formulare per pagină
- Structura docx: VML groups cu `v:textbox` elements

---

## Structura Formularului

### Jumătatea stângă — "Se completează de expeditorul trimiterii"

Câmpuri completate automat din Excel (variază per beneficiar):

| Câmp formular | Coloană Excel | Obligatoriu | Note |
|---------------|--------------|-------------|------|
| Nume și prenume destinatar | `Nume Destinatar` | da | |
| Strada | `Strada` | da | |
| Nr. | `Nr` | da | |
| Bl. | `Bl` | nu | gol dacă lipsește |
| Sc. | `Sc` | nu | gol dacă lipsește |
| Et. | `Et` | nu | gol dacă lipsește |
| Ap. | `Ap` | nu | gol dacă lipsește |
| Localitate | `Localitate` | nu | default: `București` |
| Cod postal | `Cod Postal` | nu | |
| Jud/Sector | `Sector` | da | |
| Nr. înregistrare | `Nr Inregistrare` | da | din Registrul Google Sheets |
| Greutate | `Greutate` | nu | gol dacă lipsește |
| Valoare | `Valoare` | nu | gol dacă lipsește |
| Ramburs | `Ramburs` | nu | gol dacă lipsește |

Câmpuri globale (setate o singură dată la generare, nu per rând):
- **Felul trimiterii** = întotdeauna `recomandat` (fix, nu editabil)
- **Data prezentării** = input global în UI, default data curentă

### Jumătatea dreaptă — "A se înapoia la expeditor"

Câmpuri **fixe**, configurate în Tab Configurare și salvate în `config.json`:

| Câmp formular | Configurare |
|---------------|-------------|
| Indicativ serviciu | `SSFAH` — label fix, needitabil |
| Strada expeditor | editabil în ConfigTab |
| Nr. expeditor | editabil în ConfigTab |
| Cod postal expeditor | editabil în ConfigTab |
| Sector expeditor | editabil în ConfigTab |

Câmp variabil (din Excel):
- **Nr. înregistrare** → coloana `Nr Inregistrare` (același câmp ca stânga)

---

## Maparea Text Box-urilor (DocxFiller)

Documentul template conține 3 VML groups (un group = un formular). Fiecare group are ~26-27 `v:textbox` elemente poziționate absolut prin atributul `style="position:absolute;left:X;top:Y;..."`.

### Strategia de mapare

1. La prima rulare, `DocxFiller` extrage toate textbox-urile din primul group VML
2. Le sortează după coordonata `top` (Y), apoi `left` (X) — ordine vizuală de sus în jos, stânga la dreapta
3. Le numerotează și construiește un dict `{index: câmp_semantic}`
4. Maparea este **hard-coded** în cod după ce a fost determinată empiric prin inspecție
5. Același offset de index se aplică pentru group-ul 2 și 3 (formularul 2 și 3 pe pagină)

### Mapare index → câmp (de determinat la implementare)

Se va executa un script de diagnosticare care extrage pozițiile tuturor textbox-urilor și le afișează vizual, apoi se completează maparea. Maparea finală va fi un dict constant în `DocxFiller`.

---

## Arhitectura Aplicației

**Fișier:** `C:\EXCELURI\Tipizate\confirmare_primire_app.py`
**Stack:** Python 3.x, CustomTkinter, openpyxl, python-docx, lxml

### Module (în același fișier)

```
confirmare_primire_app.py
├── Config (dataclass)     — setări persistente, load/save config.json
├── ExcelReader            — citire, validare coloane, returnează List[dict]
├── DocxFiller             — mapare VML textbox-uri, completare template
├── ConfigTab (CTk Frame)  — UI configurare adresă expeditor
└── GenerareTab (CTk Frame)— UI generare: upload, preview, selectare, generare
```

### `config.json` (lângă script)
```json
{
  "expeditor_strada": "",
  "expeditor_nr": "",
  "expeditor_cod_postal": "",
  "expeditor_sector": "",
  "template_path": "C:\\EXCELURI\\Tipizate\\Model completare Confirmare de primire 2024 (1).docx",
  "output_folder": ""
}
```

---

## Fluxul UI

### Tab 1 — Configurare
- Label fix: `SSFAH` (needitabil)
- Entry: Strada, Nr., Cod Postal, Sector — adresa DGASMB
- Entry: cale template (cu Browse)
- Buton **Salvează** → scrie `config.json`

### Tab 2 — Generare
1. **Browse Excel** → validare coloane obligatorii → afișare tabel preview
2. **Data prezentării** — DateEntry sau Entry simplu (format `DD.MM.YYYY`, default azi)
3. **Tabel preview** — CTkScrollableFrame cu rânduri + checkbox per rând + buton "Selectează tot"
4. **Buton Generează** → progresbar → mesaj succes cu calea fișierului
5. Click pe cale → deschide folderul în Explorer

---

## ExcelReader

- Coloane obligatorii: `Nume Destinatar`, `Strada`, `Nr`, `Sector`, `Nr Inregistrare`
- Coloane opționale: `Bl`, `Sc`, `Et`, `Ap`, `Localitate`, `Cod Postal`, `Greutate`, `Valoare`, `Ramburs`
- Coloane lipsă opționale → completate cu `""` automat
- `Localitate` lipsă sau goală → default `"București"`
- La validare eșuată: messagebox cu lista coloanelor lipsă

---

## DocxFiller — Logica de Generare

```
input: List[dict] beneficiari, config, data_prezentarii
output: docx bytes / fișier salvat
```

1. Determină numărul de pagini necesare: `ceil(len(beneficiari) / 3)`
2. Pentru fiecare pagină:
   a. Copiază template-ul în memorie (`copy.deepcopy` pe Document)
   b. Obține cele 3 VML groups
   c. Pentru fiecare group (formular 1, 2, 3):
      - Extrage lista de `v:textbox` sortată după poziție
      - Aplică maparea index → câmp → valoare din beneficiarul curent
      - Setează textul în `txbxContent` → `w:p` → `w:r` → `w:t`
      - Formulare goale (beneficiar lipsă): nu se modifică textbox-urile
3. Concatenează paginile într-un singur document (via `python-docx` section breaks sau `docxcompose`)
4. Salvează fișierul output

---

## Edge Cases

| Situație | Comportament |
|----------|-------------|
| Excel fără rânduri valide | Messagebox eroare, nu se generează |
| Selectare goală (0 rânduri bifate) | Buton Generează dezactivat |
| N % 3 != 0 | Ultimele 1-2 formulare pe ultima pagină rămân goale |
| Fișier output există deja | Dialog "Suprascrie?" înainte de salvare |
| Coloană obligatorie lipsă din Excel | Messagebox cu lista coloanelor lipsă |
| Template lipsă / corupt | Messagebox eroare cu calea configurată |

---

## Output

- **Filename:** `Confirmare_primire_YYYY-MM-DD.docx`
- **Folder:** același cu Excel-ul input (sau cel din config dacă e setat)
- **Conținut:** N/3 pagini (rotunjit în sus), 3 formulare per pagină, layout identic cu template-ul

---

## Excluderi (out of scope)

- Integrare directă cu Google Sheets (nr. înregistrare se copiază manual în Excel)
- Generare PDF sau print direct din app
- Editare vizuală a template-ului din UI
