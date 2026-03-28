# Confirmare de Primire App — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Aplicație CustomTkinter care citește o listă Excel cu beneficiari și generează automat formulare Word "Confirmare de Primire (A.R.)" completate, păstrând layout-ul original al template-ului.

**Architecture:** Aplicație single-file cu 5 module logice (Config, ExcelReader, DocxFiller, ConfigTab, GenerareTab). DocxFiller manipulează direct XML-ul VML din template pentru a completa text box-urile poziționate absolut. Output-ul este un `.docx` cu N/3 pagini.

**Tech Stack:** Python 3.x, CustomTkinter, openpyxl, python-docx, lxml

> **NOTE:** `docxcompose` NU se folosește — corupe VML la append. Concatenarea se face prin copiere directă a elementelor `<w:body>` cu `<w:sectPr>` separator. Template-ul se redeschide de pe disk per pagină (NU `deepcopy` pe Document).

---

## File Structure

```
C:\EXCELURI\Tipizate\
├── confirmare_primire_app.py     # aplicația completă (toate modulele)
├── config.json                   # creat automat la prima rulare
├── tests\
│   ├── test_excel_reader.py      # teste ExcelReader
│   ├── test_docx_filler.py       # teste DocxFiller
│   └── sample_lista.xlsx         # fișier Excel de test
```

**Template (existent, nu se modifică):**
`C:\EXCELURI\Tipizate\Model completare Confirmare de primire 2024 (1).docx`

---

## Task 1: Diagnostic — Maparea Text Box-urilor VML

**Scop:** Determina empiric care text box index corespunde cărui câmp din formular.

**Files:**
- Create: `C:\EXCELURI\Tipizate\diagnostic_textboxes.py`

- [ ] **Step 1: Creează scriptul de diagnosticare**

```python
# diagnostic_textboxes.py
import zipfile, re
from lxml import etree

TEMPLATE = r"C:\EXCELURI\Tipizate\Model completare Confirmare de primire 2024 (1).docx"

def extract_textboxes():
    with zipfile.ZipFile(TEMPLATE) as z:
        xml = z.read('word/document.xml')
    root = etree.fromstring(xml)
    ns = {
        'v': 'urn:schemas-microsoft-com:vml',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    }
    # Găsește TOATE v:group-urile, inclusiv imbricate
    all_groups = root.findall('.//v:group', ns)
    print(f"Total VML groups (toate nivelele): {len(all_groups)}")

    # Identifică grupurile care conțin direct textbox-uri (= formularele)
    form_groups = []
    for gi, group in enumerate(all_groups):
        direct_textboxes = group.findall('.//v:textbox', ns)
        child_groups = group.findall('v:group', ns)  # sub-grupuri directe
        print(f"Group {gi}: {len(direct_textboxes)} textboxes, {len(child_groups)} sub-grupuri directe")
        if len(direct_textboxes) > 5:  # grup cu câmpuri reale
            form_groups.append((gi, group))

    print(f"\nGrupuri formular detectate: {[g[0] for g in form_groups]}")
    for gi, group in form_groups:
        # Extrage toate shape-urile cu textbox, indiferent de tip
        shapes_with_tb = []
        for child in group.iter():
            if child.tag == f'{{{VML_NS}}}textbox':
                parent = child.getparent()
                style = parent.get('style', '')
                top  = re.search(r'top:([\d.]+)', style)
                left = re.search(r'left:([\d.]+)', style)
                t = top.group(1) if top else '?'
                l = left.group(1) if left else '?'
                shapes_with_tb.append((float(t) if t != '?' else 0,
                                       float(l) if l != '?' else 0,
                                       style[:60]))
        shapes_with_tb.sort()
        print(f"\n=== GROUP {gi} ({len(shapes_with_tb)} textboxes) ===")
        for ti, (t, l, s) in enumerate(shapes_with_tb):
            print(f"  [{ti:02d}] top={t:>8.1f} left={l:>8.1f}  style={s}")

if __name__ == '__main__':
    extract_textboxes()
```

- [ ] **Step 2: Rulează scriptul**

```bash
cd "C:\EXCELURI\Tipizate"
python diagnostic_textboxes.py
```

Expected: identifică grupurile formular (cele cu >5 textboxes) și listează pozițiile lor sortate.

**IMPORTANT:** Verifică dacă template-ul are 3 grupuri la același nivel sau un grup outer cu 3 inner. Dacă are structură imbricată, `form_groups` va arăta asta — ajustează `DocxFiller._get_form_groups()` în Task 3 corespunzător.

- [ ] **Step 3: Completează maparea**

Sortează textbox-urile după `top` (Y) → identifică vizual care index corespunde cărui câmp. Construiește dict-ul de mapare (va fi folosit în Task 3):

```python
# Completat după inspecție vizuală:
FIELD_MAP = {
    0: 'data_prezentarii',
    1: 'felul_trimiterii',
    2: 'greutate',
    3: 'valoare',
    4: 'ramburs',
    5: 'nr_inregistrare',        # dreapta sus
    6: 'nume_destinatar',
    7: 'strada_dest',
    8: 'nr_dest',
    9: 'bl_dest',
    10: 'sc_dest',
    11: 'et_dest',
    12: 'ap_dest',
    13: 'localitate_dest',
    14: 'cod_postal_dest',
    15: 'sector_dest',
    # expeditor (dreapta) - câmpuri fixe
    16: 'expeditor_strada',
    17: 'expeditor_nr',
    18: 'expeditor_cod_postal',
    19: 'expeditor_sector',
    # restul: goale sau ștampile
}
```

**NOTĂ:** Indexurile exacte se stabilesc după rularea diagnosticului — ajustează dict-ul înainte de Task 3.

- [ ] **Step 4: Commit**

```bash
git add "C:\EXCELURI\Tipizate\diagnostic_textboxes.py"
git commit -m "feat: add VML textbox diagnostic script"
git push
```

---

## Task 2: Config + ExcelReader (cu teste)

**Files:**
- Create: `C:\EXCELURI\Tipizate\confirmare_primire_app.py` (schelet + Config + ExcelReader)
- Create: `C:\EXCELURI\Tipizate\tests\test_excel_reader.py`
- Create: `C:\EXCELURI\Tipizate\tests\sample_lista.xlsx`

- [ ] **Step 1: Creează fișierul Excel de test**

```python
# rulează o dată pentru a crea sample_lista.xlsx
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Nume Destinatar','Strada','Nr','Bl','Sc','Et','Ap',
           'Localitate','Cod Postal','Sector','Greutate','Valoare',
           'Ramburs','Nr Inregistrare'])
ws.append(['Ionescu Maria','Mihai Eminescu','12','A','1','2','5',
           'Bucuresti','010100','1','','','','37342'])
ws.append(['Popescu Ioan','Unirii','5','','','','8',
           '','','2','','','','37343'])
wb.save(r'C:\EXCELURI\Tipizate\tests\sample_lista.xlsx')
print("saved")
```

- [ ] **Step 2: Scrie testele pentru ExcelReader**

```python
# tests/test_excel_reader.py
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from confirmare_primire_app import ExcelReader

SAMPLE = os.path.join(os.path.dirname(__file__), 'sample_lista.xlsx')

def test_reads_rows():
    rows = ExcelReader.read(SAMPLE)
    assert len(rows) == 2

def test_mandatory_fields_present():
    rows = ExcelReader.read(SAMPLE)
    r = rows[0]
    assert r['Nume Destinatar'] == 'Ionescu Maria'
    assert r['Strada'] == 'Mihai Eminescu'
    assert r['Nr Inregistrare'] == '37342'

def test_optional_fields_default_empty():
    rows = ExcelReader.read(SAMPLE)
    r = rows[1]
    assert r['Bl'] == ''
    assert r['Sc'] == ''

def test_localitate_defaults_to_bucuresti():
    rows = ExcelReader.read(SAMPLE)
    r = rows[1]  # Localitate goală în sample
    assert r['Localitate'] == 'Bucuresti'

def test_missing_mandatory_column_raises():
    import pytest, openpyxl, tempfile, os
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Nume Destinatar', 'Strada'])  # lipsesc Nr, Sector, Nr Inregistrare
    ws.append(['Test', 'Str'])
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    wb.save(path)
    with pytest.raises(ValueError, match="Coloane lipsă"):
        ExcelReader.read(path)
    os.unlink(path)
```

- [ ] **Step 3: Rulează testele — trebuie să EȘUEZE (confirmare că testele funcționează)**

```bash
cd "C:\EXCELURI\Tipizate"
python -m pytest tests/test_excel_reader.py -v
```

Expected: ImportError sau AttributeError — `ExcelReader` nu există încă.

- [ ] **Step 4: Implementează Config + ExcelReader în `confirmare_primire_app.py`**

```python
# confirmare_primire_app.py
import json, os
from dataclasses import dataclass, asdict
from typing import List, Dict
import openpyxl

CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.json')

MANDATORY_COLS = {'Nume Destinatar', 'Strada', 'Nr', 'Sector', 'Nr Inregistrare'}
OPTIONAL_COLS  = {'Bl', 'Sc', 'Et', 'Ap', 'Localitate', 'Cod Postal',
                  'Greutate', 'Valoare', 'Ramburs'}
ALL_COLS = MANDATORY_COLS | OPTIONAL_COLS

@dataclass
class Config:
    expeditor_strada: str = ''
    expeditor_nr: str = ''
    expeditor_cod_postal: str = ''
    expeditor_sector: str = ''
    template_path: str = r'C:\EXCELURI\Tipizate\Model completare Confirmare de primire 2024 (1).docx'
    output_folder: str = ''

    @staticmethod
    def load() -> 'Config':
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, encoding='utf-8') as f:
                return Config(**json.load(f))
        return Config()

    def save(self):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(asdict(self), f, ensure_ascii=False, indent=2)


class ExcelReader:
    @staticmethod
    def read(path: str) -> List[Dict[str, str]]:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(max_row=1))]
        missing = MANDATORY_COLS - set(headers)
        if missing:
            raise ValueError(f"Coloane lipsă: {', '.join(sorted(missing))}")
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            d = {h: (str(v).strip() if v is not None else '') for h, v in zip(headers, row)}
            for col in OPTIONAL_COLS:
                d.setdefault(col, '')
            if not d.get('Localitate'):
                d['Localitate'] = 'Bucuresti'
            rows.append(d)
        return rows
```

- [ ] **Step 5: Rulează testele — trebuie să TREACĂ**

```bash
python -m pytest tests/test_excel_reader.py -v
```

Expected: toate 5 teste PASS.

- [ ] **Step 6: Commit**

```bash
git add confirmare_primire_app.py tests/
git commit -m "feat: Config + ExcelReader cu teste"
git push
```

---

## Task 3: DocxFiller

**Files:**
- Modify: `C:\EXCELURI\Tipizate\confirmare_primire_app.py` (adaugă clasa DocxFiller)
- Create: `C:\EXCELURI\Tipizate\tests\test_docx_filler.py`

**PREREQUISIT:** Task 1 complet — `FIELD_MAP` definit cu indexurile corecte.

- [ ] **Step 1: Scrie testele pentru DocxFiller**

```python
# tests/test_docx_filler.py
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from confirmare_primire_app import DocxFiller, Config, ExcelReader
from docx import Document
import tempfile

TEMPLATE = r"C:\EXCELURI\Tipizate\Model completare Confirmare de primire 2024 (1).docx"
SAMPLE   = os.path.join(os.path.dirname(__file__), 'sample_lista.xlsx')

def test_output_is_valid_docx():
    cfg = Config()
    rows = ExcelReader.read(SAMPLE)
    out = tempfile.mktemp(suffix='.docx')
    DocxFiller.generate(rows, cfg, '23.03.2026', out)
    doc = Document(out)
    assert doc is not None
    os.unlink(out)

def test_output_file_created():
    cfg = Config()
    rows = ExcelReader.read(SAMPLE)
    out = tempfile.mktemp(suffix='.docx')
    DocxFiller.generate(rows, cfg, '23.03.2026', out)
    assert os.path.exists(out)
    os.unlink(out)

def test_first_beneficiar_name_in_output():
    """Verifică că numele primului beneficiar apare în XML-ul output."""
    cfg = Config()
    rows = ExcelReader.read(SAMPLE)
    out = tempfile.mktemp(suffix='.docx')
    DocxFiller.generate(rows, cfg, '23.03.2026', out)
    # Verifică structural că textul a fost scris în docx
    import zipfile
    with zipfile.ZipFile(out) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    assert 'Ionescu Maria' in xml, "Numele primului beneficiar nu apare în document.xml"
    assert '37342' in xml, "Nr. înregistrare nu apare în document.xml"
    os.unlink(out)
```

- [ ] **Step 2: Rulează testele — trebuie să EȘUEZE**

```bash
python -m pytest tests/test_docx_filler.py -v
```

Expected: ImportError — `DocxFiller` nu există.

- [ ] **Step 3: Implementează DocxFiller**

```python
# adaugă în confirmare_primire_app.py, după ExcelReader

import copy, re, math
from lxml import etree
from docx import Document

# =====================================================================
# FIELD_MAP: completat după rularea diagnostic_textboxes.py (Task 1)
# Cheia = index textbox în VML group (sortat după top, apoi left)
# Valoarea = cheia din dict-ul beneficiarului (sau cheie specială)
# =====================================================================
FIELD_MAP = {
    # Ajustează după diagnostic:
    0:  'data_prezentarii',
    1:  'felul_trimiterii',   # fix: 'recomandat'
    2:  'greutate',
    3:  'valoare',
    4:  'ramburs',
    5:  'nr_inregistrare_display',  # ex: 'Nr. 37342'
    6:  'Nume Destinatar',
    7:  'Strada',
    8:  'Nr',
    9:  'Bl',
    10: 'Sc',
    11: 'Et',
    12: 'Ap',
    13: 'Localitate',
    14: 'Cod Postal',
    15: 'Sector',
    16: 'expeditor_strada',
    17: 'expeditor_nr',
    18: 'expeditor_cod_postal',
    19: 'expeditor_sector',
}

VML_NS  = 'urn:schemas-microsoft-com:vml'
W_NS    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

class DocxFiller:

    @staticmethod
    def _get_form_groups(doc_element):
        """
        Returnează exact 3 VML groups corespunzătoare celor 3 formulare.
        Gestionează atât structura flat (3 grupuri la același nivel)
        cât și cea imbricată (1 grup outer + 3 inner).
        Rezultatul diagnosticului din Task 1 determină care ramură se folosește.
        """
        all_groups = doc_element.findall(f'.//{{{VML_NS}}}group')
        # Filtrare: grupuri cu >5 textboxes directe = grupuri formular
        form_groups = [
            g for g in all_groups
            if len(list(g.iter(f'{{{VML_NS}}}textbox'))) > 5
        ]
        if len(form_groups) < 3:
            raise RuntimeError(
                f"Găsite doar {len(form_groups)} grupuri formular (expected 3). "
                "Verifică structura VML cu diagnostic_textboxes.py."
            )
        return form_groups[:3]  # primele 3 grupuri formular

    @staticmethod
    def _get_textboxes_sorted(group_el):
        """
        Returnează lista de elemente v:textbox sortate după top, left.
        Iterează prin TOȚI descendenții grupului (nu doar copiii direcți)
        pentru a acoperi v:shape, v:rect, v:oval și alte tipuri.
        """
        result = []
        for child in group_el.iter():
            if child.tag != f'{{{VML_NS}}}textbox':
                continue
            parent = child.getparent()
            style = parent.get('style', '')
            m_top  = re.search(r'top:([\d.]+)', style)
            m_left = re.search(r'left:([\d.]+)', style)
            top  = float(m_top.group(1))  if m_top  else 0.0
            left = float(m_left.group(1)) if m_left else 0.0
            result.append((top, left, child))
        result.sort(key=lambda x: (x[0], x[1]))
        return [tb for _, _, tb in result]

    @staticmethod
    def _set_textbox_text(tb_el, text: str):
        """
        Inserează text în textbox păstrând formatarea originală (font, size, bold).
        Copiază rPr (run properties) din primul run existent înainte de a șterge conținutul.
        """
        content = tb_el.find(f'{{{W_NS}}}txbxContent')
        if content is None:
            return

        # Salvează rPr și pPr din primul run/paragraf existent
        existing_rpr = None
        existing_ppr = None
        first_p = content.find(f'{{{W_NS}}}p')
        if first_p is not None:
            existing_ppr = first_p.find(f'{{{W_NS}}}pPr')
            first_r = first_p.find(f'{{{W_NS}}}r')
            if first_r is not None:
                existing_rpr = first_r.find(f'{{{W_NS}}}rPr')

        # Șterge conținut existent
        for p in content.findall(f'{{{W_NS}}}p'):
            content.remove(p)

        # Reconstruiește cu formatare păstrată
        p = etree.SubElement(content, f'{{{W_NS}}}p')
        if existing_ppr is not None:
            p.insert(0, existing_ppr)
        r = etree.SubElement(p, f'{{{W_NS}}}r')
        if existing_rpr is not None:
            r.insert(0, existing_rpr)
        t = etree.SubElement(r, f'{{{W_NS}}}t')
        t.text = text
        if text.startswith(' ') or text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    @staticmethod
    def _fill_group(group_el, row: dict, config: Config, data: str, form_idx: int):
        """Completează un VML group (un formular) cu datele unui beneficiar."""
        textboxes = DocxFiller._get_textboxes_sorted(group_el)
        special = {
            'data_prezentarii': data,
            'felul_trimiterii': 'recomandat',
            'nr_inregistrare_display': f"Nr. {row.get('Nr Inregistrare', '')}",
            'expeditor_strada': config.expeditor_strada,
            'expeditor_nr': config.expeditor_nr,
            'expeditor_cod_postal': config.expeditor_cod_postal,
            'expeditor_sector': config.expeditor_sector,
            'greutate': row.get('Greutate', ''),
            'valoare': row.get('Valoare', ''),
            'ramburs': row.get('Ramburs', ''),
        }
        for idx, tb in enumerate(textboxes):
            field = FIELD_MAP.get(idx)
            if field is None:
                continue
            value = special.get(field) or row.get(field, '')
            DocxFiller._set_textbox_text(tb, str(value))

    @staticmethod
    def _fill_page(config: Config, page_rows: list, data: str) -> Document:
        """
        Deschide template-ul de pe disk (NU deepcopy), completează
        cele 3 formulare cu datele din page_rows (1-3 beneficiari).
        Returnează Document-ul modificat.
        """
        doc = Document(config.template_path)  # redeschide de pe disk per pagină
        form_groups = DocxFiller._get_form_groups(doc.element)
        for form_pos, group in enumerate(form_groups):
            if form_pos >= len(page_rows):
                break  # formular gol — lasă necompletat
            DocxFiller._fill_group(group, page_rows[form_pos], config, data, form_pos)
        return doc

    @staticmethod
    def generate(rows: list, config: Config, data: str, output_path: str):
        """
        Generează docx complet. Concatenare prin copiere body XML (nu docxcompose
        care corupe VML). Fiecare pagină = un Document separat deschis de pe disk.
        """
        pages_needed = math.ceil(len(rows) / 3)

        # Prima pagină devine documentul master
        first_page_rows = rows[0:3]
        master_doc = DocxFiller._fill_page(config, first_page_rows, data)

        if pages_needed == 1:
            master_doc.save(output_path)
            return

        # Adaugă paginile următoare prin inserare body elements în master
        master_body = master_doc.element.body
        # Asigură că body-ul master are sectPr (page break)
        master_sect = master_body.find(f'{{{W_NS}}}sectPr')

        for page_idx in range(1, pages_needed):
            page_rows = rows[page_idx * 3 : page_idx * 3 + 3]
            page_doc = DocxFiller._fill_page(config, page_rows, data)
            page_body = page_doc.element.body

            # Inserează page break la finalul ultimului paragraf din master
            # prin adăugarea unui <w:p><w:r><w:br w:type="page"/></w:r></w:p>
            pb_p = etree.SubElement(master_body, f'{{{W_NS}}}p')
            pb_r = etree.SubElement(pb_p, f'{{{W_NS}}}r')
            pb_br = etree.SubElement(pb_r, f'{{{W_NS}}}br')
            pb_br.set(f'{{{W_NS}}}type', 'page')
            # Mută pb_p înainte de sectPr dacă există
            if master_sect is not None:
                master_body.remove(pb_p)
                master_sect.addprevious(pb_p)

            # Copiază toți copiii body-ului paginii noi în master (mai puțin sectPr final)
            for child in list(page_body):
                if child.tag == f'{{{W_NS}}}sectPr':
                    continue
                master_body.append(child)

        master_doc.save(output_path)
```

- [ ] **Step 4: Verifică dependințe instalate**

```bash
pip install python-docx lxml openpyxl customtkinter
```

`docxcompose` NU este necesar.

- [ ] **Step 5: Rulează testele DocxFiller**

```bash
python -m pytest tests/test_docx_filler.py -v
```

Expected: toate 3 teste PASS.

- [ ] **Step 6: Test manual — deschide fișierul generat și verifică vizual**

```bash
python -c "
from confirmare_primire_app import DocxFiller, Config, ExcelReader
rows = ExcelReader.read(r'C:\EXCELURI\Tipizate\tests\sample_lista.xlsx')
cfg = Config.load()
DocxFiller.generate(rows, cfg, '23.03.2026', r'C:\EXCELURI\Tipizate\tests\output_test.docx')
print('Generat: output_test.docx')
"
```

Deschide `output_test.docx` în Word și verifică că textul apare în pozițiile corecte. Dacă nu, ajustează `FIELD_MAP`.

- [ ] **Step 7: Commit**

```bash
git add confirmare_primire_app.py tests/test_docx_filler.py
git commit -m "feat: DocxFiller — completare VML textboxes din Excel"
git push
```

---

## Task 4: UI — ConfigTab

**Files:**
- Modify: `C:\EXCELURI\Tipizate\confirmare_primire_app.py` (adaugă ConfigTab + App skeleton)

- [ ] **Step 1: Adaugă ConfigTab și App în `confirmare_primire_app.py`**

```python
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ConfigTab(ctk.CTkFrame):
    def __init__(self, parent, config: Config):
        super().__init__(parent)
        self.config = config
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text="Indicativ serviciu:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        ctk.CTkLabel(self, text="SSFAH", font=ctk.CTkFont(weight='bold')).grid(row=0, column=1, sticky='w', padx=10)

        fields = [
            ('Strada expeditor:', 'expeditor_strada'),
            ('Nr. expeditor:', 'expeditor_nr'),
            ('Cod postal expeditor:', 'expeditor_cod_postal'),
            ('Sector expeditor:', 'expeditor_sector'),
            ('Cale template:', 'template_path'),
        ]
        self._entries = {}
        for i, (label, key) in enumerate(fields, start=1):
            ctk.CTkLabel(self, text=label).grid(row=i, column=0, sticky='w', padx=10, pady=4)
            entry = ctk.CTkEntry(self, width=400)
            entry.insert(0, getattr(self.config, key))
            entry.grid(row=i, column=1, padx=10, pady=4)
            self._entries[key] = entry
            if key == 'template_path':
                ctk.CTkButton(self, text='Browse', width=80,
                              command=self._browse_template).grid(row=i, column=2, padx=5)

        ctk.CTkButton(self, text='Salvează configurare',
                      command=self._save).grid(row=len(fields)+1, column=1, pady=15, sticky='e')

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[('Word', '*.docx')])
        if path:
            self._entries['template_path'].delete(0, 'end')
            self._entries['template_path'].insert(0, path)

    def _save(self):
        for key, entry in self._entries.items():
            setattr(self.config, key, entry.get().strip())
        self.config.save()
        messagebox.showinfo('Salvat', 'Configurare salvată.')


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Confirmare de Primire — Generator')
        self.geometry('900x600')
        self.config_obj = Config.load()
        tabs = ctk.CTkTabview(self)
        tabs.pack(fill='both', expand=True, padx=10, pady=10)
        tabs.add('Configurare')
        tabs.add('Generare')
        ConfigTab(tabs.tab('Configurare'), self.config_obj).pack(fill='both', expand=True)
        # GenerareTab adăugat în Task 5
        self.gen_tab_frame = tabs.tab('Generare')

if __name__ == '__main__':
    app = App()
    app.mainloop()
```

- [ ] **Step 2: Testează vizual ConfigTab**

```bash
python "C:\EXCELURI\Tipizate\confirmare_primire_app.py"
```

Expected: fereastră cu 2 tab-uri, tab Configurare funcțional (poți completa și salva).

- [ ] **Step 3: Commit**

```bash
git add confirmare_primire_app.py
git commit -m "feat: ConfigTab UI — adresa expeditor SSFAH"
git push
```

---

## Task 5: UI — GenerareTab

**Files:**
- Modify: `C:\EXCELURI\Tipizate\confirmare_primire_app.py` (adaugă GenerareTab, conectează în App)

- [ ] **Step 1: Implementează GenerareTab**

```python
import threading
from datetime import date

class GenerareTab(ctk.CTkFrame):
    def __init__(self, parent, config: Config):
        super().__init__(parent)
        self.config = config
        self._rows = []
        self._checkvars = []
        self._build()

    def _build(self):
        # --- Rând 1: Browse Excel ---
        top = ctk.CTkFrame(self)
        top.pack(fill='x', padx=10, pady=(10, 0))
        self._path_var = ctk.StringVar(value='')
        ctk.CTkLabel(top, text='Fișier Excel:').pack(side='left')
        ctk.CTkEntry(top, textvariable=self._path_var, width=450).pack(side='left', padx=5)
        ctk.CTkButton(top, text='Browse', command=self._browse_excel).pack(side='left')

        # --- Rând 2: Data prezentării ---
        row2 = ctk.CTkFrame(self)
        row2.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(row2, text='Data prezentării:').pack(side='left')
        self._data_entry = ctk.CTkEntry(row2, width=120)
        self._data_entry.insert(0, date.today().strftime('%d.%m.%Y'))
        self._data_entry.pack(side='left', padx=5)

        # --- Tabel preview ---
        ctk.CTkLabel(self, text='Destinatari:').pack(anchor='w', padx=10)
        self._scroll = ctk.CTkScrollableFrame(self, height=300)
        self._scroll.pack(fill='both', expand=True, padx=10, pady=5)

        # --- Buton selectare + generare ---
        bot = ctk.CTkFrame(self)
        bot.pack(fill='x', padx=10, pady=5)
        ctk.CTkButton(bot, text='Selectează tot', command=self._select_all).pack(side='left', padx=5)
        ctk.CTkButton(bot, text='Deselectează', command=self._deselect_all).pack(side='left', padx=5)
        self._gen_btn = ctk.CTkButton(bot, text='Generează', command=self._generate,
                                       fg_color='green', state='disabled')
        self._gen_btn.pack(side='right', padx=5)
        self._progress = ctk.CTkProgressBar(bot)
        self._progress.pack(side='right', padx=10, fill='x', expand=True)
        self._progress.set(0)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx *.xls')])
        if not path:
            return
        try:
            self._rows = ExcelReader.read(path)
        except ValueError as e:
            messagebox.showerror('Eroare Excel', str(e))
            return
        self._path_var.set(path)
        self._excel_path = path
        self._build_preview()
        self._gen_btn.configure(state='normal')

    def _build_preview(self):
        for w in self._scroll.winfo_children():
            w.destroy()
        self._checkvars = []
        headers = ['', 'Nume Destinatar', 'Strada', 'Nr', 'Sector', 'Nr Inregistrare']
        for ci, h in enumerate(headers):
            ctk.CTkLabel(self._scroll, text=h, font=ctk.CTkFont(weight='bold')).grid(
                row=0, column=ci, padx=5, pady=2, sticky='w')
        for ri, row in enumerate(self._rows, start=1):
            var = ctk.BooleanVar(value=True)
            self._checkvars.append(var)
            ctk.CTkCheckBox(self._scroll, text='', variable=var, width=30).grid(
                row=ri, column=0, padx=5)
            for ci, key in enumerate(['Nume Destinatar', 'Strada', 'Nr', 'Sector', 'Nr Inregistrare'], 1):
                ctk.CTkLabel(self._scroll, text=row.get(key, '')).grid(
                    row=ri, column=ci, padx=5, sticky='w')

    def _select_all(self):
        for v in self._checkvars:
            v.set(True)

    def _deselect_all(self):
        for v in self._checkvars:
            v.set(False)

    def _generate(self):
        selected = [r for r, v in zip(self._rows, self._checkvars) if v.get()]
        if not selected:
            messagebox.showwarning('Atenție', 'Nu ai selectat niciun rând.')
            return
        data = self._data_entry.get().strip()
        out_dir = os.path.dirname(self._excel_path)
        out_file = os.path.join(out_dir,
                    f"Confirmare_primire_{date.today().strftime('%Y-%m-%d')}.docx")
        if os.path.exists(out_file):
            if not messagebox.askyesno('Fișier existent', f'{out_file}\nSuprascriem?'):
                return
        self._gen_btn.configure(state='disabled')
        self._progress.set(0.1)

        def run():
            try:
                DocxFiller.generate(selected, self.config, data, out_file)
                self.after(0, lambda: self._on_success(out_file))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror('Eroare', str(e)))
            finally:
                self.after(0, lambda: self._gen_btn.configure(state='normal'))
                self.after(0, lambda: self._progress.set(0))

        threading.Thread(target=run, daemon=True).start()

    def _on_success(self, path):
        self._progress.set(1)
        if messagebox.askyesno('Succes', f'Generat:\n{path}\n\nDeschizi folderul?'):
            os.startfile(os.path.dirname(path))
```

- [ ] **Step 2: Conectează GenerareTab în App**

```python
# în clasa App.__init__, după ConfigTab:
GenerareTab(self.gen_tab_frame, self.config_obj).pack(fill='both', expand=True)
```

- [ ] **Step 3: Test vizual complet end-to-end**

```bash
python "C:\EXCELURI\Tipizate\confirmare_primire_app.py"
```

Pași de verificare:
1. Tab Configurare → completează adresa DGASMB → Salvează
2. Tab Generare → Browse → selectează `tests\sample_lista.xlsx`
3. Verifică preview tabel (2 rânduri)
4. Apasă Generează
5. Deschide fișierul generat în Word → verifică că textele apar corect

- [ ] **Step 4: Commit final**

```bash
git add confirmare_primire_app.py
git commit -m "feat: GenerareTab UI — upload Excel, preview, generare docx"
git push
```

---

## Task 6: Polish + Ajustare FIELD_MAP

**Scop:** După testul manual din Task 3/5, ajustează `FIELD_MAP` dacă textele nu au apărut în pozițiile corecte.

- [ ] **Step 1: Deschide `output_test.docx` generat și compară cu formularul real (foto)**

Verifică fiecare câmp:
- Nume destinatar în locul corect?
- Adresă completă (strada, nr, bl, sc, et, ap, sector)?
- Nr. înregistrare în dreapta jos?
- Data prezentării în locul corect?

- [ ] **Step 2: Ajustează `FIELD_MAP` în cod pentru orice câmp greșit**

Reiterează: rulează diagnostic → reordoneaza indexuri → retestează.

- [ ] **Step 3: Rulează toate testele**

```bash
python -m pytest tests/ -v
```

Expected: toate testele PASS.

- [ ] **Step 4: Commit final cu FIELD_MAP corectat**

```bash
git add confirmare_primire_app.py
git commit -m "fix: ajustare FIELD_MAP dupa verificare vizuala"
git push
```

---

## Checklist Final

- [ ] `diagnostic_textboxes.py` rulat și FIELD_MAP completat
- [ ] `tests/test_excel_reader.py` → toate PASS
- [ ] `tests/test_docx_filler.py` → toate PASS
- [ ] App pornește fără erori
- [ ] Formular generat vizual corect (câmpuri în pozițiile corecte)
- [ ] Fișier salvat în folderul Excel-ului
- [ ] Config salvat persistent în `config.json`
- [ ] Push pe GitHub
