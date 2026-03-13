# Design Spec: Tab "Notariat" — Adeverință pentru Birouri Notariale

**Data:** 2026-03-13
**Fișier vizat:** `C:\EXCELURI\Adeverinte\adeverinta_notar_app.py`

---

## Problemă

Instituția primește solicitări de la birouri notariale pentru adeverințe privind stimulentul neîncasat al beneficiarilor decedați. Aceste adeverințe au un format diferit față de cele standard: sunt adresate direct biroului notarial, conțin datele de contact ale notariatului și referința la solicitarea externă. Generarea manuală e laborioasă.

---

## Soluție

Tab dedicat **"Notariat"** cu:
1. Încărcare opțională PDF solicitare → extragere câmpuri via **Claude Vision API** (Haiku 4.5)
2. Completare/corectare manuală — câmpurile sunt **întotdeauna editabile** indiferent dacă s-a folosit OCR
3. Căutare CNP din `engine.raw_data` (date deja încărcate din tab-ul Principal)
4. Generare `.docx` via template separat

---

## Confirmare dependențe

- `engine.raw_data` și `engine.personal_data` — populate de `load_excel()` din tab Principal; reutilizate direct
- `engine.exceptii_op` — dict `{mk: suma}` setat din tab Principal; reutilizat de `compute_outstanding()`
- `generate_all(filter_cnp=cnp)` — parametru deja existent
- `_replace_para_content()` — metoda existentă de substituție placeholder, folosită și pentru template-ul notariat
- PyMuPDF (`fitz`) — deja instalat
- Claude API: model `claude-haiku-4-5-20251001` (**verificat** — ID corect pentru Haiku 4.5 conform documentației Anthropic aug 2025)

---

## Arhitectură

### 1. Layout tab

**Secțiunea "Template & Output"** (sus):
- `CTkEntry` + buton `"Browse"` → `notariat_template_path` (selectat de utilizator)
- Nu se persistă între sesiuni (comportament identic cu tab-ul Principal)
- `out_dir` — refolosit din `self.out_dir` (setat din tab Principal). Dacă `out_dir` e `None` → butonul "Generează" rămâne dezactivat + mesaj: `"Setați folderul output în tab-ul Principal"`

**Secțiunea "Solicitare & Birou Notarial"**:
- Buton `"Încarcă PDF solicitare"` + label nume fișier
- OCR analizează **doar prima pagină** (pagina cu antetul biroului — în practică toate solicitările notariale sunt pe o singură pagină). Dacă PDF are mai multe pagini: label status afișează `"OCR: pagina 1 din N analizată"`.
- Câmpuri (CTkEntry, toate editabile):
  - `Denumire birou`
  - `Email birou`
  - `Nume notar` (pentru salut, ex: `Violeta Elena Manea`)
  - Gen notar: `CTkSegmentedButton` cu `["Doamnă", "Domnule"]`
  - `Nr. solicitare externă` (ex: `498`)
  - `Dată solicitare` (ex: `12.03.2026`, format `DD.MM.YYYY`)
  - `Nr. înregistrare intern` (ex: `43/659`)
  - `Dată înregistrare intern` (ex: `13.03.2026`, format `DD.MM.YYYY`)

**Secțiunea "Beneficiar"**:
- `CTkEntry` CNP (13 cifre) + buton `"Caută"`
- Date afișate read-only după căutare reușită (din `engine.raw_data[cnp]` și `engine.personal_data.get(cnp)`):
  - Nume + Prenume — din `raw_data[cnp]['nume']`, `raw_data[cnp]['prenume']`
  - Nr. dosar — din `raw_data[cnp]['nr_dosar']`
  - Data înregistrare — din `format_data_inreg(raw_data[cnp]['data_inreg_raw'])`
  - CI serie + nr + dată — din `personal_data[cnp]['ci_serie']`, `['ci_nr']`, `['data_ci']`
  - Adresă — din `personal_data[cnp]['adresa']`
  - Perioadă + sumă restantă — din `compute_outstanding(cnp)` (poate fi `None`)
- Mesaje:
  - CNP negăsit în `raw_data`: `"CNP negăsit — verificați că raportul Excel este încărcat"`
  - CNP găsit dar `compute_outstanding()` returnează `None`: `"CNP găsit — fără restanțe detectate. Adeverința se va genera cu suma 0."`

**Jos**:
- Buton `"Generează adeverință notarială"` — dezactivat implicit
- Se activează când TOATE: `notariat_template_path` setat + `out_dir` setat + CNP găsit în `raw_data` + `denumire_birou` nevid + `nr_solicitare` nevid
- Label status generare

### 2. Extragere OCR via Claude Vision

OCR rulează în **thread separat** (UI nu îngheață). Pre-completarea câmpurilor se face via `self.after(0, ...)` din thread.

```python
def _ocr_extract_pdf(self, pdf_path):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        self.after(0, lambda: self._set_ocr_status("ANTHROPIC_API_KEY negăsit — completați manual"))
        return

    try:
        import fitz, anthropic, base64, json
        doc = fitz.open(pdf_path)
        pix = doc[0].get_pixmap(dpi=200)
        img_b64 = base64.standard_b64encode(pix.tobytes("png")).decode()

        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                {"type": "text", "text": PROMPT_OCR_NOTARIAT}
            ]}]
        )
        data = json.loads(response.content[0].text)
        self.after(0, lambda d=data: self._apply_ocr_result(d))
    except Exception as e:
        self.after(0, lambda: self._set_ocr_status(f"OCR indisponibil: {e} — completați manual"))
```

**`PROMPT_OCR_NOTARIAT`** (constantă în cod):
```
Extrage din această imagine a unui document notarial românesc câmpurile de mai jos.
Returnează DOAR un obiect JSON valid, fără text suplimentar:
{
  "denumire_birou": "denumirea completă a biroului/societății notariale",
  "email": "adresa de email",
  "nr_solicitare": "numărul din antet (doar cifre, ex: 498)",
  "data_solicitare": "data din antet format DD.MM.YYYY",
  "cnp_beneficiar": "CNP-ul persoanei decedate (13 cifre)",
  "nume_beneficiar": "numele complet al defunctului"
}
Dacă un câmp nu poate fi identificat cu certitudine, folosește null.
```

**`_apply_ocr_result(data)`** — mapează câmpurile JSON la widget-uri:

| Câmp JSON | Widget |
|-----------|--------|
| `denumire_birou` | `self.entry_birou` |
| `email` | `self.entry_email` |
| `nr_solicitare` | `self.entry_nr_sol` |
| `data_solicitare` | `self.entry_data_sol` |
| `cnp_beneficiar` | `self.entry_cnp` → **auto-trigger `_cauta_cnp_notariat()`** |
| `nume_beneficiar` | ignorat (va apărea din lookup CNP) |

Fiecare câmp non-null: `.delete(0, 'end'); .insert(0, val)`. Câmpuri null: nemodificate.

### 3. Căutare CNP

`_cauta_cnp_notariat()`: lookup în `engine.raw_data`. Dacă găsit → populează secțiunea Beneficiar (read-only labels). Dacă `personal_data.get(cnp)` lipsește → câmpurile CI/adresă afișează `[LIPSĂ]`.

### 4. Generare document

Rulează în **thread separat**. Pe durata rulării: butonul "Generează" din TOATE tab-urile dezactivat (via `_set_busy(True)` — metoda existentă de coordonare cross-tab). La finalizare: `_set_busy(False)`.

```python
def _run_notariat_generation(self):
    cnp = self.entry_cnp.get().strip()
    outstanding = self.engine.compute_outstanding(cnp)
    # outstanding poate fi None → suma=0, luni_text='—', perioada='—'

    birou = self.entry_birou.get().strip()
    nr_sol = self.entry_nr_sol.get().strip()
    data_sol = self.entry_data_sol.get().strip()
    nr_inreg = self.entry_nr_inreg.get().strip()
    data_inreg = self.entry_data_inreg.get().strip()

    # Construiește placeholderele compuse
    nr_sol_ext = f"Nr.{nr_sol}/{data_sol}" if nr_sol and data_sol else nr_sol
    nr_inreg_int = f"{nr_inreg}/{data_inreg}" if nr_inreg and data_inreg else nr_inreg
    salut = f"Stimat{'ă' if self.seg_gen.get()=='Doamnă' else 'e'} {self.seg_gen.get()} {self.entry_nume_notar.get().strip()}"

    self.engine.generate_notariat(
        cnp=cnp,
        template_path=self.notariat_template_path,
        out_dir=self.out_dir,
        outstanding=outstanding,
        birou=birou,
        email=self.entry_email.get().strip(),
        salut=salut,
        nr_solicitare_ext=nr_sol_ext,
        nr_inreg_int=nr_inreg_int,
    )
```

### 5. Metodă `generate_notariat()` în engine

**Implementare:**
```python
def generate_notariat(self, cnp, template_path, out_dir, outstanding,
                      birou, email, salut, nr_solicitare_ext, nr_inreg_int):
    p = self.raw_data[cnp]
    pd = self.personal_data.get(cnp, {})

    # Valori din outstanding sau fallback
    if outstanding:
        suma       = str(outstanding['suma'])
        luni_text  = outstanding['luni_text']
        perioada   = f"{outstanding['period_start']} – {outstanding['period_end']}"
        data_deces = outstanding['data_deces']
    else:
        suma = '0'; luni_text = '—'; perioada = '—'
        data_deces = format_data_deces(p['data_deces_raw']) or '—'

    replacements = {
        '{{DENUMIRE_BIROU}}': birou,
        '{{EMAIL_BIROU}}':    email,
        '{{SALUT}}':          salut,
        '{{NR_SOLICITARE_EXT}}': nr_solicitare_ext,
        '{{NR_INREG_INT}}':   nr_inreg_int,
        '{{NUME_PRENUME}}':   f"{p['prenume']} {p['nume']}",
        '{{CNP}}':            cnp,
        '{{DOMICILIU}}':      pd.get('adresa', '[LIPSĂ]'),
        '{{CI_SERIE}}':       pd.get('ci_serie', '[LIPSĂ]'),
        '{{CI_NR}}':          pd.get('ci_nr', '[LIPSĂ]'),
        '{{CI_DATA}}':        pd.get('data_ci', '[LIPSĂ]'),
        '{{PERIOADA}}':       perioada,
        '{{SUMA}}':           suma,
        '{{LUNI_TEXT}}':      luni_text,
        '{{NR_DOSAR}}':       p['nr_dosar'],
        '{{DATA_INREG}}':     format_data_inreg(p['data_inreg_raw']),
        '{{DATA_DECES}}':     data_deces,
    }

    doc = Document(template_path)
    for para in doc.paragraphs:
        for ph, val in replacements.items():
            if ph in para.text:
                _replace_para_content(para, ph, val)  # metoda existentă
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for ph, val in replacements.items():
                        if ph in para.text:
                            _replace_para_content(para, ph, val)

    nr_sol_raw = nr_solicitare_ext.replace('Nr.', '').split('/')[0].strip()
    filename = f"NOTARIAT - {p['prenume']} {p['nume']} - {nr_sol_raw}.docx"
    doc.save(os.path.join(out_dir, filename))
    return filename
```

**Dacă `personal_data.get(cnp)` lipsește:** placeholderele CI/adresă se înlocuiesc cu `[LIPSĂ]` — documentul se generează totuși, utilizatorul vede explicit ce lipsește.

**Filename:** `nr_sol_raw` = numărul brut extras din `nr_solicitare_ext` (ex: `Nr.498/12.03.2026` → `498`). Nu conține slash-uri sau caractere invalide pentru filename.

### 6. Validare date calendaristice

Câmpurile `Dată solicitare` și `Dată înregistrare intern` sunt text liber. Nu se validează la input (utilizatorul poate folosi orice format). Se inserează as-is în placeholder; responsabilitatea formatării e a utilizatorului. Mesaj hint sub câmp: `(ex: 12.03.2026)`.

### 7. Stare inițială tab

Buton "Generează" disabled. Mesaj: `"Încărcați raportul Excel din tab-ul Principal, apoi completați câmpurile."`

---

## Ce NU se modifică

- Tab-ul Principal, tab-ul Omise
- `load_excel()`, `compute_outstanding()`, `generate_all()`
- Formatul adeverințelor existente (moștenitori)

---

## Fișiere afectate

| Fișier | Modificare |
|--------|-----------|
| `adeverinta_notar_app.py` | Adăugare tab Notariat + `generate_notariat()` în engine |
| `Adeverinta Notariat master.docx` | De creat de utilizator cu placeholder-ele specificate |

**Metode noi în UI:** `_load_pdf_notariat()`, `_ocr_extract_pdf()`, `_apply_ocr_result()`, `_cauta_cnp_notariat()`, `_run_notariat_generation()`
**Metodă nouă în engine:** `generate_notariat(cnp, template_path, out_dir, outstanding, birou, email, salut, nr_solicitare_ext, nr_inreg_int)`

---

## Criterii de acceptare

1. Buton "Încarcă PDF" → file picker → OCR pre-completează câmpurile detectate; CNP detectat auto-declanșează căutarea beneficiarului
2. Toate câmpurile rămân editabile după OCR
3. Câmpuri nedetectate de OCR rămân goale (nu blochează utilizatorul)
4. Eroare API Claude → mesaj clar în label status, aplicația nu cade
5. `ANTHROPIC_API_KEY` lipsă → mesaj clar, fără crash
6. Buton "Caută" CNP → găsit: afișare date beneficiar; negăsit: mesaj eroare
7. CNP găsit fără restanțe → mesaj informativ, generarea e totuși posibilă (suma 0)
8. Buton "Generează" activ doar când: template setat + out_dir setat + CNP găsit + denumire birou + nr. solicitare completate
9. `out_dir` nesetat → buton dezactivat + mesaj "Setați folderul output în tab-ul Principal"
10. Generarea produce `.docx` cu toate placeholder-ele înlocuite corect (inclusiv cele cu runs fragmentate)
11. Filename: `NOTARIAT - {Prenume} {Nume} - {nr_solicitare}.docx`
12. Tab-ul Principal și Omise nu sunt afectate
13. Generarea și OCR rulează în thread-uri separate (UI nu îngheață)
14. Pe durata generării, butoanele "Generează" din toate tab-urile sunt dezactivate via `_set_busy()`
15. OP-partial exceptions (`engine.exceptii_op`) din tab Principal sunt reutilizate automat în calculul sumei
