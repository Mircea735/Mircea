# Options Manager — Tab "Caută Beneficiar" Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Adaugă un tab "Caută" în SidebarV2.html care permite căutarea documentelor unui beneficiar după CNP sau email, cu vizualizare inline și export pe sheet "Sumar Beneficiar".

**Architecture:** 3 funcții noi în Code.gs (backend GAS) + modificare SidebarV2.html (frontend tab UI). Deploy prin script Python care citește proiectul existent și aplică modificările via Apps Script API. Nu se modifică logica existentă — doar adăugiri.

**Tech Stack:** Google Apps Script (JavaScript), HTML/CSS/JS inline, Python 3 + google-api-python-client pentru deploy.

---

## Referințe critice

- **Spreadsheet ID**: `1xFTA80PFIlByy29dBTNfcgeJn9FOXmeEOLCi18Gy6hs`
- **Script ID**: `18BkQLG-iV98F69WgAJ-KQXlGAXTvoL3Qc9EmzoQy3FPyZSnd3ndgaPKq`
- **Token OAuth**: `C:\Users\dgasm\token_apps_script.json`
- **Credentials**: `C:\Users\dgasm\credentials.json`
- **Script deploy existent**: `C:\Users\dgasm\update_options_manager_v2.py`
- **Design doc**: `C:\Users\dgasm\docs\plans\2026-03-02-options-manager-search-tab-design.md`

## Coloane registru (0-indexed)

| Col | Idx | Conținut |
|-----|-----|----------|
| E | 4 | Data înregistrării |
| L | 11 | Conținut pe scurt |
| M | 12 | CNP |
| Q | 16 | Email expeditor |

`getRange(2, 1, lastRow-1, 17)` — 17 coloane obligatoriu pentru a ajunge la col Q.

---

## Task 1: Creează scriptul de deploy pentru search tab

**Files:**
- Create: `C:\Users\dgasm\update_search_tab.py`

**Step 1: Creează fișierul cu structura de bază**

```python
"""
Deploy Tab "Caută Beneficiar" în SidebarV2.html + funcții backend în Code.gs
"""
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/script.projects']
SCRIPT_ID = '18BkQLG-iV98F69WgAJ-KQXlGAXTvoL3Qc9EmzoQy3FPyZSnd3ndgaPKq'
TOKEN_FILE = 'token_apps_script.json'
CREDENTIALS_FILE = 'credentials.json'

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds

if __name__ == '__main__':
    creds = get_credentials()
    print("Autentificat OK")
```

**Step 2: Rulează să verifici autentificarea**

```bash
cd C:\Users\dgasm
python update_search_tab.py
```

Expected: `Autentificat OK` (poate deschide browser la prima rulare)

**Step 3: Commit**

```bash
cd C:\Users\dgasm
git add update_search_tab.py
git commit -m "feat: add deploy script skeleton for search tab"
```

---

## Task 2: Adaugă funcțiile backend GAS în scriptul de deploy

**Files:**
- Modify: `C:\Users\dgasm\update_search_tab.py`

**Step 1: Adaugă constanta cu cele 3 funcții noi după imports**

```python
SEARCH_FUNCTIONS = '''

// ===== Options Manager — Caută Beneficiar =====

function getRegistrySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  return sheets
    .map(function(s) { return s.getName(); })
    .filter(function(name) { return name.indexOf('Registru') !== -1; });
}

function searchBeneficiar(query, sheetNames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  var q = query.trim();

  // Auto-detect: CNP (13 cifre) sau email (contine @)
  var isCnp    = /^\\d{13}$/.test(q);
  var isEmail  = q.indexOf('@') !== -1;

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    data.forEach(function(row) {
      var cnp   = (row[12] || '').toString().trim();
      var email = (row[16] || '').toString().trim();
      var match = false;

      if (isCnp)        match = cnp   === q;
      else if (isEmail) match = email.toLowerCase() === q.toLowerCase();
      else              match = cnp   === q || email.toLowerCase() === q.toLowerCase();

      if (match) {
        results.push({
          data:     (row[4]  || '').toString().trim(),
          continut: (row[11] || '').toString().trim(),
          registru: name
        });
      }
    });
  });

  return results;
}

function exportSumarBeneficiar(query, results) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Sumar Beneficiar';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

  // Header bloc
  sheet.appendRow(['CNP/Email căutat: ' + query, '', 'Generat: ' + today]);
  sheet.appendRow(['Data', 'Conținut', 'Registru']);

  results.forEach(function(r) {
    sheet.appendRow([r.data, r.continut, r.registru]);
  });

  // Separator
  sheet.appendRow(['']);

  return { rows: results.length, sheet: sheetName };
}
'''
```

**Step 2: Adaugă funcția `update_code_gs` în script**

```python
def update_code_gs(service):
    print("Citesc proiectul curent...")
    content = service.projects().getContent(scriptId=SCRIPT_ID).execute()
    files = content.get('files', [])

    updated = []
    for f in files:
        if f.get('name') == 'Code' and f.get('type') == 'SERVER_JS':
            # Evită duplicare dacă funcțiile există deja
            if 'getRegistrySheets' not in f.get('source', ''):
                f['source'] = f['source'] + SEARCH_FUNCTIONS
                print("  ✓ Code.gs: funcții backend adăugate")
            else:
                print("  ⚠ Code.gs: funcțiile există deja, skip")
        updated.append({'name': f['name'], 'type': f['type'], 'source': f.get('source', '')})

    service.projects().updateContent(
        scriptId=SCRIPT_ID,
        body={'files': updated}
    ).execute()
    print("  ✓ Code.gs salvat")
```

**Step 3: Actualizează `__main__` să apeleze funcția**

```python
if __name__ == '__main__':
    creds = get_credentials()
    service = build('script', 'v1', credentials=creds)
    update_code_gs(service)
    print("\n✓ Done! Reîncarcă Google Sheets și verifică meniul.")
```

**Step 4: Rulează și verifică**

```bash
cd C:\Users\dgasm
python update_search_tab.py
```

Expected:
```
Citesc proiectul curent...
  ✓ Code.gs: funcții backend adăugate
  ✓ Code.gs salvat

✓ Done! Reîncarcă Google Sheets și verifică meniul.
```

**Step 5: Test manual backend**
- Deschide Google Sheets → Extensii → Apps Script
- În consola GAS, rulează: `getRegistrySheets()`
- Expected: array cu sheet-urile "Registru*"
- Rulează: `searchBeneficiar("test@email.com", ["Registru E mail 43 2026"])`
- Expected: array gol sau cu rezultate (fără erori)

**Step 6: Commit**

```bash
git add update_search_tab.py
git commit -m "feat: add GAS backend functions for beneficiar search"
```

---

## Task 3: Construiește HTML-ul noului tab "Caută"

**Files:**
- Modify: `C:\Users\dgasm\update_search_tab.py` (adaugă constanta `SEARCH_TAB_HTML`)

**Step 1: Adaugă constanta cu fragmentul HTML al tab-ului "Caută"**

Acest fragment va fi injectat în SidebarV2.html existent.

```python
SEARCH_TAB_HTML = '''\
<!-- TAB BAR — adăugat deasupra search-box-ului existent -->
<div id="tab-bar" style="display:flex;border-bottom:2px solid #e0e0e0;margin-bottom:10px;">
  <button id="tab-optiuni" onclick="switchTab('optiuni')"
    style="flex:1;padding:7px;border:none;border-bottom:3px solid #1a73e8;
           background:#fff;font-size:13px;cursor:pointer;color:#1a73e8;font-weight:600;">
    Opțiuni
  </button>
  <button id="tab-cauta" onclick="switchTab('cauta')"
    style="flex:1;padding:7px;border:none;border-bottom:3px solid transparent;
           background:#fff;font-size:13px;cursor:pointer;color:#5f6368;">
    🔍 Caută
  </button>
</div>

<!-- PANEL OPȚIUNI — conținut existent învelit -->
<div id="panel-optiuni">
  <!-- PLACEHOLDER: conținut existent SidebarV2 se mută aici -->
</div>

<!-- PANEL CAUTĂ -->
<div id="panel-cauta" style="display:none;">
  <input id="search-query" type="text" placeholder="CNP (13 cifre) sau email..."
    style="width:100%;padding:7px 10px;border:1px solid #dadce0;border-radius:4px;
           font-size:13px;outline:none;margin-bottom:8px;box-sizing:border-box;">

  <div style="font-size:12px;color:#5f6368;margin-bottom:4px;">Caută în:</div>
  <div id="sheet-list" style="max-height:120px;overflow-y:auto;border:1px solid #e0e0e0;
       border-radius:4px;padding:4px 8px;margin-bottom:6px;font-size:12px;">
    <div style="color:#9aa0a6;font-style:italic;">Se încarcă...</div>
  </div>
  <div style="display:flex;gap:6px;margin-bottom:8px;">
    <button onclick="selectAllSheets(true)"
      style="flex:1;padding:4px;font-size:11px;border:1px solid #dadce0;
             border-radius:4px;cursor:pointer;background:#fff;">
      Selectează toate
    </button>
    <button onclick="selectAllSheets(false)"
      style="flex:1;padding:4px;font-size:11px;border:1px solid #dadce0;
             border-radius:4px;cursor:pointer;background:#fff;">
      Deselectează
    </button>
  </div>

  <button onclick="runSearch()"
    style="width:100%;padding:8px;background:#1a73e8;color:#fff;border:none;
           border-radius:4px;cursor:pointer;font-size:13px;margin-bottom:8px;">
    🔍 Caută
  </button>

  <div id="search-status" style="font-size:12px;color:#5f6368;min-height:18px;
       margin-bottom:6px;"></div>

  <div id="results-container" style="display:none;">
    <table id="results-table"
      style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:8px;">
      <thead>
        <tr style="background:#f8f9fa;">
          <th style="padding:4px 6px;border:1px solid #e0e0e0;text-align:left;">Data</th>
          <th style="padding:4px 6px;border:1px solid #e0e0e0;text-align:left;">Conținut</th>
          <th style="padding:4px 6px;border:1px solid #e0e0e0;text-align:left;">Registru</th>
        </tr>
      </thead>
      <tbody id="results-body"></tbody>
    </table>
    <button onclick="doExport()"
      style="width:100%;padding:7px;background:#fff;border:1px solid #34a853;
             color:#34a853;border-radius:4px;cursor:pointer;font-size:12px;">
      📋 Export în sheet
    </button>
  </div>
</div>

<script>
// ── Tab switching ──────────────────────────────────
function switchTab(tab) {
  var isOptiuni = tab === 'optiuni';
  document.getElementById('panel-optiuni').style.display = isOptiuni ? '' : 'none';
  document.getElementById('panel-cauta').style.display   = isOptiuni ? 'none' : '';
  document.getElementById('tab-optiuni').style.borderBottomColor = isOptiuni ? '#1a73e8' : 'transparent';
  document.getElementById('tab-optiuni').style.color = isOptiuni ? '#1a73e8' : '#5f6368';
  document.getElementById('tab-optiuni').style.fontWeight = isOptiuni ? '600' : 'normal';
  document.getElementById('tab-cauta').style.borderBottomColor = isOptiuni ? 'transparent' : '#1a73e8';
  document.getElementById('tab-cauta').style.color = isOptiuni ? '#5f6368' : '#1a73e8';
  document.getElementById('tab-cauta').style.fontWeight = isOptiuni ? 'normal' : '600';
  if (!isOptiuni && !window._sheetsLoaded) loadSheets();
}

// ── Sheet selector ─────────────────────────────────
window._sheetsLoaded = false;
window._lastResults  = [];

function loadSheets() {
  google.script.run
    .withSuccessHandler(function(names) {
      window._sheetsLoaded = true;
      var list = document.getElementById('sheet-list');
      if (!names || names.length === 0) {
        list.innerHTML = '<div style="color:#9aa0a6;">Niciun registru găsit</div>';
        return;
      }
      list.innerHTML = names.map(function(n) {
        var esc = n.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        return '<label style="display:block;padding:2px 0;cursor:pointer;">' +
          '<input type="checkbox" value="' + esc + '" checked style="margin-right:5px;">' +
          esc + '</label>';
      }).join('');
    })
    .withFailureHandler(function(e) {
      document.getElementById('sheet-list').innerHTML =
        '<div style="color:#c5221f;">Eroare: ' + e.message + '</div>';
    })
    .getRegistrySheets();
}

function selectAllSheets(checked) {
  var boxes = document.querySelectorAll('#sheet-list input[type=checkbox]');
  boxes.forEach(function(b) { b.checked = checked; });
}

function getSelectedSheets() {
  return Array.from(document.querySelectorAll('#sheet-list input[type=checkbox]'))
    .filter(function(b) { return b.checked; })
    .map(function(b) { return b.value; });
}

// ── Search ─────────────────────────────────────────
function runSearch() {
  var q = document.getElementById('search-query').value.trim();
  if (!q) { setStatus('Introdu CNP sau email.', 'red'); return; }

  var sheets = getSelectedSheets();
  if (sheets.length === 0) { setStatus('Selectează cel puțin un registru.', 'red'); return; }

  setStatus('Se caută...', 'gray');
  document.getElementById('results-container').style.display = 'none';

  google.script.run
    .withSuccessHandler(function(results) {
      window._lastResults = results;
      renderResults(results, q);
    })
    .withFailureHandler(function(e) {
      setStatus('Eroare: ' + e.message, 'red');
    })
    .searchBeneficiar(q, sheets);
}

function renderResults(results, q) {
  if (results.length === 0) {
    setStatus('Niciun document găsit pentru: ' + q, 'gray');
    document.getElementById('results-container').style.display = 'none';
    return;
  }
  setStatus(results.length + ' document(e) găsite pentru: ' + q, 'green');
  var tbody = document.getElementById('results-body');
  tbody.innerHTML = results.map(function(r) {
    var d = r.data.replace(/&/g,'&amp;').replace(/</g,'&lt;');
    var c = r.continut.replace(/&/g,'&amp;').replace(/</g,'&lt;');
    var reg = r.registru.replace(/&/g,'&amp;').replace(/</g,'&lt;');
    // Scurtăm numele registrului: "Registru E mail 43 2026" → "E43/2026"
    var short = reg.replace(/Registru\s*/i,'').replace(/\s+/g,' ');
    return '<tr>' +
      '<td style="padding:3px 6px;border:1px solid #e0e0e0;">' + d + '</td>' +
      '<td style="padding:3px 6px;border:1px solid #e0e0e0;">' + c + '</td>' +
      '<td style="padding:3px 6px;border:1px solid #e0e0e0;font-size:10px;color:#5f6368;" title="' + reg + '">' + short + '</td>' +
      '</tr>';
  }).join('');
  document.getElementById('results-container').style.display = '';
}

// ── Export ─────────────────────────────────────────
function doExport() {
  var q = document.getElementById('search-query').value.trim();
  if (!window._lastResults || window._lastResults.length === 0) return;

  setStatus('Se exportă...', 'gray');
  google.script.run
    .withSuccessHandler(function(res) {
      setStatus('✓ ' + res.rows + ' rânduri exportate pe sheet "' + res.sheet + '"', 'green');
    })
    .withFailureHandler(function(e) {
      setStatus('Eroare export: ' + e.message, 'red');
    })
    .exportSumarBeneficiar(q, window._lastResults);
}

// ── Utilitar ───────────────────────────────────────
function setStatus(msg, color) {
  var el = document.getElementById('search-status');
  el.textContent = msg;
  el.style.color = color === 'red' ? '#c5221f' : color === 'green' ? '#2e7d32' : '#5f6368';
}
</script>
'''
```

**Step 2: Commit fragmentul**

```bash
git add update_search_tab.py
git commit -m "feat: add search tab HTML fragment"
```

---

## Task 4: Injectează tab-ul în SidebarV2.html existent via API

**Files:**
- Modify: `C:\Users\dgasm\update_search_tab.py`

**Step 1: Adaugă funcția `update_sidebar_v2`**

Strategia: citește SidebarV2 existent, înfășoară conținutul `<body>` în `#panel-optiuni`, și injectează tab bar + panel "Caută" deasupra.

```python
def update_sidebar_v2(service):
    print("Citesc SidebarV2 existent...")
    content = service.projects().getContent(scriptId=SCRIPT_ID).execute()
    files = content.get('files', [])

    updated = []
    sidebar_found = False

    for f in files:
        if f.get('name') == 'SidebarV2' and f.get('type') == 'HTML':
            sidebar_found = True
            src = f.get('source', '')

            # Verifică dacă tab-ul există deja
            if 'panel-cauta' in src:
                print("  ⚠ SidebarV2: tab-ul există deja, skip")
                updated.append({'name': f['name'], 'type': f['type'], 'source': src})
                continue

            # Înfășoară conținut body existent în #panel-optiuni
            # SidebarV2 existent are <body>..conținut..</body>
            # Injectăm tab bar + panel-cauta ÎNAINTE de body content
            # și înfășurăm body content în div#panel-optiuni
            body_start = src.find('<body>') + len('<body>')
            body_end   = src.rfind('</body>')

            if body_start == -1 or body_end == -1:
                print("  ✗ Nu găsesc <body>...</body> în SidebarV2!")
                updated.append({'name': f['name'], 'type': f['type'], 'source': src})
                continue

            existing_body = src[body_start:body_end]

            # Construim noul body: tab bar + panel-optiuni (cu conținut existent) + panel-cauta
            new_body = (
                SEARCH_TAB_HTML.replace(
                    '<!-- PLACEHOLDER: conținut existent SidebarV2 se mută aici -->',
                    existing_body
                )
            )

            new_src = src[:body_start] + '\n' + new_body + '\n' + src[body_end:]
            f['source'] = new_src
            print("  ✓ SidebarV2.html: tab bar + panel Caută injectate")

        updated.append({'name': f['name'], 'type': f['type'], 'source': f.get('source', '')})

    if not sidebar_found:
        print("  ✗ SidebarV2 nu a fost găsit în proiect!")
        return False

    service.projects().updateContent(
        scriptId=SCRIPT_ID,
        body={'files': updated}
    ).execute()
    print("  ✓ SidebarV2 salvat")
    return True
```

**Step 2: Actualizează `__main__`**

```python
if __name__ == '__main__':
    creds = get_credentials()
    service = build('script', 'v1', credentials=creds)
    update_code_gs(service)
    update_sidebar_v2(service)
    print("\n✓ Done! Reîncarcă Google Sheets → Extensii → Options Manager → Open Options Manager v2")
```

**Step 3: Rulează deploy-ul complet**

```bash
cd C:\Users\dgasm
python update_search_tab.py
```

Expected:
```
Citesc proiectul curent...
  ✓ Code.gs: funcții backend adăugate
  ✓ Code.gs salvat
Citesc SidebarV2 existent...
  ✓ SidebarV2.html: tab bar + panel Caută injectate
  ✓ SidebarV2 salvat

✓ Done! Reîncarcă Google Sheets → Extensii → Options Manager → Open Options Manager v2
```

**Step 4: Commit**

```bash
git add update_search_tab.py
git commit -m "feat: inject search tab into SidebarV2 via Apps Script API"
```

---

## Task 5: Test manual complet în Google Sheets

**Step 1: Deschide sidebar**
- Google Sheets → Extensii → Options Manager → Open Options Manager v2
- Verifică că apare bara cu tab-urile `[ Opțiuni ] [ 🔍 Caută ]`
- Tab "Opțiuni" trebuie să funcționeze exact ca înainte

**Step 2: Test tab "Caută" — încărcare sheet-uri**
- Click pe tab "🔍 Caută"
- Sheet-urile din checkboxes trebuie să se încarce (toate cu "Registru" în nume)
- Verifică "Selectează toate" și "Deselectează"

**Step 3: Test căutare după email**
- Copiază un email real din col Q a unui rând din registru
- Introdu-l în câmpul de căutare
- Click "Caută"
- Expected: tabelul cu documetele acelui expeditor + status "X documente găsite"

**Step 4: Test căutare după CNP**
- Introdu un CNP de 13 cifre existent în col M
- Click "Caută"
- Expected: rezultate cu documentele acelui beneficiar

**Step 5: Test export**
- Cu rezultate vizibile, click "📋 Export în sheet"
- Expected: status "✓ X rânduri exportate pe sheet Sumar Beneficiar"
- Verifică în spreadsheet: sheet "Sumar Beneficiar" creat cu datele corecte

**Step 6: Test edge cases**
- Caută CNP inexistent → "Niciun document găsit"
- Caută fără text → "Introdu CNP sau email."
- Caută fără sheet selectat → "Selectează cel puțin un registru."
- Export de două ori → verifică că se adaugă un bloc nou (nu se suprascrie)

---

## Task 6: Fix eventuele probleme de injecție HTML

> **Notă**: Dacă structura `<body>` din SidebarV2 nu se potrivește exact, injecția automată poate eșua. În acest caz, folossim abordarea manuală.

**Dacă Task 4 eșuează** (mesajul `✗ Nu găsesc <body>...`):

**Step 1: Citește SidebarV2 curent**

```python
# script temporar de diagnostic
from update_search_tab import get_credentials
from googleapiclient.discovery import build

SCRIPT_ID = '18BkQLG-iV98F69WgAJ-KQXlGAXTvoL3Qc9EmzoQy3FPyZSnd3ndgaPKq'
creds = get_credentials()
service = build('script', 'v1', credentials=creds)
content = service.projects().getContent(scriptId=SCRIPT_ID).execute()
for f in content['files']:
    if f['name'] == 'SidebarV2':
        print(f['source'][:500])  # primele 500 caractere
```

**Step 2: Ajustează `body_start` / `body_end`** în `update_sidebar_v2` la structura reală găsită

**Step 3: Re-rulează deploy**

```bash
python update_search_tab.py
```

---

## Checklist final

- [ ] Backend: `getRegistrySheets()` returnează sheet-urile corecte
- [ ] Backend: `searchBeneficiar()` găsește după CNP, email, și ambele
- [ ] Backend: `exportSumarBeneficiar()` creează/append sheet "Sumar Beneficiar"
- [ ] UI: Tab bar vizibil cu ambele tab-uri
- [ ] UI: Tab "Opțiuni" funcționează exact ca înainte
- [ ] UI: Tab "Caută" încarcă sheet-urile dinamic
- [ ] UI: Tabelul rezultate se populează corect
- [ ] UI: Buton "Export" apare doar cu rezultate
- [ ] Edge cases: input gol, niciun sheet selectat, 0 rezultate
