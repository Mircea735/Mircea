---
name: google-sheets-registru
description: Integrare Google Sheets pentru registru de intrări — citire, scriere la rândul corect, viewer cu scroll la data curentă, selecție rând, detecție limită grid și extindere automată.
type: project
---

# Google Sheets Registru Integration

Integrare completă cu Google Sheets pentru gestionarea unui registru de documente primite prin email.

## Sheet țintă

- **Spreadsheet ID**: `1xFTA80PFIlByy29dBTNfcgeJn9FOXmeEOLCi18Gy6hs`
- **Sheet**: `Registru E mail 43 2026`

## Mapping coloane (A–Y)

| Col | Index | Conținut |
|-----|-------|---------|
| A | 0 | Extras cont |
| B | 1 | Luna |
| C | 2 | Ziua |
| D | 3 | Anul |
| E | 4 | Data înregistrării |
| F | 5 | Nr. și data documentului |
| G | 6 | Indicativ (R43/NR/DATA) |
| H | 7 | Nume |
| I | 8 | Prenume |
| J | 9 | Nume Prenume |
| L | 11 | Conținut document |
| M | 12 | CNP |
| N | 13 | Sector |
| O | 14 | Operator |
| Q | 16 | Email expeditor |
| R | 17 | Observații |
| T | 19 | Eroare email (Apps Script) |
| U | 20 | Data trimiterii (Apps Script) |
| Y | 24 | Denumire atașamente |

## Detecție rând corect pentru scriere

Evită rânduri "fantomă" (ex: R43/955 la poziție fizică greșită) prin validare multi-coloană:

```javascript
async function getLastSheetRow() {
  // Un rând e real dacă are cel puțin 2 din: Data(E), Conținut(L), CNP(M), Email(Q)
  const range = encodeURIComponent(`'${SHEET_NAME}'!A:Q`);
  const data = await api(`${SHEETS_BASE}/${SHEETS_ID}/values/${range}`);
  const rows = (data.values || []);
  let lastRow = 1;
  rows.forEach((r, i) => {
    const filled = [r[4], r[11], r[12], r[16]]
      .filter(v => v && String(v).trim()).length;
    if (filled >= 2) lastRow = i + 1;
  });
  return lastRow;
}
```

## Scriere cu extindere automată la limita grid

```javascript
async function writeRow(sheetRow, rowData) {
  const r = encodeURIComponent(`'${SHEET_NAME}'!A${sheetRow}:Y${sheetRow}`);
  const url = `${SHEETS_BASE}/${SHEETS_ID}/values/${r}?valueInputOption=USER_ENTERED`;
  try {
    await api(url, { method: 'PUT', body: JSON.stringify({ values: [rowData] }) });
  } catch (e) {
    if (e.message && e.message.includes('exceeds grid limits')) {
      // Adaugă 200 rânduri și reîncearcă
      await api(`${SHEETS_BASE}/${SHEETS_ID}:batchUpdate`, {
        method: 'POST',
        body: JSON.stringify({
          requests: [{ appendDimension: { sheetId: 1630405736, dimension: 'ROWS', length: 200 } }]
        })
      });
      await api(url, { method: 'PUT', body: JSON.stringify({ values: [rowData] }) });
    } else throw e;
  }
}
```

## Viewer cu scroll la data curentă

```javascript
async function refreshRegistruViewer() {
  const rows = await fetchAllRows();
  const today = new Date();
  const todayStr = `${today.getDate()}.${today.getMonth()+1}.${today.getFullYear()}`;

  let firstTodayId = null;
  rows.forEach((row, i) => {
    const sheetRow = i + 2; // row 1 = header
    const isToday = (row[4] || '') === todayStr;
    if (isToday && !firstTodayId) firstTodayId = `rvr-${sheetRow}`;
    // render cu highlight albastru pentru azi
  });

  // Auto-scroll la primul rând de azi
  if (firstTodayId) {
    setTimeout(() => {
      document.getElementById(firstTodayId)?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }, 150);
  }
}
```

## Număr de înregistrare — detecție corectă

Scanează doar până la ultimul rând real (nu întregul sheet):

```javascript
async function getNextRegNumber() {
  const lastReal = await getLastSheetRow();
  const range = encodeURIComponent(`'${SHEET_NAME}'!G1:G${lastReal}`);
  const data = await api(`${SHEETS_BASE}/${SHEETS_ID}/values/${range}`);
  const col = (data.values || []).flat();
  let maxN = 0;
  col.forEach(v => {
    const m = String(v).match(/R43\/(\d+)\//);
    if (m) maxN = Math.max(maxN, parseInt(m[1]));
  });
  const t = new Date();
  return `R43/${maxN + 1}/${t.getDate()}.${t.getMonth()+1}.${t.getFullYear()}`;
}
```
