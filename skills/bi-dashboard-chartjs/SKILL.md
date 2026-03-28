---
name: bi-dashboard-chartjs
description: Dashboard BI single-file HTML cu Chart.js 4.x, Google Sheets API v4, OAuth2 GIS, auto-refresh countdown, filtre perioadă/operator — palette beige DGASMC
type: project
---

# BI Dashboard cu Chart.js + Google Sheets

Dashboard analitic single-file HTML pentru Registrul Email 43/2026 al DGASMC București.
Citește date live din Google Sheets via API v4, le filtrează în browser și redă 5 vizualizări
Chart.js + tabel recent. Auto-refresh la 60s cu countdown vizual în topbar.

## Ce face
- **KPI cards**: Total, Azi, Trimise (col W="Trimis"), Fără email (col Q gol)
- **Line chart**: înregistrări pe zi (ultimele N zile, area fill)
- **Doughnut chart**: top 10 tipuri document din col L (comma-split)
- **Bar chart operator**: înregistrări per operator (col O), top 12
- **Bar chart sector**: distribuție sectoare 1–6 (col N, culori fixe)
- **Tabel recent**: ultimele 30 rânduri cu badge Status email
- **Filtre**: perioadă (30/60/90 zile / tot 2026) + dropdown operator
- **Auto-refresh**: countdown 60s în topbar, buton ⏸/▶ toggle

## Stack
- Chart.js 4.4.4 (CDN umd)
- Google Identity Services (GIS) OAuth2 — token din localStorage, shared cu dashboard Gmail
- Google Sheets API v4 `/values/{range}` (GET, read-only scope)
- IBM Plex Sans + IBM Plex Mono (Google Fonts)
- Servit de PowerShell HttpListener pe `/bi`

## Pattern cheie — destroyChart safe re-render

Chart.js 4.x aruncă eroare dacă creezi un chart pe un canvas deja folosit.
Soluția corectă:

```javascript
function destroyChart(canvasId) {
  const existing = Chart.getChart(canvasId); // Chart.js 4.x registry lookup
  if (existing) existing.destroy();
}
// Apelat înainte de orice new Chart(...)
destroyChart('chartLine');
charts.line = new Chart(document.getElementById('chartLine'), { ... });
```

**Nu** folosi `charts.line && charts.line.destroy()` — poate fi null dacă pagina
s-a reîncărcat sau canvas-ul a fost re-montat.

## Pattern cheie — auto-refresh cu countdown vizual

```javascript
const AUTO_REFRESH_SEC = 60;
let _arInterval = null, _arCountdown = 0, _arActive = false;

function startAutoRefresh() {
  _arActive = true; _arCountdown = AUTO_REFRESH_SEC;
  _tickCountdown();
  _arInterval = setInterval(() => {
    if (--_arCountdown <= 0) { _arCountdown = AUTO_REFRESH_SEC; loadData(); }
    _tickCountdown();
  }, 1000);
}
function _tickCountdown() {
  document.getElementById('countdownTxt').textContent = _arActive ? `↻ ${_arCountdown}s` : '';
}
// În loadData() success: resetează countdown
if (_arActive) { _arCountdown = AUTO_REFRESH_SEC; _tickCountdown(); }
else if (!_arInterval) { startAutoRefresh(); } // pornire automată după primul load
```

## Pattern cheie — filtrare rânduri reale vs. stray rows

Sheets-ul poate conține rânduri "fantomă" (indicativ fără alte date). Filtru dual:

```javascript
allRows = rows.filter(r =>
  (r[COL.IND] || '').match(/R43\/\d+/) ||
  (r[COL.DATA]|| '').match(/\d+\.\d+\.\d+/)
);
```

Suficient pentru BI (mai permisiv decât getLastSheetRow care cere ≥2 câmpuri).

## Pattern cheie — doughnut top-10 tipuri document (comma-split)

Col L poate conține `"CI, Extras cont, Cerere"` — split și count per token:

```javascript
filteredRows.forEach(r => {
  (r[COL.CONT] || '').split(',').forEach(t => {
    const k = t.trim();
    if (k) counts[k] = (counts[k] || 0) + 1;
  });
});
const top10 = Object.entries(counts).sort((a,b) => b[1]-a[1]).slice(0, 10);
```

## Lecții învățate
- **Token shared**: GIS token stocat în `localStorage['gToken'/'gTokenExp']` e același între
  dashboard-ul principal și cel BI — nu cere re-autentificare dacă ambele rulează în același origin.
- **`Chart.getChart()` vs referință directă**: registrul intern Chart.js 4.x e mai fiabil decât
  a păstra referințe manuale — funcționează corect și dacă canvas-ul a fost recreat de DOM.
- **Auto-refresh pornit automat**: utilizatorii nu pornesc manual timere. Mai bun UX: pornește
  din oficiu după primul load reușit, cu posibilitate de oprire.
- **`applyFilters()` separat de `loadData()`**: datele brute (`allRows`) se încarcă o singură
  dată; filtrele re-calculează `filteredRows` și re-randează graficele în browser fără fetch nou.
  Crucial pentru responsivitate la schimbarea filtrelor.
