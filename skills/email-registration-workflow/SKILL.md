---
name: email-registration-workflow
description: Flux complet de înregistrare documente primite prin email — OCR atașament → extragere câmpuri → modal înregistrare pre-completat → scriere Sheets → reply automat confirmare → confirmare în sidebar.
type: project
---

# Email Registration Workflow

Flux end-to-end pentru înregistrarea documentelor administrative primite prin email.

## Fluxul complet

```
1. Email primit cu atașament (PDF/imagine)
        ↓
2. Click pe atașament → vizualizare în sidebar
        ↓
3. Selectare câmpuri OCR dorite (checklist)
        ↓
4. OCR via Claude Opus 4.6 Vision (proxy PowerShell)
        ↓
5. Câmpuri extrase persistă în sidebar (CNP, Nume, Prenume, IBAN etc.)
        ↓
6. Click "📋 Reg." → modal pre-completat automat:
   - Nume/Prenume/CNP din OCR
   - Email expeditor din metadate email
   - Operator din localStorage
   - Număr înregistrare calculat automat (R43/NR/DATA)
        ↓
7. Bifează tipuri documente primite
   → Denumire atașamente generată automat: "R43/823/26.03.2026/CI, Extras cont.pdf"
        ↓
8. Click "✓ Înregistrează" → scriere în Sheets la rândul corect
        ↓
9. Reply automat trimis expeditorului (Template T7 + footer standard)
        ↓
10. Col T="Trimis", Col U=data în Sheets (bypass Apps Script)
        ↓
11. Confirmare verde în sidebar OCR:
    "✓ Înregistrat: R43/823 | Email trimis la expeditor@email.com"
```

## Auto-fill modal din OCR + email meta

```javascript
async function showRegModal() {
  const ocr = persistOcrData || {};
  document.getElementById('regNume').value    = ocr.NUME || '';
  document.getElementById('regPrenume').value = ocr.PRENUME || '';
  document.getElementById('regCnp').value     = ocr.CNP || '';
  // Email din metadate email curent sau ultimul vizualizat
  document.getElementById('regEmail').value   = curEmailMeta
    ? extractEmail(curEmailMeta.from)
    : (_lastSenderEmail || '');
  // Atașamente din email curent
  document.getElementById('regAtas').value    = curEmailMeta?.attachments?.join(', ') || '';
}
```

## Auto-generare denumire atașamente

```javascript
function syncDocOptInput() {
  const checked = [...document.querySelectorAll('#regDocOpts input[data-docopt]:checked')]
    .map(el => el.dataset.docopt);
  document.getElementById('regContinut').value = checked.join(', ');
  // Format: R43/NR/DATA/DocType1, DocType2.pdf
  if (checked.length) {
    const nr = document.getElementById('regNextNr').dataset.nr || '';
    document.getElementById('regAtas').value = `${nr ? nr + '/' : ''}${checked.join(', ')}.pdf`;
  }
}
```

## Reply automat după înregistrare

```javascript
// Template T7: "Numărul de înregistrare al documentelor transmise instituției
// noastre este: {NR}, iar acestea vor fi procesate în cel mai scurt timp posibil."
const replyBody = `Numărul de înregistrare al documentelor transmise instituției noastre este: ${nr}, `
  + `iar acestea vor fi procesate în cel mai scurt timp posibil.`
  + FOOTER.replace('{OP}', operator);

const raw = buildMime({ to: email, subject: replySubj, body: replyBody, inReplyTo: _lastMsgId });
await api(`${BASE}/messages/send`, { method: 'POST', body: JSON.stringify({ raw, threadId: _lastThreadId }) });

// Marchează în Sheets: T="Trimis", U=data
const dateU = `${dd}.${mm}.${yyyy}`;
await api(sheetsUrl, { method: 'PUT', body: JSON.stringify({ values: [['Trimis', dateU]] }) });
```

## Confirmare în sidebar

```javascript
function showRegConfirmInSidebar(nr, toEmail, emailStatus) {
  const fields = document.getElementById('ocrPersistFields');
  const confHtml = `<div style="padding:8px;background:rgba(34,197,94,.1);border:1.5px solid rgba(34,197,94,.4)">
    <div>✓ Înregistrat în Registru</div>
    <div>Indicativ: ${nr}</div>
    <div>Email: ${toEmail}</div>
    <div>${emailStatus}</div>
  </div>`;
  fields.innerHTML = confHtml + fields.innerHTML; // prepend
}
```

## Globale persistente (supraviețuiesc navigării)

```javascript
let persistOcrData   = null;  // câmpuri OCR extrase
let curEmailMeta     = null;  // metadate email curent
let _lastSenderEmail = '';    // email expeditor — persistent
let _lastThreadId    = '';    // thread Gmail — pentru reply în același fir
let _lastMsgId       = '';    // Message-ID — pentru In-Reply-To
let _lastSubject     = '';    // subiect — pentru Re: subiect
```
