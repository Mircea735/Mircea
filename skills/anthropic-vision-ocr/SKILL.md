---
name: anthropic-vision-ocr
description: OCR documente PDF și imagini via Claude Vision API (Opus 4.6), cu selecție câmpuri de extras, proxy CORS PowerShell și persistență date extrase în sidebar.
type: project
---

# Anthropic Vision OCR pentru Documente Administrative

OCR structurat pe documente administrative românești (CI, extras cont, certificat deces, CAF etc.) folosind Claude Opus 4.6 Vision.

## Ce face

- Suportă **PDF** (via `anthropic-beta: pdfs-2024-09-25`) și **imagini** (JPEG, PNG, WebP)
- **Checklist câmpuri** selectabile înainte de OCR — extrage doar ce e necesar
- Date extrase persistă în sidebar chiar când navighezi la alte secțiuni
- Integrare directă în email dashboard: click pe atașament → OCR → câmpuri în sidebar

## Câmpuri suportate

| Cheie | Label |
|-------|-------|
| `CNP` | CNP beneficiar |
| `IBAN` | IBAN (extras de cont) |
| `NUME` | Nume beneficiar |
| `PRENUME` | Prenume beneficiar |
| `data_deces` | Data deces |
| `nr_certificat_deces` | Nr. certificat deces |
| `titular_CI` | Nume titular CI |
| `titular_cont` | Titular cont (extras) |
| `adresa` | Adresă domiciliu |
| `nr_CI` | Serie și nr. CI |
| `data_nastere` | Data naștere |
| `sector` | Sector București |

## Problema CORS și soluția

Anthropic API blochează cererile direct din browser (CORS). Soluție: **PowerShell HTTP server ca proxy**.

API key transmis ca query parameter (nu header) pentru a evita coruperea în .NET HttpListener:

```javascript
const ocrUrl = '/api/ocr?k=' + encodeURIComponent(apiKey) + (isPDF ? '&beta=pdfs-2024-09-25' : '');
const resp = await fetch(ocrUrl, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({
    model: 'claude-opus-4-6',
    max_tokens: 1024,
    messages: [{ role: 'user', content: [
      { type: isPDF ? 'document' : 'image',
        source: { type: 'base64', media_type: mime, data: base64data } },
      { type: 'text', text: prompt }
    ]}]
  })
});
```

## Prompt structurat pentru extragere selectivă

```javascript
function buildOCRPrompt(selectedFields) {
  const fieldList = selectedFields.map(f => `- ${f.key}: ${f.label}`).join('\n');
  return `Ești un asistent OCR pentru documente administrative românești.
Extrage DOAR aceste câmpuri din document (dacă există):
${fieldList}

Răspunde STRICT în format JSON, fără explicații:
{"${selectedFields[0].key}": "valoare", ...}
Dacă un câmp nu există în document, omite-l din JSON.`;
}
```

## Persistență în sidebar

```javascript
let persistOcrData = null; // supraviețuiește navigării între secțiuni

function renderOCRFields(data) {
  persistOcrData = { ...data };
  updateOcrPersistPanel(); // actualizează panoul din template sidebar
}
```

## Model recomandat

`claude-opus-4-6` — acuratețe 85-90% pe documente românești scanate.
