---
name: reply-templates
description: Sistem de template-uri email cu categorii, variabile placeholder, preview live, import .docx și persistență în localStorage.
type: project
---

# Reply Templates System

Sistem complet de template-uri pentru răspunsuri email, cu categorii, variabile dinamice și import din documente Word.

## Template-uri built-in (exemple DGASMC)

| ID | Nume | Categorie |
|----|------|-----------|
| T1 | Lipsă fonduri — scurt | Restanțe |
| T2 | PMB nu a alocat fondurile | Restanțe |
| T7 | Confirmare înregistrare documente | Înregistrare |
| T8 | Acte necesare deces | Deces |
| T9 | Reînnoire certificat — nu e necesar | Dosar |

## Variabile suportate

```javascript
const VAR_MAP = {
  '{NR_INREGISTRARE}': () => document.getElementById('regNextNr').dataset.nr || '',
  '{LUNA_ELIGIBILA}': () => /* calcul lună eligibilă */,
  '{NR_CERERE}': () => /* din OCR sau manual */,
  '{OP}': () => document.getElementById('opInput').value || 'Operator',
};
```

## Footer standard

```javascript
const FOOTER = `\n\nCu stimă,\n{OP}\nServiciul Stimulent Financiar Adulți cu Handicap\n
Direcția Generală de Asistență Socială a Municipiului București\n
Str. Constantin Mille, nr. 10, sector 1, București\n
Tel. +4 021 314.23.15, tasta 1\n
https://www.dgas.ro/...\n\n
⚠ Modificări (CI, domiciliu, deces, cont, tutore) — anunțați în 48h.\n
⚠ Verificați situația fiscală la DITL în lunile martie și septembrie.`;
```

## Persistență localStorage

```javascript
function saveCustomTemplates() {
  const custom = TEMPLATES.filter(t => t.custom);
  localStorage.setItem('dgasmc_templates', JSON.stringify(custom));
}

function loadCustomTemplates() {
  try {
    const saved = JSON.parse(localStorage.getItem('dgasmc_templates') || '[]');
    TEMPLATES = [...TEMPLATES_BUILTIN, ...saved];
  } catch { TEMPLATES = [...TEMPLATES_BUILTIN]; }
}
```

## Import din .docx

```javascript
async function importDocxTemplate() {
  // Citește fișier .docx, extrage text via FileReader
  // Aplică ca body nou template
  const file = input.files[0];
  const reader = new FileReader();
  reader.onload = async e => {
    // Parsare simplă XML din .docx (word/document.xml)
    const zip = await JSZip.loadAsync(e.target.result);
    const xml = await zip.file('word/document.xml').async('string');
    const text = xml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    addTemplate({ name: file.name.replace('.docx',''), body: text, cat: 'Import', custom: true });
  };
}
```

## Filtrare și căutare

```javascript
function filterTmpl() {
  const q = document.getElementById('tmplSearch').value.toLowerCase();
  renderTemplates(TEMPLATES.filter(t =>
    t.name.toLowerCase().includes(q) ||
    t.body.toLowerCase().includes(q) ||
    t.cat.toLowerCase().includes(q)
  ));
}
```
