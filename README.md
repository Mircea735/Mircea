# DGASMC Platform Skills

Skills extrase din platforma internă DGASMC București — dashboard Gmail cu OCR, înregistrare Registru, și reply automat.

Fiecare skill este independent și reutilizabil în alte proiecte.

## Skills disponibile

| Skill | Descriere | Stack |
|-------|-----------|-------|
| [gmail-oauth-dashboard](skills/gmail-oauth-dashboard/) | Inbox Gmail complet într-un singur fișier HTML | Vanilla JS, Google Identity Services |
| [anthropic-vision-ocr](skills/anthropic-vision-ocr/) | OCR documente PDF/imagine via Claude Vision API | JS, Anthropic API, PowerShell proxy |
| [google-sheets-registru](skills/google-sheets-registru/) | Citire/scriere registru în Google Sheets cu detecție rând | JS, Sheets API v4 |
| [reply-templates](skills/reply-templates/) | Sistem template-uri email cu variabile și import/export | Vanilla JS |
| [powershell-cors-proxy](skills/powershell-cors-proxy/) | Server HTTP PowerShell cu proxy CORS pentru API-uri externe | PowerShell, System.Net.HttpListener |
| [email-registration-workflow](skills/email-registration-workflow/) | Flux complet: OCR → extragere câmpuri → înregistrare → reply automat | JS, Gmail API, Sheets API, Anthropic API |

## Arhitectură platformă

```
Email Gmail (inbox)
    ↓
Vizualizare atașamente (PDF/imagine)
    ↓
OCR Claude Vision API (proxy PowerShell)
    ↓
Extragere câmpuri: CNP, IBAN, Nume, Prenume, data deces etc.
    ↓
Înregistrare Registru Sheets (R43/NR/DATA)
    ↓
Reply automat confirmare (Gmail API)
```

## Context

Dezvoltat pentru **Serviciul Stimulent Financiar Adulți cu Handicap** — DGASMC București.
Rulează local pe `http://localhost:8080` via PowerShell, fără backend cloud.
