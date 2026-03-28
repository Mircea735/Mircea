# DGASMC Platform — Skills & Docs

Repo central pentru platforma internă DGASMC București.

## Structură

```
├── skills/          — skill-uri reutilizabile extrase din platformă
└── docs/            — spec-uri și planuri de implementare
    ├── plans/       — platformă principală (Sheets + Python CLI)
    └── superpowers/ — app-uri desktop standalone (CustomTkinter)
```

---

## Skills

Skill-uri independente extrase din dashboard-ul Gmail + Registru Sheets.

| Skill | Descriere | Stack |
|-------|-----------|-------|
| [gmail-oauth-dashboard](skills/gmail-oauth-dashboard/) | Inbox Gmail complet într-un singur fișier HTML | Vanilla JS, Google Identity Services |
| [anthropic-vision-ocr](skills/anthropic-vision-ocr/) | OCR PDF/imagine via Claude Vision API cu proxy CORS | JS, Anthropic API, PowerShell |
| [google-sheets-registru](skills/google-sheets-registru/) | Citire/scriere registru Sheets cu detecție rând corect | JS, Sheets API v4 |
| [reply-templates](skills/reply-templates/) | Template-uri email cu variabile și import .docx | Vanilla JS |
| [powershell-cors-proxy](skills/powershell-cors-proxy/) | Server HTTP local PowerShell cu proxy API extern | PowerShell |
| [email-registration-workflow](skills/email-registration-workflow/) | Flux complet OCR → Sheets → reply automat | JS, Gmail + Sheets + Anthropic API |
| [bi-dashboard-chartjs](skills/bi-dashboard-chartjs/) | Dashboard BI cu Chart.js, Sheets API, auto-refresh countdown | JS, Chart.js 4, Google Sheets API v4 |

---

## Docs

### Platform principală (`docs/plans/`)

| Data | Feature |
|------|---------|
| 2026-03-02 | Options Manager — Search Tab (caută beneficiar după CNP/email) |
| 2026-03-03 | Reinterogare Processor (CustomTkinter, TDD 10 tasks) |
| 2026-03-09 | Export Fișiere Separate |
| 2026-03-09 | Verificare Post-Reinterogare |
| 2026-03-10 | Interogare Termen |
| 2026-03-10 | Loop Processor |

### Superpowers (`docs/superpowers/`)

| Data | Feature |
|------|---------|
| 2026-03-13 | Adeverință Notariat — Tab Omise + Tab Notariat (Claude Vision OCR) |
| 2026-03-23 | Confirmare Primire CP14 — auto-fill formulare poștale din Excel |

---

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

Dezvoltat pentru **Serviciul Stimulent Financiar Adulți cu Handicap** — DGASMC București.
