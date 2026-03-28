---
name: gmail-oauth-dashboard
description: Inbox Gmail complet într-un singur fișier HTML static, cu OAuth2 Google Identity Services, filtre, labels, reply, draft, forward și statistici email.
type: project
---

# Gmail OAuth Dashboard (Single-File)

Inbox Gmail funcțional complet într-un **singur fișier HTML** (~2500 linii), servit local via PowerShell HTTP server.

## Ce face

- Autentificare OAuth2 cu Google Identity Services (GIS) — token stocat în `localStorage`
- Listare emailuri cu filtre: label, zile, necitite
- **Stats bar** — statistici automate pe categorii (Restanțe, Conturi, Deces, Dosar, Necitite) cu filtrare click
- Vizualizare email complet: HTML sandbox + plain text cu quoted content ascuns
- Reply, draft, forward cu atașamente
- Operator field persistent în `localStorage`
- Marcare citit/necitit, trash

## Stack

- Vanilla JavaScript (no framework)
- Google Identity Services (`accounts.google.com/gsi/client`)
- Gmail REST API v1 (`gmail.googleapis.com`)
- OAuth2 implicit flow cu refresh automat

## OAuth Scopes necesare

```javascript
const SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.compose',
  'https://www.googleapis.com/auth/spreadsheets'
];
```

## Setup

1. Google Cloud Console → OAuth2 Client ID (Web application)
2. Authorized JavaScript origins: `http://localhost:8080`
3. Client ID introdus în ecranul de setup al dashboard-ului
4. Token salvat automat în `localStorage`, refresh la expirare

## Pattern cheie — Token refresh

```javascript
async function ensureToken() {
  if (TOKEN && Date.now() < TOKEN_EXP - 60000) return;
  return new Promise((res, rej) => {
    tokenClient.callback = r => {
      if (r.error) return rej(r);
      TOKEN = r.access_token;
      TOKEN_EXP = Date.now() + r.expires_in * 1000;
      res();
    };
    tokenClient.requestAccessToken({prompt: TOKEN ? '' : 'consent'});
  });
}
```

## Pattern cheie — API call cu auth

```javascript
async function api(url, opts = {}) {
  await ensureToken();
  const r = await fetch(url, {
    ...opts,
    headers: { Authorization: 'Bearer ' + TOKEN, 'Content-Type': 'application/json', ...(opts.headers || {}) }
  });
  if (!r.ok) throw new Error(`${r.status}: ${await r.text()}`);
  return r.json();
}
```

## Stats bar — categorii automate

```javascript
const CATS = [
  { key:'rest',  label:'Restanțe', keywords:['restanță','fonduri','buget','plată','amânare'] },
  { key:'cont',  label:'Conturi',  keywords:['cont','iban','extras','bancă','transfer'] },
  { key:'deces', label:'Deces',    keywords:['deces','decedat','moștenitor','certificat deces'] },
  { key:'dosar', label:'Dosar',    keywords:['dosar','cerere','certificat','handicap','înregistrare'] },
];
```
