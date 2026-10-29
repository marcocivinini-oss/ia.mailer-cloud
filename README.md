# IA Mailer — Cloudflare

Newsletter e mailing dalla casella Google di ogni utente, all'indirizzo **mailer.ia-advisors-hub.com**
(accesso dal portale IA Advisors Hub). Funzionamento e uso: vedi `README-originale.md`.

- Interfaccia: cartella `public` (invariata, più il ritorno al portale).
- Server: `worker/index.js` esegue le funzioni originali (`netlify/functions`): registro condiviso, HubSpot, disiscrizione.
- Registro (disiscritti, storico invii, campagne): spazio R2 `ia-mailer-data`.
- `/disiscrizione` è pubblica: in Cloudflare Access serve un'applicazione Bypass sul path `disiscrizione`.

Pannello Cloudflare → Worker `ia-mailer` → Variables and Secrets:
- `GOOGLE_CLIENT_ID` (Text): Client ID OAuth di Google Cloud — viene inserito anche in `config.js` in automatico
- `UNSUBSCRIBE_SECRET` (Secret): stringa casuale lunga; non cambiarla dopo il primo invio
- `HUBSPOT_TOKEN` (Secret, facoltativo): token dell'app privata HubSpot

Pubblicazione: Build command vuoto · Deploy command `npx wrangler deploy`.
