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

## Campagne graduali (a tappe)

Per le liste grandi: si carica il CSV completo e il Mailer lo invia in più giorni.
- Server: `netlify/functions/tappe.mjs` (archivio R2 `tappe/`): quota crescente (es. 50 → +25 al giorno → max 200),
  giorni e fascia oraria nel fuso dei destinatari, massimo per azienda al giorno, alternanza dei domini,
  esclusione di disiscritti e non recapitabili, tetto giornaliero per casella (`MAILBOX_DAILY_CAP`, default 1500),
  un solo invio attivo per campagna.
- Ogni giorno il proprietario apre il Mailer e preme «Invia la tappa di oggi»: il browser invia con pause casuali.
- Prima di ogni tappa (e ogni 50 invii) il server legge solo le intestazioni dei messaggi in arrivo collegati alla campagna
  (permesso Google `gmail.metadata`): segna risposte e risposte automatiche, registra i non recapitabili
  (archivio `nonrecapitabili`, esclusi da tutti gli invii futuri) e sospende la campagna oltre il 5% di mancati recapiti.
- Una campagna sospesa si riprende con «Riprendi: conta solo i nuovi invii» (il 5% vale solo per le email inviate dopo la ripresa)
  oppure con «Riprendi senza limite» (nessuna sospensione automatica; il limite si può riattivare dalla scheda della campagna).
