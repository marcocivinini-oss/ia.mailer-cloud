# IA Mailer — guida all'installazione

Web app per inviare newsletter e mailing personalizzati **dalla propria casella Google Workspace** (o da una casella dedicata), partendo da un testo, da un design Canva esportato in immagini o da un HTML già pronto. Destinatari da Excel/CSV, elenco incollato o liste HubSpot.

Ogni advisor accede con il proprio account Google: le email partono dalla sua casella, restano nei suoi "Inviati" e le risposte tornano a lui. Nessuna password viene salvata.

## Cosa contiene

| File | Ruolo |
|---|---|
| `index.html`, `style.css`, `app.js` | Interfaccia (stile IA: navy, arancio, Poppins) |
| `mailer-core.js` | Composizione email, campi personalizzati, costruzione MIME |
| `config.js` | **Da compilare**: Client ID Google, domini autorizzati, footer |
| `netlify/functions/hubspot.mjs` | Lettura sicura delle liste HubSpot (il token resta sul server) |
| `netlify/functions/registro.mjs` | Registro condiviso di invii, campagne e disiscritti (Netlify Blobs) |
| `netlify/functions/disiscrizione.mjs` | Pagina pubblica di disiscrizione raggiunta dal link nelle email |
| `netlify/lib/common.mjs` | Autenticazione, firma dei link, archivio |
| `assets/` | Logo ufficiale IA |

## Pubblicazione su mailer.ia-advisors-hub.com

IA Mailer è un sito Netlify a sé, raggiungibile dal sottodominio **mailer.ia-advisors-hub.com** e collegato all'hub con una voce di menu. L'hub non va modificato, salvo il link.

1. **GitHub** – crea il repository privato `ia-mailer` e caricaci il contenuto della cartella (da github.com: *Add file → Upload files*, trascinando tutti i file e le cartelle).
2. **Netlify** – *Add new site → Import an existing project → GitHub → ia-mailer*. Le impostazioni di build vengono lette da `netlify.toml`: non serve modificarle. Rinomina il sito in `ia-mailer` (*Site configuration → Change site name*).
3. **Variabili d'ambiente** – vedi il punto 3 più sotto (`GOOGLE_CLIENT_ID`, `ALLOWED_DOMAINS`, `UNSUBSCRIBE_SECRET`, `HUBSPOT_TOKEN`).
4. **Dominio** – *Domain management → Add a domain → `mailer.ia-advisors-hub.com`*.
   - Se il DNS di ia-advisors-hub.com è gestito da Netlify (stesso account dell'hub), il sottodominio si configura da solo.
   - Altrimenti, nel pannello del registrar aggiungi un record **CNAME** `mailer` → `ia-mailer.netlify.app`.
   Attendi il certificato HTTPS (pochi minuti), poi imposta `mailer.ia-advisors-hub.com` come **dominio principale** del sito.
5. **Google Cloud** – nell'ID client OAuth aggiungi l'origine `https://mailer.ia-advisors-hub.com` (punto 1.4).
6. **Voce di menu nell'hub** – nel repository dell'hub aggiungi un link, ad esempio:
   ```html
   <a href="https://mailer.ia-advisors-hub.com" target="_blank" rel="noopener">IA Mailer · Newsletter</a>
   ```
7. **Prova** – apri il Mailer, accedi, invia una prova a te stesso e clicca «annulla l'iscrizione»: la pagina deve aprirsi su mailer.ia-advisors-hub.com. Poi riattiva il tuo indirizzo dalla sezione 5.

> Importante: completa il punto 4 **prima del primo invio reale**. I link di disiscrizione usano il dominio principale del sito al momento dell'invio.

## 1. Google Cloud (una tantum, ~10 minuti)

1. Vai su https://console.cloud.google.com, crea il progetto **IA Mailer** (con l'account amministratore Workspace).
2. **API e servizi → Libreria**: abilita **Gmail API**.
3. **Schermata consenso OAuth**: tipo utente **Interno** (solo account @iainternationaladvisors.com, nessuna verifica Google richiesta). Nome app "IA Mailer", logo IA.
   Ambiti: `gmail.send`, `gmail.settings.basic`, `userinfo.email`, `userinfo.profile`.
4. **Credenziali → Crea credenziali → ID client OAuth → Applicazione web**.
   Origini JavaScript autorizzate: `https://mailer.ia-advisors-hub.com` (e, solo per i test iniziali, `https://ia-mailer.netlify.app`).
5. Copia il **Client ID** in `config.js` → `GOOGLE_CLIENT_ID`.

> Casella dedicata (es. `newsletter@iainternationaladvisors.com`): basta accedere con quell'account. In alternativa, se la casella è configurata come alias "Invia come" in Gmail, compare nel menu "Indirizzo mittente".
> Account @gmail.com esterni: impostare la schermata consenso su "Esterno" e aggiungerli come utenti di test, e svuotare `ALLOWED_DOMAINS`.

## 2. HubSpot (facoltativo)

1. HubSpot → Impostazioni → Integrazioni → **App private** → Crea app "IA Mailer".
2. Ambiti: `crm.lists.read`, `crm.objects.contacts.read`.
3. Copia il token di accesso.

## 3. Netlify

1. Repository GitHub e sito Netlify come descritto in «Pubblicazione su mailer.ia-advisors-hub.com».
2. **Site configuration → Environment variables**:
   - `HUBSPOT_TOKEN` = token dell'app privata HubSpot
   - `GOOGLE_CLIENT_ID` = stesso Client ID di `config.js`
   - `ALLOWED_DOMAINS` = `iainternationaladvisors.com`
   - `ALLOWED_EMAILS` = (facoltativo) eventuali indirizzi esterni autorizzati, separati da virgola
   - `UNSUBSCRIBE_SECRET` = una stringa casuale lunga (es. generata con un password manager): firma i link di disiscrizione. **Non cambiarla dopo il primo invio**, altrimenti i link già spediti smettono di funzionare.

   L'archivio del registro (Netlify Blobs) si attiva da solo al primo deploy: non serve alcun database.
3. Deploy. Aggiungi l'URL definitivo tra le origini autorizzate in Google Cloud (punto 1.4).

## 4. Collegamento all'hub

Voce di menu verso https://mailer.ia-advisors-hub.com (vedi sopra). In alternativa si può incorporare in una pagina dell'hub:
`<iframe src="https://mailer.ia-advisors-hub.com" style="width:100%;height:100vh;border:0"></iframe>` (già autorizzato in `netlify.toml`).

## Uso

1. **Mittente** – accedi con Google, scegli indirizzo, nome visualizzato e firma (vengono ricordati).
2. **Contenuto** – oggetto e preheader; poi:
   - *Testo*: editor con grassetto, titoli, elenchi, link, campi `{{Nome}}`;
   - *Da Canva*: in Canva → Condividi → Scarica → PNG (tutte le pagine); carica le immagini, aggiungi link e testo alternativo; il PDF può andare negli allegati;
   - *HTML*: incolla un HTML pronto.
   Opzioni: cornice grafica IA, pulsante call-to-action, riga di disiscrizione.
3. **Destinatari** – Excel/CSV (colonna Email obbligatoria, le altre diventano campi), elenco incollato o lista HubSpot (i disiscritti HubSpot sono esclusi). Duplicati e indirizzi non validi vengono scartati; le esclusioni (indirizzi o `@dominio`) vengono ricordate.
4. **Anteprima e invio** – controlla l'anteprima per ogni destinatario, invia una prova, poi "Invia". Un'email individuale per destinatario, con pausa tra gli invii, pausa/ripresa, nuovo tentativo sui falliti e report CSV.

## Registro condiviso: disiscritti e storico invii

Il registro è unico per tutti gli advisor e sta sul sito Netlify (Netlify Blobs).

- **Disiscrizioni**: ogni email contiene un link personale e firmato «annulla l'iscrizione». Il destinatario conferma su una pagina con logo IA e da quel momento è escluso da *tutti* gli invii futuri, di qualunque advisor. È attivo anche il pulsante «Annulla iscrizione» che Gmail e Outlook mostrano accanto al mittente (one-click, RFC 8058), oggi richiesto da Google per chi invia a molti destinatari.
- **Chi risponde «CANCELLAMI»**: nella sezione 5 → Disiscritti → «Aggiungi disiscritti a mano».
- **Storico**: ogni invio riuscito viene registrato (destinatario, oggetto, mittente, data). Caricando una lista, la colonna *Storico* mostra quante email ha ricevuto ciascun contatto e quando. Puoi escludere automaticamente chi ha già ricevuto un'email con lo stesso oggetto (attivo di default) o chi è stato contattato negli ultimi 7/14/30/90 giorni.
- **Campagne**: la sezione 5 elenca tutte le campagne (data, oggetto, mittente, inviati/errori) con l'elenco dei destinatari scaricabile in CSV. Elenco disiscritti esportabile in CSV.
- «Riattiva» rimuove un indirizzo dai disiscritti: usarlo solo su richiesta esplicita dell'interessato.
- Le email di prova non vengono registrate.

**Campi personalizzati**: `{{Nome}}`, `{{Cognome}}`, `{{Azienda}}`, `{{Email}}` o qualsiasi colonna del file. Con valore di riserva: `{{Nome|cliente}}`.

## Limiti e buone pratiche

- Google Workspace: **2.000 email/giorno per casella** (Gmail personale: 500). Per invii 100–500 non ci sono problemi.
- Tenere la pagina aperta durante l'invio (~40 email al minuto). Se si chiude, ricaricando la lista con lo stesso oggetto chi ha già ricevuto l'email viene escluso automaticamente.
- GDPR: inviare solo a contatti con base giuridica valida e lasciare attiva la riga di disiscrizione. Il registro conserva la prova della disiscrizione (data e origine). I disiscritti non vengono aggiornati in HubSpot in automatico: per allinearlo, esporta il CSV dei disiscritti e importalo in HubSpot come opt-out.
- Per volumi molto superiori o statistiche di apertura/clic, valutare un servizio dedicato (es. Brevo, HubSpot Marketing Email).
