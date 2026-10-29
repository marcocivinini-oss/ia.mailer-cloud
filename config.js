// Configurazione IA Mailer — modificare prima del deploy
window.IA_MAILER_CONFIG = {
  // Client ID OAuth creato in Google Cloud Console (tipo "Applicazione web")
  GOOGLE_CLIENT_ID: "INSERIRE_CLIENT_ID.apps.googleusercontent.com",

  // Domini autorizzati all'accesso (vuoto = qualsiasi account Google, es. caselle dedicate @gmail.com)
  ALLOWED_DOMAINS: ["iainternationaladvisors.com"],

  // Pausa tra un invio e l'altro (millisecondi). 1500 = ~40 email/minuto
  SEND_DELAY_MS: 1500,

  // Dati del footer aziendale
  COMPANY_FOOTER:
    "IA International Advisors S.r.l. · Via Fratelli Gabba 1/a, 20121 Milano<br>" +
    "IA International Advisors Inc. · 450 7th Ave. Suite 1504, New York NY",
  WEBSITE: "https://www.iainternationaladvisors.com",

  // Integrazione HubSpot attiva (richiede la variabile HUBSPOT_TOKEN su Netlify)
  HUBSPOT_ENABLED: true
};
