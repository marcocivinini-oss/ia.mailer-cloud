// IA Mailer su Cloudflare Workers.
// - L'app è protetta da Cloudflare Access (portale IA Advisors Hub); in più ogni chiamata al server
//   presenta il token Google dell'utente, verificato come prima.
// - /disiscrizione è PUBBLICA (link nelle email e pulsante "annulla iscrizione" di Gmail/Outlook):
//   su Cloudflare va esclusa dal login con un'applicazione Access di tipo Bypass sul path "disiscrizione".
import { bindBlobs } from "./blobs-shim.js";
import registro from "../netlify/functions/registro.mjs";
import hubspot from "../netlify/functions/hubspot.mjs";
import disiscrizione from "../netlify/functions/disiscrizione.mjs";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    bindBlobs(env.DATA);
    if (url.pathname === "/disiscrizione/logo.png") {
      return env.ASSETS.fetch(new Request(new URL("/assets/logo-email.png", url), request));
    }
    if (url.pathname === "/disiscrizione") return disiscrizione(request);
    // config.js: il Client ID Google arriva dalla variabile GOOGLE_CLIENT_ID del pannello Cloudflare
    if (url.pathname === "/config.js") {
      const r = await env.ASSETS.fetch(request);
      let js = await r.text();
      if (env.GOOGLE_CLIENT_ID) js = js.replace("INSERIRE_CLIENT_ID.apps.googleusercontent.com", env.GOOGLE_CLIENT_ID);
      return new Response(js, { headers: { "content-type": "application/javascript; charset=utf-8", "cache-control": "no-store" } });
    }
    if (url.pathname === "/.netlify/functions/registro") return registro(request);
    if (url.pathname === "/.netlify/functions/hubspot") return hubspot(request);
    if (url.pathname.startsWith("/.netlify/")) return new Response("Not found", { status: 404 });
    return env.ASSETS.fetch(request);
  },
};
