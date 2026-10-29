// IA Mailer su Cloudflare Workers.
// - L'app è protetta da Cloudflare Access (portale IA Advisors Hub); in più ogni chiamata al server
//   presenta il token Google dell'utente, verificato come prima.
// - /disiscrizione è PUBBLICA (link nelle email e pulsante "annulla iscrizione" di Gmail/Outlook):
//   su Cloudflare va esclusa dal login con un'applicazione Access di tipo Bypass sul path "disiscrizione".
import { bindBlobs } from "./blobs-shim.js";
import registro from "../netlify/functions/registro.mjs";
import hubspot from "../netlify/functions/hubspot.mjs";
import disiscrizione from "../netlify/functions/disiscrizione.mjs";
import { verifyGoogle } from "../netlify/lib/common.mjs";

const jsonRes = (status, obj) => new Response(JSON.stringify(obj), { status, headers: { "content-type": "application/json" } });
// Cache breve del controllo del token Google (evita una verifica a ogni singola email)
const tokenCache = new Map();
async function checkUser(request) {
  const token = (request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token) return null;
  const hit = tokenCache.get(token);
  if (hit && hit.exp > Date.now()) return hit.email;
  const email = await verifyGoogle(request);
  if (email) tokenCache.set(token, { email, exp: Date.now() + 5 * 60 * 1000 });
  if (tokenCache.size > 500) tokenCache.clear();
  return email;
}

// Invio di un'email: il browser consegna il messaggio al Worker, che lo passa a Gmail
// con il token Google dell'utente. Gmail lo spedisce dalla casella dell'utente, come prima.
async function invia(request) {
  if (request.method !== "POST") return jsonRes(405, { error: "Metodo non consentito" });
  const email = await checkUser(request);
  if (!email) return jsonRes(401, { error: "Accesso non autorizzato" });
  const mime = await request.arrayBuffer();
  if (mime.byteLength > 35 * 1024 * 1024) return jsonRes(413, { error: "Messaggio troppo grande (max 35 MB)" });
  const r = await fetch("https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media", {
    method: "POST",
    headers: { Authorization: request.headers.get("authorization"), "Content-Type": "message/rfc822" },
    body: mime,
  });
  return new Response(await r.text(), { status: r.status, headers: { "content-type": "application/json" } });
}

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
    if (url.pathname === "/.netlify/functions/invia") {
      try { return await invia(request); }
      catch (e) { return jsonRes(502, { error: "Servizio Google non raggiungibile: " + (e.message || e) }); }
    }
    if (url.pathname.startsWith("/.netlify/")) return new Response("Not found", { status: 404 });
    return env.ASSETS.fetch(request);
  },
};
