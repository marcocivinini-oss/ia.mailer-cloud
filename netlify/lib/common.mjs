// Utilità condivise dalle funzioni Netlify di IA Mailer
import { getStore as netlifyGetStore } from "@netlify/blobs";
import { createHmac, timingSafeEqual } from "node:crypto";

export const json = (status, body) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });

// I test possono iniettare un archivio in memoria
export function getStore(name) {
  if (globalThis.__IA_TEST_STORES__) return globalThis.__IA_TEST_STORES__(name);
  return netlifyGetStore({ name, consistency: "strong" });
}

export const normEmail = (e) => String(e || "").trim().toLowerCase();
export const keyOf = (email) => encodeURIComponent(normEmail(email));
export const emailOf = (key) => decodeURIComponent(key);

/** Verifica il token Google dell'utente e il dominio autorizzato. Restituisce l'email o null */
export async function verifyGoogle(req) {
  const token = (req.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token) return null;
  const r = await fetch("https://oauth2.googleapis.com/tokeninfo?access_token=" + encodeURIComponent(token));
  if (!r.ok) return null;
  const info = await r.json();
  const cid = process.env.GOOGLE_CLIENT_ID;
  if (cid && info.azp !== cid && info.aud !== cid) return null;
  if (!info.email || String(info.email_verified) !== "true") return null;
  const list = (v) => (process.env[v] || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  const email = normEmail(info.email), domains = list("ALLOWED_DOMAINS"), emails = list("ALLOWED_EMAILS");
  if (domains.length && !domains.includes(email.split("@")[1]) && !emails.includes(email)) return null;
  return email;
}

// ---------- link di disiscrizione firmati ----------
const secret = () => {
  if (!process.env.UNSUBSCRIBE_SECRET) throw new Error("Manca la variabile UNSUBSCRIBE_SECRET");
  return process.env.UNSUBSCRIBE_SECRET;
};
const b64url = (s) => Buffer.from(s, "utf8").toString("base64url");
export const sign = (email) => createHmac("sha256", secret()).update(normEmail(email)).digest("base64url").slice(0, 32);
export function unsubscribeUrl(base, email, campaign) {
  const q = new URLSearchParams({ e: b64url(normEmail(email)), t: sign(email) });
  if (campaign) q.set("c", campaign);
  return base.replace(/\/$/, "") + "/disiscrizione?" + q;
}
export function verifyUnsub(e, t) {
  try {
    const email = normEmail(Buffer.from(String(e || ""), "base64url").toString("utf8"));
    const a = Buffer.from(sign(email)), b = Buffer.from(String(t || ""));
    return a.length === b.length && timingSafeEqual(a, b) ? email : null;
  } catch { return null; }
}

/** Registra una disiscrizione (idempotente) */
export async function addUnsubscribe(email, source, by, campaign) {
  const store = getStore("disiscritti"), key = keyOf(email);
  const prev = await store.get(key, { type: "json" });
  if (prev) return prev;
  const rec = { email: normEmail(email), date: new Date().toISOString(), source, by: by || "", campaign: campaign || "" };
  await store.setJSON(key, rec);
  return rec;
}

export async function listAll(store) {
  const { blobs } = await store.list();
  return blobs.map((b) => b.key);
}

/** Esegue fn su items con concorrenza limitata */
export async function pool(items, n, fn) {
  const out = new Array(items.length); let i = 0;
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, async () => {
    while (i < items.length) { const k = i++; out[k] = await fn(items[k], k); }
  }));
  return out;
}
