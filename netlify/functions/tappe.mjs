// IA Mailer — campagne graduali ("a tappe").
// Il CSV completo viene caricato una volta; ogni giorno il browser chiede il lotto del giorno e lo invia
// dalla casella Google dell'utente, con pause casuali. Il server decide QUANTI e QUALI destinatari:
// quota crescente (riscaldamento), giorni e fascia oraria, tetto per azienda (dominio), alternanza dei domini,
// esclusione di disiscritti e non recapitabili, tetto giornaliero della casella, un solo invio attivo per campagna.
//
// Archivio "tappe":
//   <id>            dati della campagna (impostazioni, avanzamento, contatori)
//   <id>/stato      elenco compatto dei destinatari: [email, dominio, stato, data, threadId, risposta, errore]
//   <id>/contenuto  il messaggio (testo, immagini, allegati) — salvato e restituito così com'è
//   <id>/campi      i campi personalizzati per destinatario (Nome, Azienda, …) — salvati così come sono
import { json, verifyGoogle, getStore, keyOf, normEmail, listAll, pool, unsubscribeUrl, logSends, addBounced } from "../lib/common.mjs";
import { rawStore } from "../../worker/blobs-shim.js";

// stati del destinatario
export const ST = { PEND: 0, OK: 1, ERR: 2, UNSUB: 3, BOUNCE: 4, NOREC: 5 };
const MAX_RECIPIENTS = 20000;
const MAX_CONTENT = 40 * 1024 * 1024, MAX_FIELDS = 30 * 1024 * 1024;
const LOCK_MS = 20 * 60 * 1000;
const BOUNCE_LIMIT = 0.05, BOUNCE_MIN_SENT = 20;

const newId = () => "g" + new Date().toISOString().replace(/[-:T]/g, "").slice(0, 12) + "-" + Math.random().toString(36).slice(2, 7);
const isEmail = (e) => /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(e);
const domainOf = (e) => String(e).split("@")[1] || "";

// Domini di posta personale: niente tetto per dominio (sono persone diverse), ma si alternano comunque
const CONSUMER = /^(gmail|googlemail|yahoo|ymail|hotmail|outlook|live|msn|icloud|me|mac|aol|libero|virgilio|alice|tim|tin|tiscali|fastwebnet|email|inwind|iol|gmx|web|protonmail|proton|yandex|mail|qq|163|126|uol|bol|terra|zoho)\.(com|net|it|de|fr|es|co\.uk|uk|com\.br|br|com\.ar|ar|com\.mx|mx|ru|cn|ch|at|be|nl|ca|jp|eu|me)$/i;
export const isConsumer = (d) => CONSUMER.test(d);

/** Alterna i domini: un destinatario per azienda a giro, così le email alla stessa azienda non arrivano di seguito */
export function interleave(list, dom) {
  const groups = new Map();
  for (const x of list) { const d = dom(x); if (!groups.has(d)) groups.set(d, []); groups.get(d).push(x); }
  const gs = [...groups.values()], out = [];
  while (out.length < list.length) for (const g of gs) if (g.length) out.push(g.shift());
  return out;
}

// ---------- date nel fuso orario della campagna ----------
export function nowIn(tz, d = new Date()) {
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-GB", {
    timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23", weekday: "short",
  }).formatToParts(d).map((x) => [x.type, x.value]));
  const wd = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[p.weekday];
  return { date: `${p.year}-${p.month}-${p.day}`, hm: `${p.hour}:${p.minute}`, wd };
}
const addDay = (date) => { const d = new Date(date + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + 1); return d.toISOString().slice(0, 10); };
const WD = ["", "lun", "mar", "mer", "gio", "ven", "sab", "dom"];
const fmtDay = (date, wd) => `${WD[wd]} ${Number(date.slice(8, 10))}/${Number(date.slice(5, 7))}`;

function cleanSettings(s = {}) {
  const n = (v, min, max, d) => { v = Math.round(Number(v)); return Number.isFinite(v) ? Math.min(max, Math.max(min, v)) : d; };
  const hm = (v, d) => (/^([01]\d|2[0-3]):[0-5]\d$/.test(v) ? v : d);
  let tz = "Europe/Rome";
  try { if (s.tz) { new Intl.DateTimeFormat("en", { timeZone: s.tz }); tz = s.tz; } } catch { /* fuso non valido */ }
  const o = {
    start: n(s.start, 5, 1500, 50), step: n(s.step, 0, 500, 25), max: n(s.max, 5, 1500, 200), perDomain: n(s.perDomain, 1, 100, 2),
    winStart: hm(s.winStart, "09:00"), winEnd: hm(s.winEnd, "17:30"), tz,
    days: [...new Set((Array.isArray(s.days) ? s.days : [1, 2, 3, 4, 5]).map(Number).filter((d) => d >= 1 && d <= 7))].sort(),
    gapMin: n(s.gapMin, 2, 600, 6), gapMax: n(s.gapMax, 2, 900, 15), pauseEvery: n(s.pauseEvery, 0, 1000, 25),
  };
  if (o.start > o.max) o.start = o.max;
  if (o.gapMax < o.gapMin) o.gapMax = o.gapMin;
  if (o.winEnd <= o.winStart) { o.winStart = "09:00"; o.winEnd = "17:30"; }
  if (!o.days.length) o.days = [1, 2, 3, 4, 5];
  return o;
}

const dayIndex = (meta, date) => meta.sendDays.length + (meta.sendDays.includes(date) ? 0 : 1);
const quotaFor = (s, k) => Math.min(s.max, s.start + s.step * (k - 1));

/** Prossimo giorno utile di invio dopo "date" */
function nextSendDay(s, date, wd) {
  for (let i = 0; i < 8; i++) { date = addDay(date); wd = (wd % 7) + 1; if (s.days.includes(wd)) return fmtDay(date, wd); }
  return "";
}

/** Data prevista di fine: simula le tappe future con quote, giorni e tetti per azienda */
export function simulate(meta, recs, now = new Date()) {
  const s = meta.s;
  let pend = recs.filter((r) => r[2] === ST.PEND).map((r) => r[1]);
  if (!pend.length) return null;
  const t = nowIn(s.tz, now);
  let date = t.date, wd = t.wd, k = dayIndex(meta, date), last = null;
  for (let i = 0; i < 1000 && pend.length; i++) {
    const today = i === 0;
    let took = 0;
    if (s.days.includes(wd) && !(today && t.hm >= s.winEnd)) {
      const quota = quotaFor(s, k) - (today ? meta.days[date] || 0 : 0);
      if (quota > 0) {
        const cnt = today && meta.dom && meta.dom.date === date ? Object.assign({}, meta.dom.c) : {};
        const keep = [];
        for (const d of pend) {
          if (took < quota && (isConsumer(d) || (cnt[d] || 0) < s.perDomain)) { cnt[d] = (cnt[d] || 0) + 1; took++; } else keep.push(d);
        }
        if (took) { pend = keep; last = date; }
      }
    }
    if (took || (today && meta.sendDays.includes(date))) k++;
    date = addDay(date); wd = (wd % 7) + 1;
  }
  return pend.length ? null : last;
}

function countRecs(recs) {
  const c = { total: recs.length, pending: 0, ok: 0, err: 0, unsub: 0, bounce: 0, norec: 0, reply: 0, auto: 0 };
  for (const r of recs) {
    c[["pending", "ok", "err", "unsub", "bounce", "norec"][r[2]]]++;
    if (r[5] === 1) c.reply++; else if (r[5] === 2) c.auto++;
  }
  return c;
}

/** Limite dei mancati recapiti (5%). Dopo una ripresa forzata si contano solo le email inviate da quel momento
 *  (bounceGuard.mode "nuovi"), oppure il limite è disattivato per la campagna (mode "off").
 *  Restituisce i numeri se il limite è superato, altrimenti null. */
function bounceOver(meta, recs) {
  const g = meta.bounceGuard || null;
  if (g && g.mode === "off") return null;
  const since = g && g.mode === "nuovi" ? String(g.at || "") : "";
  let sent = 0, bounce = 0;
  for (const r of recs) {
    if (r[2] !== ST.OK && r[2] !== ST.BOUNCE) continue;
    if (since && String(r[3] || "") < since) continue;   // inviata prima della ripresa
    sent++; if (r[2] === ST.BOUNCE) bounce++;
  }
  return sent >= BOUNCE_MIN_SENT && bounce / sent > BOUNCE_LIMIT ? { sent, bounce, since } : null;
}

/** Situazione di oggi (per l'elenco campagne e per il messaggio del lotto) */
function todayInfo(meta, now = new Date()) {
  const s = meta.s, t = nowIn(s.tz, now), k = dayIndex(meta, t.date);
  const quota = quotaFor(s, k), sent = meta.days[t.date] || 0;
  let open = false, message;
  if (!s.days.includes(t.wd)) message = `Oggi non è un giorno di invio. Prossima tappa: ${nextSendDay(s, t.date, t.wd)}`;
  else if (t.hm < s.winStart) message = `Tappa di oggi: ${quota} email dalle ${s.winStart}`;
  else if (t.hm >= s.winEnd) message = sent ? `Tappa di oggi conclusa: ${sent} inviate. Prossima tappa: ${nextSendDay(s, t.date, t.wd)}` : `Fascia oraria di oggi terminata. Prossima tappa: ${nextSendDay(s, t.date, t.wd)}`;
  else if (sent >= quota) message = `Tappa di oggi completata: ${sent} di ${quota}. Prossima tappa: ${nextSendDay(s, t.date, t.wd)}`;
  else { open = true; message = `Oggi: ${sent} di ${quota} inviate · fascia ${s.winStart}–${s.winEnd}`; }
  return { date: t.date, hm: t.hm, quota, sent, open, message, dayIndex: k };
}

function publicMeta(meta, user) {
  const { lock, scan, ...m } = meta;
  return Object.assign(m, {
    mine: meta.owner === user,
    running: !!(lock && lock.until > Date.now()),
    today: ["attiva"].includes(meta.status) ? todayInfo(meta) : null,
  });
}

// ---------- archivio ----------
const plans = () => getStore("tappe");
async function loadPlan(id) {
  if (!/^g[\w-]+$/.test(String(id || ""))) return {};
  const st = plans();
  const meta = await st.get(id, { type: "json" });
  if (!meta) return {};
  const recs = (await st.get(id + "/stato", { type: "json" })) || [];
  return { meta, recs };
}
async function savePlan(meta, recs) {
  meta.updatedAt = new Date().toISOString();
  if (recs) { meta.counts = countRecs(recs); await plans().setJSON(meta.id + "/stato", recs); }
  await plans().setJSON(meta.id, meta);
}
async function listPlanMetas() {
  const st = plans();
  const keys = (await listAll(st)).filter((k) => !k.includes("/"));
  return (await pool(keys, 20, (k) => st.get(k, { type: "json" }))).filter(Boolean);
}

// ---------- lettura dei messaggi di risposta e di mancato recapito (solo intestazioni) ----------
const GMAIL = "https://gmail.googleapis.com/gmail/v1/users/me/messages";
const HEADERS = ["From", "Subject", "X-Failed-Recipients", "Auto-Submitted", "X-Autoreply", "X-Autorespond", "Precedence", "Content-Type"];
async function gget(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (r.status === 401 || r.status === 403) {
    const j = await r.json().catch(() => ({}));
    const e = new Error((j.error && j.error.message) || "Permesso negato"); e.scope = true; throw e;
  }
  if (!r.ok) throw new Error("Gmail " + r.status);
  return r.json();
}
export function classify(h) {
  const v = (n) => (h[n.toLowerCase()] || "");
  const from = v("From"), subj = v("Subject"), xfr = v("X-Failed-Recipients"), ctype = v("Content-Type");
  if (/\b(delay|delayed|ritard|verzögert|retrasad|retard)/i.test(subj) && !xfr) return { kind: "ignore" };
  if (xfr || /report-type=["']?delivery-status/i.test(ctype) ||
      /mailer-daemon|postmaster|mail delivery (subsystem|system)/i.test(from) ||
      /^(undeliver|non recapitabil|unzustellbar|no se puede entregar|non remis|delivery status notification \(failure\)|mail delivery failed|returned mail|failure notice|message not delivered)/i.test(subj))
    return { kind: "bounce", failed: xfr.split(/[,\s]+/).map(normEmail).filter(isEmail) };
  const auto = v("Auto-Submitted");
  if ((auto && !/^no$/i.test(auto.trim())) || v("X-Autoreply") || v("X-Autorespond") || /auto[_-]?reply/i.test(v("Precedence")) ||
      /^(automatic reply|auto(matic)?[- ]?reply|autoreply|risposta automatica|out of office|fuori sede|assente|abwesenheit|respuesta autom|réponse automatique|autosvar)/i.test(subj))
    return { kind: "auto" };
  return { kind: "reply" };
}

/** Cerca nella posta in arrivo i messaggi collegati alle email della campagna */
async function scanInbox(token, meta, recs) {
  const byThread = new Map();
  recs.forEach((r) => { if (r[4]) byThread.set(r[4], r); });
  const res = { scanned: 0, bounces: [], replies: 0, autos: 0 };
  if (!byThread.size) return res;
  const scan = meta.scan || (meta.scan = { seen: [] });
  const seen = new Set(scan.seen);
  const since = Math.max(Date.parse(meta.createdAt), Date.now() - 14 * 864e5);
  const matched = [];
  let pageToken = "";
  for (let page = 0; page < 10; page++) {
    const u = new URL(GMAIL);
    u.searchParams.set("labelIds", "INBOX"); u.searchParams.set("maxResults", "100");
    if (pageToken) u.searchParams.set("pageToken", pageToken);
    const j = await gget(token, u.toString());
    const msgs = j.messages || [];
    res.scanned += msgs.length;
    for (const m of msgs) if (byThread.has(m.threadId) && !seen.has(m.id)) matched.push(m);
    if (!msgs.length || !j.nextPageToken) break;
    const last = await gget(token, `${GMAIL}/${msgs[msgs.length - 1].id}?format=minimal`);
    if (Number(last.internalDate) < since) break;
    pageToken = j.nextPageToken;
  }
  const emails = new Map(recs.map((r) => [r[0], r]));
  await pool(matched, 8, async (m) => {
    const q = HEADERS.map((h) => "metadataHeaders=" + encodeURIComponent(h)).join("&");
    const g = await gget(token, `${GMAIL}/${m.id}?format=metadata&${q}`);
    const h = {};
    ((g.payload && g.payload.headers) || []).forEach((x) => (h[x.name.toLowerCase()] = x.value));
    const c = classify(h), rec = byThread.get(m.threadId);
    seen.add(m.id);
    if (c.kind === "bounce") {
      const hit = c.failed.map((e) => emails.get(e)).filter(Boolean);
      for (const r of hit.length ? hit : [rec]) {
        if (r[2] === ST.OK) { r[2] = ST.BOUNCE; r[6] = String(h.subject || "").slice(0, 160); res.bounces.push(r[0]); }
      }
    } else if (c.kind === "auto") { if (!rec[5]) { rec[5] = 2; res.autos++; } }
    else if (c.kind === "reply") { if (rec[5] !== 1) { rec[5] = 1; res.replies++; } }
  });
  scan.seen = [...seen].slice(-3000);
  scan.at = new Date().toISOString();
  return res;
}

// =====================================================================
export default async (req) => {
  const user = await verifyGoogle(req);
  if (!user) return json(401, { error: "Accesso non autorizzato" });
  const url = new URL(req.url);
  const base = process.env.URL || url.origin;

  // ---- contenuto e campi: caricati e restituiti così come sono ----
  const raw = url.searchParams.get("azione");
  if (raw === "contenuto" || raw === "campi") {
    const { meta } = await loadPlan(url.searchParams.get("id"));
    if (!meta) return json(404, { error: "Campagna non trovata" });
    if (meta.owner !== user) return json(403, { error: "Solo chi ha creato la campagna può usarne il contenuto" });
    const store = rawStore("tappe"), key = meta.id + "/" + raw;
    if (req.method === "PUT") {
      if (meta.status !== "preparazione") return json(409, { error: "La campagna è già stata avviata: il contenuto non si può più cambiare" });
      const len = Number(req.headers.get("content-length") || 0);
      if (!len) return json(411, { error: "Dimensione mancante" });
      if (len > (raw === "contenuto" ? MAX_CONTENT : MAX_FIELDS)) return json(413, { error: "Contenuto troppo grande" });
      await store.put(key, req.body, { httpMetadata: { contentType: "application/json" } });
      return json(200, { ok: true });
    }
    const obj = await store.get(key);
    if (!obj) return json(404, { error: "Contenuto non trovato" });
    return new Response(obj.body, { headers: { "Content-Type": "application/json", "Cache-Control": "no-store" } });
  }

  if (req.method !== "POST") return json(405, { error: "Metodo non consentito" });
  let body;
  try { body = await req.json(); } catch { return json(400, { error: "Richiesta non valida" }); }
  const token = (req.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");

  try {
    switch (body.action) {
      // Crea la campagna con l'elenco completo (già ordinato e filtrato dal browser)
      case "crea": {
        const c = body.campaign || {};
        const emails = [...new Set((body.recipients || []).map(normEmail))].filter(isEmail);
        if (!emails.length) return json(400, { error: "Nessun destinatario valido" });
        if (emails.length > MAX_RECIPIENTS) return json(400, { error: `Massimo ${MAX_RECIPIENTS} destinatari per campagna` });
        const s = cleanSettings(body.settings);
        const id = newId(), campaignId = id.slice(1);
        const subject = String(c.subject || "").slice(0, 300), sender = normEmail(c.sender || user);
        // compare anche nello storico delle campagne del registro condiviso
        await getStore("campagne").setJSON(campaignId, { id: campaignId, subject, sender, user, mode: c.mode || "", source: c.source || "",
          planned: emails.length, ok: 0, ko: 0, graduale: id, createdAt: new Date().toISOString() });
        await getStore("campagne").setJSON(campaignId + "/destinatari", []);
        const recs = emails.map((e) => [e, domainOf(e), ST.PEND, "", "", 0, ""]);
        const meta = { id, campaignId, owner: user, subject, sender, fromName: String(c.fromName || "").slice(0, 120),
          status: "preparazione", reason: "", s, sendDays: [], days: {}, dom: null, createdAt: new Date().toISOString() };
        meta.eta = simulate(meta, recs);
        await savePlan(meta, recs);
        return json(200, { plan: publicMeta(meta, user) });
      }

      case "elenco": {
        const metas = await listPlanMetas();
        metas.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
        return json(200, { plans: metas.slice(0, 100).map((m) => publicMeta(m, user)) });
      }

      // Pausa, ripresa, annullamento, avvio dopo il caricamento del contenuto
      case "stato": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Solo chi ha creato la campagna può modificarla" });
        const to = body.status;
        if (!["attiva", "in pausa", "annullata"].includes(to)) return json(400, { error: "Stato non valido" });
        if (["completata", "annullata"].includes(meta.status)) return json(409, { error: "La campagna è chiusa" });
        if (to === "attiva" && meta.status === "preparazione") {
          const store = rawStore("tappe");
          if (!(await store.get(meta.id + "/contenuto")) || !(await store.get(meta.id + "/campi"))) return json(409, { error: "Contenuto non caricato: ricrea la campagna" });
        }
        // Ripresa di una campagna sospesa per mancati recapiti: i mancati recapiti già registrati non la bloccano più.
        // "nuovi" = il limite del 5% vale solo per le email inviate da adesso; "off" = nessun limite per questa campagna.
        if (to === "attiva" && meta.status === "sospesa") {
          meta.bounceGuard = { mode: body.bounceGuard === "off" ? "off" : "nuovi", at: new Date().toISOString(), by: user };
        }
        meta.status = to; meta.reason = to === "in pausa" ? String(body.reason || "Messa in pausa da " + user).slice(0, 300) : "";
        if (to !== "attiva") meta.lock = null;
        if (to === "attiva") meta.eta = simulate(meta, recs);
        await savePlan(meta);
        return json(200, { plan: publicMeta(meta, user) });
      }

      // Riattiva (o disattiva) il limite dei mancati recapiti; riattivato, vale per le email inviate da adesso
      case "limite": {
        const { meta } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Solo chi ha creato la campagna può modificarla" });
        if (["completata", "annullata"].includes(meta.status)) return json(409, { error: "La campagna è chiusa" });
        meta.bounceGuard = { mode: body.mode === "off" ? "off" : "nuovi", at: new Date().toISOString(), by: user };
        await savePlan(meta);
        return json(200, { plan: publicMeta(meta, user) });
      }

      // Rimette in coda gli invii falliti per errore temporaneo
      case "riprova": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Solo chi ha creato la campagna può modificarla" });
        let n = 0;
        recs.forEach((r) => { if (r[2] === ST.ERR) { r[2] = ST.PEND; r[6] = ""; n++; } });
        if (n && meta.status === "completata") meta.status = "attiva";
        meta.eta = simulate(meta, recs);
        await savePlan(meta, recs);
        return json(200, { requeued: n, plan: publicMeta(meta, user) });
      }

      // Il lotto di oggi: chi ricevere adesso, nel rispetto di quote, fascia oraria e tetti
      case "lotto": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Le email partono dalla casella di chi ha creato la campagna: solo " + meta.owner + " può inviarla" });
        const empty = (message) => json(200, { batch: [], message, plan: publicMeta(meta, user) });
        const STATO = { preparazione: "non è ancora stata avviata", "in pausa": "è in pausa", sospesa: "è sospesa", annullata: "è stata annullata", completata: "è completata" };
        if (meta.status !== "attiva") return empty(`La campagna ${STATO[meta.status] || meta.status}${meta.reason ? ": " + meta.reason : ""}`);
        if (meta.lock && meta.lock.until > Date.now() && meta.lock.session !== body.session)
          return json(409, { error: "Questa campagna è già in invio in un'altra finestra o su un altro computer. Attendi che finisca (o chiudi l'altra finestra: si sblocca da sola entro 20 minuti)." });
        const s = meta.s, info = todayInfo(meta);
        if (!info.open) return empty(info.message);

        // aggiorna le esclusioni: disiscritti e non recapitabili nel frattempo
        const unsubKeys = new Set(await listAll(getStore("disiscritti"))), norecKeys = new Set(await listAll(getStore("nonrecapitabili")));
        for (const r of recs) if (r[2] === ST.PEND) {
          const k = keyOf(r[0]);
          if (unsubKeys.has(k)) r[2] = ST.UNSUB; else if (norecKeys.has(k)) r[2] = ST.NOREC;
        }

        // quota di oggi e tetto della casella (tutte le campagne graduali dello stesso mittente)
        let remaining = info.quota - info.sent;
        const cap = Number(process.env.MAILBOX_DAILY_CAP) || 1500;
        const others = (await listPlanMetas()).filter((m) => m.owner === meta.owner && m.id !== meta.id);
        const mailboxToday = info.sent + others.reduce((a, m) => a + ((m.days || {})[info.date] || 0), 0);
        remaining = Math.min(remaining, cap - mailboxToday);
        if (remaining <= 0) {
          await savePlan(meta, recs);
          return empty(mailboxToday >= cap ? `Raggiunto il tetto giornaliero della casella (${cap} email con tutte le campagne). Si riprende domani.` : info.message);
        }

        // selezione: tetto per azienda, poi alternanza dei domini
        const cnt = meta.dom && meta.dom.date === info.date ? Object.assign({}, meta.dom.c) : {};
        const cand = [];
        for (const r of recs) {
          if (r[2] !== ST.PEND) continue;
          if (!isConsumer(r[1]) && (cnt[r[1]] || 0) >= s.perDomain) continue;
          cnt[r[1]] = (cnt[r[1]] || 0) + 1; cand.push(r);
          if (cand.length >= remaining) break;
        }
        const counts = countRecs(recs);
        if (!cand.length) {
          if (!counts.pending) meta.status = "completata";
          await savePlan(meta, recs);
          return empty(counts.pending ? `Per oggi restano solo aziende che hanno già raggiunto il limite di ${s.perDomain} email al giorno. Prossima tappa: ${nextSendDay(s, info.date, nowIn(s.tz).wd)}` : "Tutti i destinatari sono stati raggiunti: campagna completata");
        }
        const batch = interleave(cand, (r) => r[1]).map((r) => ({ email: r[0], unsubUrl: unsubscribeUrl(base, r[0], meta.campaignId) }));
        const session = body.session || Math.random().toString(36).slice(2);
        meta.lock = { session, until: Date.now() + LOCK_MS };
        meta.eta = simulate(meta, recs);
        await savePlan(meta, recs);
        return json(200, { batch, session, quota: info.quota, sentToday: info.sent, dayIndex: info.dayIndex, winEnd: s.winEnd, tz: s.tz,
          gap: [s.gapMin, s.gapMax], pauseEvery: s.pauseEvery, plan: publicMeta(meta, user) });
      }

      // Esiti di un blocco di invii
      case "esito": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Non autorizzato" });
        if (meta.lock && meta.lock.until > Date.now() && meta.lock.session !== body.session) return json(409, { error: "Sessione di invio non valida" });
        const idx = new Map(recs.map((r) => [r[0], r]));
        const t = nowIn(meta.s.tz), forLog = [];
        if (!meta.dom || meta.dom.date !== t.date) meta.dom = { date: t.date, c: {} };
        let ok = 0;
        for (const x of body.results || []) {
          const r = idx.get(normEmail(x.email));
          if (!r || r[2] === ST.OK || r[2] === ST.BOUNCE) continue;   // già registrato: niente doppioni
          if (x.status === "ok") {
            r[2] = ST.OK; r[3] = x.time || new Date().toISOString(); r[4] = String(x.threadId || ""); r[6] = "";
            meta.dom.c[r[1]] = (meta.dom.c[r[1]] || 0) + 1; ok++;
          } else { r[2] = ST.ERR; r[3] = x.time || ""; r[6] = String(x.error || "errore").slice(0, 200); }
          forLog.push({ email: r[0], status: x.status === "ok" ? "ok" : "error", id: x.id || "", time: x.time, error: x.error || "" });
        }
        if (ok) {
          meta.days[t.date] = (meta.days[t.date] || 0) + ok;
          if (!meta.sendDays.includes(t.date)) meta.sendDays.push(t.date);
        }
        meta.lock = body.final ? null : { session: body.session, until: Date.now() + LOCK_MS };
        const counts = countRecs(recs);
        if (!counts.pending && meta.status === "attiva") meta.status = "completata";
        if (body.final) meta.eta = simulate(meta, recs);
        await savePlan(meta, recs);
        if (forLog.length) await logSends(meta.campaignId, forLog);
        return json(200, { ok: true, plan: publicMeta(meta, user) });
      }

      // Controllo di risposte e mancati recapiti nella casella di chi invia
      case "recapiti": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        if (meta.owner !== user) return json(403, { error: "Solo chi invia la campagna può controllarne la casella" });
        let res;
        try { res = await scanInbox(token, meta, recs); }
        catch (e) {
          if (e.scope) return json(200, { needScope: true, message: "Per riconoscere mancati recapiti e risposte serve il permesso Gmail «metadata»: esci e rientra con Google per concederlo." });
          throw e;
        }
        await pool(res.bounces, 5, (e) => addBounced(e, meta.campaignId, user, "mancato recapito"));
        let suspended = false;
        const over = meta.status === "attiva" ? bounceOver(meta, recs) : null;
        if (over) {
          meta.status = "sospesa"; meta.lock = null; suspended = true;
          const d = over.since ? nowIn(meta.s.tz, new Date(over.since)).date : "";
          meta.reason = `mancati recapiti: ${String(Math.round((over.bounce / over.sent) * 1000) / 10).replace(".", ",")}% (${over.bounce} su ${over.sent}` +
            (d ? ` tra le email inviate dalla ripresa del ${Number(d.slice(8, 10))}/${Number(d.slice(5, 7))}` : "") + `): controlla la qualità della lista prima di riprendere`;
        }
        await savePlan(meta, recs);
        return json(200, Object.assign(res, { suspended, plan: publicMeta(meta, user) }));
      }

      // Esito per destinatario (per il report CSV)
      case "report": {
        const { meta, recs } = await loadPlan(body.id);
        if (!meta) return json(404, { error: "Campagna non trovata" });
        return json(200, { plan: publicMeta(meta, user), recipients: recs.map((r) => ({ e: r[0], st: r[2], t: r[3], r: r[5], err: r[6] })) });
      }

      default:
        return json(400, { error: "Azione non valida" });
    }
  } catch (e) {
    console.error(e);
    return json(500, { error: e.message });
  }
};
