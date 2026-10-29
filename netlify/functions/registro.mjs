// Funzione Netlify: registro centrale condiviso di invii e disiscrizioni (Netlify Blobs).
// Archivi: "disiscritti" (chiave = email), "invii" (storico per email), "campagne" (metadati + destinatari).
import { json, verifyGoogle, getStore, keyOf, emailOf, normEmail, listAll, pool, addUnsubscribe, unsubscribeUrl } from "../lib/common.mjs";

const MAX_HISTORY = 100;
const newId = () => new Date().toISOString().replace(/[-:T]/g, "").slice(0, 12) + "-" + Math.random().toString(36).slice(2, 7);

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "Metodo non consentito" });
  const user = await verifyGoogle(req);
  if (!user) return json(401, { error: "Accesso non autorizzato" });
  let body;
  try { body = await req.json(); } catch { return json(400, { error: "Richiesta non valida" }); }
  const base = process.env.URL || new URL(req.url).origin;
  const unsubStore = getStore("disiscritti"), sendStore = getStore("invii"), campStore = getStore("campagne");

  try {
    switch (body.action) {
      // Stato di una lista di destinatari: disiscritti e storico invii
      case "check": {
        const emails = [...new Set((body.emails || []).map(normEmail))].slice(0, 5000);
        const unsubKeys = new Set(await listAll(unsubStore));
        const unsubscribed = {}, history = {};
        await pool(emails, 25, async (e) => {
          if (unsubKeys.has(keyOf(e))) unsubscribed[e] = await unsubStore.get(keyOf(e), { type: "json" });
          const h = await sendStore.get(keyOf(e), { type: "json" });
          if (h && h.sends && h.sends.length) {
            const last = h.sends[h.sends.length - 1];
            history[e] = { count: h.sends.length, last, subjects: h.sends.map((s) => s.s) };
          }
        });
        return json(200, { unsubscribed, history });
      }

      // Apre una campagna e restituisce i link di disiscrizione personali
      case "prepare": {
        const c = body.campaign || {};
        const id = newId();
        const meta = { id, subject: String(c.subject || "").slice(0, 300), sender: normEmail(c.sender || user), user,
          mode: c.mode || "", source: c.source || "", planned: Number(c.total) || 0, ok: 0, ko: 0, createdAt: new Date().toISOString() };
        await campStore.setJSON(id, meta);
        await campStore.setJSON(id + "/destinatari", []);
        const links = {};
        (body.emails || []).forEach((e) => (links[normEmail(e)] = unsubscribeUrl(base, e, id)));
        return json(200, { campaignId: id, links });
      }

      // Registra l'esito di un blocco di invii
      case "log": {
        const id = body.campaignId, results = body.results || [];
        const meta = await campStore.get(id, { type: "json" });
        if (!meta) return json(404, { error: "Campagna non trovata" });
        await pool(results.filter((r) => r.status === "ok"), 20, async (r) => {
          const k = keyOf(r.email), h = (await sendStore.get(k, { type: "json" })) || { email: normEmail(r.email), sends: [] };
          h.sends.push({ c: id, s: meta.subject, f: meta.sender, t: r.time || new Date().toISOString() });
          h.sends = h.sends.slice(-MAX_HISTORY);
          await sendStore.setJSON(k, h);
        });
        const rec = (await campStore.get(id + "/destinatari", { type: "json" })) || [];
        results.forEach((r) => rec.push({ e: normEmail(r.email), st: r.status, t: r.time, err: r.error || "", id: r.id || "" }));
        meta.ok += results.filter((r) => r.status === "ok").length;
        meta.ko += results.filter((r) => r.status !== "ok").length;
        meta.updatedAt = new Date().toISOString();
        await campStore.setJSON(id + "/destinatari", rec);
        await campStore.setJSON(id, meta);
        return json(200, { ok: true, meta });
      }

      case "campaigns": {
        const keys = (await listAll(campStore)).filter((k) => !k.includes("/"));
        const metas = (await pool(keys, 25, (k) => campStore.get(k, { type: "json" }))).filter(Boolean);
        metas.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
        return json(200, { campaigns: metas.slice(0, 200) });
      }

      case "campaign": {
        const meta = await campStore.get(body.id, { type: "json" });
        if (!meta) return json(404, { error: "Campagna non trovata" });
        return json(200, { meta, recipients: (await campStore.get(body.id + "/destinatari", { type: "json" })) || [] });
      }

      case "unsubscribes": {
        const keys = await listAll(unsubStore);
        const recs = (await pool(keys, 25, (k) => unsubStore.get(k, { type: "json" })))
          .map((r, i) => r || { email: emailOf(keys[i]) });
        recs.sort((a, b) => String(b.date).localeCompare(String(a.date)));
        return json(200, { unsubscribes: recs });
      }

      case "unsubscribe-add": {
        const emails = (body.emails || []).map(normEmail).filter((e) => /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(e));
        await pool(emails, 10, (e) => addUnsubscribe(e, body.source || "manuale", user, ""));
        return json(200, { added: emails.length });
      }

      case "unsubscribe-remove": {
        await unsubStore.delete(keyOf(body.email));
        console.log(`[ia-mailer] ${user} ha riattivato ${normEmail(body.email)}`);
        return json(200, { ok: true });
      }

      default:
        return json(400, { error: "Azione non valida" });
    }
  } catch (e) {
    console.error(e);
    return json(500, { error: e.message });
  }
};
