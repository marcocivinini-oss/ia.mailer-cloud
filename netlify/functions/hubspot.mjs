// Funzione Netlify: legge liste e contatti HubSpot per IA Mailer.
// Il token HubSpot resta sul server (variabile HUBSPOT_TOKEN) e non arriva mai al browser.
// Ogni chiamata deve presentare un token Google valido di un utente del dominio autorizzato.

import { json, verifyGoogle } from "../lib/common.mjs";

const HS = "https://api.hubapi.com";
const PROPS = ["email", "firstname", "lastname", "company", "country", "jobtitle", "hs_email_optout"];
const MAX_CONTACTS = 5000;

async function hs(path, opts = {}) {
  const r = await fetch(HS + path, {
    ...opts,
    headers: { Authorization: "Bearer " + process.env.HUBSPOT_TOKEN, "Content-Type": "application/json", ...(opts.headers || {}) }
  });
  if (r.status === 429) { await new Promise((s) => setTimeout(s, 1200)); return hs(path, opts); }
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.message || "HubSpot HTTP " + r.status);
  return j;
}

export default async (req) => {
  if (!process.env.HUBSPOT_TOKEN) return json(500, { error: "Integrazione HubSpot non configurata (manca HUBSPOT_TOKEN)" });
  const user = await verifyGoogle(req);
  if (!user) return json(401, { error: "Accesso non autorizzato" });

  const url = new URL(req.url);
  const action = url.searchParams.get("action");
  try {
    if (action === "lists") {
      const j = await hs("/crm/v3/lists/search", {
        method: "POST",
        body: JSON.stringify({ query: url.searchParams.get("q") || "", count: 50, offset: 0, additionalProperties: ["hs_list_size"] })
      });
      const lists = (j.lists || [])
        .filter((l) => !l.objectTypeId || l.objectTypeId === "0-1")
        .map((l) => ({ listId: String(l.listId), name: l.name, size: l.additionalProperties ? Number(l.additionalProperties.hs_list_size) : null }));
      return json(200, { lists });
    }

    if (action === "members") {
      const listId = url.searchParams.get("listId");
      if (!/^\d+$/.test(listId || "")) return json(400, { error: "listId non valido" });
      const ids = [];
      let after;
      do {
        const q = new URLSearchParams({ limit: "250" });
        if (after) q.set("after", after);
        const j = await hs(`/crm/v3/lists/${listId}/memberships?` + q);
        (j.results || []).forEach((m) => ids.push(String(m.recordId)));
        after = j.paging && j.paging.next && j.paging.next.after;
      } while (after && ids.length < MAX_CONTACTS);

      const contacts = [];
      for (let i = 0; i < ids.length; i += 100) {
        const j = await hs("/crm/v3/objects/contacts/batch/read", {
          method: "POST",
          body: JSON.stringify({ properties: PROPS, inputs: ids.slice(i, i + 100).map((id) => ({ id })) })
        });
        (j.results || []).forEach((c) => {
          const p = c.properties || {};
          if (!p.email) return;
          contacts.push({
            email: p.email, firstname: p.firstname || "", lastname: p.lastname || "", company: p.company || "",
            country: p.country || "", jobtitle: p.jobtitle || "", optout: String(p.hs_email_optout) === "true"
          });
        });
      }
      console.log(`[ia-mailer] ${user} ha caricato la lista ${listId} (${contacts.length} contatti)`);
      return json(200, { contacts, truncated: ids.length >= MAX_CONTACTS });
    }
    return json(400, { error: "Azione non valida" });
  } catch (e) {
    return json(502, { error: e.message });
  }
};
