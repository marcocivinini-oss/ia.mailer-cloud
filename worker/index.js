// IA Suite — server su Cloudflare Workers.
// - L'accesso è protetto da Cloudflare Access: ogni richiesta porta l'email
//   certificata di chi ha fatto login (header Cf-Access-Authenticated-User-Email).
// - I dati stanno in uno spazio R2 dedicato alla società (binding DATA).
// - La chiave API di Anthropic è un "Secret" del Worker e non arriva mai al browser.

const PROMPTS = {
  contract: `Sei un assistente che estrae dati da un contratto di consulenza. Rispondi SOLO con un oggetto JSON valido, senza testo o backtick, con i campi: client (ragione sociale del cliente), currency ("EUR" o "USD"), startDate (YYYY-MM-DD o null), endDate (YYYY-MM-DD o null), paymentTermsDays (giorni per il pagamento come numero, null se non chiaro), vatExempt (true se il cliente è in esenzione IVA — es. cliente estero, non imponibile art. 7-ter/art. 8, reverse charge; altrimenti false), vatRate (aliquota IVA numerica quando indicata, null se non specificata o se vatExempt=true), notes (nota breve max 160 caratteri), billingItems (array delle componenti di compenso). Ogni elemento di billingItems ha: type — usa "recurring" per retainer o canoni periodici ricorrenti, "one_shot" per un compenso UNICO da fatturare una sola volta (es. fee fissa forfettaria per un progetto o consulenza spot), "trigger" per importi attivati da un evento specifico (firma, kickoff, milestone, closing, success fee) — label (descrizione breve della voce), amount (importo netto come numero), triggerLabel (solo per type "trigger": l'evento che lo attiva; null altrimenti), date (solo per type "one_shot": data prevista di fatturazione YYYY-MM-DD; null altrimenti). Se il contratto prevede solo un retainer mensile, restituisci un unico elemento con type "recurring". Usa null dove un dato non è presente.`,
  expense: `Estrai il totale da un foglio spese di viaggio. Rispondi SOLO con JSON valido: description (breve, es. "Trasferta Milano-Roma"), amount (totale come numero), currency ("EUR"/"USD"), date (YYYY-MM-DD o null).`,
  invoice: `Estrai i dati da una fattura fornitore. Rispondi SOLO con un oggetto JSON valido, senza testo o backtick, con i campi:
- supplier: ragione sociale del fornitore
- supplierCountry: codice paese ISO-2 del fornitore (es. "IT", "US", "GB", "DE", "CH", "FR"), null se non chiaro
- number: numero fattura (stringa)
- date: data emissione YYYY-MM-DD, null se non chiara
- dueDate: data scadenza pagamento YYYY-MM-DD, null se non presente
- amount: importo DA PAGARE al fornitore come numero (SEMPRE positivo, anche per note di credito — usa il campo isCreditNote per distinguerle). ATTENZIONE alle parcelle e agli avvisi di parcella di professionisti (commercialisti, avvocati, consulenti) con RITENUTA D'ACCONTO: in quei documenti usa il NETTO A PAGARE (voci "TOTALE NETTO", "Netto a pagare", "Netto da pagare"), cioè il totale lordo meno la ritenuta — NON il "Totale documento" e NON il "Totale lordo". Per tutti gli altri documenti usa il totale documento IVA inclusa
- withholding: importo della ritenuta d'acconto come numero (es. 511.00), null se il documento non ha ritenuta
- netToPay: il valore della riga "TOTALE NETTO" / "Netto a pagare" / "Netto da pagare" se presente nel documento, altrimenti null
- grossTotal: il valore della riga "TOTALE LORDO" (totale comprensivo di IVA, prima della ritenuta) se presente, altrimenti null
- taxableAmount: costo al netto dell'IVA come numero, cioè la somma degli imponibili (inclusa la cassa previdenza e le voci escluse da IVA come le anticipazioni art. 15); null se non chiaro
- currency: codice valuta ISO-4217 (es. "EUR", "USD", "GBP", "CHF", "JPY", "SEK", "NOK", "DKK", "PLN", "CAD", "AUD"). Cerca il simbolo (€ $ £ ¥ CHF Fr. kr zł) o codice esplicito nel documento — NON dare per scontato "EUR"
- isCreditNote: true se il documento è una nota di credito / credit note / Gutschrift / NC / TD04, false se è una fattura normale
- vatExempt: true se il documento non presenta IVA/VAT/tax (tipico per fatture UE B2B in reverse charge, fatture extra-UE, fatture US/UK verso l'estero, o italiane non imponibili art. 7-ter/art. 8/art. 41); false se l'IVA è addebitata
- vatRate: aliquota IVA numerica se addebitata (es. 22 per Italia, 19 per Germania, 20 per UK, 7.7 per CH), null se vatExempt=true o non trovata
- category: categoria di spesa più adatta scelta tra "Consulenze", "Viaggi", "Uffici", "Utenze", "Software", "Compensi amministratori", "Altro" (usa "Altro" se non riesci a determinarla)
- notes: eventuale nota breve (max 100 caratteri) sul motivo dell'esenzione IVA o sulla natura del documento`,
};

PROMPTS.issued_invoice = `Extract the data of an invoice ISSUED by IA International Advisors to a client. Reply ONLY with a valid JSON object, no text or backticks, with the fields:
- client: the client company name (the recipient, NOT IA International Advisors)
- number: invoice number (string)
- date: issue date YYYY-MM-DD, null if unclear
- dueDate: payment due date YYYY-MM-DD, null if absent
- amount: document TOTAL as a number, always positive
- currency: ISO-4217 code ("USD", "EUR", …); look for the symbol or code, do not assume
- isCreditNote: true if the document is a credit note, otherwise false
- notes: short note (max 100 characters), null if none`;

PROMPTS.bank_statement = `You are reading a US bank account statement (e.g. Chase). List EVERY transaction in the statement. Reply ONLY with a valid JSON object, no text or backticks:
{"currency":"USD","transactions":[{"date":"YYYY-MM-DD","description":"…","amount":123.45}]}
Rules:
- Use the statement period to determine the full year of each date.
- amount is POSITIVE for money coming in (deposits, additions, credits) and NEGATIVE for money going out (card purchases, ATM, electronic withdrawals, checks, fees).
- description: a short readable text keeping the counterparty name (e.g. "Cep USA Inc – Invpymnt", "Intuit QuickBooks", "Same-Day ACH payment to Yazcc").
- Do NOT include summary lines, totals, daily ending balances or beginning/ending balance.`;

const CATEGORIES_EN = '"Consulting", "Travel", "Office", "Utilities", "Software", "Directors\' fees", "Other"';
const MAX_TOKENS = { bank_statement: 8000 };

const KEY_RE = /^[a-z0-9_]{1,80}$/;

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    // Installazione sul telefono: nome e icona dipendono dall'installazione (S.r.l. o Inc.)
    if (url.pathname.startsWith("/pwa/")) {
      const ent = env.ENTITY === "inc" ? "inc" : "srl";
      if (url.pathname === "/pwa/manifest.webmanifest") {
        const name = ent === "inc" ? "IA Suite Inc." : "IA Suite S.r.l.";
        return new Response(JSON.stringify({
          name, short_name: ent === "inc" ? "IA Suite Inc" : "IA Suite", start_url: "/", scope: "/", display: "standalone",
          background_color: "#131D4B", theme_color: "#131D4B",
          icons: [
            { src: "/pwa/icon-192.png", sizes: "192x192", type: "image/png" },
            { src: "/pwa/icon-512.png", sizes: "512x512", type: "image/png" },
            { src: "/pwa/icon-maskable-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
          ],
        }), { headers: { "content-type": "application/manifest+json", "cache-control": "public, max-age=3600" } });
      }
      const m = url.pathname.match(/^\/pwa\/(apple-touch-icon|icon-192|icon-512|icon-maskable-512)\.png$/);
      if (m) return env.ASSETS.fetch(new Request(new URL(`/pwa/${ent}-${m[1]}.png`, url), request));
      return new Response("Not found", { status: 404 });
    }
    if (!url.pathname.startsWith("/api/")) return env.ASSETS.fetch(request);

    const email = (request.headers.get("Cf-Access-Authenticated-User-Email") || "").toLowerCase();
    if (!email) return json(401, { error: "Accesso non autenticato" });

    try {
      if (url.pathname === "/api/me") return json(200, { email, entity: env.ENTITY || "srl" });

      if (url.pathname === "/api/data") {
        const key = url.searchParams.get("key") || "";
        if (!KEY_RE.test(key)) return json(400, { error: "Parametro 'key' non valido" });
        const objKey = key + ".json";
        if (request.method === "GET") {
          const obj = await env.DATA.get(objKey);
          return json(200, { value: obj ? await obj.json() : null });
        }
        if (request.method === "PUT") {
          const body = await request.json().catch(() => ({}));
          await env.DATA.put(objKey, JSON.stringify(body.value ?? null), {
            httpMetadata: { contentType: "application/json" },
            customMetadata: { updatedBy: email },
          });
          return json(200, { ok: true });
        }
        if (request.method === "DELETE") {
          await env.DATA.delete(objKey);
          return json(200, { ok: true });
        }
        return json(405, { error: "Method not allowed" });
      }

      // Registro fatture IA Inc.: numerazione progressiva assegnata dal server + archivio PDF
      if (url.pathname.startsWith("/api/invoices")) {
        if (env.ENTITY !== "inc") return json(404, { error: "Invoice register available only in IA Suite Inc." });
        return await invoices(request, env, url, email);
      }

      if (url.pathname === "/api/extract" && request.method === "POST") {
        if (!env.ANTHROPIC_API_KEY) return json(500, { error: "ANTHROPIC_API_KEY non configurata" });
        const body = await request.json().catch(() => null);
        if (!body) return json(400, { error: "JSON non valido" });
        const { kind, pdfB64, text } = body;
        let prompt = PROMPTS[kind];
        if (prompt && kind === "invoice" && env.ENTITY === "inc") {
          prompt = prompt.replace('"Consulenze", "Viaggi", "Uffici", "Utenze", "Software", "Compensi amministratori", "Altro" (usa "Altro" se non riesci a determinarla)', CATEGORIES_EN + ' (use "Other" if unsure)');
        }
        if (!prompt) return json(400, { error: "kind non valido" });
        const content = pdfB64
          ? [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: pdfB64 } }, { type: "text", text: prompt }]
          : [{ type: "text", text: prompt + "\n\n--- TESTO DOCUMENTO ---\n" + String(text || "").slice(0, 12000) }];
        const r = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-api-key": env.ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
          },
          body: JSON.stringify({ model: env.ANTHROPIC_MODEL || "claude-sonnet-5", max_tokens: MAX_TOKENS[kind] || 1000, messages: [{ role: "user", content }] }),
        });
        const data = await r.json().catch(() => ({}));
        // Non inoltriamo lo status di Anthropic: un suo 401 non deve sembrare una sessione scaduta.
        if (!r.ok) return json(502, { error: data?.error?.message || `Errore API (${r.status})` });
        const out = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
        return json(200, { text: out });
      }

      return json(404, { error: "Richiesta non riconosciuta" });
    } catch (e) {
      return json(500, { error: String(e.message || e) });
    }
  },
};

// ============ REGISTRO FATTURE IA INC. ============
// Un unico oggetto R2 contiene il registro e i contatori per anno. Ogni modifica
// passa da un aggiornamento condizionato (etag): se due persone emettono nello
// stesso istante, la seconda riprova sul registro aggiornato, quindi lo stesso
// numero non viene mai assegnato due volte. I PDF stanno sotto invoices/AAAA/.
const REG_KEY = "inc_invoice_register.json";
const MAX_PDF = 8 * 1024 * 1024;

class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }

const fmtNumber = (seq, year) => String(seq).padStart(3, "0") + "/" + year;
const yearOf = (d) => { const m = /^(\d{4})-\d{2}-\d{2}$/.exec(String(d || "")); return m ? Number(m[1]) : null; };
const parseNumber = (n) => { const m = /^\s*0*(\d{1,6})\s*[\/\-]\s*(\d{4})\s*$/.exec(String(n || "")); return m ? { seq: Number(m[1]), year: Number(m[2]) } : null; };
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

function totalsOf(inv) {
  const services = (inv.services || []).reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const admin = services * (Number(inv.adminCostPercent) || 0) / 100;
  const subtotal = services + admin;
  const tax = subtotal * (Number(inv.taxPercent) || 0) / 100;
  return r2(subtotal + tax + (Number(inv.otherAmount) || 0));
}

// Campi che chi emette può scrivere; numero, stato e storico li gestisce solo il server.
function cleanInvoice(src) {
  const s = (v, n) => String(v ?? "").slice(0, n);
  const services = (Array.isArray(src.services) ? src.services : [])
    .map(x => ({ description: s(x.description, 600), amount: r2(x.amount) }))
    .filter(x => x.description.trim() || x.amount)
    .slice(0, 60);
  const inv = {
    client: s(src.client, 200).trim(),
    clientAddress: s(src.clientAddress, 600),
    date: s(src.date, 10),
    services,
    adminCostPercent: r2(src.adminCostPercent),
    taxPercent: r2(src.taxPercent),
    otherAmount: r2(src.otherAmount),
    terms: s(src.terms || "Due upon receipt", 200),
    notes: s(src.notes, 1000),
    currency: "USD",
  };
  if (!inv.client) throw new HttpError(400, "Client name is missing");
  if (!yearOf(inv.date)) throw new HttpError(400, "Invalid date");
  if (!services.some(x => x.description.trim() && x.amount)) throw new HttpError(400, "Add at least one line with an amount");
  inv.total = totalsOf(inv);
  return inv;
}

function nextSeq(reg, year) {
  const used = reg.invoices.filter(i => i.year === year && i.seq).map(i => i.seq);
  return Math.max(Number(reg.next[year]) || 1, used.length ? Math.max(...used) + 1 : 1);
}

async function readRegister(env) {
  const obj = await env.DATA.get(REG_KEY);
  const reg = obj ? await obj.json() : null;
  return { obj, reg: reg && Array.isArray(reg.invoices) ? { next: reg.next || {}, invoices: reg.invoices } : { next: {}, invoices: [] } };
}

async function mutateRegister(env, fn) {
  for (let attempt = 0; attempt < 8; attempt++) {
    const { obj, reg } = await readRegister(env);
    const out = fn(reg);
    const opts = { httpMetadata: { contentType: "application/json" } };
    if (obj) opts.onlyIf = { etagMatches: obj.etag };
    const res = await env.DATA.put(REG_KEY, JSON.stringify(reg), opts);
    if (res) return out;
    await new Promise(r => setTimeout(r, 40 + Math.random() * 120 * (attempt + 1)));
  }
  throw new HttpError(409, "The register is busy with another operation: please try again in a moment");
}

const findInv = (reg, id) => {
  const inv = reg.invoices.find(i => i.id === id);
  if (!inv) throw new HttpError(404, "Invoice not found in the register");
  return inv;
};

async function invoices(request, env, url, email) {
  const path = url.pathname;
  const now = new Date().toISOString();
  try {
    if (path === "/api/invoices" && request.method === "GET") {
      const { reg } = await readRegister(env);
      return json(200, reg);
    }

    if (path === "/api/invoices/pdf" && request.method === "GET") {
      const key = url.searchParams.get("key") || "";
      if (!/^invoices\/\d{4}\/[A-Za-z0-9._-]+\.pdf$/.test(key)) return json(400, { error: "Invalid path" });
      const obj = await env.DATA.get(key);
      if (!obj) return json(404, { error: "PDF not found" });
      const name = (url.searchParams.get("name") || key.split("/").pop()).replace(/[^A-Za-z0-9._-]/g, "_");
      const disp = url.searchParams.get("download") ? "attachment" : "inline";
      return new Response(obj.body, { headers: { "content-type": "application/pdf", "content-disposition": `${disp}; filename="${name}"`, "cache-control": "private, max-age=300" } });
    }

    if (path === "/api/invoices/pdf" && request.method === "PUT") {
      const id = url.searchParams.get("id") || "";
      const buf = await request.arrayBuffer();
      if (!buf.byteLength || buf.byteLength > MAX_PDF) return json(400, { error: "PDF is empty or too large" });
      const head = new TextDecoder().decode(new Uint8Array(buf, 0, Math.min(5, buf.byteLength)));
      if (head !== "%PDF-") return json(400, { error: "The file is not a PDF" });
      const { reg } = await readRegister(env);
      const cur = findInv(reg, id);
      if (!cur.seq) return json(400, { error: "This invoice has no sequential number" });
      const rev = Number(url.searchParams.get("rev")) || 0;
      if (rev !== (cur.rev || 1)) return json(409, { error: "The invoice was changed in the meantime: reopen the register and regenerate the PDF" });
      const key = `invoices/${cur.year}/${String(cur.seq).padStart(3, "0")}-${cur.year}_r${rev}.pdf`;
      await env.DATA.put(key, buf, { httpMetadata: { contentType: "application/pdf" }, customMetadata: { invoice: cur.number, by: email } });
      const inv = await mutateRegister(env, (r) => {
        const x = findInv(r, id);
        x.pdfs = (x.pdfs || []).filter(p => p.key !== key);
        x.pdfs.push({ rev, key, at: now, by: email, size: buf.byteLength });
        return x;
      });
      return json(200, { invoice: inv });
    }

    if (request.method !== "POST") return json(405, { error: "Method not allowed" });
    const body = await request.json().catch(() => null);
    if (!body) return json(400, { error: "JSON non valido" });

    if (path === "/api/invoices/issue") {
      const data = cleanInvoice(body.invoice || {});
      const id = String(body.invoice?.id || "").match(/^[A-Za-z0-9_-]{4,64}$/) ? body.invoice.id : crypto.randomUUID();
      const inv = await mutateRegister(env, (reg) => {
        if (reg.invoices.some(i => i.id === id)) throw new HttpError(409, "This invoice has already been issued");
        const year = yearOf(data.date);
        const seq = nextSeq(reg, year);
        const x = { id, ...data, year, seq, number: fmtNumber(seq, year), status: "issued", rev: 1, issuedAt: now, issuedBy: email, pdfs: [], history: [{ at: now, by: email, action: "issued" }] };
        reg.invoices.push(x);
        reg.next[year] = seq + 1;
        return x;
      });
      return json(200, { invoice: inv });
    }

    if (path === "/api/invoices/update") {
      const data = cleanInvoice(body.invoice || {});
      const inv = await mutateRegister(env, (reg) => {
        const x = findInv(reg, String(body.id || ""));
        if (x.status === "void") throw new HttpError(400, "This invoice is void and cannot be edited");
        if (x.seq && yearOf(data.date) !== x.year) throw new HttpError(400, `The date must stay in ${x.year}: to change year, void this invoice and issue a new one`);
        Object.assign(x, data, { rev: (x.rev || 1) + 1, updatedAt: now, updatedBy: email });
        x.history = [...(x.history || []), { at: now, by: email, action: "updated", rev: x.rev }];
        return x;
      });
      return json(200, { invoice: inv });
    }

    if (path === "/api/invoices/void") {
      const reason = String(body.reason || "").slice(0, 300).trim();
      if (!reason) return json(400, { error: "Please give a reason for voiding" });
      const inv = await mutateRegister(env, (reg) => {
        const x = findInv(reg, String(body.id || ""));
        if (x.status === "void") return x;
        Object.assign(x, { status: "void", voidReason: reason, voidedAt: now, voidedBy: email });
        x.history = [...(x.history || []), { at: now, by: email, action: "void", reason }];
        return x;
      });
      return json(200, { invoice: inv });
    }

    if (path === "/api/invoices/next") {
      const year = Number(body.year), next = Number(body.next);
      if (!(year >= 2000 && year <= 2100) || !Number.isInteger(next) || next < 1 || next > 999999) return json(400, { error: "Invalid year or number" });
      const out = await mutateRegister(env, (reg) => {
        const used = reg.invoices.filter(i => i.year === year && i.seq).map(i => i.seq);
        const min = used.length ? Math.max(...used) + 1 : 1;
        if (next < min) throw new HttpError(400, `${year} already has invoices up to no. ${min - 1}: the next number must be at least ${min}`);
        reg.next[year] = next;
        return { year, next };
      });
      return json(200, out);
    }

    if (path === "/api/invoices/import") {
      const list = Array.isArray(body.invoices) ? body.invoices.slice(0, 500) : [];
      const out = await mutateRegister(env, (reg) => {
        const added = [], skipped = [];
        for (const src of list) {
          let data;
          try { data = cleanInvoice(src); } catch (e) { skipped.push({ id: src.id, number: src.number, why: e.message }); continue; }
          const id = String(src.id || "").match(/^[A-Za-z0-9_-]{4,64}$/) ? src.id : crypto.randomUUID();
          if (reg.invoices.some(i => i.id === id)) { skipped.push({ id, number: src.number, why: "already in the register" }); continue; }
          const pn = parseNumber(src.number);
          const year = pn ? pn.year : yearOf(data.date);
          if (pn && reg.invoices.some(i => i.year === pn.year && i.seq === pn.seq)) { skipped.push({ id, number: src.number, why: "number already used" }); continue; }
          const x = { id, ...data, year, seq: pn ? pn.seq : null, number: pn ? fmtNumber(pn.seq, pn.year) : String(src.number || "s.n."), status: "issued", rev: 1, legacy: true, issuedAt: now, issuedBy: email, pdfs: [], history: [{ at: now, by: email, action: "imported", from: String(src.number || "") }] };
          reg.invoices.push(x);
          if (pn) reg.next[pn.year] = Math.max(Number(reg.next[pn.year]) || 1, pn.seq + 1);
          added.push(x);
        }
        return { added, skipped };
      });
      return json(200, out);
    }

    return json(404, { error: "Richiesta non riconosciuta" });
  } catch (e) {
    if (e instanceof HttpError) return json(e.status, { error: e.message });
    throw e;
  }
}

function json(status, obj) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}
