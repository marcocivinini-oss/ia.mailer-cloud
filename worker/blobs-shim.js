// Sostituto di "@netlify/blobs" per Cloudflare: stessi metodi (get, setJSON, list, delete),
// ma i dati stanno nello spazio R2 del Mailer. Ogni archivio ("disiscritti", "invii", "campagne")
// è una cartella dello spazio R2.
let BUCKET = null;
export function bindBlobs(bucket) { BUCKET = bucket; }

export function getStore(opts) {
  if (!BUCKET) throw new Error("Archivio non collegato");
  const name = typeof opts === "string" ? opts : opts.name;
  const p = name + "/";
  const b = BUCKET;
  return {
    async get(key, o = {}) {
      const obj = await b.get(p + key);
      if (!obj) return null;
      if (o.type === "json") return obj.json();
      return obj.text();
    },
    async setJSON(key, value) {
      await b.put(p + key, JSON.stringify(value), { httpMetadata: { contentType: "application/json" } });
    },
    async set(key, value) { await b.put(p + key, value); },
    async delete(key) { await b.delete(p + key); },
    async list() {
      const blobs = [];
      let cursor;
      do {
        const r = await b.list({ prefix: p, cursor });
        blobs.push(...r.objects.map((x) => ({ key: x.key.slice(p.length) })));
        cursor = r.truncated ? r.cursor : undefined;
      } while (cursor);
      return { blobs };
    },
  };
}

// Accesso diretto ai file di un archivio: per contenuti grandi che il server salva e restituisce
// così come sono, senza rielaborarli (meno tempo di calcolo per ogni richiesta).
export function rawStore(name) {
  if (!BUCKET) throw new Error("Archivio non collegato");
  const p = name + "/", b = BUCKET;
  return {
    put: (key, body, opts) => b.put(p + key, body, opts),
    get: (key) => b.get(p + key),
    delete: (key) => b.delete(p + key),
  };
}
