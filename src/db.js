// Data layer — Cloudflare (Worker + R2).
// L'accesso è gestito da Cloudflare Access (email + codice): l'app non ha più
// un proprio login. Il server riconosce l'utente dall'email certificata da Access.

async function api(path, opts = {}) {
  const res = await fetch(path, {
    ...opts,
    credentials: "same-origin",
    headers: { "content-type": "application/json", ...(opts.headers || {}) },
  });
  // Sessione Cloudflare Access scaduta: ricarico la pagina per rifare l'accesso.
  if (res.status === 401 && path.startsWith("/api/data")) {
    location.reload();
    throw new Error("Sessione scaduta");
  }
  return res;
}

export async function me() {
  const res = await api("/api/me");
  return res.ok ? res.json() : null;
}

export function logout() {
  location.href = "/cdn-cgi/access/logout";
}

const dataUrl = (key) => "/api/data?key=" + encodeURIComponent(key);

export async function kvGet(key, def) {
  const res = await api(dataUrl(key));
  if (!res.ok) throw new Error("Lettura dati non riuscita");
  const data = await res.json();
  return data.value != null ? data.value : def;
}

export async function kvSet(key, value) {
  const res = await api(dataUrl(key), { method: "PUT", body: JSON.stringify({ value }) });
  if (!res.ok) throw new Error("Salvataggio non riuscito");
  return true;
}

export async function kvDelete(key) {
  await api(dataUrl(key), { method: "DELETE" });
  return true;
}

export async function extractDoc(kind, payload) {
  const res = await api("/api/extract", { method: "POST", body: JSON.stringify({ kind, ...payload }) });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Estrazione non riuscita");
  return data.text;
}

// ---------- Registro fatture IA Inc. (numerazione assegnata dal server) ----------
async function regCall(path, body) {
  const res = await api(path, body ? { method: "POST", body: JSON.stringify(body) } : {});
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Error ${res.status}`);
  return data;
}
export const regList = () => regCall("/api/invoices");
export const regIssue = (invoice) => regCall("/api/invoices/issue", { invoice });
export const regUpdate = (id, invoice) => regCall("/api/invoices/update", { id, invoice });
export const regVoid = (id, reason) => regCall("/api/invoices/void", { id, reason });
export const regSetNext = (year, next) => regCall("/api/invoices/next", { year, next });
export const regImport = (invoices) => regCall("/api/invoices/import", { invoices });
export async function regUploadPdf(id, rev, bytes) {
  const res = await fetch(`/api/invoices/pdf?id=${encodeURIComponent(id)}&rev=${encodeURIComponent(rev)}`, {
    method: "PUT", credentials: "same-origin", headers: { "content-type": "application/pdf" }, body: bytes,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Error ${res.status}`);
  return data;
}
export const regPdfUrl = (key, name, download) =>
  `/api/invoices/pdf?key=${encodeURIComponent(key)}&name=${encodeURIComponent(name || "")}${download ? "&download=1" : ""}`;
