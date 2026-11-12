/* IA Mailer — interfaccia e invio tramite Gmail API */
(function () {
  "use strict";
  const C = window.IA_MAILER_CONFIG || {};
  const M = window.IAMailer;
  const $ = (id) => document.getElementById(id);
  const SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/gmail.metadata",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile"
  ].join(" ");

  const state = {
    token: null, tokenClient: null, tokenWaiters: [], user: null,
    mode: "text", source: "file",
    recipients: [], columns: [], included: new Set(),
    canvaImages: [], attachments: [], logo: null,
    results: new Map(), sending: false, paused: false, stopped: false,
    unsub: {}, history: {}, bounced: {}, registryOk: false, unsubList: [], bouncedList: [],
    plans: [], planCache: {}
  };

  // ---------- utilità ----------
  const store = {
    get(k, d) { try { const v = localStorage.getItem("iaMailer." + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem("iaMailer." + k, JSON.stringify(v)); } catch (e) { /* storage non disponibile */ } }
  };
  let toastT;
  function toast(msg, err) {
    const t = $("toast"); t.textContent = msg; t.className = "toast" + (err ? " err" : "");
    clearTimeout(toastT); toastT = setTimeout(() => t.classList.add("hidden"), err ? 7000 : 3500);
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const fileToB64 = (file) => new Promise((res, rej) => {
    const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = rej; r.readAsDataURL(file);
  });

  // ---------- tabs ----------
  document.querySelectorAll(".tabs").forEach((tabs) => {
    const section = tabs.closest("section");
    tabs.addEventListener("click", (e) => {
      const b = e.target.closest(".tab"); if (!b) return;
      tabs.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === b));
      section.querySelectorAll(".tabpane").forEach((p) => p.classList.toggle("hidden", p.dataset.pane !== b.dataset.tab));
      if (tabs.dataset.group === "mode") { state.mode = b.dataset.tab; refreshPreview(); }
      else if (tabs.dataset.group === "source") state.source = b.dataset.tab;
    });
  });
  if (!C.HUBSPOT_ENABLED) $("tabHubspot").classList.add("hidden");

  // ---------- autenticazione Google ----------
  function initGoogle() {
    if (!window.google || !google.accounts || !google.accounts.oauth2) return setTimeout(initGoogle, 200);
    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: C.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      hd: C.ALLOWED_DOMAINS && C.ALLOWED_DOMAINS.length === 1 ? C.ALLOWED_DOMAINS[0] : undefined,
      callback: (resp) => {
        const waiters = state.tokenWaiters.splice(0);
        if (resp.error) { waiters.forEach((w) => w.reject(new Error(resp.error))); return toast("Accesso non riuscito: " + resp.error, true); }
        state.token = resp.access_token;
        state.tokenExp = Date.now() + (resp.expires_in - 120) * 1000;
        waiters.forEach((w) => w.resolve(state.token));
      },
      error_callback: (err) => {
        state.tokenWaiters.splice(0).forEach((w) => w.reject(new Error(err.type || "popup")));
      }
    });
  }
  function getToken(prompt) {
    if (state.token && Date.now() < state.tokenExp && prompt == null) return Promise.resolve(state.token);
    return new Promise((resolve, reject) => {
      state.tokenWaiters.push({ resolve, reject });
      state.tokenClient.requestAccessToken({ prompt: prompt == null ? "" : prompt, login_hint: state.user && state.user.email });
    });
  }
  async function gfetch(url, opts, retry) {
    const token = await getToken();
    const r = await fetch(url, Object.assign({}, opts, { headers: Object.assign({ Authorization: "Bearer " + token }, (opts || {}).headers) }));
    if (r.status === 401 && !retry) { state.token = null; return gfetch(url, opts, true); }
    return r;
  }

  $("btnLogin").onclick = async () => {
    for (let i = 0; i < 25 && !state.tokenClient; i++) await sleep(200);
    if (!state.tokenClient) return toast("Servizio Google non raggiungibile: verifica la connessione e ricarica la pagina", true);
    if (C.GOOGLE_CLIENT_ID.startsWith("INSERIRE")) return toast("Configurare GOOGLE_CLIENT_ID in config.js", true);
    try {
      await getToken("select_account");
      const info = await (await gfetch("https://www.googleapis.com/oauth2/v3/userinfo")).json();
      const dom = (info.email || "").split("@")[1];
      if (C.ALLOWED_DOMAINS && C.ALLOWED_DOMAINS.length && !C.ALLOWED_DOMAINS.includes(dom)) {
        try { google.accounts.oauth2.revoke(state.token, () => {}); } catch (e) {}
        state.token = null;
        return toast("Account non autorizzato: usa un indirizzo " + C.ALLOWED_DOMAINS.join(", "), true);
      }
      state.user = info;
      await loadSendAs();
      onLoggedIn();
    } catch (e) { toast("Accesso annullato o non riuscito", true); }
  };

  async function loadSendAs() {
    const sel = $("fromAddress"); sel.innerHTML = "";
    let list = [];
    try {
      const r = await gfetch("https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs");
      if (r.ok) list = ((await r.json()).sendAs || []).filter((s) => s.isPrimary || s.verificationStatus === "accepted");
    } catch (e) { /* senza alias usiamo l'indirizzo principale */ }
    if (!list.length) list = [{ sendAsEmail: state.user.email, displayName: state.user.name, isDefault: true }];
    list.forEach((s) => {
      const o = document.createElement("option"); o.value = s.sendAsEmail; o.textContent = s.sendAsEmail;
      o.dataset.name = s.displayName || ""; if (s.isDefault) o.selected = true; sel.appendChild(o);
    });
    const saved = store.get("sender." + state.user.email, {});
    if (saved.from && list.some((s) => s.sendAsEmail === saved.from)) sel.value = saved.from;
    $("fromName").value = saved.name || (sel.selectedOptions[0] && sel.selectedOptions[0].dataset.name) || state.user.name || "";
    $("replyTo").value = saved.replyTo || "";
    $("signature").value = saved.signature != null ? saved.signature : (state.user.name || "") + "\nIA International Advisors";
  }
  function saveSender() {
    if (!state.user) return;
    store.set("sender." + state.user.email, { from: $("fromAddress").value, name: $("fromName").value, replyTo: $("replyTo").value, signature: $("signature").value });
  }
  ["fromAddress", "fromName", "replyTo", "signature"].forEach((id) => $(id).addEventListener("change", () => { saveSender(); refreshPreview(); }));

  function onLoggedIn() {
    $("loginArea").classList.add("hidden"); $("senderArea").classList.remove("hidden");
    $("userBox").innerHTML = M.escapeHtml(state.user.email) + '<br><a id="lnkLogout">Esci</a>';
    $("lnkLogout").onclick = () => { try { google.accounts.oauth2.revoke(state.token, () => {}); } catch (e) {} location.reload(); };
    if (!$("testTo").value) $("testTo").value = state.user.email;
    refreshPreview(); updateSendInfo(); loadCampaigns(); loadPlans();
  }

  // ---------- editor testo ----------
  $("toolbar").addEventListener("click", (e) => {
    const b = e.target.closest("button"); if (!b) return; e.preventDefault();
    const cmd = b.dataset.cmd; $("editor").focus();
    if (cmd === "h2" || cmd === "p") document.execCommand("formatBlock", false, cmd);
    else if (cmd === "link") { const u = prompt("Indirizzo del link (https://...)"); if (u) document.execCommand("createLink", false, u); }
    else document.execCommand(cmd, false, null);
    refreshPreviewSoon();
  });
  $("mergeInsert").addEventListener("change", (e) => {
    if (!e.target.value) return; $("editor").focus();
    document.execCommand("insertText", false, "{{" + e.target.value + "}}"); e.target.value = ""; refreshPreviewSoon();
  });
  $("editor").innerHTML = store.get("draft.editor", $("editor").innerHTML);
  $("subject").value = store.get("draft.subject", "");
  $("editor").addEventListener("input", () => { store.set("draft.editor", $("editor").innerHTML); refreshPreviewSoon(); });
  $("subject").addEventListener("input", () => { store.set("draft.subject", $("subject").value); refreshPreviewSoon(); });
  ["preheader", "canvaIntro", "canvaWebLink", "htmlSource", "ctaText", "ctaUrl"].forEach((id) => $(id).addEventListener("input", refreshPreviewSoon));
  ["useTemplate", "addUnsub", "ctaOn"].forEach((id) => $(id).addEventListener("change", refreshPreview));

  // ---------- immagini Canva ----------
  function optimizeImage(file) {
    return new Promise((resolve, reject) => {
      const img = new Image(), url = URL.createObjectURL(file);
      img.onload = () => {
        const w = Math.min(1200, img.naturalWidth), h = Math.round(img.naturalHeight * w / img.naturalWidth);
        const cv = document.createElement("canvas"); cv.width = w; cv.height = h;
        const ctx = cv.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, w, h); ctx.drawImage(img, 0, 0, w, h);
        const dataUrl = cv.toDataURL("image/jpeg", 0.86); URL.revokeObjectURL(url);
        resolve({ dataUrl, b64: dataUrl.split(",")[1], mime: "image/jpeg" });
      };
      img.onerror = reject; img.src = url;
    });
  }
  $("canvaImages").addEventListener("change", async (e) => {
    const files = [...e.target.files].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
    for (const f of files) {
      try {
        const o = await optimizeImage(f);
        state.canvaImages.push(Object.assign(o, { cid: "img" + Date.now() + Math.random().toString(36).slice(2, 7) + "@ia", name: f.name.replace(/\.\w+$/, "") + ".jpg", link: "", alt: "" }));
      } catch (err) { toast("Impossibile leggere " + f.name, true); }
    }
    e.target.value = ""; renderCanvaList(); refreshPreview();
  });
  function renderCanvaList() {
    const box = $("canvaList"); box.innerHTML = "";
    state.canvaImages.forEach((im, i) => {
      const d = document.createElement("div"); d.className = "imgitem";
      d.innerHTML = '<img src="' + im.dataUrl + '" alt=""><div style="flex:1;display:flex;flex-direction:column;gap:4px">' +
        '<input placeholder="Link al clic (facoltativo)" value="' + M.escapeHtml(im.link) + '" data-k="link">' +
        '<input placeholder="Testo alternativo (consigliato)" value="' + M.escapeHtml(im.alt) + '" data-k="alt"></div>' +
        '<button data-a="up" title="Su">↑</button><button data-a="down" title="Giù">↓</button><button data-a="del" title="Rimuovi">✕</button>';
      d.querySelectorAll("input").forEach((inp) => inp.addEventListener("input", () => { im[inp.dataset.k] = inp.value.trim(); refreshPreviewSoon(); }));
      d.querySelectorAll("button").forEach((b) => b.onclick = () => {
        const a = b.dataset.a, arr = state.canvaImages;
        if (a === "del") arr.splice(i, 1);
        if (a === "up" && i > 0) [arr[i - 1], arr[i]] = [arr[i], arr[i - 1]];
        if (a === "down" && i < arr.length - 1) [arr[i + 1], arr[i]] = [arr[i], arr[i + 1]];
        renderCanvaList(); refreshPreview();
      });
      box.appendChild(d);
    });
  }

  // ---------- allegati ----------
  $("attachments").addEventListener("change", async (e) => {
    for (const f of e.target.files) {
      const total = state.attachments.reduce((s, a) => s + a.size, 0) + f.size;
      if (total > 15 * 1024 * 1024) { toast("Superato il limite di 15 MB per gli allegati", true); break; }
      state.attachments.push({ name: f.name, mime: f.type || "application/octet-stream", size: f.size, b64: await fileToB64(f) });
    }
    e.target.value = ""; renderAttachments();
  });
  function renderAttachments() {
    $("attList").innerHTML = state.attachments.map((a, i) =>
      '<span class="chip">' + M.escapeHtml(a.name) + " (" + Math.round(a.size / 1024) + ' KB) <button data-i="' + i + '">✕</button></span>').join("");
    $("attList").querySelectorAll("button").forEach((b) => b.onclick = () => { state.attachments.splice(+b.dataset.i, 1); renderAttachments(); });
  }

  // ---------- destinatari ----------
  function setRecipients(res, label) {
    state.recipients = res.recipients; state.columns = res.columns;
    state.included = new Set(res.recipients.map((r) => r.email));
    let s = "<b>" + res.recipients.length + "</b> destinatari caricati" + (label ? " da " + M.escapeHtml(label) : "");
    if (res.duplicates) s += " · " + res.duplicates + " duplicati rimossi";
    if (res.invalid.length) s += " · " + res.invalid.length + " indirizzi non validi scartati";
    if (res.optout) s += " · " + res.optout + " contatti HubSpot esclusi (disiscritti)";
    state.baseSummary = s; $("recSummary").innerHTML = s; $("recSummary").classList.remove("hidden");
    state.results = new Map();
    state.unsub = {}; state.history = {}; state.bounced = {};
    selectSendMode(res.recipients.length > 150 ? "gradual" : "now");
    renderRecipients(); refreshMergeOptions(); refreshPreview(); updateSendInfo();
    checkRegistry();
  }

  // ---------- registro condiviso ----------
  async function reg(action, payload) {
    const token = await getToken();
    const r = await fetch("/.netlify/functions/registro", {
      method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({ action }, payload || {}))
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error || "Registro non raggiungibile (" + r.status + ")");
    return j;
  }
  async function checkRegistry() {
    if (!state.user || !state.recipients.length) return;
    try {
      const all = state.recipients.map((r) => r.email), j = { unsubscribed: {}, history: {}, bounced: {} };
      for (let i = 0; i < all.length; i += 2000) {
        const part = await reg("check", { emails: all.slice(i, i + 2000) });
        Object.assign(j.unsubscribed, part.unsubscribed); Object.assign(j.history, part.history); Object.assign(j.bounced, part.bounced || {});
      }
      state.unsub = j.unsubscribed; state.history = j.history; state.bounced = j.bounced; state.registryOk = true;
      const nU = Object.keys(state.unsub).length, nH = Object.keys(state.history).length, nB = Object.keys(state.bounced).length;
      $("recSummary").innerHTML = state.baseSummary + (nU ? " · <b>" + nU + "</b> disiscritti esclusi" : "") + (nB ? " · " + nB + " non recapitabili esclusi" : "") +
        (nH ? " · " + nH + " hanno già ricevuto nostre email" : "");
    } catch (e) {
      state.registryOk = false;
      toast("Registro condiviso non raggiungibile: disiscritti e storico non verificati (" + e.message + ")", true);
    }
    renderRecipients(); updateSendInfo();
  }
  // Motivo per cui un destinatario non riceverà l'email (null = verrà inviato)
  function skipReason(r) {
    if (state.unsub[r.email]) return "disiscritto";
    if (state.bounced[r.email]) return "non recapitabile";
    if (M.isExcluded(r.email, $("exclusions").value)) return "escluso";
    const h = state.history[r.email];
    if (h) {
      const subj = $("subject").value.trim().toLowerCase();
      if ($("skipSameSubject").checked && subj && h.subjects.some((s) => String(s).trim().toLowerCase() === subj)) return "già ricevuta";
      const days = +$("recentDays").value;
      if (days && Date.now() - new Date(h.last.t).getTime() < days * 864e5) return "contattato di recente";
    }
    return null;
  }
  const fmtDate = (t) => { try { return new Date(t).toLocaleDateString("it-IT"); } catch (e) { return ""; } };
  ["recentDays", "skipSameSubject"].forEach((id) => $(id).addEventListener("change", () => { renderRecipients(); updateSendInfo(); }));
  $("subject").addEventListener("change", () => { renderRecipients(); updateSendInfo(); });

  function renderRecipients() {
    const excl = $("exclusions").value;
    const pref = ["Email", "Nome", "Cognome", "Azienda"];
    const cols = pref.filter((c) => state.columns.includes(c))
      .concat(state.columns.filter((c) => !pref.includes(c) && !/^e-?mail$/i.test(c)).slice(0, 3));
    const head = '<tr><th><input type="checkbox" id="chkAll" checked></th>' + cols.map((c) => "<th>" + M.escapeHtml(c) + "</th>").join("") + "<th>Storico</th><th>Stato</th></tr>";
    const body = state.recipients.slice(0, 1000).map((r, i) => {
      const why = skipReason(r), res = state.results.get(r.email), h = state.history[r.email];
      const hist = h ? '<span title="' + M.escapeHtml(h.subjects.slice(-5).join("\n")) + '">' + h.count + (h.count === 1 ? " invio" : " invii") + " · ult. " + fmtDate(h.last.t) + "</span>" : "";
      return '<tr class="' + (why ? "excluded" : "") + '"><td><input type="checkbox" data-i="' + i + '"' + (state.included.has(r.email) ? " checked" : "") + (why ? " disabled" : "") + "></td>" +
        cols.map((c) => "<td>" + M.escapeHtml(r.fields[c] || "") + "</td>").join("") + "<td>" + hist + "</td>" +
        "<td>" + (res ? (res.status === "ok" ? "✓ inviato" : "✕ " + M.escapeHtml(res.error || "errore")) : why || "") + "</td></tr>";
    }).join("");
    $("recTable").innerHTML = head + body;
    $("recTableWrap").classList.toggle("hidden", !state.recipients.length);
    $("recTable").querySelectorAll("tbody input, tr input[data-i]").forEach((c) => c.onchange = () => {
      const r = state.recipients[+c.dataset.i]; c.checked ? state.included.add(r.email) : state.included.delete(r.email); updateSendInfo();
    });
    const all = $("chkAll"); if (all) all.onchange = () => {
      state.included = all.checked ? new Set(state.recipients.map((r) => r.email)) : new Set(); renderRecipients(); updateSendInfo();
    };
    const sel = $("previewFor"), cur = sel.value;
    sel.innerHTML = '<option value="-1">Destinatario di esempio</option>' +
      state.recipients.slice(0, 200).map((r, i) => '<option value="' + i + '">' + M.escapeHtml(r.email) + "</option>").join("");
    sel.value = state.recipients[+cur] ? cur : (state.recipients.length ? "0" : "-1");
  }
  function refreshMergeOptions() {
    const cols = ["Nome", "Cognome", "Azienda", "Email"].concat(state.columns).filter((v, i, a) => a.indexOf(v) === i);
    $("mergeInsert").innerHTML = '<option value="">+ Campo personalizzato</option>' + cols.map((c) => '<option>' + M.escapeHtml(c) + "</option>").join("");
  }
  const queue = () => state.recipients.filter((r) => state.included.has(r.email) && !skipReason(r));

  $("recFile").addEventListener("change", async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", codepage: 65001 });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
      const res = M.normalizeRows(rows);
      if (!res.recipients.length) return toast("Nessun indirizzo email trovato nel file", true);
      setRecipients(res, f.name);
    } catch (err) { toast("File non leggibile: " + err.message, true); }
    e.target.value = ""; // consente di ricaricare lo stesso file
  });
  $("btnPaste").onclick = () => {
    const res = M.parsePasted($("pasteList").value);
    if (!res.recipients.length) return toast("Nessun indirizzo valido", true);
    setRecipients(res, "elenco incollato");
  };
  $("exclusions").value = store.get("exclusions", "");
  $("exclusions").addEventListener("input", () => { store.set("exclusions", $("exclusions").value); renderRecipients(); updateSendInfo(); });

  // ---------- HubSpot ----------
  async function hs(params) {
    const token = await getToken();
    const r = await fetch("/.netlify/functions/hubspot?" + new URLSearchParams(params), { headers: { Authorization: "Bearer " + token } });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error || "Errore HubSpot " + r.status);
    return j;
  }
  $("btnHsSearch").onclick = async () => {
    if (!state.user) return toast("Accedi prima con Google", true);
    $("hsLists").innerHTML = '<span class="muted">Ricerca in corso…</span>';
    try {
      const j = await hs({ action: "lists", q: $("hsQuery").value.trim() });
      if (!j.lists.length) { $("hsLists").innerHTML = '<span class="muted">Nessuna lista trovata</span>'; return; }
      $("hsLists").innerHTML = j.lists.map((l) => '<div class="hslist"><span><b>' + M.escapeHtml(l.name) + "</b> " +
        (l.size != null ? '<span class="muted">· ' + l.size + " contatti</span>" : "") +
        '</span><button class="btn" data-id="' + M.escapeHtml(l.listId) + '" data-name="' + M.escapeHtml(l.name) + '">Usa questa lista</button></div>').join("");
      $("hsLists").querySelectorAll("button").forEach((b) => b.onclick = async () => {
        b.disabled = true; b.textContent = "Caricamento…";
        try {
          const m = await hs({ action: "members", listId: b.dataset.id });
          const rows = m.contacts.filter((c) => !c.optout).map((c) => ({ Email: c.email, Nome: c.firstname, Cognome: c.lastname, Azienda: c.company, Paese: c.country, Ruolo: c.jobtitle }));
          const res = M.normalizeRows(rows); res.optout = m.contacts.length - rows.length;
          setRecipients(res, "HubSpot · " + b.dataset.name);
        } catch (err) { toast(err.message, true); }
        b.disabled = false; b.textContent = "Usa questa lista";
      });
    } catch (err) { $("hsLists").innerHTML = ""; toast(err.message, true); }
  };

  // ---------- composizione ----------
  function sampleFields() {
    const i = +$("previewFor").value;
    if (state.recipients[i]) return state.recipients[i].fields;
    return { Nome: "Mario", Cognome: "Rossi", Azienda: "Azienda Esempio S.p.A.", Email: "mario.rossi@esempio.it" };
  }
  function composeOpts() {
    return {
      mode: state.mode,
      bodyHtml: state.mode === "html" ? $("htmlSource").value : $("editor").innerHTML,
      canvaIntro: $("canvaIntro").value, canvaImages: state.canvaImages, canvaWebLink: $("canvaWebLink").value.trim(),
      preheader: $("preheader").value, useTemplate: $("useTemplate").checked, addUnsub: $("addUnsub").checked,
      cta: $("ctaOn").checked ? { text: $("ctaText").value.trim(), url: $("ctaUrl").value.trim() } : null,
      signature: $("signature").value.trim(),
      fromEmail: $("fromAddress").value || (state.user && state.user.email) || "info@iainternationaladvisors.com",
      footer: C.COMPANY_FOOTER, website: C.WEBSITE, logoCid: state.logo ? "ialogo@ia" : null
    };
  }
  function buildMessage(fields, to) {
    const o = composeOpts();
    const html = M.buildEmailHtml(o, fields);
    const inline = [];
    if (o.useTemplate && state.logo) inline.push({ cid: "ialogo@ia", mime: "image/png", b64: state.logo, name: "logo-ia.png" });
    if (o.mode === "canva") state.canvaImages.forEach((im) => inline.push({ cid: im.cid, mime: im.mime, b64: im.b64, name: im.name }));
    return {
      fromName: $("fromName").value.trim(), fromEmail: o.fromEmail, to, replyTo: $("replyTo").value.trim(), unsubUrl: fields.__unsubUrl,
      subject: M.merge($("subject").value, fields, false).trim(),
      html, text: M.htmlToText(html), inline, attachments: state.attachments
    };
  }
  function validate() {
    if (!state.user) return "Accedi con Google (sezione 1)";
    if (!$("subject").value.trim()) return "Inserisci l'oggetto";
    if (state.mode === "canva" && !state.canvaImages.length) return "Carica almeno un'immagine da Canva";
    if (state.mode === "html" && !$("htmlSource").value.trim()) return "Incolla il codice HTML";
    if (state.mode === "text" && !$("editor").innerText.trim()) return "Scrivi il testo dell'email";
    if ($("ctaOn").checked && !/^https?:\/\//.test($("ctaUrl").value.trim())) return "Il link del pulsante deve iniziare con https://";
    return null;
  }

  let pvT;
  function refreshPreviewSoon() { clearTimeout(pvT); pvT = setTimeout(refreshPreview, 350); }
  function refreshPreview() {
    const f = sampleFields(), o = composeOpts();
    let html = M.buildEmailHtml(o, f);
    if (state.logo) html = html.split("cid:ialogo@ia").join("data:image/png;base64," + state.logo);
    state.canvaImages.forEach((im) => { html = html.split("cid:" + im.cid).join(im.dataUrl); });
    $("previewFrame").srcdoc = html;
    $("previewSubject").innerHTML = "<b>Oggetto:</b> " + M.escapeHtml(M.merge($("subject").value, f, false) || "—") +
      ($("preheader").value ? ' <span class="muted">— ' + M.escapeHtml(M.merge($("preheader").value, f, false)) + "</span>" : "") +
      (state.attachments.length ? "<br><b>Allegati:</b> " + state.attachments.map((a) => M.escapeHtml(a.name)).join(", ") : "");
  }
  $("btnPreview").onclick = refreshPreview;
  $("previewFor").onchange = refreshPreview;

  // ---------- invio ----------
  async function sendOne(msg) {
    const mime = M.buildMime(msg);
    // L'email viene consegnata al server del Mailer, che la invia tramite Gmail con il token
    // dell'utente: parte comunque dalla sua casella. (Le chiamate dirette dal browser a Gmail
    // spedivano l'email ma la risposta veniva bloccata: "Failed to fetch".)
    const url = "/.netlify/functions/invia";
    const opts = { method: "POST", headers: { "Content-Type": "message/rfc822" }, body: mime };
    for (let attempt = 0; attempt < 4; attempt++) {
      const r = await gfetch(url, opts);
      if (r.ok) { const j = await r.json(); return { id: j.id, threadId: j.threadId }; }
      const j = await r.json().catch(() => ({}));
      const reason = (j.error && (j.error.message || (j.error.errors && j.error.errors[0] && j.error.errors[0].reason))) || "HTTP " + r.status;
      if (/limit exceeded|quota|daily/i.test(reason) && r.status !== 429) { const e = new Error("Limite giornaliero di invio raggiunto"); e.fatal = true; throw e; }
      if (r.status === 429 || r.status >= 500) { await sleep(2000 * Math.pow(2, attempt)); continue; }
      throw new Error(reason);
    }
    throw new Error("Troppi tentativi, riprovare più tardi");
  }

  $("btnTest").onclick = async () => {
    const err = validate(); if (err) return toast(err, true);
    const to = $("testTo").value.trim(); if (!M.isValidEmail(to)) return toast("Indirizzo di prova non valido", true);
    $("btnTest").disabled = true;
    try {
      const msg = buildMessage(sampleFields(), to); msg.subject = "[PROVA] " + msg.subject;
      await sendOne(msg); toast("Email di prova inviata a " + to);
    } catch (e) { toast("Invio di prova non riuscito: " + e.message, true); }
    $("btnTest").disabled = false;
  };

  function updateSendInfo() {
    updateEstimate();
    const n = queue().length;
    $("btnSend").disabled = !n || !state.user || state.sending;
    $("btnSend").textContent = n ? "Invia a " + n + " destinatari" : "Invia a tutti";
    if (!state.recipients.length) return;
    const mins = Math.ceil(n * ((C.SEND_DELAY_MS || 1500) + 600) / 60000);
    const skipped = state.recipients.filter((r) => skipReason(r)).length;
    $("sendInfo").innerHTML = (skipped ? skipped + " destinatari esclusi (disiscritti, esclusioni o già contattati). " : "") + "<b>" + n + "</b> email personalizzate, una per destinatario (nessuno vede gli altri indirizzi). Tempo stimato: circa " + mins +
      " min. Tieni aperta questa pagina durante l'invio. Limite Google Workspace: 2.000 email al giorno per casella." +
      (n > 300 ? '<br><b>Per una lista di questa dimensione è consigliata la modalità «Graduale, a tappe».</b>' : "");
  }

  function log(line, cls) {
    const d = document.createElement("div"); d.className = cls || ""; d.textContent = new Date().toLocaleTimeString("it-IT") + "  " + line;
    $("log").prepend(d);
  }
  async function runQueue(list, campaignId, links) {
    state.sending = true; state.paused = false; state.stopped = false;
    ["btnPause", "btnStop", "progress"].forEach((id) => $(id).classList.remove("hidden"));
    ["btnRetry"].forEach((id) => $(id).classList.add("hidden"));
    $("btnReport").classList.remove("hidden"); $("btnSend").disabled = true;
    let ok = 0, ko = 0, pending = [];
    const flush = async () => {
      if (!campaignId || !pending.length) return;
      const batch = pending; pending = [];
      try { await reg("log", { campaignId, results: batch }); }
      catch (e) { pending = batch.concat(pending); log("Registro non aggiornato, nuovo tentativo al prossimo blocco: " + e.message, "err"); }
    };
    for (let i = 0; i < list.length; i++) {
      while (state.paused && !state.stopped) await sleep(300);
      if (state.stopped) { log("Invio interrotto dall'utente", "err"); break; }
      const r = list[i];
      try {
        const fields = Object.assign({}, r.fields, { __unsubUrl: links && links[r.email] });
        const g = await sendOne(buildMessage(fields, r.email));
        const res = { status: "ok", id: g.id, threadId: g.threadId, time: new Date().toISOString() };
        state.results.set(r.email, res); pending.push(Object.assign({ email: r.email }, res)); ok++; log("✓ " + r.email, "ok");
      } catch (e) {
        const res = { status: "error", error: e.message, time: new Date().toISOString() };
        state.results.set(r.email, res); pending.push(Object.assign({ email: r.email }, res)); ko++;
        log("✕ " + r.email + " — " + e.message, "err");
        if (e.fatal) { toast(e.message + ". Riprendi domani con «Riprova falliti» o ricaricando la lista: i già inviati verranno saltati.", true); break; }
      }
      $("progressBar").style.width = Math.round((i + 1) / list.length * 100) + "%";
      $("progressText").textContent = (i + 1) + " di " + list.length + " · inviati " + ok + " · errori " + ko;
      if (pending.length >= 10) await flush();
      if (i < list.length - 1) await sleep(C.SEND_DELAY_MS || 1500);
    }
    await flush(); if (pending.length) await flush();
    state.sending = false; state.lastCampaign = { id: campaignId, links };
    if (campaignId) { checkRegistry(); loadCampaigns(); }
    ["btnPause", "btnStop"].forEach((id) => $(id).classList.add("hidden"));
    const failed = [...state.results.values()].filter((x) => x.status === "error").length;
    $("btnRetry").classList.toggle("hidden", !failed);
    toast("Invio completato: " + ok + " inviati" + (ko ? ", " + ko + " errori" : ""));
    renderRecipients(); updateSendInfo();
  }

  $("btnSend").onclick = async () => {
    const err = validate(); if (err) return toast(err, true);
    const list = queue();
    if (!list.length) return toast("Nessun destinatario da inviare");
    if (!state.registryOk && !confirm("Il registro condiviso non è raggiungibile: disiscritti e storico non sono stati verificati e questo invio non verrà registrato. Procedere comunque?")) return;
    if (!confirm("Confermi l'invio di " + list.length + " email da " + $("fromAddress").value + "?\nOggetto: " + $("subject").value)) return;
    saveSender(); $("log").innerHTML = "";
    let campaignId = null, links = null;
    if (state.registryOk) {
      try {
        const j = await reg("prepare", { emails: list.map((r) => r.email), campaign: { subject: $("subject").value.trim(), sender: $("fromAddress").value, mode: state.mode, source: state.source, total: list.length } });
        campaignId = j.campaignId; links = j.links;
      } catch (e) {
        if (!confirm("Impossibile aprire la campagna nel registro (" + e.message + "). Inviare senza registrazione?")) return;
      }
    }
    runQueue(list, campaignId, links);
  };
  $("btnPause").onclick = () => { state.paused = !state.paused; $("btnPause").textContent = state.paused ? "Riprendi" : "Pausa"; };
  $("btnStop").onclick = () => { if (confirm("Interrompere l'invio?")) state.stopped = true; };
  $("btnRetry").onclick = () => {
    const list = state.recipients.filter((r) => (state.results.get(r.email) || {}).status === "error");
    if (list.length && confirm("Riprovare l'invio a " + list.length + " destinatari?")) runQueue(list, state.lastCampaign && state.lastCampaign.id, state.lastCampaign && state.lastCampaign.links);
  };
  $("btnReport").onclick = () => {
    const rows = state.recipients.filter((r) => state.results.has(r.email)).map((r) => {
      const x = state.results.get(r.email);
      return { Email: r.email, Nome: r.fields.Nome || "", Azienda: r.fields.Azienda || "", Esito: x.status === "ok" ? "Inviato" : "Errore", Dettaglio: x.error || "", IdGmail: x.id || "", DataOra: x.time };
    });
    const blob = new Blob(["﻿" + M.toCsv(rows)], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = "report-invio-" + new Date().toISOString().slice(0, 10) + ".csv"; a.click();
  };
  window.addEventListener("beforeunload", (e) => { if (state.sending) { e.preventDefault(); e.returnValue = ""; } });

  // ---------- pannello registro ----------
  function downloadCsv(rows, name) {
    const blob = new Blob(["\ufeff" + M.toCsv(rows)], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; a.click();
  }
  async function loadCampaigns() {
    if (!state.user) return;
    try {
      const { campaigns } = await reg("campaigns");
      $("campTable").innerHTML = "<tr><th>Data</th><th>Oggetto</th><th>Mittente</th><th>Inviati</th><th>Errori</th><th></th></tr>" +
        (campaigns.length ? campaigns.map((c) => "<tr><td>" + fmtDate(c.createdAt) + "</td><td>" + M.escapeHtml(c.subject) + "</td><td>" + M.escapeHtml(c.sender) +
          "</td><td>" + c.ok + "</td><td>" + c.ko + '</td><td><button class="btn" data-id="' + M.escapeHtml(c.id) + '">Destinatari CSV</button></td></tr>').join("")
          : '<tr><td colspan="6" class="muted">Nessuna campagna registrata</td></tr>');
      $("campTable").querySelectorAll("button").forEach((b) => b.onclick = async () => {
        try {
          const j = await reg("campaign", { id: b.dataset.id });
          downloadCsv(j.recipients.map((r) => ({ Email: r.e, Esito: r.st === "ok" ? "Inviato" : "Errore", Dettaglio: r.err, DataOra: r.t })),
            "campagna-" + j.meta.createdAt.slice(0, 10) + ".csv");
        } catch (e) { toast(e.message, true); }
      });
    } catch (e) { $("campTable").innerHTML = '<tr><td class="muted">' + M.escapeHtml(e.message) + "</td></tr>"; }
  }
  async function loadUnsubs() {
    if (!state.user) return toast("Accedi prima con Google", true);
    try { state.unsubList = (await reg("unsubscribes")).unsubscribes; renderUnsubs(); }
    catch (e) { toast(e.message, true); }
  }
  const SRC = { link: "link nell'email", manuale: "inserito a mano" };
  function renderUnsubs() {
    const q = $("unsubSearch").value.trim().toLowerCase();
    const rows = state.unsubList.filter((u) => !q || u.email.includes(q));
    $("unsubTable").innerHTML = "<tr><th>Email</th><th>Data</th><th>Origine</th><th>Da</th><th></th></tr>" +
      (rows.length ? rows.slice(0, 1000).map((u) => "<tr><td>" + M.escapeHtml(u.email) + "</td><td>" + fmtDate(u.date) + "</td><td>" + (SRC[u.source] || M.escapeHtml(u.source || "")) +
        "</td><td>" + M.escapeHtml(u.by || "") + '</td><td><button class="btn" data-e="' + M.escapeHtml(u.email) + '">Riattiva</button></td></tr>').join("")
        : '<tr><td colspan="5" class="muted">Nessun disiscritto</td></tr>');
    $("unsubTable").querySelectorAll("button").forEach((b) => b.onclick = async () => {
      if (!confirm("Riattivare " + b.dataset.e + "? Riattivalo solo se ha chiesto esplicitamente di tornare a ricevere le comunicazioni.")) return;
      try { await reg("unsubscribe-remove", { email: b.dataset.e }); toast("Indirizzo riattivato"); loadUnsubs(); checkRegistry(); } catch (e) { toast(e.message, true); }
    });
  }
  $("btnLoadCampaigns").onclick = loadCampaigns;
  $("btnLoadUnsubs").onclick = loadUnsubs;
  $("unsubSearch").addEventListener("input", renderUnsubs);
  document.querySelector('[data-group="registry"] [data-tab="unsubs"]').addEventListener("click", () => { if (!state.unsubList.length) loadUnsubs(); });
  $("btnExportUnsubs").onclick = () => downloadCsv(state.unsubList.map((u) => ({ Email: u.email, Data: u.date, Origine: u.source, InseritoDa: u.by || "" })), "disiscritti.csv");
  $("btnUnsubAdd").onclick = async () => {
    const emails = $("unsubAdd").value.split(/[\s,;]+/).filter(M.isValidEmail);
    if (!emails.length) return toast("Nessun indirizzo valido", true);
    try { const j = await reg("unsubscribe-add", { emails }); toast(j.added + " indirizzi segnati come disiscritti"); $("unsubAdd").value = ""; loadUnsubs(); checkRegistry(); }
    catch (e) { toast(e.message, true); }
  };


  // =====================================================================
  // ---------- campagne graduali ("a tappe") ----------
  // =====================================================================
  const TZS = [["Europe/Rome", "Italia"], ["Europe/Madrid", "Spagna"], ["Europe/London", "Regno Unito"], ["America/New_York", "USA – costa est"],
    ["America/Chicago", "USA – centro"], ["America/Los_Angeles", "USA – costa ovest"], ["America/Mexico_City", "Messico"], ["America/Sao_Paulo", "Brasile"],
    ["America/Argentina/Buenos_Aires", "Argentina"], ["Asia/Dubai", "Emirati Arabi"], ["Asia/Bangkok", "Thailandia / Sud-est asiatico"]];
  const DAYN = ["", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"];
  const domainOf = (e) => String(e).split("@")[1] || "";
  const CONSUMER = /^(gmail|googlemail|yahoo|ymail|hotmail|outlook|live|msn|icloud|me|mac|aol|libero|virgilio|alice|tim|tin|tiscali|fastwebnet|email|inwind|iol|gmx|web|protonmail|proton|yandex|mail|qq|163|126|uol|bol|terra|zoho)\.(com|net|it|de|fr|es|co\.uk|uk|com\.br|br|com\.ar|ar|com\.mx|mx|ru|cn|ch|at|be|nl|ca|jp|eu|me)$/i;
  const isConsumer = (d) => CONSUMER.test(d);
  function interleave(list, dom) {
    const groups = new Map();
    list.forEach((x) => { const d = dom(x); if (!groups.has(d)) groups.set(d, []); groups.get(d).push(x); });
    const gs = [...groups.values()], out = [];
    while (out.length < list.length) gs.forEach((g) => { if (g.length) out.push(g.shift()); });
    return out;
  }
  function nowIn(tz, d) {
    const p = {};
    new Intl.DateTimeFormat("en-GB", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23", weekday: "short" })
      .formatToParts(d || new Date()).forEach((x) => (p[x.type] = x.value));
    return { date: p.year + "-" + p.month + "-" + p.day, hm: p.hour + ":" + p.minute, wd: { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[p.weekday] };
  }
  const addDay = (date) => { const d = new Date(date + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + 1); return d.toISOString().slice(0, 10); };
  const fmtDay = (date) => { try { return new Date(date + "T12:00:00Z").toLocaleDateString("it-IT", { weekday: "short", day: "numeric", month: "short" }); } catch (e) { return date; } };
  const rand = (a, b) => a + Math.random() * (b - a);

  // Timer che continuano anche con la scheda in secondo piano (il browser rallenta i timer normali)
  const bgSleep = (() => {
    try {
      const w = new Worker(URL.createObjectURL(new Blob(["onmessage=function(e){setTimeout(function(){postMessage(e.data.id)},e.data.ms)}"], { type: "text/javascript" })));
      const waits = new Map(); let n = 0;
      w.onmessage = (e) => { const f = waits.get(e.data); waits.delete(e.data); if (f) f(); };
      return (ms) => new Promise((r) => { const id = ++n; waits.set(id, r); w.postMessage({ id, ms }); });
    } catch (e) { return sleep; }
  })();

  // ---- impostazioni ----
  $("gTz").innerHTML = TZS.map(([v, l]) => '<option value="' + v + '">' + l + "</option>").join("");
  $("gDays").innerHTML = [1, 2, 3, 4, 5, 6, 7].map((d) => '<label><input type="checkbox" value="' + d + '"' + (d <= 5 ? " checked" : "") + "> " + DAYN[d] + "</label>").join("");
  const G = { start: "gStart", step: "gStep", max: "gMax", perDomain: "gPerDomain", winStart: "gWinStart", winEnd: "gWinEnd", tz: "gTz", gapMin: "gGapMin", gapMax: "gGapMax" };
  (function restoreSettings() {
    const saved = store.get("gradual.settings", null); if (!saved) return;
    Object.entries(G).forEach(([k, id]) => { if (saved[k] != null) $(id).value = saved[k]; });
    if (Array.isArray(saved.days)) $("gDays").querySelectorAll("input").forEach((c) => (c.checked = saved.days.includes(+c.value)));
  })();
  function readSettings() {
    // stessi limiti applicati dal server
    const n = (id, min, max, d) => { const v = Math.round(+$(id).value); return Number.isFinite(v) && $(id).value !== "" ? Math.min(max, Math.max(min, v)) : d; };
    const s = { start: n("gStart", 5, 1500, 50), step: n("gStep", 0, 500, 25), max: n("gMax", 5, 1500, 200), perDomain: n("gPerDomain", 1, 100, 2),
      winStart: $("gWinStart").value || "09:00", winEnd: $("gWinEnd").value || "17:30", tz: $("gTz").value || "Europe/Rome",
      days: [...$("gDays").querySelectorAll("input:checked")].map((c) => +c.value),
      gapMin: n("gGapMin", 2, 600, 6), gapMax: n("gGapMax", 2, 900, 15), pauseEvery: 25 };
    if (s.start > s.max) s.start = s.max;
    if (s.gapMax < s.gapMin) s.gapMax = s.gapMin;
    return s;
  }
  function settingsProblem(s) {
    if (!s.days.length) return "Scegli almeno un giorno di invio";
    if (s.winEnd <= s.winStart) return "L'orario di fine deve essere successivo a quello di inizio";
    if (s.max > 1500) return "Massimo 1.500 email al giorno (il limite di Google Workspace è 2.000, comprese le email normali)";
    return null;
  }
  Object.values(G).forEach((id) => $(id).addEventListener("change", () => { store.set("gradual.settings", readSettings()); updateEstimate(); }));
  $("gDays").addEventListener("change", () => { store.set("gradual.settings", readSettings()); updateEstimate(); });
  function selectSendMode(mode) {
    const tab = document.querySelector('[data-group="sendmode"] [data-tab="' + mode + '"]'); if (tab) tab.click();
  }

  // ---- stima: simula le tappe con quote crescenti, giorni, fascia oraria e tetti per azienda ----
  function simulateClient(doms, s) {
    let pend = doms.slice(), t = nowIn(s.tz), date = t.date, wd = t.wd, k = 1, last = null, sendDays = 0, firstQuota = 0, maxDuration = 0;
    for (let i = 0; i < 1000 && pend.length; i++) {
      let took = 0;
      if (s.days.includes(wd) && !(i === 0 && t.hm >= s.winEnd)) {
        const quota = Math.min(s.max, s.start + s.step * (k - 1)), cnt = {}, keep = [];
        pend.forEach((d) => { if (took < quota && (isConsumer(d) || (cnt[d] || 0) < s.perDomain)) { cnt[d] = (cnt[d] || 0) + 1; took++; } else keep.push(d); });
        if (took) { pend = keep; last = date; sendDays++; k++; if (!firstQuota) firstQuota = { date, n: took }; maxDuration = Math.max(maxDuration, took); }
      }
      date = addDay(date); wd = (wd % 7) + 1;
    }
    return { eta: pend.length ? null : last, sendDays, first: firstQuota, maxPerDay: maxDuration };
  }
  function updateEstimate() {
    const box = $("gEstimate"), list = queue(), s = readSettings(), prob = settingsProblem(s);
    $("btnPlan").disabled = !list.length || !state.user || state.sending || !!prob;
    if (prob) { box.innerHTML = '<span class="warn">' + M.escapeHtml(prob) + "</span>"; return; }
    if (!list.length) { box.textContent = "Carica i destinatari per vedere la durata della campagna."; return; }
    const doms = interleave(list.map((r) => domainOf(r.email)), (d) => d), sim = simulateClient(doms, s);
    const avg = (s.gapMin + s.gapMax) / 2 + 1.5, longest = Math.ceil((sim.maxPerDay * avg + Math.floor(sim.maxPerDay / 25) * 67) / 60);
    const counts = {}; list.forEach((r) => { const d = domainOf(r.email); if (!isConsumer(d)) counts[d] = (counts[d] || 0) + 1; });
    const big = Object.entries(counts).filter(([, n]) => n > 10).sort((a, b) => b[1] - a[1]);
    const daysTxt = s.days.map((d) => DAYN[d]).join(", "), tzTxt = (TZS.find((x) => x[0] === s.tz) || [0, s.tz])[1];
    box.innerHTML = "<b>" + list.length + "</b> destinatari · " + s.start + " email il primo giorno, poi +" + s.step + " al giorno fino a " + s.max +
      " · " + daysTxt + " " + s.winStart + "–" + s.winEnd + " (ora " + M.escapeHtml(tzTxt) + ")<br>" +
      (sim.eta ? "<b>" + sim.sendDays + "</b> tappe · prima tappa " + fmtDay(sim.first.date) + " · fine prevista <b>" + fmtDay(sim.eta) + "</b>" : "Durata superiore a un anno: aumenta le quote") +
      " · la tappa più lunga dura " + (longest <= 1 ? "<b>circa un minuto</b>" : "circa <b>" + longest + " minuti</b>") + " (tieni aperta la pagina)" +
      (big.length ? '<span class="warn">' + big.length + (big.length === 1 ? " azienda ha" : " aziende hanno") + " più di 10 contatti (" +
        big.slice(0, 3).map(([d, n]) => M.escapeHtml(d) + ": " + n).join(", ") + "): con " + s.perDomain + " email al giorno per azienda allungano la campagna. Se sono reparti diversi puoi alzare il limite per azienda.</span>" : "");
  }

  // ---- chiamate al server ----
  async function tp(action, payload) {
    const token = await getToken();
    const r = await fetch("/.netlify/functions/tappe", {
      method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({ action }, payload || {}))
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error || "Server non raggiungibile (" + r.status + ")");
    return j;
  }
  async function tpRaw(method, azione, id, body) {
    const token = await getToken();
    const r = await fetch("/.netlify/functions/tappe?azione=" + azione + "&id=" + encodeURIComponent(id), {
      method, headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body
    });
    if (!r.ok) { const j = await r.json().catch(() => ({})); throw new Error(j.error || "Errore " + r.status); }
    return r.json();
  }

  // ---- il messaggio viene fissato alla creazione: le modifiche successive all'editor non cambiano la campagna ----
  function snapshot() {
    const o = composeOpts();
    return {
      v: 1, subject: $("subject").value, fromName: $("fromName").value.trim(), replyTo: $("replyTo").value.trim(),
      opts: Object.assign({}, o, { logoCid: null, canvaImages: state.canvaImages.map((im) => ({ cid: im.cid, mime: im.mime, b64: im.b64, name: im.name, link: im.link, alt: im.alt })) }),
      attachments: state.attachments.map((a) => ({ name: a.name, mime: a.mime, size: a.size, b64: a.b64 }))
    };
  }
  function fromSnapshot(snap, fields, to) {
    const o = Object.assign({}, snap.opts, { logoCid: state.logo ? "ialogo@ia" : null });
    const html = M.buildEmailHtml(o, fields), inline = [];
    if (o.useTemplate && state.logo) inline.push({ cid: "ialogo@ia", mime: "image/png", b64: state.logo, name: "logo-ia.png" });
    if (o.mode === "canva") (o.canvaImages || []).forEach((im) => inline.push({ cid: im.cid, mime: im.mime, b64: im.b64, name: im.name }));
    return { fromName: snap.fromName, fromEmail: o.fromEmail, to, replyTo: snap.replyTo, unsubUrl: fields.__unsubUrl,
      subject: M.merge(snap.subject, fields, false).trim(), html, text: M.htmlToText(html), inline, attachments: snap.attachments || [] };
  }

  // ---- creazione ----
  $("btnPlan").onclick = async () => {
    const err = validate(); if (err) return toast(err, true);
    const list = queue(); if (!list.length) return toast("Nessun destinatario da inviare", true);
    const s = readSettings(), prob = settingsProblem(s); if (prob) return toast(prob, true);
    if (!state.registryOk && !confirm("Il registro condiviso non è raggiungibile: disiscritti e storico non sono stati verificati. Procedere comunque?")) return;
    const sim = simulateClient(interleave(list.map((r) => domainOf(r.email)), (d) => d), s);
    if (!confirm("Creare la campagna graduale?\n\nOggetto: " + $("subject").value.trim() + "\nDa: " + $("fromAddress").value + "\nDestinatari: " + list.length +
      "\nTappe: " + s.start + " email il primo giorno, poi +" + s.step + " al giorno fino a " + s.max +
      "\nGiorni: " + s.days.map((d) => DAYN[d]).join(", ") + ", " + s.winStart + "–" + s.winEnd +
      "\nFine prevista: " + (sim.eta ? fmtDay(sim.eta) + " (" + sim.sendDays + " tappe)" : "oltre un anno") +
      "\n\nIl messaggio viene fissato adesso: le modifiche successive all'editor non cambiano questa campagna.")) return;
    saveSender();
    const btn = $("btnPlan"); btn.disabled = true; btn.textContent = "Creazione in corso…";
    try {
      const ordered = interleave(list, (r) => domainOf(r.email));
      const j = await tp("crea", { campaign: { subject: $("subject").value.trim(), sender: $("fromAddress").value, fromName: $("fromName").value.trim(), mode: state.mode, source: state.source },
        settings: s, recipients: ordered.map((r) => r.email) });
      const id = j.plan.id;
      await tpRaw("PUT", "contenuto", id, JSON.stringify(snapshot()));
      await tpRaw("PUT", "campi", id, JSON.stringify(Object.fromEntries(ordered.map((r) => [r.email, r.fields]))));
      await tp("stato", { id, status: "attiva" });
      toast("Campagna graduale creata: " + list.length + " destinatari");
      await loadPlans();
      $("sec-plans").scrollIntoView({ behavior: "smooth" });
    } catch (e) { toast("Creazione non riuscita: " + e.message, true); loadPlans(); }
    btn.textContent = "Crea campagna graduale"; updateEstimate();
  };

  // ---- elenco ----
  const STATUS = { attiva: "Attiva", "in pausa": "In pausa", sospesa: "Sospesa", completata: "Completata", annullata: "Annullata", preparazione: "Incompleta" };
  async function loadPlans() {
    if (!state.user) return;
    try { state.plans = (await tp("elenco")).plans || []; renderPlans(); }
    catch (e) { $("planList").innerHTML = '<p class="muted">' + M.escapeHtml(e.message) + "</p>"; }
  }
  function renderPlans() {
    const box = $("planList");
    if (!state.plans.length) { box.innerHTML = '<p class="muted">Nessuna campagna graduale. Creala dalla sezione 4, scegliendo «Graduale, a tappe».</p>'; return; }
    box.innerHTML = state.plans.map((p) => {
      const c = p.counts || {}, s = p.s || {}, sent = (c.ok || 0) + (c.bounce || 0), done = (c.total || 0) - (c.pending || 0);
      const pct = c.total ? Math.round(done / c.total * 100) : 0, bpct = sent ? Math.round((c.bounce || 0) / sent * 1000) / 10 : 0;
      const stats = ["Inviate <b>" + sent + "</b> di " + c.total, c.pending ? "in coda " + c.pending : "", c.reply ? "<b>risposte " + c.reply + "</b>" : "",
        c.auto ? "risposte automatiche " + c.auto : "", c.bounce ? "non recapitate " + c.bounce + " (" + String(bpct).replace(".", ",") + "%)" : "",
        c.unsub ? "disiscritti " + c.unsub : "", c.norec ? "esclusi perché non recapitabili " + c.norec : "", c.err ? "errori " + c.err : ""].filter(Boolean).join(" · ");
      const today = p.status === "attiva" && p.today ? p.today.message + (p.eta ? " · fine prevista " + fmtDay(p.eta) : "") : (p.reason ? M.escapeHtml(p.reason) : "");
      const b = (a, label, primary, dis) => '<button class="btn' + (primary ? " primary" : "") + '" data-a="' + a + '" data-id="' + p.id + '"' + (dis ? " disabled" : "") + ">" + label + "</button>";
      let btns = "";
      if (p.mine) {
        if (p.status === "attiva") btns = b("run", "Invia la tappa di oggi", true, !(p.today && p.today.open) || state.sending) + b("inbox", "Controlla risposte e recapiti") + b("pause", "Pausa");
        if (p.status === "in pausa" || p.status === "sospesa") btns = b("resume", "Riprendi", true) + b("inbox", "Controlla risposte e recapiti");
        if (c.err && p.status !== "annullata") btns += b("retry", "Rimetti in coda gli errori (" + c.err + ")");
      }
      btns += b("report", "Report CSV");
      if (p.mine && !["completata", "annullata"].includes(p.status)) btns += b("cancel", p.status === "preparazione" ? "Elimina" : "Annulla");
      return '<div class="plan"><h3>' + M.escapeHtml(p.subject || "(senza oggetto)") + ' <span class="badge ' + p.status.replace(" ", "-") + '">' + (STATUS[p.status] || p.status) + "</span>" +
        (p.running ? ' <span class="badge">invio in corso</span>' : "") + "</h3>" +
        '<div class="meta">da ' + M.escapeHtml(p.sender) + " · creata il " + fmtDate(p.createdAt) + " · " + s.start + "→" + s.max + " al giorno, max " + s.perDomain +
        " per azienda · " + (s.days || []).map((d) => DAYN[d]).join(" ") + " " + s.winStart + "–" + s.winEnd + "</div>" +
        '<div class="progress"><div style="width:' + pct + '%"></div></div><div class="stats">' + stats + "</div>" +
        (today ? '<div class="today">' + today + "</div>" : "") + '<div class="row">' + btns + "</div></div>";
    }).join("");
  }
  $("btnLoadPlans").onclick = loadPlans;
  $("planList").addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-a]"); if (!btn) return;
    const id = btn.dataset.id, a = btn.dataset.a, p = state.plans.find((x) => x.id === id) || {};
    if (a === "run") return runPlan(id);   // nessuna attesa prima: il clic può aprire la finestra di Google
    try {
      if (a === "pause") await tp("stato", { id, status: "in pausa" });
      if (a === "resume") {
        if (p.status === "sospesa" && !confirm("La campagna è stata sospesa: " + (p.reason || "") + "\n\nHai controllato la lista? Riprendere comunque?")) return;
        await tp("stato", { id, status: "attiva" });
      }
      if (a === "cancel" && confirm(p.status === "preparazione" ? "Eliminare questa campagna incompleta?" : "Annullare la campagna? I destinatari non ancora raggiunti non riceveranno l'email.")) await tp("stato", { id, status: "annullata" });
      if (a === "retry") { const j = await tp("riprova", { id }); toast(j.requeued + " destinatari rimessi in coda"); }
      if (a === "inbox") { await checkPlanInbox(id, false); }
      if (a === "report") await planReport(p);
    } catch (err) { toast(err.message, true); }
    loadPlans();
  });

  async function planReport(p) {
    const j = await tp("report", { id: p.id });
    let campi = {};
    if (p.mine) { try { campi = (state.planCache[p.id] && state.planCache[p.id].campi) || await tpRaw("GET", "campi", p.id); } catch (e) { /* senza nomi */ } }
    const ESITO = ["In coda", "Inviata", "Errore", "Disiscritto", "Non recapitata", "Esclusa (non recapitabile)"], RISP = ["", "Risposta", "Risposta automatica"];
    downloadCsv(j.recipients.map((r) => {
      const f = campi[r.e] || {};
      return { Email: r.e, Nome: f.Nome || "", Cognome: f.Cognome || "", Azienda: f.Azienda || "", Esito: ESITO[r.st] || "", Risposta: RISP[r.r] || "", DataOra: r.t || "", Dettaglio: r.err || "" };
    }), "campagna-graduale-" + p.createdAt.slice(0, 10) + ".csv");
  }

  // ---- controllo della casella: risposte e mancati recapiti ----
  function plog(line, cls) {
    const d = document.createElement("div"); d.className = cls || ""; d.textContent = new Date().toLocaleTimeString("it-IT") + "  " + line;
    $("planLog").prepend(d);
  }
  async function checkPlanInbox(id, quiet) {
    try {
      const j = await tp("recapiti", { id });
      if (j.needScope) {
        if (!state.scopeWarned) { state.scopeWarned = true; plog(j.message, "err"); toast(j.message, true); }
        return j;
      }
      const msg = "Controllo casella: " + j.bounces.length + " nuovi non recapitati, " + j.replies + " nuove risposte" + (j.autos ? ", " + j.autos + " risposte automatiche" : "");
      plog(msg, j.bounces.length ? "err" : ""); if (!quiet) toast(msg);
      if (j.suspended) { plog("Campagna sospesa: " + j.plan.reason, "err"); toast("Campagna sospesa: " + j.plan.reason, true); }
      return j;
    } catch (e) { plog("Controllo casella non riuscito: " + e.message, "err"); if (!quiet) toast(e.message, true); return null; }
  }

  // ---- invio della tappa di oggi ----
  let renewResolve = null;
  $("btnRenew").onclick = () => {
    getToken("").then(() => { $("renewBox").classList.add("hidden"); document.title = state.title; if (renewResolve) { renewResolve(); renewResolve = null; } })
      .catch(() => toast("Accesso a Google non rinnovato", true));
  };
  function waitRenew() {
    $("renewBox").classList.remove("hidden"); document.title = "⏸ Azione richiesta — IA Mailer";
    plog("Accesso Google in scadenza: premi «Rinnova accesso e continua»", "err");
    return new Promise((res) => { renewResolve = res; });
  }
  async function waitMs(ms) {
    const end = Date.now() + ms;
    while (Date.now() < end && !state.stopped) await bgSleep(Math.min(1000, end - Date.now()));
  }
  $("btnPlanPause").onclick = () => { state.paused = !state.paused; $("btnPlanPause").textContent = state.paused ? "Riprendi" : "Pausa"; plog(state.paused ? "In pausa" : "Ripresa"); };
  $("btnPlanStop").onclick = () => { if (confirm("Interrompere la tappa? Le email non ancora inviate partiranno alla prossima tappa.")) { state.stopped = true; if (renewResolve) { renewResolve(); renewResolve = null; } } };

  function runPlan(id) {
    if (state.sending) return toast("C'è già un invio in corso in questa pagina", true);
    // siamo ancora nel clic: se l'accesso Google scade a breve lo rinnoviamo subito (la finestra di Google può aprirsi)
    const renew = !state.token || Date.now() > state.tokenExp - 15 * 60e3 ? getToken("") : Promise.resolve();
    renew.then(() => runPlanSession(id), () => toast("Accesso a Google non rinnovato", true));
  }
  async function runPlanSession(id) {
    const p = state.plans.find((x) => x.id === id) || {};
    state.sending = true; state.paused = false; state.stopped = false; state.title = document.title;
    $("planSession").classList.remove("hidden"); $("planLog").innerHTML = ""; $("planProgressBar").style.width = "0";
    $("planSessionTitle").textContent = "Tappa di oggi — " + (p.subject || ""); $("planProgressText").textContent = "Preparazione…";
    $("btnPlanPause").textContent = "Pausa"; ["btnPlanPause", "btnPlanStop"].forEach((x) => $(x).classList.remove("hidden"));
    renderPlans(); updateEstimate();
    $("planSession").scrollIntoView({ behavior: "smooth", block: "nearest" });
    let session = null, okN = 0, koN = 0, consec = 0, note = "", wake = null;
    const pending = [];
    const flush = async (final) => {
      if (!session) return;
      const batch = pending.splice(0);
      try { await tp("esito", { id, session, results: batch, final: !!final }); }
      catch (e) { pending.unshift.apply(pending, batch); plog("Registro non aggiornato, nuovo tentativo al prossimo blocco: " + e.message, "err"); if (final) throw e; }
    };
    try {
      if (!state.planCache[id]) {
        plog("Caricamento del messaggio e dei destinatari…");
        state.planCache[id] = { snap: await tpRaw("GET", "contenuto", id), campi: await tpRaw("GET", "campi", id) };
      }
      const cache = state.planCache[id];
      plog("Controllo risposte e mancati recapiti…");
      const rc = await checkPlanInbox(id, true);
      if (rc && rc.suspended) { note = "Campagna sospesa per troppi mancati recapiti"; return; }
      const L = await tp("lotto", { id, session: Math.random().toString(36).slice(2) });
      if (!L.batch.length) { note = L.message; plog(L.message); return; }
      session = L.session;
      plog("Tappa " + L.dayIndex + ": " + L.batch.length + " email (già inviate oggi " + L.sentToday + " su " + L.quota + ")");
      try { if (navigator.wakeLock) wake = await navigator.wakeLock.request("screen"); } catch (e) { /* non supportato */ }
      for (let i = 0; i < L.batch.length; i++) {
        while (state.paused && !state.stopped) await bgSleep(300);
        if (state.stopped) { note = "Tappa interrotta: le email rimanenti partiranno alla prossima tappa"; break; }
        if (nowIn(L.tz).hm >= L.winEnd) { note = "Fascia oraria terminata alle " + L.winEnd + ": si riprende alla prossima tappa"; break; }
        if (Date.now() > state.tokenExp - 60e3) { await waitRenew(); if (state.stopped) { note = "Tappa interrotta"; break; } }
        const x = L.batch[i];
        const fields = Object.assign({ Email: x.email }, cache.campi[x.email] || {}, { __unsubUrl: x.unsubUrl });
        try {
          const g = await sendOne(fromSnapshot(cache.snap, fields, x.email));
          pending.push({ email: x.email, status: "ok", id: g.id, threadId: g.threadId, time: new Date().toISOString() });
          okN++; consec = 0; plog("✓ " + x.email, "ok");
        } catch (e) {
          if (e.fatal) { note = "Google ha segnalato il limite giornaliero della casella: si riprende domani"; plog("✕ " + e.message, "err"); break; }
          pending.push({ email: x.email, status: "error", error: e.message, time: new Date().toISOString() });
          koN++; plog("✕ " + x.email + " — " + e.message, "err");
          if (++consec >= 5) {
            note = "5 errori di invio consecutivi: campagna messa in pausa";
            await flush(); await tp("stato", { id, status: "in pausa", reason: "5 errori di invio consecutivi (ultimo: " + e.message + ")" });
            break;
          }
        }
        $("planProgressBar").style.width = Math.round((i + 1) / L.batch.length * 100) + "%";
        $("planProgressText").textContent = (i + 1) + " di " + L.batch.length + " · inviate " + okN + (koN ? " · errori " + koN : "");
        if (pending.length >= 10) await flush();
        if ((i + 1) % 50 === 0 && i < L.batch.length - 1) {
          await flush();
          const r2 = await checkPlanInbox(id, true);
          if (r2 && r2.suspended) { note = "Campagna sospesa per troppi mancati recapiti"; break; }
        }
        if (i < L.batch.length - 1) {
          let ms = rand(L.gap[0], L.gap[1]) * 1000;
          if (L.pauseEvery && (i + 1) % L.pauseEvery === 0) { ms += rand(45, 90) * 1000; plog("Pausa più lunga: " + Math.round(ms / 1000) + " secondi"); }
          await waitMs(ms);
        }
      }
    } catch (e) { note = e.message; plog("Errore: " + e.message, "err"); }
    finally {
      try { await flush(true); } catch (e) { toast("Esiti non registrati: " + e.message + ". Non ripetere la tappa prima di aver ricaricato la pagina.", true); }
      try { if (wake) await wake.release(); } catch (e) { /* già rilasciato */ }
      state.sending = false; $("renewBox").classList.add("hidden"); document.title = state.title;
      ["btnPlanPause", "btnPlanStop"].forEach((x) => $(x).classList.add("hidden"));
      const msg = (session ? "Tappa conclusa: " + okN + " email inviate" + (koN ? ", " + koN + " errori" : "") + (note ? ". " + note : "") : note) || "Nessun invio";
      $("planProgressText").textContent = msg; plog(msg); toast(msg, !!(koN && !okN));
      loadPlans(); updateSendInfo();
    }
  }

  // ---- non recapitabili (registro) ----
  async function loadBounced() {
    if (!state.user) return toast("Accedi prima con Google", true);
    try {
      state.bouncedList = (await reg("nonrecapitabili")).bounced || [];
      $("bouncedTable").innerHTML = "<tr><th>Email</th><th>Data</th><th>Rilevato da</th><th></th></tr>" + (state.bouncedList.length ? state.bouncedList.slice(0, 1000).map((u) =>
        "<tr><td>" + M.escapeHtml(u.email) + "</td><td>" + fmtDate(u.date) + "</td><td>" + M.escapeHtml(u.by || "") + '</td><td><button class="btn" data-e="' + M.escapeHtml(u.email) + '">Riattiva</button></td></tr>').join("")
        : '<tr><td colspan="4" class="muted">Nessun indirizzo non recapitabile</td></tr>');
      $("bouncedTable").querySelectorAll("button").forEach((b) => b.onclick = async () => {
        if (!confirm("Riattivare " + b.dataset.e + "? Fallo solo se sei sicuro che l'indirizzo sia corretto.")) return;
        try { await reg("nonrecapitabili-remove", { email: b.dataset.e }); toast("Indirizzo riattivato"); loadBounced(); checkRegistry(); } catch (e) { toast(e.message, true); }
      });
    } catch (e) { toast(e.message, true); }
  }
  $("btnLoadBounced").onclick = loadBounced;
  document.querySelector('[data-group="registry"] [data-tab="bounced"]').addEventListener("click", () => { if (!state.bouncedList.length) loadBounced(); });

  // ---------- avvio ----------
  fetch("assets/logo-email.png").then((r) => r.blob()).then((b) => new Promise((res) => {
    const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(",")[1]); fr.readAsDataURL(b);
  })).then((b64) => { state.logo = b64; refreshPreview(); }).catch(() => refreshPreview());
  refreshMergeOptions(); initGoogle(); refreshPreview();
})();
