/* IA Mailer — interfaccia e invio tramite Gmail API */
(function () {
  "use strict";
  const C = window.IA_MAILER_CONFIG || {};
  const M = window.IAMailer;
  const $ = (id) => document.getElementById(id);
  const SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile"
  ].join(" ");

  const state = {
    token: null, tokenClient: null, tokenWaiters: [], user: null,
    mode: "text", source: "file",
    recipients: [], columns: [], included: new Set(),
    canvaImages: [], attachments: [], logo: null,
    results: new Map(), sending: false, paused: false, stopped: false,
    unsub: {}, history: {}, registryOk: false, unsubList: []
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
    refreshPreview(); updateSendInfo(); loadCampaigns();
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
    state.unsub = {}; state.history = {};
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
      const j = await reg("check", { emails: state.recipients.map((r) => r.email) });
      state.unsub = j.unsubscribed || {}; state.history = j.history || {}; state.registryOk = true;
      const nU = Object.keys(state.unsub).length, nH = Object.keys(state.history).length;
      $("recSummary").innerHTML = state.baseSummary + (nU ? " · <b>" + nU + "</b> disiscritti esclusi" : "") +
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
    for (let attempt = 0; attempt < 4; attempt++) {
      const r = await gfetch("https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media", {
        method: "POST", headers: { "Content-Type": "message/rfc822" }, body: mime
      });
      if (r.ok) return (await r.json()).id;
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
    const n = queue().length;
    $("btnSend").disabled = !n || !state.user || state.sending;
    $("btnSend").textContent = n ? "Invia a " + n + " destinatari" : "Invia a tutti";
    if (!state.recipients.length) return;
    const mins = Math.ceil(n * ((C.SEND_DELAY_MS || 1500) + 600) / 60000);
    const skipped = state.recipients.filter((r) => skipReason(r)).length;
    $("sendInfo").innerHTML = (skipped ? skipped + " destinatari esclusi (disiscritti, esclusioni o già contattati). " : "") + "<b>" + n + "</b> email personalizzate, una per destinatario (nessuno vede gli altri indirizzi). Tempo stimato: circa " + mins +
      " min. Tieni aperta questa pagina durante l'invio. Limite Google Workspace: 2.000 email al giorno per casella.";
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
        const id = await sendOne(buildMessage(fields, r.email));
        const res = { status: "ok", id, time: new Date().toISOString() };
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

  // ---------- avvio ----------
  fetch("assets/logo-email.png").then((r) => r.blob()).then((b) => new Promise((res) => {
    const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(",")[1]); fr.readAsDataURL(b);
  })).then((b64) => { state.logo = b64; refreshPreview(); }).catch(() => refreshPreview());
  refreshMergeOptions(); initGoogle(); refreshPreview();
})();
