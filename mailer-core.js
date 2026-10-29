/* IA Mailer — funzioni di base (senza DOM), usate dal browser e dai test */
(function (root) {
  "use strict";

  const EMAIL_RE = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+\.[a-z]{2,}$/i;
  const isValidEmail = (e) => EMAIL_RE.test(String(e || "").trim());

  const escapeHtml = (s) =>
    String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  const normKey = (k) =>
    String(k || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");

  // Sinonimi: il file può avere "First Name", "firstname", "Nome"... tutti diventano {{Nome}}
  const ALIASES = {
    nome: ["firstname", "first", "nome", "name", "givenname"],
    cognome: ["lastname", "last", "cognome", "surname", "familyname"],
    azienda: ["company", "azienda", "societa", "companyname", "ragionesociale", "organization"],
    email: ["email", "mail", "emailaddress", "indirizzoemail", "posta"]
  };
  const CANONICAL = { nome: "Nome", cognome: "Cognome", azienda: "Azienda", email: "Email" };

  /** Trasforma righe {colonna: valore} in destinatari {email, fields} con deduplica */
  function normalizeRows(rows) {
    const out = [], seen = new Set(), invalid = [];
    let dup = 0;
    if (!rows || !rows.length) return { recipients: out, invalid, duplicates: 0, columns: [] };
    const headers = Object.keys(rows[0]);
    let emailCol = headers.find((h) => ALIASES.email.includes(normKey(h)));
    if (!emailCol) emailCol = headers.find((h) => /mail/i.test(h));
    if (!emailCol) emailCol = headers.find((h) => rows.slice(0, 20).some((r) => isValidEmail(r[h])));
    const columns = new Set();
    rows.forEach((r) => {
      const raw = String(r[emailCol] || "").trim().replace(/^mailto:/i, "");
      if (!raw) return;
      const email = raw.toLowerCase();
      if (!isValidEmail(email)) { invalid.push(raw); return; }
      if (seen.has(email)) { dup++; return; }
      seen.add(email);
      const fields = {};
      Object.keys(r).forEach((h) => {
        const v = r[h] == null ? "" : String(r[h]).trim();
        fields[h] = v; columns.add(h);
        const nk = normKey(h);
        Object.keys(ALIASES).forEach((canon) => {
          if (ALIASES[canon].includes(nk) && !fields[CANONICAL[canon]]) {
            fields[CANONICAL[canon]] = v; columns.add(CANONICAL[canon]);
          }
        });
      });
      fields.Email = email;
      out.push({ email, fields });
    });
    columns.add("Email");
    return { recipients: out, invalid, duplicates: dup, columns: [...columns] };
  }

  /** Elenco incollato: "email" oppure "email; Nome; Azienda" */
  function parsePasted(text) {
    const rows = String(text || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean).map((l) => {
      const p = l.split(/[;\t,]/).map((x) => x.trim());
      return { Email: p[0], Nome: p[1] || "", Azienda: p[2] || "" };
    });
    return normalizeRows(rows);
  }

  /** Esclusioni: indirizzi completi o domini (@dominio.com) */
  function isExcluded(email, exclusionsText) {
    const e = String(email).toLowerCase();
    return String(exclusionsText || "").split(/[\r\n,;]+/).map((x) => x.trim().toLowerCase()).filter(Boolean)
      .some((x) => (x.startsWith("@") ? e.endsWith(x) : e === x));
  }

  /** Sostituisce {{Campo}} e {{Campo|alternativa}}; escape HTML se richiesto */
  function merge(template, fields, html) {
    const map = {};
    Object.keys(fields || {}).forEach((k) => (map[normKey(k)] = fields[k]));
    return String(template || "").replace(/\{\{\s*([^}|]+?)\s*(?:\|\s*([^}]*?)\s*)?\}\}/g, (m, key, fb) => {
      const v = map[normKey(key)];
      const val = v != null && String(v).trim() !== "" ? String(v) : (fb != null ? fb : "");
      return html ? escapeHtml(val) : val;
    });
  }

  /** {{Nome}} dentro l'HTML dell'editor può contenere entità: le ripuliamo prima del merge */
  function unescapeMergeTags(html) {
    return String(html || "").replace(/\{\{([^}]*)\}\}/g, (m, inner) =>
      "{{" + inner.replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/<[^>]+>/g, "") + "}}");
  }

  function textToHtml(t) {
    return String(t || "").split(/\n{2,}/).map((p) => p.trim()).filter(Boolean)
      .map((p) => '<p style="margin:0 0 14px">' + escapeHtml(p).replace(/\n/g, "<br>") + "</p>").join("");
  }

  /** Stili inline per il testo dell'editor (molti client email ignorano i CSS esterni) */
  function inlineEditorStyles(html) {
    return String(html || "")
      .replace(/<p(\s[^>]*)?>/gi, '<p style="margin:0 0 14px">')
      .replace(/<h2(\s[^>]*)?>/gi, '<h2 style="margin:18px 0 10px;font-size:20px;line-height:1.3;color:#131D4B;font-family:Arial,sans-serif">')
      .replace(/<ul(\s[^>]*)?>/gi, '<ul style="margin:0 0 14px;padding-left:22px">')
      .replace(/<ol(\s[^>]*)?>/gi, '<ol style="margin:0 0 14px;padding-left:22px">')
      .replace(/<a\s+href=/gi, '<a style="color:#0E2289" href=');
  }

  /**
   * Costruisce l'HTML completo dell'email.
   * opts: {mode:'text'|'canva'|'html', bodyHtml, canvaIntro, canvaImages:[{cid,link,alt}], canvaWebLink,
   *        preheader, useTemplate, addUnsub, cta:{text,url}, signature, fromEmail, footer, website, logoCid}
   */
  function buildEmailHtml(o, fields) {
    const M = (s) => merge(s, fields, true);
    let content = "";
    if (o.mode === "canva") {
      if (o.canvaIntro) content += '<div style="padding:24px 32px 8px">' + M(textToHtml(o.canvaIntro)) + "</div>";
      content += (o.canvaImages || []).map((im) => {
        const img = '<img src="cid:' + im.cid + '" width="600" alt="' + escapeHtml(im.alt || "") +
          '" style="display:block;width:100%;max-width:600px;height:auto;border:0">';
        return im.link ? '<a href="' + escapeHtml(im.link) + '" target="_blank">' + img + "</a>" : img;
      }).join("");
    } else if (o.mode === "html") {
      content = merge(unescapeMergeTags(o.bodyHtml), fields, true);
      if (!o.useTemplate) return wrapRaw(content, o, fields);
    } else {
      content = '<div style="padding:28px 32px 8px">' + inlineEditorStyles(M(unescapeMergeTags(o.bodyHtml))) + "</div>";
    }

    if (o.cta && o.cta.text && o.cta.url) {
      content += '<div style="padding:8px 32px 20px"><table role="presentation" cellpadding="0" cellspacing="0"><tr>' +
        '<td bgcolor="#F57547" style="background:#F57547"><a href="' + escapeHtml(o.cta.url) +
        '" target="_blank" style="display:inline-block;padding:12px 26px;color:#ffffff;font-family:Arial,sans-serif;font-size:15px;font-weight:bold;text-decoration:none">' +
        M(o.cta.text) + "</a></td></tr></table></div>";
    }
    if (o.signature) {
      const lines = M(o.signature).split(/\n/);
      content += '<div style="padding:4px 32px 28px;font-size:14px;line-height:1.5">' +
        '<strong style="color:#131D4B">' + lines[0] + "</strong>" +
        (lines.length > 1 ? '<br><span style="color:#666c80">' + lines.slice(1).join("<br>") + "</span>" : "") + "</div>";
    }

    if (!o.useTemplate) return wrapRaw(content, o, fields);

    const webLink = o.mode === "canva" && o.canvaWebLink
      ? '<div style="text-align:center;font-family:Arial,sans-serif;font-size:12px;color:#666c80;padding:8px 0">Non visualizzi correttamente questa email? <a href="' +
        escapeHtml(o.canvaWebLink) + '" style="color:#0E2289">Apri la versione online</a></div>' : "";

    return '<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<meta name="color-scheme" content="light"><title></title></head>' +
      '<body style="margin:0;padding:0;background:#F4F5F9">' + preheaderHtml(o.preheader, fields) +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#F4F5F9"><tr><td align="center" style="padding:16px 8px">' +
      webLink +
      '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.6;color:#1d2233">' +
      '<tr><td style="padding:22px 32px 18px;border-bottom:3px solid #F57547">' +
      (o.logoCid ? '<img src="cid:' + o.logoCid + '" width="170" alt="International Advisors" style="display:block;width:170px;height:auto;border:0">'
                 : '<strong style="color:#0E2289;font-size:18px">INTERNATIONAL ADVISORS</strong>') +
      "</td></tr>" +
      "<tr><td>" + content + "</td></tr>" +
      '<tr><td bgcolor="#131D4B" style="background:#131D4B;padding:20px 32px;color:#c9cde0;font-size:12px;line-height:1.6">' +
      (o.footer || "") +
      (o.website ? '<br><a href="' + escapeHtml(o.website) + '" style="color:#F57547;text-decoration:none">' + escapeHtml(o.website.replace(/^https?:\/\//, "")) + "</a>" : "") +
      (o.addUnsub ? '<br><br>' + unsubHtml(o.fromEmail, "#ffffff", fields && fields.__unsubUrl) : "") +
      "</td></tr></table></td></tr></table></body></html>";
  }

  function preheaderHtml(p, fields) {
    return p ? '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent">' +
      merge(p, fields, true) + "&#847;&zwnj;&nbsp;".repeat(30) + "</div>" : "";
  }
  function unsubHtml(from, color, url) {
    const href = url || "mailto:" + from + "?subject=Cancellami";
    return "Ricevi questa email perché sei in contatto con IA International Advisors. Se non desideri più ricevere comunicazioni, " +
      '<a href="' + escapeHtml(href) + '" style="color:' + color + '">annulla l\'iscrizione</a>' + (url ? "." : " o rispondi con «CANCELLAMI».");
  }
  function wrapRaw(content, o, fields) {
    return '<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:0">' + preheaderHtml(o.preheader, fields) +
      content + (o.addUnsub ? '<p style="font-family:Arial,sans-serif;font-size:11px;color:#888;padding:16px">' + unsubHtml(o.fromEmail, "#888", fields && fields.__unsubUrl) + "</p>" : "") +
      "</body></html>";
  }

  /** Versione solo testo (migliora la consegna e l'accessibilità) */
  function htmlToText(html) {
    return String(html || "")
      .replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, "")
      .replace(/<div style="display:none[\s\S]*?<\/div>/i, "")
      .replace(/<a [^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, (m, href, txt) => {
        const t = txt.replace(/<[^>]+>/g, "").trim();
        return href.startsWith("mailto:") || !t ? t : t + " (" + href + ")";
      })
      .replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|h\d|li|tr)>/gi, "\n").replace(/<li[^>]*>/gi, "• ")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#847;|&zwnj;/g, "")
      .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  // ---------- MIME ----------
  function utf8ToB64(str) {
    if (typeof Buffer !== "undefined") return Buffer.from(str, "utf8").toString("base64");
    const bytes = new TextEncoder().encode(str);
    let bin = ""; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  const wrap76 = (b64) => b64.replace(/(.{76})/g, "$1\r\n");
  const needsEnc = (s) => /[^\x20-\x7e]/.test(s);
  function encodeWord(s) { return needsEnc(s) ? "=?UTF-8?B?" + utf8ToB64(s) + "?=" : s; }
  function encodeSubject(s) {
    if (!needsEnc(s)) return s;
    // spezza in blocchi per rispettare la lunghezza delle righe
    const chars = Array.from(s), parts = [];
    for (let i = 0; i < chars.length; i += 30) parts.push("=?UTF-8?B?" + utf8ToB64(chars.slice(i, i + 30).join("")) + "?=");
    return parts.join("\r\n ");
  }
  function address(name, email) {
    if (!name) return email;
    const n = needsEnc(name) ? encodeWord(name) : '"' + name.replace(/["\\]/g, "") + '"';
    return n + " <" + email + ">";
  }
  const boundary = (tag) => "=_ia_" + tag + "_" + Math.random().toString(36).slice(2, 12);

  /**
   * msg: {fromName, fromEmail, to, replyTo, subject, html, text, inline:[{cid,mime,b64,name}], attachments:[{name,mime,b64}]}
   * Restituisce il messaggio RFC 822 come stringa
   */
  function buildMime(msg) {
    const CRLF = "\r\n";
    const h = [
      "From: " + address(msg.fromName, msg.fromEmail),
      "To: " + msg.to,
      "Subject: " + encodeSubject(msg.subject || ""),
      "MIME-Version: 1.0",
      "List-Unsubscribe: " + (msg.unsubUrl ? "<" + msg.unsubUrl + ">, " : "") + "<mailto:" + msg.fromEmail + "?subject=Cancellami>"
    ];
    if (msg.unsubUrl) h.push("List-Unsubscribe-Post: List-Unsubscribe=One-Click");
    if (msg.replyTo) h.push("Reply-To: " + msg.replyTo);

    const altB = boundary("alt");
    const alt = "Content-Type: multipart/alternative; boundary=\"" + altB + "\"" + CRLF + CRLF +
      "--" + altB + CRLF + "Content-Type: text/plain; charset=UTF-8" + CRLF + "Content-Transfer-Encoding: base64" + CRLF + CRLF +
      wrap76(utf8ToB64(msg.text || "")) + CRLF +
      "--" + altB + CRLF + "Content-Type: text/html; charset=UTF-8" + CRLF + "Content-Transfer-Encoding: base64" + CRLF + CRLF +
      wrap76(utf8ToB64(msg.html || "")) + CRLF + "--" + altB + "--";

    let body = alt;
    if (msg.inline && msg.inline.length) {
      const relB = boundary("rel");
      body = "Content-Type: multipart/related; boundary=\"" + relB + "\"" + CRLF + CRLF +
        "--" + relB + CRLF + alt + CRLF +
        msg.inline.map((p) => "--" + relB + CRLF + "Content-Type: " + p.mime + "; name=\"" + p.name + "\"" + CRLF +
          "Content-Transfer-Encoding: base64" + CRLF + "Content-ID: <" + p.cid + ">" + CRLF +
          "Content-Disposition: inline; filename=\"" + p.name + "\"" + CRLF + CRLF + wrap76(p.b64) + CRLF).join("") +
        "--" + relB + "--";
    }
    if (msg.attachments && msg.attachments.length) {
      const mixB = boundary("mix");
      body = "Content-Type: multipart/mixed; boundary=\"" + mixB + "\"" + CRLF + CRLF +
        "--" + mixB + CRLF + body + CRLF +
        msg.attachments.map((a) => "--" + mixB + CRLF + "Content-Type: " + (a.mime || "application/octet-stream") +
          "; name=\"" + encodeWord(a.name) + "\"" + CRLF + "Content-Transfer-Encoding: base64" + CRLF +
          "Content-Disposition: attachment; filename=\"" + encodeWord(a.name) + "\"" + CRLF + CRLF + wrap76(a.b64) + CRLF).join("") +
        "--" + mixB + "--";
    }
    return h.join(CRLF) + CRLF + body + CRLF;
  }

  function toCsv(rows) {
    if (!rows.length) return "";
    const cols = Object.keys(rows[0]);
    const q = (v) => '"' + String(v == null ? "" : v).replace(/"/g, '""') + '"';
    return [cols.map(q).join(";")].concat(rows.map((r) => cols.map((c) => q(r[c])).join(";"))).join("\r\n");
  }

  const api = { isValidEmail, escapeHtml, normalizeRows, parsePasted, isExcluded, merge, unescapeMergeTags,
    textToHtml, buildEmailHtml, htmlToText, buildMime, utf8ToB64, encodeSubject, toCsv };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.IAMailer = api;
})(typeof window !== "undefined" ? window : globalThis);
