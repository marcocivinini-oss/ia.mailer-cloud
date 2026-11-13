// PDF della fattura IA Inc. — generato nel browser con pdf-lib.
// Stesso impianto della fattura HTML: logo originale IA Inc., blocco TO a destra,
// tabella servizi, totali, termini e coordinate bancarie. Formato US Letter.
// Il modulo si carica solo quando serve (import dinamico), così non pesa sull'avvio della Suite.
import { PDFDocument, rgb, degrees } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";

const NAVY = rgb(41 / 255, 56 / 255, 141 / 255);   // blu del logo IA Inc.
const RED = rgb(179 / 255, 3 / 255, 60 / 255);     // cornice del logo IA Inc.
const INK = rgb(0.2, 0.2, 0.2);
const MUTED = rgb(0.45, 0.47, 0.55);

const PAGE_W = 612, PAGE_H = 792, M = 43.2, W = PAGE_W - 2 * M;

let assetsPromise = null;
function loadAssets() {
  if (!assetsPromise) {
    const get = (u) => fetch(u, { credentials: "same-origin" }).then(r => { if (!r.ok) throw new Error("Asset non trovato: " + u); return r.arrayBuffer(); });
    assetsPromise = Promise.all([
      get("/logo-inc-fattura.png"),
      get("/fonts/poppins-latin-400-normal.woff"),
      get("/fonts/poppins-latin-600-normal.woff"),
      get("/fonts/poppins-latin-700-normal.woff"),
    ]).catch(e => { assetsPromise = null; throw e; });
  }
  return assetsPromise;
}

const fmtUSD = (n) => "$ " + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDateUS = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || "")); return m ? `${m[2]}/${m[3]}/${m[1]}` : String(iso || ""); };

function incTotals(inv) {
  const services = (inv.services || []).filter(s => String(s.description || "").trim() || Number(s.amount));
  const servicesTotal = services.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const adminCost = servicesTotal * (Number(inv.adminCostPercent) || 0) / 100;
  const subtotal = servicesTotal + adminCost;
  const taxAmount = subtotal * (Number(inv.taxPercent) || 0) / 100;
  const otherAmount = Number(inv.otherAmount) || 0;
  return { services, servicesTotal, adminCost, subtotal, taxAmount, otherAmount, total: subtotal + taxAmount + otherAmount };
}

/**
 * @param inv  fattura: { number, date, client, clientAddress, services[], adminCostPercent, taxPercent, otherAmount, terms, status }
 * @param info dati fissi della sede (IA_INC_INFO)
 * @param opts { preview: true } → filigrana "PREVIEW" e numero non ancora assegnato
 * @returns Uint8Array con il PDF
 */
export async function buildIncInvoicePdf(inv, info, opts = {}) {
  const [logoBytes, regBytes, semiBytes, boldBytes] = await loadAssets();
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const fReg = await doc.embedFont(regBytes, { subset: true });
  const fSemi = await doc.embedFont(semiBytes, { subset: true });
  const fBold = await doc.embedFont(boldBytes, { subset: true });
  const logo = await doc.embedPng(logoBytes);
  const preview = !!opts.preview;
  const number = preview ? "" : String(inv.number || "");

  doc.setTitle(preview ? "Invoice preview" : `Invoice ${number}`);
  doc.setAuthor(info.companyName);
  doc.setSubject(`Invoice to ${inv.client || ""}`);
  doc.setCreator("IA Suite");
  doc.setProducer("IA Suite");

  // Caratteri non presenti nel font: provo la lettera senza accento, altrimenti "?"
  const clean = (font, s) => {
    const kit = font.embedder.font;
    return Array.from(String(s ?? "").replace(/\r/g, "").replace(/\t/g, " ")).map(ch => {
      if (ch === "\n") return ch;
      const cp = ch.codePointAt(0);
      if (kit.hasGlyphForCodePoint(cp)) return ch;
      const base = ch.normalize("NFD").replace(/[̀-ͯ]/g, "");
      if (base && Array.from(base).every(c => kit.hasGlyphForCodePoint(c.codePointAt(0)))) return base;
      const special = { "Ł": "L", "ł": "l", "Đ": "D", "đ": "d", "Ħ": "H", "ħ": "h", "ı": "i" }[ch];
      if (special) return special;
      if (/[‘’]/.test(ch)) return "'";
      if (/[“”]/.test(ch)) return '"';
      return "?";
    }).join("");
  };
  const width = (font, s, size) => font.widthOfTextAtSize(clean(font, s), size);
  const wrap = (font, s, size, maxW) => {
    const out = [];
    for (const para of clean(font, s).split("\n")) {
      let line = "";
      for (const word of para.split(/\s+/).filter(Boolean)) {
        const test = line ? line + " " + word : word;
        if (font.widthOfTextAtSize(test, size) <= maxW) { line = test; continue; }
        if (line) out.push(line);
        // parola più lunga della colonna: la spezzo
        let w = word;
        while (font.widthOfTextAtSize(w, size) > maxW) {
          let k = w.length - 1;
          while (k > 1 && font.widthOfTextAtSize(w.slice(0, k), size) > maxW) k--;
          out.push(w.slice(0, k)); w = w.slice(k);
        }
        line = w;
      }
      out.push(line);
    }
    return out.length ? out : [""];
  };

  const pages = [];
  let page, y; // y = distanza dall'alto
  const text = (s, x, yTop, font, size, color = INK, align = "left") => {
    const t = clean(font, s);
    const w = font.widthOfTextAtSize(t, size);
    const xx = align === "right" ? x - w : align === "center" ? x - w / 2 : x;
    page.drawText(t, { x: xx, y: PAGE_H - yTop - size * 0.8, size, font, color });
    return w;
  };
  const dotted = (x1, x2, yTop) => page.drawLine({ start: { x: x1, y: PAGE_H - yTop }, end: { x: x2, y: PAGE_H - yTop }, thickness: 0.7, color: RED, dashArray: [1.2, 1.6] });
  const watermark = (label) => {
    const size = 64, w = fBold.widthOfTextAtSize(label, size);
    page.drawText(label, { x: PAGE_W / 2 - (w / 2) * Math.cos(Math.PI / 6), y: PAGE_H / 2 - (w / 2) * Math.sin(Math.PI / 6), size, font: fBold, color: rgb(0.85, 0.86, 0.9), rotate: degrees(30), opacity: 0.55 });
  };
  const newPage = () => {
    page = doc.addPage([PAGE_W, PAGE_H]); pages.push(page);
    if (preview) watermark("PREVIEW");
    else if (inv.status === "void") watermark("VOID");
  };

  // ---------- Intestazione ----------
  newPage();
  const logoH = 61, logoW = logoH * logo.width / logo.height;
  page.drawImage(logo, { x: M, y: PAGE_H - M - logoH, width: logoW, height: logoH });
  let yl = M + logoH + 14;
  text(info.companyName, M, yl, fSemi, 8.8); yl += 13;
  for (const l of [info.address1, info.address2, info.email, info.phone]) { text(l, M, yl, fReg, 8.3); yl += 11.5; }

  // Riquadro TO
  const toLines = wrap(fReg, inv.clientAddress || "", 9.2, 230);
  const nameLines = wrap(fSemi, inv.client || "", 9.5, 230);
  const boxW = Math.max(190, ...nameLines.map(l => fSemi.widthOfTextAtSize(l, 9.5) + 30), ...toLines.map(l => fReg.widthOfTextAtSize(l, 9.2) + 30));
  const boxX = M + W - boxW, boxTop = M + 15;
  let yb = boxTop + 6;
  text("TO", boxX + 10, yb, fBold, 9); yb += 14;
  for (const l of nameLines) { text(l, M + W - 12, yb, fSemi, 9.5, NAVY, "right"); yb += 13; }
  for (const l of toLines) { if (l) { text(l, M + W - 12, yb, fReg, 9.2, NAVY, "right"); } yb += 12.5; }
  const boxH = yb - boxTop + 4;
  page.drawRectangle({ x: boxX, y: PAGE_H - boxTop - boxH, width: boxW, height: boxH, borderColor: INK, borderWidth: 0.8 });

  y = Math.max(yl, boxTop + boxH) + 30;
  const title = preview ? "Invoice # (assigned when issued)" : `Invoice # ${number}`;
  text(title, M + W / 2, y, fReg, 13, NAVY, "center");
  y += 40;

  // ---------- Tabella servizi ----------
  const cDate = M, cDesc = M + W * 0.2, cAmtR = M + W, descW = W * 0.55 - 10;
  const tableHead = () => {
    text("INVOICE DATE", cDate, y, fBold, 8.3, NAVY);
    text("SERVICE DESCRIPTION", cDesc, y, fBold, 8.3, NAVY);
    text("AMOUNT", cAmtR, y, fBold, 8.3, NAVY, "right");
    y += 20;
  };
  const bottomLimit = PAGE_H - M - 30;
  tableHead();
  const t = incTotals(inv);
  const rows = t.services.map((s, i) => ({ date: i === 0 ? fmtDateUS(inv.date) : "", desc: s.description, amt: fmtUSD(Number(s.amount) || 0) }));
  if (t.adminCost > 0) rows.push({ date: "", desc: ` + ${Number(inv.adminCostPercent) || 0}% administrative cost`, amt: fmtUSD(t.adminCost) });
  for (const r of rows) {
    const lines = wrap(fReg, r.desc, 9.2, descW);
    const h = lines.length * 12.5 + 10;
    if (y + h > bottomLimit) { newPage(); y = M + 10; tableHead(); }
    if (r.date) text(r.date, cDate + 2, y, fReg, 9.2);
    lines.forEach((l, k) => text(l, cDesc + 2, y + k * 12.5, fReg, 9.2));
    text(r.amt, cAmtR - 2, y, fReg, 9.2, INK, "right");
    y += h - 4;
    dotted(cDesc, cAmtR, y);
    y += 8;
  }

  // ---------- Totali ----------
  const totRows = [
    ["SUBTOTAL", fmtUSD(t.subtotal), true],
    ["TAX", `${(Number(inv.taxPercent) || 0).toFixed(2)}%`, false],
    ["OTHER", t.otherAmount ? fmtUSD(t.otherAmount) : "-", false],
    ["TOTAL", fmtUSD(t.total), true],
  ];
  if (y + totRows.length * 17 + 20 > bottomLimit) { newPage(); y = M + 10; }
  y += 10;
  for (const [l, v, strong] of totRows) {
    text(l, M + W * 0.8, y, fBold, 8.8, NAVY, "right");
    text(v, cAmtR - 2, y, strong ? fBold : fReg, 9.2, INK, "right");
    y += 17;
  }

  // ---------- Termini e pagamento ----------
  if (y + 150 > bottomLimit + 30) { newPage(); y = M + 10; }
  y += 38;
  const termsTxt = inv.terms || "Due upon receipt";
  const lw = width(fBold, "Terms", 9.2), tw = width(fReg, termsTxt, 9.2), gap = 30;
  const tx = M + W / 2 - (lw + gap + tw) / 2;
  text("Terms", tx, y, fBold, 9.2, NAVY);
  text(termsTxt, tx + lw + gap, y, fReg, 9.2);
  y += 38;
  text("PAYMENT INFORMATION", cAmtR, y, fBold, 9, NAVY, "right"); y += 14;
  text(info.bank, cAmtR, y, fBold, 9, NAVY, "right"); y += 14;
  for (const l of [`Account number ${info.bankAccount}`, `Routing ${info.bankRouting}`, `SWIFT ${info.bankSwift}`]) { text(l, cAmtR, y, fReg, 9, INK, "right"); y += 13.5; }

  // Piè di pagina (solo se più pagine)
  if (pages.length > 1) {
    pages.forEach((p, i) => {
      const s = `${info.companyName} · Invoice ${number || "preview"} · Page ${i + 1} of ${pages.length}`;
      const c = clean(fReg, s);
      p.drawText(c, { x: PAGE_W / 2 - fReg.widthOfTextAtSize(c, 7.5) / 2, y: M / 2, size: 7.5, font: fReg, color: MUTED });
    });
  }
  return await doc.save();
}
