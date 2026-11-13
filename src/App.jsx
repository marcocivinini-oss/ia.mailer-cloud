import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import mammoth from "mammoth";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  LayoutDashboard, FileSignature, FileUp, ArrowDownToLine, ArrowUpFromLine,
  Landmark, Settings, Upload, Trash2, Check, X, Plus, RefreshCw, LogOut,
  AlertTriangle, Link2, Unlink, Loader2, Pencil, ChevronRight, CircleDollarSign, Receipt,
  Users, Briefcase, Calendar, ListChecks, Phone, Mail, MessageSquare, StickyNote, Building2, UserCircle, Clock, BarChart3, Download, Globe, ArrowLeft, Menu,
  BookOpen, Ban, FileDown, History, Search, FileArchive, Eye
} from "lucide-react";
import { kvGet, kvSet, kvDelete, extractDoc, logout as accessLogout, regList, regIssue, regUpdate, regVoid, regSetNext, regImport, regUploadPdf, regPdfUrl } from "./db";
import { __t, IS_INC } from "./i18n.js";

/* ============================================================
   IA — Suite Fatturazione & Riconciliazione (prototipo)
   Valida il flusso: login a ruoli, portale upload con estrazione
   automatica, compilazione mensile fatture attive/passive, import
   CSV Banco BPM e riconciliazione pagamenti.
   ============================================================ */

// ---- Brand ----
const C = {
  navy: "#131D4B", blue: "#0E2289", logoblue: "#132389", orange: "#F57547",
  bg: "#F5F6F8", panel: "#FFFFFF", ink: "#1B2033", muted: "#6B7280",
  line: "#E4E7EE", green: "#1F8A5B", amber: "#B7791F", red: "#C0392B",
};

const STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
.ia-app *{ box-sizing:border-box; }
.ia-app{ font-family:"Calibri","Segoe UI",system-ui,-apple-system,sans-serif; color:${C.ink};
  background:${C.bg}; -webkit-font-smoothing:antialiased; }
.ia-h{ font-family:"Poppins",sans-serif; font-weight:600; letter-spacing:.2px; }
.ia-panel{ background:${C.panel}; border:1px solid ${C.line}; }
.ia-bracket{ position:relative; }
.ia-bracket::before,.ia-bracket::after{ content:""; position:absolute; width:16px; height:16px; }
.ia-bracket::before{ top:-1px; left:-1px; border-top:3px solid ${C.orange}; border-left:3px solid ${C.orange}; }
.ia-bracket::after{ bottom:-1px; right:-1px; border-bottom:3px solid ${C.orange}; border-right:3px solid ${C.orange}; }
.ia-btn{ font-family:"Poppins",sans-serif; font-weight:500; font-size:13px; border:1px solid transparent;
  padding:8px 14px; cursor:pointer; display:inline-flex; align-items:center; gap:7px; transition:filter .15s, background .15s; }
.ia-btn:disabled{ opacity:.4; cursor:not-allowed; }
.ia-btn-primary{ background:${C.orange}; color:#fff; }
.ia-btn-primary:hover:not(:disabled){ filter:brightness(.94); }
.ia-btn-dark{ background:${C.navy}; color:#fff; }
.ia-btn-dark:hover:not(:disabled){ filter:brightness(1.15); }
.ia-btn-ghost{ background:transparent; color:${C.navy}; border:1px solid ${C.line}; }
.ia-btn-ghost:hover:not(:disabled){ background:#EEF0F5; }
.ia-btn-danger{ background:transparent; color:${C.red}; border:1px solid #E7C9C9; }
.ia-btn-danger:hover:not(:disabled){ background:#FBECEC; }
.ia-input{ font-family:inherit; font-size:14px; padding:8px 10px; border:1px solid ${C.line};
  background:#fff; color:${C.ink}; width:100%; }
.ia-input:focus{ outline:2px solid ${C.blue}; outline-offset:-1px; }
.ia-table{ width:100%; border-collapse:collapse; font-size:13.5px; }
.ia-table th{ font-family:"Poppins",sans-serif; font-weight:600; font-size:11.5px; text-transform:uppercase;
  letter-spacing:.4px; color:${C.muted}; text-align:left; padding:10px 12px; border-bottom:1px solid ${C.line}; background:#FafBfc; }
.ia-table td{ padding:11px 12px; border-bottom:1px solid ${C.line}; vertical-align:middle; }
.ia-table tr:hover td{ background:#FAFBFD; }
.ia-nav{ display:flex; align-items:center; gap:11px; padding:10px 16px; font-family:"Poppins",sans-serif;
  font-size:13.5px; font-weight:500; color:#C7CBDD; cursor:pointer; border-left:3px solid transparent; }
.ia-nav:hover{ color:#fff; background:rgba(255,255,255,.05); }
.ia-nav-active{ color:#fff; background:rgba(245,117,71,.14); border-left:3px solid ${C.orange}; }
.ia-chip{ font-family:"Poppins",sans-serif; font-size:11px; font-weight:600; padding:3px 9px; display:inline-flex; align-items:center; gap:5px; }
.ia-drop{ border:2px dashed #C4CAD9; background:#FBFCFE; padding:34px; text-align:center; transition:.15s; }
.ia-drop.drag{ border-color:${C.orange}; background:#FFF6F1; }
.ia-modal-bg{ position:fixed; inset:0; background:rgba(19,29,75,.55); display:flex; align-items:center;
  justify-content:center; padding:20px; z-index:50; }
.ia-kpi-num{ font-family:"Poppins",sans-serif; font-weight:700; font-size:30px; line-height:1; color:${C.navy}; }
.ia-scroll::-webkit-scrollbar{ height:9px; width:9px; } .ia-scroll::-webkit-scrollbar-thumb{ background:#C9CEDC; }
.ia-menu-btn{ display:none; background:none; border:1px solid ${C.line}; color:${C.navy}; width:40px; height:40px; align-items:center; justify-content:center; cursor:pointer; flex-shrink:0; }
.ia-backdrop{ display:none; }
.ia-tl-mob{ display:none; }
/* ---- Smartphone e tablet stretti ---- */
@media (max-width: 820px){
  .ia-side{ position:fixed !important; left:0; top:0; bottom:0; height:100dvh !important; width:min(280px, 84vw) !important; z-index:80;
    transform:translateX(-100%); transition:transform .22s ease; padding-top:env(safe-area-inset-top,0px); padding-bottom:env(safe-area-inset-bottom,0px); overflow-y:auto; }
  .ia-side.open{ transform:none; box-shadow:0 0 40px rgba(0,0,0,.35); }
  .ia-backdrop.open{ display:block; position:fixed; inset:0; background:rgba(19,29,75,.5); z-index:75; }
  .ia-menu-btn{ display:inline-flex; }
  .ia-nav{ padding:13px 18px; font-size:15px; }
  .ia-head{ padding:10px 14px !important; padding-top:calc(10px + env(safe-area-inset-top,0px)) !important; flex-wrap:wrap; gap:10px; }
  .ia-head h1{ font-size:18px !important; flex:1; }
  .ia-content{ padding:14px !important; padding-bottom:calc(24px + env(safe-area-inset-bottom,0px)) !important; }
  /* griglie a più colonne → una colonna */
  .ia-main [style*="grid-template-columns"]{ grid-template-columns:1fr !important; }
  /* barre con elementi ai due lati → vanno a capo */
  .ia-main [style*="justify-content: space-between"]{ flex-wrap:wrap; gap:8px; }
  /* tabelle e riquadri larghi: scorrono di lato dentro il proprio riquadro */
  .ia-main .ia-panel{ overflow-x:auto; -webkit-overflow-scrolling:touch; }
  .ia-table.ia-keep{ min-width:640px; font-size:13px; }
  .ia-table:not(.ia-keep):not(.ia-plain) thead{ display:none; }
  .ia-table:not(.ia-keep):not(.ia-plain), .ia-table:not(.ia-keep):not(.ia-plain) tbody, .ia-table:not(.ia-keep):not(.ia-plain) tr, .ia-table:not(.ia-keep):not(.ia-plain) td{ display:block; width:100%; }
  .ia-table:not(.ia-keep):not(.ia-plain) tr{ padding:10px 14px; border-bottom:1px solid ${C.line}; }
  .ia-table:not(.ia-keep):not(.ia-plain) tr:hover td{ background:transparent; }
  .ia-table:not(.ia-keep):not(.ia-plain) td{ border:none; padding:3px 0; display:flex; justify-content:space-between; align-items:center; gap:12px; text-align:right; font-size:13.5px; min-height:24px; }
  .ia-table:not(.ia-keep):not(.ia-plain) td > *{ max-width:70%; }
  .ia-table:not(.ia-keep):not(.ia-plain) td::before{ content:attr(data-label); font-family:"Poppins",sans-serif; font-size:10.5px; font-weight:600; text-transform:uppercase; letter-spacing:.4px; color:${C.muted}; text-align:left; flex-shrink:0; }
  .ia-table:not(.ia-keep):not(.ia-plain) td[data-label=""]::before{ display:none; }
  .ia-table:not(.ia-keep):not(.ia-plain) td[data-label=""]{ justify-content:flex-end; flex-wrap:wrap; }
  .ia-table:not(.ia-keep):not(.ia-plain) td:first-child{ font-weight:600; }
  .ia-table:not(.ia-keep):not(.ia-plain) td.ia-detail{ display:block; text-align:left; font-weight:400; padding:8px 0; }
  .ia-table:not(.ia-keep):not(.ia-plain) td.ia-detail::before{ display:none; }
  .ia-input{ font-size:16px; }
  .ia-btn{ padding:10px 14px; font-size:13.5px; }
  .ia-kpi-num{ font-size:26px; }
  .ia-tabs{ overflow-x:auto; flex-wrap:nowrap !important; -webkit-overflow-scrolling:touch; }
  .ia-main .ia-kpis[style]{ grid-template-columns:1fr 1fr !important; gap:10px !important; }
  .ia-kpis > *{ padding:12px !important; min-width:0; }
  .ia-kpis .ia-kpi-num{ font-size:19px; }
  .ia-hscroll{ overflow-x:auto; -webkit-overflow-scrolling:touch; }
  .ia-chart{ min-width:620px; }
  .ia-main .ia-aging[style]{ grid-template-columns:repeat(5, minmax(96px, 1fr)) !important; overflow-x:auto; padding-bottom:4px; }
  .ia-tl-desk{ display:none !important; }
  .ia-tl-mob{ display:block !important; }
  .ia-tabs > *{ flex:0 0 auto; padding-left:12px !important; padding-right:12px !important; }
  /* finestre a tutto schermo */
  .ia-modal-bg, .ia-modal-overlay{ padding:0 !important; align-items:stretch !important; }
  .ia-modal-bg > *, .ia-modal-overlay > *{ width:100% !important; max-width:100% !important; max-height:100dvh !important; height:100dvh; overflow-y:auto !important;
    padding-top:env(safe-area-inset-top,0px); padding-bottom:env(safe-area-inset-bottom,0px); border-radius:0 !important; }
  .ia-toast{ left:12px !important; right:12px !important; bottom:calc(12px + env(safe-area-inset-bottom,0px)) !important; }
}
`;

// ---- Storage helpers (shared ledger on Supabase) ----
const K = { contracts: "ia_contracts", active: "ia_active_invoices", expenses: "ia_expenses", passive: "ia_passive_invoices", bank: "ia_bank_tx", issued: "ia_issued_invoices", dues: "ia_manual_dues", cash: "ia_cash_balance", prospects: "ia_prospects", crmActivities: "ia_crm_activities", appointments: "ia_appointments", tasks: "ia_tasks", bankRules: "ia_bank_category_rules", customBankCats: "ia_custom_bank_cats" };
// Team IA: due soci con accesso paritario. L'identità si salva nel browser
// (localStorage) dopo il login, così l'app sa chi sta facendo cosa nel CRM.
const TEAM = IS_INC
  ? [
      { id: "marco", name: "Marco Civinini", initials: "MC" },
      { id: "yazmin", name: "Yazmin Sanchez", initials: "YS" },
    ]
  : [
      { id: "marco", name: "Marco Civinini", initials: "MC" },
      { id: "guglielmo", name: "Guglielmo Iodice", initials: "GI" },
    ];
const IDENTITY_KEY = "ia_identity";
const getIdentity = () => { try { return JSON.parse(localStorage.getItem(IDENTITY_KEY) || "null"); } catch { return null; } };
const setIdentity = (u) => localStorage.setItem(IDENTITY_KEY, JSON.stringify(u));
const clearIdentity = () => localStorage.removeItem(IDENTITY_KEY);
const CRM_STAGES = [
  { id: "prospect", label: "Prospect", color: "#8891B5" },
  { id: "active", label: "Attivo", color: "#F57547" },
  { id: "won", label: "Chiuso vinto", color: "#2F8259" },
  { id: "lost", label: "Chiuso perso", color: "#B23A3A" },
];
const COST_CATEGORIES = ["Consulenze", "Viaggi", "Uffici", "Utenze", "Software", "Compensi amministratori", "Altro"];

// Categorie per la registrazione diretta di movimenti bancari nel P&L, quando
// non esiste una fattura corrispondente (Telepass, addebiti carta, biglietteria, ecc.).
// group: "cost" → costo operativo · "admin" → compensi amministratori
//        "contributi" → contributi INPS · "passthrough" → non conteggiato nel P&L
const BANK_DIRECT_CATEGORIES = [
  { id: "telepass", label: "Telepass", group: "cost" },
  { id: "credit_card", label: "Carta di credito", group: "cost" },
  { id: "tickets", label: "Biglietteria (treni/aerei)", group: "cost" },
  { id: "software", label: "Software / SaaS", group: "cost" },
  { id: "utenze", label: "Utenze (bollette)", group: "cost" },
  { id: "uffici", label: "Uffici / affitti", group: "cost" },
  { id: "consulenze", label: "Consulenze", group: "cost" },
  { id: "viaggi", label: "Viaggi (hotel/ristoranti)", group: "cost" },
  { id: "altro_costo", label: "Altro costo", group: "cost" },
  { id: "admin", label: "Compensi amministratori", group: "admin" },
  { id: "f24_inps", label: "F24 contributi INPS", group: "contributi" },
  { id: "iva", label: "IVA", group: "passthrough" },
  { id: "f24_ritenute", label: "F24 ritenute", group: "passthrough" },
];
const bankCatMeta = id => BANK_DIRECT_CATEGORIES.find(c => c.id === id) || null;
// Cerca in built-in + custom
const bankCatMetaAll = (id, custom) => BANK_DIRECT_CATEGORIES.find(c => c.id === id) || (custom || []).find(c => c.id === id) || null;
const allBankCats = (custom) => [...BANK_DIRECT_CATEGORIES, ...(custom || [])];

// Parole tipiche delle descrizioni bancarie che non identificano il fornitore
const BANK_NOISE_WORDS = new Set([
  "PAGAMENTO", "BONIFICO", "SEPA", "TRANSFER", "TRSF", "CREDIT", "DEBIT", "ADDEBITO", "ACCREDITO",
  "COMMISSIONE", "COMMISSIONI", "SPESE", "OPERAZIONE", "TRAMITE", "CARTA", "FATTURA", "FATT", "NOTA",
  "FORNITORE", "FORNITORI", "CLIENTE", "PAGATA", "GIROCONTO", "PRELIEVO", "VERSAMENTO", "IBAN",
  "AUTORIZZAZIONE", "ORDINE", "DISPOSIZIONE", "VALUTA", "DIVISA", "IMPORTO", "SALDO", "MOVIMENTO",
  "STORNO", "IMPOSTA", "TASSA", "RITENUTA", "ITALIA", "ITALY", "EUROPE", "TRAMITE",
  "GENN", "FEBB", "MARZ", "APRI", "MAGG", "GIUG", "LUGL", "AGOS", "SETT", "OTTO", "NOVE", "DICE",
  "MESE", "MESI", "ANNO", "PERIODO", "DELLA", "DELLO", "DELLE", "DEGLI",
]);

// Estrae 1-2 parole distintive dalla descrizione (per creare una regola di categorizzazione)
function extractBankKeywords(description) {
  const clean = String(description || "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  const words = clean.split(" ").filter(w => w.length >= 4 && !BANK_NOISE_WORDS.has(w) && !/^\d+$/.test(w) && !/^S[RPA][A]?$/.test(w));
  return words.slice(0, 2);
}

// Una regola combacia se TUTTE le sue keyword compaiono nella descrizione
function matchesBankRule(description, rule) {
  if (!rule || !rule.keywords || rule.keywords.length === 0) return false;
  const desc = String(description || "").toUpperCase();
  return rule.keywords.every(kw => desc.includes(kw));
}
const sload = (key, def) => kvGet(key, def);
const ssave = (key, val) => kvSet(key, val);

// ---- Utils ----
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const nowMonth = () => new Date().toISOString().slice(0, 7);
function monthLabel(m) {
  if (!m) return "";
  const [y, mo] = m.split("-");
  const names = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"];
  return `${names[+mo - 1]} ${y}`;
}
function monthOptions() {
  const out = []; const d = new Date(); d.setDate(1);
  for (let i = -6; i <= 6; i++) { const x = new Date(d.getFullYear(), d.getMonth() + i, 1); out.push(x.toISOString().slice(0, 7)); }
  return out;
}
function fmtMoney(a, cur = "EUR") {
  const n = Number(a) || 0;
  try { return new Intl.NumberFormat(cur === "USD" ? "en-US" : "it-IT", { style: "currency", currency: cur }).format(n); }
  catch { return `${n.toFixed(2)} ${cur}`; }
}
const today = () => new Date().toISOString().slice(0, 10);
function lastDayOfMonth(m) { const [y, mo] = m.split("-").map(Number); return new Date(y, mo, 0).toISOString().slice(0, 10); }
function addDays(iso, n) { const d = new Date(iso + "T00:00:00"); d.setDate(d.getDate() + (Number(n) || 0)); return d.toISOString().slice(0, 10); }
function fmtDate(iso) { if (!iso) return "—"; const [y, m, d] = iso.split("-"); return IS_INC ? `${m}/${d}/${y}` : `${d}/${m}/${y}`; }
function parseItAmount(raw) {
  if (raw == null) return NaN;
  let s = String(raw).trim().replace(/[€$\s]/g, "");
  const neg = /^\(.*\)$/.test(s) || s.includes("-");
  s = s.replace(/[()-]/g, "");
  if (IS_INC) { s = s.replace(/,/g, ""); const n = parseFloat(s); return isNaN(n) ? NaN : (neg ? -n : n); } // formato USA: 1,234.56
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", ".");
  else if (s.includes(",")) s = s.replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? NaN : (neg ? -n : n);
}
function contractActiveInMonth(c, m) {
  const s = c.startDate ? c.startDate.slice(0, 7) : null;
  const e = c.endDate ? c.endDate.slice(0, 7) : null;
  if (s && m < s) return false;
  if (e && m > e) return false;
  return true;
}
async function fileToB64(file) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = () => rej(new Error("read")); r.readAsDataURL(file); });
}

// ---- Claude extraction (via Cloudflare Worker, chiave API lato server) ----
function parseJSON(text) {
  let t = String(text).trim().replace(/```json/gi, "").replace(/```/g, "").trim();
  const s = t.indexOf("{"), e = t.lastIndexOf("}");
  if (s >= 0 && e >= 0) t = t.slice(s, e + 1);
  return JSON.parse(t);
}
async function extractFromFile(file, kind) {
  const name = (file.name || "").toLowerCase();
  if (name.endsWith(".pdf")) {
    const pdfB64 = await fileToB64(file);
    return parseJSON(await extractDoc(kind, { pdfB64 }));
  }
  if (name.endsWith(".docx") || name.endsWith(".doc")) {
    const ab = await file.arrayBuffer(); const { value } = await mammoth.extractRawText({ arrayBuffer: ab });
    return parseJSON(await extractDoc(kind, { text: value.slice(0, 12000) }));
  }
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const ab = await file.arrayBuffer(); const wb = XLSX.read(ab, { type: "array", cellDates: true });
    const text = wb.SheetNames.map(n => XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n\n");
    return parseJSON(await extractDoc(kind, { text: text.slice(0, 12000) }));
  }
  throw new Error("Formato non supportato per l'estrazione automatica (usa PDF o Word).");
}

// ---- Bank file parsing ----
function readTabular(file) {
  return new Promise(async (res, rej) => {
    const name = (file.name || "").toLowerCase();
    try {
      if (name.endsWith(".csv")) {
        const text = await file.text();
        Papa.parse(text, { skipEmptyLines: true, complete: r => res(r.data) });
      } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
        const ab = await file.arrayBuffer(); const wb = XLSX.read(ab, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        res(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }));
      } else rej(new Error("Formato estratto conto non supportato (CSV o Excel)."));
    } catch (e) { rej(e); }
  });
}
// find header row + map columns for Banco BPM-style exports
function detectBank(rows) {
  const norm = s => String(s || "").toLowerCase().trim();
  let headerIdx = -1, cols = null;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const r = rows[i].map(norm);
    const dateI = r.findIndex(h => h.includes("data") && !h.includes("valuta") || h === "data contabile" || h === "posting date" || h === "transaction date");
    const anyDate = r.findIndex(h => h.includes("data") || h.includes("date"));
    const amtI = r.findIndex(h => h.includes("importo") || h === "amount");
    const descI = r.findIndex(h => h.includes("descr") || h.includes("dettagli") || h.includes("causale") || h.includes("operazione"));
    if ((dateI >= 0 || anyDate >= 0) && (amtI >= 0 || r.some(h => h.includes("dare") || h.includes("avere") || h.includes("entrate") || h.includes("uscite")))) {
      const dareI = r.findIndex(h => h.includes("dare") || h.includes("uscite") || h.includes("addebiti"));
      const avereI = r.findIndex(h => h.includes("avere") || h.includes("entrate") || h.includes("accrediti"));
      headerIdx = i;
      cols = { date: dateI >= 0 ? dateI : anyDate, desc: descI, amount: amtI, dare: dareI, avere: avereI, headers: rows[i] };
      break;
    }
  }
  return { headerIdx, cols };
}
function rowsToTx(rows, cols, headerIdx) {
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || r.every(c => String(c).trim() === "")) continue;
    let amount = NaN;
    if (cols.amount >= 0) amount = parseItAmount(r[cols.amount]);
    if (isNaN(amount) && cols.avere >= 0) {
      const av = parseItAmount(r[cols.avere]); const da = parseItAmount(r[cols.dare]);
      if (!isNaN(av) && av !== 0) amount = Math.abs(av);
      else if (!isNaN(da) && da !== 0) amount = -Math.abs(da);
    }
    if (isNaN(amount)) continue;
    let dRaw = String(r[cols.date] || "").trim();
    let date = dRaw;
    const m = dRaw.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (m) { let y = m[3].length === 2 ? "20" + m[3] : m[3]; const [dd, mm] = IS_INC ? [m[2], m[1]] : [m[1], m[2]]; date = `${y}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`; }
    out.push({ id: uid(), date, description: String(cols.desc >= 0 ? r[cols.desc] : "").trim() || "(mov. bancario)", amount, currency: "EUR", matchedId: null, matchedType: null });
  }
  return out;
}

// Estratto conto in PDF (es. Chase): Claude legge le sezioni e restituisce i movimenti con segno.
async function importStatementPdf(file) {
  const pdfB64 = await fileToB64(file);
  const data = parseJSON(await extractDoc("bank_statement", { pdfB64 }));
  const list = Array.isArray(data.transactions) ? data.transactions : [];
  return list
    .filter(x => x && x.date && !isNaN(Number(x.amount)))
    .map(x => ({ id: uid(), date: String(x.date).slice(0, 10), description: String(x.description || "").trim() || "(bank transaction)", amount: Number(x.amount), currency: data.currency || "USD", matchedId: null, matchedType: null }));
}

// ---- Invoice-register imports (ATTIVA / PASSIVA exports from the management system) ----
function excelISO(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return new Date(v.getTime() - v.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  if (typeof v === "number") return new Date(Date.UTC(1899, 11, 30) + v * 86400000).toISOString().slice(0, 10);
  const m = String(v).match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; const [dd, mm] = IS_INC ? [m[2], m[1]] : [m[1], m[2]]; return `${y}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`; }
  return "";
}
function headerMapReq(rows, spec, required) {
  const norm = s => String(s || "").toLowerCase().trim();
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const r = rows[i].map(norm); const idx = {};
    for (const [key, kws] of Object.entries(spec)) idx[key] = r.findIndex(h => kws.some(k => h.includes(k)));
    if (required.every(k => idx[k] >= 0)) return { headerIdx: i, idx };
  }
  return { headerIdx: -1, idx: {} };
}
// Rileva se una riga rappresenta una nota di credito:
// - "Tipo documento" contiene "credito"/"NC"/"TD04"/"credit note"
// - Numero fattura inizia con "NC" o contiene "nota"
// - Importo è già negativo nell'export
function isCreditNoteRow(docType, number, amount) {
  const dt = String(docType || "").toLowerCase();
  const nm = String(number || "").toLowerCase();
  if (dt.includes("credito") || dt.includes("credit note") || dt === "td04" || /\bnc\b/.test(dt)) return true;
  if (/^nc[\s\-\/]/i.test(nm) || nm.includes("nota credito") || nm.includes("nota di credito")) return true;
  if (Number(amount) < 0) return true;
  return false;
}

function importAttiva(rows) {
  const { headerIdx, idx } = headerMapReq(rows,
    { client: ["cliente"], date: ["data"], number: ["numero"], amount: ["totale", "importo"], status: ["stato"], docType: ["tipo documento", "tipo doc", "tipo"] },
    ["client", "amount"]);
  if (headerIdx < 0) return [];
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || r.every(c => String(c).trim() === "")) continue;
    const rawAmount = parseItAmount(r[idx.amount]); const client = String(idx.client >= 0 ? r[idx.client] : "").trim();
    if (isNaN(rawAmount) || !client) continue;
    const date = idx.date >= 0 ? excelISO(r[idx.date]) : "";
    const number = String(idx.number >= 0 ? r[idx.number] : "").trim();
    const docTypeRaw = String(idx.docType >= 0 ? r[idx.docType] : "").trim();
    const isCredit = isCreditNoteRow(docTypeRaw, number, rawAmount);
    // Le note di credito riducono il fatturato: memorizzo l'importo con segno negativo,
    // così tutti i calcoli (cruscotto, P&L, top clienti) le sottraggono correttamente.
    const amount = isCredit ? -Math.abs(rawAmount) : Math.abs(rawAmount);
    out.push({ id: uid(), client, number, docType: docTypeRaw || (isCredit ? "Nota di credito" : "Fattura"), isCreditNote: isCredit, date, month: date.slice(0, 7), amount, currency: "EUR", fteStatus: String(idx.status >= 0 ? r[idx.status] : "").trim(), status: "issued", matchedId: null, source: "import" });
  }
  return out;
}
function importPassiva(rows) {
  const { headerIdx, idx } = headerMapReq(rows,
    { supplier: ["fornitore"], date: ["data"], number: ["numero"], amount: ["totale", "importo"], status: ["stato"], docType: ["tipo documento", "tipo doc", "tipo"] },
    ["supplier", "amount"]);
  if (headerIdx < 0) return [];
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r || r.every(c => String(c).trim() === "")) continue;
    const rawAmount = parseItAmount(r[idx.amount]); const supplier = String(idx.supplier >= 0 ? r[idx.supplier] : "").trim();
    if (isNaN(rawAmount) || !supplier) continue;
    const date = idx.date >= 0 ? excelISO(r[idx.date]) : "";
    const number = String(idx.number >= 0 ? r[idx.number] : "").trim();
    const docTypeRaw = String(idx.docType >= 0 ? r[idx.docType] : "").trim();
    const isCredit = isCreditNoteRow(docTypeRaw, number, rawAmount);
    // Nota di credito passiva = rimborso da fornitore. Riduce i costi, quindi
    // la memorizzo con importo negativo: nella somma dei costi si sottrae.
    const amount = isCredit ? -Math.abs(rawAmount) : Math.abs(rawAmount);
    out.push({ id: uid(), supplier, number, docType: docTypeRaw || (isCredit ? "Nota di credito" : "Fattura"), isCreditNote: isCredit, date, dueDate: "", month: date.slice(0, 7), amount, currency: "EUR", fteStatus: String(idx.status >= 0 ? r[idx.status] : "").trim(), status: "to_pay", matchedId: null, fileName: "(import Excel)" });
  }
  return out;
}
function mergeById(existing, incoming, keyFn) {
  const seen = new Set(existing.map(keyFn));
  const add = incoming.filter(x => !seen.has(keyFn(x)));
  return { merged: [...existing, ...add], added: add.length, dup: incoming.length - add.length };
}
// Chiavi di deduplica robuste per l'import.
// - Nome normalizzato (ignora spazi/punteggiatura/forma societaria)
// - Numero fattura in minuscolo (spazi collassati)
// - Valore assoluto (una NC e la sua fattura hanno segno opposto ma stesso |importo| non è problema perché il number differisce)
// - Data di emissione
// Così re-importare lo stesso file non crea duplicati anche se cambia il segno (fix note di credito).
const invNumKey = n => String(n || "").toLowerCase().replace(/\s+/g, " ").trim();
const issuedKey = x => normName(x.client) + "|" + invNumKey(x.number) + "|" + Math.abs(Number(x.amount) || 0).toFixed(2) + "|" + (x.date || "").slice(0, 10);
const passiveKey = x => normName(x.supplier) + "|" + invNumKey(x.number) + "|" + Math.abs(Number(x.amount) || 0).toFixed(2) + "|" + (x.date || "").slice(0, 10);
// normalizza nomi cliente/fornitore per il confronto (contratto ↔ fattura emessa)
const normName = s => String(s || "").toLowerCase().replace(/\s+/g, " ").replace(/[.,;:()"'`\-\/&]+/g, "").replace(/\s(s\.?r\.?l\.?|s\.?p\.?a\.?|s\.?a\.?s\.?|s\.?n\.?c\.?|srl|spa|sas|snc|ltd|inc|llc)\b.*$/, "").trim();

// Estrae il "core" numerico di un numero fattura (es. "NC 73/2025" → "73", "0044/2025" → "44").
// Serve per il matching robusto banca ↔ fattura: nelle descrizioni bancarie italiane il
// numero fattura è quasi sempre citato come cifra pura (RIF. 53, FT 73, N. 89).
function extractInvoiceCore(number) {
  const m = String(number || "").match(/(\d{1,5})/);
  return m ? m[1] : "";
}

// Verifica se il numero fattura compare nella descrizione bancaria, evitando
// falsi positivi con la data (es. "RIF. 53 30/06/2025" — il 53 è il numero, non 30 o 06).
// Cerca il numero solo se preceduto da un marker di riferimento fattura: RIF., FT, FATT,
// FATTURA, N., N.RO, NUMERO, INV, INVOICE, DOC.
function invoiceNumberInDesc(invNumber, desc) {
  const core = extractInvoiceCore(invNumber);
  if (!core) return false;
  const upper = String(desc || "").toUpperCase();
  const tokens = upper.split(/[^A-Z0-9]+/).filter(Boolean);
  const markers = new Set(["RIF", "RIFERIMENTO", "FT", "FATT", "FATTURA", "FATTURE", "N", "NRO", "NUM", "NUMERO", "INV", "INVOICE", "DOC"]);
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i] !== core) continue;
    const prev = tokens[i - 1] || "";
    if (markers.has(prev)) return true;
  }
  return false;
}
// due nomi combaciano se sono identici oppure condividono una parola distintiva (>=4 caratteri)
const namesMatch = (a, b) => {
  const na = normName(a), nb = normName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  const wa = na.split(/\s+/).filter(w => w.length >= 4);
  const wb = nb.split(/\s+/).filter(w => w.length >= 4);
  if (wa.length === 0 || wb.length === 0) return false;
  if (wa.some(w => wb.includes(w))) return true;
  if (wa.some(w => nb.includes(w))) return true;
  if (wb.some(w => na.includes(w))) return true;
  return false;
};

// React error boundary — evita che un errore in un componente faccia sparire tutta l'app
class ErrorBoundary extends React.Component {
  constructor(p) { super(p); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  componentDidCatch(err, info) { console.error("IA Suite error:", err, info); }
  render() {
    if (!this.state.err) return this.props.children;
    const msg = String(this.state.err?.message || this.state.err || "Errore imprevisto");
    return (
      <div className="ia-app" style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 24, background: "#F3F5F9" }}>
        <style>{STYLE}</style>
        <div className="ia-panel ia-bracket" style={{ padding: 26, maxWidth: 520, width: "100%" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
            <AlertTriangle size={20} color="#B23A3A" />
            <h2 className="ia-h" style={{ margin: 0, fontSize: 17, color: "#B23A3A" }}>Si è verificato un errore</h2>
          </div>
          <p style={{ fontSize: 13, color: "#5F6789", margin: "6px 0 12px", lineHeight: 1.5 }}>L'app ha intercettato un problema e ha evitato il blocco totale. Puoi ricaricare la pagina per riprendere — i dati salvati sono al sicuro.</p>
          <div style={{ background: "#FBECEC", border: "1px solid #E7C9C9", padding: "10px 12px", fontSize: 12, color: "#7A2E2E", fontFamily: "monospace", marginBottom: 14, wordBreak: "break-word" }}>{msg}</div>
          <div style={{ display: "flex", gap: 10 }}>
            <button className="ia-btn ia-btn-primary" onClick={() => location.reload()}>Ricarica pagina</button>
            <button className="ia-btn ia-btn-ghost" onClick={() => this.setState({ err: null })}>Riprova</button>
          </div>
        </div>
      </div>
    );
  }
}

function Splash() {
  return <div style={{ minHeight: "100vh", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center" }}>
    <style>{STYLE}</style><Loader2 className="ia-app" size={30} color="#fff" style={{ animation: "spin 1s linear infinite" }} />
    <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
  </div>;
}

// ============================================================
export default function App() {
  const authed = true; // l'accesso è garantito da Cloudflare Access
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState(() => {
    try {
      const t = new URLSearchParams(location.search).get("tab");
      const ok = ["dashboard", "contracts", "active", "passive", "dues", "bank", "pl", "crm", "upload", "settings", ...(IS_INC ? ["register"] : [])];
      return ok.includes(t) ? t : "dashboard";
    } catch { return "dashboard"; }
  });
  const [month, setMonth] = useState(nowMonth());
  const [contracts, setContracts] = useState([]);
  const [active, setActive] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [passive, setPassive] = useState([]);
  const [bank, setBank] = useState([]);
  const [issued, setIssued] = useState([]);
  const [dues, setDues] = useState([]);           // manual scadenze (F24 e simili)
  const [cash, setCash] = useState(null);         // saldo cassa manuale { balance, updatedAt }
  const [prospects, setProspects] = useState([]);
  const [crmActivities, setCrmActivities] = useState([]);
  const [appointments, setAppointments] = useState([]);
  const [tasks, setTasks] = useState([]);
  const [bankRules, setBankRules] = useState([]);
  const [customBankCats, setCustomBankCats] = useState([]);
  const [identity, setIdentityState] = useState(() => getIdentity());
  const [toast, setToast] = useState(null);
  const [loadError, setLoadError] = useState(null);
  const [navOpen, setNavOpen] = useState(false);
  useTableLabels();

  // Load data once authenticated
  useEffect(() => {
    if (!authed) { setReady(false); return; }
    (async () => {
      try {
        const [c, a, e, p, b, iss, du, ca, pr, ca2, ap, tk, br, cbc] = await Promise.all([
          sload(K.contracts, []), sload(K.active, []), sload(K.expenses, []), sload(K.passive, []), sload(K.bank, []), sload(K.issued, []), sload(K.dues, []), sload(K.cash, null),
          sload(K.prospects, []), sload(K.crmActivities, []), sload(K.appointments, []), sload(K.tasks, []), sload(K.bankRules, []), sload(K.customBankCats, []),
        ]);
        setContracts(c); setActive(a); setExpenses(e); setPassive(p); setBank(b); setIssued(iss); setDues(du); setCash(ca);
        setProspects(pr); setCrmActivities(ca2); setAppointments(ap); setTasks(tk); setBankRules(br); setCustomBankCats(cbc);
        setReady(true);
      } catch (err) { setLoadError(String(err.message || err)); }
    })();
  }, [authed]);

  const notify = (msg, kind = "ok") => { setToast({ msg, kind }); setTimeout(() => setToast(null), 3200); };
  const persist = useCallback((key, val, setter) => { setter(val); ssave(key, val); }, []);
  const logout = () => { clearIdentity(); accessLogout(); };
  const pickIdentity = (u) => { setIdentity(u); setIdentityState(u); };
  const switchIdentity = () => { clearIdentity(); setIdentityState(null); };

  if (!identity) return <IdentityPicker onPick={pickIdentity} onLogout={logout} />;
  if (loadError) return <LoadError msg={loadError} />;
  if (!ready) return <Splash />;

  // Accesso condiviso: identità scelta al login → attribuzione azioni CRM
  const user = { name: identity.name, role: "admin", initials: identity.initials, id: identity.id };
  const isAdmin = true;

  const ctx = {
    user, isAdmin, month, notify,
    contracts, setContracts: v => persist(K.contracts, v, setContracts),
    active, setActive: v => persist(K.active, v, setActive),
    expenses, setExpenses: v => persist(K.expenses, v, setExpenses),
    passive, setPassive: v => persist(K.passive, v, setPassive),
    bank, setBank: v => persist(K.bank, v, setBank),
    issued, setIssued: v => persist(K.issued, v, setIssued),
    dues, setDues: v => persist(K.dues, v, setDues),
    cash, setCash: v => persist(K.cash, v, setCash),
    prospects, setProspects: v => persist(K.prospects, v, setProspects),
    crmActivities, setCrmActivities: v => persist(K.crmActivities, v, setCrmActivities),
    appointments, setAppointments: v => persist(K.appointments, v, setAppointments),
    tasks, setTasks: v => persist(K.tasks, v, setTasks),
    bankRules, setBankRules: v => persist(K.bankRules, v, setBankRules),
    customBankCats, setCustomBankCats: v => persist(K.customBankCats, v, setCustomBankCats),
    switchIdentity,
  };

  const NAV = [
    { id: "dashboard", label: "Cruscotto", icon: LayoutDashboard },
    { id: "contracts", label: "Contratti", icon: FileSignature },
    { id: "active", label: "Fatture da emettere", icon: ArrowUpFromLine },
    ...(IS_INC ? [{ id: "register", label: "Invoice register", icon: BookOpen }] : []),
    { id: "passive", label: "Fatture da ricevere", icon: ArrowDownToLine },
    { id: "dues", label: "Scadenze (F24)", icon: Receipt },
    { id: "bank", label: "Banca & Riconciliazione", icon: Landmark },
    { id: "pl", label: "Conto Economico", icon: BarChart3 },
    { id: "crm", label: "CRM", icon: Users },
    { id: "upload", label: "Carica documenti", icon: FileUp },
    { id: "settings", label: "Impostazioni", icon: Settings },
  ];

  return (
    <div className="ia-app" style={{ minHeight: "100vh", display: "flex" }}>
      <style>{STYLE}</style>

      {/* Sidebar */}
      <div className={"ia-backdrop" + (navOpen ? " open" : "")} onClick={() => setNavOpen(false)} />
      <aside className={"ia-side" + (navOpen ? " open" : "")} style={{ width: 236, background: C.navy, color: "#fff", flexShrink: 0, position: "sticky", top: 0, height: "100vh", display: "flex", flexDirection: "column" }}>
        <div style={{ padding: "22px 18px 18px" }}><Logo variant="light" /></div>
        <div style={{ fontFamily: "Poppins", fontSize: 10.5, letterSpacing: ".6px", color: "#8891B5", padding: "4px 18px 12px", textTransform: "uppercase" }}>Fatturazione &amp; Riconciliazione</div>
        <nav style={{ flex: 1 }}>
          {NAV.map(n => (
            <div key={n.id} className={"ia-nav" + (tab === n.id ? " ia-nav-active" : "")} onClick={() => { setTab(n.id); setNavOpen(false); window.scrollTo(0, 0); }}>
              <n.icon size={17} /> {n.label}
            </div>
          ))}
        </nav>
        <div style={{ padding: 16, borderTop: "1px solid rgba(255,255,255,.09)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 34, height: 34, background: C.orange, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Poppins", fontWeight: 600, fontSize: 13 }}>{user.initials}</div>
            <div style={{ lineHeight: 1.2, flex: 1, minWidth: 0 }}>
              <div style={{ fontFamily: "Poppins", fontSize: 12.5, fontWeight: 500, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{user.name}</div>
              <div style={{ fontSize: 11, color: "#8891B5" }}>IA Suite · <button style={{ background: "none", border: "none", color: "#C7CBDD", cursor: "pointer", padding: 0, fontSize: 11, textDecoration: "underline" }} onClick={switchIdentity}>cambia</button></div>
            </div>
          </div>
          <a className="ia-btn ia-btn-ghost" href="https://ia-advisors-hub.com/" style={{ width: "100%", marginTop: 12, justifyContent: "center", color: "#FFFFFF", borderColor: "rgba(245,117,71,.6)", background: "transparent", textDecoration: "none", boxSizing: "border-box" }}>
            <ArrowLeft size={15} /> Torna al portale
          </a>
          <button className="ia-btn ia-btn-ghost" style={{ width: "100%", marginTop: 8, justifyContent: "center", color: "#C7CBDD", borderColor: "rgba(255,255,255,.15)", background: "transparent" }} onClick={logout}>
            <LogOut size={15} /> Esci
          </button>
        </div>
      </aside>

      {/* Main */}
      <main className="ia-main" style={{ flex: 1, minWidth: 0 }}>
        <header className="ia-head" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 26px", borderBottom: "1px solid " + C.line, background: "#fff", position: "sticky", top: 0, zIndex: 20 }}>
          <button className="ia-menu-btn" aria-label="Menu" onClick={() => setNavOpen(true)}><Menu size={20} /></button>
          <h1 className="ia-h" style={{ fontSize: 20, margin: 0, color: C.navy }}>{NAV.find(n => n.id === tab)?.label}</h1>
          {["active", "passive"].includes(tab) && (
            <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
              <span style={{ fontSize: 12.5, color: C.muted, fontFamily: "Poppins" }}>Mese di riferimento</span>
              <select className="ia-input" style={{ width: "auto", fontFamily: "Poppins", fontWeight: 600, color: C.navy }} value={month} onChange={e => setMonth(e.target.value)}>
                {monthOptions().map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}
              </select>
            </div>
          )}
        </header>

        <div className="ia-content" style={{ padding: 26 }}>
          <ErrorBoundary key={tab}>
            {tab === "dashboard" && <Dashboard ctx={ctx} go={setTab} />}
            {tab === "contracts" && <Contracts ctx={ctx} />}
            {tab === "active" && <ActiveInvoices ctx={ctx} />}
            {tab === "register" && IS_INC && <IncRegister ctx={ctx} />}
            {tab === "passive" && <PassiveInvoices ctx={ctx} />}
            {tab === "dues" && <ManualDues ctx={ctx} />}
            {tab === "bank" && <Bank ctx={ctx} />}
            {tab === "pl" && <ProfitLoss ctx={ctx} />}
            {tab === "crm" && <CRM ctx={ctx} />}
            {tab === "upload" && <UploadPortal ctx={ctx} go={setTab} />}
            {tab === "settings" && isAdmin && <SettingsView ctx={ctx} />}
          </ErrorBoundary>
        </div>
      </main>

      {toast && (
        <div className="ia-toast" style={{ position: "fixed", bottom: 22, right: 22, zIndex: 90, background: toast.kind === "err" ? C.red : C.navy, color: "#fff", padding: "12px 18px", fontFamily: "Poppins", fontSize: 13.5, display: "flex", alignItems: "center", gap: 9, boxShadow: "0 8px 24px rgba(0,0,0,.2)" }}>
          {toast.kind === "err" ? <AlertTriangle size={16} /> : <Check size={16} />} {toast.msg}
        </div>
      )}
    </div>
  );
}

// ---- Logo (vectorial recreation) ----
// Su telefono le tabelle diventano schede: ogni cella riceve il nome della sua colonna.
function useTableLabels() {
  useEffect(() => {
    const label = () => document.querySelectorAll("table.ia-table:not(.ia-keep)").forEach((t) => {
      const heads = [...t.querySelectorAll("thead th")].map((th) => (th.textContent || "").trim());
      if (!heads.length) { t.classList.add("ia-plain"); return; } // tabelle semplici senza intestazione: restano tabelle
      t.querySelectorAll("tbody tr").forEach((tr) => {
        let col = 0;
        [...tr.children].forEach((td) => {
          const want = heads[col] || "";
          if (td.getAttribute("data-label") !== want) td.setAttribute("data-label", want);
          col += Number(td.getAttribute("colspan")) || 1;
        });
      });
    });
    label();
    const mo = new MutationObserver(() => { clearTimeout(mo._t); mo._t = setTimeout(label, 60); });
    mo.observe(document.body, { childList: true, subtree: true });
    return () => mo.disconnect();
  }, []);
}

function Logo({ variant = "dark" }) {
  // Logo ufficiale IA: versione bianca su fondo scuro, versione blu su fondo chiaro
  const src = IS_INC
    ? (variant === "light" ? "/logo-inc-bianco.png" : "/logo-inc.png")
    : (variant === "light" ? "/logo-ia-bianco.png" : "/logo-ia.png");
  return <img src={src} alt="International Advisors" style={{ display: "block", height: 52, width: "auto" }} />;
}

function LoadError({ msg }) {
  return (
    <div className="ia-app" style={{ minHeight: "100vh", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <style>{STYLE}</style>
      <div className="ia-panel ia-bracket" style={{ padding: 28, maxWidth: 440 }}>
        <h2 className="ia-h" style={{ margin: "0 0 6px", fontSize: 18, color: C.navy }}>Dati non disponibili</h2>
        <p style={{ fontSize: 13.5, color: C.muted, margin: "0 0 16px", lineHeight: 1.55 }}>Non è stato possibile caricare i dati della Suite ({msg}). Ricarica la pagina tra qualche istante; se il problema resta, contatta l'amministratore.</p>
        <button className="ia-btn ia-btn-primary" onClick={() => location.reload()}><RefreshCw size={15} /> Ricarica</button>
      </div>
    </div>
  );
}

function IdentityPicker({ onPick, onLogout }) {
  return (
    <div className="ia-app" style={{ minHeight: "100vh", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <style>{STYLE}</style>
      <div style={{ width: 420, maxWidth: "100%" }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 26 }}><Logo variant="light" /></div>
        <div className="ia-panel ia-bracket" style={{ padding: 26 }}>
          <h2 className="ia-h" style={{ margin: "0 0 4px", fontSize: 18, color: C.navy }}>Chi sei?</h2>
          <p style={{ margin: "0 0 20px", fontSize: 13, color: C.muted }}>Serve per attribuire correttamente le attività CRM (chi ha chiamato, chi ha in carico un task, ecc.). Puoi cambiare identità in qualsiasi momento dal menu.</p>
          {TEAM.map(u => (
            <button key={u.id} className="ia-btn" onClick={() => onPick(u)} style={{ width: "100%", justifyContent: "flex-start", gap: 12, padding: "13px 15px", marginBottom: 10, background: "#fff", border: `1px solid ${C.line}` }}>
              <span style={{ width: 38, height: 38, background: C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Poppins", fontWeight: 600, fontSize: 14 }}>{u.initials}</span>
              <span style={{ textAlign: "left", flex: 1 }}>
                <span style={{ display: "block", fontFamily: "Poppins", fontWeight: 600, fontSize: 14, color: C.ink }}>{u.name}</span>
                <span style={{ fontSize: 12, color: C.muted }}>Accedi come {u.name.split(" ")[0]}</span>
              </span>
              <ChevronRight size={17} style={{ color: C.muted }} />
            </button>
          ))}
          <div style={{ textAlign: "center", marginTop: 14 }}>
            <button onClick={onLogout} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 12, textDecoration: "underline" }}>Esci completamente</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---- shared bits ----
function Chip({ status }) {
  const map = {
    draft: { bg: "#EEF1F6", c: C.muted, t: "Bozza" },
    issued: { bg: "#FFF1E9", c: C.orange, t: "Emessa" },
    paid: { bg: "#E7F4EE", c: C.green, t: "Incassata" },
    to_pay: { bg: "#FFF7E6", c: C.amber, t: "Da pagare" },
    ppaid: { bg: "#E7F4EE", c: C.green, t: "Pagata" },
  };
  const s = map[status] || map.draft;
  return <span className="ia-chip" style={{ background: s.bg, color: s.c }}>{s.t}</span>;
}
function Empty({ icon: Icon, title, hint, action }) {
  return <div className="ia-panel" style={{ padding: 46, textAlign: "center", color: C.muted }}>
    <Icon size={30} style={{ opacity: .4 }} />
    <div className="ia-h" style={{ fontSize: 15, color: C.ink, margin: "12px 0 5px" }}>{title}</div>
    <div style={{ fontSize: 13, maxWidth: 360, margin: "0 auto 4px" }}>{hint}</div>
    {action}
  </div>;
}
// Stable field wrapper (defined at top level so inputs keep focus across re-renders)
function Field({ label, children }) {
  return <div style={{ marginBottom: 13 }}><label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 5 }}>{label}</label>{children}</div>;
}
function ColSelect({ label, value, headers, onChange }) {
  return (
    <div style={{ marginBottom: 11 }}>
      <label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 5 }}>{label}</label>
      <select className="ia-input" value={value} onChange={e => onChange(Number(e.target.value))}>
        <option value={-1}>— nessuna —</option>
        {headers.map((h, i) => <option key={i} value={i}>{`Col ${i + 1}: ${String(h).slice(0, 24)}`}</option>)}
      </select>
    </div>
  );
}

// ============ DASHBOARD ============
const RECV_DAYS = 30, PAY_DAYS = 30; // default terms when an imported invoice has no explicit due date
const M_LABELS = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"];

function Dashboard({ ctx, go }) {
  const { contracts, active, expenses, passive, bank, issued, dues, cash } = ctx;
  const [yearSel, setYearSel] = useState(new Date().getFullYear());
  const T = today();
  const nowY = new Date().getFullYear();
  const nowMo = new Date().getMonth() + 1; // 1..12
  const monthOf = d => (d || "").slice(0, 7);
  const sum = a => a.reduce((s, x) => s + (Number(x.amount) || 0), 0);

  // === Ricevibili / pagabili / scadenze fiscali ===
  const recv = issued.map(x => ({ ...x, due: x.dueDate || addDays(x.date, RECV_DAYS), party: x.client }));
  const recvOpen = recv.filter(x => x.status !== "paid");
  const recvOverdue = recvOpen.filter(x => x.due < T);
  const pay = passive.map(p => ({ ...p, due: p.dueDate || addDays(p.date, PAY_DAYS), party: p.supplier }));
  const payOpen = pay.filter(p => p.status !== "ppaid");
  const payOverdue = payOpen.filter(p => p.due < T);
  const dueList = dues.map(x => ({ ...x, due: x.dueDate, party: x.label }));
  const dueOpen = dueList.filter(x => x.status !== "paid");
  const dueOverdue = dueOpen.filter(x => x.due && x.due < T);

  // === Cassa e posizione ===
  const cashBalance = Number(cash?.balance) || 0;
  const cashDate = cash?.updatedAt || null;
  const netPosition = cashBalance + sum(recvOpen) - sum(payOpen) - sum(dueOpen);

  // === Fatturato YTD e confronto anno precedente ===
  const ytd = sum(issued.filter(x => monthOf(x.date).slice(0, 4) === String(nowY)));
  const ytdPrev = sum(issued.filter(x => {
    const m = monthOf(x.date);
    return m.slice(0, 4) === String(nowY - 1) && Number(m.slice(5, 7) || 0) <= nowMo;
  }));
  const ytdDelta = ytdPrev > 0 ? (ytd - ytdPrev) / ytdPrev * 100 : null;

  // === Grafico Gen-Dic per anno selezionato ===
  // "Da emettere" (piano dai contratti) viene mostrato solo per i mesi
  // correnti e futuri: sui mesi passati mostriamo solo il fatturato reale.
  const months12 = Array.from({ length: 12 }, (_, i) => `${yearSel}-${String(i + 1).padStart(2, "0")}`);
  const monthlyChart = months12.map((m, i) => {
    const em = issued.filter(x => (x.month || monthOf(x.date)) === m);
    const emSum = sum(em);
    const emCount = em.length;
    const isPast = m < nowMonth();
    const isCurrent = m === nowMonth();
    const isFuture = m > nowMonth();
    // Piano solo per corrente/futuro; sui mesi passati resta 0
    let planned = 0;
    if (!isPast) {
      planned = compileActive(contracts, active, expenses, m, issued)
        .filter(r => r.grossTotal > 0 && !r.hasIssued)
        .reduce((s, r) => s + r.grossTotal, 0);
    }
    return { m, monthIdx: i, emSum, emCount, planned, isPast, isCurrent, isFuture };
  });
  const chartMax = Math.max(1, ...monthlyChart.map(x => Math.max(x.emSum, x.planned)));
  const yearTotals = {
    em: monthlyChart.reduce((s, x) => s + x.emSum, 0),
    planned: monthlyChart.reduce((s, x) => s + x.planned, 0),
  };

  // Confronto anno precedente per stesso periodo (fino al mese corrente se yearSel = anno corrente)
  const prevYear = yearSel - 1;
  const cutMonth = yearSel === nowY ? nowMo : 12;
  const prevTotalSame = sum(issued.filter(x => {
    const m = monthOf(x.date);
    return m.slice(0, 4) === String(prevYear) && Number(m.slice(5, 7) || 0) <= cutMonth;
  }));
  const curTotalSame = sum(issued.filter(x => {
    const m = monthOf(x.date);
    return m.slice(0, 4) === String(yearSel) && Number(m.slice(5, 7) || 0) <= cutMonth;
  }));
  const yoyDelta = prevTotalSame > 0 ? (curTotalSame - prevTotalSame) / prevTotalSame * 100 : null;

  return (
    <div>
      <div style={{ fontSize: 12.5, color: C.muted, marginBottom: 14, fontFamily: "Poppins" }}>Aggiornato a {fmtDate(T)}</div>

      {/* KPI in cima */}
      <div className="ia-kpis" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 16, marginBottom: 18 }}>
        <div className="ia-panel ia-bracket" style={{ padding: 18 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".6px" }}>Cassa attuale</div>
          <div className="ia-kpi-num" style={{ color: C.navy, margin: "4px 0" }}>{fmtMoney(cashBalance)}</div>
          <div style={{ fontSize: 11, color: C.muted }}>{cashDate ? "al " + fmtDate(cashDate) : "non impostata"} · <button style={{ border: "none", background: "none", color: C.orange, cursor: "pointer", fontSize: 11, padding: 0, textDecoration: "underline" }} onClick={() => go("settings")}>aggiorna</button></div>
          <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid " + C.line, fontSize: 11.5, color: C.muted }}>Posizione netta prevista<br /><b style={{ fontFamily: "Poppins", color: netPosition >= 0 ? C.green : C.red, fontSize: 16 }}>{netPosition >= 0 ? "+" : ""}{fmtMoney(netPosition)}</b></div>
        </div>

        <div className="ia-panel" style={{ padding: 18 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".6px" }}>Fatturato {nowY} (YTD)</div>
          <div className="ia-kpi-num" style={{ color: C.green, margin: "4px 0" }}>{fmtMoney(ytd)}</div>
          {ytdDelta !== null ? (
            <>
              <div style={{ fontSize: 12, color: ytdDelta >= 0 ? C.green : C.red, fontFamily: "Poppins", fontWeight: 600 }}>{ytdDelta >= 0 ? "▲" : "▼"} {Math.abs(ytdDelta).toFixed(1)}% vs {nowY - 1}</div>
              <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>Stesso periodo {nowY - 1}: {fmtMoney(ytdPrev)}</div>
            </>
          ) : <div style={{ fontSize: 11.5, color: C.muted }}>Nessun dato per confronto {nowY - 1}</div>}
        </div>

        <div className="ia-panel" style={{ padding: 18, cursor: "pointer" }} onClick={() => go("active")}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".6px" }}>Da incassare</div>
          <div className="ia-kpi-num" style={{ color: C.green, margin: "4px 0" }}>{fmtMoney(sum(recvOpen))}</div>
          <div style={{ fontSize: 12, color: C.muted }}>{recvOpen.length} fatture aperte</div>
          {recvOverdue.length > 0 && <div style={{ marginTop: 8, padding: "5px 8px", background: "#FBECEC", color: C.red, fontSize: 12, fontFamily: "Poppins", fontWeight: 600 }}>{recvOverdue.length} scadute · {fmtMoney(sum(recvOverdue))}</div>}
        </div>

        <div className="ia-panel" style={{ padding: 18, cursor: "pointer" }} onClick={() => go("passive")}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".6px" }}>Da pagare</div>
          <div className="ia-kpi-num" style={{ color: C.red, margin: "4px 0" }}>{fmtMoney(sum(payOpen) + sum(dueOpen))}</div>
          <div style={{ fontSize: 12, color: C.muted }}>{payOpen.length} fornitori{dueOpen.length > 0 && ` + ${dueOpen.length} F24`}</div>
          {(payOverdue.length + dueOverdue.length) > 0 && <div style={{ marginTop: 8, padding: "5px 8px", background: "#FBECEC", color: C.red, fontSize: 12, fontFamily: "Poppins", fontWeight: 600 }}>{payOverdue.length + dueOverdue.length} scaduti · {fmtMoney(sum(payOverdue) + sum(dueOverdue))}</div>}
        </div>
      </div>

      {/* Grafico Gen-Dic: Emesse vs Da emettere */}
      <div className="ia-panel" style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 20px", borderBottom: "1px solid " + C.line, gap: 12, flexWrap: "wrap" }}>
          <div>
            <div className="ia-h" style={{ fontSize: 14.5, color: C.navy }}>Fatturato mensile — emesse vs pianificato</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Barra piena = fattura reale emessa · Barra tratteggiata = piano dai contratti (retainer, una tantum, eventi maturati)</div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {yoyDelta !== null && (
              <span style={{ fontSize: 12, color: yoyDelta >= 0 ? C.green : C.red, fontFamily: "Poppins", fontWeight: 600, background: yoyDelta >= 0 ? "#EAF6EE" : "#FBECEC", padding: "3px 8px" }}>
                {yoyDelta >= 0 ? "▲" : "▼"} {Math.abs(yoyDelta).toFixed(1)}% vs {prevYear}
              </span>
            )}
            <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 10px" }} onClick={() => setYearSel(yearSel - 1)}>‹</button>
            <span className="ia-h" style={{ fontSize: 15, color: C.navy, minWidth: 50, textAlign: "center" }}>{yearSel}</span>
            <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 10px" }} onClick={() => setYearSel(yearSel + 1)}>›</button>
          </div>
        </div>
        <div className="ia-hscroll" ref={el => { if (el && el.scrollWidth > el.clientWidth && !el.dataset.init) { el.scrollLeft = el.scrollWidth; el.dataset.init = "1"; } }} style={{ padding: "18px 20px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 14, fontSize: 12.5, flexWrap: "wrap", gap: 12 }}>
            <div><span style={{ color: C.muted }}>Emesse: </span><b style={{ color: C.navy, fontFamily: "Poppins" }}>{fmtMoney(yearTotals.em)}</b></div>
            <div><span style={{ color: C.muted }}>Piano corrente + futuro: </span><b style={{ color: C.orange, fontFamily: "Poppins" }}>{fmtMoney(yearTotals.planned)}</b></div>
          </div>
          <div className="ia-chart" style={{ display: "flex", gap: 4, alignItems: "flex-end" }}>
            {monthlyChart.map(x => (
              <div key={x.m} style={{ flex: 1, textAlign: "center", background: x.isCurrent ? "#FFF9F6" : "transparent", padding: "6px 2px", border: x.isCurrent ? "1px solid " + C.orange : "1px solid transparent" }}>
                <div style={{ display: "flex", gap: 2, alignItems: "flex-end", justifyContent: "center", height: 150 }}>
                  <div title={`Emesse ${fmtMoney(x.emSum)} · ${x.emCount} fatt.`} style={{ width: 14, height: `${Math.max(x.emSum ? 4 : 0, Math.round(x.emSum / chartMax * 100))}%`, background: C.navy }} />
                  {!x.isPast && <div title={`Pianificato ${fmtMoney(x.planned)}`} style={{ width: 14, height: `${Math.max(x.planned ? 4 : 0, Math.round(x.planned / chartMax * 100))}%`, background: `repeating-linear-gradient(45deg, ${C.orange} 0, ${C.orange} 3px, #FFEEDC 3px, #FFEEDC 6px)`, border: `1px solid ${C.orange}` }} />}
                </div>
                <div style={{ fontFamily: "Poppins", fontSize: 10.5, color: x.isCurrent ? C.orange : C.muted, marginTop: 8, fontWeight: x.isCurrent ? 700 : 500 }}>{M_LABELS[x.monthIdx]}</div>
                {x.emSum > 0 && <div style={{ fontFamily: "Poppins", fontSize: 10.5, color: C.navy, fontWeight: 600 }}>{shortEur(x.emSum)}</div>}
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 18, marginTop: 14, fontSize: 11, color: C.muted, borderTop: "1px solid " + C.line, paddingTop: 12, flexWrap: "wrap" }}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><i style={{ width: 12, height: 12, background: C.navy, display: "inline-block" }} /> Emesse (reali)</span>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><i style={{ width: 12, height: 12, background: `repeating-linear-gradient(45deg, ${C.orange} 0, ${C.orange} 3px, #FFEEDC 3px, #FFEEDC 6px)`, border: `1px solid ${C.orange}`, display: "inline-block" }} /> Da emettere (piano contratti, solo mese corrente e futuri)</span>
            <span style={{ color: C.orange }}>■ mese corrente</span>
          </div>
        </div>
      </div>

      {/* Timeline crediti */}
      <TimelinePanel
        title="Crediti da incassare — linea del tempo"
        subtitle={`${recvOpen.length} fatture aperte · totale ${fmtMoney(sum(recvOpen))}`}
        items={recvOpen.map(x => ({ due: x.due, amount: x.amount, party: x.party, ref: x.number, _id: x.id }))}
        today={T} color={C.green} overdueSum={sum(recvOverdue)} onGo={() => go("active")}
        showAging={true} kind="credit"
        onMarkPaid={(item) => {
          ctx.setIssued(issued.map(x => x.id === item._id ? { ...x, status: "paid", paidAt: today(), paidBy: "manual" } : x));
          ctx.notify("Credito segnato come incassato");
        }}
      />

      {/* Timeline debiti (fornitori + F24) */}
      <TimelinePanel
        title="Debiti da pagare — linea del tempo"
        subtitle={`${payOpen.length} fatture fornitori + ${dueOpen.length} scadenze fiscali · totale ${fmtMoney(sum(payOpen) + sum(dueOpen))}`}
        items={[
          ...payOpen.map(x => ({ due: x.due, amount: x.amount, party: x.party, ref: x.number, _src: "passive", _id: x.id })),
          ...dueOpen.filter(x => x.due).map(x => ({ due: x.due, amount: x.amount, party: x.party || x.label, ref: x.category, isTax: true, _src: "due", _id: x.id })),
        ]}
        today={T} color={C.red} overdueSum={sum(payOverdue) + sum(dueOverdue)} onGo={() => go("passive")}
        showAging={true} kind="debit"
        onMarkPaid={(item) => {
          if (item._src === "passive") {
            ctx.setPassive(passive.map(p => p.id === item._id ? { ...p, status: "ppaid", paidAt: today(), paidBy: "manual" } : p));
            ctx.notify("Fattura fornitore segnata come pagata");
          } else {
            ctx.setDues(dues.map(d => d.id === item._id ? { ...d, status: "paid", paidAt: today(), paidBy: "manual" } : d));
            ctx.notify("Scadenza segnata come pagata");
          }
        }}
        onDelete={(item) => {
          if (item._src === "passive") {
            ctx.setPassive(passive.filter(p => p.id !== item._id));
            ctx.notify("Fattura eliminata");
          } else {
            ctx.setDues(dues.filter(d => d.id !== item._id));
            ctx.notify("Scadenza eliminata");
          }
        }}
      />

      {/* Riconciliazione */}
      <div className="ia-panel" style={{ padding: "16px 20px", marginBottom: 18 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 24 }}>
          <ReconBar label="Emesse incassate" done={issued.filter(x => x.status === "paid").length} tot={issued.length} color={C.green} />
          <ReconBar label="Passive pagate" done={passive.filter(p => p.status === "ppaid").length} tot={passive.length} color={C.blue} />
          <ReconBar label="Movimenti abbinati" done={bank.filter(t => t.matchedId).length} tot={bank.length} color={C.orange} />
        </div>
      </div>

      {/* Sintesi Conto Economico YTD */}
      <PLSummary ctx={ctx} go={go} />

    </div>
  );
}

// Sintesi compatta del P&L da mostrare in cruscotto
function PLSummary({ ctx, go }) {
  const { issued, passive, dues } = ctx;
  const yearOf = d => (d || "").slice(0, 4);
  const nowY = new Date().getFullYear();
  const netOfPassive = (p) => {
    const amt = Number(p.amount) || 0;
    const sign = amt < 0 ? -1 : 1;
    if (p.netAmount != null && p.netAmount !== "" && isFinite(Number(p.netAmount))) return sign * Math.abs(Number(p.netAmount)); // imponibile letto dal documento
    const gross = Math.abs(amt);
    if (p.vatExempt) return sign * gross;
    const rate = Number(p.vatRate);
    if (!rate || isNaN(rate)) return sign * (gross / 1.22);
    return sign * (gross / (1 + rate / 100));
  };
  const isDueCost = (d) => {
    const c = (d.category || "").toLowerCase();
    return c.includes("inps") || c.includes("contribut") || c.includes("bolli") || c.includes("diritt") || c === "altro" || c === "f24";
  };
  const yIss = issued.filter(x => yearOf(x.month || x.date) === String(nowY));
  const yPas = passive.filter(p => yearOf(p.date) === String(nowY));
  const yDu = dues.filter(d => yearOf(d.dueDate || d.paidAt) === String(nowY));
  const rev = yIss.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const admin = yPas.filter(p => (p.category || "").toLowerCase().includes("amministrator")).reduce((s, p) => s + netOfPassive(p), 0);
  const other = yPas.filter(p => !(p.category || "").toLowerCase().includes("amministrator")).reduce((s, p) => s + netOfPassive(p), 0);
  const contrib = yDu.filter(isDueCost).reduce((s, d) => s + Math.abs(Number(d.amount) || 0), 0);
  const margin = rev - admin - other - contrib;
  const pct = rev > 0 ? (margin / rev) * 100 : 0;
  return (
    <div className="ia-panel" style={{ padding: "16px 20px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div>
          <div className="ia-h" style={{ fontSize: 14, color: C.navy, display: "inline-flex", alignItems: "center", gap: 8 }}><BarChart3 size={16} color={C.orange} /> Conto economico gestionale {nowY}</div>
          <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Vista sintetica YTD · <b>non ufficiale</b> · <button style={{ border: "none", background: "none", color: C.orange, cursor: "pointer", padding: 0, fontSize: 11.5, textDecoration: "underline" }} onClick={() => go("pl")}>dettaglio completo →</button></div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <div>
          <div style={{ fontSize: 10.5, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".4px" }}>Ricavi</div>
          <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 18, color: C.green, marginTop: 2 }}>{fmtMoney(rev)}</div>
        </div>
        <div>
          <div style={{ fontSize: 10.5, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".4px" }}>Costi + contributi</div>
          <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 18, color: C.red, marginTop: 2 }}>{fmtMoney(other + contrib)}</div>
        </div>
        <div>
          <div style={{ fontSize: 10.5, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".4px" }}>Comp. amministratori</div>
          <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 18, color: C.orange, marginTop: 2 }}>{fmtMoney(admin)}</div>
        </div>
        <div>
          <div style={{ fontSize: 10.5, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".4px" }}>Margine ({pct.toFixed(1)}%)</div>
          <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 18, color: margin >= 0 ? C.green : C.red, marginTop: 2 }}>{fmtMoney(margin)}</div>
        </div>
      </div>
    </div>
  );
}

// ============ TIMELINE PANEL ============
function AgingBar({ items, today: T, onBucketClick }) {
  const buckets = { notDue: [], "1-30": [], "31-60": [], "61-90": [], "90+": [] };
  const dayDiff = (a, b) => Math.floor((new Date(a + "T00:00:00") - new Date(b + "T00:00:00")) / 86400000);
  items.forEach(x => {
    if (!x.due) return;
    const late = dayDiff(T, x.due);
    if (late <= 0) buckets.notDue.push(x);
    else if (late <= 30) buckets["1-30"].push(x);
    else if (late <= 60) buckets["31-60"].push(x);
    else if (late <= 90) buckets["61-90"].push(x);
    else buckets["90+"].push(x);
  });
  const sums = Object.fromEntries(Object.entries(buckets).map(([k, arr]) => [k, arr.reduce((s, x) => s + Math.abs(Number(x.amount) || 0), 0)]));
  const total = Object.values(sums).reduce((s, v) => s + v, 0);
  if (total === 0) return null;
  const cfg = [
    { key: "notDue", label: "Non scaduto", color: C.green },
    { key: "1-30", label: "1-30 gg", color: "#F5C56A" },
    { key: "31-60", label: "31-60 gg", color: C.amber },
    { key: "61-90", label: "61-90 gg", color: "#E88C4C" },
    { key: "90+", label: "> 90 gg", color: C.red },
  ];
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>
        <span>Aging — clicca un bucket per vedere le fatture</span>
        <span>Totale <b style={{ color: C.navy }}>{fmtMoney(total)}</b></span>
      </div>
      <div style={{ display: "flex", height: 8, background: "#EEF1F6", overflow: "hidden", marginBottom: 8, borderRadius: 2 }}>
        {cfg.map(c => sums[c.key] > 0 ? (
          <div key={c.key} style={{ width: `${(sums[c.key] / total) * 100}%`, background: c.color }} title={`${c.label}: ${fmtMoney(sums[c.key])}`} />
        ) : null)}
      </div>
      <div className="ia-aging" style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 6, fontSize: 11 }}>
        {cfg.map(c => {
          const active = sums[c.key] > 0;
          return (
            <button key={c.key} disabled={!active} onClick={() => active && onBucketClick && onBucketClick(c, buckets[c.key])} style={{ textAlign: "center", padding: "6px 4px", background: active ? "#FAFBFD" : "transparent", border: "1px solid " + (active ? C.line : "transparent"), cursor: active ? "pointer" : "default", font: "inherit", transition: "background .1s" }} onMouseEnter={e => { if (active) e.currentTarget.style.background = "#F0F3F8"; }} onMouseLeave={e => { if (active) e.currentTarget.style.background = "#FAFBFD"; }}>
              <div style={{ display: "inline-flex", alignItems: "center", gap: 4, marginBottom: 2 }}>
                <i style={{ width: 8, height: 8, background: c.color, display: "inline-block", borderRadius: 4 }} />
                <span style={{ color: C.muted, fontSize: 10.5 }}>{c.label}</span>
              </div>
              <div style={{ fontFamily: "Poppins", fontWeight: 600, color: active ? c.color : C.muted, fontSize: 12 }}>{fmtMoney(sums[c.key])}</div>
              {buckets[c.key].length > 0 && <div style={{ color: C.muted, fontSize: 10 }}>{buckets[c.key].length} {buckets[c.key].length === 1 ? "voce" : "voci"}</div>}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function AgingBucketModal({ bucket, items, today: T, kind, onMarkPaid, onDelete, onClose }) {
  const [live, setLive] = useState(items);
  const sorted = [...live].sort((a, b) => (a.due || "").localeCompare(b.due || ""));
  const total = sorted.reduce((s, x) => s + Math.abs(Number(x.amount) || 0), 0);
  const markPaid = (item) => {
    if (onMarkPaid) onMarkPaid(item);
    setLive(live.filter(x => x !== item));
  };
  const del = (item) => {
    if (onDelete) onDelete(item);
    setLive(live.filter(x => x !== item));
  };
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 720, maxWidth: "100%", maxHeight: "88vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ width: 12, height: 12, background: bucket.color, borderRadius: 6, display: "inline-block" }} />
              <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{kind === "credit" ? "Crediti" : "Debiti"} — {bucket.label}</h3>
            </div>
            <div style={{ fontSize: 12.5, color: C.muted, marginTop: 4 }}>{sorted.length} {sorted.length === 1 ? "voce" : "voci"} · totale <b style={{ color: C.navy, fontFamily: "Poppins" }}>{fmtMoney(total)}</b></div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        {sorted.length === 0 ? (
          <div style={{ padding: 30, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessuna voce rimasta in questo bucket.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="ia-table">
              <thead><tr><th>{kind === "credit" ? "Cliente" : "Fornitore / voce"}</th><th>Riferimento</th><th>Scadenza</th><th style={{ textAlign: "right" }}>Importo</th><th>Ritardo</th><th style={{ textAlign: "right" }}>Azioni</th></tr></thead>
              <tbody>
                {sorted.map((x, i) => {
                  const dayDiff = Math.floor((new Date(T + "T00:00:00") - new Date(x.due + "T00:00:00")) / 86400000);
                  const late = dayDiff > 0;
                  return (
                    <tr key={i}>
                      <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{x.party}{x.isTax && <span className="ia-chip" style={{ background: "#FFF8E6", color: C.amber, marginLeft: 8, fontSize: 10 }}>fiscale</span>}</td>
                      <td style={{ fontSize: 12.5, color: C.muted }}>{x.ref || "—"}</td>
                      <td style={{ whiteSpace: "nowrap", color: late ? C.red : C.ink }}>{fmtDate(x.due)}</td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: kind === "credit" ? C.green : C.red }}>{fmtMoney(Math.abs(Number(x.amount) || 0))}</td>
                      <td style={{ fontSize: 12.5, color: late ? C.red : C.muted }}>{late ? dayDiff + " gg" : dayDiff === 0 ? "oggi" : Math.abs(dayDiff) + " gg mancanti"}</td>
                      <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                        {onMarkPaid && <button className="ia-btn ia-btn-dark" style={{ padding: "5px 10px", fontSize: 12 }} onClick={() => markPaid(x)}><Check size={12} /> {kind === "credit" ? "Segna incassata" : "Segna pagato"}</button>}
                        {onDelete && <>{" "}<button className="ia-btn ia-btn-danger" style={{ padding: "5px 8px", fontSize: 12 }} onClick={() => { if (confirm(`Eliminare "${x.party}"?`)) del(x); }} title="Elimina"><Trash2 size={12} /></button></>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "flex-end", padding: "12px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Chiudi</button>
        </div>
      </div>
    </div>
  );
}

function TimelinePanel({ title, subtitle, items, today: T, color, overdueSum, onGo, showAging, onMarkPaid, onDelete, kind }) {
  const [bucket, setBucket] = useState(null);
  if (!items || items.length === 0) {
    return (
      <div className="ia-panel" style={{ marginBottom: 18 }}>
        <PanelHead title={title} />
        <div style={{ padding: 22, fontSize: 13, color: C.muted, textAlign: "center" }}>Nessun elemento aperto.</div>
      </div>
    );
  }
  const past = 30, future = 90, tot = past + future;
  const dayDiff = (a, b) => Math.floor((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 86400000);
  const ticks = [-30, 0, 30, 60, 90];
  const valid = items.filter(x => x.due);
  const sorted = [...valid].sort((a, b) => a.due.localeCompare(b.due));
  const veryOverdue = sorted.filter(x => dayDiff(T, x.due) < -past);
  const inRange = sorted.filter(x => { const d = dayDiff(T, x.due); return d >= -past && d <= future; });
  const beyond = sorted.filter(x => dayDiff(T, x.due) > future);
  const maxAmt = Math.max(1, ...sorted.map(x => Math.abs(Number(x.amount) || 0)));

  return (
    <div className="ia-panel" style={{ marginBottom: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 20px", borderBottom: "1px solid " + C.line, gap: 10 }}>
        <div>
          <div className="ia-h" style={{ fontSize: 14.5, color: C.navy }}>{title}</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{subtitle}{overdueSum > 0 && <> · <b style={{ color: C.red }}>scaduto: {fmtMoney(overdueSum)}</b></>}</div>
        </div>
        {onGo && <button className="ia-btn ia-btn-ghost" onClick={onGo}>Vedi tutte</button>}
      </div>
      <div style={{ padding: "18px 22px" }}>
        {showAging && <AgingBar items={items} today={T} onBucketClick={(cfg, arr) => setBucket({ cfg, arr })} />}
        {veryOverdue.length > 0 && (
          <div style={{ padding: "8px 12px", background: "#FBECEC", border: "1px solid #E7C9C9", marginBottom: 14, fontSize: 12, color: C.red, fontFamily: "Poppins" }}>
            <b>{veryOverdue.length} voci</b> scadute da oltre 30 giorni (non mostrate nella linea del tempo) · totale <b>{fmtMoney(veryOverdue.reduce((s, x) => s + (Number(x.amount) || 0), 0))}</b>
          </div>
        )}
        {/* Asse temporale con tick */}
        <div className="ia-tl-desk" style={{ position: "relative", height: 26, marginBottom: 6 }}>
          {ticks.map(d => {
            const pos = ((d + past) / tot) * 100;
            const isToday = d === 0;
            return (
              <div key={d} style={{ position: "absolute", left: `${pos}%`, transform: "translateX(-50%)", textAlign: "center" }}>
                <div style={{ width: 1, height: 12, background: isToday ? C.orange : C.line, margin: "0 auto" }} />
                <div style={{ fontSize: 10.5, color: isToday ? C.orange : C.muted, fontFamily: "Poppins", fontWeight: isToday ? 700 : 500, marginTop: 2 }}>{d === 0 ? "OGGI" : d < 0 ? d + "gg" : "+" + d + "gg"}</div>
              </div>
            );
          })}
        </div>
        {/* Su telefono: elenco compatto ordinato per scadenza (al posto della linea del tempo) */}
        <div className="ia-tl-mob">
          {inRange.slice().sort((a, b) => (a.due || "").localeCompare(b.due || "")).slice(0, 18).map((x, i) => {
            const d = dayDiff(T, x.due);
            const isOver = d < 0, isSoon = d >= 0 && d <= 7;
            const dotCol = isOver ? C.red : isSoon ? C.amber : color;
            return (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 2px", borderBottom: "1px solid " + C.line, fontSize: 13 }}>
                <i style={{ width: 10, height: 10, borderRadius: 5, background: dotCol, flexShrink: 0 }} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{x.party}</div>
                  <div style={{ fontSize: 11.5, color: isOver ? C.red : C.muted }}>{fmtDate(x.due)} · {isOver ? Math.abs(d) + " gg fa" : d === 0 ? "oggi" : "tra " + d + " gg"}</div>
                </div>
                <b style={{ fontFamily: "Poppins", fontSize: 13, color: isOver ? C.red : C.ink, whiteSpace: "nowrap" }}>{fmtMoney(Number(x.amount) || 0)}</b>
              </div>
            );
          })}
          {inRange.length > 18 && <div style={{ padding: "8px 2px", fontSize: 11.5, color: C.muted, fontStyle: "italic" }}>… e altre {inRange.length - 18} voci</div>}
        </div>
        {/* Zone e barre */}
        <div className="ia-tl-desk" style={{ position: "relative", background: "#FAFBFD", border: "1px solid " + C.line, minHeight: 60, padding: "10px 0" }}>
          {/* Zona scaduta */}
          <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${(past / tot) * 100}%`, background: "#FBECEC", opacity: .35, pointerEvents: "none" }} />
          {/* Linea "oggi" */}
          <div style={{ position: "absolute", left: `${(past / tot) * 100}%`, top: 0, bottom: 0, width: 2, background: C.orange, pointerEvents: "none" }} />
          {/* Elementi */}
          {inRange.slice(0, 18).map((x, i) => {
            const d = dayDiff(T, x.due);
            const pos = ((d + past) / tot) * 100;
            const amt = Number(x.amount) || 0;
            const dotSize = 10 + Math.round((amt / maxAmt) * 18);
            const isOver = d < 0, isSoon = d >= 0 && d <= 7;
            const dotCol = isOver ? C.red : isSoon ? C.amber : color;
            const showLabelLeft = pos > 60;
            return (
              <div key={i} style={{ position: "relative", height: 26, marginBottom: 3 }} title={`${x.party} · ${fmtMoney(amt)} · scad. ${fmtDate(x.due)} · ${d < 0 ? Math.abs(d) + " gg fa" : d === 0 ? "oggi" : "tra " + d + " gg"}${x.isTax ? " · scadenza fiscale" : ""}`}>
                <div style={{ position: "absolute", left: `calc(${pos}% - ${dotSize / 2}px)`, top: `calc(50% - ${dotSize / 2}px)`, width: dotSize, height: dotSize, borderRadius: dotSize, background: dotCol, border: "2px solid #fff", boxShadow: "0 1px 3px rgba(0,0,0,.15)" }} />
                <div style={{ position: "absolute", top: "50%", transform: "translateY(-50%)", fontSize: 11, color: isOver ? C.red : C.ink, fontFamily: "Poppins", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 220, ...(showLabelLeft ? { right: `calc(100% - ${pos}% + ${dotSize / 2 + 4}px)`, textAlign: "right" } : { left: `calc(${pos}% + ${dotSize / 2 + 4}px)` }) }}>
                  <b>{x.party}</b> · {fmtMoney(amt)}
                </div>
              </div>
            );
          })}
          {inRange.length > 18 && <div style={{ padding: "6px 12px", fontSize: 11, color: C.muted, textAlign: "center", fontStyle: "italic" }}>+ {inRange.length - 18} altre voci</div>}
        </div>
        {beyond.length > 0 && (
          <div style={{ marginTop: 10, fontSize: 11.5, color: C.muted, textAlign: "right" }}>+ {beyond.length} voci oltre 90 giorni · {fmtMoney(beyond.reduce((s, x) => s + (Number(x.amount) || 0), 0))}</div>
        )}
        <div style={{ display: "flex", gap: 16, marginTop: 12, fontSize: 11, color: C.muted, borderTop: "1px solid " + C.line, paddingTop: 10, flexWrap: "wrap" }}>
          <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><i style={{ width: 10, height: 10, borderRadius: 5, background: C.red, display: "inline-block" }} /> Scaduto</span>
          <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><i style={{ width: 10, height: 10, borderRadius: 5, background: C.amber, display: "inline-block" }} /> Entro 7 giorni</span>
          <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}><i style={{ width: 10, height: 10, borderRadius: 5, background: color, display: "inline-block" }} /> Successivi (fino a 90gg)</span>
          <span>· Dimensione cerchio proporzionale all'importo · clicca un bucket aging per dettagli e imputazione pagamenti</span>
        </div>
      </div>
      {bucket && <AgingBucketModal bucket={bucket.cfg} items={bucket.arr} today={T} kind={kind} onMarkPaid={onMarkPaid} onDelete={onDelete} onClose={() => setBucket(null)} />}
    </div>
  );
}
function StatCluster({ title, accent, stats }) {
  return (
    <div className="ia-panel" style={{ padding: 18 }}>
      <div className="ia-h" style={{ fontSize: 13, color: C.navy, marginBottom: 14, display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 8, height: 8, background: accent, display: "inline-block" }} />{title}</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        {stats.map((s, i) => (
          <div key={i}>
            <div style={{ fontSize: 11, fontFamily: "Poppins", fontWeight: 500, color: C.muted, marginBottom: 4 }}>{s.label}</div>
            <div className="ia-h" style={{ fontSize: 19, fontWeight: 700, color: s.bad ? C.red : s.good ? C.green : C.navy }}>{fmtMoney(s.val)}</div>
            {s.sub && <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{s.sub}</div>}
          </div>
        ))}
      </div>
    </div>
  );
}
function shortEur(n) {
  n = Number(n) || 0;
  if (Math.abs(n) >= 1000) return "€" + (n / 1000).toFixed(n >= 10000 ? 0 : 1) + "k";
  return "€" + Math.round(n);
}
function ReconBar({ label, done, tot, color }) {
  const pct = tot ? Math.round((done / tot) * 100) : 0;
  return <div style={{ marginBottom: 18 }}>
    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 6 }}>
      <span>{label}</span><span style={{ fontFamily: "Poppins", fontWeight: 600, color: C.navy }}>{done}/{tot}</span>
    </div>
    <div style={{ height: 8, background: "#EEF1F6" }}><div style={{ height: "100%", width: pct + "%", background: color }} /></div>
  </div>;
}
function PanelHead({ title, action }) {
  return <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", borderBottom: "1px solid " + C.line }}>
    <h3 className="ia-h" style={{ margin: 0, fontSize: 14.5, color: C.navy }}>{title}</h3>{action}
  </div>;
}

// normalize billing items (migrate legacy single retainer)
function getBillingItems(c) {
  if (Array.isArray(c.billingItems) && c.billingItems.length) return c.billingItems;
  if (c.retainerDefault != null && c.retainerDefault !== "")
    return [{ id: "legacy", type: "recurring", label: "Retainer mensile", amount: Number(c.retainerDefault) }];
  return [];
}
const emptyItem = (type) => {
  const base = { id: uid(), type, label: type === "recurring" ? "Retainer mensile" : type === "one_shot" ? "Fattura una tantum" : "Fee a evento", amount: "", triggerLabel: null, status: "pending", triggeredDate: null, date: null };
  if (type === "trigger") base.triggerLabel = "Firma del contratto";
  if (type === "one_shot") base.date = today().slice(0, 7) + "-01";
  return base;
};

// aliquota IVA applicata a un contratto (0 se il cliente è in esenzione)
const contractVatRate = c => (c.vatExempt ? 0 : (c.vatRate != null && c.vatRate !== "" ? Number(c.vatRate) : 22));

// compile monthly active invoice rows — recurring items every active month,
// trigger items only in the month their trigger fired, one_shot items in
// the month of their planned date. Deduplica: se per lo stesso cliente in
// quel mese esiste già una fattura importata (ATTIVA), il piano non entra
// in pipeline (evita doppio conteggio nelle proiezioni).
function compileActive(contracts, active, expenses, month, issued = []) {
  // Fatture importate/manuali del mese (per la deduplica con il piano)
  const issuedThisMonth = issued.filter(x => (x.month || (x.date || "").slice(0, 7)) === month);
  return contracts.filter(c => contractActiveInMonth(c, month)).map(c => {
    const rec = active.find(a => a.contractId === c.id && a.month === month);
    const ov = rec?.lineOverrides || {};
    const lines = [];
    getBillingItems(c).forEach(it => {
      const fired = it.type === "trigger" && it.status === "triggered" && (it.triggeredDate || "").slice(0, 7) === month;
      const oneShotFires = it.type === "one_shot" && (it.date || "").slice(0, 7) === month;
      if (it.type === "recurring" || fired || oneShotFires) {
        const amount = ov[it.id] != null ? Number(ov[it.id]) : (Number(it.amount) || 0);
        lines.push({ id: it.id, kind: it.type, label: it.label || (it.type === "recurring" ? "Retainer" : it.type === "one_shot" ? "Una tantum" : "Fee"), triggerLabel: it.triggerLabel, amount });
      }
    });
    const exps = expenses.filter(e => e.contractId === c.id && e.month === month);
    const expTotal = exps.reduce((s, e) => s + (Number(e.amount) || 0), 0);
    // Spese amministrative: voce automatica calcolata come % del retainer (o del retainer + spese di viaggio).
    // Se pct=0 o vuoto, non viene aggiunta. Override manuale per mese possibile via lineOverrides.
    const adminPct = Number(c.adminFeePct) || 0;
    if (adminPct > 0) {
      const recurringBase = lines.filter(l => l.kind === "recurring").reduce((s, l) => s + l.amount, 0);
      const base = c.adminFeeBase === "retainer_travel" ? recurringBase + expTotal : recurringBase;
      const computed = base * adminPct / 100;
      const amount = ov["__admin_fee__"] != null ? Number(ov["__admin_fee__"]) : computed;
      if (amount > 0) lines.push({ id: "__admin_fee__", kind: "admin_fee", label: `Spese amministrative (${adminPct}%)`, amount });
    }
    const feeTotal = lines.reduce((s, l) => s + l.amount, 0);
    const status = rec?.status || "draft";
    const issueDate = rec?.issueDate || lastDayOfMonth(month);
    const dueDate = rec?.dueDate || addDays(issueDate, c.paymentTermsDays);
    const overdue = status !== "paid" && dueDate < today();
    const vatRate = contractVatRate(c);
    const total = feeTotal + expTotal;
    const grossTotal = feeTotal * (1 + vatRate / 100) + expTotal;
    // Dedup: se per questo cliente esistono fatture importate in questo
    // mese (matching flessibile sul nome), la pianificazione è "coperta"
    const matched = issuedThisMonth.filter(x => namesMatch(x.client, c.client));
    const hasIssued = matched.length > 0;
    const issuedSum = matched.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    return { contractId: c.id, client: c.client, currency: c.currency || "EUR", lines, expenses: exps, expTotal, feeTotal, total, grossTotal, vatRate, status, matchedId: rec?.matchedId || null, issueDate, dueDate, terms: Number(c.paymentTermsDays) || 0, overdue, hasIssued, issuedCount: matched.length, issuedSum };
  });
}

// ============ CONTRACTS ============
function Contracts({ ctx }) {
  const { contracts, setContracts, isAdmin, notify } = ctx;
  const [edit, setEdit] = useState(null);
  const upsert = (c) => {
    const exists = contracts.some(x => x.id === c.id);
    setContracts(exists ? contracts.map(x => x.id === c.id ? c : x) : [...contracts, c]);
    setEdit(null); notify(exists ? "Contratto aggiornato" : "Contratto aggiunto");
  };
  const del = (id) => { if (confirm("Eliminare il contratto?")) { setContracts(contracts.filter(c => c.id !== id)); notify("Contratto eliminato"); } };
  const setItem = (cId, itemId, patch) => setContracts(contracts.map(c => c.id !== cId ? c : { ...c, billingItems: getBillingItems(c).map(it => it.id !== itemId ? it : { ...it, ...patch }) }));
  const fireTrigger = (cId, itemId) => {
    const d = prompt("Data dell'evento che attiva il compenso (YYYY-MM-DD):", new Date().toISOString().slice(0, 10));
    if (!d) return;
    setItem(cId, itemId, { status: "triggered", triggeredDate: d });
    notify("Compenso attivato: comparirà tra le fatture da emettere di " + monthLabel(d.slice(0, 7)));
  };
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
        <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), client: "", currency: "EUR", startDate: "", endDate: "", paymentTermsDays: 30, notes: "", vatExempt: false, vatRate: 22, billingItems: [emptyItem("recurring")] })}><Plus size={16} /> Nuovo contratto</button>
      </div>
      {contracts.length === 0 ? (
        <Empty icon={FileSignature} title="Nessun contratto" hint="Aggiungi un contratto manualmente oppure caricane uno dalla sezione «Carica documenti» per estrarre i dati in automatico." />
      ) : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th>Cliente</th><th>Compensi</th><th>Val.</th><th>Periodo</th><th>Pag.</th><th></th></tr></thead>
            <tbody>{contracts.map(c => (
              <tr key={c.id}>
                <td style={{ fontFamily: "Poppins", fontWeight: 600, verticalAlign: "top" }}>{c.client}<div style={{ fontWeight: 400, fontSize: 12, color: C.muted }}>{c.notes}</div></td>
                <td style={{ verticalAlign: "top" }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {getBillingItems(c).length === 0 && <span style={{ color: C.muted }}>—</span>}
                    {getBillingItems(c).map(it => it.type === "recurring" ? (
                      <span key={it.id} className="ia-chip" style={{ background: "#EEF1F6", color: C.navy, alignSelf: "flex-start" }}>{it.label}: {fmtMoney(it.amount, c.currency)}/mese</span>
                    ) : (
                      <span key={it.id} style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                        <span className="ia-chip" style={{ background: "#FFF1E9", color: C.orange }}>▸ {it.triggerLabel || it.label}: {fmtMoney(it.amount, c.currency)}</span>
                        {it.status === "triggered"
                          ? <span style={{ fontSize: 11.5, color: C.green, display: "inline-flex", alignItems: "center", gap: 5 }}><Check size={12} /> attivato {it.triggeredDate}<button title="Annulla attivazione" onClick={() => setItem(c.id, it.id, { status: "pending", triggeredDate: null })} style={{ border: "none", background: "none", cursor: "pointer", color: C.muted }}><Unlink size={12} /></button></span>
                          : <button className="ia-btn ia-btn-ghost" style={{ padding: "3px 9px", fontSize: 11.5 }} onClick={() => fireTrigger(c.id, it.id)}>Attiva</button>}
                      </span>
                    ))}
                  </div>
                </td>
                <td style={{ verticalAlign: "top" }}>{c.currency}</td>
                <td style={{ fontSize: 12.5, verticalAlign: "top" }}>{c.startDate || "—"} → {c.endDate || "in corso"}</td>
                <td style={{ verticalAlign: "top" }}>{c.paymentTermsDays ? c.paymentTermsDays + " gg" : "—"}</td>
                <td style={{ textAlign: "right", whiteSpace: "nowrap", verticalAlign: "top" }}>
                  <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 9px" }} onClick={() => setEdit(c)}><Pencil size={14} /></button>{" "}
                  {isAdmin && <button className="ia-btn ia-btn-danger" style={{ padding: "6px 9px" }} onClick={() => del(c.id)}><Trash2 size={14} /></button>}
                </td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
      {edit && <ContractModal c={edit} onSave={upsert} onClose={() => setEdit(null)} />}
    </div>
  );
}

// shared editor for billing items (recurring + trigger)
function BillingItemsEditor({ items, currency, onChange }) {
  const upd = (id, patch) => onChange(items.map(it => it.id === id ? { ...it, ...patch } : it));
  const add = (type) => onChange([...items, emptyItem(type)]);
  const rm = (id) => onChange(items.filter(it => it.id !== id));
  const onTypeChange = (it, newType) => {
    const patch = { type: newType };
    patch.triggerLabel = newType === "trigger" ? (it.triggerLabel || "Firma del contratto") : null;
    patch.status = newType === "trigger" ? (it.status || "pending") : "pending";
    if (newType === "one_shot" && !it.date) patch.date = today().slice(0, 7) + "-01";
    upd(it.id, patch);
  };
  return (
    <div style={{ marginBottom: 13 }}>
      <label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 6 }}>Voci di fatturazione</label>
      {items.length === 0 && <div style={{ fontSize: 12.5, color: C.muted, marginBottom: 8 }}>Nessuna voce. Aggiungi un retainer mensile, una fattura una tantum o un compenso a evento.</div>}
      {items.map(it => (
        <div key={it.id} style={{ border: "1px solid " + C.line, padding: 10, marginBottom: 8, background: "#FBFCFE" }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            <select className="ia-input" style={{ width: 200 }} value={it.type} onChange={e => onTypeChange(it, e.target.value)}>
              <option value="recurring">Mensile (ricorrente)</option>
              <option value="one_shot">Una tantum (one shot)</option>
              <option value="trigger">A evento (trigger)</option>
            </select>
            <input className="ia-input" type="number" placeholder={"Importo " + currency + " (netto)"} style={{ flex: 1 }} value={it.amount} onChange={e => upd(it.id, { amount: e.target.value })} />
            <button className="ia-btn ia-btn-ghost" style={{ padding: "0 11px" }} onClick={() => rm(it.id)}><X size={14} /></button>
          </div>
          <input className="ia-input" placeholder="Descrizione voce" value={it.label} onChange={e => upd(it.id, { label: e.target.value })} style={{ marginBottom: (it.type === "trigger" || it.type === "one_shot") ? 8 : 0 }} />
          {it.type === "trigger" && <input className="ia-input" placeholder="Evento che attiva il compenso (es. Firma, Milestone 1, Closing)" value={it.triggerLabel || ""} onChange={e => upd(it.id, { triggerLabel: e.target.value })} />}
          {it.type === "one_shot" && (
            <div>
              <label style={{ fontSize: 11, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 4 }}>Data prevista di fatturazione</label>
              <input className="ia-input" type="date" value={it.date || ""} onChange={e => upd(it.id, { date: e.target.value })} />
            </div>
          )}
        </div>
      ))}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <button className="ia-btn ia-btn-ghost" style={{ fontSize: 12 }} onClick={() => add("recurring")}><Plus size={13} /> Voce mensile</button>
        <button className="ia-btn ia-btn-ghost" style={{ fontSize: 12 }} onClick={() => add("one_shot")}><Plus size={13} /> Una tantum</button>
        <button className="ia-btn ia-btn-ghost" style={{ fontSize: 12 }} onClick={() => add("trigger")}><Plus size={13} /> Voce a evento</button>
      </div>
    </div>
  );
}

function ContractModal({ c, onSave, onClose }) {
  const [f, setF] = useState(() => ({ vatExempt: false, vatRate: 22, ...c, billingItems: getBillingItems(c).length ? getBillingItems(c) : [emptyItem("recurring")] }));
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const save = () => onSave({ ...f, vatRate: f.vatExempt ? 0 : (Number(f.vatRate) || 22), retainerDefault: null, billingItems: f.billingItems.map(it => ({ ...it, amount: it.amount === "" ? 0 : Number(it.amount) })) });
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 520, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{c.client ? "Modifica contratto" : "Nuovo contratto"}</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
            <F label="Cliente *"><input className="ia-input" value={f.client} onChange={e => set("client", e.target.value)} /></F>
            <F label="Valuta"><select className="ia-input" value={f.currency} onChange={e => set("currency", e.target.value)}><option>EUR</option><option>USD</option></select></F>
          </div>
          <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 13, background: "#FBFCFE" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
              <input type="checkbox" checked={!!f.vatExempt} onChange={e => set("vatExempt", e.target.checked)} />
              Cliente in <b>esenzione IVA</b> (es. estero art. 7-ter, non imponibile art. 8, ecc.)
            </label>
            {!f.vatExempt && (
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                <label style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins" }}>Aliquota IVA (%)</label>
                <input className="ia-input" type="number" style={{ width: 90 }} value={f.vatRate} onChange={e => set("vatRate", e.target.value)} />
                <span style={{ fontSize: 12, color: C.muted }}>Applicata solo ai compensi; le spese di viaggio sono trattate come rimborso.</span>
              </div>
            )}
          </div>
          <BillingItemsEditor items={f.billingItems} currency={f.currency} onChange={v => set("billingItems", v)} />
          <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 13, background: "#FBFCFE" }}>
            <div style={{ fontSize: 13, fontFamily: "Poppins", color: C.ink, marginBottom: 8, fontWeight: 500 }}>Spese amministrative</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 12 }}>
              <F label="Percentuale (%)">
                <input className="ia-input" type="number" step="0.1" value={f.adminFeePct == null ? "" : f.adminFeePct} onChange={e => set("adminFeePct", e.target.value === "" ? "" : Number(e.target.value))} placeholder="es. 8" />
              </F>
              <F label="Base di calcolo">
                <select className="ia-input" value={f.adminFeeBase || "retainer"} onChange={e => set("adminFeeBase", e.target.value)}>
                  <option value="retainer">Solo retainer</option>
                  <option value="retainer_travel">Retainer + spese di viaggio</option>
                </select>
              </F>
            </div>
            <div style={{ fontSize: 11.5, color: C.muted, marginTop: 6 }}>Voce automatica aggiunta ogni mese alle fatture pianificate. Lasciare vuoto o 0 per disattivare. L'importo calcolato è editabile a mano per singolo mese.</div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Inizio"><input className="ia-input" type="date" value={f.startDate} onChange={e => set("startDate", e.target.value)} /></F>
            <F label="Fine (vuoto = in corso)"><input className="ia-input" type="date" value={f.endDate} onChange={e => set("endDate", e.target.value)} /></F>
          </div>
          <F label="Termini di pagamento (giorni)"><input className="ia-input" type="number" value={f.paymentTermsDays} onChange={e => set("paymentTermsDays", e.target.value)} /></F>
          <F label="Note"><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.client} onClick={save}>Salva contratto</button>
        </div>
      </div>
    </div>
  );
}

// ============ ACTIVE INVOICES ============
function ActiveInvoices({ ctx }) {
  const { contracts, active, setActive, expenses, setExpenses, issued, setIssued, month, notify } = ctx;
  const rows = compileActive(contracts, active, expenses, month, issued).filter(r => r.lines.length || r.expTotal);
  const upsertRec = (contractId, patch) => {
    const i = active.findIndex(a => a.contractId === contractId && a.month === month);
    if (i >= 0) setActive(active.map((a, idx) => idx === i ? { ...a, ...patch } : a));
    else setActive([...active, { contractId, month, status: "draft", matchedId: null, lineOverrides: {}, ...patch }]);
  };
  const setLineAmount = (contractId, lineId, val) => {
    const rec = active.find(a => a.contractId === contractId && a.month === month);
    upsertRec(contractId, { lineOverrides: { ...(rec?.lineOverrides || {}), [lineId]: val === "" ? 0 : Number(val) } });
  };
  const addExpManual = (contractId) => {
    const d = prompt("Descrizione spesa (es. Trasferta Milano-Roma):"); if (d == null) return;
    const a = prompt("Importo (€):"); const amount = parseItAmount(a); if (isNaN(amount)) { notify("Importo non valido", "err"); return; }
    setExpenses([...expenses, { id: uid(), contractId, month, description: d || "Spesa", amount, currency: "EUR", fileName: "(manuale)" }]);
    notify("Spesa aggiunta");
  };
  const delExp = (id) => setExpenses(expenses.filter(e => e.id !== id));
  const total = rows.reduce((s, r) => s + r.total, 0);

  const impRef = useRef();
  const [newIssued, setNewIssued] = useState(null); // modale "nuova fattura una tantum"
  const [iaIncInvoice, setIaIncInvoice] = useState(null); // modale IA Inc. (US)
  const [reportOpen, setReportOpen] = useState(false);
  const importAttivaFile = async (file) => {
    try {
      const parsed = importAttiva(await readTabular(file));
      if (!parsed.length) { notify("Nessuna fattura riconosciuta nel file", "err"); return; }
      const { merged, added, dup } = mergeById(issued, parsed, issuedKey);
      setIssued(merged); notify(`${added} fatture emesse importate${dup ? `, ${dup} già presenti` : ""}`);
    } catch (e) { notify(e.message || "Import non riuscito", "err"); }
  };
  const saveNewIssued = (x) => {
    setIssued([...issued, { ...x, id: uid(), month: (x.date || "").slice(0, 7), status: "issued", matchedId: null, source: "manual", currency: x.currency || "EUR" }]);
    setNewIssued(null); notify("Fattura una tantum registrata");
  };
  // Upsert per fatture IA Inc. (modifica esistente oppure creazione nuova)
  const saveIAInc = (record) => {
    const exists = issued.some(y => y.id === record.id);
    setIssued(exists ? issued.map(y => y.id === record.id ? record : y) : [record, ...issued]);
    setIaIncInvoice(null); notify(exists ? "Fattura IA Inc. aggiornata" : "Fattura IA Inc. registrata");
  };
  // Suite Inc.: fattura emessa o aggiornata nel registro (numero assegnato dal server)
  const saveIAIncReg = (inv, pdfErr) => {
    setIssued(syncIssuedWithReg(issued, inv));
    setIaIncInvoice(null);
    notify(pdfErr ? "Invoice " + inv.number + " saved, but the PDF was not archived: " + pdfErr : "Invoice " + inv.number + " issued and archived in the register", pdfErr ? "err" : "ok");
  };
  const voidRegFromList = async (x) => {
    const reason = prompt("Void invoice " + x.number + "?\nIt stays in the invoice register with its number, marked VOID.\n\nReason:");
    if (!reason || !reason.trim()) return;
    try { const { invoice } = await regVoid(x.id, reason.trim()); setIssued(syncIssuedWithReg(issued, invoice)); notify("Invoice " + x.number + " voided"); }
    catch (e) { notify(e.message, "err"); }
  };
  // Trova e rimuove duplicati esatti: stesso cliente (normalizzato), stesso numero,
  // stessa data, stesso |importo|. Tiene il primo record di ogni gruppo.
  const cleanupIssuedDuplicates = () => {
    const groups = new Map();
    issued.forEach(x => {
      const key = issuedKey(x);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(x);
    });
    const duplicates = [];
    groups.forEach(arr => { if (arr.length > 1) duplicates.push(...arr.slice(1)); });
    if (duplicates.length === 0) { notify("Nessun duplicato trovato"); return; }
    const preview = duplicates.slice(0, 5).map(x => `• ${x.client} — ${x.number || "s.n."} — ${fmtDate(x.date)} — ${fmtMoney(x.amount)}`).join("\n");
    if (!confirm(`Trovati ${duplicates.length} duplicati (stesso cliente + numero + data + importo).\n\nPrimi da eliminare:\n${preview}${duplicates.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toRemove = new Set(duplicates.map(x => x.id));
    setIssued(issued.filter(x => !toRemove.has(x.id)));
    notify(`${duplicates.length} duplicati rimossi`);
  };
  // Corregge le NC legacy: record con "credito"/"credit note"/"NC"/"TD04" nel docType (o numero)
  // ma amount POSITIVO — importati prima del fix di detection automatica.
  const fixLegacyCreditNotes = () => {
    const toFix = issued.filter(x => {
      const dt = String(x.docType || "").toLowerCase();
      const nm = String(x.number || "").toLowerCase();
      const isNCByDoc = dt.includes("credito") || dt.includes("credit note") || dt === "td04" || /\bnc\b/.test(dt);
      const isNCByNumber = /^nc[\s\-\/]/i.test(x.number || "") || nm.includes("nota credito") || nm.includes("nota di credito");
      const isNC = isNCByDoc || isNCByNumber || x.isCreditNote;
      return isNC && (Number(x.amount) || 0) > 0;
    });
    if (toFix.length === 0) { notify("Nessuna nota di credito legacy da correggere"); return; }
    const preview = toFix.slice(0, 5).map(x => `• ${x.client} — ${x.number || "s.n."} — ${fmtDate(x.date)} — ${fmtMoney(x.amount)}`).join("\n");
    if (!confirm(`Trovate ${toFix.length} note di credito con importo positivo (importate prima del fix).\nLe converto a importo NEGATIVO e le marco come "nota di credito".\n\nPrime da correggere:\n${preview}${toFix.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toFixIds = new Set(toFix.map(x => x.id));
    setIssued(issued.map(x => toFixIds.has(x.id) ? { ...x, amount: -Math.abs(Number(x.amount) || 0), isCreditNote: true } : x));
    notify(`${toFix.length} note di credito corrette`);
  };
  const monthIssued = issued.filter(x => (x.month || (x.date || "").slice(0, 7)) === month);
  const issuedTotal = monthIssued.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const delIssued = (id) => setIssued(issued.filter(x => x.id !== id));

  return (
    <div>
      {/* Fatture emesse importate (export ATTIVA) */}
      <div className="ia-panel" style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", borderBottom: "1px solid " + C.line }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 14.5, color: C.navy }}>Fatture emesse importate — {monthLabel(month)}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Dall'export ATTIVA del gestionale · usate per la riconciliazione con gli incassi.</div>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
            {monthIssued.length > 0 && <span className="ia-h" style={{ fontSize: 15, color: C.navy }}>{fmtMoney(issuedTotal)}</span>}
            <button className="ia-btn ia-btn-ghost" onClick={() => setReportOpen(true)}><FileSignature size={15} /> Riepilogo mese</button>
            <button className="ia-btn ia-btn-ghost" onClick={cleanupIssuedDuplicates} title="Trova e rimuove fatture doppie (stesso cliente + numero + data + importo)"><RefreshCw size={14} /> Ripulisci duplicati</button>
            <button className="ia-btn ia-btn-ghost" onClick={fixLegacyCreditNotes} title="Converte a importo negativo le note di credito importate prima del fix automatico"><RefreshCw size={14} /> Correggi NC legacy</button>
            <button className="ia-btn ia-btn-ghost" onClick={() => setNewIssued({ client: "", number: "", date: today(), dueDate: "", amount: "", currency: "EUR", vatExempt: false, vatRate: 22, notes: "" })}><Plus size={15} /> Nuova una tantum</button>
            {IS_INC
              ? <button className="ia-btn ia-btn-dark" onClick={() => setIaIncInvoice({})} title="Issue a new invoice: the number is assigned automatically and the PDF is archived in the invoice register"><Globe size={15} /> Fattura IA Inc.</button>
              : <a className="ia-btn ia-btn-dark" href="https://inc-suite.ia-advisors-hub.com/?tab=register" target="_blank" rel="noopener" style={{ textDecoration: "none" }} title="Le fatture IA Inc. si emettono dalla IA Suite Inc., che assegna il numero progressivo e archivia il PDF nel registro"><Globe size={15} /> Fattura IA Inc.</a>}
            <button className="ia-btn ia-btn-primary" onClick={() => impRef.current.click()}><Upload size={15} /> Importa Excel</button>
            <input ref={impRef} type="file" accept=".xlsx,.xls" hidden onChange={e => { if (e.target.files[0]) importAttivaFile(e.target.files[0]); e.target.value = ""; }} />
          </div>
        </div>
        {monthIssued.length === 0 ? (
          <div style={{ padding: 18, fontSize: 13, color: C.muted }}>Nessuna fattura emessa importata per {monthLabel(month)}. Importa l'elenco ATTIVA (Excel) qui o da «Carica documenti».</div>
        ) : (
          <div style={{ overflowX: "auto" }}><table className="ia-table">
            <thead><tr><th>Cliente</th><th>Numero</th><th>Data</th><th style={{ textAlign: "right" }}>Totale</th><th>Stato</th><th></th></tr></thead>
            <tbody>{monthIssued.map(x => {
              const isCredit = !!x.isCreditNote || (Number(x.amount) || 0) < 0;
              const isIAInc = !!x.iaInc;
              return (
                <tr key={x.id} style={isCredit ? { background: "#FFF5F5" } : (isIAInc ? { background: "#F5F9FE" } : undefined)}>
                  <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>
                    {x.client}
                    {isCredit && <span className="ia-chip" style={{ background: "#FBECEC", color: C.red, marginLeft: 8, fontSize: 10 }}>nota di credito</span>}
                    {isIAInc && <span className="ia-chip" style={{ background: "#EBF2FB", color: C.blue, marginLeft: 8, fontSize: 10 }}>IA Inc. · USD</span>}
                    <div style={{ fontWeight: 400, fontSize: 11.5, color: C.muted }}>{x.fteStatus}</div>
                  </td>
                  <td>{x.number || "—"}</td><td>{fmtDate(x.date)}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: isCredit ? C.red : C.ink }}>{fmtMoney(x.amount, x.currency)}</td>
                  <td><Chip status={x.status === "paid" ? "paid" : "issued"} /></td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    {isIAInc && IS_INC && x.reg && (
                      <>
                        <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => setIaIncInvoice(x)} title="Edit and archive a new version (the number does not change)"><Pencil size={13} /></button>{" "}
                        <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => { if (!openRegPdf(x)) notify("No archived PDF yet: open the Invoice register to create it", "err"); }} title="Open the archived PDF"><FileSignature size={13} /></button>{" "}
                      </>
                    )}
                    {isIAInc && !(IS_INC && x.reg) && (
                      <>
                        <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => setIaIncInvoice(x)} title="Apri, modifica e rigenera il PDF"><Pencil size={13} /></button>{" "}
                        <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => { const w = window.open("", "_blank"); if (w) { w.document.write(buildIAIncInvoiceHTML(x)); w.document.close(); w.focus(); } }} title="Rigenera PDF senza modifiche"><FileSignature size={13} /></button>{" "}
                      </>
                    )}
                    {IS_INC && x.reg
                      ? <button className="ia-btn ia-btn-danger" style={{ padding: "5px 8px" }} onClick={() => voidRegFromList(x)} title="Void: the invoice stays in the register with its number"><Ban size={13} /></button>
                      : <button className="ia-btn ia-btn-danger" style={{ padding: "5px 8px" }} onClick={() => delIssued(x.id)}><Trash2 size={13} /></button>}
                  </td>
                </tr>
              );
            })}</tbody>
          </table></div>
        )}
      </div>

      {/* Pianificazione da contratti */}
      <h3 className="ia-h" style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: ".6px", margin: "0 0 10px" }}>Pianificazione da contratti</h3>
      {contracts.length === 0 ? (
        <Empty icon={ArrowUpFromLine} title="Nessun contratto per la pianificazione" hint="Aggiungi contratti per generare il piano mensile (retainer, compensi a evento, spese). Le fatture già emesse le importi qui sopra dall'Excel." />
      ) : (<>
        <div className="ia-panel ia-bracket" style={{ padding: "16px 20px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div><div style={{ fontSize: 12.5, color: C.muted, fontFamily: "Poppins" }}>Totale pianificato da emettere · {monthLabel(month)}</div><div className="ia-kpi-num" style={{ marginTop: 4 }}>{fmtMoney(total)}</div></div>
          <div style={{ fontSize: 12.5, color: C.muted, textAlign: "right" }}>Retainer ricorrenti + compensi a evento attivati nel mese.<br />Importi modificabili riga per riga; spese sommate dai PDF.</div>
        </div>
        {rows.map(r => (
        <div key={r.contractId} className="ia-panel" style={{ marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 18px", borderBottom: "1px solid " + C.line }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
              <span className="ia-h" style={{ fontSize: 15, color: C.navy }}>{r.client}</span><Chip status={r.status} />
              {r.overdue && <span className="ia-chip" style={{ background: "#FBECEC", color: C.red }}><AlertTriangle size={12} /> scaduta</span>}
              {r.hasIssued && <span className="ia-chip" style={{ background: "#EAF6EE", color: C.green }} title={`Già presenti ${r.issuedCount} fatture emesse in ATTIVA per ${fmtMoney(r.issuedSum)} — la pianificazione non viene contata due volte in cassa`}><Check size={12} /> già fatturato ({r.issuedCount})</span>}
              {r.vatRate === 0 && <span className="ia-chip" style={{ background: "#EEF1F6", color: C.muted }}>esente IVA</span>}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ textAlign: "right", lineHeight: 1.15 }}>
                <div className="ia-h" style={{ fontSize: 17, color: C.navy }}>{fmtMoney(r.grossTotal, r.currency)}</div>
                {r.vatRate > 0 && <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins" }}>{fmtMoney(r.total, r.currency)} netto · +{r.vatRate}% IVA</div>}
              </div>
              {r.status === "draft" && <button className="ia-btn ia-btn-dark" onClick={() => upsertRec(r.contractId, { status: "issued" })}><Check size={15} /> Segna emessa</button>}
              {r.status === "issued" && <button className="ia-btn ia-btn-ghost" onClick={() => upsertRec(r.contractId, { status: "draft" })}>Riporta a bozza</button>}
            </div>
          </div>
          <div style={{ display: "flex", gap: 24, flexWrap: "wrap", alignItems: "flex-end", padding: "12px 18px", background: "#FAFBFD", borderBottom: "1px solid " + C.line }}>
            <div>
              <label style={{ fontSize: 11, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 4 }}>Data emissione</label>
              <input className="ia-input" type="date" style={{ width: 160 }} value={r.issueDate} onChange={e => upsertRec(r.contractId, { issueDate: e.target.value, dueDate: null })} />
            </div>
            <div>
              <label style={{ fontSize: 11, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 4 }}>Data scadenza {r.terms ? `(+${r.terms} gg)` : ""}</label>
              <input className="ia-input" type="date" style={{ width: 160, color: r.overdue ? C.red : C.ink, fontWeight: 600, borderColor: r.overdue ? "#E7C9C9" : C.line }} value={r.dueDate} onChange={e => upsertRec(r.contractId, { dueDate: e.target.value })} />
            </div>
            <div style={{ fontSize: 11.5, color: C.muted, paddingBottom: 8 }}>Scadenza calcolata dai termini del contratto.<br />Modificabile a mano per singola fattura.</div>
          </div>
          <div style={{ padding: "14px 18px" }}>
            {r.lines.length === 0 ? <div style={{ fontSize: 12.5, color: C.muted, marginBottom: 12 }}>Nessun compenso questo mese — solo spese.</div> : (
              <div style={{ marginBottom: 14 }}>
                {r.lines.map(l => (
                  <div key={l.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, padding: "7px 0", borderBottom: "1px dashed " + C.line }}>
                    <span style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5 }}>
                      {l.kind === "trigger"
                        ? <span className="ia-chip" style={{ background: "#FFF1E9", color: C.orange }}>a evento</span>
                        : l.kind === "admin_fee"
                          ? <span className="ia-chip" style={{ background: "#F1EEF8", color: "#5B4A8C" }}>amministrative</span>
                          : <span className="ia-chip" style={{ background: "#EEF1F6", color: C.navy }}>mensile</span>}
                      {l.label}{l.kind === "trigger" && l.triggerLabel ? <span style={{ color: C.muted, fontSize: 12 }}>· {l.triggerLabel}</span> : null}
                    </span>
                    <input className="ia-input" type="number" style={{ width: 150 }} value={l.amount} onChange={e => setLineAmount(r.contractId, l.id, e.target.value)} />
                  </div>
                ))}
              </div>
            )}
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <label style={{ fontSize: 11.5, fontFamily: "Poppins", fontWeight: 500, color: C.muted }}>Spese di viaggio ({fmtMoney(r.expTotal, r.currency)})</label>
                <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 9px", fontSize: 12 }} onClick={() => addExpManual(r.contractId)}><Plus size={13} /> Aggiungi manuale</button>
              </div>
              {r.expenses.length === 0 ? <div style={{ fontSize: 12.5, color: C.muted }}>Nessuna spesa. Carica un PDF spese dalla sezione «Carica documenti».</div> :
                r.expenses.map(e => (
                  <div key={e.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "5px 0", borderBottom: "1px dashed " + C.line }}>
                    <span>{e.description} <span style={{ color: C.muted, fontSize: 11.5 }}>· {e.fileName}</span></span>
                    <span style={{ display: "flex", alignItems: "center", gap: 8 }}><b style={{ fontFamily: "Poppins" }}>{fmtMoney(e.amount, e.currency)}</b>
                      <button onClick={() => delExp(e.id)} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={14} color={C.muted} /></button></span>
                  </div>
                ))}
            </div>
          </div>
        </div>
      ))}
      </>)}
      {newIssued && <NewIssuedModal x={newIssued} onSave={saveNewIssued} onClose={() => setNewIssued(null)} />}
      {iaIncInvoice && <IAIncInvoiceModal x={iaIncInvoice} existingIssued={issued} register={IS_INC && (!iaIncInvoice.id || !!iaIncInvoice.reg)} onSave={saveIAInc} onRegSaved={saveIAIncReg} onClose={() => setIaIncInvoice(null)} />}
      {reportOpen && <MonthlyReport month={month} ctx={ctx} onClose={() => setReportOpen(false)} />}
    </div>
  );
}

function NewIssuedModal({ x, onSave, onClose }) {
  const [f, setF] = useState(x);
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const vatRate = f.vatExempt ? 0 : (Number(f.vatRate) || 0);
  const gross = (Number(f.amount) || 0) * (1 + vatRate / 100);
  const ok = f.client && f.date && f.amount !== "" && Number(f.amount) > 0;
  const isCredit = !!f.isCreditNote;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 480, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{isCredit ? "Nuova nota di credito" : "Nuova fattura una tantum"}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>{isCredit ? "L'importo verrà sottratto dai ricavi." : "Fattura emessa non collegata a un contratto ricorrente."}</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <div style={{ background: isCredit ? "#FBECEC" : "#FBFCFE", border: "1px solid " + (isCredit ? "#E7C9C9" : C.line), padding: 10, marginBottom: 12 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
              <input type="checkbox" checked={isCredit} onChange={e => set("isCreditNote", e.target.checked)} />
              È una <b>nota di credito</b> (riduce il fatturato)
            </label>
          </div>
          <F label="Cliente *"><input className="ia-input" value={f.client} onChange={e => set("client", e.target.value)} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
            <F label={isCredit ? "Numero nota" : "Numero fattura"}><input className="ia-input" value={f.number} onChange={e => set("number", e.target.value)} placeholder={isCredit ? "es. NC 2026/3" : "es. 2026/12"} /></F>
            <F label="Valuta"><select className="ia-input" value={f.currency} onChange={e => set("currency", e.target.value)}><option>EUR</option><option>USD</option></select></F>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Data emissione *"><input className="ia-input" type="date" value={f.date} onChange={e => set("date", e.target.value)} /></F>
            <F label="Data scadenza"><input className="ia-input" type="date" value={f.dueDate} onChange={e => set("dueDate", e.target.value)} /></F>
          </div>
          <F label="Imponibile (netto) *"><input className="ia-input" type="number" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
          <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 13, background: "#FBFCFE" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
              <input type="checkbox" checked={!!f.vatExempt} onChange={e => set("vatExempt", e.target.checked)} />
              Cliente in <b>esenzione IVA</b>
            </label>
            {!f.vatExempt && (
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                <label style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins" }}>Aliquota IVA (%)</label>
                <input className="ia-input" type="number" style={{ width: 90 }} value={f.vatRate} onChange={e => set("vatRate", e.target.value)} />
              </div>
            )}
            <div style={{ marginTop: 10, fontSize: 12.5, color: C.muted, display: "flex", justifyContent: "space-between", paddingTop: 8, borderTop: "1px solid " + C.line }}>
              <span>{isCredit ? "Storno atteso (lordo)" : "Incasso atteso (lordo)"}</span><b style={{ fontFamily: "Poppins", color: isCredit ? C.red : C.navy }}>{isCredit ? "−" : ""}{fmtMoney(gross, f.currency)}</b>
            </div>
          </div>
          <F label="Note"><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!ok} onClick={() => {
            const amt = Number(f.amount);
            onSave({ ...f, amount: isCredit ? -Math.abs(amt) : Math.abs(amt), vatRate: f.vatExempt ? 0 : (Number(f.vatRate) || 0), docType: isCredit ? "Nota di credito" : (f.docType || "Fattura") });
          }}>Salva{isCredit ? " nota di credito" : ""}</button>
        </div>
      </div>
    </div>
  );
}

// ============ FATTURA IA INC. — MODALE DEDICATA + GENERAZIONE PDF ============
// Layout dedicato per l'affiliata US (IA International Advisors Inc., New York).
// Salva la fattura come emessa (in USD) con flag iaInc:true e produce un PDF
// coerente con il template ufficiale della sede US (logo originale IA Inc. + coordinate NY + JP Morgan).

// Dati fissi della sede US — se cambiano, aggiornali qui.
const IA_INC_INFO = {
  companyName: "IA INTERNATIONAL ADVISORS INC.",
  address1: "450 7th Ave., Suite 1504 New York, NY",
  address2: "10123 - USA",
  email: "yazmin.sanchez@iainternationaladvisors.com",
  phone: "Phone: +1 (778) 939 -1646",
  bank: "JP Morgan Chase",
  bankAccount: "599725077",
  bankRouting: "021000021",
  bankSwift: "CHASUS33",
};

function fmtUSD(n) {
  const v = Number(n) || 0;
  return "$ " + v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Logo originale IA Inc. (public/logo-inc-fattura.png). URL assoluto: la fattura si apre in una finestra vuota.
function IA_INC_LOGO_URL() { return (typeof window !== "undefined" ? window.location.origin : "") + "/logo-inc-fattura.png"; }

function buildIAIncInvoiceHTML(f) {
  const services = (f.services || []).filter(s => s.description || Number(s.amount));
  const servicesTotal = services.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const adminCost = servicesTotal * (Number(f.adminCostPercent) || 0) / 100;
  const subtotal = servicesTotal + adminCost;
  const taxAmount = subtotal * (Number(f.taxPercent) || 0) / 100;
  const otherAmount = Number(f.otherAmount) || 0;
  const total = subtotal + taxAmount + otherAmount;
  const esc = s => String(s || "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const addrHtml = esc(f.clientAddress).replace(/\n/g, "<br>");
  const linesHtml = services.map((s, i) => `
    <tr>
      <td class="c-date">${i === 0 ? esc(f.date) : ""}</td>
      <td class="c-desc">${esc(s.description)}</td>
      <td class="c-amt">${fmtUSD(Number(s.amount) || 0)}</td>
    </tr>`).join("");
  const adminRow = adminCost > 0 ? `
    <tr>
      <td class="c-date"></td>
      <td class="c-desc"> + ${Number(f.adminCostPercent) || 0}% administrative cost</td>
      <td class="c-amt">${fmtUSD(adminCost)}</td>
    </tr>` : "";
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Invoice ${esc(f.number)}</title>
<style>
  @page { size: letter portrait; margin: 0.6in; }
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Calibri, sans-serif; color: #333; margin: 0; font-size: 10.5pt; line-height: 1.35; }
  .logo { display: block; height: 0.85in; width: auto; }
  .header { display: table; width: 100%; margin-bottom: 26px; }
  .h-left { display: table-cell; width: 60%; vertical-align: top; }
  .h-right { display: table-cell; width: 40%; vertical-align: top; text-align: right; padding-top: 20px; }
  .co-info { margin-top: 14px; font-size: 9pt; line-height: 1.4; }
  .co-info b { display: block; font-size: 9.5pt; margin-bottom: 2px; }
  .to-box { display: inline-block; border: 1px solid #333; padding: 5px 15px 8px; text-align: right; min-width: 240px; }
  .to-lbl { text-align: left; font-weight: bold; margin-bottom: 3px; font-size: 10pt; }
  .to-body { color: #29388D; font-size: 10pt; }
  .to-body .name { font-weight: 700; }
  .inv-title { text-align: center; color: #29388D; font-family: 'Poppins', sans-serif; font-size: 13pt; font-weight: 400; margin: 30px 0 25px; }
  table.items { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
  table.items th { text-align: left; color: #29388D; font-family: 'Poppins', sans-serif; font-size: 9.5pt; font-weight: 700; padding: 4px 0 8px; }
  table.items th.amt { text-align: right; }
  table.items td { padding: 8px 4px; font-size: 10pt; vertical-align: top; }
  table.items td.c-date { color: #333; white-space: nowrap; padding-right: 12px; }
  table.items td.c-desc { color: #333; border-bottom: 1px dotted #B3033C; }
  table.items td.c-amt { text-align: right; white-space: nowrap; border-bottom: 1px dotted #B3033C; padding-left: 12px; }
  .totals { margin-top: 10px; width: 100%; }
  .totals td { padding: 3px 0; font-size: 10pt; }
  .totals td.lbl { text-align: right; color: #29388D; font-family: 'Poppins', sans-serif; font-weight: 700; padding-right: 20px; width: 88%; letter-spacing: 0.4px; }
  .totals td.val { text-align: right; white-space: nowrap; min-width: 100px; }
  .totals td.val.strong { font-weight: 700; }
  .footer { margin-top: 45px; }
  .terms { text-align: center; font-size: 10pt; margin-bottom: 25px; }
  .terms .l { color: #29388D; font-family: 'Poppins', sans-serif; font-weight: 700; margin-right: 40px; letter-spacing: 0.4px; }
  .payment { text-align: right; font-size: 10pt; line-height: 1.5; }
  .payment .title { color: #29388D; font-family: 'Poppins', sans-serif; font-weight: 700; letter-spacing: 0.4px; }
  .payment .bank { color: #29388D; font-family: 'Poppins', sans-serif; font-weight: 700; }
  .no-print { position: fixed; top: 12px; right: 12px; background: #B3033C; color: #fff; border: none; padding: 8px 14px; font-family: 'Poppins', sans-serif; font-weight: 600; cursor: pointer; font-size: 11pt; box-shadow: 0 2px 6px rgba(0,0,0,.15); }
  @media print { .no-print { display: none !important; } body { margin: 0; } }
</style></head>
<body>
  <button class="no-print" onclick="window.print()">Stampa / Salva PDF</button>
  <div class="header">
    <div class="h-left">
      <img class="logo" src="${IA_INC_LOGO_URL()}" alt="International Advisors Inc.">
      <div class="co-info">
        <b>${esc(IA_INC_INFO.companyName)}</b>
        ${esc(IA_INC_INFO.address1)}<br>
        ${esc(IA_INC_INFO.address2)}<br>
        ${esc(IA_INC_INFO.email)}<br>
        ${esc(IA_INC_INFO.phone)}
      </div>
    </div>
    <div class="h-right">
      <div class="to-box">
        <div class="to-lbl">TO</div>
        <div class="to-body"><div class="name">${esc(f.clientName)}</div>${addrHtml}</div>
      </div>
    </div>
  </div>
  <div class="inv-title">Invoice # ${esc(f.number)}</div>
  <table class="items">
    <thead><tr><th style="width:20%">INVOICE DATE</th><th style="width:55%">SERVICE DESCRIPTION</th><th class="amt" style="width:25%">AMOUNT</th></tr></thead>
    <tbody>${linesHtml}${adminRow}</tbody>
  </table>
  <table class="totals">
    <tr><td class="lbl">SUBTOTAL</td><td class="val strong">${fmtUSD(subtotal)}</td></tr>
    <tr><td class="lbl">TAX</td><td class="val">${(Number(f.taxPercent) || 0).toFixed(2)}%</td></tr>
    <tr><td class="lbl">OTHER</td><td class="val">${otherAmount > 0 ? fmtUSD(otherAmount) : "-"}</td></tr>
    <tr><td class="lbl">TOTAL</td><td class="val strong">${fmtUSD(total)}</td></tr>
  </table>
  <div class="footer">
    <div class="terms"><span class="l">Terms</span>${esc(f.terms || "Due upon receipt")}</div>
    <div class="payment">
      <div class="title">PAYMENT INFORMATION</div>
      <div class="bank">${esc(IA_INC_INFO.bank)}</div>
      Account number ${esc(IA_INC_INFO.bankAccount)}<br>
      Routing ${esc(IA_INC_INFO.bankRouting)}<br>
      SWIFT ${esc(IA_INC_INFO.bankSwift)}
    </div>
  </div>
</body></html>`;
}

function computeIAIncTotal(f) {
  const services = (f.services || []).filter(s => s.description || Number(s.amount));
  const servicesTotal = services.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const adminCost = servicesTotal * (Number(f.adminCostPercent) || 0) / 100;
  const subtotal = servicesTotal + adminCost;
  const taxAmount = subtotal * (Number(f.taxPercent) || 0) / 100;
  const otherAmount = Number(f.otherAmount) || 0;
  return { servicesTotal, adminCost, subtotal, taxAmount, otherAmount, total: subtotal + taxAmount + otherAmount };
}

function IAIncInvoiceModal({ x, existingIssued, onSave, onClose, register = false, onRegSaved }) {
  // register = Suite Inc.: numero assegnato dal server all'emissione, PDF archiviato nel registro
  const isEdit = register ? !!x?.reg : !!(x && x.id && existingIssued.some(i => i.id === x.id));
  const [busy, setBusy] = useState(false);
  const [regErr, setRegErr] = useState("");
  const [nextPreview, setNextPreview] = useState(null);
  const [f, setF] = useState(() => ({
    id: x?.id || uid(),
    clientName: x?.client || "",
    clientAddress: x?.clientAddress || "",
    number: x?.number || "",
    date: x?.date || today(),
    services: x?.services && x.services.length ? x.services : [{ description: "", amount: "" }],
    adminCostPercent: x?.adminCostPercent ?? 5,
    taxPercent: x?.taxPercent ?? 0,
    otherAmount: x?.otherAmount ?? 0,
    terms: x?.terms || "Due upon receipt",
    notes: x?.notes || "",
  }));
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const updSvc = (i, k, v) => setF(prev => ({ ...prev, services: prev.services.map((s, idx) => idx === i ? { ...s, [k]: v } : s) }));
  const addSvc = () => setF(prev => ({ ...prev, services: [...prev.services, { description: "", amount: "" }] }));
  const rmSvc = (i) => setF(prev => ({ ...prev, services: prev.services.length > 1 ? prev.services.filter((_, idx) => idx !== i) : prev.services }));

  // Auto-suggerimento del prossimo numero fattura IA Inc. (formato "NNN/YYYY")
  useEffect(() => {
    if (register || f.number || isEdit) return;
    const nums = existingIssued.filter(i => i.iaInc).map(i => {
      const m = String(i.number || "").match(/^0*(\d+)\s*[/\-]/);
      return m ? parseInt(m[1], 10) : 0;
    });
    const max = nums.length ? Math.max(...nums) : 0;
    const year = new Date().getFullYear();
    setF(prev => ({ ...prev, number: String(max + 1).padStart(3, "0") + "/" + year }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Anteprima del prossimo numero (indicativa: quello vero lo assegna il server all'emissione)
  const dateYear = Number(String(f.date || "").slice(0, 4)) || new Date().getFullYear();
  useEffect(() => {
    if (!register || isEdit) return;
    let alive = true;
    regList().then(r => { if (alive) setNextPreview(fmtSeq(regNextFor(r, dateYear), dateYear)); }).catch(() => {});
    return () => { alive = false; };
  }, [register, isEdit, dateYear]);

  const totals = computeIAIncTotal(f);
  const ok = f.clientName && (register || f.number) && f.date && f.services.some(s => s.description && Number(s.amount) > 0);

  const previewRegPdf = async () => {
    const win = window.open("", "_blank");
    try {
      const { buildIncInvoicePdf } = await loadIncPdf();
      const bytes = await buildIncInvoicePdf({ ...regPayload(f), number: f.number }, IA_INC_INFO, { preview: true });
      showPdfBytes(bytes, "IA-Inc-Invoice-preview.pdf", win);
    } catch (e) { if (win) win.close(); setRegErr(e.message); }
  };
  const issueReg = async () => {
    if (!ok || busy) return;
    setBusy(true); setRegErr("");
    const win = window.open("", "_blank");
    try {
      const payload = regPayload(f);
      const { invoice } = isEdit ? await regUpdate(f.id, payload) : await regIssue(payload);
      let saved = invoice, pdfErr = null;
      try {
        const r = await archiveIncPdf(invoice);
        saved = r.invoice;
        showPdfBytes(r.bytes, incPdfName(saved), win);
      } catch (e) { pdfErr = e.message; if (win) win.close(); }
      onRegSaved(saved, pdfErr);
    } catch (e) { if (win) win.close(); setRegErr(e.message); setBusy(false); }
  };

  const generatePDF = () => {
    const html = buildIAIncInvoiceHTML(f);
    const w = window.open("", "_blank");
    if (!w) { alert("Impossibile aprire una nuova finestra. Verifica il blocco pop-up del browser."); return; }
    w.document.write(html); w.document.close(); w.focus();
    // Attendere il render prima della stampa (opzionale — l'utente clicca Stampa)
  };

  const buildRecord = () => ({
    id: f.id,
    client: f.clientName,
    clientAddress: f.clientAddress,
    number: f.number,
    docType: "Fattura IA Inc.",
    isCreditNote: false,
    date: f.date,
    dueDate: "",
    month: (f.date || "").slice(0, 7),
    amount: totals.total,
    currency: "USD",
    fteStatus: "",
    status: "issued",
    matchedId: null,
    source: "manual",
    iaInc: true,
    services: f.services.filter(s => s.description || Number(s.amount)),
    adminCostPercent: Number(f.adminCostPercent) || 0,
    taxPercent: Number(f.taxPercent) || 0,
    otherAmount: Number(f.otherAmount) || 0,
    terms: f.terms,
    notes: f.notes,
    vatExempt: true, // Fatture IA Inc. sono fuori dal perimetro IVA italiano
    vatRate: 0,
  });

  const saveAndPDF = () => { onSave(buildRecord()); generatePDF(); };
  const saveOnly = () => onSave(buildRecord());

  const F = Field;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 720, maxWidth: "100%", maxHeight: "92vh", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{register ? (isEdit ? "Edit invoice " + f.number : "New invoice") : (isEdit ? "Modifica fattura IA Inc." : "Nuova fattura IA Inc.")}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>{register
              ? (isEdit ? "Saving archives a new version of the PDF. The invoice number does not change." : "The number is assigned automatically when you issue the invoice, and the PDF is archived in the invoice register.")
              : "Fattura in USD emessa dalla consociata IA International Advisors Inc. (New York). Salvala nel database e genera il PDF pronto da inviare al cliente."}</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>

        <div style={{ overflowY: "auto", flex: 1, padding: 20 }}>
          <div style={{ background: "#FBFCFE", border: "1px solid " + C.line, padding: 12, marginBottom: 16, fontSize: 11.5, color: C.muted }}>
            <div style={{ color: C.navy, fontFamily: "Poppins", fontWeight: 600, fontSize: 12, marginBottom: 4 }}>Intestazione fissa</div>
            {IA_INC_INFO.companyName} · {IA_INC_INFO.address1} · {IA_INC_INFO.address2} · {IA_INC_INFO.email}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 180px", gap: 12 }}>
            <F label="Cliente *"><input className="ia-input" value={f.clientName} onChange={e => set("clientName", e.target.value)} placeholder="Es. Cledan SRL" /></F>
            {register
              ? <F label="Invoice number"><div className="ia-input" style={{ background: "#F3F5F9", color: isEdit ? C.navy : C.muted, fontFamily: "Poppins", fontWeight: isEdit ? 600 : 400, fontSize: 13 }} title="Assigned by the system: it cannot be typed or changed">{isEdit ? f.number : (nextPreview ? nextPreview + " · on issue" : "Assigned on issue")}</div></F>
              : <F label="Numero fattura *"><input className="ia-input" value={f.number} onChange={e => set("number", e.target.value)} placeholder="045/2026" /></F>}
          </div>
          <F label="Indirizzo cliente (una riga per capo)"><textarea className="ia-input" rows={3} value={f.clientAddress} onChange={e => set("clientAddress", e.target.value)} placeholder={"via Leuciana Pontecorvo\n03037 Italy"} style={{ resize: "vertical", fontFamily: "inherit" }} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "180px 1fr", gap: 12 }}>
            <F label="Data emissione *"><input className="ia-input" type="date" value={f.date} onChange={e => set("date", e.target.value)} /></F>
            <F label="Termini di pagamento"><input className="ia-input" value={f.terms} onChange={e => set("terms", e.target.value)} placeholder="Due upon receipt" /></F>
          </div>

          {/* Righe di servizio */}
          <div style={{ marginTop: 6, marginBottom: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 6 }}>
              <label style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins", fontWeight: 600 }}>Voci fattura *</label>
              <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={addSvc}><Plus size={12} /> Aggiungi riga</button>
            </div>
            {f.services.map((s, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 130px 30px", gap: 8, marginBottom: 6, alignItems: "center" }}>
                <input className="ia-input" value={s.description} onChange={e => updSvc(i, "description", e.target.value)} placeholder={i === 0 ? "Es. Contract Jan 28 th, 2025 IA Inc retainer fee" : "Descrizione servizio"} style={{ marginBottom: 0 }} />
                <input className="ia-input" type="number" step="0.01" value={s.amount} onChange={e => updSvc(i, "amount", e.target.value)} placeholder="0.00" style={{ marginBottom: 0, textAlign: "right" }} />
                <button className="ia-btn ia-btn-ghost" style={{ padding: 6 }} onClick={() => rmSvc(i)} disabled={f.services.length === 1} title="Rimuovi riga"><Trash2 size={13} /></button>
              </div>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            <F label="% administrative cost"><input className="ia-input" type="number" step="0.01" value={f.adminCostPercent} onChange={e => set("adminCostPercent", e.target.value)} /></F>
            <F label="TAX %"><input className="ia-input" type="number" step="0.01" value={f.taxPercent} onChange={e => set("taxPercent", e.target.value)} /></F>
            <F label="OTHER (USD)"><input className="ia-input" type="number" step="0.01" value={f.otherAmount} onChange={e => set("otherAmount", e.target.value)} /></F>
          </div>

          {/* Anteprima totali */}
          <div style={{ background: "#FBFCFE", border: "1px solid " + C.line, padding: 14, marginTop: 8, fontSize: 13 }}>
            <div style={{ color: C.muted, fontSize: 11, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".4px", marginBottom: 8 }}>Anteprima totali</div>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "2px 0", color: C.muted }}><span>Servizi</span><span style={{ fontFamily: "Poppins" }}>{fmtUSD(totals.servicesTotal)}</span></div>
            {totals.adminCost > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "2px 0", color: C.muted }}><span>+ {f.adminCostPercent}% administrative cost</span><span style={{ fontFamily: "Poppins" }}>{fmtUSD(totals.adminCost)}</span></div>}
            <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0", borderTop: "1px solid " + C.line, marginTop: 4, fontFamily: "Poppins", fontWeight: 600, color: C.navy }}><span>SUBTOTAL</span><span>{fmtUSD(totals.subtotal)}</span></div>
            {totals.taxAmount > 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "2px 0", color: C.muted }}><span>TAX {f.taxPercent}%</span><span style={{ fontFamily: "Poppins" }}>{fmtUSD(totals.taxAmount)}</span></div>}
            {totals.otherAmount !== 0 && <div style={{ display: "flex", justifyContent: "space-between", padding: "2px 0", color: C.muted }}><span>OTHER</span><span style={{ fontFamily: "Poppins" }}>{fmtUSD(totals.otherAmount)}</span></div>}
            <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0 0", borderTop: "1px solid " + C.line, marginTop: 4, fontFamily: "Poppins", fontWeight: 700, color: C.orange, fontSize: 15 }}><span>TOTAL</span><span>{fmtUSD(totals.total)}</span></div>
          </div>

          <F label="Note interne (non stampate)" ><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>
        </div>

        {register && regErr && <div style={{ margin: "0 20px 0", padding: "9px 12px", background: "#FBECEC", border: "1px solid #E7C9C9", color: C.red, fontSize: 12.5 }}>{regErr}</div>}
        {register ? (
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line, flexWrap: "wrap" }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose} disabled={busy}>Cancel</button>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button className="ia-btn ia-btn-ghost" disabled={!ok || busy} onClick={previewRegPdf} title="See the PDF before issuing it (nothing is saved)"><Eye size={14} /> Preview PDF</button>
            <button className="ia-btn ia-btn-primary" disabled={!ok || busy} onClick={issueReg}>
              {busy ? <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> : <FileSignature size={14} />}
              {isEdit ? "Save and archive new version" : "Issue invoice"}
            </button>
          </div>
        </div>
        ) : (
        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line, flexWrap: "wrap" }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button className="ia-btn ia-btn-ghost" disabled={!ok} onClick={() => { if (ok) generatePDF(); }} title="Genera solo il PDF senza salvare la fattura nel database">Solo PDF</button>
            <button className="ia-btn ia-btn-ghost" disabled={!ok} onClick={saveOnly}>Salva senza PDF</button>
            <button className="ia-btn ia-btn-primary" disabled={!ok} onClick={saveAndPDF}><FileSignature size={14} /> Salva e genera PDF</button>
          </div>
        </div>
        )}
      </div>
    </div>
  );
}

// ============ REGISTRO FATTURE IA INC. ============
// Solo nella Suite Inc. (testi in inglese). Il registro sta sul server (R2): il numero
// progressivo lo assegna il server all'emissione, quindi non può ripetersi né essere digitato.
// Una fattura emessa non si cancella: si annulla e resta nel registro con il suo numero.
const loadIncPdf = () => import("./incInvoicePdf.js");
const fmtSeq = (seq, year) => String(seq).padStart(3, "0") + "/" + year;
const incPdfName = (inv) => "IA-Inc-Invoice-" + String(inv.number || "draft").replace(/[^A-Za-z0-9]+/g, "-") + ".pdf";
const latestPdf = (inv) => (inv?.pdfs || []).slice().sort((a, b) => (b.rev || 0) - (a.rev || 0))[0] || null;
function regNextFor(reg, year) {
  const used = (reg?.invoices || []).filter(i => i.year === year && i.seq).map(i => i.seq);
  return Math.max(Number(reg?.next?.[year]) || 1, used.length ? Math.max(...used) + 1 : 1);
}
// Dati del modulo → fattura per il server
const regPayload = (f) => ({
  id: f.id,
  client: f.clientName ?? f.client ?? "",
  clientAddress: f.clientAddress || "",
  date: f.date,
  services: (f.services || []).filter(s => String(s.description || "").trim() || Number(s.amount)).map(s => ({ description: s.description || "", amount: Number(s.amount) || 0 })),
  adminCostPercent: Number(f.adminCostPercent) || 0,
  taxPercent: Number(f.taxPercent) || 0,
  otherAmount: Number(f.otherAmount) || 0,
  terms: f.terms || "Due upon receipt",
  notes: f.notes || "",
});
// Fattura del registro → riga dell'elenco "fatture emesse" (riconciliazione, cruscotto, conto economico)
function issuedFromReg(inv, prev = {}) {
  return {
    ...prev,
    id: inv.id, client: inv.client, clientAddress: inv.clientAddress, number: inv.number,
    docType: "Fattura IA Inc.", isCreditNote: false, date: inv.date, dueDate: prev.dueDate || "",
    month: String(inv.date || "").slice(0, 7), amount: inv.total, currency: "USD", fteStatus: prev.fteStatus || "",
    status: prev.status === "paid" ? "paid" : "issued", matchedId: prev.matchedId ?? null,
    source: "register", iaInc: true, reg: true,
    services: inv.services, adminCostPercent: inv.adminCostPercent, taxPercent: inv.taxPercent, otherAmount: inv.otherAmount,
    terms: inv.terms, notes: inv.notes, vatExempt: true, vatRate: 0,
    pdfKey: latestPdf(inv)?.key || "",
  };
}
function syncIssuedWithReg(issued, inv) {
  const prev = issued.find(x => x.id === inv.id);
  if (inv.status === "void") return prev ? issued.filter(x => x.id !== inv.id) : issued; // annullata: fuori dai ricavi
  const rec = issuedFromReg(inv, prev);
  return prev ? issued.map(x => x.id === inv.id ? rec : x) : [rec, ...issued];
}
async function archiveIncPdf(inv) {
  const { buildIncInvoicePdf } = await loadIncPdf();
  const bytes = await buildIncInvoicePdf(inv, IA_INC_INFO);
  const r = await regUploadPdf(inv.id, inv.rev || 1, bytes);
  return { invoice: r.invoice, bytes };
}
function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
function showPdfBytes(bytes, name, win) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  if (win && !win.closed) { const url = URL.createObjectURL(blob); win.location.href = url; setTimeout(() => URL.revokeObjectURL(url), 120000); }
  else saveBlob(blob, name);
}
// x = fattura del registro oppure riga dell'elenco (pdfKey)
function openRegPdf(x, download) {
  const key = latestPdf(x)?.key || x.pdfKey;
  if (!key) return false;
  window.open(regPdfUrl(key, incPdfName(x), download), "_blank");
  return true;
}

const REG_STATUS = {
  issued: { bg: "#EBF2FB", c: "#1D4E9E", t: "Issued" },
  paid: { bg: "#EAF6EE", c: "#1F8A5B", t: "Paid" },
  void: { bg: "#FBECEC", c: "#C0392B", t: "Void" },
};

function IncRegister({ ctx }) {
  const { issued, setIssued, notify } = ctx;
  const issuedRef = useRef(issued); issuedRef.current = issued;
  const [reg, setReg] = useState(null);
  const [err, setErr] = useState("");
  const [year, setYear] = useState(new Date().getFullYear());
  const [q, setQ] = useState("");
  const [modal, setModal] = useState(null);
  const [openRow, setOpenRow] = useState(null);
  const [busy, setBusy] = useState("");
  const [nextEdit, setNextEdit] = useState(null);

  const load = useCallback(() => regList().then(r => { setReg(r); setErr(""); }).catch(e => setErr(e.message)), []);
  useEffect(() => { load(); }, [load]);

  // L'elenco "fatture emesse" segue il registro (anche per fatture emesse da un collega su un altro dispositivo)
  useEffect(() => {
    if (!reg) return;
    let iss = issuedRef.current, changed = false;
    for (const inv of reg.invoices) {
      const prev = iss.find(x => x.id === inv.id);
      if (inv.status === "void") { if (prev) { iss = iss.filter(x => x.id !== inv.id); changed = true; } continue; }
      const want = issuedFromReg(inv, prev);
      if (!prev || !prev.reg || prev.number !== want.number || prev.amount !== want.amount || prev.date !== want.date || prev.client !== want.client || prev.pdfKey !== want.pdfKey) {
        iss = syncIssuedWithReg(iss, inv); changed = true;
      }
    }
    if (changed) setIssued(iss);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reg]);

  const applyInv = (inv) => {
    setReg(r => {
      const invoices = r.invoices.some(i => i.id === inv.id) ? r.invoices.map(i => i.id === inv.id ? inv : i) : [...r.invoices, inv];
      const next = { ...r.next };
      if (inv.seq) next[inv.year] = Math.max(Number(next[inv.year]) || 1, inv.seq + 1);
      return { ...r, invoices, next };
    });
  };

  if (err && !reg) return (
    <div className="ia-panel" style={{ padding: 22 }}>
      <div style={{ color: C.red, fontFamily: "Poppins", fontWeight: 600, marginBottom: 6 }}>The invoice register could not be loaded</div>
      <div style={{ fontSize: 13, color: C.muted, marginBottom: 12 }}>{err}</div>
      <button className="ia-btn ia-btn-ghost" onClick={load}><RefreshCw size={14} /> Try again</button>
    </div>
  );
  if (!reg) return <div style={{ padding: 40, textAlign: "center", color: C.muted }}><Loader2 size={22} style={{ animation: "spin 1s linear infinite" }} /><style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>;

  const thisYear = new Date().getFullYear();
  const years = [...new Set([thisYear, ...reg.invoices.map(i => i.year).filter(Boolean), ...Object.keys(reg.next || {}).map(Number)])].sort((a, b) => b - a);
  const statusOf = (inv) => inv.status === "void" ? "void" : (issued.find(x => x.id === inv.id)?.status === "paid" ? "paid" : "issued");
  const ql = q.trim().toLowerCase();
  const yearRows = reg.invoices.filter(i => i.year === year)
    .sort((a, b) => (b.seq || 0) - (a.seq || 0) || String(b.date).localeCompare(String(a.date)));
  const rows = ql ? yearRows.filter(i => (i.client || "").toLowerCase().includes(ql) || (i.number || "").toLowerCase().includes(ql)) : yearRows;
  const live = yearRows.filter(i => i.status !== "void");
  const totalYear = live.reduce((s, i) => s + (Number(i.total) || 0), 0);
  const unpaid = live.filter(i => statusOf(i) !== "paid").reduce((s, i) => s + (Number(i.total) || 0), 0);
  const nextNo = regNextFor(reg, year);
  const regIds = new Set(reg.invoices.map(i => i.id));
  const legacy = issued.filter(x => x.iaInc && !regIds.has(x.id));

  const onSaved = (inv, pdfErr) => {
    applyInv(inv);
    setIssued(syncIssuedWithReg(issuedRef.current, inv));
    setModal(null); setYear(inv.year || year);
    notify(pdfErr ? "Invoice " + inv.number + " saved, but the PDF was not archived: " + pdfErr : "Invoice " + inv.number + " issued and archived", pdfErr ? "err" : "ok");
  };
  const voidInv = async (inv) => {
    const reason = prompt("Void invoice " + inv.number + " (" + inv.client + ")?\nIt stays in the register with its number, marked VOID, and is removed from revenue.\n\nReason:");
    if (!reason || !reason.trim()) return;
    setBusy("void-" + inv.id);
    try { const { invoice } = await regVoid(inv.id, reason.trim()); applyInv(invoice); setIssued(syncIssuedWithReg(issuedRef.current, invoice)); notify("Invoice " + inv.number + " voided"); }
    catch (e) { notify(e.message, "err"); } finally { setBusy(""); }
  };
  const genPdf = async (inv) => {
    const win = window.open("", "_blank");
    setBusy("pdf-" + inv.id);
    try { const r = await archiveIncPdf(inv); applyInv(r.invoice); setIssued(syncIssuedWithReg(issuedRef.current, r.invoice)); showPdfBytes(r.bytes, incPdfName(r.invoice), win); notify("PDF archived for invoice " + inv.number); }
    catch (e) { if (win) win.close(); notify(e.message, "err"); } finally { setBusy(""); }
  };
  const importLegacy = async () => {
    setBusy("import");
    try {
      const list = legacy.map(x => ({
        id: x.id, number: x.number, client: x.client, clientAddress: x.clientAddress || "", date: x.date,
        services: x.services && x.services.length ? x.services : [{ description: x.notes || "Professional services", amount: Number(x.amount) || 0 }],
        adminCostPercent: x.services && x.services.length ? (x.adminCostPercent || 0) : 0,
        taxPercent: x.services && x.services.length ? (x.taxPercent || 0) : 0,
        otherAmount: x.services && x.services.length ? (x.otherAmount || 0) : 0,
        terms: x.terms, notes: x.notes,
      }));
      const { added, skipped } = await regImport(list);
      let iss = issuedRef.current; for (const inv of added) iss = syncIssuedWithReg(iss, inv); setIssued(iss);
      await load();
      notify(added.length + " invoice(s) added to the register" + (skipped.length ? " · " + skipped.length + " skipped: " + skipped.map(s => (s.number || "?") + " (" + s.why + ")").join(", ") : ""), skipped.length ? "err" : "ok");
    } catch (e) { notify(e.message, "err"); } finally { setBusy(""); }
  };
  const saveNext = async () => {
    const n = parseInt(nextEdit, 10);
    if (!n || n < 1) { notify("Enter a valid number", "err"); return; }
    setBusy("next");
    try { await regSetNext(year, n); await load(); setNextEdit(null); notify("Next invoice number for " + year + ": " + fmtSeq(n, year)); }
    catch (e) { notify(e.message, "err"); } finally { setBusy(""); }
  };
  const exportExcel = () => {
    const data = yearRows.slice().reverse().map(i => ({
      "Invoice no.": i.number, "Date": fmtDate(i.date), "Client": i.client, "Total (USD)": Number(i.total) || 0,
      "Status": REG_STATUS[statusOf(i)].t, "Version": i.rev || 1, "Issued by": i.issuedBy || "", "Issued on": String(i.issuedAt || "").slice(0, 10),
      "Void reason": i.voidReason || "", "Archived PDF": latestPdf(i) ? "yes" : "no",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 12 }, { wch: 11 }, { wch: 34 }, { wch: 13 }, { wch: 9 }, { wch: 8 }, { wch: 34 }, { wch: 11 }, { wch: 30 }, { wch: 12 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Register " + year);
    XLSX.writeFile(wb, "IA-Inc-Invoice-register-" + year + ".xlsx");
  };
  const zipYear = async () => {
    const withPdf = yearRows.filter(i => latestPdf(i));
    if (!withPdf.length) { notify("No archived PDFs for " + year, "err"); return; }
    setBusy("zip");
    try {
      const { default: JSZip } = await import("jszip");
      const zip = new JSZip();
      for (const inv of withPdf) {
        const r = await fetch(regPdfUrl(latestPdf(inv).key, incPdfName(inv)), { credentials: "same-origin" });
        if (!r.ok) throw new Error("Download failed for invoice " + inv.number);
        zip.file((inv.status === "void" ? "VOID_" : "") + incPdfName(inv), await r.arrayBuffer());
      }
      saveBlob(await zip.generateAsync({ type: "blob" }), "IA-Inc-Invoices-" + year + ".zip");
      const missing = yearRows.filter(i => !latestPdf(i) && i.status !== "void").length;
      notify(withPdf.length + " PDF(s) downloaded" + (missing ? " · " + missing + " invoice(s) without an archived PDF" : ""));
    } catch (e) { notify(e.message, "err"); } finally { setBusy(""); }
  };

  const kpi = (label, value, sub, color = C.navy) => (
    <div className="ia-panel" style={{ padding: 16 }}>
      <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".6px" }}>{label}</div>
      <div className="ia-kpi-num" style={{ color, margin: "4px 0", fontSize: 22 }}>{value}</div>
      {sub && <div style={{ fontSize: 11.5, color: C.muted }}>{sub}</div>}
    </div>
  );
  const small = { padding: "5px 8px" };

  return (
    <div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div className="ia-kpis" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 14, marginBottom: 16 }}>
        {kpi("Invoices issued " + year, live.length, (yearRows.length - live.length) ? (yearRows.length - live.length) + " void" : "none void")}
        {kpi("Total invoiced " + year, fmtMoney(totalYear, "USD"), "void invoices excluded")}
        {kpi("Not yet paid", fmtMoney(unpaid, "USD"), "from bank reconciliation", unpaid > 0 ? C.amber : C.green)}
        {kpi("Next number", fmtSeq(nextNo, year), "assigned automatically on issue", C.blue)}
      </div>

      {legacy.length > 0 && (
        <div className="ia-panel" style={{ padding: "12px 16px", marginBottom: 14, background: "#FFF8E6", borderColor: "#F5DFA6", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div style={{ fontSize: 13, color: "#6B5A21" }}><b>{legacy.length} invoice(s)</b> created before the register are not in it yet ({legacy.slice(0, 4).map(x => x.number || "no number").join(", ")}{legacy.length > 4 ? "…" : ""}). Add them so the numbering continues after them.</div>
          <button className="ia-btn ia-btn-dark" disabled={busy === "import"} onClick={importLegacy}>{busy === "import" ? <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> : <History size={14} />} Add to register</button>
        </div>
      )}

      <div className="ia-panel">
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "14px 18px", borderBottom: "1px solid " + C.line, flexWrap: "wrap" }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 14.5, color: C.navy }}>Invoices {year}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Sequential numbering by year · every issued invoice is archived as a PDF</div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <select className="ia-input" style={{ width: "auto", fontFamily: "Poppins", fontWeight: 600, color: C.navy }} value={year} onChange={e => { setYear(Number(e.target.value)); setNextEdit(null); }}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <div style={{ position: "relative" }}>
              <Search size={14} color={C.muted} style={{ position: "absolute", left: 9, top: 11 }} />
              <input className="ia-input" style={{ width: 190, paddingLeft: 28 }} placeholder="Client or number" value={q} onChange={e => setQ(e.target.value)} />
            </div>
            <button className="ia-btn ia-btn-ghost" onClick={() => setNextEdit(nextEdit == null ? String(nextNo) : null)} title="Set the next number (only to continue a sequence started outside the app)"><Settings size={14} /> Numbering</button>
            <button className="ia-btn ia-btn-ghost" onClick={exportExcel} disabled={!yearRows.length}><Download size={14} /> Excel</button>
            <button className="ia-btn ia-btn-ghost" onClick={zipYear} disabled={!yearRows.length || busy === "zip"}>{busy === "zip" ? <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> : <FileArchive size={14} />} PDFs (.zip)</button>
            <button className="ia-btn ia-btn-primary" onClick={() => setModal({})}><Plus size={15} /> New invoice</button>
          </div>
        </div>

        {nextEdit != null && (
          <div style={{ padding: "12px 18px", background: "#FAFBFD", borderBottom: "1px solid " + C.line, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", fontSize: 13 }}>
            <span style={{ fontFamily: "Poppins", fontWeight: 500, color: C.navy }}>Next invoice number for {year}:</span>
            <input className="ia-input" type="number" min="1" style={{ width: 100 }} value={nextEdit} onChange={e => setNextEdit(e.target.value)} />
            <span style={{ color: C.muted }}>→ {fmtSeq(parseInt(nextEdit, 10) || 1, year)}</span>
            <button className="ia-btn ia-btn-dark" disabled={busy === "next"} onClick={saveNext}>Save</button>
            <button className="ia-btn ia-btn-ghost" onClick={() => setNextEdit(null)}>Cancel</button>
            <div style={{ flexBasis: "100%", fontSize: 11.5, color: C.muted }}>Use this only to continue a sequence started outside the app (e.g. the last invoice issued elsewhere was 044 → set 45). It can never go back below a number already issued.</div>
          </div>
        )}

        {rows.length === 0 ? (
          <div style={{ padding: 28, textAlign: "center", fontSize: 13, color: C.muted }}>
            {ql ? "No invoices match your search." : "No invoices in the register for " + year + " yet. The first one issued will be " + fmtSeq(nextNo, year) + "."}
          </div>
        ) : (
          <div style={{ overflowX: "auto" }}><table className="ia-table">
            <thead><tr><th>No.</th><th>Date</th><th>Client</th><th style={{ textAlign: "right" }}>Total</th><th>Status</th><th>PDF</th><th></th></tr></thead>
            <tbody>{rows.map(inv => {
              const st = statusOf(inv), S = REG_STATUS[st], isVoid = st === "void", pdf = latestPdf(inv), expanded = openRow === inv.id;
              return (
                <React.Fragment key={inv.id}>
                  <tr style={isVoid ? { color: C.muted } : undefined}>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600, color: isVoid ? C.muted : C.navy, whiteSpace: "nowrap" }}><span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                      <button onClick={() => setOpenRow(expanded ? null : inv.id)} title="Show history" style={{ border: "none", background: "none", padding: 0, cursor: "pointer", color: "inherit", font: "inherit", display: "inline-flex", alignItems: "center", gap: 4 }}>
                        <ChevronRight size={13} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform .15s" }} />{inv.number}
                      </button>
                      {inv.legacy && <span className="ia-chip" style={{ background: "#EEF1F6", color: C.muted, fontSize: 10 }}>imported</span>}
                    </span></td>
                    <td style={{ whiteSpace: "nowrap" }}>{fmtDate(inv.date)}</td>
                    <td style={{ fontFamily: "Poppins", fontWeight: 500 }}>{inv.client}</td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, whiteSpace: "nowrap" }}><span style={{ textDecoration: isVoid ? "line-through" : "none" }}>{fmtMoney(inv.total, "USD")}</span></td>
                    <td><span className="ia-chip" style={{ background: S.bg, color: S.c }}>{S.t}</span></td>
                    <td style={{ whiteSpace: "nowrap" }}><span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                      {pdf
                        ? <><button className="ia-btn ia-btn-ghost" style={small} onClick={() => openRegPdf(inv)} title="Open the archived PDF"><FileSignature size={13} /> PDF</button>
                            <button className="ia-btn ia-btn-ghost" style={small} onClick={() => openRegPdf(inv, true)} title="Download"><FileDown size={13} /></button>
                            {(inv.rev || 1) > 1 && <span style={{ fontSize: 11, color: C.muted, marginLeft: 4 }}>v{inv.rev}</span>}</>
                        : (isVoid ? <span style={{ fontSize: 12, color: C.muted }}>—</span>
                          : <button className="ia-btn ia-btn-ghost" style={small} disabled={busy === "pdf-" + inv.id} onClick={() => genPdf(inv)} title="Create and archive the PDF for this invoice">{busy === "pdf-" + inv.id ? <Loader2 size={13} style={{ animation: "spin 1s linear infinite" }} /> : <Plus size={13} />} Create PDF</button>)}
                    </span></td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                      {!isVoid && <>
                        <button className="ia-btn ia-btn-ghost" style={small} onClick={() => setModal(issuedFromReg(inv, issued.find(x => x.id === inv.id)))} title="Edit and archive a new version (the number does not change)"><Pencil size={13} /></button>{" "}
                        <button className="ia-btn ia-btn-danger" style={small} disabled={busy === "void-" + inv.id} onClick={() => voidInv(inv)} title="Void: it stays in the register with its number"><Ban size={13} /></button>
                      </>}
                    </td>
                  </tr>
                  {expanded && (
                    <tr><td className="ia-detail" colSpan={7} style={{ background: "#FAFBFD", fontSize: 12.5, color: C.ink }}>
                      <div style={{ display: "grid", gap: 4 }}>
                        {(inv.history || []).map((h, k) => (
                          <div key={k}>
                            <span style={{ color: C.muted }}>{fmtDate(String(h.at).slice(0, 10))} {String(h.at).slice(11, 16)} UTC · {h.by}</span> — {h.action === "issued" ? "issued" : h.action === "imported" ? "added to the register (created before it)" : h.action === "updated" ? "edited, version " + h.rev : h.action === "void" ? "voided: " + h.reason : h.action}
                          </div>
                        ))}
                        {(inv.pdfs || []).length > 0 && (
                          <div style={{ marginTop: 4 }}>
                            <span style={{ color: C.muted }}>Archived PDFs: </span>
                            {(inv.pdfs || []).slice().sort((a, b) => a.rev - b.rev).map(p => (
                              <a key={p.key} href={regPdfUrl(p.key, incPdfName(inv).replace(".pdf", "_v" + p.rev + ".pdf"))} target="_blank" rel="noopener" style={{ marginRight: 10, color: C.blue }}>version {p.rev}</a>
                            ))}
                          </div>
                        )}
                        {inv.clientAddress && <div style={{ color: C.muted }}>Bill to: {inv.clientAddress.replace(/\n/g, ", ")}</div>}
                        {inv.notes && <div style={{ color: C.muted }}>Internal note: {inv.notes}</div>}
                      </div>
                    </td></tr>
                  )}
                </React.Fragment>
              );
            })}</tbody>
          </table></div>
        )}
      </div>

      <div style={{ fontSize: 12, color: C.muted, marginTop: 12, lineHeight: 1.55 }}>
        The number is assigned by the system when an invoice is issued, so two people can never get the same number and it cannot be typed by hand.
        An issued invoice is never deleted: if it is wrong, void it (it stays here with its number) and issue a new one. Editing an invoice keeps its number and archives a new version of the PDF; earlier versions stay available in the history.
      </div>

      {modal && <IAIncInvoiceModal x={modal} existingIssued={issuedRef.current} register onRegSaved={onSaved} onSave={() => {}} onClose={() => setModal(null)} />}
    </div>
  );
}

// ============ MONTHLY REPORT ============
function MonthlyReport({ month, ctx, onClose }) {
  const { contracts, active, expenses, issued, notify } = ctx;
  const monthIssued = issued.filter(x => (x.month || (x.date || "").slice(0, 7)) === month);
  const planned = compileActive(contracts, active, expenses, month, issued).filter(r => !r.hasIssued && r.grossTotal > 0);

  // Per ogni fattura emessa, cerca il contratto corrispondente (matching flessibile)
  // per portarsi dietro anche le note del contratto quando la fattura in sé non ne ha.
  const findContractFor = (clientName) => contracts.find(c => namesMatch(c.client, clientName));

  const rows = [
    ...monthIssued.map(x => {
      const vRate = x.vatExempt ? 0 : (x.vatRate != null && x.vatRate !== "" ? Number(x.vatRate) : 22);
      const net = Number(x.amount) || 0; // preserva segno: NC hanno importo negativo
      const vat = net * (vRate / 100);
      const contract = findContractFor(x.client);
      const isCredit = !!x.isCreditNote || net < 0;
      const noteParts = [];
      if (x.notes) noteParts.push(x.notes);
      if (contract?.notes) noteParts.push("Contratto: " + contract.notes);
      return { type: isCredit ? "nota credito" : "emessa", client: x.client, number: x.number || "", date: x.date, net, vatRate: vRate, vat, gross: net + vat, notes: noteParts.join(" — "), isCredit };
    }),
    ...planned.map(r => {
      const contract = contracts.find(c => c.id === r.contractId);
      const detail = r.lines.map(l => `${l.label} ${fmtMoney(l.amount)}`).join(" · ") + (r.expTotal > 0 ? ` · spese ${fmtMoney(r.expTotal)}` : "");
      const noteParts = [detail];
      if (contract?.notes) noteParts.push("Contratto: " + contract.notes);
      const netFee = r.feeTotal;
      const vat = netFee * (r.vatRate / 100);
      return { type: "da emettere", client: r.client, number: "", date: r.issueDate, net: r.total, vatRate: r.vatRate, vat, gross: r.grossTotal, notes: noteParts.join(" — ") };
    }),
  ];
  rows.sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const totalNet = rows.reduce((s, r) => s + r.net, 0);
  const totalVat = rows.reduce((s, r) => s + r.vat, 0);
  const totalGross = rows.reduce((s, r) => s + r.gross, 0);

  const asText = () => {
    let t = `RIEPILOGO FATTURE — ${monthLabel(month).toUpperCase()}\n`;
    t += `IA International Advisors\n\n`;
    const emRows = rows.filter(r => r.type === "emessa");
    const ncRows = rows.filter(r => r.type === "nota credito");
    const plRows = rows.filter(r => r.type === "da emettere");
    if (emRows.length > 0) {
      t += `FATTURE EMESSE (${emRows.length})\n`;
      emRows.forEach((r, i) => {
        t += `${i + 1}. ${r.client}${r.number ? " (n. " + r.number + ")" : ""} — ${fmtDate(r.date)}\n`;
        t += `   Netto ${fmtMoney(r.net)} · ${r.vatRate > 0 ? "IVA " + r.vatRate + "% = " + fmtMoney(r.gross) : "esente IVA"}\n`;
        if (r.notes) t += `   Note: ${r.notes}\n`;
      });
      t += `\n`;
    }
    if (ncRows.length > 0) {
      t += `NOTE DI CREDITO (${ncRows.length}) — sottratte dai ricavi\n`;
      ncRows.forEach((r, i) => {
        t += `${i + 1}. ${r.client}${r.number ? " (n. " + r.number + ")" : ""} — ${fmtDate(r.date)}\n`;
        t += `   Storno netto ${fmtMoney(r.net)} · ${r.vatRate > 0 ? "IVA " + r.vatRate + "% = " + fmtMoney(r.gross) : "esente IVA"}\n`;
        if (r.notes) t += `   Note: ${r.notes}\n`;
      });
      t += `\n`;
    }
    if (plRows.length > 0) {
      t += `DA EMETTERE — piano dai contratti (${plRows.length})\n`;
      plRows.forEach((r, i) => {
        t += `${i + 1}. ${r.client} — ${fmtDate(r.date)}\n`;
        t += `   Netto ${fmtMoney(r.net)} · ${r.vatRate > 0 ? "IVA " + r.vatRate + "% = " + fmtMoney(r.gross) : "esente IVA"}\n`;
        if (r.notes) t += `   Note: ${r.notes}\n`;
      });
      t += `\n`;
    }
    t += `TOTALI\n`;
    t += `Imponibile: ${fmtMoney(totalNet)}\n`;
    t += `IVA:        ${fmtMoney(totalVat)}\n`;
    t += `Lordo:      ${fmtMoney(totalGross)}\n`;
    return t;
  };

  const copyText = async () => {
    try { await navigator.clipboard.writeText(asText()); notify("Testo copiato negli appunti"); }
    catch { notify("Copia non riuscita — seleziona il testo a mano", "err"); }
  };

  const printPdf = () => {
    const w = window.open("", "_blank", "width=900,height=1000");
    if (!w) { notify("Il browser ha bloccato l'apertura della finestra — abilita i popup", "err"); return; }
    const rowsHtml = rows.map(r => `
      <tr>
        <td><span class="tag ${r.isCredit ? "credit" : r.type === "emessa" ? "emessa" : "piano"}">${r.type}</span></td>
        <td><b>${escapeHtml(r.client)}</b></td>
        <td>${escapeHtml(r.number || "—")}</td>
        <td>${fmtDate(r.date)}</td>
        <td class="num">${fmtMoney(r.net)}</td>
        <td class="num">${r.vatRate === 0 ? "esente" : fmtMoney(r.vat) + " <span class='muted'>(" + r.vatRate + "%)</span>"}</td>
        <td class="num strong">${fmtMoney(r.gross)}</td>
        <td class="notes">${escapeHtml(r.notes)}</td>
      </tr>`).join("");
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>Riepilogo ${monthLabel(month)}</title>
      <style>
        @page { size: A4 landscape; margin: 14mm; }
        body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #1A1F35; margin: 0; padding: 20px; }
        .head { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 3px solid #131D4B; padding-bottom: 10px; margin-bottom: 18px; }
        .head h1 { margin: 0; font-size: 22px; color: #131D4B; letter-spacing: .3px; }
        .head .sub { color: #6B7392; font-size: 12px; }
        .head .brand { color: #0E2289; font-weight: bold; }
        .head .brand .orange { color: #F57547; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        thead th { background: #131D4B; color: #fff; text-align: left; padding: 8px 10px; font-weight: 600; }
        tbody td { padding: 8px 10px; border-bottom: 1px solid #E1E5EE; }
        .num { text-align: right; font-variant-numeric: tabular-nums; }
        .strong { font-weight: 700; color: #131D4B; }
        .muted { color: #6B7392; font-size: 10px; }
        .notes { color: #4A5378; font-size: 11px; max-width: 300px; word-wrap: break-word; white-space: normal; line-height: 1.4; }
        .tag { display: inline-block; padding: 2px 7px; font-size: 10px; text-transform: uppercase; letter-spacing: .3px; }
        .tag.emessa { background: #E7F4EE; color: #2F8259; }
        .tag.piano { background: #FFF1E9; color: #F57547; }
        .tag.credit { background: #FBECEC; color: #B23A3A; }
        tfoot td { padding: 10px; background: #FAFBFD; border-top: 2px solid #131D4B; font-weight: 700; }
        tfoot .lbl { text-align: right; }
        tfoot .grand { font-size: 15px; color: #131D4B; }
        .foot { margin-top: 22px; font-size: 10px; color: #8891B5; text-align: right; }
      </style>
    </head><body>
      <div class="head">
        <div>
          <h1>Riepilogo fatture — ${monthLabel(month)}</h1>
          <div class="sub">Fatture emesse + piano da contratti (una tantum, retainer, eventi)</div>
        </div>
        ${IS_INC ? `<img src="${IA_INC_LOGO_URL()}" alt="International Advisors Inc." style="height:46px;width:auto;display:block">` : `<div class="brand"><span class="orange">[</span> IA INTERNATIONAL ADVISORS</div>`}
      </div>
      ${rows.length === 0 ? '<div style="padding:40px;text-align:center;color:#8891B5;">Nessuna fattura per questo mese.</div>' : `<table>
        <thead><tr><th>Tipo</th><th>Cliente</th><th>Numero</th><th>Data</th><th class="num">Netto</th><th class="num">IVA</th><th class="num">Lordo</th><th>Note</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
        <tfoot>
          <tr><td colspan="4" class="lbl">TOTALI (${rows.length} voci)</td>
              <td class="num">${fmtMoney(totalNet)}</td>
              <td class="num">${fmtMoney(totalVat)}</td>
              <td class="num grand">${fmtMoney(totalGross)}</td>
              <td></td></tr>
        </tfoot>
      </table>`}
      <div class="foot">Generato il ${fmtDate(today())} da IA Suite</div>
      <script>window.onload = () => setTimeout(() => window.print(), 250);</script>
    </body></html>`;
    w.document.open(); w.document.write(html); w.document.close();
  };

  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 820, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Riepilogo fatture — {monthLabel(month)}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>Elenco combinato: fatture emesse + piano dai contratti (una tantum, retainer, eventi).</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: "16px 20px" }}>
          {rows.length === 0 ? (
            <div style={{ padding: 30, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessuna fattura per questo mese.</div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table className="ia-table" style={{ fontSize: 12.5 }}>
                <thead><tr><th>Tipo</th><th>Cliente</th><th>N.</th><th>Data</th><th style={{ textAlign: "right" }}>Netto</th><th style={{ textAlign: "right" }}>IVA</th><th style={{ textAlign: "right" }}>Lordo</th><th>Note</th></tr></thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i}>
                      <td><span className="ia-chip" style={{ background: r.isCredit ? "#FBECEC" : r.type === "emessa" ? "#E7F4EE" : "#FFF1E9", color: r.isCredit ? C.red : r.type === "emessa" ? C.green : C.orange, fontSize: 10 }}>{r.type}</span></td>
                      <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{r.client}</td>
                      <td>{r.number || "—"}</td>
                      <td style={{ whiteSpace: "nowrap" }}>{fmtDate(r.date)}</td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins" }}>{fmtMoney(r.net)}</td>
                      <td style={{ textAlign: "right", fontSize: 11.5, color: r.vatRate === 0 ? C.muted : C.ink }}>{r.vatRate === 0 ? "esente" : <>{fmtMoney(r.vat)} <span style={{ color: C.muted, fontSize: 10 }}>({r.vatRate}%)</span></>}</td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700, color: C.navy }}>{fmtMoney(r.gross)}</td>
                      <td style={{ fontSize: 11.5, color: C.muted, maxWidth: 220 }}>{r.notes}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: "2px solid " + C.navy, background: "#FAFBFD" }}>
                    <td colSpan={4} style={{ textAlign: "right", padding: "10px 12px", fontFamily: "Poppins", fontWeight: 700 }}>TOTALI ({rows.length} voci)</td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{fmtMoney(totalNet)}</td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{fmtMoney(totalVat)}</td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700, color: C.navy, fontSize: 15 }}>{fmtMoney(totalGross)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )}
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line, flexWrap: "wrap" }}>
          <div style={{ fontSize: 11.5, color: C.muted }}>Il testo è pronto per essere incollato in email o WhatsApp. Il PDF apre la finestra di stampa: scegli "Salva come PDF".</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="ia-btn ia-btn-ghost" onClick={onClose}>Chiudi</button>
            <button className="ia-btn ia-btn-ghost" disabled={rows.length === 0} onClick={copyText}>Copia testo</button>
            <button className="ia-btn ia-btn-primary" disabled={rows.length === 0} onClick={printPdf}>Stampa / PDF</button>
          </div>
        </div>
      </div>
    </div>
  );
}
function escapeHtml(s) { return String(s || "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch])); }

// ============ PASSIVE INVOICES ============
function PassiveInvoices({ ctx }) {
  const { passive, setPassive, month, notify } = ctx;
  const [edit, setEdit] = useState(null);
  const impRef = useRef();
  const list = passive.filter(p => (p.month || (p.date || "").slice(0, 7)) === month);
  const total = list.reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const upsert = (p) => {
    const exists = passive.some(x => x.id === p.id);
    setPassive(exists ? passive.map(x => x.id === p.id ? p : x) : [...passive, p]);
    setEdit(null); notify(exists ? "Fattura aggiornata" : "Fattura registrata");
  };
  const del = (id) => setPassive(passive.filter(p => p.id !== id));
  const importPassivaFile = async (file) => {
    try {
      const parsed = importPassiva(await readTabular(file));
      if (!parsed.length) { notify("Nessuna fattura riconosciuta nel file", "err"); return; }
      const { merged, added, dup } = mergeById(passive, parsed, passiveKey);
      setPassive(merged); notify(`${added} fatture ricevute importate${dup ? `, ${dup} già presenti` : ""}`);
    } catch (e) { notify(e.message || "Import non riuscito", "err"); }
  };
  const cleanupPassiveDuplicates = () => {
    const groups = new Map();
    passive.forEach(x => {
      const key = passiveKey(x);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(x);
    });
    const duplicates = [];
    groups.forEach(arr => { if (arr.length > 1) duplicates.push(...arr.slice(1)); });
    if (duplicates.length === 0) { notify("Nessun duplicato trovato"); return; }
    const preview = duplicates.slice(0, 5).map(x => `• ${x.supplier} — ${x.number || "s.n."} — ${fmtDate(x.date)} — ${fmtMoney(x.amount)}`).join("\n");
    if (!confirm(`Trovati ${duplicates.length} duplicati (stesso fornitore + numero + data + importo).\n\nPrimi da eliminare:\n${preview}${duplicates.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toRemove = new Set(duplicates.map(x => x.id));
    setPassive(passive.filter(x => !toRemove.has(x.id)));
    notify(`${duplicates.length} duplicati rimossi`);
  };
  // Corregge NC passive legacy (importate come positive prima del fix di detection)
  const fixLegacyPassiveCreditNotes = () => {
    const toFix = passive.filter(p => {
      const dt = String(p.docType || "").toLowerCase();
      const nm = String(p.number || "").toLowerCase();
      const isNCByDoc = dt.includes("credito") || dt.includes("credit note") || dt === "td04" || /\bnc\b/.test(dt);
      const isNCByNumber = /^nc[\s\-\/]/i.test(p.number || "") || nm.includes("nota credito") || nm.includes("nota di credito");
      const isNC = isNCByDoc || isNCByNumber || p.isCreditNote;
      return isNC && (Number(p.amount) || 0) > 0;
    });
    if (toFix.length === 0) { notify("Nessuna nota di credito passiva legacy da correggere"); return; }
    const preview = toFix.slice(0, 5).map(p => `• ${p.supplier} — ${p.number || "s.n."} — ${fmtDate(p.date)} — ${fmtMoney(p.amount)}`).join("\n");
    if (!confirm(`Trovate ${toFix.length} note di credito passive con importo positivo (importate prima del fix).\nLe converto a importo NEGATIVO e le marco come "nota di credito" (ridurranno i costi nel P&L).\n\nPrime da correggere:\n${preview}${toFix.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toFixIds = new Set(toFix.map(p => p.id));
    setPassive(passive.map(p => toFixIds.has(p.id) ? { ...p, amount: -Math.abs(Number(p.amount) || 0), isCreditNote: true } : p));
    notify(`${toFix.length} note di credito corrette`);
  };
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 13.5, color: C.muted }}>Totale da ricevere {monthLabel(month)}: <b style={{ color: C.navy, fontFamily: "Poppins" }}>{fmtMoney(total)}</b></div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="ia-btn ia-btn-ghost" onClick={cleanupPassiveDuplicates} title="Trova e rimuove fatture doppie (stesso fornitore + numero + data + importo)"><RefreshCw size={14} /> Ripulisci duplicati</button>
          <button className="ia-btn ia-btn-ghost" onClick={fixLegacyPassiveCreditNotes} title="Converte a importo negativo le note di credito importate prima del fix automatico"><RefreshCw size={14} /> Correggi NC legacy</button>
          <button className="ia-btn ia-btn-ghost" onClick={() => impRef.current.click()}><Upload size={15} /> Importa Excel</button>
          <input ref={impRef} type="file" accept=".xlsx,.xls" hidden onChange={e => { if (e.target.files[0]) importPassivaFile(e.target.files[0]); e.target.value = ""; }} />
          <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), supplier: "", number: "", date: month + "-01", dueDate: "", amount: "", currency: "EUR", month, status: "to_pay", category: "Altro", vatExempt: false, vatRate: 22, fileName: "(manuale)" })}><Plus size={16} /> Registra fattura</button>
        </div>
      </div>
      {list.length === 0 ? <Empty icon={ArrowDownToLine} title="Nessuna fattura fornitore per questo mese" hint="Importa l'elenco PASSIVA (Excel), carica i PDF dei fornitori da «Carica documenti», oppure registra a mano." /> : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th>Fornitore</th><th>Categoria</th><th>Numero</th><th>Data</th><th>Scadenza</th><th style={{ textAlign: "right" }}>Importo (lordo)</th><th>Stato</th><th></th></tr></thead>
            <tbody>{list.map(p => {
              const isCredit = !!p.isCreditNote || (Number(p.amount) || 0) < 0;
              const isForeign = p.currency && p.currency !== "EUR";
              return (
                <tr key={p.id} style={isCredit ? { background: "#FFF5F5" } : (isForeign ? { background: "#FFFCF3" } : undefined)}>
                  <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>
                    {p.supplier}
                    {isCredit && <span className="ia-chip" style={{ background: "#FBECEC", color: C.red, marginLeft: 6, fontSize: 10 }}>nota di credito</span>}
                    {isForeign && <span className="ia-chip" style={{ background: "#FFF8E6", color: C.amber, marginLeft: 6, fontSize: 10 }}>{p.currency}{p.supplierCountry ? ` · ${p.supplierCountry}` : ""}</span>}
                    <div style={{ fontWeight: 400, fontSize: 11.5, color: C.muted }}>{p.fileName}</div>
                  </td>
                  <td><span className="ia-chip" style={{ background: "#EEF1F6", color: C.muted }}>{p.category || "Altro"}</span>{p.vatExempt && <div style={{ fontSize: 10.5, color: C.amber, marginTop: 2 }}>senza IVA</div>}</td>
                  <td>{p.number || "—"}</td><td>{fmtDate(p.date)}</td><td>{fmtDate(p.dueDate)}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: isCredit ? C.red : C.ink }}>{fmtMoney(p.amount, p.currency)}</td>
                  <td><Chip status={p.status} /></td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 9px" }} onClick={() => setEdit(p)}><Pencil size={14} /></button>{" "}
                    <button className="ia-btn ia-btn-danger" style={{ padding: "6px 9px" }} onClick={() => del(p.id)}><Trash2 size={14} /></button>
                  </td>
                </tr>
              );
            })}</tbody>
          </table>
        </div>
      )}
      {edit && <PassiveModal p={edit} onSave={upsert} onClose={() => setEdit(null)} />}
    </div>
  );
}
function PassiveModal({ p, onSave, onClose }) {
  const [f, setF] = useState({ vatExempt: false, vatRate: 22, category: "Altro", ...p });
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const vRate = f.vatExempt ? 0 : (Number(f.vatRate) || 0);
  const gross = Number(f.amount) || 0;
  const hasNet = f.netAmount != null && f.netAmount !== "" && isFinite(Number(f.netAmount));
  const net = hasNet ? Number(f.netAmount) : (vRate > 0 ? gross / (1 + vRate / 100) : gross);
  const vatAmt = hasNet ? 0 : gross - net;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 500, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Fattura fornitore</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Fornitore *"><input className="ia-input" value={f.supplier} onChange={e => set("supplier", e.target.value)} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Numero"><input className="ia-input" value={f.number} onChange={e => set("number", e.target.value)} /></F>
            <F label="Importo da pagare *"><input className="ia-input" type="number" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Ritenuta d'acconto"><input className="ia-input" type="number" value={f.withholding ?? ""} onChange={e => set("withholding", e.target.value)} placeholder="solo per parcelle" /></F>
            <F label="Imponibile (netto IVA)"><input className="ia-input" type="number" value={f.netAmount ?? ""} onChange={e => set("netAmount", e.target.value)} placeholder="se vuoto è stimato" /></F>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            <F label="Data"><input className="ia-input" type="date" value={f.date} onChange={e => set("date", e.target.value)} /></F>
            <F label="Scadenza"><input className="ia-input" type="date" value={f.dueDate} onChange={e => set("dueDate", e.target.value)} /></F>
            <F label="Valuta"><select className="ia-input" value={f.currency} onChange={e => set("currency", e.target.value)}><option>EUR</option><option>USD</option></select></F>
          </div>
          <F label="Categoria di costo"><select className="ia-input" value={f.category} onChange={e => set("category", e.target.value)}>{COST_CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></F>
          <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 8, background: "#FBFCFE" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
              <input type="checkbox" checked={!!f.vatExempt} onChange={e => set("vatExempt", e.target.checked)} />
              Fornitore <b>senza IVA</b> (esente, reverse charge, split payment, ecc.)
            </label>
            {!f.vatExempt && (
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                <label style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins" }}>Aliquota IVA (%)</label>
                <input className="ia-input" type="number" style={{ width: 90 }} value={f.vatRate} onChange={e => set("vatRate", e.target.value)} />
              </div>
            )}
            {gross > 0 && (
              <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid " + C.line, fontSize: 12, display: "flex", justifyContent: "space-between", color: C.muted }}>
                <span>{hasNet ? "Imponibile (dal documento)" : "Netto imponibile stimato"}</span>
                <span><b style={{ fontFamily: "Poppins", color: C.navy }}>{fmtMoney(net)}</b>{vatAmt > 0 && <span style={{ color: C.muted, marginLeft: 6 }}>+ {fmtMoney(vatAmt)} IVA</span>}</span>
              </div>
            )}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.supplier || f.amount === ""} onClick={() => onSave({ ...f, amount: Number(f.amount), vatRate: f.vatExempt ? 0 : (Number(f.vatRate) || 0), withholding: f.withholding === "" || f.withholding == null ? null : Math.abs(Number(f.withholding)) || null, netAmount: f.netAmount === "" || f.netAmount == null ? null : Math.abs(Number(f.netAmount)) || null, month: (f.date || "").slice(0, 7) || f.month })}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ============ BANK & RECONCILIATION ============
function Bank({ ctx }) {
  const { bank, setBank, contracts, active, setActive, expenses, passive, setPassive, issued, setIssued, dues, setDues, notify } = ctx;
  const [busy, setBusy] = useState(false);
  const [mapNeeded, setMapNeeded] = useState(null);
  const [filter, setFilter] = useState("unmatched"); // unmatched | all | matched
  const fileRef = useRef();

  const handleFile = async (file) => {
    setBusy(true);
    try {
      if ((file.name || "").toLowerCase().endsWith(".pdf")) {
        const tx = await importStatementPdf(file);
        if (!tx.length) throw new Error(IS_INC ? "No transactions found in the statement" : "Nessun movimento trovato nell'estratto conto");
        const seen = new Set(bank.map(b => `${b.date}|${b.amount}|${b.description}`));
        const fresh = tx.filter(b => !seen.has(`${b.date}|${b.amount}|${b.description}`));
        setBank([...bank, ...fresh]);
        notify(IS_INC ? `${fresh.length} transactions imported${tx.length - fresh.length ? `, ${tx.length - fresh.length} already present` : ""}` : `${fresh.length} movimenti importati`);
        setBusy(false); return;
      }
      const rows = await readTabular(file);
      const { headerIdx, cols } = detectBank(rows);
      if (headerIdx < 0) { setMapNeeded({ rows }); setBusy(false); return; }
      const tx = rowsToTx(rows, cols, headerIdx);
      if (tx.length === 0) { setMapNeeded({ rows }); setBusy(false); return; }
      setBank([...bank, ...tx]); notify(`${tx.length} movimenti importati`);
    } catch (e) { notify(e.message || "Errore lettura file", "err"); }
    setBusy(false);
  };

  // reconciliation: build candidate invoices
  const openInvoices = useMemo(() => {
    const out = [];
    // Solo fatture REALMENTE emesse (import ATTIVA + una tantum registrate a mano):
    // le pianificate da contratto sono importi indicativi e NON devono entrare tra i candidati
    // di riconciliazione, altrimenti sporcano il matching (es. proiezione €44.450 contro bonifico
    // reale €41.297,45 → nessun match). Restano visibili in dashboard/forecast come previsione.
    issued.filter(x => x.status !== "paid").forEach(x => out.push({ kind: "issued", id: x.id, label: x.client, sub: "emessa · " + (x.number || monthLabel((x.date || "").slice(0, 7))), amount: Number(x.amount), currency: x.currency || "EUR", sign: 1, number: x.number, date: x.date }));
    passive.filter(p => p.status !== "ppaid").forEach(p => out.push({ kind: "passive", id: p.id, label: p.supplier, sub: "passiva · " + (p.number || ""), amount: Number(p.amount), currency: p.currency, sign: -1, number: p.number, date: p.date }));
    dues.filter(d => d.status !== "paid").forEach(d => out.push({ kind: "due", id: d.id, label: d.label, sub: (d.category || "Scadenza"), amount: Number(d.amount), currency: "EUR", sign: -1, number: null, date: d.dueDate || d.paidAt }));
    return out;
  }, [passive, issued, dues]);

  // Suggerimento con confidence: alta = importo + nome nella descrizione,
  // media = importo esatto, bassa = nome nella descrizione con importo vicino ±5%
  const suggestFor = (tx) => {
    const wantSign = tx.amount >= 0 ? 1 : -1;
    const desc = String(tx.description || "").toLowerCase();
    return openInvoices
      .filter(inv => inv.sign === wantSign)
      .map(inv => {
        const txAbs = Math.abs(tx.amount), invAbs = Math.abs(inv.amount);
        const amtDelta = Math.abs(txAbs - invAbs);
        const amountClose = amtDelta < 1.0;
        const amountLoose = invAbs > 0 && amtDelta / invAbs < 0.05;
        const nameWords = normName(inv.label).split(/\s+/).filter(w => w.length >= 4);
        const nameHit = nameWords.length > 0 && nameWords.some(w => desc.includes(w));
        let conf = 0;
        if (amountClose && nameHit) conf = 3;
        else if (amountClose) conf = 2;
        else if (amountLoose && nameHit) conf = 1;
        if (conf === 0) return null;
        return { ...inv, _conf: conf, _amtDelta: amtDelta, _nameHit: nameHit };
      })
      .filter(Boolean)
      .sort((a, b) => b._conf - a._conf || a._amtDelta - b._amtDelta);
  };
  const match = (tx, inv) => {
    setBank(bank.map(t => t.id === tx.id ? { ...t, matchedId: inv.id, matchedType: inv.kind, matchedLabel: inv.label } : t));
    if (inv.kind === "issued") setIssued(issued.map(x => x.id === inv.id ? { ...x, status: "paid", matchedId: tx.id } : x));
    else if (inv.kind === "due") setDues(dues.map(d => d.id === inv.id ? { ...d, status: "paid", matchedId: tx.id } : d));
    else if (inv.kind === "active") {
      const i = active.findIndex(a => a.contractId === inv.contractId && a.month === inv.month);
      if (i >= 0) setActive(active.map((a, idx) => idx === i ? { ...a, status: "paid", matchedId: tx.id } : a));
      else setActive([...active, { contractId: inv.contractId, month: inv.month, status: "paid", matchedId: tx.id }]);
    } else setPassive(passive.map(p => p.id === inv.id ? { ...p, status: "ppaid", matchedId: tx.id } : p));
    notify("Pagamento riconciliato");
  };
  const unmatch = (tx) => {
    setBank(bank.map(t => t.id === tx.id ? { ...t, matchedId: null, matchedType: null, matchedLabel: null } : t));
    if (tx.matchedType === "issued") setIssued(issued.map(x => x.matchedId === tx.id ? { ...x, status: "issued", matchedId: null } : x));
    if (tx.matchedType === "due") setDues(dues.map(d => d.matchedId === tx.id ? { ...d, status: "to_pay", matchedId: null } : d));
    if (tx.matchedType === "active") setActive(active.map(a => a.matchedId === tx.id ? { ...a, status: "issued", matchedId: null } : a));
    if (tx.matchedType === "passive") setPassive(passive.map(p => p.matchedId === tx.id ? { ...p, status: "to_pay", matchedId: null } : p));
  };

  // Auto-abbina tutti i movimenti con suggerimento ad alta confidenza (importo + nome)
  const autoMatchAll = () => {
    const used = new Set();
    let done = 0;
    let newBank = [...bank], newIssued = [...issued], newPassive = [...passive], newDues = [...dues], newActive = [...active];
    for (let i = 0; i < newBank.length; i++) {
      const tx = newBank[i];
      if (tx.matchedId) continue;
      const sugs = suggestFor(tx).filter(s => !used.has(s.kind + ":" + s.id) && s._conf >= 3);
      if (sugs.length === 0) continue;
      const inv = sugs[0];
      used.add(inv.kind + ":" + inv.id);
      newBank[i] = { ...tx, matchedId: inv.id, matchedType: inv.kind, matchedLabel: inv.label };
      if (inv.kind === "issued") newIssued = newIssued.map(x => x.id === inv.id ? { ...x, status: "paid", matchedId: tx.id } : x);
      else if (inv.kind === "due") newDues = newDues.map(d => d.id === inv.id ? { ...d, status: "paid", matchedId: tx.id } : d);
      else if (inv.kind === "active") {
        const j = newActive.findIndex(a => a.contractId === inv.contractId && a.month === inv.month);
        if (j >= 0) newActive = newActive.map((a, idx) => idx === j ? { ...a, status: "paid", matchedId: tx.id } : a);
        else newActive = [...newActive, { contractId: inv.contractId, month: inv.month, status: "paid", matchedId: tx.id }];
      } else newPassive = newPassive.map(p => p.id === inv.id ? { ...p, status: "ppaid", matchedId: tx.id } : p);
      done++;
    }
    if (done === 0) { notify("Nessun abbinamento ad alta confidenza (importo + nome nella descrizione)"); return; }
    setBank(newBank); setIssued(newIssued); setPassive(newPassive); setDues(newDues); setActive(newActive);
    notify(`${done} ${done === 1 ? "movimento abbinato" : "movimenti abbinati"} automaticamente`);
  };

  // Rimuove i movimenti bancari duplicati (stessa data + stesso importo + stessa descrizione normalizzata).
  // Utile quando lo stesso estratto conto è stato importato più volte per errore. Mantiene il primo di ogni gruppo.
  const cleanupBankDuplicates = () => {
    const bankKey = t => (t.date || "").slice(0, 10) + "|" + (Number(t.amount) || 0).toFixed(2) + "|" + normName(t.description || "");
    const groups = new Map();
    bank.forEach(t => {
      const key = bankKey(t);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(t);
    });
    const duplicates = [];
    groups.forEach(arr => {
      if (arr.length > 1) {
        // Preferisci tenere quello già gestito (abbinato o categorizzato), scarta gli altri
        const sorted = [...arr].sort((a, b) => {
          const aHandled = (a.matchedId || a.directCategory || a.ignoredForPL) ? 0 : 1;
          const bHandled = (b.matchedId || b.directCategory || b.ignoredForPL) ? 0 : 1;
          return aHandled - bHandled;
        });
        duplicates.push(...sorted.slice(1));
      }
    });
    if (duplicates.length === 0) { notify("Nessun duplicato bancario trovato"); return; }
    const preview = duplicates.slice(0, 5).map(t => `• ${fmtDate(t.date)} · ${fmtMoney(t.amount)} · ${String(t.description || "").slice(0, 60)}`).join("\n");
    if (!confirm(`Trovati ${duplicates.length} movimenti bancari duplicati (stessa data + stesso importo + stessa descrizione).\n\nPrimi da eliminare:\n${preview}${duplicates.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toRemove = new Set(duplicates.map(t => t.id));
    setBank(bank.filter(t => !toRemove.has(t.id)));
    notify(`${duplicates.length} duplicati bancari rimossi`);
  };

  const filtered = bank.filter(tx => filter === "matched" ? tx.matchedId : filter === "unmatched" ? !tx.matchedId : true);
  const stats = { total: bank.length, matched: bank.filter(t => t.matchedId).length, unmatched: bank.filter(t => !t.matchedId).length };
  const highConfCount = bank.filter(t => !t.matchedId && suggestFor(t).some(s => s._conf >= 3)).length;

  return (
    <div>
      <div className="ia-panel" style={{ padding: 18, marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div><div className="ia-h" style={{ fontSize: 14.5, color: C.navy }}>Estratto conto Banco BPM</div><div style={{ fontSize: 12.5, color: C.muted, marginTop: 3 }}>Importa il CSV (o Excel). Le colonne vengono riconosciute in automatico.</div></div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {bank.length > 0 && <button className="ia-btn ia-btn-ghost" onClick={cleanupBankDuplicates} title="Rimuove i movimenti bancari doppi (stessa data + stesso importo + stessa descrizione), utile se hai importato lo stesso CSV due volte."><RefreshCw size={14} /> Ripulisci duplicati</button>}
          {bank.length > 0 && <button className="ia-btn ia-btn-danger" onClick={() => { if (confirm("Svuotare i movimenti importati?")) setBank([]); }}>Svuota movimenti</button>}
          <button className="ia-btn ia-btn-primary" disabled={busy} onClick={() => fileRef.current.click()}>{busy ? <Loader2 size={15} className="spin" /> : <Upload size={15} />} Importa CSV</button>
          <input ref={fileRef} type="file" accept={IS_INC ? ".pdf,.csv,.xlsx,.xls" : ".csv,.xlsx,.xls"} hidden onChange={e => { if (e.target.files[0]) handleFile(e.target.files[0]); e.target.value = ""; }} />
        </div>
      </div>

      {bank.length === 0 ? <Empty icon={Landmark} title="Nessun movimento importato" hint="Carica l'estratto conto CSV di Banco BPM per iniziare la riconciliazione con le fatture attive e passive." /> : (
        <>
          <div className="ia-panel" style={{ padding: "12px 16px", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
            <div style={{ display: "flex", gap: 4 }}>
              {[
                { id: "unmatched", label: `Da riconciliare (${stats.unmatched})` },
                { id: "matched", label: `Abbinati (${stats.matched})` },
                { id: "all", label: `Tutti (${stats.total})` },
              ].map(f => (
                <button key={f.id} className={"ia-btn " + (filter === f.id ? "ia-btn-dark" : "ia-btn-ghost")} style={{ padding: "6px 12px", fontSize: 12.5 }} onClick={() => setFilter(f.id)}>{f.label}</button>
              ))}
            </div>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              {highConfCount > 0 && <span style={{ fontSize: 12, color: C.muted }}><b style={{ color: C.green }}>{highConfCount}</b> abbinamenti sicuri disponibili</span>}
              <button className="ia-btn ia-btn-primary" disabled={highConfCount === 0} onClick={autoMatchAll} title="Applica tutti gli abbinamenti ad alta confidenza (importo + nome nella descrizione)"><Check size={14} /> Auto-abbina tutti</button>
            </div>
          </div>
          <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
            <table className="ia-table">
              <thead><tr><th>Data</th><th>Descrizione</th><th style={{ textAlign: "right" }}>Importo</th><th>Riconciliazione</th></tr></thead>
              <tbody>
                {filtered.map(tx => {
                  const sug = tx.matchedId ? [] : suggestFor(tx);
                  return (
                    <tr key={tx.id}>
                      <td style={{ whiteSpace: "nowrap" }}>{tx.date}</td>
                      <td style={{ maxWidth: 340, fontSize: 12.5 }}>{tx.description}</td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: tx.amount >= 0 ? C.green : C.ink }}>{fmtMoney(tx.amount, tx.currency)}</td>
                      <td>
                        {tx.matchedId ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
                            <span className="ia-chip" style={{ background: "#E7F4EE", color: C.green }}><Link2 size={12} /> {tx.matchedLabel}</span>
                            <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 12 }} onClick={() => unmatch(tx)}><Unlink size={12} /> Annulla</button>
                          </span>
                        ) : sug.length ? (
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                            {sug.slice(0, 3).map(inv => {
                              const confLabel = inv._conf === 3 ? "Sicuro" : inv._conf === 2 ? "Importo OK" : "Ipotesi";
                              const confBg = inv._conf === 3 ? "#E7F4EE" : inv._conf === 2 ? "#EEF1F6" : "#FFF8E6";
                              const confColor = inv._conf === 3 ? C.green : inv._conf === 2 ? C.muted : C.amber;
                              const btnClass = inv._conf === 3 ? "ia-btn-dark" : "ia-btn-ghost";
                              return (
                                <button key={inv.kind + inv.id} className={"ia-btn " + btnClass} style={{ padding: "5px 10px", fontSize: 12 }} title={`Confidence: ${confLabel}${inv._nameHit ? " · nome nella descrizione" : ""}`} onClick={() => match(tx, inv)}>
                                  <span className="ia-chip" style={{ background: confBg, color: confColor, fontSize: 10, padding: "1px 5px" }}>{confLabel}</span>
                                  {inv.label} <span style={{ opacity: .7 }}>({inv.sub})</span>
                                </button>
                              );
                            })}
                          </div>
                        ) : <span style={{ fontSize: 12.5, color: C.muted }}>Nessun abbinamento automatico</span>}
                      </td>
                    </tr>
                  );
                })}
                {filtered.length === 0 && <tr><td colSpan={4} style={{ padding: 24, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessun movimento in questo filtro.</td></tr>}
              </tbody>
            </table>
          </div>
        </>
      )}
      <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      {mapNeeded && <ColumnMapper data={mapNeeded.rows} onCancel={() => setMapNeeded(null)} onConfirm={(cols, headerIdx) => {
        const tx = rowsToTx(mapNeeded.rows, cols, headerIdx); setBank([...bank, ...tx]); setMapNeeded(null);
        notify(tx.length ? `${tx.length} movimenti importati` : "Nessun movimento riconosciuto", tx.length ? "ok" : "err");
      }} />}
    </div>
  );
}
// manual column mapper when auto-detect fails
function ColumnMapper({ data, onCancel, onConfirm }) {
  const [headerIdx, setHeaderIdx] = useState(0);
  const headers = data[headerIdx] || [];
  const [cols, setCols] = useState({ date: 0, desc: 1, amount: 2, dare: -1, avere: -1 });
  const setCol = (k, v) => setCols({ ...cols, [k]: v });
  return (
    <div className="ia-modal-bg" onClick={onCancel}>
      <div className="ia-panel" style={{ width: 520, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Associa le colonne dell'estratto conto</h3>
          <p style={{ margin: "5px 0 0", fontSize: 12.5, color: C.muted }}>Il tracciato non è stato riconosciuto automaticamente. Indica quali colonne usare.</p>
        </div>
        <div style={{ padding: 20 }}>
          <div style={{ marginBottom: 11 }}>
            <label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 5 }}>Riga di intestazione</label>
            <select className="ia-input" value={headerIdx} onChange={e => setHeaderIdx(Number(e.target.value))}>
              {data.slice(0, 12).map((r, i) => <option key={i} value={i}>{`Riga ${i + 1}: ${r.slice(0, 4).join(" | ").slice(0, 40)}`}</option>)}
            </select>
          </div>
          <ColSelect label="Colonna data" value={cols.date} headers={headers} onChange={v => setCol("date", v)} />
          <ColSelect label="Colonna descrizione / causale" value={cols.desc} headers={headers} onChange={v => setCol("desc", v)} />
          <ColSelect label="Colonna importo (con segno)" value={cols.amount} headers={headers} onChange={v => setCol("amount", v)} />
          <div style={{ fontSize: 12, color: C.muted, margin: "4px 0 8px" }}>oppure, se dare/avere sono separati:</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <ColSelect label="Uscite / Dare" value={cols.dare} headers={headers} onChange={v => setCol("dare", v)} />
            <ColSelect label="Entrate / Avere" value={cols.avere} headers={headers} onChange={v => setCol("avere", v)} />
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onCancel}>Annulla</button>
          <button className="ia-btn ia-btn-primary" onClick={() => onConfirm(cols, headerIdx)}>Importa movimenti</button>
        </div>
      </div>
    </div>
  );
}

// ============ UPLOAD PORTAL ============
const DOC_TYPES = [
  { id: "contract", label: "Contratto", icon: FileSignature, kind: "contract", accept: ".pdf,.docx,.doc", desc: "Estrae cliente, retainer, compensi a evento, durata e termini." },
  { id: "expense", label: "Spese di viaggio", icon: CircleDollarSign, kind: "expense", accept: ".pdf,.docx,.doc", desc: "Estrae il totale e lo associa a un contratto e a un mese." },
  { id: "invoice", label: "Fattura fornitore (PDF)", icon: ArrowDownToLine, kind: "invoice", accept: ".pdf,.docx,.doc", desc: "Estrae fornitore, numero, importo e scadenza da un singolo PDF." },
  { id: "attiva", label: "Fatture emesse (Excel)", icon: ArrowUpFromLine, kind: "attiva", accept: ".xlsx,.xls", desc: "Importa l'elenco fatture attive (export ATTIVA)." },
  { id: "passiva", label: "Fatture ricevute (Excel)", icon: ArrowDownToLine, kind: "passiva", accept: ".xlsx,.xls", desc: "Importa l'elenco fatture passive (export PASSIVA)." },
  { id: "bank", label: "Estratto conto (CSV)", icon: Landmark, kind: "bank", accept: ".csv,.xlsx,.xls", desc: "Importa i movimenti Banco BPM per la riconciliazione." },
];
// Parcelle con ritenuta d'acconto: l'importo da pagare è il TOTALE NETTO (lordo − ritenuta),
// non il "Totale documento". La regola è applicata qui, indipendentemente da come risponde la lettura.
function normalizeInvoice(d) {
  const num = (v) => { if (v == null || v === "") return null; const n = typeof v === "number" ? v : Number(String(v).replace(/[^\d.,-]/g, "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".")); return isFinite(n) && n !== 0 ? Math.abs(n) : null; };
  const wh = num(d.withholding), netToPay = num(d.netToPay), gross = num(d.grossTotal);
  if (wh) {
    if (netToPay) d.amount = netToPay;
    else if (gross) d.amount = Math.round((gross - wh) * 100) / 100;
    d.withholding = wh;
  }
  return d;
}

const DOC_TYPES_ACTIVE = IS_INC
  ? [
      ...DOC_TYPES.filter(d => !["attiva", "passiva", "bank"].includes(d.id)),
      { id: "issued", label: "Fattura emessa (Word/Excel/PDF)", icon: ArrowUpFromLine, kind: "issued_invoice", accept: ".pdf,.docx,.doc,.xlsx,.xls", desc: "Estrae cliente, numero, data e importo da una fattura emessa." },
      { ...DOC_TYPES.find(d => d.id === "bank"), accept: ".pdf,.csv,.xlsx,.xls" },
    ]
  : DOC_TYPES;
function UploadPortal({ ctx, go }) {
  const [type, setType] = useState("contract");
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const [review, setReview] = useState(null); // {kind, data, fileName}
  const inputRef = useRef();
  const t = DOC_TYPES_ACTIVE.find(d => d.id === type);
  const { notify, contracts, issued, setIssued, passive, setPassive } = ctx;

  const process = async (file) => {
    if (!file) return;
    if (type === "bank") { go("bank"); notify("Carica l'estratto conto dalla sezione Banca", "ok"); return; }
    if (type === "expense" && contracts.length === 0) {
      notify("Prima crea un contratto: le spese vanno associate a un cliente", "err");
      return;
    }
    if (type === "attiva" || type === "passiva") {
      setBusy(true);
      try {
        const rows = await readTabular(file);
        if (type === "attiva") {
          const parsed = importAttiva(rows);
          if (!parsed.length) throw new Error("Nessuna fattura riconosciuta nel file");
          const { merged, added, dup } = mergeById(issued, parsed, issuedKey);
          setIssued(merged); notify(`${added} fatture emesse importate${dup ? `, ${dup} già presenti` : ""}`); go("active");
        } else {
          const parsed = importPassiva(rows);
          if (!parsed.length) throw new Error("Nessuna fattura riconosciuta nel file");
          const { merged, added, dup } = mergeById(passive, parsed, passiveKey);
          setPassive(merged); notify(`${added} fatture ricevute importate${dup ? `, ${dup} già presenti` : ""}`); go("passive");
        }
      } catch (e) { notify(e.message || "Import non riuscito", "err"); }
      setBusy(false); return;
    }
    if (type === "issued") {
      setBusy(true);
      try {
        const d = await extractFromFile(file, "issued_invoice");
        const amt = Math.abs(Number(d.amount));
        if (!d.client || isNaN(amt)) throw new Error("Could not read client and amount from the invoice");
        const date = d.date || today();
        const isCredit = !!d.isCreditNote;
        const ok = confirm(`Add this ${isCredit ? "credit note" : "invoice"}?\n\nClient: ${d.client}\nNumber: ${d.number || "—"}\nDate: ${fmtDate(date)}\nDue: ${d.dueDate ? fmtDate(d.dueDate) : "—"}\nAmount: ${fmtMoney(amt, d.currency || "USD")}`);
        if (ok) {
          const rec = { id: uid(), client: String(d.client).trim(), number: String(d.number || "").trim(), docType: isCredit ? "Credit note" : "Invoice", isCreditNote: isCredit, date, month: date.slice(0, 7), dueDate: d.dueDate || null, amount: isCredit ? -amt : amt, currency: d.currency || "USD", fteStatus: "", status: "issued", matchedId: null, source: "document" };
          const { merged, added, dup } = mergeById(issued, [rec], issuedKey);
          setIssued(merged); notify(added ? "Invoice added" : `Already present (${dup})`); go("active");
        }
      } catch (e) { notify(e.message || "Extraction failed", "err"); }
      setBusy(false); return;
    }
    setBusy(true);
    try {
      const raw = await extractFromFile(file, t.kind);
      if (!raw || typeof raw !== "object") throw new Error("Risposta di estrazione non valida");
      const data = t.kind === "invoice" ? normalizeInvoice(raw) : raw;
      setReview({ kind: t.kind, data, fileName: file.name });
    } catch (e) { notify(e.message || "Estrazione non riuscita", "err"); }
    setBusy(false);
  };

  return (
    <div>
      <div className="ia-kpis" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12, marginBottom: 20 }}>
        {DOC_TYPES_ACTIVE.map(d => (
          <button key={d.id} onClick={() => setType(d.id)} className={"ia-panel" + (type === d.id ? " ia-bracket" : "")}
            style={{ padding: 16, textAlign: "left", cursor: "pointer", border: type === d.id ? `1px solid ${C.orange}` : `1px solid ${C.line}`, background: type === d.id ? "#FFF9F6" : "#fff" }}>
            <d.icon size={20} color={type === d.id ? C.orange : C.navy} />
            <div className="ia-h" style={{ fontSize: 14, color: C.navy, margin: "9px 0 4px" }}>{d.label}</div>
            <div style={{ fontSize: 12, color: C.muted, lineHeight: 1.4 }}>{d.desc}</div>
          </button>
        ))}
      </div>

      <div className={"ia-drop" + (drag ? " drag" : "")}
        onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) process(e.dataTransfer.files[0]); }}>
        {busy ? (
          <div style={{ color: C.navy }}><Loader2 size={28} className="spin" /><div className="ia-h" style={{ marginTop: 12, fontSize: 15 }}>Elaborazione in corso…</div><div style={{ fontSize: 12.5, color: C.muted }}>{["attiva", "passiva"].includes(type) ? "Import dell'elenco Excel" : "Lettura del documento con Claude"}</div></div>
        ) : (
          <>
            <Upload size={30} color={C.navy} style={{ opacity: .7 }} />
            <div className="ia-h" style={{ fontSize: 15.5, color: C.navy, margin: "12px 0 5px" }}>Trascina qui il file «{t.label}»</div>
            <div style={{ fontSize: 13, color: C.muted, marginBottom: 16 }}>Formati: {t.accept.replace(/\./g, " ").toUpperCase()}</div>
            <button className="ia-btn ia-btn-primary" onClick={() => inputRef.current.click()}><FileUp size={15} /> Seleziona file</button>
            <input ref={inputRef} type="file" accept={t.accept} hidden onChange={e => { if (e.target.files[0]) process(e.target.files[0]); e.target.value = ""; }} />
          </>
        )}
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>

      {contracts.length === 0 && (type === "expense" || type === "invoice") && (
        <div className="ia-panel" style={{ padding: 14, marginTop: 14, fontSize: 13, color: C.amber, display: "flex", gap: 9, alignItems: "center", borderColor: "#F0E2C6", background: "#FFFBF2" }}>
          <AlertTriangle size={16} /> Nessun contratto ancora: le spese vanno associate a un contratto. Carica prima un contratto.
        </div>
      )}

      {review && <ReviewModal review={review} ctx={ctx} onClose={() => setReview(null)} go={go} />}
    </div>
  );
}
// review + confirm extracted data before saving
function ReviewModal({ review, ctx, onClose, go }) {
  const { contracts, setContracts, expenses, setExpenses, passive, setPassive, month, notify } = ctx;
  const [f, setF] = useState(() => {
    const d = review.data || {};
    if (review.kind === "contract") {
      let items;
      if (Array.isArray(d.billingItems) && d.billingItems.length)
        items = d.billingItems.map(it => {
          const type = it.type === "trigger" ? "trigger" : it.type === "one_shot" ? "one_shot" : "recurring";
          return {
            id: uid(), type,
            label: it.label || (type === "trigger" ? "Fee a evento" : type === "one_shot" ? "Fattura una tantum" : "Retainer mensile"),
            amount: it.amount ?? "",
            triggerLabel: type === "trigger" ? (it.triggerLabel || "Firma del contratto") : null,
            status: "pending", triggeredDate: null,
            date: type === "one_shot" ? (it.date || (d.startDate || today()).slice(0, 7) + "-01") : null,
          };
        });
      else if (d.retainer != null) items = [{ ...emptyItem("recurring"), amount: d.retainer }];
      else items = [emptyItem("recurring")];
      return { client: d.client || "", currency: d.currency || "EUR", startDate: d.startDate || "", endDate: d.endDate || "", paymentTermsDays: d.paymentTermsDays ?? 30, notes: d.notes || "", vatExempt: !!d.vatExempt, vatRate: d.vatRate ?? 22, billingItems: items };
    }
    if (review.kind === "expense") {
      const amt = Number(d.amount);
      return { contractId: contracts[0]?.id || "", month: (d.date || "").slice(0, 7) || month, description: d.description || "Spesa di viaggio", amount: isFinite(amt) ? amt : "", currency: d.currency || "EUR" };
    }
    return {
      supplier: d.supplier || "",
      supplierCountry: d.supplierCountry || null,
      number: d.number || "",
      date: d.date || month + "-01",
      dueDate: d.dueDate || "",
      amount: d.amount ?? "",
      currency: (d.currency || "EUR").toUpperCase(),
      isCreditNote: !!d.isCreditNote,
      vatExempt: !!d.vatExempt,
      vatRate: d.vatRate ?? 22,
      category: d.category || "Altro",
      notes: d.notes || "",
    };
  });
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const titleMap = { contract: "Dati contratto estratti", expense: "Spesa di viaggio estratta", invoice: "Fattura fornitore estratta" };

  const save = () => {
    try {
      if (review.kind === "contract") {
        setContracts([...contracts, { id: uid(), ...f, vatRate: f.vatExempt ? 0 : (Number(f.vatRate) || 22), retainerDefault: null, billingItems: f.billingItems.map(it => ({ ...it, amount: it.amount === "" ? 0 : Number(it.amount) })) }]);
        notify("Contratto salvato"); go("contracts");
      } else if (review.kind === "expense") {
        if (!f.contractId) { notify("Prima associa la spesa a un contratto", "err"); return; }
        const amt = Number(f.amount);
        if (!isFinite(amt) || amt <= 0) { notify("Importo non valido", "err"); return; }
        setExpenses([...expenses, { id: uid(), ...f, amount: amt, fileName: review.fileName }]);
        notify("Spesa aggiunta al mese"); go("active");
      } else {
        const rawAmount = Number(f.amount) || 0;
        // Nota di credito → memorizzata con importo negativo (riduce i costi nel P&L)
        const amount = f.isCreditNote ? -Math.abs(rawAmount) : Math.abs(rawAmount);
        setPassive([...passive, {
          id: uid(),
          supplier: f.supplier,
          supplierCountry: f.supplierCountry || null,
          number: f.number,
          docType: f.isCreditNote ? "Nota di credito" : "Fattura",
          isCreditNote: !!f.isCreditNote,
          date: f.date,
          dueDate: f.dueDate,
          month: (f.date || "").slice(0, 7) || month,
          amount,
          currency: (f.currency || "EUR").toUpperCase(),
          vatExempt: !!f.vatExempt,
          vatRate: f.vatExempt ? 0 : (Number(f.vatRate) || 0),
          withholding: f.withholding != null && f.withholding !== "" ? Math.abs(Number(f.withholding)) || null : null,
          netAmount: f.taxableAmount != null && f.taxableAmount !== "" ? Math.abs(Number(f.taxableAmount)) || null : null,
          category: f.category || "Altro",
          notes: f.notes || "",
          status: "to_pay",
          matchedId: null,
          fileName: review.fileName,
        }]);
        notify(f.isCreditNote ? "Nota di credito registrata" : "Fattura registrata"); go("passive");
      }
      onClose();
    } catch (err) { notify("Errore nel salvataggio: " + (err.message || err), "err"); }
  };

  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel ia-bracket" style={{ width: 470, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{titleMap[review.kind]}</h3>
          <p style={{ margin: "5px 0 0", fontSize: 12, color: C.muted }}>Da «{review.fileName}» · verifica prima di salvare.</p>
        </div>
        <div style={{ padding: 20 }}>
          {review.kind === "contract" && <>
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
              <F label="Cliente"><input className="ia-input" value={f.client} onChange={e => set("client", e.target.value)} /></F>
              <F label="Valuta"><select className="ia-input" value={f.currency} onChange={e => set("currency", e.target.value)}><option>EUR</option><option>USD</option></select></F>
            </div>
            <BillingItemsEditor items={f.billingItems} currency={f.currency} onChange={v => set("billingItems", v)} />
            <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 13, background: "#FBFCFE" }}>
              <div style={{ fontSize: 13, fontFamily: "Poppins", color: C.ink, marginBottom: 8, fontWeight: 500 }}>Spese amministrative</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 12 }}>
                <F label="Percentuale (%)">
                  <input className="ia-input" type="number" step="0.1" value={f.adminFeePct == null ? "" : f.adminFeePct} onChange={e => set("adminFeePct", e.target.value === "" ? "" : Number(e.target.value))} placeholder="es. 8" />
                </F>
                <F label="Base di calcolo">
                  <select className="ia-input" value={f.adminFeeBase || "retainer"} onChange={e => set("adminFeeBase", e.target.value)}>
                    <option value="retainer">Solo retainer</option>
                    <option value="retainer_travel">Retainer + spese di viaggio</option>
                  </select>
                </F>
              </div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 6 }}>Voce automatica aggiunta ogni mese alle fatture pianificate. Lasciare vuoto o 0 per disattivare.</div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <F label="Inizio"><input className="ia-input" type="date" value={f.startDate} onChange={e => set("startDate", e.target.value)} /></F>
              <F label="Fine"><input className="ia-input" type="date" value={f.endDate} onChange={e => set("endDate", e.target.value)} /></F>
            </div>
            <F label="Note"><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>
          </>}
          {review.kind === "expense" && <>
            <F label="Associa al contratto"><select className="ia-input" value={f.contractId} onChange={e => set("contractId", e.target.value)}>
              {contracts.length === 0 && <option value="">— nessun contratto —</option>}
              {contracts.map(c => <option key={c.id} value={c.id}>{c.client}</option>)}
            </select></F>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <F label="Mese di competenza"><select className="ia-input" value={f.month} onChange={e => set("month", e.target.value)}>{monthOptions().map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}</select></F>
              <F label="Importo"><input className="ia-input" type="number" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
            </div>
            <F label="Descrizione"><input className="ia-input" value={f.description} onChange={e => set("description", e.target.value)} /></F>
          </>}
          {review.kind === "invoice" && <>
            {/* Banner esplicito per fatture in valuta ≠ EUR — l'utente deve verificare */}
            {f.currency !== "EUR" && (
              <div style={{ background: "#FFF8E6", border: "1px solid " + C.amber, padding: 10, marginBottom: 12, fontSize: 12.5, color: C.ink }}>
                <b style={{ color: C.amber }}>⚠ Fattura in valuta estera ({f.currency})</b>
                <div style={{ marginTop: 3 }}>Verifica la valuta rilevata. L'importo verrà registrato in {f.currency}; per il P&L in EUR applica manualmente il cambio in fase di riconciliazione bancaria.</div>
              </div>
            )}
            {/* Toggle nota di credito */}
            <div style={{ background: f.isCreditNote ? "#FBECEC" : "#FBFCFE", border: "1px solid " + (f.isCreditNote ? "#E7C9C9" : C.line), padding: 10, marginBottom: 12 }}>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={!!f.isCreditNote} onChange={e => set("isCreditNote", e.target.checked)} />
                È una <b>nota di credito</b> (verrà sottratta dai costi)
              </label>
            </div>

            <F label="Fornitore *"><input className="ia-input" value={f.supplier} onChange={e => set("supplier", e.target.value)} /></F>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <F label="Numero"><input className="ia-input" value={f.number} onChange={e => set("number", e.target.value)} /></F>
              <F label={f.supplierCountry ? `Paese fornitore (rilevato: ${f.supplierCountry})` : "Paese fornitore"}><input className="ia-input" value={f.supplierCountry || ""} onChange={e => set("supplierCountry", e.target.value.toUpperCase().slice(0, 2) || null)} placeholder="es. IT, US, GB, DE, CH…" maxLength={2} /></F>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <F label="Data"><input className="ia-input" type="date" value={f.date} onChange={e => set("date", e.target.value)} /></F>
              <F label="Scadenza"><input className="ia-input" type="date" value={f.dueDate} onChange={e => set("dueDate", e.target.value)} /></F>
            </div>
            {/* Ritenuta d'acconto e imponibile: compaiono quando il documento li riporta (parcelle) */}
            {(f.withholding != null || f.taxableAmount != null) && (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, background: "#F6F8FC", border: "1px solid " + C.line, padding: "10px 12px 0", marginBottom: 12 }}>
                <F label="Ritenuta d'acconto"><input className="ia-input" type="number" step="0.01" value={f.withholding ?? ""} onChange={e => set("withholding", e.target.value)} /></F>
                <F label="Imponibile (netto IVA)"><input className="ia-input" type="number" step="0.01" value={f.taxableAmount ?? ""} onChange={e => set("taxableAmount", e.target.value)} /></F>
              </div>
            )}
            {/* Importo + valuta — visibili e verificabili */}
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
              <F label={`Importo da pagare (${f.currency}) *`}><input className="ia-input" type="number" step="0.01" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
              <F label="Valuta">
                <select className="ia-input" value={f.currency} onChange={e => set("currency", e.target.value.toUpperCase())}>
                  <option value="EUR">EUR — Euro</option>
                  <option value="USD">USD — Dollaro USA</option>
                  <option value="GBP">GBP — Sterlina UK</option>
                  <option value="CHF">CHF — Franco svizzero</option>
                  <option value="JPY">JPY — Yen giapponese</option>
                  <option value="CAD">CAD — Dollaro canadese</option>
                  <option value="AUD">AUD — Dollaro australiano</option>
                  <option value="SEK">SEK — Corona svedese</option>
                  <option value="NOK">NOK — Corona norvegese</option>
                  <option value="DKK">DKK — Corona danese</option>
                  <option value="PLN">PLN — Zloty polacco</option>
                  <option value="CZK">CZK — Corona ceca</option>
                  <option value="HUF">HUF — Fiorino ungherese</option>
                  <option value="CNY">CNY — Yuan cinese</option>
                  <option value="HKD">HKD — Dollaro HK</option>
                  <option value="SGD">SGD — Dollaro Singapore</option>
                </select>
              </F>
            </div>
            {/* Blocco IVA — importante per fatture estere in reverse charge / non imponibile */}
            <div style={{ border: "1px solid " + C.line, padding: 12, marginBottom: 13, background: "#FBFCFE" }}>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontFamily: "Poppins", color: C.ink, cursor: "pointer" }}>
                <input type="checkbox" checked={!!f.vatExempt} onChange={e => set("vatExempt", e.target.checked)} />
                Fornitore <b>senza IVA</b> {f.currency !== "EUR" && <span style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>(tipico per fatture estere in reverse charge)</span>}
              </label>
              {!f.vatExempt && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                  <label style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins" }}>Aliquota IVA (%)</label>
                  <input className="ia-input" type="number" style={{ width: 90 }} value={f.vatRate} onChange={e => set("vatRate", e.target.value)} />
                </div>
              )}
            </div>
            <F label="Categoria di spesa">
              <select className="ia-input" value={f.category} onChange={e => set("category", e.target.value)}>
                {COST_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </F>
            {f.notes && <F label="Note (dall'estrazione)"><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>}
          </>}
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Scarta</button>
          <button className="ia-btn ia-btn-primary" onClick={save} disabled={review.kind === "expense" && !f.contractId}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ============ MANUAL DUES (F24 e simili) ============
const DUE_CATS = ["F24", "IVA", "Contributi INPS", "Ritenute", "Diritti/Bolli", "Altro"];
function ManualDues({ ctx }) {
  const { dues, setDues, notify } = ctx;
  const [edit, setEdit] = useState(null);
  const T = today();
  const list = [...dues].sort((a, b) => (a.dueDate || "").localeCompare(b.dueDate || ""));
  const openTot = dues.filter(d => d.status !== "paid").reduce((s, d) => s + (Number(d.amount) || 0), 0);
  const upsert = (d) => {
    const exists = dues.some(x => x.id === d.id);
    setDues(exists ? dues.map(x => x.id === d.id ? d : x) : [...dues, d]);
    setEdit(null); notify(exists ? "Scadenza aggiornata" : "Scadenza aggiunta");
  };
  const del = (id) => setDues(dues.filter(d => d.id !== id));
  const togglePaid = (d) => setDues(dues.map(x => x.id === d.id ? { ...x, status: x.status === "paid" ? "to_pay" : "paid", matchedId: x.status === "paid" ? null : x.matchedId } : x));
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 13.5, color: C.muted }}>Scadenze aperte: <b style={{ color: C.navy, fontFamily: "Poppins" }}>{fmtMoney(openTot)}</b></div>
        <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), label: "", category: "F24", amount: "", dueDate: today(), status: "to_pay", notes: "" })}><Plus size={16} /> Nuova scadenza</button>
      </div>
      {list.length === 0 ? <Empty icon={Receipt} title="Nessuna scadenza inserita" hint="Aggiungi qui gli F24, l'IVA, i contributi e ogni altra uscita programmata: entrano nel calendario e nella proiezione di cassa." /> : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th>Scadenza</th><th>Categoria</th><th>Data</th><th style={{ textAlign: "right" }}>Importo</th><th>Stato</th><th></th></tr></thead>
            <tbody>{list.map(d => {
              const late = d.status !== "paid" && d.dueDate && d.dueDate < T;
              return (
                <tr key={d.id}>
                  <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{d.label || "—"}<div style={{ fontWeight: 400, fontSize: 11.5, color: C.muted }}>{d.notes}</div></td>
                  <td>{d.category}</td>
                  <td style={{ color: late ? C.red : C.ink }}>{fmtDate(d.dueDate)}{late && <div style={{ fontSize: 10.5, color: C.red }}>scaduta</div>}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600 }}>{fmtMoney(d.amount)}</td>
                  <td><button onClick={() => togglePaid(d)} style={{ border: "none", background: "none", cursor: "pointer", padding: 0 }}><Chip status={d.status === "paid" ? "ppaid" : "to_pay"} /></button></td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 9px" }} onClick={() => setEdit(d)}><Pencil size={14} /></button>{" "}
                    <button className="ia-btn ia-btn-danger" style={{ padding: "6px 9px" }} onClick={() => del(d.id)}><Trash2 size={14} /></button>
                  </td>
                </tr>
              );
            })}</tbody>
          </table>
        </div>
      )}
      {edit && <DueModal d={edit} onSave={upsert} onClose={() => setEdit(null)} />}
    </div>
  );
}
function DueModal({ d, onSave, onClose }) {
  const [f, setF] = useState(d);
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 440, maxWidth: "100%" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Scadenza di pagamento</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Descrizione *"><input className="ia-input" placeholder="es. F24 agosto 2026" value={f.label} onChange={e => set("label", e.target.value)} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Categoria"><select className="ia-input" value={f.category} onChange={e => set("category", e.target.value)}>{DUE_CATS.map(c => <option key={c}>{c}</option>)}</select></F>
            <F label="Importo *"><input className="ia-input" type="number" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
          </div>
          <F label="Data scadenza *"><input className="ia-input" type="date" value={f.dueDate} onChange={e => set("dueDate", e.target.value)} /></F>
          <F label="Note"><input className="ia-input" value={f.notes} onChange={e => set("notes", e.target.value)} /></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.label || f.amount === "" || !f.dueDate} onClick={() => onSave({ ...f, amount: Number(f.amount) })}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// CONTO ECONOMICO gestionale (non ufficiale)
// ============================================================
function ProfitLoss({ ctx }) {
  const { contracts, active, setActive, expenses, issued, setIssued, passive, setPassive, dues, setDues, bank, setBank, bankRules, setBankRules, customBankCats, setCustomBankCats, notify } = ctx;
  const [yearSel, setYearSel] = useState(new Date().getFullYear());
  const [bankThreshold, setBankThreshold] = useState(300);
  const [reconcileTx, setReconcileTx] = useState(null);
  const [drillCategory, setDrillCategory] = useState(null); // categoria di cui vediamo il dettaglio
  const [drillMonth, setDrillMonth] = useState(null);       // mese di cui vediamo il dettaglio completo
  const [editIssued, setEditIssued] = useState(null);       // fattura emessa in modifica
  const [editPassive, setEditPassive] = useState(null);     // fattura passiva in modifica
  const [editDue, setEditDue] = useState(null);             // scadenza F24 in modifica
  const [bankFilter, setBankFilter] = useState("unmatched"); // "unmatched" | "matched" | "all"
  const [bulkCatUpdate, setBulkCatUpdate] = useState(null); // { kind, newCategoryId, items[] } — modale bulk categoria

  const combinedCats = useMemo(() => allBankCats(customBankCats), [customBankCats]);

  const monthOf = d => (d || "").slice(0, 7);
  const yearOf = d => (d || "").slice(0, 4);
  const sum = arr => arr.reduce((s, x) => s + (Number(x.amount) || 0), 0);

  // Netto passiva: se vatExempt → amount; altrimenti amount / (1 + rate/100)
  // Netto passiva: preservo il segno — le note di credito da fornitore hanno
  // importo negativo e devono ridurre i costi (riduzione = negativo).
  const netOfPassive = (p) => {
    const amt = Number(p.amount) || 0;
    const sign = amt < 0 ? -1 : 1;
    if (p.netAmount != null && p.netAmount !== "" && isFinite(Number(p.netAmount))) return sign * Math.abs(Number(p.netAmount)); // imponibile letto dal documento
    const gross = Math.abs(amt);
    if (p.vatExempt) return sign * gross;
    const rate = Number(p.vatRate);
    if (!rate || isNaN(rate)) return sign * (gross / 1.22); // default 22%
    return sign * (gross / (1 + rate / 100));
  };
  // Netto emesse: preservo il segno — le note di credito hanno importo negativo
  // e devono sottrarre dal fatturato, non aggiungere.
  const netOfIssued = (x) => Number(x.amount) || 0;

  // === F24: split per natura ===
  // INPS + Diritti/Bolli + Altro = costo. IVA + Ritenute = pass-through.
  const isDueCost = (d) => {
    const c = (d.category || "").toLowerCase();
    if (c.includes("inps") || c.includes("contribut")) return true;
    if (c.includes("bolli") || c.includes("diritt")) return true;
    if (c === "altro" || c === "f24") return true;
    return false;
  };
  const isDuePassthrough = (d) => {
    const c = (d.category || "").toLowerCase();
    return c.includes("iva") || c.includes("ritenut");
  };

  // === Costruzione dati mese per mese ===
  const yearIssued = issued.filter(x => yearOf(x.month || x.date) === String(yearSel));
  const yearPassive = passive.filter(p => yearOf(p.date) === String(yearSel));
  const yearDues = dues.filter(d => yearOf(d.dueDate || d.paidAt) === String(yearSel));
  // Movimenti banca categorizzati direttamente (Telepass, carta, biglietti, ecc.)
  const yearBankDirect = bank.filter(t => t.directCategory && yearOf(t.date) === String(yearSel));

  const months = Array.from({ length: 12 }, (_, i) => `${yearSel}-${String(i + 1).padStart(2, "0")}`);
  const monthly = months.map(m => {
    const iss = yearIssued.filter(x => (x.month || monthOf(x.date)) === m);
    const pas = yearPassive.filter(p => monthOf(p.date) === m);
    const du = yearDues.filter(d => monthOf(d.dueDate || d.paidAt) === m);
    const bnk = yearBankDirect.filter(t => monthOf(t.date) === m);
    const revenues = iss.reduce((s, x) => s + netOfIssued(x), 0);
    // Passive: split compensi amministratori
    const admin = pas.filter(p => (p.category || "").toLowerCase().includes("amministrator"));
    const otherPas = pas.filter(p => !(p.category || "").toLowerCase().includes("amministrator"));
    const adminCostPas = admin.reduce((s, p) => s + netOfPassive(p), 0);
    const otherCostPas = otherPas.reduce((s, p) => s + netOfPassive(p), 0);
    // Movimenti banca diretti: preservo il segno. Se marchi come "Software"
    // un rimborso ricevuto (importo bancario positivo), riduce i costi Software.
    // Le uscite sono negative in banca → devono sommare in valore assoluto ai costi.
    const bnkCost = bnk.filter(t => (bankCatMetaAll(t.directCategory, customBankCats)?.group) === "cost").reduce((s, t) => s - (Number(t.amount) || 0), 0);
    const bnkAdmin = bnk.filter(t => (bankCatMetaAll(t.directCategory, customBankCats)?.group) === "admin").reduce((s, t) => s - (Number(t.amount) || 0), 0);
    const bnkContrib = bnk.filter(t => (bankCatMetaAll(t.directCategory, customBankCats)?.group) === "contributi").reduce((s, t) => s - (Number(t.amount) || 0), 0);
    const otherCost = otherCostPas + bnkCost;
    const adminCost = adminCostPas + bnkAdmin;
    const contributi = du.filter(isDueCost).reduce((s, d) => s + Math.abs(Number(d.amount) || 0), 0) + bnkContrib;
    const margin = revenues - otherCost - adminCost - contributi;
    return { m, monthIdx: parseInt(m.slice(5, 7), 10) - 1, revenues, otherCost, adminCost, contributi, margin, marginPct: revenues > 0 ? (margin / revenues) * 100 : 0, issued: iss, passive: pas, dues: du, bank: bnk };
  });

  const tot = {
    revenues: monthly.reduce((s, x) => s + x.revenues, 0),
    otherCost: monthly.reduce((s, x) => s + x.otherCost, 0),
    adminCost: monthly.reduce((s, x) => s + x.adminCost, 0),
    contributi: monthly.reduce((s, x) => s + x.contributi, 0),
  };
  tot.margin = tot.revenues - tot.otherCost - tot.adminCost - tot.contributi;
  tot.marginPct = tot.revenues > 0 ? (tot.margin / tot.revenues) * 100 : 0;
  // Note di credito emesse nell'anno (importi negativi già sottratti dai ricavi)
  const creditNotes = yearIssued.filter(x => x.isCreditNote || (Number(x.amount) || 0) < 0);
  const creditNotesTotal = Math.abs(creditNotes.reduce((s, x) => s + (Number(x.amount) || 0), 0));

  // === Confronto anno precedente ===
  const prevYear = yearSel - 1;
  const prevIssued = issued.filter(x => yearOf(x.month || x.date) === String(prevYear));
  const prevPassive = passive.filter(p => yearOf(p.date) === String(prevYear));
  const prevDues = dues.filter(d => yearOf(d.dueDate || d.paidAt) === String(prevYear));
  const prevRevenues = prevIssued.reduce((s, x) => s + netOfIssued(x), 0);
  const prevAdmin = prevPassive.filter(p => (p.category || "").toLowerCase().includes("amministrator")).reduce((s, p) => s + netOfPassive(p), 0);
  const prevOther = prevPassive.filter(p => !(p.category || "").toLowerCase().includes("amministrator")).reduce((s, p) => s + netOfPassive(p), 0);
  const prevContrib = prevDues.filter(isDueCost).reduce((s, d) => s + Math.abs(Number(d.amount) || 0), 0);
  const prevMargin = prevRevenues - prevAdmin - prevOther - prevContrib;
  const yoyRev = prevRevenues > 0 ? ((tot.revenues - prevRevenues) / prevRevenues) * 100 : null;
  const yoyMar = prevMargin !== 0 ? ((tot.margin - prevMargin) / Math.abs(prevMargin)) * 100 : null;

  // === Top clienti / fornitori ===
  const clientAgg = {};
  yearIssued.forEach(x => { const k = x.client || "—"; clientAgg[k] = (clientAgg[k] || 0) + netOfIssued(x); });
  const topClients = Object.entries(clientAgg).sort((a, b) => b[1] - a[1]).slice(0, 5);
  const supplierAgg = {};
  yearPassive.filter(p => !(p.category || "").toLowerCase().includes("amministrator"))
    .forEach(p => { const k = p.supplier || "—"; supplierAgg[k] = (supplierAgg[k] || 0) + netOfPassive(p); });
  const topSuppliers = Object.entries(supplierAgg).sort((a, b) => b[1] - a[1]).slice(0, 5);

  // === Ripartizione F24 ===
  const f24Breakdown = { inps: 0, iva: 0, ritenute: 0, altro: 0 };
  yearDues.forEach(d => {
    const c = (d.category || "").toLowerCase();
    const amt = Math.abs(Number(d.amount) || 0);
    if (c.includes("inps") || c.includes("contribut")) f24Breakdown.inps += amt;
    else if (c.includes("iva")) f24Breakdown.iva += amt;
    else if (c.includes("ritenut")) f24Breakdown.ritenute += amt;
    else f24Breakdown.altro += amt;
  });

  // === Movimenti banca non riconciliati sopra soglia (esclusi ignorati e categorizzati) ===
  const unmatchedBank = bank.filter(t => !t.matchedId && !t.ignoredForPL && !t.directCategory && yearOf(t.date) === String(yearSel) && Math.abs(Number(t.amount) || 0) >= bankThreshold)
    .sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
  const matchedBank = bank.filter(t => t.matchedId && yearOf(t.date) === String(yearSel) && Math.abs(Number(t.amount) || 0) >= bankThreshold)
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const displayedBank = bankFilter === "unmatched" ? unmatchedBank : bankFilter === "matched" ? matchedBank : [...unmatchedBank, ...matchedBank];

  // === Fatture / scadenze aperte candidabili all'abbinamento ===
  // Solo fatture REALMENTE emesse (import ATTIVA + una tantum): le pianificate da
  // contratto sono proiezioni indicative, non devono partecipare alla riconciliazione.
  const openInvoices = useMemo(() => {
    const out = [];
    issued.filter(x => x.status !== "paid").forEach(x => out.push({ kind: "issued", id: x.id, label: x.client, sub: "emessa · " + (x.number || monthLabel((x.date || "").slice(0, 7))), amount: Number(x.amount), currency: x.currency || "EUR", sign: 1, number: x.number, date: x.date }));
    passive.filter(p => p.status !== "ppaid").forEach(p => out.push({ kind: "passive", id: p.id, label: p.supplier, sub: "passiva · " + (p.number || ""), amount: Number(p.amount), currency: p.currency, sign: -1, number: p.number, date: p.date }));
    dues.filter(d => d.status !== "paid").forEach(d => out.push({ kind: "due", id: d.id, label: d.label, sub: (d.category || "Scadenza"), amount: Number(d.amount), currency: "EUR", sign: -1, number: null, date: d.dueDate || d.paidAt }));
    return out;
  }, [passive, issued, dues]);

  const suggestFor = (tx) => {
    const wantSign = tx.amount >= 0 ? 1 : -1;
    const desc = String(tx.description || "").toLowerCase();
    const descNorm = normName(tx.description);
    return openInvoices
      .filter(inv => inv.sign === wantSign)
      .map(inv => {
        const txAbs = Math.abs(tx.amount), invAbs = Math.abs(inv.amount);
        const amtDelta = Math.abs(txAbs - invAbs);
        const amountExact = amtDelta < 1.0;
        const amountLoose = invAbs > 0 && amtDelta / invAbs < 0.05;
        // Matching sul nome — cerca ogni parola distintiva del cliente/fornitore nella descrizione bancaria
        const nameWords = normName(inv.label).split(/\s+/).filter(w => w.length >= 4);
        const nameHits = nameWords.filter(w => descNorm.includes(w)).length;
        const nameStrongHit = nameWords.length > 0 && (nameHits >= 2 || (nameHits === 1 && nameWords.length === 1));
        const nameWeakHit = nameHits >= 1;
        // Matching sul numero fattura — usa l'helper robusto che riconosce marker italiani
        // ("RIF. 53", "FT 73", "N. 89") ed evita di confondere il numero con la data
        const numInDesc = invoiceNumberInDesc(inv.number, tx.description);
        // Confidenza 4 (sicuro): importo esatto + (nome forte O numero)
        // Confidenza 3 (probabile): importo esatto SENZA nome/numero, OPPURE loose + nome forte
        // Confidenza 2 (possibile): importo loose + nome debole
        // Confidenza 1 (ipotesi): solo importo esatto MA fra molte fatture simili? no, saltato
        let conf = 0;
        if (amountExact && (nameStrongHit || numInDesc)) conf = 4;
        else if (amountExact && nameWeakHit) conf = 3;
        else if (amountExact) conf = 2;
        else if (amountLoose && (nameStrongHit || numInDesc)) conf = 3;
        else if (amountLoose && nameWeakHit) conf = 1;
        if (conf === 0) return null;
        return { ...inv, _conf: conf, _amtDelta: amtDelta, _nameHit: nameStrongHit || nameWeakHit, _numInDesc: numInDesc };
      })
      .filter(Boolean)
      .sort((a, b) => b._conf - a._conf || a._amtDelta - b._amtDelta);
  };

  const matchBank = (tx, inv) => {
    setBank(bank.map(t => t.id === tx.id ? { ...t, matchedId: inv.id, matchedType: inv.kind, matchedLabel: inv.label } : t));
    if (inv.kind === "issued") setIssued(issued.map(x => x.id === inv.id ? { ...x, status: "paid", matchedId: tx.id } : x));
    else if (inv.kind === "due") setDues(dues.map(d => d.id === inv.id ? { ...d, status: "paid", matchedId: tx.id } : d));
    else if (inv.kind === "active") {
      const rec = active.find(a => a.contractId === inv.contractId && a.month === inv.month);
      if (rec) setActive(active.map(a => a === rec ? { ...a, status: "paid", matchedId: tx.id } : a));
      else setActive([...active, { contractId: inv.contractId, month: inv.month, status: "paid", matchedId: tx.id }]);
    } else setPassive(passive.map(p => p.id === inv.id ? { ...p, status: "ppaid", matchedId: tx.id } : p));
    notify("Movimento abbinato");
  };

  const ignoreBank = (tx) => {
    setBank(bank.map(t => t.id === tx.id ? { ...t, ignoredForPL: true } : t));
    notify("Movimento contrassegnato come non P&L (giroconto/personale)");
  };

  // Conta i movimenti auto-abbinabili con criterio rilassato:
  // - Importo esatto (< €1 di scarto)
  // - AND (almeno una parola del nome cliente/fornitore in descrizione OR numero fattura in descrizione)
  // - AND: candidato unico, OPPURE più candidati ma uno solo ha il numero fattura in descrizione (tiebreak)
  const autoMatchCount = useMemo(() => {
    const used = new Set(); let count = 0;
    unmatchedBank.forEach(t => {
      const wantSign = t.amount >= 0 ? 1 : -1;
      const descNorm = normName(t.description);
      const scored = openInvoices.filter(inv => inv.sign === wantSign && !used.has(inv.kind + ":" + inv.id))
        .map(inv => {
          const txAbs = Math.abs(t.amount), invAbs = Math.abs(inv.amount);
          const amtDelta = Math.abs(txAbs - invAbs);
          if (amtDelta >= 1) return null;
          const nameWords = normName(inv.label).split(/\s+/).filter(w => w.length >= 4);
          const nameHits = nameWords.filter(w => descNorm.includes(w)).length;
          const anyNameHit = nameHits >= 1;
          const numInDesc = invoiceNumberInDesc(inv.number, t.description);
          return (anyNameHit || numInDesc) ? { inv, key: inv.kind + ":" + inv.id, numInDesc } : null;
        }).filter(Boolean);
      if (scored.length === 0) return;
      // Unico candidato → match
      if (scored.length === 1) { count++; used.add(scored[0].key); return; }
      // Più candidati ma uno solo ha il numero fattura → tiebreak vinto da quello
      const withNum = scored.filter(s => s.numInDesc);
      if (withNum.length === 1) { count++; used.add(withNum[0].key); }
    });
    return count;
  }, [unmatchedBank, openInvoices]);

  // Conta quanti movimenti dell'anno selezionato (sopra soglia, non gestiti) matchano
  // le regole di categorizzazione automatica già memorizzate — pronti per l'auto-applicazione.
  const autoRulesCount = useMemo(() => {
    if (!bankRules.length) return 0;
    let count = 0;
    bank.forEach(t => {
      if (t.matchedId || t.directCategory || t.ignoredForPL) return;
      if (yearOf(t.date) !== String(yearSel)) return;
      if (Math.abs(Number(t.amount) || 0) < bankThreshold) return;
      if (bankRules.some(r => matchesBankRule(t.description, r))) count++;
    });
    return count;
  }, [bank, bankRules, yearSel, bankThreshold]);

  // Applica in blocco le regole memorizzate a tutti i movimenti storici non gestiti che matchano.
  // Utile per catturare movimenti antecedenti alla creazione della regola (Telepass, canoni, ecc.).
  const autoApplyRules = () => {
    if (bankRules.length === 0) { notify("Nessuna regola memorizzata ancora"); return; }
    const toApply = [];
    bank.forEach(t => {
      if (t.matchedId || t.directCategory || t.ignoredForPL) return;
      if (yearOf(t.date) !== String(yearSel)) return;
      if (Math.abs(Number(t.amount) || 0) < bankThreshold) return;
      for (const rule of bankRules) {
        if (matchesBankRule(t.description, rule)) {
          toApply.push({ id: t.id, category: rule.category, desc: t.description, amount: t.amount, date: t.date });
          break;
        }
      }
    });
    if (toApply.length === 0) { notify("Nessun movimento non gestito matcha le regole esistenti"); return; }
    const byCat = {};
    toApply.forEach(x => { byCat[x.category] = (byCat[x.category] || 0) + 1; });
    const summary = Object.entries(byCat).map(([cid, n]) => {
      const meta = bankCatMetaAll(cid, customBankCats);
      return `• ${meta?.label || cid}: ${n} movimenti`;
    }).join("\n");
    if (!confirm(`Applicare le regole di categorizzazione a ${toApply.length} movimenti non gestiti dell'anno ${yearSel} (sopra €${bankThreshold})?\n\n${summary}\n\nPuoi sempre annullare le categorizzazioni con "Annulla categorie" nella toolbar.`)) return;
    const idToCat = new Map(toApply.map(x => [x.id, x.category]));
    setBank(bank.map(t => idToCat.has(t.id) ? { ...t, directCategory: idToCat.get(t.id), directCategoryFromRule: true } : t));
    notify(`${toApply.length} movimenti auto-categorizzati da regole memorizzate`);
  };

  const categorizeBank = (tx, categoryId) => {
    const meta = bankCatMetaAll(categoryId, customBankCats);
    // Impara una regola per il futuro
    const keywords = extractBankKeywords(tx.description);
    let updatedRules = bankRules;
    let extraApplied = 0;
    let updatedBank = bank.map(t => t.id === tx.id ? { ...t, directCategory: categoryId } : t);
    if (keywords.length > 0) {
      const kwKey = keywords.join("|");
      updatedRules = bankRules.filter(r => (r.keywords || []).join("|") !== kwKey);
      updatedRules.push({ id: uid(), keywords, category: categoryId, createdAt: new Date().toISOString(), lastUsedAt: new Date().toISOString() });
      updatedBank = updatedBank.map(t => {
        if (t.id === tx.id) return t;
        if (t.matchedId || t.directCategory || t.ignoredForPL) return t;
        if (matchesBankRule(t.description, { keywords })) {
          extraApplied++;
          return { ...t, directCategory: categoryId, directCategoryFromRule: true };
        }
        return t;
      });
    }
    setBank(updatedBank);
    if (updatedRules !== bankRules) setBankRules(updatedRules);
    notify(meta ? `Categorizzato come ${meta.label}${extraApplied > 0 ? ` (+ ${extraApplied} movimenti simili)` : ""}` : "Categorizzato");
  };

  const addCustomCategory = (label, group) => {
    const clean = String(label || "").trim();
    if (!clean) { notify("Inserisci un nome per la categoria", "err"); return null; }
    // Evita duplicati (case-insensitive)
    const existing = combinedCats.find(c => c.label.toLowerCase() === clean.toLowerCase());
    if (existing) { notify(`Categoria "${existing.label}" già esistente`, "err"); return existing.id; }
    const id = "custom_" + uid();
    const newCat = { id, label: clean, group: group || "cost" };
    setCustomBankCats([...customBankCats, newCat]);
    return id;
  };

  const deleteCustomCategory = (id) => {
    if (!id.startsWith("custom_")) return;
    const inUse = bank.some(t => t.directCategory === id);
    if (inUse && !confirm("Alcuni movimenti usano questa categoria. Eliminarla comunque? Quei movimenti torneranno senza categoria.")) return;
    setCustomBankCats(customBankCats.filter(c => c.id !== id));
    if (inUse) setBank(bank.map(t => t.directCategory === id ? { ...t, directCategory: null, directCategoryFromRule: false } : t));
    notify("Categoria eliminata");
  };

  // === Auto-abbina in massa i movimenti con confidenza SICURO ===
  const autoMatchAll = () => {
    const used = new Set();
    const proposed = [];
    unmatchedBank.forEach(t => {
      const sug = suggestFor(t).filter(s => !used.has(s.kind + ":" + s.id));
      if (sug.length === 0) return;
      // Filtra ai candidati "forti": importo esatto + (nome nella descrizione OR numero fattura nella descrizione)
      const strong = sug.filter(s => {
        const amtDelta = Math.abs(Math.abs(s.amount) - Math.abs(t.amount));
        return amtDelta < 1 && (s._nameHit || s._numInDesc);
      });
      if (strong.length === 0) return;
      let winner = null;
      if (strong.length === 1) {
        // Unico candidato forte → vince
        winner = strong[0];
      } else {
        // Più candidati con stesso importo e nome — tiebreak: uno solo ha il numero fattura riconosciuto
        const withNum = strong.filter(s => s._numInDesc);
        if (withNum.length === 1) winner = withNum[0];
      }
      if (winner) {
        proposed.push({ tx: t, inv: winner });
        used.add(winner.kind + ":" + winner.id);
      }
    });
    if (proposed.length === 0) { notify("Nessun abbinamento automatico trovato"); return; }
    const preview = proposed.slice(0, 4).map(({ tx, inv }) => `• ${fmtDate(tx.date)} · ${fmtMoney(tx.amount)} → ${inv.label}`).join("\n");
    if (!confirm(`Auto-abbinare ${proposed.length} movimenti (importo esatto + nome cliente/numero fattura nella descrizione)?\n\n${preview}${proposed.length > 4 ? `\n… e altri ${proposed.length - 4}` : ""}`)) return;
    let newBank = bank, newIssued = issued, newDues = dues, newActive = active, newPassive = passive;
    proposed.forEach(({ tx, inv }) => {
      newBank = newBank.map(t => t.id === tx.id ? { ...t, matchedId: inv.id, matchedType: inv.kind, matchedLabel: inv.label } : t);
      if (inv.kind === "issued") newIssued = newIssued.map(x => x.id === inv.id ? { ...x, status: "paid", matchedId: tx.id } : x);
      else if (inv.kind === "due") newDues = newDues.map(d => d.id === inv.id ? { ...d, status: "paid", matchedId: tx.id } : d);
      else if (inv.kind === "active") {
        const rec = newActive.find(a => a.contractId === inv.contractId && a.month === inv.month);
        if (rec) newActive = newActive.map(a => a === rec ? { ...a, status: "paid", matchedId: tx.id } : a);
        else newActive = [...newActive, { contractId: inv.contractId, month: inv.month, status: "paid", matchedId: tx.id }];
      } else newPassive = newPassive.map(p => p.id === inv.id ? { ...p, status: "ppaid", matchedId: tx.id } : p);
    });
    setBank(newBank); setIssued(newIssued); setDues(newDues); setActive(newActive); setPassive(newPassive);
    notify(`${proposed.length} movimenti abbinati automaticamente`);
  };

  const unmatchBank = (tx) => {
    // Rimuovi l'abbinamento di questo movimento (torna tra "da chiarire") e riporta la fattura a stato aperto
    const invKind = tx.matchedType, invId = tx.matchedId;
    setBank(bank.map(t => t.id === tx.id ? { ...t, matchedId: null, matchedType: null, matchedLabel: null } : t));
    if (invKind === "issued") setIssued(issued.map(x => x.id === invId ? { ...x, status: "issued", matchedId: null } : x));
    else if (invKind === "due") setDues(dues.map(d => d.id === invId ? { ...d, status: "todo", matchedId: null } : d));
    else if (invKind === "passive") setPassive(passive.map(p => p.id === invId ? { ...p, status: "to_pay", matchedId: null } : p));
    else if (invKind === "active") setActive(active.map(a => (a.matchedId === tx.id) ? { ...a, status: "planned", matchedId: null } : a));
    notify("Abbinamento annullato");
  };

  // Reset globale di TUTTE le riconciliazioni: sblocca ogni fattura marcata come pagata,
  // annulla ogni abbinamento bancario, riporta tutto allo stato "aperto".
  // NON tocca: anagrafica fatture/movimenti, categorizzazioni dirette, movimenti ignorati.
  const resetAllReconciliations = () => {
    const paidIssued = issued.filter(x => x.status === "paid" || x.matchedId).length;
    const paidPassive = passive.filter(p => p.status === "ppaid" || p.matchedId).length;
    const paidDues = dues.filter(d => d.status === "paid" || d.matchedId).length;
    const matchedBankCount = bank.filter(t => t.matchedId).length;
    const totalToReset = paidIssued + paidPassive + paidDues + matchedBankCount;
    if (totalToReset === 0) { notify("Nessuna riconciliazione da resettare"); return; }
    if (!confirm(`ATTENZIONE: sto per resettare TUTTE le riconciliazioni esistenti.\n\n• ${paidIssued} fatture emesse tornano "aperte"\n• ${paidPassive} fatture passive tornano "da pagare"\n• ${paidDues} scadenze F24 tornano "da pagare"\n• ${matchedBankCount} movimenti bancari tornano "da chiarire"\n\nLe anagrafiche di fatture e movimenti restano invariate.\nLe categorizzazioni dirette dei movimenti (Telepass, carte, ecc.) NON vengono toccate.\n\nDopo il reset potrai ricliccare "Auto-abbina fatture" con il matcher aggiornato.\n\nProcedere?`)) return;
    setIssued(issued.map(x => ({ ...x, status: "issued", matchedId: null })));
    setPassive(passive.map(p => ({ ...p, status: "to_pay", matchedId: null })));
    setDues(dues.map(d => ({ ...d, status: "todo", matchedId: null })));
    setBank(bank.map(t => ({ ...t, matchedId: null, matchedType: null, matchedLabel: null })));
    setActive(active.map(a => a.matchedId ? { ...a, status: "planned", matchedId: null } : a));
    notify(`Reset completato: ${totalToReset} riconciliazioni azzerate`);
  };

  // Ripulisce i movimenti bancari duplicati (stesso CSV importato due volte).
  const cleanupBankDuplicatesPL = () => {
    const bankKey = t => (t.date || "").slice(0, 10) + "|" + (Number(t.amount) || 0).toFixed(2) + "|" + normName(t.description || "");
    const groups = new Map();
    bank.forEach(t => {
      const key = bankKey(t);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(t);
    });
    const duplicates = [];
    groups.forEach(arr => {
      if (arr.length > 1) {
        const sorted = [...arr].sort((a, b) => {
          const aHandled = (a.matchedId || a.directCategory || a.ignoredForPL) ? 0 : 1;
          const bHandled = (b.matchedId || b.directCategory || b.ignoredForPL) ? 0 : 1;
          return aHandled - bHandled;
        });
        duplicates.push(...sorted.slice(1));
      }
    });
    if (duplicates.length === 0) { notify("Nessun duplicato bancario trovato"); return; }
    const preview = duplicates.slice(0, 5).map(t => `• ${fmtDate(t.date)} · ${fmtMoney(t.amount)} · ${String(t.description || "").slice(0, 60)}`).join("\n");
    if (!confirm(`Trovati ${duplicates.length} movimenti bancari duplicati (stessa data + stesso importo + stessa descrizione).\n\nPrimi da eliminare:\n${preview}${duplicates.length > 5 ? "\n…" : ""}\n\nProcedere?`)) return;
    const toRemove = new Set(duplicates.map(t => t.id));
    setBank(bank.filter(t => !toRemove.has(t.id)));
    notify(`${duplicates.length} duplicati bancari rimossi`);
  };

  const recategorizeBank = (tx, newCategoryId) => {
    const oldCategoryId = tx.directCategory;
    setBank(bank.map(t => t.id === tx.id ? { ...t, directCategory: newCategoryId, directCategoryFromRule: false } : t));
    const meta = bankCatMetaAll(newCategoryId, customBankCats);
    notify(meta ? `Riclassificato come ${meta.label}` : "Riclassificato");
    // Cerca altri movimenti con descrizione simile (keyword estratte) categorizzati diversamente
    // nell'anno corrente. Se ne trova almeno uno, propone la modifica in blocco.
    if (newCategoryId === oldCategoryId) return;
    const keywords = extractBankKeywords(tx.description);
    if (keywords.length === 0) return;
    const similar = bank.filter(t => {
      if (t.id === tx.id) return false;
      if (!t.directCategory) return false;
      if (t.directCategory === newCategoryId) return false;
      if (yearOf(t.date) !== String(yearSel)) return false;
      return matchesBankRule(t.description, { keywords });
    });
    if (similar.length === 0) return;
    setBulkCatUpdate({
      kind: "bank",
      newCategoryId,
      items: similar.map(t => ({
        id: t.id,
        primary: t.description,
        secondary: `${fmtDate(t.date)} · ${fmtMoney(t.amount, t.currency)} · attuale: ${bankCatMetaAll(t.directCategory, customBankCats)?.label || t.directCategory}`,
        currentCategoryId: t.directCategory,
      })),
    });
  };

  const uncategorizeBank = (tx) => {
    setBank(bank.map(t => t.id === tx.id ? { ...t, directCategory: null, directCategoryFromRule: false } : t));
    notify("Categoria rimossa — il movimento torna tra quelli da chiarire");
  };

  const deleteRule = (ruleId) => {
    if (!confirm("Eliminare questa regola di categorizzazione automatica? I movimenti già categorizzati non verranno modificati.")) return;
    setBankRules(bankRules.filter(r => r.id !== ruleId));
    notify("Regola eliminata");
  };

  // Upsert e delete per la drill-down mensile (modifica anomalie)
  const upsertIssued = (x) => {
    const exists = issued.some(y => y.id === x.id);
    setIssued(exists ? issued.map(y => y.id === x.id ? x : y) : [x, ...issued]);
    setEditIssued(null); notify(exists ? "Fattura aggiornata" : "Fattura salvata");
  };
  const deleteIssuedRow = (id) => { if (!confirm("Eliminare questa fattura?")) return; setIssued(issued.filter(y => y.id !== id)); notify("Fattura eliminata"); };
  const upsertPassive = (p) => {
    const exists = passive.some(y => y.id === p.id);
    const previous = exists ? passive.find(y => y.id === p.id) : null;
    setPassive(exists ? passive.map(y => y.id === p.id ? p : y) : [p, ...passive]);
    setEditPassive(null); notify(exists ? "Fattura passiva aggiornata" : "Fattura passiva salvata");
    // Se ha cambiato la categoria, cerca altre fatture stesso fornitore con categoria diversa
    // nell'anno corrente. Se ne trova, propone la modifica in blocco.
    if (!exists || !previous || previous.category === p.category || !p.category) return;
    const supplierNorm = normName(p.supplier);
    if (!supplierNorm) return;
    const similar = passive.filter(y => {
      if (y.id === p.id) return false;
      if (y.category === p.category) return false;
      if (normName(y.supplier) !== supplierNorm) return false;
      if (yearOf(y.date) !== String(yearSel)) return false;
      return true;
    });
    if (similar.length === 0) return;
    setBulkCatUpdate({
      kind: "passive",
      newCategoryId: p.category,
      items: similar.map(y => ({
        id: y.id,
        primary: `${y.supplier} — ${y.number || "s.n."}`,
        secondary: `${fmtDate(y.date)} · ${fmtMoney(y.amount, y.currency)} · attuale: ${y.category || "senza categoria"}`,
        currentCategoryId: y.category,
      })),
    });
  };
  const deletePassiveRow = (id) => { if (!confirm("Eliminare questa fattura passiva?")) return; setPassive(passive.filter(y => y.id !== id)); notify("Fattura eliminata"); };
  const upsertDue = (d) => {
    const exists = dues.some(y => y.id === d.id);
    setDues(exists ? dues.map(y => y.id === d.id ? d : y) : [d, ...dues]);
    setEditDue(null); notify(exists ? "Scadenza aggiornata" : "Scadenza salvata");
  };
  const deleteDueRow = (id) => { if (!confirm("Eliminare questa scadenza?")) return; setDues(dues.filter(y => y.id !== id)); notify("Scadenza eliminata"); };

  // === PDF export ===
  const printPdf = () => {
    const w = window.open("", "_blank", "width=1000,height=1000");
    if (!w) { notify("Il browser ha bloccato la finestra — abilita i popup", "err"); return; }
    const rowsHtml = monthly.map(x => `
      <tr>
        <td>${M_LABELS[x.monthIdx]}</td>
        <td class="num">${fmtMoney(x.revenues)}</td>
        <td class="num">${x.otherCost > 0 ? fmtMoney(x.otherCost) : "—"}</td>
        <td class="num">${x.adminCost > 0 ? fmtMoney(x.adminCost) : "—"}</td>
        <td class="num">${x.contributi > 0 ? fmtMoney(x.contributi) : "—"}</td>
        <td class="num ${x.margin >= 0 ? 'pos' : 'neg'}"><b>${fmtMoney(x.margin)}</b></td>
        <td class="num muted">${x.marginPct.toFixed(1)}%</td>
      </tr>`).join("");
    const clientsHtml = topClients.length ? topClients.map(([c, v]) => `<tr><td>${escapeHtml(c)}</td><td class="num">${fmtMoney(v)}</td><td class="num muted">${((v / tot.revenues) * 100).toFixed(1)}%</td></tr>`).join("") : `<tr><td colspan="3" class="muted">—</td></tr>`;
    const suppliersHtml = topSuppliers.length ? topSuppliers.map(([s, v]) => `<tr><td>${escapeHtml(s)}</td><td class="num">${fmtMoney(v)}</td><td class="num muted">${((v / (tot.otherCost || 1)) * 100).toFixed(1)}%</td></tr>`).join("") : `<tr><td colspan="3" class="muted">—</td></tr>`;
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>Conto Economico ${yearSel}</title>
      <style>
        @page { size: A4; margin: 14mm; }
        body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #1A1F35; padding: 20px; }
        .head { display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 3px solid #131D4B; padding-bottom: 10px; margin-bottom: 6px; }
        .head h1 { margin: 0; font-size: 22px; color: #131D4B; }
        .head .sub { color: #6B7392; font-size: 12px; margin-top: 4px; }
        .brand { color: #0E2289; font-weight: bold; font-size: 12px; }
        .brand .orange { color: #F57547; }
        .warn { background: #FFF8E6; border-left: 3px solid #F5C56A; padding: 10px 14px; font-size: 11.5px; color: #6B5A21; margin: 14px 0; }
        h2 { font-size: 14px; color: #131D4B; margin: 22px 0 8px; letter-spacing: .3px; }
        table { width: 100%; border-collapse: collapse; font-size: 11.5px; margin-bottom: 8px; }
        thead th { background: #131D4B; color: #fff; text-align: left; padding: 7px 9px; font-weight: 600; }
        tbody td { padding: 6px 9px; border-bottom: 1px solid #E1E5EE; }
        .num { text-align: right; font-variant-numeric: tabular-nums; }
        .pos { color: #2F8259; }
        .neg { color: #B23A3A; }
        .muted { color: #6B7392; }
        tfoot td { padding: 9px; background: #F0F3F8; border-top: 2px solid #131D4B; font-weight: 700; font-size: 12px; }
        .kpi { display: flex; gap: 12px; margin: 12px 0 6px; }
        .kpi div { flex: 1; background: #FAFBFD; border: 1px solid #E1E5EE; padding: 10px; }
        .kpi .lbl { font-size: 10px; color: #6B7392; text-transform: uppercase; letter-spacing: .4px; }
        .kpi .val { font-size: 18px; font-weight: 700; color: #131D4B; margin-top: 3px; }
        .cols { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
        .foot { margin-top: 16px; font-size: 10px; color: #8891B5; text-align: right; }
      </style>
    </head><body>
      <div class="head">
        <div>
          <h1>Conto economico gestionale ${yearSel}</h1>
          <div class="sub">Vista non ufficiale, ricostruita dai dati inseriti in IA Suite</div>
        </div>
        ${IS_INC ? `<img src="${IA_INC_LOGO_URL()}" alt="International Advisors Inc." style="height:46px;width:auto;display:block">` : `<div class="brand"><span class="orange">[</span> IA INTERNATIONAL ADVISORS</div>`}
      </div>
      <div class="warn"><b>Non è un bilancio ufficiale.</b> Basato su fatture emesse e ricevute inserite nell'app. Non include ammortamenti, TFR, dividendi ai soci non fatturati, o altre poste non tracciate. L'IVA sui costi è stimata per differenza.</div>
      <div class="kpi">
        <div><div class="lbl">Ricavi</div><div class="val">${fmtMoney(tot.revenues)}</div></div>
        <div><div class="lbl">Costi (esclusi comp. ammin.)</div><div class="val">${fmtMoney(tot.otherCost)}</div></div>
        <div><div class="lbl">Compensi amministratori</div><div class="val">${fmtMoney(tot.adminCost)}</div></div>
        <div><div class="lbl">Margine (${tot.marginPct.toFixed(1)}%)</div><div class="val" style="color:${tot.margin >= 0 ? '#2F8259' : '#B23A3A'}">${fmtMoney(tot.margin)}</div></div>
      </div>
      <h2>Dettaglio mensile</h2>
      <table>
        <thead><tr><th>Mese</th><th class="num">Ricavi</th><th class="num">Costi</th><th class="num">Comp. ammin.</th><th class="num">Contributi</th><th class="num">Margine</th><th class="num">%</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
        <tfoot><tr><td>TOTALE ${yearSel}</td><td class="num">${fmtMoney(tot.revenues)}</td><td class="num">${fmtMoney(tot.otherCost)}</td><td class="num">${fmtMoney(tot.adminCost)}</td><td class="num">${fmtMoney(tot.contributi)}</td><td class="num ${tot.margin >= 0 ? 'pos' : 'neg'}">${fmtMoney(tot.margin)}</td><td class="num">${tot.marginPct.toFixed(1)}%</td></tr></tfoot>
      </table>
      <div class="cols">
        <div>
          <h2>Top clienti per fatturato</h2>
          <table><thead><tr><th>Cliente</th><th class="num">Ricavi</th><th class="num">%</th></tr></thead><tbody>${clientsHtml}</tbody></table>
        </div>
        <div>
          <h2>Top fornitori (escl. compensi ammin.)</h2>
          <table><thead><tr><th>Fornitore</th><th class="num">Costi</th><th class="num">%</th></tr></thead><tbody>${suppliersHtml}</tbody></table>
        </div>
      </div>
      <h2>Ripartizione scadenze F24</h2>
      <table><thead><tr><th>Voce</th><th class="num">Importo</th><th>Trattamento</th></tr></thead><tbody>
        <tr><td>Contributi INPS</td><td class="num">${fmtMoney(f24Breakdown.inps)}</td><td class="muted">Costo</td></tr>
        <tr><td>IVA</td><td class="num">${fmtMoney(f24Breakdown.iva)}</td><td class="muted">Pass-through (non conteggiata)</td></tr>
        <tr><td>Ritenute</td><td class="num">${fmtMoney(f24Breakdown.ritenute)}</td><td class="muted">Pass-through (non conteggiata)</td></tr>
        <tr><td>Altro F24 / Bolli</td><td class="num">${fmtMoney(f24Breakdown.altro)}</td><td class="muted">Costo</td></tr>
      </tbody></table>
      <div class="foot">Generato il ${fmtDate(today())} da IA Suite — Conto economico gestionale, non ufficiale.</div>
      <script>window.onload = () => setTimeout(() => window.print(), 300);</script>
    </body></html>`;
    w.document.open(); w.document.write(html); w.document.close();
  };

  return (
    <div>
      {/* Toolbar */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 10px" }} onClick={() => setYearSel(yearSel - 1)}>‹</button>
          <span className="ia-h" style={{ fontSize: 18, color: C.navy, minWidth: 60, textAlign: "center" }}>{yearSel}</span>
          <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 10px" }} onClick={() => setYearSel(yearSel + 1)}>›</button>
        </div>
        <button className="ia-btn ia-btn-primary" onClick={printPdf}><Download size={15} /> Esporta PDF</button>
      </div>

      {/* Avviso "non ufficiale" */}
      <div style={{ background: "#FFF8E6", border: "1px solid #F5C56A", padding: "10px 14px", marginBottom: 16, fontSize: 12.5, color: "#6B5A21", display: "flex", gap: 10 }}>
        <AlertTriangle size={16} style={{ marginTop: 2, flexShrink: 0 }} />
        <div><b>Conto economico gestionale, non ufficiale.</b> Ricavi = netto imponibile fatture emesse. Costi = netto stimato fatture ricevute (dall'importo lordo tolta l'IVA usando l'aliquota indicata su ogni fattura, 22% se non specificata, 0% se marcata "senza IVA"). Non include ammortamenti, TFR, dividendi non fatturati e altre poste non tracciate.</div>
      </div>

      {/* KPI */}
      <div className="ia-kpis" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 14, marginBottom: 18 }}>
        <div className="ia-panel ia-bracket" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Ricavi {yearSel}</div>
          <div className="ia-kpi-num" style={{ color: C.green, margin: "4px 0" }}>{fmtMoney(tot.revenues)}</div>
          {creditNotesTotal > 0 && <div style={{ fontSize: 11, color: C.red, marginTop: 2 }}>Al netto di <b>{fmtMoney(creditNotesTotal)}</b> in note di credito ({creditNotes.length})</div>}
          {yoyRev !== null && <div style={{ fontSize: 11.5, color: yoyRev >= 0 ? C.green : C.red, fontFamily: "Poppins", fontWeight: 600, marginTop: 3 }}>{yoyRev >= 0 ? "▲" : "▼"} {Math.abs(yoyRev).toFixed(1)}% vs {prevYear}</div>}
        </div>
        <div className="ia-panel" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Costi (esclusi comp. ammin.)</div>
          <div className="ia-kpi-num" style={{ color: C.red, margin: "4px 0" }}>{fmtMoney(tot.otherCost)}</div>
          <div style={{ fontSize: 11, color: C.muted }}>Fornitori esterni + F24 contributi</div>
        </div>
        <div className="ia-panel" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Compensi amministratori</div>
          <div className="ia-kpi-num" style={{ color: C.orange, margin: "4px 0" }}>{fmtMoney(tot.adminCost)}</div>
          <div style={{ fontSize: 11, color: C.muted }}>Voce distinta nel P&L</div>
        </div>
        <div className="ia-panel" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Margine</div>
          <div className="ia-kpi-num" style={{ color: tot.margin >= 0 ? C.green : C.red, margin: "4px 0" }}>{fmtMoney(tot.margin)}</div>
          <div style={{ fontSize: 12, color: tot.margin >= 0 ? C.green : C.red, fontFamily: "Poppins", fontWeight: 600 }}>{tot.marginPct.toFixed(1)}% sui ricavi{yoyMar !== null && <span style={{ color: C.muted, fontWeight: 400, marginLeft: 6 }}>{yoyMar >= 0 ? "▲" : "▼"} {Math.abs(yoyMar).toFixed(1)}% vs {prevYear}</span>}</div>
        </div>
      </div>

      {/* Tabella mensile */}
      <div className="ia-panel" style={{ marginBottom: 18, overflowX: "auto" }}>
        <div style={{ padding: "12px 20px", borderBottom: "1px solid " + C.line }}>
          <div className="ia-h" style={{ fontSize: 14, color: C.navy }}>Dettaglio mensile {yearSel}</div>
          <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Clicca una riga per aprire il dettaglio del mese e correggere eventuali anomalie (fatture, spese, F24, movimenti banca).</div>
        </div>
        <table className="ia-table ia-keep">
          <thead><tr><th>Mese</th><th style={{ textAlign: "right" }}>Ricavi</th><th style={{ textAlign: "right" }}>Costi</th><th style={{ textAlign: "right" }}>Comp. ammin.</th><th style={{ textAlign: "right" }}>Contributi</th><th style={{ textAlign: "right" }}>Margine</th><th style={{ textAlign: "right" }}>%</th></tr></thead>
          <tbody>
            {monthly.map(x => {
              const hasData = x.revenues > 0 || x.otherCost > 0 || x.adminCost > 0 || x.contributi > 0;
              return (
                <tr key={x.m} onClick={() => hasData && setDrillMonth(x.m)} style={{ cursor: hasData ? "pointer" : "default" }} title={hasData ? "Clicca per vedere il dettaglio del mese e modificare eventuali anomalie" : undefined}>
                  <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{M_LABELS[x.monthIdx]}{hasData && <ChevronRight size={13} style={{ verticalAlign: "middle", marginLeft: 4, color: C.muted, opacity: .6 }} />}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", color: x.revenues > 0 ? C.green : C.muted }}>{x.revenues > 0 ? fmtMoney(x.revenues) : "—"}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", color: C.muted }}>{x.otherCost > 0 ? fmtMoney(x.otherCost) : "—"}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", color: C.muted }}>{x.adminCost > 0 ? fmtMoney(x.adminCost) : "—"}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", color: C.muted }}>{x.contributi > 0 ? fmtMoney(x.contributi) : "—"}</td>
                  <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: x.margin >= 0 ? C.green : C.red }}>{fmtMoney(x.margin)}</td>
                  <td style={{ textAlign: "right", fontSize: 12, color: C.muted }}>{x.revenues > 0 ? x.marginPct.toFixed(1) + "%" : "—"}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{ borderTop: "2px solid " + C.navy, background: "#FAFBFD" }}>
              <td style={{ padding: "10px 12px", fontFamily: "Poppins", fontWeight: 700 }}>TOTALE</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700, color: C.green }}>{fmtMoney(tot.revenues)}</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{fmtMoney(tot.otherCost)}</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{fmtMoney(tot.adminCost)}</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{fmtMoney(tot.contributi)}</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700, color: tot.margin >= 0 ? C.green : C.red, fontSize: 15 }}>{fmtMoney(tot.margin)}</td>
              <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 700 }}>{tot.marginPct.toFixed(1)}%</td>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* Top clienti / fornitori */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 18 }}>
        <div className="ia-panel">
          <PanelHead title="Top clienti per fatturato" />
          {topClients.length === 0 ? <div style={{ padding: 20, fontSize: 12.5, color: C.muted, textAlign: "center" }}>Nessun ricavo registrato.</div> : (
            <table className="ia-table"><tbody>{topClients.map(([c, v]) => (
              <tr key={c}>
                <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{c}</td>
                <td style={{ textAlign: "right", fontFamily: "Poppins" }}>{fmtMoney(v)}</td>
                <td style={{ textAlign: "right", fontSize: 12, color: C.muted, width: 60 }}>{((v / tot.revenues) * 100).toFixed(1)}%</td>
              </tr>
            ))}</tbody></table>
          )}
        </div>
        <div className="ia-panel">
          <PanelHead title="Top fornitori (escl. compensi ammin.)" />
          {topSuppliers.length === 0 ? <div style={{ padding: 20, fontSize: 12.5, color: C.muted, textAlign: "center" }}>Nessun costo registrato.</div> : (
            <table className="ia-table"><tbody>{topSuppliers.map(([s, v]) => (
              <tr key={s}>
                <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{s}</td>
                <td style={{ textAlign: "right", fontFamily: "Poppins" }}>{fmtMoney(v)}</td>
                <td style={{ textAlign: "right", fontSize: 12, color: C.muted, width: 60 }}>{tot.otherCost > 0 ? ((v / tot.otherCost) * 100).toFixed(1) + "%" : "—"}</td>
              </tr>
            ))}</tbody></table>
          )}
        </div>
      </div>

      {/* F24 breakdown */}
      <div className="ia-panel" style={{ marginBottom: 18 }}>
        <PanelHead title="Ripartizione scadenze F24" />
        <div style={{ padding: 16, display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
          <div style={{ padding: 12, border: "1px solid " + C.line, background: "#FBFCFE" }}>
            <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase" }}>Contributi INPS</div>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: C.red, marginTop: 4 }}>{fmtMoney(f24Breakdown.inps)}</div>
            <div style={{ fontSize: 10.5, color: C.muted, marginTop: 2 }}>Conteggiato come costo</div>
          </div>
          <div style={{ padding: 12, border: "1px solid " + C.line, background: "#FBFCFE" }}>
            <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase" }}>IVA</div>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: C.muted, marginTop: 4 }}>{fmtMoney(f24Breakdown.iva)}</div>
            <div style={{ fontSize: 10.5, color: C.muted, marginTop: 2 }}>Pass-through, non conteggiata</div>
          </div>
          <div style={{ padding: 12, border: "1px solid " + C.line, background: "#FBFCFE" }}>
            <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase" }}>Ritenute</div>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: C.muted, marginTop: 4 }}>{fmtMoney(f24Breakdown.ritenute)}</div>
            <div style={{ fontSize: 10.5, color: C.muted, marginTop: 2 }}>Pass-through, non conteggiate</div>
          </div>
          <div style={{ padding: 12, border: "1px solid " + C.line, background: "#FBFCFE" }}>
            <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase" }}>Altro F24 / Bolli</div>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: C.red, marginTop: 4 }}>{fmtMoney(f24Breakdown.altro)}</div>
            <div style={{ fontSize: 10.5, color: C.muted, marginTop: 2 }}>Conteggiato come costo</div>
          </div>
        </div>
      </div>

      {/* Spese per categoria — drill-down sui movimenti banca già categorizzati */}
      <SpesePerCategoria bank={bank} yearSel={yearSel} customCats={customBankCats} onOpenCategory={setDrillCategory} />

      {/* Regole di categorizzazione automatica memorizzate */}
      {bankRules.length > 0 && (
        <div className="ia-panel" style={{ marginBottom: 18 }}>
          <div style={{ padding: "12px 20px", borderBottom: "1px solid " + C.line, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <div className="ia-h" style={{ fontSize: 14, color: C.navy }}>Regole di categorizzazione automatica</div>
              <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Quando categorizzi manualmente un movimento, il sistema impara le parole distintive della descrizione e le riapplica ai movimenti simili futuri.</div>
            </div>
          </div>
          <div style={{ padding: "10px 20px", display: "flex", flexWrap: "wrap", gap: 8 }}>
            {bankRules.map(r => {
              const meta = bankCatMetaAll(r.category, customBankCats);
              return (
                <div key={r.id} style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "6px 10px", background: "#FAFBFD", border: "1px solid " + C.line, fontSize: 12 }}>
                  <span style={{ fontFamily: "Poppins", fontWeight: 600 }}>{(r.keywords || []).join(" + ")}</span>
                  <span style={{ color: C.muted }}>→</span>
                  <span style={{ color: C.orange, fontFamily: "Poppins", fontWeight: 600 }}>{meta?.label || r.category}</span>
                  <button onClick={() => deleteRule(r.id)} style={{ border: "none", background: "none", cursor: "pointer", padding: 2, marginLeft: 4, color: C.muted }} title="Elimina regola"><X size={13} /></button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Movimenti banca non riconciliati */}
      <div className="ia-panel">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 20px", borderBottom: "1px solid " + C.line, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div className="ia-h" style={{ fontSize: 14, color: C.navy }}>Movimenti banca — riconciliazione</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>Movimenti {yearSel} sopra <b>{fmtMoney(bankThreshold)}</b>. I candidati sono ordinati per confidenza: <b>Sicuro</b> = importo esatto + nome/numero riconosciuti in descrizione, <b>Probabile</b> = importo esatto e nome parziale, <b>Importo OK</b> = solo importo esatto, <b>Ipotesi</b> = ~5% e nome.</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, flexWrap: "wrap" }}>
            <label style={{ color: C.muted }}>Soglia €</label>
            <input className="ia-input" type="number" style={{ width: 90 }} value={bankThreshold} onChange={e => setBankThreshold(Math.max(0, Number(e.target.value) || 0))} />
            {bank.some(t => t.ignoredForPL && yearOf(t.date) === String(yearSel)) && (
              <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 10px", fontSize: 11.5 }} title="Rimetti tra i movimenti da chiarire quelli che avevi ignorato" onClick={() => {
                const n = bank.filter(t => t.ignoredForPL && yearOf(t.date) === String(yearSel)).length;
                setBank(bank.map(t => (t.ignoredForPL && yearOf(t.date) === String(yearSel)) ? { ...t, ignoredForPL: false } : t));
                notify(`${n} movimenti ripristinati`);
              }}>Ripristina ignorati ({bank.filter(t => t.ignoredForPL && yearOf(t.date) === String(yearSel)).length})</button>
            )}
            {bank.some(t => t.directCategory && yearOf(t.date) === String(yearSel)) && (
              <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 10px", fontSize: 11.5 }} title="Annulla la categorizzazione diretta e rimetti tra i movimenti da chiarire" onClick={() => {
                if (!confirm(`Annullare la categorizzazione diretta di ${bank.filter(t => t.directCategory && yearOf(t.date) === String(yearSel)).length} movimenti?`)) return;
                const n = bank.filter(t => t.directCategory && yearOf(t.date) === String(yearSel)).length;
                setBank(bank.map(t => (t.directCategory && yearOf(t.date) === String(yearSel)) ? { ...t, directCategory: null } : t));
                notify(`${n} movimenti scategorizzati`);
              }}>Annulla categorie ({bank.filter(t => t.directCategory && yearOf(t.date) === String(yearSel)).length})</button>
            )}
          </div>
        </div>

        {/* Toolbar filtri + auto-match */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 20px", borderBottom: "1px solid " + C.line, background: "#FAFBFD", flexWrap: "wrap", gap: 10 }}>
          <div style={{ display: "flex", gap: 4 }}>
            {[
              { id: "unmatched", label: "Da riconciliare", n: unmatchedBank.length },
              { id: "matched", label: "Abbinati", n: matchedBank.length },
              { id: "all", label: "Tutti", n: unmatchedBank.length + matchedBank.length },
            ].map(f => (
              <button key={f.id} onClick={() => setBankFilter(f.id)} style={{ padding: "6px 12px", border: "1px solid " + (bankFilter === f.id ? C.navy : C.line), background: bankFilter === f.id ? C.navy : "#fff", color: bankFilter === f.id ? "#fff" : C.ink, fontSize: 12, fontFamily: "Poppins", fontWeight: 600, cursor: "pointer" }}>{f.label} <span style={{ opacity: .7, fontWeight: 400 }}>({f.n})</span></button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 12px", fontSize: 12 }} onClick={cleanupBankDuplicatesPL} title="Rimuove movimenti bancari doppi (stesso giorno + stesso importo + stessa descrizione). Utile se lo stesso CSV bancario è stato importato più volte.">
              <RefreshCw size={13} /> Ripulisci duplicati banca
            </button>
            <button className="ia-btn ia-btn-danger" style={{ padding: "6px 12px", fontSize: 12 }} onClick={resetAllReconciliations} title="Azzera TUTTE le riconciliazioni: sblocca ogni fattura marcata come pagata e ogni bonifico abbinato. Anagrafiche e categorizzazioni dirette restano invariate. Utile se si sono accumulati match sbagliati da versioni precedenti del matcher.">
              <RefreshCw size={13} /> Reset riconciliazioni
            </button>
            {bankFilter !== "matched" && autoRulesCount > 0 && (
              <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 12px", fontSize: 12 }} onClick={autoApplyRules} title={`Applica le ${bankRules.length} regole di categorizzazione già memorizzate ai movimenti ancora non gestiti.`}>
                <RefreshCw size={13} /> Auto-categorizza da regole ({autoRulesCount})
              </button>
            )}
            {bankFilter !== "matched" && autoMatchCount > 0 && (
              <button className="ia-btn ia-btn-primary" style={{ padding: "6px 12px", fontSize: 12 }} onClick={autoMatchAll} title="Abbina automaticamente i movimenti con importo esatto e nome cliente/fornitore (o numero fattura) riconosciuti nella descrizione. In caso di più candidati con stesso importo e nome, vince quello col numero fattura in descrizione. Puoi sempre annullare i singoli abbinamenti nella scheda Abbinati.">
                <Check size={13} /> Auto-abbina fatture ({autoMatchCount})
              </button>
            )}
          </div>
        </div>

        {displayedBank.length === 0 ? (
          <div style={{ padding: 20, fontSize: 12.5, color: C.muted, textAlign: "center" }}>
            {bankFilter === "unmatched" ? "Nessun movimento sopra soglia non riconciliato." : bankFilter === "matched" ? "Nessun movimento abbinato ancora." : "Nessun movimento sopra soglia."}
          </div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="ia-table">
              <thead><tr><th>Data</th><th>Descrizione</th><th style={{ textAlign: "right" }}>Importo</th><th>Riconciliazione</th></tr></thead>
              <tbody>
                {displayedBank.slice(0, 30).map(t => {
                  // Riga già abbinata
                  if (t.matchedId) {
                    return (
                      <tr key={t.id} style={{ background: "#F5F9F5" }}>
                        <td style={{ whiteSpace: "nowrap" }}>{fmtDate(t.date)}</td>
                        <td style={{ fontSize: 12.5, maxWidth: 340 }}>{t.description}</td>
                        <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: t.amount >= 0 ? C.green : C.red }}>{fmtMoney(t.amount, t.currency)}</td>
                        <td>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                            <span className="ia-chip" style={{ background: "#E7F4EE", color: C.green, fontSize: 10.5 }}>abbinato</span>
                            <span style={{ fontSize: 12, color: C.ink, fontFamily: "Poppins", fontWeight: 600 }}>{t.matchedLabel}</span>
                            <button className="ia-btn ia-btn-ghost" style={{ padding: "3px 8px", fontSize: 11 }} onClick={() => unmatchBank(t)} title="Annulla l'abbinamento — il movimento torna tra quelli da chiarire e la fattura torna aperta">Annulla</button>
                          </div>
                        </td>
                      </tr>
                    );
                  }
                  // Riga non abbinata: mostra candidati
                  const sug = suggestFor(t);
                  return (
                    <tr key={t.id}>
                      <td style={{ whiteSpace: "nowrap" }}>{fmtDate(t.date)}</td>
                      <td style={{ fontSize: 12.5, maxWidth: 340 }}>{t.description}</td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: t.amount >= 0 ? C.green : C.red }}>{fmtMoney(t.amount, t.currency)}</td>
                      <td>
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 5, alignItems: "center" }}>
                          {sug.slice(0, 2).map(inv => {
                            const meta =
                              inv._conf === 4 ? { label: "Sicuro", bg: "#DFF3E7", col: C.green } :
                              inv._conf === 3 ? { label: "Probabile", bg: "#E7F4EE", col: C.green } :
                              inv._conf === 2 ? { label: "Importo OK", bg: "#EEF1F6", col: C.muted } :
                              { label: "Ipotesi", bg: "#FFF8E6", col: C.amber };
                            return (
                              <button key={inv.kind + inv.id} className="ia-btn ia-btn-dark" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => matchBank(t, inv)} title={`Abbinamento ${meta.label}${inv._numInDesc ? " · numero fattura in descrizione" : inv._nameHit ? " · nome trovato in descrizione" : ""}`}>
                                <span className="ia-chip" style={{ background: meta.bg, color: meta.col, fontSize: 9.5, padding: "1px 5px" }}>{meta.label}</span>
                                {inv.label}
                              </button>
                            );
                          })}
                          <button className="ia-btn ia-btn-primary" style={{ padding: "4px 10px", fontSize: 11.5 }} onClick={() => setReconcileTx(t)} title="Scegli manualmente la fattura o scadenza da abbinare">Riconcilia…</button>
                          <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => ignoreBank(t)} title="Non è un'operazione di ricavo/costo (giroconto, personale, ecc.)">Ignora</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {displayedBank.length > 30 && <tr><td colSpan={4} style={{ padding: 10, textAlign: "center", fontSize: 12, color: C.muted, fontStyle: "italic" }}>+ {displayedBank.length - 30} altri movimenti — abbassa la soglia o vai in Banca per la vista completa</td></tr>}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {reconcileTx && (
        <ReconcileModal
          tx={reconcileTx}
          openInvoices={openInvoices}
          combinedCats={combinedCats}
          customCats={customBankCats}
          onClose={() => setReconcileTx(null)}
          onMatch={(inv) => { matchBank(reconcileTx, inv); setReconcileTx(null); }}
          onCategorize={(catId) => {
            if (reconcileTx.directCategory) recategorizeBank(reconcileTx, catId);
            else categorizeBank(reconcileTx, catId);
            setReconcileTx(null);
          }}
          onAddCategory={(label, group) => addCustomCategory(label, group)}
          onDeleteCustomCategory={deleteCustomCategory}
        />
      )}
      {drillCategory && (
        <CategoryDrillModal
          categoryId={drillCategory}
          bank={bank}
          yearSel={yearSel}
          customCats={customBankCats}
          onClose={() => setDrillCategory(null)}
          onEdit={(tx) => { setDrillCategory(null); setReconcileTx(tx); }}
          onUncategorize={uncategorizeBank}
        />
      )}
      {drillMonth && (
        <MonthDrillModal
          month={drillMonth}
          issued={issued}
          passive={passive}
          dues={dues}
          bank={bank}
          contracts={contracts}
          customCats={customBankCats}
          netOfPassive={netOfPassive}
          isDueCost={isDueCost}
          onClose={() => setDrillMonth(null)}
          onEditIssued={(x) => setEditIssued(x)}
          onDeleteIssued={deleteIssuedRow}
          onEditPassive={(p) => setEditPassive(p)}
          onDeletePassive={deletePassiveRow}
          onEditDue={(d) => setEditDue(d)}
          onDeleteDue={deleteDueRow}
          onEditBank={(t) => setReconcileTx(t)}
          onUncategorizeBank={uncategorizeBank}
        />
      )}
      {editIssued && <NewIssuedModal x={editIssued} onSave={upsertIssued} onClose={() => setEditIssued(null)} />}
      {editPassive && <PassiveModal p={editPassive} onSave={upsertPassive} onClose={() => setEditPassive(null)} />}
      {editDue && <DueModal d={editDue} onSave={upsertDue} onClose={() => setEditDue(null)} />}
      {bulkCatUpdate && (
        <BulkCategoryUpdateModal
          payload={bulkCatUpdate}
          customBankCats={customBankCats}
          onClose={() => setBulkCatUpdate(null)}
          onConfirm={(selectedIds) => {
            const idSet = new Set(selectedIds);
            if (bulkCatUpdate.kind === "bank") {
              setBank(bank.map(t => idSet.has(t.id) ? { ...t, directCategory: bulkCatUpdate.newCategoryId, directCategoryFromRule: false } : t));
              notify(`${selectedIds.length} movimenti riclassificati`);
            } else if (bulkCatUpdate.kind === "passive") {
              setPassive(passive.map(p => idSet.has(p.id) ? { ...p, category: bulkCatUpdate.newCategoryId } : p));
              notify(`${selectedIds.length} fatture passive riclassificate`);
            }
            setBulkCatUpdate(null);
          }}
        />
      )}
    </div>
  );
}

// Modale per riconciliazione manuale: mostra tutte le fatture/scadenze aperte
// dello stesso segno (entrata vs uscita), con ricerca e ordinamento per prossimità di importo.
function ReconcileModal({ tx, openInvoices, combinedCats, customCats, onClose, onMatch, onCategorize, onAddCategory, onDeleteCustomCategory }) {
  const [mode, setMode] = useState("match"); // "match" | "categorize"
  const [q, setQ] = useState("");
  const [newCatOpen, setNewCatOpen] = useState(false);
  const [newCatLabel, setNewCatLabel] = useState("");
  const [newCatGroup, setNewCatGroup] = useState("cost");
  const wantSign = tx.amount >= 0 ? 1 : -1;
  const txAbs = Math.abs(tx.amount);
  const desc = String(tx.description || "").toLowerCase();
  const descNorm = normName(tx.description);
  const candidates = openInvoices
    .filter(inv => inv.sign === wantSign)
    .filter(inv => {
      if (!q.trim()) return true;
      const s = q.toLowerCase();
      return (inv.label || "").toLowerCase().includes(s) || (inv.sub || "").toLowerCase().includes(s);
    })
    .map(inv => {
      const invAbs = Math.abs(inv.amount);
      const amtDelta = Math.abs(invAbs - txAbs);
      const amountExact = amtDelta < 1.0;
      const amountLoose = invAbs > 0 && amtDelta / invAbs < 0.05;
      // Name matching
      const nameWords = normName(inv.label).split(/\s+/).filter(w => w.length >= 4);
      const nameHits = nameWords.filter(w => descNorm.includes(w)).length;
      const nameStrongHit = nameWords.length > 0 && (nameHits >= 2 || (nameHits === 1 && nameWords.length === 1));
      const nameWeakHit = nameHits >= 1;
      // Number matching — helper robusto: riconosce "RIF. 53", "FT 73", "N. 89"
      // e non confonde il numero fattura con la data
      const numInDesc = invoiceNumberInDesc(inv.number, tx.description);
      let score = 0;
      if (amountExact && (nameStrongHit || numInDesc)) score = 4;
      else if (amountExact && nameWeakHit) score = 3;
      else if (amountExact) score = 2;
      else if (amountLoose && (nameStrongHit || numInDesc)) score = 3;
      else if (amountLoose && nameWeakHit) score = 1;
      else score = 0;
      return { ...inv, _delta: amtDelta, _score: score, _nameHit: nameStrongHit || nameWeakHit, _numInDesc: numInDesc };
    })
    // Ordina per punteggio decrescente, poi per delta importo crescente
    .sort((a, b) => b._score - a._score || a._delta - b._delta);

  const kindLabel = {
    issued: { text: "Fattura emessa", color: C.green, bg: "#E7F4EE" },
    active: { text: "Fattura pianificata", color: C.orange, bg: "#FFF1E9" },
    passive: { text: "Fattura passiva", color: C.blue, bg: "#EBF2FB" },
    due: { text: "Scadenza F24", color: C.amber, bg: "#FFF8E6" },
  };

  const cats = combinedCats || BANK_DIRECT_CATEGORIES;
  const groupIds = { cost: [], admin: [], contributi: [], passthrough: [] };
  cats.forEach(c => { if (groupIds[c.group]) groupIds[c.group].push(c.id); });
  const catGroups = [
    { key: "cost", label: "Costi operativi", ids: groupIds.cost },
    { key: "admin", label: "Compensi", ids: groupIds.admin },
    { key: "contributi", label: "Contributi", ids: groupIds.contributi },
    { key: "passthrough", label: "Pass-through (non conteggiati)", ids: groupIds.passthrough },
  ];
  const groupColor = { cost: C.red, admin: C.orange, contributi: C.red, passthrough: C.muted };
  const isCustom = id => (customCats || []).some(c => c.id === id);

  const saveNewCat = () => {
    if (!newCatLabel.trim()) return;
    const id = onAddCategory && onAddCategory(newCatLabel.trim(), newCatGroup);
    if (id) onCategorize(id);
    else { setNewCatOpen(false); setNewCatLabel(""); }
  };

  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 680, maxWidth: "100%", maxHeight: "88vh", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
            <div>
              <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Riconcilia movimento</h3>
              <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>Abbina il movimento a una fattura/scadenza esistente, oppure registralo come spesa diretta (Telepass, addebito carta, biglietteria, ecc.) se non c'è una fattura corrispondente.</div>
            </div>
            <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
          </div>
          <div style={{ marginTop: 12, padding: 10, background: "#FAFBFD", border: "1px solid " + C.line, fontSize: 12.5 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontFamily: "Poppins", fontWeight: 600, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tx.description}</div>
                <div style={{ color: C.muted, fontSize: 11.5, marginTop: 2 }}>{fmtDate(tx.date)}</div>
              </div>
              <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: tx.amount >= 0 ? C.green : C.red, whiteSpace: "nowrap" }}>{fmtMoney(tx.amount, tx.currency)}</div>
            </div>
          </div>
          {/* Tabs */}
          <div style={{ display: "flex", gap: 4, marginTop: 12, borderBottom: "1px solid " + C.line, marginBottom: -1 }}>
            <button onClick={() => setMode("match")} style={{ padding: "8px 14px", border: "none", background: "none", cursor: "pointer", fontFamily: "Poppins", fontSize: 13, fontWeight: 600, color: mode === "match" ? C.navy : C.muted, borderBottom: mode === "match" ? "2px solid " + C.orange : "2px solid transparent", marginBottom: -1 }}>Abbina a voce esistente</button>
            {wantSign < 0 && (
              <button onClick={() => setMode("categorize")} style={{ padding: "8px 14px", border: "none", background: "none", cursor: "pointer", fontFamily: "Poppins", fontSize: 13, fontWeight: 600, color: mode === "categorize" ? C.navy : C.muted, borderBottom: mode === "categorize" ? "2px solid " + C.orange : "2px solid transparent", marginBottom: -1 }}>Registra come spesa diretta</button>
            )}
          </div>
        </div>

        {mode === "match" ? (
          <>
            <div style={{ padding: "12px 20px", borderBottom: "1px solid " + C.line }}>
              <input className="ia-input" autoFocus placeholder="Cerca cliente / fornitore / numero fattura…" value={q} onChange={e => setQ(e.target.value)} />
            </div>
            <div style={{ overflowY: "auto", flex: 1 }}>
              {candidates.length === 0 ? (
                <div style={{ padding: 30, textAlign: "center", color: C.muted, fontSize: 13 }}>
                  {openInvoices.filter(i => i.sign === wantSign).length === 0
                    ? "Nessuna voce aperta di questo segno da riconciliare."
                    : "Nessuna voce trovata con questo filtro."}
                  {wantSign < 0 && <div style={{ marginTop: 12 }}><button className="ia-btn ia-btn-primary" onClick={() => setMode("categorize")}>Registralo come spesa diretta →</button></div>}
                </div>
              ) : candidates.map(inv => {
                const meta = kindLabel[inv.kind] || kindLabel.issued;
                const perfect = inv._delta < 1;
                const close = inv._delta / (Math.abs(inv.amount) || 1) < 0.05;
                const confMeta =
                  inv._score === 4 ? { label: "Sicuro", bg: "#DFF3E7", col: C.green } :
                  inv._score === 3 ? { label: "Probabile", bg: "#E7F4EE", col: C.green } :
                  inv._score === 2 ? { label: "Importo OK", bg: "#EEF1F6", col: C.muted } :
                  inv._score === 1 ? { label: "Ipotesi", bg: "#FFF8E6", col: C.amber } : null;
                return (
                  <button key={inv.kind + ":" + inv.id} onClick={() => onMatch(inv)} style={{ display: "flex", width: "100%", padding: "12px 20px", border: "none", background: "transparent", borderBottom: "1px solid " + C.line, cursor: "pointer", textAlign: "left", alignItems: "center", gap: 12, font: "inherit" }} onMouseEnter={e => e.currentTarget.style.background = "#F5F7FB"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    <span className="ia-chip" style={{ background: meta.bg, color: meta.color, fontSize: 10, minWidth: 110, textAlign: "center" }}>{meta.text}</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 13, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                        {inv.label}
                        {confMeta && <span className="ia-chip" style={{ background: confMeta.bg, color: confMeta.col, fontSize: 9.5, marginLeft: 8, padding: "1px 6px" }}>{confMeta.label}</span>}
                        {inv._numInDesc && <span className="ia-chip" style={{ background: "#EBF2FB", color: C.blue, fontSize: 9.5, marginLeft: 4, padding: "1px 6px" }}>numero in descr.</span>}
                        {inv._nameHit && !inv._numInDesc && <span className="ia-chip" style={{ background: "#EBF2FB", color: C.blue, fontSize: 9.5, marginLeft: 4, padding: "1px 6px" }}>nome in descr.</span>}
                      </div>
                      <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>{inv.sub}</div>
                    </div>
                    <div style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                      <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 13, color: C.ink }}>{fmtMoney(inv.amount, inv.currency)}</div>
                      <div style={{ fontSize: 10.5, color: perfect ? C.green : close ? C.amber : C.muted, marginTop: 2, fontFamily: "Poppins", fontWeight: 600 }}>
                        {perfect ? "importo esatto" : inv._delta === 0 ? "" : `Δ ${fmtMoney(inv._delta)}`}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </>
        ) : (
          <div style={{ overflowY: "auto", flex: 1, padding: "16px 20px" }}>
            <div style={{ fontSize: 12.5, color: C.muted, marginBottom: 14, lineHeight: 1.5 }}>Scegli la categoria: il movimento sarà conteggiato nel Conto Economico in base al gruppo (costo operativo, compensi amministratori, contributi INPS). Le voci pass-through (IVA, ritenute) sono registrate ma <b>non</b> conteggiate come costo.</div>
            {catGroups.map(group => group.ids.length === 0 ? null : (
              <div key={group.key} style={{ marginBottom: 18 }}>
                <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px", marginBottom: 8 }}>{group.label}</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 8 }}>
                  {group.ids.map(id => {
                    const c = cats.find(x => x.id === id);
                    if (!c) return null;
                    const custom = isCustom(id);
                    return (
                      <div key={id} style={{ position: "relative", display: "flex" }}>
                        <button onClick={() => onCategorize(id)} className="ia-btn ia-btn-ghost" style={{ flex: 1, padding: "10px 12px", justifyContent: "flex-start", background: "#fff", border: "1px solid " + C.line, fontSize: 12.5, color: C.ink, textAlign: "left", paddingRight: custom ? 26 : 12 }} onMouseEnter={e => e.currentTarget.style.background = "#F5F7FB"} onMouseLeave={e => e.currentTarget.style.background = "#fff"}>
                          <span style={{ display: "inline-block", width: 6, height: 6, background: groupColor[c.group], borderRadius: 3, marginRight: 8 }} />
                          <span style={{ fontFamily: "Poppins", fontWeight: 600 }}>{c.label}</span>
                          {custom && <span style={{ marginLeft: 6, fontSize: 9.5, color: C.muted, fontStyle: "italic" }}>custom</span>}
                        </button>
                        {custom && onDeleteCustomCategory && (
                          <button onClick={(e) => { e.stopPropagation(); onDeleteCustomCategory(id); }} title="Elimina questa categoria personalizzata" style={{ position: "absolute", right: 4, top: "50%", transform: "translateY(-50%)", border: "none", background: "none", cursor: "pointer", color: C.muted, padding: 3 }}><X size={12} /></button>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}

            {/* Aggiungi nuova categoria */}
            <div style={{ marginTop: 8, padding: 12, background: "#FAFBFD", border: "1px dashed " + C.line }}>
              {!newCatOpen ? (
                <button className="ia-btn ia-btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => setNewCatOpen(true)}><Plus size={14} /> Nuova categoria personalizzata</button>
              ) : (
                <div>
                  <div style={{ fontSize: 12, color: C.muted, fontFamily: "Poppins", marginBottom: 8, fontWeight: 600 }}>Nuova categoria — salvata per il futuro</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 160px", gap: 8, marginBottom: 8 }}>
                    <input className="ia-input" autoFocus placeholder="Nome (es. Formazione, Interessi, Commissioni…)" value={newCatLabel} onChange={e => setNewCatLabel(e.target.value)} onKeyDown={e => e.key === "Enter" && saveNewCat()} />
                    <select className="ia-input" value={newCatGroup} onChange={e => setNewCatGroup(e.target.value)}>
                      <option value="cost">Costo operativo</option>
                      <option value="admin">Compensi amministratori</option>
                      <option value="contributi">Contributi</option>
                      <option value="passthrough">Pass-through (non conteggiato)</option>
                    </select>
                  </div>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                    <button className="ia-btn ia-btn-ghost" onClick={() => { setNewCatOpen(false); setNewCatLabel(""); }}>Annulla</button>
                    <button className="ia-btn ia-btn-primary" disabled={!newCatLabel.trim()} onClick={saveNewCat}><Check size={13} /> Salva e applica</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        <div style={{ padding: "12px 20px", borderTop: "1px solid " + C.line, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 11.5, color: C.muted }}>{mode === "match" ? `${candidates.length} voci disponibili` : "Le categorie personalizzate restano disponibili per i prossimi movimenti"}</div>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
        </div>
      </div>
    </div>
  );
}

// Pannello riepilogo "Spese per categoria" (movimenti banca già categorizzati)
function SpesePerCategoria({ bank, yearSel, customCats, onOpenCategory }) {
  const yearOf = d => (d || "").slice(0, 4);
  const items = bank.filter(t => t.directCategory && yearOf(t.date) === String(yearSel));
  if (items.length === 0) return null;
  // Aggrega per categoria
  const byCat = {};
  items.forEach(t => {
    const cid = t.directCategory;
    if (!byCat[cid]) byCat[cid] = { total: 0, count: 0 };
    byCat[cid].total += Math.abs(Number(t.amount) || 0);
    byCat[cid].count++;
  });
  const rows = Object.entries(byCat).map(([cid, v]) => ({ ...v, cid, meta: bankCatMetaAll(cid, customCats) })).sort((a, b) => b.total - a.total);
  const totalAll = rows.reduce((s, r) => s + r.total, 0);
  const groupColor = { cost: C.red, admin: C.orange, contributi: C.red, passthrough: C.muted };
  return (
    <div className="ia-panel" style={{ marginBottom: 18 }}>
      <div style={{ padding: "12px 20px", borderBottom: "1px solid " + C.line }}>
        <div className="ia-h" style={{ fontSize: 14, color: C.navy }}>Spese dirette da banca per categoria</div>
        <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>Clicca una categoria per vedere i movimenti dentro e modificare la classificazione di ognuno. Totale {yearSel}: <b style={{ color: C.navy }}>{fmtMoney(totalAll)}</b></div>
      </div>
      <div style={{ padding: 12, display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 8 }}>
        {rows.map(r => (
          <button key={r.cid} onClick={() => onOpenCategory(r.cid)} className="ia-btn ia-btn-ghost" style={{ padding: "12px 14px", background: "#FAFBFD", border: "1px solid " + C.line, flexDirection: "column", alignItems: "flex-start", textAlign: "left", height: "auto", gap: 4 }} onMouseEnter={e => e.currentTarget.style.background = "#F0F3F8"} onMouseLeave={e => e.currentTarget.style.background = "#FAFBFD"}>
            <div style={{ display: "flex", alignItems: "center", gap: 6, width: "100%" }}>
              <span style={{ width: 8, height: 8, background: groupColor[r.meta?.group] || C.muted, borderRadius: 4, display: "inline-block" }} />
              <span style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 13, color: C.ink, flex: 1 }}>{r.meta?.label || r.cid}</span>
              <ChevronRight size={14} color={C.muted} />
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", width: "100%", alignItems: "baseline" }}>
              <span style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 15, color: r.meta?.group === "passthrough" ? C.muted : C.navy }}>{fmtMoney(r.total)}</span>
              <span style={{ fontSize: 11, color: C.muted }}>{r.count} {r.count === 1 ? "voce" : "voci"}</span>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

// Modale bulk update categoria: elenca voci simili con checkbox individuali;
// applica in blocco la nuova categoria a quelle selezionate. Chiuso senza conferma =
// nessuna modifica.
function BulkCategoryUpdateModal({ payload, customBankCats, onConfirm, onClose }) {
  const [selected, setSelected] = useState(() => new Set(payload.items.map(i => i.id)));
  const toggle = (id) => {
    const next = new Set(selected);
    if (next.has(id)) next.delete(id); else next.add(id);
    setSelected(next);
  };
  const all = () => setSelected(new Set(payload.items.map(i => i.id)));
  const none = () => setSelected(new Set());
  const kindLabel = payload.kind === "passive" ? "fatture passive stesso fornitore" : "movimenti con descrizione simile";
  const newCatLabel = payload.kind === "bank"
    ? (bankCatMetaAll(payload.newCategoryId, customBankCats)?.label || payload.newCategoryId)
    : payload.newCategoryId;
  return (
    <div className="ia-modal-overlay" onClick={onClose}>
      <div className="ia-modal" style={{ maxWidth: 640 }} onClick={e => e.stopPropagation()}>
        <div className="ia-modal-header">
          <div>
            <div style={{ fontFamily: "Poppins", fontSize: 16, fontWeight: 600, color: C.navy }}>Applica anche a voci simili?</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
              Trovati <b>{payload.items.length} {kindLabel}</b> nell'anno corrente con categoria diversa da <b style={{ color: C.orange }}>{newCatLabel}</b>. Seleziona quelle a cui applicare la nuova categoria.
            </div>
          </div>
          <button className="ia-btn ia-btn-ghost" onClick={onClose} style={{ padding: 6 }}><X size={16} /></button>
        </div>
        <div style={{ padding: "8px 16px", display: "flex", gap: 8, borderBottom: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 10px", fontSize: 11 }} onClick={all}>Seleziona tutti</button>
          <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 10px", fontSize: 11 }} onClick={none}>Deseleziona tutti</button>
          <div style={{ marginLeft: "auto", fontSize: 12, color: C.muted, alignSelf: "center" }}>{selected.size} / {payload.items.length} selezionati</div>
        </div>
        <div style={{ maxHeight: 380, overflowY: "auto", padding: "4px 0" }}>
          {payload.items.map(item => (
            <label key={item.id} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "10px 16px", cursor: "pointer", borderBottom: "1px solid " + C.line }}>
              <input type="checkbox" checked={selected.has(item.id)} onChange={() => toggle(item.id)} style={{ marginTop: 3 }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, color: C.ink, fontWeight: 500, wordBreak: "break-word" }}>{item.primary}</div>
                <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>{item.secondary}</div>
              </div>
            </label>
          ))}
        </div>
        <div style={{ padding: 14, borderTop: "1px solid " + C.line, display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={selected.size === 0} onClick={() => onConfirm(Array.from(selected))}>
            Applica a {selected.size} {selected.size === 1 ? "voce" : "voci"}
          </button>
        </div>
      </div>
    </div>
  );
}

// Modale dettaglio di una categoria: elenco movimenti + modifica classificazione + rimozione categoria
function CategoryDrillModal({ categoryId, bank, yearSel, customCats, onClose, onEdit, onUncategorize }) {
  const meta = bankCatMetaAll(categoryId, customCats);
  const yearOf = d => (d || "").slice(0, 4);
  const items = bank.filter(t => t.directCategory === categoryId && yearOf(t.date) === String(yearSel))
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const total = items.reduce((s, t) => s + Math.abs(Number(t.amount) || 0), 0);
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 760, maxWidth: "100%", maxHeight: "88vh", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line, display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{meta?.label || categoryId}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>{items.length} {items.length === 1 ? "movimento" : "movimenti"} · totale <b style={{ color: C.navy, fontFamily: "Poppins" }}>{fmtMoney(total)}</b> nel {yearSel}. Clicca "Modifica categoria" su una riga per riclassificarla, oppure "Rimuovi" per rimetterla tra i movimenti da chiarire.</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ overflowY: "auto", flex: 1 }}>
          {items.length === 0 ? <div style={{ padding: 30, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessun movimento in questa categoria.</div> : (
            <table className="ia-table">
              <thead><tr><th>Data</th><th>Descrizione</th><th style={{ textAlign: "right" }}>Importo</th><th></th></tr></thead>
              <tbody>
                {items.map(t => (
                  <tr key={t.id}>
                    <td style={{ whiteSpace: "nowrap" }}>{fmtDate(t.date)}</td>
                    <td style={{ fontSize: 12.5, maxWidth: 360 }}>
                      {t.description}
                      {t.directCategoryFromRule && <div style={{ fontSize: 10.5, color: C.orange, marginTop: 2 }}>categorizzato automaticamente da una regola</div>}
                    </td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: t.amount >= 0 ? C.green : C.red, whiteSpace: "nowrap" }}>{fmtMoney(t.amount, t.currency)}</td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                      <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => onEdit(t)}>Modifica categoria</button>{" "}
                      <button className="ia-btn ia-btn-danger" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => onUncategorize(t)} title="Rimuovi la categoria — il movimento torna tra quelli da chiarire">Rimuovi</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
        <div style={{ padding: "12px 20px", borderTop: "1px solid " + C.line, display: "flex", justifyContent: "flex-end" }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Chiudi</button>
        </div>
      </div>
    </div>
  );
}

// Modale dettaglio di un mese: mostra ricavi, costi, compensi, contributi e movimenti banca,
// con azioni per modificare o eliminare ogni voce (correzione anomalie).
function MonthDrillModal({ month, issued, passive, dues, bank, contracts, customCats, netOfPassive, isDueCost, onClose, onEditIssued, onDeleteIssued, onEditPassive, onDeletePassive, onEditDue, onDeleteDue, onEditBank, onUncategorizeBank }) {
  const monthOf = d => (d || "").slice(0, 7);
  const iss = issued.filter(x => (x.month || monthOf(x.date)) === month).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const pas = passive.filter(p => monthOf(p.date) === month).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const du = dues.filter(d => monthOf(d.dueDate || d.paidAt) === month).sort((a, b) => (a.dueDate || "").localeCompare(b.dueDate || ""));
  const bnk = bank.filter(t => t.directCategory && monthOf(t.date) === month).sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  const revTotal = iss.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const admin = pas.filter(p => (p.category || "").toLowerCase().includes("amministrator"));
  const otherPas = pas.filter(p => !(p.category || "").toLowerCase().includes("amministrator"));
  const otherCostPas = otherPas.reduce((s, p) => s + netOfPassive(p), 0);
  const adminCostPas = admin.reduce((s, p) => s + netOfPassive(p), 0);
  const bnkByGroup = { cost: 0, admin: 0, contributi: 0, passthrough: 0 };
  bnk.forEach(t => { const g = bankCatMetaAll(t.directCategory, customCats)?.group || "cost"; bnkByGroup[g] += Math.abs(Number(t.amount) || 0); });
  const totalCosts = otherCostPas + bnkByGroup.cost;
  const totalAdmin = adminCostPas + bnkByGroup.admin;
  const totalContrib = du.filter(isDueCost).reduce((s, d) => s + Math.abs(Number(d.amount) || 0), 0) + bnkByGroup.contributi;
  const margin = revTotal - totalCosts - totalAdmin - totalContrib;

  const Section = ({ title, count, total, color, children }) => (
    <div style={{ marginBottom: 22 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, paddingBottom: 6, borderBottom: "1px solid " + C.line }}>
        <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 13, color: C.navy }}>{title} <span style={{ color: C.muted, fontWeight: 400, marginLeft: 4 }}>({count})</span></div>
        <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 14, color: color || C.navy }}>{fmtMoney(total)}</div>
      </div>
      {children}
    </div>
  );

  const rowActions = (onE, onD) => (
    <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
      <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={onE}><Pencil size={12} /> Modifica</button>
      <button className="ia-btn ia-btn-danger" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={onD} title="Elimina"><Trash2 size={12} /></button>
    </div>
  );

  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 880, maxWidth: "100%", maxHeight: "90vh", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid " + C.line, display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Dettaglio {monthLabel(month)}</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>Tutte le voci che compongono il conto economico del mese. Clicca "Modifica" per correggere anomalie o "Elimina" per rimuovere una voce sbagliata.</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>

        <div style={{ padding: "12px 20px", borderBottom: "1px solid " + C.line, background: "#FAFBFD", display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 12, fontSize: 12 }}>
          <div><div style={{ color: C.muted, fontSize: 10.5, textTransform: "uppercase", fontFamily: "Poppins", letterSpacing: ".4px" }}>Ricavi</div><b style={{ fontFamily: "Poppins", color: C.green, fontSize: 14 }}>{fmtMoney(revTotal)}</b></div>
          <div><div style={{ color: C.muted, fontSize: 10.5, textTransform: "uppercase", fontFamily: "Poppins", letterSpacing: ".4px" }}>Costi</div><b style={{ fontFamily: "Poppins", color: C.red, fontSize: 14 }}>{fmtMoney(totalCosts)}</b></div>
          <div><div style={{ color: C.muted, fontSize: 10.5, textTransform: "uppercase", fontFamily: "Poppins", letterSpacing: ".4px" }}>Comp. ammin.</div><b style={{ fontFamily: "Poppins", color: C.orange, fontSize: 14 }}>{fmtMoney(totalAdmin)}</b></div>
          <div><div style={{ color: C.muted, fontSize: 10.5, textTransform: "uppercase", fontFamily: "Poppins", letterSpacing: ".4px" }}>Contributi</div><b style={{ fontFamily: "Poppins", color: C.red, fontSize: 14 }}>{fmtMoney(totalContrib)}</b></div>
          <div><div style={{ color: C.muted, fontSize: 10.5, textTransform: "uppercase", fontFamily: "Poppins", letterSpacing: ".4px" }}>Margine</div><b style={{ fontFamily: "Poppins", color: margin >= 0 ? C.green : C.red, fontSize: 14 }}>{fmtMoney(margin)}</b></div>
        </div>

        <div style={{ overflowY: "auto", flex: 1, padding: "18px 20px" }}>
          {/* Ricavi */}
          <Section title="Ricavi — fatture emesse" count={iss.length} total={revTotal} color={C.green}>
            {iss.length === 0 ? <div style={{ fontSize: 12.5, color: C.muted, padding: 12, textAlign: "center", fontStyle: "italic" }}>Nessuna fattura emessa in questo mese.</div> : (
              <table className="ia-table" style={{ fontSize: 12.5 }}><tbody>
                {iss.map(x => {
                  const isCredit = !!x.isCreditNote || (Number(x.amount) || 0) < 0;
                  return (
                    <tr key={x.id} style={isCredit ? { background: "#FFF5F5" } : undefined}>
                      <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{x.client}{isCredit && <span className="ia-chip" style={{ background: "#FBECEC", color: C.red, marginLeft: 6, fontSize: 9.5 }}>nota di credito</span>}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{x.number || "—"} · {fmtDate(x.date)}</div></td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: isCredit ? C.red : C.ink, whiteSpace: "nowrap" }}>{fmtMoney(x.amount, x.currency)}</td>
                      <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>{rowActions(() => onEditIssued(x), () => onDeleteIssued(x.id))}</td>
                    </tr>
                  );
                })}
              </tbody></table>
            )}
          </Section>

          {/* Costi fornitori (esclusi compensi amministratori) */}
          <Section title="Costi — fornitori" count={otherPas.length} total={otherCostPas} color={C.red}>
            {otherPas.length === 0 ? <div style={{ fontSize: 12.5, color: C.muted, padding: 12, textAlign: "center", fontStyle: "italic" }}>Nessuna fattura passiva.</div> : (
              <table className="ia-table" style={{ fontSize: 12.5 }}><tbody>
                {otherPas.map(p => (
                  <tr key={p.id}>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{p.supplier}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{p.number || "—"} · {fmtDate(p.date)} · <i>{p.category || "senza categoria"}</i>{p.vatExempt ? " · senza IVA" : ""}</div></td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtMoney(p.amount, p.currency)}<div style={{ fontSize: 10.5, color: C.muted, fontWeight: 400 }}>netto {fmtMoney(netOfPassive(p))}</div></td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>{rowActions(() => onEditPassive(p), () => onDeletePassive(p.id))}</td>
                  </tr>
                ))}
              </tbody></table>
            )}
          </Section>

          {/* Compensi amministratori */}
          {admin.length > 0 && (
            <Section title="Compensi amministratori" count={admin.length} total={adminCostPas} color={C.orange}>
              <table className="ia-table" style={{ fontSize: 12.5 }}><tbody>
                {admin.map(p => (
                  <tr key={p.id}>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{p.supplier}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{p.number || "—"} · {fmtDate(p.date)}</div></td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, whiteSpace: "nowrap" }}>{fmtMoney(p.amount, p.currency)}</td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>{rowActions(() => onEditPassive(p), () => onDeletePassive(p.id))}</td>
                  </tr>
                ))}
              </tbody></table>
            </Section>
          )}

          {/* Scadenze F24 */}
          {du.length > 0 && (
            <Section title="Scadenze F24" count={du.length} total={du.reduce((s, d) => s + Math.abs(Number(d.amount) || 0), 0)} color={C.amber}>
              <table className="ia-table" style={{ fontSize: 12.5 }}><tbody>
                {du.map(d => {
                  const isCost = isDueCost(d);
                  return (
                    <tr key={d.id}>
                      <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{d.label}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{d.category || "—"} · {d.dueDate ? "scad. " + fmtDate(d.dueDate) : ""} · <i>{isCost ? "conteggiato come costo" : "pass-through"}</i></div></td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, whiteSpace: "nowrap", color: isCost ? C.red : C.muted }}>{fmtMoney(d.amount)}</td>
                      <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>{rowActions(() => onEditDue(d), () => onDeleteDue(d.id))}</td>
                    </tr>
                  );
                })}
              </tbody></table>
            </Section>
          )}

          {/* Movimenti banca categorizzati direttamente */}
          {bnk.length > 0 && (
            <Section title="Spese dirette da banca" count={bnk.length} total={bnk.reduce((s, t) => s + Math.abs(Number(t.amount) || 0), 0)} color={C.blue}>
              <table className="ia-table" style={{ fontSize: 12.5 }}><tbody>
                {bnk.map(t => {
                  const meta = bankCatMetaAll(t.directCategory, customCats);
                  const gCol = { cost: C.red, admin: C.orange, contributi: C.red, passthrough: C.muted }[meta?.group] || C.muted;
                  return (
                    <tr key={t.id}>
                      <td style={{ fontFamily: "Poppins", fontWeight: 600, maxWidth: 300 }}>{t.description}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{fmtDate(t.date)} · <span style={{ color: gCol }}>{meta?.label || t.directCategory}</span>{t.directCategoryFromRule && " · auto"}</div></td>
                      <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600, color: t.amount >= 0 ? C.green : C.red, whiteSpace: "nowrap" }}>{fmtMoney(t.amount, t.currency)}</td>
                      <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                        <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => onEditBank(t)}>Cambia categoria</button>{" "}
                        <button className="ia-btn ia-btn-danger" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => onUncategorizeBank(t)} title="Rimuovi categoria — il movimento torna tra quelli da chiarire">Rimuovi</button>
                      </td>
                    </tr>
                  );
                })}
              </tbody></table>
            </Section>
          )}
        </div>

        <div style={{ padding: "12px 20px", borderTop: "1px solid " + C.line, display: "flex", justifyContent: "flex-end" }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Chiudi</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// CRM — Prospect, Attività, Agenda, Task
// ============================================================
const ACT_TYPES = [
  { id: "call", label: "Chiamata", icon: Phone, color: "#3B7ED4" },
  { id: "email", label: "Email", icon: Mail, color: "#8891B5" },
  { id: "meeting", label: "Incontro", icon: MessageSquare, color: C.orange },
  { id: "note", label: "Nota", icon: StickyNote, color: C.muted },
];
const actMeta = id => ACT_TYPES.find(a => a.id === id) || ACT_TYPES[3];
const stageMeta = id => CRM_STAGES.find(s => s.id === id) || CRM_STAGES[0];
const teamMeta = id => TEAM.find(u => u.id === id) || null;
const fmtDateTime = iso => { if (!iso) return "—"; const d = new Date(iso); return d.toLocaleString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }); };

function CRM({ ctx }) {
  const [sub, setSub] = useState("dashboard");
  const [openProspectId, setOpenProspectId] = useState(null);
  const TABS = [
    { id: "dashboard", label: "Cruscotto", icon: LayoutDashboard },
    { id: "prospects", label: "Prospect", icon: Building2 },
    { id: "agenda", label: "Agenda", icon: Calendar },
    { id: "tasks", label: "Task", icon: ListChecks },
  ];
  const openProspect = id => { setOpenProspectId(id); setSub("prospects"); };
  return (
    <div>
      <div className="ia-tabs" style={{ display: "flex", gap: 4, marginBottom: 20, borderBottom: "1px solid " + C.line, paddingBottom: 0 }}>
        {TABS.map(t => (
          <button key={t.id} onClick={() => { setSub(t.id); if (t.id !== "prospects") setOpenProspectId(null); }} style={{ padding: "10px 18px", border: "none", background: "none", cursor: "pointer", fontFamily: "Poppins", fontSize: 13.5, fontWeight: 600, color: sub === t.id ? C.navy : C.muted, borderBottom: sub === t.id ? "2px solid " + C.orange : "2px solid transparent", marginBottom: -1, display: "inline-flex", alignItems: "center", gap: 7 }}>
            <t.icon size={15} /> {t.label}
          </button>
        ))}
      </div>
      {sub === "dashboard" && <CrmDashboard ctx={ctx} onOpenProspect={openProspect} go={setSub} />}
      {sub === "prospects" && <Prospects ctx={ctx} openId={openProspectId} onOpen={openProspect} onClose={() => setOpenProspectId(null)} />}
      {sub === "agenda" && <Agenda ctx={ctx} onOpenProspect={openProspect} />}
      {sub === "tasks" && <Tasks ctx={ctx} onOpenProspect={openProspect} />}
    </div>
  );
}

// ============ CRM DASHBOARD ============
function CrmDashboard({ ctx, onOpenProspect, go }) {
  const { prospects, appointments, tasks, crmActivities, user } = ctx;
  const T = today();
  const in7 = addDays(T, 7);
  const in30 = addDays(T, 30);
  const activeProspects = prospects.filter(p => p.stage === "prospect" || p.stage === "active");
  const pipelineValue = activeProspects.reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const wonYtd = prospects.filter(p => p.stage === "won" && (p.closedAt || "").slice(0, 4) === String(new Date().getFullYear())).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const upcomingAppts = [...appointments].filter(a => a.startAt && a.startAt.slice(0, 10) >= T && a.startAt.slice(0, 10) <= in7).sort((a, b) => a.startAt.localeCompare(b.startAt));
  const openTasks = tasks.filter(t => t.status !== "done");
  const myTasks = openTasks.filter(t => t.assignee === user.id).sort((a, b) => (a.dueDate || "9999").localeCompare(b.dueDate || "9999"));
  const overdueTasks = openTasks.filter(t => t.dueDate && t.dueDate < T);
  const recentActivities = [...crmActivities].sort((a, b) => (b.at || "").localeCompare(a.at || "")).slice(0, 8);

  const byStage = CRM_STAGES.map(s => ({
    ...s,
    items: prospects.filter(p => (p.stage || "prospect") === s.id),
    total: prospects.filter(p => (p.stage || "prospect") === s.id).reduce((sum, p) => sum + (Number(p.amount) || 0), 0),
  }));

  return (
    <div>
      {/* KPI CRM */}
      <div className="ia-kpis" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 16, marginBottom: 20 }}>
        <div className="ia-panel ia-bracket" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Prospect attivi</div>
          <div className="ia-kpi-num" style={{ color: C.navy, margin: "4px 0" }}>{activeProspects.length}</div>
          <div style={{ fontSize: 11.5, color: C.muted }}>Pipeline totale: <b style={{ color: C.orange }}>{fmtMoney(pipelineValue)}</b></div>
        </div>
        <div className="ia-panel" style={{ padding: 16 }}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Vinti YTD</div>
          <div className="ia-kpi-num" style={{ color: C.green, margin: "4px 0" }}>{fmtMoney(wonYtd)}</div>
          <div style={{ fontSize: 11.5, color: C.muted }}>{prospects.filter(p => p.stage === "won").length} deal chiusi in totale</div>
        </div>
        <div className="ia-panel" style={{ padding: 16, cursor: "pointer" }} onClick={() => go("agenda")}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Appuntamenti 7gg</div>
          <div className="ia-kpi-num" style={{ color: C.orange, margin: "4px 0" }}>{upcomingAppts.length}</div>
          <div style={{ fontSize: 11.5, color: C.muted }}>{appointments.length} totali in agenda</div>
        </div>
        <div className="ia-panel" style={{ padding: 16, cursor: "pointer" }} onClick={() => go("tasks")}>
          <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase", letterSpacing: ".5px" }}>Task aperti</div>
          <div className="ia-kpi-num" style={{ color: overdueTasks.length > 0 ? C.red : C.navy, margin: "4px 0" }}>{openTasks.length}</div>
          <div style={{ fontSize: 11.5, color: C.muted }}>{myTasks.length} assegnati a te{overdueTasks.length > 0 && <span style={{ color: C.red, fontWeight: 600 }}> · {overdueTasks.length} scaduti</span>}</div>
        </div>
      </div>

      {/* Pipeline Kanban */}
      <div className="ia-panel" style={{ marginBottom: 20 }}>
        <PanelHead title="Pipeline vendite" />
        <div style={{ padding: 16, display: "grid", gridTemplateColumns: `repeat(${CRM_STAGES.length}, 1fr)`, gap: 12 }}>
          {byStage.map(col => (
            <div key={col.id} style={{ background: "#FAFBFD", border: "1px solid " + C.line, padding: 10, minHeight: 220 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 8, borderBottom: "2px solid " + col.color }}>
                <span style={{ width: 8, height: 8, background: col.color, display: "inline-block", borderRadius: 4 }} />
                <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 12.5, color: C.ink, flex: 1 }}>{col.label}</div>
                <div style={{ fontSize: 11, color: C.muted }}>{col.items.length}</div>
              </div>
              <div style={{ fontSize: 10.5, color: C.muted, marginBottom: 8, fontFamily: "Poppins" }}>{fmtMoney(col.total)}</div>
              {col.items.length === 0 && <div style={{ fontSize: 11.5, color: C.muted, textAlign: "center", padding: "20px 0", fontStyle: "italic" }}>—</div>}
              {col.items.slice(0, 8).map(p => (
                <div key={p.id} onClick={() => onOpenProspect(p.id)} style={{ background: "#fff", border: "1px solid " + C.line, padding: "8px 10px", marginBottom: 6, cursor: "pointer" }}>
                  <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 12, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.company}</div>
                  <div style={{ fontSize: 10.5, color: C.muted, display: "flex", justifyContent: "space-between", marginTop: 2 }}>
                    <span>{teamMeta(p.owner)?.initials || "—"}</span>
                    <b style={{ color: col.color, fontFamily: "Poppins" }}>{p.amount ? shortEur(Number(p.amount)) : ""}</b>
                  </div>
                </div>
              ))}
              {col.items.length > 8 && <div style={{ fontSize: 10.5, color: C.muted, textAlign: "center", marginTop: 6 }}>+ {col.items.length - 8}</div>}
            </div>
          ))}
        </div>
      </div>

      {/* 3 colonne: prossimi appuntamenti, task tuoi, attività recenti */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", gap: 16 }}>
        <div className="ia-panel">
          <PanelHead title="Prossimi appuntamenti (7 giorni)" action={<button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => go("agenda")}>Agenda →</button>} />
          {upcomingAppts.length === 0 ? <div style={{ padding: 22, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessun appuntamento nei prossimi 7 giorni.</div> : (
            <div>
              {upcomingAppts.slice(0, 6).map(a => {
                const p = a.prospectId ? prospects.find(x => x.id === a.prospectId) : null;
                return (
                  <div key={a.id} style={{ padding: "10px 16px", borderTop: "1px solid " + C.line, cursor: p ? "pointer" : "default" }} onClick={() => p && onOpenProspect(p.id)}>
                    <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 12.5, color: C.ink }}>{a.title}</div>
                    <div style={{ fontSize: 11.5, color: C.orange, marginTop: 2 }}><Clock size={11} style={{ verticalAlign: "middle", marginRight: 4 }} />{fmtDateTime(a.startAt)}</div>
                    {p && <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>{p.company}</div>}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="ia-panel">
          <PanelHead title={"Task assegnati a te (" + (user.name.split(" ")[0]) + ")"} action={<button className="ia-btn ia-btn-ghost" style={{ padding: "4px 8px", fontSize: 11.5 }} onClick={() => go("tasks")}>Tutti →</button>} />
          {myTasks.length === 0 ? <div style={{ padding: 22, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessun task assegnato a te.</div> : (
            <div>
              {myTasks.slice(0, 6).map(t => {
                const p = t.prospectId ? prospects.find(x => x.id === t.prospectId) : null;
                const late = t.dueDate && t.dueDate < T;
                return (
                  <div key={t.id} style={{ padding: "10px 16px", borderTop: "1px solid " + C.line, cursor: p ? "pointer" : "default" }} onClick={() => p && onOpenProspect(p.id)}>
                    <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 12.5, color: C.ink }}>{t.title}</div>
                    <div style={{ fontSize: 11.5, color: late ? C.red : C.muted, marginTop: 2 }}>{t.dueDate ? "scade " + fmtDate(t.dueDate) : "senza scadenza"}{p && " · " + p.company}</div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="ia-panel">
          <PanelHead title="Attività recenti" />
          {recentActivities.length === 0 ? <div style={{ padding: 22, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessuna attività ancora registrata.</div> : (
            <div>
              {recentActivities.map(a => {
                const p = prospects.find(x => x.id === a.prospectId);
                const meta = actMeta(a.type);
                const who = teamMeta(a.by);
                return (
                  <div key={a.id} style={{ padding: "10px 16px", borderTop: "1px solid " + C.line, cursor: p ? "pointer" : "default", display: "flex", gap: 10 }} onClick={() => p && onOpenProspect(p.id)}>
                    <meta.icon size={16} color={meta.color} style={{ marginTop: 2, flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontFamily: "Poppins", fontSize: 12, fontWeight: 600, color: C.ink }}>{p?.company || "—"} <span style={{ color: C.muted, fontWeight: 400 }}>· {meta.label.toLowerCase()}</span></div>
                      <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.notes}</div>
                      <div style={{ fontSize: 10.5, color: C.muted, marginTop: 2 }}>{fmtDateTime(a.at)}{who && " · " + who.initials}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ============ PROSPECTS ============
function Prospects({ ctx, openId, onOpen, onClose }) {
  const { prospects, setProspects, notify } = ctx;
  const [edit, setEdit] = useState(null);
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const openProspect = openId ? prospects.find(p => p.id === openId) : null;

  const list = prospects.filter(p => {
    if (filter !== "all" && p.stage !== filter) return false;
    if (q) {
      const s = q.toLowerCase();
      return (p.company || "").toLowerCase().includes(s) || (p.sector || "").toLowerCase().includes(s) || (p.contacts || []).some(c => (c.name || "").toLowerCase().includes(s));
    }
    return true;
  }).sort((a, b) => (b.updatedAt || "").localeCompare(a.updatedAt || ""));

  const upsert = (p) => {
    const exists = prospects.some(x => x.id === p.id);
    const next = { ...p, updatedAt: new Date().toISOString(), closedAt: (p.stage === "won" || p.stage === "lost") ? (p.closedAt || today()) : null };
    setProspects(exists ? prospects.map(x => x.id === p.id ? next : x) : [next, ...prospects]);
    setEdit(null); notify(exists ? "Prospect aggiornato" : "Prospect creato");
  };
  const del = (id) => {
    if (!confirm("Eliminare questo prospect e tutte le sue attività collegate?")) return;
    setProspects(prospects.filter(p => p.id !== id));
    onClose();
  };

  if (openProspect) return <ProspectDetail ctx={ctx} p={openProspect} onClose={onClose} onEdit={() => setEdit(openProspect)} onDelete={() => del(openProspect.id)} edit={edit} setEdit={setEdit} upsert={upsert} />;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, gap: 12, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <input className="ia-input" placeholder="Cerca azienda, settore, contatto…" value={q} onChange={e => setQ(e.target.value)} style={{ width: 260 }} />
          <div style={{ display: "flex", gap: 4 }}>
            <button className={"ia-btn " + (filter === "all" ? "ia-btn-dark" : "ia-btn-ghost")} style={{ padding: "6px 10px", fontSize: 12 }} onClick={() => setFilter("all")}>Tutti ({prospects.length})</button>
            {CRM_STAGES.map(s => (
              <button key={s.id} className={"ia-btn " + (filter === s.id ? "ia-btn-dark" : "ia-btn-ghost")} style={{ padding: "6px 10px", fontSize: 12 }} onClick={() => setFilter(s.id)}>{s.label} ({prospects.filter(p => p.stage === s.id).length})</button>
            ))}
          </div>
        </div>
        <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), company: "", sector: "", website: "", geo: "", stage: "prospect", amount: "", owner: ctx.user.id, notes: "", contacts: [] })}><Plus size={16} /> Nuovo prospect</button>
      </div>
      {list.length === 0 ? <Empty icon={Building2} title="Nessun prospect trovato" hint={q || filter !== "all" ? "Prova a cambiare i filtri." : "Aggiungi il primo prospect per iniziare a tracciare le opportunità."} /> : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th>Azienda</th><th>Settore</th><th>Contatti</th><th>Owner</th><th>Stage</th><th style={{ textAlign: "right" }}>Valore</th><th>Aggiornato</th></tr></thead>
            <tbody>
              {list.map(p => {
                const st = stageMeta(p.stage);
                const own = teamMeta(p.owner);
                return (
                  <tr key={p.id} style={{ cursor: "pointer" }} onClick={() => onOpen(p.id)}>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{p.company}<div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>{p.website}</div></td>
                    <td style={{ fontSize: 12.5 }}>{p.sector || "—"}<div style={{ fontSize: 11, color: C.muted }}>{p.geo || ""}</div></td>
                    <td style={{ fontSize: 12.5 }}>{p.contacts?.length || 0}</td>
                    <td style={{ fontSize: 12.5 }}>{own?.initials || "—"}</td>
                    <td><span className="ia-chip" style={{ background: st.color + "20", color: st.color, fontFamily: "Poppins", fontWeight: 600 }}>{st.label}</span></td>
                    <td style={{ textAlign: "right", fontFamily: "Poppins", fontWeight: 600 }}>{p.amount ? fmtMoney(p.amount) : "—"}</td>
                    <td style={{ fontSize: 11.5, color: C.muted, whiteSpace: "nowrap" }}>{p.updatedAt ? fmtDate(p.updatedAt.slice(0, 10)) : "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {edit && !openProspect && <ProspectModal p={edit} onSave={upsert} onClose={() => setEdit(null)} />}
    </div>
  );
}

function ProspectDetail({ ctx, p, onClose, onEdit, onDelete, edit, setEdit, upsert }) {
  const { crmActivities, setCrmActivities, appointments, tasks, notify, user } = ctx;
  const [activityOpen, setActivityOpen] = useState(null);
  const acts = crmActivities.filter(a => a.prospectId === p.id).sort((a, b) => (b.at || "").localeCompare(a.at || ""));
  const appts = appointments.filter(a => a.prospectId === p.id).sort((a, b) => (b.startAt || "").localeCompare(a.startAt || ""));
  const tks = tasks.filter(t => t.prospectId === p.id).sort((a, b) => (a.status === "done" ? 1 : 0) - (b.status === "done" ? 1 : 0) || (a.dueDate || "").localeCompare(b.dueDate || ""));
  const st = stageMeta(p.stage);
  const own = teamMeta(p.owner);
  const saveActivity = (a) => {
    const exists = crmActivities.some(x => x.id === a.id);
    setCrmActivities(exists ? crmActivities.map(x => x.id === a.id ? a : x) : [a, ...crmActivities]);
    setActivityOpen(null); notify(exists ? "Attività aggiornata" : "Attività registrata");
  };
  const delActivity = (id) => setCrmActivities(crmActivities.filter(a => a.id !== id));

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <button className="ia-btn ia-btn-ghost" style={{ padding: "6px 10px" }} onClick={onClose}>‹ Elenco</button>
        <div style={{ flex: 1 }} />
        <button className="ia-btn ia-btn-ghost" onClick={onEdit}><Pencil size={14} /> Modifica</button>
        <button className="ia-btn ia-btn-danger" style={{ padding: "6px 10px" }} onClick={onDelete}><Trash2 size={14} /></button>
      </div>
      <div className="ia-panel ia-bracket" style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 14, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 260 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
              <span className="ia-h" style={{ fontSize: 20, color: C.navy }}>{p.company}</span>
              <span className="ia-chip" style={{ background: st.color + "20", color: st.color, fontFamily: "Poppins", fontWeight: 600 }}>{st.label}</span>
            </div>
            <div style={{ fontSize: 12.5, color: C.muted }}>{[p.sector, p.geo, p.website].filter(Boolean).join(" · ") || "—"}</div>
            {p.notes && <div style={{ marginTop: 10, fontSize: 13, color: C.ink, lineHeight: 1.5 }}>{p.notes}</div>}
          </div>
          <div style={{ minWidth: 180 }}>
            <div style={{ fontSize: 11, color: C.muted, fontFamily: "Poppins", textTransform: "uppercase" }}>Valore potenziale</div>
            <div className="ia-kpi-num" style={{ color: C.orange, fontSize: 22, margin: "4px 0" }}>{p.amount ? fmtMoney(p.amount) : "—"}</div>
            <div style={{ fontSize: 12, color: C.muted }}>Owner: <b style={{ color: C.ink }}>{own?.name || "—"}</b></div>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 340px", gap: 16 }}>
        {/* Timeline */}
        <div className="ia-panel">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 18px", borderBottom: "1px solid " + C.line }}>
            <div className="ia-h" style={{ fontSize: 14, color: C.navy }}>Timeline attività</div>
            <button className="ia-btn ia-btn-primary" style={{ padding: "6px 12px", fontSize: 12.5 }} onClick={() => setActivityOpen({ id: uid(), prospectId: p.id, type: "call", at: new Date().toISOString().slice(0, 16), by: user.id, notes: "" })}><Plus size={13} /> Nuova attività</button>
          </div>
          {acts.length === 0 ? <div style={{ padding: 24, textAlign: "center", color: C.muted, fontSize: 13 }}>Nessuna attività registrata. Aggiungi la prima chiamata, email o nota.</div> : (
            <div style={{ padding: "6px 0" }}>
              {acts.map(a => {
                const meta = actMeta(a.type);
                const who = teamMeta(a.by);
                return (
                  <div key={a.id} style={{ padding: "12px 18px", borderTop: "1px solid " + C.line, display: "flex", gap: 12 }}>
                    <meta.icon size={18} color={meta.color} style={{ marginTop: 2, flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                        <div style={{ fontFamily: "Poppins", fontSize: 12.5, fontWeight: 600, color: C.ink }}>{meta.label} <span style={{ color: C.muted, fontWeight: 400 }}>· {who?.initials || "—"} · {fmtDateTime(a.at)}</span></div>
                        <div style={{ display: "flex", gap: 4 }}>
                          <button className="ia-btn ia-btn-ghost" style={{ padding: "3px 6px" }} onClick={() => setActivityOpen(a)}><Pencil size={11} /></button>
                          <button className="ia-btn ia-btn-danger" style={{ padding: "3px 6px" }} onClick={() => delActivity(a.id)}><Trash2 size={11} /></button>
                        </div>
                      </div>
                      <div style={{ fontSize: 12.5, color: C.ink, whiteSpace: "pre-wrap", lineHeight: 1.4 }}>{a.notes}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Sidebar: contatti, appuntamenti, task */}
        <div>
          <div className="ia-panel" style={{ marginBottom: 14 }}>
            <PanelHead title={`Contatti (${p.contacts?.length || 0})`} />
            {(p.contacts || []).length === 0 ? <div style={{ padding: 16, fontSize: 12.5, color: C.muted, textAlign: "center" }}>Aggiungi contatti dalla modifica.</div> : (
              <div>
                {(p.contacts || []).map((c, i) => (
                  <div key={i} style={{ padding: "10px 14px", borderTop: "1px solid " + C.line }}>
                    <div style={{ fontFamily: "Poppins", fontWeight: 600, fontSize: 12.5 }}>{c.name}</div>
                    <div style={{ fontSize: 11.5, color: C.muted }}>{c.role}</div>
                    {c.email && <div style={{ fontSize: 11.5, marginTop: 2 }}><a href={"mailto:" + c.email} style={{ color: C.orange }}>{c.email}</a></div>}
                    {c.phone && <div style={{ fontSize: 11.5, color: C.muted, marginTop: 2 }}>{c.phone}</div>}
                  </div>
                ))}
              </div>
            )}
          </div>
          <div className="ia-panel" style={{ marginBottom: 14 }}>
            <PanelHead title={`Appuntamenti (${appts.length})`} />
            {appts.length === 0 ? <div style={{ padding: 16, fontSize: 12.5, color: C.muted, textAlign: "center" }}>Nessun appuntamento.</div> : appts.slice(0, 5).map(a => (
              <div key={a.id} style={{ padding: "10px 14px", borderTop: "1px solid " + C.line }}>
                <div style={{ fontFamily: "Poppins", fontSize: 12, fontWeight: 600 }}>{a.title}</div>
                <div style={{ fontSize: 11.5, color: a.startAt.slice(0, 10) >= today() ? C.orange : C.muted }}>{fmtDateTime(a.startAt)}</div>
              </div>
            ))}
          </div>
          <div className="ia-panel">
            <PanelHead title={`Task (${tks.length})`} />
            {tks.length === 0 ? <div style={{ padding: 16, fontSize: 12.5, color: C.muted, textAlign: "center" }}>Nessun task.</div> : tks.slice(0, 6).map(t => (
              <div key={t.id} style={{ padding: "10px 14px", borderTop: "1px solid " + C.line, display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ flex: 1, textDecoration: t.status === "done" ? "line-through" : "none", opacity: t.status === "done" ? .5 : 1 }}>
                  <div style={{ fontFamily: "Poppins", fontSize: 12, fontWeight: 600 }}>{t.title}</div>
                  <div style={{ fontSize: 11.5, color: t.dueDate && t.dueDate < today() && t.status !== "done" ? C.red : C.muted }}>{t.dueDate ? fmtDate(t.dueDate) : "senza data"} · {teamMeta(t.assignee)?.initials || "—"}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {edit && <ProspectModal p={edit} onSave={upsert} onClose={() => setEdit(null)} />}
      {activityOpen && <ActivityModal a={activityOpen} onSave={saveActivity} onClose={() => setActivityOpen(null)} />}
    </div>
  );
}

function ProspectModal({ p, onSave, onClose }) {
  const [f, setF] = useState({ ...p, contacts: p.contacts || [] });
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const addContact = () => set("contacts", [...(f.contacts || []), { id: uid(), name: "", role: "", email: "", phone: "" }]);
  const updContact = (id, patch) => set("contacts", f.contacts.map(c => c.id === id ? { ...c, ...patch } : c));
  const rmContact = (id) => set("contacts", f.contacts.filter(c => c.id !== id));
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 620, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>{p.company ? "Modifica prospect" : "Nuovo prospect"}</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Azienda *"><input className="ia-input" value={f.company} onChange={e => set("company", e.target.value)} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 12 }}>
            <F label="Settore"><input className="ia-input" placeholder="es. Meccanica di precisione" value={f.sector} onChange={e => set("sector", e.target.value)} /></F>
            <F label="Area geografica"><input className="ia-input" placeholder="es. USA, DACH" value={f.geo} onChange={e => set("geo", e.target.value)} /></F>
          </div>
          <F label="Sito web"><input className="ia-input" value={f.website} onChange={e => set("website", e.target.value)} placeholder="https://…" /></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            <F label="Stage"><select className="ia-input" value={f.stage} onChange={e => set("stage", e.target.value)}>{CRM_STAGES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}</select></F>
            <F label="Valore potenziale (€)"><input className="ia-input" type="number" value={f.amount} onChange={e => set("amount", e.target.value)} /></F>
            <F label="Owner"><select className="ia-input" value={f.owner} onChange={e => set("owner", e.target.value)}>{TEAM.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}</select></F>
          </div>
          <F label="Note"><textarea className="ia-input" rows={3} value={f.notes} onChange={e => set("notes", e.target.value)} style={{ fontFamily: "inherit", resize: "vertical" }} /></F>
          <div style={{ marginTop: 6, marginBottom: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted }}>Contatti in azienda</label>
              <button className="ia-btn ia-btn-ghost" style={{ padding: "4px 10px", fontSize: 12 }} onClick={addContact}><Plus size={13} /> Aggiungi contatto</button>
            </div>
            {(f.contacts || []).map(c => (
              <div key={c.id} style={{ border: "1px solid " + C.line, padding: 10, marginBottom: 8, background: "#FBFCFE" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 8, marginBottom: 6 }}>
                  <input className="ia-input" placeholder="Nome cognome" value={c.name} onChange={e => updContact(c.id, { name: e.target.value })} />
                  <input className="ia-input" placeholder="Ruolo (es. Export Manager)" value={c.role} onChange={e => updContact(c.id, { role: e.target.value })} />
                  <button className="ia-btn ia-btn-ghost" style={{ padding: "0 10px" }} onClick={() => rmContact(c.id)}><X size={13} /></button>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <input className="ia-input" placeholder="Email" value={c.email} onChange={e => updContact(c.id, { email: e.target.value })} />
                  <input className="ia-input" placeholder="Telefono" value={c.phone} onChange={e => updContact(c.id, { phone: e.target.value })} />
                </div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.company} onClick={() => onSave({ ...f, amount: f.amount === "" ? null : Number(f.amount) })}>Salva</button>
        </div>
      </div>
    </div>
  );
}

function ActivityModal({ a, onSave, onClose }) {
  const [f, setF] = useState(a);
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 460, maxWidth: "100%" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Attività</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Tipo"><select className="ia-input" value={f.type} onChange={e => set("type", e.target.value)}>{ACT_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}</select></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Data e ora"><input className="ia-input" type="datetime-local" value={(f.at || "").slice(0, 16)} onChange={e => set("at", e.target.value)} /></F>
            <F label="Chi ha fatto"><select className="ia-input" value={f.by} onChange={e => set("by", e.target.value)}>{TEAM.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}</select></F>
          </div>
          <F label="Note / esito"><textarea className="ia-input" rows={5} value={f.notes} onChange={e => set("notes", e.target.value)} placeholder="Descrivi cosa è emerso, prossimi passi, ecc." style={{ fontFamily: "inherit", resize: "vertical" }} /></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" onClick={() => onSave({ ...f, at: f.at ? new Date(f.at).toISOString() : new Date().toISOString() })}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ============ AGENDA (appuntamenti) ============
function Agenda({ ctx, onOpenProspect }) {
  const { appointments, setAppointments, prospects, notify, user } = ctx;
  const [edit, setEdit] = useState(null);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [filter, setFilter] = useState("upcoming");
  const T = today();
  const list = [...appointments].sort((a, b) => (a.startAt || "").localeCompare(b.startAt || ""));
  const filtered = list.filter(a => {
    if (filter === "upcoming") return a.startAt && a.startAt.slice(0, 10) >= T;
    if (filter === "past") return a.startAt && a.startAt.slice(0, 10) < T;
    return true;
  });
  const upsert = (a) => {
    const exists = appointments.some(x => x.id === a.id);
    setAppointments(exists ? appointments.map(x => x.id === a.id ? a : x) : [...appointments, a]);
    setEdit(null); notify(exists ? "Appuntamento aggiornato" : "Appuntamento creato");
  };
  const del = (id) => setAppointments(appointments.filter(a => a.id !== id));

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, gap: 10, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 4 }}>
          {[{ id: "upcoming", label: "Prossimi (" + filtered.filter(a => a.startAt.slice(0, 10) >= T).length + ")" }, { id: "past", label: "Passati" }, { id: "all", label: "Tutti (" + appointments.length + ")" }].map(f => (
            <button key={f.id} className={"ia-btn " + (filter === f.id ? "ia-btn-dark" : "ia-btn-ghost")} style={{ padding: "6px 12px", fontSize: 12.5 }} onClick={() => setFilter(f.id)}>{f.label}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="ia-btn ia-btn-ghost" onClick={() => setPasteOpen(true)} title="Incolla il testo di un invito .ics o di una email di conferma meeting"><Mail size={14} /> Incolla invito</button>
          <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), title: "", startAt: new Date().toISOString().slice(0, 16), endAt: "", location: "", attendees: "", prospectId: "", notes: "", createdBy: user.id })}><Plus size={16} /> Nuovo appuntamento</button>
        </div>
      </div>
      {filtered.length === 0 ? <Empty icon={Calendar} title="Nessun appuntamento in questo filtro" hint="Crea un appuntamento manuale o incolla un invito." /> : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th>Quando</th><th>Titolo</th><th>Prospect</th><th>Luogo</th><th>Partecipanti</th><th>Creato da</th><th></th></tr></thead>
            <tbody>
              {filtered.map(a => {
                const p = a.prospectId ? prospects.find(x => x.id === a.prospectId) : null;
                const who = teamMeta(a.createdBy);
                const isFuture = a.startAt && a.startAt.slice(0, 10) >= T;
                return (
                  <tr key={a.id}>
                    <td style={{ whiteSpace: "nowrap", fontSize: 12.5, color: isFuture ? C.orange : C.muted, fontFamily: "Poppins", fontWeight: 600 }}>{fmtDateTime(a.startAt)}{a.endAt && <div style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>fino a {fmtDateTime(a.endAt).split(" ")[1]}</div>}</td>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600 }}>{a.title}<div style={{ fontWeight: 400, fontSize: 11.5, color: C.muted }}>{a.notes}</div></td>
                    <td>{p ? <button style={{ border: "none", background: "none", color: C.orange, cursor: "pointer", padding: 0, fontFamily: "Poppins", fontSize: 12.5, textDecoration: "underline" }} onClick={() => onOpenProspect(p.id)}>{p.company}</button> : "—"}</td>
                    <td style={{ fontSize: 12.5 }}>{a.location || "—"}</td>
                    <td style={{ fontSize: 12, color: C.muted }}>{a.attendees || "—"}</td>
                    <td style={{ fontSize: 12.5 }}>{who?.initials || "—"}</td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                      <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => setEdit(a)}><Pencil size={13} /></button>{" "}
                      <button className="ia-btn ia-btn-danger" style={{ padding: "5px 8px" }} onClick={() => del(a.id)}><Trash2 size={13} /></button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {edit && <AppointmentModal a={edit} prospects={prospects} onSave={upsert} onClose={() => setEdit(null)} />}
      {pasteOpen && <PasteInviteModal onParsed={(a) => { setPasteOpen(false); setEdit({ ...a, id: uid(), createdBy: user.id, prospectId: "" }); }} onClose={() => setPasteOpen(false)} />}
    </div>
  );
}

function AppointmentModal({ a, prospects, onSave, onClose }) {
  const [f, setF] = useState(a);
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  const startISO = (f.startAt || "").slice(0, 16);
  const endISO = (f.endAt || "").slice(0, 16);
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 500, maxWidth: "100%", maxHeight: "90vh", overflow: "auto" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Appuntamento</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Titolo *"><input className="ia-input" value={f.title} onChange={e => set("title", e.target.value)} placeholder="es. Prima call conoscitiva" /></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Inizio *"><input className="ia-input" type="datetime-local" value={startISO} onChange={e => set("startAt", e.target.value)} /></F>
            <F label="Fine"><input className="ia-input" type="datetime-local" value={endISO} onChange={e => set("endAt", e.target.value)} /></F>
          </div>
          <F label="Prospect collegato"><select className="ia-input" value={f.prospectId || ""} onChange={e => set("prospectId", e.target.value)}><option value="">— nessuno —</option>{prospects.map(p => <option key={p.id} value={p.id}>{p.company}</option>)}</select></F>
          <F label="Luogo / link"><input className="ia-input" value={f.location} onChange={e => set("location", e.target.value)} placeholder="Milano / Zoom / …" /></F>
          <F label="Partecipanti"><input className="ia-input" value={f.attendees} onChange={e => set("attendees", e.target.value)} placeholder="es. Mario Rossi (Vulcaflex), Marco, Guglielmo" /></F>
          <F label="Note"><textarea className="ia-input" rows={3} value={f.notes} onChange={e => set("notes", e.target.value)} style={{ fontFamily: "inherit", resize: "vertical" }} /></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.title || !f.startAt} onClick={() => onSave({ ...f, startAt: f.startAt ? new Date(f.startAt).toISOString() : null, endAt: f.endAt ? new Date(f.endAt).toISOString() : null })}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// Parser semplice per invito .ics o testo generico
function parseInvite(text) {
  const t = String(text || "");
  const out = { title: "", startAt: "", endAt: "", location: "", attendees: "", notes: "" };
  // .ics tags (SUMMARY, LOCATION, DTSTART, DTEND, ATTENDEE)
  const grab = (re) => { const m = t.match(re); return m ? m[1].trim() : ""; };
  out.title = grab(/SUMMARY[^:]*:(.+)/i) || out.title;
  out.location = grab(/LOCATION[^:]*:(.+)/i) || out.location;
  const dtStart = grab(/DTSTART[^:]*:([0-9TZ]+)/i);
  const dtEnd = grab(/DTEND[^:]*:([0-9TZ]+)/i);
  const icsToISO = s => {
    if (!s) return "";
    const m = s.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z?$/);
    if (!m) return "";
    return `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}`;
  };
  out.startAt = icsToISO(dtStart);
  out.endAt = icsToISO(dtEnd);
  const atts = [...t.matchAll(/ATTENDEE[^:]*:mailto:([^\s\r\n]+)/gi)].map(m => m[1]);
  if (atts.length) out.attendees = atts.join(", ");
  // Fallback: cerca pattern data italiana e ora
  if (!out.startAt) {
    // esempi: 15/11/2026 14:30, 15 novembre 2026 14:30
    const dm = t.match(/(\d{1,2})[\/\-\. ](\d{1,2}|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)[\/\-\. ](\d{4}).{0,20}?(\d{1,2}):(\d{2})/i);
    if (dm) {
      const mesi = { gennaio: 1, febbraio: 2, marzo: 3, aprile: 4, maggio: 5, giugno: 6, luglio: 7, agosto: 8, settembre: 9, ottobre: 10, novembre: 11, dicembre: 12 };
      const d = String(dm[1]).padStart(2, "0");
      const mo = isNaN(Number(dm[2])) ? String(mesi[dm[2].toLowerCase()] || 1).padStart(2, "0") : String(Number(dm[2])).padStart(2, "0");
      const y = dm[3], h = String(dm[4]).padStart(2, "0"), mi = dm[5];
      out.startAt = `${y}-${mo}-${d}T${h}:${mi}`;
    }
  }
  // Se non c'è titolo esplicito, prendi la prima riga
  if (!out.title) out.title = (t.split(/\r?\n/).find(l => l.trim()) || "").slice(0, 100);
  out.notes = t.length < 500 ? t : t.slice(0, 500) + "…";
  return out;
}

function PasteInviteModal({ onParsed, onClose }) {
  const [txt, setTxt] = useState("");
  const parse = () => {
    const a = parseInvite(txt);
    if (!a.title && !a.startAt) { alert("Non sono riuscito a estrarre nulla. Puoi comunque creare l'appuntamento a mano."); return; }
    onParsed(a);
  };
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 560, maxWidth: "100%" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <div>
            <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Incolla invito</h3>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 3 }}>Incolla il testo di un invito .ics (da Google Calendar / Outlook) o di un'email di conferma meeting.</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <textarea className="ia-input" rows={12} value={txt} onChange={e => setTxt(e.target.value)} placeholder="BEGIN:VCALENDAR&#10;SUMMARY:Call con Vulcaflex&#10;DTSTART:20261115T143000Z&#10;...&#10;&#10;oppure un testo tipo:&#10;Riunione con Mario Rossi&#10;Quando: 15/11/2026 14:30&#10;Dove: Zoom" style={{ fontFamily: "monospace", fontSize: 12, resize: "vertical", width: "100%" }} />
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!txt.trim()} onClick={parse}>Estrai e apri</button>
        </div>
      </div>
    </div>
  );
}

// ============ TASKS ============
function Tasks({ ctx, onOpenProspect }) {
  const { tasks, setTasks, prospects, notify, user } = ctx;
  const [edit, setEdit] = useState(null);
  const [filter, setFilter] = useState("mine");
  const T = today();
  const withDefaults = tasks.map(t => ({ ...t, status: t.status || "open" }));
  const filtered = withDefaults.filter(t => {
    if (filter === "mine") return t.status !== "done" && t.assignee === user.id;
    if (filter === "open") return t.status !== "done";
    if (filter === "overdue") return t.status !== "done" && t.dueDate && t.dueDate < T;
    if (filter === "done") return t.status === "done";
    return true;
  }).sort((a, b) => (a.status === "done" ? 1 : 0) - (b.status === "done" ? 1 : 0) || (a.dueDate || "9999").localeCompare(b.dueDate || "9999"));

  const upsert = (t) => {
    const exists = tasks.some(x => x.id === t.id);
    setTasks(exists ? tasks.map(x => x.id === t.id ? t : x) : [t, ...tasks]);
    setEdit(null); notify(exists ? "Task aggiornato" : "Task creato");
  };
  const del = (id) => setTasks(tasks.filter(t => t.id !== id));
  const toggle = (t) => upsert({ ...t, status: t.status === "done" ? "open" : "done", completedAt: t.status === "done" ? null : new Date().toISOString() });

  const counts = { mine: withDefaults.filter(t => t.status !== "done" && t.assignee === user.id).length, open: withDefaults.filter(t => t.status !== "done").length, overdue: withDefaults.filter(t => t.status !== "done" && t.dueDate && t.dueDate < T).length, done: withDefaults.filter(t => t.status === "done").length };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, gap: 10, flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
          {[{ id: "mine", label: "Tuoi (" + counts.mine + ")" }, { id: "open", label: "Aperti (" + counts.open + ")" }, { id: "overdue", label: "Scaduti (" + counts.overdue + ")" }, { id: "done", label: "Chiusi (" + counts.done + ")" }, { id: "all", label: "Tutti" }].map(f => (
            <button key={f.id} className={"ia-btn " + (filter === f.id ? "ia-btn-dark" : "ia-btn-ghost")} style={{ padding: "6px 12px", fontSize: 12.5 }} onClick={() => setFilter(f.id)}>{f.label}</button>
          ))}
        </div>
        <button className="ia-btn ia-btn-primary" onClick={() => setEdit({ id: uid(), title: "", description: "", dueDate: "", assignee: user.id, status: "open", prospectId: "", createdBy: user.id, createdAt: new Date().toISOString() })}><Plus size={16} /> Nuovo task</button>
      </div>
      {filtered.length === 0 ? <Empty icon={ListChecks} title="Nessun task in questo filtro" hint="Crea un task da fare con o senza scadenza." /> : (
        <div className="ia-panel" style={{ padding: 0, overflowX: "auto" }}>
          <table className="ia-table">
            <thead><tr><th></th><th>Task</th><th>Scadenza</th><th>Assegnato a</th><th>Prospect</th><th></th></tr></thead>
            <tbody>
              {filtered.map(t => {
                const p = t.prospectId ? prospects.find(x => x.id === t.prospectId) : null;
                const late = t.status !== "done" && t.dueDate && t.dueDate < T;
                return (
                  <tr key={t.id} style={{ opacity: t.status === "done" ? .55 : 1 }}>
                    <td style={{ width: 30 }}><input type="checkbox" checked={t.status === "done"} onChange={() => toggle(t)} style={{ cursor: "pointer", transform: "scale(1.2)" }} /></td>
                    <td style={{ fontFamily: "Poppins", fontWeight: 600, textDecoration: t.status === "done" ? "line-through" : "none" }}>{t.title}<div style={{ fontWeight: 400, fontSize: 11.5, color: C.muted }}>{t.description}</div></td>
                    <td style={{ whiteSpace: "nowrap", color: late ? C.red : C.ink, fontSize: 12.5 }}>{t.dueDate ? fmtDate(t.dueDate) : "—"}</td>
                    <td style={{ fontSize: 12.5 }}>{teamMeta(t.assignee)?.name || "—"}</td>
                    <td>{p ? <button style={{ border: "none", background: "none", color: C.orange, cursor: "pointer", padding: 0, fontFamily: "Poppins", fontSize: 12.5, textDecoration: "underline" }} onClick={() => onOpenProspect(p.id)}>{p.company}</button> : "—"}</td>
                    <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                      <button className="ia-btn ia-btn-ghost" style={{ padding: "5px 8px" }} onClick={() => setEdit(t)}><Pencil size={13} /></button>{" "}
                      <button className="ia-btn ia-btn-danger" style={{ padding: "5px 8px" }} onClick={() => del(t.id)}><Trash2 size={13} /></button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {edit && <TaskModal t={edit} prospects={prospects} onSave={upsert} onClose={() => setEdit(null)} />}
    </div>
  );
}

function TaskModal({ t, prospects, onSave, onClose }) {
  const [f, setF] = useState(t);
  const set = (k, v) => setF({ ...f, [k]: v });
  const F = Field;
  return (
    <div className="ia-modal-bg" onClick={onClose}>
      <div className="ia-panel" style={{ width: 480, maxWidth: "100%" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid " + C.line }}>
          <h3 className="ia-h" style={{ margin: 0, fontSize: 16, color: C.navy }}>Task</h3>
          <button onClick={onClose} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={18} color={C.muted} /></button>
        </div>
        <div style={{ padding: 20 }}>
          <F label="Titolo *"><input className="ia-input" value={f.title} onChange={e => set("title", e.target.value)} placeholder="es. Preparare proposta Vulcaflex" /></F>
          <F label="Descrizione"><textarea className="ia-input" rows={3} value={f.description} onChange={e => set("description", e.target.value)} style={{ fontFamily: "inherit", resize: "vertical" }} /></F>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <F label="Scadenza"><input className="ia-input" type="date" value={f.dueDate} onChange={e => set("dueDate", e.target.value)} /></F>
            <F label="Assegnato a"><select className="ia-input" value={f.assignee} onChange={e => set("assignee", e.target.value)}>{TEAM.map(u => <option key={u.id} value={u.id}>{u.name}</option>)}</select></F>
          </div>
          <F label="Prospect collegato"><select className="ia-input" value={f.prospectId || ""} onChange={e => set("prospectId", e.target.value)}><option value="">— nessuno —</option>{prospects.map(p => <option key={p.id} value={p.id}>{p.company}</option>)}</select></F>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 20px", borderTop: "1px solid " + C.line }}>
          <button className="ia-btn ia-btn-ghost" onClick={onClose}>Annulla</button>
          <button className="ia-btn ia-btn-primary" disabled={!f.title} onClick={() => onSave(f)}>Salva</button>
        </div>
      </div>
    </div>
  );
}

// ============ SETTINGS ============
function SettingsView({ ctx }) {
  const { notify, cash, setCash } = ctx;
  const [bal, setBal] = useState(cash?.balance != null ? String(cash.balance) : "");
  const saveCash = () => {
    const n = parseItAmount(bal);
    if (isNaN(n)) { notify("Importo non valido", "err"); return; }
    setCash({ balance: n, updatedAt: today() });
    notify("Saldo cassa aggiornato");
  };
  const [busy, setBusy] = useState(false);
  const fileRef = useRef(null);
  const exportBackup = async () => {
    setBusy(true);
    try {
      const data = {};
      for (const k of Object.values(K)) data[k] = await kvGet(k, null);
      const blob = new Blob([JSON.stringify({ app: "ia-suite", exportedAt: new Date().toISOString(), data }, null, 1)], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `ia-suite-backup-${today()}.json`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      notify("Backup scaricato");
    } catch (e) { notify("Backup non riuscito: " + e.message, "err"); }
    setBusy(false);
  };
  const importBackup = async (ev) => {
    const f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    let json;
    try { json = JSON.parse(await f.text()); } catch { notify("File non valido", "err"); return; }
    const data = json && json.data;
    if (!data || json.app !== "ia-suite") { notify("Il file non è un backup della IA Suite", "err"); return; }
    const keys = Object.values(K).filter(k => k in data);
    if (!confirm(`Ripristinare ${keys.length} archivi dal backup del ${String(json.exportedAt || "").slice(0, 10)}? I dati attuali verranno sostituiti.`)) return;
    setBusy(true);
    try {
      for (const k of keys) { if (data[k] == null) await kvDelete(k); else await kvSet(k, data[k]); }
      notify("Dati ripristinati — ricarico"); setTimeout(() => location.reload(), 900);
    } catch (e) { notify("Ripristino non riuscito: " + e.message, "err"); setBusy(false); }
  };
  const reset = async () => {
    if (!confirm("Azzerare TUTTI i dati (contratti, fatture, movimenti, scadenze)? Operazione irreversibile.")) return;
    for (const k of Object.values(K)) { try { await kvDelete(k); } catch (e) {} }
    notify("Dati azzerati — ricarica la pagina"); setTimeout(() => location.reload(), 900);
  };
  return (
    <div style={{ maxWidth: 640 }}>
      <div className="ia-panel ia-bracket" style={{ padding: 22, marginBottom: 16 }}>
        <h3 className="ia-h" style={{ margin: "0 0 4px", fontSize: 15, color: C.navy }}>Saldo cassa attuale</h3>
        <p style={{ fontSize: 13, color: C.muted, margin: "0 0 14px", lineHeight: 1.55 }}>Inserisci il saldo di cassa reale (somma dei conti). Serve al cruscotto per calcolare la posizione netta prevista e la proiezione di cassa. Aggiornalo quando cambia in modo rilevante.</p>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end" }}>
          <div style={{ flex: 1 }}>
            <label style={{ fontSize: 12, fontFamily: "Poppins", fontWeight: 500, color: C.muted, display: "block", marginBottom: 5 }}>Importo (€)</label>
            <input className="ia-input" type="text" placeholder="es. 45.230,00" value={bal} onChange={e => setBal(e.target.value)} onKeyDown={e => e.key === "Enter" && saveCash()} />
          </div>
          <button className="ia-btn ia-btn-primary" onClick={saveCash}><Check size={15} /> Salva</button>
        </div>
        {cash?.updatedAt && <div style={{ fontSize: 12, color: C.muted, marginTop: 10 }}>Ultimo aggiornamento: <b>{fmtDate(cash.updatedAt)}</b> · valore attuale: <b style={{ fontFamily: "Poppins", color: C.navy }}>{fmtMoney(cash.balance)}</b></div>}
      </div>
      <div className="ia-panel" style={{ padding: 22, marginBottom: 16 }}>
        <h3 className="ia-h" style={{ margin: "0 0 4px", fontSize: 15, color: C.navy }}>Accesso</h3>
        <p style={{ fontSize: 13, color: C.muted, margin: "0 0 8px", lineHeight: 1.55 }}>L'accesso è gestito dal portale IA Advisors Hub (Cloudflare Access): si entra con la propria email e un codice ricevuto via posta. Chi può entrare si decide nella policy <b>IA Srl</b> del pannello Zero Trust.</p>
        <p style={{ fontSize: 13, color: C.muted, margin: 0, lineHeight: 1.55 }}>I dati sono condivisi e salvati su Cloudflare (R2), in uno spazio riservato a questa società.</p>
      </div>
      <div className="ia-panel" style={{ padding: 22, marginBottom: 16 }}>
        <h3 className="ia-h" style={{ margin: "0 0 4px", fontSize: 15, color: C.navy }}>Backup dei dati</h3>
        <p style={{ fontSize: 13, color: C.muted, margin: "0 0 14px", lineHeight: 1.55 }}>Scarica una copia completa dei dati (contratti, fatture, movimenti, scadenze, CRM) in un file JSON, oppure ripristina i dati da un file di backup. Il ripristino sostituisce i dati attuali.</p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="ia-btn ia-btn-primary" onClick={exportBackup} disabled={busy}><Download size={15} /> Scarica backup</button>
          <button className="ia-btn" style={{ border: `1px solid ${C.line}`, background: "#fff" }} onClick={() => fileRef.current && fileRef.current.click()} disabled={busy}><Upload size={15} /> Ripristina da backup</button>
          <input ref={fileRef} type="file" accept="application/json,.json" style={{ display: "none" }} onChange={importBackup} />
        </div>
      </div>
      <div className="ia-panel" style={{ padding: 22, borderColor: "#E7CFCF" }}>
        <h3 className="ia-h" style={{ margin: "0 0 4px", fontSize: 15, color: C.red }}>Area dati</h3>
        <p style={{ fontSize: 13, color: C.muted, margin: "0 0 14px" }}>Azzera tutti i dati condivisi. Operazione irreversibile.</p>
        <button className="ia-btn ia-btn-danger" onClick={reset}><RefreshCw size={15} /> Azzera tutti i dati</button>
      </div>
    </div>
  );
}
