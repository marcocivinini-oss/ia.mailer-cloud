// Pagina pubblica di disiscrizione (link nelle email + "one-click" di Gmail/Outlook, RFC 8058)
import { verifyUnsub, addUnsubscribe } from "../lib/common.mjs";

export const config = { path: "/disiscrizione" };

const page = (title, body) => new Response(`<!doctype html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex"><title>${title}</title>
<style>body{margin:0;background:#F4F5F9;font:16px/1.6 Arial,sans-serif;color:#1d2233}
.box{max-width:520px;margin:48px auto;background:#fff;border-top:4px solid #F57547;padding:32px 28px}
img{width:170px;margin-bottom:20px}h1{font-size:20px;color:#131D4B;margin:0 0 12px}
button{background:#131D4B;color:#fff;border:0;padding:12px 24px;font-size:15px;font-weight:bold;cursor:pointer}
.muted{color:#666c80;font-size:14px}@media(max-width:560px){.box{margin:16px}}</style></head>
<body><div class="box"><img src="/disiscrizione/logo.png" alt="International Advisors">${body}</div></body></html>`,
  { status: 200, headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" } });

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

export default async (req) => {
  const url = new URL(req.url);
  const email = verifyUnsub(url.searchParams.get("e"), url.searchParams.get("t"));
  if (!email) return page("Link non valido", `<h1>Link non valido</h1><p>Il link di disiscrizione non è valido o è incompleto. Per essere rimosso dalle nostre comunicazioni scrivi a <a href="mailto:info@iainternationaladvisors.com">info@iainternationaladvisors.com</a>.</p>`);

  if (req.method === "POST") {
    await addUnsubscribe(email, "link", "", url.searchParams.get("c") || "");
    return page("Disiscrizione confermata", `<h1>Disiscrizione confermata</h1><p>L'indirizzo <b>${esc(email)}</b> non riceverà più le nostre comunicazioni.</p><p class="muted">Se si tratta di un errore, rispondi a una delle nostre email e ti reinseriremo.</p>`);
  }
  // GET: chiediamo conferma (evita disiscrizioni accidentali dovute ai filtri antispam che aprono i link)
  return page("Annulla iscrizione", `<h1>Vuoi annullare l'iscrizione?</h1><p>L'indirizzo <b>${esc(email)}</b> non riceverà più le comunicazioni di IA International Advisors.</p>
<form method="post"><button type="submit">Conferma la disiscrizione</button></form>`);
};
