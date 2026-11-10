// TicketUp — integrazione Outlook (Microsoft Graph) v2 — accesso solo utenti in tu_utenti_abilitati
// Azioni: status | send | poll | solleciti
// Secret richiesti (forniti dall'IT): MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, TICKETUP_MAILBOX
// Secret opzionale per il cron: TICKETUP_CRON_KEY
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "jsr:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-cron-key",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const json = (b: unknown, s = 200) =>
  new Response(JSON.stringify(b), { status: s, headers: { ...CORS, "Content-Type": "application/json" } });

const env = (k: string) => Deno.env.get(k) ?? "";
const TENANT = env("MS_TENANT_ID"), CLIENT = env("MS_CLIENT_ID"), SECRET = env("MS_CLIENT_SECRET");
const MAILBOX = env("TICKETUP_MAILBOX");
const CONFIGURED = !!(TENANT && CLIENT && SECRET && MAILBOX);
const MAX_SOLLECITI = 3;
const CODE_RE = /\[TU-(\d{4,})\]/i;

const admin = createClient(env("SUPABASE_URL"), env("SUPABASE_SERVICE_ROLE_KEY"));

let tokenCache: { t: string; exp: number } | null = null;
async function graphToken() {
  if (tokenCache && tokenCache.exp > Date.now() + 60_000) return tokenCache.t;
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT, client_secret: SECRET, grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  const d = await r.json();
  if (!r.ok) throw new Error(`Token Microsoft: ${d.error_description ?? d.error}`);
  tokenCache = { t: d.access_token, exp: Date.now() + d.expires_in * 1000 };
  return tokenCache.t;
}
async function graph(path: string, init: RequestInit = {}) {
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MAILBOX)}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${await graphToken()}`, "Content-Type": "application/json", ...(init.headers ?? {}) },
  });
  if (r.status === 202 || r.status === 204) return null;
  const d = await r.json().catch(() => null);
  if (!r.ok) throw new Error(`Graph ${r.status}: ${d?.error?.message ?? "errore"}`);
  return d;
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]!));
const htmlBody = (corpo: string, codice: string) =>
  `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px">${esc(corpo).replace(/\n/g, "<br>")}</div>` +
  `<p style="font-size:11px;color:#888;margin-top:24px">Rif. ${codice} — per rispondere mantenere il codice tra parentesi quadre nell'oggetto.</p>`;

// Invia una mail per il ticket e registra il messaggio
async function sendForTicket(ticketId: number, corpo: string, tipo: "inviato" | "sollecito", autore?: { id: string; email: string }) {
  const { data: t, error } = await admin.from("tu_ticket")
    .select("*, ufficio:tu_uffici(categoria,email,cc,giorni_sollecito)").eq("id", ticketId).single();
  if (error || !t) throw new Error("Ticket non trovato");
  const to = t.destinatario || t.ufficio.email;
  const cc: string[] = (t.cc?.length ? t.cc : t.ufficio.cc) ?? [];
  const oggetto = `${tipo === "sollecito" ? "Sollecito – " : ""}[${t.codice}] ${t.oggetto}`;

  const draft = await graph("/messages", {
    method: "POST",
    body: JSON.stringify({
      subject: oggetto,
      body: { contentType: "HTML", content: htmlBody(corpo, t.codice) },
      toRecipients: [{ emailAddress: { address: to } }],
      ccRecipients: cc.map((a) => ({ emailAddress: { address: a } })),
    }),
  });
  await graph(`/messages/${draft.id}/send`, { method: "POST" });

  const now = new Date();
  const next = new Date(now.getTime() + t.ufficio.giorni_sollecito * 86400_000);
  await admin.from("tu_messaggi").insert({
    ticket_id: t.id, tipo, canale: "graph", mittente: MAILBOX, destinatari: [to, ...cc],
    oggetto, corpo, autore: autore?.id ?? null, autore_email: autore?.email ?? "TicketUp (automatico)",
  });
  await admin.from("tu_ticket").update({
    stato: "in_attesa", ultimo_invio_at: now.toISOString(), prossimo_sollecito_at: next.toISOString(),
    graph_conversation_id: t.graph_conversation_id ?? draft.conversationId,
    n_solleciti: tipo === "sollecito" ? t.n_solleciti + 1 : t.n_solleciti,
  }).eq("id", t.id);
  return { codice: t.codice, oggetto };
}

// Legge la casella e aggancia le risposte ai ticket
async function poll() {
  const d = await graph(
    "/mailFolders/inbox/messages?$filter=isRead eq false&$top=50&$orderby=receivedDateTime asc" +
      "&$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,uniqueBody,conversationId",
    { headers: { Prefer: 'outlook.body-content-type="text"' } },
  );
  let agganciati = 0, ignorati = 0;
  for (const m of d?.value ?? []) {
    let ticket: { id: number; stato: string } | null = null;
    const code = m.subject?.match(CODE_RE);
    if (code) {
      const r = await admin.from("tu_ticket").select("id,stato").eq("codice", `TU-${code[1]}`).maybeSingle();
      ticket = r.data;
    }
    if (!ticket && m.conversationId) {
      const r = await admin.from("tu_ticket").select("id,stato").eq("graph_conversation_id", m.conversationId).maybeSingle();
      ticket = r.data;
    }
    if (!ticket) { ignorati++; continue; } // mail non legata a un ticket: resta non letta in casella

    const { error } = await admin.from("tu_messaggi").insert({
      ticket_id: ticket.id, tipo: "ricevuto", canale: "graph",
      mittente: m.from?.emailAddress?.address,
      destinatari: [...(m.toRecipients ?? []), ...(m.ccRecipients ?? [])].map((x: any) => x.emailAddress.address),
      oggetto: m.subject, corpo: (m.uniqueBody?.content ?? "").trim(),
      graph_message_id: m.id, created_at: m.receivedDateTime, autore_email: m.from?.emailAddress?.address,
    });
    if (!error || error.code === "23505") { // 23505 = già importata
      await admin.from("tu_ticket").update({
        stato: "risposto", ultima_risposta_at: m.receivedDateTime, prossimo_sollecito_at: null,
      }).eq("id", ticket.id);
      await graph(`/messages/${m.id}`, { method: "PATCH", body: JSON.stringify({ isRead: true }) });
      agganciati++;
    }
  }
  return { agganciati, ignorati };
}

async function solleciti() {
  const { data } = await admin.from("tu_ticket").select("id,codice,oggetto,n_solleciti")
    .eq("stato", "in_attesa").lt("prossimo_sollecito_at", new Date().toISOString()).lt("n_solleciti", MAX_SOLLECITI);
  const inviati: string[] = [];
  for (const t of data ?? []) {
    const corpo = `Buongiorno,\nvi sollecitiamo gentilmente un riscontro sulla richiesta ${t.codice} "${t.oggetto}".\n\nGrazie,\nPDE`;
    await sendForTicket(t.id, corpo, "sollecito");
    inviati.push(t.codice);
  }
  return { inviati };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const body = await req.json().catch(() => ({}));
    const action = body.action ?? "status";
    if (action === "status") return json({ configured: CONFIGURED, mailbox: CONFIGURED ? MAILBOX : null });

    // Autorizzazione: cron con chiave, oppure utente admin
    const cronKey = env("TICKETUP_CRON_KEY");
    const isCron = !!cronKey && req.headers.get("x-cron-key") === cronKey;
    let user: { id: string; email: string } | undefined;
    if (!isCron) {
      const jwt = (req.headers.get("Authorization") ?? "").replace("Bearer ", "");
      const { data } = await admin.auth.getUser(jwt);
      if (!data.user) return json({ error: "Non autorizzato" }, 403);
      const { data: ok } = await admin.from("tu_utenti_abilitati").select("user_id").eq("user_id", data.user.id).maybeSingle();
      if (!ok) return json({ error: "Non autorizzato" }, 403);
      user = { id: data.user.id, email: data.user.email ?? "" };
    }
    if (!CONFIGURED) return json({ error: "Integrazione Outlook non ancora configurata", configured: false }, 409);

    if (action === "send") {
      if (!body.ticket_id || !body.corpo) return json({ error: "ticket_id e corpo obbligatori" }, 400);
      return json(await sendForTicket(body.ticket_id, body.corpo, body.tipo === "sollecito" ? "sollecito" : "inviato", user));
    }
    if (action === "poll") return json(await poll());
    if (action === "solleciti") return json({ ...(await poll()), ...(await solleciti()) });
    return json({ error: "Azione sconosciuta" }, 400);
  } catch (e) {
    return json({ error: (e as Error).message }, 500);
  }
});
