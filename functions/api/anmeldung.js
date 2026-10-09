// Cloudflare Pages Function: /api/anmeldung
// Nimmt das Formular „Zum Start benachrichtigen“ entgegen und schickt die Angaben
// über Microsoft Graph als E-Mail an support@cubarc.at. Es wird nichts gespeichert.
//
// Benötigte Umgebungsvariablen (Cloudflare → Workers & Pages → cubarc → Einstellungen → Variablen und Geheimnisse):
//   MS_CLIENT_ID      Anwendungs-ID (Client-ID) der App-Registrierung „cubarc Website“ in Microsoft Entra
//   MS_CLIENT_SECRET  geheimer Clientschlüssel dieser App (als „Geheimnis“ anlegen)
// Optional:
//   MS_TENANT_ID      Mandant; Standard ist die Domain cubarc.at
//   MAIL_TO           Empfänger; Standard support@cubarc.at
//   MAIL_FROM         Absender-Postfach; Standard support@cubarc.at

const LAENDER = ['Burgenland', 'Kärnten', 'Niederösterreich', 'Oberösterreich', 'Salzburg', 'Steiermark', 'Tirol', 'Vorarlberg', 'Wien', 'Deutschland', 'Schweiz', 'Anderes Land'];
const TAETIG = ['Bauträgerunternehmen', 'Planungs- oder Architekturbüro', 'Projektentwicklung', 'Immobilienvermittlung', 'Grundstückseigentum', 'Sonstiges'];

const json = (obj, status = 200) => new Response(JSON.stringify(obj), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const clean = (v, max) => String(v ?? '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '').trim().slice(0, max);
const bereit = env => !!(env.MS_CLIENT_ID && env.MS_CLIENT_SECRET);

let token = null, tokenBis = 0;
async function graphToken(env) {
  if (token && Date.now() < tokenBis) return token;
  const tenant = env.MS_TENANT_ID || 'cubarc.at';
  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: env.MS_CLIENT_ID, client_secret: env.MS_CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }) });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new Error('token ' + r.status + ' ' + (j.error || ''));
  token = j.access_token; tokenBis = Date.now() + (Number(j.expires_in || 3600) - 120) * 1000;
  return token;
}

function mailHtml(d, zeit) {
  const zeile = (k, v) => v ? `<tr><td style="padding:7px 16px 7px 0;color:#7a7880;white-space:nowrap;vertical-align:top">${k}</td><td style="padding:7px 0;color:#1d1c20">${v}</td></tr>` : '';
  return `<div style="font-family:Segoe UI,Helvetica,Arial,sans-serif;font-size:15px;line-height:1.5;color:#1d1c20;max-width:560px">
<p style="margin:0 0 4px;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#c27c3c;font-weight:600">cubarc · Startbenachrichtigung</p>
<h2 style="margin:0 0 18px;font-size:22px;font-weight:600">Neue Anmeldung von ${esc(d.name)}</h2>
<table style="border-collapse:collapse;font-size:15px">
${zeile('Name', esc(d.name))}
${zeile('E-Mail', `<a href="mailto:${esc(d.email)}" style="color:#c27c3c">${esc(d.email)}</a>`)}
${zeile('Unternehmen', esc(d.firma))}
${zeile('Bundesland', esc(d.land))}
${zeile('Tätig in', esc(d.taetig))}
${zeile('Nachricht', esc(d.nachricht).replace(/\n/g, '<br>'))}
${zeile('Einwilligung', 'erteilt am ' + esc(zeit))}
</table>
<p style="margin:22px 0 0;font-size:13px;color:#7a7880">Gesendet über das Formular auf cubarc.at/benachrichtigen. „Antworten“ geht direkt an die angegebene Adresse.</p>
</div>`;
}

async function senden(env, d) {
  const an = env.MAIL_TO || 'support@cubarc.at', von = env.MAIL_FROM || 'support@cubarc.at';
  const zeit = new Date().toLocaleString('de-AT', { timeZone: 'Europe/Vienna', dateStyle: 'medium', timeStyle: 'short' });
  const t = await graphToken(env);
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(von)}/sendMail`, {
    method: 'POST',
    headers: { authorization: 'Bearer ' + t, 'content-type': 'application/json' },
    body: JSON.stringify({ saveToSentItems: false, message: {
      subject: `Startbenachrichtigung: ${d.name}${d.firma ? ' (' + d.firma + ')' : ''}`,
      body: { contentType: 'HTML', content: mailHtml(d, zeit) },
      toRecipients: [{ emailAddress: { address: an } }],
      replyTo: [{ emailAddress: { address: d.email, name: d.name } }] } }) });
  if (r.status !== 202) throw new Error('graph ' + r.status + ' ' + (await r.text()).slice(0, 300));
}

export async function onRequestGet({ env }) {
  return json({ bereit: bereit(env) });
}

export async function onRequestPost({ request, env }) {
  const alsFormular = !(request.headers.get('content-type') || '').includes('application/json');
  const zurueck = (ziel, j, status) => alsFormular
    ? Response.redirect(new URL('/benachrichtigen#' + ziel, request.url).toString(), 303)
    : json(j, status);

  const origin = request.headers.get('origin');
  if (origin && new URL(origin).host !== new URL(request.url).host) return zurueck('fehler', { fehler: 'Ungültige Herkunft.' }, 403);
  if (Number(request.headers.get('content-length') || 0) > 20000) return zurueck('fehler', { fehler: 'Die Angaben sind zu lang.' }, 413);

  let roh;
  try { roh = alsFormular ? Object.fromEntries(await request.formData()) : await request.json(); }
  catch { return zurueck('fehler', { fehler: 'Die Angaben konnten nicht gelesen werden.' }, 400); }

  // Schutz vor automatischen Einträgen: verstecktes Feld und Mindestzeit auf der Seite
  const t = Number(roh.t || 0);
  if (roh.website || (t && Date.now() - t < 2500)) return zurueck('danke', { ok: true }, 200);

  const d = {
    name: clean(roh.name, 120), email: clean(roh.email, 254).toLowerCase(), firma: clean(roh.firma, 160),
    land: LAENDER.includes(roh.land) ? roh.land : '', taetig: TAETIG.includes(roh.taetig) ? roh.taetig : '',
    nachricht: clean(roh.nachricht, 2000),
  };
  const zustimmung = roh.zustimmung === true || roh.zustimmung === 'ja' || roh.zustimmung === 'on';
  if (d.name.length < 2) return zurueck('fehler', { fehler: 'Bitte geben Sie Ihren Namen an.', feld: 'name' }, 422);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) return zurueck('fehler', { fehler: 'Bitte prüfen Sie Ihre E-Mail-Adresse.', feld: 'email' }, 422);
  if (!zustimmung) return zurueck('fehler', { fehler: 'Bitte bestätigen Sie die Einwilligung.', feld: 'zustimmung' }, 422);

  if (!bereit(env)) return zurueck('fehler', { fehler: 'Das Formular ist noch nicht freigeschaltet.', bereit: false }, 503);

  // höchstens eine Anmeldung pro Adresse und Minute (je Rechenzentrum)
  try {
    const ip = request.headers.get('cf-connecting-ip') || '';
    const key = new Request(new URL('/__rl/' + encodeURIComponent(ip), request.url).toString());
    if (ip && await caches.default.match(key)) return zurueck('fehler', { fehler: 'Bitte warten Sie kurz und versuchen Sie es dann noch einmal.' }, 429);
    if (ip) await caches.default.put(key, new Response('1', { headers: { 'cache-control': 'max-age=60' } }));
  } catch (e) {}

  try { await senden(env, d); }
  catch (e) {
    console.log('anmeldung', String(e));
    return zurueck('fehler', { fehler: 'Die Anmeldung konnte gerade nicht gesendet werden.' }, 502);
  }
  return zurueck('danke', { ok: true }, 200);
}
