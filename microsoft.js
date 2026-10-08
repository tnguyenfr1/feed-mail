// Microsoft sign-in (OAuth) for personal Outlook/Hotmail accounts, plus
// Microsoft Graph helpers and Outlook Calendar.
import { loadSecrets, saveSecrets } from './store.js';

const AUTHORITY = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['openid', 'email', 'profile', 'offline_access', 'User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Calendars.ReadWrite'];

let secrets = loadSecrets();
let persist = () => {};

// Microsoft hands out a new refresh token on every refresh; the server
// passes a function here that saves accounts when that happens.
export function onRefreshTokenChange(fn) { persist = fn; }

export const microsoftConfigured = () => !!(secrets.microsoft?.clientId && secrets.microsoft?.clientSecret);

export function setMicrosoftClient(clientId, clientSecret) {
  secrets = { ...loadSecrets(), microsoft: { clientId: clientId.trim(), clientSecret: clientSecret.trim() } };
  saveSecrets(secrets);
}

export function authUrl(redirectUri, state) {
  const q = new URLSearchParams({
    client_id: secrets.microsoft.clientId,
    redirect_uri: redirectUri,
    response_type: 'code',
    response_mode: 'query',
    scope: SCOPES.join(' '),
    prompt: 'select_account',
    state,
  });
  return `${AUTHORITY}/authorize?${q}`;
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTHORITY}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: secrets.microsoft.clientId, client_secret: secrets.microsoft.clientSecret, scope: SCOPES.join(' '), ...params,
    }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const expired = data.error === 'invalid_grant';
    const secretGone = /AADSTS7000222|AADSTS7000215/.test(data.error_description || '');
    const e = new Error(secretGone
      ? 'The Microsoft client secret has expired or is wrong. Create a new one in Azure and paste it in Accounts.'
      : expired
        ? 'Microsoft sign-in has expired or was revoked. Reconnect this account in Accounts.'
        : `Microsoft refused the request (${(data.error_description || data.error || res.status).toString().split('\r\n')[0]}).`);
    e.status = 401;
    throw e;
  }
  return data;
}

export async function exchangeCode(code, redirectUri) {
  const t = await tokenRequest({ code, redirect_uri: redirectUri, grant_type: 'authorization_code' });
  if (!t.refresh_token) throw new Error('Microsoft did not return a refresh token. Please try again.');
  const claims = JSON.parse(Buffer.from(t.id_token.split('.')[1], 'base64url').toString('utf8'));
  const email = claims.email || claims.preferred_username;
  if (!email) throw new Error('Could not read the email address from Microsoft.');
  return { email, refreshToken: t.refresh_token };
}

const tokens = new Map(); // account id -> { token, exp } or { pending }

export function forgetToken(acctId) { tokens.delete(acctId); }

export async function accessToken(acct) {
  const t = tokens.get(acct.id);
  if (t?.token && t.exp - 60000 > Date.now()) return t.token;
  if (t?.pending) return t.pending;
  const pending = tokenRequest({ refresh_token: acct.oauth.refreshToken, grant_type: 'refresh_token' })
    .then((d) => {
      tokens.set(acct.id, { token: d.access_token, exp: Date.now() + d.expires_in * 1000 });
      if (d.refresh_token && d.refresh_token !== acct.oauth.refreshToken) {
        acct.oauth.refreshToken = d.refresh_token;
        try { persist(acct); } catch {}
      }
      return d.access_token;
    })
    .catch((err) => { tokens.delete(acct.id); throw err; });
  tokens.set(acct.id, { pending });
  return pending;
}

// Call Microsoft Graph. Times come back in UTC and message ids stay the same
// when a message is moved.
export async function graph(acct, path, { method = 'GET', body, query, raw = false } = {}) {
  const url = path.startsWith('https://') ? path : `${GRAPH}${path}${query ? '?' + new URLSearchParams(query) : ''}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${await accessToken(acct)}`,
      Prefer: 'outlook.timezone="UTC", IdType="ImmutableId"',
      ...(body && { 'Content-Type': 'application/json' }),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (raw && res.ok) return res;
  const data = res.status === 202 || res.status === 204 ? {} : await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error(`Outlook: ${data.error?.message || res.status}`);
    e.status = res.status === 404 ? 404 : res.status === 401 ? 401 : 502;
    throw e;
  }
  return data;
}

// Graph gives UTC times without a "Z".
export const utc = (dt) => (dt ? new Date(dt.endsWith('Z') ? dt : dt + 'Z').toISOString() : null);

// ---------- calendar ----------

export async function outlookEvents(acct, from, to) {
  const cals = await graph(acct, '/me/calendars', { query: { $select: 'id,name,canEdit', $top: '50' } });
  const out = [];
  for (const cal of cals.value || []) {
    let next = `${GRAPH}/me/calendars/${encodeURIComponent(cal.id)}/calendarView?` + new URLSearchParams({
      startDateTime: from.toISOString(), endDateTime: to.toISOString(), $top: '250',
      $select: 'id,type,subject,start,end,isAllDay,location,iCalUId,isCancelled,responseStatus',
    });
    while (next) {
      const r = await graph(acct, next);
      for (const e of r.value || []) {
        if (e.isCancelled || e.responseStatus?.response === 'declined') continue;
        out.push({
          uid: e.iCalUId,
          ref: e.id, // what's needed to remove it
          recurring: e.type === 'occurrence' || e.type === 'exception',
          readOnly: cal.canEdit === false,
          title: e.subject || '(no title)',
          location: e.location?.displayName || '',
          allDay: !!e.isAllDay,
          start: e.isAllDay ? e.start.dateTime.slice(0, 10) : utc(e.start.dateTime),
          end: e.isAllDay ? e.end.dateTime.slice(0, 10) : utc(e.end.dateTime),
          calendar: cal.name || 'Calendar',
        });
      }
      next = r['@odata.nextLink'];
    }
  }
  return out;
}

export async function testOutlookCalendar(acct) {
  return ((await graph(acct, '/me/calendars', { query: { $select: 'id' } })).value || []).length;
}

// ---------- add / remove events ----------

export async function outlookCreate(acct, ev) {
  const at = (iso, date) => (ev.allDay
    ? { dateTime: `${date}T00:00:00`, timeZone: 'UTC' }
    : { dateTime: new Date(iso).toISOString().replace('Z', ''), timeZone: 'UTC' });
  await graph(acct, '/me/events', {
    method: 'POST',
    body: {
      subject: ev.title,
      isAllDay: !!ev.allDay,
      start: at(ev.start, ev.date),
      end: at(ev.end, ev.endDate),
      ...(ev.location && { location: { displayName: ev.location } }),
      ...(ev.notes && { body: { contentType: 'Text', content: ev.notes } }),
    },
  });
  return 'Outlook Calendar';
}

export async function outlookDelete(acct, ref) {
  await graph(acct, `/me/events/${encodeURIComponent(ref)}`, { method: 'DELETE' });
}
