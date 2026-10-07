// Google sign-in (OAuth) for Gmail, plus Google Calendar through its API.
import { loadSecrets, saveSecrets } from './store.js';

const AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const TOKEN_URL = 'https://oauth2.googleapis.com/token';
const CAL_API = 'https://www.googleapis.com/calendar/v3';
const SCOPES = ['openid', 'email', 'https://mail.google.com/', 'https://www.googleapis.com/auth/calendar'];

let secrets = loadSecrets();

export const googleConfigured = () => !!(secrets.google?.clientId && secrets.google?.clientSecret);

export function setGoogleClient(clientId, clientSecret) {
  secrets = { ...secrets, google: { clientId: clientId.trim(), clientSecret: clientSecret.trim() } };
  saveSecrets(secrets);
}

export function authUrl(redirectUri, state) {
  const q = new URLSearchParams({
    client_id: secrets.google.clientId,
    redirect_uri: redirectUri,
    response_type: 'code',
    scope: SCOPES.join(' '),
    access_type: 'offline',
    prompt: 'consent select_account',
    include_granted_scopes: 'true',
    state,
  });
  return `${AUTH_URL}?${q}`;
}

async function tokenRequest(params) {
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: secrets.google.clientId, client_secret: secrets.google.clientSecret, ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error(data.error === 'invalid_grant'
      ? 'Google sign-in has expired or was revoked. Reconnect this account in Accounts.'
      : `Google refused the request (${data.error_description || data.error || res.status}).`);
    e.status = 401;
    throw e;
  }
  return data;
}

export async function exchangeCode(code, redirectUri) {
  const t = await tokenRequest({ code, redirect_uri: redirectUri, grant_type: 'authorization_code' });
  // The id_token comes straight from Google over TLS, so reading its payload is enough here.
  const claims = JSON.parse(Buffer.from(t.id_token.split('.')[1], 'base64url').toString('utf8'));
  if (!t.refresh_token) throw new Error('Google did not return a refresh token. Remove Feed Mail from your Google account permissions and try again.');
  const scopes = (t.scope || '').split(' ');
  for (const s of ['https://mail.google.com/', 'https://www.googleapis.com/auth/calendar']) {
    if (!scopes.includes(s)) throw new Error('Please tick all the permission boxes on the Google screen (mail and calendar).');
  }
  return { email: claims.email, refreshToken: t.refresh_token };
}

const tokens = new Map(); // account id -> { token, exp } or pending promise

export function forgetToken(acctId) {
  tokens.delete(acctId);
}

export async function accessToken(acct) {
  const t = tokens.get(acct.id);
  if (t?.token && t.exp - 60000 > Date.now()) return t.token;
  if (t?.pending) return t.pending;
  const pending = tokenRequest({ refresh_token: acct.oauth.refreshToken, grant_type: 'refresh_token' })
    .then((d) => {
      tokens.set(acct.id, { token: d.access_token, exp: Date.now() + d.expires_in * 1000 });
      return d.access_token;
    })
    .catch((err) => { tokens.delete(acct.id); throw err; });
  tokens.set(acct.id, { pending });
  return pending;
}

async function gapi(acct, path, { method = 'GET', body, query } = {}) {
  const url = `${CAL_API}${path}${query ? '?' + new URLSearchParams(query) : ''}`;
  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${await accessToken(acct)}`, ...(body && { 'Content-Type': 'application/json' }) },
    body: body ? JSON.stringify(body) : undefined,
  });
  const data = res.status === 204 ? {} : await res.json().catch(() => ({}));
  if (!res.ok) {
    const e = new Error(`Google Calendar: ${data.error?.message || res.status}`);
    e.status = res.status;
    throw e;
  }
  return data;
}

// ---------- calendar ----------

const toInstant = (t) => (t.date ? t.date : new Date(t.dateTime).toISOString());

export async function googleCalendars(acct) {
  const list = await gapi(acct, '/users/me/calendarList', { query: { minAccessRole: 'reader', maxResults: '100' } });
  return (list.items || []).filter((c) => c.primary || (c.selected && !c.hidden));
}

export async function googleEvents(acct, from, to) {
  const out = [];
  for (const cal of await googleCalendars(acct)) {
    let pageToken;
    do {
      const r = await gapi(acct, `/calendars/${encodeURIComponent(cal.id)}/events`, {
        query: {
          timeMin: from.toISOString(), timeMax: to.toISOString(), singleEvents: 'true',
          orderBy: 'startTime', maxResults: '250', ...(pageToken && { pageToken }),
        },
      });
      for (const e of r.items || []) {
        if (e.status === 'cancelled' || !e.start) continue;
        const me = (e.attendees || []).find((a) => a.self);
        if (me?.responseStatus === 'declined') continue;
        out.push({
          uid: e.iCalUID || e.id,
          title: e.summary || '(no title)',
          location: e.location || '',
          allDay: !!e.start.date,
          start: toInstant(e.start),
          end: toInstant(e.end || e.start),
          calendar: cal.summaryOverride || cal.summary || 'Calendar',
        });
      }
      pageToken = r.nextPageToken;
    } while (pageToken);
  }
  return out;
}

const GOOGLE_STATUS = { ACCEPTED: 'accepted', TENTATIVE: 'tentative', DECLINED: 'declined' };

// If Google already put the invite in the calendar (usual for Gmail), answer
// through the API; Google then notifies the organizer itself.
export async function googleRsvp(acct, invite, partstat) {
  const found = await gapi(acct, '/calendars/primary/events', { query: { iCalUID: invite.uid, maxResults: '1' } });
  const ev = found.items?.[0];
  if (!ev) return false;
  const attendees = (ev.attendees || []).map((a) =>
    a.self || a.email?.toLowerCase() === acct.email.toLowerCase() ? { ...a, responseStatus: GOOGLE_STATUS[partstat] } : a);
  if (!attendees.some((a) => a.self || a.email?.toLowerCase() === acct.email.toLowerCase())) {
    attendees.push({ email: acct.email, responseStatus: GOOGLE_STATUS[partstat] });
  }
  await gapi(acct, `/calendars/primary/events/${encodeURIComponent(ev.id)}`, {
    method: 'PATCH', query: { sendUpdates: 'all' }, body: { attendees },
  });
  return true;
}

// Add an invite that Google doesn't know about yet (e.g. from another provider).
export async function googleImport(acct, invite, rrules, partstat) {
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const when = (v) => (invite.allDay ? { date: v } : { dateTime: v, timeZone: tz });
  await gapi(acct, '/calendars/primary/events/import', {
    method: 'POST',
    body: {
      iCalUID: invite.uid,
      summary: invite.title,
      location: invite.location || undefined,
      start: when(invite.start),
      end: when(invite.end || invite.start),
      recurrence: rrules.length ? rrules : undefined,
      organizer: invite.organizer?.address ? { email: invite.organizer.address, displayName: invite.organizer.name || undefined } : undefined,
      attendees: [{ email: acct.email, responseStatus: GOOGLE_STATUS[partstat] }],
    },
  });
  return 'Google Calendar';
}
