import crypto from 'node:crypto';
import fs from 'node:fs';
import { execFile } from 'node:child_process';
import express from 'express';
import compression from 'compression';
import addressparser from 'nodemailer/lib/addressparser';
import { loadAccounts, saveAccounts, loadConfig, saveConfig } from './store.js';
import { PROVIDERS, caldavUrlFor, calendarKind } from './providers.js';
import * as mail from './mail.js';
import * as cal from './calendar.js';
import * as google from './google.js';
import * as microsoft from './microsoft.js';
import * as outlook from './outlook.js';

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '127.0.0.1';
const STALE_MS = 45 * 1000;
const SESSION_DAYS = 90;
const MAX_ATTACH = 20 * 1024 * 1024;

const config = loadConfig();
let accounts = loadAccounts();
const cache = new Map(); // account id -> { at, items, error, pending }
const junkCache = new Map(); // same, for each account's junk / spam folder

// Every account gets its own colour, so each email shows where it came from.
const PALETTE = ['#3b6cf6', '#ea4335', '#16a34a', '#7c3aed', '#ea580c', '#0891b2', '#db2777', '#ca8a04', '#64748b'];
const freeColor = (wanted) => {
  const used = new Set(accounts.map((a) => a.color));
  return wanted && !used.has(wanted) ? wanted : PALETTE.find((c) => !used.has(c)) || PALETTE[accounts.length % PALETTE.length];
};
{
  // Older accounts were often all left on the default blue: spread them out once.
  const seen = new Set();
  let changed = false;
  for (const a of accounts) {
    if (a.color && !seen.has(a.color)) { seen.add(a.color); continue; }
    a.color = PALETTE.find((c) => !seen.has(c) && !accounts.some((b) => b.color === c)) || PALETTE.find((c) => !seen.has(c)) || '#64748b';
    seen.add(a.color);
    changed = true;
  }
  if (changed) saveAccounts(accounts);
}

const app = express();
app.disable('x-powered-by');
app.use(compression());
app.use(express.json({ limit: '40mb' })); // replies can carry attachments
app.use((req, res, next) => {
  res.set('X-Content-Type-Options', 'nosniff');
  res.set('Referrer-Policy', 'no-referrer');
  next();
});

// ---------- login ----------

function hashPassword(pw, salt = crypto.randomBytes(16).toString('hex')) {
  return { salt, hash: crypto.scryptSync(pw, salt, 64).toString('hex') };
}

function sign(value) {
  return crypto.createHmac('sha256', config.sessionSecret).update(value).digest('base64url');
}

function issueSession(res) {
  const exp = String(Date.now() + SESSION_DAYS * 864e5);
  res.cookie('fm_session', `${exp}.${sign(exp)}`, {
    httpOnly: true, secure: true, sameSite: 'strict', maxAge: SESSION_DAYS * 864e5, path: '/',
  });
}

function readCookie(req, name) {
  const m = (req.headers.cookie || '').match(new RegExp(`(?:^|;\\s*)${name}=([^;]+)`));
  return m ? decodeURIComponent(m[1]) : null;
}

function loggedIn(req) {
  const v = readCookie(req, 'fm_session');
  if (!v) return false;
  const [exp, sig] = v.split('.');
  if (!exp || !sig || Number(exp) < Date.now()) return false;
  const good = sign(exp);
  return sig.length === good.length && crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(good));
}

app.get('/api/session', (req, res) => {
  res.json({ setup: !config.password, loggedIn: loggedIn(req) });
});

app.post('/api/setup', (req, res) => {
  if (config.password) return res.status(403).json({ error: 'Already set up' });
  const pw = String(req.body.password || '');
  if (pw.length < 8) return res.status(400).json({ error: 'Use at least 8 characters.' });
  config.password = hashPassword(pw);
  saveConfig(config);
  issueSession(res);
  res.json({ ok: true });
});

let lastFail = 0;
app.post('/api/login', async (req, res) => {
  const wait = lastFail + 1500 - Date.now();
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
  const pw = String(req.body.password || '');
  const { hash } = hashPassword(pw, config.password?.salt || 'x');
  if (config.password && crypto.timingSafeEqual(Buffer.from(hash), Buffer.from(config.password.hash))) {
    issueSession(res);
    return res.json({ ok: true });
  }
  lastFail = Date.now();
  res.status(401).json({ error: 'Wrong password' });
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('fm_session', { path: '/' });
  res.json({ ok: true });
});

// ---------- Google / Microsoft sign-in (browser redirects, so outside /api) ----------

const OAUTH = {
  google: { mod: google, provider: 'gmail', configured: () => google.googleConfigured(), color: '#ea4335' },
  microsoft: { mod: microsoft, provider: 'outlook', configured: () => microsoft.microsoftConfigured(), color: '#0891b2' },
};

// Microsoft rotates refresh tokens; keep the newest one on disk.
microsoft.onRefreshTokenChange(() => saveAccounts(accounts));

const oauthRedirect = (req, kind) =>
  `https://${process.env.PUBLIC_HOST || req.headers['x-forwarded-host'] || req.headers.host}/oauth/${kind}/callback`;

app.get('/oauth/:kind/start', (req, res) => {
  const o = OAUTH[req.params.kind];
  if (!o || !loggedIn(req) || !o.configured()) return res.redirect('/');
  const state = crypto.randomBytes(16).toString('hex');
  // Lax so the cookie survives the trip to the provider and back; the session cookie is Strict.
  res.cookie('fm_oauth', state, { httpOnly: true, secure: true, sameSite: 'lax', maxAge: 15 * 60 * 1000, path: '/oauth' });
  res.redirect(o.mod.authUrl(oauthRedirect(req, req.params.kind), state));
});

app.get('/oauth/:kind/callback', async (req, res) => {
  const back = (msg) => res.type('html').send(
    `<!doctype html><meta charset="utf-8"><script>location.replace(${JSON.stringify('/?' + msg)})</script>`);
  const o = OAUTH[req.params.kind];
  if (!o) return back('oauthError=' + encodeURIComponent('Unknown sign-in.'));
  const state = readCookie(req, 'fm_oauth');
  res.clearCookie('fm_oauth', { path: '/oauth' });
  if (!state || state !== req.query.state) return back('oauthError=' + encodeURIComponent('Sign-in expired, please try again.'));
  if (req.query.error) return back('oauthError=' + encodeURIComponent('Sign-in was cancelled.'));
  try {
    const { email, refreshToken } = await o.mod.exchangeCode(String(req.query.code || ''), oauthRedirect(req, req.params.kind));
    const existing = accounts.find((a) => a.provider === o.provider && a.email.toLowerCase() === email.toLowerCase());
    const acct = existing
      ? { ...existing, oauth: { refreshToken } }
      : { id: crypto.randomBytes(4).toString('hex'), provider: o.provider, email, oauth: { refreshToken }, name: '',
          color: freeColor(o.color) };
    mail.dropClient(acct.id);
    o.mod.forgetToken(acct.id);
    cal.forgetCalendar(acct.id);
    await mail.testLogin(acct);
    if (existing) Object.assign(existing, acct);
    else accounts.push(acct);
    saveAccounts(accounts);
    back('added=' + encodeURIComponent(email));
  } catch (err) {
    console.error(req.params.kind, 'callback', err.message);
    back('oauthError=' + encodeURIComponent(err.message));
  }
});

app.use('/api', (req, res, next) => (loggedIn(req) ? next() : res.status(401).json({ error: 'Not logged in' })));

app.get('/api/google', (req, res) => res.json({ configured: google.googleConfigured(), redirectUri: oauthRedirect(req, 'google') }));

app.post('/api/google', (req, res) => {
  const clientId = String(req.body.clientId || '').trim();
  const clientSecret = String(req.body.clientSecret || '').trim();
  if (!/\.apps\.googleusercontent\.com$/.test(clientId)) return res.status(400).json({ error: 'The Client ID should end with .apps.googleusercontent.com' });
  if (clientSecret.length < 10) return res.status(400).json({ error: 'Paste the Client secret too.' });
  google.setGoogleClient(clientId, clientSecret);
  res.json({ ok: true });
});

app.get('/api/microsoft', (req, res) => res.json({ configured: microsoft.microsoftConfigured(), redirectUri: oauthRedirect(req, 'microsoft') }));

app.post('/api/microsoft', (req, res) => {
  const clientId = String(req.body.clientId || '').trim();
  const clientSecret = String(req.body.clientSecret || '').trim();
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId)) {
    return res.status(400).json({ error: 'The Application (client) ID looks like 1234abcd-12ab-34cd-56ef-1234567890ab.' });
  }
  if (clientSecret.length < 30 || /\s/.test(clientSecret) || /^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(clientSecret)) {
    return res.status(400).json({ error: 'Paste the secret\'s Value (not the Secret ID).' });
  }
  microsoft.setMicrosoftClient(clientId, clientSecret);
  res.json({ ok: true });
});

// ---------- accounts ----------

const publicAccount = ({ password, calPassword, oauth, ...a }) => ({
  ...a, hasCalPassword: !!calPassword, calendar: !!calendarKind(a), calPasswordAllowed: !!caldavUrlFor(a),
});
const findAccount = (id) => {
  const a = accounts.find((x) => x.id === id);
  if (!a) throw Object.assign(new Error('Unknown account'), { status: 404 });
  return a;
};

app.get('/api/providers', (req, res) => {
  res.json(Object.entries(PROVIDERS).map(([id, p]) => ({ id, label: p.label, oauth: p.oauth || null })));
});

app.get('/api/accounts', (req, res) => res.json(accounts.map(publicAccount)));

app.post('/api/accounts', async (req, res, next) => {
  try {
    const { provider, email, password, name, color } = req.body;
    if (!PROVIDERS[provider] || PROVIDERS[provider].oauth) return res.status(400).json({ error: 'Pick a provider.' });
    if (!/^\S+@\S+\.\S+$/.test(email || '')) return res.status(400).json({ error: 'Enter a valid email address.' });
    if (!password) return res.status(400).json({ error: 'Enter the app password.' });
    const acct = {
      id: crypto.randomBytes(4).toString('hex'),
      provider, email: email.trim(), password: password.replace(/\s+/g, ''),
      name: (name || '').trim(), color: freeColor(color),
    };
    const count = await mail.testLogin(acct);
    accounts.push(acct);
    saveAccounts(accounts);
    res.json({ ...publicAccount(acct), messages: count });
  } catch (err) { next(err); }
});

app.patch('/api/accounts/:id', async (req, res, next) => {
  try {
    const a = findAccount(req.params.id);
    if (typeof req.body.name === 'string') a.name = req.body.name.trim();
    if (typeof req.body.color === 'string') a.color = req.body.color;
    if (typeof req.body.calPassword === 'string') {
      const calPassword = req.body.calPassword.replace(/\s+/g, '');
      await cal.testCalendar({ ...a, calPassword: calPassword || undefined });
      a.calPassword = calPassword || undefined;
      cal.forgetCalendar(a.id);
    }
    saveAccounts(accounts);
    res.json(publicAccount(a));
  } catch (err) { next(err); }
});

app.delete('/api/accounts/:id', (req, res, next) => {
  try {
    findAccount(req.params.id);
    accounts = accounts.filter((a) => a.id !== req.params.id);
    saveAccounts(accounts);
    mail.dropClient(req.params.id);
    cal.forgetCalendar(req.params.id);
    google.forgetToken(req.params.id);
    microsoft.forgetToken(req.params.id);
    cache.delete(req.params.id);
    junkCache.delete(req.params.id);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// ---------- feed ----------

function refresh(acct, store = cache, fetch = mail.fetchFeed) {
  const entry = store.get(acct.id) || {};
  if (entry.pending) return entry.pending;
  entry.pending = fetch(acct)
    .then((items) => Object.assign(entry, { at: Date.now(), items, error: null }))
    .catch((err) => Object.assign(entry, { at: Date.now(), error: err.message }))
    .finally(() => { entry.pending = null; });
  store.set(acct.id, entry);
  return entry.pending;
}
const refreshJunk = (acct) => refresh(acct, junkCache, mail.fetchJunk);

// Answer from memory, refreshing what's stale; wait only when there's nothing yet.
async function collect(req, store, refreshOne) {
  const force = req.query.refresh === '1';
  await Promise.all(accounts.map((a) => {
    const e = store.get(a.id);
    if (force || !e?.items) {
      // Nothing to show yet (or the refresh button): wait, but not forever.
      return Promise.race([refreshOne(a), new Promise((r) => setTimeout(r, force ? 8000 : 20000))]);
    }
    // Otherwise answer from memory now and refresh behind the scenes;
    // the page asks again shortly when something was still loading.
    if (Date.now() - e.at > STALE_MS) refreshOne(a);
  }));
  const items = [];
  const status = {};
  for (const a of accounts) {
    const e = store.get(a.id) || {};
    status[a.id] = { error: e.error || null, at: e.at || null, loading: !!e.pending };
    for (const it of e.items || []) items.push(it);
  }
  items.sort((x, y) => (x.date < y.date ? 1 : -1));
  return { items, status };
}

app.get('/api/feed', async (req, res) => {
  const data = await collect(req, cache, (a) => refresh(a));
  // How much unread junk each account has, so the page can point at it.
  const junk = {};
  for (const a of accounts) {
    const items = junkCache.get(a.id)?.items || [];
    junk[a.id] = { count: items.length, unread: items.filter((i) => !i.seen).length };
  }
  res.json({ ...data, junk });
});

app.get('/api/junk', async (req, res) => {
  res.json(await collect(req, junkCache, refreshJunk));
});

const boxOf = (req) => (['sent', 'all', 'junk'].includes(req.query.box) ? req.query.box : 'inbox');

app.get('/api/search', async (req, res) => {
  const q = String(req.query.q || '').trim().slice(0, 200);
  if (!q) return res.json({ items: [], status: {} });
  const which = req.query.acct ? accounts.filter((a) => a.id === req.query.acct) : accounts;
  const items = [];
  const status = {};
  await Promise.all(which.map(async (a) => {
    try {
      const found = await Promise.race([
        mail.search(a, q),
        new Promise((_, rej) => setTimeout(() => rej(new Error('Search took too long')), 25000)),
      ]);
      items.push(...found);
      status[a.id] = { count: found.length };
    } catch (err) {
      status[a.id] = { error: err.message };
    }
  }));
  items.sort((x, y) => (x.date < y.date ? 1 : -1));
  res.json({ items, status });
});

app.get('/api/thread/:acct/:key', async (req, res, next) => {
  try {
    res.json(await mail.getThread(findAccount(req.params.acct), req.params.key));
  } catch (err) { next(err); }
});

function patchCached(acctId, uid, fn, box = 'inbox') {
  const it = (box === 'junk' ? junkCache : cache).get(acctId)?.items?.find((x) => String(x.uid) === String(uid));
  if (it) fn(it);
}
function dropCached(acctId, uid, box = 'inbox') {
  const e = (box === 'junk' ? junkCache : cache).get(acctId);
  if (e?.items) e.items = e.items.filter((x) => String(x.uid) !== String(uid));
}

app.get('/api/msg/:acct/:uid', async (req, res, next) => {
  try {
    const box = boxOf(req);
    const msg = await mail.getMessage(findAccount(req.params.acct), req.params.uid, box);
    if (box === 'inbox' || box === 'junk') patchCached(req.params.acct, req.params.uid, (it) => { it.seen = true; }, box);
    res.json(msg);
  } catch (err) { next(err); }
});

app.get('/api/msg/:acct/:uid/att/:idx', async (req, res, next) => {
  try {
    const a = await mail.getAttachment(findAccount(req.params.acct), req.params.uid, req.params.idx, boxOf(req));
    res.set('Content-Type', a.contentType || 'application/octet-stream');
    res.set('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(a.filename || 'attachment')}`);
    res.send(a.content);
  } catch (err) { next(err); }
});

app.get('/api/sent/:acct/:uid', async (req, res, next) => {
  try {
    res.json(await mail.getSentReply(findAccount(req.params.acct), req.params.uid));
  } catch (err) { next(err); }
});

app.post('/api/msg/:acct/:uid/seen', async (req, res, next) => {
  try {
    const seen = !!req.body.seen;
    const box = boxOf(req);
    await mail.setSeen(findAccount(req.params.acct), req.params.uid, seen, box);
    patchCached(req.params.acct, req.params.uid, (it) => { it.seen = seen; }, box);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.delete('/api/msg/:acct/:uid', async (req, res, next) => {
  try {
    const box = boxOf(req);
    await mail.deleteMessage(findAccount(req.params.acct), req.params.uid, box);
    dropCached(req.params.acct, req.params.uid, box);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// Move a message from Junk back to the inbox.
app.post('/api/msg/:acct/:uid/notjunk', async (req, res, next) => {
  try {
    const acct = findAccount(req.params.acct);
    await mail.notJunk(acct, req.params.uid);
    dropCached(acct.id, req.params.uid, 'junk');
    refresh(acct);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// Turn "Anna <a@b.com>, c@d.fr" into address objects; refuse anything malformed.
function parseAddresses(text, label) {
  const list = addressparser(String(text || '')).flatMap((a) => (a.group ? a.group : [a]));
  for (const a of list) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a.address || '')) {
      throw Object.assign(new Error(`"${a.address || a.name}" in ${label} isn't a valid email address.`), { status: 400 });
    }
  }
  return list.map((a) => ({ name: a.name || '', address: a.address }));
}

function readAttachments(body) {
  const attachments = (Array.isArray(body.attachments) ? body.attachments : []).map((a) => ({
    filename: String(a.filename || 'file'),
    contentType: String(a.contentType || 'application/octet-stream'),
    content: Buffer.from(String(a.data || ''), 'base64'),
  }));
  if (attachments.reduce((n, a) => n + a.content.length, 0) > MAX_ATTACH) {
    throw Object.assign(new Error('Attachments are too big (20 MB max in total).'), { status: 413 });
  }
  return attachments;
}

app.post('/api/compose', async (req, res, next) => {
  try {
    const acct = findAccount(String(req.body.acct || ''));
    const to = parseAddresses(req.body.to, 'To');
    const cc = parseAddresses(req.body.cc, 'Cc');
    const bcc = parseAddresses(req.body.bcc, 'Bcc');
    if (!to.length && !cc.length && !bcc.length) return res.status(400).json({ error: 'Add at least one recipient.' });
    const subject = String(req.body.subject || '').trim();
    const body = String(req.body.body || '');
    if (!subject && !body.trim()) return res.status(400).json({ error: 'Write a subject or a message first.' });
    await mail.sendNew(acct, { to, cc, bcc, subject, body, attachments: readAttachments(req.body) });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.post('/api/msg/:acct/:uid/reply', async (req, res, next) => {
  try {
    const body = String(req.body.body || '').trim();
    if (!body) return res.status(400).json({ error: 'Write something first.' });
    const attachments = readAttachments(req.body);
    const box = boxOf(req);
    await mail.sendReply(findAccount(req.params.acct), req.params.uid, { body, all: !!req.body.all, attachments, box });
    if (box === 'inbox') patchCached(req.params.acct, req.params.uid, (it) => { it.answered = true; it.replied = { at: new Date().toISOString(), uid: null }; });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// ---------- the Pi itself ----------
// The Pi has no power button: let the owner shut it down cleanly before unplugging.

app.post('/api/system/:action', (req, res) => {
  const cmd = { poweroff: 'poweroff', reboot: 'reboot' }[req.params.action];
  if (!cmd) return res.status(404).json({ error: 'Unknown action' });
  res.json({ ok: true });
  setTimeout(() => execFile('sudo', ['-n', 'systemctl', cmd], (err) => { if (err) console.error('system', cmd, err.message); }), 1000);
});

// ---------- calendar ----------

app.get('/api/calendar', async (req, res, next) => {
  try {
    const from = new Date(req.query.from), to = new Date(req.query.to);
    if (isNaN(from) || isNaN(to) || to <= from || to - from > 62 * 864e5) return res.status(400).json({ error: 'Bad date range' });
    const events = [];
    const status = {};
    await Promise.all(accounts.filter((a) => calendarKind(a)).map(async (a) => {
      try {
        events.push(...await cal.fetchEvents(a, from, to));
        status[a.id] = { ok: true };
      } catch (err) {
        status[a.id] = { error: err.message };
      }
    }));
    events.sort((x, y) => (x.start < y.start ? -1 : 1));
    // The same event can come from two accounts (shared calendars, invites
    // between your own addresses): show it once.
    const seen = new Set();
    const unique = events.filter((ev) => {
      const k = `${ev.uid}|${ev.start}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
    res.json({ events: unique, status });
  } catch (err) { next(err); }
});

// Add an event by hand (e.g. from a date mentioned in an email) or remove one.
app.post('/api/events', async (req, res, next) => {
  try {
    const acct = findAccount(String(req.body.acct || ''));
    if (!calendarKind(acct)) return res.status(400).json({ error: 'This account has no calendar.' });
    const b = req.body;
    const title = String(b.title || '').trim().slice(0, 300);
    if (!title) return res.status(400).json({ error: 'Give the event a title.' });
    const ymdOk = (d) => /^\d{4}-\d{2}-\d{2}$/.test(String(d || ''));
    const ev = { title, allDay: !!b.allDay, location: String(b.location || '').slice(0, 300), notes: String(b.notes || '').slice(0, 2000) };
    if (ev.allDay) {
      if (!ymdOk(b.date) || !ymdOk(b.endDate) || b.endDate <= b.date) return res.status(400).json({ error: 'Check the dates.' });
      Object.assign(ev, { date: b.date, endDate: b.endDate });
    } else {
      const s0 = new Date(b.start), e0 = new Date(b.end);
      if (isNaN(s0) || isNaN(e0) || e0 <= s0) return res.status(400).json({ error: 'The end must be after the start.' });
      Object.assign(ev, { start: s0.toISOString(), end: e0.toISOString() });
    }
    res.json({ ok: true, calendar: await cal.createEvent(acct, ev) });
  } catch (err) { next(err); }
});

app.delete('/api/events', async (req, res, next) => {
  try {
    const acct = findAccount(String(req.body.acct || ''));
    if (!req.body.ref) return res.status(400).json({ error: 'Missing event' });
    await cal.deleteEvent(acct, { ref: String(req.body.ref), etag: req.body.etag ? String(req.body.etag) : undefined });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// An invitation that only informs (no RSVP): just put it in the calendar.
app.post('/api/msg/:acct/:uid/addinvite', async (req, res, next) => {
  try {
    const acct = findAccount(req.params.acct);
    const { ics } = await mail.getInviteIcs(acct, req.params.uid);
    res.json({ ok: true, savedTo: await cal.addToCalendar(acct, ics, 'ACCEPTED') });
  } catch (err) { next(err); }
});

const RSVP = { ACCEPTED: 'Accepted', TENTATIVE: 'Tentative', DECLINED: 'Declined' };

app.post('/api/msg/:acct/:uid/rsvp', async (req, res, next) => {
  try {
    const partstat = String(req.body.response || '').toUpperCase();
    if (!RSVP[partstat]) return res.status(400).json({ error: 'Bad response' });
    const acct = findAccount(req.params.acct);
    if (acct.provider === 'outlook') {
      return res.json({ ok: true, savedTo: await outlook.rsvp(acct, req.params.uid, partstat), calendarError: null });
    }
    const { ics, parsed } = await mail.getInviteIcs(acct, req.params.uid);
    if (await cal.rsvpViaProvider(acct, ics, partstat).catch(() => false)) {
      return res.json({ ok: true, savedTo: 'Google Calendar', calendarError: null });
    }
    await mail.sendInviteReply(acct, parsed, cal.buildReply(ics, acct.email, partstat), RSVP[partstat]);
    let savedTo = null, calendarError = null;
    if (partstat !== 'DECLINED' && calendarKind(acct)) {
      try { savedTo = await cal.addToCalendar(acct, ics, partstat); } catch (err) { calendarError = err.message; }
    }
    res.json({ ok: true, savedTo, calendarError });
  } catch (err) { next(err); }
});

app.use('/api', (err, req, res, next) => {
  console.error(req.method, req.path, err.message);
  res.status(err.status || 500).json({ error: err.message || 'Something went wrong' });
});

const PUBLIC = new URL('./public', import.meta.url).pathname;

// The page's version (its last change), so open apps notice updates.
app.get('/version', (req, res) => {
  res.set('Cache-Control', 'no-store');
  res.json({ v: String(fs.statSync(PUBLIC + '/index.html').mtimeMs) });
});

app.use(express.static(PUBLIC, {
  index: 'index.html',
  // Always fetch the page itself fresh; installed phone apps otherwise keep old copies.
  setHeaders: (res, path) => { if (path.endsWith('.html') || path.endsWith('sw.js')) res.set('Cache-Control', 'no-cache, no-store'); },
}));

app.listen(PORT, HOST, () => console.log(`Feed Mail on http://${HOST}:${PORT}`));

// Keep every inbox and calendar warm in the background so the app is instant.
const warm = () => accounts.forEach((a) => refresh(a));
// Junk changes less and matters less: every 2 minutes, after the inboxes.
const warmJunk = () => accounts.forEach((a) => (cache.get(a.id)?.pending || Promise.resolve()).finally(() => refreshJunk(a)));
const warmCalendars = () => accounts.filter((a) => calendarKind(a)).forEach((a) => cal.warmCalendar(a));
setTimeout(() => { warm(); warmCalendars(); warmJunk(); }, 2000);
setInterval(warm, 60 * 1000);
setInterval(warmJunk, 2 * 60 * 1000);
setInterval(warmCalendars, 5 * 60 * 1000);
