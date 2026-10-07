import crypto from 'node:crypto';
import express from 'express';
import { loadAccounts, saveAccounts, loadConfig, saveConfig } from './store.js';
import { PROVIDERS, caldavUrlFor } from './providers.js';
import * as mail from './mail.js';
import * as cal from './calendar.js';

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '127.0.0.1';
const STALE_MS = 45 * 1000;
const SESSION_DAYS = 90;

const config = loadConfig();
let accounts = loadAccounts();
const cache = new Map(); // account id -> { at, items, error, pending }

const app = express();
app.disable('x-powered-by');
app.use(express.json({ limit: '1mb' }));
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

app.use('/api', (req, res, next) => (loggedIn(req) ? next() : res.status(401).json({ error: 'Not logged in' })));

// ---------- accounts ----------

const publicAccount = ({ password, calPassword, ...a }) => ({ ...a, hasCalPassword: !!calPassword, calendar: !!caldavUrlFor(a) });
const findAccount = (id) => {
  const a = accounts.find((x) => x.id === id);
  if (!a) throw Object.assign(new Error('Unknown account'), { status: 404 });
  return a;
};

app.get('/api/providers', (req, res) => {
  res.json(Object.entries(PROVIDERS).map(([id, p]) => ({ id, label: p.label })));
});

app.get('/api/accounts', (req, res) => res.json(accounts.map(publicAccount)));

app.post('/api/accounts', async (req, res, next) => {
  try {
    const { provider, email, password, name, color } = req.body;
    if (!PROVIDERS[provider]) return res.status(400).json({ error: 'Pick a provider.' });
    if (!/^\S+@\S+\.\S+$/.test(email || '')) return res.status(400).json({ error: 'Enter a valid email address.' });
    if (!password) return res.status(400).json({ error: 'Enter the app password.' });
    const acct = {
      id: crypto.randomBytes(4).toString('hex'),
      provider, email: email.trim(), password: password.replace(/\s+/g, ''),
      name: (name || '').trim(), color: color || '#4f7cff',
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
    cache.delete(req.params.id);
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// ---------- feed ----------

function refresh(acct) {
  const entry = cache.get(acct.id) || {};
  if (entry.pending) return entry.pending;
  entry.pending = mail.fetchFeed(acct)
    .then((items) => Object.assign(entry, { at: Date.now(), items, error: null }))
    .catch((err) => Object.assign(entry, { at: Date.now(), error: err.message }))
    .finally(() => { entry.pending = null; });
  cache.set(acct.id, entry);
  return entry.pending;
}

app.get('/api/feed', async (req, res) => {
  const force = req.query.refresh === '1';
  await Promise.all(accounts.map((a) => {
    const e = cache.get(a.id);
    if (force || !e || !e.items || Date.now() - e.at > STALE_MS) {
      // Don't make the page wait forever for one slow server.
      return Promise.race([refresh(a), new Promise((r) => setTimeout(r, 20000))]);
    }
  }));
  const items = [];
  const status = {};
  for (const a of accounts) {
    const e = cache.get(a.id) || {};
    status[a.id] = { error: e.error || null, at: e.at || null, loading: !!e.pending };
    for (const it of e.items || []) items.push(it);
  }
  items.sort((x, y) => (x.date < y.date ? 1 : -1));
  res.json({ items, status });
});

function patchCached(acctId, uid, fn) {
  const it = cache.get(acctId)?.items?.find((x) => x.uid === Number(uid));
  if (it) fn(it);
}

app.get('/api/msg/:acct/:uid', async (req, res, next) => {
  try {
    const msg = await mail.getMessage(findAccount(req.params.acct), req.params.uid);
    patchCached(req.params.acct, req.params.uid, (it) => { it.seen = true; });
    res.json(msg);
  } catch (err) { next(err); }
});

app.get('/api/msg/:acct/:uid/att/:idx', async (req, res, next) => {
  try {
    const a = await mail.getAttachment(findAccount(req.params.acct), req.params.uid, req.params.idx);
    res.set('Content-Type', a.contentType || 'application/octet-stream');
    res.set('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(a.filename || 'attachment')}`);
    res.send(a.content);
  } catch (err) { next(err); }
});

app.post('/api/msg/:acct/:uid/seen', async (req, res, next) => {
  try {
    const seen = !!req.body.seen;
    await mail.setSeen(findAccount(req.params.acct), req.params.uid, seen);
    patchCached(req.params.acct, req.params.uid, (it) => { it.seen = seen; });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.delete('/api/msg/:acct/:uid', async (req, res, next) => {
  try {
    await mail.deleteMessage(findAccount(req.params.acct), req.params.uid);
    const e = cache.get(req.params.acct);
    if (e?.items) e.items = e.items.filter((x) => x.uid !== Number(req.params.uid));
    res.json({ ok: true });
  } catch (err) { next(err); }
});

app.post('/api/msg/:acct/:uid/reply', async (req, res, next) => {
  try {
    const body = String(req.body.body || '').trim();
    if (!body) return res.status(400).json({ error: 'Write something first.' });
    await mail.sendReply(findAccount(req.params.acct), req.params.uid, { body, all: !!req.body.all });
    patchCached(req.params.acct, req.params.uid, (it) => { it.answered = true; });
    res.json({ ok: true });
  } catch (err) { next(err); }
});

// ---------- calendar ----------

app.get('/api/calendar', async (req, res, next) => {
  try {
    const from = new Date(req.query.from), to = new Date(req.query.to);
    if (isNaN(from) || isNaN(to) || to <= from || to - from > 62 * 864e5) return res.status(400).json({ error: 'Bad date range' });
    const events = [];
    const status = {};
    await Promise.all(accounts.filter((a) => caldavUrlFor(a)).map(async (a) => {
      try {
        events.push(...await cal.fetchEvents(a, from, to));
        status[a.id] = { ok: true };
      } catch (err) {
        status[a.id] = { error: err.message };
      }
    }));
    events.sort((x, y) => (x.start < y.start ? -1 : 1));
    res.json({ events, status });
  } catch (err) { next(err); }
});

const RSVP = { ACCEPTED: 'Accepted', TENTATIVE: 'Tentative', DECLINED: 'Declined' };

app.post('/api/msg/:acct/:uid/rsvp', async (req, res, next) => {
  try {
    const partstat = String(req.body.response || '').toUpperCase();
    if (!RSVP[partstat]) return res.status(400).json({ error: 'Bad response' });
    const acct = findAccount(req.params.acct);
    const { ics, parsed } = await mail.getInviteIcs(acct, req.params.uid);
    await mail.sendInviteReply(acct, parsed, cal.buildReply(ics, acct.email, partstat), RSVP[partstat]);
    let savedTo = null, calendarError = null;
    if (partstat !== 'DECLINED' && caldavUrlFor(acct)) {
      try { savedTo = await cal.addToCalendar(acct, ics, partstat); } catch (err) { calendarError = err.message; }
    }
    res.json({ ok: true, savedTo, calendarError });
  } catch (err) { next(err); }
});

app.use('/api', (err, req, res, next) => {
  console.error(req.method, req.path, err.message);
  res.status(err.status || 500).json({ error: err.message || 'Something went wrong' });
});

app.use(express.static(new URL('./public', import.meta.url).pathname, { index: 'index.html' }));

app.listen(PORT, HOST, () => console.log(`Feed Mail on http://${HOST}:${PORT}`));
