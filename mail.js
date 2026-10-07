// Talks to the mail servers. Messages are fetched on demand and kept only in
// memory; nothing is saved to disk.
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import MailComposer from 'nodemailer/lib/mail-composer/index.js';
import { serversFor } from './providers.js';
import { parseInvite, findDates } from './calendar.js';
import { accessToken } from './google.js';

const FEED_SIZE = 50; // newest messages per account
const PREVIEW_BYTES = 20000; // enough of each message to build a preview

const clients = new Map(); // account id -> Promise<ImapFlow>

async function newClient(acct) {
  const { imap } = serversFor(acct);
  const c = new ImapFlow({
    ...imap,
    auth: acct.oauth
      ? { user: acct.email, accessToken: await accessToken(acct) }
      : { user: acct.email, pass: acct.password },
    logger: false,
    emitLogs: false,
  });
  c.on('error', () => {}); // handled through the 'close' / usable checks
  return c;
}

async function getClient(acct) {
  const pending = clients.get(acct.id);
  if (pending) {
    const c = await pending.catch(() => null);
    if (c && c.usable) return c;
    clients.delete(acct.id);
  }
  const p = (async () => {
    const c = await newClient(acct);
    await c.connect();
    c.on('close', () => {
      if (clients.get(acct.id) === p) clients.delete(acct.id);
    });
    return c;
  })();
  clients.set(acct.id, p);
  try {
    return await p;
  } catch (err) {
    clients.delete(acct.id);
    throw friendly(err);
  }
}

export function dropClient(acctId) {
  known.delete(acctId);
  sentCache.delete(acctId);
  const p = clients.get(acctId);
  clients.delete(acctId);
  p?.then((c) => c.logout().catch(() => c.close())).catch(() => {});
}

// Some servers (Yahoo) never tell a long-open session about new mail, so
// re-select INBOX when the view may be stale.
const RESELECT_MS = 15 * 1000;

// Run fn with a mailbox selected; retry once on a dead connection.
async function withMailbox(acct, path, fn, { fresh = false } = {}, retried = false) {
  const c = await getClient(acct);
  let lock;
  try {
    lock = await c.getMailboxLock(path);
    // Switching mailbox already re-selects; otherwise refresh a stale view.
    if (c.fmSelectedPath === path && (fresh || Date.now() - c.fmSelectedAt > RESELECT_MS)) {
      await c.mailboxOpen(path);
    }
    c.fmSelectedPath = path;
    c.fmSelectedAt = Date.now();
    return await fn(c);
  } catch (err) {
    if (!retried && !c.usable) {
      clients.delete(acct.id);
      return withMailbox(acct, path, fn, { fresh }, true);
    }
    throw friendly(err);
  } finally {
    lock?.release();
  }
}

const withInbox = (acct, fn, opts) => withMailbox(acct, 'INBOX', fn, opts);

// ---------- Sent mail: replies and conversations ----------
// Scan recent Sent mail so replies made anywhere (phone app, webmail) are
// recognised, and so my own messages can be shown inside conversations.

const SENT_SCAN = 300;
const SENT_STALE_MS = 60 * 1000;
const sentCache = new Map(); // account id -> { at, map, items }

const isGmail = (acct) => acct.provider === 'gmail';
const msgIds = (s) => String(s || '').match(/<[^>]+>/g) || [];

function refsFrom(headers) {
  // Raw header block -> Message-IDs listed in References (may be folded).
  const m = String(headers || '').replace(/\r?\n[ \t]+/g, ' ').match(/^references:(.*)$/im);
  return m ? msgIds(m[1]) : [];
}

async function sentPath(acct) {
  const c = await getClient(acct);
  if (c.fmSentPath === undefined) c.fmSentPath = (await findSpecial(c, '\\Sent')) || null;
  return c.fmSentPath;
}

async function boxPath(acct, box) {
  if (box !== 'sent') return 'INBOX';
  const path = await sentPath(acct);
  if (!path) throw Object.assign(new Error('No Sent folder found.'), { status: 404 });
  return path;
}

async function sentState(acct) {
  const hit = sentCache.get(acct.id);
  if (hit && Date.now() - hit.at < SENT_STALE_MS) return hit;
  const map = new Map();
  const items = [];
  const path = await sentPath(acct);
  if (path) {
    await withMailbox(acct, path, async (c) => {
      const exists = c.mailbox.exists;
      if (!exists) return;
      for await (const m of c.fetch(`${Math.max(1, exists - SENT_SCAN + 1)}:*`, {
        uid: true, envelope: true, internalDate: true, headers: ['references'], ...(isGmail(acct) && { threadId: true }),
      })) {
        const env = m.envelope || {};
        const date = (env.date || m.internalDate || new Date()).toISOString();
        const inReplyTo = msgIds(env.inReplyTo);
        for (const id of inReplyTo) {
          const prev = map.get(id);
          if (!prev || prev.at < date) map.set(id, { at: date, uid: m.uid });
        }
        // Keep previews already fetched for this message.
        const old = hit?.items.find((x) => x.uid === m.uid);
        items.push({
          box: 'sent', uid: m.uid, date, messageId: env.messageId || null, inReplyTo, refs: refsFrom(m.headers),
          gThread: m.threadId || null, from: addr(env.from?.[0]), to: (env.to || []).map(addr),
          subject: env.subject || '(no subject)', preview: old?.preview,
        });
      }
    });
  }
  const state = { at: Date.now(), map, items };
  sentCache.set(acct.id, state);
  return state;
}

// Group messages into conversations: Gmail's own thread id, otherwise the
// Message-ID / In-Reply-To / References links (union-find).
function assignThreads(items) {
  const parent = new Map();
  const add = (x) => { if (!parent.has(x)) parent.set(x, x); };
  const find = (x) => {
    while (parent.get(x) !== x) { parent.set(x, parent.get(parent.get(x))); x = parent.get(x); }
    return x;
  };
  const union = (a, b) => {
    add(a); add(b);
    const ra = find(a), rb = find(b);
    if (ra !== rb) ra < rb ? parent.set(rb, ra) : parent.set(ra, rb);
  };
  const own = (it) => it.messageId || `${it.box}:${it.uid}`;
  for (const it of items) {
    if (it.gThread) continue;
    add(own(it));
    for (const r of [...(it.refs || []), ...(it.inReplyTo || [])]) union(own(it), r);
  }
  for (const it of items) it.thread = it.gThread ? `g${it.gThread}` : `m${find(own(it))}`;
}

export async function getThread(acct, key) {
  const inbox = [...(known.get(acct.id)?.byUid.values() || [])].filter((i) => i.thread === key);
  const sent = (sentCache.get(acct.id)?.items || []).filter((i) => i.thread === key);
  const need = sent.filter((i) => i.preview === undefined);
  if (need.length) {
    const path = await sentPath(acct);
    const fetched = await withMailbox(acct, path, async (c) => {
      const out = [];
      for await (const m of c.fetch(need.map((i) => i.uid).join(','), { uid: true, source: { maxLength: PREVIEW_BYTES } }, { uid: true })) out.push(m);
      return out;
    });
    for (const m of fetched) {
      const it = need.find((x) => x.uid === m.uid);
      try {
        const p = await simpleParser(m.source, { skipImageLinks: true, skipTextToHtml: true });
        it.preview = stripQuote(p.text || '').replace(/\[?(https?:\/\/|mailto:)\S+\]?/g, '').replace(/\s+/g, ' ').trim().slice(0, 240);
      } catch { it.preview = ''; }
    }
  }
  const pick = (i, box) => ({
    box, uid: i.uid, date: i.date, from: i.from, to: i.to, subject: i.subject, preview: i.preview || '',
    seen: box === 'sent' ? true : !!i.seen, attach: !!i.attach, acct: acct.id,
  });
  return [...inbox.map((i) => pick(i, 'inbox')), ...sent.map((i) => pick(i, 'sent'))]
    .sort((a, b) => (a.date < b.date ? -1 : 1));
}

// Keep only what I wrote, not the quoted original below it.
function stripQuote(text) {
  const out = [];
  const lines = text.split('\n');
  const intro = /^(On .+wrote:|Le .+a écrit\s*:|-----Original Message-----|-{2,} ?Original)/i;
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    // Gmail often wraps "On <date> <name> wrote:" over two lines.
    if (intro.test(t) || intro.test(`${t} ${(lines[i + 1] || '').trim()}`)) break;
    if (/^>/.test(t)) continue;
    out.push(lines[i]);
  }
  return out.join('\n').trim();
}

export async function getSentReply(acct, uid) {
  const msg = await getMessage(acct, uid, 'sent');
  return { date: msg.date, to: msg.to, text: stripQuote(msg.text || '') };
}

async function smtpTransport(acct) {
  const { smtp } = serversFor(acct);
  const auth = acct.oauth
    ? { type: 'OAuth2', user: acct.email, accessToken: await accessToken(acct) }
    : { user: acct.email, pass: acct.password };
  return nodemailer.createTransport({ ...smtp, auth });
}

function friendly(err) {
  if (err.authenticationFailed || /AUTHENTICATIONFAILED|Invalid credentials|LOGIN failed/i.test(err.responseText || err.message)) {
    const e = new Error('Login refused: check the email address and app password, and that IMAP is switched on.');
    e.status = 401;
    return e;
  }
  return err;
}

export async function testLogin(acct) {
  const c = await newClient(acct);
  try {
    await c.connect();
    const status = await c.status('INBOX', { messages: true });
    await c.logout();
    return status.messages;
  } catch (err) {
    c.close();
    throw friendly(err);
  }
}

function makePreview(parsed) {
  const text = parsed.text || '';
  return text
    .replace(/\[?(https?:\/\/|mailto:)\S+\]?/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 240);
}

function hasAttachment(node) {
  if (!node) return false;
  if (node.disposition === 'attachment') return true;
  return (node.childNodes || []).some(hasAttachment);
}

function addr(a) {
  return a ? { name: a.name || '', address: a.address || '' } : null;
}

// Per account: what we already know about each message (headers + preview),
// so a refresh only downloads new mail. Memory only.
const known = new Map(); // account id -> { validity, byUid: Map(uid -> item) }

export async function fetchFeed(acct) {
  const { flags, fresh, validity } = await withInbox(acct, async (c) => {
    const exists = c.mailbox.exists;
    const validity = String(c.mailbox.uidValidity);
    if (!exists) return { flags: [], fresh: [], validity };
    const start = Math.max(1, exists - FEED_SIZE + 1);
    const flags = [];
    for await (const m of c.fetch(`${start}:*`, { uid: true, flags: true })) flags.push(m);

    const k = known.get(acct.id);
    const byUid = k && k.validity === validity ? k.byUid : new Map();
    const missing = flags.filter((m) => !byUid.has(m.uid)).map((m) => m.uid);
    const fresh = [];
    if (missing.length) {
      for await (const m of c.fetch(missing.join(','), {
        uid: true, envelope: true, internalDate: true, bodyStructure: true, headers: ['references'],
        source: { maxLength: PREVIEW_BYTES }, ...(isGmail(acct) && { threadId: true }),
      }, { uid: true })) fresh.push(m);
    }
    return { flags, fresh, validity };
  }, { fresh: true });

  const prev = known.get(acct.id);
  const byUid = prev && prev.validity === validity ? prev.byUid : new Map();
  for (const m of fresh) {
    let preview = '';
    try {
      preview = makePreview(await simpleParser(m.source, { skipImageLinks: true, skipTextToHtml: true }));
    } catch {}
    const env = m.envelope || {};
    byUid.set(m.uid, {
      id: `${acct.id}:${m.uid}`,
      acct: acct.id,
      uid: m.uid,
      date: (env.date || m.internalDate || new Date()).toISOString(),
      from: addr(env.from?.[0]),
      to: (env.to || []).map(addr),
      subject: env.subject || '(no subject)',
      messageId: env.messageId || null,
      inReplyTo: msgIds(env.inReplyTo),
      refs: refsFrom(m.headers),
      gThread: m.threadId || null,
      attach: hasAttachment(m.bodyStructure),
      preview,
    });
  }
  // Forget messages that dropped out of the window (deleted, moved, or old).
  const current = new Set(flags.map((m) => m.uid));
  for (const uid of byUid.keys()) if (!current.has(uid)) byUid.delete(uid);
  known.set(acct.id, { validity, byUid });

  let sent = { map: new Map(), items: [] };
  try { sent = await sentState(acct); } catch {}

  const inbox = flags.filter((m) => byUid.has(m.uid)).map((m) => {
    const base = byUid.get(m.uid);
    base.box = 'inbox';
    base.seen = m.flags?.has('\\Seen') || false;
    base.answered = m.flags?.has('\\Answered') || false;
    base.flagged = m.flags?.has('\\Flagged') || false;
    return base;
  });
  assignThreads([...inbox, ...sent.items]);

  // Per conversation: how many of my own messages, and the latest one.
  const mine = new Map();
  for (const s of sent.items) {
    const t = mine.get(s.thread) || { count: 0, lastAt: null };
    t.count++;
    if (!t.lastAt || s.date > t.lastAt) t.lastAt = s.date;
    mine.set(s.thread, t);
  }

  return inbox.map((base) => {
    const { refs, inReplyTo, gThread, ...item } = base;
    return {
      ...item,
      replied: sent.map.get(base.messageId) || (base.answered ? { at: null, uid: null } : null),
      threadSent: mine.get(base.thread) || null,
    };
  });
}

async function fetchParsed(acct, uid, c) {
  const m = await c.fetchOne(String(uid), { source: true, flags: true }, { uid: true });
  if (!m) {
    const e = new Error('Message not found (it may have been moved or deleted).');
    e.status = 404;
    throw e;
  }
  return { parsed: await simpleParser(m.source), flags: m.flags };
}

const list = (v) => (v ? v.value || [] : []).map(addr);

export async function getMessage(acct, uid, box = 'inbox') {
  const { parsed } = await withMailbox(acct, await boxPath(acct, box), async (c) => {
    const r = await fetchParsed(acct, uid, c);
    if (box === 'inbox' && !r.flags?.has('\\Seen')) await c.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
    return r;
  });

  let html = parsed.html || null;
  if (html) {
    // Inline embedded (cid:) images as data URIs so they display.
    let budget = 8 * 1024 * 1024;
    for (const a of parsed.attachments) {
      if (!a.contentId || a.size > budget) continue;
      const cid = a.contentId.replace(/^<|>$/g, '');
      const ref = `cid:${cid}`;
      if (!html.includes(ref)) continue;
      budget -= a.size;
      html = html.split(ref).join(`data:${a.contentType};base64,${a.content.toString('base64')}`);
    }
  }

  const ics = findIcs(parsed);
  const invite = ics ? parseInvite(ics) : null;

  let replied = null;
  if (box === 'inbox') try { replied = (await sentState(acct)).map.get(parsed.messageId) || null; } catch {}

  return {
    acct: acct.id,
    uid: Number(uid),
    box,
    messageId: parsed.messageId || null,
    replied,
    invite: box === 'inbox' ? invite : null,
    dates: findDates(parsed.subject, parsed.text, parsed.date || new Date()),
    subject: parsed.subject || '(no subject)',
    date: (parsed.date || new Date()).toISOString(),
    from: list(parsed.from)[0] || null,
    to: list(parsed.to),
    cc: list(parsed.cc),
    html,
    text: parsed.text || '',
    attachments: parsed.attachments
      .map((a, i) => ({ idx: i, filename: a.filename || 'attachment', size: a.size, contentType: a.contentType, related: a.related }))
      .filter((a) => !a.related),
  };
}

// The calendar invite (text/calendar part or .ics file) inside a message, if any.
function findIcs(parsed) {
  const a = parsed.attachments.find((x) => /text\/calendar/i.test(x.contentType) || /\.ics$/i.test(x.filename || ''));
  return a ? a.content.toString('utf8') : null;
}

export async function getInviteIcs(acct, uid) {
  const { parsed } = await withInbox(acct, (c) => fetchParsed(acct, uid, c));
  const ics = findIcs(parsed);
  if (!ics) throw Object.assign(new Error('This email has no invitation.'), { status: 404 });
  return { ics, parsed };
}

export async function sendInviteReply(acct, parsed, icsReply, label) {
  const servers = serversFor(acct);
  const invite = parseInvite(icsReply);
  const to = invite?.organizer?.address || list(parsed.replyTo)[0]?.address || list(parsed.from)[0]?.address;
  const mail = {
    from: acct.name ? { name: acct.name, address: acct.email } : acct.email,
    to,
    subject: `${label}: ${invite?.title || parsed.subject || ''}`,
    text: `${label}: ${invite?.title || ''}`,
    icalEvent: { method: 'REPLY', content: icsReply },
  };
  const raw = await new MailComposer(mail).compile().build();
  const transport = await smtpTransport(acct);
  try {
    await transport.sendMail({ envelope: { from: acct.email, to: [to] }, raw });
  } catch (err) {
    throw friendly(err);
  }
}

export async function getAttachment(acct, uid, idx, box = 'inbox') {
  const { parsed } = await withMailbox(acct, await boxPath(acct, box), (c) => fetchParsed(acct, uid, c));
  const a = parsed.attachments[Number(idx)];
  if (!a) {
    const e = new Error('Attachment not found');
    e.status = 404;
    throw e;
  }
  return a;
}

async function findSpecial(c, use) {
  const boxes = await c.list();
  return boxes.find((b) => b.specialUse === use)?.path;
}

export async function setSeen(acct, uid, seen) {
  await withInbox(acct, (c) =>
    seen
      ? c.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true })
      : c.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true }));
}

export async function deleteMessage(acct, uid) {
  await withInbox(acct, async (c) => {
    const trash = await findSpecial(c, '\\Trash');
    if (trash) await c.messageMove(String(uid), trash, { uid: true });
    else await c.messageDelete(String(uid), { uid: true });
  });
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
const fmtAddr = (a) => (a.name ? `${a.name} <${a.address}>` : a.address);

// Who a reply goes to. Replying to my own sent message is a follow-up to
// the same people.
export function replyRecipients(acct, parsed, box, all) {
  const me = acct.email.toLowerCase();
  const same = (a, b) => a.address?.toLowerCase() === b.address?.toLowerCase();
  const notMe = (a) => a.address && a.address.toLowerCase() !== me;
  let to, rest;
  if (box === 'sent') {
    to = list(parsed.to).filter(notMe);
    rest = list(parsed.cc);
  } else {
    to = list(parsed.replyTo).length ? list(parsed.replyTo) : list(parsed.from);
    rest = [...list(parsed.to), ...list(parsed.cc)];
  }
  const cc = all
    ? rest.filter((a, i, arr) => notMe(a) && !to.some((t) => same(t, a)) && arr.findIndex((x) => same(x, a)) === i)
    : [];
  return { to, cc };
}

export async function sendReply(acct, uid, { body, all, attachments = [], box = 'inbox' }) {
  const servers = serversFor(acct);
  const { parsed } = await withMailbox(acct, await boxPath(acct, box), (c) => fetchParsed(acct, uid, c));
  const { to, cc } = replyRecipients(acct, parsed, box, all);
  if (!to.length) throw Object.assign(new Error('No one to reply to.'), { status: 400 });

  const subject = parsed.subject || '';
  const refs = [].concat(parsed.references || [], parsed.messageId || []);
  const sender = list(parsed.from)[0] || { name: '', address: '' };
  const when = (parsed.date || new Date()).toUTCString();
  const intro = `On ${when}, ${fmtAddr(sender)} wrote:`;
  const quotedText = (parsed.text || '').split('\n').map((l) => '> ' + l).join('\n');
  const quotedHtml = parsed.html || `<pre style="white-space:pre-wrap">${esc(parsed.text || '')}</pre>`;

  const mail = {
    from: acct.name ? { name: acct.name, address: acct.email } : acct.email,
    to,
    cc: cc.length ? cc : undefined,
    subject: /^re:/i.test(subject) ? subject : `Re: ${subject}`,
    inReplyTo: parsed.messageId,
    references: refs.length ? refs : undefined,
    text: `${body}\n\n${intro}\n${quotedText}`,
    html: `<div style="white-space:pre-wrap">${esc(body)}</div><br><div>${esc(intro)}</div>`
      + `<blockquote style="margin:0 0 0 .8ex;border-left:1px solid #ccc;padding-left:1ex">${quotedHtml}</blockquote>`,
    attachments: attachments.length ? attachments : undefined,
  };

  const raw = await new MailComposer(mail).compile().build();
  const transport = await smtpTransport(acct);
  try {
    await transport.sendMail({
      envelope: { from: acct.email, to: [...to, ...cc].map((a) => a.address) },
      raw,
    });
  } catch (err) {
    throw friendly(err);
  }

  sentCache.delete(acct.id); // so the next refresh picks up the new Sent copy

  await withInbox(acct, async (c) => {
    if (box === 'inbox') await c.messageFlagsAdd(String(uid), ['\\Answered'], { uid: true });
    if (servers.appendSent) {
      const sent = await findSpecial(c, '\\Sent');
      if (sent) await c.append(sent, raw, ['\\Seen']);
    }
  });
}
