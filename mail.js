// Talks to the mail servers. Messages are fetched on demand and kept only in
// memory; nothing is saved to disk.
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import MailComposer from 'nodemailer/lib/mail-composer/index.js';
import { serversFor } from './providers.js';
import { parseInvite, findDates } from './calendar.js';

const FEED_SIZE = 50; // newest messages per account
const PREVIEW_BYTES = 20000; // enough of each message to build a preview

const clients = new Map(); // account id -> Promise<ImapFlow>

function newClient(acct) {
  const { imap } = serversFor(acct);
  const c = new ImapFlow({
    ...imap,
    auth: { user: acct.email, pass: acct.password },
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
    const c = newClient(acct);
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
  const p = clients.get(acctId);
  clients.delete(acctId);
  p?.then((c) => c.logout().catch(() => c.close())).catch(() => {});
}

// Some servers (Yahoo) never tell a long-open session about new mail, so
// re-select INBOX when the view may be stale.
const RESELECT_MS = 15 * 1000;

// Run fn with INBOX selected; retry once on a dead connection.
async function withInbox(acct, fn, { fresh = false } = {}, retried = false) {
  const c = await getClient(acct);
  let lock;
  try {
    lock = await c.getMailboxLock('INBOX');
    if (fresh || !c.fmSelectedAt || Date.now() - c.fmSelectedAt > RESELECT_MS) {
      await c.mailboxOpen('INBOX');
    }
    c.fmSelectedAt = Date.now();
    return await fn(c);
  } catch (err) {
    if (!retried && !c.usable) {
      clients.delete(acct.id);
      return withInbox(acct, fn, { fresh }, true);
    }
    throw friendly(err);
  } finally {
    lock?.release();
  }
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
  const c = newClient(acct);
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

export async function fetchFeed(acct) {
  const raw = await withInbox(acct, async (c) => {
    const exists = c.mailbox.exists;
    if (!exists) return [];
    const start = Math.max(1, exists - FEED_SIZE + 1);
    const out = [];
    for await (const m of c.fetch(`${start}:*`, {
      uid: true, flags: true, envelope: true, internalDate: true, bodyStructure: true,
      source: { maxLength: PREVIEW_BYTES },
    })) out.push(m);
    return out;
  }, { fresh: true });

  const items = [];
  for (const m of raw) {
    let preview = '';
    try {
      preview = makePreview(await simpleParser(m.source, { skipImageLinks: true, skipTextToHtml: true }));
    } catch {}
    const env = m.envelope || {};
    items.push({
      id: `${acct.id}:${m.uid}`,
      acct: acct.id,
      uid: m.uid,
      date: (env.date || m.internalDate || new Date()).toISOString(),
      from: addr(env.from?.[0]),
      to: (env.to || []).map(addr),
      subject: env.subject || '(no subject)',
      seen: m.flags?.has('\\Seen') || false,
      answered: m.flags?.has('\\Answered') || false,
      flagged: m.flags?.has('\\Flagged') || false,
      attach: hasAttachment(m.bodyStructure),
      preview,
    });
  }
  return items;
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

export async function getMessage(acct, uid) {
  const { parsed } = await withInbox(acct, async (c) => {
    const r = await fetchParsed(acct, uid, c);
    if (!r.flags?.has('\\Seen')) await c.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
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

  return {
    acct: acct.id,
    uid: Number(uid),
    invite,
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
  const transport = nodemailer.createTransport({ ...servers.smtp, auth: { user: acct.email, pass: acct.password } });
  try {
    await transport.sendMail({ envelope: { from: acct.email, to: [to] }, raw });
  } catch (err) {
    throw friendly(err);
  }
}

export async function getAttachment(acct, uid, idx) {
  const { parsed } = await withInbox(acct, (c) => fetchParsed(acct, uid, c));
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

export async function sendReply(acct, uid, { body, all }) {
  const servers = serversFor(acct);
  const { parsed } = await withInbox(acct, (c) => fetchParsed(acct, uid, c));

  const me = acct.email.toLowerCase();
  const same = (a, b) => a.address?.toLowerCase() === b.address?.toLowerCase();
  const to = list(parsed.replyTo).length ? list(parsed.replyTo) : list(parsed.from);
  const cc = all
    ? [...list(parsed.to), ...list(parsed.cc)].filter(
        (a, i, arr) => a.address && a.address.toLowerCase() !== me && !to.some((t) => same(t, a)) && arr.findIndex((x) => same(x, a)) === i)
    : [];

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
  };

  const raw = await new MailComposer(mail).compile().build();
  const transport = nodemailer.createTransport({ ...servers.smtp, auth: { user: acct.email, pass: acct.password } });
  try {
    await transport.sendMail({
      envelope: { from: acct.email, to: [...to, ...cc].map((a) => a.address) },
      raw,
    });
  } catch (err) {
    throw friendly(err);
  }

  await withInbox(acct, async (c) => {
    await c.messageFlagsAdd(String(uid), ['\\Answered'], { uid: true });
    if (servers.appendSent) {
      const sent = await findSpecial(c, '\\Sent');
      if (sent) await c.append(sent, raw, ['\\Seen']);
    }
  });
}
