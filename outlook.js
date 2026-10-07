// Outlook / Hotmail mail through Microsoft Graph (IMAP sign-in is unreliable
// for personal Microsoft accounts). Same shapes as mail.js returns.
import { graph, utc } from './microsoft.js';
import { findDates } from './calendar.js';

const FEED_SIZE = 50;
const SENT_SCAN = 100;
const SENT_STALE_MS = 60 * 1000;

const LIST_FIELDS = 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,bodyPreview,conversationId,internetMessageId,parentFolderId,isDraft';

const person = (r) => (r?.emailAddress ? { name: r.emailAddress.name || '', address: r.emailAddress.address || '' } : null);
const people = (list) => (list || []).map(person).filter(Boolean);
const preview = (s) => String(s || '').replace(/\s+/g, ' ').trim().slice(0, 240);

const sentCache = new Map(); // account id -> { at, items, folderId }

export function dropOutlook(acctId) {
  sentCache.delete(acctId);
}

export async function testLogin(acct) {
  const f = await graph(acct, '/me/mailFolders/inbox', { query: { $select: 'totalItemCount' } });
  return f.totalItemCount;
}

async function sentState(acct) {
  const hit = sentCache.get(acct.id);
  if (hit && Date.now() - hit.at < SENT_STALE_MS) return hit;
  const folder = hit?.folderId || (await graph(acct, '/me/mailFolders/sentitems', { query: { $select: 'id' } })).id;
  const r = await graph(acct, '/me/mailFolders/sentitems/messages', {
    query: { $top: String(SENT_SCAN), $orderby: 'sentDateTime desc', $select: LIST_FIELDS },
  });
  const items = (r.value || []).map((m) => ({
    box: 'sent', uid: m.id, date: utc(m.sentDateTime || m.receivedDateTime), thread: `o${m.conversationId}`,
    from: person(m.from), to: people(m.toRecipients), subject: m.subject || '(no subject)', preview: preview(m.bodyPreview),
  }));
  const state = { at: Date.now(), items, folderId: folder };
  sentCache.set(acct.id, state);
  return state;
}

export async function fetchFeed(acct) {
  const r = await graph(acct, '/me/mailFolders/inbox/messages', {
    query: { $top: String(FEED_SIZE), $orderby: 'receivedDateTime desc', $select: LIST_FIELDS },
  });
  let sent = { items: [] };
  try { sent = await sentState(acct); } catch {}

  const mine = new Map();
  for (const s of sent.items) {
    const t = mine.get(s.thread) || { count: 0, lastAt: null };
    t.count++;
    if (!t.lastAt || s.date > t.lastAt) t.lastAt = s.date;
    mine.set(s.thread, t);
  }

  return (r.value || []).filter((m) => !m.isDraft).map((m) => {
    const thread = `o${m.conversationId}`;
    const date = utc(m.receivedDateTime);
    // Graph doesn't say which message a reply answers; count it as replied
    // when I sent something in the same conversation afterwards.
    const after = sent.items.filter((s) => s.thread === thread && s.date > date).sort((a, b) => (a.date < b.date ? -1 : 1))[0];
    return {
      id: `${acct.id}:${m.id}`,
      acct: acct.id,
      uid: m.id,
      box: 'inbox',
      date,
      from: person(m.from),
      to: people(m.toRecipients),
      subject: m.subject || '(no subject)',
      messageId: m.internetMessageId || null,
      attach: !!m.hasAttachments,
      preview: preview(m.bodyPreview),
      seen: !!m.isRead,
      answered: !!after,
      flagged: false,
      thread,
      replied: after ? { at: after.date, uid: after.uid } : null,
      threadSent: mine.get(thread) || null,
    };
  });
}

export async function getThread(acct, key) {
  const conversationId = key.replace(/^o/, '');
  const sent = await sentState(acct);
  const r = await graph(acct, '/me/messages', {
    query: { $filter: `conversationId eq '${conversationId.replace(/'/g, "''")}'`, $select: LIST_FIELDS, $top: '50' },
  });
  return (r.value || []).filter((m) => !m.isDraft).map((m) => {
    const isSent = m.parentFolderId === sent.folderId;
    return {
      box: isSent ? 'sent' : 'inbox', uid: m.id, acct: acct.id,
      date: utc(isSent ? m.sentDateTime : m.receivedDateTime),
      from: person(m.from), to: people(m.toRecipients), subject: m.subject || '(no subject)',
      preview: preview(m.bodyPreview), seen: isSent ? true : !!m.isRead, attach: !!m.hasAttachments,
    };
  }).sort((a, b) => (a.date < b.date ? -1 : 1));
}

const htmlToText = (html) => String(html || '')
  .replace(/<(style|script)[^>]*>[\s\S]*?<\/\1>/gi, '')
  .replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|tr|li|h\d)>/gi, '\n')
  .replace(/<[^>]+>/g, '')
  .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
  .replace(/\n{3,}/g, '\n\n').trim();

async function loadMessage(acct, id) {
  return graph(acct, `/me/messages/${encodeURIComponent(id)}`, {
    query: {
      $select: 'id,subject,from,toRecipients,ccRecipients,replyTo,receivedDateTime,sentDateTime,body,hasAttachments,internetMessageId,isRead,parentFolderId,meetingMessageType,conversationId',
      $expand: 'attachments($select=id,name,contentType,size,isInline)',
    },
  });
}

function inviteFrom(m) {
  const e = m.event;
  if (!e || m.meetingMessageType !== 'meetingRequest') return null;
  return {
    method: 'REQUEST',
    uid: e.iCalUId,
    graphEventId: e.id,
    title: e.subject || m.subject || '(no title)',
    location: e.location?.displayName || '',
    allDay: !!e.isAllDay,
    start: e.isAllDay ? e.start.dateTime.slice(0, 10) : utc(e.start.dateTime),
    end: e.isAllDay ? e.end.dateTime.slice(0, 10) : utc(e.end.dateTime),
    organizer: person(e.organizer),
    recurring: !!e.recurrence,
  };
}

export async function getMessage(acct, uid, box = 'inbox') {
  const m = await loadMessage(acct, uid);
  if (box === 'inbox' && !m.isRead) {
    graph(acct, `/me/messages/${encodeURIComponent(uid)}`, { method: 'PATCH', body: { isRead: true } }).catch(() => {});
  }

  let html = m.body?.contentType === 'html' ? m.body.content : null;
  const text = m.body?.contentType === 'html' ? htmlToText(m.body.content) : (m.body?.content || '');
  const atts = m.attachments || [];
  if (html && /cid:/i.test(html) && atts.some((a) => a.isInline)) {
    // Fetch inline images and embed them.
    let budget = 8 * 1024 * 1024;
    for (const a of atts.filter((x) => x.isInline && x.size < budget)) {
      try {
        const full = await graph(acct, `/me/messages/${encodeURIComponent(uid)}/attachments/${encodeURIComponent(a.id)}`);
        if (full.contentId && full.contentBytes) {
          budget -= a.size;
          html = html.split(`cid:${full.contentId.replace(/^<|>$/g, '')}`).join(`data:${full.contentType};base64,${full.contentBytes}`);
        }
      } catch {}
    }
  }

  let invite = null;
  if (box === 'inbox' && m.meetingMessageType === 'meetingRequest') {
    try {
      const ev = await graph(acct, `/me/messages/${encodeURIComponent(uid)}`, {
        query: { $select: 'meetingMessageType,subject', $expand: 'microsoft.graph.eventMessage/event' },
      });
      invite = inviteFrom(ev);
    } catch {}
  }

  // Use what the last feed refresh learned about my sent mail.
  let replied = null;
  const sent = sentCache.get(acct.id);
  if (box === 'inbox' && sent) {
    const date = utc(m.receivedDateTime);
    const after = sent.items.filter((s) => s.thread === `o${m.conversationId}` && s.date > date).sort((a, b) => (a.date < b.date ? -1 : 1))[0];
    if (after) replied = { at: after.date, uid: after.uid };
  }

  const date = utc(box === 'sent' ? m.sentDateTime : m.receivedDateTime);
  return {
    acct: acct.id,
    uid,
    box,
    messageId: m.internetMessageId || null,
    replied,
    invite,
    dates: findDates(m.subject, text, date),
    subject: m.subject || '(no subject)',
    date,
    from: person(m.from),
    to: people(m.toRecipients),
    cc: people(m.ccRecipients),
    html,
    text,
    attachments: atts.filter((a) => !a.isInline)
      .map((a, i) => ({ idx: i, filename: a.name || 'attachment', size: a.size, contentType: a.contentType })),
  };
}

export async function getAttachment(acct, uid, idx) {
  const m = await graph(acct, `/me/messages/${encodeURIComponent(uid)}/attachments`, { query: { $select: 'id,name,contentType,size,isInline' } });
  const a = (m.value || []).filter((x) => !x.isInline)[Number(idx)];
  if (!a) throw Object.assign(new Error('Attachment not found'), { status: 404 });
  const full = await graph(acct, `/me/messages/${encodeURIComponent(uid)}/attachments/${encodeURIComponent(a.id)}`);
  return { filename: full.name, contentType: full.contentType, content: Buffer.from(full.contentBytes || '', 'base64') };
}

export async function setSeen(acct, uid, seen) {
  await graph(acct, `/me/messages/${encodeURIComponent(uid)}`, { method: 'PATCH', body: { isRead: !!seen } });
}

export async function deleteMessage(acct, uid) {
  await graph(acct, `/me/messages/${encodeURIComponent(uid)}/move`, { method: 'POST', body: { destinationId: 'deleteditems' } });
}

const MAX_SIMPLE_ATTACH = 3 * 1024 * 1024; // larger files need an upload session

export async function sendReply(acct, uid, { body, all, attachments = [], box = 'inbox' }) {
  const id = encodeURIComponent(uid);
  for (const a of attachments) {
    if (a.content.length > MAX_SIMPLE_ATTACH) {
      throw Object.assign(new Error(`"${a.filename}" is over 3 MB; Outlook replies can't carry files that big yet.`), { status: 413 });
    }
  }
  // Outlook builds the reply (subject, quoting, threading); we add our text.
  const action = all || box === 'sent' ? 'createReplyAll' : 'createReply';
  const draft = await graph(acct, `/me/messages/${id}/${action}`, { method: 'POST', body: { comment: body } });
  const did = encodeURIComponent(draft.id);
  try {
    if (box === 'sent') {
      // Following up my own message: send to the same people, not to myself.
      const orig = await graph(acct, `/me/messages/${id}`, { query: { $select: 'toRecipients,ccRecipients' } });
      const me = acct.email.toLowerCase();
      const notMe = (r) => r.emailAddress?.address?.toLowerCase() !== me;
      await graph(acct, `/me/messages/${did}`, {
        method: 'PATCH',
        body: { toRecipients: (orig.toRecipients || []).filter(notMe), ccRecipients: all ? (orig.ccRecipients || []).filter(notMe) : [] },
      });
    }
    for (const a of attachments) {
      await graph(acct, `/me/messages/${did}/attachments`, {
        method: 'POST',
        body: { '@odata.type': '#microsoft.graph.fileAttachment', name: a.filename, contentType: a.contentType, contentBytes: a.content.toString('base64') },
      });
    }
    await graph(acct, `/me/messages/${did}/send`, { method: 'POST' });
  } catch (err) {
    graph(acct, `/me/messages/${did}`, { method: 'DELETE' }).catch(() => {});
    throw err;
  }
  sentCache.delete(acct.id);
}

// Answer an Outlook meeting request; Outlook tells the organizer and updates the calendar.
export async function rsvp(acct, uid, partstat) {
  const ev = await graph(acct, `/me/messages/${encodeURIComponent(uid)}`, {
    query: { $select: 'meetingMessageType,subject', $expand: 'microsoft.graph.eventMessage/event' },
  });
  const invite = inviteFrom(ev);
  if (!invite?.graphEventId) throw Object.assign(new Error('Open this invitation in Outlook to answer it.'), { status: 400 });
  const verb = { ACCEPTED: 'accept', TENTATIVE: 'tentativelyAccept', DECLINED: 'decline' }[partstat];
  await graph(acct, `/me/events/${encodeURIComponent(invite.graphEventId)}/${verb}`, { method: 'POST', body: { sendResponse: true } });
  return 'Outlook Calendar';
}

export async function getSentReply(acct, uid) {
  const msg = await getMessage(acct, uid, 'sent');
  // Outlook puts the quoted original after a "From:" block or a long rule.
  const cut = msg.text.search(/\n\s*(_{10,}|-{10,}|From: .+\n\s*Sent: |De : .+\n\s*Envoyé : )/);
  return { date: msg.date, to: msg.to, text: (cut > 0 ? msg.text.slice(0, cut) : msg.text).trim() };
}
