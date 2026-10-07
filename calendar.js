// Calendars over CalDAV: read events from every account, add accepted invites.
// Events are only held in memory for a couple of minutes.
import { createDAVClient } from 'tsdav';
import ICAL from 'ical.js';
import * as chrono from 'chrono-node';
import { caldavUrlFor, calendarKind } from './providers.js';
import { googleEvents, googleCalendars, googleRsvp, googleImport } from './google.js';

const CAL_LIST_MS = 30 * 60 * 1000;
const EVENTS_MS = 2 * 60 * 1000;
const MAX_OCCURRENCES = 400;

const clients = new Map(); // account id -> { at, promise: { client, calendars } }
const eventCache = new Map(); // `${acct}|${from}|${to}` -> { at, events }

export function forgetCalendar(acctId) {
  clients.delete(acctId);
  for (const k of eventCache.keys()) if (k.startsWith(acctId + '|')) eventCache.delete(k);
}

const calPassword = (acct) => acct.calPassword || acct.password;

function friendly(err) {
  if (/401|Unauthorized|Invalid credentials/i.test(err.message)) {
    const e = new Error('Calendar login refused. GMX needs its own calendar password: add it in Accounts.');
    e.status = 401;
    return e;
  }
  return err;
}

async function connect(acct) {
  const serverUrl = caldavUrlFor(acct);
  if (!serverUrl) throw new Error('No calendar support for this provider yet.');
  const client = await createDAVClient({
    serverUrl,
    credentials: { username: acct.email, password: calPassword(acct) },
    authMethod: 'Basic',
    defaultAccountType: 'caldav',
  });
  const calendars = (await client.fetchCalendars())
    .filter((c) => !c.components || c.components.includes('VEVENT'));
  return { client, calendars };
}

async function getCal(acct) {
  const e = clients.get(acct.id);
  if (e && Date.now() - e.at < CAL_LIST_MS) return e.promise;
  const promise = connect(acct).catch((err) => {
    clients.delete(acct.id);
    throw friendly(err);
  });
  clients.set(acct.id, { at: Date.now(), promise });
  return promise;
}

export async function testCalendar(acct) {
  const { calendars } = await connect(acct).catch((err) => { throw friendly(err); });
  return calendars.length;
}

// ---------- iCalendar parsing ----------

function registerTimezones(root) {
  for (const tz of root.getAllSubcomponents('vtimezone')) {
    try { ICAL.TimezoneService.register(new ICAL.Timezone(tz)); } catch {}
  }
}

function toOut(time) {
  // All-day values stay as plain dates; timed values become absolute instants.
  return time.isDate ? time.toString().slice(0, 10) : time.toJSDate().toISOString();
}

function personOf(prop) {
  if (!prop) return null;
  const v = String(prop.getFirstValue() || '').replace(/^mailto:/i, '');
  return { address: v, name: prop.getParameter('cn') || '' };
}

function eventsFromIcs(ics, from, to) {
  const out = [];
  let root;
  try { root = new ICAL.Component(ICAL.parse(ics)); } catch { return out; }
  registerTimezones(root);
  const vevents = root.getAllSubcomponents('vevent');
  const masters = vevents.filter((v) => !v.hasProperty('recurrence-id'));
  const exceptions = vevents.filter((v) => v.hasProperty('recurrence-id'));

  for (const m of masters) {
    const ev = new ICAL.Event(m);
    for (const x of exceptions) if (x.getFirstPropertyValue('uid') === ev.uid) ev.relateException(x);
    const push = (start, end, item) => {
      const s = start.toJSDate(), en = (end || start).toJSDate();
      if (en > from && s < to) {
        out.push({
          uid: ev.uid,
          title: item.summary || '(no title)',
          location: item.location || '',
          allDay: start.isDate,
          start: toOut(start),
          end: toOut(end || start),
        });
      }
    };
    if (ev.isRecurring()) {
      const it = ev.iterator();
      let next, n = 0;
      while ((next = it.next()) && n++ < MAX_OCCURRENCES) {
        if (next.toJSDate() >= to) break;
        const d = ev.getOccurrenceDetails(next);
        push(d.startDate, d.endDate, d.item);
      }
    } else {
      push(ev.startDate, ev.endDate, ev);
    }
  }
  // Exceptions whose master is missing (rare) are shown on their own.
  for (const x of exceptions) {
    if (masters.some((m) => m.getFirstPropertyValue('uid') === x.getFirstPropertyValue('uid'))) continue;
    const ev = new ICAL.Event(x);
    const s = ev.startDate.toJSDate(), en = (ev.endDate || ev.startDate).toJSDate();
    if (en > from && s < to) out.push({ uid: ev.uid, title: ev.summary || '(no title)', location: ev.location || '', allDay: ev.startDate.isDate, start: toOut(ev.startDate), end: toOut(ev.endDate || ev.startDate) });
  }
  return out;
}

export async function fetchEvents(acct, from, to) {
  const key = `${acct.id}|${from.toISOString()}|${to.toISOString()}`;
  const hit = eventCache.get(key);
  if (hit && Date.now() - hit.at < EVENTS_MS) return hit.events;

  if (calendarKind(acct) === 'google') {
    const events = (await googleEvents(acct, from, to)).map((ev) => ({ ...ev, acct: acct.id }));
    eventCache.set(key, { at: Date.now(), events });
    return events;
  }

  const { client, calendars } = await getCal(acct);
  const events = [];
  // Ask the server for a slightly wider window so recurring masters are included.
  const timeRange = { start: new Date(from - 864e5).toISOString(), end: new Date(+to + 864e5).toISOString() };
  for (const cal of calendars) {
    const objects = await client.fetchCalendarObjects({ calendar: cal, timeRange });
    for (const o of objects) {
      for (const ev of eventsFromIcs(o.data || '', from, to)) {
        events.push({ ...ev, acct: acct.id, calendar: cal.displayName || 'Calendar' });
      }
    }
  }
  eventCache.set(key, { at: Date.now(), events });
  return events;
}

// ---------- invites inside emails ----------

export function parseInvite(ics) {
  let root;
  try { root = new ICAL.Component(ICAL.parse(ics)); } catch { return null; }
  registerTimezones(root);
  const v = root.getFirstSubcomponent('vevent');
  if (!v) return null;
  const ev = new ICAL.Event(v);
  return {
    method: (root.getFirstPropertyValue('method') || 'PUBLISH').toUpperCase(),
    uid: ev.uid,
    title: ev.summary || '(no title)',
    location: ev.location || '',
    allDay: ev.startDate?.isDate || false,
    start: ev.startDate ? toOut(ev.startDate) : null,
    end: ev.endDate ? toOut(ev.endDate) : ev.startDate ? toOut(ev.startDate) : null,
    organizer: personOf(v.getFirstProperty('organizer')),
    recurring: ev.isRecurring(),
  };
}

const pad = (n) => String(n).padStart(2, '0');
const icsStamp = (d = new Date()) =>
  `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;

function setMyStatus(vevent, email, partstat) {
  const mine = vevent.getAllProperties('attendee')
    .find((p) => String(p.getFirstValue()).replace(/^mailto:/i, '').toLowerCase() === email.toLowerCase());
  if (mine) {
    mine.setParameter('partstat', partstat);
    mine.removeParameter('rsvp');
    return mine;
  }
  const p = new ICAL.Property('attendee', vevent);
  p.setValue(`mailto:${email}`);
  p.setParameter('partstat', partstat);
  vevent.addProperty(p);
  return p;
}

// Build the iTIP REPLY that tells the organizer our answer.
export function buildReply(ics, email, partstat) {
  const src = new ICAL.Component(ICAL.parse(ics));
  const v = src.getFirstSubcomponent('vevent');
  const cal = new ICAL.Component(['vcalendar', [], []]);
  cal.addPropertyWithValue('prodid', '-//Feed Mail//EN');
  cal.addPropertyWithValue('version', '2.0');
  cal.addPropertyWithValue('method', 'REPLY');
  for (const tz of src.getAllSubcomponents('vtimezone')) cal.addSubcomponent(new ICAL.Component(tz.toJSON()));
  const r = new ICAL.Component('vevent');
  for (const name of ['uid', 'sequence', 'dtstart', 'dtend', 'duration', 'summary', 'organizer', 'recurrence-id']) {
    const p = v.getFirstProperty(name);
    if (p) r.addProperty(new ICAL.Property(p.toJSON()));
  }
  r.updatePropertyWithValue('dtstamp', ICAL.Time.fromString(new Date().toISOString().slice(0, 19) + 'Z'));
  const mine = setMyStatus(v, email, partstat);
  r.addProperty(new ICAL.Property(mine.toJSON()));
  cal.addSubcomponent(r);
  return cal.toString();
}

// Google accounts: answer in Google Calendar directly when the invite is
// already there. Returns true when Google will notify the organizer itself.
export async function rsvpViaProvider(acct, ics, partstat) {
  if (calendarKind(acct) !== 'google') return false;
  const invite = parseInvite(ics);
  const done = invite && await googleRsvp(acct, invite, partstat);
  if (done) forgetCalendar(acct.id);
  return done;
}

export async function testGoogleCalendar(acct) {
  return (await googleCalendars(acct)).length;
}

// Save the invited event into this account's main calendar.
export async function addToCalendar(acct, ics, partstat) {
  if (calendarKind(acct) === 'google') {
    const invite = parseInvite(ics);
    const vevent = new ICAL.Component(ICAL.parse(ics)).getFirstSubcomponent('vevent');
    const rrules = vevent.getAllProperties('rrule').map((p) => p.toICALString());
    const name = await googleImport(acct, invite, rrules, partstat);
    forgetCalendar(acct.id);
    return name;
  }
  const { client, calendars } = await getCal(acct);
  if (!calendars.length) throw new Error('No calendar found in this account.');
  const root = new ICAL.Component(ICAL.parse(ics));
  root.removeAllProperties('method');
  for (const v of root.getAllSubcomponents('vevent')) setMyStatus(v, acct.email, partstat);
  const uid = root.getFirstSubcomponent('vevent').getFirstPropertyValue('uid') || icsStamp();
  const filename = uid.replace(/[^\w.-]/g, '_').slice(0, 120) + '.ics';
  // Avoid special calendars like GMX's "Birthdays" when choosing where to save.
  const cal = calendars.find((c) => !/birthday|anniversaire|geburtstag|holiday|férié/i.test(c.displayName || '')) || calendars[0];
  const iCalString = root.toString();
  let res = await client.createCalendarObject({ calendar: cal, filename, iCalString });
  if (!res.ok) {
    // Already there (e.g. updated invite): overwrite it.
    res = await client.updateCalendarObject({ calendarObject: { url: new URL(filename, cal.url).href, data: iCalString } });
  }
  if (!res.ok) throw new Error(`Calendar refused the event (${res.status}).`);
  forgetCalendar(acct.id);
  return cal.displayName || 'Calendar';
}

// ---------- dates mentioned in an email ----------

function stripQuoted(text) {
  const lines = [];
  for (const l of text.split('\n')) {
    if (/^\s*>/.test(l)) continue;
    if (/^(On .+wrote:|Le .+a écrit\s*:|-----Original Message-----|De ?: .+)$/i.test(l.trim())) break;
    lines.push(l);
  }
  return lines.join('\n').slice(0, 4000);
}

const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

export function findDates(subject, text, sentAt) {
  const ref = new Date(sentAt);
  const body = `${subject || ''}\n${stripQuoted(text || '')}`;
  const results = [...chrono.parse(body, ref, { forwardDate: true }), ...chrono.fr.parse(body, ref, { forwardDate: true })];
  const seen = new Map();
  for (const r of results) {
    const d = r.start.date();
    const days = (d - ref) / 864e5;
    if (days < -1 || days > 365) continue;
    // A bare time ("at 3pm") with no day is too vague to be useful.
    if (!r.start.isCertain('day') && !r.start.isCertain('weekday')) continue;
    const key = ymd(d);
    if (!seen.has(key)) seen.set(key, { date: key, text: r.text.slice(0, 40) });
  }
  return [...seen.values()].sort((a, b) => (a.date < b.date ? -1 : 1)).slice(0, 4);
}
