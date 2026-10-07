// Server settings for each supported mail provider.
// appendSent: whether we must save a copy of replies into the Sent folder
// ourselves (some providers do it automatically when sending via SMTP).

export const PROVIDERS = {
  yahoo: {
    label: 'Yahoo',
    imap: { host: 'imap.mail.yahoo.com', port: 993, secure: true },
    smtp: { host: 'smtp.mail.yahoo.com', port: 465, secure: true },
    appendSent: false,
  },
  aol: {
    label: 'AOL',
    imap: { host: 'imap.aol.com', port: 993, secure: true },
    smtp: { host: 'smtp.aol.com', port: 465, secure: true },
    appendSent: false,
  },
  gmx: {
    label: 'GMX',
    imap: { host: 'imap.gmx.com', port: 993, secure: true },
    smtp: { host: 'mail.gmx.com', port: 587, secure: false },
    appendSent: true,
  },
};

const GMX_NET_DOMAINS = ['gmx.net', 'gmx.de', 'gmx.at', 'gmx.ch'];

const CALDAV = {
  yahoo: 'https://caldav.calendar.yahoo.com',
  aol: 'https://caldav.aol.com', // unverified until an AOL account is added
  gmx: 'https://caldav.gmx.com',
};

export function caldavUrlFor(account) {
  const domain = account.email.split('@')[1]?.toLowerCase();
  if (account.provider === 'gmx' && GMX_NET_DOMAINS.includes(domain)) return 'https://caldav.gmx.net';
  return CALDAV[account.provider] || null;
}

export function serversFor(account) {
  const p = PROVIDERS[account.provider];
  if (!p) throw new Error('Unknown provider ' + account.provider);
  const s = { imap: { ...p.imap }, smtp: { ...p.smtp }, appendSent: p.appendSent };
  if (account.provider === 'gmx') {
    const domain = account.email.split('@')[1]?.toLowerCase();
    if (GMX_NET_DOMAINS.includes(domain)) {
      s.imap.host = 'imap.gmx.net';
      s.smtp.host = 'mail.gmx.net';
    }
  }
  return s;
}
