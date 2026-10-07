// Encrypted storage for account credentials + app login config.
// Nothing about emails themselves is ever written to disk.
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

const DATA = process.env.FEEDMAIL_DATA || path.join(process.cwd(), 'data');
fs.mkdirSync(DATA, { recursive: true, mode: 0o700 });

const keyFile = path.join(DATA, 'secret.key');
const accountsFile = path.join(DATA, 'accounts.enc');
const configFile = path.join(DATA, 'config.json');

function writePrivate(file, content) {
  fs.writeFileSync(file + '.tmp', content, { mode: 0o600 });
  fs.renameSync(file + '.tmp', file);
}

if (!fs.existsSync(keyFile)) writePrivate(keyFile, crypto.randomBytes(32));
const KEY = fs.readFileSync(keyFile);

function encrypt(obj) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', KEY, iv);
  const data = Buffer.concat([c.update(JSON.stringify(obj), 'utf8'), c.final()]);
  return Buffer.concat([iv, c.getAuthTag(), data]);
}

function decrypt(buf) {
  const d = crypto.createDecipheriv('aes-256-gcm', KEY, buf.subarray(0, 12));
  d.setAuthTag(buf.subarray(12, 28));
  return JSON.parse(Buffer.concat([d.update(buf.subarray(28)), d.final()]).toString('utf8'));
}

export function loadAccounts() {
  if (!fs.existsSync(accountsFile)) return [];
  return decrypt(fs.readFileSync(accountsFile));
}

export function saveAccounts(accounts) {
  writePrivate(accountsFile, encrypt(accounts));
}

export function loadConfig() {
  if (!fs.existsSync(configFile)) {
    const cfg = { sessionSecret: crypto.randomBytes(32).toString('hex') };
    saveConfig(cfg);
    return cfg;
  }
  return JSON.parse(fs.readFileSync(configFile, 'utf8'));
}

export function saveConfig(cfg) {
  writePrivate(configFile, JSON.stringify(cfg, null, 2));
}

// App-level secrets such as OAuth client credentials (encrypted).
const secretsFile = path.join(DATA, 'secrets.enc');

export function loadSecrets() {
  return fs.existsSync(secretsFile) ? decrypt(fs.readFileSync(secretsFile)) : {};
}

export function saveSecrets(secrets) {
  writePrivate(secretsFile, encrypt(secrets));
}
