// Rebuild the encrypted PAYLOAD in index.html from the Excel files.
//   DASH_PASSWORD=... LIVE_URL=<apps script web app url> node build/build.mjs
//   DASH_PASSWORD=... node build/build.mjs --from DIR     (reads DIR/<tab id>.xlsx)
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const {parseWorkbook, stable} = require(path.join(ROOT, 'parse.js'));

const TABS = [
  {id: 'receivables', title: 'المستحقات والالتزامات', kind: 'receivables', source: 'حسابات العملاء.xlsx'},
  {id: 'work', title: 'أسعار الشغل', kind: 'generic', source: 'اسعار الشغل.xlsx'},
  {id: 'suppliers', title: 'أسعار البنود للموردين', kind: 'generic', source: 'اسعار البنود للموردين.xlsx',
   rename: {Sheet1: 'أسعار البنود'}},
];
const HTML = path.join(ROOT, 'index.html');
const ITER = 250000;
const PAYLOAD_RE = /const PAYLOAD = (\{.*?\});\n/;

const fail = msg => { console.error(msg); process.exit(1); };
const pw = process.env.DASH_PASSWORD || fail('DASH_PASSWORD is not set');
const from = process.argv.includes('--from') ? process.argv[process.argv.indexOf('--from') + 1] : null;
const live = process.env.LIVE_URL || null;
if (!from && !live) fail('LIVE_URL is not set');

const derive = (salt, iter) => crypto.pbkdf2Sync(pw, salt, iter, 32, 'sha256');
function decrypt(p) {
  const ct = Buffer.from(p.ct, 'base64');
  const d = crypto.createDecipheriv('aes-256-gcm', derive(Buffer.from(p.salt, 'base64'), p.iter), Buffer.from(p.iv, 'base64'));
  d.setAuthTag(ct.subarray(-16));
  return JSON.parse(Buffer.concat([d.update(ct.subarray(0, -16)), d.final()]).toString('utf8'));
}
function encrypt(doc) {
  const salt = crypto.randomBytes(16), iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', derive(salt, ITER), iv);
  const ct = Buffer.concat([c.update(JSON.stringify(doc), 'utf8'), c.final(), c.getAuthTag()]);
  return {enc: 1, salt: salt.toString('base64'), iv: iv.toString('base64'), iter: ITER, ct: ct.toString('base64')};
}

const html = fs.readFileSync(HTML, 'utf8');
const m = PAYLOAD_RE.exec(html);
const old = JSON.parse(m[1]);
let prevDoc = {tabs: []};
if (old.enc) { try { prevDoc = decrypt(old); } catch (e) { fail("DASH_PASSWORD doesn't open the current dashboard"); } }
const prev = Object.fromEntries(prevDoc.tabs.map(t => [t.id, t]));

let files = {};
if (live) {
  const res = await fetch(live, {redirect: 'follow'});
  const body = await res.text();
  try { files = JSON.parse(body).files; } catch (e) { fail(`LIVE_URL didn't return JSON (${res.status}): ${body.slice(0, 200)}`); }
}

const now = new Intl.DateTimeFormat('sv-SE', {timeZone: 'Africa/Cairo', dateStyle: 'short', timeStyle: 'short'}).format(new Date());
const changed = [];
const tabs = TABS.map(t => {
  const f = files[t.id];
  const bytes = from ? fs.readFileSync(path.join(from, `${t.id}.xlsx`)) : f && f.b64 ? Buffer.from(f.b64, 'base64') : null;
  if (!bytes) fail(`couldn't get "${t.source}" — check the file name in the Apps Script`);
  const data = parseWorkbook(XLSX.read(bytes, {type: 'buffer', sheetStubs: true}), t);
  const p = prev[t.id], same = p && stable(p.data) === stable(data);
  if (!same) changed.push(t.title);
  return {...t, modified: same ? p.modified : (f && f.modified) || now,
          edit_url: (f && f.edit_url) || (p && p.edit_url) || '', data};
});

if (!changed.length && (prevDoc.live || null) === (live || prevDoc.live || null) && stable(TABS.map(t => t.id)) === stable(Object.keys(prev))) {
  console.log('no changes');
  process.exit(0);
}
const doc = {generated: now, live: live || prevDoc.live || null, tabs};
fs.writeFileSync(HTML, html.slice(0, m.index) + `const PAYLOAD = ${JSON.stringify(encrypt(doc))};\n` + html.slice(m.index + m[0].length));
console.log(`updated ${now}: ${changed.join(', ') || 'live link'}`);
if (process.env.GITHUB_OUTPUT) fs.appendFileSync(process.env.GITHUB_OUTPUT, `stamp=${now}\n`);
