/*
 * store.js — server-side state storage for the Hydro-Wates PM app.
 *
 * Two backends, ONE synchronous interface (read / write / has):
 *   - "file"     (default): the original data/*.json files on local disk.
 *   - "supabase" : rows in the pm_app_state table (key -> jsonb). Used on Vercel,
 *                  where there is no persistent disk.
 *
 * It switches to Supabase automatically when SUPABASE_URL + SUPABASE_SERVICE_KEY
 * are set (env vars or a local .env). On first run in Supabase mode it SEEDS the
 * table from any existing local files, so migrating is automatic and lossless.
 *
 * Reads are synchronous: in Supabase mode the whole (small) state is loaded into
 * memory at init() and kept in sync; writes update memory immediately and are
 * flushed to Supabase in the background (serialized per key). Large values (the
 * Zoho/lead cache) are gzip-compressed so they stay well under request limits.
 *
 * Zero external dependencies (Node 18+ built-ins only).
 */
'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

let MODE = 'file';
let SB = null;                 // { url, key }
const mem = {};                // stateKey -> value (Supabase mode in-memory cache)
const pathToKey = {};          // absolute file path -> stateKey
const keyToPath = {};          // stateKey -> absolute file path
const writeChain = {};         // stateKey -> Promise (serialize writes per key)
let ROOT = process.cwd();

// ---- tiny .env loader (KEY=VALUE lines; does not override already-set vars) ----
function loadEnv(dir) {
  try {
    const file = path.join(dir, '.env');
    if (!fs.existsSync(file)) return;
    for (const raw of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
      const line = raw.trim();
      if (!line || line.startsWith('#')) continue;
      const eq = line.indexOf('=');
      if (eq < 0) continue;
      const k = line.slice(0, eq).trim();
      let v = line.slice(eq + 1).trim();
      if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
      if (k && process.env[k] === undefined) process.env[k] = v;
    }
  } catch (e) { /* ignore — env stays as-is */ }
}

// ---- gzip wrapper so big values fit in one row / request -----------------------
const GZIP_OVER = 512 * 1024; // compress anything whose JSON is larger than 512 KB
function wrap(obj) {
  const json = JSON.stringify(obj);
  if (json.length > GZIP_OVER) {
    return { _gz: zlib.gzipSync(Buffer.from(json, 'utf8')).toString('base64') };
  }
  return obj === undefined ? null : obj;
}
function unwrap(value) {
  if (value && typeof value === 'object' && typeof value._gz === 'string') {
    return JSON.parse(zlib.gunzipSync(Buffer.from(value._gz, 'base64')).toString('utf8'));
  }
  return value;
}

// ---- Supabase REST (service-role key; bypasses RLS on pm_app_state) ------------
async function sbGet(pathQ) {
  const r = await fetch(SB.url + '/rest/v1/' + pathQ, {
    headers: { apikey: SB.key, Authorization: 'Bearer ' + SB.key, Accept: 'application/json' }
  });
  const text = await r.text();
  if (!r.ok) throw new Error('Supabase GET ' + r.status + ': ' + text.slice(0, 200));
  return text ? JSON.parse(text) : [];
}
async function sbUpsert(key, wrappedValue) {
  const body = JSON.stringify([{ key, value: wrappedValue, updated_at: new Date().toISOString() }]);
  const r = await fetch(SB.url + '/rest/v1/pm_app_state', {
    method: 'POST',
    headers: {
      apikey: SB.key, Authorization: 'Bearer ' + SB.key,
      'Content-Type': 'application/json', Prefer: 'resolution=merge-duplicates,return=minimal'
    },
    body
  });
  if (!r.ok) { const t = await r.text(); throw new Error('Supabase upsert ' + r.status + ': ' + t.slice(0, 200)); }
}

function readDisk(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch (e) { return fallback; }
}
function writeDisk(file, obj) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(obj, null, 2), 'utf8');
}

// ---- public API ----------------------------------------------------------------

// Call once at startup. opts: { dataDir, files } where files maps stateKey -> path.
async function init(opts) {
  const files = opts.files || {};
  ROOT = path.dirname(opts.dataDir || process.cwd());
  for (const [k, p] of Object.entries(files)) { pathToKey[p] = k; keyToPath[k] = p; }

  loadEnv(ROOT);
  const url = process.env.SUPABASE_URL;
  const svc = process.env.SUPABASE_SERVICE_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !svc) { MODE = 'file'; return; }   // no service key -> stay on local files

  SB = { url: url.replace(/\/+$/, ''), key: svc };
  try {
    // Load everything already in the table.
    const rows = await sbGet('pm_app_state?select=key,value');
    for (const r of rows) mem[r.key] = unwrap(r.value);

    // One-time migration: seed any state key that exists on disk but not yet in the table.
    let seeded = 0;
    for (const k of Object.keys(keyToPath)) {
      if (k in mem) continue;
      const disk = readDisk(keyToPath[k], undefined);
      if (disk !== undefined) { mem[k] = disk; await sbUpsert(k, wrap(disk)); seeded++; }
    }
    MODE = 'supabase';
    console.log('[store] Supabase mode — loaded ' + rows.length + ' key(s)' + (seeded ? ', migrated ' + seeded + ' from disk' : ''));
  } catch (e) {
    MODE = 'file'; SB = null;
    for (const k of Object.keys(mem)) delete mem[k];
    console.error('[store] Supabase unreachable (' + e.message + ') — using local files instead.');
  }
}

function read(file, fallback) {
  if (MODE === 'file') return readDisk(file, fallback);
  const key = pathToKey[file];
  if (key === undefined) return readDisk(file, fallback);   // unknown file -> disk
  return (key in mem) ? mem[key] : fallback;
}

function write(file, obj) {
  if (MODE === 'file') return writeDisk(file, obj);
  const key = pathToKey[file];
  if (key === undefined) return writeDisk(file, obj);       // unknown file -> disk
  mem[key] = obj;
  const wrapped = wrap(obj);                                 // snapshot now (caller may mutate obj later)
  const prev = writeChain[key] || Promise.resolve();
  writeChain[key] = prev
    .then(() => sbUpsert(key, wrapped))
    .catch((e) => console.error('[store] write failed for "' + key + '": ' + e.message));
}

function has(file) {
  if (MODE === 'file') { try { return fs.existsSync(file); } catch (e) { return false; } }
  const key = pathToKey[file];
  return key !== undefined && (key in mem);
}

// Await all pending background writes (graceful shutdown / tests).
async function flush() {
  await Promise.allSettled(Object.values(writeChain));
}

function mode() { return MODE; }

module.exports = { init, read, write, has, flush, mode };
