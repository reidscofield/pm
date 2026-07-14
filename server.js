/*
 * Hydro-Wates Project Manager — local web app
 * Zero-dependency Node.js server (Node 18+ required; uses built-in fetch).
 *
 * Run:  node server.js   (or double-click "Start Project Manager.bat")
 * Then open http://localhost:8743
 */
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const net = require('net');
const tls = require('tls');
const crypto = require('crypto');
const store = require('./store');   // file-backed by default; Supabase when SUPABASE_SERVICE_KEY is set

const PORT = 8743;
const HOST = '127.0.0.1';
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
// Resolve public/ from the first place it actually exists — robust whether running
// locally or bundled into a Vercel function (cwd vs __dirname can differ there).
const PUB_DIR = (() => {
  for (const c of [path.join(ROOT, 'public'), path.join(process.cwd(), 'public')]) {
    try { if (fs.existsSync(c)) return c; } catch (e) { /* try next */ }
  }
  return path.join(ROOT, 'public');
})();

const FILES = {
  settings: path.join(DATA_DIR, 'settings.json'),
  templates: path.join(DATA_DIR, 'templates.json'),
  jobs: path.join(DATA_DIR, 'jobs.json'),       // local state per job: stage, category override, planning
  leads: path.join(DATA_DIR, 'leads.json'),     // local state per lead: completed override
  cache: path.join(DATA_DIR, 'zoho-cache.json'), // synced copy of Zoho records + SharePoint lead list
  travelModes: path.join(DATA_DIR, 'travel-modes.json') // PM's fly/drive decision per HWI (published to the travel app)
};

// ---------------------------------------------------------------- utilities

// Storage goes through store.js: local JSON files by default, Supabase (table
// pm_app_state) when the service key is configured — same synchronous interface.
function readJson(file, fallback) { return store.read(file, fallback); }
function writeJson(file, obj) { return store.write(file, obj); }
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ---------------------------------------------------------------- defaults

const DEFAULT_SETTINGS = {
  // Zoho connection
  dc: 'com',                  // data centre: com | eu | in | com.au | jp | ca | sa
  clientId: '',
  clientSecret: '',
  apiDomain: '',              // filled automatically from Zoho token response
  refreshToken: '',           // filled by the OAuth callback
  orgId: '',
  orgName: '',
  // What counts as a "job"
  modules: { estimates: true, salesorders: true, invoices: true, projects: false },
  maxPages: 3,                // pages of 200 records per module per sync
  // Categorisation
  rules: [
    { keyword: 'rental', category: 'rental' },
    { keyword: 'hire', category: 'rental' },
    { keyword: 'rent', category: 'rental' },
    { keyword: 'proof load', category: 'service' },
    { keyword: 'load test', category: 'service' },
    { keyword: 'function test', category: 'service' },
    { keyword: 'test', category: 'service' },
    { keyword: 'inspection', category: 'service' },
    { keyword: 'certif', category: 'service' },
    { keyword: 'calibrat', category: 'service' },
    { keyword: 'service', category: 'service' }
  ],
  defaultCategory: 'sales',
  // Misc
  senderName: '',
  demoMode: 'auto',           // auto = demo until Zoho is connected | on | off
  // Optional outgoing email server — enables one-click "Send now"
  smtp: { host: '', port: 587, user: '', pass: '', fromName: '', fromAddr: '' },
  // Microsoft 365 / SharePoint (PO tracker reads the "Lead List")
  ms: {
    tenant: 'organizations', clientId: '', refreshToken: '',
    siteId: '', siteName: '', listId: '', listName: '',
    map: { po: '', poMode: 'nonempty', poValue: '', company: '', value: '', date: '' }
  },
  // Shop Master (Supabase Postgres, read-only via REST) — the "received jobs" source for the Invoice tracker
  shopmaster: { url: '', key: '' }
};

const DEFAULT_TEMPLATES = {
  questions: [
    { id: 'q1', text: 'Are we going to test the equipment to 125% of WLL?', type: 'boolean' },
    { id: 'q2', text: 'Are we working in metric tonnes or short tons?', type: 'choice', options: ['Metric tonnes', 'Short tons'] },
    { id: 'q3', text: 'What water source is available on site?', type: 'choice', options: ['Mains / hydrant', 'Bowser / tanker', 'Storage tank / IBC', 'Open water (river / lake / sea)', 'Other'] },
    { id: 'q3b', text: 'What is the water connection type?', type: 'text' },
    { id: 'q4', text: 'How far away from the test site is the water source?', type: 'number', unit: 'ft', units: ['ft', 'm'] },
    { id: 'q5', text: 'How much headroom do we have below the hook?', type: 'number', unit: 'ft', units: ['ft', 'm'] },
    { id: 'q6', text: 'Length of the area we are testing in?', type: 'number', unit: 'ft', units: ['ft', 'm'] },
    { id: 'q6b', text: 'Width of the area we are testing in?', type: 'number', unit: 'ft', units: ['ft', 'm'] },
    { id: 'q7', text: 'Will we be function testing at 100%?', type: 'boolean' }
  ],
  emailSubject: 'Planning questions — Job {job}',
  emailIntro: 'Hi {customer},\n\nWe are getting ready for the upcoming job ({job}). To help us plan everything before our team arrives, could you please answer the questions below?',
  emailOutro: 'If anything is unclear, just reply to this email or give us a call.\n\nMany thanks,',
  // Hydro-Wates staff roster — used to fill the Responsibilities table in procedures.
  team: [
    { name: 'Vanoy Harris', contact: '832-367-9279' },
    { name: 'Kaylee Kim', contact: '281-838-5475' }
  ],
  // Standard equipment list — pre-loaded into a procedure's "Equipment & materials" (step 4), then edited per job.
  equipment: [
    'PPE: coveralls (long sleeve), steel-toed boots, safety glasses, gloves, hearing protection, hard hat',
    'Water weight bag(s)',
    'Fill fire hose',
    'Drain hose',
    'Load cell with shackles',
    'Master link',
    'Ball valve',
    'Tool bag',
    'Standard adapter kit',
    'Data logging kit'
  ]
};

// ---- structured-answer helpers (keystone for the answers → loadout/procedure/cert cascade)
const Q_TYPES = ['text', 'number', 'choice', 'boolean'];

// Pull the type/unit/options shape off a raw question, dropping anything irrelevant to its type.
function normQShape(q) {
  q = q || {};
  const type = Q_TYPES.includes(q.type) ? q.type : 'text';
  const out = { type };
  if (type === 'number') {
    const units = (Array.isArray(q.units) ? q.units : []).map(u => String(u).trim()).filter(Boolean).slice(0, 8);
    if (units.length) out.units = units;
    if (q.unit) out.unit = String(q.unit).slice(0, 16);
    else if (units.length) out.unit = units[0];
  }
  if (type === 'choice') {
    out.options = (Array.isArray(q.options) ? q.options : [])
      .map(o => String(o).trim()).filter(Boolean).slice(0, 30);
  }
  return out;
}

// Human-readable form of a structured value — kept on the record so the email
// builder, chips and any plain-text consumer keep working unchanged.
function displayAnswer(q) {
  const v = q.value == null ? '' : String(q.value).trim();
  if (!v) return '';
  if (q.type === 'number') return q.unit ? v + ' ' + q.unit : v;
  if (q.type === 'boolean') return v === 'yes' ? 'Yes' : v === 'no' ? 'No' : '';
  return v;
}

// Normalize one per-job planning question (carries shape + structured value + derived display).
function normJobQ(q) {
  q = q || {};
  const base = {
    id: String(q.id || ('q' + Math.random().toString(36).slice(2, 8))),
    text: String(q.text || '')
  };
  Object.assign(base, normQShape(q));
  // value is the structured answer; fall back to a legacy free-text `answer` for old records.
  base.value = q.value != null ? String(q.value) : (q.answer != null ? String(q.answer) : '');
  base.answer = displayAnswer(base);
  return base;
}

// A job is "archived" (a past project, off the main board) when it's closed —
// marked Complete, or carrying a closed/invoiced Zoho status — unless the PM has
// manually overridden it. l.archiveOverride: 'archived' | 'active' forces either way.
// Hydro-Wates job number (e.g. HWI-26-223). It's the estimate number; pull it from
// the doc number, the reference, or anywhere in the record's searchable text.
const HWI_RE = /\bHW[A-Z]?-\d{2}-\d{2,}\b/i;
function extractHwi(b) {
  // projectName (Zoho "Project") is the HWI on invoices; then the doc number (estimates), reference, text.
  for (const s of [b.projectName, b.number, b.reference, b.haystack]) {
    const m = String(s || '').match(HWI_RE);
    if (m) return m[0].toUpperCase();
  }
  return '';
}

// ---- Shop Master (Supabase Postgres, read-only via the REST API). Received-jobs feed for the Invoice tracker.
async function shopmasterGet(s, path) {
  const sm = (s && s.shopmaster) || {};
  if (!sm.url || !sm.key) { const e = new Error('Shop Master isn’t connected — add the Supabase URL and key in Settings.'); e.code = 'NOT_CONNECTED'; throw e; }
  const url = sm.url.replace(/\/+$/, '') + '/rest/v1/' + path;
  const r = await fetch(url, { headers: { apikey: sm.key, Authorization: 'Bearer ' + sm.key, Accept: 'application/json' } });
  const text = await r.text();
  if (!r.ok) throw new Error('Shop Master (Supabase) returned ' + r.status + ': ' + text.slice(0, 300));
  try { return JSON.parse(text); } catch (e) { throw new Error('Shop Master returned unexpected data.'); }
}

// Call a Supabase RPC (Postgres function) with the same connection as Shop Master.
// Used to publish the PM's fly/drive decision via the scoped set_job_travel_mode()
// function — the anon key can't write cportal_projects directly (RLS), but it is
// granted EXECUTE on that one narrow, security-definer function.
async function shopmasterRpc(s, fn, body) {
  const sm = (s && s.shopmaster) || {};
  if (!sm.url || !sm.key) { const e = new Error('Shop Master isn’t connected — add the Supabase URL and key in Settings.'); e.code = 'NOT_CONNECTED'; throw e; }
  const url = sm.url.replace(/\/+$/, '') + '/rest/v1/rpc/' + fn;
  const r = await fetch(url, {
    method: 'POST',
    headers: { apikey: sm.key, Authorization: 'Bearer ' + sm.key, 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify(body || {})
  });
  const text = await r.text();
  if (!r.ok) { const e = new Error('Supabase RPC ' + fn + ' returned ' + r.status + ': ' + text.slice(0, 300)); e.status = r.status; throw e; }
  try { return text ? JSON.parse(text) : null; } catch (e) { return text; }
}

// Normalize an HWI so "HWI-26-019", "hwi 26 019", "HWI26019" all collapse to one key.
function hwiKey(x) {
  return String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}
// Build a lookup of projected values from the SharePoint Lead List, keyed by the
// HWI number (the Lead List's QuoteNum column) -> ValueTotal. This is an exact join:
// the same HWI that Shop Master uses as job_number is stored on the lead as QuoteNum.
function leadProjectedLookup(cache) {
  const items = (cache.leads && cache.leads.items) || [];
  const byHwi = new Map();   // hwiKey -> value
  for (const it of items) {
    const f = it.fields || {};
    const k = hwiKey(f.QuoteNum);
    if (k.length < 4) continue;
    const value = parseMoney(f.ValueTotal);
    if (value == null) continue;
    if (!byHwi.has(k)) byHwi.set(k, value);   // first (newest) wins on the rare duplicate
  }
  return { byHwi };
}

// Normalize a PO so "#PO10-00877", "PO1000877", "1000877" collapse to one key.
function poNorm(x) {
  return String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^PO/, '');
}

// Index of the SharePoint Lead List, the source of truth for a job's NUMBER and
// CATEGORY. The Lead List's QuoteNum column IS the HWI; its ProjType column
// classifies the job as Service / Rental / Sales. We key by HWI (exact join) and
// also by PO, so a Zoho record that lacks an HWI can still recover one via its PO.
//   byHwi: hwiKey -> { quoteNum, cat, value, company }
//   byPo:  poKey  -> { hwi, cat }
function leadIndex(cache) {
  const items = (cache.leads && cache.leads.items) || [];
  const byHwi = new Map();
  const byPo = new Map();
  const toCat = (pt) => {
    const t = String(pt || '').trim().toLowerCase();
    return (t === 'service' || t === 'rental' || t === 'sales') ? t : null;
  };
  for (const it of items) {
    const f = it.fields || {};
    const quoteNum = String(f.QuoteNum || '').trim();
    const hk = hwiKey(quoteNum);
    const cat = toCat(f.ProjType);
    if (hk.length >= 4 && !byHwi.has(hk)) {
      byHwi.set(hk, {
        quoteNum, cat, value: parseMoney(f.ValueTotal), company: fieldToText(f.Company),
        invoiced: parseMoney(f.InvoicedTotal) != null || !!fieldToText(f.InvoiceDate).trim()   // the Lead List says this job has been billed (invoiced total OR an invoice date)
      });
    }
    const pk = poNorm(f.PONum);
    if (pk.length >= 4 && quoteNum && !byPo.has(pk)) byPo.set(pk, { hwi: quoteNum, cat });
  }
  return { byHwi, byPo };
}

// A Zoho invoice counts as "sent" (= the job is completed) unless it's still a
// draft or has been voided — a draft isn't sent yet, a void was cancelled.
function isSentInvoice(b) {
  if (b.module !== 'invoices') return false;
  const st = String(b.status || '').toLowerCase();
  return st !== 'draft' && st !== 'void';
}

// The set of HWIs that have been INVOICED (sent) in Zoho — i.e. completed jobs.
// Keyed by hwiKey so it matches a job regardless of HWI punctuation.
function invoicedHwisFrom(jobs) {
  const set = new Set();
  for (const b of jobs) {
    if (!isSentInvoice(b)) continue;
    const k = hwiKey(extractHwi(b));
    if (k.length >= 4) set.add(k);
  }
  return set;
}

// Whether a job is completed (invoiced) and how. A job is done — off the active
// board — once a SENT Zoho invoice exists for its HWI (or it IS that invoice), or
// the Lead List records an InvoicedTotal / InvoiceDate for it. "Complete" = an
// invoice was CREATED (paid or not). A manual "keep active" override always wins.
function completedState(b, l, hwi, invoicedHwis, idx) {
  const arch = archiveState(b, l);
  if (arch.how === 'forced-active') return { archived: false, how: 'forced-active' };
  const hk = hwiKey(hwi);
  const invoiced = isSentInvoice(b) ||                                  // the invoice itself (even with no HWI)
    (hk.length >= 4 && invoicedHwis.has(hk)) ||                         // its HWI was invoiced in Zoho
    (hwi ? !!(idx.byHwi.get(hk) || {}).invoiced : false);              // the Lead List says it's billed
  if (arch.archived) return { archived: true, how: arch.how };
  // Multi-invoice jobs (billed in stages) stay on the board through their invoicing —
  // one invoice doesn't mean done. The PM removes them (or clears the flag) when finished.
  if (invoiced && !l.multiInvoice) return { archived: true, how: 'invoiced' };
  return { archived: false, how: arch.how };
}

// Resolve a job's HWI (number) and category using the Lead List as the authority.
// HWI: from the record itself, else recovered from the Lead List by PO.
// Category: a manual override wins; then the Lead List's ProjType; then the old
// keyword rules as a fallback for anything the Lead List doesn't cover.
function resolveJob(b, local, s, idx) {
  let hwi = extractHwi(b);
  if (!hwi) {
    const pk = poNorm(b.reference) || poNorm(b.number);
    const hit = pk.length >= 4 ? idx.byPo.get(pk) : null;
    if (hit) hwi = hit.hwi;
  }
  const l = local[b.key] || {};
  let category, overridden = false, rule = null;
  if (l.categoryOverride) { category = l.categoryOverride; overridden = true; }
  else {
    const leadCat = hwi ? (idx.byHwi.get(hwiKey(hwi)) || {}).cat : null;
    if (leadCat) { category = leadCat; rule = 'Lead List'; }
    else { const c = categorize(b, local, s); category = c.category; rule = c.rule || null; }
  }
  return { hwi, category, categoryOverridden: overridden, categoryRule: rule };
}

// Pull received jobs and cross-reference Zoho Books for invoiced status (HWI first, then PO).
async function shopmasterReceivedJobs(s) {
  const rows = await shopmasterGet(s, 'shopmaster_loadouts?select=job_number,customer_name,created_at,phase,po_number&order=created_at.desc');
  const cache = readJson(FILES.cache, { jobs: {} });
  const zoho = [].concat(...Object.values(cache.jobs || {}).map(v => Array.isArray(v) ? v : [v])).filter(Boolean);
  const projLookup = leadProjectedLookup(cache);
  const travelModes = readJson(FILES.travelModes, {});   // { HWIKEY: 'fly'|'drive' } — the PM's decision per job
  const cleanPo = x => String(x || '').trim().toUpperCase().replace(/^PO[\s#:.-]*/, '');
  function matchZoho(hwi, po) {
    const H = String(hwi || '').toUpperCase();
    const P = cleanPo(po);
    let est = null;
    for (const b of zoho) {
      const num = String(b.number || '').toUpperCase();
      const ref = String(b.reference || '').toUpperCase();
      const proj = String(b.projectName || '').toUpperCase();
      const hay = String(b.haystack || '').toUpperCase();
      const hwiHit = H && (proj === H || num === H || ref === H || hay.includes(H));
      const poHit = P && P.length >= 4 && (cleanPo(b.reference) === P);
      if (b.module === 'invoices' && (hwiHit || poHit)) {
        return { invoiced: true, how: hwiHit ? 'hwi→invoice' : 'po→invoice', invoice: b.number, status: b.status, invoiceDate: b.date || '', total: b.total, currency: b.currency, ref: b.reference || '' };
      }
      if (b.module === 'estimates' && hwiHit) est = b;
    }
    if (est) {
      const inv = /invoic/i.test(est.status || '');
      return { invoiced: inv, how: inv ? 'estimate(invoiced)' : 'estimate(open)', invoice: est.number, status: est.status, invoiceDate: est.date || '', total: est.total, currency: est.currency, ref: est.reference || '' };
    }
    return { invoiced: false, how: 'no-match', invoice: null, status: null, invoiceDate: '', ref: '' };
  }
  const byJob = new Map();   // a job can have several loadout rows; keep the latest (rows are date-desc)
  for (const r of rows) {
    const k = String(r.job_number || '').toUpperCase();
    if (k && !byJob.has(k)) byJob.set(k, r);
  }
  return [...byJob.values()].map(r => {
    const m = matchZoho(r.job_number, r.po_number);
    const poNumber = (m.ref || r.po_number || '');   // the PO — from the matched Zoho invoice's reference, else Shop Master
    // For jobs awaiting an invoice, pull the projected value from the Lead List,
    // matched exactly by HWI (Shop Master job_number == Lead List QuoteNum).
    let projected = null, projectedHow = null;
    if (!m.invoiced && m.total == null) {
      const hk = hwiKey(r.job_number);
      const v = hk.length >= 4 ? projLookup.byHwi.get(hk) : null;
      if (v != null) { projected = v; projectedHow = 'hwi'; }
    }
    return {
      hwi: String(r.job_number || '').toUpperCase(),
      customer: r.customer_name || '', phase: r.phase || '',
      receivedAt: r.created_at || '', po: r.po_number || '',
      invoiced: m.invoiced, matchHow: m.how, invoice: m.invoice, invoiceStatus: m.status,
      invoiceDate: m.invoiceDate || '', total: m.total, currency: m.currency,
      projected, projectedHow,
      travelMode: travelModes[hwiKey(r.job_number)] || null,   // fly | drive | null (undecided)
      poNumber
    };
  });
}

// Shop Master received jobs, shaped as dashboard "blocks" so they show as cards
// alongside the Zoho jobs (used by the dashboard feed + single-job lookup). These
// are jobs that exist in Shop Master but may have no Zoho record yet (awaiting an
// invoice) — exactly the ones a fly/drive decision is about. Safe: returns [] if
// Shop Master isn't connected or the fetch fails, so the dashboard still loads.
async function receivedDashJobs(s) {
  let jobs = [];
  try { jobs = await shopmasterReceivedJobs(s); }
  catch (e) { return []; }
  return jobs.map(j => {
    const hwi = j.hwi || '';
    // Awaiting jobs carry a projected value (from the Lead List); invoiced ones the actual total.
    const total = (j.total != null ? j.total : (j.projected != null ? j.projected : null));
    return {
      key: 'received:' + hwiKey(hwi),
      module: 'received', id: hwi, number: hwi, projectName: hwi,
      customer: j.customer || '', customerId: '',
      date: String(j.receivedAt || '').slice(0, 10),
      total, currency: j.currency || '',
      status: j.invoiced ? (j.invoiceStatus || 'invoiced') : (j.phase || 'received'),
      reference: j.poNumber || j.po || '', lineItems: [],
      received: true, invoiced: !!j.invoiced, phase: j.phase || '',
      projectedValue: (j.total == null ? j.projected : null),
      travelMode: j.travelMode || null,
      haystack: [hwi, j.customer, j.poNumber, j.phase].join(' || ').toLowerCase()
    };
  });
}

const CLOSED_STATUSES = new Set(['paid', 'closed', 'invoiced', 'void', 'declined', 'expired', 'rejected']);
function archiveState(b, l) {
  l = l || {};
  if (l.archiveOverride === 'archived') return { archived: true, how: 'manual' };
  if (l.archiveOverride === 'active') return { archived: false, how: 'forced-active' };
  if ((l.stage || 'new') === 'complete') return { archived: true, how: 'complete' };
  if (CLOSED_STATUSES.has(String(b.status || '').toLowerCase())) return { archived: true, how: 'invoiced' };
  return { archived: false, how: 'auto' };
}

function getSettings() {
  const s = readJson(FILES.settings, {});
  const merged = Object.assign({}, DEFAULT_SETTINGS, s);
  merged.modules = Object.assign({}, DEFAULT_SETTINGS.modules, s.modules || {});
  merged.smtp = Object.assign({}, DEFAULT_SETTINGS.smtp, s.smtp || {});
  merged.ms = Object.assign({}, DEFAULT_SETTINGS.ms, s.ms || {});
  merged.ms.map = Object.assign({}, DEFAULT_SETTINGS.ms.map, (s.ms || {}).map || {});
  merged.shopmaster = Object.assign({}, DEFAULT_SETTINGS.shopmaster, s.shopmaster || {});
  if (!Array.isArray(merged.rules)) merged.rules = DEFAULT_SETTINGS.rules;
  return merged;
}
function getTemplates() {
  const t = readJson(FILES.templates, {});
  const merged = Object.assign({}, DEFAULT_TEMPLATES, t);
  if (!Array.isArray(merged.questions) || merged.questions.length === 0) merged.questions = DEFAULT_TEMPLATES.questions;
  if (!Array.isArray(merged.team)) merged.team = [];
  if (!Array.isArray(merged.equipment)) merged.equipment = [];
  return merged;
}
function ensureDefaults() {
  if (store.mode() === 'file') fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!store.has(FILES.settings)) writeJson(FILES.settings, DEFAULT_SETTINGS);
  if (!store.has(FILES.templates)) writeJson(FILES.templates, DEFAULT_TEMPLATES);
  if (!store.has(FILES.jobs)) writeJson(FILES.jobs, {});
  if (!store.has(FILES.leads)) writeJson(FILES.leads, {});
}

// ---------------------------------------------------------------- Zoho API

const MODULES = {
  estimates:   { list: '/estimates',   arr: 'estimates',   one: 'estimate',   id: 'estimate_id',   num: 'estimate_number' },
  salesorders: { list: '/salesorders', arr: 'salesorders', one: 'salesorder', id: 'salesorder_id', num: 'salesorder_number' },
  invoices:    { list: '/invoices',    arr: 'invoices',    one: 'invoice',    id: 'invoice_id',    num: 'invoice_number' },
  projects:    { list: '/projects',    arr: 'projects',    one: 'project',    id: 'project_id',    num: 'project_name' }
};

let tokenCache = { token: null, exp: 0 };

const accountsBase = (dc) => 'https://accounts.zoho.' + dc;
const apiBase = (s) => (s.apiDomain || 'https://www.zohoapis.' + s.dc);
const redirectUri = () => 'http://localhost:' + PORT + '/oauth/callback';

async function tokenRequest(params, dc) {
  const res = await fetch(accountsBase(dc) + '/oauth/v2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params)
  });
  const j = await res.json().catch(() => ({}));
  if (j.error) throw new Error('Zoho sign-in error: ' + j.error);
  return j;
}

async function getAccessToken() {
  const s = getSettings();
  if (!s.refreshToken) { const e = new Error('Not connected to Zoho Books'); e.code = 'NOT_CONNECTED'; throw e; }
  if (tokenCache.token && Date.now() < tokenCache.exp - 60000) return tokenCache.token;
  const j = await tokenRequest({
    refresh_token: s.refreshToken, client_id: s.clientId, client_secret: s.clientSecret,
    grant_type: 'refresh_token'
  }, s.dc);
  tokenCache = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  if (j.api_domain && j.api_domain !== s.apiDomain) {
    const cur = readJson(FILES.settings, {}); cur.apiDomain = j.api_domain; writeJson(FILES.settings, cur);
  }
  return tokenCache.token;
}

async function zohoGet(pathname, params = {}, withOrg = true) {
  const s = getSettings();
  const token = await getAccessToken();
  const u = new URL(apiBase(s) + '/books/v3' + pathname);
  if (withOrg && s.orgId) u.searchParams.set('organization_id', s.orgId);
  for (const [k, v] of Object.entries(params)) u.searchParams.set(k, String(v));
  const res = await fetch(u, { headers: { Authorization: 'Zoho-oauthtoken ' + token } });
  if (res.status === 401) { tokenCache = { token: null, exp: 0 }; throw new Error('Zoho rejected the session (401). Try Sync again, or reconnect in Settings.'); }
  if (res.status === 429) throw new Error('Zoho rate limit reached — wait a minute and sync again.');
  const j = await res.json().catch(() => ({}));
  if (j.code !== undefined && j.code !== 0) throw new Error('Zoho: ' + (j.message || ('error code ' + j.code)));
  if (!res.ok) throw new Error('Zoho HTTP ' + res.status);
  return j;
}

// ---------------------------------------------------------------- sync

let syncState = { running: false, message: '', done: 0, total: 0, error: null };

function normalizeRow(mod, m, r) {
  const b = {
    key: mod + ':' + r[m.id],
    module: mod,
    id: String(r[m.id]),
    number: String(r[m.num] || ''),
    customer: r.customer_name || r.company_name || '',
    customerId: r.customer_id ? String(r.customer_id) : '',
    date: r.date || (r.created_time ? String(r.created_time).slice(0, 10) : ''),
    total: (r.total !== undefined && r.total !== null) ? r.total : null,
    currency: r.currency_code || '',
    status: r.status || r.order_status || '',
    reference: r.reference_number || '',
    projectName: r.project_name || '',   // Zoho "Project" — this is the HWI job number on invoices
    email: r.email || '',
    lastModified: r.last_modified_time || ''
  };
  b.sig = [b.date, b.total, b.status, b.lastModified, b.projectName].join('|');
  return b;
}

async function runSync() {
  if (syncState.running) return;
  syncState = { running: true, message: 'Fetching job lists from Zoho Books…', done: 0, total: 0, error: null };
  try {
    const s = getSettings();
    if (!s.refreshToken) { const e = new Error('Not connected'); e.code = 'NOT_CONNECTED'; throw e; }
    if (!s.orgId) throw new Error('No organisation selected — pick one in Settings.');
    const cache = readJson(FILES.cache, { jobs: {}, lastSync: null });
    const found = {};

    for (const [mod, m] of Object.entries(MODULES)) {
      if (!s.modules[mod]) continue;
      let page = 1;
      const maxPages = Math.min(Math.max(Number(s.maxPages) || 3, 1), 10);
      while (page <= maxPages) {
        const params = { page, per_page: 200 };
        if (mod !== 'projects') { params.sort_column = 'date'; params.sort_order = 'D'; }
        const j = await zohoGet(m.list, params);
        for (const r of (j[m.arr] || [])) {
          const b = normalizeRow(mod, m, r);
          found[b.key] = b;
        }
        if (!(j.page_context && j.page_context.has_more_page)) break;
        page++;
      }
    }

    // Re-use cached line-item details where the record hasn't changed.
    for (const b of Object.values(found)) {
      const old = cache.jobs[b.key];
      if (old && old.sig === b.sig && old.haystack) {
        b.lineItems = old.lineItems || [];
        b.haystack = old.haystack;
        b.email = b.email || old.email || '';
      }
    }

    // Fetch details (line items) for new/changed records — needed for categorisation.
    const need = Object.values(found).filter(b => !b.haystack && b.module !== 'projects');
    syncState.total = need.length;
    if (need.length) syncState.message = 'Fetching job details…';
    let idx = 0;
    const workers = Array.from({ length: 4 }, async () => {
      while (idx < need.length) {
        const b = need[idx++];
        try {
          const m = MODULES[b.module];
          const j = await zohoGet(m.list + '/' + b.id);
          const d = j[m.one] || {};
          b.lineItems = (d.line_items || []).map(li => ({
            name: li.name || '', description: li.description || '',
            quantity: li.quantity, rate: li.rate, total: li.item_total
          }));
          if (!b.email) {
            b.email = d.email || ((d.contact_persons_details || [])[0] || {}).email || '';
          }
          const cf = (d.custom_fields || []).map(c => ((c.label ? c.label + ': ' : '') + (c.value === undefined || c.value === null ? '' : c.value))).join(' | ');
          b.haystack = [b.number, b.customer, b.reference, d.notes || '', d.subject || '', cf,
            ...b.lineItems.map(li => li.name + ' ' + li.description)].join(' || ').toLowerCase();
        } catch (e) {
          b.lineItems = b.lineItems || [];
          b.haystack = [b.number, b.customer, b.reference].join(' || ').toLowerCase();
          b.detailError = String(e.message || e);
        }
        syncState.done++;
        await sleep(250); // stay well inside Zoho's rate limits
      }
    });
    await Promise.all(workers);

    for (const b of Object.values(found)) {
      if (!b.haystack) b.haystack = [b.number, b.customer, b.reference, b.status].join(' || ').toLowerCase();
      if (!b.lineItems) b.lineItems = [];
    }

    cache.jobs = found;
    cache.lastSync = new Date().toISOString();
    writeJson(FILES.cache, cache);

    // Refresh the SharePoint lead list too, if connected (failures show on the PO tracker's own Refresh).
    if (s.ms && s.ms.refreshToken && s.ms.siteId && s.ms.listId) {
      syncState.message = 'Refreshing the SharePoint lead list…';
      try { await runLeadsSync(); } catch (e) { /* non-fatal */ }
    }

    syncState = { running: false, done: 0, total: 0, error: null, message: '' };
  } catch (e) {
    const msg = (e && e.code === 'NOT_CONNECTED')
      ? 'Not connected to Zoho Books yet — open Settings to connect.'
      : String((e && e.message) || e);
    syncState = { running: false, message: '', done: 0, total: 0, error: msg };
  }
}

// ---------------------------------------------------------------- categorisation

function categorize(b, local, s) {
  const l = local[b.key] || {};
  if (l.categoryOverride) return { category: l.categoryOverride, overridden: true };
  const hay = (b.haystack || (b.number + ' ' + b.customer + ' ' + b.reference)).toLowerCase();
  for (const r of (s.rules || [])) {
    if (!r || !r.keyword) continue;
    if (hay.includes(String(r.keyword).toLowerCase())) {
      return { category: r.category || s.defaultCategory, overridden: false, rule: r.keyword };
    }
  }
  return { category: s.defaultCategory || 'sales', overridden: false };
}

// ---------------------------------------------------------------- contacts on file

async function contactsForCustomer(customerId) {
  const cache = readJson(FILES.cache, { jobs: {}, lastSync: null });
  const store = cache.contactsByCustomer || {};
  const hit = store[customerId];
  if (hit && (Date.now() - new Date(hit.at).getTime()) < 24 * 3600 * 1000) return hit.contacts;
  const j = await zohoGet('/contacts/' + customerId);
  const ct = j.contact || {};
  const out = [];
  for (const p of (ct.contact_persons || [])) {
    if (!p.email) continue;
    out.push({
      name: [p.salutation, p.first_name, p.last_name].filter(Boolean).join(' ').trim() || p.email,
      email: p.email,
      phone: p.phone || p.mobile || '',
      designation: p.designation || '',
      isPrimary: !!p.is_primary_contact
    });
  }
  if (!out.length && ct.email) {
    out.push({ name: ct.contact_name || ct.email, email: ct.email, phone: ct.phone || '', designation: '', isPrimary: true });
  }
  store[customerId] = { at: new Date().toISOString(), contacts: out };
  cache.contactsByCustomer = store;
  writeJson(FILES.cache, cache);
  return out;
}

// ---------------------------------------------------------------- outgoing email (zero-dependency SMTP client)

const b64 = (s) => Buffer.from(s, 'utf8').toString('base64');
const rfc2047 = (s) => /^[\x20-\x7e]*$/.test(s) ? s : '=?UTF-8?B?' + b64(s) + '?=';
const wrap76 = (s) => s.replace(/(.{76})/g, '$1\r\n');

function smtpDialogue(initialSocket, host, timeoutMs) {
  let socket = initialSocket;
  let buf = '';
  let pending = null;

  const fail = (err) => {
    const p = pending; pending = null;
    if (p) p.reject(err instanceof Error ? err : new Error(String(err)));
    try { socket.destroy(); } catch (e) { /* already gone */ }
  };
  const tryResolve = () => {
    if (!pending) return;
    const lines = buf.split(/\r?\n/);
    for (let i = 0; i < lines.length - 1; i++) {
      if (/^\d{3}( |$)/.test(lines[i])) {       // final line of a (possibly multi-line) reply
        const reply = lines.slice(0, i + 1);
        buf = lines.slice(i + 1).join('\r\n');
        const p = pending; pending = null;
        p.resolve({ code: Number(reply[i].slice(0, 3)), text: reply.join(' ').trim() });
        return;
      }
    }
  };
  const hook = (s) => {
    s.on('data', (d) => { buf += d.toString('utf8'); tryResolve(); });
    s.on('error', fail);
    s.setTimeout(timeoutMs, () => fail(new Error('The email server stopped responding (timeout).')));
  };
  hook(socket);

  return {
    read: () => new Promise((resolve, reject) => { pending = { resolve, reject }; tryResolve(); }),
    write: (line) => socket.write(line + '\r\n'),
    cmd(line) { this.write(line); return this.read(); },
    upgradeTls: (servername) => new Promise((resolve, reject) => {
      const clear = socket;
      clear.removeAllListeners('data'); clear.removeAllListeners('error'); clear.setTimeout(0);
      const secure = tls.connect({ socket: clear, servername }, () => resolve());
      secure.once('error', reject);
      socket = secure;
      buf = '';
      hook(secure);
    }),
    end: () => { try { socket.end(); } catch (e) { /* fine */ } }
  };
}

async function sendSmtpMail(smtp, mail) {
  const host = String((smtp && smtp.host) || '').trim();
  const port = Number((smtp && smtp.port) || 587);
  if (!host || !smtp.user || !smtp.pass) {
    throw new Error('Email sending is not set up — fill in “Email sending” under Settings (or use “Open email draft”).');
  }
  const implicitTls = port === 465;
  const timeoutMs = 25000;

  const socket = await new Promise((resolve, reject) => {
    const s = implicitTls
      ? tls.connect({ host, port, servername: host }, () => resolve(s))
      : net.connect({ host, port }, () => resolve(s));
    s.once('error', reject);
    s.setTimeout(timeoutMs, () => { s.destroy(); reject(new Error('Could not reach ' + host + ':' + port + ' — check the server name and port.')); });
  });

  const c = smtpDialogue(socket, host, timeoutMs);
  const expect = (r, codes, what) => {
    if (!codes.includes(r.code)) {
      const friendly = (r.code === 535 || r.code === 534)
        ? 'The email server rejected the sign-in. Check the user and password — Microsoft 365 and Gmail usually need an “app password”, and your IT admin may need to enable Authenticated SMTP for the mailbox. (Server said: ' + r.text + ')'
        : 'Email server said: ' + r.text + (what ? ' [during ' + what + ']' : '');
      throw new Error(friendly);
    }
    return r;
  };

  try {
    expect(await c.read(), [220], 'connect');
    expect(await c.cmd('EHLO [127.0.0.1]'), [250], 'EHLO');
    if (!implicitTls) {
      expect(await c.cmd('STARTTLS'), [220], 'STARTTLS');
      await c.upgradeTls(host);
      expect(await c.cmd('EHLO [127.0.0.1]'), [250], 'EHLO');
    }
    const auth = await c.cmd('AUTH LOGIN');
    if (auth.code === 334) {
      expect(await c.cmd(b64(smtp.user)), [334], 'sign-in (user)');
      expect(await c.cmd(b64(smtp.pass)), [235], 'sign-in (password)');
    } else {
      expect(await c.cmd('AUTH PLAIN ' + b64('\0' + smtp.user + '\0' + smtp.pass)), [235], 'sign-in');
    }
    const fromAddr = (smtp.fromAddr || smtp.user).trim();
    expect(await c.cmd('MAIL FROM:<' + fromAddr + '>'), [250], 'sender');
    for (const rcpt of mail.to) {
      expect(await c.cmd('RCPT TO:<' + rcpt + '>'), [250, 251], 'recipient ' + rcpt);
    }
    expect(await c.cmd('DATA'), [354], 'DATA');
    const headers = [
      'From: ' + (smtp.fromName ? rfc2047(smtp.fromName) + ' <' + fromAddr + '>' : fromAddr),
      'To: ' + mail.to.join(', '),
      'Subject: ' + rfc2047(mail.subject),
      'Date: ' + new Date().toUTCString(),
      'Message-ID: <hwpm-' + Date.now() + '-' + Math.random().toString(36).slice(2) + '@' + host + '>',
      'MIME-Version: 1.0',
      'Content-Type: text/' + (mail.html ? 'html' : 'plain') + '; charset=utf-8',
      'Content-Transfer-Encoding: base64'
    ];
    c.write(headers.join('\r\n') + '\r\n\r\n' + wrap76(b64(mail.html || mail.body)) + '\r\n.');
    expect(await c.read(), [250], 'delivery');
    c.write('QUIT');
    c.end();
    return true;
  } catch (e) {
    c.end();
    throw e;
  }
}

// ---------------------------------------------------------------- Microsoft 365 / SharePoint (PO tracker)

const MS_SCOPES = 'offline_access openid profile https://graph.microsoft.com/Sites.Read.All https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Send';
const msRedirectUri = () => 'http://localhost:' + PORT + '/ms/callback';
let msAuthPending = null;                      // { verifier, state } during sign-in
let msTokenCache = { token: null, exp: 0 };

const msLoginBase = (tenant) => 'https://login.microsoftonline.com/' + (tenant || 'organizations') + '/oauth2/v2.0';

async function getMsToken() {
  const s = getSettings();
  if (!(s.ms && s.ms.refreshToken)) { const e = new Error('Not connected to Microsoft 365 — connect in Settings.'); e.code = 'NOT_CONNECTED'; throw e; }
  if (msTokenCache.token && Date.now() < msTokenCache.exp - 60000) return msTokenCache.token;
  const r = await fetch(msLoginBase(s.ms.tenant) + '/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: s.ms.clientId, grant_type: 'refresh_token',
      refresh_token: s.ms.refreshToken, scope: MS_SCOPES
    })
  });
  const j = await r.json().catch(() => ({}));
  if (!j.access_token) {
    const why = String(j.error_description || j.error || 'no token returned').split(/\r?\n/)[0].slice(0, 200);
    throw new Error('Microsoft 365 sign-in has expired — reconnect in Settings. (' + why + ')');
  }
  msTokenCache = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
  if (j.refresh_token && j.refresh_token !== s.ms.refreshToken) {   // Microsoft rotates refresh tokens
    const cur = readJson(FILES.settings, {});
    cur.ms = Object.assign({}, cur.ms, { refreshToken: j.refresh_token });
    writeJson(FILES.settings, cur);
  }
  return msTokenCache.token;
}

async function graphGet(pathOrUrl) {
  const token = await getMsToken();
  const url = pathOrUrl.startsWith('https://') ? pathOrUrl : 'https://graph.microsoft.com/v1.0' + pathOrUrl;
  const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) {
    if (r.status === 401) msTokenCache = { token: null, exp: 0 };
    const msg = (j.error && (j.error.message || j.error.code)) || ('HTTP ' + r.status);
    if (r.status === 403) throw new Error('Microsoft 365 refused access: ' + msg + ' — check that the app registration has the Sites.Read.All (Delegated) permission.');
    throw new Error('Microsoft 365: ' + msg);
  }
  return j;
}

// Send email via Microsoft Graph as the signed-in user (delegated Mail.Send). Preferred over SMTP.
async function graphSendMail(s, mail) {
  const token = await getMsToken();
  const isHtml = !!mail.html;
  const payload = {
    message: {
      subject: String(mail.subject || ''),
      body: { contentType: isHtml ? 'HTML' : 'Text', content: String(mail.html || mail.body || '') },
      toRecipients: (mail.to || []).map(a => ({ emailAddress: { address: String(a) } }))
    },
    saveToSentItems: true
  };
  const r = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (r.status === 202) return true;          // Graph returns 202 Accepted on success
  const j = await r.json().catch(() => ({}));
  if (r.status === 401) msTokenCache = { token: null, exp: 0 };
  const msg = (j.error && (j.error.message || j.error.code)) || ('HTTP ' + r.status);
  const e = new Error('Microsoft 365 (Graph) could not send the email: ' + msg +
    (r.status === 403 ? ' — the app needs the Mail.Send (Delegated) permission; ask IT to add it to the app registration, then reconnect Microsoft 365 in Settings.' : ''));
  e.code = 'GRAPH_SEND';
  throw e;
}

// Pick the send method: Microsoft Graph if M365 is connected (preferred), else SMTP fallback.
// Send as a SPECIFIC person, using THEIR own Microsoft access token (obtained when
// they signed in with Microsoft). The email leaves from — and lands in the Sent
// Items of — that person's mailbox, never a shared account.
async function graphSendMailAsUser(accessToken, mail) {
  const isHtml = !!mail.html;
  const payload = {
    message: {
      subject: String(mail.subject || ''),
      body: { contentType: isHtml ? 'HTML' : 'Text', content: String(mail.html || mail.body || '') },
      toRecipients: (mail.to || []).map(a => ({ emailAddress: { address: String(a) } }))
    },
    saveToSentItems: true
  };
  const r = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  if (r.status === 202) return true;
  const j = await r.json().catch(() => ({}));
  const msg = (j.error && (j.error.message || j.error.code)) || ('HTTP ' + r.status);
  throw new Error('Microsoft 365 couldn’t send the email: ' + msg +
    (r.status === 401 ? ' — your Microsoft sign-in expired; sign out and back in with Microsoft.'
     : r.status === 403 ? ' — your account is missing the Mail.Send permission.' : ''));
}

// Email goes out as the LOGGED-IN user via their own Microsoft token. No shared
// account: if the person didn't sign in with Microsoft there's no mailbox to send
// from, so we ask them to — nothing is ever sent from someone else's address.
async function sendMail(s, mail, userToken) {
  if (userToken) return await graphSendMailAsUser(userToken, mail);
  throw new Error('To send from your own mailbox, sign in with Microsoft first (then try again).');
}

function fieldToText(v) {
  if (v === undefined || v === null) return '';
  if (Array.isArray(v)) return v.map(fieldToText).filter(Boolean).join(', ');
  if (typeof v === 'object') return String(v.LookupValue || v.Label || v.Email || v.DisplayName || v.displayName || '');
  return String(v);
}
function parseMoney(v) {
  const t = fieldToText(v);
  if (!t) return null;
  const n = Number(t.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}

async function runLeadsSync() {
  const s = getSettings();
  if (!(s.ms && s.ms.refreshToken)) throw new Error('Not connected to Microsoft 365 — connect in Settings.');
  if (!s.ms.siteId || !s.ms.listId) throw new Error('Pick the SharePoint site and the lead list in Settings first.');
  let url = '/sites/' + encodeURIComponent(s.ms.siteId) + '/lists/' + encodeURIComponent(s.ms.listId) + '/items?$expand=fields&$top=200';
  const items = [];
  while (url) {
    const j = await graphGet(url);
    for (const it of (j.value || [])) {
      items.push({ id: String(it.id), created: it.createdDateTime || '', modified: it.lastModifiedDateTime || '', fields: it.fields || {} });
    }
    url = j['@odata.nextLink'] || '';
    if (items.length >= 5000) break;
  }
  const cache = readJson(FILES.cache, { jobs: {}, lastSync: null });
  cache.leads = { at: new Date().toISOString(), items, listName: s.ms.listName };
  writeJson(FILES.cache, cache);
  return items.length;
}

function leadsAreDemo(s) {
  if (s.demoMode === 'on') return true;
  if (s.demoMode === 'off') return false;
  return !(s.ms && s.ms.refreshToken);
}

function demoLeads() {
  const d = (n) => new Date(Date.now() - n * 86400000).toISOString().slice(0, 10);
  return [
    { id: 'L1', title: 'Berth crane proof load test', company: 'Port Terminal Services LLC', po: 'PO-88412', value: 12400, date: d(9) },
    { id: 'L2', title: 'Pedestal crane recert + function test', company: 'Acadia Offshore', po: 'ACA-PO-1190', value: 15200, date: d(26) },
    { id: 'L3', title: '10 t water weights purchase', company: 'Marina Del Sol', po: 'MDS-2071', value: 11600, date: d(36) },
    { id: 'L4', title: '40 t mobile crane test', company: 'Gulf Crane Services', po: 'PO-7741', value: 4850, date: d(7) },
    { id: 'L5', title: 'Davit tests — dry dock 2', company: 'Henderson Shipyards', po: 'HS-PO-3327', value: 6300, date: d(13) },
    { id: 'L6', title: 'Load cell + shackles', company: 'Texas Lifting Solutions', po: 'TL-2210', value: 9750, date: d(16) },
    { id: 'L7', title: '3-week water bag rental', company: 'Pelican Energy Partners', po: 'PEP-0457', value: 21000, date: d(25) },
    { id: 'L8', title: '50 t kit purchase', company: 'Bayou Marine Group', po: 'BM-PO-1102', value: 38900, date: d(8) }
  ];
}

// Builds the lead rows (PO-filtered, invoice-matched, overrides applied) — used by
// both the PO tracker page and the dashboard.
function computeLeads(s) {
  const overrides = readJson(FILES.leads, {});
  const cache = readJson(FILES.cache, {});
  const demo = leadsAreDemo(s);
  let rows = [];
  let mapWarning = '';
  if (demo) {
    rows = demoLeads();
  } else {
    const items = (cache.leads && cache.leads.items) || [];
    const map = s.ms.map || {};
    if (!map.po) mapWarning = 'Pick which column holds the PO (Settings → SharePoint lead list) so only jobs with a PO are shown — showing every lead for now.';
    for (const it of items) {
      const f = it.fields || {};
      if (map.po) {
        const v = f[map.po];
        const txt = fieldToText(v).trim();
        let okPo;
        if (map.poMode === 'yes') okPo = v === true || /^(yes|true|1)$/i.test(txt);
        else if (map.poMode === 'equals') okPo = txt !== '' && txt.toLowerCase() === String(map.poValue || '').trim().toLowerCase();
        else okPo = txt !== '' && !/^(no|false|0)$/i.test(txt);
        if (!okPo) continue;
      }
      rows.push({
        id: String(it.id),
        title: fieldToText(f.Title),
        hwi: String(f.QuoteNum || '').trim().toUpperCase(),   // the job number (Lead List QuoteNum)
        company: map.company ? fieldToText(f[map.company]) : '',
        po: fieldToText(map.po ? f[map.po] : f.PONum),   // PO received — from the mapped column, else the Lead List PONum

        value: map.value ? parseMoney(f[map.value]) : null,
        date: (map.date ? fieldToText(f[map.date]) : String(it.created || '')).slice(0, 10)
      });
    }
  }
  const invoices = currentJobs(s).filter(b => b.module === 'invoices');
  const norm = (x) => String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const leads = rows.map(r => {
    const o = overrides[r.id] || {};
    const np = norm(r.po);
    let inv = null;
    if (np.length >= 3) {
      inv = invoices.find(b => {
        const nr = norm(b.reference), nn = norm(b.number);
        return (nr && (nr === np || (np.length >= 4 && nr.includes(np)) || (nr.length >= 4 && np.includes(nr)))) || (nn && nn === np);
      }) || null;
    }
    let completed, how;
    if (o.completedOverride === true) { completed = true; how = 'manual'; }
    else if (o.completedOverride === false) { completed = false; how = 'forced-open'; }
    else { completed = !!inv; how = inv ? 'invoice' : null; }
    return Object.assign({}, r, {
      completed, how,
      overridden: o.completedOverride === true || o.completedOverride === false,
      invoice: inv ? { number: inv.number, date: inv.date, total: inv.total, currency: inv.currency, status: inv.status } : null
    });
  });
  leads.sort((a, b) => (a.completed === b.completed)
    ? String(b.date || '').localeCompare(String(a.date || ''))
    : (a.completed ? 1 : -1));
  return {
    leads, demo,
    listName: demo ? 'Lead List (demo)' : (s.ms.listName || ''),
    lastSync: demo ? null : ((cache.leads && cache.leads.at) || null),
    mapWarning
  };
}

// Leads that are still open ("PO received, not yet invoiced") shown as dashboard jobs.
function leadAsJob(r, demo) {
  return {
    key: 'lead:' + r.id, module: 'lead', id: r.id,
    number: r.po || r.hwi || ('Lead ' + r.id),
    projectName: r.hwi || '',                       // so extractHwi() picks up the job number
    customer: r.company || '(no customer)',         // Title is a timestamp, not a name — don't use it
    customerId: '', date: r.date || '',
    total: (r.value === undefined ? null : r.value), currency: '',
    status: 'PO received', reference: r.po || '', email: '', lastModified: '',
    lineItems: [], demo: !!demo, lead: true, leadTitle: '',   // no subtitle line — keep cards uniform
    haystack: [r.hwi, r.title, r.company, r.po].join(' || ').toLowerCase()
  };
}
function openLeadJobs(s) {
  try {
    const c = computeLeads(s);
    const local = readJson(FILES.jobs, {});
    // A dashboard card must be a real, active job: it carries a job number (HWI)
    // AND a PO (the customer has committed), and it hasn't been invoiced yet. Leads
    // with no QuoteNum are raw inquiries; leads with no PO are quotes not yet won —
    // neither gets a card. EXCEPTION: a multi-invoice job stays on the board through
    // its invoicing even once the Lead List/PO reads "completed". (Demo leads exempt.)
    return c.leads
      .filter(l => {
        if (!(c.demo || (l.hwi && String(l.po || '').trim() !== ''))) return false;
        const multi = !!(local['lead:' + l.id] && local['lead:' + l.id].multiInvoice);
        return !l.completed || multi;
      })
      .map(l => leadAsJob(l, c.demo));
  } catch (e) { return []; }
}

// ---------------------------------------------------------------- demo data

function demoJobs() {
  const day = 86400000;
  const d = (n) => new Date(Date.now() - n * day).toISOString().slice(0, 10);
  const mk = (key, module, number, customer, date, total, status, email, items, reference, contacts) => {
    const lineItems = items.map(([name, quantity, rate]) => ({ name, description: '', quantity, rate, total: quantity * rate }));
    return {
      key, module, id: key.split(':')[1], number, customer, date, total, currency: 'USD', status,
      reference: reference || '', email, lastModified: '', sig: 'demo', demo: true, lineItems,
      contacts: contacts || [],
      haystack: [number, customer, reference || '', ...items.map(i => i[0])].join(' || ').toLowerCase()
    };
  };
  const person = (name, designation, email, isPrimary) => ({ name, designation, email, phone: '', isPrimary: !!isPrimary });
  return [
    mk('demo:1', 'invoices', 'INV-1058', 'Port Terminal Services LLC', d(3), 12400, 'paid', 'ops@portterminal.example',
      [['Water bag rental — 6 × 35 t bags, 4 weeks', 6, 1800], ['Pump & hose hire', 1, 1600]], 'PO-88412',
      [person('Maria Sandoval', 'Operations Manager', 'maria@portterminal.example', true),
       person('Greg Holt', 'Dock Foreman', 'greg.holt@portterminal.example')]),
    mk('demo:2', 'estimates', 'EST-00214', 'Gulf Crane Services', d(5), 4850, 'accepted', 'projects@gulfcrane.example',
      [['Proof load test to 125% WLL — 40 t mobile crane', 1, 4850]], 'PO-7741',
      [person('Dana Whitfield', 'Projects Manager', 'dana@gulfcrane.example', true),
       person('Luis Ortega', 'Yard Supervisor', 'luis@gulfcrane.example')]),
    mk('demo:3', 'salesorders', 'SO-0092', 'Bayou Marine Group', d(7), 38900, 'open', 'purchasing@bayoumarine.example',
      [['50 t water weight kit (bag, harness, hoses)', 2, 17500], ['Sling set 25 t', 2, 1950]], '',
      [person('Paul Thibodeaux', 'Purchasing Manager', 'paul@bayoumarine.example', true)]),
    mk('demo:4', 'invoices', 'INV-1062', 'Stark Industrial Cranes', d(10), 7150, 'sent', 'mike@starkindustrial.example',
      [['100 t water bag hire — 2 weeks', 1, 5400], ['Delivery & collection', 1, 1750]], '',
      [person('Mike Stark', 'Owner', 'mike@starkindustrial.example', true),
       person('Renee Calloway', 'Office Administrator', 'renee@starkindustrial.example')]),
    mk('demo:5', 'estimates', 'EST-00217', 'Henderson Shipyards', d(12), 6300, 'sent', 'docks@hendersonship.example',
      [['Davit proof load test to 125% WLL — 8 stations', 1, 6300]], 'Dry dock 2',
      [person('Alan Reyes', 'Dockmaster', 'alan.reyes@hendersonship.example', true)]),
    mk('demo:6', 'salesorders', 'SO-0095', 'Texas Lifting Solutions', d(15), 9750, 'open', 'sales@texaslifting.example',
      [['Load cell 55 t with display', 1, 7900], ['Shackles 25 t (pair)', 2, 925]], '',
      [person('Katie Brunner', 'Sales Coordinator', 'katie@texaslifting.example', true)]),
    mk('demo:7', 'invoices', 'INV-1066', 'Acadia Offshore', d(20), 15200, 'sent', 'maintenance@acadiaoffshore.example',
      [['Pedestal crane recertification', 1, 11800], ['Function test at 100% SWL', 1, 3400]], 'ACA-PO-1190',
      [person('Joe Fontenot', 'Maintenance Lead', 'joe@acadiaoffshore.example', true),
       person('Sandra Mills', 'HSE Advisor', 'sandra@acadiaoffshore.example')]),
    mk('demo:8', 'estimates', 'EST-00220', 'Pelican Energy Partners', d(24), 21000, 'draft', 'logistics@pelicanenergy.example',
      [['Water bag rental — 12 × 20 t, 3 weeks', 12, 1500], ['Spreader bar hire', 1, 3000]], '',
      [person('Tom Landry', 'Logistics Coordinator', 'tom@pelicanenergy.example', true)]),
    mk('demo:9', 'invoices', 'INV-1070', 'Marina Del Sol', d(31), 11600, 'paid', 'harbour@marinadelsol.example',
      [['10 t water weights — set of 4 (sale)', 1, 11600]], 'MDS-2071',
      [person('Isabel Cruz', 'Harbour Master', 'isabel@marinadelsol.example', true)]),
    mk('demo:10', 'salesorders', 'SO-0098', 'Crescent Rigging Co', d(38), 5400, 'open', 'office@crescentrigging.example',
      [['Water bag hire + dynamometer — 1 week', 1, 5400]], '',
      [person('Denise Boudreaux', 'Office Manager', 'denise@crescentrigging.example', true)])
  ];
}

function effectiveDemo(s) {
  if (s.demoMode === 'on') return true;
  if (s.demoMode === 'off') return false;
  return !s.refreshToken; // auto
}

function currentJobs(s) {
  if (effectiveDemo(s)) return demoJobs();
  const cache = readJson(FILES.cache, { jobs: {}, lastSync: null });
  return Object.values(cache.jobs || {});
}

// ---------------------------------------------------------------- HTTP server

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.webp': 'image/webp', '.gif': 'image/gif',
  '.ico': 'image/x-icon', '.json': 'application/json; charset=utf-8'
};

function send(res, code, body, type) {
  const buf = typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body);
  res.writeHead(code, {
    'Content-Type': type || 'application/json; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  res.end(buf);
}
const ok = (res, obj) => send(res, 200, obj);
const bad = (res, code, msg) => send(res, code, { error: msg });

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => { size += c.length; if (size > 2 * 1024 * 1024) { reject(new Error('Body too large')); req.destroy(); } else chunks.push(c); });
    req.on('end', () => {
      if (!chunks.length) return resolve({});
      try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf8'))); }
      catch (e) { reject(new Error('Invalid JSON body')); }
    });
    req.on('error', reject);
  });
}

// ---- procedure attachments (rigging-drawing PDFs + site photos), stored in Supabase
// Storage and kept OUT of app state so pm_app_state stays small — only a tiny reference
// (name/size/path) lives on the procedure.
const DRAW_BUCKET = 'procedure-drawings';
const PHOTO_BUCKET = 'procedure-photos';
const DRAW_MAX = 4 * 1024 * 1024;    // 4 MB per PDF — under the serverless body limit
const PHOTO_MAX = 4 * 1024 * 1024;   // photos are downscaled client-side; this is a backstop
function storageCfg() {
  return { url: (process.env.SUPABASE_URL || '').replace(/\/+$/, ''), key: process.env.SUPABASE_SERVICE_KEY || '' };
}
function storageOn() { const c = storageCfg(); return !!(c.url && c.key); }
async function storagePut(bucket, path, buf, contentType) {
  const c = storageCfg();
  const r = await fetch(c.url + '/storage/v1/object/' + bucket + '/' + path, {
    method: 'POST',
    headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Content-Type': contentType, 'x-upsert': 'true' },
    body: buf
  });
  if (!r.ok) throw new Error('Storage upload failed (' + r.status + ')');
}
async function storageSign(bucket, path, expiresIn) {
  const c = storageCfg();
  const r = await fetch(c.url + '/storage/v1/object/sign/' + bucket + '/' + path, {
    method: 'POST',
    headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Content-Type': 'application/json' },
    body: JSON.stringify({ expiresIn: expiresIn || 3600 })
  });
  if (!r.ok) throw new Error('Could not create a view link (' + r.status + ')');
  const d = await r.json();
  return c.url + '/storage/v1' + d.signedURL;
}
async function storageDel(bucket, path) {
  const c = storageCfg();
  try {
    await fetch(c.url + '/storage/v1/object/' + bucket + '/' + path, {
      method: 'DELETE', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key }
    });
  } catch (e) {}
}
// Attach fresh signed view URLs to a photo list — thumbnails need a src at render time.
async function photosWithUrls(photos) {
  const list = Array.isArray(photos) ? photos : [];
  return Promise.all(list.map(async p => {
    try { return Object.assign({}, p, { url: await storageSign(PHOTO_BUCKET, p.path, 3600) }); }
    catch (e) { return Object.assign({}, p, { url: '' }); }
  }));
}
function readRawBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => { size += c.length; if (size > maxBytes) { reject(Object.assign(new Error('Body too large'), { code: 'TOO_LARGE' })); req.destroy(); } else chunks.push(c); });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function publicSettings(s) {
  const out = Object.assign({}, s);
  delete out.refreshToken;
  out.connected = !!s.refreshToken;
  out.redirectUri = redirectUri();
  out.ms = Object.assign({}, s.ms);
  delete out.ms.refreshToken;
  out.msConnected = !!(s.ms && s.ms.refreshToken);
  out.msRedirectUri = msRedirectUri();
  out.shopmasterConnected = !!(s.shopmaster && s.shopmaster.url && s.shopmaster.key);
  return out;
}

const SETTINGS_EDITABLE = ['dc', 'clientId', 'clientSecret', 'senderName', 'modules', 'rules',
  'defaultCategory', 'maxPages', 'demoMode', 'orgId', 'orgName', 'smtp', 'ms', 'shopmaster'];

// ---------------------------------------------------------------- auth (Supabase)
// Login reuses the shared Supabase project's Auth — the SAME accounts as the travel
// app. The browser signs in with supabase-js and sends the user's access token on
// every /api request; we verify it here against Supabase and restrict to staff by
// work-email domain. Enforced only when REQUIRE_AUTH is on, so local dev keeps
// working until you flip it (it is set ON in the Vercel deployment).
function authRequired() { return /^(1|true|yes|on)$/i.test(String(process.env.REQUIRE_AUTH || '')); }
function supabaseAnonKey() { return process.env.SUPABASE_ANON_KEY || (getSettings().shopmaster || {}).key || ''; }
function allowedDomain() { return String(process.env.ALLOWED_EMAIL_DOMAIN || 'hydrowates.com').trim().toLowerCase(); }
function bearerToken(req) { const a = req.headers.authorization || ''; return a.startsWith('Bearer ') ? a.slice(7).trim() : ''; }

const authCache = new Map();   // access token -> { user, exp }
async function verifyUser(token) {
  if (!token) return null;
  const now = Date.now();
  const hit = authCache.get(token);
  if (hit && hit.exp > now) return hit.user;
  const url = process.env.SUPABASE_URL, anon = supabaseAnonKey();
  if (!url || !anon) return null;
  let user = null;
  try {
    const r = await fetch(url.replace(/\/+$/, '') + '/auth/v1/user', {
      headers: { apikey: anon, Authorization: 'Bearer ' + token }
    });
    if (r.ok) user = await r.json();
  } catch (e) { return null; }
  if (!user || !user.id) return null;
  const email = String(user.email || '').toLowerCase();
  const dom = allowedDomain();
  if (dom && !email.endsWith('@' + dom)) return null;   // not internal staff
  if (authCache.size > 500) authCache.clear();
  authCache.set(token, { user, exp: now + 5 * 60 * 1000 });
  return user;
}

// ---- Shop Master loadout -> procedure equipment -------------------------------
// Clean the WLL unit — Shop Master syncs it from SharePoint, so it can arrive as a
// nested "expanded reference" blob; pull out the readable unit and normalize it.
function cleanWllUnit(u) {
  const s = String(u == null ? '' : u);
  if (/pound|\blbs?\b/i.test(s)) return 'lb';
  if (/metric ton|tonne|\bte\b/i.test(s)) return 't';
  const clean = s.replace(/[{}"@]/g, '').trim();
  return (clean && clean.length <= 12 && !/odata|reference|sharepoint|value|id\s*:/i.test(clean)) ? clean : '';
}
function fmtEquipLine(e) {
  const wll = (e.wll != null && e.wll !== '') ? ' (WLL ' + Number(e.wll).toLocaleString() + (e.wllUnit ? ' ' + e.wllUnit : '') + ')' : '';
  const ser = e.serials.length ? ' — ' + e.serials.join(', ') : '';
  return e.qty + '× ' + e.name + wll + ser;
}
// Read a job's ACTUAL equipment from Shop Master's loadout, matched by HWI:
// loadout header -> line items -> inventory record (name + WLL). Aggregated per
// item, with serials collected. Read-only.
async function shopmasterLoadoutEquipment(s, hwi) {
  const H = String(hwi || '').trim();
  if (!H) return { found: false };
  const los = await shopmasterGet(s, 'shopmaster_loadouts?select=id,mobilization_number,job_number,phase,approved_at,shipped_at,created_at&job_number=eq.' + encodeURIComponent(H) + '&order=created_at.desc');
  if (!Array.isArray(los) || !los.length) return { found: false };
  const byItem = new Map();
  let pieces = 0;
  for (const lo of los) {
    let items;
    try { items = await shopmasterGet(s, 'shopmaster_loadout_items?select=quantity,assembly_serial_number,shopmaster_inventory_items(description,working_load_limit,wll_unit)&loadout_id=eq.' + lo.id); }
    catch (e) { continue; }
    if (!Array.isArray(items)) continue;
    for (const it of items) {
      const inv = it.shopmaster_inventory_items || {};
      const name = String(inv.description || '').trim() || 'Unnamed item';
      const wll = (inv.working_load_limit != null && inv.working_load_limit !== '') ? inv.working_load_limit : null;
      const wllUnit = cleanWllUnit(inv.wll_unit);
      const qty = Number(it.quantity) || 1;
      const key = name.toLowerCase();                       // merge case variants of the same item
      let e = byItem.get(key);
      if (!e) { e = { name, wll, wllUnit, qty: 0, serials: [] }; byItem.set(key, e); }
      e.qty += qty; pieces += qty;
      if (e.wll == null && wll != null) { e.wll = wll; e.wllUnit = wllUnit; }   // adopt a WLL if this row has one
      const sn = String(it.assembly_serial_number || '').trim();
      // real serials only — bulk items store their own name in that field
      if (sn && sn !== '-' && sn.toLowerCase() !== name.toLowerCase() && !e.serials.includes(sn)) e.serials.push(sn);
    }
  }
  const items = [...byItem.values()].sort((a, b) => (b.qty - a.qty) || a.name.localeCompare(b.name));
  const top = los[0];
  return {
    found: items.length > 0,
    source: { hwi: H, loadoutRef: top.mobilization_number || top.job_number || H, phase: top.phase || '', at: top.approved_at || top.shipped_at || top.created_at || '', pieces, shipments: los.length },
    items,
    lines: items.map(fmtEquipLine)
  };
}

async function handleApi(req, res, u) {
  const p = u.pathname;
  const method = req.method;

  // ---- storage backend diagnostic (no secrets) — confirms file vs Supabase
  if (p === '/api/storage' && method === 'GET') {
    return ok(res, { mode: store.mode() });
  }

  // ---- auth config for the browser (public): how to reach Supabase + whether login is on
  if (p === '/api/auth/config' && method === 'GET') {
    return ok(res, { url: process.env.SUPABASE_URL || '', anonKey: supabaseAnonKey(), required: authRequired() });
  }

  // ---- jobs list
  if (p === '/api/jobs' && method === 'GET') {
    const s = getSettings();
    const local = readJson(FILES.jobs, {});
    const cache = readJson(FILES.cache, { jobs: {}, lastSync: null });
    const demo = effectiveDemo(s);
    // Dashboard sources: the SharePoint Lead List + Zoho Books ONLY. Shop Master is
    // not used here (it's the Invoices page's received-jobs feed) — the Lead List and
    // Zoho carry everything the dashboard needs.
    const allBlocks = currentJobs(s).concat(openLeadJobs(s));
    const idx = leadIndex(cache);
    const invoicedHwis = invoicedHwisFrom(allBlocks);
    const mapped = allBlocks.map(b => {
      const r = resolveJob(b, local, s, idx);
      const l = local[b.key] || {};
      const done = completedState(b, l, r.hwi, invoicedHwis, idx);
      return {
        key: b.key, module: b.module, number: b.number, customer: b.customer, date: b.date,
        total: b.total, currency: b.currency, status: b.status, reference: b.reference,
        hwi: r.hwi,
        lead: !!b.lead, leadTitle: b.leadTitle || '',
        received: !!b.received, projected: (b.projectedValue != null ? b.projectedValue : null),
        category: r.category, categoryOverridden: r.categoryOverridden,
        stage: l.stage || 'new', hidden: !!l.hidden,
        planningStatus: (l.planning && l.planning.status) || 'none',
        preHeld: !!(l.meetings && l.meetings.pre && l.meetings.pre.held),
        postHeld: !!(l.meetings && l.meetings.post && l.meetings.post.held),
        multiInvoice: !!l.multiInvoice,
        archived: done.archived, archivedHow: done.how
      };
    });
    // "Made a job" date: the EARLIEST date across every block that shares this
    // HWI (lead quote, estimate, sales order …). A job often surfaces first as a
    // Lead List quote and only later as Zoho documents, so we want when it first
    // entered the pipeline — not the latest document date on the winning card.
    const earliestByHwi = new Map();
    for (const j of mapped) {
      const k = hwiKey(j.hwi);
      if (k.length < 4 || !j.date) continue;
      const cur = earliestByHwi.get(k);
      if (!cur || j.date < cur) earliestByHwi.set(k, j.date);
    }
    for (const j of mapped) {
      const k = hwiKey(j.hwi);
      j.createdDate = ((k.length >= 4 && earliestByHwi.get(k)) || j.date || '');
    }
    // The active board = jobs that carry a PO (the customer has committed) AND have
    // NOT been invoiced (an invoice = the job is complete, paid or not). Anything
    // invoiced, or with no PO on record, drops off entirely — the Invoices page
    // keeps the invoiced history.
    const activeMapped = mapped.filter(j => !j.archived && String(j.reference || '').trim() !== '');
    // One card per job (HWI). The same job can surface as a lead, an estimate, a
    // sales order, and an invoice — keep the most useful single card. Lead cards
    // win (they carry the Lead List category + value); cards with no HWI can't be
    // deduped so they're all kept.
    const PRIO = { lead: 5, received: 4, estimates: 3, salesorders: 2, invoices: 1 };
    const bestAt = new Map();   // hwiKey -> index in jobs[]
    const jobs = [];
    for (const j of activeMapped) {
      const k = hwiKey(j.hwi);
      if (k.length < 4) { jobs.push(j); continue; }
      if (!bestAt.has(k)) { bestAt.set(k, jobs.length); jobs.push(j); continue; }
      const i = bestAt.get(k);
      if ((PRIO[j.module] || 0) > (PRIO[jobs[i].module] || 0)) jobs[i] = j;   // replace with the better card
    }
    return ok(res, {
      jobs, demo, connected: !!s.refreshToken, orgName: s.orgName || '',
      lastSync: demo ? null : cache.lastSync, sync: syncState
    });
  }

  // ---- Recently deleted: jobs manually removed from the board (archiveOverride='archived').
  if (p === '/api/removed' && method === 'GET') {
    const s = getSettings();
    const local = readJson(FILES.jobs, {});
    const keys = Object.keys(local).filter(k => local[k] && local[k].archiveOverride === 'archived' && !local[k].purged);
    if (!keys.length) return ok(res, { removed: [] });
    const byKey = new Map(currentJobs(s).concat(openLeadJobs(s)).map(b => [b.key, b]));
    const idx = leadIndex(readJson(FILES.cache, {}));
    const removed = keys.map(k => {
      const b = byKey.get(k); const l = local[k] || {};
      if (!b) return { key: k, hwi: '', customer: '(no longer in the sync window)', po: '', total: null, currency: '', category: '', removedAt: l.removedAt || null };
      const r = resolveJob(b, local, s, idx);
      return { key: k, hwi: r.hwi, customer: b.customer || '', po: b.reference || '', total: b.total, currency: b.currency, category: r.category, removedAt: l.removedAt || null };
    }).sort((a, bb) => String(bb.removedAt || '').localeCompare(String(a.removedAt || '')));
    return ok(res, { removed });
  }

  // ---- Bulk permanently-delete from Recently deleted. Body: { keys: [...] }.
  if (p === '/api/removed/purge' && method === 'POST') {
    const body = await readBody(req);
    const local = readJson(FILES.jobs, {});
    const keys = Array.isArray(body.keys) ? body.keys : [];
    let n = 0;
    for (const k of keys) {
      const l = local[k];
      if (l && l.archiveOverride === 'archived' && !l.purged) {
        l.purged = true;
        if (!l.removedAt) l.removedAt = new Date().toISOString();
        n++;
      }
    }
    if (n) writeJson(FILES.jobs, local);
    return ok(res, { purged: n });
  }

  // ---- Post-job meetings still owed: jobs whose pre-job meeting was held but the
  // post-job meeting hasn't been (works even after the job is invoiced / off the board).
  if (p === '/api/meetings/todo' && method === 'GET') {
    const local = readJson(FILES.jobs, {});
    const todo = Object.keys(local).filter(k => {
      const m = local[k] && local[k].meetings;
      return m && m.pre && m.pre.held && !(m.post && m.post.held);
    }).map(k => {
      const m = local[k].meetings; const jb = m.job || {};
      return { key: k, hwi: jb.hwi || '', customer: jb.customer || '', po: jb.po || '', preDate: (m.pre && m.pre.date) || '' };
    }).sort((a, b) => String(b.preDate || '').localeCompare(String(a.preDate || '')));
    return ok(res, { todo });
  }

  // ---- single job + local edits
  if (p.startsWith('/api/job/')) {
    const rest = decodeURIComponent(p.slice('/api/job/'.length));
    const m2 = rest.match(/^(.*)\/(planning|procedure|contacts|send|loadout-equipment|meetings)$/);
    const key = m2 ? m2[1] : rest;
    const sub = m2 ? m2[2] : '';
    const s = getSettings();
    const local = readJson(FILES.jobs, {});

    // ---- rigging-drawing PDF attachments on a job's procedure ----
    const mDraw = rest.match(/^(.*)\/drawings(?:\/([^/]+))?$/);
    if (mDraw) {
      const dkey = mDraw[1];
      const drawId = mDraw[2] ? decodeURIComponent(mDraw[2]) : '';
      const procOf = () => (local[dkey] && local[dkey].procedure) || null;
      const persist = (drawings) => {
        const procedure = Object.assign({}, procOf() || {}, { drawings, updatedAt: new Date().toISOString() });
        local[dkey] = Object.assign({}, local[dkey], { procedure });
        writeJson(FILES.jobs, local);
      };
      if (!storageOn()) return bad(res, 501, 'File storage is not configured on this server.');

      if (!drawId && method === 'POST') {                 // upload a PDF
        let buf;
        try { buf = await readRawBody(req, DRAW_MAX + 8192); }
        catch (e) { return bad(res, 413, 'That PDF is over the 4 MB limit — please attach a smaller file.'); }
        if (buf.length > DRAW_MAX) return bad(res, 413, 'That PDF is over the 4 MB limit — please attach a smaller file.');
        if (buf.length < 5 || buf.slice(0, 5).toString('latin1') !== '%PDF-') return bad(res, 400, 'That file is not a PDF.');
        const name = String(req.headers['x-filename'] || 'drawing.pdf').replace(/[^A-Za-z0-9._ -]/g, '_').replace(/\.pdf$/i, '').slice(0, 100) + '.pdf';
        const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
        const path = dkey.replace(/[^A-Za-z0-9._-]/g, '_') + '/' + id + '-' + name;
        try { await storagePut(DRAW_BUCKET, path, buf, 'application/pdf'); }
        catch (e) { return bad(res, 502, String(e.message || e)); }
        const prev = procOf() || {};
        const drawings = (Array.isArray(prev.drawings) ? prev.drawings : []).concat([{ id, name, size: buf.length, path, at: new Date().toISOString() }]);
        persist(drawings);
        return ok(res, { drawings });
      }
      if (drawId && method === 'GET') {                   // signed view link
        const d = (procOf() || {}).drawings && (procOf().drawings || []).find(x => x.id === drawId);
        if (!d) return bad(res, 404, 'Drawing not found');
        try { return ok(res, { url: await storageSign(DRAW_BUCKET, d.path, 3600), name: d.name }); }
        catch (e) { return bad(res, 502, String(e.message || e)); }
      }
      if (drawId && method === 'DELETE') {                // remove a drawing
        const prev = procOf();
        if (!prev) return bad(res, 404, 'No procedure for this job');
        const d = (prev.drawings || []).find(x => x.id === drawId);
        if (d) await storageDel(DRAW_BUCKET, d.path);
        persist((prev.drawings || []).filter(x => x.id !== drawId));
        return ok(res, { drawings: (prev.drawings || []).filter(x => x.id !== drawId) });
      }
      return bad(res, 405, 'Method not allowed');
    }

    // ---- site-photo (image) attachments on a job's procedure ----
    const mPhoto = rest.match(/^(.*)\/photos(?:\/([^/]+))?$/);
    if (mPhoto) {
      const pkey = mPhoto[1];
      const photoId = mPhoto[2] ? decodeURIComponent(mPhoto[2]) : '';
      const procOf = () => (local[pkey] && local[pkey].procedure) || null;
      const persist = (photos) => {
        const procedure = Object.assign({}, procOf() || {}, { photos, updatedAt: new Date().toISOString() });
        local[pkey] = Object.assign({}, local[pkey], { procedure });
        writeJson(FILES.jobs, local);
      };
      if (!storageOn()) return bad(res, 501, 'File storage is not configured on this server.');

      if (!photoId && method === 'POST') {                // upload an image (downscaled to JPEG client-side)
        let buf;
        try { buf = await readRawBody(req, PHOTO_MAX + 8192); }
        catch (e) { return bad(res, 413, 'That photo is too large.'); }
        if (buf.length > PHOTO_MAX) return bad(res, 413, 'That photo is too large.');
        const isJpg = buf.length > 3 && buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF;
        const isPng = buf.length > 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47;
        if (!isJpg && !isPng) return bad(res, 400, 'That file is not a JPEG or PNG image.');
        const ext = isPng ? 'png' : 'jpg';
        const rawName = String(req.headers['x-filename'] || 'photo').replace(/[^A-Za-z0-9._ -]/g, '_').replace(/\.[A-Za-z0-9]+$/, '').slice(0, 80);
        const name = (rawName || 'photo') + '.' + ext;
        const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
        const path = pkey.replace(/[^A-Za-z0-9._-]/g, '_') + '/' + id + '.' + ext;
        try { await storagePut(PHOTO_BUCKET, path, buf, isPng ? 'image/png' : 'image/jpeg'); }
        catch (e) { return bad(res, 502, String(e.message || e)); }
        const prev = procOf() || {};
        const photos = (Array.isArray(prev.photos) ? prev.photos : []).concat([{ id, name, size: buf.length, path, at: new Date().toISOString() }]);
        persist(photos);
        return ok(res, { photos: await photosWithUrls(photos) });
      }
      if (photoId && method === 'GET') {                  // signed full-size link
        const ph = (procOf() || {}).photos && (procOf().photos || []).find(x => x.id === photoId);
        if (!ph) return bad(res, 404, 'Photo not found');
        try { return ok(res, { url: await storageSign(PHOTO_BUCKET, ph.path, 3600), name: ph.name }); }
        catch (e) { return bad(res, 502, String(e.message || e)); }
      }
      if (photoId && method === 'DELETE') {
        const prev = procOf();
        if (!prev) return bad(res, 404, 'No procedure for this job');
        const ph = (prev.photos || []).find(x => x.id === photoId);
        if (ph) await storageDel(PHOTO_BUCKET, ph.path);
        const photos = (prev.photos || []).filter(x => x.id !== photoId);
        persist(photos);
        return ok(res, { photos: await photosWithUrls(photos) });
      }
      return bad(res, 405, 'Method not allowed');
    }

    if (sub === 'contacts' && method === 'GET') {
      if (key.startsWith('received:')) return ok(res, { contacts: [] }); // Shop Master jobs carry no Zoho contact link
      const b = currentJobs(s).find(j => j.key === key) ||
        (key.startsWith('lead:') ? openLeadJobs(s).find(j => j.key === key) : null);
      if (!b) return bad(res, 404, 'Job not found');
      if (b.lead) return ok(res, { contacts: [] }); // SharePoint leads carry no Zoho contact link
      if (b.demo) return ok(res, { contacts: b.contacts || [] });
      if (!b.customerId) return bad(res, 409, 'This job needs a fresh sync before its contacts can be looked up — press Sync, then reopen it.');
      try {
        return ok(res, { contacts: await contactsForCustomer(b.customerId) });
      } catch (e) { return bad(res, 502, String(e.message || e)); }
    }

    // The job's real equipment from its Shop Master loadout (by HWI) — used to
    // auto-fill the procedure's Equipment section.
    if (sub === 'loadout-equipment' && method === 'GET') {
      const b = currentJobs(s).find(j => j.key === key) ||
        (key.startsWith('lead:') ? openLeadJobs(s).find(j => j.key === key) : null);
      if (!b) return bad(res, 404, 'Job not found');
      const hwi = extractHwi(b);
      if (!hwi) return ok(res, { found: false, reason: 'no-hwi' });
      try { return ok(res, await shopmasterLoadoutEquipment(s, hwi)); }
      catch (e) { return bad(res, e.code === 'NOT_CONNECTED' ? 409 : 502, String(e.message || e)); }
    }

    if (sub === 'send' && method === 'POST') {
      const body = await readBody(req);
      const b = currentJobs(s).find(j => j.key === key) ||
        (key.startsWith('lead:') ? openLeadJobs(s).find(j => j.key === key) : null);
      if (!b) return bad(res, 404, 'Job not found');
      if (b.demo) return bad(res, 400, 'This is a demo job, so nothing was sent. Connect Zoho Books to send for real — or use “Open email draft” to see how it looks.');
      const to = Array.isArray(body.to)
        ? [...new Set(body.to.map(x => String(x).trim()).filter(x => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x)))]
        : [];
      if (!to.length) return bad(res, 400, 'No valid recipient email address — tick a contact or type one in.');
      const kind = body.kind === 'procedure' ? 'procedure' : body.kind === 'meeting-report' ? 'meeting-report' : 'planning';
      const subject = String(body.subject || '').slice(0, 500) || (kind === 'procedure' ? 'Load Test Procedure' : kind === 'meeting-report' ? 'Meeting report' : 'Planning questions');
      const text = String(body.body || '');
      const html = body.html ? String(body.html) : '';
      if (!text.trim() && !html.trim()) return bad(res, 400, 'The email body is empty.');
      try {
        await sendMail(s, { to, subject, body: text, html: html || undefined }, body.msToken);
      } catch (e) { return bad(res, 502, String(e.message || e)); }
      const now = new Date().toISOString();
      if (kind === 'meeting-report') {
        const prev = (local[key] && local[key].meetings) || {};
        const meetings = Object.assign({}, prev);
        meetings.reportLog = (Array.isArray(prev.reportLog) ? prev.reportLog : []).concat([{ at: now, to, subject }]).slice(-20);
        local[key] = Object.assign({}, local[key], { meetings });
        writeJson(FILES.jobs, local);
        return ok(res, { sent: true, to, meetings });
      }
      if (kind === 'procedure') {
        const prev = (local[key] && local[key].procedure) || {};
        const procedure = Object.assign({}, prev);
        procedure.status = 'final';
        procedure.recipients = to;
        if (!procedure.sentAt) procedure.sentAt = now;
        procedure.sentLog = (prev.sentLog || []).concat([{ at: now, to, subject }]);
        procedure.updatedAt = now;
        local[key] = Object.assign({}, local[key], { procedure });
        writeJson(FILES.jobs, local);
        return ok(res, { sent: true, to, procedure });
      }
      const prev = (local[key] && local[key].planning) || {};
      const planning = Object.assign({ questions: [], notes: '', email: '' }, prev);
      planning.status = 'sent';
      if (!planning.sentAt) planning.sentAt = now;
      planning.recipients = to;
      planning.sentLog = (prev.sentLog || []).concat([{ at: now, to, subject }]);
      planning.updatedAt = now;
      local[key] = Object.assign({}, local[key], { planning });
      writeJson(FILES.jobs, local);
      return ok(res, { sent: true, to, planning });
    }

    if (sub === 'planning' && method === 'PUT') {
      const body = await readBody(req);
      const prev = (local[key] && local[key].planning) || {};
      const planning = {
        questions: Array.isArray(body.questions) ? body.questions.map(normJobQ) : (prev.questions || []),
        status: ['none', 'prepared', 'sent', 'answered'].includes(body.status) ? body.status : (prev.status || 'prepared'),
        notes: body.notes !== undefined ? String(body.notes) : (prev.notes || ''),
        email: body.email !== undefined ? String(body.email) : (prev.email || ''),
        recipients: Array.isArray(body.recipients) ? body.recipients.map(x => String(x).trim()).filter(Boolean) : (prev.recipients || []),
        sentLog: prev.sentLog || [],
        sentAt: prev.sentAt || null,
        updatedAt: new Date().toISOString()
      };
      if (planning.status === 'sent' && !planning.sentAt) planning.sentAt = new Date().toISOString();
      local[key] = Object.assign({}, local[key], { planning });
      writeJson(FILES.jobs, local);
      return ok(res, { saved: true, planning });
    }

    if (sub === 'procedure' && method === 'PUT') {
      const body = await readBody(req);
      const prev = (local[key] && local[key].procedure) || {};
      const arr = (v, fb) => Array.isArray(v) ? v.map(x => String(x)).slice(0, 200) : (fb || []);
      const str = (v, fb) => v !== undefined ? String(v).slice(0, 20000) : (fb || '');
      const procedure = {
        status: ['draft', 'reviewed', 'final'].includes(body.status) ? body.status : (prev.status || 'draft'),
        startDate: str(body.startDate, prev.startDate),
        scope: str(body.scope, prev.scope),
        jobSite: str(body.jobSite, prev.jobSite),
        projectRef: str(body.projectRef, prev.projectRef),
        objective: str(body.objective, prev.objective),
        coordination: str(body.coordination, prev.coordination),
        responsibilities: arr(body.responsibilities, prev.responsibilities),
        equipment: arr(body.equipment, prev.equipment),
        equipmentSource: (body.equipmentSource !== undefined ? body.equipmentSource : prev.equipmentSource) || null,
        ppe: arr(body.ppe, prev.ppe),
        drawings: Array.isArray(prev.drawings) ? prev.drawings : [],   // managed via the /drawings endpoints, never clobbered by a form save
        photos: Array.isArray(prev.photos) ? prev.photos.map(p => { const q = Object.assign({}, p); delete q.url; return q; }) : [],   // same for /photos; strip any signed url before persisting
        setup: (body.setup && typeof body.setup === 'object' && !Array.isArray(body.setup)) ? body.setup : (prev.setup || null),   // the "Procedure setup" answers that drove generation
        preJob: str(body.preJob, prev.preJob),
        setupSteps: arr(body.setupSteps, prev.setupSteps),
        executionSteps: arr(body.executionSteps, prev.executionSteps),
        logging: str(body.logging, prev.logging),
        approvedBy: str(body.approvedBy, prev.approvedBy),
        approvedDate: str(body.approvedDate, prev.approvedDate),
        recipients: Array.isArray(body.recipients) ? body.recipients.map(x => String(x).trim()).filter(Boolean) : (prev.recipients || []),
        email: str(body.email, prev.email),
        sentLog: Array.isArray(prev.sentLog) ? prev.sentLog : [],
        sentAt: prev.sentAt || null,
        updatedAt: new Date().toISOString()
      };
      local[key] = Object.assign({}, local[key], { procedure });
      writeJson(FILES.jobs, local);
      return ok(res, { saved: true, procedure });
    }

    if (sub === 'meetings' && method === 'PUT') {
      const body = await readBody(req);
      const prev = (local[key] && local[key].meetings) || {};
      const blk = (v, p) => {
        v = v || {}; p = p || {};
        const actions = Array.isArray(v.actions)
          ? v.actions.slice(0, 200).map(a => ({
              text: String((a && a.text) || '').slice(0, 500),
              assignee: String((a && a.assignee) || '').slice(0, 120),
              due: String((a && a.due) || '').slice(0, 20),
              done: !!(a && a.done)
            }))
          : (Array.isArray(p.actions) ? p.actions : []);
        return {
          held: v.held !== undefined ? !!v.held : !!p.held,
          date: v.date !== undefined ? String(v.date).slice(0, 20) : (p.date || ''),
          notes: v.notes !== undefined ? String(v.notes).slice(0, 8000) : (p.notes || ''),
          actions: actions
        };
      };
      const jb = body.job || {};
      const meetings = {
        pre: blk(body.pre, prev.pre),
        post: blk(body.post, prev.post),
        notes: body.notes !== undefined ? String(body.notes).slice(0, 8000) : (prev.notes || ''),
        // a light snapshot so the "post-job to-do" list can show the job even after it leaves the board
        job: (jb.hwi || jb.customer || jb.po)
          ? { hwi: String(jb.hwi || ''), customer: String(jb.customer || ''), po: String(jb.po || '') }
          : (prev.job || {}),
        updatedAt: new Date().toISOString()
      };
      local[key] = Object.assign({}, local[key], { meetings });
      writeJson(FILES.jobs, local);
      return ok(res, { saved: true, meetings });
    }

    if (!sub && method === 'PATCH') {
      const body = await readBody(req);
      const l = Object.assign({}, local[key]);
      if (body.stage !== undefined) l.stage = String(body.stage);
      if (body.hidden !== undefined) l.hidden = !!body.hidden;
      if (body.categoryOverride !== undefined) {
        if (body.categoryOverride) l.categoryOverride = String(body.categoryOverride);
        else delete l.categoryOverride;
      }
      if (body.wll !== undefined) {
        const n = parseFloat(body.wll);
        if (body.wll === '' || body.wll === null || isNaN(n) || n < 0) delete l.wll;
        else l.wll = n;
      }
      if (body.wllUnit !== undefined) l.wllUnit = String(body.wllUnit).slice(0, 16);
      if (body.multiInvoice !== undefined) { if (body.multiInvoice) l.multiInvoice = true; else delete l.multiInvoice; }
      if (body.archiveOverride !== undefined) {
        if (body.archiveOverride === 'archived' || body.archiveOverride === 'active') l.archiveOverride = body.archiveOverride;
        else { delete l.archiveOverride; delete l.purged; }   // clearing = restore to the board
        if (body.archiveOverride === 'archived') l.removedAt = new Date().toISOString();   // stamp for the Recently deleted list
        else delete l.removedAt;
      }
      // Permanently drop a removed job out of "Recently deleted": stays off the board
      // (archived) but no longer restorable from the list. The underlying SharePoint/
      // Zoho record can't be deleted from here — this just hides it everywhere.
      if (body.purge === true) {
        l.archiveOverride = 'archived';
        l.purged = true;
        if (!l.removedAt) l.removedAt = new Date().toISOString();
      }
      local[key] = l;
      writeJson(FILES.jobs, local);
      return ok(res, { saved: true });
    }

    if (!sub && method === 'GET') {
      const b = currentJobs(s).find(j => j.key === key) ||
        (key.startsWith('lead:') ? openLeadJobs(s).find(j => j.key === key) : null) ||
        (key.startsWith('received:') ? (await receivedDashJobs(s)).find(j => j.key === key) : null);
      if (!b) return bad(res, 404, 'Job not found (it may have dropped out of the sync window)');
      const l = local[key] || {};
      const idx = leadIndex(readJson(FILES.cache, {}));
      const r = resolveJob(b, local, s, idx);
      const hwi = r.hwi;
      const done = completedState(b, l, hwi, invoicedHwisFrom(currentJobs(s)), idx);
      const travelModes = readJson(FILES.travelModes, {});
      return ok(res, {
        job: {
          key: b.key, module: b.module, id: b.id, number: b.number, customer: b.customer,
          date: b.date, total: b.total, currency: b.currency, status: b.status,
          reference: b.reference, hwi, email: b.email || '', lineItems: b.lineItems || [],
          lead: !!b.lead, leadTitle: b.leadTitle || '', received: !!b.received,
          phase: b.phase || '', invoiced: !!b.invoiced, projected: (b.projectedValue != null ? b.projectedValue : null),
          demo: !!b.demo, category: r.category, categoryOverridden: r.categoryOverridden,
          matchedRule: r.categoryRule, stage: l.stage || 'new', hidden: !!l.hidden,
          wll: (l.wll !== undefined ? l.wll : null), wllUnit: l.wllUnit || 't',
          travelMode: (hwi ? (travelModes[hwiKey(hwi)] || null) : null),
          multiInvoice: !!l.multiInvoice,
          archived: done.archived, archivedHow: done.how
        },
        planning: l.planning || null,
        procedure: l.procedure ? Object.assign({}, l.procedure, { photos: await photosWithUrls(l.procedure.photos) }) : null,
        meetings: l.meetings || null,
        zoho: { dc: s.dc, orgId: s.orgId }
      });
    }
    return bad(res, 405, 'Method not allowed');
  }

  // ---- templates
  if (p === '/api/templates' && method === 'GET') return ok(res, getTemplates());
  if (p === '/api/templates' && method === 'PUT') {
    const body = await readBody(req);
    const t = getTemplates();
    if (Array.isArray(body.questions)) {
      t.questions = body.questions
        .map(q => Object.assign(
          { id: String(q.id || ('q' + Math.random().toString(36).slice(2, 8))), text: String(q.text || '').trim() },
          normQShape(q)
        ))
        .filter(q => q.text);
    }
    for (const k of ['emailSubject', 'emailIntro', 'emailOutro']) {
      if (body[k] !== undefined) t[k] = String(body[k]);
    }
    if (Array.isArray(body.team)) {
      t.team = body.team
        .map(m => ({ name: String(m.name || '').trim().slice(0, 120), contact: String(m.contact || '').trim().slice(0, 120), email: String(m.email || '').trim().slice(0, 160) }))
        .filter(m => m.name || m.contact || m.email)
        .slice(0, 200);
    }
    if (Array.isArray(body.equipment)) {
      t.equipment = body.equipment.map(x => String(x).trim()).filter(Boolean).slice(0, 200);
    }
    writeJson(FILES.templates, t);
    return ok(res, t);
  }

  // ---- settings
  if (p === '/api/settings' && method === 'GET') return ok(res, publicSettings(getSettings()));
  if (p === '/api/settings' && method === 'PUT') {
    const body = await readBody(req);
    const cur = readJson(FILES.settings, {});
    const prevMsToken = (cur.ms && cur.ms.refreshToken) || '';
    for (const k of SETTINGS_EDITABLE) {
      if (body[k] !== undefined) cur[k] = body[k];
    }
    if (cur.ms && typeof cur.ms === 'object') {
      const m = cur.ms, mm = m.map || {};
      cur.ms = {
        tenant: String(m.tenant || 'organizations').trim() || 'organizations',
        clientId: String(m.clientId || '').trim(),
        siteId: String(m.siteId || ''), siteName: String(m.siteName || ''),
        listId: String(m.listId || ''), listName: String(m.listName || ''),
        map: {
          po: String(mm.po || ''),
          poMode: ['nonempty', 'yes', 'equals'].includes(mm.poMode) ? mm.poMode : 'nonempty',
          poValue: String(mm.poValue || ''),
          company: String(mm.company || ''),
          value: String(mm.value || ''),
          date: String(mm.date || '')
        },
        refreshToken: prevMsToken
      };
    }
    if (Array.isArray(cur.rules)) {
      cur.rules = cur.rules
        .map(r => ({ keyword: String(r.keyword || '').trim(), category: ['rental', 'service', 'sales'].includes(r.category) ? r.category : 'sales' }))
        .filter(r => r.keyword);
    }
    if (cur.smtp && typeof cur.smtp === 'object') {
      cur.smtp = {
        host: String(cur.smtp.host || '').trim(),
        port: Math.min(Math.max(Number(cur.smtp.port) || 587, 1), 65535),
        user: String(cur.smtp.user || '').trim(),
        pass: String(cur.smtp.pass || ''),
        fromName: String(cur.smtp.fromName || ''),
        fromAddr: String(cur.smtp.fromAddr || '').trim()
      };
    }
    if (cur.shopmaster && typeof cur.shopmaster === 'object') {
      cur.shopmaster = {
        url: String(cur.shopmaster.url || '').trim().replace(/\/+$/, ''),
        key: String(cur.shopmaster.key || '').trim()
      };
    }
    writeJson(FILES.settings, cur);
    tokenCache = { token: null, exp: 0 };   // dc / client may have changed
    msTokenCache = { token: null, exp: 0 };
    return ok(res, publicSettings(getSettings()));
  }

  // ---- Shop Master (Supabase) — received jobs for the Invoice tracker
  if (p === '/api/shopmaster/test' && method === 'GET') {
    try {
      const rows = await shopmasterGet(getSettings(), 'shopmaster_loadouts?select=job_number&limit=1');
      return ok(res, { ok: true, reachable: true, sample: rows.length });
    } catch (e) { return bad(res, e.code === 'NOT_CONNECTED' ? 409 : 502, String(e.message || e)); }
  }
  if (p === '/api/shopmaster/jobs' && method === 'GET') {
    try {
      const jobs = await shopmasterReceivedJobs(getSettings());
      return ok(res, { jobs, count: jobs.length });
    } catch (e) { return bad(res, e.code === 'NOT_CONNECTED' ? 409 : 502, String(e.message || e)); }
  }
  // The PM's fly/drive decision for one job. Saved locally (so the PM app always
  // shows it) AND published to the shared portal table the travel app reads, via
  // the scoped set_job_travel_mode() Supabase function. Body: { hwi, mode }.
  if (p === '/api/shopmaster/travel-mode' && method === 'POST') {
    const body = await readBody(req);
    const hwi = String(body.hwi || '').trim();
    let mode = body.mode === null ? null : String(body.mode || '').trim().toLowerCase();
    const k = hwiKey(hwi);
    if (!hwi || k.length < 4) return bad(res, 400, 'That doesn’t look like a valid HWI number.');
    if (mode !== 'fly' && mode !== 'drive' && mode !== null) return bad(res, 400, "Mode must be 'fly', 'drive', or null.");
    // 1) Save locally first — this is what the PM app displays, and it never fails.
    const store = readJson(FILES.travelModes, {});
    if (mode === null) delete store[k]; else store[k] = mode;
    writeJson(FILES.travelModes, store);
    // 2) Publish to the shared Supabase project so the travel app picks it up.
    let published = false, matched = null, publishError = null;
    try {
      const r = await shopmasterRpc(getSettings(), 'set_job_travel_mode', { p_job: hwi, p_mode: mode });
      published = true;
      matched = (r && typeof r === 'object' && 'matched' in r) ? r.matched : (typeof r === 'number' ? r : null);
    } catch (e) {
      publishError = String(e.message || e);
    }
    return ok(res, { ok: true, hwi, mode, saved: true, published, matched, publishError });
  }

  // ---- Zoho connection
  if (p === '/api/zoho/authurl' && method === 'GET') {
    const s = getSettings();
    if (!s.clientId || !s.clientSecret) return bad(res, 400, 'Enter your Zoho Client ID and Client Secret first, then press Save & Connect.');
    const q = new URLSearchParams({
      scope: 'ZohoBooks.fullaccess.all',
      client_id: s.clientId,
      response_type: 'code',
      redirect_uri: redirectUri(),
      access_type: 'offline',
      prompt: 'consent'
    });
    return ok(res, { url: accountsBase(s.dc) + '/oauth/v2/auth?' + q.toString() });
  }

  if (p === '/api/zoho/orgs' && method === 'GET') {
    try {
      const j = await zohoGet('/organizations', {}, false);
      const orgs = (j.organizations || []).map(o => ({ id: String(o.organization_id), name: o.name }));
      return ok(res, { orgs });
    } catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  if (p === '/api/zoho/disconnect' && method === 'POST') {
    const cur = readJson(FILES.settings, {});
    delete cur.refreshToken; delete cur.apiDomain;
    writeJson(FILES.settings, cur);
    tokenCache = { token: null, exp: 0 };
    return ok(res, { disconnected: true });
  }

  // ---- sync
  if (p === '/api/sync' && method === 'POST') {
    const s = getSettings();
    if (!s.refreshToken) return bad(res, 400, 'Connect to Zoho Books first (Settings).');
    if (!syncState.running) runSync(); // fire and forget
    return ok(res, { started: true });
  }
  if (p === '/api/sync/status' && method === 'GET') {
    const cache = readJson(FILES.cache, { lastSync: null });
    return ok(res, Object.assign({ lastSync: cache.lastSync }, syncState));
  }

  // ---- outgoing email test
  if (p === '/api/smtp/test' && method === 'POST') {
    const body = await readBody(req);
    const st = getSettings();
    try {
      let to = String(body.to || (st.smtp && (st.smtp.fromAddr || st.smtp.user)) || '').trim();
      if (!to && st.ms && st.ms.refreshToken) {   // Graph: send the test to the signed-in user
        const me = await graphGet('/me').catch(() => null);
        if (me) to = me.mail || me.userPrincipalName || '';
      }
      if (!to) return bad(res, 400, 'Connect Microsoft 365 (recommended) or fill in the email settings, then try the test again.');
      await sendMail(st, {
        to: [to],
        subject: 'Hydro-Wates Project Manager — test email',
        body: 'Success!\n\nEmail sending is set up correctly.\nSent ' + new Date().toLocaleString() + ' from the Project Manager app on this PC.'
      });
      return ok(res, { sent: true, to });
    } catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  // ---- PO tracker: leads from the SharePoint lead list
  if (p === '/api/leads' && method === 'GET') {
    const s = getSettings();
    const c = computeLeads(s);
    return ok(res, Object.assign({
      msConnected: !!(s.ms && s.ms.refreshToken),
      zohoConnected: !!s.refreshToken,
      configured: !!(s.ms && s.ms.siteId && s.ms.listId)
    }, c));
  }

  if (p === '/api/leads/sync' && method === 'POST') {
    try { const n = await runLeadsSync(); return ok(res, { synced: true, count: n }); }
    catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  if (p.startsWith('/api/lead/') && method === 'PATCH') {
    const id = decodeURIComponent(p.slice('/api/lead/'.length));
    const body = await readBody(req);
    const overrides = readJson(FILES.leads, {});
    const cur = Object.assign({}, overrides[id]);
    if (body.completedOverride === true || body.completedOverride === false) cur.completedOverride = body.completedOverride;
    else if (body.completedOverride === null) delete cur.completedOverride;
    overrides[id] = cur;
    writeJson(FILES.leads, overrides);
    return ok(res, { saved: true });
  }

  // ---- Microsoft 365 connection
  if (p === '/api/ms/authurl' && method === 'GET') {
    const s = getSettings();
    if (!s.ms.clientId) return bad(res, 400, 'Enter the Application (client) ID first, then press Save & Connect.');
    const verifier = crypto.randomBytes(48).toString('base64url');
    const stateTok = crypto.randomBytes(12).toString('base64url');
    msAuthPending = { verifier, state: stateTok };
    const q = new URLSearchParams({
      client_id: s.ms.clientId, response_type: 'code', redirect_uri: msRedirectUri(),
      response_mode: 'query', scope: MS_SCOPES, state: stateTok,
      code_challenge: crypto.createHash('sha256').update(verifier).digest('base64url'),
      code_challenge_method: 'S256', prompt: 'select_account'
    });
    return ok(res, { url: msLoginBase(s.ms.tenant) + '/authorize?' + q.toString() });
  }

  if (p === '/api/ms/disconnect' && method === 'POST') {
    const cur = readJson(FILES.settings, {});
    if (cur.ms) delete cur.ms.refreshToken;
    writeJson(FILES.settings, cur);
    msTokenCache = { token: null, exp: 0 };
    return ok(res, { disconnected: true });
  }

  if (p === '/api/ms/sites' && method === 'GET') {
    const q = (u.searchParams.get('q') || '').trim();
    if (!q) return bad(res, 400, 'Type part of the site name first.');
    try {
      const j = await graphGet('/sites?search=' + encodeURIComponent(q));
      return ok(res, { sites: (j.value || []).map(x => ({ id: x.id, name: x.displayName || x.name || '(unnamed site)', url: x.webUrl || '' })) });
    } catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  if (p === '/api/ms/lists' && method === 'GET') {
    const s = getSettings();
    const siteId = u.searchParams.get('siteId') || s.ms.siteId;
    if (!siteId) return bad(res, 400, 'Pick a site first.');
    try {
      const j = await graphGet('/sites/' + encodeURIComponent(siteId) + '/lists?$top=200');
      const lists = (j.value || []).filter(x => !(x.list && x.list.hidden)).map(x => ({ id: x.id, name: x.displayName || '(unnamed list)' }));
      return ok(res, { lists });
    } catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  if (p === '/api/ms/columns' && method === 'GET') {
    const s = getSettings();
    const siteId = u.searchParams.get('siteId') || s.ms.siteId;
    const listId = u.searchParams.get('listId') || s.ms.listId;
    if (!siteId || !listId) return bad(res, 400, 'Pick the site and list first.');
    try {
      const j = await graphGet('/sites/' + encodeURIComponent(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200');
      const columns = (j.value || []).filter(c => !c.hidden).map(c => ({
        name: c.name, displayName: c.displayName || c.name,
        type: c.boolean ? 'yesno' : (c.dateTime ? 'date' : (c.number || c.currency ? 'number' : 'text'))
      }));
      return ok(res, { columns });
    } catch (e) { return bad(res, 502, String(e.message || e)); }
  }

  return bad(res, 404, 'Unknown API endpoint');
}

function oauthCallbackPage(title, body, okFlag) {
  return `<!doctype html><meta charset="utf-8"><title>${title}</title>
<body style="font-family:Segoe UI,system-ui,sans-serif;background:#f4f6f8;display:grid;place-items:center;height:100vh;margin:0">
<div style="background:#fff;padding:36px 44px;border-radius:12px;box-shadow:0 8px 30px rgba(15,40,60,.12);max-width:560px;text-align:center">
<div style="font-size:40px">${okFlag ? '✅' : '⚠️'}</div>
<h2 style="margin:10px 0 6px;color:#1c2733">${title}</h2>
<p style="color:#52606d;line-height:1.5">${body}</p>
${okFlag ? '<p style="color:#52606d">Taking you back to the app…</p><script>setTimeout(function(){location.href="/#settings"},1800)</script>'
         : '<p><a href="/#settings" style="color:#0b5e8a">Back to Settings</a></p>'}
</div></body>`;
}

async function handleOauthCallback(req, res, u) {
  const code = u.searchParams.get('code');
  const error = u.searchParams.get('error');
  if (error) return send(res, 200, oauthCallbackPage('Zoho sign-in was cancelled', 'Zoho said: <b>' + error + '</b>. No changes were made.', false), MIME['.html']);
  if (!code) return send(res, 200, oauthCallbackPage('Missing code', 'Zoho did not send an authorisation code.', false), MIME['.html']);
  try {
    const s = getSettings();
    const j = await tokenRequest({
      grant_type: 'authorization_code', code,
      client_id: s.clientId, client_secret: s.clientSecret, redirect_uri: redirectUri()
    }, s.dc);
    if (!j.refresh_token) {
      return send(res, 200, oauthCallbackPage('Nearly there', 'Zoho did not return a long-lived token. In Zoho, remove this app\'s access (My Account → Security → Connected apps) and connect again.', false), MIME['.html']);
    }
    const cur = readJson(FILES.settings, {});
    cur.refreshToken = j.refresh_token;
    if (j.api_domain) cur.apiDomain = j.api_domain;
    writeJson(FILES.settings, cur);
    tokenCache = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };

    // Auto-select the organisation if there is exactly one.
    try {
      const o = await zohoGet('/organizations', {}, false);
      const orgs = o.organizations || [];
      if (orgs.length === 1) {
        const cur2 = readJson(FILES.settings, {});
        cur2.orgId = String(orgs[0].organization_id);
        cur2.orgName = orgs[0].name;
        writeJson(FILES.settings, cur2);
      }
    } catch (e) { /* org can be picked manually in Settings */ }

    return send(res, 200, oauthCallbackPage('Connected to Zoho Books', 'The app can now read your jobs. Next: check the organisation in Settings, then press <b>Sync now</b>.', true), MIME['.html']);
  } catch (e) {
    return send(res, 200, oauthCallbackPage('Connection failed', String(e.message || e), false), MIME['.html']);
  }
}

async function handleMsCallback(req, res, u) {
  const code = u.searchParams.get('code');
  const error = u.searchParams.get('error');
  const desc = String(u.searchParams.get('error_description') || '').split(/\r?\n/)[0];
  if (error) return send(res, 200, oauthCallbackPage('Microsoft sign-in was cancelled', 'Microsoft said: <b>' + error + '</b><br>' + desc, false), MIME['.html']);
  if (!code || !msAuthPending || u.searchParams.get('state') !== msAuthPending.state) {
    return send(res, 200, oauthCallbackPage('Sign-in could not be completed', 'The sign-in attempt lost its security code along the way. Go back to Settings and press Connect again.', false), MIME['.html']);
  }
  try {
    const s = getSettings();
    const r = await fetch(msLoginBase(s.ms.tenant) + '/token', {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: s.ms.clientId, grant_type: 'authorization_code', code,
        redirect_uri: msRedirectUri(), code_verifier: msAuthPending.verifier, scope: MS_SCOPES
      })
    });
    const j = await r.json().catch(() => ({}));
    msAuthPending = null;
    if (!j.refresh_token) {
      const why = String(j.error_description || j.error || 'Microsoft did not return a long-lived token.').split(/\r?\n/)[0].slice(0, 300);
      return send(res, 200, oauthCallbackPage('Connection failed', why, false), MIME['.html']);
    }
    const cur = readJson(FILES.settings, {});
    cur.ms = Object.assign({}, cur.ms, { refreshToken: j.refresh_token });
    writeJson(FILES.settings, cur);
    msTokenCache = { token: j.access_token, exp: Date.now() + (j.expires_in || 3600) * 1000 };
    return send(res, 200, oauthCallbackPage('Connected to Microsoft 365', 'The app can now read SharePoint. Next: pick your site and the lead list in Settings.', true), MIME['.html']);
  } catch (e) {
    return send(res, 200, oauthCallbackPage('Connection failed', String(e.message || e), false), MIME['.html']);
  }
}

function serveStatic(req, res, u) {
  let rel = decodeURIComponent(u.pathname);
  if (rel === '/' || rel === '') rel = '/index.html';
  if (rel === '/favicon.ico') { res.writeHead(204); return res.end(); }
  const file = path.normalize(path.join(PUB_DIR, rel));
  if (!file.startsWith(PUB_DIR)) return bad(res, 403, 'Forbidden');
  fs.readFile(file, (err, buf) => {
    if (err) return bad(res, 404, 'Not found');
    send(res, 200, buf, MIME[path.extname(file).toLowerCase()] || 'application/octet-stream');
  });
}

// The one request handler — shared by the local server (below) and the Vercel
// serverless entry (api/index.js). Plain (req, res) so it works in both.
async function handleRequest(req, res) {
  const u = new URL(req.url, 'http://' + (req.headers.host || 'localhost'));
  try {
    if (u.pathname.startsWith('/api/')) {
      // Gate every API call behind a valid Supabase session (except the public
      // auth-config probe the login screen needs before anyone is signed in).
      if (authRequired() && u.pathname !== '/api/auth/config') {
        const user = await verifyUser(bearerToken(req));
        if (!user) return bad(res, 401, 'Please sign in.');
        req.user = user;
      }
      return await handleApi(req, res, u);
    }
    if (u.pathname === '/oauth/callback') return await handleOauthCallback(req, res, u);
    if (u.pathname === '/ms/callback') return await handleMsCallback(req, res, u);
    return serveStatic(req, res, u);
  } catch (e) {
    try { bad(res, 500, String((e && e.message) || e)); } catch (_) { /* response already sent */ }
  }
}

// Initialize storage once (loads + migrates from Supabase if configured). Memoized
// so the serverless entry can cheaply await it on every invocation; on Vercel this
// runs once per cold start.
let _ready = null;
function ensureReady() {
  if (!_ready) _ready = (async () => {
    try { await store.init({ dataDir: DATA_DIR, files: FILES }); }
    catch (e) { console.error('[store] init error: ' + (e && e.message)); }
    ensureDefaults();
  })();
  return _ready;
}

// Run directly (`node server.js`) -> start an always-on HTTP server for local use.
// When this module is IMPORTED (by api/index.js on Vercel) we do NOT listen — the
// serverless platform invokes handleRequest() per request instead.
if (require.main === module) {
  const server = http.createServer(handleRequest);
  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE') {
      console.log('\nThe app looks like it is already running.');
      console.log('Open  http://localhost:' + PORT + '  in your browser.\n');
      process.exit(0);
    }
    throw e;
  });
  ensureReady().then(() => {
    server.listen(PORT, HOST, () => {
      console.log('\n  Hydro-Wates Project Manager  [storage: ' + store.mode() + ']');
      console.log('  ---------------------------');
      console.log('  Open:  http://localhost:' + PORT);
      console.log('  Stop:  close this window (or press Ctrl+C)\n');
    });
  });
}

module.exports = { handleRequest, ensureReady };
