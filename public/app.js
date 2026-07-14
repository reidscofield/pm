/* Hydro-Wates Project Manager — front end */
'use strict';

const BUILD = 'build 2026-07-14 · 68';

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));
const esc = (s) => String(s === undefined || s === null ? '' : s)
  .replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// Can the CURRENT user send email? Only if they signed in with Microsoft (so we
// have their token to send as them). No shared account, no SMTP.
function emailReady() {
  return !!msGraphToken;
}

// The email sign-off = the signed-in person's name (so it matches who's actually
// sending), falling back to the company name when a login carries no name.
function userFromSession(session) {
  if (!session || !session.user) return null;
  const u = session.user, md = u.user_metadata || {};
  return { email: u.email || '', name: String(md.full_name || md.name || md.display_name || '').trim() };
}
function signOff() {
  return (currentUser && currentUser.name) || 'Hydro-Wates';
}

const CATS = [
  ['rental', 'Rental'],
  ['service', 'Service'],
  ['sales', 'Sales']
];
const STAGES = [
  ['new', 'New'], ['planning', 'Planning'], ['scheduled', 'Scheduled'],
  ['inprogress', 'In progress'], ['complete', 'Complete'], ['onhold', 'On hold']
];
const PLAN_LABELS = { none: 'No planning yet', prepared: 'Questions prepared', sent: 'Sent to customer', answered: 'Answers received' };
const Q_TYPES = [['text', 'Text'], ['number', 'Number'], ['choice', 'Choice (dropdown)'], ['boolean', 'Yes / No']];
const Q_TYPE_LABEL = { text: 'Text', number: 'Number', choice: 'Choice', boolean: 'Yes / No' };

// Customer-facing hint appended to a question in the email, so they answer in
// the shape we'll record it (purely cosmetic — the typed capture is internal).
function answerHint(q) {
  const type = q.type || 'text';
  if (type === 'boolean') return '  (Yes / No)';
  if (type === 'choice' && (q.options || []).length) return '  (' + q.options.join(' / ') + ')';
  if (type === 'number') return q.unit ? '  (answer in ' + q.unit + ')' : '';
  return '';
}

// Human-readable form of a structured answer (mirrors the server's displayAnswer).
function answerDisplay(q) {
  const v = q.value == null ? '' : String(q.value).trim();
  if (!v) return '';
  if (q.type === 'number') return q.unit ? v + ' ' + q.unit : v;
  if (q.type === 'boolean') return v === 'yes' ? 'Yes' : v === 'no' ? 'No' : '';
  return v;
}

// The typed input a customer's answer is captured through, chosen by the question's type.
function answerInputHtml(q) {
  const type = q.type || 'text';
  // legacy records stored the answer under `answer`; fall back to it for text.
  const raw = q.value != null ? q.value : (q.answer || '');
  if (type === 'boolean') {
    const v = q.value || '';
    return '<select data-qval class="answer-input" style="margin-top:6px">' +
      '<option value=""' + (v === '' ? ' selected' : '') + '>— not answered —</option>' +
      '<option value="yes"' + (v === 'yes' ? ' selected' : '') + '>Yes</option>' +
      '<option value="no"' + (v === 'no' ? ' selected' : '') + '>No</option>' +
      '</select>';
  }
  if (type === 'choice') {
    const v = q.value || '';
    return '<select data-qval class="answer-input" style="margin-top:6px">' +
      '<option value="">— not answered —</option>' +
      (q.options || []).map(o => '<option value="' + esc(o) + '"' + (v === o ? ' selected' : '') + '>' + esc(o) + '</option>').join('') +
      '</select>';
  }
  if (type === 'number') {
    const units = (q.units && q.units.length) ? q.units : null;
    const unitCtrl = units
      ? '<select data-qunit class="unit-sel">' + units.map(u => '<option' + (q.unit === u ? ' selected' : '') + '>' + esc(u) + '</option>').join('') + '</select>'
      : (q.unit ? '<span class="unit">' + esc(q.unit) + '</span>' : '');
    return '<span class="num-answer">' +
      '<input data-qval type="number" step="any" class="answer-input" style="margin-top:6px" placeholder="Customer\'s answer…" value="' + esc(raw) + '">' +
      unitCtrl +
      '</span>';
  }
  return '<input data-qval class="answer-input" style="margin-top:6px" placeholder="Customer\'s answer…" value="' + esc(raw) + '">';
}
const MOD_BADGE = { estimates: 'EST', salesorders: 'SO', invoices: 'INV', projects: 'PRJ', lead: 'PO' };
const MOD_NAME = { estimates: 'Estimate', salesorders: 'Sales order', invoices: 'Invoice', projects: 'Project', lead: 'PO lead' };

const state = {
  view: 'dashboard',
  jobs: [], meta: { demo: true, connected: false, lastSync: null, sync: {} },
  templates: null, settings: null,
  search: '', stageFilter: '', showHidden: false, compact: false,
  open: null, detail: null, planningDraft: null, procedureDraft: null, procSendConfirm: null, modalTab: 'details',
  contacts: null, contactsError: '',
  orgs: null, syncTimer: null,
  leads: null, leadsMeta: {}, poSearch: '', poShowCompleted: true,
  smJobs: null, smSearch: '', smShowAll: false, smInvoicedCollapsed: true,
  msSites: null, msLists: null, msCols: null, _msColsLoading: false
};

/* ---------------- api ---------------- */
async function api(method, url, body) {
  const opt = { method, headers: {} };
  if (authToken) opt.headers['Authorization'] = 'Bearer ' + authToken;
  if (body !== undefined) { opt.headers['Content-Type'] = 'application/json'; opt.body = JSON.stringify(body); }
  const r = await fetch(url, opt);
  const j = await r.json().catch(() => ({}));
  if (r.status === 401) { renderLogin('Your session has expired — please sign in again.'); throw new Error('Not signed in'); }
  if (!r.ok) throw new Error(j.error || ('Request failed (' + r.status + ')'));
  return j;
}

/* ---------------- auth (Supabase login) ---------------- */
let SB = null;          // supabase-js client (only when login is required)
let authToken = null;   // current access token, sent on every api() call
let msGraphToken = null; // the logged-in user's Microsoft token — used to send email AS them
let currentUser = null;  // { email, name } of the signed-in user — used for the email sign-off

// Returns true if the app may proceed, false if the login screen is showing.
async function initAuth() {
  let cfg;
  try { cfg = await (await fetch('/api/auth/config')).json(); }
  catch (e) { return true; }                 // can't reach config — let the app load and surface errors
  if (!cfg.required) return true;            // login turned off (local dev) — proceed
  if (!window.supabase || !cfg.url || !cfg.anonKey) {
    renderLogin('Login is required but Supabase isn’t fully configured yet.');
    return false;
  }
  SB = window.supabase.createClient(cfg.url, cfg.anonKey, {
    auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true }
  });
  SB.auth.onAuthStateChange((_e, session) => {
    authToken = session ? session.access_token : null;
    currentUser = userFromSession(session);
    if (session && session.provider_token) msGraphToken = session.provider_token;   // Microsoft sign-in token
  });
  const { data: { session } } = await SB.auth.getSession();
  if (!session) { renderLogin(); return false; }
  authToken = session.access_token;
  currentUser = userFromSession(session);
  if (session.provider_token) msGraphToken = session.provider_token;   // present right after a Microsoft sign-in
  const lb = document.getElementById('logoutBtn');
  if (lb) { lb.hidden = false; lb.textContent = 'Sign out' + (session.user && session.user.email ? ' · ' + session.user.email : ''); }
  return true;
}

function renderLogin(msg) {
  const gate = document.getElementById('authGate');
  if (!gate) return;
  gate.innerHTML =
    '<div class="auth-overlay"><div class="auth-card">' +
      '<img class="auth-logo" id="authLogo" src="/logo.jpg" alt="Hydro-Wates — Proof-Load Testing Services">' +
      '<div class="auth-title">Project Management</div>' +
      '<p class="auth-sub">Sign in with your Hydro-Wates account.</p>' +
      '<p id="authErr" class="auth-err"' + (msg ? '' : ' hidden') + '>' + (msg ? esc(msg) : '') + '</p>' +
      '<button class="btn auth-ms" id="authMs">' +
        '<svg class="ms-logo" viewBox="0 0 21 21" width="17" height="17" aria-hidden="true">' +
          '<rect x="1" y="1" width="9" height="9" fill="#F25022"></rect>' +
          '<rect x="11" y="1" width="9" height="9" fill="#7FBA00"></rect>' +
          '<rect x="1" y="11" width="9" height="9" fill="#00A4EF"></rect>' +
          '<rect x="11" y="11" width="9" height="9" fill="#FFB900"></rect>' +
        '</svg>Sign in with Microsoft</button>' +
      '<p class="auth-hint">Use your Hydro-Wates Microsoft account. New staff are signed in automatically — no account setup needed.</p>' +
      '<button type="button" class="auth-alt" id="authShowEmail">Can’t use Microsoft? Sign in with email</button>' +
      '<div id="authEmailWrap" hidden>' +
        '<div class="auth-or"><span>work email</span></div>' +
        '<form id="authForm">' +
          '<input id="authEmail" type="email" placeholder="you@hydrowates.com" required autocomplete="username">' +
          '<input id="authPass" type="password" placeholder="Password" required autocomplete="current-password">' +
          '<button class="btn primary" type="submit">Sign in</button>' +
        '</form>' +
      '</div>' +
    '</div></div>';
  const form = document.getElementById('authForm');
  if (form) form.addEventListener('submit', doEmailLogin);
  const ms = document.getElementById('authMs');
  if (ms) ms.addEventListener('click', doMicrosoftLogin);
  const showEmail = document.getElementById('authShowEmail');
  if (showEmail) showEmail.addEventListener('click', () => {
    const w = document.getElementById('authEmailWrap'); if (w) w.hidden = false;
    showEmail.hidden = true;
  });
  const logo = document.getElementById('authLogo');
  if (logo) logo.onerror = function () { this.outerHTML = '<div class="auth-brand"><span class="brand-drop">💧</span> Hydro-Wates</div>'; };
}

async function doEmailLogin(e) {
  e.preventDefault();
  const email = document.getElementById('authEmail').value.trim();
  const password = document.getElementById('authPass').value;
  const err = document.getElementById('authErr');
  const btn = e.target.querySelector('button[type="submit"]');
  if (err) err.hidden = true;
  if (btn) { btn.disabled = true; btn.textContent = 'Signing in…'; }
  const { error } = await SB.auth.signInWithPassword({ email, password });
  if (error) {
    if (err) { err.textContent = error.message; err.hidden = false; }
    if (btn) { btn.disabled = false; btn.textContent = 'Sign in'; }
    return;
  }
  location.reload();   // re-init with the new session
}
async function doMicrosoftLogin() {
  if (!SB) return;
  // Ask for Mail.Send so the app can email AS the person who's signing in.
  await SB.auth.signInWithOAuth({
    provider: 'azure',
    options: { redirectTo: location.origin, scopes: 'openid email profile offline_access https://graph.microsoft.com/Mail.Send' }
  });
}

// ---- Silent, on-demand Microsoft re-auth (to send email AS you) --------------
// The Microsoft "send-mail" token is handed over only at sign-in and isn't kept
// across reloads. Rather than store a long-lived refresh token, we fetch a fresh
// one on demand: a small popup does the Microsoft round-trip (usually silent, since
// you already have an M365 session) and hands the token back via postMessage.
// Nothing long-lived is stored — the token stays in memory for this session only.
async function ensureGraphToken() {
  if (msGraphToken) return msGraphToken;
  return await reauthMicrosoftPopup();
}
async function reauthMicrosoftPopup() {
  if (!SB || !window.supabase) return null;
  const width = 520, height = 690;
  const left = Math.max(0, (window.screenX || 0) + (((window.outerWidth || 1024) - width) / 2));
  const top = Math.max(0, (window.screenY || 0) + (((window.outerHeight || 768) - height) / 2));
  // Open synchronously (inside the click gesture) or the browser blocks the popup.
  const popup = window.open('about:blank', 'hw-ms-reauth', 'width=' + width + ',height=' + height + ',left=' + left + ',top=' + top);
  if (!popup) { toast('Please allow popups so the app can confirm your Microsoft sign-in.', true); return null; }
  try { popup.document.write('<p style="font:15px system-ui,-apple-system,sans-serif;color:#334;padding:28px">Opening Microsoft sign-in…</p>'); } catch (e) {}
  try { localStorage.setItem('hw_ms_reauth', String(Date.now())); } catch (e) {}
  let url = null;
  try {
    const { data } = await SB.auth.signInWithOAuth({
      provider: 'azure',
      options: { skipBrowserRedirect: true, redirectTo: location.origin, scopes: 'openid email profile offline_access https://graph.microsoft.com/Mail.Send' }
    });
    url = data && data.url;
  } catch (e) {}
  if (!url) { try { popup.close(); } catch (e) {} try { localStorage.removeItem('hw_ms_reauth'); } catch (e) {} return null; }
  try { popup.location.href = url; } catch (e) { try { popup.close(); } catch (x) {} return null; }
  return await new Promise((resolve) => {
    let done = false;
    const finish = (tok) => {
      if (done) return; done = true;
      window.removeEventListener('message', onMsg);
      clearInterval(watch); clearTimeout(timer);
      try { localStorage.removeItem('hw_ms_reauth'); } catch (e) {}
      if (tok) msGraphToken = tok;
      resolve(tok || null);
    };
    const onMsg = (ev) => {
      if (ev.origin !== location.origin || !ev.data || ev.data.type !== 'ms-graph-token') return;
      try { popup.close(); } catch (x) {}
      finish(ev.data.token || null);
    };
    window.addEventListener('message', onMsg);
    const watch = setInterval(() => { if (popup.closed) finish(null); }, 500);
    const timer = setTimeout(() => { try { popup.close(); } catch (x) {} finish(null); }, 120000);
  });
}
// Runs INSIDE the re-auth popup: finish the round-trip, hand the token back, close.
// Guarded by a recent localStorage flag + window.opener so it never fires on a normal load.
async function handleMsPopupCallback() {
  let flag = null;
  try { flag = localStorage.getItem('hw_ms_reauth'); } catch (e) {}
  if (!(flag && window.opener && (Date.now() - Number(flag) < 300000))) return false;
  try { localStorage.removeItem('hw_ms_reauth'); } catch (e) {}
  try { if (document.body) document.body.innerHTML = '<p style="font:15px system-ui,-apple-system,sans-serif;color:#334;padding:28px">Finishing Microsoft sign-in… this window will close automatically.</p>'; } catch (e) {}
  let token = null;
  try {
    const cfg = await (await fetch('/api/auth/config')).json();
    if (window.supabase && cfg.url && cfg.anonKey) {
      const sb = window.supabase.createClient(cfg.url, cfg.anonKey, { auth: { persistSession: true, autoRefreshToken: false, detectSessionInUrl: true } });
      for (let i = 0; i < 30 && !token; i++) {
        const { data: { session } } = await sb.auth.getSession();
        if (session && session.provider_token) token = session.provider_token;
        else await new Promise(r => setTimeout(r, 200));
      }
    }
  } catch (e) {}
  try { if (window.opener) window.opener.postMessage({ type: 'ms-graph-token', token: token }, location.origin); } catch (e) {}
  setTimeout(() => { try { window.close(); } catch (e) {} }, 300);
  return true;
}
async function doLogout() {
  try { if (SB) await SB.auth.signOut(); } catch (e) {}
  authToken = null;
  location.reload();
}

/* ---------------- formatting ---------------- */
function fmtMoney(total, currency) {
  if (total === null || total === undefined) return '—';
  try { return new Intl.NumberFormat(undefined, { style: 'currency', currency: currency || 'USD', maximumFractionDigits: 0 }).format(total); }
  catch (e) { return (currency ? currency + ' ' : '') + Number(total).toLocaleString(); }
}
function fmtDate(d) {
  if (!d) return '—';
  const dt = new Date(d + (d.length === 10 ? 'T00:00:00' : ''));
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
}
function stageChip(stage) {
  const label = (STAGES.find(s => s[0] === stage) || ['', stage])[1];
  return '<span class="chip stage-' + esc(stage) + '">' + esc(label) + '</span>';
}
function planChip(st) {
  const icon = { none: '·', prepared: '✎', sent: '✉', answered: '✓' }[st] || '·';
  return '<span class="chip plan-' + esc(st) + '">' + icon + ' ' + esc(PLAN_LABELS[st] || st) + '</span>';
}

function toast(msg, isError) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.toggle('error', !!isError);
  t.hidden = false;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.hidden = true; }, isError ? 6500 : 3200);
}

/* ---------------- data loading ---------------- */
async function loadJobs() {
  const j = await api('GET', '/api/jobs');
  state.jobs = j.jobs || [];
  state.meta = j;
  try { const t = await api('GET', '/api/meetings/todo'); state.mtgTodo = t.todo || []; } catch (e) { state.mtgTodo = state.mtgTodo || []; }
  updateSyncStatus();
}
async function loadAll() {
  const [t, s] = await Promise.all([api('GET', '/api/templates'), api('GET', '/api/settings')]);
  state.templates = t; state.settings = s;
  await loadJobs();
}

/* ---------------- header / sync ---------------- */
function updateSyncStatus() {
  const el = $('#syncStatus');
  const sync = state.meta.sync || {};
  if (sync.running) {
    el.innerHTML = '<span class="spin">⟳</span> ' + esc(sync.message || 'Syncing…') +
      (sync.total ? ' (' + sync.done + '/' + sync.total + ')' : '');
  } else if (state.meta.demo) {
    el.textContent = 'Demo data';
  } else if (state.meta.lastSync) {
    const dt = new Date(state.meta.lastSync);
    el.textContent = 'Last sync ' + dt.toLocaleDateString(undefined, { day: 'numeric', month: 'short' }) +
      ' ' + dt.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
  } else {
    el.textContent = state.meta.connected ? 'Not synced yet' : 'Not connected';
  }
}

function startSyncPoll() {
  if (state.syncTimer) return;
  state.syncTimer = setInterval(async () => {
    try {
      const st = await api('GET', '/api/sync/status');
      state.meta.sync = st;
      state.meta.lastSync = st.lastSync;
      updateSyncStatus();
      if (!st.running) {
        clearInterval(state.syncTimer); state.syncTimer = null;
        if (st.error) toast(st.error, true);
        else { await loadJobs(); render(); toast('Sync finished — ' + state.jobs.length + ' jobs loaded.'); }
      }
    } catch (e) { /* server briefly busy — keep polling */ }
  }, 1200);
}

async function syncNow() {
  if (state.meta.demo && !state.meta.connected) {
    toast('Demo mode — connect Zoho Books in Settings first.', true);
    location.hash = '#settings';
    return;
  }
  try {
    await api('POST', '/api/sync');
    state.meta.sync = { running: true, message: 'Starting sync…' };
    updateSyncStatus();
    startSyncPoll();
  } catch (e) { toast(e.message, true); }
}

/* ---------------- dashboard ---------------- */
function filteredJobs() {
  const q = state.search.trim().toLowerCase();
  return state.jobs.filter(j => {
    if (j.hidden && !state.showHidden) return false;
    if (state.stageFilter && j.stage !== state.stageFilter) return false;
    if (q && !((j.customer || '').toLowerCase().includes(q) ||
               (j.number || '').toLowerCase().includes(q) ||
               (j.reference || '').toLowerCase().includes(q))) return false;
    return true;
  });
}

function jobCard(j) {
  return '<div class="card" data-action="open-job" data-key="' + esc(j.key) + '">' +
    '<button class="card-x" data-action="job-remove-card" data-key="' + esc(j.key) + '" title="Remove from board">✕</button>' +
    (j.hwi ? '<div class="card-hwi">' + esc(j.hwi) + '</div>'
           : '<div class="card-hwi missing" title="No job number found in Zoho or the Lead List — add the HWI to this job">No job&nbsp;#</div>') +
    '<div class="card-top">' +
      '<span class="card-customer">' + esc(j.customer || '(no customer)') + '</span>' +
      '<span class="card-total">' + esc(fmtMoney(j.total, j.currency)) + '</span>' +
    '</div>' +
    '<div class="card-mid">' +
      '<span class="chip mod">' + esc(MOD_BADGE[j.module] || j.module) + '</span>' +
      (j.number && j.number !== j.hwi ? '<span>' + esc(j.number) + '</span>' : '') +
      '<span>·</span><span>' + esc(fmtDate(j.createdDate || j.date)) + '</span>' +
      (j.status ? '<span class="chip zstatus">' + esc(j.status) + '</span>' : '') +
      (j.hidden ? '<span class="chip zstatus">hidden</span>' : '') +
    '</div>' +
    '<div class="card-bottom">' +
      stageChip(j.stage) + planChip(j.planningStatus) +
      (j.category === 'service'
        ? '<span class="chip mtg' + (j.preHeld ? ' on' : '') + '" title="Pre-job meeting ' + (j.preHeld ? 'held' : 'not held yet') + '">Pre ' + (j.preHeld ? '✓' : '–') + '</span>' +
          '<span class="chip mtg' + (j.postHeld ? ' on' : '') + '" title="Post-job meeting ' + (j.postHeld ? 'held' : 'not held yet') + '">Post ' + (j.postHeld ? '✓' : '–') + '</span>'
        : '') +
      (j.multiInvoice ? '<span class="chip multi" title="Multi-invoice job — stays on the board through staged billing">multi-invoice</span>' : '') +
      (j.categoryOverridden ? '<span class="chip override">manual</span>' : '') +
      (j.archived ? '<span class="chip finished">finished</span>' : '') +
    '</div>' +
  '</div>';
}

// One dense line per job for the Compact dashboard view.
function jobCardCompact(j) {
  return '<div class="card compact" data-action="open-job" data-key="' + esc(j.key) + '">' +
    (j.hwi ? '<span class="cc-hwi">' + esc(j.hwi) + '</span>'
           : '<span class="cc-hwi missing" title="No job number found — add the HWI to this job">No job&nbsp;#</span>') +
    '<span class="cc-cust">' + esc(j.customer || '(no customer)') + '</span>' +
    (j.planningStatus && j.planningStatus !== 'none' ? planChip(j.planningStatus) : '') +
    (j.archived ? '<span class="chip finished">finished</span>' : '') +
    stageChip(j.stage) +
    (j.category === 'service' ? '<span class="cc-mtg" title="Pre / Post meeting (✓ = held)">' + (j.preHeld ? '✓' : '·') + (j.postHeld ? '✓' : '·') + '</span>' : '') +
    '<span class="cc-total">' + esc(fmtMoney(j.total, j.currency)) + '</span>' +
    '<button class="card-x" data-action="job-remove-card" data-key="' + esc(j.key) + '" title="Remove from board">✕</button>' +
  '</div>';
}

function colsHtml() {
  const jobs = filteredJobs().slice().sort((a, b) =>
    ((b.createdDate || b.date) || '').localeCompare((a.createdDate || a.date) || ''));
  const card = state.compact ? jobCardCompact : jobCard;
  return '<div class="cols' + (state.compact ? ' compact' : '') + '">' + CATS.map(([cat, label]) => {
    const list = jobs.filter(j => j.category === cat);
    const sum = list.reduce((acc, j) => acc + (j.total || 0), 0);
    return '<div class="col col-' + cat + '">' +
      '<div class="col-head">' + esc(label) +
        ' <span class="count">' + list.length + '</span>' +
        (sum ? '<span class="sum">' + esc(fmtMoney(sum, (list[0] || {}).currency)) + '</span>' : '') +
      '</div>' +
      (list.length ? list.map(card).join('') : '<div class="empty-col">No ' + esc(label.toLowerCase()) + ' jobs match</div>') +
    '</div>';
  }).join('') + '</div>';
}

function renderDashboard() {
  const banner = state.meta.demo
    ? '<div class="banner">🔎 <b>Showing demo data.</b> Connect your Zoho Books account to load real jobs.' +
      '<button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>'
    : (state.meta.sync && state.meta.sync.error
      ? '<div class="banner">⚠️ Last sync problem: ' + esc(state.meta.sync.error) +
        '<button class="btn" data-action="sync-now">Try again</button></div>' : '');

  $('#view').innerHTML = banner +
    '<div class="toolbar">' +
      '<input type="search" id="searchBox" placeholder="Search customer, number or reference…" value="' + esc(state.search) + '">' +
      '<select id="stageFilter" data-change="stage-filter">' +
        '<option value="">All stages</option>' +
        STAGES.map(([v, l]) => '<option value="' + v + '"' + (state.stageFilter === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select>' +
      '<label class="chk"><input type="checkbox" data-change="show-hidden"' + (state.showHidden ? ' checked' : '') + '> show hidden</label>' +
      '<label class="chk"><input type="checkbox" data-change="compact"' + (state.compact ? ' checked' : '') + '> compact</label>' +
      '<button class="btn small" data-action="open-removed" title="Jobs removed from the board — restore them here">🗑 Recently deleted</button>' +
      '<span class="muted" style="margin-left:auto;font-size:13px">' + filteredJobs().length + ' jobs</span>' +
    '</div>' +
    '<div class="dash-body">' +
      '<div id="cols">' + colsHtml() + '</div>' +
      postJobSideHtml() +
    '</div>';
}

/* ---------------- job modal ---------------- */
async function openJob(key) {
  try {
    const d = await api('GET', '/api/job/' + encodeURIComponent(key));
    state.open = key;
    state.detail = d;
    state.planningDraft = d.planning ? JSON.parse(JSON.stringify(d.planning)) : null;
    if (state.planningDraft && !state.planningDraft.email) state.planningDraft.email = d.job.email || '';
    state.procedureDraft = d.procedure ? JSON.parse(JSON.stringify(d.procedure)) : null;
    state.procedureSetup = (state.procedureDraft && state.procedureDraft.setup)
      ? Object.assign(defaultSetup(d.job, d.planning), state.procedureDraft.setup)
      : defaultSetup(d.job, d.planning);
    state.meetingsDraft = d.meetings ? JSON.parse(JSON.stringify(d.meetings)) : { pre: {}, post: {}, notes: '' };
    state.procSendConfirm = null;
    state.modalTab = (d.planning && d.planning.status !== 'none') ? 'planning' : 'details';
    state.contacts = null; state.contactsError = '';
    renderModal();
    loadContacts(key);
  } catch (e) { toast(e.message, true); }
}

function loadContacts(key) {
  api('GET', '/api/job/' + encodeURIComponent(key) + '/contacts')
    .then(r => {
      if (state.open !== key) return;
      state.contacts = r.contacts || [];
      const box = $('#contactsBox');
      if (box) box.outerHTML = contactsBoxHtml();
    })
    .catch(e => {
      if (state.open !== key) return;
      state.contacts = [];
      state.contactsError = e.message;
      const box = $('#contactsBox');
      if (box) box.outerHTML = contactsBoxHtml();
    });
}

function closeModal() {
  if (state.modalTab === 'planning' && state.planningDraft) collectPlanning(); // keep typed text in memory
  else if (state.modalTab === 'procedure' && state.procedureDraft) collectProcedure();
  else if (state.modalTab === 'meetings') collectMeetings();
  state.open = null; state.detail = null; state.planningDraft = null; state.procedureDraft = null; state.meetingsDraft = null; state.procSendConfirm = null;
  $('#modal').innerHTML = '';
  if (state.view === 'dashboard') renderDashboard();
}

function detailTabHtml(d) {
  const j = d.job;
  const zohoLink = (!j.demo && j.module !== 'lead' && d.zoho && d.zoho.orgId)
    ? '<a class="btn small" target="_blank" href="https://books.zoho.' + esc(d.zoho.dc) + '/app/' + esc(d.zoho.orgId) + '#/' + esc(j.module) + '/' + esc(j.id) + '">Open in Zoho Books ↗</a>'
    : '';
  const items = (j.lineItems && j.lineItems.length)
    ? '<table class="items"><tr><th>Item</th><th class="num">Qty</th><th class="num">Rate</th><th class="num">Amount</th></tr>' +
      j.lineItems.map(li =>
        '<tr><td>' + esc(li.name) + (li.description && li.description !== li.name ? '<div class="muted" style="font-size:12px">' + esc(li.description) + '</div>' : '') + '</td>' +
        '<td class="num">' + esc(li.quantity === undefined ? '' : li.quantity) + '</td>' +
        '<td class="num">' + esc(li.rate === undefined ? '' : Number(li.rate).toLocaleString()) + '</td>' +
        '<td class="num">' + esc(li.total === undefined ? '' : Number(li.total).toLocaleString()) + '</td></tr>'
      ).join('') + '</table>'
    : (j.module === 'lead'
      ? '<p class="muted">This job comes from the SharePoint lead list (PO ' + esc(j.reference || '—') + '). ' +
        'It will leave the dashboard automatically once an invoice carrying that PO number appears in Zoho Books — see the Invoices page.</p>'
      : '<p class="muted">No line items available for this record.</p>');

  return '<div class="detail-grid">' +
    (j.hwi ? '<div><div class="dt">HWI job no.</div><div class="dd"><b>' + esc(j.hwi) + '</b></div></div>' : '') +
    '<div><div class="dt">Type</div><div class="dd">' + esc(MOD_NAME[j.module] || j.module) + ' ' + esc(j.number) + '</div></div>' +
    '<div><div class="dt">Date</div><div class="dd">' + esc(fmtDate(j.date)) + '</div></div>' +
    '<div><div class="dt">Amount</div><div class="dd">' + esc(fmtMoney(j.total, j.currency)) + '</div></div>' +
    '<div><div class="dt">Zoho status</div><div class="dd">' + esc(j.status || '—') + '</div></div>' +
    (j.reference ? '<div><div class="dt">Reference</div><div class="dd">' + esc(j.reference) + '</div></div>' : '') +
    '<div><div class="dt">Category</div><div class="dd">' +
      '<select data-change="job-cat" style="width:auto;padding:4px 8px;font-size:13px">' +
        '<option value=""' + (!j.categoryOverridden ? ' selected' : '') + '>Auto' + (!j.categoryOverridden ? ' (' + esc(j.category) + ')' : '') + '</option>' +
        CATS.map(([v, l]) => '<option value="' + v + '"' + (j.categoryOverridden && j.category === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select></div></div>' +
    '<div><div class="dt">Stage</div><div class="dd">' +
      '<select data-change="job-stage" style="width:auto;padding:4px 8px;font-size:13px">' +
        STAGES.map(([v, l]) => '<option value="' + v + '"' + (j.stage === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select></div></div>' +
    '<div style="grid-column:1/-1"><div class="dt">Invoicing</div><div class="dd">' +
      '<label class="chk"><input type="checkbox" data-change="job-multi"' + (j.multiInvoice ? ' checked' : '') + '> <b>Multi-invoice job</b> — billed in stages; keep it on the board even after an invoice goes out</label>' +
    '</div></div>' +
    '<div><div class="dt">WLL · entered internally</div><div class="dd">' +
      '<input id="jobWll" data-change="job-wll" type="number" step="any" min="0" placeholder="—" style="width:78px;padding:4px 8px;font-size:13px" value="' + esc(j.wll == null ? '' : j.wll) + '"> ' +
      '<select id="jobWllUnit" data-change="job-wll" style="width:auto;padding:4px 6px;font-size:13px">' +
        ['t', 'short tons'].map(u => '<option value="' + u + '"' + ((j.wllUnit || 't') === u ? ' selected' : '') + '>' + u + '</option>').join('') +
      '</select></div></div>' +
    '<div><div class="dt">Test load · 125% proof</div><div class="dd" id="testLoad">' + testLoadText(j) + '</div></div>' +
  '</div>' +
  '<div class="dt" style="font-size:12px;color:var(--muted);margin-bottom:2px">Line items</div>' + items +
  '<div class="plan-actions">' + zohoLink +
    '<button class="btn small" data-action="job-hide">' + (j.hidden ? 'Unhide job' : 'Hide job') + '</button>' +
    (j.archivedHow === 'manual'
      ? '<button class="btn small" data-action="job-restore">↩ Restore to board</button>'
      : '<button class="btn small danger" data-action="job-remove">🗑 Remove from board</button>') +
  '</div>';
}

function contactsBoxHtml() {
  const proc = state.modalTab === 'procedure';
  const p = (proc ? state.procedureDraft : state.planningDraft) || {};
  const attr = proc ? 'data-pcontact' : 'data-contact';
  const heading = proc ? 'Send procedure to — contacts on file for this customer' : 'Send to — contacts on file for this customer';
  const cs = state.contacts;
  let inner;
  if (cs === null) {
    inner = '<span class="muted" style="font-size:13px"><span class="spin">⟳</span> Looking up the contacts on file…</span>';
  } else if (state.contactsError) {
    inner = '<span class="muted" style="font-size:13px">⚠️ ' + esc(state.contactsError) + '</span>';
  } else if (!cs.length) {
    inner = '<span class="muted" style="font-size:13px">No contacts on file for this customer — type an address below, or add a contact person in Zoho Books.</span>';
  } else {
    let sel = (p.recipients && p.recipients.length) ? p.recipients : cs.filter(c => c.isPrimary).map(c => c.email);
    if (!sel.length) sel = [cs[0].email];
    inner = cs.map(c =>
      '<label class="chk contact-row"><input type="checkbox" ' + attr + ' value="' + esc(c.email) + '"' + (sel.includes(c.email) ? ' checked' : '') + '> ' +
      '<span><b>' + esc(c.name) + '</b>' +
      (c.designation ? ' <span class="muted">· ' + esc(c.designation) + '</span>' : '') +
      ' <span class="muted">&lt;' + esc(c.email) + '&gt;</span>' +
      (c.isPrimary ? ' <span class="chip plan-prepared" style="font-size:10.5px">primary</span>' : '') +
      '</span></label>').join('');
  }
  return '<div id="contactsBox" class="contacts-box">' +
    '<label>' + heading + '</label>' + inner + '</div>';
}

function planningTabHtml(d) {
  const p = state.planningDraft;
  if (!p) {
    return '<p class="muted" style="line-height:1.6">No planning questions yet for this job.<br>' +
      'Load your standard question list — you can edit, add or remove questions for this customer before sending.</p>' +
      '<button class="btn primary" data-action="plan-init">Load standard questions</button>';
  }
  const qs = (p.questions || []).map((q, i) =>
    '<div class="qrow" data-qrow>' +
      '<div class="qnum">' + (i + 1) + '</div>' +
      '<div style="flex:1">' +
        '<textarea data-qtext rows="2">' + esc(q.text) + '</textarea>' +
        '<div class="qtype-tag">' + esc(Q_TYPE_LABEL[q.type] || 'Text') +
          (q.type === 'choice' && (q.options || []).length ? ': ' + esc(q.options.join(' / ')) : '') +
          (q.type === 'number' ? (q.units && q.units.length ? ' (' + esc(q.units.join(' / ')) + ')' : (q.unit ? ' (' + esc(q.unit) + ')' : '')) : '') + '</div>' +
        answerInputHtml(q) +
      '</div>' +
      '<div class="qbtns"><button class="btn-icon" title="Remove question" data-action="plan-del-q" data-i="' + i + '">✕</button></div>' +
    '</div>'
  ).join('');

  const smtpReady = emailReady();
  const log = (p.sentLog || []).slice().reverse().map(l =>
    '<div>' + esc(new Date(l.at).toLocaleString()) + ' → ' + esc((l.to || []).join(', ')) + '</div>').join('');

  return contactsBoxHtml() +
    '<div class="frow">' +
      '<div><label>Extra recipient (optional)</label><input id="planEmail" type="email" placeholder="name@customer.com" value="' + esc(p.email || '') + '"></div>' +
      '<div><label>Planning status</label><select id="planStatus">' +
        Object.entries(PLAN_LABELS).filter(([v]) => v !== 'none').map(([v, l]) =>
          '<option value="' + v + '"' + (p.status === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select></div>' +
    '</div>' +
    '<label style="margin-top:6px">Questions for this job</label>' + qs +
    '<button class="btn small" data-action="plan-add-q">+ Add question</button>' +
    '<div style="margin-top:14px"><label>Internal notes</label><textarea id="planNotes" rows="2" placeholder="Anything to remember for this job…">' + esc(p.notes || '') + '</textarea></div>' +
    '<div class="plan-actions">' +
      '<button class="btn primary" data-action="plan-send">✈ Send now</button>' +
      '<button class="btn" data-action="plan-email">✉ Open email draft</button>' +
      '<button class="btn" data-action="plan-copy">Copy questions</button>' +
      '<button class="btn" data-action="plan-print">🖨 Print / save PDF</button>' +
      '<button class="btn" data-action="plan-save">Save</button>' +
    '</div>' +
    (smtpReady ? '' : '<div class="sent-note">“Send now” emails the ticked contacts from <b>your own</b> mailbox once you <b>sign in with Microsoft</b>. Until then, use “Open email draft”.</div>') +
    (log ? '<div class="sent-note"><b>Send history</b>' + log + '</div>'
         : (p.sentAt ? '<div class="sent-note">First sent to customer: ' + esc(new Date(p.sentAt).toLocaleString()) + '</div>' : ''));
}

// Travel tab — the PM decides Fly or Drive for this job. The choice publishes to
// the shared portal table the travel app reads (cportal_projects.travel_mode by
// HWI), which then locks the travel app's trip planner to that mode.
function travelTabHtml(d) {
  const j = d.job;
  const hwi = j.hwi || '';
  if (!hwi) {
    return '<div class="travel-tab">' +
      '<p class="hint">This job doesn’t have an <b>HWI</b> number yet, so a travel decision can’t be linked to the travel app. ' +
      'Once the job has an HWI (after the next sync), Fly/Drive can be set here.</p></div>';
  }
  const m = j.travelMode || null;
  const btn = (val, icon, label) =>
    '<button class="tmode-btn' + (m === val ? ' active' : '') + '" ' +
      'data-action="set-travel-mode" data-hwi="' + esc(hwi) + '" data-mode="' + val + '" ' +
      'title="' + (m === val ? 'Click again to clear this decision' : 'Set to ' + label + ' — sends to the travel app') + '">' +
      '<span class="tmode-ico">' + icon + '</span><span class="tmode-label">' + label + '</span></button>';
  const decided = m
    ? '<div class="travel-status set"><b>' + (m === 'fly' ? '✈ Flying' : '🚗 Driving') + '</b> for ' + esc(hwi) +
        '. <span class="muted">Sent to the travel app — its trip planner is locked to this mode.</span></div>'
    : '<div class="travel-status">No decision yet. <span class="muted">Until you choose, the travel app lets the booker pick fly or drive.</span></div>';
  return '<div class="travel-tab">' +
    '<h3 style="margin:2px 0 4px">How is the crew getting there?</h3>' +
    '<p class="hint" style="margin-top:0">Decide how the team travels for <b>' + esc(hwi) + '</b>. Your choice is sent straight to the <b>travel app</b>, which locks its trip planner so the booker can’t pick the other mode by mistake.</p>' +
    '<div class="tmode tmode-lg" style="margin:14px 0">' + btn('fly', '✈', 'Fly') + btn('drive', '🚗', 'Drive') + '</div>' +
    decided +
    (m ? '<button class="btn ghost" data-action="set-travel-mode" data-hwi="' + esc(hwi) + '" data-mode="' + m + '" style="margin-top:14px">Clear decision</button>' : '') +
    '</div>';
}

// ---- Recently deleted: jobs manually removed from the board, restorable here ----
function removedOverlayHtml() {
  const list = state.removed || [];
  const rows = list.length ? list.map(r =>
    '<div class="rm-row">' +
      '<label class="rm-check"><input type="checkbox" data-rmsel value="' + esc(r.key) + '"></label>' +
      '<div class="rm-info"><b>' + esc(r.hwi || '(no HWI)') + '</b> · ' + esc(r.customer || '(no customer)') +
        '<div class="muted" style="font-size:12px">PO ' + esc(r.po || '—') +
          (r.removedAt ? ' · removed ' + esc(new Date(r.removedAt).toLocaleDateString()) : '') + '</div></div>' +
      '<div class="rm-actions">' +
        '<button class="btn small" data-action="removed-restore" data-key="' + esc(r.key) + '">↩ Restore</button>' +
        '<button class="btn small danger" data-action="removed-purge" data-key="' + esc(r.key) + '" title="Delete permanently — won’t appear on the board or here again">🗑 Delete</button>' +
      '</div>' +
    '</div>'
  ).join('') : '<div class="muted" style="padding:14px 2px">Nothing removed yet. Open a job and use <b>🗑 Remove from board</b> to send oddball jobs here.</div>';
  const bar = list.length
    ? '<div class="rm-bar">' +
        '<label class="chk"><input type="checkbox" id="rmSelAll" data-change="removed-selall"> Select all</label>' +
        '<button class="btn small" data-action="removed-del-selected" style="margin-left:auto">Delete selected</button>' +
        '<button class="btn small danger" data-action="removed-del-all">Delete all</button>' +
      '</div>'
    : '';
  return '<div class="overlay" data-action="removed-bg">' +
    '<div class="dialog" style="max-width:560px">' +
      '<div class="dialog-head"><div><h2>🗑 Recently deleted</h2>' +
        '<div class="sub">' + list.length + ' removed · restore, or permanently delete</div></div>' +
        '<button class="btn-icon dialog-close" data-action="removed-close" title="Close">✕</button></div>' +
      bar +
      '<div class="dialog-body">' + rows + '</div>' +
    '</div>' +
  '</div>';
}
async function openRemoved() {
  try { const r = await api('GET', '/api/removed'); state.removed = r.removed || []; }
  catch (e) { toast(e.message, true); return; }
  let host = document.getElementById('removedModal');
  if (!host) { host = document.createElement('div'); host.id = 'removedModal'; document.body.appendChild(host); }
  host.innerHTML = removedOverlayHtml();
}
function closeRemoved() { const h = document.getElementById('removedModal'); if (h) h.innerHTML = ''; }

// ---- Meetings tab: pre-job + post-job meeting tracking + notes ----
function meetingsTabHtml(d) {
  const m = state.meetingsDraft || { pre: {}, post: {}, notes: '' };
  const block = (key, title, hint) => {
    const b = m[key] || {};
    return '<div class="mtg-block' + (b.held ? ' held' : '') + '">' +
      '<div class="mtg-top">' +
        '<label class="chk"><input type="checkbox" id="mtg' + key + 'Held" data-change="mtg-held"' + (b.held ? ' checked' : '') + '> <b>' + title + '</b> — held</label>' +
        '<span class="mtg-date"><label>on</label><input id="mtg' + key + 'Date" type="date" value="' + esc(b.date || '') + '"></span>' +
      '</div>' +
      '<textarea id="mtg' + key + 'Notes" rows="4" placeholder="' + esc(hint) + '">' + esc(b.notes || '') + '</textarea>' +
    '</div>';
  };
  return '<div class="mtg">' +
    '<p class="muted" style="line-height:1.5;margin-bottom:6px">Track the meetings for this job — tick <b>held</b> and it stamps the date. A job with a pre-job meeting but no post-job one shows under <b>📋 Post-job to-do</b> on the dashboard.</p>' +
    block('pre', 'Pre-job meeting', 'Planning, coordination, safety brief — who attended, key points…') +
    mtgActionListHtml((m.pre && m.pre.actions) || []) +
    block('post', 'Post-job meeting', 'Debrief / lessons learned — what went well, what to improve, follow-ups…') +
    '<label style="margin-top:10px">General job notes</label>' +
    '<textarea id="mtgNotes" rows="4" placeholder="Anything else worth noting about this job…">' + esc(m.notes || '') + '</textarea>' +
    '<div class="plan-actions" style="margin-top:12px">' +
      '<button class="btn primary" data-action="mtg-save">Save</button>' +
      '<button class="btn" data-action="mtg-report-open" title="Email the action items to the team (except Mike Scofield)">✉ Send meeting report</button>' +
    '</div>' +
    (m.reportLog && m.reportLog.length
      ? '<div class="muted" style="font-size:12px;margin-top:6px">Report last sent ' + esc(String(m.reportLog[m.reportLog.length - 1].at || '').slice(0, 10)) +
        ' to ' + (m.reportLog[m.reportLog.length - 1].to || []).length + ' recipient' + ((m.reportLog[m.reportLog.length - 1].to || []).length === 1 ? '' : 's') + '.</div>'
      : '') +
  '</div>';
}
// ---- Pre-job action items: a live task list, each assignable to a Hydro-Wates staff member ----
function mtgActionListHtml(actions) {
  actions = actions || [];
  const team = (state.templates && state.templates.team) || [];
  const names = team.map(m => m.name);
  const open = actions.filter(a => !a.done).length;
  const rows = actions.map((a, i) => {
    let opts = '<option value="">— unassigned —</option>' +
      team.map(m => '<option' + (a.assignee === m.name ? ' selected' : '') + '>' + esc(m.name) + '</option>').join('');
    if (a.assignee && names.indexOf(a.assignee) === -1) opts += '<option selected>' + esc(a.assignee) + '</option>';
    return '<div class="ma-row' + (a.done ? ' done' : '') + '" data-mtg-action-row>' +
      '<label class="ma-check" title="Mark done"><input type="checkbox" data-ma-done data-change="mtg-action-done"' + (a.done ? ' checked' : '') + '></label>' +
      '<input class="ma-text" data-ma-text placeholder="Action item — what needs doing" value="' + esc(a.text || '') + '">' +
      '<select class="ma-assignee" data-ma-assignee title="Assign to">' + opts + '</select>' +
      '<input class="ma-due" type="date" data-ma-due value="' + esc(a.due || '') + '" title="Due date">' +
      '<button class="btn-icon" data-action="mtg-action-del" data-i="' + i + '" title="Remove">✕</button>' +
    '</div>';
  }).join('');
  return '<div class="ma-wrap">' +
    '<div class="ma-head">📝 <b>Action items</b>' +
      '<span class="ma-count muted">' + (actions.length ? open + ' open · ' + actions.length + ' total' : 'assign tasks to your team') + '</span>' +
    '</div>' +
    (actions.length
      ? '<div class="ma-list">' +
          '<div class="ma-row ma-hdr"><span></span><span>Task</span><span>Assigned to</span><span>Due</span><span></span></div>' +
          rows +
        '</div>'
      : '<div class="muted" style="font-size:12.5px;margin:2px 0 8px">No action items yet — add tasks and assign each to a team member.</div>') +
    '<button class="btn small" data-action="mtg-action-add">+ Add action item</button>' +
    (team.length ? '' : '<span class="muted" style="font-size:12px;margin-left:8px">Add staff under <b>Templates → Team</b> to assign people.</span>') +
  '</div>';
}
function collectMtgActions() {
  return $$('[data-mtg-action-row]').map(row => {
    const q = s => row.querySelector(s);
    return {
      text: (q('[data-ma-text]') ? q('[data-ma-text]').value : '').trim(),
      assignee: q('[data-ma-assignee]') ? q('[data-ma-assignee]').value : '',
      due: q('[data-ma-due]') ? q('[data-ma-due]').value : '',
      done: !!(q('[data-ma-done]') && q('[data-ma-done]').checked)
    };
  });
}
function collectMeetings() {
  const m = state.meetingsDraft || (state.meetingsDraft = { pre: {}, post: {}, notes: '' });
  if (!document.getElementById('mtgpreHeld')) return m;
  const g = id => { const el = document.getElementById(id); return el ? el.value : ''; };
  const c = id => { const el = document.getElementById(id); return el ? el.checked : false; };
  m.pre = { held: c('mtgpreHeld'), date: g('mtgpreDate'), notes: g('mtgpreNotes'), actions: collectMtgActions() };
  m.post = { held: c('mtgpostHeld'), date: g('mtgpostDate'), notes: g('mtgpostNotes') };
  m.notes = g('mtgNotes');
  ['pre', 'post'].forEach(k => { if (m[k].held && !m[k].date) m[k].date = new Date().toISOString().slice(0, 10); });
  return m;
}
async function saveMeetings(silent) {
  const m = collectMeetings();
  if (!m || !state.open) return;
  const j = state.detail.job;
  try {
    const preOut = Object.assign({}, m.pre, {
      actions: (m.pre.actions || []).filter(a => (a.text && a.text.trim()) || a.assignee)
    });
    const r = await api('PUT', '/api/job/' + encodeURIComponent(state.open) + '/meetings',
      { pre: preOut, post: m.post, notes: m.notes, job: { hwi: j.hwi, customer: j.customer, po: j.reference } });
    state.meetingsDraft = JSON.parse(JSON.stringify(r.meetings));
    state.detail.meetings = r.meetings;
    const row = state.jobs.find(x => x.key === state.open);
    if (row) { row.preHeld = !!(r.meetings.pre && r.meetings.pre.held); row.postHeld = !!(r.meetings.post && r.meetings.post.held); }
    // Refresh the post-job to-do list and update the dashboard side panel live (it sits behind the modal).
    try {
      const t = await api('GET', '/api/meetings/todo'); state.mtgTodo = t.todo || [];
      const side = document.querySelector('.dash-side'); if (side) side.outerHTML = postJobSideHtml();
    } catch (e) {}
    renderModal();
    if (!silent) toast('Meetings saved.');
  } catch (e) { toast(e.message, true); }
}

// ---- Post-job meetings still owed (dashboard panel) ----
function mtgTodoOverlayHtml() {
  const list = state.mtgTodo || [];
  const rows = list.length ? list.map(t =>
    '<div class="rm-row">' +
      '<div class="rm-info"><b>' + esc(t.hwi || '(no HWI)') + '</b> · ' + esc(t.customer || '(no customer)') +
        '<div class="muted" style="font-size:12px">PO ' + esc(t.po || '—') + (t.preDate ? ' · pre-job meeting ' + esc(t.preDate) : '') + '</div></div>' +
      '<button class="btn small" data-action="mtg-todo-open" data-key="' + esc(t.key) + '">Open</button>' +
    '</div>'
  ).join('') : '<div class="muted" style="padding:14px 2px">🎉 No post-job meetings outstanding.<br>A job lands here once its pre-job meeting is held but the post-job one isn’t.</div>';
  return '<div class="overlay" data-action="mtg-todo-bg">' +
    '<div class="dialog" style="max-width:560px">' +
      '<div class="dialog-head"><div><h2>📋 Post-job meetings to do</h2>' +
        '<div class="sub">' + list.length + ' job' + (list.length === 1 ? '' : 's') + ' awaiting a post-job meeting</div></div>' +
        '<button class="btn-icon dialog-close" data-action="mtg-todo-close" title="Close">✕</button></div>' +
      '<div class="dialog-body">' + rows + '</div>' +
    '</div>' +
  '</div>';
}
async function openMtgTodo() {
  try { const r = await api('GET', '/api/meetings/todo'); state.mtgTodo = r.todo || []; } catch (e) { toast(e.message, true); return; }
  let host = document.getElementById('mtgTodoModal');
  if (!host) { host = document.createElement('div'); host.id = 'mtgTodoModal'; document.body.appendChild(host); }
  host.innerHTML = mtgTodoOverlayHtml();
}
function closeMtgTodo() { const h = document.getElementById('mtgTodoModal'); if (h) h.innerHTML = ''; }

// Persistent side panel on the dashboard: Service jobs whose pre-job meeting is
// held but the post-job (R&R) one isn't. Clicking an item opens that job's Meetings tab.
function postJobSideHtml() {
  const list = state.mtgTodo || [];
  const items = list.length
    ? list.map(t =>
        '<button class="pjt-item" data-action="mtg-todo-open" data-key="' + esc(t.key) + '" title="Open ' + esc(t.hwi || '') + ' — record the post-job meeting">' +
          '<span class="pjt-hwi">' + esc(t.hwi || '(no HWI)') + '</span>' +
          '<span class="pjt-cust">' + esc(t.customer || '(no customer)') + '</span>' +
          (t.preDate ? '<span class="pjt-date">pre-job meeting ' + esc(t.preDate) + '</span>' : '') +
        '</button>').join('')
    : '<div class="pjt-empty">🎉 All caught up — no post-job meetings outstanding.</div>';
  return '<aside class="dash-side"><div class="pjt-panel">' +
    '<div class="pjt-head">📋 Post-job to-do' + (list.length ? ' <span class="todo-count">' + list.length + '</span>' : '') + '</div>' +
    '<p class="pjt-sub">Service jobs whose pre-job meeting is done but the post-job (R&R) one isn’t. Click one to record it.</p>' +
    '<div class="pjt-list">' + items + '</div>' +
  '</div></aside>';
}

// ---- Meeting report: email the action items to the team (everyone EXCEPT Mike Scofield),
//      sent AS the logged-in user via Microsoft Graph (same path as "Send to customer"). ----
var MTG_REPORT_EXCLUDE = /^(mike|michael)\s+scofield$/i;
function isEmailAddr(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim()); }
function meetingReportRecipients() {
  const team = (state.templates && state.templates.team) || [];
  return team
    .filter(m => m.name)
    .map(m => ({
      name: m.name,
      email: isEmailAddr(m.email) ? m.email.trim() : (isEmailAddr(m.contact) ? m.contact.trim() : ''),
      checked: !MTG_REPORT_EXCLUDE.test(String(m.name).trim())   // everyone ticked by default except Mike Scofield
    }));
}
function buildMeetingReport(d) {
  const j = d.job;
  const m = state.meetingsDraft || {};
  const pre = m.pre || {}, post = m.post || {};
  const actions = pre.actions || [];
  const sender = signOff();
  const ref = j.hwi || j.number || '';
  const heading = (j.customer || 'Job') + (ref ? ' — ' + ref : '');
  const subject = 'Pre-job meeting report — ' + heading;
  const open = actions.filter(a => !a.done).length;
  const base = 'padding:6px 9px;border:1px solid #dfe6ec;font-size:13px;vertical-align:top';
  const thS = 'padding:6px 9px;border:1px solid #dfe6ec;background:#f2f6f9;text-align:left;font-size:12px;color:#5a6b7b';
  const rows = actions.length
    ? actions.map((a, i) => {
        const taskS = base + (a.done ? ';text-decoration:line-through;color:#8a97a5' : '');
        return '<tr>' +
          '<td style="' + base + '">' + (i + 1) + '</td>' +
          '<td style="' + taskS + '">' + esc(a.text || '(no description)') + '</td>' +
          '<td style="' + base + '"><b>' + esc(a.assignee || 'Unassigned') + '</b></td>' +
          '<td style="' + base + '">' + esc(a.due || '—') + '</td>' +
          '<td style="' + base + '">' + (a.done ? '✓ Done' : 'Open') + '</td>' +
        '</tr>';
      }).join('')
    : '<tr><td style="' + base + '" colspan="5"><i>No action items were recorded for this meeting.</i></td></tr>';
  const statusLine = 'Pre-job meeting: ' + (pre.held ? 'held' + (pre.date ? ' on ' + esc(pre.date) : '') : 'not yet held') +
    ' &nbsp;·&nbsp; Post-job meeting: ' + (post.held ? 'held' + (post.date ? ' on ' + esc(post.date) : '') : 'not yet held');
  const html = '<div style="font:14px/1.55 Arial,Segoe UI,sans-serif;color:#1a2733">' +
    '<h2 style="margin:0 0 2px;font-size:18px">Pre-job meeting report</h2>' +
    '<div style="color:#5a6b7b;margin-bottom:10px">' + esc(heading) + (j.reference ? ' &nbsp;·&nbsp; PO ' + esc(j.reference) : '') + '</div>' +
    '<p style="margin:6px 0">' + statusLine + '</p>' +
    '<p style="margin:14px 0 6px"><b>Action items</b> — ' + open + ' open of ' + actions.length + '</p>' +
    '<table style="border-collapse:collapse;border:1px solid #dfe6ec;width:100%">' +
      '<thead><tr><th style="' + thS + '">#</th><th style="' + thS + '">Task</th><th style="' + thS + '">Assigned to</th><th style="' + thS + '">Due</th><th style="' + thS + '">Status</th></tr></thead>' +
      '<tbody>' + rows + '</tbody></table>' +
    (pre.notes ? '<p style="margin:14px 0 4px"><b>Meeting notes</b></p><div style="white-space:pre-wrap;background:#f7fafd;border:1px solid #e3ebf1;border-radius:6px;padding:8px 10px;font-size:13px">' + esc(pre.notes) + '</div>' : '') +
    '<p style="color:#8a97a5;font-size:12px;margin-top:16px">Sent from the Hydro-Wates Project Manager by ' + esc(sender) + '.</p>' +
  '</div>';
  const lines = actions.length
    ? actions.map((a, i) => (i + 1) + '. ' + (a.text || '(no description)') + ' — ' + (a.assignee || 'Unassigned') + (a.due ? ' (due ' + a.due + ')' : '') + (a.done ? ' [DONE]' : '')).join('\n')
    : 'No action items were recorded for this meeting.';
  const text = 'Pre-job meeting report\n' + heading + (j.reference ? ' · PO ' + j.reference : '') + '\n\n' +
    statusLine.replace(/&nbsp;·&nbsp;/g, ' · ') + '\n\nAction items (' + open + ' open of ' + actions.length + '):\n' + lines +
    (pre.notes ? '\n\nMeeting notes:\n' + pre.notes : '') + '\n\nSent from the Hydro-Wates Project Manager.';
  return { subject, html, text };
}
function mtgReportOverlayHtml(d) {
  const recips = state.mtgReportRecipients || [];
  const m = state.meetingsDraft || {};
  const actions = (m.pre && m.pre.actions) || [];
  const open = actions.filter(a => !a.done).length;
  const ready = emailReady();
  const missing = recips.filter(r => r.checked && !r.email).length;
  const recipRows = recips.length
    ? recips.map(r =>
        '<div class="rpt-recip">' +
          '<label class="rpt-check" title="Include ' + esc(r.name) + '"><input type="checkbox" data-rpt-pick' + (r.checked ? ' checked' : '') + '></label>' +
          '<span class="rpt-name">' + esc(r.name) + '</span>' +
          '<input data-rpt-email type="email" placeholder="name@hydrowates.com" value="' + esc(r.email || '') + '">' +
        '</div>').join('')
    : '<div class="muted">No team members found. Add staff under Templates → Team.</div>';
  const itemRows = actions.length
    ? actions.map(a => '<li>' + esc(a.text || '(no description)') + ' — <b>' + esc(a.assignee || 'Unassigned') + '</b>' + (a.due ? ' · due ' + esc(a.due) : '') + (a.done ? ' · <span style="color:#1a7f45">done</span>' : '') + '</li>').join('')
    : '<li class="muted">No action items recorded yet.</li>';
  return '<div class="overlay" data-action="mtg-report-bg">' +
    '<div class="dialog" style="max-width:600px">' +
      '<div class="dialog-head"><div><h2>✉ Send meeting report</h2>' +
        '<div class="sub">' + esc((d.job.customer || '') + (d.job.hwi ? ' · ' + d.job.hwi : '')) + '</div></div>' +
        '<button class="btn-icon dialog-close" data-action="mtg-report-close" title="Close">✕</button></div>' +
      '<div class="dialog-body">' +
        (ready ? '' : '<div class="banner" style="background:#eef4fb;border:1px solid #cfe0f2;color:#25507d">ℹ️ When you press Send, a quick Microsoft popup confirms it’s you (usually no password needed) — then the report goes out from your account.</div>') +
        '<p class="hint" style="margin:2px 0 8px">Sent from <b>your</b> account to everyone ticked below. <b>Everyone is selected except Mike Scofield</b> — tick or untick anyone.' + (missing ? ' <b>' + missing + '</b> selected still need an email.' : '') + '</p>' +
        '<label class="chk rpt-selall-row"><input type="checkbox" id="rptSelAll" data-change="rpt-selall"> Select / clear all</label>' +
        '<div class="rpt-recips">' + recipRows + '</div>' +
        '<p style="margin:12px 0 4px"><b>Report contents</b> — ' + actions.length + ' action item' + (actions.length === 1 ? '' : 's') + ' (' + open + ' open)</p>' +
        '<ul class="rpt-items">' + itemRows + '</ul>' +
        '<div class="plan-actions" style="margin-top:12px">' +
          '<button class="btn primary" data-action="mtg-report-send">✈ Send report</button>' +
          '<button class="btn" data-action="mtg-report-close">Cancel</button>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div>';
}
function openMtgReport() {
  let host = document.getElementById('mtgReportModal');
  if (!host) { host = document.createElement('div'); host.id = 'mtgReportModal'; document.body.appendChild(host); }
  host.innerHTML = mtgReportOverlayHtml(state.detail);
}
function closeMtgReport() { const h = document.getElementById('mtgReportModal'); if (h) h.innerHTML = ''; }

function renderModal() {
  const d = state.detail;
  if (!d) { $('#modal').innerHTML = ''; return; }
  const j = d.job;
  const mtg = d.meetings || {};
  const preH = !!(mtg.pre && mtg.pre.held), postH = !!(mtg.post && mtg.post.held);
  $('#modal').innerHTML =
    '<div class="overlay" data-action="overlay-close">' +
      '<div class="dialog">' +
        '<div class="dialog-head">' +
          '<div><h2>' + esc(j.customer || '(no customer)') + '</h2>' +
          '<div class="sub">' + esc(MOD_NAME[j.module] || j.module) + ' ' + esc(j.number) + ' · ' + esc(fmtMoney(j.total, j.currency)) + (j.demo ? ' · demo' : '') + '</div></div>' +
          '<button class="btn-icon dialog-close" data-action="close-modal" title="Close">✕</button>' +
        '</div>' +
        '<div class="tabs">' +
          '<button class="tab' + (state.modalTab === 'details' ? ' active' : '') + '" data-action="modal-tab" data-tab="details">Details</button>' +
          '<button class="tab' + (state.modalTab === 'planning' ? ' active' : '') + '" data-action="modal-tab" data-tab="planning">Planning</button>' +
          '<button class="tab' + (state.modalTab === 'procedure' ? ' active' : '') + '" data-action="modal-tab" data-tab="procedure">Procedure</button>' +
          '<button class="tab' + (state.modalTab === 'travel' ? ' active' : '') + '" data-action="modal-tab" data-tab="travel">Travel' +
            (j.travelMode ? ' <span class="tab-dot">' + (j.travelMode === 'fly' ? '✈' : '🚗') + '</span>' : '') + '</button>' +
          (j.category === 'service'
            ? '<button class="tab' + (state.modalTab === 'meetings' ? ' active' : '') + '" data-action="modal-tab" data-tab="meetings">Meetings' +
                (preH || postH ? ' <span class="tab-dot">' + (preH && postH ? '✓✓' : '✓') + '</span>' : '') + '</button>'
            : '') +
        '</div>' +
        '<div class="dialog-body">' +
          (state.modalTab === 'planning' ? planningTabHtml(d)
            : state.modalTab === 'procedure' ? procedureTabHtml(d)
            : state.modalTab === 'travel' ? travelTabHtml(d)
            : (state.modalTab === 'meetings' && j.category === 'service') ? meetingsTabHtml(d)
            : detailTabHtml(d)) +
        '</div>' +
      '</div>' +
    '</div>';
}

/* ----- planning helpers ----- */
function collectPlanning() {
  const p = state.planningDraft;
  if (!p) return null;
  const rows = $$('[data-qrow]');
  if (rows.length || $('#planEmail')) {
    const qs = rows.map((row, i) => {
      const cur = p.questions[i] || {};
      const valEl = row.querySelector('[data-qval]');
      const q = {
        id: cur.id || ('q' + Math.random().toString(36).slice(2, 8)),
        text: row.querySelector('[data-qtext]').value,
        type: cur.type || 'text',
        value: valEl ? valEl.value : ''
      };
      if (q.type === 'number') {
        if (cur.units) q.units = cur.units;
        const us = row.querySelector('[data-qunit]');
        q.unit = us ? us.value : (cur.unit || '');
      }
      if (q.type === 'choice') q.options = cur.options || [];
      q.answer = answerDisplay(q);
      return q;
    }).filter(q => q.text.trim());
    p.questions = qs;
    if ($('#planEmail')) p.email = $('#planEmail').value.trim();
    if ($('#planStatus')) p.status = $('#planStatus').value;
    if ($('#planNotes')) p.notes = $('#planNotes').value;
  }
  const boxes = $$('[data-contact]');
  if (boxes.length) p.recipients = boxes.filter(b => b.checked).map(b => b.value);
  return p;
}

function plannedRecipients() {
  const p = state.planningDraft;
  const to = (p.recipients || []).slice();
  const extra = (p.email || '').trim();
  if (extra && !to.includes(extra)) to.push(extra);
  return to;
}

async function savePlanning(silent) {
  const p = collectPlanning();
  if (!p || !state.open) return;
  try {
    const r = await api('PUT', '/api/job/' + encodeURIComponent(state.open) + '/planning', p);
    state.planningDraft = JSON.parse(JSON.stringify(r.planning));
    state.detail.planning = r.planning;
    const row = state.jobs.find(x => x.key === state.open);
    if (row) row.planningStatus = r.planning.status;
    if (!silent) toast('Planning saved.');
  } catch (e) { toast(e.message, true); }
}

function buildEmail() {
  const p = state.planningDraft, t = state.templates, j = state.detail.job;
  const fill = (s) => String(s || '')
    .split('{job}').join(j.number || 'upcoming job')
    .split('{customer}').join(j.customer || '');
  const numbered = (p.questions || []).map((q, i) => (i + 1) + ') ' + q.text + answerHint(q)).join('\n\n');
  const subject = fill(t.emailSubject);
  const body = fill(t.emailIntro) + '\n\n' + numbered + '\n\n' + fill(t.emailOutro) + '\n' + signOff();
  return { subject, body };
}

// Proof-test load shown on the job: WLL × 1.25. WLL is entered internally (not asked of the customer).
function testLoadText(j) {
  if (j.wll === null || j.wll === undefined || j.wll === '') return '<span class="muted">— enter WLL —</span>';
  const wll = Number(j.wll), u = j.wllUnit || 't';
  const tl = +(wll * 1.25).toFixed(3);
  return '<b>' + tl + ' ' + esc(u) + '</b> <span class="muted">(' + wll + ' × 1.25)</span>';
}

// Open a clean, printable Q&A sheet for the job pack (the customer's planning answers).
function printQuestionnaire() {
  const p = collectPlanning(); const j = state.detail.job;
  if (!p || !(p.questions || []).length) { toast('Load the questions first.', true); return; }
  const win = window.open('', '_blank');
  if (!win) { toast('Allow pop-ups for this site to print.', true); return; }
  const tl = (j.wll !== null && j.wll !== undefined && j.wll !== '')
    ? (+(Number(j.wll) * 1.25).toFixed(3)) + ' ' + (j.wllUnit || 't') : null;
  const rows = p.questions.map((q, i) => {
    const a = answerDisplay(q);
    return '<li><div class="q">' + (i + 1) + '. ' + esc(q.text) + '</div>' +
      '<div class="a">' + (a ? esc(a) : '<span class="na">— not answered —</span>') + '</div></li>';
  }).join('');
  const meta = [
    j.number ? 'Job ' + esc(j.number) : '',
    j.reference ? 'PO ' + esc(j.reference) : '',
    j.wll != null && j.wll !== '' ? 'WLL ' + esc(j.wll) + ' ' + esc(j.wllUnit || 't') : '',
    tl ? 'Test load ' + esc(tl) + ' (125%)' : ''
  ].filter(Boolean).join('  ·  ');
  const notes = (p.notes || '').trim();
  const letterhead =
    '<div class="lh">' +
      '<div class="lh-brand">' +
        '<img class="lh-logo" src="/logo.jpg" alt="Hydro-Wates — Proof-Load Testing Services" ' +
          'onerror="this.style.display=\'none\';var f=this.nextElementSibling;if(f)f.style.display=\'block\'">' +
        '<div class="lh-fallback" style="display:none">' +
          '<div class="lh-name">Hydro-Wates</div><div class="lh-tag">Proof-Load Testing Services</div>' +
        '</div>' +
      '</div>' +
      '<div class="lh-doc">Planning questionnaire</div>' +
    '</div>';
  const html =
    '<!doctype html><html><head><meta charset="utf-8"><title>Planning — ' + esc(j.customer || '') + '</title>' +
    '<style>' +
    'body{font:14px/1.5 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1a2733;max-width:720px;margin:32px auto;padding:0 24px}' +
    '.lh{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;border-bottom:3px solid #0b5e8a;padding-bottom:10px;margin-bottom:16px}' +
    '.lh-logo{max-height:62px;max-width:320px;display:block}' +
    '.lh-name{font-size:22px;font-weight:700;letter-spacing:.5px;color:#2b5c9e}.lh-tag{font-size:13px;font-weight:700;color:#1a2733}' +
    '.lh-doc{font-size:13px;color:#5b6770;text-transform:uppercase;letter-spacing:.08em;white-space:nowrap}' +
    'h1{font-size:20px;margin:0 0 2px}.meta{color:#5b6770;font-size:13px;margin:0 0 18px}' +
    'ol{list-style:none;padding:0;margin:0}li{padding:10px 0;border-bottom:1px solid #e9eef2}' +
    '.q{font-weight:600;margin-bottom:3px}.a{white-space:pre-wrap}.na{color:#9aa6b1}' +
    '.notes{margin-top:18px;padding:12px 14px;background:#f6f8fa;border:1px solid #e3e9ee;border-radius:6px}' +
    '.notes-h{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:#5b6770;margin-bottom:4px}.notes-b{white-space:pre-wrap}' +
    '.foot{margin-top:22px;color:#9aa6b1;font-size:12px}@media print{body{margin:0}}' +
    '</style></head><body>' +
    letterhead +
    '<h1>' + esc(j.customer || '(no customer)') + '</h1>' +
    (meta ? '<p class="meta">' + meta + '</p>' : '') +
    '<ol>' + rows + '</ol>' +
    (notes ? '<div class="notes"><div class="notes-h">Internal notes</div><div class="notes-b">' + esc(notes) + '</div></div>' : '') +
    '<p class="foot">Hydro-Wates Project Manager · ' + esc(new Date().toLocaleString()) + '</p>' +
    '</body></html>';
  win.document.write(html);
  win.document.close();
  win.focus();
  win.onload = () => win.print();
}

/* ---------------- procedure (load-test procedure generator) ---------------- */
const PROC_COMPANY = {
  tagline: 'Providing Proof-Load Testing Services to the Maritime, Petroleum, & Heavy Construction Industries — Worldwide',
  address: '8100 Lockheed Avenue, Houston, Texas 77061',
  tel: '(713) 643-9990'
};
const PROC_STATUS = { draft: 'Draft', reviewed: 'Reviewed', final: 'Final' };
const PROC_LOGO_URL = 'https://www.hydrowates.com/assets/images/template/hydro-logo.jpg';
const PROC_PREJOB =
  'Conduct a pre-job safety briefing with all Hydro-Wates and assisting personnel. The following topics must be reviewed:\n' +
  '1. Overview of lifting and load testing activities.\n' +
  '2. Equipment to be used for load testing.\n' +
  '3. Sequence of events and step-by-step testing procedure.\n' +
  '4. Safety measures and Job Safety Analysis (JSA) action items.\n' +
  '5. Personnel assignments, including responsibilities, communication methods, PPE requirements, and emergency plans.\n\n' +
  'Safety topics must include: dropped-object hazards, identification of equipment Working Load Limits (WLL), proper use of tag lines, ' +
  'crane operator communication, and maintaining a minimum 10 ft distance from suspended loads. Address any concurrent operations on ' +
  'site and coordinate with operation leaders. Allow an open forum for personnel to identify and discuss potential hazards before ' +
  'proceeding. All concerns must be resolved prior to beginning testing.';
const PROC_LOGGING =
  'Record all relevant data, including fill times, static hold duration, and any observations. Maintain manual notes as a backup in the ' +
  'event of electronic data loss (dongle error).';
// The PPE checklist shown on every procedure. Tick what applies per job; add options here.
const PPE_OPTIONS = [
  'Hard hat',
  'Safety glasses',
  'Steel-toe safety boots',
  'High-visibility vest',
  'Gloves',
  'Hearing protection',
  'Fall-arrest harness',
  'FR coveralls',
  'Respirator / dust mask'
];
// Ticked by default on a new procedure (the rest are one click away in the checklist).
const PROC_PPE = ['Hard hat', 'Safety glasses', 'Steel-toe safety boots'];

// Render the PPE picker: each standard option as a checkbox (ticked if already on the
// procedure), plus a small "Other" box for anything custom not in the standard list.
function ppeChecklistHtml(p) {
  // Tick the standard defaults until this procedure has its own PPE selection
  // (so older procedures that predate the checklist still show sensible defaults).
  const sel = (Array.isArray(p.ppe) && p.ppe.length) ? p.ppe : PROC_PPE.slice();
  const boxes = PPE_OPTIONS.map(o =>
    '<label class="chk ppe-opt"><input type="checkbox" data-ppe value="' + esc(o) + '"' +
    (sel.includes(o) ? ' checked' : '') + '> ' + esc(o) + '</label>'
  ).join('');
  const extra = sel.filter(x => !PPE_OPTIONS.includes(x));
  return '<div class="ppe-grid">' + boxes + '</div>' +
    '<label class="ppe-other-lbl">Other <span class="muted">(one per line)</span></label>' +
    '<textarea id="procPPEOther" rows="' + Math.max(1, extra.length) +
      '" placeholder="Any PPE not in the list above">' + esc(extra.join('\n')) + '</textarea>';
}

function fmtBytes(n) {
  n = Number(n) || 0;
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
  return (n / 1024 / 1024).toFixed(1) + ' MB';
}
// Rigging-drawing PDFs attached to the procedure (stored in Supabase Storage).
function drawingsTabHtml(p) {
  const list = (p && p.drawings) || [];
  const rows = list.map(d =>
    '<div class="draw-row">' +
      '<span class="draw-ic">📄</span>' +
      '<span class="draw-name" title="' + esc(d.name) + '">' + esc(d.name) + '</span>' +
      '<span class="muted draw-size">' + fmtBytes(d.size) + '</span>' +
      '<button type="button" class="btn small" data-action="draw-view" data-id="' + esc(d.id) + '">View</button>' +
      '<button type="button" class="btn small ghost" data-action="draw-remove" data-id="' + esc(d.id) + '">Remove</button>' +
    '</div>'
  ).join('');
  return '<div class="draw-box">' +
    (rows || '<div class="muted" style="font-size:13px;padding:2px 0 6px">No drawings attached yet.</div>') +
    '<div class="draw-add">' +
      '<label class="btn small">➕ Add PDF drawing' +
        '<input type="file" accept="application/pdf,.pdf" data-change="draw-file" style="display:none"></label>' +
      '<span class="muted" style="font-size:12px">PDF only, up to 4 MB each.</span>' +
    '</div>' +
  '</div>';
}

// Deep link to the Hydro-Wates Rigging Planner (our own Vercel app). ?job=<HWI>
// opens that job's plan or creates a new one. Update this URL if the planner is renamed.
const RIGGING_PLANNER_URL = 'https://rigging-planner-gamma.vercel.app/';
function riggingPlannerLinkHtml(j) {
  const hwi = (j && j.hwi) ? String(j.hwi) : '';
  const url = RIGGING_PLANNER_URL + (hwi ? '?job=' + encodeURIComponent(hwi) : '');
  return '<div class="rig-plan">' +
    '<a class="btn small" href="' + esc(url) + '" target="_blank" rel="noopener">🏗 Open in Rigging Planner</a>' +
    '<span class="muted" style="font-size:12px">' +
      (hwi ? 'Opens <b>' + esc(hwi) + '</b>’s plan (or starts a new one)' : 'Add an HWI to this job to link its plan') +
      ' — then print → Save as PDF and drop it below.</span>' +
  '</div>';
}

// Upload a PDF drawing to the current job's procedure — shared by the file picker and drag-drop.
async function uploadDrawingFile(file) {
  if (!file || !state.open) return;
  if (file.type !== 'application/pdf' && !/\.pdf$/i.test(file.name)) { toast('Please choose a PDF file.', true); return; }
  if (file.size > 4 * 1024 * 1024) { toast('That PDF is over the 4 MB limit — please attach a smaller file.', true); return; }
  if ($('#procObjective')) collectProcedure();
  toast('Uploading ' + file.name + '…');
  try {
    const opt = { method: 'POST', headers: { 'Content-Type': 'application/pdf', 'x-filename': file.name }, body: file };
    if (authToken) opt.headers['Authorization'] = 'Bearer ' + authToken;
    const resp = await fetch('/api/job/' + encodeURIComponent(state.open) + '/drawings', opt);
    const j = await resp.json().catch(() => ({}));
    if (resp.status === 401) { renderLogin('Your session has expired — please sign in again.'); return; }
    if (!resp.ok) throw new Error(j.error || ('Upload failed (' + resp.status + ')'));
    if (state.procedureDraft) state.procedureDraft.drawings = j.drawings;
    renderModal();
    toast('Attached ' + file.name + '.');
  } catch (err) { toast(err.message || 'Upload failed.', true); }
}

// Site photos attached to the procedure. Thumbnails need a signed URL (server adds
// `url` to each on load / upload); clicking a thumb opens the full-size image.
function photosTabHtml(p) {
  const list = (p && p.photos) || [];
  const thumbs = list.map(ph =>
    '<div class="photo-thumb" data-action="photo-view" data-id="' + esc(ph.id) + '" title="' + esc(ph.name) + '">' +
      (ph.url ? '<img src="' + esc(ph.url) + '" alt="' + esc(ph.name) + '">' : '<div class="photo-broken">image</div>') +
      '<button type="button" class="photo-x" data-action="photo-remove" data-id="' + esc(ph.id) + '" title="Remove">✕</button>' +
    '</div>'
  ).join('');
  return '<div class="photo-box">' +
    (thumbs ? '<div class="photo-grid">' + thumbs + '</div>' : '<div class="muted" style="font-size:13px;padding:2px 0 6px">No photos attached yet.</div>') +
    '<div class="photo-add">' +
      '<label class="btn small">➕ Add photo' +
        '<input type="file" accept="image/*" data-change="photo-file" style="display:none"></label>' +
      '<span class="muted" style="font-size:12px">JPG/PNG — resized automatically to keep it small.</span>' +
    '</div>' +
  '</div>';
}
// Downscale + re-encode an image in the browser so uploads (and storage) stay small.
function downscaleImage(file, maxDim, quality) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      let w = img.naturalWidth || img.width, h = img.naturalHeight || img.height;
      const scale = Math.min(1, maxDim / Math.max(w, h));
      w = Math.max(1, Math.round(w * scale)); h = Math.max(1, Math.round(h * scale));
      const c = document.createElement('canvas'); c.width = w; c.height = h;
      c.getContext('2d').drawImage(img, 0, 0, w, h);
      c.toBlob(b => b ? resolve(b) : reject(new Error('Could not process that image.')), 'image/jpeg', quality || 0.82);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Could not read that image.')); };
    img.src = url;
  });
}

// Read a structured planning answer by question id (drives the procedure cascade).
function planAns(planning, id) {
  if (!planning || !planning.questions) return '';
  const q = planning.questions.find(x => x.id === id);
  return q && q.value != null ? String(q.value).trim() : '';
}

// Build a full draft procedure from the job + its WLL/test load + planning answers + standard boilerplate.
const JOB_TYPES = [['waterbags', '💧 Water bags'], ['steelweights', '🏋 Steel weights'], ['padeye', '⚓ Pad eye testing']];

// The answers that drive procedure generation. Pre-fills a couple of fields from the
// planning answers when they exist.
function defaultSetup(j, planning) {
  const q4 = (planning && planning.questions) ? planning.questions.find(x => x.id === 'q4') : null;
  const dist = q4 && q4.value != null ? String(q4.value).trim() : '';
  const q7 = planAns(planning, 'q7');
  return {
    jobType: 'waterbags',
    testPct: '125',          // % of WLL to test to (100 for a proof-to-100% job, etc.)
    testLoad: '',            // blank = auto (WLL × testPct)
    functionTest: q7 ? q7 === 'yes' : true,
    staticHold: true,
    holdMinutes: '10',
    bagCount: '', bagSize: '',
    waterSource: '', waterDistance: dist,
    weightConfig: '', landingArea: '',
    padEyeCount: '', pullDirection: '', tester: '', ramArea: ''
  };
}

// Deterministic, no-AI procedure builder: branches on the job type from the setup
// answers and assembles the objective, equipment, and steps. Same answers -> same doc.
function generateProcedure(j, planning, setup) {
  setup = setup || defaultSetup(j, planning);
  const shortTons = /short/i.test(planAns(planning, 'q2'));
  const unitWord = shortTons ? 'short tons' : 'metric tons';
  const factor = shortTons ? 2000 : 2204.62;
  const wll = (j.wll != null && j.wll !== '') ? Number(j.wll) : null;
  const pctRaw = (setup.testPct != null && String(setup.testPct).trim() !== '') ? Number(setup.testPct) : 125;
  const pct = isNaN(pctRaw) ? 125 : pctRaw;
  const ov = (setup.testLoad != null && String(setup.testLoad).trim() !== '') ? Number(setup.testLoad) : null;
  const testLoad = (ov != null && !isNaN(ov)) ? ov : (wll != null ? +(wll * pct / 100).toFixed(3) : null);
  const lbs = n => Math.round(n).toLocaleString();
  const pctWord = pct + '%';
  const loadWord = testLoad != null ? (testLoad + ' ' + unitWord + ' (' + lbs(testLoad * factor) + ' lbs, ' + pctWord + ' WLL)') : (pctWord + ' of WLL');
  const holdMin = String(setup.holdMinutes || '').trim();
  const x = {
    wll: wll, testLoad: testLoad, unitWord: unitWord, factor: factor, lbs: lbs, loadWord: loadWord, pctWord: pctWord,
    fnTest: !!setup.functionTest,
    cust: j.customer || 'the customer',
    holdStr: setup.staticHold
      ? ('Hold the test load' + (holdMin ? ' for ' + holdMin + ' minutes' : ' for the required duration') + ' as a static proof hold, monitoring for any movement, slippage, or deformation.')
      : null
  };
  const t = setup.jobType || 'waterbags';
  const build = t === 'steelweights' ? genSteelWeights(j, setup, x)
              : t === 'padeye' ? genPadEye(j, setup, x)
              : genWaterBags(j, setup, x);
  return {
    status: 'draft',
    startDate: j.date ? fmtDate(j.date) : '',
    scope: '',
    jobSite: '',
    projectRef: j.number || '',
    objective: build.objective,
    coordination: '',
    responsibilities: [],
    equipment: build.equipment,
    ppe: ((state.templates && state.templates.ppe) || PROC_PPE).slice(),
    preJob: PROC_PREJOB,
    setupSteps: build.setupSteps,
    executionSteps: build.executionSteps,
    logging: PROC_LOGGING,
    setup: Object.assign({}, setup),
    approvedBy: '',
    approvedDate: ''
  };
}

function genWaterBags(j, s, x) {
  const n = String(s.bagCount || '').trim(), size = String(s.bagSize || '').trim();
  const bagDesc = (n ? n + ' × ' : '') + (size ? size + ' ' : '') + 'water-weight bag' + ((n && Number(n) === 1) ? '' : 's');
  const src = String(s.waterSource || '').trim();
  const dist = String(s.waterDistance || '').trim();
  const distStr = dist ? (/[a-z]/i.test(dist) ? dist : dist + ' ft') : '';
  const srcStr = (src || 'the water source') + (distStr ? ' (' + distStr + ' away)' : '');
  return {
    objective: 'Perform a proof load test to ' + x.loadWord + ' using ' + bagDesc + '.',
    equipment: ((state.templates && state.templates.equipment) || []).slice(),
    setupSteps: [
      'Mark the test area with red danger tape and establish an exclusion zone.',
      'Position the water bag(s) under the hook.',
      'Unroll the bag(s) and connect the ball valve to the drain trunk.',
      'Unroll and stage fill hoses to ' + srcStr + '.',
      'Unroll and stage drain hoses to the discharge location.',
      'Attach fill hose to the bag’s fill fitting.',
      'Connect ' + (src || 'the water source') + ' to the fill hoses.',
      'Attach master link to the hook if needed.',
      'Attach shackles to the hook or master link.',
      'Attach the tared-out load cell to the shackles.',
      'Connect the lower shackle to the large master links on the water bag.',
      'Ensure rigging shackles are properly pinned to master links.',
      'Review all rigging; check for and correct pinched slings, missing retaining pins, etc.',
      'Begin recording load data in the OSCAR app using the dongle.',
      'Slowly lift the bag approximately 1 ft off the ground and verify all rigging.'
    ],
    executionSteps: [
      'Confirm the drain ball valve is closed.',
      'Turn on the water source and monitor the load increase carefully.',
      x.wll != null ? 'Fill water bags to a total of ' + x.wll + ' ' + x.unitWord + ' (' + x.lbs(x.wll * x.factor) + ' lbs).' : 'Fill water bags to 100% of WLL.',
      x.fnTest ? 'Perform a function test as directed by ' + x.cust + '.' : null,
      'Continue filling to the test load of ' + x.loadWord + '.',
      x.holdStr,
      'After completion, connect drain hoses.',
      'Open the ball valve and drain the bag completely, monitoring the drain location and reducing flow as needed to prevent flooding.',
      'Once empty, carefully lay the water bag down.',
      'Repeat these steps for each additional unit under test.'
    ].filter(Boolean)
  };
}

function genSteelWeights(j, s, x) {
  const cfg = String(s.weightConfig || '').trim(), landing = String(s.landingArea || '').trim();
  const tmpl = (state.templates && state.templates.equipment) || [];
  return {
    objective: 'Load test the customer’s equipment to ' + x.loadWord + ' using calibrated steel test weights' + (cfg ? ' (' + cfg + ')' : '') + '.',
    equipment: tmpl.length ? tmpl.slice() : [
      '(2) 2,000 lb steel plates', '(2) 500 lb steel plates', '(10) 50 lb hand weights',
      'Test stand', '6.5 Te load cell + OSCAR dongle', '17 Te shackles', '3/4 Te shackles',
      'Wire rope slings (2 ft, 4 ft)', 'Tool bag'
    ],
    setupSteps: [
      'Conduct the pre-job safety briefing and JSA; confirm all personnel understand their assignments.',
      'Mark the test area with red danger tape and establish an exclusion zone.',
      'Confirm the ' + (landing ? landing + ' ' : '') + 'test stand / lay-down area is level and rated for the full test load.',
      'Stack the required steel plates and hand weights on the test stand to make up the test load' + (cfg ? ' (' + cfg + ')' : '') + '.',
      'Attach the shackle(s) and rig the wire rope sling(s) to the weights.',
      'Fit the tared-out load cell in line with the rigging.',
      'Review all rigging; check shackles are properly pinned and slings are not pinched.'
    ],
    executionSteps: [
      x.fnTest ? 'Perform a function test of the equipment as directed by ' + x.cust + '.' : null,
      'Have the equipment under test pick up the weights and take the load to ' + x.loadWord + ', monitoring the load cell.',
      x.holdStr || 'Hold the test load for the required duration.',
      'Log the start and stop times of the hold.',
      'Land the weights back down in a controlled manner and de-rig.',
      'Account for all test weights and record the result.',
      'Repeat for each additional unit under test — position the required plates/weights, rig, lift, and hold.'
    ].filter(Boolean)
  };
}

function genPadEye(j, s, x) {
  const cnt = String(s.padEyeCount || '').trim();
  const dir = String(s.pullDirection || '').trim();
  const ram = String(s.tester || '').trim() || 'hydraulic ram';
  const many = cnt && Number(cnt) > 1;
  const cntWord = cnt ? (cnt + ' lifting lug' + (many ? 's' : '') + ' / pad eye' + (many ? 's' : '')) : 'the lifting lug(s) / pad eye(s)';
  const area = String(s.ramArea || '').trim();
  const areaN = area ? Number(area.replace(/[^0-9.]/g, '')) : null;
  const testLbs = x.testLoad != null ? Math.round(x.testLoad * x.factor) : null;
  const psiStep = 'Calculate the target PSI for the test load:  PSI = test load (lbs) ÷ ram surface area (in²)' +
    ((areaN && testLbs) ? '  =  ' + x.lbs(testLbs) + ' ÷ ' + area + '  =  ' + Math.round(testLbs / areaN).toLocaleString() + ' psi.' : '  (enter the ram model / surface area).');
  const tmpl = (state.templates && state.templates.equipment) || [];
  return {
    objective: 'Load test ' + cntWord + ' to a test load of ' + x.loadWord + ' using a hydraulic ram and reaction frame.',
    equipment: tmpl.length ? tmpl.slice() : [
      (ram === 'hydraulic ram' ? '30 Te hydraulic ram' : ram),
      'Threaded pulling mandrel w/ washers & nuts', '(2) reaction pins', 'Hydraulic hoses (10 ft)',
      'Enerpac hydraulic hand pump w/ oil', 'Slotted I-beam', 'Assorted metal shims', 'Reaction frame',
      'Shackles (bolt-type + screw-pin, assorted)', '10,000 psi pressure gauge w/ calibration cert', 'Assorted wooden blocks'
    ],
    setupSteps: [
      'Conduct the pre-job safety briefing and JSA; review working at heights, dropped objects, and tag lines.',
      'Attach the shackle to the lifting lug / pad eye.',
      'Build the reaction frame around the shackle using the I-beam and metal shims (stack wooden blocks under the I-beam for height if needed).',
      'Set the ' + ram + ' on top of the I-beam.',
      'Insert the pulling mandrel through the ram and I-beam.',
      'Attach the pulling mandrel to the shackle and adjust by tightening the wing nut.',
      'Attach the hydraulic hose to the hand pump and the ram.',
      'Attach the calibrated pressure gauge to the hand pump.' + (dir ? ' Confirm the pull direction / angle (' + dir + ') matches the design load direction.' : '')
    ],
    executionSteps: [
      psiStep,
      'Erect red danger tape and clear the area of all personnel.',
      'Slowly pump the hand pump until 500 psi.',
      'A Hydro-Wates representative enters the test area to inspect the apparatus (ensure the clevis is not contacting the underside of the reaction frame), then returns to the safe area outside the barrier.',
      x.fnTest ? 'Perform a function test as directed by ' + x.cust + '.' : null,
      'Slowly pump until the target PSI is reached.',
      x.holdStr || 'Once the target PSI is reached, hold for 2 minutes.',
      'Bleed off pressure via the bleed-off line on the hand pump; confirm zero pressure and that the system is de-energized.',
      'Remove the test equipment and note any damage or issues.',
      'Flag the lug with pink flagging tape for the weld inspector to indicate it has been load tested.',
      'Complete a field load test report for each lug and collect the customer witness signature.',
      cnt ? 'Reposition and repeat for each of the ' + cnt + ' lugs.' : 'Reposition and repeat for each additional lug / pad eye.'
    ].filter(Boolean)
  };
}

// The "Procedure setup" question panel — pick the job type, answer the factors, Generate.
function setupPanelHtml(s, j) {
  s = s || defaultSetup(j);
  const wll = (j.wll != null && j.wll !== '') ? Number(j.wll) : null;
  const pctN = (s.testPct != null && String(s.testPct).trim() !== '' && !isNaN(Number(s.testPct))) ? Number(s.testPct) : 125;
  const autoTL = wll != null ? (+(wll * pctN / 100).toFixed(3) + ' ' + (j.wllUnit || 't')) : 'set WLL on the Details tab';
  const fld = (id, val, ph) => '<input id="' + id + '" value="' + esc(val == null ? '' : val) + '" placeholder="' + esc(ph || '') + '">';
  const radios = JOB_TYPES.map(([v, l]) =>
    '<label class="setup-type' + (s.jobType === v ? ' on' : '') + '"><input type="radio" name="setupJobType" value="' + v + '" data-change="setup-jobtype"' + (s.jobType === v ? ' checked' : '') + '> ' + l + '</label>'
  ).join('');
  let typeFields;
  if (s.jobType === 'steelweights') {
    typeFields =
      '<div class="frow"><div><label>Weights / configuration</label>' + fld('setupWeightConfig', s.weightConfig, 'e.g. 5 × 10 Te plates') + '</div>' +
      '<div><label>Lift &amp; landing area</label>' + fld('setupLandingArea', s.landingArea, 'e.g. quay apron, level') + '</div></div>';
  } else if (s.jobType === 'padeye') {
    typeFields =
      '<div class="frow"><div><label>Number of lugs / pad eyes</label>' + fld('setupPadCount', s.padEyeCount, 'e.g. 4') + '</div>' +
      '<div><label>Pull direction / angle</label>' + fld('setupPullDir', s.pullDirection, 'e.g. vertical, 45°') + '</div></div>' +
      '<div class="frow"><div><label>Ram / tester</label>' + fld('setupTester', s.tester, 'e.g. 30 Te hydraulic ram') + '</div>' +
      '<div><label>Ram surface area (in²)</label>' + fld('setupRamArea', s.ramArea, 'e.g. 7.22 (RC302) — for PSI') + '</div></div>';
  } else {
    typeFields =
      '<div class="frow"><div><label>Number of bags</label>' + fld('setupBagCount', s.bagCount, 'e.g. 4') + '</div>' +
      '<div><label>Bag size</label>' + fld('setupBagSize', s.bagSize, 'e.g. 35 Te') + '</div></div>' +
      '<div class="frow"><div><label>Water source</label>' + fld('setupWaterSource', s.waterSource, 'e.g. quay hydrant') + '</div>' +
      '<div><label>Distance to fill point</label>' + fld('setupWaterDist', s.waterDistance, 'e.g. 150 ft') + '</div></div>';
  }
  return '<div class="setup-box" id="setupPanel">' +
    '<div class="setup-h">Procedure setup <span class="muted">— answer these, then Generate (no AI — fully deterministic)</span></div>' +
    '<div class="setup-types">' + radios + '</div>' +
    '<div class="frow">' +
      '<div><label>Test as % of WLL</label>' + fld('setupTestPct', s.testPct, '125') + '</div>' +
      '<div><label>Test load <span class="muted">(override)</span></label>' + fld('setupTestLoad', s.testLoad, 'auto: ' + autoTL) + '</div>' +
    '</div>' +
    '<div class="frow">' +
      '<div><label>Testing</label><div class="setup-toggles">' +
        '<label class="chk"><input type="checkbox" id="setupFunc"' + (s.functionTest ? ' checked' : '') + '> Function test</label>' +
        '<label class="chk"><input type="checkbox" id="setupStatic" data-change="setup-static"' + (s.staticHold ? ' checked' : '') + '> Static hold</label>' +
        (s.staticHold ? '<span class="setup-hold"><input id="setupHold" value="' + esc(s.holdMinutes || '') + '"> min</span>' : '') +
      '</div></div>' +
    '</div>' +
    typeFields +
    '<button class="btn primary" data-action="proc-generate">' + (state.procedureDraft ? '↻ Regenerate procedure' : 'Generate procedure') + '</button>' +
  '</div>';
}
// Read the setup panel back into state.procedureSetup.
function collectSetup() {
  const s = state.procedureSetup || (state.procedureSetup = defaultSetup(state.detail && state.detail.job, state.detail && state.detail.planning));
  const v = id => { const el = $('#' + id); return el ? el.value.trim() : undefined; };
  const jt = $$('input[name="setupJobType"]').find(r => r.checked);
  if (jt) s.jobType = jt.value;
  const func = $('#setupFunc'); if (func) s.functionTest = func.checked;
  const stat = $('#setupStatic'); if (stat) s.staticHold = stat.checked;
  const set = (id, key) => { const val = v(id); if (val !== undefined) s[key] = val; };
  set('setupTestPct', 'testPct'); set('setupTestLoad', 'testLoad'); set('setupHold', 'holdMinutes');
  set('setupBagCount', 'bagCount'); set('setupBagSize', 'bagSize');
  set('setupWaterSource', 'waterSource'); set('setupWaterDist', 'waterDistance');
  set('setupWeightConfig', 'weightConfig'); set('setupLandingArea', 'landingArea');
  set('setupPadCount', 'padEyeCount'); set('setupPullDir', 'pullDirection'); set('setupTester', 'tester'); set('setupRamArea', 'ramArea');
  return s;
}

// Responsibilities editor: one card per person, with a dropdown of Hydro-Wates staff
// that auto-fills name + contact. Stored as "Name | Company | Role | Contact" lines.
function respRowsHtml(p) {
  const team = (state.templates && state.templates.team) || [];
  const opts = '<option value="">— pick Hydro-Wates staff —</option>' +
    team.map(m => '<option data-name="' + esc(m.name) + '" data-contact="' + esc(m.contact || '') + '">' + esc(m.name) + (m.contact ? ' · ' + esc(m.contact) : '') + '</option>').join('');
  const rows = (p.responsibilities || []).map((line, i) => {
    const c = line.split('|').map(x => x.trim());
    return '<div class="resp-row" data-resp-row>' +
      '<div class="resp-top">' +
        '<select class="resp-pick" data-change="resp-pick">' + opts + '</select>' +
        '<button class="btn-icon" data-action="resp-del" data-i="' + i + '" title="Remove">✕</button>' +
      '</div>' +
      '<div class="resp-grid">' +
        '<input data-rname placeholder="Name" value="' + esc(c[0] || '') + '">' +
        '<input data-rcompany placeholder="Company" value="' + esc(c[1] || '') + '">' +
        '<input data-rrole placeholder="Responsibility (e.g. Run testing)" value="' + esc(c[2] || '') + '">' +
        '<input data-rcontact placeholder="Contact" value="' + esc(c[3] || '') + '">' +
      '</div>' +
    '</div>';
  }).join('');
  return '<div class="resp-list">' + rows + '</div>' +
    '<button class="btn small" data-action="resp-add">+ Add person</button>';
}

// Pull the send-screen recipient inputs into the draft (contacts ticked + typed email).
function collectProcSend() {
  const p = state.procedureDraft; if (!p) return;
  if ($('#procEmail')) p.email = $('#procEmail').value.trim();
  const boxes = $$('[data-pcontact]'); if (boxes.length) p.recipients = boxes.filter(b => b.checked).map(b => b.value);
}

// Guided send: STEP 1 — must preview the procedure; STEP 2 — enter the customer email and send.
function procSendConfirmHtml(d) {
  const p = state.procedureDraft;
  const conf = state.procSendConfirm || {};
  const m = buildProcedureEmail(p, d.job);
  const head =
    '<div class="banner" style="background:#fdf0d9;border:1px solid #f0d49a;color:#92600a">✉️ <b>Send procedure to the customer</b></div>' +
    '<h2 style="margin:6px 0 4px">' + esc(d.job.customer || '') + ' · ' + esc(p.projectRef || d.job.number || '') + '</h2>';

  if (!conf.previewed) {
    return '<div class="proc-confirm">' + head +
      '<p class="hint" style="margin:10px 0"><b>Step 1 of 2 — review.</b> You must <b>preview the procedure</b> before you can send it. Opening the preview unlocks the next step.</p>' +
      '<div class="plan-actions">' +
        '<button class="btn primary" data-action="proc-send-preview">👁 Preview the procedure</button>' +
        '<button class="btn" data-action="proc-send-cancel">← Cancel</button>' +
      '</div>' +
    '</div>';
  }
  return '<div class="proc-confirm">' + head +
    '<div style="color:#1d6f37;font-weight:600;margin:8px 0">✓ Procedure previewed</div>' +
    '<p class="hint" style="margin-bottom:8px"><b>Step 2 of 2 — send.</b> Tick a contact on file, or type the customer’s email address.</p>' +
    contactsBoxHtml() +
    '<div style="margin-top:8px"><label>Customer email</label><input id="procEmail" type="email" placeholder="name@customer.com" value="' + esc(p.email || '') + '"></div>' +
    '<div style="margin-top:10px"><b>Subject:</b> ' + esc(m.subject) + '</div>' +
    '<p class="hint" style="margin:8px 0">Once you press send, the full procedure emails to that address from your account.</p>' +
    '<div class="plan-actions">' +
      '<button class="btn" data-action="proc-send-preview">👁 Preview again</button>' +
      '<button class="btn primary" data-action="proc-send-confirm">✈ Send now</button>' +
      '<button class="btn" data-action="proc-send-cancel">← Cancel</button>' +
    '</div>' +
  '</div>';
}

// The little source tag above the equipment list — shows where the list came from.
function equipmentSourceHtml(p) {
  const src = p && p.equipmentSource;
  if (!src) return '';
  const when = src.at ? fmtDate(src.at) : '';
  return '<div class="equip-source">🚛 From <b>' + esc(src.loadoutRef || src.hwi || '') + '</b> loadout · ' +
    esc(String(src.pieces || 0)) + ' pieces' +
    (src.phase ? ' · ' + esc(src.phase) : '') + (when ? ' · ' + esc(when) : '') + '</div>';
}

function procedureTabHtml(d) {
  const p = state.procedureDraft;
  if (!state.procedureSetup) state.procedureSetup = (p && p.setup) ? Object.assign(defaultSetup(d.job, d.planning), p.setup) : defaultSetup(d.job, d.planning);
  if (!p) {
    return '<p class="muted" style="line-height:1.6;margin-bottom:10px">Pick the job type and answer a few questions, then <b>Generate</b> — it builds a tailored, fully editable draft. No AI; the same answers always produce the same procedure.</p>' +
      setupPanelHtml(state.procedureSetup, d.job);
  }
  if (state.procSendConfirm) return procSendConfirmHtml(d);
  const ta = (id, val, rows, ph) => '<textarea id="' + id + '" rows="' + rows + '"' + (ph ? ' placeholder="' + esc(ph) + '"' : '') + '>' + esc(val || '') + '</textarea>';
  const lines = arr => (arr || []).join('\n');
  const rowsFor = arr => Math.max(3, (arr || []).length + 1);
  const smtpReady = emailReady();
  const sendLog = (p.sentLog || []).slice().reverse().map(l => '<div>' + esc(new Date(l.at).toLocaleString()) + ' → ' + esc((l.to || []).join(', ')) + '</div>').join('');
  return '<div class="proc">' +
    setupPanelHtml(state.procedureSetup, d.job) +
    '<div class="frow">' +
      '<div><label>Status</label><select id="procStatus">' +
        Object.entries(PROC_STATUS).map(([v, l]) => '<option value="' + v + '"' + (p.status === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select></div>' +
      '<div><label>Expected start date</label><input id="procStartDate" value="' + esc(p.startDate || '') + '"></div>' +
      '<div><label>Project reference</label><input id="procRef" value="' + esc(p.projectRef || '') + '"></div>' +
    '</div>' +
    '<div class="frow">' +
      '<div><label>Scope / equipment under test</label><input id="procScope" placeholder="e.g. 2x 10 Te overhead cranes" value="' + esc(p.scope || '') + '"></div>' +
      '<div><label>Job site</label><input id="procSite" placeholder="Site address" value="' + esc(p.jobSite || '') + '"></div>' +
    '</div>' +
    '<label style="margin-top:8px">1. Objective</label>' + ta('procObjective', p.objective, 2) +
    '<label style="margin-top:8px">2. Field coordination notes <span class="muted">(one per line)</span></label>' + ta('procCoordination', p.coordination, 3) +
    '<label style="margin-top:8px">3. Responsibilities <span class="muted">(pick a staff member, or type)</span></label>' + respRowsHtml(p) +
    '<label style="margin-top:8px">4. Equipment &amp; materials</label>' +
    equipmentSourceHtml(p) +
    '<div class="equip-pull"><button class="btn small" data-action="proc-pull-equipment">↻ Pull from Shop Master</button>' +
      '<span class="muted" style="font-size:12px">' + (p.equipmentSource ? 'Live from the job’s loadout — edit any line.' : 'One per line — from your standard list; edit per job.') + '</span></div>' +
    ta('procEquip', lines(p.equipment), rowsFor(p.equipment), 'Add equipment items, one per line') +
    '<label style="margin-top:8px">5. Required PPE <span class="muted">(tick all that apply)</span></label>' + ppeChecklistHtml(p) +
    '<label style="margin-top:8px">6.1 Pre-job preparation</label>' + ta('procPreJob', p.preJob, 6) +
    '<label style="margin-top:8px">6.2 Test setup steps <span class="muted">(one per line)</span></label>' + ta('procSetup', lines(p.setupSteps), rowsFor(p.setupSteps)) +
    '<label style="margin-top:8px">6.3 Load test execution steps <span class="muted">(one per line)</span></label>' + ta('procExec', lines(p.executionSteps), rowsFor(p.executionSteps)) +
    '<label style="margin-top:8px">6.4 Information logging</label>' + ta('procLogging', p.logging, 3) +
    '<label style="margin-top:8px">7. Project drawings <span class="muted">(PDF rigging drawings)</span></label>' +
    riggingPlannerLinkHtml(d.job) +
    drawingsTabHtml(p) +
    '<label style="margin-top:8px">8. Site photos <span class="muted">(JPG/PNG — shown in the printed procedure)</span></label>' +
    photosTabHtml(p) +
    '<div class="frow" style="margin-top:8px">' +
      '<div><label>Approved by</label><input id="procApprovedBy" value="' + esc(p.approvedBy || '') + '"></div>' +
      '<div><label>Approval date</label><input id="procApprovedDate" value="' + esc(p.approvedDate || '') + '"></div>' +
    '</div>' +
    '<div class="plan-actions" style="margin-top:14px;border-top:1px solid var(--line);padding-top:12px">' +
      '<button class="btn" data-action="proc-preview">👁 Preview</button>' +
      '<button class="btn primary" data-action="proc-send">✈ Send to customer</button>' +
      '<button class="btn" data-action="proc-print">🖨 Print / save PDF</button>' +
      '<button class="btn" data-action="proc-copy">Copy</button>' +
      '<button class="btn" data-action="proc-save">Save</button>' +
      '<button class="btn" data-action="proc-generate">↻ Regenerate</button>' +
    '</div>' +
    (smtpReady ? '' : '<div class="sent-note">“Send to customer” emails from <b>your own</b> mailbox once you <b>sign in with Microsoft</b>. Until then, use “Copy” or “Print / save PDF”.</div>') +
    (sendLog ? '<div class="sent-note"><b>Send history</b>' + sendLog + '</div>'
             : (p.sentAt ? '<div class="sent-note">First sent to customer: ' + esc(new Date(p.sentAt).toLocaleString()) + '</div>' : '')) +
    '<div class="sent-note">Steps 1–3 (objective, coordination, responsibilities) are filled in by you. Equipment (step 4) auto-fills from the job’s <b>Shop Master loadout</b> (or your <a href="#templates">Standard equipment</a> list). Tick the PPE for the job (step 5); the procedure steps (step 6) come from the standard template with this job’s WLL / answers — edit any of it.</div>' +
  '</div>';
}

function collectProcedure() {
  const p = state.procedureDraft;
  if (!p || !$('#procObjective')) return p;
  const v = id => { const el = $('#' + id); return el ? el.value : ''; };
  const ls = id => { const el = $('#' + id); return el ? el.value.split('\n').map(s => s.trim()).filter(Boolean) : []; };
  p.status = v('procStatus') || p.status;
  p.startDate = v('procStartDate'); p.projectRef = v('procRef');
  p.scope = v('procScope'); p.jobSite = v('procSite');
  p.objective = v('procObjective'); p.coordination = v('procCoordination');
  p.responsibilities = $$('[data-resp-row]').map(r => {
    const g = sel => { const el = r.querySelector(sel); return el ? el.value.trim() : ''; };
    return [g('[data-rname]'), g('[data-rcompany]'), g('[data-rrole]'), g('[data-rcontact]')].join(' | ');
  }).filter(s => s.replace(/\|/g, '').trim());
  p.equipment = ls('procEquip');
  p.ppe = $$('[data-ppe]').filter(b => b.checked).map(b => b.value).concat(ls('procPPEOther'));
  p.preJob = v('procPreJob'); p.setupSteps = ls('procSetup'); p.executionSteps = ls('procExec');
  p.logging = v('procLogging'); p.approvedBy = v('procApprovedBy'); p.approvedDate = v('procApprovedDate');
  if ($('#procEmail')) p.email = $('#procEmail').value.trim();
  const boxes = $$('[data-pcontact]');
  if (boxes.length) p.recipients = boxes.filter(b => b.checked).map(b => b.value);
  if ($('#setupPanel')) { collectSetup(); p.setup = JSON.parse(JSON.stringify(state.procedureSetup)); }
  return p;
}

async function saveProcedure(silent) {
  const p = collectProcedure();
  if (!p || !state.open) return;
  const keepPhotos = p.photos, keepDrawings = p.drawings;   // attachments are managed separately; keep their live copies (incl. signed photo URLs)
  try {
    const r = await api('PUT', '/api/job/' + encodeURIComponent(state.open) + '/procedure', p);
    state.procedureDraft = JSON.parse(JSON.stringify(r.procedure));
    if (keepPhotos) state.procedureDraft.photos = keepPhotos;
    if (keepDrawings) state.procedureDraft.drawings = keepDrawings;
    state.detail.procedure = r.procedure;
    if (!silent) toast('Procedure saved.');
  } catch (e) { toast(e.message, true); }
}

// Pull the Equipment section straight from this job's Shop Master loadout (by HWI).
// Called automatically after "Generate" (onGenerate) and by the "Pull from Shop
// Master" button (manual). Won't silently clobber edits on a manual re-pull.
async function pullLoadoutEquipment(key, opts) {
  opts = opts || {};
  if (opts.manual && state.procedureDraft && (state.procedureDraft.equipment || []).length &&
      !confirm('Replace the equipment list with this job’s Shop Master loadout?')) return;
  let r;
  try { r = await api('GET', '/api/job/' + encodeURIComponent(key) + '/loadout-equipment'); }
  catch (e) { if (opts.manual) toast(e.message, true); return; }
  if (!state.procedureDraft || state.open !== key) return;
  if (r && r.found && Array.isArray(r.lines) && r.lines.length) {
    if ($('#procEquip')) collectProcedure();          // keep any other in-progress edits
    state.procedureDraft.equipment = r.lines;
    state.procedureDraft.equipmentSource = r.source;
    renderModal();                                    // form now shows the pulled list
    await saveProcedure(true);
    toast('Pulled ' + r.source.pieces + ' pieces from ' + r.source.hwi + '’s Shop Master loadout');
  } else if (opts.manual) {
    toast('No Shop Master loadout for this job yet — pull again once it’s been loaded.', false);
  }
}

// Render the full procedure as a printable / PDF document on the Hydro-Wates letterhead.
function printProcedure(preview) {
  const p = collectProcedure(); const j = state.detail.job;
  if (!p) { toast('Generate the procedure first.', true); return; }
  const win = window.open('', '_blank');
  if (!win) { toast('Allow pop-ups for this site to print.', true); return; }
  const para = s => String(s || '').split('\n').map(l => l.trim() ? '<p>' + esc(l) + '</p>' : '').join('');
  const bullets = s => '<ul class="bul">' + String(s || '').split('\n').map(l => l.trim() ? '<li>' + esc(l) + '</li>' : '').join('') + '</ul>';
  const steps = arr => '<ol class="steps">' + (arr || []).map(s => '<li>' + esc(s) + '</li>').join('') + '</ol>';
  const respRows = (p.responsibilities || []).map(line => {
    const c = line.split('|').map(x => x.trim());
    return '<tr><td>' + esc(c[0] || '') + '</td><td>' + esc(c[1] || '') + '</td><td>' + esc(c[2] || '') + '</td><td>' + esc(c[3] || '') + '</td></tr>';
  }).join('');
  const equip = (p.equipment || []).length ? steps(p.equipment) : '<p class="muted">[ Equipment list to be completed ]</p>';
  const field = (lbl, val) => '<div><span class="fl">' + esc(lbl) + '</span> ' + esc(val || '—') + '</div>';
  const html =
    '<!doctype html><html><head><meta charset="utf-8"><title>Load Test Procedure — ' + esc(j.customer || '') + '</title><style>' +
    'body{font:13.5px/1.55 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1a2733;max-width:760px;margin:28px auto;padding:0 26px}' +
    '.lh{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;border-bottom:3px solid #0b5e8a;padding-bottom:10px;margin-bottom:6px}' +
    '.lh img{max-height:58px;max-width:300px}.co{font-size:11px;color:#5b6770;text-align:right;line-height:1.4}' +
    'h1{font-size:19px;margin:14px 0 8px;color:#0b3a52}h2{font-size:15px;margin:18px 0 6px;color:#0b3a52;border-bottom:1px solid #e9eef2;padding-bottom:3px}' +
    'h3{font-size:13.5px;margin:12px 0 4px}.fields{display:grid;grid-template-columns:1fr 1fr;gap:2px 18px;font-size:13px;margin:6px 0 4px}' +
    '.fl{color:#5b6770}p{margin:5px 0}ol.steps,ul.bul{margin:5px 0;padding-left:22px}ol.steps li,ul.bul li{margin:3px 0}' +
    'table{border-collapse:collapse;width:100%;font-size:13px;margin:6px 0}th,td{border:1px solid #d7dee4;padding:5px 8px;text-align:left}th{background:#f1f5f8}' +
    '.appr{margin-top:24px;border-top:1px solid #e9eef2;padding-top:10px;font-size:13px;display:flex;justify-content:space-between}' +
    '.foot{margin-top:18px;color:#9aa6b1;font-size:11px}' +
    '.photos{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:8px 0}.photos figure{margin:0;break-inside:avoid}.photos img{width:100%;border:1px solid #d7dee4;border-radius:4px}.photos figcaption{font-size:11px;color:#5b6770;margin-top:2px}' +
    '@media print{body{margin:0}h2{break-after:avoid}}' +
    '</style></head><body>' +
    '<div class="lh"><img src="/logo.jpg" alt="Hydro-Wates">' +
      '<div class="co">' + esc(PROC_COMPANY.tagline) + '<br>' + esc(PROC_COMPANY.address) + '<br>Tel: ' + esc(PROC_COMPANY.tel) + '</div></div>' +
    '<h1>Load Test Procedure</h1>' +
    '<div class="fields">' +
      field('Expected start date:', p.startDate) + field('Customer:', j.customer) +
      field('Project reference:', p.projectRef || j.number) + field('Scope:', p.scope) +
      field('Job site:', p.jobSite) + field('Test load (125%):', (j.wll != null && j.wll !== '') ? (+(Number(j.wll) * 1.25).toFixed(3)) + ' ' + (j.wllUnit || 't') : '—') +
    '</div>' +
    '<h2>1. Objective</h2>' + para(p.objective) +
    '<h2>2. Field coordination notes</h2>' + bullets(p.coordination) +
    '<h2>3. Responsibilities</h2><table><tr><th>Name</th><th>Company</th><th>Responsibilities</th><th>Contact</th></tr>' + respRows + '</table>' +
    '<h2>4. Equipment &amp; materials</h2>' + equip +
    '<h2>5. Required PPE</h2>' + ((p.ppe || []).length ? steps(p.ppe) : '<p class="muted">[ PPE to be confirmed ]</p>') +
    '<h2>6. Procedure</h2>' +
    '<h3>6.1 Pre-job preparation</h3>' + para(p.preJob) +
    '<h3>6.2 Test setup — 100% function test &amp; 125% overload</h3>' + steps(p.setupSteps) +
    '<h3>6.3 Load test execution</h3>' + steps(p.executionSteps) +
    '<h3>6.4 Information logging</h3>' + para(p.logging) +
    '<h2>7. Project drawings</h2>' + ((p.drawings || []).length
      ? '<ul class="bul">' + p.drawings.map(d => '<li>📄 ' + esc(d.name) + '</li>').join('') + '</ul>' +
        '<p class="muted" style="font-size:11px">Rigging drawing PDF(s) attached separately.</p>'
      : '<p class="muted">[ Attach rigging drawing ]</p>') +
    ((p.photos || []).filter(ph => ph.url).length
      ? '<h2>8. Site photos</h2><div class="photos">' +
        p.photos.filter(ph => ph.url).map(ph => '<figure><img src="' + esc(ph.url) + '"><figcaption>' + esc(ph.name) + '</figcaption></figure>').join('') +
        '</div>'
      : '') +
    '<div class="appr"><div><b>Approved:</b> ' + esc(p.approvedBy || '________________') + '</div><div><b>Date:</b> ' + esc(p.approvedDate || '____________') + '</div></div>' +
    '<p class="foot">Hydro-Wates Project Manager · ' + esc(new Date().toLocaleString()) + '</p>' +
    '</body></html>';
  win.document.write(html);
  win.document.close();
  win.focus();
  if (!preview) win.onload = () => win.print();
}

function procRecipients() {
  const p = state.procedureDraft || {};
  const to = (p.recipients || []).slice();
  const extra = (p.email || '').trim();
  if (extra && !to.includes(extra)) to.push(extra);
  return to;
}

// Plain-text rendering of the procedure (for Copy and as the email's text fallback).
function procedureText(p, j) {
  const tl = (j.wll != null && j.wll !== '') ? (+(Number(j.wll) * 1.25).toFixed(3)) + ' ' + (j.wllUnit || 't') : '';
  const L = ['LOAD TEST PROCEDURE'];
  L.push('Project reference: ' + (p.projectRef || j.number || ''));
  L.push('Customer: ' + (j.customer || ''));
  if (p.startDate) L.push('Expected start date: ' + p.startDate);
  if (p.scope) L.push('Scope: ' + p.scope);
  if (p.jobSite) L.push('Job site: ' + p.jobSite);
  if (tl) L.push('Test load (125%): ' + tl);
  L.push('', '1. OBJECTIVE', p.objective || '', '', '2. FIELD COORDINATION NOTES');
  (p.coordination || '').split('\n').filter(x => x.trim()).forEach(x => L.push('- ' + x.trim()));
  L.push('', '3. RESPONSIBILITIES');
  (p.responsibilities || []).forEach(line => { const c = line.split('|').map(s => s.trim()); L.push([c[0], c[1], c[2], c[3]].filter(Boolean).join(' — ')); });
  L.push('', '4. EQUIPMENT & MATERIALS');
  if ((p.equipment || []).length) p.equipment.forEach((e, i) => L.push((i + 1) + '. ' + e)); else L.push('To be confirmed.');
  L.push('', '5. REQUIRED PPE');
  if ((p.ppe || []).length) p.ppe.forEach(e => L.push('- ' + e)); else L.push('To be confirmed.');
  L.push('', '6. PROCEDURE', '6.1 Pre-job preparation', p.preJob || '', '', '6.2 Test setup');
  (p.setupSteps || []).forEach((s, i) => L.push((i + 1) + '. ' + s));
  L.push('', '6.3 Load test execution');
  (p.executionSteps || []).forEach((s, i) => L.push((i + 1) + '. ' + s));
  L.push('', '6.4 Information logging', p.logging || '');
  if (p.approvedBy || p.approvedDate) L.push('', 'Approved: ' + (p.approvedBy || '') + '   Date: ' + (p.approvedDate || ''));
  return L.join('\n');
}

// Inline-styled HTML of the procedure for the email body (logo loads from the public website).
function procedureHtml(p, j) {
  const tl = (j.wll != null && j.wll !== '') ? (+(Number(j.wll) * 1.25).toFixed(3)) + ' ' + (j.wllUnit || 't') : '';
  const para = s => String(s || '').split('\n').map(l => l.trim() ? '<p style="margin:5px 0">' + esc(l) + '</p>' : '').join('');
  const bullets = s => '<ul style="margin:5px 0;padding-left:20px">' + String(s || '').split('\n').map(l => l.trim() ? '<li>' + esc(l) + '</li>' : '').join('') + '</ul>';
  const steps = arr => '<ol style="margin:5px 0;padding-left:22px">' + (arr || []).map(s => '<li style="margin:3px 0">' + esc(s) + '</li>').join('') + '</ol>';
  const td = 'style="border:1px solid #d7dee4;padding:5px 8px;text-align:left"';
  const th = 'style="border:1px solid #d7dee4;padding:5px 8px;text-align:left;background:#f1f5f8"';
  const respRows = (p.responsibilities || []).map(line => { const c = line.split('|').map(x => x.trim()); return '<tr><td ' + td + '>' + esc(c[0] || '') + '</td><td ' + td + '>' + esc(c[1] || '') + '</td><td ' + td + '>' + esc(c[2] || '') + '</td><td ' + td + '>' + esc(c[3] || '') + '</td></tr>'; }).join('');
  const equip = (p.equipment || []).length ? steps(p.equipment) : '<p style="color:#888">To be confirmed.</p>';
  const h2 = 'style="margin:18px 0 6px;color:#0b3a52;font-size:16px;border-bottom:1px solid #e9eef2;padding-bottom:3px"';
  const h3 = 'style="margin:12px 0 4px;font-size:14px"';
  return '<div style="font:14px/1.55 Arial,Segoe UI,sans-serif;color:#1a2733;max-width:680px">' +
    '<div style="border-bottom:3px solid #0b5e8a;padding-bottom:8px;margin-bottom:8px">' +
      '<img src="' + PROC_LOGO_URL + '" alt="Hydro-Wates" style="max-height:54px">' +
      '<div style="font-size:11px;color:#5b6770;margin-top:4px">' + esc(PROC_COMPANY.address) + ' · Tel: ' + esc(PROC_COMPANY.tel) + '</div></div>' +
    '<h1 style="font-size:19px;color:#0b3a52;margin:10px 0 6px">Load Test Procedure</h1>' +
    '<div style="font-size:13px;color:#444;margin-bottom:6px">' +
      [p.projectRef ? 'Project ' + esc(p.projectRef) : '', j.customer ? esc(j.customer) : '', tl ? 'Test load ' + esc(tl) + ' (125%)' : ''].filter(Boolean).join(' &nbsp;·&nbsp; ') + '</div>' +
    '<h2 ' + h2 + '>1. Objective</h2>' + para(p.objective) +
    '<h2 ' + h2 + '>2. Field coordination notes</h2>' + bullets(p.coordination) +
    '<h2 ' + h2 + '>3. Responsibilities</h2><table style="border-collapse:collapse;width:100%;font-size:13px"><tr><th ' + th + '>Name</th><th ' + th + '>Company</th><th ' + th + '>Responsibilities</th><th ' + th + '>Contact</th></tr>' + respRows + '</table>' +
    '<h2 ' + h2 + '>4. Equipment &amp; materials</h2>' + equip +
    '<h2 ' + h2 + '>5. Required PPE</h2>' + ((p.ppe || []).length ? steps(p.ppe) : '<p style="color:#888">To be confirmed.</p>') +
    '<h2 ' + h2 + '>6. Procedure</h2>' +
    '<h3 ' + h3 + '>6.1 Pre-job preparation</h3>' + para(p.preJob) +
    '<h3 ' + h3 + '>6.2 Test setup</h3>' + steps(p.setupSteps) +
    '<h3 ' + h3 + '>6.3 Load test execution</h3>' + steps(p.executionSteps) +
    '<h3 ' + h3 + '>6.4 Information logging</h3>' + para(p.logging) +
    ((p.approvedBy || p.approvedDate) ? '<p style="margin-top:18px;border-top:1px solid #e9eef2;padding-top:8px"><b>Approved:</b> ' + esc(p.approvedBy || '') + ' &nbsp; <b>Date:</b> ' + esc(p.approvedDate || '') + '</p>' : '') +
  '</div>';
}

function buildProcedureEmail(p, j) {
  const ref = p.projectRef || j.number || 'upcoming job';
  const sender = signOff();
  const subject = 'Load Test Procedure — ' + ref;
  const coverLine = 'Please find below the load test procedure for ' + (j.customer || 'your site') + ' (' + ref + '). Please review and let us know if you have any questions or require changes ahead of the test.';
  const text = 'Hi,\n\n' + coverLine + '\n\nMany thanks,\n' + sender + '\n\n----------------------------------------\n\n' + procedureText(p, j);
  const html = '<div style="font:14px/1.55 Arial,Segoe UI,sans-serif;color:#1a2733">' +
    '<p>Hi,</p><p>' + esc(coverLine) + '</p>' +
    '<hr style="border:none;border-top:1px solid #e9eef2;margin:14px 0">' +
    procedureHtml(p, j) +
    '<p style="margin-top:12px">Many thanks,<br>' + esc(sender) + '</p></div>';
  return { subject, text, html };
}

/* ---------------- PO tracker ---------------- */
async function loadLeads() {
  const j = await api('GET', '/api/leads');
  state.leads = j.leads || [];
  state.leadsMeta = j;
}

function leadStatusChip(l) {
  if (l.completed) {
    const what = l.how === 'manual' ? 'Completed (manual)' : ('Invoiced' + (l.invoice ? ' · ' + l.invoice.number : ''));
    return '<span class="chip plan-answered">✓ ' + esc(what) + '</span>';
  }
  return '<span class="chip stage-new">Open</span>' + (l.how === 'forced-open' ? ' <span class="chip override">manual</span>' : '');
}

function leadRow(l) {
  const inv = l.invoice;
  const valueTxt = (l.value === null || l.value === undefined) ? '—'
    : (inv ? fmtMoney(l.value, inv.currency) : Number(l.value).toLocaleString());
  return '<tr class="' + (l.completed ? 'done' : '') + '">' +
    '<td><b>' + esc(l.company || l.title || '(no name)') + '</b>' +
      (l.company && l.title ? '<div class="muted" style="font-size:12px">' + esc(l.title) + '</div>' : '') + '</td>' +
    '<td>' + esc(l.po || '—') + '</td>' +
    '<td class="num">' + esc(valueTxt) + '</td>' +
    '<td>' + esc(fmtDate(l.date)) + '</td>' +
    '<td>' + leadStatusChip(l) +
      (inv && l.completed && l.how === 'invoice'
        ? '<div class="muted" style="font-size:11.5px">' + esc(fmtDate(inv.date)) + ' · ' + esc(fmtMoney(inv.total, inv.currency)) + '</div>' : '') + '</td>' +
    '<td style="white-space:nowrap;text-align:right">' +
      '<button class="btn small" data-action="lead-toggle" data-id="' + esc(l.id) + '" data-next="' + (!l.completed) + '">' +
        (l.completed ? 'Reopen' : 'Mark completed') + '</button>' +
      (l.overridden ? ' <button class="btn-icon" title="Back to automatic (follow Zoho Books)" data-action="lead-auto" data-id="' + esc(l.id) + '">↺</button>' : '') +
    '</td></tr>';
}

function poBodyHtml() {
  const q = state.poSearch.trim().toLowerCase();
  let rows = state.leads.filter(l => !q || [l.company, l.title, l.po].some(x => String(x || '').toLowerCase().includes(q)));
  const openCount = rows.filter(l => !l.completed).length;
  const doneCount = rows.length - openCount;
  if (!state.poShowCompleted) rows = rows.filter(l => !l.completed);
  const m = state.leadsMeta || {};
  return '<div style="margin:0 0 10px;display:flex;gap:8px;align-items:center">' +
      '<span class="chip stage-new">Open ' + openCount + '</span>' +
      '<span class="chip plan-answered">✓ Completed ' + doneCount + '</span>' +
    '</div>' +
    '<div class="panel" style="max-width:none;padding:8px 10px 4px">' +
    '<table class="items leads-table">' +
      '<tr><th>Company / job</th><th>PO number</th><th class="num">Value</th><th>PO date</th><th>Status</th><th></th></tr>' +
      (rows.length ? rows.map(leadRow).join('')
        : '<tr><td colspan="6" class="muted" style="text-align:center;padding:22px">Nothing matches</td></tr>') +
    '</table></div>' +
    '<p class="hint" style="max-width:none">' +
      (m.demo ? '' : ('List: <b>' + esc(m.listName || '—') + '</b>' + (m.lastSync ? ' · refreshed ' + esc(new Date(m.lastSync).toLocaleString()) : ' · not refreshed yet') + ' · ')) +
      'A job goes green automatically when a Zoho Books <b>invoice</b> carries its PO number (in the invoice reference). The buttons override that when needed. ' +
      'Open jobs here also appear on the <b>Dashboard</b> (look for the PO badge) until they are invoiced.</p>';
}

/* ----- Invoice tracker: Shop Master received jobs ----- */
async function loadShopmasterJobs() {
  const j = await api('GET', '/api/shopmaster/jobs');
  state.smJobs = j.jobs || [];
}

function smStatusChip(j) {
  if (j.invoiced) return '<span class="chip plan-answered">✓ Invoiced' + (j.invoice ? ' · ' + esc(j.invoice) : '') + '</span>';
  if (j.matchHow === 'no-match') return '<span class="chip zstatus">not matched</span>';
  return '<span class="chip stage-new">Awaiting invoice</span>';
}

const SM_RECENT_DAYS = 60;
function smIsRecent(j) {
  if (!j.invoiced || !j.invoiceDate) return false;
  const t = new Date(j.invoiceDate).getTime();
  return !isNaN(t) && t >= Date.now() - SM_RECENT_DAYS * 864e5;
}

// Value cell: exact invoice amount for invoiced jobs; projected Lead-List ValueTotal
// (styled muted/italic with a ~ and "projected") for jobs still awaiting an invoice.
function smValueCell(j) {
  if (j.total != null) return esc(fmtMoney(j.total, j.currency));
  if (j.projected != null) {
    return '<span class="proj-val" title="Projected value from the SharePoint Lead List (ValueTotal), matched by HWI — not yet invoiced">~' +
      esc(fmtMoney(j.projected, j.currency)) + '<span class="proj-tag">projected</span></span>';
  }
  return '—';
}

function smStatusCell(j) {
  if (j.invoiced) {
    return '<span class="chip plan-answered">✓ Invoiced' + (j.invoice ? ' · ' + esc(j.invoice) : '') + '</span>' +
      (j.invoiceDate ? '<div class="muted" style="font-size:11.5px">' + esc(fmtDate(j.invoiceDate)) + '</div>' : '');
  }
  return '<span class="chip stage-new">Awaiting invoice</span>';
}

function smRowHtml(j) {
  return '<tr class="' + (j.invoiced ? 'done' : '') + '">' +
    '<td><b>' + esc(j.hwi) + '</b></td>' +
    '<td>' + esc(j.customer || '—') + '</td>' +
    '<td>' + esc(j.poNumber || '—') + '</td>' +
    '<td class="num">' + smValueCell(j) + '</td>' +
    '<td>' + esc(j.phase || '—') + '</td>' +
    '<td>' + esc(fmtDate(j.receivedAt)) + '</td>' +
    '<td>' + smStatusCell(j) + '</td>' +
  '</tr>';
}
function smBodyHtml() {
  const all = state.smJobs || [];
  const q = state.smSearch.trim().toLowerCase();
  let rows = all.filter(j => !q || [j.hwi, j.customer, j.poNumber, j.invoice].some(x => String(x || '').toLowerCase().includes(q)));
  const awaiting = rows.filter(j => !j.invoiced).length;
  const recent = rows.filter(smIsRecent).length;
  const invoiced = rows.filter(j => j.invoiced).length;
  // Default view = the worklist: not-yet-invoiced + recently invoiced. "Show all" reveals older invoiced jobs.
  if (!state.smShowAll) rows = rows.filter(j => !j.invoiced || smIsRecent(j));
  const byRecent = (a, b) => String(b.receivedAt || '').localeCompare(String(a.receivedAt || ''));
  const awaitingRows = rows.filter(j => !j.invoiced).sort(byRecent);
  const invoicedRows = rows.filter(j => j.invoiced).sort(byRecent);
  const collapsed = state.smInvoicedCollapsed !== false;   // invoiced (done) rows are collapsed by default
  let body;
  if (!awaitingRows.length && !invoicedRows.length) {
    body = '<tr><td colspan="7" class="muted" style="text-align:center;padding:22px">Nothing matches</td></tr>';
  } else {
    body = awaitingRows.map(smRowHtml).join('');
    if (invoicedRows.length) {
      body += '<tr class="sm-group"><td colspan="7">' +
        '<button class="sm-toggle" data-action="sm-toggle-invoiced">' +
          '<span class="sm-caret">' + (collapsed ? '▾' : '▴') + '</span> ✓ Invoiced' + (state.smShowAll ? '' : ' (last ' + SM_RECENT_DAYS + ' days)') + ' · ' + invoicedRows.length +
        '</button></td></tr>';
      if (!collapsed) body += invoicedRows.map(smRowHtml).join('');
    }
  }
  return '<div style="margin:0 0 10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">' +
      '<span class="chip stage-new">Awaiting invoice ' + awaiting + '</span>' +
      '<span class="chip plan-answered">✓ Invoiced (last ' + SM_RECENT_DAYS + ' days) ' + recent + '</span>' +
      '<span class="muted" style="font-size:13px">' + invoiced + ' invoiced of ' + all.length + ' received jobs</span>' +
    '</div>' +
    '<div class="panel" style="max-width:none;padding:8px 10px 4px">' +
    '<table class="items leads-table">' +
      '<tr><th>HWI job no.</th><th>Customer</th><th>PO number</th><th class="num">Value</th><th>Phase</th><th>Received</th><th>Invoice status</th></tr>' +
      body +
    '</table></div>';
}

function renderShopmaster() {
  if (state.smJobs === null) {
    $('#view').innerHTML = '<div class="panel"><p class="hint"><span class="spin">⟳</span> Loading received jobs from Shop Master…</p></div>';
    loadShopmasterJobs()
      .then(() => { if (state.view === 'po') renderShopmaster(); })
      .catch(e => { if (state.view === 'po') $('#view').innerHTML = '<div class="panel"><h2>Invoices</h2><p class="hint">⚠️ ' + esc(e.message) + '</p><button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>'; });
    return;
  }
  $('#view').innerHTML =
    '<div class="toolbar">' +
      '<input type="search" id="smSearch" placeholder="Search HWI or customer…" value="' + esc(state.smSearch) + '">' +
      '<label class="chk"><input type="checkbox" data-change="sm-show-all"' + (state.smShowAll ? ' checked' : '') + '> show all invoiced</label>' +
      '<button class="btn" data-action="sm-refresh" style="margin-left:auto">⟳ Refresh</button>' +
    '</div>' +
    '<div id="smBody">' + smBodyHtml() + '</div>';
}

function renderPo() {
  if (state.settings && state.settings.shopmasterConnected) return renderShopmaster();
  if (state.leads === null) {
    $('#view').innerHTML = '<div class="panel"><p class="hint"><span class="spin">⟳</span> Loading the lead list…</p></div>';
    loadLeads()
      .then(() => { if (state.view === 'po') renderPo(); })
      .catch(e => { if (state.view === 'po') $('#view').innerHTML = '<div class="panel"><h2>Invoices</h2><p class="hint">' + esc(e.message) + '</p></div>'; });
    return;
  }
  const m = state.leadsMeta || {};
  let banner = '';
  if (m.demo) {
    banner = '<div class="banner">🔎 <b>Showing demo leads.</b> Connect SharePoint in Settings to see your real lead list.' +
      '<button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>';
  } else if (!m.configured) {
    banner = '<div class="banner">⚙️ Connected to Microsoft 365 — now pick the site and your “Lead List” in Settings.' +
      '<button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>';
  } else if (!m.zohoConnected) {
    banner = '<div class="banner">ℹ️ Connect Zoho Books too, so jobs complete automatically when they are invoiced.' +
      '<button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>';
  }
  if (m.mapWarning) {
    banner += '<div class="banner">⚠️ ' + esc(m.mapWarning) +
      '<button class="btn" data-action="nav" data-view="settings">Open Settings</button></div>';
  }
  $('#view').innerHTML = banner +
    '<div class="toolbar">' +
      '<input type="search" id="poSearch" placeholder="Search company, job or PO number…" value="' + esc(state.poSearch) + '">' +
      '<label class="chk"><input type="checkbox" data-change="po-show-done"' + (state.poShowCompleted ? ' checked' : '') + '> show completed</label>' +
      '<button class="btn" data-action="po-refresh" style="margin-left:auto"' + (m.demo ? ' title="Demo data — connect SharePoint in Settings"' : '') + '>⟳ Refresh' + (m.demo ? '' : ' from SharePoint') + '</button>' +
    '</div>' +
    '<div id="poBody">' + poBodyHtml() + '</div>';
}

/* ---------------- templates view ---------------- */
function collectTemplates() {
  const t = state.templates;
  const texts = $$('[data-tq]');
  if (texts.length) {
    t.questions = texts.map((el, i) => {
      const prev = t.questions[i] || {};
      const row = el.closest('.qrow');
      const typeEl = row && row.querySelector('[data-tqtype]');
      const type = typeEl ? typeEl.value : (prev.type || 'text');
      const q = { id: prev.id || ('q' + Math.random().toString(36).slice(2, 8)), text: el.value, type };
      if (type === 'number') {
        const u = row && row.querySelector('[data-tqunit]');
        const raw = u ? u.value : (prev.units && prev.units.length ? prev.units.join(', ') : (prev.unit || ''));
        const parts = String(raw).split(',').map(s => s.trim()).filter(Boolean);
        if (parts.length > 1) { q.units = parts; q.unit = parts[0]; }
        else { q.unit = parts[0] || ''; }
      } else if (type === 'choice') {
        const o = row && row.querySelector('[data-tqopts]');
        q.options = o ? o.value.split(',').map(s => s.trim()).filter(Boolean) : (prev.options || []);
      }
      return q;
    }).filter(q => q.text.trim());
  }
  if ($('#tplSubject')) t.emailSubject = $('#tplSubject').value;
  if ($('#tplIntro')) t.emailIntro = $('#tplIntro').value;
  if ($('#tplOutro')) t.emailOutro = $('#tplOutro').value;
  if ($('.team-list')) {
    t.team = $$('[data-team-row]').map(r => ({
      name: r.querySelector('[data-tmname]').value.trim(),
      contact: r.querySelector('[data-tmcontact]').value.trim(),
      email: (r.querySelector('[data-tmemail]') ? r.querySelector('[data-tmemail]').value.trim() : '')
    })).filter(m => m.name || m.contact || m.email);
  }
  if ($('#tplEquip')) t.equipment = $('#tplEquip').value.split('\n').map(s => s.trim()).filter(Boolean);
  return t;
}

function teamRowsHtml(t) {
  const team = t.team || [];
  return '<div class="team-list">' + team.map((m, i) =>
    '<div class="team-row" data-team-row>' +
      '<input data-tmname placeholder="Name" value="' + esc(m.name || '') + '">' +
      '<input data-tmcontact placeholder="Phone" value="' + esc(m.contact || '') + '">' +
      '<input data-tmemail type="email" placeholder="Email" value="' + esc(m.email || '') + '">' +
      '<button class="btn-icon" data-action="tpl-team-del" data-i="' + i + '" title="Remove">✕</button>' +
    '</div>').join('') + '</div>';
}

// The Hydro-Wates staff roster — now lives in Settings. Name feeds the procedure
// Responsibilities + meeting Action-item dropdowns; Email is where the meeting report is sent.
function teamPanelHtml() {
  const t = state.templates || {};
  return '<div class="panel">' +
    '<h2>Team / employees <span class="muted" style="font-weight:400;font-size:13px">(Hydro-Wates staff)</span></h2>' +
    '<p class="hint">Your staff roster. The <b>name</b> appears in the dropdowns when you assign <b>Responsibilities</b> in a procedure and <b>Action items</b> in a meeting. The <b>email</b> is where the <b>meeting report</b> is sent — to everyone here <b>except Mike Scofield</b>.</p>' +
    '<div class="team-head"><span>Name</span><span>Phone</span><span>Email</span><span></span></div>' +
    teamRowsHtml(t) +
    '<button class="btn small" data-action="tpl-team-add">+ Add person</button>' +
    '<div style="margin-top:12px"><button class="btn primary" data-action="tpl-save">Save team</button></div>' +
  '</div>';
}

function renderTemplates() {
  const t = state.templates;
  const qs = t.questions.map((q, i) => {
    const type = q.type || 'text';
    const typeSel = '<select data-tqtype data-change="tpl-type">' +
      Q_TYPES.map(([v, l]) => '<option value="' + v + '"' + (type === v ? ' selected' : '') + '>' + l + '</option>').join('') +
      '</select>';
    const extra = type === 'number'
      ? '<input data-tqunit placeholder="unit — e.g. ft  (or ft, m for a choice)" value="' + esc(q.units && q.units.length ? q.units.join(', ') : (q.unit || '')) + '">'
      : type === 'choice'
        ? '<input data-tqopts placeholder="options, comma-separated" value="' + esc((q.options || []).join(', ')) + '">'
        : '';
    return '<div class="qrow">' +
      '<div class="qnum">' + (i + 1) + '</div>' +
      '<div style="flex:1">' +
        '<textarea data-tq rows="2">' + esc(q.text) + '</textarea>' +
        '<div class="tpl-type"><label>Answer type</label>' + typeSel + extra + '</div>' +
      '</div>' +
      '<div class="qbtns">' +
        '<button class="btn-icon" title="Move up" data-action="tpl-up" data-i="' + i + '">▲</button>' +
        '<button class="btn-icon" title="Move down" data-action="tpl-down" data-i="' + i + '">▼</button>' +
        '<button class="btn-icon" title="Remove" data-action="tpl-del" data-i="' + i + '">✕</button>' +
      '</div>' +
    '</div>';
  }).join('');

  $('#view').innerHTML =
    '<div class="panel">' +
      '<h2>Standard planning questions</h2>' +
      '<p class="hint">This is the master list loaded for every new job. You can still edit the questions per job before sending — changes here only affect jobs planned from now on.</p>' +
      qs +
      '<button class="btn small" data-action="tpl-add">+ Add question</button>' +
    '</div>' +
    '<div class="panel">' +
      '<h2>Email wording</h2>' +
      '<p class="hint">Used when you press “Open email draft” on a job. Placeholders: <b>{customer}</b> = customer name, <b>{job}</b> = job number.</p>' +
      '<div style="margin-bottom:12px"><label>Subject</label><input id="tplSubject" value="' + esc(t.emailSubject) + '"></div>' +
      '<div style="margin-bottom:12px"><label>Opening</label><textarea id="tplIntro" rows="3">' + esc(t.emailIntro) + '</textarea></div>' +
      '<div style="margin-bottom:12px"><label>Closing</label><textarea id="tplOutro" rows="3">' + esc(t.emailOutro) + '</textarea></div>' +
      '<button class="btn primary" data-action="tpl-save">Save questions & wording</button>' +
    '</div>' +
    '<div class="panel">' +
      '<h2>Standard equipment <span class="muted" style="font-weight:400;font-size:13px">(load-test kit)</span></h2>' +
      '<p class="hint">This list is pre-loaded into the <b>Equipment &amp; materials</b> section (step 4) of every new procedure — you then edit it per job (sizes, quantities). One item per line.</p>' +
      '<textarea id="tplEquip" rows="' + Math.max(6, (t.equipment || []).length + 1) + '">' + esc((t.equipment || []).join('\n')) + '</textarea>' +
      '<div style="margin-top:12px"><button class="btn primary" data-action="tpl-save">Save equipment</button></div>' +
    '</div>';
}

/* ---------------- settings view ---------------- */
function collectSettings() {
  const s = state.settings;
  const get = (id) => { const el = $(id); return el ? el.value : undefined; };
  if ($('#setDc')) s.dc = get('#setDc');
  if ($('#setCid')) s.clientId = get('#setCid').trim();
  if ($('#setCsec')) s.clientSecret = get('#setCsec').trim();
  if ($('#setMaxPages')) s.maxPages = Number(get('#setMaxPages')) || 3;
  if ($('#setDemo')) s.demoMode = get('#setDemo');
  if ($('#setDefCat')) s.defaultCategory = get('#setDefCat');
  $$('[data-mod]').forEach(el => { s.modules[el.dataset.mod] = el.checked; });
  if ($('#smtpHost')) {
    s.smtp = {
      host: $('#smtpHost').value.trim(),
      port: Number($('#smtpPort').value) || 587,
      user: $('#smtpUser').value.trim(),
      pass: $('#smtpPass').value,
      fromName: $('#smtpFromName').value,
      fromAddr: $('#smtpFromAddr').value.trim()
    };
  }
  if ($('#msTenant')) {
    s.ms.tenant = $('#msTenant').value.trim() || 'organizations';
    s.ms.clientId = $('#msCid').value.trim();
  }
  if ($('#mapPo')) {
    s.ms.map = {
      po: $('#mapPo').value, poMode: $('#mapPoMode').value, poValue: $('#mapPoValue').value,
      company: $('#mapCompany').value, value: $('#mapValue').value, date: $('#mapDate').value
    };
  }
  if ($('#smUrl')) s.shopmaster = { url: $('#smUrl').value.trim(), key: $('#smKey').value.trim() };
  const rows = $$('.rule-row');
  if (rows.length || $('#rulesEmpty')) {
    s.rules = rows.map(r => ({
      keyword: r.querySelector('[data-rk]').value.trim(),
      category: r.querySelector('[data-rc]').value
    })).filter(r => r.keyword);
  }
  return s;
}

function settingsPayload(s) {
  return {
    dc: s.dc, clientId: s.clientId, clientSecret: s.clientSecret, senderName: s.senderName,
    modules: s.modules, rules: s.rules, defaultCategory: s.defaultCategory,
    maxPages: s.maxPages, demoMode: s.demoMode, orgId: s.orgId, orgName: s.orgName,
    smtp: s.smtp,
    ms: {
      tenant: s.ms.tenant, clientId: s.ms.clientId,
      siteId: s.ms.siteId, siteName: s.ms.siteName,
      listId: s.ms.listId, listName: s.ms.listName, map: s.ms.map
    },
    shopmaster: s.shopmaster || { url: '', key: '' }
  };
}

function renderSettings() {
  const s = state.settings;
  const dcs = [['com', 'zoho.com (US)'], ['eu', 'zoho.eu (Europe)'], ['in', 'zoho.in (India)'],
    ['com.au', 'zoho.com.au (Australia)'], ['jp', 'zoho.jp (Japan)'], ['ca', 'zoho.ca (Canada)'], ['sa', 'zoho.sa (Saudi Arabia)']];

  const connHtml = s.connected
    ? '<span class="conn-pill on">● Connected' + (s.orgName ? ' — ' + esc(s.orgName) : '') + '</span>' +
      '<div class="plan-actions" style="margin-top:14px">' +
        '<button class="btn" data-action="set-orgs-load">Choose organisation…</button>' +
        (state.orgs ? '<select id="orgPick" data-change="org-pick" style="width:auto">' +
          '<option value="">— pick organisation —</option>' +
          state.orgs.map(o => '<option value="' + esc(o.id) + '"' + (s.orgId === o.id ? ' selected' : '') + '>' + esc(o.name) + '</option>').join('') +
        '</select>' : '') +
        '<button class="btn" data-action="sync-now">Sync now</button>' +
        '<button class="btn danger" data-action="set-disconnect">Disconnect</button>' +
      '</div>'
    : '<span class="conn-pill off">● Not connected — showing demo data</span>' +
      '<p class="hint" style="margin-top:14px">One-time setup (about 5 minutes):</p>' +
      '<ol class="hint" style="margin:0 0 14px;padding-left:20px;line-height:1.8">' +
        '<li>Open <b>api-console.zoho.com</b> (sign in with your Zoho account) → <b>Add Client</b> → <b>Server-based Applications</b>.</li>' +
        '<li>Homepage URL: <b>http://localhost:8743</b> &nbsp;·&nbsp; Authorized Redirect URI: copy the box below.</li>' +
        '<li>Copy the <b>Client ID</b> and <b>Client Secret</b> it gives you into the fields below.</li>' +
        '<li>Press <b>Save &amp; Connect</b> and approve access when Zoho asks.</li>' +
      '</ol>' +
      '<div class="code-box" style="margin-bottom:14px"><span>' + esc(s.redirectUri || 'http://localhost:8743/oauth/callback') + '</span>' +
        '<button class="btn small" data-action="copy-redirect">Copy</button></div>' +
      '<div class="frow">' +
        '<div><label>Zoho region</label><select id="setDc">' +
          dcs.map(([v, l]) => '<option value="' + v + '"' + (s.dc === v ? ' selected' : '') + '>' + l + '</option>').join('') +
        '</select></div>' +
        '<div><label>Client ID</label><input id="setCid" value="' + esc(s.clientId) + '" autocomplete="off"></div>' +
        '<div><label>Client Secret</label><input id="setCsec" type="password" value="' + esc(s.clientSecret) + '" autocomplete="off"></div>' +
      '</div>' +
      '<button class="btn primary" data-action="set-connect">Save &amp; Connect to Zoho Books</button>';

  const mods = [['estimates', 'Estimates (quotes)'], ['salesorders', 'Sales orders'], ['invoices', 'Invoices'], ['projects', 'Projects (timesheets)']];

  // ----- SharePoint lead list panel -----
  const ms = s.ms || { map: {} };
  const mapOpt = (sel, none) => '<option value="">' + none + '</option>' +
    (state.msCols || []).map(c => '<option value="' + esc(c.name) + '"' + (sel === c.name ? ' selected' : '') + '>' + esc(c.displayName) + '</option>').join('');
  let spHtml;
  if (!s.msConnected) {
    spHtml = '<span class="conn-pill off">● Not connected — the Invoices page shows demo leads</span>' +
      '<p class="hint" style="margin-top:14px">One-time setup — if your company locks down Azure, ask IT to do steps 1–4 (takes them 2 minutes):</p>' +
      '<ol class="hint" style="margin:0 0 14px;padding-left:20px;line-height:1.8">' +
        '<li>Go to <b>portal.azure.com</b> → <b>Microsoft Entra ID</b> → <b>App registrations</b> → <b>New registration</b>.</li>' +
        '<li>Name: <b>HW Project Manager</b> · accounts: <b>this organisation only</b>.</li>' +
        '<li>Redirect URI: pick platform <b>Public client/native (mobile &amp; desktop)</b> and paste the box below.</li>' +
        '<li>Then <b>API permissions</b> → Add a permission → Microsoft Graph → <b>Delegated</b> → tick <b>Sites.Read.All</b>.</li>' +
        '<li>Copy the <b>Application (client) ID</b> from its Overview page into the field below, then connect.</li>' +
      '</ol>' +
      '<div class="code-box" style="margin-bottom:14px"><span>' + esc(s.msRedirectUri || 'http://localhost:8743/ms/callback') + '</span>' +
        '<button class="btn small" data-action="copy-ms-redirect">Copy</button></div>' +
      '<div class="frow">' +
        '<div><label>Tenant (leave as-is unless IT says otherwise)</label><input id="msTenant" value="' + esc(ms.tenant || 'organizations') + '"></div>' +
        '<div><label>Application (client) ID</label><input id="msCid" value="' + esc(ms.clientId || '') + '" autocomplete="off"></div>' +
      '</div>' +
      '<button class="btn primary" data-action="ms-connect">Save &amp; Connect to Microsoft 365</button>';
  } else {
    const siteRow = ms.siteId
      ? '<div style="margin-bottom:10px">Site: <b>' + esc(ms.siteName || ms.siteId) + '</b> <button class="btn small" data-action="ms-site-clear">change</button></div>'
      : '<div class="frow" style="align-items:flex-end;margin-bottom:4px">' +
          '<div><label>Find your SharePoint site</label><input id="msSiteQ" placeholder="part of the site name, e.g. Sales"></div>' +
          '<div style="flex:0 0 auto"><button class="btn" data-action="ms-site-search">Search</button></div>' +
        '</div>' +
        (state.msSites ? '<div style="margin:0 0 10px"><select data-change="ms-site-pick"><option value="">— pick the site (' + state.msSites.length + ' found) —</option>' +
          state.msSites.map(x => '<option value="' + esc(x.id) + '" data-name="' + esc(x.name) + '">' + esc(x.name) + ' — ' + esc(x.url) + '</option>').join('') +
        '</select></div>' : '');
    const listRow = !ms.siteId ? '' : (ms.listId
      ? '<div style="margin-bottom:10px">List: <b>' + esc(ms.listName || ms.listId) + '</b> <button class="btn small" data-action="ms-list-clear">change</button></div>'
      : (state.msLists
        ? '<div style="margin:0 0 10px"><select data-change="ms-list-pick"><option value="">— pick the list —</option>' +
            state.msLists.map(x => '<option value="' + esc(x.id) + '" data-name="' + esc(x.name) + '">' + esc(x.name) + '</option>').join('') +
          '</select></div>'
        : '<button class="btn" data-action="ms-lists-load" style="margin-bottom:10px">Choose the list…</button>'));
    const mapBlock = !ms.listId ? '' : (state.msCols
      ? '<hr><p class="hint" style="margin-bottom:10px"><b>Which columns mean what?</b> Adjust if the guesses are wrong, then press <b>Save settings</b> at the bottom.</p>' +
        '<div class="frow">' +
          '<div><label>“PO received” column</label><select id="mapPo">' + mapOpt(ms.map.po, '— pick column —') + '</select></div>' +
          '<div><label>Counts as received when…</label><select id="mapPoMode">' +
            '<option value="nonempty"' + (ms.map.poMode !== 'yes' && ms.map.poMode !== 'equals' ? ' selected' : '') + '>it has any value (e.g. PO number typed in)</option>' +
            '<option value="yes"' + (ms.map.poMode === 'yes' ? ' selected' : '') + '>it is ticked Yes</option>' +
            '<option value="equals"' + (ms.map.poMode === 'equals' ? ' selected' : '') + '>it equals this text →</option>' +
          '</select></div>' +
          '<div style="max-width:170px"><label>…text (for “equals”)</label><input id="mapPoValue" value="' + esc(ms.map.poValue || '') + '" placeholder="e.g. PO received"></div>' +
        '</div>' +
        '<div class="frow">' +
          '<div><label>Company column</label><select id="mapCompany">' + mapOpt(ms.map.company, '— use the Title column —') + '</select></div>' +
          '<div><label>Value column (optional)</label><select id="mapValue">' + mapOpt(ms.map.value, '— none —') + '</select></div>' +
          '<div><label>PO date column (optional)</label><select id="mapDate">' + mapOpt(ms.map.date, '— use created date —') + '</select></div>' +
        '</div>'
      : '<p class="hint"><span class="spin">⟳</span> Loading the list’s columns…</p>');
    spHtml = '<span class="conn-pill on">● Connected to Microsoft 365</span>' +
      '<div style="margin-top:14px">' + siteRow + listRow + mapBlock + '</div>' +
      '<div class="plan-actions"><button class="btn danger" data-action="ms-disconnect">Disconnect</button></div>';
  }

  const zohoPanel =
    '<div class="panel"><h2>Zoho Books connection</h2>' +
      '<p class="hint">The app only <b>reads</b> from Zoho Books — it never changes anything there.</p>' + connHtml +
    '</div>';

  const spPanel =
    '<div class="panel"><h2>SharePoint lead list</h2>' +
      '<p class="hint">Feeds the <b>Invoices</b> page: every lead in your list that has received a PO, marked completed automatically once it is invoiced in Zoho Books. Read-only — the app never changes the list.</p>' +
      spHtml +
    '</div>';

  const shopmasterPanel =
    '<div class="panel"><h2>Shop Master <span class="muted" style="font-weight:400;font-size:13px">(received-jobs feed for the Invoices page)</span></h2>' +
      '<p class="hint">Read-only connection to Shop Master’s Supabase database. The <b>Invoices</b> page lists your received jobs from here and cross-references Zoho Books to show which are invoiced. ' +
        (s.shopmasterConnected ? '<b style="color:#1d6f37">● Connected</b>' : '<b style="color:#a02626">● Not connected</b>') + '</p>' +
      '<div class="frow">' +
        '<div><label>Supabase project URL</label><input id="smUrl" value="' + esc((s.shopmaster && s.shopmaster.url) || '') + '" placeholder="https://xxxx.supabase.co"></div>' +
        '<div><label>anon / public API key</label><input id="smKey" type="password" value="' + esc((s.shopmaster && s.shopmaster.key) || '') + '" autocomplete="off" placeholder="eyJ…"></div>' +
      '</div>' +
      '<div class="plan-actions">' +
        '<button class="btn" data-action="sm-test">Test connection</button>' +
        '<span class="muted" style="font-size:12.5px;align-self:center">Read-only. Use the <b>anon / public</b> key — never the service_role key.</span>' +
      '</div>' +
    '</div>';

  const jobCountPanel =
    '<div class="panel"><h2>What counts as a job?</h2>' +
      '<p class="hint">Tick the Zoho Books record types that should appear on the dashboard.</p>' +
      '<div class="frow tight">' +
        mods.map(([v, l]) => '<div><label class="chk" style="font-size:14px;color:var(--ink)"><input type="checkbox" data-mod="' + v + '"' + (s.modules[v] ? ' checked' : '') + '> ' + l + '</label></div>').join('') +
      '</div>' +
      '<div class="frow" style="margin-top:8px"><div style="max-width:220px"><label>How far back to fetch (pages of 200 records)</label>' +
      '<input id="setMaxPages" type="number" min="1" max="10" value="' + esc(s.maxPages) + '"></div></div>' +
    '</div>';

  const rulesPanel =
    '<div class="panel"><h2>Rental / Service / Sales rules</h2>' +
      '<p class="hint">Each job is matched against these keywords (checked against its line items, reference, notes and customer name). ' +
      '<b>First match wins</b>, top to bottom. Anything that matches nothing goes to the default. You can always override a single job from its Details tab.</p>' +
      (s.rules.length ? s.rules.map((r, i) =>
        '<div class="rule-row">' +
          '<input data-rk value="' + esc(r.keyword) + '" placeholder="keyword, e.g. rental">' +
          '<select data-rc>' + CATS.map(([v, l]) => '<option value="' + v + '"' + (r.category === v ? ' selected' : '') + '>' + l + '</option>').join('') + '</select>' +
          '<button class="btn-icon" data-action="set-rule-up" data-i="' + i + '" title="Move up">▲</button>' +
          '<button class="btn-icon" data-action="set-rule-down" data-i="' + i + '" title="Move down">▼</button>' +
          '<button class="btn-icon" data-action="set-rule-del" data-i="' + i + '" title="Remove">✕</button>' +
        '</div>').join('') : '<p class="hint" id="rulesEmpty">No rules yet.</p>') +
      '<div class="plan-actions">' +
        '<button class="btn small" data-action="set-rule-add">+ Add rule</button>' +
        '<div style="display:flex;align-items:center;gap:8px;margin-left:auto"><label style="margin:0">If nothing matches:</label>' +
        '<select id="setDefCat" style="width:auto">' + CATS.map(([v, l]) => '<option value="' + v + '"' + (s.defaultCategory === v ? ' selected' : '') + '>' + l + '</option>').join('') + '</select></div>' +
      '</div>' +
    '</div>';

  const demoPanel =
    '<div class="panel"><h2>Demo data</h2>' +
      '<p class="hint"><b>Demo data is fake sample jobs — for demonstrations and testing only.</b> It never touches or shows your real jobs. Keep it <b>off</b> for normal use; turn it <b>on</b> only when you want to show the app with placeholder data.</p>' +
      '<select id="setDemo" style="max-width:380px">' +
        '<option value="off"' + (s.demoMode === 'on' ? '' : ' selected') + '>Off — use my real jobs</option>' +
        '<option value="on"' + (s.demoMode === 'on' ? ' selected' : '') + '>On — show demo data only (demos &amp; testing)</option>' +
      '</select>' +
    '</div>';

  const saveBar = '<div class="set-savebar"><button class="btn primary" data-action="set-save">Save settings</button></div>';
  const groups = {
    connections: zohoPanel + spPanel + shopmasterPanel + saveBar,
    rules: jobCountPanel + rulesPanel + saveBar,
    team: teamPanelHtml(),
    prefs: demoPanel + saveBar
  };
  const TABS = [['connections', 'Connections'], ['rules', 'Job rules'], ['team', 'Team'], ['prefs', 'Preferences']];
  const tab = groups[state.settingsTab] ? state.settingsTab : 'connections';
  const tabBar = '<div class="subtabs">' +
    TABS.map(([k, l]) => '<button class="subtab' + (tab === k ? ' active' : '') + '" data-action="set-tab" data-tab="' + k + '">' + l + '</button>').join('') +
  '</div>';
  $('#view').innerHTML = tabBar + groups[tab];

  // Fetch the lead list's columns once, then guess sensible mappings.
  if (s.msConnected && ms.listId && state.msCols === null && !state._msColsLoading) {
    state._msColsLoading = true;
    api('GET', '/api/ms/columns').then(r => {
      state.msCols = r.columns || [];
      guessMapping();
      state._msColsLoading = false;
      if (state.view === 'settings') renderSettings();
    }).catch(e => {
      state.msCols = [];
      state._msColsLoading = false;
      toast(e.message, true);
      if (state.view === 'settings') renderSettings();
    });
  }
}

function guessMapping() {
  const map = state.settings.ms.map;
  const cols = state.msCols || [];
  const find = (re) => { const c = cols.find(x => re.test(x.displayName)); return c ? c.name : ''; };
  if (!map.po) map.po = find(/\bp\.?o\.?\b|purchase\s*order/i);
  if (!map.company) map.company = find(/company|customer|client|account/i);
  if (!map.value) map.value = find(/value|amount|price|total|quote/i);
  if (!map.date) map.date = find(/\bdate\b/i);
  if (map.po) {
    const c = cols.find(x => x.name === map.po);
    if (c && c.type === 'yesno' && map.poMode === 'nonempty') map.poMode = 'yes';
  }
}

/* ---------------- router / render ---------------- */
function render() {
  $$('.nav-btn').forEach(b => b.classList.toggle('active', b.dataset.view === state.view));
  if (state.view === 'templates') renderTemplates();
  else if (state.view === 'settings') renderSettings();
  else if (state.view === 'po') renderPo();
  else renderDashboard();
  updateSyncStatus();
}

function routeFromHash() {
  const h = (location.hash || '').replace('#', '');
  state.view = ['dashboard', 'po', 'templates', 'settings'].includes(h) ? h : 'dashboard';
}

/* ---------------- actions ---------------- */
document.addEventListener('click', async (e) => {
  const el = e.target.closest('[data-action]');
  if (!el) return;
  const act = el.dataset.action;

  if (act === 'overlay-close') { if (e.target === el) closeModal(); return; }

  switch (act) {
    case 'nav':
      location.hash = '#' + el.dataset.view;
      return;
    case 'sync-now': syncNow(); return;
    case 'logout': doLogout(); return;
    case 'open-job': openJob(el.dataset.key); return;
    case 'job-remove-card': {
      const key = el.dataset.key;
      const j = (state.jobs || []).find(x => x.key === key);
      const label = (j && (j.hwi || j.customer)) || 'this job';
      if (!confirm('Remove ' + label + ' from the board?\n\nIt goes to “Recently deleted”, where you can restore it anytime.')) return;
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(key), { archiveOverride: 'archived' });
        await loadJobs(); render();
        toast('Removed — find it under 🗑 Recently deleted.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'close-modal': closeModal(); return;

    case 'modal-tab': {
      if (state.modalTab === 'planning') collectPlanning();
      else if (state.modalTab === 'procedure') collectProcedure();
      else if (state.modalTab === 'meetings') collectMeetings();
      state.procSendConfirm = null;
      state.modalTab = el.dataset.tab;
      renderModal();
      return;
    }
    case 'proc-generate': {
      if (state.procedureDraft && !confirm('Generate a fresh draft? This replaces the current procedure content for this job.')) return;
      collectSetup();
      const old = state.procedureDraft;
      state.procedureDraft = generateProcedure(state.detail.job, state.detail.planning, state.procedureSetup);
      if (old) { state.procedureDraft.drawings = old.drawings || []; state.procedureDraft.photos = old.photos || []; }   // keep attachments across regenerate
      await saveProcedure(true);
      renderModal();
      toast('Draft procedure generated — edit freely, then print.');
      pullLoadoutEquipment(state.open, { onGenerate: true });   // auto-fill equipment from Shop Master if a loadout exists
      return;
    }
    case 'proc-pull-equipment': pullLoadoutEquipment(state.open, { manual: true }); return;
    case 'draw-view': {
      const old = el.textContent; el.textContent = '…'; el.disabled = true;
      try {
        const r = await api('GET', '/api/job/' + encodeURIComponent(state.open) + '/drawings/' + encodeURIComponent(el.dataset.id));
        window.open(r.url, '_blank');
      } catch (err) { toast(err.message, true); }
      el.textContent = old; el.disabled = false;
      return;
    }
    case 'draw-remove': {
      if (!confirm('Remove this drawing from the procedure?')) return;
      if ($('#procObjective')) collectProcedure();   // keep any in-progress edits
      try {
        const r = await api('DELETE', '/api/job/' + encodeURIComponent(state.open) + '/drawings/' + encodeURIComponent(el.dataset.id));
        if (state.procedureDraft) state.procedureDraft.drawings = r.drawings;
        renderModal();
        toast('Drawing removed.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'photo-view': {
      const ph = state.procedureDraft && (state.procedureDraft.photos || []).find(x => x.id === el.dataset.id);
      if (ph && ph.url) { window.open(ph.url, '_blank'); return; }
      try {
        const r = await api('GET', '/api/job/' + encodeURIComponent(state.open) + '/photos/' + encodeURIComponent(el.dataset.id));
        window.open(r.url, '_blank');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'photo-remove': {
      if (!confirm('Remove this photo from the procedure?')) return;
      if ($('#procObjective')) collectProcedure();
      try {
        const r = await api('DELETE', '/api/job/' + encodeURIComponent(state.open) + '/photos/' + encodeURIComponent(el.dataset.id));
        if (state.procedureDraft) state.procedureDraft.photos = r.photos;
        renderModal();
        toast('Photo removed.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'proc-save': await saveProcedure(); return;
    case 'proc-print': collectProcedure(); printProcedure(); return;
    case 'proc-preview': collectProcedure(); printProcedure(true); return;
    case 'resp-add': collectProcedure(); state.procedureDraft.responsibilities.push(' |  |  | '); renderModal(); return;
    case 'resp-del': collectProcedure(); state.procedureDraft.responsibilities.splice(Number(el.dataset.i), 1); renderModal(); return;
    case 'proc-copy': {
      const p = collectProcedure();
      if (!p) return;
      const m = buildProcedureEmail(p, state.detail.job);
      try { await navigator.clipboard.writeText(m.text); toast('Procedure copied — paste into an email.'); }
      catch (err) { toast('Could not copy automatically.', true); }
      return;
    }
    case 'proc-send': {   // start the guided send flow (preview required first — does NOT send)
      const p = collectProcedure();
      if (!p) { toast('Generate the procedure first.', true); return; }
      const smtpReady = emailReady();
      if (!smtpReady) { toast('Sign in with Microsoft to send from your own mailbox — or use Copy / Print.', true); return; }
      await saveProcedure(true);
      state.procSendConfirm = { previewed: false };
      renderModal();
      return;
    }
    case 'proc-send-preview': {   // step 1 — open the preview, which unlocks the send step
      collectProcSend();
      printProcedure(true);
      state.procSendConfirm = Object.assign({}, state.procSendConfirm, { previewed: true });
      renderModal();
      return;
    }
    case 'proc-send-cancel': { state.procSendConfirm = null; renderModal(); return; }
    case 'proc-send-confirm': {   // step 2 — the PM entered the email; now actually send
      collectProcSend();
      const to = procRecipients();
      if (!to.length) { toast('Enter the customer’s email (or tick a contact) to send to.', true); return; }
      const m = buildProcedureEmail(state.procedureDraft, state.detail.job);
      el.disabled = true; el.textContent = 'Sending…';
      try {
        const r = await api('POST', '/api/job/' + encodeURIComponent(state.open) + '/send', { to, subject: m.subject, html: m.html, body: m.text, kind: 'procedure', msToken: msGraphToken });
        state.procSendConfirm = null;
        state.procedureDraft = JSON.parse(JSON.stringify(r.procedure));
        state.detail.procedure = r.procedure;
        renderModal();
        toast('Procedure sent to ' + r.to.join(', ') + ' ✓');
      } catch (err) {
        el.disabled = false; el.textContent = '✈ Send now';
        toast(err.message, true);
      }
      return;
    }
    case 'plan-init': {
      const t = state.templates;
      state.planningDraft = {
        questions: t.questions.map(q => {
          const jq = { id: q.id, text: q.text, type: q.type || 'text', value: '', answer: '' };
          if (jq.type === 'number') { if (q.unit) jq.unit = q.unit; if (q.units) jq.units = q.units; }
          if (jq.type === 'choice') jq.options = q.options || [];
          return jq;
        }),
        status: 'prepared', notes: '', email: state.detail.job.email || '', sentAt: null
      };
      await savePlanning(true);
      renderModal();
      toast('Standard questions loaded — edit freely, then send.');
      return;
    }
    case 'plan-add-q': collectPlanning(); state.planningDraft.questions.push({ id: 'q' + Math.random().toString(36).slice(2, 8), text: '', type: 'text', value: '', answer: '' }); renderModal(); return;
    case 'plan-del-q': collectPlanning(); state.planningDraft.questions.splice(Number(el.dataset.i), 1); renderModal(); return;
    case 'plan-save': await savePlanning(); renderModal(); return;
    case 'plan-print': printQuestionnaire(); return;
    case 'plan-copy': {
      collectPlanning();
      const m = buildEmail();
      try { await navigator.clipboard.writeText(m.body); toast('Questions copied — paste anywhere.'); }
      catch (err) { toast('Could not copy automatically.', true); }
      return;
    }
    case 'plan-send': {
      collectPlanning();
      const p = state.planningDraft;
      if (!(p.questions || []).some(q => q.text.trim())) { toast('Add at least one question first.', true); return; }
      const to = plannedRecipients();
      if (!to.length) { toast('Tick at least one contact (or type an address).', true); return; }
      const smtpReady = emailReady();
      if (!smtpReady) {
        toast('Sign in with Microsoft to send from your own mailbox — or use “Open email draft”.', true);
        return;
      }
      await savePlanning(true);
      const m = buildEmail();
      el.disabled = true; el.textContent = 'Sending…';
      try {
        const r = await api('POST', '/api/job/' + encodeURIComponent(state.open) + '/send', { to, subject: m.subject, body: m.body, msToken: msGraphToken });
        state.planningDraft = JSON.parse(JSON.stringify(r.planning));
        state.detail.planning = r.planning;
        const row = state.jobs.find(x => x.key === state.open);
        if (row) row.planningStatus = 'sent';
        renderModal();
        toast('Sent to ' + r.to.join(', ') + ' ✓');
      } catch (err) {
        el.disabled = false; el.textContent = '✈ Send now';
        toast(err.message, true);
      }
      return;
    }
    case 'plan-email': {
      collectPlanning();
      const p = state.planningDraft;
      if (!(p.questions || []).some(q => q.text.trim())) { toast('Add at least one question first.', true); return; }
      const to = plannedRecipients();
      if (p.status === 'prepared') p.status = 'sent';
      await savePlanning(true);
      const m = buildEmail();
      const href = 'mailto:' + to.map(encodeURIComponent).join(',') +
        '?subject=' + encodeURIComponent(m.subject) + '&body=' + encodeURIComponent(m.body);
      window.location.href = href;
      renderModal();
      toast('Email draft opened — job marked as “Sent to customer”.');
      return;
    }
    case 'job-hide': {
      const j = state.detail.job;
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { hidden: !j.hidden });
        j.hidden = !j.hidden;
        const row = state.jobs.find(x => x.key === j.key);
        if (row) row.hidden = j.hidden;
        renderModal();
        toast(j.hidden ? 'Job hidden (tick “show hidden” to see it).' : 'Job visible again.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'job-remove': {
      const j = state.detail.job;
      if (!confirm('Remove ' + (j.hwi || j.customer || 'this job') + ' from the board?\n\nIt goes to “Recently deleted”, where you can restore it anytime.')) return;
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { archiveOverride: 'archived' });
        closeModal();
        await loadJobs(); render();
        toast('Removed — find it under 🗑 Recently deleted.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'job-restore': {
      const j = state.detail.job;
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { archiveOverride: '' });
        closeModal();
        await loadJobs(); render();
        toast('Restored to the board.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'open-removed': await openRemoved(); return;
    case 'removed-close': closeRemoved(); return;
    case 'removed-bg': if (e.target === el) closeRemoved(); return;
    case 'removed-restore': {
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(el.dataset.key), { archiveOverride: '' });
        await openRemoved();          // refresh the panel in place
        await loadJobs(); render();   // refresh the board underneath
        toast('Restored to the board.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'removed-purge': {
      if (!confirm('Permanently delete this job?\n\nIt won’t appear on the board or in Recently deleted anymore, and can’t be restored from here.')) return;
      try {
        await api('PATCH', '/api/job/' + encodeURIComponent(el.dataset.key), { purge: true });
        await openRemoved();   // refresh the panel (it drops off the list)
        toast('Permanently removed.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'removed-del-selected': {
      const keys = $$('[data-rmsel]').filter(c => c.checked).map(c => c.value);
      if (!keys.length) { toast('Tick the jobs you want to delete first.', true); return; }
      if (!confirm('Permanently delete ' + keys.length + ' selected job' + (keys.length === 1 ? '' : 's') + '?\n\nThey won’t appear on the board or in Recently deleted anymore.')) return;
      try {
        const r = await api('POST', '/api/removed/purge', { keys });
        await openRemoved();
        toast('Permanently removed ' + (r.purged != null ? r.purged : keys.length) + '.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'removed-del-all': {
      const keys = (state.removed || []).map(x => x.key);
      if (!keys.length) return;
      if (!confirm('Permanently delete ALL ' + keys.length + ' job' + (keys.length === 1 ? '' : 's') + ' in Recently deleted?\n\nThis clears the whole list and can’t be undone from here.')) return;
      try {
        const r = await api('POST', '/api/removed/purge', { keys });
        await openRemoved();
        toast('Permanently removed ' + (r.purged != null ? r.purged : keys.length) + '.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'mtg-action-add':
      collectMeetings();
      state.meetingsDraft.pre = state.meetingsDraft.pre || {};
      state.meetingsDraft.pre.actions = (state.meetingsDraft.pre.actions || []).concat([{ text: '', assignee: '', due: '', done: false }]);
      renderModal();
      return;
    case 'mtg-action-del':
      collectMeetings();
      if (state.meetingsDraft.pre && state.meetingsDraft.pre.actions) state.meetingsDraft.pre.actions.splice(Number(el.dataset.i), 1);
      renderModal();
      return;
    case 'mtg-save': await saveMeetings(); return;
    case 'mtg-report-open':
      collectMeetings();
      state.mtgReportRecipients = meetingReportRecipients();
      openMtgReport();
      return;
    case 'mtg-report-close': closeMtgReport(); return;
    case 'mtg-report-bg': if (e.target === el) closeMtgReport(); return;
    case 'mtg-report-send': {
      const to = [...new Set($$('.rpt-recip')
        .filter(row => { const cb = row.querySelector('[data-rpt-pick]'); return cb && cb.checked; })
        .map(row => { const inp = row.querySelector('[data-rpt-email]'); return inp ? inp.value.trim() : ''; })
        .filter(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)))];
      if (!to.length) { toast('Tick at least one recipient with a valid email address.', true); return; }
      el.disabled = true; el.textContent = emailReady() ? 'Sending…' : 'Confirming Microsoft…';
      const tok = await ensureGraphToken();   // silent Microsoft re-auth if we don't have a send token yet
      if (!tok) { el.disabled = false; el.textContent = '✈ Send report'; toast('Couldn’t confirm your Microsoft sign-in — allow the popup and try again.', true); return; }
      el.textContent = 'Sending…';
      const rep = buildMeetingReport(state.detail);
      try {
        const r = await api('POST', '/api/job/' + encodeURIComponent(state.open) + '/send', { to, subject: rep.subject, html: rep.html, body: rep.text, kind: 'meeting-report', msToken: tok });
        if (r.meetings) { state.detail.meetings = r.meetings; if (state.meetingsDraft) state.meetingsDraft.reportLog = r.meetings.reportLog; }
        closeMtgReport();
        renderModal();
        toast('Meeting report sent to ' + r.to.length + ' recipient' + (r.to.length === 1 ? '' : 's') + ' ✓');
      } catch (err) {
        el.disabled = false; el.textContent = '✈ Send report';
        if (/expired|Mail\.Send|401|Microsoft/i.test(err.message || '')) msGraphToken = null;   // stale token -> re-auth next time
        toast(err.message, true);
      }
      return;
    }
    case 'open-mtg-todo': await openMtgTodo(); return;
    case 'mtg-todo-close': closeMtgTodo(); return;
    case 'mtg-todo-bg': if (e.target === el) closeMtgTodo(); return;
    case 'mtg-todo-open': {
      const key = el.dataset.key; closeMtgTodo();
      await openJob(key);
      if (state.detail) { state.modalTab = 'meetings'; renderModal(); }
      return;
    }
    case 'tpl-add': collectTemplates(); state.templates.questions.push({ id: 'q' + Math.random().toString(36).slice(2, 8), text: '' }); renderTemplates(); return;
    case 'tpl-team-add': collectTemplates(); state.templates.team = (state.templates.team || []).concat([{ name: '', contact: '', email: '' }]); render(); return;
    case 'tpl-team-del': collectTemplates(); state.templates.team.splice(Number(el.dataset.i), 1); render(); return;
    case 'tpl-del': collectTemplates(); state.templates.questions.splice(Number(el.dataset.i), 1); renderTemplates(); return;
    case 'tpl-up': case 'tpl-down': {
      collectTemplates();
      const i = Number(el.dataset.i), j = act === 'tpl-up' ? i - 1 : i + 1;
      const qs = state.templates.questions;
      if (j < 0 || j >= qs.length) return;
      [qs[i], qs[j]] = [qs[j], qs[i]];
      renderTemplates(); return;
    }
    case 'tpl-save': {
      try {
        state.templates = await api('PUT', '/api/templates', collectTemplates());
        render();
        toast(state.view === 'settings' ? 'Team saved.' : 'Saved — new jobs will use the updated questions.');
      } catch (err) { toast(err.message, true); }
      return;
    }

    case 'copy-redirect': {
      try { await navigator.clipboard.writeText(state.settings.redirectUri || 'http://localhost:8743/oauth/callback'); toast('Redirect URI copied.'); }
      catch (err) { toast('Could not copy automatically.', true); }
      return;
    }
    case 'set-connect': {
      try {
        state.settings = Object.assign(state.settings, await api('PUT', '/api/settings', settingsPayload(collectSettings())));
        const r = await api('GET', '/api/zoho/authurl');
        window.location.href = r.url;
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'set-orgs-load': {
      collectSettings();
      try {
        const r = await api('GET', '/api/zoho/orgs');
        state.orgs = r.orgs;
        renderSettings();
        if (!r.orgs.length) toast('Zoho returned no organisations for this account.', true);
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'set-disconnect': {
      if (!confirm('Disconnect from Zoho Books? Your planning data stays — only the connection is removed.')) return;
      try {
        await api('POST', '/api/zoho/disconnect');
        state.settings = await api('GET', '/api/settings');
        state.orgs = null;
        await loadJobs(); render();
        toast('Disconnected. Demo data is shown again.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'set-tab': collectSettings(); collectTemplates(); state.settingsTab = el.dataset.tab; renderSettings(); return;
    case 'set-rule-add': collectSettings(); state.settings.rules.push({ keyword: '', category: 'rental' }); renderSettings(); return;
    case 'set-rule-del': collectSettings(); state.settings.rules.splice(Number(el.dataset.i), 1); renderSettings(); return;
    case 'set-rule-up': case 'set-rule-down': {
      collectSettings();
      const i = Number(el.dataset.i), j2 = act === 'set-rule-up' ? i - 1 : i + 1;
      const rs = state.settings.rules;
      if (j2 < 0 || j2 >= rs.length) return;
      [rs[i], rs[j2]] = [rs[j2], rs[i]];
      renderSettings(); return;
    }
    case 'sm-refresh': { state.smJobs = null; renderShopmaster(); return; }
    case 'sm-toggle-invoiced': { state.smInvoicedCollapsed = (state.smInvoicedCollapsed === false); const b = $('#smBody'); if (b) b.innerHTML = smBodyHtml(); return; }
    case 'set-travel-mode': {
      const hwi = el.dataset.hwi;
      const clicked = el.dataset.mode;          // 'fly' | 'drive'
      const job = state.detail && state.detail.job;
      if (!job || !hwi) return;
      const prev = job.travelMode || null;
      const next = prev === clicked ? null : clicked;   // click the active mode again = clear
      job.travelMode = next;                            // optimistic
      renderModal();
      try {
        const r = await api('POST', '/api/shopmaster/travel-mode', { hwi, mode: next });
        if (next === null) toast('Travel decision cleared for ' + hwi);
        else if (r.published && (r.matched === null || r.matched > 0)) toast((next === 'fly' ? '✈ Flying' : '🚗 Driving') + ' — sent to the travel app for ' + hwi);
        else if (r.published) toast('Saved. No matching job in the travel app yet for ' + hwi + ' — it’ll apply once the job appears there.');
        else toast('Saved locally, but couldn’t reach the travel app: ' + (r.publishError || 'unknown error'), true);
      } catch (err) {
        job.travelMode = prev;                          // revert
        renderModal();
        toast(err.message, true);
      }
      return;
    }
    case 'po-refresh': {
      if (state.leadsMeta && state.leadsMeta.demo) { state.leads = null; renderPo(); return; }
      el.disabled = true; el.textContent = 'Refreshing…';
      try { await api('POST', '/api/leads/sync'); state.leads = null; renderPo(); }
      catch (err) { toast(err.message, true); if (state.view === 'po') renderPo(); }
      return;
    }
    case 'lead-toggle': {
      try {
        await api('PATCH', '/api/lead/' + encodeURIComponent(el.dataset.id), { completedOverride: el.dataset.next === 'true' });
        state.leads = null; renderPo();
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'lead-auto': {
      try {
        await api('PATCH', '/api/lead/' + encodeURIComponent(el.dataset.id), { completedOverride: null });
        state.leads = null; renderPo();
      } catch (err) { toast(err.message, true); }
      return;
    }

    case 'copy-ms-redirect': {
      try { await navigator.clipboard.writeText(state.settings.msRedirectUri || 'http://localhost:8743/ms/callback'); toast('Redirect URI copied.'); }
      catch (err) { toast('Could not copy automatically.', true); }
      return;
    }
    case 'ms-connect': {
      try {
        state.settings = Object.assign(state.settings, await api('PUT', '/api/settings', settingsPayload(collectSettings())));
        const r = await api('GET', '/api/ms/authurl');
        window.location.href = r.url;
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'ms-disconnect': {
      if (!confirm('Disconnect from Microsoft 365? The Invoices page goes back to demo leads — nothing in SharePoint is affected.')) return;
      try {
        await api('POST', '/api/ms/disconnect');
        state.settings = await api('GET', '/api/settings');
        state.msSites = state.msLists = state.msCols = null;
        state.leads = null;
        renderSettings();
        toast('Disconnected from Microsoft 365.');
      } catch (err) { toast(err.message, true); }
      return;
    }
    case 'ms-site-search': {
      const q = ($('#msSiteQ') ? $('#msSiteQ').value : '').trim();
      if (!q) { toast('Type part of the site name first.', true); return; }
      collectSettings();
      el.disabled = true; el.textContent = 'Searching…';
      try {
        const r = await api('GET', '/api/ms/sites?q=' + encodeURIComponent(q));
        state.msSites = r.sites;
        renderSettings();
        if (!r.sites.length) toast('No sites found for “' + q + '” — try another word.', true);
      } catch (err) { toast(err.message, true); el.disabled = false; el.textContent = 'Search'; }
      return;
    }
    case 'ms-lists-load': {
      collectSettings();
      el.disabled = true; el.textContent = 'Loading…';
      try { const r = await api('GET', '/api/ms/lists'); state.msLists = r.lists; renderSettings(); }
      catch (err) { toast(err.message, true); el.disabled = false; el.textContent = 'Choose the list…'; }
      return;
    }
    case 'ms-site-clear': {
      collectSettings();
      Object.assign(state.settings.ms, { siteId: '', siteName: '', listId: '', listName: '' });
      state.msSites = state.msLists = state.msCols = null;
      state.leads = null;
      try { await api('PUT', '/api/settings', settingsPayload(state.settings)); } catch (err) { toast(err.message, true); }
      renderSettings();
      return;
    }
    case 'ms-list-clear': {
      collectSettings();
      Object.assign(state.settings.ms, { listId: '', listName: '' });
      state.msLists = null; state.msCols = null;
      state.leads = null;
      try { await api('PUT', '/api/settings', settingsPayload(state.settings)); } catch (err) { toast(err.message, true); }
      renderSettings();
      return;
    }

    case 'smtp-test': {
      const btn = el;
      try {
        state.settings = Object.assign(state.settings, await api('PUT', '/api/settings', settingsPayload(collectSettings())));
        btn.disabled = true; btn.textContent = 'Sending…';
        const r = await api('POST', '/api/smtp/test', {});
        toast('Test email sent to ' + r.to + ' — check that inbox.');
      } catch (err) { toast(err.message, true); }
      btn.disabled = false; btn.textContent = 'Send test email';
      return;
    }
    case 'sm-test': {
      const btn = el;
      try {
        state.settings = Object.assign(state.settings, await api('PUT', '/api/settings', settingsPayload(collectSettings())));
        btn.disabled = true; btn.textContent = 'Testing…';
        await api('GET', '/api/shopmaster/test');
        state.settings.shopmasterConnected = true;
        state.smJobs = null;
        toast('Shop Master connected ✓ — open the Invoices page.');
        renderSettings();
      } catch (err) { toast(err.message, true); }
      btn.disabled = false; btn.textContent = 'Test connection';
      return;
    }
    case 'set-save': {
      try {
        state.settings = Object.assign(state.settings, await api('PUT', '/api/settings', settingsPayload(collectSettings())));
        await loadJobs();
        state.leads = null; // category rules or PO column mapping may have changed
        renderSettings();
        toast('Settings saved.');
      } catch (err) { toast(err.message, true); }
      return;
    }
  }
});

document.addEventListener('change', async (e) => {
  const el = e.target.closest('[data-change]');
  if (!el) return;
  const what = el.dataset.change;

  if (what === 'stage-filter') { state.stageFilter = el.value; $('#cols').innerHTML = colsHtml(); }
  if (what === 'show-hidden') { state.showHidden = el.checked; $('#cols').innerHTML = colsHtml(); }
  if (what === 'compact') { state.compact = el.checked; try { localStorage.setItem('pmCompact', el.checked ? '1' : '0'); } catch (e) {} $('#cols').innerHTML = colsHtml(); }
  if (what === 'po-show-done') { state.poShowCompleted = el.checked; const b = $('#poBody'); if (b) b.innerHTML = poBodyHtml(); }
  if (what === 'sm-show-all') { state.smShowAll = el.checked; const b = $('#smBody'); if (b) b.innerHTML = smBodyHtml(); }
  if (what === 'setup-jobtype' || what === 'setup-static') {
    collectSetup();
    const panel = $('#setupPanel');
    if (panel && state.detail) panel.outerHTML = setupPanelHtml(state.procedureSetup, state.detail.job);
    return;
  }
  if (what === 'mtg-held') {
    const which = e.target.id === 'mtgpreHeld' ? 'pre' : 'post';
    const dateEl = document.getElementById('mtg' + which + 'Date');
    if (e.target.checked && dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0, 10);
    return;
  }
  if (what === 'mtg-action-done') {
    const row = e.target.closest('[data-mtg-action-row]');
    if (row) row.classList.toggle('done', e.target.checked);
    const rows = $$('[data-mtg-action-row]');
    const open = rows.filter(r => { const c = r.querySelector('[data-ma-done]'); return !(c && c.checked); }).length;
    const cnt = document.querySelector('.ma-count');
    if (cnt) cnt.textContent = open + ' open · ' + rows.length + ' total';
    return;
  }
  if (what === 'rpt-selall') {
    const on = e.target.checked;
    $$('[data-rpt-pick]').forEach(c => { c.checked = on; });
    return;
  }
  if (what === 'removed-selall') {
    const on = e.target.checked;
    $$('[data-rmsel]').forEach(c => { c.checked = on; });
    return;
  }
  if (what === 'draw-file') {
    const file = el.files && el.files[0];
    el.value = '';                                  // let the same file be re-picked later
    await uploadDrawingFile(file);
    return;
  }
  if (what === 'photo-file') {
    const file = el.files && el.files[0];
    el.value = '';
    if (!file) return;
    if (!/^image\//.test(file.type || '') && !/\.(jpe?g|png|webp|heic|heif)$/i.test(file.name)) { toast('Please choose an image.', true); return; }
    if ($('#procObjective')) collectProcedure();
    toast('Processing ' + file.name + '…');
    let blob;
    try { blob = await downscaleImage(file, 1600, 0.82); }
    catch (err) { toast(err.message || 'Could not process that image.', true); return; }
    try {
      const opt = { method: 'POST', headers: { 'Content-Type': 'image/jpeg', 'x-filename': file.name }, body: blob };
      if (authToken) opt.headers['Authorization'] = 'Bearer ' + authToken;
      const resp = await fetch('/api/job/' + encodeURIComponent(state.open) + '/photos', opt);
      const j = await resp.json().catch(() => ({}));
      if (resp.status === 401) { renderLogin('Your session has expired — please sign in again.'); return; }
      if (!resp.ok) throw new Error(j.error || ('Upload failed (' + resp.status + ')'));
      if (state.procedureDraft) state.procedureDraft.photos = j.photos;
      renderModal();
      toast('Photo added.');
    } catch (err) { toast(err.message || 'Upload failed.', true); }
    return;
  }
  if (what === 'tpl-type') { collectTemplates(); renderTemplates(); }
  if (what === 'resp-pick') {
    const opt = el.selectedOptions[0];
    const row = el.closest('[data-resp-row]');
    if (row && opt && opt.dataset.name) {
      row.querySelector('[data-rname]').value = opt.dataset.name;
      if (opt.dataset.contact) row.querySelector('[data-rcontact]').value = opt.dataset.contact;
      const comp = row.querySelector('[data-rcompany]');
      if (!comp.value.trim()) comp.value = 'Hydro-Wates';
    }
    el.value = '';
  }

  if (what === 'ms-site-pick' && el.value) {
    const opt = el.selectedOptions[0];
    collectSettings();
    Object.assign(state.settings.ms, { siteId: el.value, siteName: opt.dataset.name || opt.textContent, listId: '', listName: '' });
    state.msLists = null; state.msCols = null; state.leads = null;
    try {
      await api('PUT', '/api/settings', settingsPayload(state.settings));
      const r = await api('GET', '/api/ms/lists');
      state.msLists = r.lists;
      const hit = r.lists.filter(l => /lead/i.test(l.name));
      if (hit.length === 1) {
        Object.assign(state.settings.ms, { listId: hit[0].id, listName: hit[0].name });
        await api('PUT', '/api/settings', settingsPayload(state.settings));
        toast('Found “' + hit[0].name + '” — loading its columns…');
      }
      renderSettings();
    } catch (err) { toast(err.message, true); renderSettings(); }
  }

  if (what === 'ms-list-pick' && el.value) {
    const opt = el.selectedOptions[0];
    collectSettings();
    Object.assign(state.settings.ms, { listId: el.value, listName: opt.dataset.name || opt.textContent });
    state.msCols = null; state.leads = null;
    try { await api('PUT', '/api/settings', settingsPayload(state.settings)); } catch (err) { toast(err.message, true); }
    renderSettings();
  }

  if (what === 'job-wll' && state.detail) {
    const j = state.detail.job;
    const wllEl = $('#jobWll'), unitEl = $('#jobWllUnit');
    const wll = wllEl ? wllEl.value : '';
    const wllUnit = unitEl ? unitEl.value : 't';
    try {
      await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { wll: wll, wllUnit: wllUnit });
      j.wll = (wll === '' ? null : Number(wll));
      j.wllUnit = wllUnit;
      const tl = $('#testLoad'); if (tl) tl.innerHTML = testLoadText(j);
      toast('WLL saved.');
    } catch (err) { toast(err.message, true); }
  }

  if (what === 'job-stage' && state.detail) {
    const j = state.detail.job;
    try {
      await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { stage: el.value });
      j.stage = el.value;
      const row = state.jobs.find(x => x.key === j.key);
      if (row) row.stage = el.value;
      toast('Stage updated.');
    } catch (err) { toast(err.message, true); }
  }
  if (what === 'job-multi' && state.detail) {
    const j = state.detail.job;
    try {
      await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { multiInvoice: el.checked });
      j.multiInvoice = el.checked;
      const row = state.jobs.find(x => x.key === j.key);
      if (row) row.multiInvoice = el.checked;
      await loadJobs(); render();   // board membership can change (an invoiced multi-invoice job stays / leaves)
      toast(el.checked ? 'Marked multi-invoice — it stays on the board through invoicing.' : 'No longer marked multi-invoice.');
    } catch (err) { toast(err.message, true); }
  }

  if (what === 'job-cat' && state.detail) {
    const j = state.detail.job;
    try {
      await api('PATCH', '/api/job/' + encodeURIComponent(j.key), { categoryOverride: el.value });
      const d = await api('GET', '/api/job/' + encodeURIComponent(j.key));
      state.detail.job = d.job;
      const row = state.jobs.find(x => x.key === j.key);
      if (row) { row.category = d.job.category; row.categoryOverridden = d.job.categoryOverridden; }
      renderModal();
      toast(el.value ? 'Moved to ' + el.value + '.' : 'Back to automatic category.');
    } catch (err) { toast(err.message, true); }
  }

  if (what === 'org-pick') {
    const opt = el.selectedOptions[0];
    if (!el.value) return;
    state.settings.orgId = el.value;
    state.settings.orgName = opt.textContent;
    try {
      await api('PUT', '/api/settings', settingsPayload(collectSettings()));
      renderSettings();
      toast('Organisation saved — press Sync now.');
    } catch (err) { toast(err.message, true); }
  }
});

document.addEventListener('input', (e) => {
  if (e.target.id === 'searchBox') {
    state.search = e.target.value;
    $('#cols').innerHTML = colsHtml();
  }
  if (e.target.id === 'poSearch') {
    state.poSearch = e.target.value;
    const b = $('#poBody');
    if (b) b.innerHTML = poBodyHtml();
  }
  if (e.target.id === 'smSearch') {
    state.smSearch = e.target.value;
    const b = $('#smBody');
    if (b) b.innerHTML = smBodyHtml();
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && state.open) closeModal();
});

// Drag-and-drop a PDF straight onto the drawings box to attach it.
document.addEventListener('dragover', (e) => {
  const box = e.target.closest && e.target.closest('.draw-box');
  if (box) { e.preventDefault(); box.classList.add('drag'); }
});
document.addEventListener('dragleave', (e) => {
  const box = e.target.closest && e.target.closest('.draw-box');
  if (box && !box.contains(e.relatedTarget)) box.classList.remove('drag');
});
document.addEventListener('drop', async (e) => {
  const box = e.target.closest && e.target.closest('.draw-box');
  if (!box) return;
  e.preventDefault(); box.classList.remove('drag');
  const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) await uploadDrawingFile(file);
});

/* ---------------- boot ---------------- */
window.addEventListener('hashchange', () => { if (state.open) closeModal(); routeFromHash(); render(); });

(async function init() {
  if (await handleMsPopupCallback()) return;   // we're the Microsoft re-auth popup — handled; don't boot the app
  try { state.compact = localStorage.getItem('pmCompact') === '1'; } catch (e) {}
  const bt = document.getElementById('buildTag'); if (bt) bt.textContent = BUILD;
  const authed = await initAuth();
  if (!authed) return;   // login screen is showing — stop until the user signs in
  routeFromHash();
  try {
    await loadAll();
    render();
    if (state.meta.sync && state.meta.sync.running) startSyncPoll();
  } catch (e) {
    $('#view').innerHTML = '<div class="panel"><h2>Could not load</h2><p class="hint">' + esc(e.message) + '</p></div>';
  }
})();
