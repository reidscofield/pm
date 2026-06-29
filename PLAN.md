# Project Management App — Plan

*v3 — 15 June 2026 · Status: phase 1 built; full PM workflow mapped (phases 2+ not yet built)*

> v1 (11 June) planned a generic standalone tracker with manually-created
> projects. Superseded on 12 June: jobs now come **from Zoho Books** instead of
> being typed in by hand, which means the app runs as a small local server
> (double-click `Start Project Manager.bat`) rather than a plain HTML file —
> a browser page on its own isn't allowed to talk to Zoho's API.
>
> **v3 (15 June)** maps the project manager's full end-to-end job workflow
> (PO → planning → loadout → procedure → travel → certificate → invoice) and how
> the app should grow to cover it. **Objective: same-day invoicing.** Phase 1
> stays as built; the new map is the "Target workflow" section below.

## Goal

One place to see and run all jobs:

1. every job created in **Zoho Books** appears automatically,
2. split into **Rental / Service / Sales**,
3. each job carries its workflow — starting with **planning questionnaires**
   sent to the customer.

## Built in phase 1 (12 June 2026)

- **Zoho Books sync** — pulls Estimates, Sales Orders, Invoices (each
  toggleable; Projects optional). Read-only. Guided one-time connection in
  Settings; demo data until connected.
- **Rental / Service / Sales dashboard** — three columns, auto-categorised by
  editable keyword rules (first match wins) + per-job manual override. Search,
  stage filter, hide jobs.
- **Job stages** — New → Planning → Scheduled → In progress → Complete / On hold.
- **Planning questionnaires** — master list of the 7 standard questions
  (editable under "Planning questions"), copied into each job where it can be
  tweaked per customer; one-click pre-filled email draft; answers + internal
  notes recorded per job; status tracked (Prepared → Sent → Answers received).
- **Send to contacts on file** *(added later on 12 June)* — each job's
  Planning tab lists the contact persons stored against that customer in Zoho
  Books (primary pre-ticked); **Send now** emails the questions to the ticked
  contacts directly from the app (one-time SMTP setup in Settings → Email
  sending, with send-test button); per-job send history. Email-draft and
  copy buttons remain as the no-setup fallback.
- **PO tracker** *(added later on 12 June)* — new page fed by the SharePoint
  **Lead List** (Microsoft 365 connection via Entra app, read-only, Settings
  wizard with site/list pickers and column mapping). Shows every lead whose
  PO condition is met; each is matched against Zoho Books **invoices** by PO
  number ↔ invoice reference (or invoice number) and flips to **Completed**
  automatically when invoiced. Manual completed/open overrides per row.
  Lead refresh is bundled into the main Sync and has its own Refresh button.
- **Auto-hide finished jobs** *(18 June; Archive page removed 19 June)* — a job is
  treated as finished (hidden from the dashboard) when **stage = Complete** or it
  carries a **closed/invoiced Zoho status** (paid, closed, invoiced, void, declined,
  expired, rejected). This keeps the board to active jobs only — of 532 synced, 481
  finished are hidden, leaving ~51 active. The separate **"Archive" page + manual
  Archive/Restore buttons were removed 19 June** at the user's request (they just
  wanted the active board, no separate view). The server still computes the
  `archived` flag (`archiveState`) and the dashboard filters on it; there is no
  in-app way to browse the hidden jobs now (they live in Zoho).

### The 7 standard questions (master copy lives in `data/templates.json`)

1. Are we going to test the equipment to 125% of WLL?
2. Are we working in metric tonnes or short tons?
3. What kind of water source is available on site, and what is the connection type?
4. How far away from the test site is the water source?
5. How much headroom do we have below the hook?
6. What are the dimensions of the area we are testing in?
7. Will we be function testing at 100%?

## Target workflow — the PM's full job pipeline

**Objective: invoice the customer the same day.** That only works if everything
that *can* be prepared before the crew travels is locked in advance, so the only
new input onsite is the measured test results — which then flow straight into the
certificate (within 4 h) and the invoice (same day).

### The job record (the spine)

One record per job, created at intake and carried to invoice. Fields fill in as
it moves through the stages:

`PO# · customer · site · value · contacts · answers · loadout · procedure ·
travel/freight · results · certificate · invoice`

Today these live in ~7 separate tools (Power Apps, "the other PM app", Word,
email, a travel site, a cert system, Zoho), so every handoff is a re-key. The
whole streamline is: **one record, data captured once, cascading downstream.**

### The cascade

The customer's answers to the standard questions are the same numbers that drive
every later step, so they're entered once (as **structured fields**, not free text):

```
answers (WLL, units, water source + distance, headroom, area, function-test %)
  → loadout (bag count/size, pumps, hoses, slings, total weight)
  → procedure (test load = 125% WLL, units, sequence, equipment)
  → certificate (finalized procedure + measured results)
  → invoice (PO value, pre-drafted at intake)
```

### Stage-by-stage

**Phase 1 — Intake**

| # | Step | Needs | Produces | Status in app |
|---|---|---|---|---|
| 1 | PO in + create job | PO#, customer, value, scope, site | the job record | partial — Zoho sync + PO tracker |
| 2 | Customer questions | 7 standard Qs, contacts | the answers (above) | built + structured (18 Jun) — typed answers (text/number+unit/choice/yes-no); email shows type hints; Q3 (source→choice + connection) & Q6 (→ length × width) split into structured fields; answered sheet prints to PDF for the job pack |

**Phase 2 — Pre-stage** *(locked before the crew leaves)*

| # | Step | Needs | Produces | Status in app |
|---|---|---|---|---|
| 3 | Loadout | WLL/test load + answers | equipment list + total weight | **done in "Shop Master" — do NOT rebuild.** App's job is to capture its output + feed it the inputs. WLL + auto test-load (×1.25) now captured on the job (18 Jun); WLL entered internally, not asked of customer |
| 4 | Procedure | answers + loadout | load-test procedure | **built 18 Jun** — Procedure tab generates a full draft from WLL/test load + answers + standard boilerplate (objective, coordination, responsibilities, 5.1–5.4 steps with computed fill tonnages & water distance); editable; prints to PDF on the letterhead. Modelled on HWI-26-223. Equipment list left blank → from Shop Master loadout. **Responsibilities (22 Jun):** structured rows with a **team-roster dropdown** (Hydro-Wates staff + contacts, managed under Planning questions → Team contacts) that auto-fills name/company/contact |
| 5 | Prejob review | draft procedure | finalized + signed-off | partial — PM edits the draft + status (draft/reviewed/final) on the Procedure tab |
| 6 | Send procedure | finalized doc + contacts | sent record | **built 19 Jun** — "Send to customer" on the Procedure tab: ticks contacts on file + extra recipient, emails the procedure as a **formatted HTML email** (logo from the public website, all sections), marks status Final, records send history. Reuses SMTP. Copy + Print/PDF fallbacks. (No PDF attachment — zero-dependency app) |
| 7 | Travel/freight plan | site, dates, loadout weight | travel + freight plan | gap |
| 8 | Book travel | travel plan | confirmations | external — app only stores the result |

**Phase 3 — Execute & close** *(the objective)*

| # | Step | Needs | Produces | Status in app |
|---|---|---|---|---|
| 9 | Certificate (within 4 h) | finalized procedure + measured results + equipment serials | test certificate | gap — separate/manual today, 4 h is a struggle. **Biggest payoff** |
| 10 | Invoice (same day) | PO value + complete + variations | invoice in Zoho | tracked only; plan = app **pre-drafts**, PM reviews + posts in Zoho |

### Decisions captured (15 June)

- **Procedures:** templating partly exists, lots still manual → goal is to auto-fill
  a draft from answers + loadout, leaving the PM to fine-tune.
- **Certificates:** separate/manual today and the 4 h SLA is a struggle → highest-value
  gap to close; the cert is mostly "finalized procedure + measured results".
- **Invoicing:** app **pre-drafts** the invoice; the PM reviews and posts it in Zoho
  (keeps Zoho effectively read-only — lower risk than auto-creating).
- **Travel booking** stays external; the app tracks the plan + confirmations only.

### Recommended build sequence (dependency order, not yet agreed)

1. ~~**Structured answers** — turn the Planning answers into typed fields. Unblocks
   the whole cascade. Smallest change, highest leverage.~~ **DONE (18 Jun 2026).**
   Each master question now carries an answer **type** (text / number+unit /
   choice / yes-no), editable under *Planning questions*; per-job answers are
   captured through the matching typed input and stored as a structured `value`
   (plus a derived display string for the email/back-compat). Defaults seeded
   for the 7. Also built 18 Jun: customer email shows answer-type hints; Q3/Q6
   split into structured fields; answered questionnaire prints to PDF for the
   job pack; **WLL + auto test-load (×1.25) captured on the job** (the input to
   the loadout, entered internally).
2. ~~**Loadout builder**~~ **— NOT being built.** Loadouts are done in the
   existing **Shop Master** program; rebuilding would re-fragment the workflow.
   The app instead **captures Shop Master's output** (kit list / total weight →
   freight + job pack) and **feeds it the inputs** (WLL/test load, now captured).
   TODO: learn how Shop Master exports/connects to wire the capture (see open Qs).
3. ~~**Procedure generator**~~ **DONE (18 Jun)** — Procedure tab: one-click draft from
   WLL/test load + planning answers + standard boilerplate (matches HWI-26-223),
   fully editable, prints to PDF on the Hydro-Wates letterhead. Equipment section
   left blank pending the Shop Master loadout feed.
4. **Certificate** — results entry + cert render from the finalized procedure, plus a
   4 h timer/alert from "test complete". (The objective's bottleneck.)
5. **Invoice pre-draft** — build the draft from the PO value at intake; "raise draft"
   on cert issue; PM posts in Zoho.
6. **Pre-stage glue** — review/sign-off status, send-procedure (reuse SMTP),
   travel/freight tracking, and an expanded stage model / per-job checklist.

The current stage model (New → Planning → Scheduled → In progress → Complete /
On hold) would grow a **Prep** stage (loadout/procedure/review/sent) between
Planning and Scheduled, and **Cert issued → Invoiced** milestones at the close —
or stay coarse with a per-job checklist (the phase-3 idea below).

### Earlier phase ideas (still relevant, folded into the above)

| Phase | Idea |
|---|---|
| 2 | Scheduling: job dates, calendar/timeline view, deadline surfacing |
| 3 | Per-stage checklists beyond planning (kit prep, test day, sign-off, certs) |
| 4 | One-click status reports; .ics calendar export |

## Where things live

- `server.js` — the app (zero dependencies, Node 18+)
- `public/` — the web pages
- `data/` — settings, question templates, planning answers, Zoho cache
- `README.md` — how to start, connect Zoho, troubleshoot

## Open questions

1. Which Zoho Books record really marks "a job exists" for Hydro-Wates —
   accepted estimate? sales order? invoice? (All three shown for now;
   toggleable in Settings.)
2. Is there a custom field in Zoho Books that already says rental/service/sales?
   If yes, a rule on that field beats keyword guessing.
3. ~~Should answered questionnaires export to PDF/Word for the job pack?~~
   **RESOLVED (18 Jun):** yes — Planning tab has a "Print / save PDF" button.
4. **Intake trigger:** is the PO the moment a job is created, or the Zoho
   estimate/sales order? (Affects step 1 — does the PM key the PO, or does it
   come from the SharePoint lead list / Zoho?)
5. ~~**Equipment catalog**~~ **MOOT (18 Jun):** loadouts (and the kit catalog)
   live in **Shop Master** — the app won't hold a catalog. New question: **how
   does Shop Master export/connect** (file / API / screen-only), and which
   outputs (total weight, bag count, a document?) should the app capture and
   feed downstream? Awaiting the PM (out of office 18 Jun).
6. **Certificate format:** is there an accreditation-body template the cert must
   reproduce exactly, or is it a Hydro-Wates in-house format? (Shapes step 9.)
7. **Structured answers:** what's the typed shape behind each standard question
   (e.g. WLL = number + unit, water source = choice, function-test = yes/% )?
   This is the keystone that lets answers cascade — worth nailing down first.
