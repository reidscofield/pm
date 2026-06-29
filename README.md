# Hydro-Wates Project Manager

A small local web app that pulls your jobs from **Zoho Books**, sorts them into
**Rental / Service / Sales**, and runs your **planning questionnaire** workflow
(standard questions → email to customer → record the answers).

Everything runs on your own PC. Nothing is uploaded anywhere; the app only
*reads* from Zoho Books and never changes anything there.

---

## Starting the app

Double-click **`Start Project Manager.bat`**.

A black window opens (that's the app running — minimise it, don't close it) and
your browser opens `http://localhost:8743`.

To stop the app, close the black window.

> Requires Node.js (already installed on this PC). On a new PC, install the LTS
> version from <https://nodejs.org> first.

## Demo mode

Until you connect Zoho Books the dashboard shows **demo jobs** so you can try
everything: open a job, load the planning questions, draft the email, record
answers. Demo jobs are clearly flagged and disappear once you connect.

## Connecting Zoho Books (one-time, ~5 minutes)

1. Go to **<https://api-console.zoho.com>** and sign in with the Zoho account.
2. **Add Client** → **Server-based Applications**.
3. Enter:
   - Homepage URL: `http://localhost:8743`
   - Authorized Redirect URI: `http://localhost:8743/oauth/callback`
4. Copy the **Client ID** and **Client Secret** it creates.
5. In the app: **Settings** → pick your Zoho region → paste Client ID and
   Secret → **Save & Connect** → approve access on the Zoho page.
6. Back in Settings, check the right **organisation** is selected, then press
   **Sync**.

If your Zoho account is hosted in Europe/India/etc., make sure the region
dropdown matches (it changes which Zoho servers the app talks to).

## How jobs are categorised

Every synced job is checked against the **keyword rules** in Settings — the
keywords are matched against the job's line items, reference, notes and
customer name. **The first rule that matches wins**, top to bottom. Jobs that
match nothing go to the default category (Sales out of the box).

Out of the box: `rental`, `hire`, `rent` → Rental; `proof load`, `load test`,
`test`, `inspection`, `certif`, `calibrat`, `service` → Service; everything
else → Sales.

Got one wrong? Open the job → **Details** tab → change the **Category**
dropdown. That override sticks and survives future syncs.

## The planning workflow

1. Open a job → **Planning** tab → **Load standard questions**.
2. The app looks up the **contacts on file** for that customer in Zoho Books
   and lists them with tick-boxes (the primary contact is pre-ticked). Add a
   one-off address in *Extra recipient* if needed.
3. Edit, add or remove questions for this particular customer if needed.
4. Send:
   - **Send now** — the app emails the questions to the ticked contacts
     directly (needs the one-time *Email sending* setup below), or
   - **Open email draft** — your email program opens pre-addressed and
     pre-filled, or **Copy questions** to paste anywhere.
   Either way the job is marked *Sent to customer*; direct sends are kept in
   a per-job send history.
5. When the customer replies, type their answers under each question and set
   the status to *Answers received*.

## Email sending setup (optional, for "Send now")

Settings → **Email sending**. For Microsoft 365 / Outlook accounts:
server `smtp.office365.com`, port `587`, user = your full email address, and
your normal password — though many companies require an **app password**
(create one at account.microsoft.com → Security) or need IT to enable
*Authenticated SMTP* for the mailbox. Use **Send test email** to check it
works — the app sends a test message to your own address. If your company
blocks SMTP entirely, just keep using **Open email draft**.

The master question list and the email wording live under **Planning
questions** in the top bar. Editing them affects newly planned jobs only —
questions already attached to a job are kept as they are.

## The PO tracker (SharePoint lead list)

The **PO tracker** page shows every lead in your SharePoint **Lead List**
that has received a PO, and marks each one **completed automatically as soon
as a Zoho Books invoice carries its PO number** (in the invoice's reference
field). Open jobs sit at the top; invoiced ones drop to the bottom in green
with the matching invoice shown. "Mark completed" / "Reopen" buttons override
the automatic result for odd cases, and ↺ puts a row back on automatic.

One-time connection (Settings → **SharePoint lead list**):

1. Someone with Azure access (you or IT) registers an app:
   **portal.azure.com** → Microsoft Entra ID → App registrations → New
   registration → name it, pick *this organisation only*, add a redirect URI
   under platform **Public client/native (mobile & desktop)** using the
   address shown in the app, then under API permissions add **Microsoft
   Graph → Delegated → Sites.Read.All**.
2. Paste the **Application (client) ID** into Settings → **Save & Connect**
   and sign in with your normal Microsoft 365 account.
3. Search for your SharePoint site, pick the **Lead List**, and tell the app
   which columns hold the **PO**, the **company**, and (optionally) value and
   date — it guesses these for you.

The app only ever *reads* the list. The PO column condition is flexible: "has
any value" (a PO number typed in), "is ticked Yes", or "equals some text"
(e.g. a Status column set to "PO received").

## What counts as a "job"?

By default the app pulls **Estimates, Sales Orders and Invoices** (Zoho Books
has no single "job" record, so you choose which of these represent jobs for
you). Toggle them in Settings. If an estimate later becomes an invoice you'll
see both — use **Hide job** on the one you don't want to track.

## Where your data lives

| File | Contents |
|---|---|
| `data/jobs.json` | Your planning answers, stages, category overrides |
| `data/leads.json` | Manual completed/open overrides on the PO tracker |
| `data/templates.json` | The standard questions and email wording |
| `data/settings.json` | Settings **including your Zoho and Microsoft 365 connection keys** |
| `data/zoho-cache.json` | Local copy of the synced Zoho records and the SharePoint lead list |

The folder lives in OneDrive, so all of this is backed up automatically.
Note that `settings.json` contains the key that grants read access to your
Zoho Books — don't email or share that file.

## Troubleshooting

- **Browser says "can't connect"** — the app isn't running; double-click
  `Start Project Manager.bat`.
- **"The app looks like it is already running"** — it is! Just open
  <http://localhost:8743>.
- **Sync error about rate limits** — Zoho briefly throttled us; wait a minute
  and press Sync again.
- **Wrong/expired connection** — Settings → Disconnect, then connect again.
- **A job is missing** — it may be older than the sync window; increase
  "how far back to fetch" in Settings and sync again.
