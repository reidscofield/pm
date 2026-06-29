# Deploying the Hydro-Wates PM app to Vercel

The app is ready for Vercel. Storage is in Supabase, login uses Supabase Auth, and
the serverless entry is `api/index.js`. These are the remaining steps — most are
clicks in the Vercel and Supabase dashboards (which you control).

## 0. One-time prerequisite (done if you've tested Supabase mode)
- Run the `pm_app_state` table SQL in Supabase (already provided).
- Confirm local Supabase mode works: `http://localhost:8743/api/storage` → `{"mode":"supabase"}`.

## 1. Put the project in a GitHub repo
From `C:\Users\Kaylee Kim\hydrowates-pm`:
```
git init
git add .
git commit -m "Hydro-Wates PM app — ready for Vercel"
```
Create a new **private** repo on GitHub, then:
```
git remote add origin https://github.com/<you>/hydrowates-pm.git
git branch -M main
git push -u origin main
```
(`.gitignore` keeps `.env` and `data/` out, so no secrets are pushed.)

## 2. Import the repo into Vercel
Vercel dashboard → **Add New… → Project** → import the GitHub repo.
- Framework preset: **Other** (no build step needed).
- Leave Build/Output settings default. Deploy.

## 3. Set the Environment Variables (Vercel → Project → Settings → Environment Variables)
| Name | Value |
|------|-------|
| `SUPABASE_URL` | `https://vpdcikiyaifppkkantrb.supabase.co` |
| `SUPABASE_SERVICE_KEY` | your Supabase **service_role** secret |
| `REQUIRE_AUTH` | `true` |
| `ALLOWED_EMAIL_DOMAIN` | `hydrowates.com` (optional; this is the default) |

Apply them to **Production** (and Preview if you want). Then **redeploy** so they take effect.

## 4. Point Supabase login at the Vercel URL
Supabase dashboard → **Authentication → URL Configuration → Redirect URLs** → add your
Vercel URL, e.g. `https://hydrowates-pm.vercel.app` (and keep `http://localhost:8743`
for local). This is what makes "Sign in with Microsoft" return to the app.

## 5. Test
- Open the Vercel URL → you should get the **login screen**.
- Sign in with a Hydro-Wates account (same as the travel app).
- Confirm the dashboard loads (it reads the same Supabase data as local).

## Notes
- **Zoho / Microsoft email** keep working on Vercel using the connections already
  saved (migrated into Supabase) — no need to reconnect. Re-connecting from the
  deployed URL would need its redirect URI added in Zoho/Entra; not required for
  normal use.
- The dashboard and Vercel share ONE Supabase dataset, so local and production stay
  in sync.
- Background **Sync** runs within Vercel's 300s function limit (Pro). If a future
  sync ever times out, we'll switch it to a chunked/cron approach.
