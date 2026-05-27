# Deployment Guide — Braudit Audit Tool

This guide takes you from "I have the project folder" to "my team can use it
at a private URL". The whole process is GitHub web UI and Streamlit Cloud web
UI — **no command line required**. Total time: about 15 minutes.

You will need:
- A GitHub account (free) — you said yours is `ponnyg1980@gmail.com`
- A Streamlit Cloud account (free, sign in with GitHub)
- The `braudit_audit_tool` folder I packaged for you

---

## Step 1 — Create a GitHub repository (5 min)

1. Go to https://github.com/new while logged in as `ponnyg1980@gmail.com`.
2. Fill in:
   - **Repository name:** `braudit-audit-tool`
   - **Description:** `Internal Streamlit tool for Braudit Steps 2–5`
   - **Public** (required for free Streamlit Cloud; no client data ever
     lives in the repo, so this is safe)
   - Tick **Add a README file** — we'll overwrite it in a minute
   - Leave `.gitignore` as **None** and license as **None**
3. Click **Create repository**.

## Step 2 — Upload the project files (5 min)

1. On the new repo's page, click **Add file → Upload files** (the button is
   near the top right of the file list).
2. Drag the contents of the `braudit_audit_tool` folder into the upload
   area. Make sure you drag the **contents**, not the folder itself — you
   want `app.py` and `requirements.txt` at the root of the repo, not inside
   a `braudit_audit_tool/` sub-folder.
3. Important — also upload the `.streamlit/` folder (drag it in). GitHub's
   web UI shows folders fine.
4. **Do NOT upload `.streamlit/secrets.toml`** if you used the development
   password. Upload `.streamlit/secrets.toml.example` instead and we'll set
   the real password directly in Streamlit Cloud (see Step 4).
5. At the bottom of the upload page, write a commit message like
   `Initial commit` and click **Commit changes**.

## Step 3 — Create a Streamlit Cloud account (2 min)

1. Go to https://share.streamlit.io
2. Click **Sign in** → **Continue with GitHub**
3. Authorise Streamlit Cloud to read your GitHub repos.

## Step 4 — Deploy the app (3 min)

1. From the Streamlit Cloud dashboard, click **Create app** →
   **Deploy a public app from GitHub**.
2. Fill in:
   - **Repository:** `ponnyg1980/braudit-audit-tool`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL:** pick something like `braudit-audit-tool` — your URL will
     be `https://braudit-audit-tool.streamlit.app`
3. Click **Advanced settings** before deploying:
   - Under **Secrets**, paste:
     ```toml
     app_password = "YOUR-CHOSEN-PASSWORD-HERE"
     ```
     Pick something memorable but not guessable. This is what your team
     will type to access the tool.
4. Click **Deploy**.

The first deploy takes about 2 minutes (Streamlit installs the requirements).
Once it's up, you'll see the login screen at your chosen URL.

## Step 5 — Test it (2 min)

1. Visit your URL: `https://braudit-audit-tool.streamlit.app`
2. Enter your password.
3. Upload a real scraped-results spreadsheet, fill the form, click
   **Run Audit**, download the docx, confirm it looks right.

## Step 6 — Share with your team

1. Send the URL and the password to your audit team in whatever channel you
   normally use (Slack, email, internal wiki).
2. That's it. They don't need accounts. They just visit the URL.

---

## Updating the tool later

Streamlit Cloud auto-redeploys on every git push to `main`. So if you (or
someone helping you) updates `filters.py` or `app.py`:

- **Via the GitHub web UI:** open the file, click the pencil icon, edit,
  commit. Streamlit picks it up automatically within 60 seconds.
- **Via Claude Code on your laptop:** make changes locally, commit, push.
  Same auto-redeploy.

If something breaks after a deploy, the Streamlit Cloud dashboard shows the
build log so you can see the error.

---

## Changing the password

Go to your app on https://share.streamlit.io → **Settings** → **Secrets** →
edit `app_password`. Save and it takes effect in about 30 seconds.

## When you outgrow v1

When you're ready to add Step 6 (the forensic appendix):

1. You'll need an Anthropic API key.
2. Add it as a second secret in Streamlit Cloud (`anthropic_api_key = "sk-..."`).
3. We'd extend `pipeline/` with a `forensic.py` module that calls the
   Claude Agent SDK. The Run Card prompt we built earlier becomes the
   system prompt.
4. The Run Audit button would gain a checkbox: "Also run forensic appendix
   (~3 minutes, ~£0.50 in API costs)."

That's the v2 lift, and we'd do it in a separate working session.

---

## Troubleshooting

**Build fails with "ModuleNotFoundError"** — the requirements.txt didn't
upload, or the `pipeline/` folder didn't upload. Check the repo on GitHub.

**App loads but uploads time out** — Streamlit Cloud's free tier has a
~200MB upload limit. The scraped-results .xlsx is typically well under 1MB,
so this should never bite you.

**Multiple team members hitting the app simultaneously** — Streamlit Cloud
spins up one process per session. The free tier allows ~10 concurrent users
which is plenty for an internal team. If you need more, the paid tier is
$20/month.

**You forgot the password** — go to Streamlit Cloud → Settings → Secrets,
read it (or change it).

---

*Deployment guide v1.0 · 21 May 2026 · For The Trademark Helpline / Braudit.*
