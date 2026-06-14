# SiteDown Web App — Deployment Guide

## Project structure

```
sitedown_web/
├── app.py            ← Streamlit UI
├── processor.py      ← All data processing logic
├── requirements.txt  ← Dependencies (auto-installed by Streamlit Cloud)
└── DEPLOY.md         ← This file
```

---

## Step 1 — Push to GitHub (one-time setup)

1. Create a free account at https://github.com if you don't have one.
2. Create a **new private repository** (e.g. `sitedown-report`).
3. Upload the three files: `app.py`, `processor.py`, `requirements.txt`.

   You can drag-and-drop them in the GitHub web UI — no Git command line needed.

---

## Step 2 — Deploy on Streamlit Cloud (free)

1. Go to https://share.streamlit.io and sign in with your GitHub account.
2. Click **"New app"**.
3. Fill in:
   - **Repository:** `your-username/sitedown-report`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **"Deploy!"**

Streamlit will install dependencies from `requirements.txt` automatically.
Deployment takes about 1–2 minutes.

5. You get a public URL like:
   ```
   https://your-username-sitedown-report-app-xxxxx.streamlit.app
   ```
   Share this link with your colleagues. That's it.

---

## Step 3 — Share with colleagues

Send the URL via email, Teams, or WhatsApp.

**No installation required** — colleagues open the link in any browser
(Chrome, Edge, Firefox, Safari) on any device (Windows, Mac, phone).

---

## Updating the app later

If you change `processor.py` or `app.py`:
1. Edit the file on GitHub (click the file → pencil icon → edit → commit).
2. Streamlit Cloud auto-detects the change and redeploys within ~30 seconds.

---

## Making the app private (so only your team can access it)

On the free Streamlit Cloud plan, apps are public by default.

Options to restrict access:
- **Streamlit Teams plan** (~$250/mo) – adds email-based access control.
- **Password gate (free workaround)** – add a simple password check:

  ```python
  # Add at the top of app.py, before any other content
  pwd = st.text_input("Enter access password", type="password")
  if pwd != "your-secret-password":
      st.stop()
  ```

- **Host privately** – run on your company's internal server:
  ```bash
  pip install streamlit pandas openpyxl
  streamlit run app.py --server.port 8501
  ```
  Then share the internal IP (e.g. `http://192.168.1.100:8501`) on the office network.

---

## Running locally (for testing before deploying)

```bash
# 1. Install dependencies
pip install streamlit pandas openpyxl

# 2. Run the app
cd sitedown_web
streamlit run app.py
```

Your browser will open automatically at http://localhost:8501.
