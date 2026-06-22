# Deploying to Streamlit Community Cloud

## 0. Before you start
- A free GitHub account: https://github.com
- A free Streamlit account: https://share.streamlit.io (sign in with GitHub)
- Your API keys (already in your local `.streamlit/secrets.toml`)

> ⚠️ **Never commit `.streamlit/secrets.toml`.** It holds your real keys.
> The included `.gitignore` already excludes it. You paste the keys into
> Streamlit Cloud's **Secrets** box instead (step 3).

## 1. Push the project to GitHub
From the project folder (`D:\Python_projects\CVup`):

```bash
git init
git add .
git commit -m "CVup resume tailoring app"
git branch -M main
# create an empty repo on github.com first (e.g. "CVup"), then:
git remote add origin https://github.com/<your-username>/CVup.git
git push -u origin main
```

**Check:** on GitHub, confirm `app.py`, `requirements.txt`, `assets/corgi.png`
are there, and that `.streamlit/secrets.toml` is **NOT** (only the `.example`).

## 2. Create the app on Streamlit Cloud
1. Go to https://share.streamlit.io → **Create app** → **Deploy from GitHub**.
2. Repository: `<your-username>/CVup`  ·  Branch: `main`  ·  Main file: `app.py`
3. **Advanced settings → Python version: 3.12** (do NOT pick 3.14 — Cloud
   doesn't support it yet).

## 3. Add your API keys (Secrets)
In the same **Advanced settings**, find the **Secrets** box and paste this
(use your real keys — copy them from your local `.streamlit/secrets.toml`):

```toml
GEMINI_API_KEY = "AQ.xxxxx-your-real-gemini-key"
GROQ_API_KEY   = "gsk_xxxxx-your-real-groq-key"
# optional: GEMINI_MODEL = "gemini-2.5-flash"
```

Then click **Deploy**. First build takes a couple of minutes.

> You can edit secrets any time later: app page → **⋮ menu → Settings →
> Secrets**. Saving secrets reboots the app automatically.

## 4. The personalised greeting
- The app greets the **owner by name** when the signed-in viewer's email is
  `sallosdorka@gmail.com` (set as `OWNER_EMAIL` in `app.py`).
- Everyone else is greeted by their own username; anonymous visitors get a
  neutral "there".
- For this to work, viewers must be **signed in to Streamlit**. To guarantee
  every viewer is identified, set the app to **private** (Settings → Sharing →
  restrict who can view) — then Streamlit always knows the viewer's email.

## 5. Updating the app later
Just push to GitHub — Streamlit redeploys automatically:

```bash
git add -A
git commit -m "your change"
git push
```
