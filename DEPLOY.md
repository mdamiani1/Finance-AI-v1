# Financial Forecasting Tool — Deployment Guide

## Run locally (2 minutes)

```bash
cd web_app
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501 in your browser.

Your API key can be entered in the sidebar, or set as an env variable:
```bash
export ANTHROPIC_API_KEY=sk-ant-...
streamlit run app.py
```

---

## Deploy to Streamlit Community Cloud (free, public URL)

**Step 1 — Push to GitHub**

Create a new GitHub repo and push the `web_app/` folder contents to the root:
```
your-repo/
  app.py
  extractor.py
  excel_builder.py
  requirements.txt
```

**Step 2 — Deploy**

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **New app**
4. Select your repo, branch `main`, and file `app.py`
5. Click **Deploy**

You'll get a public URL like `https://your-app.streamlit.app` in ~2 minutes.

**Step 3 — Add your API key as a secret (so it's not exposed)**

In the Streamlit Cloud dashboard → your app → **Settings → Secrets**, add:
```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

The app reads `os.environ.get("ANTHROPIC_API_KEY")` automatically.

---

## Deploy to other platforms

### Render (free tier)
1. Push to GitHub
2. New Web Service on [render.com](https://render.com)
3. Build command: `pip install -r requirements.txt`
4. Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. Add `ANTHROPIC_API_KEY` as an environment variable

### Railway / Fly.io
Same approach — they all support Python web services with env variables.

---

## Notes

- The app uses **Claude's PDF document API** — it sends the full PDF to Claude for intelligent extraction. This works with any company's financial statements, any layout.
- Each PDF extraction costs ~$0.01–0.05 in API credits depending on PDF size and model used.
- `claude-sonnet-4-6` is the best balance of speed, accuracy, and cost for this use case.
- Scanned/image PDFs may have lower extraction accuracy than text-based PDFs.
