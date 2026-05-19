# TRULYPDF

Browser-based PDF toolbox built with React + Vite.

## Local run

```bash
npm install
npm run dev
```

## Production build check

```bash
npm run build
npm run preview
```

## Argos Translate setup (Translate PDF)

`Translate PDF` now uses an Argos Translate backend API (Python service in `backend/`).

Frontend env var:
- `VITE_TRANSLATE_API_URL` (required in production), example: `https://trulypdf-argos-api.onrender.com/api/translate`

Local dev:
1. Start backend:
```bash
cd backend
pip install -r requirements.txt
python app.py
```
2. Start frontend in repo root:
```bash
npm run dev
```

Backend env vars (optional):
- `ARGOS_PREWARM_PAIRS` default: `en-hi,hi-en,en-es,es-en,en-fr,fr-en`
- `ARGOS_TRANSLATE_TIMEOUT_SEC` default: `45`
- `ARGOS_TRANSLATE_RETRIES` default: `2`

Backend endpoints:
- `POST /api/translate` translate text with Argos
- `GET /api/languages` list installed languages and available targets
- `GET /api/health` health check

## Free Render deploy

This repo is ready for Render Blueprint deploy using [`render.yaml`](./render.yaml).

### Option A: Blueprint (recommended)
1. Push this repo to GitHub.
2. In Render, choose `New +` -> `Blueprint`.
3. Select this repo.
4. Render reads `render.yaml` and creates:
   - `trulypdf` (static frontend)
   - `trulypdf-argos-api` (python backend)
5. Deploy.
6. Set `VITE_TRANSLATE_API_URL` in the `trulypdf` service to your API service URL:
   - `https://trulypdf-argos-api.onrender.com/api/translate`

### Option B: Manual setup
Create two services:
1. Static site (`trulypdf`)
   - Build Command: `npm install && npm run build`
   - Publish Directory: `dist`
   - Environment: `Static Site`
   - Env var: `VITE_TRANSLATE_API_URL=https://<your-api-service>.onrender.com/api/translate`
2. Python web service (`trulypdf-argos-api`)
   - Root Directory: `backend`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python app.py`

## Routing note

Client-side routes (`/tool/...`, `/convert`, etc.) are supported in production by Render rewrite:

`/* -> /index.html`

This is already configured in `render.yaml`.

## Important browser/runtime notes

- Camera features (Scan to PDF) require user permission and secure context (Render HTTPS is fine).
- File conversion quality can vary by input complexity because conversions run in-browser.
