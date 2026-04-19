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

## Free Render deploy

This repo is ready for Render Blueprint deploy using [`render.yaml`](./render.yaml).

### Option A: Blueprint (recommended)
1. Push this repo to GitHub.
2. In Render, choose `New +` -> `Blueprint`.
3. Select this repo.
4. Render reads `render.yaml` and creates a free static service.
5. Deploy.

### Option B: Manual static site setup
Use these values:
- Build Command: `npm install && npm run build`
- Publish Directory: `dist`
- Environment: `Static Site`

## Routing note

Client-side routes (`/tool/...`, `/convert`, etc.) are supported in production by Render rewrite:

`/* -> /index.html`

This is already configured in `render.yaml`.

## Important browser/runtime notes

- Camera features (Scan to PDF) require user permission and secure context (Render HTTPS is fine).
- File conversion quality can vary by input complexity because conversions run in-browser.
