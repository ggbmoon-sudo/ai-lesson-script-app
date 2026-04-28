# Deployment Guide

This app is best deployed as a Node web service, not as GitHub Pages. GitHub Pages can serve the static HTML, but it cannot run `server.js`, so AI proxy, PPTX export, Office/PDF parsing, and `/api/config` would not work.

## Render Deployment

1. Push the repository to GitHub.
2. Open Render and create a new Blueprint.
3. Connect the GitHub repository.
4. Render will read `render.yaml` and create the `eduscript-ai-studio` web service.
5. Add environment variables in Render:

```text
OPENAI_API_KEY=your_openai_key_if_you_want_server_ai
OPENAI_MODEL=gpt-5.2
GOOGLE_DRIVE_CLIENT_ID=your_google_oauth_client_id.apps.googleusercontent.com
PUBLIC_BASE_URL=https://your-service-name.onrender.com
```

`OPENAI_API_KEY` is optional. If it is empty, the frontend falls back to local rule-based generation.

## Google Drive OAuth

After Render gives you a public URL, add it to the Google OAuth Client:

```text
https://your-service-name.onrender.com
```

Keep the local origins too if you still test locally:

```text
http://localhost:4173
http://127.0.0.1:4173
```

Only use the OAuth Client ID in the app. Do not put Google Client Secret in the frontend, GitHub, or Render unless you later build a backend OAuth code flow.

## Health Check

Render uses:

```text
/api/health
```

Expected response includes:

```json
{
  "ok": true,
  "aiEnabled": false,
  "model": "gpt-5.2"
}
```

## Local Production Check

```bash
npm run check
npm start
```

Then open:

```text
http://localhost:4173
```
