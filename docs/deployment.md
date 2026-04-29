# Deployment Guide

This app is best deployed as a Node web service, not as GitHub Pages. GitHub Pages can serve the static HTML, but it cannot run `server.js`, so AI proxy, PPTX export, Office/PDF parsing, and `/api/config` would not work.

## Render Deployment

1. Push the repository to GitHub.
2. Open Render and create a new Blueprint.
3. Connect the GitHub repository.
4. Render will read `render.yaml` and create the `eduscript-ai-studio` web service.
5. Add environment variables in Render:

```text
AI_PROVIDER=openai-compatible
OPENAI_COMPAT_BASE_URL=https://api.newcoin.top
OPENAI_COMPAT_API_KEY=your_newcoin_key
OPENAI_COMPAT_MODEL=qwen3.6-plus
OPENAI_COMPAT_TEMPERATURE=0.25
OPENAI_COMPAT_MAX_TOKENS=16384
OPENAI_COMPAT_SCRIPT_MAX_TOKENS=32768
GOOGLE_DRIVE_CLIENT_ID=your_google_oauth_client_id.apps.googleusercontent.com
GAMMA_API_KEY=your_gamma_api_key_if_you_want_gamma_ppt_export
GAMMA_EXPORT_AS=pptx
GAMMA_TEXT_MODE=generate
PUBLIC_BASE_URL=https://your-service-name.onrender.com
```

`OPENAI_COMPAT_API_KEY` is required for AI generation. The app uses OpenAI-compatible chat completions at `OPENAI_COMPAT_BASE_URL`; with NewCoin the server will call `/v1/chat/completions` automatically. `OPENAI_COMPAT_MODEL` can be `qwen3.6-plus` or `qwen3.5-plus`. If no AI key is configured, generation stops with an error rather than using local fallback rules.

`GAMMA_API_KEY` is optional. If it is empty, the `Gamma PPT` button exports a Gamma-ready Markdown prompt instead of calling Gamma. If it is set, `server.js` calls Gamma from the backend so the key is never exposed to the browser.

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
  "provider": "local",
  "model": ""
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
