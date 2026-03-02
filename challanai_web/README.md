# ChallanAI — Web UI

Drag the `challanai_web/` folder into [vercel.com/new](https://vercel.com/new) to deploy instantly.

## Environment Variables

Set `BACKEND_URL` to your Cloud Run API URL (e.g. `https://challanai-xxxxx-el.a.run.app`).

## Structure

```
challanai_web/
├── index.html       # Main page
├── styles.css       # All styles
├── app.js           # Client-side logic
├── api/
│   └── proxy.js     # Vercel edge proxy → backend
├── vercel.json      # Rewrite rules
└── README.md
```

## Analytics

Vercel Web Analytics and Speed Insights are included via the standard `/_vercel/` script tags. Enable them in your Vercel dashboard under **Analytics** and **Speed Insights**.
