# Screenshot Capture Utility

A small utility service for converting **Excel spreadsheet ranges into PNG screenshots** (base64 in/out).

## What’s in this repo

- `screenshot-service/`: the HTTP service (Node.js) that performs Excel → PNG conversion
- `docker-compose.yml`: deployment for the screenshot service (e.g. Coolify)

## Quick start (local)

```bash
cd screenshot-service
npm install
npm start
```

Health check:

```bash
curl http://localhost:3000/health
```

See `screenshot-service/README.md` for full API docs.

