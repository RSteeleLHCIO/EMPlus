ERPlus (preview)

This repository contains a small Vite + React preview created to run the original ERPlus UI locally.

Quick start

1. Install dependencies

```powershell
npm install
```

2. Create your env file

```powershell
copy .env.example .env
```

3. Run the backend API

```powershell
npm run server
```

4. Run the Vite dev server (separate terminal)

```powershell
npm run dev
```

Open the app at http://localhost:5173/ in your browser.

AWS/DynamoDB setup

- This app reads and writes directly to DynamoDB through the local Node API in `server/`.
- The AWS SDK uses standard credential resolution (env vars, shared profile, SSO, etc.).
- On a new machine, configure credentials before login/data calls will work.

Option A: shared profile files (recommended)

```powershell
aws configure
```

Option B: env vars for current shell

```powershell
$env:AWS_ACCESS_KEY_ID="..."
$env:AWS_SECRET_ACCESS_KEY="..."
$env:AWS_REGION="us-east-1"
```

Required env keys in `.env`

- `AWS_REGION`
- `DYNAMODB_TABLE_NODES`
- `DYNAMODB_TABLE_RELS`
- `DYNAMODB_TABLE_DD_FIELDS`
- `DYNAMODB_TABLE_EXPORT_REPORTS`
- `DYNAMODB_TABLE_CLIENTS`
- `DYNAMODB_TABLE_USERS`

If these names do not match the tables used on your other machine, update `.env` to those exact table names.

Notes

- The preview mounts `src/App.jsx` (this file contains the ported app logic).
- Small UI shim components are under `src/components/ui/` and provide minimal styling/behavior for the preview.
- If you have the original `main.js`, it is not required for the preview and may be removed or kept for reference.

If you'd like, I can tidy the CSS or add a small test harness next.
