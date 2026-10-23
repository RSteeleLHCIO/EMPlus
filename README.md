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

- Saved tabular views retain data-dictionary column IDs while definitions load. Unavailable fields are omitted from the displayed table without removing their IDs from the saved view.
- Run the tabular-column regression tests with `node --test src\utils\tabularColumns.test.js`.
- The node edit dialog shows a read-only audit trail ("View history") sourced from the node's version history in `EMPlusNodeHistory` (`GET /api/nodes/:id/history`): who entered each change, when, and which fields changed from the prior version. The directory, hierarchy, and tabular views always display current data. Per client request, entities (everything except Ownership groups — see below) no longer support picking a per-edit effective date: every field change is recorded as effective from the entity's own creation date (`node.createdAt`), computed and enforced server-side rather than chosen in the UI, and the old "replaces/splits a prior entry" confirmation is gone. The effective-dating database fields/logic (`effectiveFrom`/`effectiveTo` on `EMPlusNodeHistory`, `persistNodeVersion`'s date-collision handling, Time Machine's "effective" reconstruction mode) are intentionally left in place for a possible future reintroduction — they're just unused by the current UI for non-ownership entities. Because repeat edits now share the same effective date, `persistNodeVersion` no longer deletes a prior history row just because its effective date matches (`replaceSameStart: false`); each edit still gets its own history row, closed out (`effectiveTo`) by the next one, so the full who/what/when audit trail survives. Run `node --test src\utils\nodeVersions.test.js` for the helper tests.
- Time Machine (date-travel snapshot of Directory/Tabular) no longer offers a choice between "effective" and "recorded" reconstruction modes — it's hard-coded to "recorded" (what a live query would have literally shown on that date, based on real audit timestamps), since entities no longer have a meaningful distinct effective date to reconstruct from. The server still accepts and supports the "effective" mode for a future reintroduction.
- Ownership is the one exception: entities can have multiple time-bounded ownership sets, and the ownership editor keeps its own dedicated timeline (`GET /api/ownership/history/:entityId`) for viewing and editing past/future ownership periods, including the "Effective as of" date on ownership edits and CSV imports. The directory/tabular/hierarchy views always show the *current* ownership set; the historical timeline is only used within the ownership editor itself.
- The tabular grid autosaves: dropdowns save immediately, text cells save after a 1-second pause or on blur. Pending edits are also saved when switching views, going Back, hiding or leaving the page. The toolbar shows the save status, with Retry when a save fails.
- Node IDs are permanent. They are derived from the name when a node is created (`kind:slug(name)`), and the server appends `-2`, `-3`, ? if that ID was ever used, including by a deleted node whose history remains. Renaming a node never changes its ID, so relationships, history and ownership stay attached. CSV/imports still match existing nodes by the name-derived ID, so importing a renamed node under its new name creates a separate node.
- Creating or renaming a node (Add dialog, tabular grid, edit dialog, new owner) warns when another node of the same kind has, or ever had, the same name, including deleted nodes. Matching ignores case, spacing and punctuation, and uses `GET /api/node-names`. In the grid, a duplicate name is held unsaved until you leave the cell and confirm. Imports don't warn.
- The preview mounts `src/App.jsx` (this file contains the ported app logic).
- Small UI shim components are under `src/components/ui/` and provide minimal styling/behavior for the preview.
- If you have the original `main.js`, it is not required for the preview and may be removed or kept for reference.

If you'd like, I can tidy the CSS or add a small test harness next.
