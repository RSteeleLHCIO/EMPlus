// One-off migration: entities (everything except Ownership groups) no longer support a
// per-edit effective date — every future field edit is recorded as effective from the
// entity's own "origin date" (its earliest recorded effectiveFrom in EMPlusNodeHistory,
// which may predate the literal `createdAt` timestamp for a backdated CSV import — e.g. a
// node imported with "Effective as of" = 12/31/2025 but whose `createdAt` is whenever the
// import actually ran).
//
// This matters because persistNodeVersion only lets an edit become the live "current"
// record if no existing history row has a LATER effectiveFrom. Any node that was ever
// edited before this policy existed has history rows dated whenever each edit happened
// (the old UI defaulted to "today"), which are almost always later than that node's origin
// date. Left alone, the very next edit after this change ships would be silently rejected
// as "historical" — the save would succeed and write to history, but the live Directory/
// Tabular view would keep showing stale data, because PUT /api/nodes/:id can no longer pick
// a later date to catch up (there's no UI for it anymore).
//
// This script fixes that by rewriting every node's EXISTING history rows so they all share
// that one origin effectiveFrom, with every row closed out (effectiveTo = origin date,
// i.e. a zero-width/inactive window) except the single most-recently-changed one (by
// changedAt), which stays open (effectiveTo = null) — exactly the steady-state shape
// persistNodeVersion now produces for new edits. changedAt/changedBy/field values/wasCurrent
// are left untouched, so the full who/what/when audit trail is preserved; only the
// effectiveFrom/effectiveTo bookkeeping (and the key, which embeds effectiveFrom) changes.
//
// The live EMPlusNodes "current" table is NOT touched — it doesn't store effectiveFrom/
// effectiveTo at all, and already reflects the latest edit regardless of this bug.
//
// Usage:
//   node server/scripts/normalizeNodeEffectiveDates.js            (dry run, no writes)
//   node server/scripts/normalizeNodeEffectiveDates.js --apply    (writes changes)
//
// Optional: --node=<nodeId> to scope the scan to a single node (its full id, e.g.
// "entity:battlefield-parkway-mob-ii-llc").
import "dotenv/config";
import { writeFileSync } from "fs";
import { pathToFileURL } from "url";
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  ScanCommand,
  PutCommand,
  DeleteCommand,
} from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";
const rawClient = new DynamoDBClient({ region });
const ddb = DynamoDBDocumentClient.from(rawClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const NODE_HISTORY_TABLE = process.env.DYNAMODB_TABLE_NODE_HISTORY || "EMPlusNodeHistory";

const APPLY = process.argv.includes("--apply");
const nodeArg = process.argv.find((a) => a.startsWith("--node="));
const NODE_FILTER = nodeArg ? nodeArg.slice("--node=".length) : null;

// Mirrors makeNodeHistoryKey in server/app.js - the date is baked into this key, so a
// changed effectiveFrom requires a delete of the old key + put under the new one.
const makeNodeHistoryKey = ({ nodeId, effectiveFrom, versionId }) =>
  `NODE#${nodeId}#AT#${effectiveFrom}#VER#${versionId}`;

const normalizeDate = (value) => {
  const s = String(value || "").slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
};

// Pure planning logic (no I/O) — kept separate so it can be unit tested without DynamoDB.
// Returns null if the node's history is already in the normalized shape (no-op).
// Otherwise returns { originDate, writes }, where each write is
// { oldKey: string|null, item } (oldKey is set only when the row's key must change).
export const planNodeNormalization = (rows, { nodeId }) => {
  const originDate = rows
    .map((r) => normalizeDate(r.effectiveFrom))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b))[0];
  if (!originDate) return null; // no usable dates on any row — nothing to anchor to

  // Already normalized: every row shares the same effectiveFrom. Nothing to do, even if
  // there's only one row (the common, unaffected case).
  const alreadyFlat = rows.every((r) => normalizeDate(r.effectiveFrom) === originDate);
  if (alreadyFlat) return null;

  // True chronological edit order, independent of the (about-to-be-rewritten) effectiveFrom.
  const ordered = [...rows].sort((a, b) => String(a.changedAt || "").localeCompare(String(b.changedAt || "")));
  const lastIdx = ordered.length - 1;

  const writes = [];
  ordered.forEach((row, idx) => {
    const newEffectiveFrom = originDate;
    const newEffectiveTo = idx === lastIdx ? null : originDate;
    if (normalizeDate(row.effectiveFrom) === newEffectiveFrom && (row.effectiveTo || null) === newEffectiveTo) {
      return; // this row is already correct
    }
    const keyChanges = normalizeDate(row.effectiveFrom) !== newEffectiveFrom;
    const updated = { ...row, effectiveFrom: newEffectiveFrom };
    if (newEffectiveTo) updated.effectiveTo = newEffectiveTo;
    else delete updated.effectiveTo;
    if (keyChanges) {
      updated.nodeHistoryKey = makeNodeHistoryKey({
        nodeId,
        effectiveFrom: newEffectiveFrom,
        versionId: row.versionId || row.nodeHistoryKey,
      });
    }
    writes.push({ oldKey: keyChanges ? row.nodeHistoryKey : null, original: row, item: updated });
  });

  if (writes.length === 0) return null;
  return { originDate, writes, openRowId: ordered[lastIdx].versionId || ordered[lastIdx].nodeHistoryKey };
};

async function scanAll(tableName) {
  const items = [];
  let lastKey;
  do {
    const result = await ddb.send(
      new ScanCommand({ TableName: tableName, ExclusiveStartKey: lastKey })
    );
    items.push(...(result.Items || []));
    lastKey = result.LastEvaluatedKey;
  } while (lastKey);
  return items;
}

async function main() {
  console.log(`Mode: ${APPLY ? "APPLY (writing changes)" : "DRY RUN (no writes)"}`);
  if (NODE_FILTER) console.log(`Scoped to node: ${NODE_FILTER}`);
  console.log("");

  const historyItems = await scanAll(NODE_HISTORY_TABLE);
  const relevant = historyItems.filter(
    (r) => r.nodeId && (!NODE_FILTER || r.nodeId === NODE_FILTER)
  );

  // Group by (clientId, nodeId).
  const byNode = new Map();
  relevant.forEach((row) => {
    const key = `${row.clientId}|${row.nodeId}`;
    if (!byNode.has(key)) byNode.set(key, []);
    byNode.get(key).push(row);
  });

  const toWrite = []; // { oldKey, clientId, item }
  const originalRows = [];
  let nodesNeedingChange = 0;

  for (const rows of byNode.values()) {
    // Read ids from the row itself: nodeIds contain "|" (e.g. "spencerco|entity:x"), so
    // splitting the composite map key would truncate them.
    const { clientId, nodeId } = rows[0];
    const plan = planNodeNormalization(rows, { nodeId });
    if (!plan) continue;

    nodesNeedingChange += 1;
    const closedCount = plan.writes.filter((w) => w.item.effectiveTo).length;
    console.log(
      `Node ${nodeId} (client=${clientId}): origin=${plan.originDate}, ${rows.length} row(s), ` +
      `${closedCount} to close, ${plan.writes.length - closedCount} to keep open.`
    );
    plan.writes.forEach(({ oldKey, original, item }) => {
      originalRows.push(original);
      toWrite.push({ oldKey, clientId, item });
    });
  }

  console.log(`\n${nodesNeedingChange} node(s) need normalizing, ${toWrite.length} history row(s) to rewrite.`);

  if (!APPLY) {
    console.log("\nDry run only - no changes written. Re-run with --apply to write changes.");
    return;
  }
  if (toWrite.length === 0) {
    console.log("\nNothing to update.");
    return;
  }

  const backupPath = `node-effective-date-backup-${Date.now()}.json`;
  writeFileSync(backupPath, JSON.stringify(originalRows, null, 2));
  console.log(`\nBacked up ${originalRows.length} original record(s) to ${backupPath}`);

  for (const { oldKey, clientId, item } of toWrite) {
    // Put the new/updated item first, then remove the old key (if it changed), so a crash
    // mid-way never loses data.
    await ddb.send(new PutCommand({ TableName: NODE_HISTORY_TABLE, Item: item }));
    if (oldKey) {
      await ddb.send(
        new DeleteCommand({
          TableName: NODE_HISTORY_TABLE,
          Key: { clientId, nodeHistoryKey: oldKey },
        })
      );
    }
  }
  console.log(`Updated ${toWrite.length} EMPlusNodeHistory record(s).`);

  console.log("\nDone.");
}

// Guarded so `planNodeNormalization` can be imported (e.g. by tests) without triggering a
// real DynamoDB scan — only run main() when this file is executed directly.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((err) => {
    console.error("normalizeNodeEffectiveDates failed:", err);
    process.exitCode = 1;
  });
}
