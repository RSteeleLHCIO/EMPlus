// One-off repair: an earlier run of normalizeNodeEffectiveDates.js rebuilt some
// EMPlusNodeHistory keys with a truncated nodeId (e.g. "NODE#spencerco#AT#..." instead of
// "NODE#spencerco|entity:x#AT#..."), because it split a "clientId|nodeId" string on "|"
// and nodeIds themselves contain "|". The row contents (including the nodeId attribute)
// are correct; only the nodeHistoryKey is wrong. This rewrites each such row under its
// correct key (put new, then delete old). No other attributes are changed.
//
// Usage:
//   node server/scripts/repairNodeHistoryKeys.js            (dry run, no writes)
//   node server/scripts/repairNodeHistoryKeys.js --apply    (writes changes)
//   optional: --client=<clientId>
import "dotenv/config";
import { writeFileSync } from "fs";
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  ScanCommand,
  PutCommand,
  DeleteCommand,
} from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";
const ddb = DynamoDBDocumentClient.from(new DynamoDBClient({ region }), {
  marshallOptions: { removeUndefinedValues: true },
});

const NODE_HISTORY_TABLE = process.env.DYNAMODB_TABLE_NODE_HISTORY || "EMPlusNodeHistory";
const APPLY = process.argv.includes("--apply");
const clientArg = process.argv.find((a) => a.startsWith("--client="));
const CLIENT_FILTER = clientArg ? clientArg.slice("--client=".length) : null;

// Mirrors makeNodeHistoryKey in server/app.js.
const makeNodeHistoryKey = ({ nodeId, effectiveFrom, versionId }) =>
  `NODE#${nodeId}#AT#${effectiveFrom}#VER#${versionId}`;

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
  if (CLIENT_FILTER) console.log(`Scoped to client: ${CLIENT_FILTER}`);

  const items = await scanAll(NODE_HISTORY_TABLE);
  const existingKeys = new Set(items.map((r) => `${r.clientId}\u0000${r.nodeHistoryKey}`));

  const fixes = [];
  const skipped = [];
  items
    .filter((r) => r.nodeId && (!CLIENT_FILTER || r.clientId === CLIENT_FILTER))
    .filter((r) => !String(r.nodeHistoryKey).startsWith(`NODE#${r.nodeId}#AT#`))
    .forEach((row) => {
      if (!row.versionId || !row.effectiveFrom) {
        skipped.push({ row, reason: "missing versionId/effectiveFrom" });
        return;
      }
      const newKey = makeNodeHistoryKey(row);
      if (existingKeys.has(`${row.clientId}\u0000${newKey}`)) {
        skipped.push({ row, reason: `target key already exists: ${newKey}` });
        return;
      }
      fixes.push({ row, newKey });
    });

  fixes.forEach(({ row, newKey }) =>
    console.log(`${row.clientId}: ${row.nodeHistoryKey}\n    -> ${newKey}`)
  );
  skipped.forEach(({ row, reason }) =>
    console.log(`SKIP ${row.clientId}: ${row.nodeHistoryKey} (${reason})`)
  );
  console.log(
    `\n${fixes.length} row(s) to re-key across ` +
    `${new Set(fixes.map((f) => `${f.row.clientId}/${f.row.nodeId}`)).size} node(s); ` +
    `${skipped.length} skipped.`
  );

  if (!APPLY) {
    console.log("\nDry run only - no changes written. Re-run with --apply to write changes.");
    return;
  }
  if (fixes.length === 0) {
    console.log("\nNothing to update.");
    return;
  }

  const backupPath = `node-history-key-repair-backup-${Date.now()}.json`;
  writeFileSync(backupPath, JSON.stringify(fixes.map((f) => f.row), null, 2));
  console.log(`\nBacked up ${fixes.length} original record(s) to ${backupPath}`);

  for (const { row, newKey } of fixes) {
    // Put first, then delete, so a crash mid-way never loses data.
    await ddb.send(new PutCommand({
      TableName: NODE_HISTORY_TABLE,
      Item: { ...row, nodeHistoryKey: newKey },
    }));
    await ddb.send(new DeleteCommand({
      TableName: NODE_HISTORY_TABLE,
      Key: { clientId: row.clientId, nodeHistoryKey: row.nodeHistoryKey },
    }));
  }
  console.log(`Re-keyed ${fixes.length} EMPlusNodeHistory record(s).\nDone.`);
}

main().catch((err) => {
  console.error("repairNodeHistoryKeys failed:", err);
  process.exitCode = 1;
});
