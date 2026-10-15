// One-off diagnostic + cleanup: finds ownership "periods" (sets of EMPlusOwnershipHistory
// rows sharing a setId) that collide on the same owned entity + effectiveFrom date.
//
// This is a data-integrity bug, not a valid state: an entity should have at most one
// ownership period starting on any given date. It could previously happen via
// PUT /api/ownership/update-period, which (before being fixed) let you move an existing
// period's effective date onto a date another period already occupied, without merging
// or removing the collision. When the timeline is rendered, each period's end date is
// computed as "the day before the next period's start" — with two periods sharing a
// start date, one of them ends up with an end date *before* its own start date (e.g.
// "2025-12-31 to 2025-12-30"), which is the visible symptom of this bug.
//
// For each colliding group, this script keeps one period and deletes the rest:
//   1. Prefer whichever setId's rows are also present in EMPlusOwnershipCurrent (i.e. it's
//      the one actually driving the live "current owners" view).
//   2. Otherwise, keep whichever setId has the most recent changedAt.
//
// Usage:
//   node server/scripts/dedupeOwnershipPeriods.js            (dry run, no writes)
//   node server/scripts/dedupeOwnershipPeriods.js --apply    (writes changes)
//
// Optional: --entity=<ownedNodeId> to scope the scan to a single owned entity (its full
// node id, e.g. "entity:battlefield-parkway-mob-ii-llc").
import "dotenv/config";
import { writeFileSync } from "fs";
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  ScanCommand,
  DeleteCommand,
} from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";
const rawClient = new DynamoDBClient({ region });
const ddb = DynamoDBDocumentClient.from(rawClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const OWNERSHIP_CURRENT_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_CURRENT || "EMPlusOwnershipCurrent";
const OWNERSHIP_HISTORY_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_HISTORY || "EMPlusOwnershipHistory";

const APPLY = process.argv.includes("--apply");
const entityArg = process.argv.find((a) => a.startsWith("--entity="));
const ENTITY_FILTER = entityArg ? entityArg.slice("--entity=".length) : null;

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
  if (ENTITY_FILTER) console.log(`Scoped to entity: ${ENTITY_FILTER}`);
  console.log("");

  const [currentItems, historyItems] = await Promise.all([
    scanAll(OWNERSHIP_CURRENT_TABLE),
    scanAll(OWNERSHIP_HISTORY_TABLE),
  ]);

  const historyOwns = historyItems.filter(
    (r) => r.type === "owns" && r.to && r.setId && (!ENTITY_FILTER || r.to === ENTITY_FILTER)
  );
  const currentSetIds = new Set(
    currentItems.filter((r) => r.type === "owns" && r.setId).map((r) => `${r.clientId}|${r.to}|${r.setId}`)
  );

  // Group by (clientId, to, effectiveFrom) -> distinct setIds sharing that start date.
  const buckets = new Map();
  historyOwns.forEach((row) => {
    const key = `${row.clientId}|${row.to}|${row.effectiveFrom || ""}`;
    if (!buckets.has(key)) buckets.set(key, new Map());
    const bySetId = buckets.get(key);
    if (!bySetId.has(row.setId)) bySetId.set(row.setId, []);
    bySetId.get(row.setId).push(row);
  });

  const toDelete = [];
  let collisionCount = 0;

  for (const [bucketKey, bySetId] of buckets.entries()) {
    if (bySetId.size <= 1) continue; // no collision
    collisionCount += 1;
    const [clientId, to, effectiveFrom] = bucketKey.split("|");

    const setEntries = [...bySetId.entries()].map(([setId, rows]) => ({
      setId,
      rows,
      isCurrent: currentSetIds.has(`${clientId}|${to}|${setId}`),
      latestChangedAt: rows.map((r) => r.changedAt || "").sort().at(-1) || "",
    }));

    // Prefer the setId reflected in EMPlusOwnershipCurrent; otherwise the most recently changed.
    setEntries.sort((a, b) => {
      if (a.isCurrent !== b.isCurrent) return a.isCurrent ? -1 : 1;
      return b.latestChangedAt.localeCompare(a.latestChangedAt);
    });
    const [keep, ...drop] = setEntries;

    console.log(`Collision: clientId=${clientId} to=${to} effectiveFrom=${effectiveFrom}`);
    console.log(`  KEEP  setId=${keep.setId} isCurrent=${keep.isCurrent} changedAt=${keep.latestChangedAt}`);
    drop.forEach((d) => {
      console.log(`  DROP  setId=${d.setId} isCurrent=${d.isCurrent} changedAt=${d.latestChangedAt} (${d.rows.length} row(s))`);
      toDelete.push(...d.rows);
    });
    console.log("");
  }

  console.log(`Found ${collisionCount} colliding period group(s), ${toDelete.length} history row(s) to remove.`);

  if (!APPLY) {
    console.log("\nDry run only - no changes written. Re-run with --apply to write changes.");
    return;
  }
  if (toDelete.length === 0) {
    console.log("\nNothing to clean up.");
    return;
  }

  const backupPath = `ownership-dedupe-backup-${Date.now()}.json`;
  writeFileSync(backupPath, JSON.stringify(toDelete, null, 2));
  console.log(`\nBacked up ${toDelete.length} original record(s) to ${backupPath}`);

  for (const record of toDelete) {
    await ddb.send(
      new DeleteCommand({
        TableName: OWNERSHIP_HISTORY_TABLE,
        Key: { clientId: record.clientId, ownershipHistoryKey: record.ownershipHistoryKey },
      })
    );
  }
  console.log(`Deleted ${toDelete.length} EMPlusOwnershipHistory record(s).`);

  console.log("\nDone.");
}

main().catch((err) => {
  console.error("dedupeOwnershipPeriods failed:", err);
  process.exitCode = 1;
});
