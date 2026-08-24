// One-off migration: clamp any ownership record's effectiveFrom that falls
// AFTER a cutoff date down to that cutoff. Records on/before the cutoff are untouched.
//
// Usage:
//   node server/scripts/clampOwnershipEffectiveDates.js            (dry run, no writes)
//   node server/scripts/clampOwnershipEffectiveDates.js --apply    (writes changes)
//
// Optional env var CLAMP_CUTOFF overrides the default cutoff of 2025-12-31.
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
const rawClient = new DynamoDBClient({ region });
const ddb = DynamoDBDocumentClient.from(rawClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const OWNERSHIP_CURRENT_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_CURRENT || "EMPlusOwnershipCurrent";
const OWNERSHIP_HISTORY_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_HISTORY || "EMPlusOwnershipHistory";

const CUTOFF = process.env.CLAMP_CUTOFF || "2025-12-31";
const APPLY = process.argv.includes("--apply");

// Mirrors makeOwnershipHistoryKey in server/app.js - the date is baked into this key,
// so a changed effectiveFrom requires a delete of the old key + put under the new one.
const makeOwnershipHistoryKey = ({ to, effectiveFrom, from, setId }) =>
  `OWNED#${to}#FROM#${from}#AT#${effectiveFrom}#SET#${setId}`;

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
  console.log(`Cutoff date: effectiveFrom > ${CUTOFF} will be clamped to ${CUTOFF}\n`);

  // Ownership Current: key (OWNS#from#to) has no date in it, so effectiveFrom can be updated in place.
  const currentItems = await scanAll(OWNERSHIP_CURRENT_TABLE);
  const currentToUpdate = currentItems.filter(
    (r) => r.type === "owns" && r.effectiveFrom && r.effectiveFrom > CUTOFF
  );
  console.log(`EMPlusOwnershipCurrent: ${currentToUpdate.length} record(s) to clamp (of ${currentItems.length} total).`);
  currentToUpdate.forEach((r) =>
    console.log(`  [current] clientId=${r.clientId} to=${r.to} from=${r.from} effectiveFrom ${r.effectiveFrom} -> ${CUTOFF}`)
  );

  // Ownership History: date is baked into ownershipHistoryKey, so requires delete + re-put under a new key.
  const historyItems = await scanAll(OWNERSHIP_HISTORY_TABLE);
  const historyToUpdate = historyItems.filter(
    (r) => r.type === "owns" && r.effectiveFrom && r.effectiveFrom > CUTOFF
  );
  console.log(`\nEMPlusOwnershipHistory: ${historyToUpdate.length} record(s) to clamp (of ${historyItems.length} total).`);
  historyToUpdate.forEach((r) =>
    console.log(`  [history] clientId=${r.clientId} to=${r.to} from=${r.from} setId=${r.setId} effectiveFrom ${r.effectiveFrom} -> ${CUTOFF}`)
  );

  if (!APPLY) {
    console.log("\nDry run only - no changes written. Re-run with --apply to write changes.");
    return;
  }
  if (currentToUpdate.length === 0 && historyToUpdate.length === 0) {
    console.log("\nNothing to update.");
    return;
  }

  // Back up originals before mutating anything, so this can be manually reversed if needed.
  const backupPath = `ownership-clamp-backup-${Date.now()}.json`;
  writeFileSync(
    backupPath,
    JSON.stringify({ cutoff: CUTOFF, current: currentToUpdate, history: historyToUpdate }, null, 2)
  );
  console.log(`\nBacked up ${currentToUpdate.length + historyToUpdate.length} original record(s) to ${backupPath}`);

  for (const record of currentToUpdate) {
    const updated = { ...record, effectiveFrom: CUTOFF };
    delete updated.effectiveTo;
    await ddb.send(new PutCommand({ TableName: OWNERSHIP_CURRENT_TABLE, Item: updated }));
  }
  console.log(`Updated ${currentToUpdate.length} EMPlusOwnershipCurrent record(s).`);

  for (const record of historyToUpdate) {
    const newKey = makeOwnershipHistoryKey({
      to: record.to,
      effectiveFrom: CUTOFF,
      from: record.from,
      setId: record.setId,
    });
    const updated = { ...record, ownershipHistoryKey: newKey, effectiveFrom: CUTOFF };
    delete updated.effectiveTo;
    // Put the new-keyed item first, then remove the old one, so a crash mid-way never loses data.
    await ddb.send(new PutCommand({ TableName: OWNERSHIP_HISTORY_TABLE, Item: updated }));
    await ddb.send(
      new DeleteCommand({
        TableName: OWNERSHIP_HISTORY_TABLE,
        Key: { clientId: record.clientId, ownershipHistoryKey: record.ownershipHistoryKey },
      })
    );
  }
  console.log(`Updated ${historyToUpdate.length} EMPlusOwnershipHistory record(s).`);

  console.log("\nDone.");
}

main().catch((err) => {
  console.error("clampOwnershipEffectiveDates failed:", err);
  process.exitCode = 1;
});
