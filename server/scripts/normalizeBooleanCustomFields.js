// One-off cleanup: Data Dictionary "boolean" fields should hold real true/false values,
// but some stored values are text (e.g. "Yes"/"No", typically entered while the field was
// still a text field, or "true"/"false" saved by older edit-dialog code). This rewrites
// "yes"/"true" -> true and "no"/"false" -> false (case-insensitive, trimmed) in both
// EMPlusNodes (current) and EMPlusNodeHistory (so the audit trail doesn't show a phantom
// change). Any other value is reported and left untouched (i.e. treated as undeclared).
// No other attributes are changed and no new history versions are written.
//
// Usage:
//   node server/scripts/normalizeBooleanCustomFields.js            (dry run, no writes)
//   node server/scripts/normalizeBooleanCustomFields.js --apply    (writes changes)
//   optional: --client=<clientId>   (default: every client in EMPlusClients)
import "dotenv/config";
import { writeFileSync } from "fs";
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import { DynamoDBDocumentClient, ScanCommand } from "@aws-sdk/lib-dynamodb";
import {
  CLIENTS_TABLE,
  queryDDFields,
  queryNodes,
  putNode,
  forEachNodeHistoryPage,
  putNodeHistory,
} from "../dynamoClient.js";

const APPLY = process.argv.includes("--apply");
const clientArg = process.argv.find((a) => a.startsWith("--client="));
const CLIENT_FILTER = clientArg ? clientArg.slice("--client=".length) : null;

const ddb = DynamoDBDocumentClient.from(
  new DynamoDBClient({ region: process.env.AWS_REGION || "us-east-1" })
);

async function listClientIds() {
  const ids = [];
  let lastKey;
  do {
    const result = await ddb.send(new ScanCommand({
      TableName: CLIENTS_TABLE,
      ProjectionExpression: "clientId",
      ExclusiveStartKey: lastKey,
    }));
    (result.Items || []).forEach((i) => i.clientId && ids.push(i.clientId));
    lastKey = result.LastEvaluatedKey;
  } while (lastKey);
  return ids.sort();
}

const TRUE_WORDS = new Set(["yes", "true"]);
const FALSE_WORDS = new Set(["no", "false"]);

// Returns true/false for a convertible text value, null for "leave as-is".
function toBoolean(v) {
  if (typeof v !== "string") return null;
  const s = v.trim().toLowerCase();
  if (TRUE_WORDS.has(s)) return true;
  if (FALSE_WORDS.has(s)) return false;
  return null;
}

function normalizeRecord(customFields, fieldIds) {
  const changes = [];
  const leftAlone = [];
  if (!customFields || typeof customFields !== "object") return { changed: false, changes, leftAlone };
  const next = { ...customFields };
  for (const fieldId of fieldIds) {
    const raw = customFields[fieldId];
    if (raw === undefined || raw === null || raw === "" || typeof raw === "boolean") continue;
    const b = toBoolean(raw);
    if (b === null) { leftAlone.push({ fieldId, value: raw }); continue; }
    next[fieldId] = b;
    changes.push({ fieldId, from: raw, to: b });
  }
  return { changed: changes.length > 0, customFields: next, changes, leftAlone };
}

async function main() {
  console.log(`Mode: ${APPLY ? "APPLY (writing changes)" : "DRY RUN (no writes)"}`);
  const clientIds = CLIENT_FILTER ? [CLIENT_FILTER] : await listClientIds();

  const nodeFixes = [];
  const historyFixes = [];
  const leftAloneAll = [];

  for (const clientId of clientIds) {
    const fields = (await queryDDFields(clientId)).filter((f) => f.dataType === "boolean");
    if (fields.length === 0) continue;
    const fieldIds = fields.map((f) => f.fieldId);
    const promptById = new Map(fields.map((f) => [f.fieldId, f.prompt || f.fieldId]));

    const tally = new Map(); // "prompt: from -> to" -> count
    for (const node of await queryNodes(clientId)) {
      const result = normalizeRecord(node.customFields, fieldIds);
      result.leftAlone.forEach((u) =>
        leftAloneAll.push({ clientId, name: node.name, field: promptById.get(u.fieldId), value: u.value })
      );
      if (!result.changed) continue;
      nodeFixes.push({ clientId, row: node, customFields: result.customFields });
      result.changes.forEach((c) => {
        const k = `${promptById.get(c.fieldId)}: ${JSON.stringify(c.from)} -> ${c.to}`;
        tally.set(k, (tally.get(k) || 0) + 1);
      });
    }

    let clientHistory = 0;
    await forEachNodeHistoryPage(clientId, (rows) => {
      for (const row of rows) {
        const result = normalizeRecord(row.customFields, fieldIds);
        if (result.changed) { historyFixes.push({ clientId, row, customFields: result.customFields }); clientHistory++; }
      }
    });

    const clientNodes = nodeFixes.filter((f) => f.clientId === clientId).length;
    console.log(
      `\n${clientId}: boolean fields [${fields.map((f) => f.prompt || f.fieldId).join(", ")}] — ` +
      `${clientNodes} node(s), ${clientHistory} history row(s) to normalize`
    );
    [...tally.entries()].sort().forEach(([k, n]) => console.log(`  ${k}  (${n})`));
  }

  if (leftAloneAll.length) {
    console.log(`\n${leftAloneAll.length} value(s) are not yes/no/true/false and will be left as-is:`);
    leftAloneAll.forEach((u) => console.log(`  ${u.clientId}: ${u.name} — ${u.field} = ${JSON.stringify(u.value)}`));
  }

  console.log(`\nTotal: ${nodeFixes.length} node record(s), ${historyFixes.length} history row(s).`);
  if (!APPLY) {
    console.log("Dry run only - no changes written. Re-run with --apply to write changes.");
    return;
  }
  if (nodeFixes.length === 0 && historyFixes.length === 0) {
    console.log("Nothing to update.");
    return;
  }

  const backupPath = `boolean-field-normalize-backup-${Date.now()}.json`;
  writeFileSync(backupPath, JSON.stringify({
    nodes: nodeFixes.map((f) => f.row),
    history: historyFixes.map((f) => f.row),
  }, null, 2));
  console.log(`Backed up original records to ${backupPath}`);

  for (const { row, customFields } of nodeFixes) await putNode({ ...row, customFields });
  for (const { row, customFields } of historyFixes) await putNodeHistory({ ...row, customFields });
  console.log(`Updated ${nodeFixes.length} node record(s) and ${historyFixes.length} history row(s).\nDone.`);
}

main().catch((err) => {
  console.error("normalizeBooleanCustomFields failed:", err);
  process.exitCode = 1;
});
