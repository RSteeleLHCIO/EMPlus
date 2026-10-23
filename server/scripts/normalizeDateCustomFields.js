// One-off cleanup: Data Dictionary "date" fields should hold ISO "YYYY-MM-DD" values (that's
// what the app's date inputs, filters and sorting expect), but some stored values are in
// other formats such as "MM/DD/YYYY" (typically entered while the field was still a text
// field). This rewrites those values to ISO in both EMPlusNodes (current) and
// EMPlusNodeHistory (so the audit trail doesn't show a phantom format "change").
// Values that can't be parsed as a date are reported and left untouched. No other
// attributes are changed and no new history versions are written.
//
// Usage:
//   node server/scripts/normalizeDateCustomFields.js            (dry run, no writes)
//   node server/scripts/normalizeDateCustomFields.js --apply    (writes changes)
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
import { normalizeDateInput } from "../dateUtils.js";

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

// Returns { changed, customFields, changes, unparseable } for one record's customFields.
function normalizeRecord(customFields, dateFieldIds) {
  const changes = [];
  const unparseable = [];
  if (!customFields || typeof customFields !== "object") return { changed: false, changes, unparseable };
  const next = { ...customFields };
  for (const fieldId of dateFieldIds) {
    const raw = customFields[fieldId];
    if (raw === undefined || raw === null || raw === "") continue;
    const values = Array.isArray(raw) ? raw : [raw];
    const fixed = values.map((v) => {
      if (v === undefined || v === null || String(v).trim() === "") return v;
      const iso = normalizeDateInput(v);
      if (!iso) { unparseable.push({ fieldId, value: v }); return v; }
      return iso;
    });
    const same = fixed.every((v, i) => v === values[i]);
    if (same) continue;
    next[fieldId] = Array.isArray(raw) ? fixed : fixed[0];
    changes.push({ fieldId, from: raw, to: next[fieldId] });
  }
  return { changed: changes.length > 0, customFields: next, changes, unparseable };
}

async function main() {
  console.log(`Mode: ${APPLY ? "APPLY (writing changes)" : "DRY RUN (no writes)"}`);
  const clientIds = CLIENT_FILTER ? [CLIENT_FILTER] : await listClientIds();

  const nodeFixes = [];
  const historyFixes = [];
  const unparseableAll = [];

  for (const clientId of clientIds) {
    const dateFields = (await queryDDFields(clientId)).filter((f) => f.dataType === "date");
    if (dateFields.length === 0) continue;
    const dateFieldIds = dateFields.map((f) => f.fieldId);
    const promptById = new Map(dateFields.map((f) => [f.fieldId, f.prompt || f.fieldId]));

    for (const node of await queryNodes(clientId)) {
      const result = normalizeRecord(node.customFields, dateFieldIds);
      result.unparseable.forEach((u) =>
        unparseableAll.push({ clientId, name: node.name, field: promptById.get(u.fieldId), value: u.value })
      );
      if (result.changed) nodeFixes.push({ clientId, row: node, customFields: result.customFields, changes: result.changes });
    }

    await forEachNodeHistoryPage(clientId, (rows) => {
      for (const row of rows) {
        const result = normalizeRecord(row.customFields, dateFieldIds);
        if (result.changed) historyFixes.push({ clientId, row, customFields: result.customFields });
      }
    });

    const clientNodeFixes = nodeFixes.filter((f) => f.clientId === clientId);
    const clientHistoryFixes = historyFixes.filter((f) => f.clientId === clientId);
    console.log(
      `\n${clientId}: date fields [${dateFields.map((f) => f.prompt || f.fieldId).join(", ")}] — ` +
      `${clientNodeFixes.length} node(s), ${clientHistoryFixes.length} history row(s) to normalize`
    );
    clientNodeFixes.slice(0, 10).forEach(({ row, changes }) =>
      changes.forEach((c) => console.log(`  ${row.name}: ${promptById.get(c.fieldId)} ${JSON.stringify(c.from)} -> ${JSON.stringify(c.to)}`))
    );
    if (clientNodeFixes.length > 10) console.log(`  … and ${clientNodeFixes.length - 10} more node(s)`);
  }

  if (unparseableAll.length) {
    console.log(`\n${unparseableAll.length} value(s) could not be parsed as dates and will be left as-is:`);
    unparseableAll.forEach((u) => console.log(`  ${u.clientId}: ${u.name} — ${u.field} = ${JSON.stringify(u.value)}`));
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

  const backupPath = `date-field-normalize-backup-${Date.now()}.json`;
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
  console.error("normalizeDateCustomFields failed:", err);
  process.exitCode = 1;
});
