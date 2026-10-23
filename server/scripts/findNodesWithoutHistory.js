// Read-only report: lists every node in EMPlusNodes that has NO rows at all in
// EMPlusNodeHistory. Such nodes predate version tracking (or had their history removed),
// so their original/creation state is not captured anywhere — the first edit made to them
// will become their earliest recorded version.
//
// Usage:
//   node server/scripts/findNodesWithoutHistory.js
//   node server/scripts/findNodesWithoutHistory.js --client=<clientId>
//   node server/scripts/findNodesWithoutHistory.js --json=<outputPath>   (also save results)
import "dotenv/config";
import { writeFileSync } from "fs";
import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import { DynamoDBDocumentClient, ScanCommand } from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";
const ddb = DynamoDBDocumentClient.from(new DynamoDBClient({ region }));

const NODES_TABLE = process.env.DYNAMODB_TABLE_NODES || "EMPlusNodes";
const NODE_HISTORY_TABLE = process.env.DYNAMODB_TABLE_NODE_HISTORY || "EMPlusNodeHistory";

const argValue = (name) => {
  const arg = process.argv.find((a) => a.startsWith(`--${name}=`));
  return arg ? arg.slice(name.length + 3) : null;
};
const CLIENT_FILTER = argValue("client");
const JSON_PATH = argValue("json");

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
  console.log(`Nodes table: ${NODES_TABLE}`);
  console.log(`History table: ${NODE_HISTORY_TABLE}`);
  if (CLIENT_FILTER) console.log(`Scoped to client: ${CLIENT_FILTER}`);
  console.log("");

  const [nodes, historyRows] = await Promise.all([
    scanAll(NODES_TABLE),
    scanAll(NODE_HISTORY_TABLE),
  ]);

  const withHistory = new Set(
    historyRows.filter((r) => r.nodeId).map((r) => `${r.clientId}|${r.nodeId}`)
  );

  const missing = nodes
    .filter((n) => !CLIENT_FILTER || n.clientId === CLIENT_FILTER)
    .filter((n) => !withHistory.has(`${n.clientId}|${n.nodeId}`))
    .map((n) => ({
      clientId: n.clientId,
      nodeId: n.nodeId,
      name: n.name || "",
      kind: n.kind || "",
      createdAt: n.createdAt || null,
      updatedAt: n.updatedAt || null,
    }))
    .sort((a, b) =>
      String(a.clientId).localeCompare(String(b.clientId))
      || String(a.kind).localeCompare(String(b.kind))
      || String(a.name).localeCompare(String(b.name))
    );

  const scopedCount = nodes.filter((n) => !CLIENT_FILTER || n.clientId === CLIENT_FILTER).length;

  if (missing.length === 0) {
    console.log(`All ${scopedCount} node(s) have at least one history record.`);
  } else {
    let currentClient = null;
    missing.forEach((n) => {
      if (n.clientId !== currentClient) {
        currentClient = n.clientId;
        console.log(`\nClient: ${currentClient}`);
      }
      console.log(
        `  [${n.kind || "?"}] ${n.name || "(no name)"}  (${n.nodeId})` +
        `  created=${n.createdAt || "?"}  updated=${n.updatedAt || "?"}`
      );
    });
    console.log(`\n${missing.length} of ${scopedCount} node(s) have NO history record.`);
  }

  if (JSON_PATH) {
    writeFileSync(JSON_PATH, JSON.stringify(missing, null, 2));
    console.log(`Saved results to ${JSON_PATH}`);
  }
}

main().catch((err) => {
  console.error("findNodesWithoutHistory failed:", err);
  process.exitCode = 1;
});
