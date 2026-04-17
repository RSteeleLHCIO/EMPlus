import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  GetCommand,
  PutCommand,
  DeleteCommand,
  QueryCommand,
  BatchWriteCommand,
} from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";

const rawClient = new DynamoDBClient({ region });
const ddb = DynamoDBDocumentClient.from(rawClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const NODES_TABLE = process.env.DYNAMODB_TABLE_NODES || "EMPlusNodes";
const RELS_TABLE = process.env.DYNAMODB_TABLE_RELS || "EMPlusRels";
const DD_TABLE = process.env.DYNAMODB_TABLE_DD_FIELDS || "EMPlusDDFields";

// Produces the sort key used for a relationship item.
// e.g. makeRelKey("owns", "test|entity:a", "test|entity:b") => "OWNS#test|entity:a#test|entity:b"
export const makeRelKey = (type, fromId, toId) =>
  `${String(type).toUpperCase()}#${fromId}#${toId}`;

// ─── Internal helper ──────────────────────────────────────────────────────────

async function paginatedQuery(params) {
  const items = [];
  let lastKey;
  do {
    const result = await ddb.send(
      new QueryCommand({ ...params, ExclusiveStartKey: lastKey })
    );
    items.push(...(result.Items || []));
    lastKey = result.LastEvaluatedKey;
  } while (lastKey);
  return items;
}

// ─── Nodes ────────────────────────────────────────────────────────────────────

export async function getNode(clientId, nodeId) {
  const result = await ddb.send(
    new GetCommand({ TableName: NODES_TABLE, Key: { clientId, nodeId } })
  );
  return result.Item || null;
}

export async function putNode(item) {
  await ddb.send(new PutCommand({ TableName: NODES_TABLE, Item: item }));
}

export async function deleteNode(clientId, nodeId) {
  await ddb.send(
    new DeleteCommand({ TableName: NODES_TABLE, Key: { clientId, nodeId } })
  );
}

// Returns all node items for a client (Query on PK clientId).
export async function queryNodes(clientId) {
  return paginatedQuery({
    TableName: NODES_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

// Batch-upsert nodes (25 per BatchWrite call — DynamoDB limit).
export async function batchPutNodes(items) {
  for (let i = 0; i < items.length; i += 25) {
    const chunk = items.slice(i, i + 25);
    await ddb.send(
      new BatchWriteCommand({
        RequestItems: {
          [NODES_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
        },
      })
    );
  }
}

// ─── Relationships ────────────────────────────────────────────────────────────

export async function getRel(clientId, relKey) {
  const result = await ddb.send(
    new GetCommand({ TableName: RELS_TABLE, Key: { clientId, relKey } })
  );
  return result.Item || null;
}

export async function putRel(item) {
  await ddb.send(new PutCommand({ TableName: RELS_TABLE, Item: item }));
}

export async function deleteRel(clientId, relKey) {
  await ddb.send(
    new DeleteCommand({ TableName: RELS_TABLE, Key: { clientId, relKey } })
  );
}

// Returns all rel items for a client (Query on PK clientId).
export async function queryRels(clientId) {
  return paginatedQuery({
    TableName: RELS_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

// GSI-1 "rels-by-from": fetch all rels where from = fromId.
// `from` is a DynamoDB reserved word — use expression attribute name alias.
export async function queryRelsByFrom(fromId) {
  return paginatedQuery({
    TableName: RELS_TABLE,
    IndexName: "rels-by-from",
    KeyConditionExpression: "#from = :f",
    ExpressionAttributeNames: { "#from": "from" },
    ExpressionAttributeValues: { ":f": fromId },
  });
}

// GSI-2 "rels-by-to": fetch all rels where to = toId.
// `to` is a DynamoDB reserved word — use expression attribute name alias.
export async function queryRelsByTo(toId) {
  return paginatedQuery({
    TableName: RELS_TABLE,
    IndexName: "rels-by-to",
    KeyConditionExpression: "#to = :t",
    ExpressionAttributeNames: { "#to": "to" },
    ExpressionAttributeValues: { ":t": toId },
  });
}

// Batch-upsert rels (25 per BatchWrite call — DynamoDB limit).
export async function batchPutRels(items) {
  for (let i = 0; i < items.length; i += 25) {
    const chunk = items.slice(i, i + 25);
    await ddb.send(
      new BatchWriteCommand({
        RequestItems: {
          [RELS_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
        },
      })
    );
  }
}

// ─── Data Dictionary ──────────────────────────────────────────────────────────

export async function getDDField(clientId, fieldId) {
  const result = await ddb.send(
    new GetCommand({ TableName: DD_TABLE, Key: { clientId, fieldId } })
  );
  return result.Item || null;
}

export async function putDDField(item) {
  await ddb.send(new PutCommand({ TableName: DD_TABLE, Item: item }));
}

export async function deleteDDField(clientId, fieldId) {
  await ddb.send(
    new DeleteCommand({ TableName: DD_TABLE, Key: { clientId, fieldId } })
  );
}

export async function queryDDFields(clientId) {
  return paginatedQuery({
    TableName: DD_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

// ─── Export Reports ───────────────────────────────────────────────────────────

export const EXPORT_REPORTS_TABLE =
  process.env.DYNAMODB_TABLE_EXPORT_REPORTS || "EMPlusExportReports";

export async function getExportReport(clientId, reportId) {
  const result = await ddb.send(
    new GetCommand({ TableName: EXPORT_REPORTS_TABLE, Key: { clientId, reportId } })
  );
  return result.Item || null;
}

export async function putExportReport(item) {
  await ddb.send(new PutCommand({ TableName: EXPORT_REPORTS_TABLE, Item: item }));
}

export async function deleteExportReport(clientId, reportId) {
  await ddb.send(
    new DeleteCommand({ TableName: EXPORT_REPORTS_TABLE, Key: { clientId, reportId } })
  );
}

export async function queryExportReports(clientId) {
  return paginatedQuery({
    TableName: EXPORT_REPORTS_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}
