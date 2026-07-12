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
const NODE_HISTORY_TABLE = process.env.DYNAMODB_TABLE_NODE_HISTORY || "EMPlusNodeHistory";
const RELS_TABLE = process.env.DYNAMODB_TABLE_RELS || "EMPlusRels";
const OWNERSHIP_CURRENT_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_CURRENT || "EMPlusOwnershipCurrent";
const OWNERSHIP_HISTORY_TABLE = process.env.DYNAMODB_TABLE_OWNERSHIP_HISTORY || "EMPlusOwnershipHistory";
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
  const MAX_RETRIES = 3;
  
  let remaining = items;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (remaining.length === 0) return;
    
    const promises = [];
    for (let i = 0; i < remaining.length; i += 25) {
      const chunk = remaining.slice(i, i + 25);
      promises.push(
        ddb.send(
          new BatchWriteCommand({
            RequestItems: {
              [NODES_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
            },
          })
        )
      );
    }
    
    const responses = await Promise.all(promises);
    
    // Collect any unprocessed items
    remaining = [];
    for (const resp of responses) {
      if (resp.UnprocessedItems?.[NODES_TABLE]?.length > 0) {
        remaining.push(
          ...resp.UnprocessedItems[NODES_TABLE]
            .filter((req) => req.PutRequest)
            .map((req) => req.PutRequest.Item)
        );
      }
    }
    
    // If there are unprocessed items, wait before retry
    if (remaining.length > 0 && attempt < MAX_RETRIES - 1) {
      const backoff = 100 * Math.pow(2, attempt);
      await new Promise((resolve) => setTimeout(resolve, backoff));
    }
  }
  
  if (remaining.length > 0) {
    throw new Error(`Failed to write ${remaining.length} node items after ${MAX_RETRIES} retries`);
  }
}

export async function putNodeHistory(item) {
  await ddb.send(new PutCommand({ TableName: NODE_HISTORY_TABLE, Item: item }));
}

export async function deleteNodeHistory(clientId, nodeHistoryKey) {
  await ddb.send(
    new DeleteCommand({ TableName: NODE_HISTORY_TABLE, Key: { clientId, nodeHistoryKey } })
  );
}

export async function queryNodeHistory(clientId) {
  return paginatedQuery({
    TableName: NODE_HISTORY_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

export async function batchPutNodeHistory(items) {
  const MAX_RETRIES = 3;
  
  let remaining = items;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (remaining.length === 0) return;
    
    const promises = [];
    for (let i = 0; i < remaining.length; i += 25) {
      const chunk = remaining.slice(i, i + 25);
      promises.push(
        ddb.send(
          new BatchWriteCommand({
            RequestItems: {
              [NODE_HISTORY_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
            },
          })
        )
      );
    }
    
    const responses = await Promise.all(promises);
    
    // Collect any unprocessed items
    remaining = [];
    for (const resp of responses) {
      if (resp.UnprocessedItems?.[NODE_HISTORY_TABLE]?.length > 0) {
        remaining.push(
          ...resp.UnprocessedItems[NODE_HISTORY_TABLE]
            .filter((req) => req.PutRequest)
            .map((req) => req.PutRequest.Item)
        );
      }
    }
    
    // If there are unprocessed items, wait before retry
    if (remaining.length > 0 && attempt < MAX_RETRIES - 1) {
      const backoff = 100 * Math.pow(2, attempt);
      await new Promise((resolve) => setTimeout(resolve, backoff));
    }
  }
  
  if (remaining.length > 0) {
    throw new Error(`Failed to write ${remaining.length} node history items after ${MAX_RETRIES} retries`);
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
  const MAX_RETRIES = 3;
  
  let remaining = items;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (remaining.length === 0) return;
    
    const promises = [];
    for (let i = 0; i < remaining.length; i += 25) {
      const chunk = remaining.slice(i, i + 25);
      promises.push(
        ddb.send(
          new BatchWriteCommand({
            RequestItems: {
              [RELS_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
            },
          })
        )
      );
    }
    
    const responses = await Promise.all(promises);
    
    // Collect any unprocessed items
    remaining = [];
    for (const resp of responses) {
      if (resp.UnprocessedItems?.[RELS_TABLE]?.length > 0) {
        remaining.push(
          ...resp.UnprocessedItems[RELS_TABLE]
            .filter((req) => req.PutRequest)
            .map((req) => req.PutRequest.Item)
        );
      }
    }
    
    // If there are unprocessed items, wait before retry
    if (remaining.length > 0 && attempt < MAX_RETRIES - 1) {
      const backoff = 100 * Math.pow(2, attempt);
      await new Promise((resolve) => setTimeout(resolve, backoff));
    }
  }
  
  if (remaining.length > 0) {
    throw new Error(`Failed to write ${remaining.length} rel items after ${MAX_RETRIES} retries`);
  }
}

// ─── Ownership Current / History ─────────────────────────────────────────────

export async function putOwnershipCurrent(item) {
  await ddb.send(new PutCommand({ TableName: OWNERSHIP_CURRENT_TABLE, Item: item }));
}

export async function deleteOwnershipCurrent(clientId, ownershipKey) {
  await ddb.send(
    new DeleteCommand({ TableName: OWNERSHIP_CURRENT_TABLE, Key: { clientId, ownershipKey } })
  );
}

export async function queryOwnershipCurrent(clientId) {
  return paginatedQuery({
    TableName: OWNERSHIP_CURRENT_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

export async function putOwnershipHistory(item) {
  await ddb.send(new PutCommand({ TableName: OWNERSHIP_HISTORY_TABLE, Item: item }));
}

export async function queryOwnershipHistory(clientId) {
  return paginatedQuery({
    TableName: OWNERSHIP_HISTORY_TABLE,
    KeyConditionExpression: "clientId = :c",
    ExpressionAttributeValues: { ":c": clientId },
  });
}

export async function deleteOwnershipHistory(clientId, ownershipHistoryKey) {
  await ddb.send(
    new DeleteCommand({ TableName: OWNERSHIP_HISTORY_TABLE, Key: { clientId, ownershipHistoryKey } })
  );
}

export async function batchPutOwnershipCurrent(items) {
  const MAX_RETRIES = 3;
  
  let remaining = items;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (remaining.length === 0) return;
    
    console.log(`[batchPutOwnershipCurrent] Attempt ${attempt + 1}/${MAX_RETRIES}: writing ${remaining.length} items`);
    
    const promises = [];
    for (let i = 0; i < remaining.length; i += 25) {
      const chunk = remaining.slice(i, i + 25);
      promises.push(
        ddb.send(
          new BatchWriteCommand({
            RequestItems: {
              [OWNERSHIP_CURRENT_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
            },
          })
        )
      );
    }
    
    const responses = await Promise.all(promises);
    
    // Collect any unprocessed items
    const previousCount = remaining.length;
    remaining = [];
    for (const resp of responses) {
      if (resp.UnprocessedItems?.[OWNERSHIP_CURRENT_TABLE]?.length > 0) {
        remaining.push(
          ...resp.UnprocessedItems[OWNERSHIP_CURRENT_TABLE]
            .filter((req) => req.PutRequest)
            .map((req) => req.PutRequest.Item)
        );
      }
    }
    
    const processed = previousCount - remaining.length;
    console.log(`[batchPutOwnershipCurrent] Attempt ${attempt + 1} result: ${processed} processed, ${remaining.length} unprocessed`);
    
    // If there are unprocessed items, wait before retry
    if (remaining.length > 0 && attempt < MAX_RETRIES - 1) {
      const backoff = 100 * Math.pow(2, attempt);
      console.log(`[batchPutOwnershipCurrent] Waiting ${backoff}ms before retry...`);
      await new Promise((resolve) => setTimeout(resolve, backoff));
    }
  }
  
  if (remaining.length > 0) {
    const err = new Error(`Failed to write ${remaining.length} ownership current items after ${MAX_RETRIES} retries`);
    console.error(`[batchPutOwnershipCurrent] ${err.message}`);
    err.failedItems = remaining;
    throw err;
  }
}

export async function batchPutOwnershipHistory(items) {
  const MAX_RETRIES = 3;
  
  let remaining = items;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    if (remaining.length === 0) return;
    
    console.log(`[batchPutOwnershipHistory] Attempt ${attempt + 1}/${MAX_RETRIES}: writing ${remaining.length} items`);
    
    const promises = [];
    for (let i = 0; i < remaining.length; i += 25) {
      const chunk = remaining.slice(i, i + 25);
      promises.push(
        ddb.send(
          new BatchWriteCommand({
            RequestItems: {
              [OWNERSHIP_HISTORY_TABLE]: chunk.map((item) => ({ PutRequest: { Item: item } })),
            },
          })
        )
      );
    }
    
    const responses = await Promise.all(promises);
    
    // Collect any unprocessed items
    const previousCount = remaining.length;
    remaining = [];
    for (const resp of responses) {
      if (resp.UnprocessedItems?.[OWNERSHIP_HISTORY_TABLE]?.length > 0) {
        remaining.push(
          ...resp.UnprocessedItems[OWNERSHIP_HISTORY_TABLE]
            .filter((req) => req.PutRequest)
            .map((req) => req.PutRequest.Item)
        );
      }
    }
    
    const processed = previousCount - remaining.length;
    console.log(`[batchPutOwnershipHistory] Attempt ${attempt + 1} result: ${processed} processed, ${remaining.length} unprocessed`);
    
    // If there are unprocessed items, wait before retry
    if (remaining.length > 0 && attempt < MAX_RETRIES - 1) {
      const backoff = 100 * Math.pow(2, attempt);
      console.log(`[batchPutOwnershipHistory] Waiting ${backoff}ms before retry...`);
      await new Promise((resolve) => setTimeout(resolve, backoff));
    }
  }
  
  if (remaining.length > 0) {
    const err = new Error(`Failed to write ${remaining.length} ownership history items after ${MAX_RETRIES} retries`);
    console.error(`[batchPutOwnershipHistory] ${err.message}`);
    err.failedItems = remaining;
    throw err;
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

// ─── Clients ─────────────────────────────────────────────────────────────────
// One record per client (PK = clientId, no sort key).

export const CLIENTS_TABLE = process.env.DYNAMODB_TABLE_CLIENTS || "EMPlusClients";

export async function getClient(clientId) {
  const result = await ddb.send(
    new GetCommand({ TableName: CLIENTS_TABLE, Key: { clientId } })
  );
  return result.Item || null;
}

export async function putClient(item) {
  await ddb.send(new PutCommand({ TableName: CLIENTS_TABLE, Item: item }));
}
