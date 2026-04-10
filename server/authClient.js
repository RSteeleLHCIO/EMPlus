import { DynamoDBClient } from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  GetCommand,
  PutCommand,
  DeleteCommand,
  QueryCommand,
} from "@aws-sdk/lib-dynamodb";

const region = process.env.AWS_REGION || "us-east-1";
const rawClient = new DynamoDBClient({ region });
const ddb = DynamoDBDocumentClient.from(rawClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const USERS_TABLE = process.env.DYNAMODB_TABLE_USERS || "EMPlusUsers";

// Get a user by loginId (email)
export async function getUser(loginId) {
  const result = await ddb.send(
    new GetCommand({ TableName: USERS_TABLE, Key: { loginId } })
  );
  return result.Item || null;
}

// Create or replace a user record
export async function putUser(item) {
  await ddb.send(new PutCommand({ TableName: USERS_TABLE, Item: item }));
  return item;
}

// Delete a user
export async function deleteUser(loginId) {
  await ddb.send(
    new DeleteCommand({ TableName: USERS_TABLE, Key: { loginId } })
  );
}

// List all users for a clientId (requires GSI "users-by-client": PK=clientId)
export async function listUsersByClient(clientId) {
  const result = await ddb.send(
    new QueryCommand({
      TableName: USERS_TABLE,
      IndexName: "users-by-client",
      KeyConditionExpression: "clientId = :c",
      ExpressionAttributeValues: { ":c": clientId },
    })
  );
  return result.Items || [];
}

// Count users for a clientId (used for bootstrap check)
export async function countUsersByClient(clientId) {
  const result = await ddb.send(
    new QueryCommand({
      TableName: USERS_TABLE,
      IndexName: "users-by-client",
      KeyConditionExpression: "clientId = :c",
      ExpressionAttributeValues: { ":c": clientId },
      Select: "COUNT",
    })
  );
  return result.Count || 0;
}
