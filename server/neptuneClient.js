import AWS from "aws-sdk";

const getConfig = () => {
  const endpoint = process.env.NEPTUNE_ENDPOINT;
  if (!endpoint) {
    throw new Error("Missing NEPTUNE_ENDPOINT env var (host only, no protocol).");
  }
  return {
    endpoint,
    port: process.env.NEPTUNE_PORT || "8182",
    region: process.env.NEPTUNE_REGION || process.env.AWS_REGION || "us-east-1",
  };
};

export async function runOpenCypher(query, parameters = {}) {
  const { endpoint, port, region } = getConfig();
  const body = JSON.stringify({ query, parameters });

  AWS.config.update({ region });
  await new Promise((resolve, reject) => {
    AWS.config.getCredentials((err) => {
      if (err) reject(err);
      else resolve();
    });
  });

  const url = `https://${endpoint}:${port}/openCypher`;
  const awsEndpoint = new AWS.Endpoint(`https://${endpoint}:${port}`);
  const request = new AWS.HttpRequest(awsEndpoint, region);
  request.method = "POST";
  request.path = "/openCypher";
  request.headers.host = `${endpoint}:${port}`;
  request.headers["Content-Type"] = "application/json";
  request.body = body;

  const signer = new AWS.Signers.V4(request, "neptune-db");
  signer.addAuthorization(AWS.config.credentials, new Date());

  const response = await fetch(url, {
    method: request.method,
    headers: request.headers,
    body: request.body,
  });

  const text = await response.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  if (!response.ok) {
    const message = data?.message || response.statusText || "Neptune error";
    const error = new Error(message);
    error.status = response.status;
    error.details = data;
    throw error;
  }

  return data;
}
