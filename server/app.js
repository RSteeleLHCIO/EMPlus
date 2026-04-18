import "dotenv/config";
import express from "express";
import cors from "cors";
import multer from "multer";
import { parse } from "csv-parse/sync";
import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";
import { randomBytes } from "crypto";
import { S3Client, PutObjectCommand } from "@aws-sdk/client-s3";
import {
  getNode,
  putNode,
  deleteNode,
  queryNodes,
  batchPutNodes,
  getRel,
  putRel,
  deleteRel,
  queryRels,
  queryRelsByFrom,
  queryRelsByTo,
  batchPutRels,
  makeRelKey,
  getDDField,
  putDDField,
  deleteDDField,
  queryDDFields,
  getExportReport,
  putExportReport,
  deleteExportReport,
  queryExportReports,
  getClient,
  putClient,
} from "./dynamoClient.js";
import {
  getUser,
  putUser,
  listUsersByClient,
  countUsersByClient,
} from "./authClient.js";

const JWT_SECRET = process.env.JWT_SECRET || "changeme-set-JWT_SECRET-env-var";
const JWT_EXPIRY = process.env.JWT_EXPIRY || "8h";
const SETUP_SECRET = process.env.SETUP_SECRET || null;

// ─── Pure business-logic helpers (no DynamoDB concerns) ──────────────────────

const normalizeEmails = (value) => {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  if (typeof value === "string") {
    return value.split(/[,;]+/).map((item) => item.trim()).filter(Boolean);
  }
  return [String(value).trim()].filter(Boolean);
};

const slugify = (value) =>
  String(value || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");

const normalizeClientId = (client, rawId) => {
  if (!rawId) return rawId;
  if (!client) return rawId;
  const trimmed = String(rawId).trim();
  if (trimmed.includes("|")) return trimmed;
  return `${client}|${trimmed}`;
};

const detectHeaderRow = (row) => {
  if (!Array.isArray(row)) return false;
  const headers = row.map((cell) => String(cell || "").trim().toLowerCase());
  return headers.some((header) =>
    [
      "name",
      "node",
      "node name",
      "kind",
      "type",
      "owner",
      "owned",
      "entity",
      "percent",
      "percentage",
      "pct",
      "from",
      "to",
      "parent",
      "child",
      "source",
      "target",
    ].includes(header)
  );
};

const getColumnIndex = (headers, candidates, fallback) => {
  const lowered = headers.map((header) => String(header || "").trim().toLowerCase());
  const found = candidates.findIndex((candidate) => lowered.includes(candidate));
  if (found !== -1) return lowered.indexOf(candidates[found]);
  return fallback;
};

const parseNodeCsv = (csvText, defaultKind = "entity", client) => {
  if (!csvText || typeof csvText !== "string") return { rows: [], skipped: 0 };
  const records = parse(csvText, {
    skip_empty_lines: true,
    relax_column_count: true,
    trim: true,
  });
  if (!Array.isArray(records) || records.length === 0) return { rows: [], skipped: 0 };

  let startIndex = 0;
  let nameIndex = 0;
  let kindIndex = 1;
  if (detectHeaderRow(records[0])) {
    const headers = records[0];
    nameIndex = getColumnIndex(headers, ["name", "node", "node name"], 0);
    kindIndex = getColumnIndex(headers, ["kind", "type"], 1);
    startIndex = 1;
  }

  const rows = [];
  let skipped = 0;
  for (let i = startIndex; i < records.length; i += 1) {
    const record = records[i] || [];
    const rawName = record[nameIndex] ?? record[0];
    const rawKind = record[kindIndex] ?? record[1];
    const name = String(rawName || "").trim();
    if (!name) {
      skipped += 1;
      continue;
    }
    const kindValue = String(rawKind || defaultKind || "entity").trim();
    const kind = /person/i.test(kindValue) ? "person" : "entity";
    const slug = slugify(name) || "node";
    const rawId = `${kind}:${slug}`;
    rows.push({ id: normalizeClientId(client, rawId), name, kind });
  }

  return { rows, skipped };
};

const parseOwnershipCsv = (csvText) => {
  if (!csvText || typeof csvText !== "string") return { rows: [], skipped: 0 };
  const records = parse(csvText, {
    skip_empty_lines: true,
    relax_column_count: true,
    trim: true,
  });
  if (!Array.isArray(records) || records.length === 0) return { rows: [], skipped: 0 };

  let startIndex = 0;
  let ownerIndex = 0;
  let ownedIndex = 1;
  let percentIndex = 2;
  if (detectHeaderRow(records[0])) {
    const headers = records[0];
    ownerIndex = getColumnIndex(headers, ["owner", "from", "parent", "source"], 0);
    ownedIndex = getColumnIndex(headers, ["owned", "entity", "to", "child", "target"], 1);
    percentIndex = getColumnIndex(headers, ["percent", "percentage", "pct"], 2);
    startIndex = 1;
  }

  const rows = [];
  let skipped = 0;
  for (let i = startIndex; i < records.length; i += 1) {
    const record = records[i] || [];
    const rawOwner = record[ownerIndex] ?? record[0];
    const rawOwned = record[ownedIndex] ?? record[1];
    const rawPercent = record[percentIndex] ?? record[2];
    const owner = String(rawOwner || "").trim();
    const owned = String(rawOwned || "").trim();
    if (!owner || !owned) {
      skipped += 1;
      continue;
    }
    const percentText = String(rawPercent ?? "").trim();
    const percent = percentText === "" ? Number.NaN : Number(percentText);
    rows.push({ owner, owned, percent });
  }

  return { rows, skipped };
};

// ─── DynamoDB item builders ───────────────────────────────────────────────────

const buildNodeItem = (
  { id, name, kind, client, address, workPhone, cellPhone, emails, photo,
    taxId, accountingUrl, hrUrl, logo, customFields, createdAt, updatedAt }
) => ({
  clientId: client,
  nodeId: id,          // DynamoDB sort key — mirrors id
  id,                  // returned to frontend
  name,
  kind,
  client,
  address: address || "",
  workPhone: workPhone || "",
  cellPhone: cellPhone || "",
  emails: normalizeEmails(emails),
  photo: photo || "",
  taxId: taxId || "",
  accountingUrl: accountingUrl || "",
  hrUrl: hrUrl || "",
  logo: logo || "",
  customFields: customFields || {},
  createdAt,
  updatedAt,
});

const buildOwnsItem = (
  { from, to, percent, startDate, endDate, client, createdAt, updatedAt }
) => ({
  clientId: client,
  relKey: makeRelKey("owns", from, to),
  id: `owns-${from}-${to}`,
  type: "owns",
  from,
  to,
  percent: percent ?? null,
  startDate: startDate || null,
  endDate: endDate || null,
  client,
  createdAt,
  updatedAt,
});

const buildEmploysItem = (
  { from, to, role, startDate, endDate, client, createdAt, updatedAt }
) => ({
  clientId: client,
  relKey: makeRelKey("employs", from, to),
  id: `employs-${from}-${to}`,
  type: "employs",
  from,
  to,
  role: role || null,
  startDate: startDate || null,
  endDate: endDate || null,
  client,
  createdAt,
  updatedAt,
});

// Strip DynamoDB-internal keys before sending to the frontend.
const toNodeResponse = ({ clientId: _c, nodeId: _n, ...rest }) => rest;
const toRelResponse = ({ clientId: _c, relKey: _r, ...rest }) => rest;

// Check whether a relationship is active on a given date string (YYYY-MM-DD).
const relActiveAt = (rel, asOf) => {
  const start = rel.startDate || "0001-01-01";
  const end = rel.endDate || "9999-12-31";
  return start <= asOf && end >= asOf;
};

// Delete every relationship that references nodeId (both directions).
const deleteRelsForNode = async (clientId, nodeId) => {
  const [fromRels, toRels] = await Promise.all([
    queryRelsByFrom(nodeId),
    queryRelsByTo(nodeId),
  ]);
  await Promise.all(
    [...fromRels, ...toRels].map((r) => deleteRel(clientId, r.relKey))
  );
};

// ─── Express app ──────────────────────────────────────────────────────────────

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// ─── S3 (optional — only active when S3_BUCKET_UPLOADS env var is set) ───────
const S3_BUCKET = process.env.S3_BUCKET_UPLOADS || null;
const s3 = S3_BUCKET
  ? new S3Client({ region: process.env.AWS_REGION || "us-east-1" })
  : null;

app.use(
  cors({
    origin: true,
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());
app.use(express.json({ limit: "1mb" }));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "dynamodb-api", version: "2026-04-18.1419p", features: ["client management"] });
});

// ─── AUTH ROUTES (no JWT required) ───────────────────────────────────────────

app.post("/api/auth/login", async (req, res) => {
  try {
    const { loginId, password } = req.body || {};
    if (!loginId || !password) {
      return res.status(400).json({ error: "loginId and password are required" });
    }
    const user = await getUser(String(loginId).toLowerCase().trim());
    if (!user) return res.status(401).json({ error: "Invalid credentials" });
    const match = await bcrypt.compare(password, user.passwordHash);
    if (!match) return res.status(401).json({ error: "Invalid credentials" });
    const token = jwt.sign(
      { loginId: user.loginId, clientId: user.clientId, personId: user.personId || null },
      JWT_SECRET,
      { expiresIn: JWT_EXPIRY }
    );
    res.json({ token, clientId: user.clientId, personId: user.personId || null, role: user.role || "user" });
  } catch (err) {
    console.error("/api/auth/login error", err);
    res.status(500).json({ error: err.message });
  }
});

// Bootstrap: register the FIRST user for a brand-new client.
// Requires SETUP_SECRET when set (guards against unauthorized client creation).
// Rejects the request if the client already has users — use POST /api/auth/users instead.
app.post("/api/auth/register", async (req, res) => {
  try {
    const { loginId, password, clientId, personId, setupSecret } = req.body || {};
    if (!loginId || !password || !clientId) {
      return res.status(400).json({ error: "loginId, password, clientId are required" });
    }
    // Block if this client already has users — they should use the in-app Add User feature
    const existingCount = await countUsersByClient(clientId);
    if (existingCount > 0) {
      return res.status(403).json({ error: "Client already exists. Add users via the in-app Settings menu." });
    }
    // Require SETUP_SECRET for new-client provisioning
    if (SETUP_SECRET && setupSecret !== SETUP_SECRET) {
      return res.status(403).json({ error: "Setup secret required to provision a new client" });
    }
    const normalizedId = String(loginId).toLowerCase().trim();
    const existing = await getUser(normalizedId);
    if (existing) return res.status(409).json({ error: "loginId already exists" });
    const passwordHash = await bcrypt.hash(password, 10);
    const user = {
      loginId: normalizedId,
      clientId,
      personId: personId || null,
      role: "admin", // bootstrap user is always an admin
      passwordHash,
      createdAt: new Date().toISOString(),
    };
    await putUser(user);
    res.status(201).json({ loginId: user.loginId, clientId: user.clientId, role: user.role });
  } catch (err) {
    console.error("/api/auth/register error", err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/auth/promote — break-glass: elevate a user to admin using SETUP_SECRET.
// Declared before JWT middleware intentionally — no token required.
app.post("/api/auth/promote", async (req, res) => {
  try {
    if (!SETUP_SECRET) {
      return res.status(403).json({ error: "SETUP_SECRET is not configured on this server" });
    }
    const { loginId, setupSecret } = req.body || {};
    if (!loginId || !setupSecret) {
      return res.status(400).json({ error: "loginId and setupSecret are required" });
    }
    if (setupSecret !== SETUP_SECRET) {
      return res.status(403).json({ error: "Invalid setup secret" });
    }
    const user = await getUser(String(loginId).toLowerCase().trim());
    if (!user) return res.status(404).json({ error: "User not found" });
    await putUser({ ...user, role: "admin" });
    res.json({ loginId: user.loginId, role: "admin" });
  } catch (err) {
    console.error("/api/auth/promote error", err);
    res.status(500).json({ error: err.message });
  }
});

// ─── JWT MIDDLEWARE (all /api/* routes below this require a valid token) ──────

app.use("/api", (req, res, next) => {
  // Skip only the unauthenticated auth endpoints
  if (req.path === "/auth/login" || req.path === "/auth/register") return next();
  if (req.path === "/health") return next();
  const authHeader = req.headers["authorization"] || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) return res.status(401).json({ error: "Unauthorized" });
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.auth = payload; // { loginId, clientId, personId }
    next();
  } catch {
    return res.status(401).json({ error: "Invalid or expired token" });
  }
});

// ─── Admin middleware ────────────────────────────────────────────────────────
// Use as a route-level middleware: app.get("/api/foo", requireAdmin, handler)
const requireAdmin = async (req, res, next) => {
  try {
    const user = await getUser(req.auth.loginId);
    if (!user || user.role !== "admin") {
      return res.status(403).json({ error: "Admin access required" });
    }
    next();
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};

// openCypher has been removed — backend now uses DynamoDB.
app.post("/api/cypher", (_req, res) => {
  res.status(410).json({ error: "openCypher endpoint removed; backend now uses DynamoDB." });
});

app.post("/api/graph", async (req, res) => {
  try {
    const clientId = req.auth.clientId;
    const { focusId, asOf } = req.body || {};
    if (!focusId || !asOf) {
      return res.status(400).json({ error: "focusId and asOf are required" });
    }
    const focusKey = normalizeClientId(clientId, focusId);

    // Fetch focus node and all rels referencing it in parallel.
    const [focusItem, fromRels, toRels] = await Promise.all([
      getNode(clientId, focusKey),
      queryRelsByFrom(focusKey),
      queryRelsByTo(focusKey),
    ]);

    if (!focusItem) {
      return res.status(404).json({ error: "node not found" });
    }

    const activeFrom = fromRels.filter((r) => relActiveAt(r, asOf));
    const activeTo = toRels.filter((r) => relActiveAt(r, asOf));

    // Fetch all neighbor nodes in parallel.
    const neighborIds = new Set([
      ...activeFrom.map((r) => r.to),
      ...activeTo.map((r) => r.from),
    ]);
    const neighborItems = await Promise.all(
      [...neighborIds].map((nid) => getNode(clientId, nid))
    );
    const neighborMap = new Map(
      neighborItems.filter(Boolean).map((item) => [item.id, toNodeResponse(item)])
    );

    res.json({
      node: toNodeResponse(focusItem),
      owners: activeTo
        .filter((r) => r.type === "owns")
        .map((r) => ({ node: neighborMap.get(r.from) || null, rel: toRelResponse(r), direction: "in" }))
        .filter((x) => x.node),
      owns: activeFrom
        .filter((r) => r.type === "owns")
        .map((r) => ({ node: neighborMap.get(r.to) || null, rel: toRelResponse(r), direction: "out" }))
        .filter((x) => x.node),
      employees: activeFrom
        .filter((r) => r.type === "employs")
        .map((r) => ({ node: neighborMap.get(r.to) || null, rel: toRelResponse(r), direction: "out" }))
        .filter((x) => x.node),
      employers: activeTo
        .filter((r) => r.type === "employs")
        .map((r) => ({ node: neighborMap.get(r.from) || null, rel: toRelResponse(r), direction: "in" }))
        .filter((x) => x.node),
    });
  } catch (err) {
    console.error("/api/graph error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/directory", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const [nodeItems, relItems] = await Promise.all([
      queryNodes(client),
      queryRels(client),
    ]);
    res.json({
      nodes: nodeItems.map(toNodeResponse),
      rels: relItems.map(toRelResponse),
    });
  } catch (err) {
    console.error("/api/directory error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/nodes", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const {
      id, name, kind, address, workPhone, cellPhone, emails,
      photo, taxId, accountingUrl, hrUrl, logo, customFields,
    } = req.body || {};
    if (!id || !name || !kind) {
      return res.status(400).json({ error: "id, name, kind are required" });
    }
    const normalizedId = normalizeClientId(client, id);
    const now = new Date().toISOString();
    const existing = await getNode(client, normalizedId);
    const item = buildNodeItem({
      id: normalizedId, name, kind, client, address, workPhone, cellPhone,
      emails, photo, taxId, accountingUrl, hrUrl, logo, customFields,
      createdAt: existing?.createdAt || now,
      updatedAt: now,
    });
    await putNode(item);
    res.json(toNodeResponse(item));
  } catch (err) {
    console.error("/api/nodes error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── Helper shared by both nodes-csv endpoints ────────────────────────────────

const importNodeRows = async (rows, client) => {
  const now = new Date().toISOString();
  const importBatchId = now;
  const items = rows.map((row) =>
    buildNodeItem({
      id: row.id, name: row.name, kind: row.kind, client,
      createdAt: now, updatedAt: now,
    })
  );
  // Preserve existing createdAt values in a single batch read + write pass.
  // For simplicity at import scale we batch-write directly (upsert semantics).
  Object.assign(items, items.map((item) => ({ ...item, importBatchId })));
  await batchPutNodes(items.map((item) => ({ ...item, importBatchId })));
  return { importBatchId, entities: rows.filter((r) => r.kind === "entity").length, persons: rows.filter((r) => r.kind === "person").length };
};

app.post("/api/import/nodes-csv", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { csv, defaultKind } = req.body || {};
    if (!csv) {
      return res.status(400).json({ error: "csv is required" });
    }
    const { rows, skipped } = parseNodeCsv(csv, defaultKind || "entity", client);
    if (!rows.length) {
      return res.status(400).json({ error: "no valid rows found in csv" });
    }
    const { importBatchId, entities, persons } = await importNodeRows(rows, client);
    res.json({ ok: true, total: rows.length, entities, persons, skipped, importBatchId });
  } catch (err) {
    console.error("/api/import/nodes-csv error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/import/nodes-csv/upload", upload.single("file"), async (req, res) => {
  try {
    const client = req.auth.clientId;
    const defaultKind = req.body?.defaultKind || "entity";
    if (!req.file?.buffer) return res.status(400).json({ error: "file is required" });

    const csv = req.file.buffer.toString("utf-8");
    const { rows, skipped } = parseNodeCsv(csv, defaultKind, client);
    if (!rows.length) {
      return res.status(400).json({ error: "no valid rows found in csv" });
    }
    const { importBatchId, entities, persons } = await importNodeRows(rows, client);
    res.json({ ok: true, total: rows.length, entities, persons, skipped, importBatchId });
  } catch (err) {
    console.error("/api/import/nodes-csv/upload error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── Helper shared by both ownerships-csv endpoints ───────────────────────────

const resolveOwnershipRows = async (rows, client) => {
  // Load all nodes for this client once and build a name→id map in memory.
  const allNodes = await queryNodes(client);
  const nameMap = new Map();
  allNodes.forEach((item) => {
    if (item.name) nameMap.set(item.name.toLowerCase(), item.id);
  });

  const resolved = [];
  const errors = [];
  rows.forEach((row, idx) => {
    const ownerRaw = String(row.owner || "").trim();
    const ownedRaw = String(row.owned || "").trim();
    const percent = Number(row.percent);

    const ownerId = ownerRaw.includes(":")
      ? normalizeClientId(client, ownerRaw)
      : nameMap.get(ownerRaw.toLowerCase());
    const ownedId = ownedRaw.includes(":")
      ? normalizeClientId(client, ownedRaw)
      : nameMap.get(ownedRaw.toLowerCase());

    if (!ownerId) { errors.push({ row: idx + 1, reason: "owner not found", owner: ownerRaw }); return; }
    if (!ownedId) { errors.push({ row: idx + 1, reason: "owned entity not found", owned: ownedRaw }); return; }
    if (!Number.isFinite(percent)) { errors.push({ row: idx + 1, reason: "percent missing or invalid", percent: row.percent }); return; }

    resolved.push({ from: ownerId, to: ownedId, percent });
  });
  return { resolved, errors };
};

const importOwnershipRows = async (resolved, client) => {
  const now = new Date().toISOString();
  const importBatchId = now;
  const items = resolved.map((r) =>
    buildOwnsItem({ ...r, client, createdAt: now, updatedAt: now, importBatchId })
  );
  await batchPutRels(items);
  return importBatchId;
};

app.post("/api/import/ownerships-csv", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { csv } = req.body || {};
    if (!csv) {
      return res.status(400).json({ error: "csv is required" });
    }
    const { rows, skipped } = parseOwnershipCsv(csv);
    if (!rows.length) {
      return res.status(400).json({ error: "no valid rows found in csv" });
    }
    const { resolved, errors } = await resolveOwnershipRows(rows, client);
    if (!resolved.length) {
      return res.status(400).json({ error: "no valid ownership rows to import", skipped: skipped + errors.length, errors });
    }
    const importBatchId = await importOwnershipRows(resolved, client);
    res.json({ ok: true, total: rows.length, imported: resolved.length, skipped: skipped + errors.length, errors, importBatchId });
  } catch (err) {
    console.error("/api/import/ownerships-csv error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/import/ownerships-csv/upload", upload.single("file"), async (req, res) => {
  try {
    const client = req.auth.clientId;
    if (!req.file?.buffer) return res.status(400).json({ error: "file is required" });

    const csv = req.file.buffer.toString("utf-8");
    const { rows, skipped } = parseOwnershipCsv(csv);
    if (!rows.length) {
      return res.status(400).json({ error: "no valid rows found in csv" });
    }
    const { resolved, errors } = await resolveOwnershipRows(rows, client);
    if (!resolved.length) {
      return res.status(400).json({ error: "no valid ownership rows to import", skipped: skipped + errors.length, errors });
    }
    const importBatchId = await importOwnershipRows(resolved, client);
    res.json({ ok: true, total: rows.length, imported: resolved.length, skipped: skipped + errors.length, errors, importBatchId });
  } catch (err) {
    console.error("/api/import/ownerships-csv/upload error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.put("/api/nodes/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const client = req.auth.clientId;
    const {
      name, kind, newId, address, workPhone, cellPhone, emails,
      photo, taxId, accountingUrl, hrUrl, logo, customFields,
    } = req.body || {};
    if (!id || !name || !kind) {
      return res.status(400).json({ error: "id, name, kind are required" });
    }
    const normalizedId = normalizeClientId(client, id);
    const normalizedNewId = newId ? normalizeClientId(client, newId) : null;
    const now = new Date().toISOString();

    const existing = await getNode(client, normalizedId);
    if (!existing) return res.status(404).json({ error: "node not found" });

    const finalId = (normalizedNewId && normalizedNewId !== normalizedId)
      ? normalizedNewId
      : normalizedId;

    const item = buildNodeItem({
      id: finalId, name, kind, client, address, workPhone, cellPhone,
      emails, photo, taxId, accountingUrl, hrUrl, logo, customFields,
      createdAt: existing.createdAt || now,
      updatedAt: now,
    });
    await putNode(item);

    // If the ID changed, re-key every relationship that referenced the old ID.
    if (finalId !== normalizedId) {
      const [fromRels, toRels] = await Promise.all([
        queryRelsByFrom(normalizedId),
        queryRelsByTo(normalizedId),
      ]);
      await Promise.all([
        ...fromRels.map((r) =>
          putRel({ ...r, clientId: client, relKey: makeRelKey(r.type, finalId, r.to), from: finalId, id: `${r.type}-${finalId}-${r.to}` })
            .then(() => deleteRel(client, r.relKey))
        ),
        ...toRels.map((r) =>
          putRel({ ...r, clientId: client, relKey: makeRelKey(r.type, r.from, finalId), to: finalId, id: `${r.type}-${r.from}-${finalId}` })
            .then(() => deleteRel(client, r.relKey))
        ),
      ]);
      await deleteNode(client, normalizedId);
    }

    res.json(toNodeResponse(item));
  } catch (err) {
    console.error("/api/nodes/:id error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.delete("/api/nodes/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const client = req.auth.clientId;
    if (!id) {
      return res.status(400).json({ error: "id is required" });
    }
    const normalizedId = normalizeClientId(client, id);
    // Detach-delete: remove all relationships first, then the node.
    await deleteRelsForNode(client, normalizedId);
    await deleteNode(client, normalizedId);
    res.json({ id: normalizedId });
  } catch (err) {
    console.error("/api/nodes/:id (delete) error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── OWNS ─────────────────────────────────────────────────────────────────────

app.post("/api/relationships/owns", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to, percent, startDate, endDate } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    const relKey = makeRelKey("owns", fromId, toId);
    const now = new Date().toISOString();
    const existing = await getRel(client, relKey);
    const item = buildOwnsItem({
      from: fromId, to: toId, percent: percent ?? null,
      startDate: startDate || null, endDate: endDate || null, client,
      createdAt: existing?.createdAt || now, updatedAt: now,
    });
    await putRel(item);
    res.json(toRelResponse(item));
  } catch (err) {
    console.error("/api/relationships/owns error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.put("/api/relationships/owns", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to, percent, startDate, endDate } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    const relKey = makeRelKey("owns", fromId, toId);
    const existing = await getRel(client, relKey);
    if (!existing) return res.status(404).json({ error: "relationship not found" });
    const item = buildOwnsItem({
      from: fromId, to: toId, percent: percent ?? null,
      startDate: startDate || null, endDate: endDate || null, client,
      createdAt: existing.createdAt, updatedAt: new Date().toISOString(),
    });
    await putRel(item);
    res.json(toRelResponse(item));
  } catch (err) {
    console.error("/api/relationships/owns (put) error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.delete("/api/relationships/owns", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    await deleteRel(client, makeRelKey("owns", fromId, toId));
    res.json({ from: fromId, to: toId });
  } catch (err) {
    console.error("/api/relationships/owns (delete) error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── EMPLOYS ──────────────────────────────────────────────────────────────────

app.post("/api/relationships/employs", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to, role, startDate, endDate } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    const relKey = makeRelKey("employs", fromId, toId);
    const now = new Date().toISOString();
    const existing = await getRel(client, relKey);
    const item = buildEmploysItem({
      from: fromId, to: toId, role: role || null,
      startDate: startDate || null, endDate: endDate || null, client,
      createdAt: existing?.createdAt || now, updatedAt: now,
    });
    await putRel(item);
    res.json(toRelResponse(item));
  } catch (err) {
    console.error("/api/relationships/employs error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.put("/api/relationships/employs", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to, role, startDate, endDate } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    const relKey = makeRelKey("employs", fromId, toId);
    const existing = await getRel(client, relKey);
    if (!existing) return res.status(404).json({ error: "relationship not found" });
    const item = buildEmploysItem({
      from: fromId, to: toId, role: role || null,
      startDate: startDate || null, endDate: endDate || null, client,
      createdAt: existing.createdAt, updatedAt: new Date().toISOString(),
    });
    await putRel(item);
    res.json(toRelResponse(item));
  } catch (err) {
    console.error("/api/relationships/employs (put) error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.delete("/api/relationships/employs", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { from, to } = req.body || {};
    if (!from || !to) {
      return res.status(400).json({ error: "from and to are required" });
    }
    const fromId = normalizeClientId(client, from);
    const toId = normalizeClientId(client, to);
    await deleteRel(client, makeRelKey("employs", fromId, toId));
    res.json({ from: fromId, to: toId });
  } catch (err) {
    console.error("/api/relationships/employs (delete) error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── FILE UPLOAD ──────────────────────────────────────────────────────────────

// POST /api/upload — upload a file to S3 as a public-read object and return its
// permanent public URL.  The object key contains 32 random hex chars so the URL
// is not guessable.  Requires S3_BUCKET_UPLOADS env var; returns 501 if not set.
//
// Bucket prerequisites (set once in the AWS console or via CLI):
//   • "Block all public access" → OFF
//   • Bucket policy granting s3:GetObject on arn:aws:s3:::BUCKET/uploads/*
app.post("/api/upload", upload.single("file"), async (req, res) => {
  try {
    if (!s3 || !S3_BUCKET) {
      return res.status(501).json({ error: "File upload not configured (set S3_BUCKET_UPLOADS env var)" });
    }
    if (!req.file?.buffer) return res.status(400).json({ error: "file is required" });

    // Derive a safe extension from the original file name.
    const original = req.file.originalname || "upload";
    const ext = original.includes(".")
      ? original.split(".").pop().replace(/[^a-zA-Z0-9]/g, "").toLowerCase().slice(0, 10)
      : "bin";
    const key = `uploads/${req.auth.clientId}/${randomBytes(16).toString("hex")}.${ext}`;
    const region = process.env.AWS_REGION || "us-east-1";

    // Derive content type — prefer multer's detection but fall back by extension
    const MIME_BY_EXT = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp", pdf: "application/pdf" };
    const contentType = (req.file.mimetype && req.file.mimetype !== "application/octet-stream")
      ? req.file.mimetype
      : (MIME_BY_EXT[ext] || "application/octet-stream");

    await s3.send(
      new PutObjectCommand({
        Bucket: S3_BUCKET,
        Key: key,
        Body: req.file.buffer,
        ContentType: contentType,
      })
    );

    // Permanent public URL — never expires.
    const url = `https://${S3_BUCKET}.s3.${region}.amazonaws.com/${key}`;
    res.json({ url, key });
  } catch (err) {
    console.error("/api/upload error", err);
    res.status(500).json({ error: err.message });
  }
});

// ─── IMAGE PROXY ─────────────────────────────────────────────────────────────

// GET /api/proxy-image?url=<encoded-s3-url>
// Fetches a file from our own S3 bucket server-side and returns it to the
// client.  This lets html2canvas render cross-origin S3 images without needing
// a CORS policy on the bucket.  Only URLs that start with our configured bucket
// hostname are allowed to prevent open-proxy abuse.
app.get("/api/proxy-image", async (req, res) => {
  try {
    const { url } = req.query;
    if (!url) return res.status(400).json({ error: "url is required" });

    if (!S3_BUCKET || !url.startsWith(`https://${S3_BUCKET}.s3.`)) {
      return res.status(403).json({ error: "URL not permitted" });
    }

    const upstream = await fetch(url);
    if (!upstream.ok) return res.status(upstream.status).json({ error: "Upstream error" });

    const buffer = await upstream.arrayBuffer();
    const upstreamType = upstream.headers.get("content-type") || "";
    const ext = url.split(".").pop().split("?")[0].toLowerCase();
    const MIME_BY_EXT = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp" };
    const mimeType = upstreamType.startsWith("image/") ? upstreamType : (MIME_BY_EXT[ext] || "image/jpeg");
    const b64 = Buffer.from(buffer).toString("base64");
    const dataUrl = `data:${mimeType};base64,${b64}`;

    // Return as JSON — avoids API Gateway binary encoding issues entirely
    res.json({ dataUrl });
  } catch (err) {
    console.error("/api/proxy-image error", err);
    res.status(500).json({ error: err.message });
  }
});

// ─── DATA DICTIONARY ──────────────────────────────────────────────────────────

const VALID_DATA_TYPES = new Set([
  "string", "textarea", "number", "currency", "percentage",
  "boolean", "date", "time", "phone", "email", "link", "file", "address", "year",
]);

const VALID_APPLIES_TO = new Set(["entity", "person", "both"]);

const toDDResponse = ({ clientId: _c, ...rest }) => rest;

// GET /api/data-dictionary?client=<clientId>
app.get("/api/data-dictionary", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const items = await queryDDFields(client);
    res.json(items.map(toDDResponse));
  } catch (err) {
    console.error("/api/data-dictionary GET error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// POST /api/data-dictionary
app.post("/api/data-dictionary", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { prompt, dataType, appliesTo, multiValue, validValues, phoneTypes } = req.body || {};
    if (!prompt || !prompt.trim()) {
      return res.status(400).json({ error: "prompt is required" });
    }
    if (!VALID_DATA_TYPES.has(dataType)) {
      return res.status(400).json({ error: `invalid dataType: ${dataType}` });
    }
    if (!VALID_APPLIES_TO.has(appliesTo)) {
      return res.status(400).json({ error: `invalid appliesTo: ${appliesTo}` });
    }
    const fieldId = `dd:${slugify(prompt.trim())}-${Date.now().toString(36)}`;
    const now = new Date().toISOString();
    const existingFields = await queryDDFields(client);
    const item = {
      clientId: client,
      fieldId,
      id: fieldId,
      prompt: prompt.trim(),
      dataType,
      appliesTo: appliesTo || "both",
      multiValue: !!multiValue,
      validValues: Array.isArray(validValues) ? validValues.filter(Boolean) : [],
      phoneTypes: Array.isArray(phoneTypes) ? phoneTypes.filter(Boolean) : [],
      sortOrder: existingFields.length,
      createdAt: now,
      updatedAt: now,
    };
    await putDDField(item);
    res.status(201).json(toDDResponse(item));
  } catch (err) {
    console.error("/api/data-dictionary POST error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// PUT /api/data-dictionary/:fieldId
app.put("/api/data-dictionary/:fieldId", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const fieldId = req.params.fieldId;
    const { prompt, dataType, appliesTo, multiValue, validValues, phoneTypes } = req.body || {};
    if (!prompt || !prompt.trim()) {
      return res.status(400).json({ error: "prompt is required" });
    }
    if (!VALID_DATA_TYPES.has(dataType)) {
      return res.status(400).json({ error: `invalid dataType: ${dataType}` });
    }
    if (!VALID_APPLIES_TO.has(appliesTo)) {
      return res.status(400).json({ error: `invalid appliesTo: ${appliesTo}` });
    }
    const existing = await getDDField(client, fieldId);
    if (!existing) return res.status(404).json({ error: "field not found" });
    const item = {
      ...existing,
      prompt: prompt.trim(),
      dataType,
      appliesTo: appliesTo || "both",
      multiValue: !!multiValue,
      validValues: Array.isArray(validValues) ? validValues.filter(Boolean) : [],
      phoneTypes: Array.isArray(phoneTypes) ? phoneTypes.filter(Boolean) : [],
      updatedAt: new Date().toISOString(),
    };
    await putDDField(item);
    res.json(toDDResponse(item));
  } catch (err) {
    console.error("/api/data-dictionary PUT error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// DELETE /api/data-dictionary/:fieldId
app.delete("/api/data-dictionary/:fieldId", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const fieldId = req.params.fieldId;
    const existing = await getDDField(client, fieldId);
    if (!existing) return res.status(404).json({ error: "field not found" });
    await deleteDDField(client, fieldId);
    res.json({ fieldId });
  } catch (err) {
    console.error("/api/data-dictionary DELETE error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// PUT /api/data-dictionary/:fieldId/reorder
app.put("/api/data-dictionary/:fieldId/reorder", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { fieldId } = req.params;
    const { direction } = req.body || {};
    if (direction !== "up" && direction !== "down") {
      return res.status(400).json({ error: "direction must be 'up' or 'down'" });
    }
    const all = await queryDDFields(client);
    // Normalize to consecutive integers so the swap is always predictable,
    // even for legacy entries created before sortOrder was introduced.
    const sorted = [...all]
      .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
      .map((f, i) => ({ ...f, sortOrder: i }));
    const idx = sorted.findIndex((f) => f.fieldId === fieldId);
    if (idx === -1) return res.status(404).json({ error: "field not found" });
    const swapIdx = direction === "up" ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= sorted.length) {
      return res.status(400).json({ error: "cannot move in that direction" });
    }
    const now = new Date().toISOString();
    // Apply the swap and persist every item so all sortOrders are normalized.
    const toSave = sorted.map((f, i) => {
      if (i === idx) return { ...f, sortOrder: swapIdx, updatedAt: now };
      if (i === swapIdx) return { ...f, sortOrder: idx, updatedAt: now };
      return f;
    });
    await Promise.all(toSave.map((f) => putDDField(f)));
    res.json(toSave.map(toDDResponse));
  } catch (err) {
    console.error("/api/data-dictionary reorder error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── EXPORT REPORTS ───────────────────────────────────────────────────────────

const toExportReportResponse = ({ clientId: _c, ...rest }) => rest;

const slugifyReportName = (name) =>
  String(name || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "report";

// GET /api/export-reports — list all saved export reports for this client
app.get("/api/export-reports", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const items = await queryExportReports(client);
    res.json(items.map(toExportReportResponse));
  } catch (err) {
    console.error("/api/export-reports GET error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// PUT /api/export-reports/:reportId — upsert (create or overwrite by name-slug key)
app.put("/api/export-reports/:reportId", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { reportId } = req.params;
    const { name, fieldIds } = req.body || {};
    if (!name || !name.trim()) {
      return res.status(400).json({ error: "name is required" });
    }
    if (!Array.isArray(fieldIds)) {
      return res.status(400).json({ error: "fieldIds must be an array" });
    }
    const now = new Date().toISOString();
    const existing = await getExportReport(client, reportId);
    const item = {
      clientId: client,
      reportId,
      id: reportId,
      name: name.trim(),
      fieldIds,
      createdAt: existing?.createdAt || now,
      updatedAt: now,
    };
    await putExportReport(item);
    res.json(toExportReportResponse(item));
  } catch (err) {
    console.error("/api/export-reports PUT error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// DELETE /api/export-reports/:reportId
app.delete("/api/export-reports/:reportId", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { reportId } = req.params;
    const existing = await getExportReport(client, reportId);
    if (!existing) return res.status(404).json({ error: "report not found" });
    await deleteExportReport(client, reportId);
    res.json({ reportId });
  } catch (err) {
    console.error("/api/export-reports DELETE error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// GET /api/auth/me — return the current user's profile (no password hash).
// Also includes clientName from the client record so all users can display it.
app.get("/api/auth/me", async (req, res) => {
  try {
    const user = await getUser(req.auth.loginId);
    if (!user) return res.status(404).json({ error: "User not found" });
    const { passwordHash, ...profile } = user;
    // Attach clientName so the UI can use it without a separate admin-only request
    const clientRecord = await getClient(req.auth.clientId).catch(() => null);
    if (clientRecord?.clientName) profile.clientName = clientRecord.clientName;
    res.json(profile);
  } catch (err) {
    console.error("/api/auth/me GET error", err);
    res.status(500).json({ error: err.message });
  }
});

// PATCH /api/auth/me — update editable profile fields.
app.patch("/api/auth/me", async (req, res) => {
  try {
    const user = await getUser(req.auth.loginId);
    if (!user) return res.status(404).json({ error: "User not found" });
    const { name, email, cellPhone, workPhone, homeScreen } = req.body || {};
    const updated = {
      ...user,
      ...(name       !== undefined && { name:       String(name).trim() }),
      ...(email      !== undefined && { email:      String(email).trim() }),
      ...(cellPhone  !== undefined && { cellPhone:  String(cellPhone).trim() }),
      ...(workPhone  !== undefined && { workPhone:  String(workPhone).trim() }),
      ...(homeScreen !== undefined && { homeScreen: homeScreen }),
    };
    await putUser(updated);
    const { passwordHash, ...profile } = updated;
    res.json(profile);
  } catch (err) {
    console.error("/api/auth/me PATCH error", err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/auth/users — add a user to the currently authenticated client.
// JWT required; admin role required; new user is scoped to req.auth.clientId.
app.post("/api/auth/users", requireAdmin, async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { loginId, password, personId, role } = req.body || {};
    if (!loginId || !password) {
      return res.status(400).json({ error: "loginId and password are required" });
    }
    const normalizedId = String(loginId).toLowerCase().trim();
    const existing = await getUser(normalizedId);
    if (existing) return res.status(409).json({ error: "loginId already exists" });
    const passwordHash = await bcrypt.hash(password, 10);
    const newRole = role === "admin" ? "admin" : "user";
    const user = {
      loginId: normalizedId,
      clientId: client,
      personId: personId || null,
      role: newRole,
      passwordHash,
      createdAt: new Date().toISOString(),
    };
    await putUser(user);
    res.status(201).json({ loginId: user.loginId, clientId: user.clientId, role: user.role });
  } catch (err) {
    console.error("/api/auth/users error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// GET /api/auth/users — list all users for this client (admin only).
app.get("/api/auth/users", requireAdmin, async (req, res) => {
  try {
    const users = await listUsersByClient(req.auth.clientId);
    const response = users.map(({ passwordHash: _, ...u }) => u);
    res.json(response);
  } catch (err) {
    console.error("/api/auth/users GET error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// PATCH /api/auth/users/:loginId — change a user's role (admin only).
// Only users belonging to req.auth.clientId may be updated.
app.patch("/api/auth/users/:loginId", requireAdmin, async (req, res) => {
  try {
    const client = req.auth.clientId;
    const targetId = String(req.params.loginId).toLowerCase().trim();
    const { role } = req.body || {};
    if (role !== "admin" && role !== "user") {
      return res.status(400).json({ error: "role must be \"admin\" or \"user\"" });
    }
    const user = await getUser(targetId);
    if (!user) return res.status(404).json({ error: "User not found" });
    if (user.clientId !== client) return res.status(403).json({ error: "User belongs to a different client" });
    await putUser({ ...user, role });
    res.json({ loginId: user.loginId, clientId: user.clientId, role });
  } catch (err) {
    console.error("/api/auth/users/:loginId PATCH error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── CLIENT INFO ──────────────────────────────────────────────────────────────

// GET /api/client — return this client's info record (admin only).
app.get("/api/client", requireAdmin, async (req, res) => {
  try {
    const clientId = req.auth.clientId;
    const record = await getClient(clientId);
    if (!record) return res.json({});
    const { clientId: _c, ...response } = record;
    res.json(response);
  } catch (err) {
    console.error("/api/client GET error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// PATCH /api/client — create or update this client's info record (admin only).
app.patch("/api/client", requireAdmin, async (req, res) => {
  try {
    const clientId = req.auth.clientId;
    const { clientName, address, billingContact, billingEmail, billingPhone, notes } = req.body || {};
    const existing = await getClient(clientId);
    const now = new Date().toISOString();
    const updated = {
      clientId,
      ...(existing || { createdAt: now }),
      ...(clientName     !== undefined && { clientName:     String(clientName).trim() }),
      ...(address        !== undefined && { address:        String(address).trim() }),
      ...(billingContact !== undefined && { billingContact: String(billingContact).trim() }),
      ...(billingEmail   !== undefined && { billingEmail:   String(billingEmail).trim() }),
      ...(billingPhone   !== undefined && { billingPhone:   String(billingPhone).trim() }),
      ...(notes          !== undefined && { notes:          String(notes).trim() }),
      updatedAt: now,
    };
    await putClient(updated);
    const { clientId: _c, ...response } = updated;
    res.json(response);
  } catch (err) {
    console.error("/api/client PATCH error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

export default app;
