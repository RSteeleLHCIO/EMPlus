import "dotenv/config";
import express from "express";
import cors from "cors";
import multer from "multer";
import { parse } from "csv-parse/sync";
import jwt from "jsonwebtoken";
import bcrypt from "bcryptjs";
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
} from "./dynamoClient.js";
import {
  getUser,
  putUser,
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
    taxId, accountingUrl, hrUrl, logo, createdAt, updatedAt }
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
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });
app.use(
  cors({
    origin: true,
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);
app.options("*", cors());
app.use(express.json({ limit: "1mb" }));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "dynamodb-api" });
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
    res.json({ token, clientId: user.clientId, personId: user.personId || null });
  } catch (err) {
    console.error("/api/auth/login error", err);
    res.status(500).json({ error: err.message });
  }
});

// Bootstrap: register the first user for a client.
// Requires header X-Setup-Secret matching SETUP_SECRET env var (if set),
// OR the client must have zero existing users.
app.post("/api/auth/register", async (req, res) => {
  try {
    const { loginId, password, clientId, personId, setupSecret } = req.body || {};
    if (!loginId || !password || !clientId) {
      return res.status(400).json({ error: "loginId, password, clientId are required" });
    }
    // Security gate: either SETUP_SECRET matches, or no users exist yet for this client
    if (SETUP_SECRET && setupSecret !== SETUP_SECRET) {
      const existing = await countUsersByClient(clientId);
      if (existing > 0) {
        return res.status(403).json({ error: "Registration not permitted" });
      }
    }
    const normalizedId = String(loginId).toLowerCase().trim();
    const existing = await getUser(normalizedId);
    if (existing) return res.status(409).json({ error: "loginId already exists" });
    const passwordHash = await bcrypt.hash(password, 10);
    const user = {
      loginId: normalizedId,
      clientId,
      personId: personId || null,
      passwordHash,
      createdAt: new Date().toISOString(),
    };
    await putUser(user);
    res.status(201).json({ loginId: user.loginId, clientId: user.clientId });
  } catch (err) {
    console.error("/api/auth/register error", err);
    res.status(500).json({ error: err.message });
  }
});

// ─── JWT MIDDLEWARE (all /api/* routes below this require a valid token) ──────

app.use("/api", (req, res, next) => {
  // Skip auth routes
  if (req.path.startsWith("/auth/")) return next();
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
      photo, taxId, accountingUrl, hrUrl, logo,
    } = req.body || {};
    if (!id || !name || !kind) {
      return res.status(400).json({ error: "id, name, kind are required" });
    }
    const normalizedId = normalizeClientId(client, id);
    const now = new Date().toISOString();
    const existing = await getNode(client, normalizedId);
    const item = buildNodeItem({
      id: normalizedId, name, kind, client, address, workPhone, cellPhone,
      emails, photo, taxId, accountingUrl, hrUrl, logo,
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
      photo, taxId, accountingUrl, hrUrl, logo,
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
      emails, photo, taxId, accountingUrl, hrUrl, logo,
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

export default app;
