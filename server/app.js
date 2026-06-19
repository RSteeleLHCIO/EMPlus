import "dotenv/config";
import express from "express";
import cors from "cors";
import multer from "multer";
import { parse } from "csv-parse/sync";
import XLSX from "xlsx";
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

// ─── Fuzzy header → DD field mapping ─────────────────────────────────────────

// Maps known user column header variants to canonical DD field names.
const FIELD_SYNONYMS = {
  // NOTE: "kind" is intentionally absent — the details upload must never alter a node's kind.
  name:            ["name", "company name", "entity name", "node name", "organization", "org name", "business name", "legal name", "entity or person s name"],
  address:         ["address", "street", "street address", "mailing address", "location", "addr"],
  workPhone:       ["work phone", "phone", "workphone", "office phone", "ph work", "business phone", "telephone", "tel", "phone number", "primary phone"],
  cellPhone:       ["cell phone", "cell", "mobile", "mobile phone", "cellphone", "cell number", "personal phone"],
  emails:          ["email", "emails", "email address", "e mail", "email addr"],
  taxId:           ["tax id", "taxid", "ein", "tin", "federal id", "tax identification", "federal tax id", "fein"],
  accountingUrl:   ["accounting url", "accounting", "accountingurl", "acctg link", "quickbooks url", "quickbooks", "accounting link", "acctg url"],
  hrUrl:           ["hr url", "hrurl", "hr link", "hr system", "hris url", "hris", "human resources url"],
  logo:            ["logo", "logo url", "logo link", "company logo"],
  photo:           ["photo", "photo url", "picture", "headshot", "profile photo", "image"],
  legalStatus:     ["legal status", "legalstatus", "entity structure", "structure", "corp type", "company type", "incorporation type"],
  operationalRole: ["operational role", "operationalrole", "role", "function", "operational function", "business role"],
  personStatus:    ["person status", "personstatus", "status", "employment status"],
};

// Strips punctuation, lowercases, and collapses whitespace in a header string.
const normalizeHeader = (str) =>
  String(str || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

// Jaccard token-overlap score between two normalized strings.
const tokenOverlap = (a, b) => {
  const ta = new Set(a.split(" ").filter(Boolean));
  const tb = new Set(b.split(" ").filter(Boolean));
  if (ta.size === 0 && tb.size === 0) return 0;
  const intersection = [...ta].filter((t) => tb.has(t)).length;
  const union = new Set([...ta, ...tb]).size;
  return intersection / union;
};

// Maps an array of raw column headers to DD field names.
// Returns:
//   fieldMap — { columnIndex: { field, confidence: "exact"|"fuzzy" } }
//   guessed  — { rawHeader: fieldName } for fuzzy-matched columns (surfaced as warnings)
//   extra    — rawHeaders with no DD match (will be stored in customFields)
const mapHeadersToFields = (headers) => {
  const fieldMap = {};
  const guessed = {};
  const extra = [];

  headers.forEach((rawHeader, idx) => {
    const normalized = normalizeHeader(rawHeader);
    if (!normalized) return; // skip blank column headers

    // Exact match against synonym lists
    let matched = null;
    for (const [field, synonyms] of Object.entries(FIELD_SYNONYMS)) {
      if (synonyms.includes(normalized)) {
        matched = { field, confidence: "exact" };
        break;
      }
    }

    // Fuzzy fallback: best token-overlap score ≥ 0.5
    if (!matched) {
      let bestScore = 0;
      let bestField = null;
      for (const [field, synonyms] of Object.entries(FIELD_SYNONYMS)) {
        for (const synonym of synonyms) {
          const score = tokenOverlap(normalized, synonym);
          if (score > bestScore) { bestScore = score; bestField = field; }
        }
      }
      if (bestScore >= 0.5) matched = { field: bestField, confidence: "fuzzy" };
    }

    if (matched) {
      fieldMap[idx] = matched;
      if (matched.confidence === "fuzzy") guessed[rawHeader] = matched.field;
    } else {
      extra.push(rawHeader);
    }
  });

  return { fieldMap, guessed, extra };
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

// ─── Node details CSV/XLSX parser ────────────────────────────────────────────

// Converts a parsed records array (header row + data rows) into patch objects.
// Each row produces { name, patch } where patch contains only fields present in the file.
// Columns that don't match any DD field are stored under customFields.
const parseRecordsToDetailRows = (records) => {
  if (!Array.isArray(records) || records.length < 2) return { rows: [], mapping: {}, skipped: 0 };

  const headers = records[0].map((h) => String(h ?? ""));
  const { fieldMap, guessed, extra } = mapHeadersToFields(headers);

  // A "name" column is required to identify each node.
  const nameColIdx = Number(
    Object.entries(fieldMap).find(([, v]) => v.field === "name")?.[0] ?? -1
  );
  if (nameColIdx === -1) {
    throw Object.assign(new Error('CSV must contain a "name" column to identify nodes'), { status: 400 });
  }

  // Map column indices for unrecognised headers → stored in customFields.
  const extraColMap = {};
  headers.forEach((rawHeader, idx) => {
    if (extra.includes(rawHeader) && rawHeader.trim()) extraColMap[idx] = rawHeader.trim();
  });

  const rows = [];
  let skipped = 0;
  for (let i = 1; i < records.length; i++) {
    const record = records[i] || [];
    const name = String(record[nameColIdx] ?? "").trim();
    if (!name || name.startsWith("[")) { skipped += 1; continue; }

    const patch = {};
    for (const [idxStr, { field }] of Object.entries(fieldMap)) {
      if (field === "name") continue;
      const value = String(record[Number(idxStr)] ?? "").trim();
      if (value !== "") patch[field] = value;
    }

    // Non-DD columns stored separately so applyDetailRows can assign proper DD fieldIds.
    const extraValues = {};
    for (const [idxStr, header] of Object.entries(extraColMap)) {
      const value = String(record[Number(idxStr)] ?? "").trim();
      if (value !== "") extraValues[header] = value;
    }

    rows.push({ name, patch, extraValues });
  }

  return { rows, extraHeaders: extra, mapping: { guessed, extra }, skipped };
};

const parseDetailsCsv = (csvText) => {
  if (!csvText || typeof csvText !== "string") return { rows: [], mapping: {}, skipped: 0 };
  const records = parse(csvText, { skip_empty_lines: true, relax_column_count: true, trim: true });
  return parseRecordsToDetailRows(records);
};

const parseDetailsXlsx = (buffer) => {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  return parseRecordsToDetailRows(records);
};

const xlsxBufferToCsv = (buffer) => {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_csv(sheet);
};

// ─── DynamoDB item builders ───────────────────────────────────────────────────

const buildNodeItem = (
  { id, name, kind, client, address, workPhone, cellPhone, emails, photo,
    taxId, accountingUrl, hrUrl, logo, customFields,
    operationalRole, legalStatus, personStatus,
    createdAt, updatedAt }
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
  operationalRole: operationalRole || "",
  legalStatus: legalStatus || "",
  personStatus: personStatus || "",
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
  res.json({ ok: true, service: "dynamodb-api", version: "2026-05-06.operationalRole", features: ["client management"] });
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
      operationalRole, legalStatus, personStatus,
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
      operationalRole, legalStatus, personStatus,
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

    const ext = (req.file.originalname || "").split(".").pop().toLowerCase();
    const isXlsx = ext === "xlsx" || ext === "xls";
    const csv = isXlsx ? xlsxBufferToCsv(req.file.buffer) : req.file.buffer.toString("utf-8");
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

    const ext = (req.file.originalname || "").split(".").pop().toLowerCase();
    const isXlsx = ext === "xlsx" || ext === "xls";
    const csv = isXlsx ? xlsxBufferToCsv(req.file.buffer) : req.file.buffer.toString("utf-8");
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

// ─── Import preview (parse-only, no writes) ───────────────────────────────────

app.post("/api/import/preview/upload", upload.single("file"), async (req, res) => {
  try {
    const client = req.auth.clientId;
    const importType = req.body?.importType || "entity";
    const defaultKind = req.body?.defaultKind || "entity";
    if (!req.file?.buffer) return res.status(400).json({ error: "file is required" });

    const ext = (req.file.originalname || "").split(".").pop().toLowerCase();
    const isXlsx = ext === "xlsx" || ext === "xls";
    const MAX_PREVIEW = 50;

    if (importType === "details") {
      let rawRecords;
      if (isXlsx) {
        const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        rawRecords = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      } else {
        rawRecords = parse(req.file.buffer.toString("utf-8"), { skip_empty_lines: true, relax_column_count: true, trim: true });
      }
      if (!rawRecords || rawRecords.length < 2) {
        return res.json({ headers: [], rows: [], total: 0, skipped: 0, truncated: false });
      }
      const { rows, skipped, mapping } = parseRecordsToDetailRows(rawRecords);
      const headers = rawRecords[0].map((h) => String(h ?? "")).filter((h) => h.trim());
      const previewRows = rawRecords.slice(1, MAX_PREVIEW + 1).map((record) =>
        headers.reduce((obj, h, i) => { obj[h] = String(record[i] ?? ""); return obj; }, {})
      );
      return res.json({ headers, rows: previewRows, total: rows.length, skipped, mapping, truncated: rows.length > MAX_PREVIEW });
    }

    if (importType === "ownership") {
      const csv = isXlsx ? xlsxBufferToCsv(req.file.buffer) : req.file.buffer.toString("utf-8");
      const { rows, skipped } = parseOwnershipCsv(csv);
      const headers = ["Owner", "Owned", "Percent"];
      const previewRows = rows.slice(0, MAX_PREVIEW).map((r) => ({
        Owner: r.owner,
        Owned: r.owned,
        Percent: Number.isNaN(r.percent) ? "" : String(r.percent),
      }));
      return res.json({ headers, rows: previewRows, total: rows.length, skipped, truncated: rows.length > MAX_PREVIEW });
    }

    // entity or person
    const csv = isXlsx ? xlsxBufferToCsv(req.file.buffer) : req.file.buffer.toString("utf-8");
    const { rows, skipped } = parseNodeCsv(csv, defaultKind, client);
    const headers = ["Name", "Kind"];
    const previewRows = rows.slice(0, MAX_PREVIEW).map((r) => ({ Name: r.name, Kind: r.kind }));
    return res.json({ headers, rows: previewRows, total: rows.length, skipped, truncated: rows.length > MAX_PREVIEW });
  } catch (err) {
    console.error("/api/import/preview/upload error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// ─── Helper shared by both details endpoints ──────────────────────────────────

// Patches existing nodes with fields extracted from a details file.
// Looks up each row by name (case-insensitive). Unknown columns that landed in
// customFields are merged (not replaced) with whatever the node already has.
// Resolves or auto-creates a DD field for each unrecognised column header.
// Returns a map of rawHeader → fieldId so values can be stored under the correct key.
const resolveExtraHeaders = async (extraHeaders, client, existingFields = null) => {
  if (!extraHeaders || extraHeaders.length === 0) return {};

  const ddFields = existingFields || await queryDDFields(client);
  // Build lookup: normalised prompt → fieldId
  const existingByPrompt = new Map(
    ddFields.map((f) => [normalizeHeader(f.prompt || ""), f.fieldId])
  );

  const headerToFieldId = {};
  const now = new Date().toISOString();
  const toCreate = [];

  for (const rawHeader of extraHeaders) {
    if (!rawHeader.trim()) continue;
    const normalized = normalizeHeader(rawHeader);
    if (existingByPrompt.has(normalized)) {
      headerToFieldId[rawHeader] = existingByPrompt.get(normalized);
    } else {
      const fieldId = `dd:${slugify(rawHeader.trim())}-${Date.now().toString(36)}-${toCreate.length}`;
      const item = {
        clientId: client,
        fieldId,
        id: fieldId,
        prompt: rawHeader.trim(),
        dataType: "string",
        appliesTo: "both",
        multiValue: false,
        validValues: [],
        phoneTypes: [],
        showInStats: false,
        sortOrder: ddFields.length + toCreate.length,
        createdAt: now,
        updatedAt: now,
      };
      toCreate.push(item);
      headerToFieldId[rawHeader] = fieldId;
      existingByPrompt.set(normalized, fieldId); // prevent duplicates within the same batch
    }
  }

  if (toCreate.length) await Promise.all(toCreate.map((f) => putDDField(f)));
  return headerToFieldId;
};

const applyDetailRows = async (rows, extraHeaders, client, guessed = {}) => {
  const [allNodes, ddFields] = await Promise.all([
    queryNodes(client),
    queryDDFields(client),
  ]);

  // If a fuzzy-matched built-in header is an EXACT match for a DD field prompt,
  // the DD field takes precedence — re-route that column back to extraValues.
  const ddNormToPrompt = new Map(ddFields.map((f) => [normalizeHeader(f.prompt || ""), f.prompt]));
  const redirect = {}; // builtInFieldName → rawHeader
  for (const [rawHeader, builtInField] of Object.entries(guessed)) {
    if (ddNormToPrompt.has(normalizeHeader(rawHeader))) {
      redirect[builtInField] = rawHeader;
    }
  }

  const hasRedirects = Object.keys(redirect).length > 0;
  const adjustedRows = hasRedirects
    ? rows.map(({ name, patch, extraValues = {} }) => {
        const newPatch = { ...patch };
        const newExtraValues = { ...extraValues };
        for (const [field, rawHeader] of Object.entries(redirect)) {
          if (field in newPatch) {
            newExtraValues[rawHeader] = newPatch[field];
            delete newPatch[field];
          }
        }
        return { name, patch: newPatch, extraValues: newExtraValues };
      })
    : rows;

  const adjustedExtraHeaders = hasRedirects
    ? [...extraHeaders, ...Object.values(redirect)]
    : extraHeaders;

  const [headerToFieldId] = await Promise.all([
    resolveExtraHeaders(adjustedExtraHeaders, client, ddFields),
  ]);
  const nameMap = new Map(allNodes.map((n) => [String(n.name || "").toLowerCase(), n]));

  const now = new Date().toISOString();
  const importBatchId = now;
  const updated = [];
  const notFoundList = [];

  for (const { name, patch, extraValues = {} } of adjustedRows) {
    // Match by display name or by raw/prefixed id value
    const existing =
      nameMap.get(name.toLowerCase()) ||
      allNodes.find((n) => n.id === normalizeClientId(client, name));
    if (!existing) { notFoundList.push(name); continue; }

    // Remap extra column values from raw header name → DD fieldId.
    const remappedExtra = {};
    for (const [rawHeader, value] of Object.entries(extraValues)) {
      const fieldId = headerToFieldId[rawHeader];
      if (fieldId) remappedExtra[fieldId] = value;
    }

    // Merge customFields: preserve existing keys; add/overwrite only keys present in this import.
    // Fields absent from the spreadsheet are NOT touched.
    const mergedCustomFields = {
      ...(existing.customFields || {}),
      ...remappedExtra,
    };

    // Never allow the details import to change a node's kind.
    delete patch.kind;

    const item = buildNodeItem({
      ...existing,
      ...patch,                  // only overwrites DD fields that were present & non-blank
      customFields: mergedCustomFields,
      emails: patch.emails !== undefined
        ? normalizeEmails(patch.emails)
        : existing.emails,
      client,
      createdAt: existing.createdAt || now,
      updatedAt: now,
    });
    updated.push({ ...item, importBatchId });
  }

  if (updated.length) await batchPutNodes(updated);
  const ddCreated = Object.entries(headerToFieldId);
  return {
    updated: updated.length,
    notFound: notFoundList.length,
    notFoundList,
    ddFieldsCreated: ddCreated.map(([header, fieldId]) => ({ header, fieldId })),
  };
};

app.post("/api/import/details-csv", async (req, res) => {
  try {
    const client = req.auth.clientId;
    const { csv } = req.body || {};
    if (!csv) return res.status(400).json({ error: "csv is required" });

    const { rows, extraHeaders, mapping, skipped } = parseDetailsCsv(csv);
    if (!rows.length) return res.status(400).json({ error: "no valid rows found", mapping, skipped });

    const result = await applyDetailRows(rows, extraHeaders, client, mapping.guessed || {});
    res.json({ ok: true, ...result, skipped, mapping });
  } catch (err) {
    console.error("/api/import/details-csv error", err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

app.post("/api/import/details-csv/upload", upload.single("file"), async (req, res) => {
  try {
    const client = req.auth.clientId;
    if (!req.file?.buffer) return res.status(400).json({ error: "file is required" });

    const ext = (req.file.originalname || "").split(".").pop().toLowerCase();
    const isXlsx = ext === "xlsx" || ext === "xls";
    const { rows, extraHeaders, mapping, skipped } = isXlsx
      ? parseDetailsXlsx(req.file.buffer)
      : parseDetailsCsv(req.file.buffer.toString("utf-8"));

    if (!rows.length) return res.status(400).json({ error: "no valid rows found", mapping, skipped });

    const result = await applyDetailRows(rows, extraHeaders, client, mapping.guessed || {});
    res.json({ ok: true, ...result, skipped, mapping });
  } catch (err) {
    console.error("/api/import/details-csv/upload error", err);
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
      operationalRole, legalStatus, personStatus,
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
      operationalRole, legalStatus, personStatus,
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
    const { prompt, dataType, appliesTo, multiValue, validValues, phoneTypes, showInStats } = req.body || {};
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
      showInStats: !!showInStats,
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
    const { prompt, dataType, appliesTo, multiValue, validValues, phoneTypes, showInStats } = req.body || {};
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
      showInStats: !!showInStats,
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
    const { name, email, cellPhone, workPhone, homeScreen, tabularViews, tabularViewsSelectedId } = req.body || {};
    const normalizedTabularViews = Array.isArray(tabularViews)
      ? tabularViews.map((view) => {
          if (!view || typeof view !== "object") return null;
          const { hidden, ...rest } = view;
          return rest;
        }).filter(Boolean)
      : tabularViews;
    const updated = {
      ...user,
      ...(name       !== undefined && { name:       String(name).trim() }),
      ...(email      !== undefined && { email:      String(email).trim() }),
      ...(cellPhone  !== undefined && { cellPhone:  String(cellPhone).trim() }),
      ...(workPhone  !== undefined && { workPhone:  String(workPhone).trim() }),
      ...(homeScreen !== undefined && { homeScreen: homeScreen }),
      ...(tabularViews !== undefined && { tabularViews: normalizedTabularViews }),
      ...(tabularViewsSelectedId !== undefined && { tabularViewsSelectedId: tabularViewsSelectedId }),
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

// ─── CLONE CLIENT ─────────────────────────────────────────────────────────────

// POST /api/admin/clone-client — admin-only; creates a new empty client with:
//   • a new client record (EMPlusClients)
//   • a single admin user whose loginId is  <first segment of caller's loginId>-<newClientId>
//     and whose password hash is copied from the caller (same password, no plaintext needed)
//   • a full copy of the source client's DataDictionary
// The new client starts with no nodes/rels — the new admin is expected to import data.
app.post("/api/admin/clone-client", requireAdmin, async (req, res) => {
  try {
    const srcClientId = req.auth.clientId;
    const { newClientId } = req.body || {};

    if (!newClientId || !String(newClientId).trim()) {
      return res.status(400).json({ error: "newClientId is required" });
    }

    // Normalise to lowercase slug (same chars allowed as clientId throughout the app).
    const cleanId = String(newClientId)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9-]/g, "-")
      .replace(/(^-|-$)/g, "");

    if (!cleanId) {
      return res.status(400).json({ error: "newClientId must contain alphanumeric characters" });
    }

    // Guard: reject if the target client already has users.
    const existingCount = await countUsersByClient(cleanId);
    if (existingCount > 0) {
      return res.status(409).json({ error: `Client "${cleanId}" already exists` });
    }

    // Fetch the current admin's record so we can copy the password hash.
    const srcUser = await getUser(req.auth.loginId);
    if (!srcUser) return res.status(404).json({ error: "Source user not found" });

    // New admin loginId: first "-"-delimited segment of current loginId + "-" + newClientId.
    const baseLoginId = req.auth.loginId.split("-")[0];
    const newLoginId = `${baseLoginId}-${cleanId}`;

    // Guard: loginId must not already exist in the users table.
    const existingLogin = await getUser(newLoginId);
    if (existingLogin) {
      return res.status(409).json({ error: `Login ID "${newLoginId}" is already taken` });
    }

    const now = new Date().toISOString();

    // 1. Create the new client record.
    await putClient({
      clientId: cleanId,
      clonedFrom: srcClientId,
      createdAt: now,
      updatedAt: now,
    });

    // 2. Create the new admin user, reusing the source user's bcrypt hash so the
    //    new admin can log in with the same password without us ever seeing it as
    //    plaintext.
    await putUser({
      loginId: newLoginId,
      clientId: cleanId,
      personId: null,
      role: "admin",
      passwordHash: srcUser.passwordHash,
      createdAt: now,
    });

    // 3. Copy all DataDictionary fields from the source client to the new client.
    const ddFields = await queryDDFields(srcClientId);
    if (ddFields.length > 0) {
      await Promise.all(
        ddFields.map((field) =>
          putDDField({ ...field, clientId: cleanId, updatedAt: now })
        )
      );
    }

    res.status(201).json({
      ok: true,
      newClientId: cleanId,
      newLoginId,
      ddFieldsCloned: ddFields.length,
    });
  } catch (err) {
    console.error("/api/admin/clone-client error", err);
    res.status(500).json({ error: err.message });
  }
});

export default app;
