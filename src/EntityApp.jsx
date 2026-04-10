import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Card, CardContent } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { CalendarDays, Link, Users, Building2, Plus, Pencil, ChevronRight, ChevronDown, Upload } from "lucide-react";
import { Calendar } from "./components/ui/calendar";
import { Input } from "./components/ui/input";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "./components/ui/dialog";
import { Label } from "./components/ui/label";

const initialNodes = [
  { id: "test|entity:A", name: "Entity A", kind: "entity", client: "test" },
  { id: "test|entity:B", name: "Entity B", kind: "entity", client: "test" },
  { id: "test|entity:C", name: "Entity C", kind: "entity", client: "test" },
  { id: "test|person:ray", name: "Ray", kind: "person", client: "test" },
  { id: "test|person:cassie", name: "Cassie", kind: "person", client: "test" },
];

const initialRelationships = [
  { id: "rel-1", type: "owns", from: "test|entity:B", to: "test|entity:A", percent: 20, startDate: "2024-01-01" },
  { id: "rel-2", type: "owns", from: "test|entity:C", to: "test|entity:A", percent: 80, startDate: "2024-01-01" },
  { id: "rel-3", type: "owns", from: "test|person:ray", to: "test|entity:C", percent: 42, startDate: "2024-01-01" },
  { id: "rel-4", type: "employs", from: "test|entity:C", to: "test|person:cassie", role: "Employee", startDate: "2024-06-01" },
];

const getNode = (list, id) => list.find((n) => n.id === id);

const getOwnersOf = (relList, entityId) =>
  relList
    .filter((r) => r.type === "owns" && r.to === entityId)
    .map((r) => ({ nodeId: r.from, rel: r }))
    .sort((a, b) => {
      const aPercent = Number(a.rel?.percent);
      const bPercent = Number(b.rel?.percent);
      const aValue = Number.isFinite(aPercent) ? aPercent : -Infinity;
      const bValue = Number.isFinite(bPercent) ? bPercent : -Infinity;
      return bValue - aValue;
    });

const getOwnedBy = (relList, ownerId) =>
  relList
    .filter((r) => r.type === "owns" && r.from === ownerId)
    .map((r) => ({ nodeId: r.to, rel: r }));

const getEmployeesOf = (relList, entityId) =>
  relList
    .filter((r) => r.type === "employs" && r.from === entityId)
    .map((r) => ({ nodeId: r.to, rel: r }));

const getEmployersOf = (relList, personId) =>
  relList
    .filter((r) => r.type === "employs" && r.to === personId)
    .map((r) => ({ nodeId: r.from, rel: r }));

const buildTree = (rootId, getChildren, path = new Set()) => {
  if (path.has(rootId)) {
    return { id: rootId, children: [], cycle: true };
  }
  const nextPath = new Set(path);
  nextPath.add(rootId);
  const children = getChildren(rootId).map((child) => ({
    rel: child.rel,
    node: buildTree(child.nodeId, getChildren, nextPath),
  }));
  return { id: rootId, children };
};

const formatRel = (rel) => {
  if (!rel) return "";
  if (rel.type === "owns") {
    return (rel.percent && rel.percent !== 100) ? `${rel.percent}%` : "";
  }
  if (rel.type === "employs") {
    return rel.role ? `employs (${rel.role})` : "employs";
  }
  return rel.type;
};

const TreeNode = ({
  tree,
  relLabel,
  rel,
  nodes,
  onRelClick,
  ownershipTotals,
  warnOwnershipTotals,
  collapsedNodes,
  onToggleCollapse,
}) => {
  const node = getNode(nodes, tree.id);
  if (!node) return null;
  const isCollapsed = Boolean(collapsedNodes?.has(tree.id));
  const hasChildren = tree.children.length > 0;
  const totalOwnership = ownershipTotals?.get(tree.id) ?? null;
  const ownershipMismatch =
    warnOwnershipTotals &&
    typeof totalOwnership === "number" &&
    totalOwnership > 0 &&
    Math.abs(totalOwnership - 100) > 0.01;
  return (
    <li className="tree-item">
      <div
        className={`tree-node${ownershipMismatch ? " tree-node-warning" : ""}`}
        style={onRelClick && rel ? { cursor: "pointer" } : undefined}
        onClick={() => {
          if (!onRelClick || !rel) return;
          onRelClick(rel);
        }}
      >
        <div className="tree-node-main">
          {hasChildren && (
            <button
              type="button"
              className="tree-toggle"
              aria-label={isCollapsed ? "Expand" : "Collapse"}
              onClick={(event) => {
                event.stopPropagation();
                onToggleCollapse?.(tree.id);
              }}
            >
              {isCollapsed ? <ChevronRight size={14} /> : <ChevronDown size={14} />}
            </button>
          )}
          <span className="tree-node-icon" aria-hidden="true">
            {node.kind === "entity" ? (
              <Building2 size={16} />
            ) : (
              <Users size={16} />
            )}
          </span>
          <div className="tree-node-text">
            <span className="tree-name">{node.name}</span>
            {relLabel && <span className="tree-rel">{relLabel}</span>}
          </div>
          {tree.cycle && <span className="tree-cycle">(cycle)</span>}
        </div>
      </div>
      {hasChildren && !isCollapsed && (
        <ul className="tree-children">
          {tree.children.map((child) => (
            <TreeNode
              key={`${tree.id}-${child.node.id}`}
              tree={child.node}
              relLabel={formatRel(child.rel)}
              rel={child.rel}
              nodes={nodes}
              onRelClick={onRelClick}
              ownershipTotals={ownershipTotals}
              warnOwnershipTotals={warnOwnershipTotals}
              collapsedNodes={collapsedNodes}
              onToggleCollapse={onToggleCollapse}
            />
          ))}
        </ul>
      )}
    </li>
  );
};

const slugify = (value) =>
  value
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

const toSentenceCase = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/\b\w/g, (match) => match.toUpperCase());

const toIsoDate = (d) => {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

export default function EntityApp() {
  const [nodeList, setNodeList] = useState(initialNodes);
  const [relList, setRelList] = useState(initialRelationships);
  const [viewMode, setViewMode] = useState("hierarchy");
  const [focusId, setFocusId] = useState(() => {
    if (typeof window === "undefined") return "entity:A";
    try {
      return localStorage.getItem("focusId") || "entity:A";
    } catch {
      return "entity:A";
    }
  });
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [showCalendarPopup, setShowCalendarPopup] = useState(false);
  const calendarButtonRef = useRef(null);
  const apiBase = import.meta.env.VITE_API_URL || "http://localhost:5174";
  const [remoteStatus, setRemoteStatus] = useState("idle");
  const [remoteError, setRemoteError] = useState("");
  const [directoryLoaded, setDirectoryLoaded] = useState(false);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [uploadFile, setUploadFile] = useState(null);
  const [uploadKind, setUploadKind] = useState("entity");
  const [uploadType, setUploadType] = useState("entity");
  const [uploadStatus, setUploadStatus] = useState("idle");
  const [uploadError, setUploadError] = useState("");
  const [uploadSummary, setUploadSummary] = useState(null);
  const [uploadDetected, setUploadDetected] = useState("");
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
    const parseCsvLine = (line) => {
      const result = [];
      let current = "";
      let inQuotes = false;
      for (let i = 0; i < line.length; i += 1) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuotes && line[i + 1] === '"') {
            current += '"';
            i += 1;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (ch === "," && !inQuotes) {
          result.push(current);
          current = "";
        } else {
          current += ch;
        }
      }
      result.push(current);
      return result.map((cell) => cell.trim());
    };

    const detectUploadType = (csvText) => {
      const lines = String(csvText || "")
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);
      if (!lines.length) return { type: "entity", label: "Entities" };

      const first = parseCsvLine(lines[0]);
      const headers = first.map((cell) => cell.toLowerCase());
      const headerSet = new Set(headers);
      const hasOwnershipHeaders = ["owner", "owned", "percent", "percentage", "pct", "from", "to"].some((h) =>
        headerSet.has(h)
      );
      if (hasOwnershipHeaders) return { type: "ownership", label: "Ownerships" };

      const sampleLines = lines.slice(0, 5).map(parseCsvLine);
      const maxCols = Math.max(...sampleLines.map((row) => row.length));

      if (maxCols >= 3) {
        const percentIdx = 2;
        const looksNumeric = sampleLines.some((row) =>
          row[percentIdx] && !Number.isNaN(Number(row[percentIdx]))
        );
        if (looksNumeric) return { type: "ownership", label: "Ownerships" };
      }

      const secondCol = sampleLines.map((row) => (row[1] || "").toLowerCase());
      const hasKind = secondCol.some((value) => value === "entity" || value === "person");
      if (hasKind) {
        const defaultKind = secondCol.includes("person") && !secondCol.includes("entity") ? "person" : "entity";
        return { type: defaultKind, label: defaultKind === "person" ? "Persons" : "Entities" };
      }

      return { type: "entity", label: "Entities" };
    };

  const clientId = useMemo(() => {
    if (typeof window === "undefined") return "test";
    const params = new URLSearchParams(window.location.search);
    const fromQuery = params.get("client");
    if (fromQuery) {
      try {
        localStorage.setItem("clientId", fromQuery);
      } catch {
        // ignore storage errors
      }
      return fromQuery;
    }
    try {
      const stored = localStorage.getItem("clientId");
      if (stored) return stored;
    } catch {
      // ignore storage errors
    }
    return "test";
  }, []);

  const [newNode, setNewNode] = useState({
    name: "",
    kind: "entity",
    address: "",
    workPhone: "",
    cellPhone: "",
    emails: [""],
    photo: "",
    taxId: "",
    accountingUrl: "",
    hrUrl: "",
    logo: "",
  });
  const [editNodeId, setEditNodeId] = useState(initialNodes[0]?.id ?? "");
  const [nodeDraft, setNodeDraft] = useState({
    name: "",
    kind: "entity",
    address: "",
    workPhone: "",
    cellPhone: "",
    emails: [""],
    photo: "",
    taxId: "",
    accountingUrl: "",
    hrUrl: "",
    logo: "",
  });

  const [newOwnership, setNewOwnership] = useState({
    from: "",
    to: "",
    percent: "",
    startDate: "",
    endDate: "",
  });
  const [editOwnershipId, setEditOwnershipId] = useState(
    initialRelationships.find((r) => r.type === "owns")?.id ?? ""
  );
  const [ownershipDraft, setOwnershipDraft] = useState({
    from: "",
    to: "",
    percent: "",
    startDate: "",
    endDate: "",
  });

  const [newEmployment, setNewEmployment] = useState({
    from: "",
    to: "",
    role: "",
    startDate: "",
    endDate: "",
  });
  const [editEmploymentId, setEditEmploymentId] = useState(
    initialRelationships.find((r) => r.type === "employs")?.id ?? ""
  );
  const [employmentDraft, setEmploymentDraft] = useState({
    from: "",
    to: "",
    role: "",
    startDate: "",
    endDate: "",
  });

  const [openDialog, setOpenDialog] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [isAddingNode, setIsAddingNode] = useState(false);
  const [isAddingOwnership, setIsAddingOwnership] = useState(false);
  const [isAddingEmployment, setIsAddingEmployment] = useState(false);
  const [collapsedOwnerNodes, setCollapsedOwnerNodes] = useState(() => new Set());
  const [collapsedOwnedNodes, setCollapsedOwnedNodes] = useState(() => new Set());

  const focusNode = useMemo(() => getNode(nodeList, focusId), [nodeList, focusId]);
  const ownerships = useMemo(
    () => relList.filter((r) => r.type === "owns"),
    [relList]
  );
  const employments = useMemo(
    () => relList.filter((r) => r.type === "employs"),
    [relList]
  );
  const entityNodes = useMemo(
    () => nodeList.filter((n) => n.kind === "entity"),
    [nodeList]
  );
  const personNodes = useMemo(
    () => nodeList.filter((n) => n.kind === "person"),
    [nodeList]
  );
  const sortedEntityNodes = useMemo(
    () =>
      [...entityNodes].sort((a, b) =>
        String(a.name || a.id || "").localeCompare(String(b.name || b.id || ""), undefined, {
          sensitivity: "base",
        })
      ),
    [entityNodes]
  );
  const sortedPersonNodes = useMemo(
    () =>
      [...personNodes].sort((a, b) =>
        String(a.name || a.id || "").localeCompare(String(b.name || b.id || ""), undefined, {
          sensitivity: "base",
        })
      ),
    [personNodes]
  );

  const apiRequest = async (path, options = {}) => {
    const response = await fetch(`${apiBase}${path}`, {
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
      ...options,
    });
    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `API error ${response.status}`);
    }
    return response.json();
  };

  const readFileText = (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("Unable to read file"));
      reader.readAsText(file);
    });

  const parseOwnershipCsvClient = (csvText) => {
    const lines = String(csvText || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    if (!lines.length) return { rows: [], skipped: 0 };

    const headerCells = parseCsvLine(lines[0]).map((cell) => cell.toLowerCase());
    const hasHeader = headerCells.some((cell) =>
      ["owner", "owned", "entity", "percent", "percentage", "pct", "from", "to"].includes(cell)
    );
    const startIndex = hasHeader ? 1 : 0;

    const rows = [];
    let skipped = 0;
    for (let i = startIndex; i < lines.length; i += 1) {
      const cols = parseCsvLine(lines[i]);
      const owner = String(cols[0] || "").trim();
      const owned = String(cols[1] || "").trim();
      const percentText = String(cols[2] ?? "").trim();
      if (!owner || !owned) {
        skipped += 1;
        continue;
      }
      const percent = percentText === "" ? Number.NaN : Number(percentText);
      rows.push({ owner, owned, percent });
    }
    return { rows, skipped };
  };

  const rowsToCsv = (rows) =>
    rows
      .map((row) =>
        [row.owner, row.owned, row.percent]
          .map((value) => `"${String(value).replace(/"/g, '""')}"`)
          .join(",")
      )
      .join("\n");

  const loadDirectory = useCallback(async ({ isActive } = {}) => {
    const canUpdate = () => (typeof isActive === "function" ? isActive() : true);
    try {
      if (!canUpdate()) return;
      setRemoteStatus("loading");
      setRemoteError("");
      const response = await fetch(`${apiBase}/api/directory`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ client: clientId, debug: false }),
      });
      if (!response.ok) {
        throw new Error(`API error ${response.status}`);
      }
      const data = await response.json();
      if (!canUpdate()) return;
      if (Array.isArray(data.nodes) && data.nodes.length > 0) {
        setNodeList(data.nodes);
        if (!data.nodes.some((n) => n.id === focusId)) {
          setFocusId(data.nodes[0].id);
        }
      }
      if (Array.isArray(data.rels)) {
        setRelList(data.rels);
      }
      if (canUpdate()) setRemoteStatus("connected");
    } catch (err) {
      if (!canUpdate()) return;
      setRemoteStatus("error");
      setRemoteError(err.message || "Unable to load directory");
    } finally {
      if (canUpdate()) setDirectoryLoaded(true);
    }
  }, [apiBase, clientId, focusId]);

  const handleUploadCsv = async () => {
    if (!uploadFile) {
      setUploadError("Please select a CSV file.");
      setUploadStatus("error");
      return;
    }
    try {
      setUploadStatus("uploading");
      setUploadError("");
      setUploadSummary(null);
      setUploadProgress({ current: 0, total: 0 });

      if (uploadType === "ownership") {
        const csvText = await readFileText(uploadFile);
        const { rows, skipped } = parseOwnershipCsvClient(csvText);
        if (!rows.length) {
          setUploadStatus("error");
          setUploadError("No valid ownership rows found in the CSV.");
          return;
        }

        const chunkSize = 10;
        const totalChunks = Math.ceil(rows.length / chunkSize);
        setUploadProgress({ current: 0, total: totalChunks });

        let imported = 0;
        let rejected = 0;
        const allErrors = [];

        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        for (let i = 0; i < totalChunks; i += 1) {
          const chunk = rows.slice(i * chunkSize, (i + 1) * chunkSize);
          const chunkCsv = rowsToCsv(chunk);
          const response = await fetch(`${apiBase}/api/import/ownerships-csv`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ csv: chunkCsv, client: clientId }),
          });
          const responseText = await response.text();
          let data = {};
          if (responseText) {
            try {
              data = JSON.parse(responseText);
            } catch {
              data = { raw: responseText };
            }
          }
          if (!response.ok) {
            const fallback = typeof data?.raw === "string" ? data.raw : responseText;
            throw new Error(data?.error || fallback || `Upload failed (${response.status})`);
          }

          imported += Number(data.imported || 0);
          rejected += Number(data.skipped || 0);
          if (Array.isArray(data.errors)) allErrors.push(...data.errors);

          setUploadProgress({ current: i + 1, total: totalChunks });
          await sleep(250);
        }

        setUploadSummary({
          total: rows.length,
          imported,
          skipped: skipped + rejected,
          errors: allErrors,
        });
        setUploadStatus("success");
        await loadDirectory();
        return;
      }

      const form = new FormData();
      form.append("file", uploadFile);
      form.append("client", clientId);
      if (uploadType !== "ownership") {
        form.append("defaultKind", uploadKind);
      }

      const endpoint = uploadType === "ownership"
        ? "/api/import/ownerships-csv/upload"
        : "/api/import/nodes-csv/upload";
      const response = await fetch(`${apiBase}${endpoint}`, {
        method: "POST",
        body: form,
      });
      const responseText = await response.text();
      let data = {};
      if (responseText) {
        try {
          data = JSON.parse(responseText);
        } catch {
          data = { raw: responseText };
        }
      }
      if (!response.ok) {
        const fallback = typeof data?.raw === "string" ? data.raw : responseText;
        throw new Error(data?.error || fallback || `Upload failed (${response.status})`);
      }
      setUploadSummary(data);
      setUploadStatus("success");
      await loadDirectory();
    } catch (err) {
      setUploadStatus("error");
      setUploadError(err.message || "Upload failed");
    }
  };

  const downloadErrorsCsv = () => {
    const errors = uploadSummary?.errors || [];
    if (!errors.length) return;
    const header = ["row", "reason", "owner", "owned", "percent"].join(",");
    const lines = errors.map((err) => [
      err.row ?? "",
      err.reason ?? "",
      err.owner ?? "",
      err.owned ?? "",
      err.percent ?? "",
    ].map((value) => `"${String(value).replace(/"/g, '""')}"`).join(","));
    const csv = [header, ...lines].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "ownership-import-errors.csv";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const makeNodeId = (kind, name, currentId = "") => {
    const slug = slugify(name || "node");
    const rawId = `${kind}:${slug || "node"}`;
    const next = normalizeClientId(clientId, rawId);
    return currentId && currentId === next ? currentId : next;
  };

  const makeRelId = () => `rel-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const getNodeName = (id) => getNode(nodeList, id)?.name || id;

  useEffect(() => {
    const node = getNode(nodeList, editNodeId);
    if (node) {
      const emails = Array.isArray(node.emails)
        ? node.emails
        : node.email
          ? [node.email]
          : [""];
      setNodeDraft({
        name: node.name,
        kind: node.kind,
        address: node.address || "",
        workPhone: node.workPhone || "",
        cellPhone: node.cellPhone || "",
        emails: emails.length > 0 ? emails : [""],
        photo: node.photo || "",
        taxId: node.taxId || "",
        accountingUrl: node.accountingUrl || "",
        hrUrl: node.hrUrl || "",
        logo: node.logo || "",
      });
    }
  }, [editNodeId, nodeList]);

  useEffect(() => {
    const rel = ownerships.find((r) => r.id === editOwnershipId);
    if (rel) {
      setOwnershipDraft({
        from: rel.from,
        to: rel.to,
        percent: rel.percent ?? "",
        startDate: rel.startDate ?? "",
        endDate: rel.endDate ?? "",
      });
    }
  }, [editOwnershipId, ownerships]);

  useEffect(() => {
    const rel = employments.find((r) => r.id === editEmploymentId);
    if (rel) {
      setEmploymentDraft({
        from: rel.from,
        to: rel.to,
        role: rel.role ?? "",
        startDate: rel.startDate ?? "",
        endDate: rel.endDate ?? "",
      });
    }
  }, [editEmploymentId, employments]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    try {
      localStorage.setItem("focusId", focusId);
    } catch {
      // ignore storage errors
    }
  }, [focusId]);

  useEffect(() => {
    if (!clientId) return;
    setFocusId((prev) => normalizeClientId(clientId, prev));
  }, [clientId]);

  useEffect(() => {
    let active = true;
    if (apiBase) {
      loadDirectory({ isActive: () => active });
    }
    return () => {
      active = false;
    };
  }, [apiBase, clientId, focusId, loadDirectory]);

  const ownerTree = useMemo(
    () => buildTree(focusId, (id) => getOwnersOf(relList, id)),
    [focusId, relList]
  );
  const ownedTree = useMemo(
    () => buildTree(focusId, (id) => getOwnedBy(relList, id)),
    [focusId, relList]
  );

  const focusEmployees = useMemo(
    () => getEmployeesOf(relList, focusId),
    [relList, focusId]
  );
  const focusEmployers = useMemo(
    () => getEmployersOf(relList, focusId),
    [relList, focusId]
  );

  const ownershipTotalsByEntity = useMemo(() => {
    const totals = new Map();
    relList
      .filter((rel) => rel.type === "owns")
      .forEach((rel) => {
        const value = Number(rel.percent);
        if (!Number.isFinite(value)) return;
        totals.set(rel.to, (totals.get(rel.to) ?? 0) + value);
      });
    return totals;
  }, [relList]);

  const toggleOwnerCollapse = (nodeId) => {
    setCollapsedOwnerNodes((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      return next;
    });
  };

  const toggleOwnedCollapse = (nodeId) => {
    setCollapsedOwnedNodes((prev) => {
      const next = new Set(prev);
      if (next.has(nodeId)) {
        next.delete(nodeId);
      } else {
        next.add(nodeId);
      }
      return next;
    });
  };

  const directOwners = useMemo(
    () => getOwnersOf(relList, focusId),
    [relList, focusId]
  );
  const directOwnerTotal = useMemo(
    () =>
      directOwners.reduce((sum, item) => {
        const value = Number(item.rel?.percent);
        return Number.isFinite(value) ? sum + value : sum;
      }, 0),
    [directOwners]
  );
  const directOwnerTotalOk = Math.abs(directOwnerTotal - 100) < 0.01;

  const niceDate = selectedDate.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
  });

  return (
    <div style={{ maxWidth: "90%", margin: "0 auto", padding: "24px 16px 120px" }} data-lpignore="true">
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 600 }}>{toSentenceCase(clientId)}</div>
          <div style={{ fontSize: 22, fontWeight: 600 }}>Entity Dashboard</div>
          <div className="header-date">{niceDate}</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div className="focus-panel">
            <span className="focus-label">Focus</span>
            <select
              className="focus-select"
              value={focusId}
              onChange={(e) => setFocusId(e.target.value)}
              data-lpignore="true"
            >
              {nodeList.map((node) => (
                <option key={node.id} value={node.id}>
                  {node.name}
                </option>
              ))}
            </select>
          </div>
          <Button
            type="button"
            variant="outline"
            className="btn-icon"
            aria-label="Select date"
            ref={calendarButtonRef}
            onClick={() => setShowCalendarPopup((prev) => !prev)}
          >
            <CalendarDays />
          </Button>
          <Button
            type="button"
            variant="outline"
            className="btn-icon"
            aria-label="Upload CSV"
            title="Import CSV"
            onClick={() => {
              setUploadOpen(true);
              setUploadStatus("idle");
              setUploadError("");
              setUploadSummary(null);
              setUploadFile(null);
              setUploadDetected("");
            }}
          >
            <Upload />
          </Button>
        </div>
      </div>

      {showCalendarPopup && (
        <>
          <div
            className="calendar-popup-overlay"
            onClick={() => setShowCalendarPopup(false)}
          />
          <div className="calendar-popup">
            <Calendar
              mode="single"
              selected={selectedDate}
              onSelect={(d) => {
                if (d) {
                  setSelectedDate(d);
                  setShowCalendarPopup(false);
                }
              }}
              disabled={(date) => date > new Date()}
              initialFocus
            />
          </div>
        </>
      )}

      <Dialog open={uploadOpen} onOpenChange={() => setUploadOpen(false)}>
        <DialogContent style={{ maxWidth: 520 }}>
          <DialogHeader>
            <DialogTitle>Import CSV</DialogTitle>
          </DialogHeader>
          <div style={{ display: "grid", gap: 12 }}>
            <div style={{ fontSize: 14, color: "#6b7280" }}>
              {uploadDetected ? `Detected: ${uploadDetected}` : "Choose a CSV file to detect its type."}
            </div>
            <div>
              <Label htmlFor="csv-upload">CSV file</Label>
              <Input
                id="csv-upload"
                type="file"
                accept=".csv,text/csv"
                onChange={(event) => {
                  const file = event.target.files?.[0] || null;
                  setUploadFile(file);
                  setUploadSummary(null);
                  setUploadError("");
                  setUploadStatus("idle");
                  if (!file) {
                    setUploadDetected("");
                    return;
                  }
                  const reader = new FileReader();
                  reader.onload = () => {
                    const text = reader.result || "";
                    const detected = detectUploadType(text);
                    setUploadType(detected.type === "person" ? "person" : detected.type === "ownership" ? "ownership" : "entity");
                    setUploadDetected(detected.label);
                    if (detected.type === "person") setUploadKind("person");
                    if (detected.type === "entity") setUploadKind("entity");
                  };
                  reader.readAsText(file);
                }}
              />
            </div>
            {uploadType !== "ownership" && uploadDetected && (
              <div>
                <Label htmlFor="csv-default-kind">Default kind</Label>
                <select
                  id="csv-default-kind"
                  className="focus-select"
                  value={uploadKind}
                  onChange={(event) => setUploadKind(event.target.value)}
                >
                  <option value="entity">Entity</option>
                  <option value="person">Person</option>
                </select>
              </div>
            )}
            {uploadStatus === "uploading" && (
              <div style={{ color: "#6b7280", fontSize: 14 }}>
                {uploadType === "ownership" && uploadProgress.total > 0
                  ? `Uploading chunk ${uploadProgress.current} of ${uploadProgress.total}...`
                  : "Uploading..."}
              </div>
            )}
            {uploadStatus === "error" && uploadError && (
              <div style={{ color: "#dc2626", fontSize: 14 }}>{uploadError}</div>
            )}
            {uploadStatus === "success" && uploadSummary && (
              <div style={{ color: "#16a34a", fontSize: 14 }}>
                {uploadType === "ownership" ? (
                  <>Imported {uploadSummary.imported || 0} ownerships. Skipped {uploadSummary.skipped || 0}.</>
                ) : (
                  <>Imported {uploadSummary.total || 0} rows ({uploadSummary.entities || 0} entities,
                    {" "}{uploadSummary.persons || 0} persons). Skipped {uploadSummary.skipped || 0}.</>
                )}
              </div>
            )}
            {uploadSummary?.errors?.length > 0 && (
              <div style={{ display: "grid", gap: 6 }}>
                <div style={{ fontSize: 13, color: "#b45309" }}>
                  {uploadSummary.errors.length} rows rejected.
                </div>
                <div style={{ maxHeight: 140, overflowY: "auto", border: "1px solid #e5e7eb", borderRadius: 8, padding: 8 }}>
                  {uploadSummary.errors.slice(0, 50).map((err, idx) => (
                    <div key={`${err.row}-${idx}`} style={{ fontSize: 12, color: "#6b7280" }}>
                      Row {err.row}: {err.reason}
                      {err.owner ? ` (owner: ${err.owner})` : ""}
                      {err.owned ? ` (owned: ${err.owned})` : ""}
                      {err.percent !== undefined ? ` (percent: ${err.percent})` : ""}
                    </div>
                  ))}
                  {uploadSummary.errors.length > 50 && (
                    <div style={{ fontSize: 12, color: "#6b7280" }}>…showing first 50</div>
                  )}
                </div>
                <Button type="button" variant="outline" onClick={downloadErrorsCsv}>
                  Download rejected rows
                </Button>
              </div>
            )}
          </div>
          <DialogFooter>
            <Button type="button" variant="outline" onClick={() => setUploadOpen(false)}>
              Close
            </Button>
            <Button
              type="button"
              onClick={handleUploadCsv}
              disabled={uploadStatus === "uploading"}
            >
              Upload
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <div style={{ display: "flex", justifyContent: "center", gap: 8, margin: "18px 0" }}>
        <Button
          type="button"
          variant={viewMode === "hierarchy" ? "default" : "outline"}
          onClick={() => setViewMode("hierarchy")}
        >
          Hierarchy
        </Button>
        <Button
          type="button"
          variant={viewMode === "directory" ? "default" : "outline"}
          onClick={() => setViewMode("directory")}
        >
          Directory
        </Button>
      </div>

      {remoteStatus === "loading" && (
        <div style={{ textAlign: "center", color: "#6b7280", marginBottom: 12 }}>
          Loading directory...
        </div>
      )}
      {remoteStatus === "error" && (
        <div style={{ textAlign: "center", color: "#dc2626", marginBottom: 12 }}>
          {remoteError || "Unable to load directory"}
        </div>
      )}

      {viewMode === "hierarchy" && (
        <div className="hierarchy-grid">
          {focusNode?.kind !== "person" && (
            <Card>
              <CardContent>
                <div className="section-title">Owners of {focusNode?.name || focusId}</div>
                <ul className="tree-root">
                  <TreeNode
                    tree={ownerTree}
                    relLabel={null}
                    nodes={nodeList}
                    onRelClick={(rel) => {
                      setEditOwnershipId(rel.id);
                      setOpenDialog({ type: "edit-ownership" });
                    }}
                    ownershipTotals={ownershipTotalsByEntity}
                    warnOwnershipTotals
                    collapsedNodes={collapsedOwnerNodes}
                    onToggleCollapse={toggleOwnerCollapse}
                  />
                </ul>
              </CardContent>
            </Card>
          )}
          <Card>
            <CardContent>
              <div className="section-title">Owned by {focusNode?.name || focusId}</div>
              <ul className="tree-root">
                <TreeNode
                  tree={ownedTree}
                  relLabel={null}
                  nodes={nodeList}
                  onRelClick={(rel) => {
                    setEditOwnershipId(rel.id);
                    setOpenDialog({ type: "edit-ownership" });
                  }}
                  collapsedNodes={collapsedOwnedNodes}
                  onToggleCollapse={toggleOwnedCollapse}
                />
              </ul>
            </CardContent>
          </Card>
          <Card>
            <CardContent>
              <div className="section-title">Relationships</div>
              {focusNode?.kind === "entity" ? (
                <ul className="relationship-list">
                  {focusEmployees.length === 0 ? (
                    <li className="empty-text">No employees listed</li>
                  ) : (
                    focusEmployees.map((item) => (
                      <li
                        key={`${item.rel.id}-${item.nodeId}`}
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditEmploymentId(item.rel.id);
                          setOpenDialog({ type: "edit-employment" });
                        }}
                      >
                        <span className="relationship-name">{getNodeName(item.nodeId)}</span>
                        <span className="relationship-meta">{formatRel(item.rel)}</span>
                      </li>
                    ))
                  )}
                </ul>
              ) : (
                <ul className="relationship-list">
                  {focusEmployers.length === 0 ? (
                    <li className="empty-text">No employers listed</li>
                  ) : (
                    focusEmployers.map((item) => (
                      <li
                        key={`${item.rel.id}-${item.nodeId}`}
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditEmploymentId(item.rel.id);
                          setOpenDialog({ type: "edit-employment" });
                        }}
                      >
                        <span className="relationship-name">{getNodeName(item.nodeId)}</span>
                        <span className="relationship-meta">{formatRel(item.rel)}</span>
                      </li>
                    ))
                  )}
                </ul>
              )}
            </CardContent>
          </Card>
        </div>
      )}

      {viewMode === "directory" && (
        <div className="directory-grid">
          <Card>
            <CardContent>
              <div className="section-title">Entities</div>
              <div className="directory-scroll">
                <ul className="directory-list">
                  {sortedEntityNodes.map((n) => (
                    <li key={n.id}>
                      <div
                        className="directory-item"
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditNodeId(n.id);
                          setOpenDialog({ type: "edit-node" });
                        }}
                      >
                        <Building2 className="directory-icon" />
                        <div>
                          <div className="directory-name">{n.name}</div>
                          <div className="directory-meta">{n.id}</div>
                        </div>
                      </div>
                    </li>
                  ))}
                </ul>
              </div>
            </CardContent>
          </Card>
          <Card>
            <CardContent>
              <div className="section-title">People</div>
              <div className="directory-scroll">
                <ul className="directory-list">
                  {sortedPersonNodes.map((n) => (
                    <li key={n.id}>
                      <div
                        className="directory-item"
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditNodeId(n.id);
                          setOpenDialog({ type: "edit-node" });
                        }}
                      >
                        <Users className="directory-icon" />
                        <div>
                          <div className="directory-name">{n.name}</div>
                          <div className="directory-meta">{n.id}</div>
                        </div>
                      </div>
                    </li>
                  ))}
                </ul>
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      {viewMode === "editor" && (
        <div className="editor-section">
          <Card>
            <CardContent>
              <div className="section-title">Nodes</div>
              <div className="editor-help">Use the Add button below to create new nodes.</div>
              <div className="editor-actions">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setOpenDialog({ type: "edit-node" })}
                >
                  Edit a node
                </Button>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardContent>
              <div className="section-title">Ownerships</div>
              <div className="editor-help">Use the Add button below to create new ownerships.</div>
              <div className="editor-actions">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setOpenDialog({ type: "edit-ownership" })}
                >
                  Edit an ownership
                </Button>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardContent>
              <div className="section-title">Employment</div>
              <div className="editor-help">Use the Add button below to create new employment records.</div>
              <div className="editor-actions">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setOpenDialog({ type: "edit-employment" })}
                >
                  Edit employment
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      <div className="fab-container">
        <Button
          type="button"
          className="fab-add"
          onClick={() => setOpenDialog({ type: "add-picker" })}
        >
          <Plus size={16} />
          <span>Add</span>
        </Button>
      </div>

      <Dialog open={Boolean(openDialog)} onOpenChange={() => setOpenDialog(null)}>
        {openDialog?.type === "add-picker" && (
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Add new record</DialogTitle>
            </DialogHeader>
            <div className="dialog-choice-grid">
              <Button type="button" variant="outline" onClick={() => setOpenDialog({ type: "add-node" })}>
                Node
              </Button>
              <Button type="button" variant="outline" onClick={() => setOpenDialog({ type: "add-ownership" })}>
                Ownership
              </Button>
              <Button type="button" variant="outline" onClick={() => setOpenDialog({ type: "add-employment" })}>
                Employment
              </Button>
            </div>
            <DialogFooter>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "add-node" && (
          <DialogContent style={{ minWidth: 1000, maxWidth: "none" }}>
            <DialogHeader>
              <DialogTitle>Add node</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Name</label>
                <input
                  className="form-input"
                  type="text"
                  placeholder="Name"
                  value={newNode.name}
                  onChange={(e) =>
                    setNewNode((prev) => ({ ...prev, name: e.target.value }))
                  }
                  autoComplete="off"
                  data-lpignore="true"
                />
              </div>
              <div className="form-row">
                <label className="form-label">Kind</label>
                <select
                  className="form-select"
                  value={newNode.kind}
                  onChange={(e) =>
                    setNewNode((prev) => ({
                      ...prev,
                      kind: e.target.value,
                      emails: prev.emails?.length ? prev.emails : [""],
                    }))
                  }
                  data-lpignore="true"
                >
                  <option value="entity">Entity</option>
                  <option value="person">Person</option>
                </select>
              </div>
              {newNode.kind === "person" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={newNode.address}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, address: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Work phone</label>
                    <input
                      className="form-input"
                      type="tel"
                      value={newNode.workPhone}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, workPhone: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Cell phone</label>
                    <input
                      className="form-input"
                      type="tel"
                      value={newNode.cellPhone}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, cellPhone: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Email addresses</label>
                    <div style={{ display: "grid", gap: 8 }}>
                      {(newNode.emails ?? [""]).map((email, idx) => (
                        <div key={idx} style={{ display: "flex", gap: 8 }}>
                          <input
                            className="form-input"
                            type="email"
                            value={email}
                            style={{ flex: 1, minWidth: 0 }}
                            onChange={(e) => {
                              const next = [...(newNode.emails ?? [""])];
                              next[idx] = e.target.value;
                              setNewNode((prev) => ({ ...prev, emails: next }));
                            }}
                            autoComplete="off"
                            data-lpignore="true"
                          />
                          {(newNode.emails ?? []).length > 1 && (
                            <Button
                              type="button"
                              variant="outline"
                              onClick={() => {
                                const next = [...(newNode.emails ?? [""])];
                                next.splice(idx, 1);
                                setNewNode((prev) => ({ ...prev, emails: next }));
                              }}
                            >
                              Remove
                            </Button>
                          )}
                        </div>
                      ))}
                      <div>
                        <Button
                          type="button"
                          variant="outline"
                          onClick={() => {
                            const next = [...(newNode.emails ?? [""]), ""];
                            setNewNode((prev) => ({ ...prev, emails: next }));
                          }}
                        >
                          Add email
                        </Button>
                      </div>
                    </div>
                  </div>
                  <div className="form-row">
                    <label className="form-label">Photo URL</label>
                    <input
                      className="form-input"
                      type="url"
                      value={newNode.photo}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, photo: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                </>
              )}
              {newNode.kind === "entity" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={newNode.address}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, address: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Tax ID</label>
                    <input
                      className="form-input"
                      type="text"
                      value={newNode.taxId}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, taxId: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Accounting system URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={newNode.accountingUrl}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, accountingUrl: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">HR system URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={newNode.hrUrl}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, hrUrl: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Logo image URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={newNode.logo}
                      onChange={(e) =>
                        setNewNode((prev) => ({ ...prev, logo: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                </>
              )}
            </div>
            <DialogFooter>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                disabled={isAddingNode}
                onClick={async () => {
                  if (isAddingNode) return;
                  if (!newNode.name.trim()) return;
                  setIsAddingNode(true);
                  const id = makeNodeId(newNode.kind, newNode.name);
                  const emails = (newNode.emails ?? [])
                    .map((e) => e.trim())
                    .filter(Boolean);
                  const payload = {
                    id,
                    name: newNode.name.trim(),
                    kind: newNode.kind,
                    client: clientId,
                    address: newNode.address?.trim() || "",
                    workPhone: newNode.workPhone?.trim() || "",
                    cellPhone: newNode.cellPhone?.trim() || "",
                    emails,
                    photo: newNode.photo?.trim() || "",
                    taxId: newNode.taxId?.trim() || "",
                    accountingUrl: newNode.accountingUrl?.trim() || "",
                    hrUrl: newNode.hrUrl?.trim() || "",
                    logo: newNode.logo?.trim() || "",
                  };
                  try {
                    await apiRequest("/api/nodes", {
                      method: "POST",
                      body: JSON.stringify(payload),
                    });
                    setNodeList((prev) => [...prev, payload]);
                    if (!focusId) setFocusId(id);
                    setNewNode({
                      name: "",
                      kind: newNode.kind,
                      address: "",
                      workPhone: "",
                      cellPhone: "",
                      emails: [""],
                      photo: "",
                      taxId: "",
                      accountingUrl: "",
                      hrUrl: "",
                      logo: "",
                    });
                    setRemoteStatus("connected");
                    setOpenDialog(null);
                  } catch (err) {
                    setRemoteStatus("error");
                    setRemoteError(err.message);
                  } finally {
                    setIsAddingNode(false);
                  }
                }}
              >
                Add
              </Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "add-ownership" && (
          <DialogContent style={{ minWidth: 1000, maxWidth: "none" }}>
            <DialogHeader>
              <DialogTitle>Add ownership</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Owner</label>
                <select
                  className="form-select"
                  value={newOwnership.from}
                  onChange={(e) =>
                    setNewOwnership((prev) => ({
                      ...prev,
                      from: e.target.value,
                    }))
                  }
                >
                  <option value="">Owner</option>
                  {nodeList.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Owned entity</label>
                <select
                  className="form-select"
                  value={newOwnership.to}
                  onChange={(e) =>
                    setNewOwnership((prev) => ({
                      ...prev,
                      to: e.target.value,
                    }))
                  }
                >
                  <option value="">Owned entity</option>
                  {entityNodes.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Percent</label>
                <input
                  className="form-input"
                  type="number"
                  min="0"
                  max="100"
                  placeholder="Percent"
                  value={newOwnership.percent}
                  onChange={(e) =>
                    setNewOwnership((prev) => ({
                      ...prev,
                      percent: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">Start date</label>
                <input
                  className="form-input"
                  type="date"
                  value={newOwnership.startDate}
                  onChange={(e) =>
                    setNewOwnership((prev) => ({
                      ...prev,
                      startDate: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">End date</label>
                <input
                  className="form-input"
                  type="date"
                  value={newOwnership.endDate}
                  onChange={(e) =>
                    setNewOwnership((prev) => ({
                      ...prev,
                      endDate: e.target.value,
                    }))
                  }
                />
              </div>
            </div>
            <DialogFooter>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                disabled={isAddingOwnership}
                onClick={async () => {
                  if (isAddingOwnership) return;
                  if (!newOwnership.from || !newOwnership.to) return;
                  setIsAddingOwnership(true);
                  const payload = {
                    from: newOwnership.from,
                    to: newOwnership.to,
                    percent: newOwnership.percent
                      ? Number(newOwnership.percent)
                      : null,
                    startDate: newOwnership.startDate || null,
                    endDate: newOwnership.endDate || null,
                    client: clientId,
                  };
                  try {
                    await apiRequest("/api/relationships/owns", {
                      method: "POST",
                      body: JSON.stringify(payload),
                    });
                    setRelList((prev) => [
                      ...prev,
                      {
                        id: makeRelId(),
                        type: "owns",
                        ...payload,
                      },
                    ]);
                    setNewOwnership({
                      from: "",
                      to: "",
                      percent: "",
                      startDate: "",
                      endDate: "",
                    });
                    setRemoteStatus("connected");
                    setOpenDialog(null);
                  } catch (err) {
                    setRemoteStatus("error");
                    setRemoteError(err.message);
                  } finally {
                    setIsAddingOwnership(false);
                  }
                }}
              >
                Add
              </Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "add-employment" && (
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Add employment</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Employer</label>
                <select
                  className="form-select"
                  value={newEmployment.from}
                  onChange={(e) =>
                    setNewEmployment((prev) => ({
                      ...prev,
                      from: e.target.value,
                    }))
                  }
                >
                  <option value="">Employer</option>
                  {entityNodes.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Person</label>
                <select
                  className="form-select"
                  value={newEmployment.to}
                  onChange={(e) =>
                    setNewEmployment((prev) => ({
                      ...prev,
                      to: e.target.value,
                    }))
                  }
                >
                  <option value="">Person</option>
                  {personNodes.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Role</label>
                <input
                  className="form-input"
                  type="text"
                  placeholder="Role"
                  value={newEmployment.role}
                  onChange={(e) =>
                    setNewEmployment((prev) => ({
                      ...prev,
                      role: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">Start date</label>
                <input
                  className="form-input"
                  type="date"
                  value={newEmployment.startDate}
                  onChange={(e) =>
                    setNewEmployment((prev) => ({
                      ...prev,
                      startDate: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">End date</label>
                <input
                  className="form-input"
                  type="date"
                  value={newEmployment.endDate}
                  onChange={(e) =>
                    setNewEmployment((prev) => ({
                      ...prev,
                      endDate: e.target.value,
                    }))
                  }
                />
              </div>
            </div>
            <DialogFooter>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                disabled={isAddingEmployment}
                onClick={async () => {
                  if (isAddingEmployment) return;
                  if (!newEmployment.from || !newEmployment.to) return;
                  setIsAddingEmployment(true);
                  const payload = {
                    from: newEmployment.from,
                    to: newEmployment.to,
                    role: newEmployment.role || null,
                    startDate: newEmployment.startDate || null,
                    endDate: newEmployment.endDate || null,
                    client: clientId,
                  };
                  try {
                    await apiRequest("/api/relationships/employs", {
                      method: "POST",
                      body: JSON.stringify(payload),
                    });
                    setRelList((prev) => [
                      ...prev,
                      {
                        id: makeRelId(),
                        type: "employs",
                        ...payload,
                      },
                    ]);
                    setNewEmployment({
                      from: "",
                      to: "",
                      role: "",
                      startDate: "",
                      endDate: "",
                    });
                    setRemoteStatus("connected");
                    setOpenDialog(null);
                  } catch (err) {
                    setRemoteStatus("error");
                    setRemoteError(err.message);
                  } finally {
                    setIsAddingEmployment(false);
                  }
                }}
              >
                Add
              </Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "edit-node" && (
          <DialogContent style={{ minWidth: 1000, maxWidth: "none" }}>
            <DialogHeader>
              <DialogTitle>Edit node</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Name</label>
                <input
                  className="form-input"
                  type="text"
                  value={nodeDraft.name}
                  onChange={(e) =>
                    setNodeDraft((prev) => ({ ...prev, name: e.target.value }))
                  }
                  autoComplete="off"
                  data-lpignore="true"
                />
              </div>
              <div className="form-row">
                <label className="form-label">Kind</label>
                <select
                  className="form-select"
                  value={nodeDraft.kind}
                  onChange={(e) =>
                    setNodeDraft((prev) => ({
                      ...prev,
                      kind: e.target.value,
                      emails: prev.emails?.length ? prev.emails : [""],
                    }))
                  }
                  data-lpignore="true"
                >
                  <option value="entity">Entity</option>
                  <option value="person">Person</option>
                </select>
              </div>
              {nodeDraft.kind === "person" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={nodeDraft.address}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, address: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Work phone</label>
                    <input
                      className="form-input"
                      type="tel"
                      value={nodeDraft.workPhone}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, workPhone: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Cell phone</label>
                    <input
                      className="form-input"
                      type="tel"
                      value={nodeDraft.cellPhone}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, cellPhone: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Email addresses</label>
                    <div style={{ display: "grid", gap: 8 }}>
                      {(nodeDraft.emails ?? [""]).map((email, idx) => (
                        <div key={idx} style={{ display: "flex", gap: 8 }}>
                          <input
                            className="form-input"
                            type="email"
                            value={email}
                            style={{ flex: 1, minWidth: 0 }}
                            onChange={(e) => {
                              const next = [...(nodeDraft.emails ?? [""])];
                              next[idx] = e.target.value;
                              setNodeDraft((prev) => ({ ...prev, emails: next }));
                            }}
                            autoComplete="off"
                            data-lpignore="true"
                          />
                          {(nodeDraft.emails ?? []).length > 1 && (
                            <Button
                              type="button"
                              variant="outline"
                              onClick={() => {
                                const next = [...(nodeDraft.emails ?? [""])];
                                next.splice(idx, 1);
                                setNodeDraft((prev) => ({ ...prev, emails: next }));
                              }}
                            >
                              Remove
                            </Button>
                          )}
                        </div>
                      ))}
                      <div>
                        <Button
                          type="button"
                          variant="outline"
                          onClick={() => {
                            const next = [...(nodeDraft.emails ?? [""]), ""];
                            setNodeDraft((prev) => ({ ...prev, emails: next }));
                          }}
                        >
                          Add email
                        </Button>
                      </div>
                    </div>
                  </div>
                  <div className="form-row">
                    <label className="form-label">Photo URL</label>
                    <input
                      className="form-input"
                      type="url"
                      value={nodeDraft.photo}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, photo: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                </>
              )}
              {nodeDraft.kind === "entity" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={nodeDraft.address}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, address: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Tax ID</label>
                    <input
                      className="form-input"
                      type="text"
                      value={nodeDraft.taxId}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, taxId: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Accounting system URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={nodeDraft.accountingUrl}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, accountingUrl: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">HR system URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={nodeDraft.hrUrl}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, hrUrl: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Logo image URL</label>
                    <input
                      className="form-input"
                      type="url"
                      placeholder="https://"
                      value={nodeDraft.logo}
                      onChange={(e) =>
                        setNodeDraft((prev) => ({ ...prev, logo: e.target.value }))
                      }
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                </>
              )}
            </div>
            <DialogFooter
              style={{
                justifyContent: "space-between",
                width: "100%",
                paddingLeft: 24,
                paddingRight: 24,
                marginLeft: 0,
                marginRight: 0,
              }}
            >
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  if (!editNodeId) return;
                  setFocusId(editNodeId);
                  setViewMode("hierarchy");
                  setOpenDialog(null);
                }}
              >
                Focus
              </Button>
              <div style={{ display: "flex", paddingRight: 48, justifyContent: "flex-end", gap: 8 }}>
                <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => {
                    if (!editNodeId) return;
                    const targetId = editNodeId;
                    setConfirmDialog({
                      title: "ERPlus says",
                      message: "Delete this node? This will also remove related relationships.",
                      onConfirm: () => {
                        apiRequest(`/api/nodes/${targetId}`, {
                          method: "DELETE",
                          body: JSON.stringify({ client: clientId }),
                        })
                          .then(() => {
                            setNodeList((prev) => prev.filter((n) => n.id !== targetId));
                            setRelList((prev) => prev.filter((r) => r.from !== targetId && r.to !== targetId));
                            if (focusId === targetId) {
                              const next = nodeList.find((n) => n.id !== targetId);
                              if (next) setFocusId(next.id);
                            }
                            setRemoteStatus("connected");
                            setOpenDialog(null);
                            setConfirmDialog(null);
                          })
                          .catch((err) => {
                            setRemoteStatus("error");
                            setRemoteError(err.message);
                          });
                      },
                    });
                  }}
                >
                  Delete
                </Button>
                <Button
                  type="button"
                  onClick={() => {
                    if (!editNodeId || !nodeDraft.name.trim()) return;
                    const newId = makeNodeId(
                      nodeDraft.kind,
                      nodeDraft.name,
                      editNodeId
                    );
                    const emails = (nodeDraft.emails ?? [])
                      .map((e) => e.trim())
                      .filter(Boolean);
                    const payload = {
                      name: nodeDraft.name.trim(),
                      kind: nodeDraft.kind,
                      client: clientId,
                      newId: newId !== editNodeId ? newId : null,
                      address: nodeDraft.address?.trim() || "",
                      workPhone: nodeDraft.workPhone?.trim() || "",
                      cellPhone: nodeDraft.cellPhone?.trim() || "",
                      emails,
                      photo: nodeDraft.photo?.trim() || "",
                      taxId: nodeDraft.taxId?.trim() || "",
                      accountingUrl: nodeDraft.accountingUrl?.trim() || "",
                      hrUrl: nodeDraft.hrUrl?.trim() || "",
                      logo: nodeDraft.logo?.trim() || "",
                    };
                    apiRequest(`/api/nodes/${editNodeId}`, {
                      method: "PUT",
                      body: JSON.stringify(payload),
                    })
                      .then(() => {
                        setNodeList((prev) =>
                          prev.map((n) =>
                            n.id === editNodeId
                              ? {
                                  ...n,
                                  id: newId,
                                  name: nodeDraft.name.trim(),
                                  kind: nodeDraft.kind,
                                  address: payload.address,
                                  workPhone: payload.workPhone,
                                  cellPhone: payload.cellPhone,
                                  emails: payload.emails,
                                  photo: payload.photo,
                                  taxId: payload.taxId,
                                  accountingUrl: payload.accountingUrl,
                                  hrUrl: payload.hrUrl,
                                  logo: payload.logo,
                                  client: n.client || clientId,
                                }
                              : n
                          )
                        );
                        if (newId !== editNodeId) {
                          setRelList((prev) =>
                            prev.map((r) => ({
                              ...r,
                              from: r.from === editNodeId ? newId : r.from,
                              to: r.to === editNodeId ? newId : r.to,
                            }))
                          );
                          if (focusId === editNodeId) {
                            setFocusId(newId);
                          }
                          setEditNodeId(newId);
                        }
                        setRemoteStatus("connected");
                        setOpenDialog(null);
                      })
                      .catch((err) => {
                        setRemoteStatus("error");
                        setRemoteError(err.message);
                      });
                  }}
                >
                  Update
                </Button>
              </div>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "edit-ownership" && (
          <DialogContent style={{ minWidth: 1000, maxWidth: "none" }}>
            <DialogHeader>
              <DialogTitle>Edit ownership</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Owner</label>
                <div style={{ display: "flex", alignItems: "center", gap: 8, width: "100%" }}>
                  <select
                    className="form-select"
                    value={ownershipDraft.from}
                    onChange={(e) =>
                      setOwnershipDraft((prev) => ({
                        ...prev,
                        from: e.target.value,
                      }))
                    }
                    style={{ flex: 1, minWidth: 0 }}
                  >
                    {nodeList.map((n) => (
                      <option key={n.id} value={n.id}>
                        {n.name}
                      </option>
                    ))}
                  </select>
                  <Button
                    type="button"
                    variant="outline"
                    className="btn-icon"
                    aria-label="Edit owner"
                    disabled={!ownershipDraft.from}
                    onClick={() => {
                      if (!ownershipDraft.from) return;
                      setEditNodeId(ownershipDraft.from);
                      setOpenDialog({ type: "edit-node" });
                    }}
                  >
                    <Pencil size={16} />
                  </Button>
                </div>
              </div>
              <div className="form-row">
                <label className="form-label">Owned entity</label>
                <div style={{ display: "flex", alignItems: "center", gap: 8, width: "100%" }}>
                  <select
                    className="form-select"
                    value={ownershipDraft.to}
                    onChange={(e) =>
                      setOwnershipDraft((prev) => ({
                        ...prev,
                        to: e.target.value,
                      }))
                    }
                    style={{ flex: 1, minWidth: 0 }}
                  >
                    {entityNodes.map((n) => (
                      <option key={n.id} value={n.id}>
                        {n.name}
                      </option>
                    ))}
                  </select>
                  <Button
                    type="button"
                    variant="outline"
                    className="btn-icon"
                    aria-label="Edit owned entity"
                    disabled={!ownershipDraft.to}
                    onClick={() => {
                      if (!ownershipDraft.to) return;
                      setEditNodeId(ownershipDraft.to);
                      setOpenDialog({ type: "edit-node" });
                    }}
                  >
                    <Pencil size={16} />
                  </Button>
                </div>
              </div>
              <div className="form-row">
                <label className="form-label">Percent</label>
                <input
                  className="form-input"
                  type="number"
                  min="0"
                  max="100"
                  value={ownershipDraft.percent}
                  onChange={(e) =>
                    setOwnershipDraft((prev) => ({
                      ...prev,
                      percent: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">Start date</label>
                <input
                  className="form-input"
                  type="date"
                  value={ownershipDraft.startDate}
                  onChange={(e) =>
                    setOwnershipDraft((prev) => ({
                      ...prev,
                      startDate: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">End date</label>
                <input
                  className="form-input"
                  type="date"
                  value={ownershipDraft.endDate}
                  onChange={(e) =>
                    setOwnershipDraft((prev) => ({
                      ...prev,
                      endDate: e.target.value,
                    }))
                  }
                />
              </div>
            </div>
            <DialogFooter
              style={{
                paddingLeft: 24,
                paddingRight: 24,
                marginLeft: 0,
                marginRight: 0,
              }}
            >
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  if (!editOwnershipId) return;
                  const relId = editOwnershipId;
                  setConfirmDialog({
                    title: "ERPlus says",
                    message: "Delete this ownership relationship?",
                    onConfirm: () => {
                      apiRequest("/api/relationships/owns", {
                        method: "DELETE",
                        body: JSON.stringify({
                          from: ownershipDraft.from,
                          to: ownershipDraft.to,
                          client: clientId,
                        }),
                      })
                        .then(() => {
                          setRelList((prev) => prev.filter((r) => r.id !== relId));
                          setRemoteStatus("connected");
                          setOpenDialog(null);
                          setConfirmDialog(null);
                        })
                        .catch((err) => {
                          setRemoteStatus("error");
                          setRemoteError(err.message);
                        });
                    },
                  });
                }}
              >
                Delete
              </Button>
              <Button
                type="button"
                onClick={() => {
                  if (!editOwnershipId) return;
                  const payload = {
                    from: ownershipDraft.from,
                    to: ownershipDraft.to,
                    percent: ownershipDraft.percent
                      ? Number(ownershipDraft.percent)
                      : null,
                    startDate: ownershipDraft.startDate || null,
                    endDate: ownershipDraft.endDate || null,
                    client: clientId,
                  };
                  apiRequest("/api/relationships/owns", {
                    method: "PUT",
                    body: JSON.stringify(payload),
                  })
                    .then(() => {
                      setRelList((prev) =>
                        prev.map((r) =>
                          r.id === editOwnershipId
                            ? {
                                ...r,
                                type: "owns",
                                ...payload,
                              }
                            : r
                        )
                      );
                      setRemoteStatus("connected");
                      setOpenDialog(null);
                    })
                    .catch((err) => {
                      setRemoteStatus("error");
                      setRemoteError(err.message);
                    });
                }}
              >
                Update
              </Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "edit-employment" && (
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Edit employment</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Employer</label>
                <select
                  className="form-select"
                  value={employmentDraft.from}
                  onChange={(e) =>
                    setEmploymentDraft((prev) => ({
                      ...prev,
                      from: e.target.value,
                    }))
                  }
                >
                  {entityNodes.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Person</label>
                <select
                  className="form-select"
                  value={employmentDraft.to}
                  onChange={(e) =>
                    setEmploymentDraft((prev) => ({
                      ...prev,
                      to: e.target.value,
                    }))
                  }
                >
                  {personNodes.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.name}
                    </option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Role</label>
                <input
                  className="form-input"
                  type="text"
                  value={employmentDraft.role}
                  onChange={(e) =>
                    setEmploymentDraft((prev) => ({
                      ...prev,
                      role: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">Start date</label>
                <input
                  className="form-input"
                  type="date"
                  value={employmentDraft.startDate}
                  onChange={(e) =>
                    setEmploymentDraft((prev) => ({
                      ...prev,
                      startDate: e.target.value,
                    }))
                  }
                />
              </div>
              <div className="form-row">
                <label className="form-label">End date</label>
                <input
                  className="form-input"
                  type="date"
                  value={employmentDraft.endDate}
                  onChange={(e) =>
                    setEmploymentDraft((prev) => ({
                      ...prev,
                      endDate: e.target.value,
                    }))
                  }
                />
              </div>
            </div>
            <DialogFooter
              style={{
                paddingLeft: 24,
                paddingRight: 24,
                marginLeft: 0,
                marginRight: 0,
              }}
            >
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  if (!editEmploymentId) return;
                  const relId = editEmploymentId;
                  setConfirmDialog({
                    title: "ERPlus says",
                    message: "Delete this employment relationship?",
                    onConfirm: () => {
                      apiRequest("/api/relationships/employs", {
                        method: "DELETE",
                        body: JSON.stringify({
                          from: employmentDraft.from,
                          to: employmentDraft.to,
                          client: clientId,
                        }),
                      })
                        .then(() => {
                          setRelList((prev) => prev.filter((r) => r.id !== relId));
                          setRemoteStatus("connected");
                          setOpenDialog(null);
                          setConfirmDialog(null);
                        })
                        .catch((err) => {
                          setRemoteStatus("error");
                          setRemoteError(err.message);
                        });
                    },
                  });
                }}
              >
                Delete
              </Button>
              <Button
                type="button"
                onClick={() => {
                  if (!editEmploymentId) return;
                  const payload = {
                    from: employmentDraft.from,
                    to: employmentDraft.to,
                    role: employmentDraft.role || null,
                    startDate: employmentDraft.startDate || null,
                    endDate: employmentDraft.endDate || null,
                    client: clientId,
                  };
                  apiRequest("/api/relationships/employs", {
                    method: "PUT",
                    body: JSON.stringify(payload),
                  })
                    .then(() => {
                      setRelList((prev) =>
                        prev.map((r) =>
                          r.id === editEmploymentId
                            ? {
                                ...r,
                                type: "employs",
                                ...payload,
                              }
                            : r
                        )
                      );
                      setRemoteStatus("connected");
                      setOpenDialog(null);
                    })
                    .catch((err) => {
                      setRemoteStatus("error");
                      setRemoteError(err.message);
                    });
                }}
              >
                Update
              </Button>
            </DialogFooter>
          </DialogContent>
        )}
      </Dialog>

      <Dialog open={Boolean(confirmDialog)} onOpenChange={() => setConfirmDialog(null)}>
        {confirmDialog && (
          <DialogContent>
            <DialogHeader>
              <DialogTitle>{confirmDialog.title || "ERPlus says"}</DialogTitle>
            </DialogHeader>
            <div style={{ paddingTop: 4 }}>{confirmDialog.message}</div>
            <DialogFooter>
              <Button variant="secondary" onClick={() => setConfirmDialog(null)}>Cancel</Button>
              <Button
                type="button"
                onClick={() => {
                  if (confirmDialog?.onConfirm) confirmDialog.onConfirm();
                }}
              >
                Delete
              </Button>
            </DialogFooter>
          </DialogContent>
        )}
      </Dialog>
    </div>
  );
}

// EOF