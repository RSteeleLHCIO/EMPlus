import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Card, CardContent } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { CalendarDays, Link, Users, Building2, Plus, Pencil, Trash2, ChevronRight, ChevronDown, Upload, X, Search, Settings, LogOut, GitFork, LayoutList, Hourglass, Home, Download, BookOpen, UserPlus } from "lucide-react";
import { generateEntityPdf, generateEntityBook } from "./utils/generateEntityPdf";
import ExportDialog from "./components/ExportDialog";
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
import { Switch } from "./components/ui/switch";

const DATA_TYPES = [
  { value: "string",     label: "Short text" },
  { value: "textarea",   label: "Long text" },
  { value: "number",     label: "Number" },
  { value: "currency",   label: "Currency" },
  { value: "percentage", label: "Percentage" },
  { value: "boolean",    label: "Yes / No" },
  { value: "date",       label: "Date" },
  { value: "time",       label: "Time" },
  { value: "phone",      label: "Phone number" },
  { value: "email",      label: "Email address" },
  { value: "link",       label: "URL / Link" },
  { value: "address",    label: "Address" },
  { value: "year",       label: "Year" },
];

const dataTypeToHtmlInput = (dt) => {
  switch (dt) {
    case "number": case "currency": case "percentage": case "year": return "number";
    case "email": return "email";
    case "link": return "url";
    case "date": return "date";
    case "time": return "time";
    case "phone": return "tel";
    default: return "text";
  }
};

const renderDdField = (field, value, onChange) => {
  const { fieldId, prompt, dataType, multiValue, validValues, phoneTypes } = field;

  if (dataType === "phone") {
    const types = phoneTypes?.length ? phoneTypes : ["Phone"];
    // Normalize stored value to [{type, number}] array; handle legacy {[type]:number} objects
    let entries;
    if (Array.isArray(value) && value.length > 0) {
      entries = value;
    } else if (value && typeof value === "object" && !Array.isArray(value)) {
      entries = Object.entries(value).map(([t, n]) => ({ type: t, number: n }));
    }
    if (!entries || entries.length === 0) entries = [{ type: types[0], number: "" }];

    const updateEntry = (idx, key, val) =>
      onChange(entries.map((e, i) => i === idx ? { ...e, [key]: val } : e));
    const removeEntry = (idx) => {
      const next = entries.filter((_, i) => i !== idx);
      onChange(next.length > 0 ? next : [{ type: types[0], number: "" }]);
    };
    const addEntry = () => onChange([...entries, { type: types[0], number: "" }]);

    return (
      <div className="form-row">
        <label className="form-label">{prompt}</label>
        <div style={{ display: "grid", gap: 8 }}>
          {entries.map((entry, idx) => (
            <div key={idx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
              {types.length > 1 && (
                <select
                  className="form-select"
                  style={{ width: 110, flexShrink: 0 }}
                  value={entry.type}
                  onChange={(e) => updateEntry(idx, "type", e.target.value)}
                >
                  {types.map((t) => <option key={t} value={t}>{t}</option>)}
                </select>
              )}
              {types.length === 1 && (
                <span style={{ fontSize: 12, color: "#6b7280", flexShrink: 0, minWidth: 50 }}>{types[0]}</span>
              )}
              <input
                className="form-input"
                style={{ flex: 1 }}
                type="tel"
                value={entry.number}
                onChange={(e) => updateEntry(idx, "number", e.target.value)}
                autoComplete="off"
                data-lpignore="true"
              />
              {multiValue && entries.length > 1 && (
                <button type="button" className="btn btn-outline" style={{ padding: "4px 8px", flexShrink: 0 }}
                  onClick={() => removeEntry(idx)}>✕</button>
              )}
            </div>
          ))}
          {multiValue && (
            <div><Button type="button" variant="outline" onClick={addEntry}>+ Add {prompt}</Button></div>
          )}
        </div>
      </div>
    );
  }

  if (dataType === "boolean") {
    return (
      <div className="form-row">
        <label className="form-label">{prompt}</label>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <Switch
            checked={value === "true" || value === true}
            onCheckedChange={(v) => onChange(v ? "true" : "false")}
          />
          <span style={{ fontSize: 14, color: (value === "true" || value === true) ? "#374151" : "#6b7280" }}>
            {(value === "true" || value === true) ? "Yes" : "No"}
          </span>
        </div>
      </div>
    );
  }

  if (multiValue) {
    const arrVal = Array.isArray(value) ? value : (value ? [value] : [""]);
    return (
      <div className="form-row">
        <label className="form-label">{prompt}</label>
        <div style={{ display: "grid", gap: 8 }}>
          {arrVal.map((v, idx) => (
            <div key={idx} style={{ display: "flex", gap: 8 }}>
              {validValues?.length > 0 ? (
                <select
                  className="form-select"
                  style={{ flex: 1 }}
                  value={v}
                  onChange={(e) => { const next = [...arrVal]; next[idx] = e.target.value; onChange(next); }}
                >
                  <option value=""></option>
                  {validValues.map((vv) => <option key={vv} value={vv}>{vv}</option>)}
                </select>
              ) : (
                <input
                  className="form-input"
                  style={{ flex: 1 }}
                  type={dataTypeToHtmlInput(dataType)}
                  value={v}
                  onChange={(e) => { const next = [...arrVal]; next[idx] = e.target.value; onChange(next); }}
                  autoComplete="off"
                  data-lpignore="true"
                />
              )}
              {arrVal.length > 1 && (
                <Button type="button" variant="outline" onClick={() => onChange(arrVal.filter((_, i) => i !== idx))}>Remove</Button>
              )}
            </div>
          ))}
          <div>
            <Button type="button" variant="outline" onClick={() => onChange([...arrVal, ""])}>Add {prompt}</Button>
          </div>
        </div>
      </div>
    );
  }

  if (validValues?.length > 0) {
    return (
      <div className="form-row">
        <label className="form-label">{prompt}</label>
        <select className="form-select" value={value || ""} onChange={(e) => onChange(e.target.value)}>
          <option value=""></option>
          {validValues.map((v) => <option key={v} value={v}>{v}</option>)}
        </select>
      </div>
    );
  }

  if (dataType === "textarea" || dataType === "address") {
    return (
      <div className="form-row">
        <label className="form-label">{prompt}</label>
        <textarea
          className="form-input"
          style={{ minHeight: 70, resize: "vertical", fontFamily: "inherit" }}
          value={value || ""}
          onChange={(e) => onChange(e.target.value)}
        />
      </div>
    );
  }

  return (
    <div className="form-row">
      <label className="form-label">{prompt}</label>
      <input
        className="form-input"
        type={dataTypeToHtmlInput(dataType)}
        value={value || ""}
        onChange={(e) => onChange(e.target.value)}
        autoComplete="off"
        data-lpignore="true"
      />
    </div>
  );
};

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

export default function EntityApp({ token, clientId: clientIdProp, onSignOut }) {
  const [nodeList, setNodeList] = useState(initialNodes);
  const [relList, setRelList] = useState(initialRelationships);
  const [homeScreen, setHomeScreen] = useState(() => {
    if (typeof window === "undefined") return null;
    try {
      const s = localStorage.getItem("homeScreen");
      return s ? JSON.parse(s) : null;
    } catch { return null; }
  });
  const [viewMode, setViewMode] = useState(() => homeScreen?.viewMode ?? "hierarchy");
  const [focusId, setFocusId] = useState(() => {
    if (typeof window === "undefined") return "entity:A";
    try {
      return homeScreen?.focusId || localStorage.getItem("focusId") || "entity:A";
    } catch {
      return "entity:A";
    }
  });
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [showCalendarPopup, setShowCalendarPopup] = useState(false);
  const calendarButtonRef = useRef(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [showDateSelection, setShowDateSelection] = useState(false);
  const [homeAnimating, setHomeAnimating] = useState(false);
  const [homeAnimOrigin, setHomeAnimOrigin] = useState("50% 50%");
  const settingsRef = useRef(null);
  const exportMenuRef = useRef(null);
  const homeButtonRef = useRef(null);
  const focusBoxRef = useRef(null);
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

  const clientId = clientIdProp || "test";

  const [newNode, setNewNode] = useState({
    name: "",
    kind: "entity",
    customFields: {},
  });
  const [dupMatches, setDupMatches] = useState([]);
  const [dirSearch, setDirSearch] = useState("");

  function checkDuplicateName(name) {
    const q = name.trim().toLowerCase();
    if (!q) { setDupMatches([]); return; }
    const matches = nodeList
      .filter((n) => {
        const t = (n.name || "").toLowerCase();
        if (t === q) return true;
        if (t.includes(q) || q.includes(t)) return true;
        // simple bigram overlap
        const bigrams = (s) => { const b = new Set(); for (let i = 0; i < s.length - 1; i++) b.add(s.slice(i, i + 2)); return b; };
        const bq = bigrams(q), bt = bigrams(t);
        let shared = 0;
        bq.forEach((g) => { if (bt.has(g)) shared++; });
        const score = (2 * shared) / (bq.size + bt.size || 1);
        return score >= 0.4;
      })
      .map((n) => n.name);
    setDupMatches(matches);
  }

  const [editNodeId, setEditNodeId] = useState(initialNodes[0]?.id ?? "");
  const [nodeDraft, setNodeDraft] = useState({
    name: "",
    kind: "entity",
    customFields: {},
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
  const [prevDialog, setPrevDialog] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [isAddingNode, setIsAddingNode] = useState(false);
  const [isAddingOwnership, setIsAddingOwnership] = useState(false);
  const [isAddingEmployment, setIsAddingEmployment] = useState(false);
  const [ownerEditorRows, setOwnerEditorRows] = useState([]);
  const [ownerEditorOriginal, setOwnerEditorOriginal] = useState([]);
  const [ownerSearch, setOwnerSearch] = useState("");
  const [ownerSearchOpen, setOwnerSearchOpen] = useState(false);
  const [isSavingOwners, setIsSavingOwners] = useState(false);
  const [isCreatingOwnerNode, setIsCreatingOwnerNode] = useState(false);
  const [isPdfExporting, setIsPdfExporting] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportReports, setExportReports] = useState([]);
  const [exportReportsLoaded, setExportReportsLoaded] = useState(false);
  const [addUserOpen, setAddUserOpen] = useState(false);
  const [addUserDraft, setAddUserDraft] = useState({ loginId: "", password: "", confirm: "" });
  const [addUserBusy, setAddUserBusy] = useState(false);
  const [addUserError, setAddUserError] = useState("");
  const [addUserSuccess, setAddUserSuccess] = useState("");
  const [collapsedOwnerNodes, setCollapsedOwnerNodes] = useState(() => new Set());
  const [collapsedOwnedNodes, setCollapsedOwnedNodes] = useState(() => new Set());

  const [dataDictionary, setDataDictionary] = useState([]);
  const emptyDdDraft = { prompt: "", dataType: "string", appliesTo: "both", multiValue: false, validValuesText: "", phoneTypesText: "" };
  const [ddEntryDraft, setDdEntryDraft] = useState(emptyDdDraft);
  const [ddEntryId, setDdEntryId] = useState(null);
  const [isSavingDdEntry, setIsSavingDdEntry] = useState(false);

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

  const dirSearchLower = dirSearch.trim().toLowerCase();
  const filteredEntityNodes = dirSearchLower
    ? sortedEntityNodes.filter((n) => (n.name || n.id || "").toLowerCase().includes(dirSearchLower))
    : sortedEntityNodes;
  const filteredPersonNodes = dirSearchLower
    ? sortedPersonNodes.filter((n) => (n.name || n.id || "").toLowerCase().includes(dirSearchLower))
    : sortedPersonNodes;

  const apiRequest = async (path, options = {}) => {
    const { headers: optHeaders, ...restOptions } = options;
    const response = await fetch(`${apiBase}${path}`, {
      headers: {
        "Content-Type": "application/json",
        ...(token ? { "Authorization": `Bearer ${token}` } : {}),
        ...(optHeaders || {}),
      },
      ...restOptions,
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
        headers: {
          "Content-Type": "application/json",
          ...(token ? { "Authorization": `Bearer ${token}` } : {}),
        },
        body: JSON.stringify({ debug: false }),
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

  const loadDataDictionary = async () => {
    try {
      const data = await apiRequest(`/api/data-dictionary?client=${encodeURIComponent(clientId)}`);
      setDataDictionary(Array.isArray(data) ? data : []);
    } catch {
      setDataDictionary([]);
    }
  };

  const openNewDdEntry = () => {
    setDdEntryId(null);
    setDdEntryDraft(emptyDdDraft);
    setOpenDialog({ type: "data-dictionary-entry" });
  };

  const openDdEntry = (entry) => {
    setDdEntryId(entry.id);
    setDdEntryDraft({
      prompt: entry.prompt,
      dataType: entry.dataType,
      appliesTo: entry.appliesTo || "both",
      multiValue: !!entry.multiValue,
      validValuesText: (entry.validValues || []).join("\n"),
      phoneTypesText: (entry.phoneTypes || []).join("\n"),
    });
    setOpenDialog({ type: "data-dictionary-entry" });
  };

  const saveDdEntry = async () => {
    if (!ddEntryDraft.prompt.trim()) return;
    setIsSavingDdEntry(true);
    const validValues = ddEntryDraft.validValuesText
      .split("\n")
      .map((v) => v.trim())
      .filter(Boolean);
    const phoneTypes = ddEntryDraft.dataType === "phone"
      ? ddEntryDraft.phoneTypesText.split("\n").map((v) => v.trim()).filter(Boolean)
      : [];
    const payload = {
      client: clientId,
      prompt: ddEntryDraft.prompt.trim(),
      dataType: ddEntryDraft.dataType,
      appliesTo: ddEntryDraft.appliesTo,
      multiValue: ddEntryDraft.multiValue,
      validValues,
      phoneTypes,
    };
    try {
      if (ddEntryId) {
        const updated = await apiRequest(`/api/data-dictionary/${encodeURIComponent(ddEntryId)}`, {
          method: "PUT",
          body: JSON.stringify(payload),
        });
        setDataDictionary((prev) => prev.map((e) => e.id === ddEntryId ? updated : e));
      } else {
        const created = await apiRequest("/api/data-dictionary", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        setDataDictionary((prev) => [...prev, created]);
      }
      setOpenDialog({ type: "data-dictionary" });
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
    } finally {
      setIsSavingDdEntry(false);
    }
  };

  const deleteDdEntry = (id) => {
    setConfirmDialog({
      title: "ERPlus says",
      message: "Delete this field definition?",
      onConfirm: async () => {
        try {
          await apiRequest(`/api/data-dictionary/${encodeURIComponent(id)}`, {
            method: "DELETE",
            body: JSON.stringify({ client: clientId }),
          });
          setDataDictionary((prev) => prev.filter((e) => e.id !== id));
          setConfirmDialog(null);
        } catch (err) {
          setRemoteStatus("error");
          setRemoteError(err.message);
          setConfirmDialog(null);
        }
      },
    });
  };

  const reorderDdEntry = async (id, direction) => {
    try {
      const updated = await apiRequest(`/api/data-dictionary/${encodeURIComponent(id)}/reorder`, {
        method: "PUT",
        body: JSON.stringify({ direction }),
      });
      setDataDictionary(Array.isArray(updated) ? updated : []);
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
    }
  };

  const makeNodeId = (kind, name, currentId = "") => {
    const slug = slugify(name || "node");
    const rawId = `${kind}:${slug || "node"}`;
    const next = normalizeClientId(clientId, rawId);
    return currentId && currentId === next ? currentId : next;
  };

  const makeRelId = () => `rel-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const openOwnerEditor = (targetId) => {
    const currentOwners = getOwnersOf(relList, targetId).map((item) => {
      const node = getNode(nodeList, item.nodeId);
      return {
        nodeId: item.nodeId,
        name: node?.name ?? item.nodeId,
        percent: item.rel?.percent != null ? String(item.rel.percent) : "",
        startDate: item.rel?.startDate ?? "",
        endDate: item.rel?.endDate ?? "",
        isNew: false,
      };
    });
    setOwnerEditorRows(currentOwners);
    setOwnerEditorOriginal(currentOwners);
    setOwnerSearch("");
    setOwnerSearchOpen(false);
    setOpenDialog({ type: "edit-owners", targetId });
  };

  const saveOwnerEditor = async () => {
    const targetId = openDialog?.targetId;
    if (!targetId) return;
    setIsSavingOwners(true);
    try {
      const removedRows = ownerEditorOriginal.filter(
        (orig) => !ownerEditorRows.find((r) => r.nodeId === orig.nodeId)
      );
      for (const row of removedRows) {
        await apiRequest("/api/relationships/owns", {
          method: "DELETE",
          body: JSON.stringify({ from: row.nodeId, to: targetId, client: clientId }),
        });
      }
      const changedRows = ownerEditorRows.filter((row) => {
        if (row.isNew) return false;
        const orig = ownerEditorOriginal.find((o) => o.nodeId === row.nodeId);
        return orig && String(row.percent) !== String(orig.percent);
      });
      for (const row of changedRows) {
        await apiRequest("/api/relationships/owns", {
          method: "PUT",
          body: JSON.stringify({
            from: row.nodeId, to: targetId,
            percent: row.percent !== "" ? Number(row.percent) : null,
            startDate: row.startDate || null, endDate: row.endDate || null,
            client: clientId,
          }),
        });
      }
      for (const row of ownerEditorRows.filter((r) => r.isNew)) {
        await apiRequest("/api/relationships/owns", {
          method: "POST",
          body: JSON.stringify({
            from: row.nodeId, to: targetId,
            percent: row.percent !== "" ? Number(row.percent) : null,
            startDate: row.startDate || null, endDate: row.endDate || null,
            client: clientId,
          }),
        });
      }
      setRelList((prev) => {
        let next = [...prev];
        for (const row of removedRows) {
          next = next.filter((r) => !(r.type === "owns" && r.from === row.nodeId && r.to === targetId));
        }
        for (const row of changedRows) {
          const idx = next.findIndex((r) => r.type === "owns" && r.from === row.nodeId && r.to === targetId);
          if (idx !== -1) next[idx] = { ...next[idx], percent: row.percent !== "" ? Number(row.percent) : null };
        }
        for (const row of ownerEditorRows.filter((r) => r.isNew)) {
          next.push({ id: makeRelId(), type: "owns", from: row.nodeId, to: targetId,
            percent: row.percent !== "" ? Number(row.percent) : null,
            startDate: row.startDate || null, endDate: row.endDate || null });
        }
        return next;
      });
      setRemoteStatus("connected");
      if (prevDialog) {
        setOpenDialog(prevDialog);
        setPrevDialog(null);
      } else {
        setOpenDialog(null);
      }
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
    } finally {
      setIsSavingOwners(false);
    }
  };

  const createOwnerNode = async (kind) => {
    const name = ownerSearch.trim();
    if (!name) return;
    setIsCreatingOwnerNode(true);
    try {
      const created = await apiRequest("/api/nodes", {
        method: "POST",
        body: JSON.stringify({ name, kind, client: clientId }),
      });
      setNodeList((prev) => [...prev, created]);
      setOwnerEditorRows((prev) => [
        ...prev,
        { nodeId: created.id, name: created.name, percent: "", startDate: "", endDate: "", isNew: true },
      ]);
      setOwnerSearch("");
      setOwnerSearchOpen(false);
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
    } finally {
      setIsCreatingOwnerNode(false);
    }
  };

  const getNodeName = (id) => getNode(nodeList, id)?.name || id;

  useEffect(() => {
    const node = getNode(nodeList, editNodeId);
    if (node) {
      setNodeDraft({
        name: node.name,
        kind: node.kind,
        customFields: node.customFields || {},
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

  // Back-button interception: close dialogs, then go home, then let browser navigate away
  const _backRef = useRef({});
  _backRef.current = { openDialog, prevDialog, uploadOpen, confirmDialog, viewMode, homeScreen, focusId };
  useEffect(() => {
    history.pushState({ emplus: true }, "");
    const handlePop = () => {
      const s = _backRef.current;
      let handled = false;
      if (s.confirmDialog) {
        setConfirmDialog(null);
        handled = true;
      } else if (s.openDialog) {
        if (s.prevDialog) {
          setOpenDialog(s.prevDialog);
          setPrevDialog(null);
          // Update ref immediately so the next back press sees the correct state
          // without waiting for React to re-render
          _backRef.current = { ...s, openDialog: s.prevDialog, prevDialog: null };
        } else {
          setOpenDialog(null);
          _backRef.current = { ...s, openDialog: null };
        }
        setDupMatches([]);
        setOwnerSearch("");
        setOwnerSearchOpen(false);
        handled = true;
      } else if (s.uploadOpen) {
        setUploadOpen(false);
        handled = true;
      } else if (s.homeScreen && (
        s.viewMode !== s.homeScreen.viewMode ||
        (s.homeScreen.focusId && s.focusId !== s.homeScreen.focusId)
      )) {
        setViewMode(s.homeScreen.viewMode);
        if (s.homeScreen.focusId) setFocusId(s.homeScreen.focusId);
        handled = true;
      }
      if (handled) {
        history.pushState({ emplus: true }, "");
      }
    };
    window.addEventListener("popstate", handlePop);
    return () => window.removeEventListener("popstate", handlePop);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!clientId) return;
    setFocusId((prev) => normalizeClientId(clientId, prev));
  }, [clientId]);

  useEffect(() => {
    let active = true;
    if (apiBase) {
      loadDirectory({ isActive: () => active });
      loadDataDictionary();
    }
    return () => {
      active = false;
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
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

  useEffect(() => {
    if (viewMode === "hierarchy" && focusBoxRef.current) {
      focusBoxRef.current.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }, [focusId, viewMode]);

  useEffect(() => {
    if (!settingsOpen) return;
    const handler = (e) => {
      if (settingsRef.current && !settingsRef.current.contains(e.target)) {
        setSettingsOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [settingsOpen]);

  useEffect(() => {
    if (!exportMenuOpen) return;
    const handler = (e) => {
      if (exportMenuRef.current && !exportMenuRef.current.contains(e.target)) {
        setExportMenuOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [exportMenuOpen]);

  const niceDate = selectedDate.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
  });

  return (
    <div style={viewMode === "hierarchy" ? {} : { paddingBottom: 120 }} data-lpignore="true">
      {homeAnimating && (
        <div className="home-anim-overlay" style={{ transformOrigin: homeAnimOrigin }} />
      )}
      <div className="app-header">
        <div style={{ maxWidth: "90%", margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 600 }}>{toSentenceCase(clientId)}</div>
          <div style={{ fontSize: 22, fontWeight: 600 }}>Entity Dashboard</div>
          {showDateSelection && <div className="header-date">{niceDate}</div>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {showDateSelection && (
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
          )}
          <Button
            type="button"
            variant="outline"
            className="btn-icon"
            aria-label={viewMode === "hierarchy" ? "Switch to Directory" : "Switch to Hierarchy"}
            title={viewMode === "hierarchy" ? "Directory" : "Hierarchy"}
            onClick={() => setViewMode(viewMode === "hierarchy" ? "directory" : "hierarchy")}
          >
            {viewMode === "hierarchy" ? <LayoutList size={18} /> : <GitFork size={18} />}
          </Button>
          <Button
            ref={homeButtonRef}
            type="button"
            variant="outline"
            className="btn-icon"
            aria-label="Go home"
            title="Home"
            onClick={() => {
              if (homeScreen) {
                setViewMode(homeScreen.viewMode);
                if (homeScreen.focusId) setFocusId(homeScreen.focusId);
              } else {
                setViewMode("hierarchy");
              }
            }}
          >
            <Home size={18} />
          </Button>
          <div className="settings-anchor" ref={exportMenuRef}>
            <Button
              type="button"
              variant="outline"
              className="btn-icon"
              aria-label="Export"
              title={viewMode === "hierarchy" ? "Export PDF" : "Export"}
              disabled={viewMode === "hierarchy" && isPdfExporting}
              onClick={async () => {
                if (viewMode === "hierarchy") {
                  if (isPdfExporting) return;
                  setIsPdfExporting(true);
                  try {
                    await generateEntityPdf({
                      nodeId: focusId,
                      nodeList,
                      relList,
                      dataDictionary,
                      clientName: toSentenceCase(clientId),
                    });
                  } finally {
                    setIsPdfExporting(false);
                  }
                } else {
                  setExportMenuOpen((prev) => !prev);
                }
              }}
            >
              <Download size={18} />
            </Button>
            {exportMenuOpen && viewMode !== "hierarchy" && (
              <div className="settings-menu">
                <button
                  className="settings-menu-item"
                  onClick={async () => {
                    setExportMenuOpen(false);
                    if (!exportReportsLoaded) {
                      try {
                        const data = await apiRequest(`/api/export-reports?client=${encodeURIComponent(clientId)}`);
                        setExportReports(Array.isArray(data) ? data : []);
                      } catch {
                        setExportReports([]);
                      }
                      setExportReportsLoaded(true);
                    }
                    setExportOpen(true);
                  }}
                >
                  <Download size={15} />
                  Export to Excel
                </button>
                <div className="settings-menu-divider" />
                <button
                  className="settings-menu-item"
                  onClick={async () => {
                    setExportMenuOpen(false);
                    setIsPdfExporting(true);
                    try {
                      const exportNodes = dirSearch.trim()
                        ? [...filteredEntityNodes, ...filteredPersonNodes]
                        : nodeList;
                      await generateEntityBook({
                        nodes: exportNodes,
                        nodeList,
                        relList,
                        dataDictionary,
                        clientName: toSentenceCase(clientId),
                        pageType: "hierarchy",
                        fileName: `${toSentenceCase(clientId)}-entity-book-hierarchy`,
                      });
                    } finally {
                      setIsPdfExporting(false);
                    }
                  }}
                >
                  <BookOpen size={15} />
                  Entity Book — Hierarchy pages
                </button>
                <button
                  className="settings-menu-item"
                  onClick={async () => {
                    setExportMenuOpen(false);
                    setIsPdfExporting(true);
                    try {
                      const exportNodes = dirSearch.trim()
                        ? [...filteredEntityNodes, ...filteredPersonNodes]
                        : nodeList;
                      await generateEntityBook({
                        nodes: exportNodes,
                        nodeList,
                        relList,
                        dataDictionary,
                        clientName: toSentenceCase(clientId),
                        pageType: "info",
                        fileName: `${toSentenceCase(clientId)}-entity-book-data`,
                      });
                    } finally {
                      setIsPdfExporting(false);
                    }
                  }}
                >
                  <BookOpen size={15} />
                  Entity Book — Data pages
                </button>
              </div>
            )}
          </div>
          <Button
            type="button"
            variant="outline"
            className={`btn-icon${remoteStatus === "loading" ? " btn-loading" : ""}`}
            aria-label="Loading…"
            title="Loading…"
            disabled
            style={{ visibility: remoteStatus === "loading" ? "visible" : "hidden" }}
          >
            <Hourglass size={18} />
          </Button>
          <div className="settings-anchor" ref={settingsRef}>
            <Button
              type="button"
              variant="outline"
              className="btn-icon"
              aria-label="Settings"
              onClick={() => setSettingsOpen((prev) => !prev)}
            >
              <Settings size={18} />
            </Button>
            {settingsOpen && (
              <div className="settings-menu">
                <button
                  className="settings-menu-item"
                  onClick={() => { setShowDateSelection((prev) => !prev); setSettingsOpen(false); }}
                >
                  <CalendarDays size={15} />
                  {showDateSelection ? "Hide date selection" : "Show date selection"}
                </button>
                <button
                  className="settings-menu-item"
                  onClick={() => {
                    const screen = {
                      viewMode,
                      focusId: viewMode === "hierarchy" ? focusId : null,
                    };
                    setHomeScreen(screen);
                    try { localStorage.setItem("homeScreen", JSON.stringify(screen)); } catch {}
                    const rect = homeButtonRef.current?.getBoundingClientRect();
                    const ox = rect ? `${Math.round(rect.left + rect.width / 2)}px` : "90vw";
                    const oy = rect ? `${Math.round(rect.top + rect.height / 2)}px` : "40px";
                    setHomeAnimOrigin(`${ox} ${oy}`);
                    setHomeAnimating(true);
                    setTimeout(() => setHomeAnimating(false), 650);
                    setSettingsOpen(false);
                  }}
                >
                  <Home size={15} />
                  Set as Home
                </button>
                <button
                  className="settings-menu-item"
                  onClick={() => {
                    setUploadOpen(true);
                    setUploadStatus("idle");
                    setUploadError("");
                    setUploadSummary(null);
                    setUploadFile(null);
                    setUploadDetected("");
                    setSettingsOpen(false);
                  }}
                >
                  <Upload size={15} />
                  Import CSV
                </button>
                <button
                  className="settings-menu-item"
                  onClick={() => {
                    setSettingsOpen(false);
                    loadDataDictionary();
                    setOpenDialog({ type: "data-dictionary" });
                  }}
                >
                  <BookOpen size={15} />
                  Data Dictionary
                </button>
                <button
                  className="settings-menu-item"
                  onClick={() => {
                    setSettingsOpen(false);
                    setAddUserDraft({ loginId: "", password: "", confirm: "" });
                    setAddUserError("");
                    setAddUserSuccess("");
                    setAddUserOpen(true);
                  }}
                >
                  <UserPlus size={15} />
                  Add User
                </button>
                <div className="settings-menu-divider" />
                <button
                  className="settings-menu-item settings-menu-item--danger"
                  onClick={() => { setSettingsOpen(false); if (onSignOut) onSignOut(); }}
                >
                  <LogOut size={15} />
                  Sign out
                </button>
              </div>
            )}
          </div>
        </div>
        </div>{/* end maxWidth wrapper */}
      </div>{/* end app-header */}

      <div className="app-content">

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

      {remoteStatus === "error" && (
        <div style={{ textAlign: "center", color: "#dc2626", marginBottom: 12 }}>
          {remoteError || "Unable to load directory"}
        </div>
      )}

      {viewMode === "hierarchy" && (
        <div className="hierarchy-vertical">

          <div className="hv-above">
          {/* ── OWNERS (above the focus box) ── */}
          {focusNode?.kind !== "person" && (
            <div className="hv-section">
              <div className="hv-section-header">
                <span className="hv-section-label">Owners</span>
                <button className="hv-section-add" title="Edit owners" onClick={() => openOwnerEditor(focusId)}>
                  <Pencil size={11} />
                </button>
              </div>
              <div className="hv-owners-row">
                {getOwnersOf(relList, focusId).length === 0 ? (
                  <div className="hv-empty">No owners recorded</div>
                ) : (
                  [...getOwnersOf(relList, focusId).map((item) => {
                    const ownerNode = getNode(nodeList, item.nodeId);
                    if (!ownerNode) return null;
                    const pct = item.rel?.percent;
                    const showPct = pct != null && Number.isFinite(Number(pct));
                    const isZero = showPct && Number(pct) === 0;
                    return (
                      <div
                        key={item.nodeId}
                        className={`hv-neighbor-box${isZero ? " hv-neighbor-box--zero" : ""}`}
                        onClick={() => setFocusId(item.nodeId)}
                        title={isZero ? "Non-economic / 0% interest" : "Click to focus"}
                      >
                        {ownerNode.kind === "person"
                          ? ownerNode.photo
                            ? <img src={ownerNode.photo} alt="" className="hv-neighbor-photo" />
                            : <Users size={22} className="hv-neighbor-icon" />
                          : ownerNode.logo
                            ? <img src={ownerNode.logo} alt="" className="hv-neighbor-logo" />
                            : <Building2 size={22} className="hv-neighbor-icon" />
                        }
                        <div className="hv-neighbor-name">{ownerNode.name}</div>
                        {showPct && (
                          <div className="hv-neighbor-pct">{Number(pct)}%</div>
                        )}
                      </div>
                    );
                  }),
                  (() => {
                    const gap = Math.round((100 - directOwnerTotal) * 10) / 10;
                    if (directOwnerTotalOk || directOwnerTotal === 0 || gap <= 0) return null;
                    return (
                      <div key="__unknown__" className="hv-neighbor-box hv-neighbor-box--unknown">
                        <div className="hv-neighbor-name">Unknown</div>
                        <div className="hv-neighbor-pct">{gap}%</div>
                      </div>
                    );
                  })()]
                )}
              </div>
            </div>
          )}

          {/* ── connector line from owners down to focus ── */}
          {focusNode?.kind !== "person" && getOwnersOf(relList, focusId).length > 0 && (
            <div className="hv-connector" />
          )}
          </div>{/* end hv-above */}

          {/* ── FOCUS BOX (centre) ── */}
          <div className="hv-focus-box" ref={focusBoxRef} onClick={() => {
            setEditNodeId(focusId);
            setOpenDialog({ type: "edit-node" });
          }}>
            {focusNode?.logo && (
              <img src={focusNode.logo} alt="" className="hv-focus-logo" />
            )}
            <div className="hv-focus-name">{focusNode?.name || focusId}</div>
            {focusNode?.kind === "person" && focusNode?.photo && (
              <img src={focusNode.photo} alt="" className="hv-focus-photo" />
            )}
          </div>

          <div className="hv-below">
          {/* ── connector line from focus down to owned ── */}
          {getOwnedBy(relList, focusId).length > 0 && (
            <div className="hv-connector" />
          )}

          {/* ── OWNED BY (below the focus box) ── */}
          {getOwnedBy(relList, focusId).length > 0 && (
            <div className="hv-section">
              <div className="hv-section-header">
                <span className="hv-section-label">Owns</span>
              </div>
              <div className="hv-owned-row">
              {getOwnedBy(relList, focusId).map((item) => {
                const ownedNode = getNode(nodeList, item.nodeId);
                if (!ownedNode) return null;
                const pct = item.rel?.percent;
                const showPct = pct != null && Number.isFinite(Number(pct));
                const isZero = showPct && Number(pct) === 0;
                return (
                  <div
                    key={item.nodeId}
                    className={`hv-neighbor-box${isZero ? " hv-neighbor-box--zero" : ""}`}
                    onClick={() => setFocusId(item.nodeId)}
                    title={isZero ? "Non-economic / 0% interest" : "Click to focus"}
                  >
                    {ownedNode.kind === "person"
                      ? ownedNode.photo
                        ? <img src={ownedNode.photo} alt="" className="hv-neighbor-photo" />
                        : <Users size={22} className="hv-neighbor-icon" />
                      : ownedNode.logo
                        ? <img src={ownedNode.logo} alt="" className="hv-neighbor-logo" />
                        : <Building2 size={22} className="hv-neighbor-icon" />
                    }
                    <div className="hv-neighbor-name">{ownedNode.name}</div>
                    {showPct && (
                      <div className="hv-neighbor-pct">{Number(pct)}%</div>
                    )}
                  </div>
                );
              })}
            </div>
            </div>
          )}
          {(focusEmployees.length > 0 || focusEmployers.length > 0) && (
            <Card style={{ marginTop: 24, alignSelf: "center", minWidth: 280, maxWidth: 480 }}>
              <CardContent>
                <div className="section-title">
                  {focusNode?.kind === "entity" ? "Employees" : "Employed by"}
                </div>
                <ul className="relationship-list">
                  {(focusNode?.kind === "entity" ? focusEmployees : focusEmployers).map((item) => (
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
                  ))}
                </ul>
              </CardContent>
            </Card>
          )}
          </div>{/* end hv-below */}
        </div>
      )}

      {viewMode === "directory" && (
        <div className="directory-grid">
          <div className="directory-search-bar">
            <Search size={15} style={{ color: "#9ca3af", flexShrink: 0 }} />
            <input
              type="text"
              placeholder="Search entities and people…"
              value={dirSearch}
              onChange={(e) => setDirSearch(e.target.value)}
              className="directory-search-input"
            />
            {dirSearch && (
              <button className="directory-search-clear" onClick={() => setDirSearch("")} title="Clear">
                <X size={13} />
              </button>
            )}
          </div>
          <Card>
            <CardContent>
              <div className="section-title">Entities {dirSearchLower ? `(${filteredEntityNodes.length})` : ""}</div>
              <div className="directory-scroll">
                <ul className="directory-list">
                  {filteredEntityNodes.map((n) => (
                    <li key={n.id}>
                      <div
                        className="directory-item"
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditNodeId(n.id);
                          setOpenDialog({ type: "edit-node" });
                        }}
                      >
                        {n.logo
                          ? <img src={n.logo} alt="" className="directory-thumb" />
                          : <Building2 className="directory-icon" />}
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
              <div className="section-title">People {dirSearchLower ? `(${filteredPersonNodes.length})` : ""}</div>
              <div className="directory-scroll">
                <ul className="directory-list">
                  {filteredPersonNodes.map((n) => (
                    <li key={n.id}>
                      <div
                        className="directory-item"
                        style={{ cursor: "pointer" }}
                        onClick={() => {
                          setEditNodeId(n.id);
                          setOpenDialog({ type: "edit-node" });
                        }}
                      >
                        {n.photo
                          ? <img src={n.photo} alt="" className="directory-thumb directory-thumb--round" />
                          : <Users className="directory-icon" />}
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

      {viewMode !== "hierarchy" && (
      <div className="fab-container">
        <Button
          type="button"
          className="fab-add"
          onClick={() => {
            setNewNode({ name: "", kind: "entity", customFields: {} });
            setOpenDialog({ type: "add-node" });
          }}
        >
          <Plus size={16} />
          <span>Add Entity</span>
        </Button>
        <Button
          type="button"
          className="fab-add"
          onClick={() => {
            setNewNode({ name: "", kind: "person", customFields: {} });
            setOpenDialog({ type: "add-node" });
          }}
        >
          <Plus size={16} />
          <span>Add Person</span>
        </Button>
      </div>
      )}

      <Dialog open={Boolean(openDialog)} onOpenChange={() => { setOpenDialog(null); setPrevDialog(null); setDupMatches([]); setOwnerSearch(""); setOwnerSearchOpen(false); }}>

        {openDialog?.type === "data-dictionary" && (
          <DialogContent className="dialog-content--tall" style={{ width: "min(900px, 92vw)", maxWidth: "none" }}>
            <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
              <DialogTitle>Data Dictionary — {toSentenceCase(clientId)}</DialogTitle>
            </DialogHeader>
            <div className="dialog-body">
              {dataDictionary.length === 0 ? (
                <div style={{ color: "#6b7280", fontSize: 14, padding: "12px 0" }}>
                  No fields defined yet. Use <strong>Add Field</strong> to create one.
                </div>
              ) : (
                <div style={{ overflowX: "clip" }}>
                  <table className="dd-table">
                    <thead>
                      <tr>
                        <th>Prompt</th>
                        <th>Type</th>
                        <th>Applies To</th>
                        <th>Multi-value</th>
                        <th>Valid Values</th>
                        <th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {/* Name — system field, always first, no edit/delete/reorder */}
                      <tr>
                        <td><em style={{ color: "#6b7280" }}>Name</em></td>
                        <td style={{ color: "#6b7280" }}>Short text</td>
                        <td style={{ color: "#6b7280" }}>Both</td>
                        <td style={{ color: "#6b7280" }}>No</td>
                        <td style={{ color: "#6b7280" }}></td>
                        <td></td>
                      </tr>
                      {[...dataDictionary]
                        .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                        .map((entry, idx, sorted) => (
                        <tr key={entry.id}>
                          <td>{entry.prompt}</td>
                          <td>{DATA_TYPES.find((t) => t.value === entry.dataType)?.label ?? entry.dataType}</td>
                          <td>
                            {entry.appliesTo === "person" ? "Person" : entry.appliesTo === "entity" ? "Entity" : "Both"}
                          </td>
                          <td>{entry.multiValue ? "Yes" : "No"}</td>
                          <td className="dd-valid-values">
                            {(entry.validValues || []).length > 0
                              ? entry.validValues.join(", ")
                              : <span style={{ color: "#9ca3af" }}>Free-form</span>}
                          </td>
                          <td>
                            <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                              <Button type="button" variant="outline" style={{ padding: "3px 7px", fontSize: 12, visibility: idx === 0 ? "hidden" : "visible" }} onClick={() => reorderDdEntry(entry.id, "up")}>↑</Button>
                              <Button type="button" variant="outline" style={{ padding: "3px 7px", fontSize: 12, visibility: idx === sorted.length - 1 ? "hidden" : "visible" }} onClick={() => reorderDdEntry(entry.id, "down")}>↓</Button>
                              <Button type="button" variant="outline" style={{ padding: "4px 8px" }} title="Edit" onClick={() => openDdEntry(entry)}><Pencil size={13} /></Button>
                              <Button type="button" variant="outline" style={{ padding: "4px 8px" }} title="Delete" onClick={() => deleteDdEntry(entry.id)}><Trash2 size={13} /></Button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
            <DialogFooter style={{ justifyContent: "space-between", marginTop: 16 }}>
              <Button type="button" variant="outline" onClick={openNewDdEntry}>
                <Plus size={14} style={{ marginRight: 6 }} />
                Add Field
              </Button>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Close</Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "data-dictionary-entry" && (
          <DialogContent style={{ width: "min(600px, 92vw)", maxWidth: "none" }}>
            <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
              <DialogTitle>{ddEntryId ? "Edit Field" : "Add Field"}</DialogTitle>
            </DialogHeader>
            <div className="form-grid">
              <div className="form-row">
                <label className="form-label">Prompt</label>
                <input
                  className="form-input"
                  type="text"
                  value={ddEntryDraft.prompt}
                  onChange={(e) => setDdEntryDraft((prev) => ({ ...prev, prompt: e.target.value }))}
                  placeholder="e.g. Annual Revenue"
                  autoComplete="off"
                  data-lpignore="true"
                />
              </div>
              <div className="form-row">
                <label className="form-label">Data type</label>
                <select
                  className="form-select"
                  value={ddEntryDraft.dataType}
                  onChange={(e) => setDdEntryDraft((prev) => ({ ...prev, dataType: e.target.value }))}
                >
                  {DATA_TYPES.map((t) => (
                    <option key={t.value} value={t.value}>{t.label}</option>
                  ))}
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Applies to</label>
                <select
                  className="form-select"
                  value={ddEntryDraft.appliesTo}
                  onChange={(e) => setDdEntryDraft((prev) => ({ ...prev, appliesTo: e.target.value }))}
                >
                  <option value="both">Entity &amp; Person</option>
                  <option value="entity">Entity only</option>
                  <option value="person">Person only</option>
                </select>
              </div>
              <div className="form-row">
                <label className="form-label">Multiple values</label>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <Switch
                    checked={ddEntryDraft.multiValue}
                    onCheckedChange={(v) => setDdEntryDraft((prev) => ({ ...prev, multiValue: v }))}
                  />
                  <span style={{ fontSize: 14, color: ddEntryDraft.multiValue ? "#374151" : "#6b7280" }}>
                    {ddEntryDraft.multiValue ? "Stores a list of values" : "Stores a single value"}
                  </span>
                </div>
              </div>
              {ddEntryDraft.dataType === "phone" ? (
                <div className="form-row">
                  <label className="form-label">Phone types (one per line — e.g. Work, Cell, Home)</label>
                  <textarea
                    className="form-input"
                    style={{ minHeight: 80, resize: "vertical", fontFamily: "inherit" }}
                    value={ddEntryDraft.phoneTypesText}
                    onChange={(e) => setDdEntryDraft((prev) => ({ ...prev, phoneTypesText: e.target.value }))}
                    placeholder={"Work\nCell\nHome\nFax"}
                  />
                </div>
              ) : (
                <div className="form-row">
                  <label className="form-label">Valid values (one per line — leave blank for free-form input)</label>
                  <textarea
                    className="form-input"
                    style={{ minHeight: 90, resize: "vertical", fontFamily: "inherit" }}
                    value={ddEntryDraft.validValuesText}
                    onChange={(e) => setDdEntryDraft((prev) => ({ ...prev, validValuesText: e.target.value }))}
                    placeholder={"Option A\nOption B\nOption C"}
                  />
                </div>
              )}
            </div>
            <DialogFooter style={{ marginTop: 16 }}>
              <Button variant="secondary" onClick={() => setOpenDialog({ type: "data-dictionary" })}>Cancel</Button>
              <Button type="button" disabled={isSavingDdEntry || !ddEntryDraft.prompt.trim()} onClick={saveDdEntry}>
                Save
              </Button>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "edit-owners" && (() => {
          const targetNode = getNode(nodeList, openDialog.targetId);
          const ownerTotal = ownerEditorRows.reduce(
            (s, r) => s + (r.percent !== "" && !isNaN(Number(r.percent)) ? Number(r.percent) : 0), 0
          );
          const overLimit = ownerTotal > 100;
          const searchQ = ownerSearch.trim().toLowerCase();
          const searchResults = searchQ
            ? nodeList.filter(
                (n) =>
                  n.name?.toLowerCase().includes(searchQ) &&
                  !ownerEditorRows.find((r) => r.nodeId === n.id) &&
                  n.id !== openDialog.targetId
              )
            : [];
          return (
            <DialogContent style={{ minWidth: 520, maxWidth: 640 }}>
              <DialogHeader>
                <DialogTitle>Owners of {targetNode?.name ?? openDialog.targetId}</DialogTitle>
              </DialogHeader>

              <div className="owner-editor">
                {ownerEditorRows.length === 0 && (
                  <div className="owner-editor-empty">No owners yet — search below to add one.</div>
                )}
                {ownerEditorRows.map((row, idx) => (
                  <div key={row.nodeId} className="owner-editor-row">
                    <div className="owner-editor-name">{row.name}</div>
                    <div className="owner-editor-pct-wrap">
                      <input
                        className="form-input owner-editor-pct-input"
                        type="number"
                        min="0"
                        max="100"
                        placeholder="%"
                        value={row.percent}
                        onChange={(e) =>
                          setOwnerEditorRows((prev) =>
                            prev.map((r, i) => i === idx ? { ...r, percent: e.target.value } : r)
                          )
                        }
                      />
                      <span className="owner-editor-pct-sign">%</span>
                    </div>
                    <button
                      className="owner-editor-remove"
                      title="Remove owner"
                      onClick={() => setOwnerEditorRows((prev) => prev.filter((_, i) => i !== idx))}
                    >
                      <X size={14} />
                    </button>
                  </div>
                ))}

                <div className={`owner-editor-total ${overLimit ? "over" : ""}`}>
                  Total:&nbsp;<strong>{ownerTotal}%</strong>
                  {overLimit && <span className="owner-editor-over-msg"> — exceeds 100%</span>}
                </div>

                <div className="owner-search-section">
                  <div className="owner-search-label">Add owner</div>
                  <div className="owner-search-container">
                    <Search size={14} className="owner-search-icon" />
                    <input
                      className="form-input owner-search-input"
                      type="text"
                      placeholder="Search entities and people…"
                      value={ownerSearch}
                      autoComplete="off"
                      data-lpignore="true"
                      onChange={(e) => { setOwnerSearch(e.target.value); setOwnerSearchOpen(true); }}
                      onFocus={() => setOwnerSearchOpen(true)}
                    />
                  </div>
                  {ownerSearchOpen && (searchResults.length > 0 || ownerSearch.trim()) && (
                    <div className="owner-search-dropdown">
                      {searchResults.map((n) => (
                        <div
                          key={n.id}
                          className="owner-search-result"
                          onMouseDown={() => {
                            setOwnerEditorRows((prev) => [
                              ...prev,
                              { nodeId: n.id, name: n.name, percent: "", startDate: "", endDate: "", isNew: true },
                            ]);
                            setOwnerSearch("");
                            setOwnerSearchOpen(false);
                          }}
                        >
                          <span className="owner-search-result-kind">{n.kind === "person" ? "Person" : "Entity"}</span>
                          {n.name}
                        </div>
                      ))}
                      {ownerSearch.trim() && (
                        <div className="owner-search-actions">
                          <button
                            className="owner-search-create"
                            disabled={isCreatingOwnerNode}
                            onMouseDown={() => createOwnerNode("entity")}
                          >
                            <Plus size={12} /> Create entity "{ownerSearch.trim()}"
                          </button>
                          <button
                            className="owner-search-create"
                            disabled={isCreatingOwnerNode}
                            onMouseDown={() => createOwnerNode("person")}
                          >
                            <Plus size={12} /> Create person "{ownerSearch.trim()}"
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>

              <DialogFooter>
                <Button variant="secondary" onClick={() => {
                  if (prevDialog) { setOpenDialog(prevDialog); setPrevDialog(null); }
                  else setOpenDialog(null);
                }}>Cancel</Button>
                <Button type="button" disabled={isSavingOwners} onClick={saveOwnerEditor}>
                  {isSavingOwners ? "Saving…" : "Save"}
                </Button>
              </DialogFooter>
            </DialogContent>
          );
        })()}
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
          <DialogContent className="dialog-content--tall" style={{ width: "min(1000px, 92vw)", maxWidth: "none" }}>
            <DialogHeader style={{marginBottom: '24px', marginLeft: 0}}>
              <DialogTitle>{"Add " + (newNode.kind === "person" ? "Person" : "Entity")}</DialogTitle>
            </DialogHeader>
            <div className="dialog-body">
            <div className="form-grid">
              {/* Name — always first, system field */}
              <div className="form-row">
                <label className="form-label">Name</label>
                <input
                  className="form-input"
                  type="text"
                  placeholder="Name"
                  value={newNode.name}
                  onChange={(e) => {
                    setNewNode((prev) => ({ ...prev, name: e.target.value }));
                    setDupMatches([]);
                  }}
                  onBlur={(e) => checkDuplicateName(e.target.value)}
                  autoComplete="off"
                  data-lpignore="true"
                />
                {dupMatches.length > 0 && (
                  <div className="dup-warning">
                    <strong>Possible duplicate{dupMatches.length > 1 ? "s" : ""}:</strong>{" "}
                    {dupMatches.join(", ")}
                  </div>
                )}
              </div>
              {[...dataDictionary]
                .filter((f) => f.appliesTo === "both" || f.appliesTo === newNode.kind)
                .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                .map((field) => (
                  <React.Fragment key={field.fieldId}>
                    {renderDdField(
                      field,
                      newNode.customFields?.[field.fieldId],
                      (val) => setNewNode((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } }))
                    )}
                  </React.Fragment>
                ))
              }
            </div>
            </div>
            <DialogFooter style={{ justifyContent: "space-between" }}>
              <Button
                type="button"
                variant="outline"
                disabled={!nodeList.some((n) => n.id === makeNodeId(newNode.kind, newNode.name))}
                onClick={() => {
                  const id = makeNodeId(newNode.kind, newNode.name);
                  openOwnerEditor(id);
                }}
              >
                Owners
              </Button>
              <div style={{ display: "flex", gap: 8 }}>
              <Button variant="secondary" onClick={() => setOpenDialog(null)}>Cancel</Button>
              <Button
                type="button"
                disabled={isAddingNode}
                onClick={async () => {
                  if (isAddingNode) return;
                  if (!newNode.name.trim()) return;
                  setIsAddingNode(true);
                  const id = makeNodeId(newNode.kind, newNode.name);
                  const payload = {
                    id,
                    name: newNode.name.trim(),
                    kind: newNode.kind,
                    client: clientId,
                    customFields: newNode.customFields || {},
                  };
                  try {
                    await apiRequest("/api/nodes", {
                      method: "POST",
                      body: JSON.stringify(payload),
                    });
                    setNodeList((prev) => [...prev, payload]);
                    if (!focusId) setFocusId(id);
                    setNewNode({ name: "", kind: newNode.kind, customFields: {} });
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
              </div>
            </DialogFooter>
          </DialogContent>
        )}

        {openDialog?.type === "add-ownership" && (
          <DialogContent style={{ width: "min(1000px, 92vw)", maxWidth: "none" }}>
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
          <DialogContent className="dialog-content--tall" style={{ width: "min(1000px, 92vw)", maxWidth: "none" }}>
            <DialogHeader  style={{marginBottom: '24px', marginLeft: 0}}>
              <DialogTitle>{nodeDraft.name || "Edit Node"}</DialogTitle>
            </DialogHeader>
            <div className="dialog-body">
            <div className="form-grid">
              {/* Name — always first, system field */}
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
              {[...dataDictionary]
                .filter((f) => f.appliesTo === "both" || f.appliesTo === nodeDraft.kind)
                .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                .map((field) => (
                  <React.Fragment key={field.fieldId}>
                    {renderDdField(
                      field,
                      nodeDraft.customFields?.[field.fieldId],
                      (val) => setNodeDraft((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } }))
                    )}
                  </React.Fragment>
                ))
              }
            </div>
            </div>
            <DialogFooter
              style={{
                justifyContent: "space-between",
                paddingLeft: 24,
                paddingRight: 24,
                marginLeft: 0,
                marginRight: 0,
                marginTop: 24
              }}
            >
              <div style={{ display: "flex", gap: 8 }}>
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
                {nodeDraft.kind !== "person" && (
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => {
                    if (!editNodeId) return;
                    setPrevDialog(openDialog);
                    openOwnerEditor(editNodeId);
                  }}
                >
                  Owners
                </Button>
                )}
                <Button
                  type="button"
                  variant="outline"
                  disabled={isPdfExporting}
                  onClick={async () => {
                    if (!editNodeId || isPdfExporting) return;
                    setIsPdfExporting(true);
                    try {
                      await generateEntityPdf({
                        nodeId: editNodeId,
                        nodeList,
                        relList,
                        dataDictionary,
                        clientName: toSentenceCase(clientId),
                      });
                    } finally {
                      setIsPdfExporting(false);
                    }
                  }}
                >
                  {isPdfExporting ? "Exporting…" : "Export"}
                </Button>
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
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
                    const payload = {
                      name: nodeDraft.name.trim(),
                      kind: nodeDraft.kind,
                      client: clientId,
                      newId: newId !== editNodeId ? newId : null,
                      customFields: nodeDraft.customFields || {},
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
                                  customFields: payload.customFields,
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
          <DialogContent style={{ width: "min(1000px, 92vw)", maxWidth: "none" }}>
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

      <ExportDialog
        open={exportOpen}
        onClose={() => setExportOpen(false)}
        exportNodes={
          viewMode === "directory" && dirSearch.trim()
            ? [...filteredEntityNodes, ...filteredPersonNodes]
            : nodeList
        }
        nodeList={nodeList}
        relList={relList}
        dataDictionary={dataDictionary}
        savedReports={exportReports}
        onReportSaved={(report) =>
          setExportReports((prev) => {
            const idx = prev.findIndex((r) => r.reportId === report.reportId);
            return idx >= 0
              ? prev.map((r) => (r.reportId === report.reportId ? report : r))
              : [...prev, report];
          })
        }
        onReportDeleted={(reportId) =>
          setExportReports((prev) => prev.filter((r) => r.reportId !== reportId))
        }
        apiRequest={apiRequest}
        clientName={toSentenceCase(clientId)}
      />

      {/* ── Add User dialog ── */}
      <Dialog open={addUserOpen} onOpenChange={(v) => { if (!v) setAddUserOpen(false); }}>
        <DialogContent style={{ maxWidth: 440 }}>
          <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
            <DialogTitle>Add User</DialogTitle>
          </DialogHeader>
          {addUserSuccess ? (
            <div style={{ padding: "16px 0" }}>
              <div style={{ color: "#16a34a", fontWeight: 600, marginBottom: 8 }}>User created!</div>
              <div style={{ fontSize: 13, color: "#374151" }}>
                Login ID: <strong>{addUserSuccess}</strong>
              </div>
              <DialogFooter style={{ marginTop: 24 }}>
                <Button onClick={() => setAddUserOpen(false)}>Close</Button>
              </DialogFooter>
            </div>
          ) : (
            <>
              <div className="form-grid">
                <div className="form-row">
                  <label className="form-label">Login ID</label>
                  <input
                    className="form-input"
                    type="text"
                    value={addUserDraft.loginId}
                    onChange={(e) => setAddUserDraft((p) => ({ ...p, loginId: e.target.value }))}
                    autoComplete="off"
                    data-lpignore="true"
                    placeholder="e.g. alice"
                  />
                </div>
                <div className="form-row">
                  <label className="form-label">Password</label>
                  <input
                    className="form-input"
                    type="password"
                    value={addUserDraft.password}
                    onChange={(e) => setAddUserDraft((p) => ({ ...p, password: e.target.value }))}
                    autoComplete="new-password"
                  />
                </div>
                <div className="form-row">
                  <label className="form-label">Confirm password</label>
                  <input
                    className="form-input"
                    type="password"
                    value={addUserDraft.confirm}
                    onChange={(e) => setAddUserDraft((p) => ({ ...p, confirm: e.target.value }))}
                    autoComplete="new-password"
                  />
                </div>
              </div>
              {addUserError && (
                <div style={{ color: "#dc2626", fontSize: 13, marginTop: 8 }}>{addUserError}</div>
              )}
              <DialogFooter style={{ marginTop: 24 }}>
                <Button variant="secondary" onClick={() => setAddUserOpen(false)}>Cancel</Button>
                <Button
                  type="button"
                  disabled={addUserBusy}
                  onClick={async () => {
                    const { loginId, password, confirm, setupSecret } = addUserDraft;
                    if (!loginId.trim()) { setAddUserError("Login ID is required."); return; }
                    if (!password) { setAddUserError("Password is required."); return; }
                    if (password !== confirm) { setAddUserError("Passwords do not match."); return; }
                    setAddUserError("");
                    setAddUserBusy(true);
                    try {
                      const data = await apiRequest("/api/auth/users", {
                        method: "POST",
                        body: JSON.stringify({
                          loginId: loginId.trim().toLowerCase(),
                          password,
                        }),
                      });
                      setAddUserSuccess(data.loginId);
                    } catch (err) {
                      setAddUserError(err.message);
                    } finally {
                      setAddUserBusy(false);
                    }
                  }}
                >
                  {addUserBusy ? "Creating…" : "Create User"}
                </Button>
              </DialogFooter>
            </>
          )}
        </DialogContent>
      </Dialog>

      </div>{/* end app-content */}
    </div>
  );
}

// EOF
