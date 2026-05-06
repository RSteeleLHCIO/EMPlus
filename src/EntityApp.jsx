import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Card, CardContent } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { Link, Users, Building2, Plus, Pencil, Trash2, ChevronRight, ChevronDown, Upload, X, Search, Settings, LogOut, GitFork, LayoutList, Home, Download, BookOpen, User, UserPlus, Loader2, Printer } from "lucide-react";
import { generateEntityPdf, generateEntityBook, generateEntityBookInterleaved } from "./utils/generateEntityPdf";
import ExportDialog from "./components/ExportDialog";
import { normalizePhone, formatPhone } from "./utils/helpers";
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
  { value: "file",       label: "File / Image" },
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

// ── Phone input: formats for display, normalises to E.164 on blur ─────────────────
const PhoneInputRow = ({ value, onCommit, style }) => {
  const [raw, setRaw] = React.useState(() => formatPhone(value));
  React.useEffect(() => { setRaw(formatPhone(value)); }, [value]);
  return (
    <input
      className="form-input"
      style={style}
      type="tel"
      value={raw}
      onChange={(e) => setRaw(e.target.value)}
      onBlur={() => {
        const normalized = normalizePhone(raw);
        onCommit(normalized);
        setRaw(formatPhone(normalized));
      }}
      autoComplete="off"
      data-lpignore="true"
    />
  );
};

// ── External URL field — user types/pastes a URL, favicon shown as preview ──
const DdLinkField = ({ prompt, value, onChange }) => {
  const faviconUrl = (url) => {
    try {
      const { hostname } = new URL(url);
      return `https://www.google.com/s2/favicons?domain=${encodeURIComponent(hostname)}&sz=64`;
    } catch {
      return null;
    }
  };
  const favicon = value ? faviconUrl(value) : null;

  return (
    <div className="form-row">
      <label className="form-label">{prompt}</label>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <input
          className="form-input"
          type="url"
          value={value || ""}
          onChange={(e) => onChange(e.target.value)}
          placeholder="https://…"
          autoComplete="off"
          data-lpignore="true"
        />
        {value && (
          <a href={value} target="_blank" rel="noopener noreferrer"
            style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "#2563eb", wordBreak: "break-all", textDecoration: "none" }}>
            {favicon && (
              <img src={favicon} alt="" width={16} height={16}
                style={{ flexShrink: 0, borderRadius: 2 }}
                onError={(e) => { e.currentTarget.style.display = "none"; }}
              />
            )}
            <span style={{ textDecoration: "underline" }}>{value}</span>
          </a>
        )}
      </div>
    </div>
  );
};

// ── Uploaded file field — URL hidden from user; shows thumbnail + Replace btn ──
const DdFileField = ({ prompt, value, onChange, apiBase, token }) => {
  const [uploading, setUploading] = React.useState(false);
  const [imgError, setImgError] = React.useState(false);
  const fileInputRef = React.useRef(null);

  React.useEffect(() => { setImgError(false); }, [value]);

  const isImage = (url) => /\.(png|jpe?g|gif|webp|svg|bmp)(\?|$)/i.test(url || "");
  const isVideo = (url) => /\.(mp4|webm|ogv)(\?|$)/i.test(url || "");

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${apiBase}/api/upload`, {
        method: "POST",
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        body: fd,
      });
      if (!res.ok) throw new Error(`Upload failed: ${res.status}`);
      const { url } = await res.json();
      onChange(url);
    } catch (err) {
      console.error("Upload error", err);
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  // Derive a display name from the S3 key (last path segment without the hex prefix).
  const displayName = (url) => {
    try {
      const seg = new URL(url).pathname.split("/").pop() || "file";
      // Strip the 32-char hex random prefix: "<hex>.<ext>" → ".<ext>"
      return seg.replace(/^[0-9a-f]{32}\./, "");
    } catch {
      return "file";
    }
  };

  return (
    <div className="form-row">
      <label className="form-label">{prompt}</label>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {value && isImage(value) && !imgError && (
          <a href={value} target="_blank" rel="noopener noreferrer">
            <img
              src={value}
              alt={prompt}
              style={{ maxWidth: 220, maxHeight: 150, borderRadius: 4, border: "1px solid #e5e7eb", display: "block" }}
              onError={() => setImgError(true)}
            />
          </a>
        )}
        {value && isVideo(value) && (
          <video src={value}
            style={{ maxWidth: 220, maxHeight: 150, borderRadius: 4, border: "1px solid #e5e7eb" }}
            controls
          />
        )}
        {value && !isImage(value) && !isVideo(value) && (
          <a href={value} target="_blank" rel="noopener noreferrer"
            style={{ fontSize: 12, color: "#2563eb", textDecoration: "underline" }}>
            {displayName(value)}
          </a>
        )}
        {value && isImage(value) && imgError && (
          <a href={value} target="_blank" rel="noopener noreferrer"
            style={{ fontSize: 12, color: "#2563eb", textDecoration: "underline" }}>
            {displayName(value)}
          </a>
        )}
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <button
            type="button"
            title={uploading ? "Uploading…" : value ? "Replace file" : "Upload file"}
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "5px 10px", border: "1px solid #d1d5db", borderRadius: 6,
              background: "#fff", cursor: uploading ? "not-allowed" : "pointer",
              fontSize: 13, color: "#374151",
            }}
            onClick={() => fileInputRef.current?.click()}
            disabled={uploading}
          >
            <Upload size={14} />
            {uploading ? "Uploading…" : value ? "Replace" : "Upload file"}
          </button>
          {value && (
            <button
              type="button"
              title="Remove"
              style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "5px 10px", border: "1px solid #fca5a5", borderRadius: 6,
                background: "#fff", cursor: "pointer", fontSize: 13, color: "#dc2626",
              }}
              onClick={() => onChange("")}
            >
              <X size={14} /> Remove
            </button>
          )}
        </div>
        <input ref={fileInputRef} type="file" style={{ display: "none" }} onChange={handleFileChange} />
      </div>
    </div>
  );
};

// ── Built-in photo / logo upload (persons → photo, entities → logo) ──────────
const NodeImageField = ({ kind, value, onChange, apiBase, token }) => {
  const [uploading, setUploading] = React.useState(false);
  const [imgError, setImgError] = React.useState(false);
  const fileInputRef = React.useRef(null);
  React.useEffect(() => { setImgError(false); }, [value]);

  const label = kind === "person" ? "Photo" : "Logo";
  const accept = "image/png, image/jpeg, image/gif, image/webp, image/svg+xml";

  const handleFileChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const res = await fetch(`${apiBase}/api/upload`, {
        method: "POST",
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        body: fd,
      });
      if (!res.ok) throw new Error(`Upload failed: ${res.status}`);
      const { url } = await res.json();
      onChange(url);
    } catch (err) {
      console.error("NodeImageField upload error", err);
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const imgStyle = kind === "person"
    ? { width: 72, height: 72, borderRadius: "50%", objectFit: "cover", border: "1px solid #e5e7eb" }
    : { width: 72, height: 72, objectFit: "contain", border: "1px solid #e5e7eb", borderRadius: 4, background: "#f9fafb" };

  return (
    <div className="form-row">
      <label className="form-label">{label}</label>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        {value && !imgError
          ? <img src={value} alt={label} style={imgStyle} onError={() => setImgError(true)} />
          : <div style={{ ...imgStyle, display: "flex", alignItems: "center", justifyContent: "center", color: "#d1d5db" }}>
              {kind === "person" ? <User size={32} /> : <Building2 size={32} />}
            </div>
        }
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          <button
            type="button"
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "5px 10px", border: "1px solid #d1d5db", borderRadius: 6,
              background: "#fff", cursor: uploading ? "not-allowed" : "pointer",
              fontSize: 13, color: "#374151",
            }}
            disabled={uploading}
            onClick={() => fileInputRef.current?.click()}
          >
            <Upload size={14} />
            {uploading ? "Uploading…" : value ? `Replace ${label}` : `Upload ${label}`}
          </button>
          {value && (
            <button
              type="button"
              style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "5px 10px", border: "1px solid #fca5a5", borderRadius: 6,
                background: "#fff", cursor: "pointer", fontSize: 13, color: "#dc2626",
              }}
              onClick={() => onChange("")}
            >
              <X size={14} /> Remove
            </button>
          )}
        </div>
        <input ref={fileInputRef} type="file" accept={accept} style={{ display: "none" }} onChange={handleFileChange} />
      </div>
    </div>
  );
};

const renderDdField = (field, value, onChange, { apiBase, token } = {}) => {
  const { fieldId, prompt, dataType, multiValue, validValues, phoneTypes } = field;

  if (dataType === "link") {
    return <DdLinkField key={fieldId} prompt={prompt} value={value} onChange={onChange} />;
  }

  if (dataType === "file") {
    return <DdFileField key={fieldId} prompt={prompt} value={value} onChange={onChange} apiBase={apiBase} token={token} />;
  }

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
              <PhoneInputRow
                value={entry.number}
                onCommit={(val) => updateEntry(idx, "number", val)}
                style={{ flex: 1 }}
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

// ── Directory stats strip ────────────────────────────────────────────────────
function StatsStrip({ allEntityNodes, allPersonNodes, filteredEntityNodes, filteredPersonNodes, dataDictionary, isFiltered }) {
  const INTERVAL_MS = 3000;

  // Build rotating slides from built-in fields + DD showInStats fields
  const slides = useMemo(() => {
    const result = [];
    const allNodes = [...allEntityNodes, ...allPersonNodes];
    const filteredNodes = [...filteredEntityNodes, ...filteredPersonNodes];

    const makeSlide = (label, nodes, getValue) => {
      const counts = {};
      let unset = 0;
      for (const n of nodes) {
        const v = getValue(n);
        if (v) { counts[v] = (counts[v] || 0) + 1; }
        else unset++;
      }
      if (nodes.length === 0) return null; // no nodes to report on
      return { label, counts, unset };
    };

    // Built-in entity fields
    const entitySlide1 = makeSlide("Operational Role", filteredEntityNodes, (n) => n.operationalRole);
    if (entitySlide1) result.push(entitySlide1);
    const entitySlide2 = makeSlide("Legal Status", filteredEntityNodes, (n) => n.legalStatus);
    if (entitySlide2) result.push(entitySlide2);

    // Built-in person field
    const personSlide = makeSlide("Person Status", filteredPersonNodes, (n) => n.personStatus);
    if (personSlide) result.push(personSlide);

    // DD showInStats fields
    const statsFields = [...dataDictionary]
      .filter((f) => f.showInStats && (f.validValues || []).length > 0)
      .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0));
    for (const field of statsFields) {
      const nodes = field.appliesTo === "person" ? filteredPersonNodes
                  : field.appliesTo === "entity" ? filteredEntityNodes
                  : filteredNodes;
      const slide = makeSlide(field.prompt, nodes, (n) => n.customFields?.[field.fieldId]);
      if (slide) result.push(slide);
    }
    return result;
  }, [allEntityNodes, allPersonNodes, filteredEntityNodes, filteredPersonNodes, dataDictionary]);

  const [slideIdx, setSlideIdx] = useState(0);
  const [visible, setVisible] = useState(true);
  const pausedRef = useRef(false);
  const timerRef = useRef(null);

  // Reset to first slide when slides change
  useEffect(() => { setSlideIdx(0); }, [slides]);

  // Advance with crossfade
  const advance = useCallback(() => {
    if (slides.length < 2) return;
    setVisible(false);
    setTimeout(() => {
      setSlideIdx((i) => (i + 1) % slides.length);
      setVisible(true);
    }, 150);
  }, [slides.length]);

  useEffect(() => {
    if (slides.length < 2) return;
    timerRef.current = setInterval(() => {
      if (!pausedRef.current) advance();
    }, INTERVAL_MS);
    return () => clearInterval(timerRef.current);
  }, [slides.length, advance]);

  const totalEntities = allEntityNodes.length;
  const totalPeople = allPersonNodes.length;
  const shownEntities = filteredEntityNodes.length;
  const shownPeople = filteredPersonNodes.length;

  const line1 = isFiltered
    ? `${shownEntities} of ${totalEntities} ${totalEntities === 1 ? "Entity" : "Entities"} · ${shownPeople} of ${totalPeople} ${totalPeople === 1 ? "Person" : "People"}`
    : `${totalEntities} ${totalEntities === 1 ? "Entity" : "Entities"} · ${totalPeople} ${totalPeople === 1 ? "Person" : "People"}`;

  const slide = slides[slideIdx];

  return (
    <div className="stats-strip">
      {isFiltered && (
        <div className="stats-filtered-badge">List is filtered</div>
      )}
      <div className="stats-line1">{line1}</div>
      {slide && (
        <div
          className="stats-line2"
          style={{ opacity: visible ? 1 : 0, transition: "opacity 150ms ease" }}
          onMouseEnter={() => { pausedRef.current = true; }}
          onMouseLeave={() => { pausedRef.current = false; }}
          onTouchStart={() => { pausedRef.current = true; }}
          onTouchEnd={() => { pausedRef.current = false; advance(); }}
          onClick={() => advance()}
          title="Tap to advance"
        >
          <span className="stats-slide-label">{slide.label}:</span>
          {Object.entries(slide.counts).map(([val, count]) => (
            <span key={val} className="stats-slide-value">{count} {val}</span>
          ))}
          {slide.unset > 0 && (
            <span className="stats-slide-unset">{slide.unset} —</span>
          )}
          {slides.length > 1 && (
            <span className="stats-dots">
              {slides.map((_, i) => (
                <span key={i} className={`stats-dot${i === slideIdx ? " stats-dot--active" : ""}`} />
              ))}
            </span>
          )}
        </div>
      )}
    </div>
  );
}

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
  // homeScreen is loaded from the server after login (see effect below)
  const [viewMode, setViewMode] = useState(() => homeScreen?.viewMode ?? "hierarchy");
  const [focusId, setFocusId] = useState(() => {
    if (typeof window === "undefined") return "entity:A";
    try {
      return homeScreen?.focusId || localStorage.getItem("focusId") || "entity:A";
    } catch {
      return "entity:A";
    }
  });
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [homeAnimating, setHomeAnimating] = useState(false);
  const [homeAnimOrigin, setHomeAnimOrigin] = useState("50% 50%");
  const settingsRef = useRef(null);
  const exportMenuRef = useRef(null);
  const homeButtonRef = useRef(null);
  const focusBoxRef = useRef(null);
  const hierarchyContainerRef = useRef(null);
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
    photo: "",
    logo: "",
    operationalRole: "",
    legalStatus: "",
    personStatus: "",
    customFields: {},
  });
  const [dupMatches, setDupMatches] = useState([]);
  const [dirSearch, setDirSearch] = useState(() => {
    if (typeof window === "undefined") return "";
    try {
      const s = localStorage.getItem("homeScreen");
      const hs = s ? JSON.parse(s) : null;
      return (hs?.viewMode === "directory" && hs?.dirFilter) ? hs.dirFilter : "";
    } catch { return ""; }
  });

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
  const [printDialogOpen, setPrintDialogOpen] = useState(false);
  const [printHierarchy, setPrintHierarchy] = useState(true);
  const [printDetail, setPrintDetail] = useState(true);
  const [exportResult, setExportResult] = useState(null); // { url, fileName }
  const exportResultRef = useRef(null);
  const pdfCancelRef = useRef(false);
  const [pdfProgress, setPdfProgress] = useState(null); // { current, total } | null
  const setExportResultAndRevoke = (result) => {
    if (exportResultRef.current?.url) URL.revokeObjectURL(exportResultRef.current.url);
    exportResultRef.current = result;
    setExportResult(result);
  };
  const [manageUsersOpen, setManageUsersOpen] = useState(false);
  const [manageUsersList, setManageUsersList] = useState([]);
  const [manageUsersLoading, setManageUsersLoading] = useState(false);
  const [manageUsersError, setManageUsersError] = useState("");
  const [manageUsersUpdating, setManageUsersUpdating] = useState(() => new Set());
  const [addUserDraft, setAddUserDraft] = useState({ loginId: "", password: "", confirm: "" });
  const [addUserBusy, setAddUserBusy] = useState(false);
  const [addUserError, setAddUserError] = useState("");
  const [addUserSuccess, setAddUserSuccess] = useState("");
  const [accountInfoOpen, setAccountInfoOpen] = useState(false);
  const [accountInfoDraft, setAccountInfoDraft] = useState({ name: "", email: "", cellPhone: "", workPhone: "" });
  const [accountInfoBusy, setAccountInfoBusy] = useState(false);
  const [accountInfoError, setAccountInfoError] = useState("");
  const [myLoginId, setMyLoginId] = useState("");
  const [myRole, setMyRole] = useState(() => {
    try { return localStorage.getItem("myRole") || "user"; } catch { return "user"; }
  });
  const [clientDisplayName, setClientDisplayName] = useState(() => {
    try { return localStorage.getItem("clientDisplayName") || ""; } catch { return ""; }
  });
  const [clientInfoOpen, setClientInfoOpen] = useState(false);
  const [clientInfoDraft, setClientInfoDraft] = useState({ clientName: "", address: "", billingContact: "", billingEmail: "", billingPhone: "", notes: "" });
  const [clientInfoBusy, setClientInfoBusy] = useState(false);
  const [clientInfoError, setClientInfoError] = useState("");
  const [serverInfoOpen, setServerInfoOpen] = useState(false);
  const [serverInfoData, setServerInfoData] = useState(null);
  const [serverInfoBusy, setServerInfoBusy] = useState(false);
  const [collapsedOwnerNodes, setCollapsedOwnerNodes] = useState(() => new Set());
  const [collapsedOwnedNodes, setCollapsedOwnedNodes] = useState(() => new Set());

  const [dataDictionary, setDataDictionary] = useState([]);
  const emptyDdDraft = { prompt: "", dataType: "string", appliesTo: "both", multiValue: false, validValuesText: "", phoneTypesText: "", showInStats: false };
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
    if (response.status === 401) {
      if (onSignOut) onSignOut(true);
      throw new Error("Session expired");
    }
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
      showInStats: !!entry.showInStats,
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
      showInStats: validValues.length > 0 ? ddEntryDraft.showInStats : false,
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
        photo: node.photo || "",
        logo: node.logo || "",
        operationalRole: node.operationalRole || "",
        legalStatus: node.legalStatus || "",
        personStatus: node.personStatus || "",
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
      // Load persisted homeScreen from user record (overrides localStorage)
      apiRequest("/api/auth/me").then((data) => {
        if (!active) return;
        if (data?.loginId) {
          setMyLoginId(data.loginId);
        }
        if (data?.role) {
          setMyRole(data.role);
          try { localStorage.setItem("myRole", data.role); } catch {}
        }
        if (data?.clientName) {
          setClientDisplayName(data.clientName);
          try { localStorage.setItem("clientDisplayName", data.clientName); } catch {}
        }
        if (!data?.homeScreen) return;
        setHomeScreen(data.homeScreen);
        try { localStorage.setItem("homeScreen", JSON.stringify(data.homeScreen)); } catch {}
        setViewMode(data.homeScreen.viewMode ?? "hierarchy");
        if (data.homeScreen.focusId) setFocusId(data.homeScreen.focusId);
        if (data.homeScreen.viewMode === "directory") setDirSearch(data.homeScreen.dirFilter ?? "");
      }).catch(() => {});
    }
    return () => {
      active = false;
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [apiBase, clientId]);

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
    if (viewMode !== "hierarchy") return;
    const center = (behavior) => {
      if (!focusBoxRef.current || !hierarchyContainerRef.current) return;
      const container = hierarchyContainerRef.current;
      const box = focusBoxRef.current;
      const containerRect = container.getBoundingClientRect();
      const boxRect = box.getBoundingClientRect();
      const delta = (boxRect.top + box.clientHeight / 2) - (containerRect.top + container.clientHeight / 2);
      container.scrollTo({ top: container.scrollTop + delta, behavior });
    };
    requestAnimationFrame(() => center("smooth"));
    // Re-center after images in the focus box have loaded (photos/logos load asynchronously)
    const t = setTimeout(() => center("instant"), 300);
    return () => clearTimeout(t);
  }, [focusId, viewMode, nodeList]);

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


  return (
    <div style={viewMode === "hierarchy" ? {} : { paddingBottom: 120 }} data-lpignore="true">
      {homeAnimating && (
        <div className="home-anim-overlay" style={{ transformOrigin: homeAnimOrigin }} />
      )}
      <div className="app-header">
        <div style={{ maxWidth: "90%", margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <button
            type="button"
            aria-label="Go home"
            title="Home"
            style={{ background: "none", border: "none", padding: 0, cursor: "pointer", lineHeight: 0 }}
            onClick={() => {
              if (homeScreen) {
                setViewMode(homeScreen.viewMode);
                if (homeScreen.focusId) setFocusId(homeScreen.focusId);
                if (homeScreen.viewMode === "directory") setDirSearch(homeScreen.dirFilter ?? "");
              } else {
                setViewMode("hierarchy");
              }
            }}
          >
            <img src="/emplus-logo.png" alt="EMPlus" style={{ height: 80, width: "auto", margin: "-10px" }} />
          </button>
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, color: "#1a1a2e" }}>{clientDisplayName || toSentenceCase(clientId)}</div>
            <div style={{ fontSize: 13, color: "#64748b" }}>Entity Dashboard</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ display: "flex", borderRadius: 6, overflow: "hidden", border: "1px solid #cbd5e1" }}>
            <button
              type="button"
              aria-label="Hierarchy view"
              title="Hierarchy"
              onClick={() => setViewMode("hierarchy")}
              style={{
                display: "flex", alignItems: "center", gap: 5, padding: "5px 12px",
                fontSize: 13, fontWeight: 500, cursor: "pointer", border: "none",
                borderRight: "1px solid #cbd5e1",
                background: viewMode === "hierarchy" ? "#1e293b" : "#fff",
                color: viewMode === "hierarchy" ? "#fff" : "#475569",
              }}
            >
              <GitFork size={14} /> Hierarchy
            </button>
            <button
              type="button"
              aria-label="Directory view"
              title="Directory"
              onClick={() => setViewMode("directory")}
              style={{
                display: "flex", alignItems: "center", gap: 5, padding: "5px 12px",
                fontSize: 13, fontWeight: 500, cursor: "pointer", border: "none",
                background: viewMode === "directory" ? "#1e293b" : "#fff",
                color: viewMode === "directory" ? "#fff" : "#475569",
              }}
            >
              <LayoutList size={14} /> Directory
            </button>
          </div>
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
                if (homeScreen.viewMode === "directory") setDirSearch(homeScreen.dirFilter ?? "");
              } else {
                setViewMode("hierarchy");
              }
            }}
          >
            <Home size={18} />
          </Button>
          <Button
            type="button"
            variant="outline"
            className="btn-icon"
            aria-label={viewMode === "hierarchy" ? "Print" : "Print Entity Book"}
            title={viewMode === "hierarchy" ? "Print" : "Print Entity Book"}
            disabled={isPdfExporting}
            onClick={async () => {
              if (viewMode === "hierarchy") {
                if (isPdfExporting) return;
                const focusNode = nodeList.find((n) => n.id === focusId);
                const pendingFileName = `${focusNode?.name || focusId}.pdf`
                  .replace(/[^\w\s.-]/g, "").replace(/\s+/g, "_");
                setExportResultAndRevoke({ status: "exporting", fileName: pendingFileName });
                setIsPdfExporting(true);
                pdfCancelRef.current = false;
                setPdfProgress(null);
                try {
                  const result = await generateEntityPdf({
                    nodeId: focusId,
                    nodeList,
                    relList,
                    dataDictionary,
                    clientName: clientDisplayName || toSentenceCase(clientId),
                    isCancelled: () => pdfCancelRef.current,
                    onProgress: (current, total) => setPdfProgress({ current, total }),
                    apiBase,
                    token,
                  });
                  if (result?.url) setExportResultAndRevoke({ status: "ready", url: result.url, fileName: result.fileName });
                } finally {
                  setIsPdfExporting(false);
                  setPdfProgress(null);
                }
              } else {
                setPrintDialogOpen(true);
              }
            }}
          >
            {isPdfExporting ? <Loader2 size={18} className="animate-spin" /> : <Printer size={18} />}
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
                  onClick={() => {
                    const screen = {
                      viewMode,
                      focusId: viewMode === "hierarchy" ? focusId : null,
                      dirFilter: viewMode === "directory" ? dirSearch : null,
                    };
                    setHomeScreen(screen);
                    try { localStorage.setItem("homeScreen", JSON.stringify(screen)); } catch {}
                    // Persist to user record so it survives across sessions/devices
                    apiRequest("/api/auth/me", {
                      method: "PATCH",
                      body: JSON.stringify({ homeScreen: screen }),
                    }).catch(() => {});
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
                  onClick={async () => {
                    setSettingsOpen(false);
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
                  Export Database
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
                  onClick={async () => {
                    setSettingsOpen(false);
                    try {
                      const data = await apiRequest("/api/auth/me");
                      setAccountInfoDraft({
                        name: data.name || "",
                        email: data.email || "",
                        cellPhone: data.cellPhone || "",
                        workPhone: data.workPhone || "",
                      });
                    } catch {
                      setAccountInfoDraft({ name: "", email: "", cellPhone: "", workPhone: "" });
                    }
                    setAccountInfoError("");
                    setAccountInfoOpen(true);
                  }}
                >
                  <User size={15} />
                  Account Info
                </button>
                {myRole === "admin" && (
                  <button
                    className="settings-menu-item"
                    onClick={async () => {
                      setSettingsOpen(false);
                      try {
                        const data = await apiRequest("/api/client");
                        setClientInfoDraft({
                          clientName: data.clientName || "",
                          address: data.address || "",
                          billingContact: data.billingContact || "",
                          billingEmail: data.billingEmail || "",
                          billingPhone: data.billingPhone || "",
                          notes: data.notes || "",
                        });
                      } catch {
                        setClientInfoDraft({ clientName: "", address: "", billingContact: "", billingEmail: "", billingPhone: "", notes: "" });
                      }
                      setClientInfoError("");
                      setClientInfoOpen(true);
                    }}
                  >
                    <Building2 size={15} />
                    Client Info
                  </button>
                )}
                {myRole === "admin" && (
                  <button
                    className="settings-menu-item"
                    onClick={async () => {
                      setSettingsOpen(false);
                      setAddUserDraft({ loginId: "", password: "", confirm: "" });
                      setAddUserError("");
                      setAddUserSuccess("");
                      setManageUsersError("");
                      setManageUsersOpen(true);
                      setManageUsersLoading(true);
                      try {
                        const data = await apiRequest("/api/auth/users");
                        setManageUsersList(Array.isArray(data) ? data : []);
                      } catch (err) {
                        setManageUsersError(err.message);
                      } finally {
                        setManageUsersLoading(false);
                      }
                    }}
                  >
                    <UserPlus size={15} />
                    Manage Users
                  </button>
                )}
                <div className="settings-menu-divider" />
                <button
                  className="settings-menu-item"
                  onClick={async () => {
                    setSettingsOpen(false);
                    setServerInfoData(null);
                    setServerInfoOpen(true);
                    setServerInfoBusy(true);
                    try {
                      const data = await apiRequest("/api/health");
                      setServerInfoData(data);
                    } catch (err) {
                      setServerInfoData({ error: err.message });
                    } finally {
                      setServerInfoBusy(false);
                    }
                  }}
                >
                  <Settings size={15} />
                  Server Info
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
                <div style={{ maxHeight: 140, overflowY: "auto", overscrollBehavior: "contain", border: "1px solid #e5e7eb", borderRadius: 8, padding: 8 }}>
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
        <div className="hierarchy-vertical" ref={hierarchyContainerRef}>

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
            {(() => {
              const ownerCount = focusNode?.kind !== "person" ? getOwnersOf(relList, focusId).length : 0;
              const ownedCount = getOwnedBy(relList, focusId).length;
              return (ownerCount > 0 || ownedCount > 0) ? (
                <div className="hv-focus-counts">
                  {ownerCount > 0 && (
                    <span>{ownerCount} {ownerCount === 1 ? "Owner" : "Owners"}</span>
                  )}
                  {ownedCount > 0 && (
                    <span>Owns {ownedCount} {ownedCount === 1 ? "entity" : "entities"}</span>
                  )}
                </div>
              ) : null;
            })()}
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
          <StatsStrip
            allEntityNodes={sortedEntityNodes}
            allPersonNodes={sortedPersonNodes}
            filteredEntityNodes={filteredEntityNodes}
            filteredPersonNodes={filteredPersonNodes}
            dataDictionary={dataDictionary}
            isFiltered={!!dirSearchLower}
          />
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
                          setFocusId(n.id);
                          setViewMode("hierarchy");
                        }}
                      >
                        {n.logo
                          ? <img src={n.logo} alt="" className="directory-thumb" />
                          : <Building2 className="directory-icon" />}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div className="directory-name">{n.name}</div>
                          <div className="directory-meta">{n.id}</div>
                        </div>
                        <button
                          className="directory-edit-btn"
                          title="Edit"
                          onClick={(e) => {
                            e.stopPropagation();
                            setEditNodeId(n.id);
                            setOpenDialog({ type: "edit-node" });
                          }}
                        >
                          <Pencil size={13} />
                        </button>
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
                          setFocusId(n.id);
                          setViewMode("hierarchy");
                        }}
                      >
                        {n.photo
                          ? <img src={n.photo} alt="" className="directory-thumb directory-thumb--round" />
                          : <Users className="directory-icon" />}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div className="directory-name">{n.name}</div>
                          <div className="directory-meta">{n.id}</div>
                        </div>
                        <button
                          className="directory-edit-btn"
                          title="Edit"
                          onClick={(e) => {
                            e.stopPropagation();
                            setEditNodeId(n.id);
                            setOpenDialog({ type: "edit-node" });
                          }}
                        >
                          <Pencil size={13} />
                        </button>
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
            setNewNode({ name: "", kind: "entity", photo: "", logo: "", customFields: {} });
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
            setNewNode({ name: "", kind: "person", photo: "", logo: "", customFields: {} });
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
              <DialogTitle>Data Dictionary — {clientDisplayName || toSentenceCase(clientId)}</DialogTitle>
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
                        <th>Stats</th>
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
                      {/* Photo / Logo — built-in image field, not editable or moveable */}
                      <tr>
                        <td><em style={{ color: "#6b7280" }}>Photo / Logo</em></td>
                        <td style={{ color: "#6b7280" }}>File / Image</td>
                        <td style={{ color: "#6b7280" }}>Both</td>
                        <td style={{ color: "#6b7280" }}>No</td>
                        <td style={{ color: "#6b7280" }}></td>
                        <td></td>
                      </tr>
                      {/* Operational Role — built-in entity field */}
                      <tr>
                        <td><em style={{ color: "#6b7280" }}>Operational Role</em></td>
                        <td style={{ color: "#6b7280" }}>Dropdown</td>
                        <td style={{ color: "#6b7280" }}>Entity</td>
                        <td style={{ color: "#6b7280" }}>No</td>
                        <td style={{ color: "#6b7280" }}>Active, Passive, Mixed</td>
                        <td style={{ color: "#6b7280" }}>✓</td>
                        <td></td>
                      </tr>
                      {/* Legal Status — built-in entity field */}
                      <tr>
                        <td><em style={{ color: "#6b7280" }}>Legal Status</em></td>
                        <td style={{ color: "#6b7280" }}>Dropdown</td>
                        <td style={{ color: "#6b7280" }}>Entity</td>
                        <td style={{ color: "#6b7280" }}>No</td>
                        <td style={{ color: "#6b7280" }}>Good Standing, Dormant, Dissolved, Suspended</td>
                        <td style={{ color: "#6b7280" }}>✓</td>
                        <td></td>
                      </tr>
                      {/* Status — built-in person field */}
                      <tr>
                        <td><em style={{ color: "#6b7280" }}>Status</em></td>
                        <td style={{ color: "#6b7280" }}>Dropdown</td>
                        <td style={{ color: "#6b7280" }}>Person</td>
                        <td style={{ color: "#6b7280" }}>No</td>
                        <td style={{ color: "#6b7280" }}>Active, Inactive, Deceased, Former</td>
                        <td style={{ color: "#6b7280" }}>✓</td>
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
                            {entry.dataType === "link"
                              ? <span style={{ color: "#9ca3af" }}>URL</span>
                              : entry.dataType === "file"
                              ? <span style={{ color: "#9ca3af" }}>File / Image</span>
                              : (entry.validValues || []).length > 0
                                ? entry.validValues.join(", ")
                                : <span style={{ color: "#9ca3af" }}>Free-form</span>}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {entry.showInStats ? "✓" : ""}
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
              {ddEntryDraft.dataType !== "link" && ddEntryDraft.dataType !== "file" && (
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
              )}
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
              ) : ddEntryDraft.dataType !== "link" && ddEntryDraft.dataType !== "file" ? (
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
              ) : null}
              {/* Show in Stats — only available when field has fixed valid values */}
              {(ddEntryDraft.validValuesText.trim().split("\n").filter(Boolean).length > 0) && (
                <div className="form-row">
                  <label className="form-label">Show in statistics</label>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <Switch
                      checked={ddEntryDraft.showInStats}
                      onCheckedChange={(v) => setDdEntryDraft((prev) => ({ ...prev, showInStats: v }))}
                    />
                    <span style={{ fontSize: 14, color: ddEntryDraft.showInStats ? "#374151" : "#6b7280" }}>
                      {ddEntryDraft.showInStats ? "Counts shown in directory stats bar" : "Not shown in stats"}
                    </span>
                  </div>
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
              {/* Photo / Logo — built-in system field */}
              <NodeImageField
                kind={newNode.kind}
                value={newNode.kind === "person" ? newNode.photo : newNode.logo}
                onChange={(url) => setNewNode((prev) => ({ ...prev, [prev.kind === "person" ? "photo" : "logo"]: url }))}
                apiBase={apiBase}
                token={token}
              />
              {/* Built-in status fields */}
              {newNode.kind === "entity" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Operational Role</label>
                    <select
                      className="form-input"
                      value={newNode.operationalRole}
                      onChange={(e) => setNewNode((prev) => ({ ...prev, operationalRole: e.target.value }))}
                    >
                      <option value="">— select —</option>
                      <option value="Active">Active</option>
                      <option value="Passive">Passive</option>
                      <option value="Mixed">Mixed</option>
                    </select>
                  </div>
                  <div className="form-row">
                    <label className="form-label">Legal Status</label>
                    <select
                      className="form-input"
                      value={newNode.legalStatus}
                      onChange={(e) => setNewNode((prev) => ({ ...prev, legalStatus: e.target.value }))}
                    >
                      <option value="">— select —</option>
                      <option value="Good Standing">Good Standing</option>
                      <option value="Dormant">Dormant</option>
                      <option value="Dissolved">Dissolved</option>
                      <option value="Suspended">Suspended</option>
                    </select>
                  </div>
                </>
              )}
              {newNode.kind === "person" && (
                <div className="form-row">
                  <label className="form-label">Status</label>
                  <select
                    className="form-input"
                    value={newNode.personStatus}
                    onChange={(e) => setNewNode((prev) => ({ ...prev, personStatus: e.target.value }))}
                  >
                    <option value="">— select —</option>
                    <option value="Active">Active</option>
                    <option value="Inactive">Inactive</option>
                    <option value="Deceased">Deceased</option>
                    <option value="Former">Former</option>
                  </select>
                </div>
              )}
              {[...dataDictionary]
                .filter((f) => f.appliesTo === "both" || f.appliesTo === newNode.kind)
                .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                .map((field) => (
                  <React.Fragment key={field.fieldId}>
                    {renderDdField(
                      field,
                      newNode.customFields?.[field.fieldId],
                      (val) => setNewNode((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } })),
                      { apiBase, token }
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
                    photo: newNode.kind === "person" ? (newNode.photo || "") : "",
                    logo: newNode.kind === "entity" ? (newNode.logo || "") : "",
                    operationalRole: newNode.kind === "entity" ? (newNode.operationalRole || "") : "",
                    legalStatus: newNode.kind === "entity" ? (newNode.legalStatus || "") : "",
                    personStatus: newNode.kind === "person" ? (newNode.personStatus || "") : "",
                    customFields: newNode.customFields || {},
                  };
                  try {
                    await apiRequest("/api/nodes", {
                      method: "POST",
                      body: JSON.stringify(payload),
                    });
                    setNodeList((prev) => [...prev, payload]);
                    if (!focusId) setFocusId(id);
                    setNewNode({ name: "", kind: newNode.kind, photo: "", logo: "", operationalRole: "", legalStatus: "", personStatus: "", customFields: {} });
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
              {/* Photo / Logo — built-in system field */}
              <NodeImageField
                kind={nodeDraft.kind}
                value={nodeDraft.kind === "person" ? nodeDraft.photo : nodeDraft.logo}
                onChange={(url) => setNodeDraft((prev) => ({ ...prev, [prev.kind === "person" ? "photo" : "logo"]: url }))}
                apiBase={apiBase}
                token={token}
              />
              {/* Built-in status fields */}
              {nodeDraft.kind === "entity" && (
                <>
                  <div className="form-row">
                    <label className="form-label">Operational Role</label>
                    <select
                      className="form-input"
                      value={nodeDraft.operationalRole}
                      onChange={(e) => setNodeDraft((prev) => ({ ...prev, operationalRole: e.target.value }))}
                    >
                      <option value="">— select —</option>
                      <option value="Active">Active</option>
                      <option value="Passive">Passive</option>
                      <option value="Mixed">Mixed</option>
                    </select>
                  </div>
                  <div className="form-row">
                    <label className="form-label">Legal Status</label>
                    <select
                      className="form-input"
                      value={nodeDraft.legalStatus}
                      onChange={(e) => setNodeDraft((prev) => ({ ...prev, legalStatus: e.target.value }))}
                    >
                      <option value="">— select —</option>
                      <option value="Good Standing">Good Standing</option>
                      <option value="Dormant">Dormant</option>
                      <option value="Dissolved">Dissolved</option>
                      <option value="Suspended">Suspended</option>
                    </select>
                  </div>
                </>
              )}
              {nodeDraft.kind === "person" && (
                <div className="form-row">
                  <label className="form-label">Status</label>
                  <select
                    className="form-input"
                    value={nodeDraft.personStatus}
                    onChange={(e) => setNodeDraft((prev) => ({ ...prev, personStatus: e.target.value }))}
                  >
                    <option value="">— select —</option>
                    <option value="Active">Active</option>
                    <option value="Inactive">Inactive</option>
                    <option value="Deceased">Deceased</option>
                    <option value="Former">Former</option>
                  </select>
                </div>
              )}
              {[...dataDictionary]
                .filter((f) => f.appliesTo === "both" || f.appliesTo === nodeDraft.kind)
                .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                .map((field) => (
                  <React.Fragment key={field.fieldId}>
                    {renderDdField(
                      field,
                      nodeDraft.customFields?.[field.fieldId],
                      (val) => setNodeDraft((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } })),
                      { apiBase, token }
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
                    const editNode = nodeList.find((n) => n.id === editNodeId);
                    const pendingFileName = `${editNode?.name || editNodeId}.pdf`
                      .replace(/[^\w\s.-]/g, "").replace(/\s+/g, "_");
                    setExportResultAndRevoke({ status: "exporting", fileName: pendingFileName });
                    setIsPdfExporting(true);
                    pdfCancelRef.current = false;
                    setPdfProgress(null);
                    try {
                      const result = await generateEntityPdf({
                        nodeId: editNodeId,
                        nodeList,
                        relList,
                        dataDictionary,
                        clientName: clientDisplayName || toSentenceCase(clientId),
                        isCancelled: () => pdfCancelRef.current,
                        onProgress: (current, total) => setPdfProgress({ current, total }),
                        apiBase,
                        token,
                      });
                      if (result?.url) setExportResultAndRevoke({ status: "ready", url: result.url, fileName: result.fileName });
                    } finally {
                      setIsPdfExporting(false);
                      setPdfProgress(null);
                    }
                  }}
                >
                  Print
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
                      photo: nodeDraft.kind === "person" ? (nodeDraft.photo || "") : "",
                      logo: nodeDraft.kind === "entity" ? (nodeDraft.logo || "") : "",
                      operationalRole: nodeDraft.kind === "entity" ? (nodeDraft.operationalRole || "") : "",
                      legalStatus: nodeDraft.kind === "entity" ? (nodeDraft.legalStatus || "") : "",
                      personStatus: nodeDraft.kind === "person" ? (nodeDraft.personStatus || "") : "",
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
                                  photo: payload.photo,
                                  logo: payload.logo,
                                  operationalRole: payload.operationalRole,
                                  legalStatus: payload.legalStatus,
                                  personStatus: payload.personStatus,
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

      {/* ── Print Entity Book dialog (directory mode) ── */}
      <Dialog open={printDialogOpen} onOpenChange={setPrintDialogOpen}>
        <DialogContent style={{ maxWidth: 360 }}>
          <DialogHeader>
            <DialogTitle>Print Entity Book</DialogTitle>
          </DialogHeader>
          <div style={{ padding: "8px 0 16px", display: "flex", flexDirection: "column", gap: 12 }}>
            <p style={{ fontSize: 13, color: "#64748b", margin: 0 }}>
              Select page types to include for each entity in the current view.
              Pages are interleaved: hierarchy then detail for each entity.
            </p>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={printHierarchy}
                onChange={(e) => setPrintHierarchy(e.target.checked)}
              />
              Hierarchy pages
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={printDetail}
                onChange={(e) => setPrintDetail(e.target.checked)}
              />
              Detail pages
            </label>
            {(printHierarchy || printDetail) && (() => {
              const scopeNodes = dirSearch.trim()
                ? [...filteredEntityNodes, ...filteredPersonNodes]
                : nodeList;
              return (
                <p style={{ fontSize: 12, color: "#94a3b8", margin: 0 }}>
                  {scopeNodes.length} {scopeNodes.length === 1 ? "item" : "items"} in scope
                </p>
              );
            })()}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setPrintDialogOpen(false)}>Cancel</Button>
            <Button
              disabled={(!printHierarchy && !printDetail) || isPdfExporting}
              onClick={async () => {
                setPrintDialogOpen(false);
                setIsPdfExporting(true);
                const suffix = printHierarchy && printDetail ? "full" : printHierarchy ? "hierarchy" : "detail";
                const safeBase = `${clientDisplayName || toSentenceCase(clientId)}-entity-book-${suffix}`
                  .replace(/[^\w\s.-]/g, "").replace(/\s+/g, "_");
                setExportResultAndRevoke({ status: "exporting", fileName: `${safeBase}.pdf` });
                pdfCancelRef.current = false;
                setPdfProgress(null);
                try {
                  const printNodes = dirSearch.trim()
                    ? [...filteredEntityNodes, ...filteredPersonNodes]
                    : nodeList;
                  const result = await generateEntityBookInterleaved({
                    nodes: printNodes,
                    nodeList,
                    relList,
                    dataDictionary,
                    clientName: clientDisplayName || toSentenceCase(clientId),
                    includeHierarchy: printHierarchy,
                    includeDetail: printDetail,
                    isCancelled: () => pdfCancelRef.current,
                    onProgress: (current, total) => setPdfProgress({ current, total }),
                    fileName: safeBase,
                    apiBase,
                    token,
                  });
                  if (result?.url) setExportResultAndRevoke({ status: "ready", url: result.url, fileName: result.fileName });
                } finally {
                  setIsPdfExporting(false);
                  setPdfProgress(null);
                }
              }}
            >
              {isPdfExporting
                ? <><Loader2 size={14} className="animate-spin" style={{ marginRight: 6 }} />Printing…</>
                : "Print"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <ExportDialog
        open={exportOpen}
        onClose={() => setExportOpen(false)}
        onExported={(result) => setExportResultAndRevoke({ status: "ready", url: result.url, fileName: result.fileName })}
        onExportStart={({ fileName }) => setExportResultAndRevoke({ status: "exporting", fileName })}
        exportNodes={nodeList}
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
        clientName={clientDisplayName || toSentenceCase(clientId)}
      />

      {/* ── Export progress modal ── */}
      <Dialog
        open={!!exportResult}
        onOpenChange={(v) => {
          if (!v && exportResult?.status === "ready") {
            if (exportResult.url) URL.revokeObjectURL(exportResult.url);
            exportResultRef.current = null;
            setExportResult(null);
          }
        }}
      >
        <DialogContent style={{ maxWidth: 400 }}>
          <DialogHeader style={{ marginBottom: 12, marginLeft: 0 }}>
            <DialogTitle>
              {exportResult?.status === "ready" ? "Document ready" : "Printing…"}
            </DialogTitle>
          </DialogHeader>

          {exportResult?.status !== "ready" ? (
            /* ── In-progress ── */
            <>
              {!pdfProgress ? (
                <p style={{ fontSize: 13, color: "#64748b", margin: 0 }}>Preparing…</p>
              ) : (
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 4 }}>
                    <span>Entity {pdfProgress.current} of {pdfProgress.total}</span>
                    <span>{Math.round((pdfProgress.current / pdfProgress.total) * 100)}%</span>
                  </div>
                  <div style={{ height: 6, borderRadius: 3, background: "#e2e8f0", overflow: "hidden" }}>
                    <div style={{
                      height: "100%", borderRadius: 3, background: "#1e293b",
                      width: `${(pdfProgress.current / pdfProgress.total) * 100}%`,
                      transition: "width 0.2s ease",
                    }} />
                  </div>
                </div>
              )}
              <DialogFooter style={{ marginTop: 16 }}>
                <Button
                  variant="outline"
                  onClick={() => {
                    pdfCancelRef.current = true;
                    setExportResultAndRevoke(null);
                    setIsPdfExporting(false);
                  }}
                >
                  Cancel
                </Button>
              </DialogFooter>
            </>
          ) : (
            /* ── Ready state ── */
            <>
              <DialogFooter style={{ gap: 8 }}>
                <Button
                  variant="outline"
                  onClick={() => {
                    if (exportResult.url) URL.revokeObjectURL(exportResult.url);
                    exportResultRef.current = null;
                    setExportResult(null);
                  }}
                >
                  Close
                </Button>
                <Button
                  variant="outline"
                  onClick={() => window.open(exportResult.url)}
                >
                  View
                </Button>
                <Button
                  onClick={() => {
                    const a = document.createElement("a");
                    a.href = exportResult.url;
                    a.download = exportResult.fileName;
                    a.click();
                  }}
                >
                  Download
                </Button>
              </DialogFooter>
            </>
          )}
        </DialogContent>
      </Dialog>

      {/* ── Account Info dialog ── */}
      <Dialog open={accountInfoOpen} onOpenChange={(v) => { if (!v) setAccountInfoOpen(false); }}>
        <DialogContent style={{ maxWidth: 440 }}>
          <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
            <DialogTitle>Account Info</DialogTitle>
          </DialogHeader>
          <div className="form-grid">
            <div className="form-row">
              <label className="form-label">Name</label>
              <input
                className="form-input"
                type="text"
                value={accountInfoDraft.name}
                onChange={(e) => setAccountInfoDraft((p) => ({ ...p, name: e.target.value }))}
                autoComplete="name"
              />
            </div>
            <div className="form-row">
              <label className="form-label">E-Mail</label>
              <input
                className="form-input"
                type="email"
                value={accountInfoDraft.email}
                onChange={(e) => setAccountInfoDraft((p) => ({ ...p, email: e.target.value }))}
                autoComplete="email"
              />
            </div>
            <div className="form-row">
              <label className="form-label">Cell Phone</label>
              <input
                className="form-input"
                type="tel"
                value={accountInfoDraft.cellPhone}
                onChange={(e) => setAccountInfoDraft((p) => ({ ...p, cellPhone: e.target.value }))}
                autoComplete="tel"
              />
            </div>
            <div className="form-row">
              <label className="form-label">Work Phone</label>
              <input
                className="form-input"
                type="tel"
                value={accountInfoDraft.workPhone}
                onChange={(e) => setAccountInfoDraft((p) => ({ ...p, workPhone: e.target.value }))}
                autoComplete="tel"
              />
            </div>
          </div>
          {accountInfoError && (
            <div style={{ color: "#dc2626", fontSize: 13, marginTop: 8 }}>{accountInfoError}</div>
          )}
          <DialogFooter style={{ marginTop: 24 }}>
            <Button variant="secondary" onClick={() => setAccountInfoOpen(false)}>Cancel</Button>
            <Button
              type="button"
              disabled={accountInfoBusy}
              onClick={async () => {
                setAccountInfoError("");
                setAccountInfoBusy(true);
                try {
                  await apiRequest("/api/auth/me", {
                    method: "PATCH",
                    body: JSON.stringify(accountInfoDraft),
                  });
                  setAccountInfoOpen(false);
                } catch (err) {
                  setAccountInfoError(err.message);
                } finally {
                  setAccountInfoBusy(false);
                }
              }}
            >
              {accountInfoBusy ? "Saving…" : "Save"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* ── Manage Users dialog (admin only) ── */}
      <Dialog open={manageUsersOpen} onOpenChange={(v) => { if (!v) setManageUsersOpen(false); }}>
        <DialogContent style={{ width: "min(560px, 95vw)", maxWidth: "none" }}>
          <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
            <DialogTitle>Manage Users</DialogTitle>
          </DialogHeader>

          {/* ─ User list ─ */}
          {manageUsersLoading && <div style={{ color: "#6b7280", fontSize: 14 }}>Loading…</div>}
          {!manageUsersLoading && manageUsersError && (
            <div style={{ color: "#dc2626", fontSize: 13, marginBottom: 12 }}>{manageUsersError}</div>
          )}
          {!manageUsersLoading && manageUsersList.length > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 14 }}>
              <thead>
                <tr style={{ borderBottom: "2px solid #e5e7eb" }}>
                  <th style={{ textAlign: "left", padding: "4px 8px 6px 0", color: "#374151", fontWeight: 600 }}>Login ID</th>
                  <th style={{ textAlign: "left", padding: "4px 8px 6px", color: "#374151", fontWeight: 600 }}>Role</th>
                  <th style={{ padding: "4px 0 6px" }}></th>
                </tr>
              </thead>
              <tbody>
                {manageUsersList.map((u) => {
                  const isSelf = u.loginId === myLoginId;
                  const isUpdating = manageUsersUpdating.has(u.loginId);
                  return (
                    <tr key={u.loginId} style={{ borderBottom: "1px solid #f1f5f9" }}>
                      <td style={{ padding: "7px 8px 7px 0", color: isSelf ? "#6b7280" : "#111827" }}>
                        {u.loginId}{isSelf && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 6 }}>(you)</span>}
                      </td>
                      <td style={{ padding: "7px 8px", color: u.role === "admin" ? "#7c3aed" : "#6b7280" }}>
                        {u.role === "admin" ? "Admin" : "User"}
                      </td>
                      <td style={{ padding: "7px 0", textAlign: "right" }}>
                        {!isSelf && (
                          <button
                            style={{
                              fontSize: 12, padding: "3px 10px", borderRadius: 4, cursor: isUpdating ? "default" : "pointer",
                              border: "1px solid #d1d5db", background: isUpdating ? "#f9fafb" : "#fff",
                              color: isUpdating ? "#9ca3af" : "#374151",
                            }}
                            disabled={isUpdating}
                            onClick={async () => {
                              const newRole = u.role === "admin" ? "user" : "admin";
                              setManageUsersUpdating((prev) => new Set([...prev, u.loginId]));
                              try {
                                await apiRequest(`/api/auth/users/${encodeURIComponent(u.loginId)}`, {
                                  method: "PATCH",
                                  body: JSON.stringify({ role: newRole }),
                                });
                                setManageUsersList((prev) =>
                                  prev.map((x) => x.loginId === u.loginId ? { ...x, role: newRole } : x)
                                );
                              } catch (err) {
                                setManageUsersError(err.message);
                              } finally {
                                setManageUsersUpdating((prev) => { const s = new Set(prev); s.delete(u.loginId); return s; });
                              }
                            }}
                          >
                            {isUpdating ? "Saving…" : u.role === "admin" ? "Make User" : "Make Admin"}
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}

          {/* ─ Add user form ─ */}
          <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 16 }}>
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12, color: "#374151" }}>Add New User</div>
            {addUserSuccess ? (
              <div style={{ marginBottom: 12 }}>
                <span style={{ color: "#16a34a", fontWeight: 600 }}>Created: </span>
                <span style={{ fontSize: 13 }}>{addUserSuccess}</span>
                <button
                  style={{ marginLeft: 12, fontSize: 12, color: "#6b7280", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}
                  onClick={() => { setAddUserSuccess(""); setAddUserDraft({ loginId: "", password: "", confirm: "" }); }}
                >Add another</button>
              </div>
            ) : (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                  <div style={{ gridColumn: "1 / -1" }}>
                    <label className="form-label" style={{ fontSize: 12 }}>Login ID</label>
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
                  <div>
                    <label className="form-label" style={{ fontSize: 12 }}>Password</label>
                    <input
                      className="form-input"
                      type="password"
                      value={addUserDraft.password}
                      onChange={(e) => setAddUserDraft((p) => ({ ...p, password: e.target.value }))}
                      autoComplete="new-password"
                    />
                  </div>
                  <div>
                    <label className="form-label" style={{ fontSize: 12 }}>Confirm</label>
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
                  <div style={{ color: "#dc2626", fontSize: 13, marginBottom: 8 }}>{addUserError}</div>
                )}
                <Button
                  type="button"
                  disabled={addUserBusy}
                  onClick={async () => {
                    const { loginId, password, confirm } = addUserDraft;
                    if (!loginId.trim()) { setAddUserError("Login ID is required."); return; }
                    if (!password) { setAddUserError("Password is required."); return; }
                    if (password !== confirm) { setAddUserError("Passwords do not match."); return; }
                    setAddUserError("");
                    setAddUserBusy(true);
                    try {
                      const data = await apiRequest("/api/auth/users", {
                        method: "POST",
                        body: JSON.stringify({ loginId: loginId.trim().toLowerCase(), password }),
                      });
                      setAddUserSuccess(data.loginId);
                      setManageUsersList((prev) => [...prev, { loginId: data.loginId, role: data.role || "user" }]);
                    } catch (err) {
                      setAddUserError(err.message);
                    } finally {
                      setAddUserBusy(false);
                    }
                  }}
                >
                  {addUserBusy ? "Creating…" : "Create User"}
                </Button>
              </>
            )}
          </div>

          <DialogFooter style={{ marginTop: 20 }}>
            <Button variant="secondary" onClick={() => setManageUsersOpen(false)}>Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* ── Client Info dialog (admin only) ── */}
      <Dialog open={clientInfoOpen} onOpenChange={(v) => { if (!v) setClientInfoOpen(false); }}>
        <DialogContent className="dialog-content--tall" style={{ width: "60vw", maxWidth: "none", minWidth: 360 }}>
          <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
            <DialogTitle>Client Info</DialogTitle>
          </DialogHeader>
          <div className="dialog-body">
          <div className="form-grid">
            <div className="form-row">
              <label className="form-label">Client ID</label>
              <div style={{ fontSize: 14, color: "#6b7280", padding: "6px 0" }}>{clientId}</div>
            </div>
            <div className="form-row">
              <label className="form-label">Client Name</label>
              <input
                className="form-input"
                type="text"
                value={clientInfoDraft.clientName}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, clientName: e.target.value }))}
              />
            </div>
            <div className="form-row">
              <label className="form-label">Address</label>
              <textarea
                className="form-input"
                style={{ minHeight: 64, resize: "vertical", fontFamily: "inherit" }}
                value={clientInfoDraft.address}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, address: e.target.value }))}
              />
            </div>
            <div className="form-row">
              <label className="form-label">Billing Contact</label>
              <input
                className="form-input"
                type="text"
                value={clientInfoDraft.billingContact}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, billingContact: e.target.value }))}
              />
            </div>
            <div className="form-row">
              <label className="form-label">Billing Email</label>
              <input
                className="form-input"
                type="email"
                value={clientInfoDraft.billingEmail}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, billingEmail: e.target.value }))}
                autoComplete="email"
              />
            </div>
            <div className="form-row">
              <label className="form-label">Billing Phone</label>
              <input
                className="form-input"
                type="tel"
                value={clientInfoDraft.billingPhone}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, billingPhone: e.target.value }))}
                autoComplete="tel"
              />
            </div>
            <div className="form-row">
              <label className="form-label">Notes</label>
              <textarea
                className="form-input"
                style={{ minHeight: 64, resize: "vertical", fontFamily: "inherit" }}
                value={clientInfoDraft.notes}
                onChange={(e) => setClientInfoDraft((p) => ({ ...p, notes: e.target.value }))}
              />
            </div>
          </div>
          {clientInfoError && (
            <div style={{ color: "#dc2626", fontSize: 13, marginTop: 8 }}>{clientInfoError}</div>
          )}
          </div>
          <DialogFooter>
            <Button variant="secondary" onClick={() => setClientInfoOpen(false)}>Cancel</Button>
            <Button
              type="button"
              disabled={clientInfoBusy}
              onClick={async () => {
                setClientInfoError("");
                setClientInfoBusy(true);
                try {
                  await apiRequest("/api/client", {
                    method: "PATCH",
                    body: JSON.stringify(clientInfoDraft),
                  });
                  setClientInfoOpen(false);
                } catch (err) {
                  setClientInfoError(err.message);
                } finally {
                  setClientInfoBusy(false);
                }
              }}
            >
              {clientInfoBusy ? "Saving…" : "Save"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={serverInfoOpen} onOpenChange={(v) => { if (!v) setServerInfoOpen(false); }}>
        <DialogContent style={{ width: "min(480px, 92vw)", maxWidth: "none" }}>
          <DialogHeader style={{ marginBottom: 16 }}>
            <DialogTitle>Server Info</DialogTitle>
          </DialogHeader>
          <div style={{ fontSize: 14 }}>
            {serverInfoBusy && <div style={{ color: "#6b7280" }}>Checking…</div>}
            {!serverInfoBusy && serverInfoData && (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <tbody>
                  {Object.entries(serverInfoData).map(([k, v]) => (
                    <tr key={k} style={{ borderBottom: "1px solid #f1f5f9" }}>
                      <td style={{ padding: "6px 12px 6px 0", fontWeight: 600, color: "#374151", whiteSpace: "nowrap", verticalAlign: "top" }}>{k}</td>
                      <td style={{ padding: "6px 0", color: "#111827", wordBreak: "break-all" }}>
                        {Array.isArray(v) ? v.join(", ") : String(v)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
          <DialogFooter style={{ marginTop: 20 }}>
            <Button variant="secondary" onClick={() => setServerInfoOpen(false)}>Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      </div>{/* end app-content */}
    </div>
  );
}

// EOF
