import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Card, CardContent } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { Link, Users, Building2, Plus, Pencil, Trash2, ChevronRight, ChevronDown, ChevronsDown, ChevronsUp, Upload, X, Search, Settings, LogOut, GitFork, LayoutList, Home, Download, BookOpen, User, UserPlus, Loader2, Crosshair } from "lucide-react";
import { generateEntityPdf, generateEntityBook, generateEntityBookInterleaved, estimatePosterPageCount, generateOrgChartPoster } from "./utils/generateEntityPdf";
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
  { value: "string", label: "Short text" },
  { value: "textarea", label: "Long text" },
  { value: "number", label: "Number" },
  { value: "currency", label: "Currency" },
  { value: "percentage", label: "Percentage" },
  { value: "boolean", label: "Yes / No" },
  { value: "date", label: "Date" },
  { value: "time", label: "Time" },
  { value: "phone", label: "Phone number" },
  { value: "email", label: "Email address" },
  { value: "link", label: "URL / Link" },
  { value: "file", label: "File / Image" },
  { value: "address", label: "Address" },
  { value: "year", label: "Year" },
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

// ── Collect all descendant nodeIds, cycle-safe ─────────────────────────────────────────
const getAllDescendants = (relList, rootId, initialVisited = new Set()) => {
  const result = new Set();
  const stack = [{ id: rootId, seen: new Set(initialVisited) }];
  while (stack.length) {
    const { id, seen } = stack.pop();
    for (const child of getOwnedBy(relList, id)) {
      if (seen.has(child.nodeId)) continue;
      result.add(child.nodeId);
      const nextSeen = new Set(seen);
      nextSeen.add(child.nodeId);
      stack.push({ id: child.nodeId, seen: nextSeen });
    }
  }
  return result;
};

// ── HvNeighborBox: a single child tile in the explodable Owns section ────────────────
const HvNeighborBox = ({
  item,
  nodeList,
  relList,
  explodedNodes,
  onExplode,
  onExplodeAll,
  onFocus,
  onFocusPrimary,
  onEdit,
  onPrintBook,
  onPrintPoster,
  showExplodeControls = true,
  isCyclic = false,
}) => {
  const node = getNode(nodeList, item.nodeId);
  if (!node) return null;
  const pct = item.rel?.percent;
  const showPct = pct != null && Number.isFinite(Number(pct));
  const isZero = showPct && Number(pct) === 0;
  const childCount = getOwnedBy(relList, item.nodeId).length;
  const isExploded = explodedNodes.has(item.nodeId);
  const classNames = [
    "hv-neighbor-box",
    isZero ? "hv-neighbor-box--zero" : "",
    isCyclic ? "hv-neighbor-box--cyclic" : "",
  ].filter(Boolean).join(" ");
  return (
    <div
      className={classNames}
      style={{ position: "relative" }}
      data-hv-node-id={item.nodeId}
      onClick={() => onFocus(item.nodeId)}
      title={isCyclic ? "Circular ownership — already shown above" : isZero ? "Non-economic / 0% interest" : undefined}
    >
      {node.kind === "person"
        ? node.photo
          ? <img src={node.photo} alt="" className="hv-neighbor-photo" />
          : <Users size={22} className="hv-neighbor-icon" />
        : node.logo
          ? <img src={node.logo} alt="" className="hv-neighbor-logo" />
          : <Building2 size={22} className="hv-neighbor-icon" />
      }
      <div className="hv-neighbor-name">{node.name}</div>
      {showPct && <div className="hv-neighbor-pct">{Number(pct)}%</div>}
      {isCyclic && <div className="hv-neighbor-cycle-badge" title="Circular reference">∞</div>}
      <div className="hv-neighbor-actions" onClick={(e) => e.stopPropagation()}>
        <button
          className="hv-neighbor-action-btn"
          type="button"
          title="Edit"
          aria-label="Edit"
          onClick={() => onEdit?.(item.nodeId)}
        >
          <Pencil size={11} />
        </button>
        <button
          className="hv-neighbor-action-btn"
          type="button"
          title="Print Book Pages (PDF)"
          aria-label="Print Book Pages"
          onClick={() => onPrintBook?.(item.nodeId)}
        >
          <BookOpen size={11} />
        </button>
        <button
          className="hv-neighbor-action-btn"
          type="button"
          title="Print Org Chart Poster"
          aria-label="Print Org Chart Poster"
          onClick={() => onPrintPoster?.(item.nodeId)}
        >
          <GitFork size={11} />
        </button>
        <button
          className="hv-neighbor-action-btn"
          type="button"
          title="Focus on me"
          aria-label="Focus on me"
          onClick={() => {
            if (onFocusPrimary) {
              onFocusPrimary(item.nodeId);
            } else {
              onFocus(item.nodeId);
            }
          }}
        >
          <Crosshair size={11} />
        </button>
      </div>
      {showExplodeControls && !isCyclic && childCount > 0 && (
        <>
          <button
            className={`hv-explode-btn${isExploded ? " hv-explode-btn--active" : ""}`}
            title={isExploded ? `Collapse ${childCount} direct child${childCount === 1 ? "" : "ren"}` : `Expand ${childCount} direct child${childCount === 1 ? "" : "ren"}`}
            onClick={(e) => {
              e.stopPropagation();
              const anchorEl = e.currentTarget.closest('[data-hv-node-id]');
              onExplode(item.nodeId, anchorEl);
            }}
          >
            <ChevronDown size={12} style={{ transition: "transform 0.2s", transform: isExploded ? "rotate(180deg)" : "none" }} />
          </button>
          {!isExploded && (() => {
            const totalDesc = getAllDescendants(relList, item.nodeId).size;
            return (
              <button
                className="hv-explode-all-btn"
                title={`Expand all — ${totalDesc} node${totalDesc === 1 ? "" : "s"} in tree`}
                onClick={(e) => {
                  e.stopPropagation();
                  const anchorEl = e.currentTarget.closest('[data-hv-node-id]');
                  onExplodeAll(item.nodeId, anchorEl);
                }}
              >
                <ChevronsDown size={12} />
              </button>
            );
          })()}
        </>
      )}
    </div>
  );
};

// ── Org-chart layout constants ────────────────────────────────────────────────
const ORG_NODE_W = 140;      // matches .hv-neighbor-box width in px
const ORG_NODE_GAP = 24;     // horizontal gap between sibling subtree columns
const ORG_V_SEG = 20;        // px — height of each vertical connector segment
const ORG_LINE_COLOR = '#d1d5db';
const DEFAULT_TABULAR_VIEW_ID = "__default__";

// Compute the full pixel width of a subtree column (for precise connector positioning)
function computeColWidth(nodeId, relList, explodedNodes, visited = new Set()) {
  if (visited.has(nodeId)) return ORG_NODE_W;          // cyclic — just the box
  if (!explodedNodes.has(nodeId)) return ORG_NODE_W;   // collapsed
  const children = getOwnedBy(relList, nodeId);
  if (!children.length) return ORG_NODE_W;
  const nv = new Set(visited); nv.add(nodeId);
  const total = children.reduce(
    (sum, c, i) => sum + computeColWidth(c.nodeId, relList, explodedNodes, nv) + (i > 0 ? ORG_NODE_GAP : 0),
    0
  );
  return Math.max(ORG_NODE_W, total);
}

// ── Recursive org-chart tree node ─────────────────────────────────────────────
const OrgChartTreeNode = ({
  item, nodeList, relList, explodedNodes,
  onExplode, onExplodeAll, onFocus, onFocusPrimary, onEdit, onPrintBook, onPrintPoster,
  visitedIds = new Set(), showTopConnector = true,
}) => {
  const { nodeId } = item;
  const isCyclic = visitedIds.has(nodeId);
  const isExploded = !isCyclic && explodedNodes.has(nodeId);
  const children = isExploded ? getOwnedBy(relList, nodeId) : [];
  const nv = new Set(visitedIds); nv.add(nodeId);
  const colW = computeColWidth(nodeId, relList, explodedNodes, visitedIds);

  // Horizontal bar: spans from centre of first child column to centre of last
  let barLeft = 0, barW = 0;
  if (children.length >= 1) {
    const cws = children.map(c => computeColWidth(c.nodeId, relList, explodedNodes, nv));
    barLeft = cws[0] / 2;
    barW = Math.max(0, colW - cws[0] / 2 - cws[cws.length - 1] / 2);
  }

  return (
    <div style={{ width: colW, display: 'flex', flexDirection: 'column', alignItems: 'center', flexShrink: 0 }}>
      {showTopConnector && (
        <div style={{ width: 2, height: ORG_V_SEG, background: ORG_LINE_COLOR, flexShrink: 0 }} />
      )}
      <HvNeighborBox
        item={item} nodeList={nodeList} relList={relList}
        explodedNodes={explodedNodes} onExplode={onExplode}
        onExplodeAll={onExplodeAll} onFocus={onFocus}
        onFocusPrimary={onFocusPrimary}
        onEdit={onEdit} onPrintBook={onPrintBook} onPrintPoster={onPrintPoster}
        isCyclic={isCyclic}
      />
      {children.length > 0 && (
        <>
          {/* Vertical drop from node down to horizontal bar */}
          <div style={{ width: 2, height: ORG_V_SEG, background: ORG_LINE_COLOR, flexShrink: 0 }} />
          {/* Horizontal bar: centre-to-centre across children */}
          <div style={{ position: 'relative', width: colW, height: 2, flexShrink: 0 }}>
            <div style={{ position: 'absolute', top: 0, left: barLeft, width: barW, height: 2, background: ORG_LINE_COLOR }} />
          </div>
          {/* Children row */}
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: ORG_NODE_GAP, flexShrink: 0 }}>
            {children.map(child => (
              <OrgChartTreeNode
                key={child.nodeId} item={child} nodeList={nodeList} relList={relList}
                explodedNodes={explodedNodes} onExplode={onExplode}
                onExplodeAll={onExplodeAll} onFocus={onFocus}
                onFocusPrimary={onFocusPrimary}
                onEdit={onEdit} onPrintBook={onPrintBook} onPrintPoster={onPrintPoster}
                visitedIds={nv} showTopConnector={true}
              />
            ))}
          </div>
        </>
      )}
    </div>
  );
};

// ── Minimap: small birds-eye overview of the full org chart ───────────────────
const OrgChartMinimap = ({ containerRef, watchKey }) => {
  const canvasRef = useRef(null);
  const layoutRef = useRef({ minX: 0, minY: 0, scale: 1, PAD: 8, offsetX: 0, offsetY: 0 });

  const redraw = useCallback(() => {
    const canvas = canvasRef.current;
    const container = containerRef.current;
    if (!canvas || !container) return;
    const ctx = canvas.getContext('2d');
    const CW = canvas.width, CH = canvas.height;
    ctx.clearRect(0, 0, CW, CH);
    ctx.fillStyle = '#f8fafc';
    ctx.fillRect(0, 0, CW, CH);

    const nodeEls = Array.from(container.querySelectorAll('[data-hv-node-id]'));
    const focusEl = container.querySelector('.hv-focus-box');
    const allEls = focusEl ? [focusEl, ...nodeEls] : nodeEls;
    if (!allEls.length) {
      layoutRef.current = { minX: 0, minY: 0, scale: 0, PAD: 8, offsetX: 0, offsetY: 0 };
      return;
    }

    const cr = container.getBoundingClientRect();
    const sl = container.scrollLeft, st = container.scrollTop;
    const rects = allEls.map(el => {
      const r = el.getBoundingClientRect();
      return { x: r.left - cr.left + sl, y: r.top - cr.top + st, w: r.width, h: r.height, isFocus: el === focusEl };
    });

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    rects.forEach(r => {
      minX = Math.min(minX, r.x); minY = Math.min(minY, r.y);
      maxX = Math.max(maxX, r.x + r.w); maxY = Math.max(maxY, r.y + r.h);
    });

    const PAD = 8;
    const worldW = Math.max(1, (maxX - minX) + PAD * 2);
    const worldH = Math.max(1, (maxY - minY) + PAD * 2);
    const scale = Math.min(CW / worldW, CH / worldH);
    const offsetX = (CW - worldW * scale) / 2;
    const offsetY = (CH - worldH * scale) / 2;
    layoutRef.current = { minX, minY, scale, PAD, offsetX, offsetY };

    rects.forEach(r => {
      const sx = offsetX + (r.x - minX + PAD) * scale;
      const sy = offsetY + (r.y - minY + PAD) * scale;
      const sw = Math.max(2, r.w * scale), sh = Math.max(2, r.h * scale);
      ctx.fillStyle = r.isFocus ? '#dbeafe' : '#e5e7eb';
      ctx.strokeStyle = r.isFocus ? '#2563eb' : '#9ca3af';
      ctx.lineWidth = r.isFocus ? 0.8 : 0.5;
      ctx.beginPath(); ctx.rect(sx, sy, sw, sh); ctx.fill(); ctx.stroke();
    });

    // Viewport indicator
    const vpX = offsetX + (sl - minX + PAD) * scale;
    const vpY = offsetY + (st - minY + PAD) * scale;
    const vpW = container.clientWidth * scale, vpH = container.clientHeight * scale;
    ctx.fillStyle = 'rgba(239,68,68,0.08)';
    ctx.fillRect(vpX, vpY, Math.max(4, vpW), Math.max(4, vpH));
    ctx.strokeStyle = '#ef4444';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(vpX, vpY, Math.max(4, vpW), Math.max(4, vpH));
  }, [containerRef]);

  useEffect(() => {
    const c = containerRef.current;
    if (!c) return;
    const onScroll = () => requestAnimationFrame(redraw);
    c.addEventListener('scroll', onScroll, { passive: true });
    const ro = new ResizeObserver(() => requestAnimationFrame(redraw));
    ro.observe(c);
    requestAnimationFrame(redraw);
    return () => { c.removeEventListener('scroll', onScroll); ro.disconnect(); };
  }, [containerRef, redraw]);

  useEffect(() => {
    const t = setTimeout(() => requestAnimationFrame(redraw), 100);
    return () => clearTimeout(t);
  }, [watchKey, redraw]);

  const handleClick = useCallback((e) => {
    const canvas = canvasRef.current, c = containerRef.current;
    if (!canvas || !c) return;
    const cr = canvas.getBoundingClientRect();
    const cx = e.clientX - cr.left, cy = e.clientY - cr.top;
    const { minX, minY, scale, PAD, offsetX, offsetY } = layoutRef.current;
    if (!scale) return;
    c.scrollTo({
      left: (cx - offsetX) / scale + minX - PAD - c.clientWidth / 2,
      top: (cy - offsetY) / scale + minY - PAD - c.clientHeight / 2,
      behavior: 'smooth',
    });
  }, [containerRef]);

  return (
    <div className="oc-minimap">
      <div className="oc-minimap-label">Overview</div>
      <canvas ref={canvasRef} width={200} height={140} onClick={handleClick} title="Click to navigate" />
    </div>
  );
};

// ── ExplodableChildRow: recursive indented tree (kept for reference) ──────────
// depth=0 primary children: horizontal wrap when nothing exploded, vertical once any exploded.
// depth>0 always vertical (leaf rows included), indented one box+gap unit per level.
// visitedIds tracks every ancestor shown so far — cyclic nodes are flagged but still rendered.
const ExplodableChildRow = ({ items, nodeList, relList, explodedNodes, onExplode, onExplodeAll, onFocus, depth = 0, visitedIds = new Set() }) => {
  const anyExploded = items.some(item => !visitedIds.has(item.nodeId) && explodedNodes.has(item.nodeId));

  // Primary level with nothing exploded → original horizontal wrap row
  if (depth === 0 && !anyExploded) {
    return (
      <div className="hv-owned-row">
        {items.map(item => (
          <HvNeighborBox
            key={item.nodeId}
            item={item}
            nodeList={nodeList}
            relList={relList}
            explodedNodes={explodedNodes}
            onExplode={onExplode}
            onExplodeAll={onExplodeAll}
            onFocus={onFocus}
            isCyclic={visitedIds.has(item.nodeId)}
          />
        ))}
      </div>
    );
  }

  // Vertical column — either primary level after explosion, or any sub-level
  return (
    <div className="hv-owned-column">
      {items.map(item => {
        const isCyclic = visitedIds.has(item.nodeId);
        // Cyclic nodes are shown (red, no explode) but never expanded further
        const isExploded = !isCyclic && explodedNodes.has(item.nodeId);
        const nextVisited = new Set(visitedIds);
        nextVisited.add(item.nodeId);
        const childItems = isExploded ? getOwnedBy(relList, item.nodeId) : [];
        return (
          <React.Fragment key={item.nodeId}>
            <HvNeighborBox
              item={item}
              nodeList={nodeList}
              relList={relList}
              explodedNodes={explodedNodes}
              onExplode={onExplode}
              onExplodeAll={onExplodeAll}
              onFocus={onFocus}
              isCyclic={isCyclic}
            />
            {isExploded && childItems.length > 0 && (
              <div className="hv-explode-indent">
                <ExplodableChildRow
                  items={childItems}
                  nodeList={nodeList}
                  relList={relList}
                  explodedNodes={explodedNodes}
                  onExplode={onExplode}
                  onExplodeAll={onExplodeAll}
                  onFocus={onFocus}
                  depth={depth + 1}
                  visitedIds={nextVisited}
                />
              </div>
            )}
          </React.Fragment>
        );
      })}
    </div>
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
  const [nodeList, setNodeList] = useState(() => (token ? [] : initialNodes));
  const [relList, setRelList] = useState(() => (token ? [] : initialRelationships));
  const [homeScreen, setHomeScreen] = useState(() => {
    if (token) return null;
    if (typeof window === "undefined") return null;
    try {
      const s = localStorage.getItem("homeScreen");
      return s ? JSON.parse(s) : null;
    } catch { return null; }
  });
  // homeScreen is loaded from the server after login (see effect below)
  const [viewMode, setViewMode] = useState(() => homeScreen?.viewMode ?? "hierarchy");
  const [explodedNodes, setExplodedNodes] = useState(new Set());
  const [explodedAnchorId, setExplodedAnchorId] = useState(null);
  const [focusId, setFocusId] = useState(() => {
    if (token) return "";
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
  const explodedAnchorSnapshotRef = useRef(null);
  const hierarchyPanRef = useRef({
    active: false,
    pointerId: null,
    moved: false,
    startX: 0,
    startY: 0,
    startLeft: 0,
    startTop: 0,
    prevScrollBehavior: "",
    suppressClickUntil: 0,
  });
  const apiBase = import.meta.env.VITE_API_URL || "http://localhost:5174";
  const [remoteStatus, setRemoteStatus] = useState("idle");
  const [remoteError, setRemoteError] = useState("");
  const [exitPromptOpen, setExitPromptOpen] = useState(false);
  const [directoryLoaded, setDirectoryLoaded] = useState(false);
  const [profileLoading, setProfileLoading] = useState(true);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [uploadFile, setUploadFile] = useState(null);
  const [uploadKind, setUploadKind] = useState("entity");
  const [uploadType, setUploadType] = useState("entity");
  const [uploadStatus, setUploadStatus] = useState("idle");
  const [uploadError, setUploadError] = useState("");
  const [uploadSummary, setUploadSummary] = useState(null);
  const [uploadDetected, setUploadDetected] = useState("");
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [uploadPreview, setUploadPreview] = useState(null);
  const [uploadPreviewLoading, setUploadPreviewLoading] = useState(false);
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

    const secondCol = sampleLines.map((row) => (row[1] || "").toLowerCase());
    const hasKind = secondCol.some((value) => value === "entity" || value === "person");
    if (hasKind) {
      const defaultKind = secondCol.includes("person") && !secondCol.includes("entity") ? "person" : "entity";
      return { type: defaultKind, label: defaultKind === "person" ? "Persons" : "Entities" };
    }

    return { type: "entity", label: "Entities" };
  };

  const clientId = clientIdProp || (token ? "" : "test");

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
    if (token) return "";
    if (typeof window === "undefined") return "";
    try {
      const s = localStorage.getItem("homeScreen");
      const hs = s ? JSON.parse(s) : null;
      return ((hs?.viewMode === "directory" || hs?.viewMode === "tabular") && hs?.dirFilter) ? hs.dirFilter : "";
    } catch { return ""; }
  });
  const [tableDrafts, setTableDrafts] = useState({});
  const [tableDirtyKeys, setTableDirtyKeys] = useState(() => new Set());
  const [tableSavingKeys, setTableSavingKeys] = useState(() => new Set());
  const [tableRowErrors, setTableRowErrors] = useState({});
  const [tableNewRows, setTableNewRows] = useState([]);
  const [tabularViews, setTabularViews] = useState([]);
  const [selectedTabularViewId, setSelectedTabularViewId] = useState(() => {
    if (homeScreen?.viewMode === "tabular" && homeScreen?.selectedTabularViewId) {
      return homeScreen.selectedTabularViewId;
    }
    return DEFAULT_TABULAR_VIEW_ID;
  });
  const [tabularViewDialogOpen, setTabularViewDialogOpen] = useState(false);
  const [tabularViewDraft, setTabularViewDraft] = useState({
    name: "",
    columnOrder: [],
  });
  const [tabularViewNameError, setTabularViewNameError] = useState("");
  const getTabularPrefsPayload = useCallback((overrides = {}) => ({
    tabularViews: overrides.tabularViews ?? tabularViews,
    tabularViewsSelectedId: overrides.tabularViewsSelectedId ?? selectedTabularViewId,
  }), [selectedTabularViewId, tabularViews]);

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

  const [editNodeId, setEditNodeId] = useState(() => (token ? "" : (initialNodes[0]?.id ?? "")));
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
  const [editOwnershipId, setEditOwnershipId] = useState(() => (
    token ? "" : (initialRelationships.find((r) => r.type === "owns")?.id ?? "")
  ));
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
  const [editEmploymentId, setEditEmploymentId] = useState(() => (
    token ? "" : (initialRelationships.find((r) => r.type === "employs")?.id ?? "")
  ));
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
  const [printDialogMode, setPrintDialogMode] = useState(null); // "book" | "poster" | null
  const [printTargetNodeId, setPrintTargetNodeId] = useState("");
  const [printHierarchy, setPrintHierarchy] = useState(true);
  const [printDetail, setPrintDetail] = useState(true);
  const [posterConfirmed, setPosterConfirmed] = useState(false);
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
    if (token) return "user";
    try { return localStorage.getItem("myRole") || "user"; } catch { return "user"; }
  });
  const [clientDisplayName, setClientDisplayName] = useState(() => {
    if (token) return "";
    try { return localStorage.getItem("clientDisplayName") || ""; } catch { return ""; }
  });
  const [clientInfoOpen, setClientInfoOpen] = useState(false);
  const [clientInfoDraft, setClientInfoDraft] = useState({ clientName: "", address: "", billingContact: "", billingEmail: "", billingPhone: "", notes: "" });
  const [clientInfoBusy, setClientInfoBusy] = useState(false);
  const [clientInfoError, setClientInfoError] = useState("");
  const [serverInfoOpen, setServerInfoOpen] = useState(false);
  const [serverInfoData, setServerInfoData] = useState(null);
  const [serverInfoBusy, setServerInfoBusy] = useState(false);
  const [cloneClientOpen, setCloneClientOpen] = useState(false);
  const [cloneClientDraft, setCloneClientDraft] = useState("");
  const [cloneClientBusy, setCloneClientBusy] = useState(false);
  const [cloneClientError, setCloneClientError] = useState("");
  const [cloneClientResult, setCloneClientResult] = useState(null);
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
  const filteredAllNodes = useMemo(() => {
    const sorted = [...nodeList].sort((a, b) =>
      String(a.name || a.id || "").localeCompare(String(b.name || b.id || ""), undefined, {
        sensitivity: "base",
      })
    );
    if (!dirSearchLower) return sorted;
    return sorted.filter((n) => (n.name || n.id || "").toLowerCase().includes(dirSearchLower));
  }, [nodeList, dirSearchLower]);
  const tableDdFields = useMemo(
    () => [...dataDictionary].sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0)),
    [dataDictionary]
  );
  const baseTabularColumns = useMemo(() => ([
    { key: "status", label: "Status", hideable: false },
    { key: "type", label: "Type", hideable: true },
    { key: "name", label: "Name", hideable: true },
    { key: "address",   label: "Address",       hideable: true },
    { key: "workPhone", label: "Primary Phone",  hideable: true },
    { key: "cellPhone", label: "Cell Phone",     hideable: true },
    { key: "emails",    label: "e-Mail",         hideable: true },
    { key: "taxId",     label: "Tax ID",         hideable: true },
    { key: "operationalRole", label: "Operational Role", hideable: true },
    { key: "legalStatus", label: "Legal Status", hideable: true },
    { key: "personStatus", label: "Person Status", hideable: true },
    { key: "actions", label: "Actions", hideable: false },
  ]), []);
  const allTabularColumns = useMemo(() => {
    const ddCols = tableDdFields.map((field) => ({
      key: field.fieldId,
      label: field.prompt,
      field,
      hideable: true,
    }));
    const selectableBaseColumns = baseTabularColumns.filter((c) => c.key !== "status" && c.key !== "actions");
    return [...selectableBaseColumns, ...ddCols].filter(Boolean);
  }, [baseTabularColumns, tableDdFields]);
  const defaultTabularOrder = useMemo(
    () => allTabularColumns.map((c) => c.key),
    [allTabularColumns]
  );

  const normalizeTabularColumnKey = useCallback((key) => {
    const text = String(key || "").trim();
    if (text.startsWith("dd:dd:")) return `dd:${text.slice(6)}`;
    return text;
  }, []);

  const sanitizeTabularView = useCallback((view) => {
    if (!view || typeof view !== "object") return null;
    const id = String(view.id || "").trim();
    const name = String(view.name || "").trim();
    if (!id || !name) return null;
    const allowed = new Set(defaultTabularOrder);
    const rawOrder = Array.isArray(view.columnOrder)
      ? view.columnOrder.map(normalizeTabularColumnKey).filter((k) => allowed.has(k))
      : [];
    const hiddenSet = new Set(Array.isArray(view.hidden)
      ? view.hidden.map(normalizeTabularColumnKey).filter((k) => allowed.has(k))
      : []);
    const visibleOrder = rawOrder.filter((k) => !hiddenSet.has(k));
    return { id, name, columnOrder: visibleOrder };
  }, [defaultTabularOrder, normalizeTabularColumnKey]);

  const captureHomeScreen = useCallback(() => {
    const screen = {
      viewMode,
      focusId,
      dirFilter: (viewMode === "directory" || viewMode === "tabular") ? dirSearch : null,
    };
    if (viewMode === "tabular") {
      screen.selectedTabularViewId = selectedTabularViewId;
    }
    return screen;
  }, [dirSearch, focusId, selectedTabularViewId, viewMode]);

  const restoreHomeScreen = useCallback((screen) => {
    if (!screen) return;
    setViewMode(screen.viewMode ?? "hierarchy");
    if (screen.focusId) setFocusId(screen.focusId);
    if (screen.viewMode === "directory" || screen.viewMode === "tabular") {
      setDirSearch(screen.dirFilter ?? "");
    }
    if (screen.viewMode === "tabular") {
      setSelectedTabularViewId(screen.selectedTabularViewId || DEFAULT_TABULAR_VIEW_ID);
    }
  }, []);

  const isSameHomeScreen = useCallback((screen) => {
    if (!screen) return false;
    if (viewMode !== (screen.viewMode ?? "hierarchy")) return false;
    if ((focusId || "") !== (screen.focusId || "")) return false;
    const currentFilter = (viewMode === "directory" || viewMode === "tabular") ? dirSearch : "";
    const screenFilter = (screen.viewMode === "directory" || screen.viewMode === "tabular") ? (screen.dirFilter ?? "") : "";
    if (currentFilter !== screenFilter) return false;
    const currentTabularId = viewMode === "tabular" ? selectedTabularViewId : DEFAULT_TABULAR_VIEW_ID;
    const screenTabularId = screen.viewMode === "tabular" ? (screen.selectedTabularViewId ?? DEFAULT_TABULAR_VIEW_ID) : DEFAULT_TABULAR_VIEW_ID;
    return currentTabularId === screenTabularId;
  }, [dirSearch, focusId, selectedTabularViewId, viewMode]);

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

  useEffect(() => {
    if (!myLoginId) return;
    apiRequest("/api/auth/me", {
      method: "PATCH",
      body: JSON.stringify({
        tabularViews,
        tabularViewsSelectedId: selectedTabularViewId,
      }),
    }).catch(() => { });
    // apiRequest is intentionally omitted to avoid firing on every render.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [myLoginId, selectedTabularViewId, tabularViews]);

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
        // Preserve the latest focus when possible; avoid overriding a just-restored home focus.
        setFocusId((prev) => (data.nodes.some((n) => n.id === prev) ? prev : data.nodes[0].id));
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
  }, [apiBase, token]);

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
        const ext = (uploadFile.name || "").split(".").pop().toLowerCase();
        let csvText;
        if (ext === "xlsx" || ext === "xls") {
          const buffer = await uploadFile.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: "array" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          csvText = XLSX.utils.sheet_to_csv(sheet);
        } else {
          csvText = await readFileText(uploadFile);
        }
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
            headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
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
        : uploadType === "details"
        ? "/api/import/details-csv/upload"
        : "/api/import/nodes-csv/upload";
      const response = await fetch(`${apiBase}${endpoint}`, {
        method: "POST",
        headers: token ? { Authorization: `Bearer ${token}` } : {},
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
        // For ownership resolution failures the server sends per-row errors — surface them.
        if (Array.isArray(data?.errors) && data.errors.length > 0) {
          setUploadSummary({ imported: 0, skipped: data.skipped || 0, errors: data.errors, total: data.errors.length });
          setUploadStatus("success");
          await loadDirectory();
          return;
        }
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

  const handlePreview = async () => {
    if (!uploadFile) {
      setUploadError("Please select a file.");
      setUploadStatus("error");
      return;
    }
    setUploadPreviewLoading(true);
    setUploadError("");
    setUploadStatus("idle");
    setUploadPreview(null);
    try {
      const form = new FormData();
      form.append("file", uploadFile);
      form.append("importType", uploadType);
      form.append("defaultKind", uploadKind);
      const response = await fetch(`${apiBase}/api/import/preview/upload`, {
        method: "POST",
        headers: token ? { Authorization: `Bearer ${token}` } : {},
        body: form,
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Preview failed");
      setUploadPreview(data);
    } catch (err) {
      setUploadError(err.message || "Preview failed");
    } finally {
      setUploadPreviewLoading(false);
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

  const downloadTemplate = () => {
    if (uploadType === "details") {
      // Built-in fields included in the details template
      const builtInCols = [
        { header: "Entity or Person's Name", helper: "[← Delete rows you don't need. Fill in the columns and upload.]" },
        { header: "Address",       helper: "" },
        { header: "Primary Phone", helper: "" },
        { header: "Cell Phone",    helper: "(People only)" },
        { header: "e-Mail",        helper: "(People only)" },
        { header: "Tax ID",        helper: "" },
      ];

      // Mirrors the server's normalizeHeader — strips punctuation, lowercases, collapses spaces.
      const normalizeH = (str) =>
        String(str || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();

      // All synonyms for the built-in fields already in the template.
      // Any DD field whose prompt normalizes to one of these is redundant and excluded.
      const BUILTIN_SYNONYMS = new Set([
        "name", "company name", "entity name", "node name", "organization", "org name", "business name", "legal name", "entity or person s name",
        "kind", "type", "node type", "entity type", "business type",
        "address", "street", "street address", "mailing address", "location", "addr",
        "work phone", "phone", "workphone", "office phone", "ph work", "business phone", "telephone", "tel", "phone number", "primary phone",
        "cell phone", "cell", "mobile", "mobile phone", "cellphone", "cell number", "personal phone",
        "email", "emails", "email address", "e mail", "email addr",
        "tax id", "taxid", "ein", "tin", "federal id", "tax identification", "federal tax id", "fein",
      ]);

      // DD custom fields — exclude file type and any field redundant with a built-in, sorted by sortOrder
      const ddFields = [...dataDictionary]
        .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
        .filter((f) => f.dataType !== "file")
        .filter((f) => !BUILTIN_SYNONYMS.has(normalizeH(f.prompt)));

      const ddCols = ddFields.map((f) => {
        const parts = [];
        if (Array.isArray(f.validValues) && f.validValues.length > 0) {
          parts.push(f.validValues.join(", "));
        }
        if (f.appliesTo === "entity") parts.push("(Entities only)");
        else if (f.appliesTo === "person") parts.push("(People only)");
        return { header: f.prompt, helper: parts.join(" ") };
      });

      const allCols = [...builtInCols, ...ddCols];
      const headers = allCols.map((c) => c.header);
      const helpers = allCols.map((c) => c.helper);

      // Pre-populate rows with existing node names: entities alpha first, then people alpha.
      const sortedNodes = [...nodeList].sort((a, b) => {
        if (a.kind !== b.kind) return a.kind === "entity" ? -1 : 1;
        return String(a.name || "").localeCompare(String(b.name || ""));
      });
      const nodeRows = sortedNodes.map((n) => [n.name, ...Array(allCols.length - 1).fill("")]);

      const ws = XLSX.utils.aoa_to_sheet([headers, helpers, ...nodeRows]);
      ws["!cols"] = allCols.map((c) => ({
        wch: Math.max(c.header.length, c.helper.length, 20) + 4,
      }));
      // Freeze top two rows (header + helper) so they stay visible while scrolling.
      ws["!freeze"] = { xSplit: 0, ySplit: 2 };

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Template");
      XLSX.writeFile(wb, "details-template.xlsx");
      return;
    }

    const TEMPLATES = {
      entity: {
        headers: ["name", "kind"],
        example: ["Example Corp", "entity"],
        fileName: "entities-template.xlsx",
      },
      person: {
        headers: ["name", "kind"],
        example: ["Jane Doe", "person"],
        fileName: "persons-template.xlsx",
      },
      ownership: {
        headers: ["owner", "owned", "percent"],
        example: ["Parent Corp", "Subsidiary LLC", "100"],
        fileName: "ownerships-template.xlsx",
      },
    };
    const tpl = TEMPLATES[uploadType] || TEMPLATES.entity;
    const ws = XLSX.utils.aoa_to_sheet([tpl.headers, tpl.example]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, tpl.fileName);
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

  const parseTableFieldValue = useCallback((field, raw) => {
    if (field?.multiValue) {
      return String(raw || "")
        .split(",")
        .map((v) => v.trim())
        .filter(Boolean);
    }
    if (field?.dataType === "boolean") {
      if (raw === "") return "";
      return raw === "true";
    }
    return raw;
  }, []);

  const tableFieldToString = useCallback((field, value) => {
    if (Array.isArray(value)) return value.join(", ");
    if (typeof value === "boolean") return value ? "true" : "false";
    return value ?? "";
  }, []);

  const normalizeCustomFields = useCallback((value) => {
    const obj = value && typeof value === "object" && !Array.isArray(value) ? value : {};
    const normalized = {};
    Object.keys(obj)
      .sort((a, b) => a.localeCompare(b))
      .forEach((key) => {
        const v = obj[key];
        normalized[key] = Array.isArray(v) ? [...v] : v;
      });
    return normalized;
  }, []);

  const hasNodeDraftChanges = useCallback((node, draft) => {
    if (!node || !draft) return false;
    if ((draft.name || "") !== (node.name || "")) return true;
    if ((draft.kind || "entity") !== (node.kind || "entity")) return true;
    if ((draft.photo || "") !== (node.photo || "")) return true;
    if ((draft.logo || "") !== (node.logo || "")) return true;
    if ((draft.operationalRole || "") !== (node.operationalRole || "")) return true;
    if ((draft.legalStatus || "") !== (node.legalStatus || "")) return true;
    if ((draft.personStatus || "") !== (node.personStatus || "")) return true;
    if ((draft.address || "") !== (node.address || "")) return true;
    if ((draft.workPhone || "") !== (node.workPhone || "")) return true;
    if ((draft.cellPhone || "") !== (node.cellPhone || "")) return true;
    if ((draft.taxId || "") !== (node.taxId || "")) return true;
    const toEmailStr = (v) => Array.isArray(v) ? v.join(", ") : (v || "");
    if (toEmailStr(draft.emails) !== toEmailStr(node.emails)) return true;
    const nodeCustom = JSON.stringify(normalizeCustomFields(node.customFields || {}));
    const draftCustom = JSON.stringify(normalizeCustomFields(draft.customFields || {}));
    return nodeCustom !== draftCustom;
  }, [normalizeCustomFields]);

  const addTableRow = useCallback((kind) => {
    const key = `__new__${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const draft = {
      id: "",
      name: "",
      kind,
      photo: "",
      logo: "",
      operationalRole: "",
      legalStatus: "",
      personStatus: "",
      address: "",
      workPhone: "",
      cellPhone: "",
      emails: "",
      taxId: "",
      customFields: {},
      client: clientId,
    };
    setTableNewRows((prev) => [{ key, draft }, ...prev]);
    setTableDirtyKeys((prev) => {
      const next = new Set(prev);
      next.add(key);
      return next;
    });
    setTableRowErrors((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  }, [clientId]);

  const removeTableNewRow = useCallback((key) => {
    setTableNewRows((prev) => prev.filter((r) => r.key !== key));
    setTableDirtyKeys((prev) => {
      const next = new Set(prev);
      next.delete(key);
      return next;
    });
    setTableSavingKeys((prev) => {
      const next = new Set(prev);
      next.delete(key);
      return next;
    });
    setTableRowErrors((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  }, []);

  const updateTableRowDraft = useCallback((row, patch) => {
    if (row.isNew) {
      setTableNewRows((prev) =>
        prev.map((r) => (r.key === row.key ? { ...r, draft: { ...r.draft, ...patch } } : r))
      );
      setTableDirtyKeys((prev) => {
        const next = new Set(prev);
        next.add(row.key);
        return next;
      });
    } else {
      const merged = { ...(tableDrafts[row.key] || row.node), ...patch };
      setTableDrafts((prev) => ({
        ...prev,
        [row.key]: merged,
      }));
      setTableDirtyKeys((prev) => {
        const next = new Set(prev);
        if (hasNodeDraftChanges(row.node, merged)) next.add(row.key);
        else next.delete(row.key);
        return next;
      });
    }
    setTableRowErrors((prev) => {
      const next = { ...prev };
      delete next[row.key];
      return next;
    });
  }, [hasNodeDraftChanges, tableDrafts]);

  const saveTableRow = useCallback(async (rowKey) => {
    const newRow = tableNewRows.find((r) => r.key === rowKey);
    const existingNode = newRow ? null : nodeList.find((n) => n.id === rowKey);
    const draft = newRow ? newRow.draft : (tableDrafts[rowKey] || existingNode);
    if (!newRow && existingNode && draft && !hasNodeDraftChanges(existingNode, draft)) {
      setTableDirtyKeys((prev) => {
        const next = new Set(prev);
        next.delete(rowKey);
        return next;
      });
      return true;
    }
    if (!draft) return false;
    if (!String(draft.name || "").trim()) {
      setTableRowErrors((prev) => ({ ...prev, [rowKey]: "Name is required." }));
      return false;
    }

    setTableSavingKeys((prev) => {
      const next = new Set(prev);
      next.add(rowKey);
      return next;
    });

    try {
      if (newRow) {
        const id = makeNodeId(draft.kind, draft.name);
        const payload = {
          id,
          name: draft.name.trim(),
          kind: draft.kind,
          client: clientId,
          photo: draft.kind === "person" ? (draft.photo || "") : "",
          logo: draft.kind === "entity" ? (draft.logo || "") : "",
          operationalRole: draft.kind === "entity" ? (draft.operationalRole || "") : "",
          legalStatus: draft.kind === "entity" ? (draft.legalStatus || "") : "",
          personStatus: draft.kind === "person" ? (draft.personStatus || "") : "",
          address: draft.address || "",
          workPhone: draft.workPhone || "",
          cellPhone: draft.cellPhone || "",
          emails: draft.emails || "",
          taxId: draft.taxId || "",
          customFields: draft.customFields || {},
        };
        await apiRequest("/api/nodes", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        setNodeList((prev) => [...prev, payload]);
        removeTableNewRow(rowKey);
      } else if (existingNode) {
        const newId = makeNodeId(draft.kind, draft.name, rowKey);
        const payload = {
          name: draft.name.trim(),
          kind: draft.kind,
          client: clientId,
          newId: newId !== rowKey ? newId : null,
          photo: draft.kind === "person" ? (draft.photo || "") : "",
          logo: draft.kind === "entity" ? (draft.logo || "") : "",
          operationalRole: draft.kind === "entity" ? (draft.operationalRole || "") : "",
          legalStatus: draft.kind === "entity" ? (draft.legalStatus || "") : "",
          personStatus: draft.kind === "person" ? (draft.personStatus || "") : "",
          address: draft.address || "",
          workPhone: draft.workPhone || "",
          cellPhone: draft.cellPhone || "",
          emails: draft.emails || "",
          taxId: draft.taxId || "",
          customFields: draft.customFields || {},
        };
        await apiRequest(`/api/nodes/${rowKey}`, {
          method: "PUT",
          body: JSON.stringify(payload),
        });

        setNodeList((prev) =>
          prev.map((n) =>
            n.id === rowKey
              ? {
                ...n,
                id: newId,
                name: payload.name,
                kind: payload.kind,
                photo: payload.photo,
                logo: payload.logo,
                operationalRole: payload.operationalRole,
                legalStatus: payload.legalStatus,
                personStatus: payload.personStatus,
                address: payload.address,
                workPhone: payload.workPhone,
                cellPhone: payload.cellPhone,
                emails: payload.emails,
                taxId: payload.taxId,
                customFields: payload.customFields,
                client: n.client || clientId,
              }
              : n
          )
        );
        if (newId !== rowKey) {
          setRelList((prev) =>
            prev.map((r) => ({
              ...r,
              from: r.from === rowKey ? newId : r.from,
              to: r.to === rowKey ? newId : r.to,
            }))
          );
          if (focusId === rowKey) setFocusId(newId);
          if (editNodeId === rowKey) setEditNodeId(newId);
        }
        setTableDrafts((prev) => {
          const next = { ...prev };
          delete next[rowKey];
          if (newId !== rowKey) delete next[newId];
          return next;
        });
        setTableDirtyKeys((prev) => {
          const next = new Set(prev);
          next.delete(rowKey);
          if (newId !== rowKey) next.delete(newId);
          return next;
        });
        setTableRowErrors((prev) => {
          const next = { ...prev };
          delete next[rowKey];
          if (newId !== rowKey) delete next[newId];
          return next;
        });
      }
      setRemoteStatus("connected");
      return true;
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
      setTableRowErrors((prev) => ({ ...prev, [rowKey]: err.message || "Unable to save." }));
      return false;
    } finally {
      setTableSavingKeys((prev) => {
        const next = new Set(prev);
        next.delete(rowKey);
        return next;
      });
    }
  }, [
    apiRequest,
    clientId,
    editNodeId,
    focusId,
    hasNodeDraftChanges,
    makeNodeId,
    nodeList,
    removeTableNewRow,
    tableDrafts,
    tableNewRows,
  ]);

  const pendingTableKeys = useMemo(() => {
    const pending = new Set();
    tableNewRows.forEach((r) => pending.add(r.key));
    Object.entries(tableDrafts).forEach(([key, draft]) => {
      const node = nodeList.find((n) => n.id === key);
      if (node && hasNodeDraftChanges(node, draft)) pending.add(key);
    });
    return pending;
  }, [hasNodeDraftChanges, nodeList, tableDrafts, tableNewRows]);

  const saveAllTableRows = useCallback(async () => {
    const keys = [...pendingTableKeys];
    for (const key of keys) {
      // eslint-disable-next-line no-await-in-loop
      await saveTableRow(key);
    }
  }, [pendingTableKeys, saveTableRow]);

  const tableRows = useMemo(() => {
    const newRows = tableNewRows.filter((r) => {
      if (!dirSearchLower) return true;
      const txt = `${r.draft.name || ""} ${r.draft.id || ""}`.toLowerCase();
      return txt.includes(dirSearchLower);
    }).map((r) => ({ key: r.key, isNew: true, node: r.draft }));
    const existingRows = filteredAllNodes.map((n) => ({ key: n.id, isNew: false, node: n }));
    return [...newRows, ...existingRows];
  }, [dirSearchLower, filteredAllNodes, tableNewRows]);

  const activeTabularView = useMemo(() => {
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) {
      return {
        id: DEFAULT_TABULAR_VIEW_ID,
        name: "Default",
        columnOrder: defaultTabularOrder,
      };
    }
    const found = tabularViews.find((v) => v.id === selectedTabularViewId);
    return found || {
      id: DEFAULT_TABULAR_VIEW_ID,
      name: "Default",
      columnOrder: defaultTabularOrder,
    };
  }, [defaultTabularOrder, selectedTabularViewId, tabularViews]);

  const visibleTabularColumns = useMemo(() => {
    const byKey = new Map(allTabularColumns.map((c) => [c.key, c]));
    const seen = new Set();
    const middle = (activeTabularView.columnOrder || [])
      .map((key) => byKey.get(key))
      .filter((col) => {
        if (!col || seen.has(col.key)) return false;
        seen.add(col.key);
        return true;
      });
    return [
      { key: "status", label: "Status" },
      ...middle,
      { key: "actions", label: "Actions" },
    ];
  }, [activeTabularView.columnOrder, allTabularColumns]);

  const availableTabularColumns = useMemo(() => {
    const selected = new Set(tabularViewDraft.columnOrder || []);
    return allTabularColumns.filter((column) => !selected.has(column.key));
  }, [allTabularColumns, tabularViewDraft.columnOrder]);

  const openTabularViewManager = useCallback(() => {
    setTabularViewNameError("");
    setTabularViewDraft({
      name: activeTabularView.name || "",
      columnOrder: [...(activeTabularView.columnOrder || defaultTabularOrder)],
    });
    setTabularViewDialogOpen(true);
  }, [activeTabularView, defaultTabularOrder]);

  const moveTabularDraftColumn = useCallback((key, direction) => {
    setTabularViewDraft((prev) => {
      const idx = prev.columnOrder.indexOf(key);
      if (idx < 0) return prev;
      const nextIdx = idx + direction;
      if (nextIdx < 0 || nextIdx >= prev.columnOrder.length) return prev;
      const columnOrder = [...prev.columnOrder];
      const tmp = columnOrder[idx];
      columnOrder[idx] = columnOrder[nextIdx];
      columnOrder[nextIdx] = tmp;
      return { ...prev, columnOrder };
    });
  }, []);

  const toggleTabularDraftSelected = useCallback((key) => {
    setTabularViewDraft((prev) => {
      const nextOrder = [...prev.columnOrder];
      const idx = nextOrder.indexOf(key);
      if (idx >= 0) {
        nextOrder.splice(idx, 1);
      } else {
        nextOrder.push(key);
      }
      return { ...prev, columnOrder: nextOrder };
    });
  }, []);

  const isDuplicateTabularViewName = useCallback((name, excludeId = "") => {
    const normalized = String(name || "").trim().toLowerCase();
    if (!normalized) return false;
    return tabularViews.some((v) =>
      v.id !== excludeId && String(v.name || "").trim().toLowerCase() === normalized
    );
  }, [tabularViews]);

  const persistTabularViewPrefs = useCallback((nextViews, nextSelectedId) => {
    apiRequest("/api/auth/me", {
      method: "PATCH",
      body: JSON.stringify({
        tabularViews: nextViews,
        tabularViewsSelectedId: nextSelectedId,
      }),
    }).catch((err) => {
      setRemoteStatus("error");
      setRemoteError(err.message || "Unable to save tabular view settings.");
    });
  }, [apiRequest]);

  const saveTabularViewAsNew = useCallback(() => {
    const name = String(tabularViewDraft.name || "").trim();
    if (!name) {
      setTabularViewNameError("Please enter a view name.");
      return;
    }
    if (isDuplicateTabularViewName(name)) {
      setTabularViewNameError("A view with this name already exists.");
      return;
    }
    setTabularViewNameError("");
    const id = `view-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const nextView = sanitizeTabularView({
      id,
      name,
      columnOrder: tabularViewDraft.columnOrder,
    });
    if (!nextView) return;
    const nextViews = [...tabularViews, nextView];
    setTabularViews(nextViews);
    setSelectedTabularViewId(nextView.id);
    persistTabularViewPrefs(nextViews, nextView.id);
    setTabularViewDialogOpen(false);
  }, [isDuplicateTabularViewName, persistTabularViewPrefs, sanitizeTabularView, tabularViewDraft.columnOrder, tabularViewDraft.name, tabularViews]);

  const updateCurrentTabularView = useCallback(() => {
    const name = String(tabularViewDraft.name || "").trim();
    if (!name) {
      setTabularViewNameError("Please enter a view name.");
      return;
    }
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) {
      setTabularViewNameError("Select a saved view first, or use Save As New.");
      return;
    }
    const selectedExists = tabularViews.some((v) => v.id === selectedTabularViewId);
    if (!selectedExists) {
      setTabularViewNameError("Selected view no longer exists. Please reselect a view.");
      return;
    }
    if (isDuplicateTabularViewName(name, selectedTabularViewId)) {
      setTabularViewNameError("A view with this name already exists.");
      return;
    }
    setTabularViewNameError("");

    const nextView = sanitizeTabularView({
      id: selectedTabularViewId,
      name,
      columnOrder: tabularViewDraft.columnOrder,
    });
    if (!nextView) return;
    const nextViews = tabularViews.map((v) => (v.id === selectedTabularViewId ? nextView : v));
    setTabularViews(nextViews);
    persistTabularViewPrefs(nextViews, selectedTabularViewId);
    setTabularViewDialogOpen(false);
  }, [isDuplicateTabularViewName, persistTabularViewPrefs, sanitizeTabularView, selectedTabularViewId, tabularViewDraft.columnOrder, tabularViewDraft.name, tabularViews]);

  const deleteCurrentTabularView = useCallback(() => {
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) return;
    const nextViews = tabularViews.filter((v) => v.id !== selectedTabularViewId);
    setTabularViews(nextViews);
    setSelectedTabularViewId(DEFAULT_TABULAR_VIEW_ID);
    persistTabularViewPrefs(nextViews, DEFAULT_TABULAR_VIEW_ID);
    // If this view was the saved home screen, reset and persist the home screen
    if (homeScreen?.viewMode === "tabular" && homeScreen?.selectedTabularViewId === selectedTabularViewId) {
      const updatedHomeScreen = { ...homeScreen, selectedTabularViewId: DEFAULT_TABULAR_VIEW_ID };
      setHomeScreen(updatedHomeScreen);
      try { localStorage.setItem("homeScreen", JSON.stringify(updatedHomeScreen)); } catch { }
      apiRequest("/api/auth/me", {
        method: "PATCH",
        body: JSON.stringify({
          homeScreen: updatedHomeScreen,
          ...getTabularPrefsPayload({
            tabularViews: nextViews,
            tabularViewsSelectedId: DEFAULT_TABULAR_VIEW_ID,
          }),
        }),
      }).catch(() => { });
    }
    setTabularViewNameError("");
    setTabularViewDialogOpen(false);
  }, [apiRequest, getTabularPrefsPayload, homeScreen, persistTabularViewPrefs, selectedTabularViewId, tabularViews]);

  const renderTabularCell = useCallback((column, { row, rowNode, isSaving, isDirty, rowError, resolvedId }) => {
    switch (column.key) {
      case "status":
        return (
          <td key={`${row.key}-status`}>
            {rowError ? (
              <span className="tabular-status tabular-status--error" title={rowError}>Error</span>
            ) : isSaving ? (
              <span className="tabular-status tabular-status--saving">Saving</span>
            ) : isDirty ? (
              <span className="tabular-status tabular-status--dirty">Unsaved</span>
            ) : row.isNew ? (
              <span className="tabular-status tabular-status--new">New</span>
            ) : (
              <span className="tabular-status tabular-status--saved">Saved</span>
            )}
          </td>
        );
      case "type":
        return (
          <td key={`${row.key}-type`}>
            <select
              className="tabular-cell-input"
              value={rowNode.kind || "entity"}
              onChange={(e) => updateTableRowDraft(row, { kind: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              disabled={isSaving}
            >
              <option value="entity">Entity</option>
              <option value="person">Person</option>
            </select>
          </td>
        );
      case "name":
        return (
          <td key={`${row.key}-name`}>
            <input
              className="tabular-cell-input"
              value={rowNode.name || ""}
              onChange={(e) => updateTableRowDraft(row, { name: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "operationalRole":
        return (
          <td key={`${row.key}-operationalRole`}>
            <select
              className="tabular-cell-input"
              value={rowNode.kind === "entity" ? (rowNode.operationalRole || "") : ""}
              onChange={(e) => updateTableRowDraft(row, { operationalRole: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              disabled={isSaving || rowNode.kind !== "entity"}
            >
              <option value="">-</option>
              <option value="Active">Active</option>
              <option value="Passive">Passive</option>
              <option value="Mixed">Mixed</option>
            </select>
          </td>
        );
      case "legalStatus":
        return (
          <td key={`${row.key}-legalStatus`}>
            <select
              className="tabular-cell-input"
              value={rowNode.kind === "entity" ? (rowNode.legalStatus || "") : ""}
              onChange={(e) => updateTableRowDraft(row, { legalStatus: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              disabled={isSaving || rowNode.kind !== "entity"}
            >
              <option value="">-</option>
              <option value="Good Standing">Good Standing</option>
              <option value="Dormant">Dormant</option>
              <option value="Dissolved">Dissolved</option>
              <option value="Suspended">Suspended</option>
            </select>
          </td>
        );
      case "personStatus":
        return (
          <td key={`${row.key}-personStatus`}>
            <select
              className="tabular-cell-input"
              value={rowNode.kind === "person" ? (rowNode.personStatus || "") : ""}
              onChange={(e) => updateTableRowDraft(row, { personStatus: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              disabled={isSaving || rowNode.kind !== "person"}
            >
              <option value="">-</option>
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
              <option value="Deceased">Deceased</option>
              <option value="Former">Former</option>
            </select>
          </td>
        );
      case "address":
        return (
          <td key={`${row.key}-address`}>
            <input
              className="tabular-cell-input"
              value={rowNode.address || ""}
              onChange={(e) => updateTableRowDraft(row, { address: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "workPhone":
        return (
          <td key={`${row.key}-workPhone`}>
            <input
              className="tabular-cell-input"
              type="tel"
              value={rowNode.workPhone || ""}
              onChange={(e) => updateTableRowDraft(row, { workPhone: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "cellPhone":
        return (
          <td key={`${row.key}-cellPhone`}>
            <input
              className="tabular-cell-input"
              type="tel"
              value={rowNode.cellPhone || ""}
              onChange={(e) => updateTableRowDraft(row, { cellPhone: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "emails":
        return (
          <td key={`${row.key}-emails`}>
            <input
              className="tabular-cell-input"
              type="email"
              value={Array.isArray(rowNode.emails) ? rowNode.emails.join(", ") : (rowNode.emails || "")}
              onChange={(e) => updateTableRowDraft(row, { emails: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "taxId":
        return (
          <td key={`${row.key}-taxId`}>
            <input
              className="tabular-cell-input"
              value={rowNode.taxId || ""}
              onChange={(e) => updateTableRowDraft(row, { taxId: e.target.value })}
              onBlur={() => saveTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      case "actions":
        return (
          <td key={`${row.key}-actions`}>
            <div className="tabular-actions">
              <Button
                type="button"
                variant="outline"
                onClick={() => saveTableRow(row.key)}
                disabled={isSaving || !isDirty}
              >
                Save
              </Button>
              {!row.isNew && (
                <>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => {
                      setFocusId(row.key);
                      setViewMode("hierarchy");
                    }}
                  >
                    Focus
                  </Button>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => openNodeBookPrintDialog(row.key)}
                  >
                    Print Book
                  </Button>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => openNodePosterPrintDialog(row.key)}
                  >
                    Print Org
                  </Button>
                </>
              )}
              {row.isNew && (
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => removeTableNewRow(row.key)}
                >
                  Cancel
                </Button>
              )}
            </div>
          </td>
        );
      default:
        if (!column.field) return <td key={`${row.key}-${column.key}`}></td>;
        {
          const field = column.field;
          const applicable = field.appliesTo === "both" || field.appliesTo === rowNode.kind;
          const validValues = Array.isArray(field.validValues)
            ? field.validValues.filter(Boolean)
            : [];
          const currentVal = rowNode.customFields?.[field.fieldId];
          if (!applicable) {
            return <td key={`${row.key}-${field.fieldId}`} className="tabular-na">-</td>;
          }
          if (field.dataType === "boolean") {
            return (
              <td key={`${row.key}-${field.fieldId}`}>
                <select
                  className="tabular-cell-input"
                  value={tableFieldToString(field, currentVal)}
                  onChange={(e) => {
                    updateTableRowDraft(row, {
                      customFields: {
                        ...(rowNode.customFields || {}),
                        [field.fieldId]: parseTableFieldValue(field, e.target.value),
                      },
                    });
                  }}
                  onBlur={() => saveTableRow(row.key)}
                  disabled={isSaving}
                >
                  <option value="">-</option>
                  <option value="true">Yes</option>
                  <option value="false">No</option>
                </select>
              </td>
            );
          }
          if (validValues.length > 0) {
            return (
              <td key={`${row.key}-${field.fieldId}`}>
                <select
                  className="tabular-cell-input"
                  value={tableFieldToString(field, currentVal)}
                  onChange={(e) => {
                    updateTableRowDraft(row, {
                      customFields: {
                        ...(rowNode.customFields || {}),
                        [field.fieldId]: parseTableFieldValue(field, e.target.value),
                      },
                    });
                  }}
                  onBlur={() => saveTableRow(row.key)}
                  disabled={isSaving}
                >
                  <option value=""></option>
                  {validValues.map((v) => (
                    <option key={v} value={v}>{v}</option>
                  ))}
                </select>
              </td>
            );
          }
          return (
            <td key={`${row.key}-${field.fieldId}`}>
              <input
                className="tabular-cell-input"
                type={dataTypeToHtmlInput(field.dataType)}
                value={tableFieldToString(field, currentVal)}
                onChange={(e) => {
                  updateTableRowDraft(row, {
                    customFields: {
                      ...(rowNode.customFields || {}),
                      [field.fieldId]: parseTableFieldValue(field, e.target.value),
                    },
                  });
                }}
                onBlur={() => saveTableRow(row.key)}
                onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
                disabled={isSaving}
                autoComplete="off"
                data-lpignore="true"
              />
            </td>
          );
        }
    }
  }, [
    parseTableFieldValue,
    removeTableNewRow,
    saveTableRow,
    setFocusId,
    setViewMode,
    tableFieldToString,
    updateTableRowDraft,
  ]);

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
          next.push({
            id: makeRelId(), type: "owns", from: row.nodeId, to: targetId,
            percent: row.percent !== "" ? Number(row.percent) : null,
            startDate: row.startDate || null, endDate: row.endDate || null
          });
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

  // Back-button behavior:
  // 1) Close transient UI first (dialogs/upload)
  // 2) Otherwise follow normal in-app history (view/focus)
  // 3) At app root only, require a second back press to exit
  const _backRef = useRef({});
  _backRef.current = {
    openDialog,
    prevDialog,
    uploadOpen,
    confirmDialog,
    viewMode,
    focusId,
    selectedTabularViewId,
  };
  const backNavRef = useRef({
    initialized: false,
    restoring: false,
    lastKey: "",
    exitArmedUntil: 0,
  });
  const buildAppHistoryState = (isRoot = false) => {
    const s = _backRef.current;
    const state = {
      emplusAppNav: true,
      emplusRoot: isRoot,
      viewMode: s.viewMode,
      focusId: s.focusId,
    };
    if (s.viewMode === "tabular") {
      state.selectedTabularViewId = s.selectedTabularViewId;
    }
    return state;
  };

  useEffect(() => {
    const nav = backNavRef.current;
    const rootState = buildAppHistoryState(true);
    const initialState = buildAppHistoryState(false);
    history.replaceState(rootState, "");
    history.pushState(initialState, "");
    nav.lastKey = JSON.stringify(initialState);
    nav.initialized = true;

    const handlePop = (event) => {
      const s = _backRef.current;

      // First consume back presses for transient UI that should close in-place.
      if (s.confirmDialog) {
        setConfirmDialog(null);
        setExitPromptOpen(false);
        history.pushState(buildAppHistoryState(), "");
        return;
      }
      if (s.openDialog) {
        if (s.prevDialog) {
          setOpenDialog(s.prevDialog);
          setPrevDialog(null);
          _backRef.current = { ...s, openDialog: s.prevDialog, prevDialog: null };
        } else {
          setOpenDialog(null);
          _backRef.current = { ...s, openDialog: null };
        }
        setDupMatches([]);
        setOwnerSearch("");
        setOwnerSearchOpen(false);
        setExitPromptOpen(false);
        history.pushState(buildAppHistoryState(), "");
        return;
      }
      if (s.uploadOpen) {
        setUploadOpen(false);
        setExitPromptOpen(false);
        history.pushState(buildAppHistoryState(), "");
        return;
      }

      // Normal in-app pop to a known app state.
      const next = event.state;
      if (next?.emplusAppNav) {
        // Root sentinel means user tried to leave app from its first history entry.
        if (next.emplusRoot) {
          const now = Date.now();
          if (now < nav.exitArmedUntil) {
            nav.exitArmedUntil = 0;
            setExitPromptOpen(false);
            history.back();
            return;
          }
          nav.exitArmedUntil = now + 2500;
          setExitPromptOpen(true);
          const restoreState = buildAppHistoryState(false);
          history.pushState(restoreState, "");
          nav.lastKey = JSON.stringify(restoreState);
          return;
        }
        setExitPromptOpen(false);
        nav.restoring = true;
        setViewMode(next.viewMode || "hierarchy");
        if (next.focusId) setFocusId(next.focusId);
        if (next.viewMode === "tabular") {
          setSelectedTabularViewId(next.selectedTabularViewId || DEFAULT_TABULAR_VIEW_ID);
        }
        requestAnimationFrame(() => {
          nav.restoring = false;
        });
        return;
      }

      // Root guard: first back warns, second back exits.
      const now = Date.now();
      if (now < nav.exitArmedUntil) {
        nav.exitArmedUntil = 0;
        setExitPromptOpen(false);
        history.back();
        return;
      }
      nav.exitArmedUntil = now + 2500;
      setExitPromptOpen(true);
      const restoreState = buildAppHistoryState(false);
      history.pushState(restoreState, "");
      nav.lastKey = JSON.stringify(restoreState);
    };

    window.addEventListener("popstate", handlePop);
    return () => window.removeEventListener("popstate", handlePop);
    // Run once: this initializes app history and pop handling.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const nav = backNavRef.current;
    if (!nav.initialized || nav.restoring) return;
    const nextState = buildAppHistoryState();
    const nextKey = JSON.stringify(nextState);
    if (nextKey === nav.lastKey) return;
    history.pushState(nextState, "");
    nav.lastKey = nextKey;
  }, [focusId, selectedTabularViewId, viewMode]);

  const handleExitPromptReturn = useCallback(() => {
    backNavRef.current.exitArmedUntil = 0;
    setExitPromptOpen(false);
  }, []);

  const handleExitPromptExit = useCallback(() => {
    backNavRef.current.exitArmedUntil = Date.now() + 2000;
    setExitPromptOpen(false);
    history.back();
  }, []);

  useEffect(() => {
    if (!clientId) return;
    setFocusId((prev) => normalizeClientId(clientId, prev));
  }, [clientId]);

  useEffect(() => {
    let active = true;
    setProfileLoading(true);
    setDirectoryLoaded(false);
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
          try { localStorage.setItem("myRole", data.role); } catch { }
        }
        if (data?.clientName) {
          setClientDisplayName(data.clientName);
          try { localStorage.setItem("clientDisplayName", data.clientName); } catch { }
        }
        if (Array.isArray(data?.tabularViews)) {
          const views = data.tabularViews.map(sanitizeTabularView).filter(Boolean);
          setTabularViews(views);
          setSelectedTabularViewId(data?.tabularViewsSelectedId || DEFAULT_TABULAR_VIEW_ID);
        }
        if (!data?.homeScreen) return;
        setHomeScreen(data.homeScreen);
        try { localStorage.setItem("homeScreen", JSON.stringify(data.homeScreen)); } catch { }
        restoreHomeScreen(data.homeScreen);
      }).catch(() => { }).finally(() => { if (active) setProfileLoading(false); });
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

  const snapshotExplodedAnchorPosition = useCallback((nodeId, anchorEl = null) => {
    const container = hierarchyContainerRef.current;
    if (!container) return;
    const selector = `[data-hv-node-id="${CSS.escape(nodeId)}"]`;
    const matches = Array.from(container.querySelectorAll(selector));
    const isValidAnchorEl = anchorEl instanceof Element && container.contains(anchorEl);
    const isFocusAnchor = isValidAnchorEl && anchorEl.classList.contains("hv-focus-box");
    if (!matches.length && !isFocusAnchor) return;
    const sourceEl =
      isValidAnchorEl
        ? anchorEl
        : matches[0] || null;
    if (!sourceEl) return;
    const el = sourceEl;
    const containerRect = container.getBoundingClientRect();
    const elRect = el.getBoundingClientRect();
    const instanceIndex = matches.indexOf(sourceEl);
    explodedAnchorSnapshotRef.current = {
      nodeId,
      instanceIndex: instanceIndex >= 0 ? instanceIndex : null,
      preferFocusBox: isFocusAnchor,
      left: elRect.left - containerRect.left,
      top: elRect.top - containerRect.top,
    };
  }, []);

  const onHierarchyPointerDown = useCallback((e) => {
    if (viewMode !== "hierarchy") return;
    if (e.pointerType === "mouse" && e.button !== 0) return;
    const container = hierarchyContainerRef.current;
    if (!container) return;

    const target = e.target;
    if (target instanceof Element) {
      const interactiveSelector = "button, input, textarea, select, a, label, [role=\"button\"], .oc-minimap, .oc-minimap *";
      if (target.closest(interactiveSelector)) return;
    }

    const pan = hierarchyPanRef.current;
    pan.active = true;
    pan.pointerId = e.pointerId;
    pan.moved = false;
    pan.startX = e.clientX;
    pan.startY = e.clientY;
    pan.startLeft = container.scrollLeft;
    pan.startTop = container.scrollTop;
    pan.prevScrollBehavior = container.style.scrollBehavior;
    container.style.scrollBehavior = "auto";
    container.classList.add("is-panning");
    try {
      container.setPointerCapture(e.pointerId);
    } catch {
      // Some browsers may throw if capture is unavailable; dragging still works while inside bounds.
    }
  }, [viewMode]);

  const onHierarchyPointerMove = useCallback((e) => {
    const container = hierarchyContainerRef.current;
    const pan = hierarchyPanRef.current;
    if (!container || !pan.active || pan.pointerId !== e.pointerId) return;

    const dx = e.clientX - pan.startX;
    const dy = e.clientY - pan.startY;
    if (!pan.moved && (Math.abs(dx) > 3 || Math.abs(dy) > 3)) {
      pan.moved = true;
    }
    if (pan.moved) {
      container.scrollLeft = pan.startLeft - dx;
      container.scrollTop = pan.startTop - dy;
      e.preventDefault();
    }
  }, []);

  const endHierarchyPointerPan = useCallback((e) => {
    const container = hierarchyContainerRef.current;
    const pan = hierarchyPanRef.current;
    if (!container || !pan.active || pan.pointerId !== e.pointerId) return;

    if (pan.moved) {
      pan.suppressClickUntil = Date.now() + 180;
    }
    pan.active = false;
    pan.pointerId = null;
    container.style.scrollBehavior = pan.prevScrollBehavior || "";
    container.classList.remove("is-panning");
    try {
      container.releasePointerCapture(e.pointerId);
    } catch {
      // Ignore if capture was never acquired.
    }
  }, []);

  const onHierarchyPointerUp = useCallback((e) => {
    endHierarchyPointerPan(e);
  }, [endHierarchyPointerPan]);

  const onHierarchyPointerCancel = useCallback((e) => {
    endHierarchyPointerPan(e);
  }, [endHierarchyPointerPan]);

  const onHierarchyClickCapture = useCallback((e) => {
    const pan = hierarchyPanRef.current;
    if (Date.now() < pan.suppressClickUntil) {
      e.preventDefault();
      e.stopPropagation();
      pan.suppressClickUntil = 0;
    }
  }, []);

  useEffect(() => {
    if (viewMode !== "hierarchy") return;
    const shouldCenterVertically = directOwners.length > 0;
    const center = (behavior) => {
      if (!focusBoxRef.current || !hierarchyContainerRef.current) return;
      const container = hierarchyContainerRef.current;
      const box = focusBoxRef.current;
      const containerRect = container.getBoundingClientRect();
      const boxRect = box.getBoundingClientRect();
      const deltaY = shouldCenterVertically
        ? (boxRect.top + box.clientHeight / 2) - (containerRect.top + container.clientHeight / 2)
        : 0;
      const deltaX = (boxRect.left + box.clientWidth / 2) - (containerRect.left + container.clientWidth / 2);
      container.scrollTo({
        left: container.scrollLeft + deltaX,
        top: container.scrollTop + deltaY,
        behavior,
      });
    };
    requestAnimationFrame(() => center("smooth"));
    // Re-center after images in the focus box have loaded (photos/logos load asynchronously)
    const t = setTimeout(() => center("instant"), 300);
    return () => clearTimeout(t);
  }, [directOwners.length, focusId, viewMode, nodeList]);

  // Reset exploded child nodes whenever the focused entity changes
  useEffect(() => {
    setExplodedNodes(new Set());
    setExplodedAnchorId(null);
    explodedAnchorSnapshotRef.current = null;
  }, [focusId]);

  // Keep the tapped node anchored to the same on-screen spot across expand/collapse.
  useEffect(() => {
    if (!explodedAnchorId || !hierarchyContainerRef.current) return;
    const container = hierarchyContainerRef.current;
    requestAnimationFrame(() => {
      const selector = `[data-hv-node-id="${CSS.escape(explodedAnchorId)}"]`;
      const matches = Array.from(container.querySelectorAll(selector));
      const anchor = explodedAnchorSnapshotRef.current;
      let el = null;
      if (anchor && anchor.nodeId === explodedAnchorId && anchor.preferFocusBox && focusBoxRef.current && container.contains(focusBoxRef.current)) {
        el = focusBoxRef.current;
      } else if (anchor && anchor.nodeId === explodedAnchorId && matches.length) {
        const idx = Math.min(Math.max(anchor.instanceIndex ?? 0, 0), matches.length - 1);
        el = matches[idx];
      } else {
        el = matches[0] || null;
      }
      if (!el) {
        explodedAnchorSnapshotRef.current = null;
        return;
      }
      if (anchor && anchor.nodeId === explodedAnchorId) {
        const containerRect = container.getBoundingClientRect();
        const elRect = el.getBoundingClientRect();
        const dx = (elRect.left - containerRect.left) - anchor.left;
        const dy = (elRect.top - containerRect.top) - anchor.top;
        if (dx !== 0 || dy !== 0) {
          container.scrollTo({
            left: container.scrollLeft + dx,
            top: container.scrollTop + dy,
            behavior: "auto",
          });
        }
      } else {
        el.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "center" });
      }
      explodedAnchorSnapshotRef.current = null;
    });
    setExplodedAnchorId(null);
  }, [explodedAnchorId]);

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


  const initialHydrationLoading = profileLoading || !directoryLoaded;
  const activePrintFocusId = printTargetNodeId || focusId;
  const activePrintNode = getNode(nodeList, activePrintFocusId);
  const openNodeEditFromHierarchy = useCallback((nodeId) => {
    setEditNodeId(nodeId);
    setOpenDialog({ type: "edit-node" });
  }, []);
  const openNodeBookPrintDialog = useCallback((nodeId) => {
    setPrintDialogMode("book");
    setPrintTargetNodeId(nodeId || "");
    setPosterConfirmed(false);
    setPrintDialogOpen(true);
  }, []);
  const openNodePosterPrintDialog = useCallback((nodeId) => {
    setPrintDialogMode("poster");
    setPrintTargetNodeId(nodeId || "");
    setPosterConfirmed(false);
    setPrintDialogOpen(true);
  }, []);
  const focusNodeAndPersistPrimary = useCallback((nodeId) => {
    const nextFocusId = String(nodeId || "");
    if (!nextFocusId) return;
    setFocusId(nextFocusId);

    let baseHome = homeScreen && typeof homeScreen === "object" ? homeScreen : null;
    if (!baseHome) {
      try {
        const stored = localStorage.getItem("homeScreen");
        const parsed = stored ? JSON.parse(stored) : null;
        if (parsed && typeof parsed === "object") baseHome = parsed;
      } catch {
        baseHome = null;
      }
    }
    if (!baseHome) return;

    const nextHome = { ...baseHome, focusId: nextFocusId };
    setHomeScreen(nextHome);
    try {
      localStorage.setItem("homeScreen", JSON.stringify(nextHome));
    } catch {
      // ignore storage errors
    }
    apiRequest("/api/auth/me", {
      method: "PATCH",
      body: JSON.stringify({
        homeScreen: nextHome,
        ...getTabularPrefsPayload(),
      }),
    }).catch(() => { });
  }, [apiRequest, getTabularPrefsPayload, homeScreen]);


  if (initialHydrationLoading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", flexDirection: "column", gap: 16, color: "#888", fontFamily: "inherit" }}>
        <div style={{ width: 32, height: 32, border: "3px solid #ddd", borderTopColor: "#555", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
        <span style={{ fontSize: 14 }}>Loading…</span>
      </div>
    );
  }

  return (
    <div style={viewMode === "hierarchy" ? {} : viewMode === "tabular" ? {} : { paddingBottom: 120 }} data-lpignore="true">
      {homeAnimating && (
        <div className="home-anim-overlay" style={{ transformOrigin: homeAnimOrigin }} />
      )}
      <div className="app-header">
        <div style={{ maxWidth: "90%", margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button
              type="button"
              aria-label="Go home"
              title="Home"
              style={{ background: "none", border: "none", padding: 0, cursor: "pointer", lineHeight: 0 }}
              onClick={() => {
                if (homeScreen) {
                  restoreHomeScreen(homeScreen);
                } else {
                  setViewMode("hierarchy");
                }
              }}
            >
              <img src="/emplus-logo.png" alt="EMPlus" style={{ height: 50, width: "auto", margin: "-10px", borderRadius: "25px" }} />
            </button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
            <div style={{ textAlign: "center" }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: "#1a1a2e" }}>{clientDisplayName || toSentenceCase(clientId)}</div>
              <div style={{ fontSize: 13, color: "#64748b" }}>Entity Dashboard</div>
            </div>
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
                borderRight: "1px solid #cbd5e1",
                background: viewMode === "directory" ? "#1e293b" : "#fff",
                color: viewMode === "directory" ? "#fff" : "#475569",
              }}
            >
              <LayoutList size={14} /> Directory
            </button>
            <button
              type="button"
              aria-label="Tabular view"
              title="Tabular"
              onClick={() => setViewMode("tabular")}
              style={{
                display: "flex", alignItems: "center", gap: 5, padding: "5px 12px",
                fontSize: 13, fontWeight: 500, cursor: "pointer", border: "none",
                background: viewMode === "tabular" ? "#1e293b" : "#fff",
                color: viewMode === "tabular" ? "#fff" : "#475569",
              }}
            >
              <BookOpen size={14} /> Tabular
            </button>
          </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <Button
              ref={homeButtonRef}
              type="button"
              variant="outline"
              className="btn-icon"
              aria-label="Go home"
              title="Home"
              onClick={() => {
                if (homeScreen) {
                  restoreHomeScreen(homeScreen);
                } else {
                  setViewMode("hierarchy");
                }
              }}
            >
              <Home size={18} />
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
                      const screen = captureHomeScreen();
                      setHomeScreen(screen);
                      try { localStorage.setItem("homeScreen", JSON.stringify(screen)); } catch { }
                      // Persist to user record so it survives across sessions/devices
                      apiRequest("/api/auth/me", {
                        method: "PATCH",
                        body: JSON.stringify({
                          homeScreen: screen,
                          ...getTabularPrefsPayload(),
                        }),
                      }).catch((err) => {
                        setRemoteStatus("error");
                        setRemoteError(err.message || "Unable to save home screen.");
                      });
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
                      setUploadPreview(null);
                      setUploadPreviewLoading(false);
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
                  {myRole === "admin" && (
                    <button
                      className="settings-menu-item"
                      onClick={() => {
                        setSettingsOpen(false);
                        setCloneClientDraft("");
                        setCloneClientError("");
                        setCloneClientResult(null);
                        setCloneClientOpen(true);
                      }}
                    >
                      <GitFork size={15} />
                      Clone this client
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

        {exitPromptOpen && (
          <div
            style={{
              position: "fixed",
              top: 60,
              left: 12,
              zIndex: 1400,
              background: "#fff",
              border: "1px solid #fca5a5",
              borderRadius: 10,
              boxShadow: "0 8px 24px rgba(15, 23, 42, 0.18)",
              padding: "10px 12px",
              minWidth: 280,
            }}
          >
            <div style={{ fontSize: 13, fontWeight: 700, color: "#991b1b", marginBottom: 8 }}>
              Leave EMPlus?
            </div>
            <div style={{ fontSize: 12, color: "#7f1d1d", marginBottom: 10 }}>
              Are you trying to exit EMPlus?
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Button type="button" variant="outline" onClick={handleExitPromptExit}>
                Exit
              </Button>
              <Button type="button" onClick={handleExitPromptReturn}>
                Return to EMPlus
              </Button>
            </div>
          </div>
        )}

      <div className="app-content">

        <Dialog open={uploadOpen} onOpenChange={() => setUploadOpen(false)}>
          <DialogContent style={{ maxWidth: uploadPreview && uploadStatus !== "success" ? 700 : 520 }}>
            <DialogHeader>
              <DialogTitle>Import</DialogTitle>
            </DialogHeader>

            {/* ── Configure view ── */}
            {!uploadPreview && uploadStatus !== "success" && (
              <div style={{ display: "grid", gap: 12 }}>
                <div>
                  <Label>Import type</Label>
                  <div style={{ display: "flex", gap: 20, marginTop: 8, flexWrap: "wrap" }}>
                    {[
                      { value: "entity", label: "Entities" },
                      { value: "person", label: "Persons" },
                      { value: "ownership", label: "Ownerships" },
                      { value: "details", label: "Details" },
                    ].map(({ value, label }) => (
                      <label key={value} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 14 }}>
                        <input
                          type="radio"
                          name="upload-type"
                          value={value}
                          checked={uploadType === value}
                          onChange={() => {
                            setUploadType(value);
                            if (value === "entity") setUploadKind("entity");
                            if (value === "person") setUploadKind("person");
                            setUploadSummary(null);
                            setUploadError("");
                            setUploadStatus("idle");
                            setUploadPreview(null);
                          }}
                        />
                        {label}
                      </label>
                    ))}
                  </div>
                </div>
                <div>
                  <Label htmlFor="csv-upload">File</Label>
                  <Input
                    id="csv-upload"
                    type="file"
                    accept=".csv,.xlsx,.xls,text/csv"
                    onChange={(event) => {
                      const file = event.target.files?.[0] || null;
                      setUploadFile(file);
                      setUploadSummary(null);
                      setUploadError("");
                      setUploadStatus("idle");
                      setUploadPreview(null);
                      if (!file) {
                        setUploadDetected("");
                        return;
                      }
                      const ext = (file.name || "").split(".").pop().toLowerCase();
                      if (ext === "xlsx" || ext === "xls") {
                        return; // keep whatever radio button type the user selected
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
                {uploadStatus === "error" && uploadError && (
                  <div style={{ color: "#dc2626", fontSize: 14 }}>{uploadError}</div>
                )}
              </div>
            )}

            {/* ── Preview view ── */}
            {uploadPreview && uploadStatus !== "success" && (
              <div style={{ display: "grid", gap: 10 }}>
                <div style={{ fontSize: 14, color: "#374151" }}>
                  Found <strong>{uploadPreview.total}</strong> row{uploadPreview.total !== 1 ? "s" : ""}
                  {uploadPreview.skipped > 0 && <>, {uploadPreview.skipped} blank or invalid skipped</>}.
                  {uploadPreview.truncated && <> Showing first {uploadPreview.rows.length}.</>}
                </div>
                <div style={{ border: "1px solid #e5e7eb", borderRadius: 8, overflow: "hidden" }}>
                  <div style={{ overflowX: "auto", maxHeight: 320, overflowY: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr>
                          {uploadPreview.headers.map((h) => (
                            <th key={h} style={{ padding: "6px 10px", textAlign: "left", background: "#f9fafb", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap", fontWeight: 600, color: "#374151", position: "sticky", top: 0 }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {uploadPreview.rows.map((row, i) => (
                          <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                            {uploadPreview.headers.map((h) => (
                              <td key={h} style={{ padding: "4px 10px", color: "#374151", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                {row[h] ?? ""}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
                {uploadPreview.mapping?.guessed && Object.keys(uploadPreview.mapping.guessed).length > 0 && (
                  <div style={{ fontSize: 12, color: "#b45309", display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                    <span style={{ fontWeight: 600 }}>Fuzzy-matched columns:</span>
                    {Object.entries(uploadPreview.mapping.guessed).map(([h, f]) => (
                      <span key={h} style={{ background: "#fef3c7", padding: "1px 6px", borderRadius: 4 }}>"{h}" → {f}</span>
                    ))}
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
              </div>
            )}

            {/* ── Success view ── */}
            {uploadStatus === "success" && uploadSummary && (
              <div style={{ display: "grid", gap: 10 }}>
                <div style={{ color: "#16a34a", fontSize: 14 }}>
                  {uploadType === "ownership" ? (
                    <>Imported {uploadSummary.imported || 0} ownerships. Skipped {uploadSummary.skipped || 0}.</>
                  ) : uploadType === "details" ? (
                    <>Updated {uploadSummary.updated || 0} records.{uploadSummary.notFound > 0 ? ` ${uploadSummary.notFound} not matched by name.` : ""}{uploadSummary.skipped > 0 ? ` ${uploadSummary.skipped} rows skipped.` : ""}</>
                  ) : (
                    <>Imported {uploadSummary.total || 0} rows ({uploadSummary.entities || 0} entities,
                      {" "}{uploadSummary.persons || 0} persons). Skipped {uploadSummary.skipped || 0}.</>
                  )}
                </div>
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
            )}

            <DialogFooter>
              {uploadStatus === "success" ? (
                <Button type="button" variant="outline" onClick={() => setUploadOpen(false)}>
                  Close
                </Button>
              ) : uploadPreview ? (
                <>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => { setUploadPreview(null); setUploadStatus("idle"); setUploadError(""); }}
                    disabled={uploadStatus === "uploading"}
                  >
                    Back
                  </Button>
                  <Button
                    type="button"
                    onClick={handleUploadCsv}
                    disabled={uploadStatus === "uploading" || uploadPreview.total === 0}
                  >
                    {uploadStatus === "uploading"
                      ? <><Loader2 size={14} className="animate-spin" style={{ marginRight: 6 }} />Importing…</>
                      : `Import ${uploadPreview.total} row${uploadPreview.total !== 1 ? "s" : ""}`}
                  </Button>
                </>
              ) : (
                <>
                  <Button type="button" variant="outline" onClick={() => setUploadOpen(false)}>
                    Close
                  </Button>
                  <Button type="button" variant="outline" onClick={downloadTemplate}>
                    Download Template
                  </Button>
                  <Button
                    type="button"
                    onClick={handlePreview}
                    disabled={!uploadFile || uploadPreviewLoading}
                  >
                    {uploadPreviewLoading
                      ? <><Loader2 size={14} className="animate-spin" style={{ marginRight: 6 }} />Loading…</>
                      : "Preview"}
                  </Button>
                </>
              )}
            </DialogFooter>
          </DialogContent>
        </Dialog>

        {remoteStatus === "error" && (
          <div style={{ textAlign: "center", color: "#dc2626", marginBottom: 12 }}>
            {remoteError || "Unable to load directory"}
          </div>
        )}

        {viewMode === "hierarchy" && (
          <>
            <div
              className={`hierarchy-vertical${directOwners.length === 0 ? " hierarchy-vertical--no-parents" : ""}`}
              ref={hierarchyContainerRef}
              onPointerDown={onHierarchyPointerDown}
              onPointerMove={onHierarchyPointerMove}
              onPointerUp={onHierarchyPointerUp}
              onPointerCancel={onHierarchyPointerCancel}
              onClickCapture={onHierarchyClickCapture}
            >
              <div className="hv-stage">

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
                          [...getOwnersOf(relList, focusId).map((item) => (
                            <HvNeighborBox
                              key={item.nodeId}
                              item={item}
                              nodeList={nodeList}
                              relList={relList}
                              explodedNodes={explodedNodes}
                              onExplode={() => { }}
                              onExplodeAll={() => { }}
                              onFocus={setFocusId}
                              onFocusPrimary={focusNodeAndPersistPrimary}
                              onEdit={openNodeEditFromHierarchy}
                              onPrintBook={openNodeBookPrintDialog}
                              onPrintPoster={openNodePosterPrintDialog}
                              showExplodeControls={false}
                            />
                          )),
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
                    const totalDesc = ownedCount > 0 ? getAllDescendants(relList, focusId).size : 0;
                    const hasMore = totalDesc > ownedCount;
                    return (ownerCount > 0 || ownedCount > 0) ? (
                      <div className="hv-focus-counts">
                        {ownerCount > 0 && (
                          <span>{ownerCount} {ownerCount === 1 ? "Owner" : "Owners"}</span>
                        )}
                        {ownedCount > 0 && (
                          <span>
                            {ownedCount} direct {ownedCount === 1 ? "subsidiary" : "subsidiaries"}
                            {hasMore && ` · ${totalDesc} total in tree`}
                          </span>
                        )}
                      </div>
                    ) : null;
                  })()}
                  <div className="hv-focus-actions" onClick={(e) => e.stopPropagation()}>
                    <button
                      className="hv-focus-action-btn"
                      type="button"
                      title="Edit"
                      aria-label="Edit"
                      onClick={() => openNodeEditFromHierarchy(focusId)}
                    >
                      <Pencil size={13} />
                    </button>
                    <button
                      className="hv-focus-action-btn"
                      type="button"
                      title="Print Book Pages (PDF)"
                      aria-label="Print Book Pages"
                      onClick={() => openNodeBookPrintDialog(focusId)}
                    >
                      <BookOpen size={13} />
                    </button>
                    <button
                      className="hv-focus-action-btn"
                      type="button"
                      title="Print Org Chart Poster"
                      aria-label="Print Org Chart Poster"
                      onClick={() => openNodePosterPrintDialog(focusId)}
                    >
                      <GitFork size={13} />
                    </button>
                  </div>
                  {getOwnedBy(relList, focusId).length > 0 && (
                    explodedNodes.size > 0 ? (
                      <button
                        className="hv-focus-explode-all-btn"
                        title="Collapse all — return to first level"
                        onClick={(e) => {
                          e.stopPropagation();
                          snapshotExplodedAnchorPosition(focusId, focusBoxRef.current);
                          setExplodedNodes(new Set());
                          setExplodedAnchorId(focusId);
                        }}
                      >
                        <ChevronsUp size={14} />
                      </button>
                    ) : (
                      <button
                        className="hv-focus-explode-all-btn"
                        title="Expand all descendants"
                        onClick={(e) => {
                          e.stopPropagation();
                          snapshotExplodedAnchorPosition(focusId, focusBoxRef.current);
                          const allDesc = getAllDescendants(relList, focusId, new Set([focusId]));
                          setExplodedNodes(prev => new Set([...prev, ...allDesc]));
                          setExplodedAnchorId(focusId);
                        }}
                      >
                        <ChevronsDown size={14} />
                      </button>
                    )
                  )}
                </div>

                <div className="hv-below">
                  {/* ── connector line from focus down to owned ── */}
                  {getOwnedBy(relList, focusId).length > 0 && (
                    <div className="hv-connector" />
                  )}

                  {/* ── OWNED BY (below the focus box) — org-chart tree ── */}
                  {getOwnedBy(relList, focusId).length > 0 && (() => {
                    const items = getOwnedBy(relList, focusId);
                    const visitedIds = new Set([focusId]);
                    const cws = items.map(item => computeColWidth(item.nodeId, relList, explodedNodes, visitedIds));
                    const totalW = cws.reduce((s, w) => s + w, 0) + Math.max(0, items.length - 1) * ORG_NODE_GAP;
                    const barLeft = items.length > 1 ? cws[0] / 2 : 0;
                    const barW = items.length > 1 ? totalW - cws[0] / 2 - cws[cws.length - 1] / 2 : 0;
                    const handleExplode = (nodeId, anchorEl) => {
                      snapshotExplodedAnchorPosition(nodeId, anchorEl);
                      setExplodedNodes(prev => {
                        const next = new Set(prev);
                        next.has(nodeId) ? next.delete(nodeId) : next.add(nodeId);
                        return next;
                      });
                      setExplodedAnchorId(nodeId);
                    };
                    const handleExplodeAll = (nodeId, anchorEl) => {
                      snapshotExplodedAnchorPosition(nodeId, anchorEl);
                      const desc = getAllDescendants(relList, nodeId, new Set([focusId]));
                      setExplodedNodes(prev => new Set([...prev, nodeId, ...desc]));
                      setExplodedAnchorId(nodeId);
                    };
                    return (
                      <>
                        <div className="hv-section-header" style={{ alignSelf: 'center' }}>
                          <span className="hv-section-label">Owns</span>
                        </div>
                        <div style={{ margin: '0 auto', width: 'max-content', display: 'flex', flexDirection: 'column', alignItems: 'center', paddingBottom: 24 }}>
                          {/* Top-level horizontal connector bar (only when 2+ children) */}
                          {items.length > 1 && (
                            <div style={{ position: 'relative', width: totalW, height: 2, flexShrink: 0 }}>
                              <div style={{ position: 'absolute', top: 0, left: barLeft, width: barW, height: 2, background: ORG_LINE_COLOR }} />
                            </div>
                          )}
                          {/* Top-level children row */}
                          <div style={{ display: 'flex', alignItems: 'flex-start', gap: ORG_NODE_GAP }}>
                            {items.map(item => (
                              <OrgChartTreeNode
                                key={item.nodeId} item={item} nodeList={nodeList} relList={relList}
                                explodedNodes={explodedNodes} onExplode={handleExplode}
                                onExplodeAll={handleExplodeAll} onFocus={setFocusId}
                                onFocusPrimary={focusNodeAndPersistPrimary}
                                onEdit={openNodeEditFromHierarchy}
                                onPrintBook={openNodeBookPrintDialog}
                                onPrintPoster={openNodePosterPrintDialog}
                                visitedIds={visitedIds} showTopConnector={items.length > 1}
                              />
                            ))}
                          </div>
                        </div>
                      </>
                    );
                  })()}
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
              </div>{/* end hv-stage */}
            </div>
            <OrgChartMinimap
              containerRef={hierarchyContainerRef}
              watchKey={`${focusId}-${explodedNodes.size}`}
            />
          </>
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
                          <div className="directory-item-actions" onClick={(e) => e.stopPropagation()}>
                            <button
                              className="directory-edit-btn"
                              title="Edit"
                              aria-label="Edit"
                              onClick={() => {
                                setEditNodeId(n.id);
                                setOpenDialog({ type: "edit-node" });
                              }}
                            >
                              <Pencil size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Print Book Pages (PDF)"
                              aria-label="Print Book Pages"
                              onClick={() => openNodeBookPrintDialog(n.id)}
                            >
                              <BookOpen size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Print Org Chart Poster"
                              aria-label="Print Org Chart Poster"
                              onClick={() => openNodePosterPrintDialog(n.id)}
                            >
                              <GitFork size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Focus in Hierarchy"
                              aria-label="Focus in Hierarchy"
                              onClick={() => {
                                setFocusId(n.id);
                                setViewMode("hierarchy");
                              }}
                            >
                              <Crosshair size={13} />
                            </button>
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
                          <div className="directory-item-actions" onClick={(e) => e.stopPropagation()}>
                            <button
                              className="directory-edit-btn"
                              title="Edit"
                              aria-label="Edit"
                              onClick={() => {
                                setEditNodeId(n.id);
                                setOpenDialog({ type: "edit-node" });
                              }}
                            >
                              <Pencil size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Print Book Pages (PDF)"
                              aria-label="Print Book Pages"
                              onClick={() => openNodeBookPrintDialog(n.id)}
                            >
                              <BookOpen size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Print Org Chart Poster"
                              aria-label="Print Org Chart Poster"
                              onClick={() => openNodePosterPrintDialog(n.id)}
                            >
                              <GitFork size={13} />
                            </button>
                            <button
                              className="directory-edit-btn"
                              title="Focus in Hierarchy"
                              aria-label="Focus in Hierarchy"
                              onClick={() => {
                                setFocusId(n.id);
                                setViewMode("hierarchy");
                              }}
                            >
                              <Crosshair size={13} />
                            </button>
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

        {viewMode === "tabular" && (
          <div className="tabular-grid">
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

            <div className="tabular-toolbar">
              <div className="tabular-toolbar-left">
                <select
                  className="tabular-view-select"
                  value={selectedTabularViewId}
                  onChange={(e) => setSelectedTabularViewId(e.target.value)}
                >
                  <option value={DEFAULT_TABULAR_VIEW_ID}>Default</option>
                  {tabularViews.map((v) => (
                    <option key={v.id} value={v.id}>{v.name}</option>
                  ))}
                </select>
                <Button type="button" variant="outline" onClick={openTabularViewManager}>
                  Customize
                </Button>
              </div>
              <Button
                type="button"
                onClick={saveAllTableRows}
                disabled={pendingTableKeys.size === 0 || tableSavingKeys.size > 0}
              >
                Save All ({pendingTableKeys.size})
              </Button>
            </div>

            <Card className="tabular-card">
              <CardContent style={{ padding: 0 }}>
                <div className="tabular-wrap">
                  <table
                    className="tabular-table"
                    style={{ minWidth: `${Math.max(visibleTabularColumns.length * 150, 1000)}px` }}
                  >
                    <thead>
                      <tr>
                        {visibleTabularColumns.map((column) => (
                          <th key={column.key}>{column.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {tableRows.length === 0 && (
                        <tr>
                          <td colSpan={visibleTabularColumns.length} className="tabular-empty">
                            No matching rows.
                          </td>
                        </tr>
                      )}
                      {tableRows.map((row) => {
                        const rowNode = row.isNew ? row.node : (tableDrafts[row.key] || row.node);
                        const isSaving = tableSavingKeys.has(row.key);
                        const isDirty = pendingTableKeys.has(row.key);
                        const rowError = tableRowErrors[row.key];
                        const resolvedId = row.isNew
                          ? makeNodeId(rowNode.kind, rowNode.name || "")
                          : makeNodeId(rowNode.kind, rowNode.name || "", row.key);
                        return (
                          <tr key={row.key} className={rowError ? "tabular-row-error" : undefined}>
                            {visibleTabularColumns.map((column) =>
                              renderTabularCell(column, { row, rowNode, isSaving, isDirty, rowError, resolvedId })
                            )}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>

            {tabularViewDialogOpen && (
              <Dialog open={tabularViewDialogOpen} onOpenChange={setTabularViewDialogOpen}>
                <DialogContent style={{ width: "min(760px, 92vw)", maxWidth: "none" }}>
                  <DialogHeader>
                    <DialogTitle>Tabular View Columns</DialogTitle>
                  </DialogHeader>
                  <div className="tabular-view-editor">
                    <div className="form-row">
                      <label className="form-label">View Name</label>
                      <input
                        className="form-input"
                        value={tabularViewDraft.name}
                        onChange={(e) => {
                          setTabularViewNameError("");
                          setTabularViewDraft((prev) => ({ ...prev, name: e.target.value }));
                        }}
                        placeholder="My custom tabular view"
                        autoComplete="off"
                        data-lpignore="true"
                      />
                      {tabularViewNameError && (
                        <div className="dup-warning" style={{ marginTop: 8 }}>
                          {tabularViewNameError}
                        </div>
                      )}
                    </div>
                    <div className="tabular-view-columns">
                      <div className="tabular-view-section-title" style={{ marginTop: 10, marginLeft: 10 }}>In View</div>
                      {tabularViewDraft.columnOrder.map((key, idx) => {
                        const col = allTabularColumns.find((c) => c.key === key);
                        if (!col) return null;
                        return (
                          <div key={key} className="tabular-view-column-row">
                            <label className="tabular-view-column-toggle">
                              <input
                                type="checkbox"
                                checked
                                onChange={() => toggleTabularDraftSelected(key)}
                              />
                              <span>{col.label}</span>
                            </label>
                            <div className="tabular-view-column-actions">
                              <Button
                                type="button"
                                variant="outline"
                                onClick={() => moveTabularDraftColumn(key, -1)}
                                disabled={idx === 0}
                              >
                                Up
                              </Button>
                              <Button
                                type="button"
                                variant="outline"
                                onClick={() => moveTabularDraftColumn(key, 1)}
                                disabled={idx === tabularViewDraft.columnOrder.length - 1}
                              >
                                Down
                              </Button>
                            </div>
                          </div>
                        );
                      })}
                      <div className="tabular-view-section-title" style={{ marginBottom: 4, marginTop: 10, marginLeft: 10 }}>Available Fields</div>
                      {availableTabularColumns.map((col) => (
                        <div key={col.key} className="tabular-view-column-row">
                          <label className="tabular-view-column-toggle">
                            <input
                              type="checkbox"
                              checked={false}
                              onChange={() => toggleTabularDraftSelected(col.key)}
                            />
                            <span>{col.label}</span>
                          </label>

                        </div>
                      ))}
                    </div>
                  </div>
                  <DialogFooter style={{ justifyContent: "space-between" }}>
                    <div style={{ display: "flex", gap: 8 }}>
                      <Button type="button" variant="outline" onClick={saveTabularViewAsNew}>
                        Save As New
                      </Button>
                      <Button
                        type="button"
                        variant="outline"
                        onClick={updateCurrentTabularView}
                      >
                        Update Selected
                      </Button>
                      <Button
                        type="button"
                        variant="outline"
                        onClick={deleteCurrentTabularView}
                        disabled={selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID}
                      >
                        Delete Selected
                      </Button>
                    </div>
                    <Button type="button" variant="secondary" onClick={() => setTabularViewDialogOpen(false)}>
                      Close
                    </Button>
                  </DialogFooter>
                </DialogContent>
              </Dialog>
            )}
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
              <DialogHeader style={{ marginBottom: '24px', marginLeft: 0 }}>
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
              <DialogHeader style={{ marginBottom: '24px', marginLeft: 0 }}>
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

        {/* ── Print dialogs ── */}
        <Dialog open={printDialogOpen} onOpenChange={(open) => {
          if (!open) {
            setPosterConfirmed(false);
            setPrintDialogMode(null);
            setPrintTargetNodeId("");
          }
          setPrintDialogOpen(open);
        }}>
          <DialogContent style={{ maxWidth: 380 }}>
            <DialogHeader>
              <DialogTitle>
                {printDialogMode === "poster" ? "Print Org Chart Poster" : "Print Entity Book"}
              </DialogTitle>
            </DialogHeader>
            <div style={{ padding: "8px 0 16px", display: "flex", flexDirection: "column", gap: 12 }}>
              {printDialogMode === "book" && (
                <>
                  <p style={{ fontSize: 13, color: "#64748b", margin: 0 }}>
                    {printTargetNodeId
                      ? "Select page types to include for the selected entity."
                      : viewMode === "hierarchy"
                      ? "Select page types to include."
                      : "Select page types to include for each entity in the current view. Pages are interleaved: hierarchy then detail for each entity."}
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
                    let scopeNodes;
                    if (printTargetNodeId) {
                      scopeNodes = activePrintNode ? [activePrintNode] : [];
                    } else if (viewMode === "hierarchy") {
                      if (explodedNodes.size > 0) {
                        const visibleIds = new Set([focusId, ...explodedNodes]);
                        scopeNodes = nodeList.filter(n => visibleIds.has(n.id));
                      } else {
                        scopeNodes = nodeList.filter(n => n.id === focusId);
                      }
                    } else {
                      scopeNodes = dirSearch.trim()
                        ? [...filteredEntityNodes, ...filteredPersonNodes]
                        : nodeList;
                    }
                    return (
                      <p style={{ fontSize: 12, color: "#94a3b8", margin: 0 }}>
                        {scopeNodes.length} {scopeNodes.length === 1 ? "item" : "items"} in scope
                      </p>
                    );
                  })()}
                </>
              )}

              {printDialogMode === "poster" && (() => {
                const { pages, cols, rows } = estimatePosterPageCount(activePrintFocusId, relList);
                return (
                  <>
                    {printTargetNodeId && (
                      <p style={{ fontSize: 12, color: "#64748b", margin: 0 }}>
                        Selected focus: <strong>{activePrintNode?.name || activePrintFocusId}</strong>
                      </p>
                    )}
                    <p style={{ fontSize: 12, color: "#64748b", margin: 0 }}>
                      Tiled landscape pages showing the complete ownership tree with every level fully horizontal. Print and assemble side-by-side.
                    </p>
                    <p style={{ fontSize: 12, color: "#94a3b8", margin: 0 }}>
                      Estimated: <strong style={{ color: "#475569" }}>{pages} {pages === 1 ? "page" : "pages"}</strong> ({cols} wide × {rows} tall, A4 landscape)
                    </p>
                    <label style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 13, cursor: "pointer", lineHeight: 1.4 }}>
                      <input
                        type="checkbox"
                        style={{ marginTop: 2, flexShrink: 0 }}
                        checked={posterConfirmed}
                        onChange={(e) => {
                          setPosterConfirmed(e.target.checked);
                          if (e.target.checked) {
                            setPrintHierarchy(false);
                            setPrintDetail(false);
                          }
                        }}
                      />
                      I understand — generate poster ({pages} {pages === 1 ? "page" : "pages"})
                    </label>
                  </>
                );
              })()}
            </div>
            <DialogFooter style={{ gap: 8 }}>
              <Button variant="outline" onClick={() => setPrintDialogOpen(false)}>Cancel</Button>
              {printDialogMode === "poster" && (
                <Button
                  disabled={isPdfExporting || !posterConfirmed || !activePrintFocusId}
                  variant="outline"
                  style={posterConfirmed ? { borderColor: "#f59e0b", color: "#92400e", background: "#fffbeb" } : {}}
                  onClick={async () => {
                    setPrintDialogOpen(false);
                    setIsPdfExporting(true);
                    const focusNodeName = activePrintNode?.name || activePrintFocusId;
                    const safeBase = `${focusNodeName}-org-chart-poster`.replace(/[^\w\s.-]/g, "").replace(/\s+/g, "_");
                    setExportResultAndRevoke({ status: "exporting", fileName: `${safeBase}.pdf` });
                    pdfCancelRef.current = false;
                    setPdfProgress(null);
                    try {
                      const result = await generateOrgChartPoster({
                        focusId: activePrintFocusId,
                        nodeList,
                        relList,
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
                      setPosterConfirmed(false);
                    }
                  }}
                >
                  Generate Poster
                </Button>
              )}
              {printDialogMode === "book" && <Button
                disabled={(!printHierarchy && !printDetail) || isPdfExporting || (printTargetNodeId && !activePrintNode)}
                onClick={async () => {
                  setPrintDialogOpen(false);
                  setIsPdfExporting(true);
                  const suffix = printHierarchy && printDetail ? "full" : printHierarchy ? "hierarchy" : "detail";
                  let printNodes;
                  let baseLabel;
                  if (printTargetNodeId) {
                    printNodes = activePrintNode ? [activePrintNode] : [];
                    const nodeName = activePrintNode?.name || printTargetNodeId;
                    baseLabel = `${nodeName}-${suffix}`;
                  } else if (viewMode === "hierarchy") {
                    if (explodedNodes.size > 0) {
                      const visibleIds = new Set([focusId, ...explodedNodes]);
                      printNodes = nodeList.filter(n => visibleIds.has(n.id));
                      const focusNodeName = nodeList.find(n => n.id === focusId)?.name || focusId;
                      baseLabel = `${focusNodeName}-expanded-${suffix}`;
                    } else {
                      printNodes = nodeList.filter(n => n.id === focusId);
                      const focusNodeName = nodeList.find(n => n.id === focusId)?.name || focusId;
                      baseLabel = `${focusNodeName}-${suffix}`;
                    }
                  } else {
                    printNodes = dirSearch.trim()
                      ? [...filteredEntityNodes, ...filteredPersonNodes]
                      : nodeList;
                    baseLabel = `${clientDisplayName || toSentenceCase(clientId)}-entity-book-${suffix}`;
                  }
                  const safeBase = baseLabel.replace(/[^\w\s.-]/g, "").replace(/\s+/g, "_");
                  setExportResultAndRevoke({ status: "exporting", fileName: `${safeBase}.pdf` });
                  pdfCancelRef.current = false;
                  setPdfProgress(null);
                  try {
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
                  : "Save PDF"}
              </Button>}
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

        <Dialog open={cloneClientOpen} onOpenChange={(v) => { if (!v && !cloneClientBusy) setCloneClientOpen(false); }}>
          <DialogContent style={{ width: "min(440px, 92vw)", maxWidth: "none" }}>
            <DialogHeader style={{ marginBottom: 16 }}>
              <DialogTitle>Clone this client</DialogTitle>
            </DialogHeader>
            <div style={{ display: "grid", gap: 14, fontSize: 14 }}>
              {cloneClientResult ? (
                <div style={{ display: "grid", gap: 10 }}>
                  <div style={{ color: "#16a34a", fontWeight: 600 }}>Client cloned successfully.</div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <tbody>
                      <tr>
                        <td style={{ padding: "4px 12px 4px 0", fontWeight: 600, color: "#374151" }}>New client ID</td>
                        <td style={{ padding: "4px 0", fontFamily: "monospace" }}>{cloneClientResult.newClientId}</td>
                      </tr>
                      <tr>
                        <td style={{ padding: "4px 12px 4px 0", fontWeight: 600, color: "#374151" }}>Admin login</td>
                        <td style={{ padding: "4px 0", fontFamily: "monospace" }}>{cloneClientResult.newLoginId}</td>
                      </tr>
                      <tr>
                        <td style={{ padding: "4px 12px 4px 0", fontWeight: 600, color: "#374151" }}>DD fields copied</td>
                        <td style={{ padding: "4px 0" }}>{cloneClientResult.ddFieldsCloned}</td>
                      </tr>
                      <tr>
                        <td style={{ padding: "4px 12px 4px 0", fontWeight: 600, color: "#374151" }}>Password</td>
                        <td style={{ padding: "4px 0" }}>Same as your current password</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              ) : (
                <>
                  <div style={{ color: "#6b7280" }}>
                    Creates a new empty client with a single admin user. The DataDictionary from this client will be copied. The new admin can log in with the same password as your current account.
                  </div>
                  <div>
                    <Label htmlFor="clone-client-id">New client ID</Label>
                    <Input
                      id="clone-client-id"
                      value={cloneClientDraft}
                      onChange={(e) => { setCloneClientDraft(e.target.value); setCloneClientError(""); }}
                      placeholder="e.g. acme-corp"
                      autoComplete="off"
                      disabled={cloneClientBusy}
                    />
                  </div>
                  {cloneClientDraft.trim() && (
                    <div style={{ fontSize: 12, color: "#6b7280" }}>
                      New admin login:{" "}
                      <span style={{ fontFamily: "monospace", color: "#111827" }}>
                        {myLoginId.split("-")[0]}-{cloneClientDraft.trim().toLowerCase().replace(/[^a-z0-9-]/g, "-").replace(/(^-|-$)/g, "")}
                      </span>
                    </div>
                  )}
                  {cloneClientError && (
                    <div style={{ color: "#dc2626", fontSize: 13 }}>{cloneClientError}</div>
                  )}
                </>
              )}
            </div>
            <DialogFooter style={{ marginTop: 20 }}>
              {cloneClientResult ? (
                <Button onClick={() => setCloneClientOpen(false)}>Close</Button>
              ) : (
                <>
                  <Button variant="secondary" onClick={() => setCloneClientOpen(false)} disabled={cloneClientBusy}>
                    Cancel
                  </Button>
                  <Button
                    onClick={async () => {
                      const val = cloneClientDraft.trim();
                      if (!val) { setCloneClientError("Please enter a client ID."); return; }
                      setCloneClientBusy(true);
                      setCloneClientError("");
                      try {
                        const result = await apiRequest("/api/admin/clone-client", {
                          method: "POST",
                          body: JSON.stringify({ newClientId: val }),
                        });
                        setCloneClientResult(result);
                      } catch (err) {
                        setCloneClientError(err.message || "Clone failed.");
                      } finally {
                        setCloneClientBusy(false);
                      }
                    }}
                    disabled={cloneClientBusy || !cloneClientDraft.trim()}
                  >
                    {cloneClientBusy ? <><Loader2 size={15} className="animate-spin" style={{ marginRight: 6 }} />Cloning…</> : "Clone client"}
                  </Button>
                </>
              )}
            </DialogFooter>
          </DialogContent>
        </Dialog>

      </div>{/* end app-content */}
    </div>
  );
}

// EOF
