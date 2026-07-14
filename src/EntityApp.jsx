import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Card, CardContent } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { Link, Users, Building2, Plus, Pencil, Trash2, ChevronRight, ChevronDown, ChevronsDown, ChevronsUp, Upload, X, Search, Settings, LogOut, GitFork, LayoutList, Home, Download, BookOpen, User, UserPlus, Loader2, Crosshair, Filter, FileSpreadsheet, Save } from "lucide-react";
import { generateEntityPdf, generateEntityBook, generateEntityBookInterleaved, estimatePosterPageCount, generateOrgChartPoster } from "./utils/generateEntityPdf";
import ExportDialog from "./components/ExportDialog";
import { normalizePhone, formatPhone, normalizeDateInput } from "./utils/helpers";
import { ENTITY_OWNERSHIP_SUMMARY_FIELD, getEntityOwnershipSummary } from "./utils/ownershipSummary";
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

const renderDdField = (field, value, onChange, { apiBase, token, node, nodeList, relList, asOfDate, ownershipTimeline } = {}) => {
  const { fieldId, prompt, dataType, multiValue, validValues, phoneTypes } = field;

  if (field?._virtual) {
    const computed = getEntityOwnershipSummary(node, nodeList, relList, asOfDate, ownershipTimeline);
    return (
      <div className="form-row" key={fieldId}>
        <label className="form-label">{prompt}</label>
        <div className="form-input" style={{ minHeight: 38, display: "flex", alignItems: "center", color: computed ? "#111827" : "#9ca3af", background: "#f8fafc" }}>
          {computed || "No ownership records"}
        </div>
      </div>
    );
  }

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

const parseOwnershipPercent = (rawValue) => {
  if (rawValue === "" || rawValue == null) {
    return { ok: true, value: null };
  }
  const parsed = Number(rawValue);
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 100) {
    return { ok: false, message: "Ownership percent must be between 0 and 100." };
  }
  return { ok: true, value: parsed };
};

// Format ownership effective date range for display
// Handles 4 cases: 1) no dates, 2) from only, 3) to only, 4) both dates
const formatOwnershipDateRange = (from, to) => {
  const formatDate = (dateStr) => {
    if (!dateStr) return null;
    try {
      // Parse as local date, not UTC (ISO date strings are dates, not datetimes)
      const [year, month, day] = dateStr.split('-').map(Number);
      const d = new Date(year, month - 1, day);
      return d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    } catch {
      return dateStr;
    }
  };

  const fromFormatted = formatDate(from);
  const toFormatted = formatDate(to);

  // Case 1: No dates
  if (!fromFormatted && !toFormatted) {
    return "Current ownership";
  }

  // Case 2: From date only (no end date = current)
  if (fromFormatted && !toFormatted) {
    return `Current owners (since ${fromFormatted})`;
  }

  // Case 3: To date only (missing start date)
  if (!fromFormatted && toFormatted) {
    return `Ownership until ${toFormatted}`;
  }

  // Case 4: Both dates
  return `Owners from ${fromFormatted} until ${toFormatted}`;
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
  onQuickView,
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
      style={{ position: "relative", cursor: "pointer" }}
      data-hv-node-id={item.nodeId}
      onClick={() => onQuickView?.(item.nodeId)}
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
      {showExplodeControls && !isCyclic && childCount > 0 && (
        <>
          {childCount > 1 && (
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
          )}
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
const DEFAULT_OWNERSHIP_TABULAR_VIEW_ID = "__default_ownership__";
const TABULAR_PREFS_STORAGE_KEY = "tabularPrefs";

function readStoredTabularPrefs() {
  if (typeof window === "undefined") {
    return { tabularViews: [], tabularViewsSelectedId: DEFAULT_TABULAR_VIEW_ID };
  }
  try {
    const raw = localStorage.getItem(TABULAR_PREFS_STORAGE_KEY);
    if (!raw) return { tabularViews: [], tabularViewsSelectedId: DEFAULT_TABULAR_VIEW_ID };
    const parsed = JSON.parse(raw);
    return {
      tabularViews: Array.isArray(parsed?.tabularViews) ? parsed.tabularViews : [],
      tabularViewsSelectedId: parsed?.tabularViewsSelectedId || DEFAULT_TABULAR_VIEW_ID,
    };
  } catch {
    return { tabularViews: [], tabularViewsSelectedId: DEFAULT_TABULAR_VIEW_ID };
  }
}

// Normalize appliesTo: old string "both" → ["entity","person"], single string → [string], array → as-is
const normalizeAppliesTo = (v) =>
  Array.isArray(v) ? v :
    v === "both" ? ["entity", "person"] :
      v && v.includes(",") ? v.split(",") :
        v ? [v] : ["entity", "person"];

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
  onExplode, onExplodeAll, onQuickView,
  onFocus = () => {}, onFocusPrimary = () => {}, onEdit = () => {}, onPrintBook = () => {}, onPrintPoster = () => {},
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
        onExplodeAll={onExplodeAll} onQuickView={onQuickView}
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
const ExplodableChildRow = ({ items, nodeList, relList, explodedNodes, onExplode, onExplodeAll, onQuickView, depth = 0, visitedIds = new Set() }) => {
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
            onQuickView={onQuickView}
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
              onQuickView={onQuickView}
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
                  onQuickView={onQuickView}
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
      const _nat = normalizeAppliesTo(field.appliesTo);
      if (_nat.includes("ownership") && !_nat.includes("entity") && !_nat.includes("person")) continue;
      const nodes = (_nat.includes("entity") && !_nat.includes("person")) ? filteredEntityNodes
        : (!_nat.includes("entity") && _nat.includes("person")) ? filteredPersonNodes
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


// ── Tabular sort / filter helpers ────────────────────────────────────────────

function isFilterEmpty(filter) {
  if (!filter) return true;
  if (filter.type === "text") return !filter.value;
  if (filter.type === "enum") return !(filter.selected?.length);
  if (filter.type === "range") return filter.min === "" && filter.max === "";
  if (filter.type === "daterange") return !filter.from && !filter.to;
  return true;
}

function formatFilterSummary(filter) {
  if (!filter) return "";
  if (filter.type === "text") return `"${filter.value}"`;
  if (filter.type === "enum") return (filter.selected || []).join(", ");
  if (filter.type === "range") {
    const parts = [];
    if (filter.min !== "" && filter.min != null) parts.push(`\u2265 ${filter.min}`);
    if (filter.max !== "" && filter.max != null) parts.push(`\u2264 ${filter.max}`);
    return parts.join(", ");
  }
  if (filter.type === "daterange") {
    const parts = [];
    if (filter.from) parts.push(`from ${filter.from}`);
    if (filter.to) parts.push(`to ${filter.to}`);
    return parts.join(" ");
  }
  return "";
}

function getColumnFilterConfig(column) {
  const key = column.key;
  if (key === "status" || key === "actions") return { filterType: null };
  if (key === "type") return { filterType: "enum", enumOptions: [{ value: "entity", label: "Entity" }, { value: "person", label: "Person" }] };
  if (key === "operationalRole") return { filterType: "enum", enumOptions: ["Active", "Passive", "Mixed"].map((v) => ({ value: v, label: v })) };
  if (key === "legalStatus") return { filterType: "enum", enumOptions: ["Good Standing", "Dormant", "Dissolved", "Suspended"].map((v) => ({ value: v, label: v })) };
  if (key === "personStatus") return { filterType: "enum", enumOptions: ["Active", "Inactive", "Deceased", "Former"].map((v) => ({ value: v, label: v })) };
  if (key === "percent") return { filterType: "range" };
  if (key === "startDate" || key === "endDate") return { filterType: "daterange" };
  if (column.field?.dataType === "boolean") return { filterType: "enum", enumOptions: [{ value: "true", label: "Yes" }, { value: "false", label: "No" }] };
  if (column.field?.dataType === "date") return { filterType: "daterange" };
  if (["number", "currency", "percentage"].includes(column.field?.dataType)) return { filterType: "range" };
  if ((column.field?.validValues || []).length > 0) return { filterType: "enum", enumOptions: column.field.validValues.map((v) => ({ value: v, label: v })) };
  return { filterType: "text" };
}

// ── Canvas-based text width measurement ───────────────────────────────────
let _measureCanvas = null;
function measureTextWidth(text, fontSize = 13) {
  const s = String(text ?? "");
  if (!s) return 0;
  try {
    if (!_measureCanvas) _measureCanvas = document.createElement("canvas");
    const ctx = _measureCanvas.getContext("2d");
    if (ctx) {
      ctx.font = `${fontSize}px Inter, system-ui, Arial, sans-serif`;
      const w = ctx.measureText(s).width;
      if (w > 0) return w; // canvas can return 0 without throwing — fall through
    }
  } catch { /* ignore */ }
  // Fallback: character-count estimation
  return s.length * fontSize * 0.65;
}

function getNodeTabularValue(node, column, nodeList = [], relList = [], asOfDate = null) {
  if (!node) return "";
  switch (column.key) {
    case "type": return node.kind || "";
    case "name": return node.name || "";
    case "address": return node.address || "";
    case "workPhone": return node.workPhone || "";
    case "cellPhone": return node.cellPhone || "";
    case "emails": return Array.isArray(node.emails) ? node.emails.join(", ") : (node.emails || "");
    case "taxId": return node.taxId || "";
    case "operationalRole": return node.operationalRole || "";
    case "legalStatus": return node.legalStatus || "";
    case "personStatus": return node.personStatus || "";
    default:
      if (!column.field) return "";
      if (column.field._virtual) {
        return getEntityOwnershipSummary(node, nodeList, relList, asOfDate);
      }
      return node.customFields?.[column.field.fieldId] ?? "";
  }
}

function getOwnershipTabularValue(rel, nodeList, column) {
  if (!rel) return "";
  switch (column.key) {
    case "owner": { const n = nodeList.find((x) => x.id === rel.from); return n?.name || rel.from || ""; }
    case "owned": { const n = nodeList.find((x) => x.id === rel.to); return n?.name || rel.to || ""; }
    case "percent": return rel.percent ?? "";
    case "startDate": return rel.startDate || "";
    case "endDate": return rel.endDate || "";
    default: return column.field ? (rel.customFields?.[column.field.fieldId] ?? "") : "";
  }
}

function matchesTabularFilter(value, filter) {
  if (isFilterEmpty(filter)) return true;
  if (filter.type === "text") return String(value ?? "").toLowerCase().includes(filter.value.toLowerCase());
  if (filter.type === "enum") return filter.selected.includes(String(value ?? ""));
  if (filter.type === "range") {
    const n = Number(value);
    if (filter.min !== "" && n < Number(filter.min)) return false;
    if (filter.max !== "" && n > Number(filter.max)) return false;
    return true;
  }
  if (filter.type === "daterange") {
    const v = String(value ?? "");
    if (filter.from && v < filter.from) return false;
    if (filter.to && v > filter.to) return false;
    return true;
  }
  return true;
}

function ColumnFilterPopover({ filterType, enumOptions, currentFilter, popoverPos, onChange, onClose }) {
  const ref = React.useRef(null);
  React.useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    document.addEventListener("mousedown", handler, true);
    return () => document.removeEventListener("mousedown", handler, true);
  }, [onClose]);

  const textVal = currentFilter?.type === "text" ? (currentFilter.value || "") : "";
  const enumSel = currentFilter?.type === "enum" ? (currentFilter.selected || []) : [];
  const rangeMin = currentFilter?.type === "range" ? (currentFilter.min || "") : "";
  const rangeMax = currentFilter?.type === "range" ? (currentFilter.max || "") : "";
  const dateFrom = currentFilter?.type === "daterange" ? (currentFilter.from || "") : "";
  const dateTo = currentFilter?.type === "daterange" ? (currentFilter.to || "") : "";
  const hasValue = !isFilterEmpty(currentFilter);

  return (
    <div
      ref={ref}
      className="tabular-filter-popover"
      style={{ position: "fixed", top: popoverPos.top, left: popoverPos.left, zIndex: 9999 }}
      onClick={(e) => e.stopPropagation()}
    >
      <button className="tabular-filter-confirm-btn" title="Done" onClick={onClose}>
        <Save size={13} />
      </button>
      {filterType === "text" && (
        <input
          className="tabular-filter-input"
          autoFocus
          placeholder="Contains…"
          value={textVal}
          onChange={(e) => onChange({ type: "text", value: e.target.value })}
          onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}
        />
      )}
      {filterType === "enum" && (
        <div className="tabular-filter-enum-list">
          {enumOptions.map((opt) => {
            const val = typeof opt === "object" ? opt.value : opt;
            const label = typeof opt === "object" ? opt.label : val;
            const checked = enumSel.includes(val);
            return (
              <label key={val} className="tabular-filter-enum-item">
                <input type="checkbox" checked={checked} onChange={() => {
                  const next = checked ? enumSel.filter((s) => s !== val) : [...enumSel, val];
                  onChange({ type: "enum", selected: next });
                }} />
                <span>{label}</span>
              </label>
            );
          })}
        </div>
      )}
      {filterType === "range" && (
        <div className="tabular-filter-range">
          <input className="tabular-filter-input" type="number" placeholder="Min" value={rangeMin}
            onChange={(e) => onChange({ type: "range", min: e.target.value, max: rangeMax })} />
          <input className="tabular-filter-input" type="number" placeholder="Max" value={rangeMax}
            onChange={(e) => onChange({ type: "range", min: rangeMin, max: e.target.value })} />
        </div>
      )}
      {filterType === "daterange" && (
        <div className="tabular-filter-range">
          <label className="tabular-filter-date-label">From</label>
          <input className="tabular-filter-input" type="date" value={dateFrom}
            onChange={(e) => onChange({ type: "daterange", from: e.target.value, to: dateTo })} />
          <label className="tabular-filter-date-label">To</label>
          <input className="tabular-filter-input" type="date" value={dateTo}
            onChange={(e) => onChange({ type: "daterange", from: dateFrom, to: e.target.value })} />
        </div>
      )}
      <div className="tabular-filter-popover-footer">
        {hasValue && (
          <button className="tabular-filter-clear-btn" onClick={() => onChange(null)}>Clear filter</button>
        )}
      </div>
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
  const [showStats, setShowStats] = useState(false);
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [homeAnimating, setHomeAnimating] = useState(false);
  const [homeAnimOrigin, setHomeAnimOrigin] = useState("50% 50%");
  const settingsRef = useRef(null);
  const exportMenuRef = useRef(null);
  const homeButtonRef = useRef(null);
  const quickFindRef = useRef(null);
  const quickFindInputRef = useRef(null);
  const focusBoxRef = useRef(null);
  const hierarchyStageRef = useRef(null);
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
  const [uploadOwnershipAsOfDate, setUploadOwnershipAsOfDate] = useState("");
  const [uploadDetected, setUploadDetected] = useState("");
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [uploadPreview, setUploadPreview] = useState(null);
  const [uploadPreviewLoading, setUploadPreviewLoading] = useState(false);
  const [notification, setNotification] = useState(null);

  // Show notification - persists until user closes it
  const showNotification = useCallback((type, title, message, details = null) => {
    setNotification({ type, title, message, details, id: Date.now() });
    // Play sound notification
    try {
      const audioContext = new (window.AudioContext || window.webkitAudioContext)();
      const oscillator = audioContext.createOscillator();
      const gainNode = audioContext.createGain();
      oscillator.connect(gainNode);
      gainNode.connect(audioContext.destination);
      
      if (type === "success") {
        oscillator.frequency.value = 800;
        gainNode.gain.setValueAtTime(0.3, audioContext.currentTime);
        gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.2);
        oscillator.start(audioContext.currentTime);
        oscillator.stop(audioContext.currentTime + 0.2);
      } else if (type === "error") {
        oscillator.frequency.value = 300;
        gainNode.gain.setValueAtTime(0.2, audioContext.currentTime);
        gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.3);
        oscillator.start(audioContext.currentTime);
        oscillator.stop(audioContext.currentTime + 0.3);
      }
    } catch (e) {
      // Audio context not supported, silently continue
    }
  }, []);

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
  const [tabularViews, setTabularViews] = useState(() => readStoredTabularPrefs().tabularViews);
  const [selectedTabularViewId, setSelectedTabularViewId] = useState(() => {
    if (homeScreen?.viewMode === "tabular" && homeScreen?.selectedTabularViewId) {
      return homeScreen.selectedTabularViewId;
    }
    return readStoredTabularPrefs().tabularViewsSelectedId || DEFAULT_TABULAR_VIEW_ID;
  });
  const [tabularViewDialogOpen, setTabularViewDialogOpen] = useState(false);
  const [tabularViewDraft, setTabularViewDraft] = useState({
    name: "",
    columnOrder: [],
  });
  const [tabularViewNameError, setTabularViewNameError] = useState("");
  const [tabularDragKey, setTabularDragKey] = useState(null);
  const [tabularDragOverKey, setTabularDragOverKey] = useState(null);
  const [tabularDraftFilterKey, setTabularDraftFilterKey] = useState(null);
  const [tabularDraftFilterPopoverPos, setTabularDraftFilterPopoverPos] = useState({ top: 0, left: 0 });
  const [tabularSaveAsNewOpen, setTabularSaveAsNewOpen] = useState(false);
  const [tabularSaveAsNewName, setTabularSaveAsNewName] = useState("");
  const [tabularSaveAsNewError, setTabularSaveAsNewError] = useState("");
  const [tabularDeleteConfirmOpen, setTabularDeleteConfirmOpen] = useState(false);
  const [ownershipDeleteConfirmOpen, setOwnershipDeleteConfirmOpen] = useState(false);

  // ── Tabular sub-mode (nodes vs ownerships) ────────────────────────────────
  const [tabularSubMode, setTabularSubMode] = useState("entities"); // "entities" | "persons" | "ownerships"

  // ── Ownership tabular view management ─────────────────────────────────────
  const [ownershipTabularViews, setOwnershipTabularViews] = useState([]);
  const [selectedOwnershipTabularViewId, setSelectedOwnershipTabularViewId] = useState(DEFAULT_OWNERSHIP_TABULAR_VIEW_ID);
  const [ownershipTabularViewDialogOpen, setOwnershipTabularViewDialogOpen] = useState(false);
  const [ownershipTabularViewDraft, setOwnershipTabularViewDraft] = useState({ name: "", columnOrder: [] });
  const [ownershipTabularViewNameError, setOwnershipTabularViewNameError] = useState("");
  const [ownershipOrderInputs, setOwnershipOrderInputs] = useState({});
  const [ownershipSaveAsNewOpen, setOwnershipSaveAsNewOpen] = useState(false);
  const [ownershipSaveAsNewName, setOwnershipSaveAsNewName] = useState("");
  const [ownershipSaveAsNewError, setOwnershipSaveAsNewError] = useState("");

  // ── Tabular sort / filter state ───────────────────────────────────────────
  const [tabularSort, setTabularSort] = useState(null);           // { key, dir:"asc"|"desc" } | null
  const [tabularFilters, setTabularFilters] = useState({});       // { [columnKey]: filter | null }
  const [openTabularFilterKey, setOpenTabularFilterKey] = useState(null);
  const [tabularFilterPopoverPos, setTabularFilterPopoverPos] = useState({ top: 0, left: 0 });

  const [ownershipTabularSort, setOwnershipTabularSort] = useState(null);
  const [ownershipTabularFilters, setOwnershipTabularFilters] = useState({});
  const [openOwnershipTabularFilterKey, setOpenOwnershipTabularFilterKey] = useState(null);
  const [ownershipTabularFilterPopoverPos, setOwnershipTabularFilterPopoverPos] = useState({ top: 0, left: 0 });

  // ── Ownership tabular row state ────────────────────────────────────────────
  const [ownershipTableDrafts, setOwnershipTableDrafts] = useState({});
  const [ownershipTableDirtyKeys, setOwnershipTableDirtyKeys] = useState(() => new Set());
  const [ownershipTableSavingKeys, setOwnershipTableSavingKeys] = useState(() => new Set());
  const [ownershipTableRowErrors, setOwnershipTableRowErrors] = useState({});
  const [asOfDate, setAsOfDate] = useState("");
  const todayIso = useMemo(() => toIsoDate(new Date()), []);

  const getTabularPrefsPayload = useCallback((overrides = {}) => ({
    tabularViews: overrides.tabularViews ?? tabularViews,
    tabularViewsSelectedId: overrides.tabularViewsSelectedId ?? selectedTabularViewId,
    ownershipTabularViews: overrides.ownershipTabularViews ?? ownershipTabularViews,
    ownershipTabularViewsSelectedId: overrides.ownershipTabularViewsSelectedId ?? selectedOwnershipTabularViewId,
  }), [selectedTabularViewId, tabularViews, ownershipTabularViews, selectedOwnershipTabularViewId]);

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
  const [editNodeEffectiveDate, setEditNodeEffectiveDate] = useState("");
  const [editNodeOwnershipTimeline, setEditNodeOwnershipTimeline] = useState([]);
  const [nodeDraft, setNodeDraft] = useState({
    id: "",
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
  const newOwnershipPercentInvalid =
    newOwnership.percent !== "" && !parseOwnershipPercent(newOwnership.percent).ok;
  const editOwnershipPercentInvalid =
    ownershipDraft.percent !== "" && !parseOwnershipPercent(ownershipDraft.percent).ok;

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
  const [ownerEditorEffectiveDate, setOwnerEditorEffectiveDate] = useState("");
  const [ownerEditorDateRange, setOwnerEditorDateRange] = useState({ from: null, to: null, isCurrent: true });
  const [ownerEditorMode, setOwnerEditorMode] = useState("view"); // 'view' | 'edit-existing' | 'create-new'
  const [ownershipTimeline, setOwnershipTimeline] = useState([]); // Array of all ownership periods
  const [ownershipSelectedPeriodSetId, setOwnershipSelectedPeriodSetId] = useState(null); // Which period is displayed
  const [ownershipDeleteConfirm, setOwnershipDeleteConfirm] = useState(false); // Confirmation for deleting group
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
  const [quickFindQuery, setQuickFindQuery] = useState("");
  const [quickFindOpen, setQuickFindOpen] = useState(false);
  const [quickFindHighlight, setQuickFindHighlight] = useState(-1);
  const [quickViewNodeId, setQuickViewNodeId] = useState("");
  const [quickViewOwnershipTimeline, setQuickViewOwnershipTimeline] = useState([]);

  const [dataDictionary, setDataDictionary] = useState([]);
  const emptyDdDraft = { prompt: "", dataType: "string", appliesTo: ["entity", "person"], multiValue: false, validValuesText: "", phoneTypesText: "", showInStats: false };
  const [ddEntryDraft, setDdEntryDraft] = useState(emptyDdDraft);
  const [ddEntryId, setDdEntryId] = useState(null);
  const [isSavingDdEntry, setIsSavingDdEntry] = useState(false);

  const focusNode = useMemo(() => getNode(nodeList, focusId), [nodeList, focusId]);
  const quickFindMatches = useMemo(() => {
    const q = String(quickFindQuery || "").trim().toLowerCase();
    if (q.length < 2) return [];
    const scored = nodeList.map((node) => {
      const name = String(node.name || "");
      const id = String(node.id || "");
      const kind = String(node.kind || "");
      const nameLc = name.toLowerCase();
      const idLc = id.toLowerCase();
      const kindLc = kind.toLowerCase();

      let score = -1;
      if (nameLc === q) score = 100;
      else if (nameLc.startsWith(q)) score = 90;
      else if (nameLc.includes(q)) score = 75;
      else if (idLc.includes(q)) score = 55;
      else if (kindLc.includes(q)) score = 35;

      return score >= 0 ? { node, score } : null;
    }).filter(Boolean);

    return scored
      .sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return String(a.node.name || a.node.id || "").localeCompare(String(b.node.name || b.node.id || ""), undefined, { sensitivity: "base" });
      })
      .slice(0, 12)
      .map((entry) => entry.node);
  }, [nodeList, quickFindQuery]);

  const quickViewNode = useMemo(() => getNode(nodeList, quickViewNodeId), [nodeList, quickViewNodeId]);
  const quickViewOwners = useMemo(
    () => (quickViewNodeId ? getOwnersOf(relList, quickViewNodeId) : []),
    [quickViewNodeId, relList]
  );
  const quickViewOwned = useMemo(
    () => (quickViewNodeId ? getOwnedBy(relList, quickViewNodeId) : []),
    [quickViewNodeId, relList]
  );
  const quickViewDescCount = useMemo(
    () => (quickViewNodeId ? getAllDescendants(relList, quickViewNodeId).size : 0),
    [quickViewNodeId, relList]
  );
  
  // Fetch ownership timeline for quick view node
  useEffect(() => {
    if (quickViewNode?.kind === "entity") {
      apiRequest(`/api/ownership/history/${encodeURIComponent(quickViewNode.id)}`)
        .then((resp) => setQuickViewOwnershipTimeline(resp.periods || []))
        .catch(() => setQuickViewOwnershipTimeline([]));
    } else {
      setQuickViewOwnershipTimeline([]);
    }
  }, [quickViewNode?.id]);
  
  const quickViewOwnershipSummary = useMemo(() => {
    if (!quickViewNode || quickViewNode.kind !== "entity") return "";
    const raw = String(getEntityOwnershipSummary(quickViewNode, nodeList, relList, asOfDate, quickViewOwnershipTimeline) || "").trim();
    if (!raw) return "";
    return raw.split(";").map((part) => part.trim()).filter(Boolean).join("\n");
  }, [nodeList, quickViewNode, relList, asOfDate, quickViewOwnershipTimeline]);
  const quickViewEmailText = useMemo(() => {
    if (!quickViewNode) return "";
    const emails = quickViewNode.emails;
    if (Array.isArray(emails)) {
      const vals = emails.map((v) => String(v || "").trim()).filter(Boolean);
      return vals.join(", ");
    }
    const txt = String(emails || "").trim();
    return txt;
  }, [quickViewNode]);
  const quickViewPhoneText = useMemo(() => {
    if (!quickViewNode) return "";
    const vals = [quickViewNode.workPhone, quickViewNode.cellPhone]
      .map((v) => String(v || "").trim())
      .filter(Boolean);
    return vals.join(" / ");
  }, [quickViewNode]);
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
    () => [...dataDictionary, ENTITY_OWNERSHIP_SUMMARY_FIELD]
      .filter((f) => { const n = normalizeAppliesTo(f.appliesTo); return n.includes("entity") || n.includes("person"); })
      .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0)),
    [dataDictionary]
  );
  const ownershipDdFields = useMemo(
    () => [...dataDictionary]
      .filter((f) => normalizeAppliesTo(f.appliesTo).includes("ownership"))
      .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0)),
    [dataDictionary]
  );
  const baseTabularColumns = useMemo(() => ([
    { key: "name", label: "Name", hideable: true },
    { key: "address", label: "Address", hideable: true },
    { key: "workPhone", label: "Primary Phone", hideable: true, width: 160 },
    { key: "cellPhone", label: "Cell Phone", hideable: true, width: 160 },
    { key: "emails", label: "e-Mail", hideable: true },
    { key: "taxId", label: "Tax ID", hideable: true, width: 140 },
    { key: "operationalRole", label: "Operational Role", hideable: true, width: 160 },
    { key: "legalStatus", label: "Legal Status", hideable: true, width: 150 },
    { key: "personStatus", label: "Person Status", hideable: true, width: 150 },
    { key: "actions", label: "Actions", hideable: false },
  ]), []);
  const allTabularColumns = useMemo(() => {
    const ddCols = tableDdFields.map((field) => ({
      key: field.fieldId,
      label: field.prompt,
      field,
      hideable: true,
    }));
    const selectableBaseColumns = baseTabularColumns.filter((c) => c.key !== "actions");
    return [...selectableBaseColumns, ...ddCols].filter(Boolean);
  }, [baseTabularColumns, tableDdFields]);
  const defaultTabularOrder = useMemo(
    () => allTabularColumns.map((c) => c.key),
    [allTabularColumns]
  );

  // ── Ownership tabular columns ──────────────────────────────────────────────
  const baseOwnershipTabularColumns = useMemo(() => ([
    { key: "owner", label: "Owner", hideable: true },
    { key: "owned", label: "Owned Entity", hideable: true },
    { key: "percent", label: "% Owned", hideable: true, width: 110 },
    { key: "startDate", label: "Start Date", hideable: true, width: 130 },
    { key: "endDate", label: "End Date", hideable: true, width: 130 },
  ]), []);
  const allOwnershipTabularColumns = useMemo(() => {
    const ddCols = ownershipDdFields.map((field) => ({
      key: field.fieldId,
      label: field.prompt,
      field,
      hideable: true,
    }));
    return [...baseOwnershipTabularColumns, ...ddCols].filter(Boolean);
  }, [baseOwnershipTabularColumns, ownershipDdFields]);
  const defaultOwnershipTabularOrder = useMemo(
    () => allOwnershipTabularColumns.map((c) => c.key),
    [allOwnershipTabularColumns]
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
    const sort = view.sort?.key && (view.sort.dir === "asc" || view.sort.dir === "desc")
      ? { key: view.sort.key, dir: view.sort.dir } : null;
    const filters = (view.filters && typeof view.filters === "object") ? view.filters : {};
    const columnWidths = {};
    if (view.columnWidths && typeof view.columnWidths === "object") {
      for (const [k, v] of Object.entries(view.columnWidths)) {
        const n = Number(v); if (n > 0) columnWidths[k] = n;
      }
    }
    return { id, name, nodeKind: view.nodeKind || null, columnOrder: visibleOrder, sort, filters, columnWidths };
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
    try {
      localStorage.setItem(TABULAR_PREFS_STORAGE_KEY, JSON.stringify({
        tabularViews,
        tabularViewsSelectedId: selectedTabularViewId,
      }));
    } catch { }
  }, [selectedTabularViewId, tabularViews]);

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
        body: JSON.stringify({ debug: false, asOf: asOfDate || null }),
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
  }, [apiBase, token, asOfDate]);

  // Reload directory whenever asOfDate changes (including to null for current data)
  useEffect(() => {
    if (token && apiBase) {
      loadDirectory();
    }
  }, [asOfDate, loadDirectory, token, apiBase]);

  // Background ownership import - starts import and returns immediately with feedback
  const startBackgroundOwnershipImport = (rows, csvSkipped) => {
    // Fire-and-forget: start import in background without waiting
    (async () => {
      try {
        console.log(`[Background Import] Starting ownership import of ${rows.length} rows...`);
        const chunkCsv = rowsToCsv(rows);
        
        const abortController = new AbortController();
        const timeoutId = setTimeout(() => abortController.abort(), 5 * 60 * 1000);
        
        const response = await fetch(`${apiBase}/api/import/ownerships-csv`, {
          method: "POST",
          headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
          body: JSON.stringify({ csv: chunkCsv, asOfDate: uploadOwnershipAsOfDate, client: clientId }),
          signal: abortController.signal,
        });
        clearTimeout(timeoutId);
        
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

        console.log(`[Background Import] Complete: ${data.imported} imported, ${data.skipped} skipped`);
        
        // Refresh directory to show new data
        await loadDirectory();
        
        // Show completion result
        const imported = Number(data.imported || 0);
        const skipped = Number(data.skipped || 0);
        const total = rows.length;
        
        setUploadSummary({
          total,
          imported,
          skipped: csvSkipped + skipped,
          errors: data.errors || [],
        });
        setUploadStatus("success");
        
        // Show success notification
        const skipMsg = csvSkipped + skipped > 0 ? `, ${csvSkipped + skipped} skipped` : "";
        showNotification(
          "success",
          "✓ Import Complete",
          `Imported ${imported} ownership records${skipMsg}`,
          { imported, skipped: csvSkipped + skipped, total }
        );
      } catch (err) {
        console.error(`[Background Import] Error:`, err);
        setUploadStatus("error");
        setUploadError(`Background import failed: ${err.message || 'Unknown error'}`);
        
        // Show error notification
        showNotification(
          "error",
          "✗ Import Failed",
          err.message || "An error occurred during import"
        );
      }
    })();
  };


  const handleUploadCsv = async () => {
    if (!uploadFile) {
      setUploadError("Please select a CSV file.");
      setUploadStatus("error");
      return;
    }
    try {
      if (!normalizeDateInput(uploadOwnershipAsOfDate)) {
        setUploadStatus("error");
        setUploadError("Effective as of date is required for imports.");
        return;
      }

      if (uploadType === "ownership" && uploadPreview?.ownershipValidation && !uploadPreview.ownershipValidation.valid) {
        setUploadStatus("error");
        setUploadError("Import blocked: each owned entity must total exactly 0% or 100% ownership.");
        return;
      }

      // For ownership imports, use background loading with immediate feedback
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

        // Enforce 500-record limit for ownership imports
        const MAX_OWNERSHIP_RECORDS = 500;
        if (rows.length > MAX_OWNERSHIP_RECORDS) {
          setUploadStatus("error");
          setUploadError(`Ownership imports are limited to ${MAX_OWNERSHIP_RECORDS} records. Your file contains ${rows.length} records. Please split into multiple files.`);
          return;
        }

        // Show immediate feedback and close dialog
        setUploadStatus("loading-background");
        setUploadError("");
        setUploadSummary(null);
        setUploadProgress({ current: 0, total: 0 });

        // Start import in background
        startBackgroundOwnershipImport(rows, skipped);
        return;
      }

      // Non-ownership imports: proceed with original blocking behavior
      setUploadStatus("uploading");
      setUploadError("");
      setUploadSummary(null);
      setUploadProgress({ current: 0, total: 0 });

      const form = new FormData();
      form.append("file", uploadFile);
      form.append("client", clientId);
      form.append("asOfDate", uploadOwnershipAsOfDate);
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

  const ownershipPreviewValidation = useMemo(() => {
    if (uploadType !== "ownership" || !uploadPreview) {
      return { valid: true, offendingEntities: [] };
    }
    const validation = uploadPreview.ownershipValidation;
    if (validation && Array.isArray(validation.offendingEntities)) {
      return {
        valid: !!validation.valid,
        offendingEntities: validation.offendingEntities,
      };
    }
    return { valid: true, offendingEntities: [] };
  }, [uploadType, uploadPreview]);

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
        { header: "Address", helper: "" },
        { header: "Primary Phone", helper: "" },
        { header: "Cell Phone", helper: "(People only)" },
        { header: "e-Mail", helper: "(People only)" },
        { header: "Tax ID", helper: "" },
      ];

      // Mirrors the server's normalizeHeader — strips punctuation, lowercases, collapses spaces.
      const normalizeH = (str) =>
        String(str || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();

      // All synonyms for the built-in fields already in the template.
      // Any DD field whose prompt normalizes to one of these is redundant and excluded.
      const BUILTIN_SYNONYMS = new Set([
        "name", "company name", "entity name", "node name", "organization", "org name", "business name", "legal name", "entity or person s name",
        "address", "street", "street address", "mailing address", "location", "addr",
        "work phone", "phone", "workphone", "office phone", "ph work", "business phone", "telephone", "tel", "phone number", "primary phone",
        "cell phone", "cell", "mobile", "mobile phone", "cellphone", "cell number", "personal phone",
        "email", "emails", "email address", "e mail", "email addr",
        "tax id", "taxid", "ein", "tin", "federal id", "tax identification", "federal tax id", "fein",
      ]);

      // DD custom fields — exclude file type and any field redundant with a built-in, sorted by sortOrder
      const ddFields = [...dataDictionary]
        .sort((a, b) => String(a.prompt || "").localeCompare(String(b.prompt || ""), undefined, { sensitivity: "base" }))
        .filter((f) => f.dataType !== "file")
        .filter((f) => !BUILTIN_SYNONYMS.has(normalizeH(f.prompt)));

      const ddCols = ddFields.map((f) => {
        const parts = [];
        if (Array.isArray(f.validValues) && f.validValues.length > 0) {
          parts.push(f.validValues.join(", "));
        }
        { const _n = normalizeAppliesTo(f.appliesTo), all3 = _n.includes("entity") && _n.includes("person") && _n.includes("ownership"); if (!all3) { const ls = [_n.includes("entity") && "Entities", _n.includes("person") && "People", _n.includes("ownership") && "Ownerships"].filter(Boolean); if (ls.length) parts.push(`(${ls.join(" & ")} only)`); } }
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
      appliesTo: normalizeAppliesTo(entry.appliesTo || "both"),
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
    if (field?.dataType === "date") {
      return normalizeDateInput(raw);
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
          asOfDate: asOfDate || todayIso,
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
          asOfDate: asOfDate || todayIso,
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
    asOfDate,
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
    todayIso,
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
    const kindFilter = tabularSubMode === "entities" ? "entity" : tabularSubMode === "persons" ? "person" : null;
    const newRows = tableNewRows.filter((r) => {
      if (kindFilter && (r.draft.kind || "entity") !== kindFilter) return false;
      if (!dirSearchLower) return true;
      const txt = `${r.draft.name || ""} ${r.draft.id || ""}`.toLowerCase();
      return txt.includes(dirSearchLower);
    }).map((r) => ({ key: r.key, isNew: true, node: r.draft }));
    const existingRows = filteredAllNodes
      .filter((n) => !kindFilter || n.kind === kindFilter)
      .map((n) => ({ key: n.id, isNew: false, node: n }));
    return [...newRows, ...existingRows];
  }, [dirSearchLower, filteredAllNodes, tableNewRows, tabularSubMode]);

  const currentTabularNodeKind = useMemo(() => {
    if (tabularSubMode === "entities") return "entity";
    if (tabularSubMode === "persons") return "person";
    return null;
  }, [tabularSubMode]);

  const tabularViewsForCurrentKind = useMemo(() => {
    if (!currentTabularNodeKind) return tabularViews;
    return tabularViews.filter((v) => !v.nodeKind || v.nodeKind === currentTabularNodeKind);
  }, [currentTabularNodeKind, tabularViews]);

  const effectiveSelectedTabularViewId = useMemo(() => {
    if (!currentTabularNodeKind) return selectedTabularViewId;
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) return DEFAULT_TABULAR_VIEW_ID;
    const existsForKind = tabularViewsForCurrentKind.some((v) => v.id === selectedTabularViewId);
    return existsForKind ? selectedTabularViewId : DEFAULT_TABULAR_VIEW_ID;
  }, [currentTabularNodeKind, selectedTabularViewId, tabularViewsForCurrentKind]);

  useEffect(() => {
    if (!currentTabularNodeKind) return;
    if (selectedTabularViewId === effectiveSelectedTabularViewId) return;
    setSelectedTabularViewId(effectiveSelectedTabularViewId);
  }, [currentTabularNodeKind, effectiveSelectedTabularViewId, selectedTabularViewId]);

  const activeTabularView = useMemo(() => {
    if (effectiveSelectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) {
      return {
        id: DEFAULT_TABULAR_VIEW_ID,
        name: "Default",
        columnOrder: defaultTabularOrder,
      };
    }
    const found = tabularViewsForCurrentKind.find((v) => v.id === effectiveSelectedTabularViewId);
    return found || {
      id: DEFAULT_TABULAR_VIEW_ID,
      name: "Default",
      columnOrder: defaultTabularOrder,
    };
  }, [defaultTabularOrder, effectiveSelectedTabularViewId, tabularViewsForCurrentKind]);

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
    return middle;
  }, [activeTabularView.columnOrder, allTabularColumns]);

  const availableTabularColumns = useMemo(() => {
    const selected = new Set(tabularViewDraft.columnOrder || []);
    return allTabularColumns.filter((column) => !selected.has(column.key));
  }, [allTabularColumns, tabularViewDraft.columnOrder]);

  const openTabularViewManager = useCallback(() => {
    const nodeKind = tabularSubMode === "entities" ? "entity" : tabularSubMode === "persons" ? "person" : null;
    setTabularViewNameError("");
    setTabularViewDraft({
      name: activeTabularView.name || "",
      nodeKind,
      columnOrder: [...(activeTabularView.columnOrder || defaultTabularOrder)],
      sort: tabularSort,
      filters: { ...tabularFilters },
      columnWidths: { ...(activeTabularView.columnWidths || {}) },
    });
    setTabularDraftFilterKey(null);
    setTabularViewDialogOpen(true);
  }, [activeTabularView, defaultTabularOrder, tabularSort, tabularFilters, tabularSubMode]);

  const handleTabularDragDrop = useCallback((dragKey, dropKey) => {
    if (!dragKey || !dropKey || dragKey === dropKey) return;
    setTabularViewDraft((prev) => {
      const order = [...prev.columnOrder];
      const fromIdx = order.indexOf(dragKey);
      const toIdx = order.indexOf(dropKey);
      if (fromIdx < 0 || toIdx < 0) return prev;
      order.splice(fromIdx, 1);
      order.splice(toIdx, 0, dragKey);
      return { ...prev, columnOrder: order };
    });
  }, []);

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
    if (normalized === "default") return true; // reserved name
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
    setTabularSaveAsNewName("");
    setTabularSaveAsNewError("");
    setTabularSaveAsNewOpen(true);
  }, []);

  const commitTabularSaveAsNew = useCallback(() => {
    const name = tabularSaveAsNewName.trim();
    if (!name) { setTabularSaveAsNewError("Please enter a name."); return; }
    if (isDuplicateTabularViewName(name)) {
      setTabularSaveAsNewError(name.toLowerCase() === "default"
        ? '"Default" is a reserved name.'
        : "A view with this name already exists.");
      return;
    }
    const id = `view-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const nextView = sanitizeTabularView({ ...tabularViewDraft, id, name });
    if (!nextView) return;
    const nextViews = [...tabularViews, nextView];
    setTabularViews(nextViews);
    setSelectedTabularViewId(nextView.id);
    persistTabularViewPrefs(nextViews, nextView.id);
    setTabularSaveAsNewOpen(false);
    setTabularViewDialogOpen(false);
  }, [isDuplicateTabularViewName, persistTabularViewPrefs, sanitizeTabularView, tabularSaveAsNewName, tabularViewDraft, tabularViews]);

  const updateCurrentTabularView = useCallback(() => {
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) return;
    const selectedExists = tabularViews.some((v) => v.id === selectedTabularViewId);
    if (!selectedExists) return;
    const name = tabularViewDraft.name || activeTabularView.name || "";
    const nextView = sanitizeTabularView({ ...tabularViewDraft, id: selectedTabularViewId, name });
    if (!nextView) return;
    const nextViews = tabularViews.map((v) => (v.id === selectedTabularViewId ? nextView : v));
    setTabularViews(nextViews);
    persistTabularViewPrefs(nextViews, selectedTabularViewId);
    setTabularFilters(nextView.filters || {});
    setTabularSort(nextView.sort || null);
    setTabularViewDialogOpen(false);
  }, [activeTabularView.name, persistTabularViewPrefs, sanitizeTabularView, selectedTabularViewId, tabularViewDraft, tabularViews]);

  const deleteCurrentTabularView = useCallback(() => {
    if (selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) return;
    setTabularDeleteConfirmOpen(true);
  }, [selectedTabularViewId]);

  const confirmTabularViewDelete = useCallback(() => {
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
    setTabularDeleteConfirmOpen(false);
    setTabularViewDialogOpen(false);
  }, [apiRequest, getTabularPrefsPayload, homeScreen, persistTabularViewPrefs, selectedTabularViewId, tabularViews]);


  // ── Ownership tabular view management ─────────────────────────────────────
  const sanitizeOwnershipTabularView = useCallback((view) => {
    if (!view || typeof view !== "object") return null;
    const id = String(view.id || "").trim();
    const name = String(view.name || "").trim();
    if (!id || !name) return null;
    const allowed = new Set(defaultOwnershipTabularOrder);
    const rawOrder = Array.isArray(view.columnOrder)
      ? view.columnOrder.filter((k) => allowed.has(k))
      : [];
    const sort = view.sort?.key && (view.sort.dir === "asc" || view.sort.dir === "desc")
      ? { key: view.sort.key, dir: view.sort.dir } : null;
    const filters = (view.filters && typeof view.filters === "object") ? view.filters : {};
    const columnWidths = {};
    if (view.columnWidths && typeof view.columnWidths === "object") {
      for (const [k, v] of Object.entries(view.columnWidths)) {
        const n = Number(v); if (n > 0) columnWidths[k] = n;
      }
    }
    return { id, name, columnOrder: rawOrder, sort, filters, columnWidths };
  }, [defaultOwnershipTabularOrder]);

  const persistOwnershipTabularViewPrefs = useCallback((nextViews, nextSelectedId) => {
    apiRequest("/api/auth/me", {
      method: "PATCH",
      body: JSON.stringify({
        ownershipTabularViews: nextViews,
        ownershipTabularViewsSelectedId: nextSelectedId,
      }),
    }).catch((err) => {
      setRemoteStatus("error");
      setRemoteError(err.message || "Unable to save ownership tabular view settings.");
    });
  }, [apiRequest]);

  const isDuplicateOwnershipTabularViewName = useCallback((name, excludeId = null) => {
    const normalized = String(name || "").trim().toLowerCase();
    if (!normalized) return false;
    if (normalized === "default") return true; // reserved name
    return ownershipTabularViews.some((v) =>
      v.id !== excludeId && String(v.name || "").trim().toLowerCase() === normalized
    );
  }, [ownershipTabularViews]);

  const activeOwnershipTabularView = useMemo(() => {
    if (selectedOwnershipTabularViewId === DEFAULT_OWNERSHIP_TABULAR_VIEW_ID) {
      return { id: DEFAULT_OWNERSHIP_TABULAR_VIEW_ID, name: "Default", columnOrder: defaultOwnershipTabularOrder };
    }
    const found = ownershipTabularViews.find((v) => v.id === selectedOwnershipTabularViewId);
    return found || { id: DEFAULT_OWNERSHIP_TABULAR_VIEW_ID, name: "Default", columnOrder: defaultOwnershipTabularOrder };
  }, [defaultOwnershipTabularOrder, ownershipTabularViews, selectedOwnershipTabularViewId]);

  const visibleOwnershipTabularColumns = useMemo(() => {
    const byKey = new Map(allOwnershipTabularColumns.map((c) => [c.key, c]));
    const seen = new Set();
    const middle = (activeOwnershipTabularView.columnOrder || [])
      .map((key) => byKey.get(key))
      .filter((col) => { if (!col || seen.has(col.key)) return false; seen.add(col.key); return true; });
    return middle;
  }, [activeOwnershipTabularView.columnOrder, allOwnershipTabularColumns]);

  const availableOwnershipTabularColumns = useMemo(() => {
    const selected = new Set(ownershipTabularViewDraft.columnOrder || []);
    return allOwnershipTabularColumns.filter((c) => !selected.has(c.key));
  }, [allOwnershipTabularColumns, ownershipTabularViewDraft.columnOrder]);

  const openOwnershipTabularViewManager = useCallback(() => {
    setOwnershipTabularViewNameError("");
    setOwnershipTabularViewDraft({
      name: activeOwnershipTabularView.name || "",
      columnOrder: [...(activeOwnershipTabularView.columnOrder || defaultOwnershipTabularOrder)],
      sort: ownershipTabularSort,
      filters: { ...ownershipTabularFilters },
      columnWidths: { ...(activeOwnershipTabularView.columnWidths || {}) },
    });
    setOwnershipOrderInputs({});
    setOwnershipTabularViewDialogOpen(true);
  }, [activeOwnershipTabularView, defaultOwnershipTabularOrder, ownershipTabularSort, ownershipTabularFilters]);

  const applyOwnershipTabularOrder = useCallback((currentOrder, inputs) => {
    const orderValues = {};
    currentOrder.forEach((key, idx) => {
      const raw = inputs[key];
      const parsed = raw !== undefined && raw !== "" ? parseFloat(raw) : NaN;
      orderValues[key] = isNaN(parsed) ? (idx + 1) : parsed;
    });
    const sorted = [...currentOrder].sort((a, b) => orderValues[a] - orderValues[b]);
    setOwnershipTabularViewDraft((prev) => ({ ...prev, columnOrder: sorted }));
    setOwnershipOrderInputs({});
  }, []);

  const moveOwnershipTabularDraftColumn = useCallback((key, direction) => {
    setOwnershipTabularViewDraft((prev) => {
      const idx = prev.columnOrder.indexOf(key);
      if (idx < 0) return prev;
      const nextIdx = idx + direction;
      if (nextIdx < 0 || nextIdx >= prev.columnOrder.length) return prev;
      const columnOrder = [...prev.columnOrder];
      const tmp = columnOrder[idx]; columnOrder[idx] = columnOrder[nextIdx]; columnOrder[nextIdx] = tmp;
      return { ...prev, columnOrder };
    });
  }, []);

  const toggleOwnershipTabularDraftSelected = useCallback((key) => {
    setOwnershipTabularViewDraft((prev) => {
      const inOrder = prev.columnOrder.includes(key);
      return inOrder
        ? { ...prev, columnOrder: prev.columnOrder.filter((k) => k !== key) }
        : { ...prev, columnOrder: [...prev.columnOrder, key] };
    });
  }, []);

  const saveOwnershipTabularViewAsNew = useCallback(() => {
    setOwnershipSaveAsNewName("");
    setOwnershipSaveAsNewError("");
    setOwnershipSaveAsNewOpen(true);
  }, []);

  const commitOwnershipSaveAsNew = useCallback(() => {
    const name = ownershipSaveAsNewName.trim();
    if (!name) { setOwnershipSaveAsNewError("Please enter a name."); return; }
    if (isDuplicateOwnershipTabularViewName(name)) {
      setOwnershipSaveAsNewError(name.toLowerCase() === "default"
        ? '"Default" is a reserved name.'
        : "A view with this name already exists.");
      return;
    }
    const id = `ov-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const nextView = sanitizeOwnershipTabularView({ ...ownershipTabularViewDraft, id, name });
    if (!nextView) return;
    const nextViews = [...ownershipTabularViews, nextView];
    setOwnershipTabularViews(nextViews);
    setSelectedOwnershipTabularViewId(nextView.id);
    persistOwnershipTabularViewPrefs(nextViews, nextView.id);
    setOwnershipSaveAsNewOpen(false);
    setOwnershipTabularViewDialogOpen(false);
  }, [isDuplicateOwnershipTabularViewName, ownershipSaveAsNewName, ownershipTabularViewDraft, ownershipTabularViews, persistOwnershipTabularViewPrefs, sanitizeOwnershipTabularView]);

  const updateCurrentOwnershipTabularView = useCallback(() => {
    if (selectedOwnershipTabularViewId === DEFAULT_OWNERSHIP_TABULAR_VIEW_ID) return;
    const name = ownershipTabularViewDraft.name || activeOwnershipTabularView.name || "";
    const nextView = sanitizeOwnershipTabularView({ ...ownershipTabularViewDraft, id: selectedOwnershipTabularViewId, name });
    if (!nextView) return;
    const nextViews = ownershipTabularViews.map((v) => (v.id === selectedOwnershipTabularViewId ? nextView : v));
    setOwnershipTabularViews(nextViews);
    persistOwnershipTabularViewPrefs(nextViews, selectedOwnershipTabularViewId);
    setOwnershipTabularViewDialogOpen(false);
  }, [activeOwnershipTabularView.name, ownershipTabularViewDraft, ownershipTabularViews, persistOwnershipTabularViewPrefs, sanitizeOwnershipTabularView, selectedOwnershipTabularViewId]);

  const deleteCurrentOwnershipTabularView = useCallback(() => {
    if (selectedOwnershipTabularViewId === DEFAULT_OWNERSHIP_TABULAR_VIEW_ID) return;
    setOwnershipDeleteConfirmOpen(true);
  }, [selectedOwnershipTabularViewId]);

  const confirmOwnershipTabularViewDelete = useCallback(() => {
    if (selectedOwnershipTabularViewId === DEFAULT_OWNERSHIP_TABULAR_VIEW_ID) return;
    const nextViews = ownershipTabularViews.filter((v) => v.id !== selectedOwnershipTabularViewId);
    setOwnershipTabularViews(nextViews);
    setSelectedOwnershipTabularViewId(DEFAULT_OWNERSHIP_TABULAR_VIEW_ID);
    persistOwnershipTabularViewPrefs(nextViews, DEFAULT_OWNERSHIP_TABULAR_VIEW_ID);
    setOwnershipTabularViewNameError("");
    setOwnershipDeleteConfirmOpen(false);
    setOwnershipTabularViewDialogOpen(false);
  }, [ownershipTabularViews, persistOwnershipTabularViewPrefs, selectedOwnershipTabularViewId]);

  // ── Ownership tabular rows & save logic ───────────────────────────────────
  const ownershipTableRows = useMemo(() => {
    return ownerships.map((r) => ({ key: r.id, rel: r }));
  }, [ownerships]);

  // ── Sort / filter callbacks ───────────────────────────────────────────────
  const handleTabularSort = useCallback((key) => {
    setTabularSort((prev) => {
      if (!prev || prev.key !== key) return { key, dir: "asc" };
      if (prev.dir === "asc") return { key, dir: "desc" };
      return null;
    });
  }, []);
  const handleOwnershipTabularSort = useCallback((key) => {
    setOwnershipTabularSort((prev) => {
      if (!prev || prev.key !== key) return { key, dir: "asc" };
      if (prev.dir === "asc") return { key, dir: "desc" };
      return null;
    });
  }, []);

  // ── Apply stored sort/filters when switching to a saved view ──────────────
  const tabularViewsRef = useRef(tabularViews);
  useEffect(() => { tabularViewsRef.current = tabularViews; }, [tabularViews]);
  const ownershipTabularViewsRef = useRef(ownershipTabularViews);
  useEffect(() => { ownershipTabularViewsRef.current = ownershipTabularViews; }, [ownershipTabularViews]);
  useEffect(() => {
    if (effectiveSelectedTabularViewId === DEFAULT_TABULAR_VIEW_ID) return;
    const view = tabularViewsRef.current.find((v) => v.id === effectiveSelectedTabularViewId);
    if (!view) return;
    if (view.sort !== undefined) setTabularSort(view.sort || null);
    if (view.filters !== undefined) setTabularFilters(view.filters || {});
  }, [effectiveSelectedTabularViewId]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (selectedOwnershipTabularViewId === DEFAULT_OWNERSHIP_TABULAR_VIEW_ID) return;
    const view = ownershipTabularViewsRef.current.find((v) => v.id === selectedOwnershipTabularViewId);
    if (!view) return;
    if (view.sort !== undefined) setOwnershipTabularSort(view.sort || null);
    if (view.filters !== undefined) setOwnershipTabularFilters(view.filters || {});
  }, [selectedOwnershipTabularViewId]); // eslint-disable-line react-hooks/exhaustive-deps

  const toggleTabularFilter = useCallback((key, e) => {
    const rect = e.currentTarget.getBoundingClientRect();
    setTabularFilterPopoverPos({ top: rect.bottom + 4, left: rect.left });
    setOpenTabularFilterKey((prev) => (prev === key ? null : key));
  }, []);
  const toggleOwnershipTabularFilter = useCallback((key, e) => {
    const rect = e.currentTarget.getBoundingClientRect();
    setOwnershipTabularFilterPopoverPos({ top: rect.bottom + 4, left: rect.left });
    setOpenOwnershipTabularFilterKey((prev) => (prev === key ? null : key));
  }, []);

  // ── Filtered + sorted row derivations ────────────────────────────────────
  const filteredSortedTableRows = useMemo(() => {
    let rows = tableRows;
    const active = Object.entries(tabularFilters).filter(([, f]) => !isFilterEmpty(f));
    if (active.length > 0) {
      const colMap = new Map(visibleTabularColumns.map((c) => [c.key, c]));
      rows = rows.filter((row) => {
        const node = row.isNew ? row.node : (tableDrafts[row.key] || row.node);
        return active.every(([key, filter]) => {
          const col = colMap.get(key);
          return col ? matchesTabularFilter(getNodeTabularValue(node, col, nodeList, relList, asOfDate), filter) : true;
        });
      });
    }
    if (tabularSort) {
      const col = visibleTabularColumns.find((c) => c.key === tabularSort.key);
      if (col) {
        const { filterType } = getColumnFilterConfig(col);
        rows = [...rows].sort((a, b) => {
          const an = a.isNew ? a.node : (tableDrafts[a.key] || a.node);
          const bn = b.isNew ? b.node : (tableDrafts[b.key] || b.node);
          const av = getNodeTabularValue(an, col, nodeList, relList, asOfDate);
          const bv = getNodeTabularValue(bn, col, nodeList, relList, asOfDate);
          const aBlank = av == null || String(av).trim() === "";
          const bBlank = bv == null || String(bv).trim() === "";
          if (aBlank && bBlank) return 0;
          if (aBlank) return 1;
          if (bBlank) return -1;
          const cmp = filterType === "range"
            ? (Number(av) || 0) - (Number(bv) || 0)
            : String(av ?? "").localeCompare(String(bv ?? ""), undefined, { sensitivity: "base" });
          return tabularSort.dir === "asc" ? cmp : -cmp;
        });
      }
    }
    return rows;
  }, [tableRows, tabularFilters, tabularSort, visibleTabularColumns, tableDrafts, nodeList, relList]);

  const filteredSortedOwnershipTableRows = useMemo(() => {
    let rows = ownershipTableRows;
    const active = Object.entries(ownershipTabularFilters).filter(([, f]) => !isFilterEmpty(f));
    if (active.length > 0) {
      const colMap = new Map(visibleOwnershipTabularColumns.map((c) => [c.key, c]));
      rows = rows.filter((row) => {
        const rel = { ...row.rel, ...(ownershipTableDrafts[row.key] || {}) };
        return active.every(([key, filter]) => {
          const col = colMap.get(key);
          return col ? matchesTabularFilter(getOwnershipTabularValue(rel, nodeList, col), filter) : true;
        });
      });
    }
    if (ownershipTabularSort) {
      const col = visibleOwnershipTabularColumns.find((c) => c.key === ownershipTabularSort.key);
      if (col) {
        const { filterType } = getColumnFilterConfig(col);
        rows = [...rows].sort((a, b) => {
          const ar = { ...a.rel, ...(ownershipTableDrafts[a.key] || {}) };
          const br = { ...b.rel, ...(ownershipTableDrafts[b.key] || {}) };
          const av = getOwnershipTabularValue(ar, nodeList, col);
          const bv = getOwnershipTabularValue(br, nodeList, col);
          const aBlank = av == null || String(av).trim() === "";
          const bBlank = bv == null || String(bv).trim() === "";
          if (aBlank && bBlank) return 0;
          if (aBlank) return 1;
          if (bBlank) return -1;
          const cmp = filterType === "range"
            ? (Number(av) || 0) - (Number(bv) || 0)
            : String(av ?? "").localeCompare(String(bv ?? ""), undefined, { sensitivity: "base" });
          return ownershipTabularSort.dir === "asc" ? cmp : -cmp;
        });
      }
    }
    return rows;
  }, [ownershipTableRows, ownershipTabularFilters, ownershipTabularSort, visibleOwnershipTabularColumns, nodeList, ownershipTableDrafts]);

  const activeTabularFilterCount = useMemo(() =>
    Object.values(tabularFilters).filter((f) => !isFilterEmpty(f)).length,
    [tabularFilters]);
  const activeOwnershipTabularFilterCount = useMemo(() =>
    Object.values(ownershipTabularFilters).filter((f) => !isFilterEmpty(f)).length,
    [ownershipTabularFilters]);

  const exportActiveTabularViewToExcel = useCallback(() => {
    const isOwnershipView = tabularSubMode === "ownerships";
    const viewName = isOwnershipView
      ? (activeOwnershipTabularView.name || "Ownerships")
      : (activeTabularView.name || (tabularSubMode === "entities" ? "Entities" : "Persons"));
    const visibleColumns = isOwnershipView ? visibleOwnershipTabularColumns : visibleTabularColumns;
    const rows = isOwnershipView ? filteredSortedOwnershipTableRows : filteredSortedTableRows;
    const headers = visibleColumns.map((column) => column.label);
    const dataRows = rows.map((row) => {
      const values = visibleColumns.map((column) => {
        const raw = isOwnershipView
          ? getOwnershipTabularValue({ ...row.rel, ...(ownershipTableDrafts[row.key] || {}) }, nodeList, column)
          : getNodeTabularValue(row.isNew ? row.node : (tableDrafts[row.key] || row.node), column, nodeList, relList, asOfDate);
        if (raw == null) return "";
        if (Array.isArray(raw)) return raw.join(", ");
        if (typeof raw === "object") return JSON.stringify(raw);
        return String(raw);
      });
      return values;
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
    worksheet["!cols"] = headers.map((header, idx) => {
      const maxLen = Math.max(
        String(header || "").length,
        ...dataRows.map((row) => String(row[idx] || "").length),
      );
      return { wch: Math.min(Math.max(maxLen + 2, 10), 60) };
    });

    const workbook = XLSX.utils.book_new();
    const sheetName = String(viewName || "Tabular View").slice(0, 31);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    const exportClientName = clientDisplayName || toSentenceCase(clientId) || "export";
    const safeFileName = `${exportClientName}-${tabularSubMode}-${viewName}`
      .replace(/[^\w\s.-]/g, "")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "") || "tabular-view";
    XLSX.writeFile(workbook, `${safeFileName}.xlsx`);
  }, [activeOwnershipTabularView.name, activeTabularView.name, clientDisplayName, clientId, filteredSortedOwnershipTableRows, filteredSortedTableRows, nodeList, ownershipTableDrafts, relList, tableDrafts, tabularSubMode, visibleOwnershipTabularColumns, visibleTabularColumns]);

  // ── Column width computation ───────────────────────────────────────────────
  // Priority: 1) column.width (hand-set), 2) max data content width, 3) never > 40% of screen
  const tabularColumnWidths = useMemo(() => {
    const maxPx = (typeof window !== "undefined" ? window.innerWidth : 1280) * 0.4;
    const CELL_FONT = 13;
    const CELL_PAD = 32;   // td padding (8×2) + input padding (7×2) + input border (1×2)
    const SELECT_ARROW = 28; // extra room for the dropdown chevron on enum columns

    return visibleTabularColumns.map((column) => {
      const viewW = activeTabularView.columnWidths?.[column.key];
      if (viewW > 0) return Math.min(viewW, maxPx);
      if (column.width) return Math.min(column.width, maxPx);

      const { filterType } = getColumnFilterConfig(column);
      const extraPad = filterType === "enum" ? SELECT_ARROW : 0;

      // Auto: start with header label width
      let maxW = 60; // minimum; header is excluded (it wraps)

      // Measure every row's display value — use nodeList directly (not tableRows/tableDrafts)
      // so that column widths are stable and don't shift when cell content changes at runtime.
      for (const node of nodeList) {
        const val = String(getNodeTabularValue(node, column, nodeList, relList, asOfDate) ?? "");
        const w = measureTextWidth(val, CELL_FONT) + CELL_PAD + extraPad;
        if (w > maxW) maxW = w;
      }

      return Math.min(Math.max(Math.ceil(maxW), 80), maxPx);
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visibleTabularColumns, activeTabularView, nodeList.length]);

  const ownershipTabularColumnWidths = useMemo(() => {
    const maxPx = (typeof window !== "undefined" ? window.innerWidth : 1280) * 0.4;
    const CELL_FONT = 13;
    const CELL_PAD = 32;
    const SELECT_ARROW = 28;

    return visibleOwnershipTabularColumns.map((column) => {
      const viewW = activeOwnershipTabularView.columnWidths?.[column.key];
      if (viewW > 0) return Math.min(viewW, maxPx);
      if (column.width) return Math.min(column.width, maxPx);

      const { filterType } = getColumnFilterConfig(column);
      const extraPad = filterType === "enum" ? SELECT_ARROW : 0;

      let maxW = 60; // minimum; header is excluded (it wraps)

      for (const row of ownershipTableRows) {
        const rel = { ...row.rel, ...(ownershipTableDrafts[row.key] || {}) };
        const val = String(getOwnershipTabularValue(rel, nodeList, column) ?? "");
        const w = measureTextWidth(val, CELL_FONT) + CELL_PAD + extraPad;
        if (w > maxW) maxW = w;
      }

      return Math.min(Math.max(Math.ceil(maxW), 80), maxPx);
    });
  }, [visibleOwnershipTabularColumns, ownershipTableRows, ownershipTableDrafts, nodeList, activeOwnershipTabularView]);

  const updateOwnershipTableDraft = useCallback((relId, patch) => {
    setOwnershipTableDrafts((prev) => ({ ...prev, [relId]: { ...(prev[relId] || {}), ...patch } }));
    setOwnershipTableDirtyKeys((prev) => { const next = new Set(prev); next.add(relId); return next; });
    setOwnershipTableRowErrors((prev) => { const next = { ...prev }; delete next[relId]; return next; });
  }, []);

  const saveOwnershipTableRow = useCallback(async (relId) => {
    const original = ownerships.find((r) => r.id === relId);
    if (!original) return false;
    const draft = ownershipTableDrafts[relId];
    if (!draft) {
      setOwnershipTableDirtyKeys((prev) => { const next = new Set(prev); next.delete(relId); return next; });
      return true;
    }

    const rawPercent = draft.percent !== undefined ? draft.percent : (original.percent ?? null);
    const parsedPercent = parseOwnershipPercent(rawPercent);
    if (!parsedPercent.ok) {
      setOwnershipTableRowErrors((prev) => ({ ...prev, [relId]: parsedPercent.message }));
      return false;
    }

    setOwnershipTableSavingKeys((prev) => { const next = new Set(prev); next.add(relId); return next; });
    try {
      const payload = {
        from: original.from,
        to: original.to,
        percent: parsedPercent.value,
        startDate: draft.startDate !== undefined ? (draft.startDate || null) : (original.startDate || null),
        endDate: draft.endDate !== undefined ? (draft.endDate || null) : (original.endDate || null),
        asOfDate: asOfDate || todayIso,
        customFields: { ...(original.customFields || {}), ...(draft.customFields || {}) },
        client: clientId,
      };
      await apiRequest("/api/relationships/owns", { method: "PUT", body: JSON.stringify(payload) });
      setRelList((prev) => prev.map((r) => r.id === relId ? { ...r, ...payload } : r));
      setOwnershipTableDrafts((prev) => { const next = { ...prev }; delete next[relId]; return next; });
      setOwnershipTableDirtyKeys((prev) => { const next = new Set(prev); next.delete(relId); return next; });
      setOwnershipTableRowErrors((prev) => { const next = { ...prev }; delete next[relId]; return next; });
      setRemoteStatus("connected");
      return true;
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
      setOwnershipTableRowErrors((prev) => ({ ...prev, [relId]: err.message || "Unable to save." }));
      return false;
    } finally {
      setOwnershipTableSavingKeys((prev) => { const next = new Set(prev); next.delete(relId); return next; });
    }
  }, [apiRequest, asOfDate, clientId, ownerships, ownershipTableDrafts, todayIso]);

  const pendingOwnershipKeys = useMemo(() => ownershipTableDirtyKeys, [ownershipTableDirtyKeys]);

  const saveAllOwnershipRows = useCallback(async () => {
    for (const key of [...pendingOwnershipKeys]) {
      // eslint-disable-next-line no-await-in-loop
      await saveOwnershipTableRow(key);
    }
  }, [pendingOwnershipKeys, saveOwnershipTableRow]);

  // ── Ownership tabular cell renderer ───────────────────────────────────────
  const renderOwnershipTabularCell = useCallback((column, { row, rowRel, isSaving, isDirty, rowError }) => {
    const draft = ownershipTableDrafts[row.key] || {};
    const percent = draft.percent !== undefined ? draft.percent : (rowRel.percent ?? "");
    const startDate = draft.startDate !== undefined ? draft.startDate : (rowRel.startDate || "");
    const endDate = draft.endDate !== undefined ? draft.endDate : (rowRel.endDate || "");
    const customFields = { ...(rowRel.customFields || {}), ...(draft.customFields || {}) };
    const ownerNode = nodeList.find((n) => n.id === rowRel.from);
    const ownedNode = nodeList.find((n) => n.id === rowRel.to);

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
            ) : (
              <span className="tabular-status tabular-status--saved">Saved</span>
            )}
          </td>
        );
      case "owner":
        return <td key={`${row.key}-owner`}><span className="tabular-cell-ro">{ownerNode?.name || rowRel.from}</span></td>;
      case "owned":
        return <td key={`${row.key}-owned`}><span className="tabular-cell-ro">{ownedNode?.name || rowRel.to}</span></td>;
      case "percent":
        return (
          <td key={`${row.key}-percent`}>
            <input
              className="tabular-cell-input"
              type="number"
              min="0" max="100" step="0.01"
              value={percent}
              onChange={(e) => updateOwnershipTableDraft(row.key, { percent: e.target.value })}
              onBlur={() => saveOwnershipTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
            />
          </td>
        );
      case "startDate":
        return (
          <td key={`${row.key}-startDate`}>
            <input
              className="tabular-cell-input"
              type="date"
              value={startDate}
              onChange={(e) => updateOwnershipTableDraft(row.key, { startDate: e.target.value })}
              onBlur={() => saveOwnershipTableRow(row.key)}
              disabled={isSaving}
            />
          </td>
        );
      case "endDate":
        return (
          <td key={`${row.key}-endDate`}>
            <input
              className="tabular-cell-input"
              type="date"
              value={endDate}
              onChange={(e) => updateOwnershipTableDraft(row.key, { endDate: e.target.value })}
              onBlur={() => saveOwnershipTableRow(row.key)}
              disabled={isSaving}
            />
          </td>
        );
      case "actions":
        return (
          <td key={`${row.key}-actions`}>
            <div className="tabular-actions">
              <Button type="button" variant="outline" onClick={() => saveOwnershipTableRow(row.key)} disabled={isSaving || !isDirty}>Save</Button>
            </div>
          </td>
        );
      default: {
        if (!column.field) return <td key={`${row.key}-${column.key}`}></td>;
        const field = column.field;
        const validValues = Array.isArray(field.validValues) ? field.validValues.filter(Boolean) : [];
        const currentVal = customFields?.[field.fieldId];
        if (field.dataType === "boolean") {
          return (
            <td key={`${row.key}-${field.fieldId}`}>
              <select
                className="tabular-cell-input"
                value={tableFieldToString(field, currentVal)}
                onChange={(e) => updateOwnershipTableDraft(row.key, { customFields: { ...customFields, [field.fieldId]: parseTableFieldValue(field, e.target.value) } })}
                onBlur={() => saveOwnershipTableRow(row.key)}
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
                onChange={(e) => updateOwnershipTableDraft(row.key, { customFields: { ...customFields, [field.fieldId]: parseTableFieldValue(field, e.target.value) } })}
                onBlur={() => saveOwnershipTableRow(row.key)}
                disabled={isSaving}
              >
                <option value=""></option>
                {validValues.map((v) => <option key={v} value={v}>{v}</option>)}
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
              onChange={(e) => updateOwnershipTableDraft(row.key, { customFields: { ...customFields, [field.fieldId]: parseTableFieldValue(field, e.target.value) } })}
              onBlur={() => saveOwnershipTableRow(row.key)}
              onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); }}
              disabled={isSaving}
              autoComplete="off"
              data-lpignore="true"
            />
          </td>
        );
      }
    }
  }, [nodeList, ownershipTableDrafts, saveOwnershipTableRow, updateOwnershipTableDraft]);

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
          <td key={`${row.key}-actions`} className="tabular-row-actions-cell">
            {row.isNew ? (
              <div className="tabular-actions">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => saveTableRow(row.key)}
                  disabled={isSaving || !isDirty}
                >
                  Save
                </Button>
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => removeTableNewRow(row.key)}
                  disabled={isSaving}
                >
                  Cancel
                </Button>
              </div>
            ) : null}
          </td>
        );
      default:
        if (!column.field) return <td key={`${row.key}-${column.key}`}></td>;
        {
          const field = column.field;
          const applicable = normalizeAppliesTo(field.appliesTo).includes(rowNode.kind);
          if (field._virtual) {
            if (!applicable) {
              return <td key={`${row.key}-${field.fieldId}`} className="tabular-na">-</td>;
            }
            const virtualValue = getEntityOwnershipSummary(rowNode, nodeList, relList, asOfDate);
            return (
              <td key={`${row.key}-${field.fieldId}`} className="tabular-cell-wrap" title={virtualValue || ""}>
                {virtualValue || ""}
              </td>
            );
          }
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
    nodeList,
    relList,
    parseTableFieldValue,
    removeTableNewRow,
    saveTableRow,
    setEditNodeId,
    setFocusId,
    setOpenDialog,
    setViewMode,
    tableFieldToString,
    updateTableRowDraft,
  ]);

  const makeRelId = () => `rel-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const openOwnerEditor = async (targetId) => {
    // Fetch ownership timeline from server
    try {
      const timelineResp = await apiRequest(`/api/ownership/history/${encodeURIComponent(targetId)}`);
      setOwnershipTimeline(timelineResp.periods || []);
    } catch (err) {
      console.error("Failed to fetch ownership timeline:", err);
      setOwnershipTimeline([]);
    }

    // Find the period that's current as of asOfDate
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

    // Capture the effective date range from the first ownership record
    const firstOwnership = getOwnersOf(relList, targetId)[0];
    const dateRange = {
      from: firstOwnership?.rel?.effectiveFrom || null,
      to: firstOwnership?.rel?.effectiveTo || null,
      isCurrent: !firstOwnership?.rel?.effectiveTo, // No end date = current
    };

    setOwnerEditorRows(currentOwners);
    setOwnerEditorOriginal(currentOwners);
    setOwnerEditorDateRange(dateRange);
    setOwnerEditorMode("view"); // Start in view mode
    setOwnershipSelectedPeriodSetId(firstOwnership?.rel?.setId || null);
    setOwnershipDeleteConfirm(false); // Reset delete confirmation
    setOwnerSearch("");
    setOwnerSearchOpen(false);
    setOwnerEditorEffectiveDate("");
    setOpenDialog({ type: "edit-owners", targetId });
  };

  const saveOwnerEditor = async () => {
    const targetId = openDialog?.targetId;
    if (!targetId) return;
    const effectiveAsOf = normalizeDateInput(ownerEditorEffectiveDate);
    if (!effectiveAsOf) {
      setRemoteStatus("error");
      setRemoteError("Effective as of date is required.");
      return;
    }

    const hasOutOfRangePercent = ownerEditorRows.some((r) => {
      if (r.percent === "") return false;
      return !parseOwnershipPercent(r.percent).ok;
    });
    if (hasOutOfRangePercent) return;

    const ownerTotal = ownerEditorRows.reduce(
      (s, r) => s + (r.percent !== "" && !isNaN(Number(r.percent)) ? Number(r.percent) : 0), 0
    );
    const ownerTotalInvalid = Math.abs(ownerTotal - 100) > 0.0001;
    if (ownerTotalInvalid) return;

    setIsSavingOwners(true);
    try {
      const ownerData = ownerEditorRows.map((row) => ({
        from: row.nodeId,
        percent: row.percent !== "" ? Number(row.percent) : null,
      }));

      // Use different endpoint based on mode
      if (ownerEditorMode === "edit-existing") {
        // Update existing period: change date and/or owners
        await apiRequest("/api/ownership/update-period", {
          method: "PUT",
          body: JSON.stringify({
            to: targetId,
            setId: ownershipSelectedPeriodSetId,
            newEffectiveFrom: effectiveAsOf,
            owners: ownerData,
            client: clientId,
          }),
        });
      } else {
        // Create new period
        await apiRequest("/api/ownership/sets", {
          method: "PUT",
          body: JSON.stringify({
            to: targetId,
            asOfDate: effectiveAsOf,
            owners: ownerData,
            client: clientId,
          }),
        });
      }

      // Reload directory data
      await loadDirectory();

      // Reload the ownership timeline for this entity
      try {
        const timelineResp = await apiRequest(`/api/ownership/history/${encodeURIComponent(targetId)}`);
        setOwnershipTimeline(timelineResp.periods || []);

        // Find and reload the period we just saved/updated
        const updatedPeriod = timelineResp.periods?.find((p) => p.setId === ownershipSelectedPeriodSetId);
        if (updatedPeriod) {
          const refreshedRows = updatedPeriod.owners.map((o) => {
            const node = getNode(nodeList, o.from);
            return {
              nodeId: o.from,
              name: node?.name ?? o.from,
              percent: String(o.percent),
              startDate: "",
              endDate: "",
              isNew: false,
            };
          });
          setOwnerEditorRows(refreshedRows);
          setOwnerEditorOriginal(refreshedRows);
          setOwnerEditorDateRange({
            from: updatedPeriod.effectiveFrom,
            to: updatedPeriod.effectiveTo,
            isCurrent: !updatedPeriod.effectiveTo,
          });
        }
      } catch (err) {
        console.error("Failed to reload ownership timeline:", err);
      }

      // Return to view mode instead of closing
      setOwnerEditorMode("view");
      setOwnerEditorEffectiveDate("");
      setRemoteStatus("connected");
    } catch (err) {
      setRemoteStatus("error");
      setRemoteError(err.message);
    } finally {
      setIsSavingOwners(false);
    }
  };

  const deleteOwnershipGroup = async () => {
    const targetId = openDialog?.targetId;
    if (!targetId || !ownershipSelectedPeriodSetId) return;

    setIsSavingOwners(true);
    try {
      await apiRequest(`/api/ownership/group/${encodeURIComponent(targetId)}/${encodeURIComponent(ownershipSelectedPeriodSetId)}`, {
        method: "DELETE",
      });

      // Reload directory data
      await loadDirectory();

      // Reload the ownership timeline for this entity
      try {
        const timelineResp = await apiRequest(`/api/ownership/history/${encodeURIComponent(targetId)}`);
        setOwnershipTimeline(timelineResp.periods || []);

        // If there are other periods, display the first one; otherwise close dialog
        if (timelineResp.periods && timelineResp.periods.length > 0) {
          const newPeriod = timelineResp.periods[0];
          const newRows = newPeriod.owners.map((o) => {
            const node = getNode(nodeList, o.from);
            return {
              nodeId: o.from,
              name: node?.name ?? o.from,
              percent: String(o.percent),
              startDate: "",
              endDate: "",
              isNew: false,
            };
          });
          setOwnershipSelectedPeriodSetId(newPeriod.setId);
          setOwnerEditorRows(newRows);
          setOwnerEditorOriginal(newRows);
          setOwnerEditorDateRange({
            from: newPeriod.effectiveFrom,
            to: newPeriod.effectiveTo,
            isCurrent: !newPeriod.effectiveTo,
          });
        } else {
          // No ownership groups left, close the dialog
          setOpenDialog(null);
        }
      } catch (err) {
        console.error("Failed to reload ownership timeline:", err);
      }

      setOwnershipDeleteConfirm(false);
      setRemoteStatus("connected");
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
        body: JSON.stringify({ name, kind, asOfDate: asOfDate || todayIso, client: clientId }),
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
        id: node.id,
        name: node.name,
        kind: node.kind,
        photo: node.photo || "",
        logo: node.logo || "",
        address: node.address || "",
        workPhone: node.workPhone || "",
        cellPhone: node.cellPhone || "",
        emails: node.emails || "",
        taxId: node.taxId || "",
        customFields: node.customFields || {},
      });
      setEditNodeEffectiveDate("");
      
      // For entities, fetch ownership timeline to show in the ownership records field
      if (node.kind === "entity") {
        apiRequest(`/api/ownership/history/${encodeURIComponent(node.id)}`)
          .then((resp) => setEditNodeOwnershipTimeline(resp.periods || []))
          .catch(() => setEditNodeOwnershipTimeline([]));
      } else {
        setEditNodeOwnershipTimeline([]);
      }
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
        if (Array.isArray(data?.ownershipTabularViews)) {
          const oviews = data.ownershipTabularViews.map(sanitizeOwnershipTabularView).filter(Boolean);
          setOwnershipTabularViews(oviews);
          setSelectedOwnershipTabularViewId(data?.ownershipTabularViewsSelectedId || DEFAULT_OWNERSHIP_TABULAR_VIEW_ID);
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
      const interactiveSelector = "button, input, textarea, select, a, label, [role=\"button\"], .oc-minimap, .oc-minimap *, [data-hv-node-id]";
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
    const center = (behavior) => {
      if (!focusBoxRef.current || !hierarchyContainerRef.current || !hierarchyStageRef.current) return;
      const container = hierarchyContainerRef.current;
      const focusBox = focusBoxRef.current;
      const stage = hierarchyStageRef.current;
      const containerRect = container.getBoundingClientRect();
      const stageRect = stage.getBoundingClientRect();
      const focusRect = focusBox.getBoundingClientRect();
      const clamp = (value, min, max) => (min <= max ? Math.min(Math.max(value, min), max) : value);
      const centeredDeltaX = (stageRect.left + stageRect.width / 2) - (containerRect.left + container.clientWidth / 2);
      const centeredDeltaY = (stageRect.top + stageRect.height / 2) - (containerRect.top + container.clientHeight / 2);

      // Keep the focused entity inside a small safety inset so it never scrolls off screen.
      const insetX = Math.min(24, Math.max(0, Math.floor(container.clientWidth / 4)));
      const insetY = Math.min(24, Math.max(0, Math.floor(container.clientHeight / 4)));
      const focusMinDeltaX = (focusRect.right - containerRect.left) - (container.clientWidth - insetX);
      const focusMaxDeltaX = (focusRect.left - containerRect.left) - insetX;
      const focusMinDeltaY = (focusRect.bottom - containerRect.top) - (container.clientHeight - insetY);
      const focusMaxDeltaY = (focusRect.top - containerRect.top) - insetY;

      const nextLeft = container.scrollLeft + clamp(centeredDeltaX, focusMinDeltaX, focusMaxDeltaX);
      const nextTop = container.scrollTop + clamp(centeredDeltaY, focusMinDeltaY, focusMaxDeltaY);
      container.scrollTo({ left: nextLeft, top: nextTop, behavior });
    };
    requestAnimationFrame(() => center("smooth"));
    // Re-center after images in the focus box have loaded (photos/logos load asynchronously)
    const t = setTimeout(() => center("auto"), 300);
    return () => clearTimeout(t);
  }, [focusId, viewMode, nodeList, explodedNodes.size]);

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

  useEffect(() => {
    if (quickFindMatches.length === 0) {
      setQuickFindHighlight(-1);
      return;
    }
    setQuickFindHighlight((prev) => {
      if (prev < 0 || prev >= quickFindMatches.length) return 0;
      return prev;
    });
  }, [quickFindMatches]);

  useEffect(() => {
    const q = String(quickFindQuery || "").trim();
    if (q.length < 2) {
      setQuickFindOpen(false);
    }
  }, [quickFindQuery]);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (quickFindRef.current && !quickFindRef.current.contains(e.target)) {
        setQuickFindOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    const onKeydown = (e) => {
      if ((e.ctrlKey || e.metaKey) && String(e.key || "").toLowerCase() === "k") {
        e.preventDefault();
        quickFindInputRef.current?.focus();
        const q = String(quickFindQuery || "").trim();
        if (q.length >= 2) setQuickFindOpen(true);
        return;
      }
      if (e.key === "Escape" && quickViewNodeId) {
        setQuickViewNodeId("");
      }
    };
    window.addEventListener("keydown", onKeydown);
    return () => window.removeEventListener("keydown", onKeydown);
  }, [quickFindQuery, quickViewNodeId]);


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
  const handleQuickFindSelect = useCallback((nodeId) => {
    setQuickViewNodeId(nodeId);
    setQuickFindQuery("");
    setQuickFindHighlight(-1);
    setQuickFindOpen(false);
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
        <div style={{ maxWidth: "90%", margin: "0 auto", display: "flex", flexDirection: "column", gap: 8 }}>
          {/* TOP ROW: Logo + Client Name | View Selector + Date + Settings */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 20 }}>
            {/* LEFT ZONE: Logo + Client Name */}
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexShrink: 0 }}>
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
                <img src="/emplus-logo.png" alt="EMPlus" style={{ height: 40, width: "auto", borderRadius: "20px" }} />
              </button>
              <div style={{ fontSize: 16, fontWeight: 600, color: "#1a1a2e" }}>
                {clientDisplayName || toSentenceCase(clientId)}
              </div>
            </div>

            {/* RIGHT ZONE: View Selector + Date + Settings */}
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexShrink: 0 }}>
              {/* View Dropdown */}
              <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
                <select
                  value={viewMode}
                  onChange={(e) => setViewMode(e.target.value)}
                  style={{
                    padding: "6px 12px",
                    fontSize: 13,
                    fontWeight: 500,
                    border: "1px solid #cbd5e1",
                    borderRadius: 6,
                    background: "#fff",
                    color: "#475569",
                    cursor: "pointer",
                    appearance: "none",
                    paddingRight: 28,
                  }}
                >
                  <option value="hierarchy">Hierarchy</option>
                  <option value="directory">Directory</option>
                  <option value="tabular">Tabular</option>
                </select>
                <span style={{ position: "absolute", right: 8, pointerEvents: "none", color: "#6b7280" }}>
                  ▼
                </span>
              </div>

              {/* As of Date — always visible */}
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ fontSize: 12, color: "#6b7280", whiteSpace: "nowrap" }}>As of</span>
                  <input
                    type="date"
                    value={asOfDate}
                    max={todayIso}
                    onChange={(e) => setAsOfDate(e.target.value)}
                    style={{
                      border: "1px solid #cbd5e1",
                      borderRadius: 6,
                      padding: "4px 8px",
                      fontSize: 12,
                      color: "#1f2937",
                      background: "#fff",
                    }}
                  />
                  {asOfDate && (
                    <button
                      type="button"
                      onClick={() => setAsOfDate("")}
                      style={{
                        border: "1px solid #cbd5e1",
                        borderRadius: 6,
                        padding: "3px 8px",
                        fontSize: 11,
                        cursor: "pointer",
                        background: "#fff",
                        color: "#475569",
                        whiteSpace: "nowrap",
                      }}
                    >
                      Current
                    </button>
                  )}
                </div>

              {/* Settings Gear */}
              <div className="settings-anchor" ref={settingsRef}>
                <Button
                  ref={homeButtonRef}
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
                    <button
                      className="settings-menu-item"
                      onClick={() => {
                        setShowStats(!showStats);
                        setSettingsOpen(false);
                      }}
                    >
                      <LayoutList size={15} />
                      {showStats ? "Hide" : "Show"} Statistics
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
            {/* end RIGHT ZONE */}
          </div>
          {/* end TOP ROW */}

          {/* SEARCH ROW */}
          <div style={{ display: "flex", alignItems: "center", width: "100%" }}>
            <div className="quick-find" ref={quickFindRef} style={{ width: "100%", maxWidth: "none", border: "none", background: "transparent", minHeight: "auto", padding: 0 }}>
              <div className="quick-find-input-wrap" style={{ border: "none", background: "transparent", padding: 0 }}>
                <Search size={14} className="quick-find-icon" style={{ color: "#6b7280" }} />
                <input
                  ref={quickFindInputRef}
                  type="text"
                  className="quick-find-input"
                  style={{ border: "none", fontSize: 14, background: "transparent", color: "#1f2937", padding: 0 }}
                  placeholder="Lightning search..."
                  value={quickFindQuery}
                  onChange={(e) => {
                    const next = e.target.value;
                    setQuickFindQuery(next);
                    setQuickFindOpen(next.trim().length >= 2);
                  }}
                  onFocus={() => {
                    if (quickFindQuery.trim().length >= 2) setQuickFindOpen(true);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "ArrowDown") {
                      e.preventDefault();
                      if (!quickFindOpen && quickFindMatches.length > 0) {
                        setQuickFindOpen(true);
                        return;
                      }
                      setQuickFindHighlight((prev) => {
                        if (quickFindMatches.length === 0) return -1;
                        return Math.min((prev < 0 ? 0 : prev + 1), quickFindMatches.length - 1);
                      });
                    } else if (e.key === "ArrowUp") {
                      e.preventDefault();
                      setQuickFindHighlight((prev) => {
                        if (quickFindMatches.length === 0) return -1;
                        return Math.max((prev < 0 ? 0 : prev - 1), 0);
                      });
                    } else if (e.key === "Enter") {
                      if (!quickFindOpen || quickFindHighlight < 0 || quickFindHighlight >= quickFindMatches.length) return;
                      e.preventDefault();
                      handleQuickFindSelect(quickFindMatches[quickFindHighlight].id);
                    } else if (e.key === "Escape") {
                      setQuickFindOpen(false);
                    }
                  }}
                />
              </div>
              {quickFindOpen && (
                <div className="quick-find-menu" style={{ left: 0, right: "auto" }}>
                  {quickFindMatches.length === 0 ? (
                    <div className="quick-find-empty">No matches. Keep typing...</div>
                  ) : (
                    quickFindMatches.map((n, idx) => (
                      <button
                        key={n.id}
                        type="button"
                        className={`quick-find-item${idx === quickFindHighlight ? " is-active" : ""}`}
                        onMouseEnter={() => setQuickFindHighlight(idx)}
                        onClick={() => handleQuickFindSelect(n.id)}
                      >
                        <span className="quick-find-item-name">{n.name || n.id}</span>
                        <span className="quick-find-item-meta">{toSentenceCase(n.kind)} · {n.id}</span>
                      </button>
                    ))
                  )}
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

        {/* Notification Toast */}
        {notification && (
          <div style={{
            position: "fixed",
            top: 16,
            right: 16,
            zIndex: 9999,
            maxWidth: 420,
            padding: "16px",
            borderRadius: "8px",
            boxShadow: "0 10px 25px rgba(0,0,0,0.2)",
            backgroundColor: notification.type === "success" ? "#ecfdf5" : "#fef2f2",
            border: `2px solid ${notification.type === "success" ? "#10b981" : "#ef4444"}`,
            animation: "slideIn 0.3s ease-out",
          }}>
            <style>{`
              @keyframes slideIn {
                from {
                  transform: translateX(420px);
                  opacity: 0;
                }
                to {
                  transform: translateX(0);
                  opacity: 1;
                }
              }
            `}</style>
            
            {/* Close button */}
            <button
              onClick={() => setNotification(null)}
              style={{
                position: "absolute",
                top: 8,
                right: 8,
                background: "none",
                border: "none",
                fontSize: "20px",
                color: notification.type === "success" ? "#047857" : "#991b1b",
                cursor: "pointer",
                padding: "4px 8px",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
              title="Dismiss"
            >
              ✕
            </button>

            <div style={{
              fontSize: "15px",
              fontWeight: 700,
              color: notification.type === "success" ? "#065f46" : "#7f1d1d",
              marginBottom: "4px",
              paddingRight: "24px",
            }}>
              {notification.title}
            </div>
            <div style={{
              fontSize: "13px",
              color: notification.type === "success" ? "#047857" : "#991b1b",
              marginBottom: notification.details ? "8px" : "0",
            }}>
              {notification.message}
            </div>
            {notification.details && (
              <div style={{
                fontSize: "13px",
                color: notification.type === "success" ? "#059669" : "#b91c1c",
                paddingTop: "8px",
                borderTop: `1px solid ${notification.type === "success" ? "#a7f3d0" : "#fecaca"}`,
              }}>
                {notification.details.imported !== undefined && (
                  <div style={{ marginTop: "6px" }}>
                    <strong>Imported:</strong> {notification.details.imported} / {notification.details.total}
                  </div>
                )}
                {notification.details.skipped > 0 && (
                  <div style={{ marginTop: "4px" }}>
                    <strong>Skipped:</strong> {notification.details.skipped}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

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
                  <div style={{ display: "flex", gap: 5, flexDirection: "column", marginTop: 8, flexWrap: "wrap" }}>
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
                    style={{ width: "90%" }}
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
                <div>
                  <Label htmlFor="ownership-import-as-of">Effective as of</Label>
                  <Input
                    id="ownership-import-as-of"
                    style={{ width: "80%" }}
                    type="date"
                    value={uploadOwnershipAsOfDate}
                    onChange={(event) => {
                      setUploadOwnershipAsOfDate(event.target.value);
                      setUploadError("");
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
                {uploadStatus === "loading-background" && (
                  <div style={{ color: "#1f2937", fontSize: 14, padding: "12px 0" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ 
                        width: 20, 
                        height: 20, 
                        border: "3px solid #e5e7eb", 
                        borderTop: "3px solid #3b82f6", 
                        borderRadius: "50%", 
                        animation: "spin 1s linear infinite"
                      }} />
                      <span>
                        <strong>Loading {uploadFile?.name || "file"} in the background...</strong><br/>
                        <span style={{ fontSize: 12, color: "#6b7280" }}>We'll notify you when complete. You can close this dialog.</span>
                      </span>
                    </div>
                    <style>{`
                      @keyframes spin {
                        from { transform: rotate(0deg); }
                        to { transform: rotate(360deg); }
                      }
                    `}</style>
                  </div>
                )}
                {uploadStatus === "error" && uploadError && (
                  <div style={{ color: "#dc2626", fontSize: 14 }}>{uploadError}</div>
                )}
                {uploadType === "ownership" && !ownershipPreviewValidation.valid && (
                  <div style={{ border: "1px solid #fca5a5", background: "#fff1f2", color: "#9f1239", borderRadius: 8, padding: 10, display: "grid", gap: 6 }}>
                    <div style={{ fontSize: 13, fontWeight: 700 }}>
                      Import blocked: each owned entity must total exactly 0% or 100% ownership.
                    </div>
                    <div style={{ fontSize: 12 }}>
                      Offending entities:
                    </div>
                    <div style={{ maxHeight: 140, overflowY: "auto", overscrollBehavior: "contain", fontSize: 12 }}>
                      {ownershipPreviewValidation.offendingEntities.map((entry, idx) => (
                        <div key={`${entry.owned}-${idx}`}>
                          {entry.owned}: {entry.total}% ({entry.reason})
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {!normalizeDateInput(uploadOwnershipAsOfDate) && (
                  <div style={{ color: "#dc2626", fontSize: 13 }}>
                    Effective as of date is required before importing rows.
                  </div>
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

            <DialogFooter style={{ marginTop: "36px" }}>
              {uploadStatus === "success" ? (
                <Button type="button" variant="outline" onClick={() => setUploadOpen(false)}>
                  Close
                </Button>
              ) : uploadStatus === "loading-background" ? (
                <Button type="button" variant="outline" onClick={() => setUploadOpen(false)}>
                  Close
                </Button>
              ) : uploadPreview ? (
                <>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => { setUploadPreview(null); setUploadStatus("idle"); setUploadError(""); }}
                    disabled={uploadStatus === "uploading" || uploadStatus === "loading-background"}
                  >
                    Back
                  </Button>
                  <Button
                    type="button"
                    onClick={handleUploadCsv}
                    disabled={
                      uploadStatus === "uploading" ||
                      uploadStatus === "loading-background" ||
                      uploadPreview.total === 0 ||
                      ((uploadType === "ownership" && !ownershipPreviewValidation.valid) || !normalizeDateInput(uploadOwnershipAsOfDate))
                    }
                  >
                    {uploadStatus === "uploading"
                      ? <><Loader2 size={14} className="animate-spin" style={{ marginRight: 6 }} />Importing…</>
                      : uploadStatus === "loading-background"
                        ? <><Loader2 size={14} className="animate-spin" style={{ marginRight: 6 }} />Loading in background…</>
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
              <div className="hv-stage" ref={hierarchyStageRef}>

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
                              onQuickView={setQuickViewNodeId}
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
                                onExplodeAll={handleExplodeAll} onQuickView={setQuickViewNodeId}
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
            {showStats && (
              <StatsStrip
                allEntityNodes={sortedEntityNodes}
                allPersonNodes={sortedPersonNodes}
                filteredEntityNodes={filteredEntityNodes}
                filteredPersonNodes={filteredPersonNodes}
                dataDictionary={dataDictionary}
                isFiltered={!!dirSearchLower}
              />
            )}
            {/* Main content: search and cards in column layout */}
            <div style={{ display: "flex", flexDirection: "column", gap: 16, width: "100%" }}>
              {/* Search bar for directory */}
              <div className="quick-find" style={{ maxWidth: 400, border: "none", background: "transparent", minHeight: "auto", padding: "8px 0" }}>
                <div className="quick-find-input-wrap" style={{ border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "center", gap: 8 }}>
                  <Search size={14} style={{ color: "#6b7280", flexShrink: 0 }} />
                  <input
                    type="text"
                    className="quick-find-input"
                    style={{ border: "none", fontSize: 14, background: "transparent", color: "#1f2937", padding: 0 }}
                    placeholder="Find in directory…"
                    value={dirSearch}
                    onChange={(e) => setDirSearch(e.target.value)}
                  />
                  {dirSearch && (
                    <button
                      type="button"
                      onClick={() => setDirSearch("")}
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        padding: 0,
                        display: "flex",
                        alignItems: "center",
                        color: "#6b7280",
                      }}
                      title="Clear"
                    >
                      <X size={14} />
                    </button>
                  )}
                </div>
              </div>
              {/* Cards container - side by side */}
              <div style={{ display: "flex", gap: 24, width: "100%" }}>
                <Card style={{ width: "45%" }}>
                  <CardContent>
                    <div className="section-title" style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <span>Entities {dirSearchLower ? `(${filteredEntityNodes.length})` : ""}</span>
                      <button
                        type="button"
                        title="Add Entity"
                        onClick={() => {
                          setNewNode({ name: "", kind: "entity", photo: "", logo: "", customFields: {} });
                          setOpenDialog({ type: "add-node" });
                        }}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          padding: "4px 8px",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          color: "#6b7280",
                          fontSize: 16,
                        }}
                      >
                        <Plus size={18} />
                      </button>
                    </div>
                    <div className="directory-scroll">
                      <ul className="directory-list">
                        {filteredEntityNodes.map((n) => (
                          <li key={n.id}>
                            <div
                              className="directory-item"
                              style={{ cursor: "pointer" }}
                              onClick={() => setQuickViewNodeId(n.id)}
                            >
                              {n.logo
                                ? <img src={n.logo} alt="" className="directory-thumb" />
                                : <Building2 className="directory-icon" />}
                              <div style={{ flex: 1, minWidth: 0 }}>
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
                <Card style={{ width: "45%" }}>
                  <CardContent>
                    <div className="section-title" style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      <span>People {dirSearchLower ? `(${filteredPersonNodes.length})` : ""}</span>
                      <button
                        type="button"
                        title="Add Person"
                        onClick={() => {
                          setNewNode({ name: "", kind: "person", photo: "", logo: "", customFields: {} });
                          setOpenDialog({ type: "add-node" });
                        }}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          padding: "4px 8px",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          color: "#6b7280",
                          fontSize: 16,
                        }}
                      >
                        <Plus size={18} />
                      </button>
                    </div>
                    <div className="directory-scroll">
                      <ul className="directory-list">
                        {filteredPersonNodes.map((n) => (
                          <li key={n.id}>
                            <div
                              className="directory-item"
                              style={{ cursor: "pointer" }}
                              onClick={() => setQuickViewNodeId(n.id)}
                            >
                              {n.photo
                                ? <img src={n.photo} alt="" className="directory-thumb directory-thumb--round" />
                                : <Users className="directory-icon" />}
                              <div style={{ flex: 1, minWidth: 0 }}>
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
              {/* End cards container */}
            </div>
            {/* End main content column */}
          </div>
        )}

        {viewMode === "tabular" && (
          <div className="tabular-grid">

            {/* ── Sub-mode toggle: Entities | Persons | Ownerships ── */}
            <div style={{ display: "flex", alignItems: "center", gap: 0, borderRadius: 6, overflow: "hidden", border: "1px solid #cbd5e1", alignSelf: "flex-start" }}>
              <button
                type="button"
                onClick={() => setTabularSubMode("entities")}
                style={{
                  padding: "5px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer", border: "none",
                  borderRight: "1px solid #cbd5e1",
                  background: tabularSubMode === "entities" ? "#1e293b" : "#fff",
                  color: tabularSubMode === "entities" ? "#fff" : "#475569",
                }}
              >
                Entities
              </button>
              <button
                type="button"
                onClick={() => setTabularSubMode("persons")}
                style={{
                  padding: "5px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer", border: "none",
                  background: tabularSubMode === "persons" ? "#1e293b" : "#fff",
                  color: tabularSubMode === "persons" ? "#fff" : "#475569",
                }}
              >
                Persons
              </button>
            </div>

            {/* ── Entities / Persons sub-mode (shared node tabular infrastructure) ── */}
            {(tabularSubMode === "entities" || tabularSubMode === "persons") && (
              <>
                {showStats && (
                  <StatsStrip
                    allEntityNodes={sortedEntityNodes}
                    allPersonNodes={sortedPersonNodes}
                    filteredEntityNodes={filteredEntityNodes}
                    filteredPersonNodes={filteredPersonNodes}
                    dataDictionary={dataDictionary}
                    isFiltered={!!dirSearchLower}
                  />
                )}

                <div className="tabular-toolbar">
                  <div className="tabular-toolbar-left">
                    <select
                      className="tabular-view-select"
                      value={effectiveSelectedTabularViewId}
                      onChange={(e) => setSelectedTabularViewId(e.target.value)}
                    >
                      <option value={DEFAULT_TABULAR_VIEW_ID}>Default</option>
                      {tabularViewsForCurrentKind.map((v) => (
                        <option key={v.id} value={v.id}>{v.name}</option>
                      ))}
                    </select>
                    <Button type="button" variant="outline" onClick={openTabularViewManager}>
                      Customize
                    </Button>
                    {/* Search input after Customize */}
                    <div className="quick-find" style={{ width: 200, maxWidth: "none", border: "none", background: "transparent", minHeight: "auto", padding: 0 }}>
                      <div className="quick-find-input-wrap" style={{ border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "center", gap: 8 }}>
                        <Search size={14} style={{ color: "#6b7280", flexShrink: 0 }} />
                        <input
                          type="text"
                          className="quick-find-input"
                          style={{ border: "none", fontSize: 13, background: "transparent", color: "#1f2937", padding: 0 }}
                          placeholder="Find in this table…"
                          value={dirSearch}
                          onChange={(e) => setDirSearch(e.target.value)}
                        />
                        {dirSearch && (
                          <button
                            type="button"
                            onClick={() => setDirSearch("")}
                            style={{
                              background: "none",
                              border: "none",
                              cursor: "pointer",
                              padding: 0,
                              display: "flex",
                              alignItems: "center",
                              color: "#6b7280",
                            }}
                            title="Clear"
                          >
                            <X size={14} />
                          </button>
                        )}
                      </div>
                    </div>
                    {activeTabularFilterCount > 0 && (
                      <Button type="button" variant="outline" onClick={() => { setTabularFilters({}); setOpenTabularFilterKey(null); }}>
                        Clear Filters ({activeTabularFilterCount})
                      </Button>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <Button
                      type="button"
                      variant="outline"
                      onClick={exportActiveTabularViewToExcel}
                      disabled={tabularSubMode === "entities" || tabularSubMode === "persons"
                        ? filteredSortedTableRows.length === 0
                        : filteredSortedOwnershipTableRows.length === 0}
                    >
                      <FileSpreadsheet size={16} />
                      Export
                    </Button>
                    <Button
                      type="button"
                      variant="outline"
                      onClick={() => {
                        const kind = tabularSubMode === "entities" ? "entity" : "person";
                        setNewNode({ name: "", kind, photo: "", logo: "", customFields: {} });
                        setOpenDialog({ type: "add-node" });
                      }}
                    >
                      <Plus size={16} />
                      Add
                    </Button>
                    <Button
                      type="button"
                      onClick={saveAllTableRows}
                      disabled={pendingTableKeys.size === 0 || tableSavingKeys.size > 0}
                    >
                      Save Changes ({pendingTableKeys.size})
                    </Button>
                  </div>
                </div>

                <Card className="tabular-card">
                  <CardContent style={{ padding: 0 }}>
                    <div className="tabular-wrap">
                      <table
                        className="tabular-table"
                        style={{ width: `${tabularColumnWidths.reduce((a, b) => a + b, 0)}px` }}
                      >
                        <colgroup>
                          {tabularColumnWidths.map((w, i) => (
                            <col key={i} style={{ width: `${w}px` }} />
                          ))}
                        </colgroup>
                        <thead>
                          <tr>
                            {visibleTabularColumns.map((column, colIdx) => {
                              const { filterType, enumOptions } = getColumnFilterConfig(column);
                              const isSorted = tabularSort?.key === column.key;
                              const hasFilter = !isFilterEmpty(tabularFilters[column.key]);
                              const headClass = `${isSorted ? "tabular-th--sorted" : ""}${hasFilter ? " tabular-th--filtered" : ""}${colIdx === 0 ? " tabular-col-frozen" : ""}`;
                              return (
                                <th key={column.key} title={column.label} className={headClass}>
                                  <div className="tabular-th-inner">
                                    {filterType ? (
                                      <span className="tabular-th-label" onClick={() => handleTabularSort(column.key)}>
                                        {column.label}{isSorted && <span className="tabular-sort-icon">{tabularSort.dir === "asc" ? " ↑" : " ↓"}</span>}
                                      </span>
                                    ) : (
                                      <span className="tabular-th-label-plain">{column.label}</span>
                                    )}
                                    {filterType && (
                                      <button type="button" className={`tabular-filter-btn${hasFilter ? " tabular-filter-btn--active" : ""}`}
                                        title={hasFilter ? "Filter active" : "Filter"}
                                        onClick={(e) => { e.stopPropagation(); toggleTabularFilter(column.key, e); }}>
                                        <Filter size={11} />
                                      </button>
                                    )}
                                  </div>
                                </th>
                              );
                            })}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredSortedTableRows.length === 0 && (
                            <tr>
                              <td colSpan={visibleTabularColumns.length} className="tabular-empty">
                                {tableRows.length === 0 ? "No matching rows." : "No rows match the active filters."}
                              </td>
                            </tr>
                          )}
                          {filteredSortedTableRows.map((row) => {
                            const rowNode = row.isNew ? row.node : (tableDrafts[row.key] || row.node);
                            const isSaving = tableSavingKeys.has(row.key);
                            const isDirty = pendingTableKeys.has(row.key);
                            const rowError = tableRowErrors[row.key];
                            const resolvedId = row.isNew
                              ? makeNodeId(rowNode.kind, rowNode.name || "")
                              : makeNodeId(rowNode.kind, rowNode.name || "", row.key);
                            return (
                              <tr key={row.key} className={rowError ? "tabular-row-error" : undefined}>
                                {visibleTabularColumns.map((column, colIdx) => {
                                  const cell = renderTabularCell(column, { row, rowNode, isSaving, isDirty, rowError, resolvedId });
                                  if (!React.isValidElement(cell)) return cell;

                                  let nextCell = cell;
                                  if (colIdx === 0 && !row.isNew) {
                                    nextCell = React.cloneElement(nextCell, {
                                      className: `${nextCell.props.className || ""} tabular-actions-anchor`.trim(),
                                      children: (
                                        <div className="tabular-actions-anchor-box">
                                          <button
                                            className="tabular-edit-btn"
                                            type="button"
                                            title="Edit"
                                            aria-label="Edit"
                                            onClick={() => {
                                              setEditNodeId(row.key);
                                              setOpenDialog({ type: "edit-node" });
                                            }}
                                            disabled={isSaving}
                                          >
                                            <Pencil size={14} />
                                          </button>
                                          {nextCell.props.children}
                                        </div>
                                      ),
                                    });
                                  }

                                  const extraClass = `${colIdx === 0 ? "tabular-col-frozen" : ""}${colIdx === 0 && isDirty ? " tabular-col-dirty" : ""}`.trim();
                                  if (!extraClass) return nextCell;
                                  return React.cloneElement(nextCell, {
                                    className: `${nextCell.props.className || ""} ${extraClass}`.trim(),
                                  });
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>

                {openTabularFilterKey && (() => {
                  const col = visibleTabularColumns.find((c) => c.key === openTabularFilterKey);
                  if (!col) return null;
                  const { filterType, enumOptions } = getColumnFilterConfig(col);
                  return (
                    <ColumnFilterPopover
                      filterType={filterType}
                      enumOptions={enumOptions || []}
                      currentFilter={tabularFilters[openTabularFilterKey] || null}
                      popoverPos={tabularFilterPopoverPos}
                      onChange={(f) => setTabularFilters((p) => ({ ...p, [openTabularFilterKey]: f }))}
                      onClose={() => setOpenTabularFilterKey(null)}
                    />
                  );
                })()}

                {tabularViewDialogOpen && (
                  <Dialog open={tabularViewDialogOpen} onOpenChange={setTabularViewDialogOpen}>
                    <DialogContent style={{ width: "min(760px, 92vw)", maxWidth: "none" }}>
                      <DialogHeader>
                        <DialogTitle>
                          Customize View — Nodes
                          <span style={{ fontWeight: 400, color: '#64748b', fontSize: '14px', marginLeft: 10 }}>{activeTabularView.name}</span>
                        </DialogTitle>
                      </DialogHeader>
                      <div className="tabular-view-editor">
                        <div className="tabular-view-meta-section">
                          <div className="tabular-view-meta-row">
                            <span className="tabular-view-label">Sort:</span>
                            <select
                              className="tabular-meta-select"
                              value={tabularViewDraft.sort?.key || ""}
                              onChange={(e) => {
                                const k = e.target.value;
                                setTabularViewDraft((prev) => k
                                  ? { ...prev, sort: { key: k, dir: prev.sort?.dir || "asc" } }
                                  : { ...prev, sort: null });
                              }}
                            >
                              <option value="">— none —</option>
                              {tabularViewDraft.columnOrder.map((k) => {
                                const col = allTabularColumns.find((c) => c.key === k);
                                if (!col) return null;
                                const { filterType } = getColumnFilterConfig(col);
                                if (!filterType) return null;
                                return <option key={k} value={k}>{col.label}</option>;
                              })}
                            </select>
                            {tabularViewDraft.sort?.key && (
                              <select
                                className="tabular-meta-select"
                                value={tabularViewDraft.sort.dir || "asc"}
                                onChange={(e) => setTabularViewDraft((prev) => ({ ...prev, sort: { ...prev.sort, dir: e.target.value } }))}
                              >
                                <option value="asc">Ascending</option>
                                <option value="desc">Descending</option>
                              </select>
                            )}
                            {tabularViewDraft.sort && (
                              <button type="button" className="tabular-meta-clear" onClick={() => setTabularViewDraft((prev) => ({ ...prev, sort: null }))}>Clear</button>
                            )}
                          </div>
                          {Object.entries(tabularViewDraft.filters || {}).filter(([, f]) => !isFilterEmpty(f)).length > 0 && (
                            <div className="tabular-view-meta-row">
                              <span className="tabular-view-label">Filters:</span>
                              <div className="tabular-view-filter-pills">
                                {Object.entries(tabularViewDraft.filters || {}).filter(([, f]) => !isFilterEmpty(f)).map(([fKey, filter]) => {
                                  const col = allTabularColumns.find((c) => c.key === fKey);
                                  if (!col) return null;
                                  return (
                                    <span key={fKey} className="tabular-filter-pill">
                                      {col.label}: {formatFilterSummary(filter)}
                                      <button type="button" className="tabular-filter-pill-x"
                                        onClick={() => setTabularViewDraft((prev) => { const f = { ...(prev.filters || {}) }; delete f[fKey]; return { ...prev, filters: f }; })}
                                      >×</button>
                                    </span>
                                  );
                                })}
                                <button type="button" className="tabular-meta-clear" onClick={() => setTabularViewDraft((prev) => ({ ...prev, filters: {} }))}>Clear all</button>
                              </div>
                            </div>
                          )}
                        </div>
                        <div className="tabular-view-columns">
                          <div className="tabular-view-section-title" style={{ marginTop: 10, marginLeft: 10 }}>In View</div>
                          {tabularViewDraft.columnOrder.map((key, idx) => {
                            const col = allTabularColumns.find((c) => c.key === key);
                            if (!col) return null;
                            return (
                              <div
                                key={key}
                                className={`tabular-view-column-row${tabularDragOverKey === key && tabularDragKey !== key ? ' tabular-drag-over' : ''}`}
                                draggable
                                onDragStart={() => setTabularDragKey(key)}
                                onDragOver={(e) => { e.preventDefault(); setTabularDragOverKey(key); }}
                                onDragLeave={() => setTabularDragOverKey(null)}
                                onDrop={() => { handleTabularDragDrop(tabularDragKey, key); setTabularDragKey(null); setTabularDragOverKey(null); }}
                                onDragEnd={() => { setTabularDragKey(null); setTabularDragOverKey(null); }}
                              >
                                <span className="tabular-drag-handle" title="Drag to reorder">⠿</span>
                                <label className="tabular-view-column-toggle">
                                  <input type="checkbox" checked onChange={() => toggleTabularDraftSelected(key)} />
                                  <span>{col.label}</span>
                                </label>
                                {(() => {
                                  const { filterType } = getColumnFilterConfig(col);
                                  if (!filterType) return null;
                                  const hasFilter = !isFilterEmpty(tabularViewDraft.filters?.[key]);
                                  return (
                                    <button
                                      type="button"
                                      className={`tabular-filter-btn${hasFilter ? " tabular-filter-btn--active" : ""}`}
                                      title={hasFilter ? "Filter active" : "Filter"}
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        const rect = e.currentTarget.getBoundingClientRect();
                                        setTabularDraftFilterPopoverPos({ top: rect.bottom + 4, left: rect.left });
                                        setTabularDraftFilterKey((prev) => (prev === key ? null : key));
                                      }}
                                    >
                                      <Filter size={11} />
                                    </button>
                                  );
                                })()}
                                <input
                                  type="number"
                                  className="tabular-width-input"
                                  placeholder="auto"
                                  min="60"
                                  step="1"
                                  value={tabularViewDraft.columnWidths?.[key] || ""}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    setTabularViewDraft((prev) => {
                                      if (!v) { const cw = { ...(prev.columnWidths || {}) }; delete cw[key]; return { ...prev, columnWidths: cw }; }
                                      return { ...prev, columnWidths: { ...(prev.columnWidths || {}), [key]: Number(v) } };
                                    });
                                  }}
                                  title="Column width in pixels (blank = auto-size from data)"
                                />
                                <span className="tabular-width-px-label">px</span>
                              </div>
                            );
                          })}
                          <div className="tabular-view-section-title" style={{ marginBottom: 4, marginTop: 10, marginLeft: 10 }}>Available Fields</div>
                          {availableTabularColumns.map((col) => (
                            <div key={col.key} className="tabular-view-column-row">
                              <label className="tabular-view-column-toggle">
                                <input type="checkbox" checked={false} onChange={() => toggleTabularDraftSelected(col.key)} />
                                <span>{col.label}</span>
                              </label>
                            </div>
                          ))}
                        </div>
                      </div>
                      {tabularDraftFilterKey && (() => {
                        const draftFilterCol = allTabularColumns.find((c) => c.key === tabularDraftFilterKey);
                        if (!draftFilterCol) return null;
                        const { filterType, enumOptions } = getColumnFilterConfig(draftFilterCol);
                        return (
                          <ColumnFilterPopover
                            filterType={filterType}
                            enumOptions={enumOptions || []}
                            currentFilter={tabularViewDraft.filters?.[tabularDraftFilterKey] || null}
                            popoverPos={tabularDraftFilterPopoverPos}
                            onChange={(f) => setTabularViewDraft((prev) => ({ ...prev, filters: { ...(prev.filters || {}), [tabularDraftFilterKey]: f } }))}
                            onClose={() => setTabularDraftFilterKey(null)}
                          />
                        );
                      })()}
                      <DialogFooter style={{ justifyContent: "space-between" }}>
                        <div style={{ display: "flex", gap: 8 }}>
                          <Button type="button" variant="outline" onClick={updateCurrentTabularView} disabled={selectedTabularViewId === DEFAULT_TABULAR_VIEW_ID}>Save</Button>
                          {selectedTabularViewId !== DEFAULT_TABULAR_VIEW_ID && (
                            <Button type="button" variant="outline" onClick={deleteCurrentTabularView}>Delete This View</Button>
                          )}
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          <Button type="button" variant="outline" onClick={saveTabularViewAsNew}>Save As New…</Button>
                          <Button type="button" variant="secondary" onClick={() => setTabularViewDialogOpen(false)}>Close</Button>
                        </div>
                      </DialogFooter>
                    </DialogContent>
                  </Dialog>
                )}

                {tabularSaveAsNewOpen && (
                  <Dialog open={tabularSaveAsNewOpen} onOpenChange={(open) => { if (!open) setTabularSaveAsNewOpen(false); }}>
                    <DialogContent style={{ width: "min(400px, 92vw)", maxWidth: "none" }}>
                      <DialogHeader>
                        <DialogTitle>Save as New View — Nodes</DialogTitle>
                      </DialogHeader>
                      <div style={{ padding: "8px 0" }}>
                        <label className="form-label" style={{ marginBottom: 6, display: "block" }}>View name</label>
                        <input
                          className="form-input"
                          value={tabularSaveAsNewName}
                          onChange={(e) => { setTabularSaveAsNewError(""); setTabularSaveAsNewName(e.target.value); }}
                          onKeyDown={(e) => { if (e.key === "Enter") commitTabularSaveAsNew(); }}
                          placeholder="e.g. My Tax View"
                          autoFocus
                          autoComplete="off"
                          data-lpignore="true"
                        />
                        {tabularSaveAsNewError && (
                          <div className="dup-warning" style={{ marginTop: 8 }}>{tabularSaveAsNewError}</div>
                        )}
                      </div>
                      <DialogFooter>
                        <Button type="button" variant="secondary" onClick={() => setTabularSaveAsNewOpen(false)}>Cancel</Button>
                        <Button type="button" onClick={commitTabularSaveAsNew}>Save</Button>
                      </DialogFooter>
                    </DialogContent>
                  </Dialog>
                )}

                {tabularDeleteConfirmOpen && (
                  <Dialog open={tabularDeleteConfirmOpen} onOpenChange={(open) => { if (!open) setTabularDeleteConfirmOpen(false); }}>
                    <DialogContent style={{ width: "min(400px, 92vw)", maxWidth: "none" }}>
                      <DialogHeader>
                        <DialogTitle>Delete View</DialogTitle>
                      </DialogHeader>
                      <p style={{ padding: "4px 0 16px", color: "#374151", margin: 0 }}>
                        Permanently delete <strong>&#8220;{activeTabularView.name}&#8221;</strong>? This cannot be undone.
                      </p>
                      <DialogFooter>
                        <Button type="button" variant="secondary" onClick={() => setTabularDeleteConfirmOpen(false)}>Cancel</Button>
                        <Button type="button" onClick={confirmTabularViewDelete} style={{ background: '#dc2626', borderColor: '#dc2626', color: '#fff' }}>Delete</Button>
                      </DialogFooter>
                    </DialogContent>
                  </Dialog>
                )}
              </>
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

        {viewMode === "directory" && (
          <div className="fab-container">
            {/* FAB buttons now integrated into section headers */}
          </div>
        )}

        <Dialog open={Boolean(openDialog)} onOpenChange={() => { setOpenDialog(null); setPrevDialog(null); setDupMatches([]); setOwnerSearch(""); setOwnerSearchOpen(false); }}>

          {openDialog?.type === "data-dictionary" && (
            <DialogContent className="dialog-content--tall" style={{ width: "min(900px, 92vw)", maxWidth: "none" }}>
              <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
                <DialogTitle>Data Dictionary — {clientDisplayName || toSentenceCase(clientId)}</DialogTitle>
              </DialogHeader>
              <div className="dialog-body">
                <>
                  {dataDictionary.length === 0 && (
                    <div style={{ color: "#6b7280", fontSize: 14, padding: "12px 0" }}>
                      No custom fields defined yet. Use <strong>Add Field</strong> to create one.
                    </div>
                  )}
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
                        {/* Name — always first since it's the record identifier */}
                        <tr>
                          <td><em style={{ color: "#6b7280" }}>Name</em></td>
                          <td style={{ color: "#6b7280" }}>Short text</td>
                          <td style={{ color: "#6b7280" }}>All</td>
                          <td style={{ color: "#6b7280" }}>No</td>
                          <td style={{ color: "#6b7280" }}></td>
                          <td></td>
                          <td></td>
                        </tr>
                        {/* Built-in fields + DD fields merged and sorted alphabetically */}
                        {(() => {
                          const builtIns = [
                            { _builtin: true, prompt: "Address", type: "Address", appliesTo: "Entity, Person", multi: "No", values: "" },
                            { _builtin: true, prompt: "Cell Phone", type: "Phone", appliesTo: "Person", multi: "No", values: "" },
                            { _builtin: true, prompt: "e-Mail", type: "Email", appliesTo: "Entity, Person", multi: "Yes", values: "" },
                            { _builtin: true, prompt: "Ownership Records", type: "Computed", appliesTo: "Entity", multi: "No", values: "Read-only" },
                            { _builtin: true, prompt: "Primary Phone", type: "Phone", appliesTo: "Entity, Person", multi: "No", values: "" },
                            { _builtin: true, prompt: "Tax ID", type: "Short text", appliesTo: "Entity", multi: "No", values: "" },
                          ];
                          const ddRows = [...dataDictionary].map((e) => ({ _builtin: false, _entry: e, prompt: e.prompt || "" }));
                          const all = [...builtIns, ...ddRows].sort((a, b) =>
                            String(a.prompt || "").localeCompare(String(b.prompt || ""), undefined, { sensitivity: "base" })
                          );
                          return all.map((row) => {
                            if (row._builtin) {
                              return (
                                <tr key={`builtin-${row.prompt}`}>
                                  <td><em style={{ color: "#6b7280" }}>{row.prompt}</em></td>
                                  <td style={{ color: "#6b7280" }}>{row.type}</td>
                                  <td style={{ color: "#6b7280" }}>{row.appliesTo}</td>
                                  <td style={{ color: "#6b7280" }}>{row.multi}</td>
                                  <td style={{ color: "#6b7280" }}>{row.values}</td>
                                  <td></td>
                                  <td></td>
                                </tr>
                              );
                            }
                            const entry = row._entry;
                            return (
                              <tr key={entry.id}>
                                <td>{entry.prompt}</td>
                                <td>{DATA_TYPES.find((t) => t.value === entry.dataType)?.label ?? entry.dataType}</td>
                                <td>
                                  {(() => { const n = normalizeAppliesTo(entry.appliesTo); return [n.includes("entity") && "Entity", n.includes("person") && "Person", n.includes("ownership") && "Ownership"].filter(Boolean).join(", ") || "—"; })()}
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
                                    <Button type="button" variant="outline" style={{ padding: "4px 8px" }} title="Edit" onClick={() => openDdEntry(entry)}><Pencil size={13} /></Button>
                                    <Button type="button" variant="outline" style={{ padding: "4px 8px" }} title="Delete" onClick={() => deleteDdEntry(entry.id)}><Trash2 size={13} /></Button>
                                  </div>
                                </td>
                              </tr>
                            );
                          });
                        })()}
                      </tbody>
                    </table>
                  </div>
                </>
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
                  <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
                    {["entity", "person", "ownership"].map((kind) => {
                      const checked = normalizeAppliesTo(ddEntryDraft.appliesTo).includes(kind);
                      return (
                        <label key={kind} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 14 }}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => {
                              const cur = normalizeAppliesTo(ddEntryDraft.appliesTo);
                              const next = checked ? cur.filter((k) => k !== kind) : [...cur, kind];
                              setDdEntryDraft((prev) => ({ ...prev, appliesTo: next.length ? next : cur }));
                            }}
                          />
                          {kind.charAt(0).toUpperCase() + kind.slice(1)}
                        </label>
                      );
                    })}
                  </div>
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
            const hasOutOfRangePercent = ownerEditorRows.some((r) => {
              if (r.percent === "") return false;
              return !parseOwnershipPercent(r.percent).ok;
            });
            const ownerTotalInvalid = Math.abs(ownerTotal - 100) > 0.0001;
            const searchQ = ownerSearch.trim().toLowerCase();
            const searchResults = searchQ
              ? nodeList.filter(
                (n) =>
                  n.name?.toLowerCase().includes(searchQ) &&
                  !ownerEditorRows.find((r) => r.nodeId === n.id) &&
                  n.id !== openDialog.targetId
              )
              : [];

            // Helper to format date for display
            const formatDateDisplay = (dateStr) => {
              if (!dateStr) return null;
              try {
                // Parse as local date, not UTC (ISO date strings are dates, not datetimes)
                const [year, month, day] = dateStr.split('-').map(Number);
                const d = new Date(year, month - 1, day);
                return d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
              } catch {
                return dateStr;
              }
            };

            const otherPeriods = ownershipTimeline.filter((p) => p.setId !== ownershipSelectedPeriodSetId);

            return (
              <DialogContent style={{ minWidth: 560, maxWidth: 700, maxHeight: '85vh', display: 'flex', flexDirection: 'column' }}>
                <DialogHeader>
                  <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12 }}>
                    <div>
                      <DialogTitle>Owners of {targetNode?.name ?? openDialog.targetId}</DialogTitle>
                      {ownerEditorRows.length > 0 && (
                        <div style={{ fontSize: '12px', color: '#64748b', marginTop: 4 }}>
                          {formatOwnershipDateRange(ownerEditorDateRange.from, ownerEditorDateRange.to)}
                        </div>
                      )}
                    </div>
                    {!ownerEditorDateRange.isCurrent && (
                      <div style={{ backgroundColor: '#fee2e2', color: '#991b1b', padding: '4px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, whiteSpace: 'nowrap', marginTop: 2 }}>
                        HISTORICAL
                      </div>
                    )}
                  </div>
                </DialogHeader>

                {/* VIEW MODE */}
                {ownerEditorMode === "view" && (
                  <div style={{ overflowY: 'auto', flex: 1, paddingRight: 16 }}>
                    {/* Current ownership display */}
                    <div style={{ marginBottom: 20 }}>
                      {ownerEditorRows.length === 0 ? (
                        <div style={{ padding: 12, backgroundColor: '#f0fdf4', color: '#166534', borderRadius: 4, fontSize: 13 }}>
                          {(() => {
                            // Find the earliest effective date in timeline if there are future periods
                            const earliestFutureDate = ownershipTimeline
                              .map((p) => p.effectiveFrom)
                              .filter(Boolean)
                              .sort((a, b) => a.localeCompare(b))[0];
                            
                            return earliestFutureDate
                              ? `No ownership records before ${formatDateDisplay(earliestFutureDate)}`
                              : "No ownership records";
                          })()}
                        </div>
                      ) : (
                        <div className="owner-editor">
                          {[...ownerEditorRows].sort((a, b) => (Number(b.percent) || 0) - (Number(a.percent) || 0)).map((row) => (
                            <div key={row.nodeId} className="owner-editor-row" style={{ opacity: 0.8 }}>
                              <div className="owner-editor-name">{row.name}</div>
                              <div className="owner-editor-pct-wrap">
                                <span style={{ fontWeight: 600 }}>{row.percent}%</span>
                              </div>
                            </div>
                          ))}
                          <div style={{ marginTop: 12, paddingTop: 12, borderTop: '1px solid #e2e8f0', fontWeight: 600, fontSize: 13 }}>
                            Total: {ownerTotal}%
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Timeline of other ownership periods */}
                    {ownershipTimeline.length > 1 && (
                      <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 12 }}>
                        <div style={{ fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 8 }}>
                          Other ownership groups {otherPeriods.length > 0 ? `(${otherPeriods.length})` : "(none)"}
                        </div>
                        {otherPeriods.length === 0 ? (
                          <div style={{ fontSize: 12, color: '#94a3b8 ' }}>No other ownership groups</div>
                        ) : (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                            {otherPeriods.map((period) => {
                              const fromDisplay = formatDateDisplay(period.effectiveFrom);
                              const toDisplay = formatDateDisplay(period.effectiveTo);
                              const label = !fromDisplay && !toDisplay ? "Unspecified"
                                : fromDisplay && !toDisplay ? `Since ${fromDisplay}`
                                  : !fromDisplay && toDisplay ? `Until ${toDisplay}`
                                    : `${fromDisplay} to ${toDisplay}`;
                              return (
                                <button
                                  key={period.setId}
                                  type="button"
                                  onClick={() => {
                                    // Reload this period's data and stay in view mode
                                    const periodRows = period.owners.map((o) => {
                                      const node = getNode(nodeList, o.from);
                                      return {
                                        nodeId: o.from,
                                        name: node?.name ?? o.from,
                                        percent: String(o.percent),
                                        startDate: "",
                                        endDate: "",
                                        isNew: false,
                                      };
                                    });
                                    setOwnershipSelectedPeriodSetId(period.setId);
                                    setOwnerEditorRows(periodRows);
                                    setOwnerEditorOriginal(periodRows); // Sync backup for cancel logic
                                    setOwnerEditorDateRange({
                                      from: period.effectiveFrom,
                                      to: period.effectiveTo,
                                      isCurrent: !period.effectiveTo,
                                    });
                                  }}
                                  style={{
                                    padding: '8px 10px',
                                    textAlign: 'left',
                                    fontSize: 12,
                                    border: '1px solid #cbd5e1',
                                    borderRadius: 4,
                                    backgroundColor: '#f8fafc',
                                    cursor: 'pointer',
                                    transition: 'all 0.2s',
                                  }}
                                  onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f1f5f9'}
                                  onMouseOut={(e) => e.currentTarget.style.backgroundColor = '#f8fafc'}
                                >
                                  <div style={{ fontWeight: 500 }}>{label}</div>
                                  <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>
                                    {period.owners.length} owner{period.owners.length !== 1 ? 's' : ''} • {period.owners.reduce((s, o) => s + (o.percent || 0), 0)}%
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}

                {/* EDIT/CREATE MODE */}
                {(ownerEditorMode === "edit-existing" || ownerEditorMode === "create-new") && (
                  <div style={{ overflowY: 'auto', flex: 1, paddingRight: 16 }}>
                    <div className="owner-editor">
                      <div className="form-row" style={{ marginBottom: 12 }}>
                        <label className="form-label">
                          {ownerEditorMode === "edit-existing" ? "Effective from (change start date)" : "Effective from"}
                        </label>
                        <input
                          className="form-input"
                          type="date"
                          value={ownerEditorEffectiveDate}
                          onChange={(e) => setOwnerEditorEffectiveDate(e.target.value)}
                          required
                        />
                      </div>

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
                        {hasOutOfRangePercent && <span className="owner-editor-over-msg"> — each owner percent must be 0-100</span>}
                        {!overLimit && ownerTotalInvalid && <span className="owner-editor-over-msg"> — must equal 100% to save</span>}
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
                                {n.kind === "person" ? <Users size={14} style={{ color: "#6b7280", flexShrink: 0 }} /> : <Building2 size={14} style={{ color: "#6b7280", flexShrink: 0 }} />}
                                {n.name}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                )}

                <DialogFooter>
                  {ownerEditorMode === "view" && (
                    <>
                      <Button variant="secondary" onClick={() => {
                        if (prevDialog) { setOpenDialog(prevDialog); setPrevDialog(null); }
                        else setOpenDialog(null);
                      }}>Cancel</Button>
                      {!ownershipDeleteConfirm && (
                        <div style={{ display: 'flex', gap: 8 }}>
                          <Button type="button" variant="outline" onClick={() => setOwnershipDeleteConfirm(true)} title="Permanently delete this ownership group" style={{ color: '#dc2626', borderColor: '#dc2626' }}>
                            Delete this group
                          </Button>
                          <Button type="button" variant="outline" onClick={() => {
                            setOwnerEditorMode("edit-existing");
                            setOwnerEditorEffectiveDate(ownerEditorDateRange.from || "");
                          }}>
                            Update this period
                          </Button>
                          <Button type="button" variant="default" onClick={() => {
                            // Save the original state for restoration on cancel
                            setOwnerEditorOriginal(ownerEditorRows);
                            setOwnerEditorMode("create-new");
                            setOwnerEditorRows([]);
                            setOwnerEditorEffectiveDate(todayIso);
                          }}>
                            Create new ownership group
                          </Button>
                        </div>
                      )}
                      {ownershipDeleteConfirm && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 12, width: '100%' }}>
                          {(() => {
                            // Check if this is the earliest (chronologically first) ownership period
                            const allPeriods = ownershipTimeline || [];
                            const sortedByDate = [...allPeriods].sort((a, b) =>
                              (a.effectiveFrom || "0001-01-01").localeCompare(b.effectiveFrom || "0001-01-01")
                            );
                            const isEarliest = sortedByDate.length > 0 && sortedByDate[0].setId === ownershipSelectedPeriodSetId;

                            if (isEarliest && sortedByDate.length > 1) {
                              const nextPeriodDate = sortedByDate[1].effectiveFrom;
                              const formatted = nextPeriodDate
                                ? new Date(nextPeriodDate + 'T00:00:00Z').toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' })
                                : 'unknown date';
                              return (
                                <div style={{ fontSize: 13, color: '#92400e', backgroundColor: '#fef3c7', padding: 10, borderRadius: 4, lineHeight: 1.5 }}>
                                  <strong>⚠️ Caution:</strong> This is the earliest ownership period. Deleting it will create a gap with no ownership information before {formatted}. You can add groups later to fill this gap if needed.
                                </div>
                              );
                            }
                            return null;
                          })()}
                          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                            <span style={{ fontSize: 13, color: '#991b1b', fontWeight: 500 }}>Delete this group? This cannot be undone.</span>
                            <Button type="button" variant="outline" onClick={() => setOwnershipDeleteConfirm(false)} disabled={isSavingOwners}>
                              Cancel
                            </Button>
                            <Button type="button" variant="outline" onClick={deleteOwnershipGroup} disabled={isSavingOwners} style={{ color: '#dc2626', borderColor: '#dc2626' }}>
                              {isSavingOwners ? "Deleting…" : "Confirm Delete"}
                            </Button>
                          </div>
                        </div>
                      )}
                    </>
                  )}
                  {(ownerEditorMode === "edit-existing" || ownerEditorMode === "create-new") && (
                    <>
                      <Button variant="secondary" onClick={() => {
                        // Restore original state when canceling
                        setOwnerEditorMode("view");
                        setOwnerEditorRows(ownerEditorOriginal);
                        setOwnerEditorEffectiveDate("");
                      }}>Cancel</Button>
                      <Button type="button" disabled={isSavingOwners || ownerTotalInvalid || hasOutOfRangePercent || !normalizeDateInput(ownerEditorEffectiveDate)} onClick={saveOwnerEditor}>
                        {isSavingOwners ? "Saving…" : "Save"}
                      </Button>
                    </>
                  )}
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
                  {/* Built-in contact fields */}
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={newNode.address || ""}
                      onChange={(e) => setNewNode((prev) => ({ ...prev, address: e.target.value }))}
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Primary Phone</label>
                    <PhoneInputRow
                      value={newNode.workPhone || ""}
                      onCommit={(val) => setNewNode((prev) => ({ ...prev, workPhone: val }))}
                    />
                  </div>
                  {newNode.kind === "person" && (
                    <div className="form-row">
                      <label className="form-label">Cell Phone</label>
                      <PhoneInputRow
                        value={newNode.cellPhone || ""}
                        onCommit={(val) => setNewNode((prev) => ({ ...prev, cellPhone: val }))}
                      />
                    </div>
                  )}
                  {newNode.kind === "person" && (
                    <div className="form-row">
                      <label className="form-label">e-Mail</label>
                      <input
                        className="form-input"
                        type="email"
                        value={Array.isArray(newNode.emails) ? newNode.emails.join(", ") : (newNode.emails || "")}
                        onChange={(e) => setNewNode((prev) => ({ ...prev, emails: e.target.value }))}
                        autoComplete="off"
                        data-lpignore="true"
                      />
                    </div>
                  )}
                  {newNode.kind === "entity" && (
                    <div className="form-row">
                      <label className="form-label">Tax ID</label>
                      <input
                        className="form-input"
                        type="text"
                        value={newNode.taxId || ""}
                        onChange={(e) => setNewNode((prev) => ({ ...prev, taxId: e.target.value }))}
                        autoComplete="off"
                        data-lpignore="true"
                      />
                    </div>
                  )}
                  {[...dataDictionary, ENTITY_OWNERSHIP_SUMMARY_FIELD]
                    .filter((f) => normalizeAppliesTo(f.appliesTo).includes(newNode.kind))
                    .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                    .map((field) => (
                      <React.Fragment key={field.fieldId}>
                        {renderDdField(
                          field,
                          newNode.customFields?.[field.fieldId],
                          (val) => setNewNode((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } })),
                          { apiBase, token, node: newNode, nodeList, relList, asOfDate, ownershipTimeline: newNode.kind === "entity" ? editNodeOwnershipTimeline : [] }
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
                        asOfDate: asOfDate || todayIso,
                        client: clientId,
                        photo: newNode.kind === "person" ? (newNode.photo || "") : "",
                        logo: newNode.kind === "entity" ? (newNode.logo || "") : "",
                        address: newNode.address || "",
                        workPhone: newNode.workPhone || "",
                        cellPhone: newNode.cellPhone || "",
                        emails: newNode.emails || "",
                        taxId: newNode.taxId || "",
                        customFields: newNode.customFields || {},
                      };
                      try {
                        await apiRequest("/api/nodes", {
                          method: "POST",
                          body: JSON.stringify(payload),
                        });
                        setNodeList((prev) => [...prev, payload]);
                        if (!focusId) setFocusId(id);
                        setNewNode({ name: "", kind: newNode.kind, photo: "", logo: "", address: "", workPhone: "", cellPhone: "", emails: "", taxId: "", customFields: {} });
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
                  {newOwnershipPercentInvalid && (
                    <div className="dup-warning">Ownership percent must be between 0 and 100.</div>
                  )}
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
                  disabled={isAddingOwnership || newOwnershipPercentInvalid}
                  onClick={async () => {
                    if (isAddingOwnership) return;
                    if (!newOwnership.from || !newOwnership.to) return;
                    const parsedPercent = parseOwnershipPercent(newOwnership.percent);
                    if (!parsedPercent.ok) {
                      setRemoteStatus("error");
                      setRemoteError(parsedPercent.message);
                      return;
                    }
                    setIsAddingOwnership(true);
                    const payload = {
                      from: newOwnership.from,
                      to: newOwnership.to,
                      percent: parsedPercent.value,
                      startDate: newOwnership.startDate || null,
                      endDate: newOwnership.endDate || null,
                      asOfDate: asOfDate || todayIso,
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
                  {/* Effective date of this change */}
                  <div className="form-row">
                    <label className="form-label" title="Leave blank to use the global As-of date, or today if none is set">
                      Change effective
                    </label>
                    <input
                      className="form-input"
                      type="date"
                      value={editNodeEffectiveDate}
                      placeholder={asOfDate || todayIso}
                      onChange={(e) => setEditNodeEffectiveDate(e.target.value)}
                    />
                    <span style={{ fontSize: 11, color: "#6b7280", marginLeft: 4 }}>
                      {editNodeEffectiveDate
                        ? `Recording change as of ${editNodeEffectiveDate}`
                        : asOfDate
                          ? `Using global As-of date (${asOfDate})`
                          : `Defaulting to today (${todayIso})`}
                    </span>
                  </div>
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
                  {/* Built-in contact fields */}
                  <div className="form-row">
                    <label className="form-label">Address</label>
                    <input
                      className="form-input"
                      type="text"
                      value={nodeDraft.address || ""}
                      onChange={(e) => setNodeDraft((prev) => ({ ...prev, address: e.target.value }))}
                      autoComplete="off"
                      data-lpignore="true"
                    />
                  </div>
                  <div className="form-row">
                    <label className="form-label">Primary Phone</label>
                    <PhoneInputRow
                      value={nodeDraft.workPhone || ""}
                      onCommit={(val) => setNodeDraft((prev) => ({ ...prev, workPhone: val }))}
                    />
                  </div>
                  {nodeDraft.kind === "person" && (
                    <div className="form-row">
                      <label className="form-label">Cell Phone</label>
                      <PhoneInputRow
                        value={nodeDraft.cellPhone || ""}
                        onCommit={(val) => setNodeDraft((prev) => ({ ...prev, cellPhone: val }))}
                      />
                    </div>
                  )}
                  {nodeDraft.kind === "person" && (
                    <div className="form-row">
                      <label className="form-label">e-Mail</label>
                      <input
                        className="form-input"
                        type="email"
                        value={Array.isArray(nodeDraft.emails) ? nodeDraft.emails.join(", ") : (nodeDraft.emails || "")}
                        onChange={(e) => setNodeDraft((prev) => ({ ...prev, emails: e.target.value }))}
                        autoComplete="off"
                        data-lpignore="true"
                      />
                    </div>
                  )}
                  {nodeDraft.kind === "entity" && (
                    <div className="form-row">
                      <label className="form-label">Tax ID</label>
                      <input
                        className="form-input"
                        type="text"
                        value={nodeDraft.taxId || ""}
                        onChange={(e) => setNodeDraft((prev) => ({ ...prev, taxId: e.target.value }))}
                        autoComplete="off"
                        data-lpignore="true"
                      />
                    </div>
                  )}
                  {[...dataDictionary, ENTITY_OWNERSHIP_SUMMARY_FIELD]
                    .filter((f) => normalizeAppliesTo(f.appliesTo).includes(nodeDraft.kind))
                    .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0))
                    .map((field) => (
                      <React.Fragment key={field.fieldId}>
                        {renderDdField(
                          field,
                          nodeDraft.customFields?.[field.fieldId],
                          (val) => setNodeDraft((prev) => ({ ...prev, customFields: { ...prev.customFields, [field.fieldId]: val } })),
                          { apiBase, token, node: nodeDraft, nodeList, relList, asOfDate, ownershipTimeline: nodeDraft.kind === "entity" ? editNodeOwnershipTimeline : [] }
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
                    onClick={() => {
                      if (!editNodeId) return;
                      openNodeBookPrintDialog(editNodeId);
                      setOpenDialog(null);
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
                            body: JSON.stringify({ asOfDate: asOfDate || todayIso, client: clientId }),
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
                        asOfDate: editNodeEffectiveDate || asOfDate || todayIso,
                        client: clientId,
                        newId: newId !== editNodeId ? newId : null,
                        photo: nodeDraft.kind === "person" ? (nodeDraft.photo || "") : "",
                        logo: nodeDraft.kind === "entity" ? (nodeDraft.logo || "") : "",
                        address: nodeDraft.address || "",
                        workPhone: nodeDraft.workPhone || "",
                        cellPhone: nodeDraft.cellPhone || "",
                        emails: nodeDraft.emails || "",
                        taxId: nodeDraft.taxId || "",
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
                  {editOwnershipPercentInvalid && (
                    <div className="dup-warning">Ownership percent must be between 0 and 100.</div>
                  )}
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
                            asOfDate: asOfDate || todayIso,
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
                  disabled={editOwnershipPercentInvalid}
                  onClick={() => {
                    if (!editOwnershipId) return;
                    const parsedPercent = parseOwnershipPercent(ownershipDraft.percent);
                    if (!parsedPercent.ok) {
                      setRemoteStatus("error");
                      setRemoteError(parsedPercent.message);
                      return;
                    }
                    const payload = {
                      from: ownershipDraft.from,
                      to: ownershipDraft.to,
                      percent: parsedPercent.value,
                      startDate: ownershipDraft.startDate || null,
                      endDate: ownershipDraft.endDate || null,
                      asOfDate: asOfDate || todayIso,
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

      {quickViewNode && (
        <aside className="quick-view-panel">
          <div className="quick-view-header">
            <div>
              <div className="quick-view-title">Quick View</div>
              <div className="quick-view-subtitle">{quickViewNode.name || quickViewNode.id}</div>
            </div>
            <button
              type="button"
              className="quick-view-close"
              onClick={() => setQuickViewNodeId("")}
              aria-label="Close quick view"
              title="Close"
            >
              <X size={14} />
            </button>
          </div>
          <div className="quick-view-section">
            <div className="quick-view-row">
              <span style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                {quickViewNode.kind === "entity" ? <Building2 size={16} /> : <Users size={16} />}
                Type
              </span>
              <strong>{toSentenceCase(quickViewNode.kind || "")}</strong>
            </div>
            {quickViewNode.address && <div className="quick-view-row"><span>Address</span><strong>{quickViewNode.address}</strong></div>}
            {quickViewEmailText && <div className="quick-view-row"><span>Email</span><strong>{quickViewEmailText}</strong></div>}
            {quickViewPhoneText && <div className="quick-view-row"><span>Phone</span><strong>{quickViewPhoneText}</strong></div>}
            {quickViewOwners.length > 0 && <div className="quick-view-row"><span>Owners</span><strong>{quickViewOwners.length}</strong></div>}
            <div className="quick-view-row"><span>Entities owned</span><strong>{quickViewOwned.length} ({quickViewDescCount})</strong></div>
            {quickViewNode.kind === "entity" && (
              <div className="quick-view-block">
                <div className="quick-view-label">Ownership Summary - {quickViewOwners.length} owner{quickViewOwners.length === 1 ? "" : "s"}</div>
                <div className="quick-view-summary">{quickViewOwnershipSummary || "No ownership records"}</div>
              </div>
            )}
          </div>
          <div className="quick-view-actions">
            <Button
              type="button"
              variant="outline"
              onClick={() => {
                setEditNodeId(quickViewNode.id);
                setOpenDialog({ type: "edit-node" });
                setQuickViewNodeId("");
              }}
            >
              <Pencil size={14} /> View/Edit
            </Button>
          </div>
        </aside>
      )}
    </div>
  );
}

// EOF
