/**
 * ExportDialog — two-step Excel export with saved report templates.
 *
 * Step 1 ("pick"):   choose a previously-saved field selection, or start fresh.
 *                    Skipped automatically when no saved reports exist.
 * Step 2 ("configure"): check/uncheck DD fields, name the report, export.
 *
 * Props:
 *   open            – boolean
 *   onClose         – () => void
 *   exportNodes     – array of node objects to include in the export
 *                     (caller decides filtering: directory filter or all nodes)
 *   dataDictionary  – DD field definitions array
 *   savedReports    – array of { id, reportId, name, fieldIds } loaded by the caller
 *   onReportSaved   – (report) => void  — called after a successful PUT
 *   onReportDeleted – (reportId) => void
 *   apiRequest      – the bound apiRequest(path, options) from EntityApp
 *   clientName      – display name for the sheet / file name
 */

import React, { useCallback, useEffect, useRef, useState } from "react";
import { Button } from "./ui/button";
import { formatPhone } from "../utils/helpers";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "./ui/dialog";
import { FileSpreadsheet, Plus, Trash2, ChevronLeft } from "lucide-react";
import * as XLSX from "xlsx";

// ─── helpers ─────────────────────────────────────────────────────────────────

const slugify = (value) =>
  String(value || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "report";

const formatFieldValue = (field, value) => {
  if (value == null || value === "" || value === false) return "";
  const { dataType } = field;
  if (dataType === "boolean") return value ? "Yes" : "No";
  if (dataType === "date") {
    try {
      return new Date(String(value) + "T00:00:00").toLocaleDateString(undefined, {
        month: "short", day: "numeric", year: "numeric",
      });
    } catch { return String(value); }
  }
  if (dataType === "currency") {
    const num = Number(value);
    return Number.isFinite(num)
      ? num.toLocaleString(undefined, { style: "currency", currency: "USD" })
      : String(value);
  }
  if (dataType === "percentage") return `${value}%`;
  if (dataType === "phone") {
    if (Array.isArray(value)) {
      return value.map((e) => `${e.type}: ${formatPhone(e.number)}`).filter((s) => s.trim() !== ": ").join(", ");
    }
    if (value && typeof value === "object") {
      return Object.entries(value).map(([t, n]) => `${t}: ${formatPhone(n)}`).join(", ");
    }
    return formatPhone(String(value));
  }
  if (Array.isArray(value)) return value.join(", ");
  return String(value);
};

const KIND_BADGE = {
  entity: { label: "Entity", color: "#2563eb" },
  person: { label: "Person", color: "#7c3aed" },
  both:   { label: "Both",   color: "#059669" },
};

// ─── Virtual fields (not in DD, computed from relationships) ──────────────────

const VIRTUAL_FIELDS = [
  {
    fieldId: "__virtual__owners",
    prompt: "Entity Owners",
    dataType: "virtual",
    appliesTo: "entity",
    sortOrder: -2,
    _virtual: true,
  },
  {
    fieldId: "__virtual__owned",
    prompt: "Owned Entities",
    dataType: "virtual",
    appliesTo: "both",
    sortOrder: -1,
    _virtual: true,
  },
];

const getVirtualValue = (fieldId, node, nodeList, relList) => {
  if (fieldId === "__virtual__owners") {
    // Only meaningful for entities
    if (node.kind === "person") return "";
    const owners = (relList || [])
      .filter((r) => r.type === "owns" && r.to === node.id)
      .sort((a, b) => (Number(b.percent) || 0) - (Number(a.percent) || 0));
    if (owners.length === 0) return "";
    return owners
      .map((r) => {
        const ownerNode = (nodeList || []).find((n) => n.id === r.from);
        const name = ownerNode?.name || r.from;
        const pct = r.percent != null && Number.isFinite(Number(r.percent))
          ? ` ${Number(r.percent)}%`
          : "";
        return `${name}${pct}`;
      })
      .join("; ");
  }
  if (fieldId === "__virtual__owned") {
    const owned = (relList || [])
      .filter((r) => r.type === "owns" && r.from === node.id)
      .sort((a, b) => (Number(b.percent) || 0) - (Number(a.percent) || 0));
    if (owned.length === 0) return "";
    return owned
      .map((r) => {
        const ownedNode = (nodeList || []).find((n) => n.id === r.to);
        const name = ownedNode?.name || r.to;
        const pct = r.percent != null && Number.isFinite(Number(r.percent))
          ? ` ${Number(r.percent)}%`
          : "";
        return `${name}${pct}`;
      })
      .join("; ");
  }
  return "";
};

// ─── ExportDialog ─────────────────────────────────────────────────────────────

export default function ExportDialog({
  open,
  onClose,
  exportNodes,
  nodeList,
  relList,
  dataDictionary,
  savedReports,
  onReportSaved,
  onReportDeleted,
  apiRequest,
  clientName,
}) {
  const [step, setStep] = useState("pick");
  const [reportName, setReportName] = useState("");
  const [checkedIds, setCheckedIds] = useState(new Set());
  const [isBusy, setIsBusy] = useState(false);
  const [error, setError] = useState("");
  const nameRef = useRef(null);

  const sortedFields = [
    ...VIRTUAL_FIELDS,
    ...[...(dataDictionary || [])].sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0)),
  ];

  // On open: decide starting step
  useEffect(() => {
    if (!open) return;
    setError("");
    setIsBusy(false);
    if (savedReports && savedReports.length > 0) {
      setStep("pick");
    } else {
      setStep("configure");
      setReportName("");
      setCheckedIds(new Set());
    }
  }, [open]); // eslint-disable-line react-hooks/exhaustive-deps

  // Focus the name field when we reach "configure"
  useEffect(() => {
    if (step === "configure" && open) {
      setTimeout(() => nameRef.current?.focus(), 50);
    }
  }, [step, open]);

  const pickReport = (report) => {
    setReportName(report.name);
    setCheckedIds(new Set(report.fieldIds || []));
    setStep("configure");
    setError("");
  };

  const skipPick = () => {
    setReportName("");
    setCheckedIds(new Set());
    setStep("configure");
    setError("");
  };

  const toggleField = (fieldId) => {
    setCheckedIds((prev) => {
      const next = new Set(prev);
      if (next.has(fieldId)) next.delete(fieldId);
      else next.add(fieldId);
      return next;
    });
  };

  const toggleAll = () => {
    if (checkedIds.size === sortedFields.length) {
      setCheckedIds(new Set());
    } else {
      setCheckedIds(new Set(sortedFields.map((f) => f.fieldId)));
    }
  };

  const handleDeleteReport = async (report, e) => {
    e.stopPropagation();
    setIsBusy(true);
    setError("");
    try {
      await apiRequest(`/api/export-reports/${encodeURIComponent(report.reportId)}`, {
        method: "DELETE",
      });
      onReportDeleted?.(report.reportId);
    } catch (err) {
      setError(err.message || "Could not delete report");
    } finally {
      setIsBusy(false);
    }
  };

  const handleExport = async () => {
    if (checkedIds.size === 0) {
      setError("Select at least one field to export.");
      return;
    }
    setIsBusy(true);
    setError("");

    const name = reportName.trim();

    try {
      // 1. Save/overwrite the report definition if a name was given
      if (name) {
        const reportId = `report:${slugify(name)}`;
        const saved = await apiRequest(`/api/export-reports/${encodeURIComponent(reportId)}`, {
          method: "PUT",
          body: JSON.stringify({ name, fieldIds: [...checkedIds] }),
        });
        onReportSaved?.(saved);
      }

      // 2. Build the XLSX
      const checkedFields = sortedFields.filter((f) => checkedIds.has(f.fieldId));
      const headers = ["Name", "Type", ...checkedFields.map((f) => f.prompt)];

      const rows = (exportNodes || []).map((node) => {
        const row = { Name: node.name, Type: node.kind };
        checkedFields.forEach((field) => {
          if (field._virtual) {
            row[field.prompt] = getVirtualValue(field.fieldId, node, nodeList, relList);
          } else {
            row[field.prompt] = formatFieldValue(field, node.customFields?.[field.fieldId]);
          }
        });
        return row;
      });

      const worksheet = XLSX.utils.json_to_sheet(rows, { header: headers });

      // Auto-width columns
      const colWidths = headers.map((h) => {
        const maxData = rows.reduce((max, row) => {
          const val = String(row[h] || "");
          return Math.max(max, val.length);
        }, h.length);
        return { wch: Math.min(maxData + 2, 60) };
      });
      worksheet["!cols"] = colWidths;

      const workbook = XLSX.utils.book_new();
      const sheetName = (clientName || "Export").slice(0, 31);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

      const fileName = `${name || clientName || "export"}.xlsx`
        .replace(/[^\w\s.-]/g, "")
        .replace(/\s+/g, "_");
      XLSX.writeFile(workbook, fileName);

      onClose();
    } catch (err) {
      setError(err.message || "Export failed");
    } finally {
      setIsBusy(false);
    }
  };

  // ── Pick step ──────────────────────────────────────────────────────────────

  const renderPick = () => (
    <>
      <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
        <DialogTitle>Export to Excel</DialogTitle>
      </DialogHeader>

      <div style={{ marginBottom: 12, fontSize: 13, color: "#6b7280" }}>
        Choose a saved report definition to pre-select fields, or start fresh.
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
        {(savedReports || [])
          .slice()
          .sort((a, b) => String(a.name).localeCompare(String(b.name)))
          .map((report) => (
            <div
              key={report.reportId}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                border: "1px solid #e5e7eb",
                borderRadius: 8,
                padding: "10px 12px",
                cursor: "pointer",
                background: "#fafafa",
              }}
              onClick={() => pickReport(report)}
            >
              <FileSpreadsheet size={16} style={{ color: "#2563eb", flexShrink: 0 }} />
              <span style={{ flex: 1, fontWeight: 500, fontSize: 13 }}>{report.name}</span>
              <span style={{ fontSize: 11, color: "#9ca3af" }}>
                {(report.fieldIds || []).length} field{(report.fieldIds || []).length !== 1 ? "s" : ""}
              </span>
              <button
                style={{
                  background: "none", border: "none", cursor: "pointer",
                  padding: 4, color: "#9ca3af", display: "flex", alignItems: "center",
                }}
                title="Delete this report"
                disabled={isBusy}
                onClick={(e) => handleDeleteReport(report, e)}
              >
                <Trash2 size={14} />
              </button>
            </div>
          ))}
      </div>

      {error && <div style={{ color: "#dc2626", fontSize: 13, marginBottom: 8 }}>{error}</div>}

      <DialogFooter style={{ justifyContent: "space-between" }}>
        <Button variant="secondary" onClick={onClose}>Cancel</Button>
        <Button type="button" onClick={skipPick}>
          <Plus size={14} style={{ marginRight: 4 }} />
          New export
        </Button>
      </DialogFooter>
    </>
  );

  // ── Configure step ─────────────────────────────────────────────────────────

  const renderConfigure = () => {
    const allChecked = sortedFields.length > 0 && checkedIds.size === sortedFields.length;

    return (
      <>
        <DialogHeader style={{ marginBottom: 16, marginLeft: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            {savedReports && savedReports.length > 0 && (
              <button
                style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", color: "#6b7280" }}
                onClick={() => { setStep("pick"); setError(""); }}
              >
                <ChevronLeft size={18} />
              </button>
            )}
            <DialogTitle style={{ margin: 0 }}>Configure export</DialogTitle>
          </div>
        </DialogHeader>

        {/* Report name */}
        <div className="form-row" style={{ marginBottom: 16 }}>
          <label className="form-label">Report name (optional — saves for reuse)</label>
          <input
            ref={nameRef}
            className="form-input"
            type="text"
            value={reportName}
            onChange={(e) => setReportName(e.target.value)}
            placeholder="e.g. Annual entity dump"
            autoComplete="off"
            data-lpignore="true"
            onKeyDown={(e) => { if (e.key === "Enter") handleExport(); }}
          />
        </div>

        {/* Field selector */}
        <div style={{ marginBottom: 8, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontSize: 13, fontWeight: 600, color: "#374151" }}>
            Fields ({checkedIds.size} of {sortedFields.length} selected)
          </span>
          {sortedFields.length > 0 && (
            <button
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 12, color: "#2563eb", padding: 0 }}
              onClick={toggleAll}
            >
              {allChecked ? "Clear all" : "Select all"}
            </button>
          )}
        </div>

        {sortedFields.length === 0 ? (
          <div style={{ fontSize: 13, color: "#9ca3af", marginBottom: 16 }}>
            No data dictionary fields defined. Configure fields in Settings → Data Dictionary.
          </div>
        ) : (
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 6,
              maxHeight: 280,
              overflowY: "auto",
              border: "1px solid #e5e7eb",
              borderRadius: 8,
              padding: "8px 10px",
              marginBottom: 16,
            }}
          >
            {sortedFields.map((field) => {
              const badge = KIND_BADGE[field.appliesTo] || KIND_BADGE.both;
              return (
                <label
                  key={field.fieldId}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    cursor: "pointer",
                    padding: "4px 0",
                    fontSize: 13,
                    userSelect: "none",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={checkedIds.has(field.fieldId)}
                    onChange={() => toggleField(field.fieldId)}
                    style={{ width: 15, height: 15, accentColor: "#2563eb", flexShrink: 0 }}
                  />
                  <span style={{ flex: 1 }}>{field.prompt}</span>
                  <span
                    style={{
                      fontSize: 10,
                      fontWeight: 600,
                      color: badge.color,
                      background: `${badge.color}15`,
                      borderRadius: 4,
                      padding: "1px 5px",
                      flexShrink: 0,
                    }}
                  >
                    {badge.label}
                  </span>
                </label>
              );
            })}
          </div>
        )}

        <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 8 }}>
          {(exportNodes || []).length} record{(exportNodes || []).length !== 1 ? "s" : ""} will be exported.
        </div>

        {error && <div style={{ color: "#dc2626", fontSize: 13, marginBottom: 8 }}>{error}</div>}

        <DialogFooter style={{ justifyContent: "flex-end" }}>
          <Button variant="secondary" onClick={onClose}>Cancel</Button>
          <Button
            type="button"
            disabled={isBusy || checkedIds.size === 0}
            onClick={handleExport}
          >
            {isBusy ? "Exporting…" : "Export"}
          </Button>
        </DialogFooter>
      </>
    );
  };

  return (
    <Dialog open={open} onOpenChange={(v) => { if (!v) onClose(); }}>
      <DialogContent style={{ width: "min(560px, 92vw)", maxWidth: "none" }}>
        {step === "pick" ? renderPick() : renderConfigure()}
      </DialogContent>
    </Dialog>
  );
}
