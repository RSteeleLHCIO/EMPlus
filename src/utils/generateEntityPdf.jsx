/**
 * generateEntityPdf.jsx
 *
 * Reusable PDF generation utilities for entity/person records.
 *
 * Exports:
 *   - HierarchyPageContent   – React component that renders the hierarchy view
 *                              for one node (owners → focus → owns). Used both
 *                              here and in the future "entity book" feature.
 *   - EntityInfoPageContent  – React component that renders the entity/person
 *                              detail fields.
 *   - generateEntityPdf      – async function that builds a jsPDF document with
 *                              page 1 = hierarchy, page 2+ = entity info.
 */

import React from "react";
import { Building2, Users } from "lucide-react";
import { renderToStaticMarkup } from "react-dom/server";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

// ─── helpers ────────────────────────────────────────────────────────────────

const getNode = (nodeList, id) => nodeList.find((n) => n.id === id);

const getOwnersOf = (relList, entityId) =>
  relList
    .filter((r) => r.type === "owns" && r.to === entityId)
    .map((r) => ({ nodeId: r.from, rel: r }))
    .sort((a, b) => {
      const aP = Number(a.rel?.percent);
      const bP = Number(b.rel?.percent);
      const aV = Number.isFinite(aP) ? aP : -Infinity;
      const bV = Number.isFinite(bP) ? bP : -Infinity;
      return bV - aV;
    });

const getOwnedBy = (relList, ownerId) =>
  relList
    .filter((r) => r.type === "owns" && r.from === ownerId)
    .map((r) => ({ nodeId: r.to, rel: r }));

const formatDate = (iso) => {
  if (!iso) return "";
  try {
    return new Date(iso + "T00:00:00").toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  } catch {
    return iso;
  }
};

const formatFieldValue = (field, value) => {
  if (value == null || value === "" || value === false) return "";
  const { dataType } = field;
  if (dataType === "boolean") return value ? "Yes" : "No";
  if (dataType === "date") return formatDate(String(value));
  if (dataType === "currency") {
    const num = Number(value);
    return Number.isFinite(num)
      ? num.toLocaleString(undefined, { style: "currency", currency: "USD" })
      : String(value);
  }
  if (dataType === "percentage") return `${value}%`;
  if (dataType === "phone") {
    if (Array.isArray(value)) {
      return value.map((e) => `${e.type}: ${e.number}`).filter((s) => s.trim() !== ": ").join(", ");
    }
    if (value && typeof value === "object") {
      return Object.entries(value)
        .map(([t, n]) => `${t}: ${n}`)
        .join(", ");
    }
    return String(value);
  }
  if (Array.isArray(value)) return value.join(", ");
  return String(value);
};

// ─── HierarchyPageContent ────────────────────────────────────────────────────

/**
 * Renders a self-contained hierarchy view for `nodeId`.
 * Pass the same nodeList / relList / dataDictionary from EntityApp.
 *
 * Props:
 *   nodeId        – the focal entity/person
 *   nodeList      – full node array
 *   relList       – full relationship array
 *   clientName    – display name for the client (shown as page title)
 *   asOf          – optional Date shown as "as of …" subtitle
 */
export function HierarchyPageContent({ nodeId, nodeList, relList, clientName, asOf }) {
  const focusNode = getNode(nodeList, nodeId);
  const owners = getOwnersOf(relList, nodeId);
  const owned = getOwnedBy(relList, nodeId);

  const ownerTotal = owners.reduce((s, o) => {
    const v = Number(o.rel?.percent);
    return s + (Number.isFinite(v) ? v : 0);
  }, 0);
  const ownerTotalOk = Math.abs(ownerTotal - 100) < 0.01;
  const gap = Math.round((100 - ownerTotal) * 10) / 10;

  const boxStyle = {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: 4,
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "10px 14px",
    minWidth: 100,
    maxWidth: 140,
    background: "#ffffff",
    fontSize: 12,
    textAlign: "center",
  };

  const focusBoxStyle = {
    ...boxStyle,
    border: "2px solid #2563eb",
    borderRadius: 10,
    padding: "14px 20px",
    minWidth: 140,
    fontWeight: 600,
    fontSize: 14,
    background: "#eff6ff",
  };

  const connectorStyle = {
    width: 2,
    height: 24,
    background: "#9ca3af",
    margin: "0 auto",
  };

  const rowStyle = {
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "center",
    gap: 10,
  };

  const sectionLabelStyle = {
    fontSize: 11,
    fontWeight: 600,
    color: "#6b7280",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    marginBottom: 8,
    textAlign: "center",
  };

  const NodeBox = ({ nodeItem, style }) => {
    const n = getNode(nodeList, nodeItem.nodeId);
    if (!n) return null;
    const pct = nodeItem.rel?.percent;
    const showPct = pct != null && Number.isFinite(Number(pct));
    return (
      <div style={style || boxStyle}>
        {n.kind === "person" ? (
          n.photo
            ? <img src={n.photo} alt="" style={{ width: 32, height: 32, borderRadius: "50%", objectFit: "cover" }} />
            : <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#6b7280" strokeWidth="2"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>
        ) : (
          n.logo
            ? <img src={n.logo} alt="" style={{ width: 32, height: 32, objectFit: "contain" }} />
            : <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#6b7280" strokeWidth="2"><rect x="3" y="9" width="18" height="12" rx="1"/><path d="M8 9V5a2 2 0 0 1 4 0v4"/></svg>
        )}
        <span style={{ fontWeight: 500, fontSize: 12, wordBreak: "break-word" }}>{n.name}</span>
        {showPct && <span style={{ fontSize: 11, color: "#6b7280" }}>{Number(pct)}%</span>}
      </div>
    );
  };

  return (
    <div
      style={{
        fontFamily: "Inter, system-ui, Arial, sans-serif",
        padding: 32,
        background: "#ffffff",
        minHeight: "100%",
      }}
    >
      {/* Page header */}
      {clientName && (
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 18, fontWeight: 700, color: "#1e293b" }}>{clientName}</div>
          <div style={{ fontSize: 13, color: "#64748b", marginTop: 2 }}>Hierarchy View</div>
          {asOf && (
            <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 2 }}>
              As of {asOf.toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" })}
            </div>
          )}
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 0 }}>

        {/* ── Owners section ── */}
        {focusNode?.kind !== "person" && owners.length > 0 && (
          <>
            <div style={sectionLabelStyle}>Owners</div>
            <div style={rowStyle}>
              {owners.map((o) => <NodeBox key={o.nodeId} nodeItem={o} />)}
              {!ownerTotalOk && ownerTotal > 0 && gap > 0 && (
                <div style={{ ...boxStyle, border: "1px dashed #9ca3af", background: "#f9fafb" }}>
                  <span style={{ fontWeight: 500, fontSize: 12, color: "#9ca3af" }}>Unknown</span>
                  <span style={{ fontSize: 11, color: "#9ca3af" }}>{gap}%</span>
                </div>
              )}
            </div>
            <div style={connectorStyle} />
          </>
        )}

        {/* ── Focus box ── */}
        <div style={focusBoxStyle}>
          {focusNode?.kind === "person" ? (
            focusNode.photo
              ? <img src={focusNode.photo} alt="" style={{ width: 40, height: 40, borderRadius: "50%", objectFit: "cover" }} />
              : <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth="2"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>
          ) : (
            focusNode?.logo
              ? <img src={focusNode.logo} alt="" style={{ width: 40, height: 40, objectFit: "contain" }} />
              : <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth="2"><rect x="3" y="9" width="18" height="12" rx="1"/><path d="M8 9V5a2 2 0 0 1 4 0v4"/></svg>
          )}
          <span>{focusNode?.name || nodeId}</span>
          <span style={{ fontSize: 11, fontWeight: 400, color: "#64748b", textTransform: "capitalize" }}>
            {focusNode?.kind || "entity"}
          </span>
        </div>

        {/* ── Owned section ── */}
        {owned.length > 0 && (
          <>
            <div style={connectorStyle} />
            <div style={sectionLabelStyle}>Owns</div>
            <div style={rowStyle}>
              {owned.map((o) => <NodeBox key={o.nodeId} nodeItem={o} />)}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── EntityInfoPageContent ───────────────────────────────────────────────────

/**
 * Renders the entity/person information (fields) as a print-ready page.
 *
 * Props:
 *   nodeId         – the focal entity/person id
 *   nodeList       – full node array
 *   relList        – full relationship array
 *   dataDictionary – DD field definitions
 *   clientName     – display name for the client
 */
export function EntityInfoPageContent({ nodeId, nodeList, relList, dataDictionary, clientName }) {
  const node = getNode(nodeList, nodeId);
  if (!node) return null;

  const owners = getOwnersOf(relList, nodeId);
  const owned = getOwnedBy(relList, nodeId);

  const fields = [...(dataDictionary || [])]
    .filter((f) => f.appliesTo === "both" || f.appliesTo === node.kind)
    .sort((a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0));

  const labelStyle = { fontSize: 11, fontWeight: 600, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.05em" };
  const valueStyle = { fontSize: 13, color: "#1e293b", marginTop: 2 };
  const rowStyle = { marginBottom: 16 };
  const dividerStyle = { border: "none", borderTop: "1px solid #e2e8f0", margin: "20px 0" };

  return (
    <div
      style={{
        fontFamily: "Inter, system-ui, Arial, sans-serif",
        padding: 32,
        background: "#ffffff",
        minHeight: "100%",
      }}
    >
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 24 }}>
        <div style={{
          width: 48, height: 48, borderRadius: node.kind === "person" ? "50%" : 8,
          background: "#eff6ff", display: "flex", alignItems: "center", justifyContent: "center",
          flexShrink: 0,
        }}>
          {node.kind === "person" ? (
            node.photo
              ? <img src={node.photo} alt="" style={{ width: 48, height: 48, borderRadius: "50%", objectFit: "cover" }} />
              : <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth="2"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>
          ) : (
            node.logo
              ? <img src={node.logo} alt="" style={{ width: 40, height: 40, objectFit: "contain" }} />
              : <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth="2"><rect x="3" y="9" width="18" height="12" rx="1"/><path d="M8 9V5a2 2 0 0 1 4 0v4"/></svg>
          )}
        </div>
        <div>
          {clientName && <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 2 }}>{clientName}</div>}
          <div style={{ fontSize: 20, fontWeight: 700, color: "#1e293b" }}>{node.name}</div>
          <div style={{ fontSize: 12, color: "#64748b", textTransform: "capitalize" }}>{node.kind}</div>
        </div>
      </div>

      <hr style={dividerStyle} />

      {/* System fields */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 24px" }}>
        <div style={rowStyle}>
          <div style={labelStyle}>Name</div>
          <div style={valueStyle}>{node.name}</div>
        </div>
        <div style={rowStyle}>
          <div style={labelStyle}>Type</div>
          <div style={{ ...valueStyle, textTransform: "capitalize" }}>{node.kind}</div>
        </div>
      </div>

      {/* Custom fields */}
      {fields.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 24px" }}>
          {fields.map((field) => {
            const raw = node.customFields?.[field.fieldId];
            const display = formatFieldValue(field, raw);
            if (!display) return null;
            return (
              <div key={field.fieldId} style={rowStyle}>
                <div style={labelStyle}>{field.prompt}</div>
                <div style={valueStyle}>{display}</div>
              </div>
            );
          })}
        </div>
      )}

      {/* Ownership summary */}
      {(owners.length > 0 || owned.length > 0) && (
        <>
          <hr style={dividerStyle} />
          <div style={{ fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 12 }}>Ownership</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 24px" }}>
            {node.kind !== "person" && owners.length > 0 && (
              <div style={rowStyle}>
                <div style={labelStyle}>Owned by</div>
                <div style={valueStyle}>
                  {owners.map((o) => {
                    const n = getNode(nodeList, o.nodeId);
                    const pct = o.rel?.percent;
                    const showPct = pct != null && Number.isFinite(Number(pct));
                    return (
                      <div key={o.nodeId} style={{ marginBottom: 2 }}>
                        {n?.name || o.nodeId}
                        {showPct ? ` (${Number(pct)}%)` : ""}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
            {owned.length > 0 && (
              <div style={rowStyle}>
                <div style={labelStyle}>Owns</div>
                <div style={valueStyle}>
                  {owned.map((o) => {
                    const n = getNode(nodeList, o.nodeId);
                    const pct = o.rel?.percent;
                    const showPct = pct != null && Number.isFinite(Number(pct));
                    return (
                      <div key={o.nodeId} style={{ marginBottom: 2 }}>
                        {n?.name || o.nodeId}
                        {showPct ? ` (${Number(pct)}%)` : ""}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </>
      )}

      {/* Print date footer */}
      <div style={{ marginTop: 40, fontSize: 11, color: "#9ca3af", borderTop: "1px solid #e2e8f0", paddingTop: 12 }}>
        Generated {new Date().toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" })}
      </div>
    </div>
  );
}

// ─── generateEntityPdf ───────────────────────────────────────────────────────

/**
 * Generates a two-page PDF for a single entity/person.
 *
 * Page 1: Hierarchy view (owners → entity → owns)
 * Page 2: Entity/person information
 *
 * @param {object} opts
 * @param {string}   opts.nodeId
 * @param {Array}    opts.nodeList
 * @param {Array}    opts.relList
 * @param {Array}    opts.dataDictionary
 * @param {string}   opts.clientName
 * @param {boolean}  [opts.download=true]  – if true, triggers browser download;
 *                                           if false, returns the jsPDF instance
 * @returns {Promise<jsPDF|void>}
 */
export async function generateEntityPdf({
  nodeId,
  nodeList,
  relList,
  dataDictionary,
  clientName,
  download = true,
}) {
  const node = getNode(nodeList, nodeId);
  const fileName = `${node?.name || nodeId}.pdf`
    .replace(/[^\w\s.-]/g, "")
    .replace(/\s+/g, "_");

  // Helper: render a React element into an off-screen div, capture it, return canvas
  const captureComponent = async (element, width = 794) => {
    const container = document.createElement("div");
    container.style.cssText = `
      position: fixed;
      left: -9999px;
      top: 0;
      width: ${width}px;
      background: #ffffff;
      font-family: Inter, system-ui, Arial, sans-serif;
    `;
    document.body.appendChild(container);

    // Render React element to static HTML string and set as innerHTML
    // (avoids needing ReactDOM.render in this utility)
    const { createRoot } = await import("react-dom/client");
    const root = createRoot(container);
    await new Promise((resolve) => {
      root.render(element);
      // Give React one tick to paint
      setTimeout(resolve, 100);
    });

    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
    });

    root.unmount();
    document.body.removeChild(container);
    return canvas;
  };

  const pageWidth = 794; // ~A4 width @ 96dpi

  const canvas1 = await captureComponent(
    <HierarchyPageContent
      nodeId={nodeId}
      nodeList={nodeList}
      relList={relList}
      clientName={clientName}
      asOf={new Date()}
    />,
    pageWidth
  );

  const canvas2 = await captureComponent(
    <EntityInfoPageContent
      nodeId={nodeId}
      nodeList={nodeList}
      relList={relList}
      dataDictionary={dataDictionary}
      clientName={clientName}
    />,
    pageWidth
  );

  // Build PDF (A4 portrait)
  const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  const addCanvasPage = (canvas, isFirst) => {
    if (!isFirst) pdf.addPage();
    const imgData = canvas.toDataURL("image/jpeg", 0.92);
    const canvasAspect = canvas.height / canvas.width;
    const imgH = pdfWidth * canvasAspect;

    if (imgH <= pdfHeight) {
      pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, imgH);
    } else {
      // Content taller than one page — let jsPDF scale to fit
      pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, pdfHeight);
    }
  };

  addCanvasPage(canvas1, true);
  addCanvasPage(canvas2, false);

  if (download) {
    pdf.save(fileName);
  } else {
    return pdf;
  }
}

// ─── generateEntityBook ──────────────────────────────────────────────────────

/**
 * Generates a multi-page "entity book" PDF — one page per node.
 *
 * @param {object} opts
 * @param {Array}   opts.nodes          – ordered list of node objects to include
 * @param {Array}   opts.nodeList       – full node list (for relationship lookups)
 * @param {Array}   opts.relList        – full relationship list
 * @param {Array}   opts.dataDictionary – DD field definitions
 * @param {string}  opts.clientName     – shown in each page header
 * @param {"hierarchy"|"info"} opts.pageType
 *                                      – which page to render per node:
 *                                        "hierarchy" = owners → entity → owns
 *                                        "info"      = entity/person detail fields
 * @param {string}  [opts.fileName]     – override the output file name
 */
export async function generateEntityBook({
  nodes,
  nodeList,
  relList,
  dataDictionary,
  clientName,
  pageType = "hierarchy",
  fileName,
}) {
  if (!nodes || nodes.length === 0) return;

  const { createRoot } = await import("react-dom/client");

  const captureComponent = async (element, width = 794) => {
    const container = document.createElement("div");
    container.style.cssText = `
      position: fixed;
      left: -9999px;
      top: 0;
      width: ${width}px;
      background: #ffffff;
      font-family: Inter, system-ui, Arial, sans-serif;
    `;
    document.body.appendChild(container);
    const root = createRoot(container);
    await new Promise((resolve) => {
      root.render(element);
      setTimeout(resolve, 100);
    });
    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
    });
    root.unmount();
    document.body.removeChild(container);
    return canvas;
  };

  const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();
  const asOf = new Date();

  const addCanvasPage = (canvas, isFirst) => {
    if (!isFirst) pdf.addPage();
    const imgData = canvas.toDataURL("image/jpeg", 0.92);
    const canvasAspect = canvas.height / canvas.width;
    const imgH = pdfWidth * canvasAspect;
    pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, Math.min(imgH, pdfHeight));
  };

  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    const element =
      pageType === "hierarchy" ? (
        <HierarchyPageContent
          nodeId={node.id}
          nodeList={nodeList}
          relList={relList}
          clientName={clientName}
          asOf={asOf}
        />
      ) : (
        <EntityInfoPageContent
          nodeId={node.id}
          nodeList={nodeList}
          relList={relList}
          dataDictionary={dataDictionary}
          clientName={clientName}
        />
      );

    const canvas = await captureComponent(element, 794);
    addCanvasPage(canvas, i === 0);
  }

  const safeName = (fileName || `${clientName || "entity-book"}-${pageType}`)
    .replace(/[^\w\s.-]/g, "")
    .replace(/\s+/g, "_");
  pdf.save(`${safeName}.pdf`);
}

