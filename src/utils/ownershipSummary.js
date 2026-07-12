export const ENTITY_OWNERSHIP_SUMMARY_FIELD = {
  fieldId: "__virtual__entity_ownership_summary",
  prompt: "Ownership Records",
  dataType: "virtual",
  appliesTo: "entity",
  multiValue: false,
  sortOrder: -1000,
  _virtual: true,
  _readOnly: true,
};

const formatPercent = (value) => {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return "";
  return `${parsed}%`;
};

export const getEntityOwnershipSummary = (node, nodeList, relList, asOfDate = null, ownershipTimeline = null) => {
  if (!node || node.kind !== "entity" || !node.id) return "";

  const rels = (relList || [])
    .filter((rel) => rel.type === "owns" && rel.to === node.id);

  const rows = rels
    .map((rel) => {
      const ownerNode = (nodeList || []).find((n) => n.id === rel.from);
      const ownerName = String(ownerNode?.name || rel.from || "").trim();
      return {
        ownerName,
        percentNum: Number(rel.percent),
        percentText: formatPercent(rel.percent),
        effectiveTo: rel.effectiveTo || "",
      };
    })
    .filter((row) => row.ownerName);

  if (rows.length === 0) {
    // When no ownership records exist, check if there's a future effective date
    if (Array.isArray(ownershipTimeline) && ownershipTimeline.length > 0) {
      const earliestFutureDate = ownershipTimeline
        .map((p) => p.effectiveFrom)
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b))[0];
      
      if (earliestFutureDate) {
        // Format date: parse as YYYY-MM-DD and return MM/DD/YYYY
        try {
          const [year, month, day] = earliestFutureDate.split('-').map(Number);
          const d = new Date(year, month - 1, day);
          const formatted = d.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
          return `No ownership records before ${formatted}`;
        } catch {
          return `No ownership records before ${earliestFutureDate}`;
        }
      }
    }
    return "";
  }

  rows.sort((a, b) => {
    const aPct = Number.isFinite(a.percentNum) ? a.percentNum : -Infinity;
    const bPct = Number.isFinite(b.percentNum) ? b.percentNum : -Infinity;
    if (aPct !== bPct) return bPct - aPct;
    return a.ownerName.localeCompare(b.ownerName, undefined, { sensitivity: "base" });
  });

  const summary = rows
    .map((row) => (row.percentText ? `${row.ownerName} (${row.percentText})` : row.ownerName))
    .join("; ");

  // Check if this ownership structure is no longer current:
  // Historical records have effectiveTo set; if any is in the past relative to today,
  // the structure shown is not the current one.
  const today = new Date().toISOString().split("T")[0];
  const isStale = rows.some((row) => row.effectiveTo && row.effectiveTo !== "9999-12-31" && row.effectiveTo < today);

  if (isStale) {
    return summary + "  - THIS IS NOT THE CURRENT OWNERSHIP STRUCTURE";
  }

  return summary;
};
