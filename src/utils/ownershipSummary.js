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

export const getEntityOwnershipSummary = (node, nodeList, relList) => {
  if (!node || node.kind !== "entity" || !node.id) return "";

  const rows = (relList || [])
    .filter((rel) => rel.type === "owns" && rel.to === node.id)
    .map((rel) => {
      const ownerNode = (nodeList || []).find((n) => n.id === rel.from);
      const ownerName = String(ownerNode?.name || rel.from || "").trim();
      return {
        ownerName,
        percentNum: Number(rel.percent),
        percentText: formatPercent(rel.percent),
      };
    })
    .filter((row) => row.ownerName);

  if (rows.length === 0) return "";

  rows.sort((a, b) => {
    const aPct = Number.isFinite(a.percentNum) ? a.percentNum : -Infinity;
    const bPct = Number.isFinite(b.percentNum) ? b.percentNum : -Infinity;
    if (aPct !== bPct) return bPct - aPct;
    return a.ownerName.localeCompare(b.ownerName, undefined, { sensitivity: "base" });
  });

  return rows
    .map((row) => (row.percentText ? `${row.ownerName} (${row.percentText})` : row.ownerName))
    .join("; ");
};
