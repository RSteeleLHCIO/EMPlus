// Duplicate-name detection against the index returned by GET /api/node-names.
import { formatInclusiveEnd, formatIsoDate } from "./nodeVersions.js";

// Case, spacing and punctuation are ignored ("Acme, Inc." matches "acme inc").
export const normalizeNodeName = (name) =>
  String(name || "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();

// Returns other nodes of the same kind whose current or former name matches `name`.
// `extraNodes` (e.g. the loaded node list) covers nodes missing from a stale index.
export const findNameConflicts = (index, { name, kind, excludeId, extraNodes = [] }) => {
  const target = normalizeNodeName(name);
  if (!target) return [];
  const entries = new Map();
  (Array.isArray(index) ? index : []).forEach((entry) => {
    if (entry?.id) entries.set(entry.id, entry);
  });
  extraNodes.forEach((node) => {
    if (!node?.id || entries.has(node.id)) return;
    entries.set(node.id, { id: node.id, kind: node.kind, currentName: node.name, deleted: false, names: [] });
  });

  const conflicts = [];
  entries.forEach((entry) => {
    if (entry.id === excludeId) return;
    if (!entry.deleted && entry.kind === kind && normalizeNodeName(entry.currentName) === target) {
      conflicts.push({ id: entry.id, kind: entry.kind, type: "current", currentName: entry.currentName });
      return;
    }
    const matches = (entry.names || []).filter(
      (v) => (v.kind || entry.kind) === kind && normalizeNodeName(v.name) === target
    );
    if (!matches.length) return;
    const from = matches.map((v) => v.effectiveFrom).filter(Boolean).sort()[0] || null;
    const openEnded = matches.some((v) => !v.effectiveTo);
    const to = openEnded ? null : matches.map((v) => v.effectiveTo).sort().slice(-1)[0];
    conflicts.push({
      id: entry.id,
      kind: entry.kind,
      type: entry.deleted ? "deleted" : "past",
      currentName: entry.currentName,
      formerName: matches[matches.length - 1].name,
      from,
      to,
    });
  });
  const order = { current: 0, past: 1, deleted: 2 };
  return conflicts.sort((a, b) => order[a.type] - order[b.type]);
};

export const describeNameConflict = (conflict) => {
  const kindLabel = conflict.kind === "person" ? "person" : "entity";
  if (conflict.type === "current") {
    return `An existing ${kindLabel} is named "${conflict.currentName}".`;
  }
  const range = [
    conflict.from ? `from ${formatIsoDate(conflict.from)}` : "",
    conflict.to ? `through ${formatInclusiveEnd(conflict.to)}` : "",
  ].filter(Boolean).join(" ");
  if (conflict.type === "deleted") {
    return `A deleted ${kindLabel} was named "${conflict.formerName}"${range ? ` ${range}` : ""}.`;
  }
  return `"${conflict.formerName}" was the former name of the ${kindLabel} now named "${conflict.currentName}"${range ? ` (${range})` : ""}.`;
};
