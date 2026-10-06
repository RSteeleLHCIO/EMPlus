// Helpers for EMPlusNodeHistory versions as returned by GET /api/nodes/:id/history.
// effectiveTo is exclusive: a version applies while effectiveFrom <= date < effectiveTo.

const MIN_DATE = "0001-01-01";
const MAX_DATE = "9999-12-31";

export const isVersionActiveAt = (version, date) => {
  if (!version || !date) return false;
  const start = version.effectiveFrom || MIN_DATE;
  const end = version.effectiveTo || MAX_DATE;
  return start <= date && date < end;
};

// Index of the version in effect on `date`, or the closest earlier one if `date` falls in a gap.
export const findVersionIndexAt = (versions, date) => {
  if (!Array.isArray(versions) || versions.length === 0) return -1;
  for (let i = versions.length - 1; i >= 0; i -= 1) {
    if (isVersionActiveAt(versions[i], date)) return i;
  }
  for (let i = versions.length - 1; i >= 0; i -= 1) {
    if ((versions[i].effectiveFrom || MIN_DATE) <= date) return i;
  }
  return 0;
};

export const getVersionStatus = (version, today) => {
  if (!version) return "current";
  if ((version.effectiveFrom || MIN_DATE) > today) return "future";
  return isVersionActiveAt(version, today) ? "current" : "historical";
};

export const addDaysIso = (iso, days) => {
  const [y, m, d] = String(iso).split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  return dt.toISOString().slice(0, 10);
};

export const formatIsoDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = String(iso).slice(0, 10).split("-");
  if (!y || !m || !d) return String(iso);
  return `${m}/${d}/${y}`;
};

// Human-readable range using an inclusive end date (the day before effectiveTo).
export const formatInclusiveEnd = (toExclusive) =>
  (toExclusive ? formatIsoDate(addDaysIso(toExclusive, -1)) : "");

export const formatVersionRange = (from, toExclusive) => {
  const start = from ? formatIsoDate(from) : "the beginning";
  if (!toExclusive) return `${start} – present`;
  return `${start} – ${formatInclusiveEnd(toExclusive)}`;
};

const normalizeValue = (value) => {
  if (value === undefined || value === null || value === "") return "";
  if (Array.isArray(value)) {
    const cleaned = value.filter((v) => v !== "" && v !== null && v !== undefined);
    return cleaned.length ? JSON.stringify(cleaned) : "";
  }
  if (typeof value === "object") {
    const keys = Object.keys(value).sort();
    return JSON.stringify(keys.reduce((acc, k) => ({ ...acc, [k]: value[k] }), {}));
  }
  return String(value);
};

const emailsValue = (v) => (Array.isArray(v) ? v.join(", ") : (v || ""));

export const NODE_VERSION_FIELDS = [
  "name", "kind", "photo", "logo", "address", "workPhone", "cellPhone",
  "emails", "taxId", "operationalRole", "legalStatus", "personStatus",
];

// Map of changed keys → previous value. Built-in fields use their own key;
// data-dictionary fields use `cf:<fieldId>`.
export const getChangedVersionFields = (previous, current) => {
  const changed = new Map();
  if (!previous || !current) return changed;
  NODE_VERSION_FIELDS.forEach((key) => {
    const prev = key === "emails" ? emailsValue(previous[key]) : previous[key];
    const next = key === "emails" ? emailsValue(current[key]) : current[key];
    if (normalizeValue(prev) !== normalizeValue(next)) changed.set(key, prev);
  });
  const prevCustom = previous.customFields || {};
  const nextCustom = current.customFields || {};
  new Set([...Object.keys(prevCustom), ...Object.keys(nextCustom)]).forEach((fieldId) => {
    if (normalizeValue(prevCustom[fieldId]) !== normalizeValue(nextCustom[fieldId])) {
      changed.set(`cf:${fieldId}`, prevCustom[fieldId]);
    }
  });
  return changed;
};

export const describeVersionValue = (value) => {
  if (value === undefined || value === null || value === "") return "(blank)";
  if (Array.isArray(value)) {
    const parts = value
      .map((v) => (v && typeof v === "object" ? (v.number || v.url || v.name || JSON.stringify(v)) : v))
      .filter((v) => v !== "" && v !== null && v !== undefined);
    return parts.length ? parts.join(", ") : "(blank)";
  }
  if (typeof value === "object") return JSON.stringify(value);
  if (value === true || value === "true") return "Yes";
  if (value === false || value === "false") return "No";
  return String(value);
};

// Mirrors the server's persistNodeVersion: what range will a save effective `date` occupy?
export const describeSaveRange = (versions, date) => {
  const list = Array.isArray(versions) ? versions : [];
  const replaces = list.find((v) => v.effectiveFrom === date) || null;
  const others = list.filter((v) => v.effectiveFrom !== date);
  const nextStart = others
    .map((v) => v.effectiveFrom)
    .filter((start) => start && start > date)
    .sort((a, b) => a.localeCompare(b))[0] || null;
  const splits = others.find(
    (v) => (v.effectiveFrom || MIN_DATE) < date && isVersionActiveAt(v, date)
  ) || null;
  return { from: date, to: nextStart, replaces, splits, isLatest: !nextStart };
};
