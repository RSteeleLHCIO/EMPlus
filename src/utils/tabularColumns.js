const normalizeColumnKeys = (keys) =>
  Array.isArray(keys)
    ? keys.filter((key) => typeof key === "string").map((key) => {
      const text = key.trim();
      return text.startsWith("dd:dd:") ? `dd:${text.slice(6)}` : text;
    }).filter(Boolean)
    : [];

export function normalizeSavedColumnOrder(columnOrder, hidden) {
  const hiddenSet = new Set(normalizeColumnKeys(hidden));
  // Keep unresolved IDs: dictionary definitions may not have loaded yet.
  return normalizeColumnKeys(columnOrder).filter((key) => !hiddenSet.has(key));
}

export function resolveTabularColumns(columnOrder, columns) {
  const byKey = new Map(columns.map((column) => [column.key, column]));
  const seen = new Set();
  return (columnOrder || [])
    .map((key) => byKey.get(key))
    .filter((column) => {
      if (!column || seen.has(column.key)) return false;
      seen.add(column.key);
      return true;
    });
}
