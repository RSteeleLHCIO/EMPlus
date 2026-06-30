// Month name → 0-based index (English full + abbreviated)
const _MONTH_NAMES = {
  january:0, february:1, march:2, april:3, may:4, june:5,
  july:6, august:7, september:8, october:9, november:10, december:11,
  jan:0, feb:1, mar:2, apr:3, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
};

function _dateToIso(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/**
 * Normalize any reasonable date input to an ISO "YYYY-MM-DD" string.
 * Returns "" for unrecognized / empty input. Never throws.
 *
 * Handles:
 *   - JS Date objects
 *   - JS timestamps (ms since Unix epoch) — numbers >= 10_000_000_000
 *   - Excel serial numbers (1–2958465, i.e. 1900-01-01 to 9999-12-31)
 *   - ISO 8601: "2024-01-15", "2024-01-15T..."
 *   - Slash-separated: "1/15/2024", "01/15/2024", "15/01/2024" (month>12 → D/M/Y)
 *   - Dash-separated: "15-01-2024", "2024-01-15"
 *   - Dot-separated:  "15.01.2024", "01.15.2024"
 *   - Long month name: "January 15, 2024", "15 January 2024"
 *   - Short month name: "Jan 15 2024", "15 Jan 2024"
 *
 * @param {*} raw
 * @returns {string} "YYYY-MM-DD" or ""
 */
export function normalizeDateInput(raw) {
  if (raw === null || raw === undefined || raw === "") return "";

  // ── JS Date object ──────────────────────────────────────────────────────────
  if (raw instanceof Date) return _dateToIso(raw);

  const n = Number(raw);

  if (!isNaN(n) && raw !== "" && String(raw).trim() !== "") {
    // ── JS timestamp (ms) — >= 10_000_000_000 to avoid colliding with Excel serials ──
    if (n >= 10_000_000_000) {
      return _dateToIso(new Date(n));
    }

    // ── Excel serial number (1 – 2958465 covers 1900-01-01 to 9999-12-31) ────
    // Excel epoch: Dec 30 1899 UTC (accounts for the 1900 leap-year bug).
    // Serials >= 61 subtract 1 extra day for the phantom Feb 29 1900.
    if (Number.isInteger(n) && n >= 1 && n <= 2_958_465) {
      const adjusted = n >= 61 ? n - 1 : n;
      const ms = Date.UTC(1899, 11, 31) + adjusted * 86_400_000;
      return _dateToIso(new Date(ms));
    }
  }

  const s = String(raw).trim();
  if (!s) return "";

  // ── ISO 8601 with optional time component: "2024-01-15" or "2024-01-15T…" ──
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    return _dateToIso(new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3])));
  }

  // ── Slash-separated ─────────────────────────────────────────────────────────
  const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    const [, a, b, y] = slash.map(Number);
    const [mo, dy] = a > 12 ? [b, a] : [a, b];
    return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
  }

  // ── Dash-separated (non-ISO): "15-01-2024" or "01-15-2024" ─────────────────
  const dash = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dash) {
    const [, a, b, y] = dash.map(Number);
    const [mo, dy] = a > 12 ? [b, a] : [a, b];
    return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
  }

  // ── Dot-separated: "15.01.2024" or "01.15.2024" ─────────────────────────────
  const dot = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (dot) {
    const [, a, b, y] = dot.map(Number);
    const [mo, dy] = a > 12 ? [b, a] : [a, b];
    return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
  }

  // ── Month name formats ───────────────────────────────────────────────────────
  // "January 15, 2024" / "January 15 2024" / "Jan 15, 2024"
  const mdy = s.match(/^([A-Za-z]+)\s+(\d{1,2})[,\s]+(\d{4})$/);
  if (mdy) {
    const mo = _MONTH_NAMES[mdy[1].toLowerCase()];
    if (mo !== undefined) return _dateToIso(new Date(Date.UTC(+mdy[3], mo, +mdy[2])));
  }
  // "15 January 2024" / "15 Jan 2024"
  const dmy = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (dmy) {
    const mo = _MONTH_NAMES[dmy[2].toLowerCase()];
    if (mo !== undefined) return _dateToIso(new Date(Date.UTC(+dmy[3], mo, +dmy[1])));
  }

  return "";
}
