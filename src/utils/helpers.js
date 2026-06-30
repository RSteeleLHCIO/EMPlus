/**
 * Format a Date object as YYYY-MM-DD string
 * @param {Date} d - The date to format
 * @returns {string} Date string in YYYY-MM-DD format
 */
export function toKey(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
}

// Month name → 0-based index lookup (English, full + abbreviated)
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

    // ── JS Date object ────────────────────────────────────────────────────────
    if (raw instanceof Date) return _dateToIso(raw);

    const n = Number(raw);

    if (!isNaN(n) && raw !== "" && String(raw).trim() !== "") {
        // ── JS timestamp (ms) — 13+ digit numbers, or >= 10_000_000_000 ───────
        // Smallest plausible: 1970-01-01 = 0, but to avoid colliding with Excel
        // serials we require >= 10_000_000_000 (≈ 1970-04-26).
        if (n >= 10_000_000_000) {
            return _dateToIso(new Date(n));
        }

        // ── Excel serial number (1 – 2958465 covers 1900-01-01 to 9999-12-31) ─
        // Excel epoch: Dec 30 1899 UTC (accounts for the 1900 leap-year bug).
        // Serials 1–60 map to Jan 1 – Feb 28 1900; serial 60 is the phantom
        // Feb 29 1900 (skip it); serials ≥ 61 subtract 1 extra day.
        if (Number.isInteger(n) && n >= 1 && n <= 2_958_465) {
            const adjusted = n >= 61 ? n - 1 : n;
            const ms = Date.UTC(1899, 11, 31) + adjusted * 86_400_000;
            return _dateToIso(new Date(ms));
        }
    }

    const s = String(raw).trim();
    if (!s) return "";

    // ── ISO 8601 with optional time component: "2024-01-15" or "2024-01-15T…" ─
    const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
        const d = new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3]));
        return _dateToIso(d);
    }

    // ── Slash-separated ───────────────────────────────────────────────────────
    const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slash) {
        const [, a, b, y] = slash.map(Number);
        // If first part > 12 it must be day; otherwise assume M/D/Y (US default)
        const [mo, dy] = a > 12 ? [b, a] : [a, b];
        return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
    }

    // ── Dash-separated (non-ISO): "15-01-2024" or "01-15-2024" ───────────────
    const dash = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (dash) {
        const [, a, b, y] = dash.map(Number);
        const [mo, dy] = a > 12 ? [b, a] : [a, b];
        return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
    }

    // ── Dot-separated: "15.01.2024" or "01.15.2024" ───────────────────────────
    const dot = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (dot) {
        const [, a, b, y] = dot.map(Number);
        const [mo, dy] = a > 12 ? [b, a] : [a, b];
        return _dateToIso(new Date(Date.UTC(y, mo - 1, dy)));
    }

    // ── Month name formats ─────────────────────────────────────────────────────
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

/**
 * Format a Date object as a localized time string
 * @param {Date|null} d - The date to format
 * @returns {string|null} Formatted time string or null if invalid
 */
export function fmtTime(d) {
    if (!d) return null;
    try {
        return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    } catch {
        return null;
    }
}

/**
 * Format a Date object as a localized date+time string
 * @param {Date|null} d - The date to format
 * @returns {string|null} Formatted date+time or null if invalid
 */
export function fmtDateTime(d) {
    if (!d) return null;
    try {
        return d.toLocaleString([], {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit'
        });
    } catch {
        return null;
    }
}

/**
 * Convert a string to sentence case (first letter uppercase, rest lowercase)
 * @param {string} str - The string to convert
 * @returns {string} Sentence-cased string
 */
export function toSentenceCase(str) {
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

// ─── Phone utilities ──────────────────────────────────────────────────────────

// 2-digit ITU-T country code prefixes (selected coverage).
const _CC2 = new Set(["20","27","30","31","32","33","34","36","39","40","41","43","44","45","46","47","48","49","51","52","53","54","55","56","57","58","60","61","62","63","64","65","66","81","82","84","86","90","91","92","93","94","95","98"]);
// 3-digit ITU-T country code prefixes (selected coverage).
const _CC3 = new Set(["212","213","216","218","220","221","222","223","224","225","226","227","228","229","230","231","232","233","234","235","236","237","238","239","240","241","242","243","244","245","246","247","248","249","250","251","252","253","254","255","256","257","258","260","261","262","263","264","265","266","267","268","269","290","291","297","298","299","350","351","352","353","354","355","356","357","358","359","370","371","372","373","374","375","376","377","378","380","381","382","385","386","387","389","420","421","423","500","501","502","503","504","505","506","507","508","509","590","591","592","593","594","595","596","597","598","599","670","672","673","674","675","676","677","678","679","680","681","682","683","685","686","687","688","689","690","691","692","850","852","853","855","856","880","886","960","961","962","963","964","965","966","967","968","970","971","972","973","974","975","976","977","992","993","994","995","996","998"]);

/**
 * Normalise a raw phone input to E.164 (+[cc][number]).
 * Assumes US (+1) when there is no leading + and the digit count is ≤ 10.
 */
export function normalizePhone(raw) {
    if (!raw) return "";
    const str = String(raw).trim();
    const hasPlus = str.startsWith("+");
    const digits = str.replace(/\D/g, "");
    if (!digits) return "";
    if (hasPlus)                                        return "+" + digits; // explicit international
    if (digits.length === 10)                           return "+1" + digits; // US 10-digit
    if (digits.length === 11 && digits[0] === "1")     return "+" + digits;  // US with leading 1
    if (digits.length < 10)                             return "+1" + digits; // partial US
    return "+" + digits;                                // long → treat as international
}

function _groupDigits(s) {
    const chunks = [];
    let i = 0;
    while (i < s.length) {
        const rem = s.length - i;
        const size = rem <= 4 ? rem : 3;
        chunks.push(s.slice(i, i + size));
        i += size;
    }
    return chunks.join(" ");
}

/**
 * Format a stored E.164 phone number for display.
 * US/NANP (+1 + 10 digits) → "(NXX) NXX-XXXX"
 * Other international  → "+CC GROUPS"
 */
export function formatPhone(stored) {
    if (!stored) return "";
    const str = String(stored).trim();
    const digits = str.replace(/\D/g, "");
    if (!digits) return str;
    // NANP: +1 followed by exactly 10 more digits
    if (digits.length === 11 && digits[0] === "1") {
        return `(${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7, 11)}`;
    }
    // Too short to be meaningful — return as-is
    if (digits.length < 7) return "+" + digits;
    // Generic international: detect CC length then group remainder
    const ccLen = _CC3.has(digits.slice(0, 3)) ? 3 : _CC2.has(digits.slice(0, 2)) ? 2 : 1;
    return `+${digits.slice(0, ccLen)} ${_groupDigits(digits.slice(ccLen))}`;
}
