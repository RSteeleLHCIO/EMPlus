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
