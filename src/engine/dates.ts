// Calendar dates as UTC day numbers (days since 1970-01-01), so there are no time-zone shifts.
// Parsing accepts what people type or paste from Excel: ISO, 15/03/2026 or 3/15/2026 (by order
// setting, auto-resolved when a part is > 12), 15 Mar 2026, Mar 15 2026, Excel serials, and
// periods: Mar 2026, 2026-03, Q1 2026, H2 2026, 2026 (start or end of the period).

export type DateOrder = "DMY" | "MDY";

const MS_DAY = 86400000;

export function toDay(y: number, m: number, d: number): number {
  return Math.round(Date.UTC(y, m - 1, d) / MS_DAY);
}

export function fromDay(n: number): { y: number; m: number; d: number; wd: number } {
  const dt = new Date(n * MS_DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: (dt.getUTCDay() + 6) % 7 }; // wd: Mon=0
}

export function isoDay(n: number): string {
  const { y, m, d } = fromDay(n);
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

export function todayDay(): number {
  const t = new Date();
  return toDay(t.getFullYear(), t.getMonth() + 1, t.getDate());
}

const daysInMonth = (y: number, m: number) => fromDay(toDay(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1).d;

const MONTHS: Record<string, number> = {};
const addMonths = (names: string[]) => names.forEach((n, i) => n.split("|").forEach((alias) => (MONTHS[alias] = i + 1)));
addMonths(["jan|january|janv|janvier|januar|ocak|oca", "feb|february|févr|fevr|février|februar|şubat|şub|sub|subat",
  "mar|march|mars|märz|marz|mär|mart", "apr|april|avr|avril|nisan|nis", "may|mai|mayıs|mayis",
  "jun|june|juin|juni|haziran|haz", "jul|july|juil|juillet|juli|temmuz|tem", "aug|august|août|aout|ağustos|ağu|agu|agustos",
  "sep|sept|september|septembre|eylül|eyl|eylul", "oct|october|okt|oktober|octobre|ekim|eki",
  "nov|november|novembre|kasım|kas|kasim", "dec|december|dez|dezember|déc|décembre|aralık|ara|aralik"]);

const year2 = (y: number) => (y < 100 ? (y < 70 ? 2000 + y : 1900 + y) : y);

function valid(y: number, m: number, d: number): number | null {
  if (m < 1 || m > 12 || d < 1 || d > daysInMonth(y, m) || y < 1900 || y > 2200) return null;
  return toDay(y, m, d);
}

/**
 * Parses a date. `edge` picks the first or last day when the text names a whole period
 * ("Q1 2026" as an end date means 31 Mar 2026). Returns a day number or null.
 */
export function parseDate(raw: string, order: DateOrder = "DMY", edge: "start" | "end" = "start"): number | null {
  const s = raw.trim().toLowerCase().replace(/\s+/g, " ").replace(/(\d)(st|nd|rd|th)\b/g, "$1");
  if (!s) return null;
  let m: RegExpMatchArray | null;

  // Excel serial number (1900 date system): 20000..80000 covers 1954..2119.
  if ((m = s.match(/^(\d{5})(\.\d+)?$/))) {
    const serial = Number(m[1]);
    return serial >= 20000 && serial <= 80000 ? serial - 25569 : null;
  }
  // ISO yyyy-mm-dd (optionally with time).
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[t ].*)?$/))) return valid(+m[1], +m[2], +m[3]);
  // yyyy/mm/dd or yyyy.mm.dd
  if ((m = s.match(/^(\d{4})[/.](\d{1,2})[/.](\d{1,2})$/))) return valid(+m[1], +m[2], +m[3]);
  // a/b/c with / . or -
  if ((m = s.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$/))) {
    const a = +m[1], b = +m[2], y = year2(+m[3]);
    let d: number, mo: number;
    if (a > 12) [d, mo] = [a, b];
    else if (b > 12) [mo, d] = [a, b];
    else [d, mo] = order === "DMY" ? [a, b] : [b, a];
    return valid(y, mo, d);
  }
  // 15 Mar 2026 / 15-Mar-26 / 15. März 2026
  if ((m = s.match(/^(\d{1,2})[.\- ]+([a-zçğıöşüéûä]+)\.?[\- ,]+(\d{2,4})$/))) {
    const mo = MONTHS[m[2]];
    return mo ? valid(year2(+m[3]), mo, +m[1]) : null;
  }
  // Mar 15, 2026 / March 15 2026
  if ((m = s.match(/^([a-zçğıöşüéûä]+)\.? (\d{1,2}),? (\d{2,4})$/))) {
    const mo = MONTHS[m[1]];
    return mo ? valid(year2(+m[3]), mo, +m[2]) : null;
  }
  // Periods: Mar 2026 / March-26 / 2026-03 / 03/2026
  let y = 0, mo = 0;
  if ((m = s.match(/^([a-zçğıöşüéûä]+)\.?[\- ']+(\d{2,4})$/)) && MONTHS[m[1]]) { mo = MONTHS[m[1]]; y = year2(+m[2]); }
  else if ((m = s.match(/^(\d{4})-(\d{1,2})$/))) { y = +m[1]; mo = +m[2]; }
  else if ((m = s.match(/^(\d{1,2})[/.](\d{4})$/))) { mo = +m[1]; y = +m[2]; }
  if (y && mo >= 1 && mo <= 12) return edge === "start" ? toDay(y, mo, 1) : toDay(y, mo, daysInMonth(y, mo));
  // Quarters and halves: Q1 2026 / Q1-26 / H2 2026 / 2026 Q1 / 2026-H2
  let kind = "", k = 0;
  if ((m = s.match(/^([qh])([1-4]) ?[\-'/]? ?(\d{2,4})$/))) { kind = m[1]; k = +m[2]; y = year2(+m[3]); }
  else if ((m = s.match(/^(\d{4}) ?[\-/]? ?([qh])([1-4])$/))) { y = +m[1]; kind = m[2]; k = +m[3]; }
  if (kind) {
    if (kind === "h" && k > 2) return null;
    const len = kind === "q" ? 3 : 6;
    const m0 = (k - 1) * len + 1;
    return edge === "start" ? toDay(y, m0, 1) : toDay(y, m0 + len - 1, daysInMonth(y, m0 + len - 1));
  }
  // A bare year.
  if ((m = s.match(/^(\d{4})$/))) return edge === "start" ? toDay(+m[1], 1, 1) : toDay(+m[1], 12, 31);
  return null;
}

const SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export function monthName(m: number): string {
  return SHORT[m - 1];
}

/** "15 Mar" (DMY) or "Mar 15" (MDY). */
export function shortDate(n: number, order: DateOrder): string {
  const { m, d } = fromDay(n);
  return order === "DMY" ? `${d} ${SHORT[m - 1]}` : `${SHORT[m - 1]} ${d}`;
}

/** Default order from the browser locale: US-style languages use month first. */
export function localeOrder(): DateOrder {
  const lang = (typeof navigator !== "undefined" && navigator.language) || "en-GB";
  return /^en-(us|ph|ca)$/i.test(lang) || /^(fil)/i.test(lang) ? "MDY" : "DMY";
}
