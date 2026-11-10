// Parsing of pasted / typed data.

/**
 * Parses a number as typed or pasted from Excel. Handles thousands separators, currency
 * symbols, %, unicode minus and accounting negatives "(12)". Returns null for blanks/garbage.
 */
export function parseNumber(raw: string, decimalSep = "."): number | null {
  let s = raw.trim();
  if (!s || s === "-" || s === "–") return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[−–]/g, "-").replace(/[%$€£¥\s  ']/g, "");
  if (decimalSep === ",") s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
  const v = parseFloat(s);
  return Number.isFinite(v) ? (neg ? -v : v) : null;
}

/** Splits clipboard text from Excel into rows of cells (tab separated, CRLF tolerant). */
export function parseTSV(text: string): string[][] {
  const rows = text.replace(/\r\n?/g, "\n").replace(/\n$/, "").split("\n");
  return rows.map((r) => r.split("\t"));
}

/** Formats a number for an editable cell (no thousands separators). */
export function cellText(v: number | null, decimalSep = "."): string {
  if (v === null || !Number.isFinite(v)) return "";
  const s = String(Math.round(v * 1e9) / 1e9);
  return decimalSep === "," ? s.replace(".", ",") : s;
}
