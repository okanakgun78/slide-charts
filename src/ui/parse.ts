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

/**
 * Splits clipboard text from Excel into rows of cells. Tab separated; a cell containing line breaks
 * or tabs arrives quoted ("a""b" for a quote), as Excel writes it.
 */
export function parseTSV(text: string): string[][] {
  const src = text.replace(/\r\n?/g, "\n");
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let i = 0;
  let atStart = true;
  while (i < src.length) {
    const ch = src[i];
    if (atStart && ch === '"') {
      // Quoted cell: read to the closing quote.
      let j = i + 1;
      let v = "";
      for (;;) {
        if (j >= src.length) break;
        if (src[j] === '"') {
          if (src[j + 1] === '"') { v += '"'; j += 2; continue; }
          j++;
          break;
        }
        v += src[j++];
      }
      // Only treat it as quoted if the cell ends right after the closing quote.
      if (j >= src.length || src[j] === "\t" || src[j] === "\n") { cell = v; i = j; atStart = false; continue; }
    }
    if (ch === "\t") { row.push(cell); cell = ""; atStart = true; i++; continue; }
    if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; atStart = true; i++; continue; }
    cell += ch;
    atStart = false;
    i++;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

/** Formats a number for an editable cell (no thousands separators). */
export function cellText(v: number | null, decimalSep = "."): string {
  if (v === null || !Number.isFinite(v)) return "";
  const s = String(Math.round(v * 1e9) / 1e9);
  return decimalSep === "," ? s.replace(".", ",") : s;
}
