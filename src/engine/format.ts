// Number formatting and colour helpers.

export function isNum(v: unknown): v is number {
  return typeof v === "number" && Number.isFinite(v);
}

/** 0 decimals when every value is an integer, otherwise 1. */
export function autoDecimals(values: number[]): number {
  return values.every((v) => Number.isInteger(v)) ? 0 : 1;
}

export function formatNumber(v: number, decimals: number, thousands = ",", decimal = "."): string {
  const fixed = Math.abs(v).toFixed(decimals);
  const isZero = Number(fixed) === 0;
  const [int, frac] = fixed.split(".");
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, thousands);
  return (v < 0 && !isZero ? "-" : "") + grouped + (frac ? decimal + frac : "");
}

/** Normalises "#abc", "abc", "#AABBCC", "AABBCC" to "#AABBCC"; returns null if not a hex colour. */
export function normHex(c: string | null | undefined): string | null {
  if (!c) return null;
  let s = c.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map((ch) => ch + ch).join("");
  return /^[0-9a-f]{6}$/i.test(s) ? "#" + s.toUpperCase() : null;
}

/** Relative luminance (WCAG) of a hex colour, 0..1. */
export function luminance(hex: string): number {
  const h = normHex(hex) ?? "#000000";
  const ch = [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
  const lin = ch.map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
}

/** Label colour that reads best on `fill`. */
export function contrastText(fill: string, dark = "#1A1A1A", light = "#FFFFFF"): string {
  const L = luminance(fill);
  const vsLight = (1.05) / (L + 0.05);
  const vsDark = (L + 0.05) / (luminance(dark) + 0.05);
  return vsLight >= vsDark ? light : dark;
}

export const DEFAULT_PALETTE = ["#1F3A5F", "#3C6E9F", "#7FA7CF", "#B9CFE4", "#8C8C8C", "#C9C9C9"];
