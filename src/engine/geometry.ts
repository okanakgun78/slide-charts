// Small geometry and text helpers shared by the layout engines.

import type { Frame, Measure, Prim } from "./types";

/**
 * Places label centres as close as possible to `desired` while keeping at least `minGap`
 * between neighbours and staying inside [lo, hi]. Returns centres in the input order.
 */
export function placeLabels(desired: number[], minGap: number, lo: number, hi: number): number[] {
  if (desired.length === 0) return [];
  const clamp = (y: number) => Math.min(Math.max(y, lo + minGap / 2), hi - minGap / 2);
  const order = desired.map((y, k) => ({ y: clamp(y), k })).sort((a, b) => a.y - b.y);
  const ys = order.map((o) => o.y);
  for (let k = 1; k < ys.length; k++) ys[k] = Math.max(ys[k], ys[k - 1] + minGap);
  const overflow = ys[ys.length - 1] - (hi - minGap / 2);
  if (overflow > 0) {
    for (let k = 0; k < ys.length; k++) ys[k] -= overflow;
    for (let k = 0; k < ys.length; k++) ys[k] = Math.max(ys[k], lo + minGap / 2 + k * minGap);
  }
  const out = new Array<number>(desired.length);
  order.forEach((o, idx) => (out[o.k] = ys[idx]));
  return out;
}

/** Bounding box of a primitive list. */
export function bbox(prims: Prim[]): Frame {
  let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
  for (const p of prims) {
    const [x1, y1, x2, y2] =
      p.kind === "line"
        ? [Math.min(p.x1, p.x2), Math.min(p.y1, p.y2), Math.max(p.x1, p.x2), Math.max(p.y1, p.y2)]
        : [p.x, p.y, p.x + p.w, p.y + p.h];
    l = Math.min(l, x1); t = Math.min(t, y1); r = Math.max(r, x2); b = Math.max(b, y2);
  }
  return { left: l, top: t, width: r - l, height: b - t };
}

/** Greedy word wrap. A single word wider than `maxW` stays on its own line. */
export function wrapLines(text: string, maxW: number, size: number, measure: Measure, bold = false): string[] {
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length === 0) return [""];
  const lines: string[] = [];
  let cur = words[0];
  for (let k = 1; k < words.length; k++) {
    const next = `${cur} ${words[k]}`;
    if (measure(next, size, bold) <= maxW) cur = next;
    else {
      lines.push(cur);
      cur = words[k];
    }
  }
  lines.push(cur);
  return lines;
}

/** Largest font size <= `size` (and >= `min`) at which every text fits in `maxW`. */
export function fitFontSize(texts: string[], maxW: number, size: number, measure: Measure, min = 7): number {
  let s = size;
  while (s > min && texts.some((t) => measure(t, s) > maxW)) s -= 0.5;
  return s;
}
