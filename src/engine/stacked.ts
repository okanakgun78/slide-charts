// Stacked column chart layout (pilot scope: non-negative values).
// Consulting conventions: no value axis, segment labels inside, totals on top,
// series labels to the right of the last column instead of a legend box.

import { autoDecimals, contrastText, formatNumber, isNum } from "./format";
import type { Frame, Layout, Measure, Prim, StackedSpec, Style } from "./types";

const GAP_RATIO = 0.38; // share of each category slot left empty between columns

export function layoutStacked(spec: StackedSpec, frame: Frame, style: Style, measure: Measure): Layout {
  const prims: Prim[] = [];
  const warnings: string[] = [];
  const n = spec.categories.length;
  const m = spec.series.length;
  if (n === 0 || m === 0) return { prims, warnings: ["No data"] };

  const fs = style.fontSize;
  const lineH = fs * 1.3;
  const pad = fs * 0.4;

  const all: number[] = [];
  spec.series.forEach((s) =>
    s.values.forEach((v) => {
      if (isNum(v)) all.push(v);
    }),
  );
  if (all.some((v) => v < 0)) warnings.push("Negative values are not drawn yet (pilot)");
  const dec = spec.decimals ?? autoDecimals(all);
  const fmt = (v: number) => formatNumber(v, dec);

  const valueAt = (j: number, i: number): number => {
    const v = spec.series[j].values[i];
    return isNum(v) && v > 0 ? v : 0;
  };
  const totals = spec.categories.map((_, i) => spec.series.reduce((acc, _s, j) => acc + valueAt(j, i), 0));
  const maxTotal = Math.max(...totals) || 1;

  // Horizontal budget: plot area + series labels on the right.
  const seriesLabelW = Math.max(...spec.series.map((s) => measure(s.name, fs))) + pad;
  const plotLeft = frame.left;
  const plotRight = frame.left + frame.width - seriesLabelW - pad;
  const slot = (plotRight - plotLeft) / n;
  const colW = slot * (1 - GAP_RATIO);

  // Vertical budget: totals above, category labels below.
  const plotTop = frame.top + (spec.showTotals ? lineH : 0);
  const catLabelTop = frame.top + frame.height - lineH;
  const baseline = catLabelTop - pad * 0.5;
  const scale = (baseline - plotTop) / maxTotal;
  if (scale <= 0) return { prims, warnings: ["Chart frame too small"] };

  const lastSegCenters: (number | null)[] = new Array(m).fill(null);

  for (let i = 0; i < n; i++) {
    const x = plotLeft + i * slot + (slot - colW) / 2;
    let y = baseline;
    for (let j = 0; j < m; j++) {
      const v = valueAt(j, i);
      if (v <= 0) continue;
      const h = v * scale;
      y -= h;
      const fill = style.palette[j % style.palette.length];
      prims.push({ kind: "rect", x, y, w: colW, h, fill, role: `seg:${j}:${i}` });
      const label = fmt(v);
      if (h >= fs * 1.05 && measure(label, fs) <= colW - 2) {
        prims.push({
          kind: "text", x, y, w: colW, h, text: label, size: fs, color: contrastText(fill),
          align: "center", valign: "middle", role: `lbl:${j}:${i}`,
        });
      }
      if (i === n - 1) lastSegCenters[j] = y + h / 2;
    }
    if (spec.showTotals) {
      prims.push({
        kind: "text", x: x - slot * GAP_RATIO / 2, y: y - lineH, w: slot, h: lineH, text: fmt(totals[i]),
        size: fs, color: style.textColor, bold: true, align: "center", valign: "bottom", role: `tot:${i}`,
      });
    }
    prims.push({
      kind: "text", x: plotLeft + i * slot, y: catLabelTop, w: slot, h: lineH, text: spec.categories[i],
      size: fs, color: style.textColor, align: "center", valign: "top", role: `cat:${i}`,
    });
  }

  prims.push({
    kind: "line", x1: plotLeft, y1: baseline, x2: plotRight, y2: baseline,
    color: style.axisColor, width: 0.75, role: "axis",
  });

  // Series labels: aligned with the last column's segments, de-overlapped.
  const desired = spec.series.map((_, j) => lastSegCenters[j] ?? baseline - (j + 0.5) * lineH);
  const ys = placeLabels(desired, lineH, plotTop, baseline);
  const lx = plotRight + pad;
  spec.series.forEach((s, j) => {
    prims.push({
      kind: "text", x: lx, y: ys[j] - lineH / 2, w: seriesLabelW, h: lineH, text: s.name,
      size: fs, color: style.textColor, align: "left", valign: "middle", role: `ser:${j}`,
    });
  });

  return { prims, warnings };
}

/**
 * Places label centres as close as possible to `desired` while keeping at least `minGap`
 * between neighbours and staying inside [lo, hi]. Returns centres in the input order.
 */
export function placeLabels(desired: number[], minGap: number, lo: number, hi: number): number[] {
  const order = desired.map((y, k) => ({ y, k })).sort((a, b) => a.y - b.y);
  const ys = order.map((o) => o.y);
  // Push down to remove overlaps.
  for (let k = 1; k < ys.length; k++) ys[k] = Math.max(ys[k], ys[k - 1] + minGap);
  // If the stack overflows the bottom, shift it up as a block, but not above `lo`.
  const overflow = ys[ys.length - 1] - (hi - minGap / 2);
  if (overflow > 0) {
    for (let k = 0; k < ys.length; k++) ys[k] -= overflow;
    for (let k = 0; k < ys.length; k++) ys[k] = Math.max(ys[k], lo + minGap / 2 + k * minGap);
  }
  const out = new Array<number>(desired.length);
  order.forEach((o, idx) => (out[o.k] = ys[idx]));
  return out;
}

/** Bounding box of a primitive list (used to detect move/resize of a drawn chart). */
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
