// Category charts: column/bar (stacked, clustered, 100%), line, and column+line combination.
// Consulting conventions: no value axis; labels inside segments; totals at the end of stacks;
// series names right of the last column (or a legend on top); first data row = top segment.

import { autoDecimals, contrastText, formatNumber, isNum } from "./format";
import { fitFontSize, placeLabels, wrapLines } from "./geometry";
import type { CategorySpec, Frame, Layout, LinePrim, Measure, Mode, Prim, RectPrim, ShapePrim, Style, TextPrim } from "./types";

const GAP = 0.38; // share of each category slot left empty between columns/bars

interface Seg {
  j: number;
  i: number;
  v0: number;
  v1: number;
  /** Value shown in the label (raw value, or share in 100% charts). */
  shown: number;
}

export function layoutCategory(spec: CategorySpec, frame: Frame, style: Style, measure: Measure): Layout {
  const warnings: string[] = [];
  const n = spec.categories.length;
  const m = spec.series.length;
  if (n === 0 || m === 0) return { prims: [], warnings: ["No data"] };

  const horizontal = spec.type === "bar";
  const fs = style.fontSize;
  const lineH = fs * 1.3;
  const pad = fs * 0.4;

  const isLine = (j: number) => spec.type === "line" || (spec.type === "combo" && !!spec.series[j].asLine);
  const barIdx = spec.series.map((_, j) => j).filter((j) => !isLine(j));
  const lineIdx = spec.series.map((_, j) => j).filter((j) => isLine(j));
  if (horizontal && lineIdx.length) warnings.push("Line series are not shown in bar charts");

  let mode: Mode = spec.mode;
  if (spec.type === "combo" && mode === "percent") mode = "stacked";
  if (barIdx.length === 1 && mode === "stacked") mode = "clustered"; // one series: label at the end

  const val = (j: number, i: number): number | null => {
    const v = spec.series[j].values[i];
    return isNum(v) ? v : null;
  };
  const all: number[] = [];
  spec.series.forEach((_, j) => spec.categories.forEach((_c, i) => { const v = val(j, i); if (v !== null) all.push(v); }));
  const dec = spec.decimals ?? autoDecimals(all);
  const fmt = (v: number) => formatNumber(v, dec, style.thousandsSep, style.decimalSep);
  const fmtPct = (v: number) => `${formatNumber(v, 0, style.thousandsSep, style.decimalSep)}%`;

  const barColor = (j: number) => spec.series[j].color ?? style.palette[barIdx.indexOf(j) % style.palette.length];
  const lineColor = (j: number) => spec.series[j].color ?? style.linePalette[lineIdx.indexOf(j) % style.linePalette.length];

  // ---- Segments in value space ----
  const segs: Seg[] = [];
  const stackTop = new Array<number>(n).fill(0);
  const stackBottom = new Array<number>(n).fill(0);
  const totals = new Array<number>(n).fill(0);
  for (let i = 0; i < n; i++) {
    if (mode === "clustered") {
      for (const j of barIdx) {
        const v = val(j, i);
        if (v !== null) segs.push({ j, i, v0: 0, v1: v, shown: v });
      }
      continue;
    }
    const sumAbs = barIdx.reduce((a, j) => a + Math.abs(val(j, i) ?? 0), 0);
    let up = 0, down = 0;
    // Columns: stack bottom-up in reverse sheet order so the first row ends up on top.
    // Bars: first row closest to the axis (left), matching reading order.
    for (const j of horizontal ? barIdx : [...barIdx].reverse()) {
      const raw = val(j, i);
      if (raw === null || raw === 0) continue;
      totals[i] += raw;
      const v = mode === "percent" ? (sumAbs ? (raw / sumAbs) * 100 : 0) : raw;
      if (v >= 0) { segs.push({ j, i, v0: up, v1: up + v, shown: v }); up += v; }
      else { segs.push({ j, i, v0: down, v1: down + v, shown: v }); down += v; }
    }
    stackTop[i] = up;
    stackBottom[i] = down;
  }

  // ---- Value domain ----
  let lo = 0, hi = 0;
  for (const s of segs) { lo = Math.min(lo, s.v0, s.v1); hi = Math.max(hi, s.v0, s.v1); }
  for (const j of lineIdx) for (let i = 0; i < n; i++) { const v = val(j, i); if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); } }
  if (hi - lo < 1e-9) hi = lo + 1;

  // ---- What goes where ----
  const legend: "right" | "top" | "none" =
    spec.labels.legend !== "auto" ? spec.labels.legend
      : m === 1 ? "none"
        : !horizontal && (mode !== "clustered" || barIdx.length === 0) ? "right" : "top";
  const showVals = spec.labels.values;
  const showTotals = spec.labels.totals && mode !== "clustered" && barIdx.length > 0;
  const outsideBarLabels = (mode === "clustered" && showVals) || showTotals;
  const lineLabels = showVals && lineIdx.length > 0;
  const hasNeg = lo < 0;

  // Legend on top.
  type LegendItem = { j: number; w: number };
  const sw = fs * 0.8;
  const items: LegendItem[] = spec.series.map((s, j) => ({ j, w: sw + fs * 0.35 + measure(s.name, fs) + fs * 1.2 }));
  const legendRows: LegendItem[][] = [];
  if (legend === "top") {
    let row: LegendItem[] = [], used = 0;
    for (const it of items) {
      if (row.length && used + it.w > frame.width) { legendRows.push(row); row = []; used = 0; }
      row.push(it); used += it.w;
    }
    if (row.length) legendRows.push(row);
  }
  const legendH = legendRows.length ? legendRows.length * lineH + pad : 0;
  const rightW = legend === "right" ? Math.min(Math.max(...spec.series.map((s) => measure(s.name, fs))) + pad, frame.width * 0.35) : 0;

  const textPrims: TextPrim[] = [];
  const rects: RectPrim[] = [];
  const lines: LinePrim[] = [];
  const marks: ShapePrim[] = [];
  const text = (t: Omit<TextPrim, "kind" | "size" | "color"> & { size?: number; color?: string }) =>
    textPrims.push({ kind: "text", size: fs, color: style.textColor, ...t });

  // Labels above/below stacks and points, de-overlapped per category at the end (columns only).
  type Float = { text: string; y: number; up: boolean; bold?: boolean; role: string };
  const floats: Float[][] = Array.from({ length: n }, () => []);
  let floatLo = frame.top + legendH;
  let floatHi = frame.top + frame.height;

  // ---- Geometry ----
  let pos: (v: number) => number;
  let slotStart: (i: number) => number;
  let slot: number;
  let plotA: number, plotB: number; // extent along the category direction
  let catLabel: (i: number) => void;

  if (!horizontal) {
    const plotLeft = frame.left;
    const plotRight = frame.left + frame.width - rightW - (rightW ? pad : 0);
    slot = (plotRight - plotLeft) / n;
    const wrapped = spec.categories.map((c) => wrapLines(c, slot * 0.96, fs, measure));
    const catLines = Math.min(3, Math.max(...wrapped.map((w) => w.length)));
    const catH = catLines * lineH;
    const topRoom = lineLabels && showTotals ? lineH * 2.1 : lineLabels ? lineH * 1.2 : outsideBarLabels ? lineH : 0;
    const top = frame.top + legendH + topRoom;
    const bottom = frame.top + frame.height - catH - pad * 0.5 - (hasNeg && (outsideBarLabels || lineLabels) ? topRoom || lineH : 0);
    const scale = (bottom - top) / (hi - lo);
    if (scale <= 0) return { prims: [], warnings: ["Chart frame too small"] };
    pos = (v) => top + (hi - v) * scale;
    slotStart = (i) => plotLeft + i * slot;
    plotA = plotLeft; plotB = plotRight;
    const catTop = frame.top + frame.height - catH;
    floatHi = catTop;
    floatLo = frame.top + legendH;
    catLabel = (i) => text({ x: slotStart(i), y: catTop, w: slot, h: catH, text: spec.categories[i], wrap: wrapped[i].length > 1, align: "center", valign: "top", role: `cat:${i}` });
  } else {
    const top = frame.top + legendH;
    const bottom = frame.top + frame.height;
    slot = (bottom - top) / n;
    const catW = Math.min(Math.max(...spec.categories.map((c) => measure(c, fs))) + 2, frame.width * 0.3);
    const outsideTexts = segs.map((s) => (mode === "clustered" ? fmt(s.shown) : "")).concat(showTotals ? totals.map(fmt) : []);
    const outW = outsideBarLabels ? Math.max(0, ...outsideTexts.map((t) => measure(t, fs, true))) + pad + 2 : 0;
    const plotLeft = frame.left + catW + pad + (hasNeg ? outW : 0);
    const plotRight = frame.left + frame.width - outW;
    const scale = (plotRight - plotLeft) / (hi - lo);
    if (scale <= 0) return { prims: [], warnings: ["Chart frame too small"] };
    pos = (v) => plotLeft + (v - lo) * scale;
    slotStart = (i) => top + i * slot;
    plotA = top; plotB = bottom;
    catLabel = (i) => text({ x: frame.left, y: slotStart(i), w: catW, h: slot, text: spec.categories[i], wrap: measure(spec.categories[i], fs) > catW, align: "right", valign: "middle", role: `cat:${i}` });
  }

  const zero = pos(0);
  const groupW = slot * (1 - GAP);
  const offset = (slot - groupW) / 2;
  const rectFor = (bandStart: number, bandSize: number, a: number, b: number) => {
    const pa = pos(a), pb = pos(b);
    return horizontal
      ? { x: Math.min(pa, pb), y: bandStart, w: Math.abs(pa - pb), h: bandSize }
      : { x: bandStart, y: Math.min(pa, pb), w: bandSize, h: Math.abs(pa - pb) };
  };
  const band = (s: Seg) => {
    if (mode !== "clustered") return { start: slotStart(s.i) + offset, size: groupW };
    const k = barIdx.indexOf(s.j);
    const size = groupW / barIdx.length;
    return { start: slotStart(s.i) + offset + k * size, size };
  };

  // Font size for clustered outside labels (must fit the bar width in columns).
  let outFs = fs;
  if (mode === "clustered" && showVals && segs.length) {
    const size = groupW / barIdx.length;
    outFs = horizontal ? Math.max(7, Math.min(fs, size / 1.1)) : fitFontSize(segs.map((s) => fmt(s.shown)), size * 1.1, fs, measure);
  }

  // ---- Bars ----
  for (const s of segs) {
    const b = band(s);
    const r = rectFor(b.start, b.size, s.v0, s.v1);
    const fill = barColor(s.j);
    rects.push({ kind: "rect", ...r, fill, role: `seg:${s.j}:${s.i}` });
    if (!showVals) continue;
    const label = mode === "percent" ? fmtPct(s.shown) : fmt(s.shown);
    if (mode === "clustered") {
      const neg = s.v1 < 0;
      const w = measure(label, outFs);
      if (!horizontal) {
        const lh = outFs * 1.3;
        const extra = Math.max(0, (w - b.size) / 2 + 2);
        text({ x: b.start - extra, y: neg ? r.y + r.h : r.y - lh, w: b.size + 2 * extra, h: lh, text: label, size: outFs, align: "center", valign: neg ? "top" : "bottom", role: `lbl:${s.j}:${s.i}` });
      } else {
        const x = neg ? r.x - pad - w - 2 : r.x + r.w + pad;
        text({ x, y: b.start, w: w + 2, h: b.size, text: label, size: outFs, align: neg ? "right" : "left", valign: "middle", role: `lbl:${s.j}:${s.i}` });
      }
    } else {
      const thick = horizontal ? r.h : r.w;
      const len = horizontal ? r.w : r.h;
      const fits = horizontal ? thick >= fs * 1.05 && measure(label, fs) <= len - 4 : len >= fs * 1.05 && measure(label, fs) <= thick - 2;
      if (fits) text({ ...r, text: label, color: contrastText(fill), align: "center", valign: "middle", role: `lbl:${s.j}:${s.i}` });
    }
  }

  // ---- Totals ----
  if (showTotals) {
    for (let i = 0; i < n; i++) {
      if (!segs.some((s) => s.i === i)) continue;
      const atTop = stackTop[i] > 0 || stackBottom[i] === 0;
      const end = atTop ? stackTop[i] : stackBottom[i];
      const label = fmt(totals[i]);
      if (!horizontal) {
        floats[i].push({ text: label, y: atTop ? pos(end) - lineH / 2 : pos(end) + lineH / 2, up: atTop, bold: true, role: `tot:${i}` });
      } else {
        const w = measure(label, fs, true) + 2;
        const x = atTop ? pos(end) + pad : pos(end) - pad - w;
        text({ x, y: slotStart(i), w, h: slot, text: label, bold: true, align: atTop ? "left" : "right", valign: "middle", role: `tot:${i}` });
      }
    }
  }

  // ---- Category axis ----
  lines.push(horizontal
    ? { kind: "line", x1: zero, y1: plotA, x2: zero, y2: plotB, color: style.axisColor, width: 0.75, role: "axis" }
    : { kind: "line", x1: plotA, y1: zero, x2: plotB, y2: zero, color: style.axisColor, width: 0.75, role: "axis" });
  for (let i = 0; i < n; i++) catLabel(i);

  // ---- Lines (vertical charts only) ----
  const lastPointY = new Map<number, number>();
  if (!horizontal && lineIdx.length) {
    const cx = (i: number) => slotStart(i) + slot / 2;
    const d = fs * 0.42;
    const perCat: { j: number; y: number; v: number }[][] = Array.from({ length: n }, () => []);
    for (const j of lineIdx) {
      const color = lineColor(j);
      let prev: { x: number; y: number } | null = null;
      for (let i = 0; i < n; i++) {
        const v = val(j, i);
        if (v === null) { prev = null; continue; }
        const p = { x: cx(i), y: pos(v) };
        if (prev) lines.push({ kind: "line", x1: prev.x, y1: prev.y, x2: p.x, y2: p.y, color, width: Math.max(1.25, fs * 0.15), role: `line:${j}:${i}` });
        marks.push({ kind: "shape", shape: "Ellipse", x: p.x - d / 2, y: p.y - d / 2, w: d, h: d, fill: color, role: `mk:${j}:${i}` });
        perCat[i].push({ j, y: p.y, v });
        lastPointY.set(j, p.y);
        prev = p;
      }
    }
    if (lineLabels) {
      const off = lineH * 0.62;
      const yAt = (j: number, i: number) => { const v = i >= 0 && i < n ? val(j, i) : null; return v === null ? null : pos(v); };
      for (let i = 0; i < n; i++) {
        const pts = [...perCat[i]].sort((a, b) => a.y - b.y);
        pts.forEach((p, k) => {
          let below: boolean;
          if (pts.length === 1) {
            // Single line here: put the label on the open side (below at a dip, above at a peak).
            const nb = [yAt(p.j, i - 1), yAt(p.j, i + 1)].filter((y): y is number => y !== null);
            below = nb.length > 0 && p.y > nb.reduce((a, b) => a + b, 0) / nb.length + 0.5;
          } else {
            below = k === pts.length - 1;
          }
          floats[i].push({ text: fmt(p.v), y: below ? p.y + off : p.y - off, up: !below, role: `plbl:${p.j}:${i}` });
        });
      }
    }
  }

  // ---- Floating labels (columns): totals and line labels, without overlaps ----
  // Labels above things are pushed further up, labels below things further down, so they never
  // slide into the bar or point they describe. A label that still lands on a bar gets contrasting ink.
  if (!horizontal) {
    const gap = lineH * 0.95;
    for (let i = 0; i < n; i++) {
      const ups = floats[i].filter((f) => f.up).sort((a, b) => b.y - a.y);
      for (let k = 1; k < ups.length; k++) ups[k].y = Math.min(ups[k].y, ups[k - 1].y - gap);
      const over = floatLo + gap / 2 - (ups.length ? ups[ups.length - 1].y : Infinity);
      if (over > 0) ups.forEach((f) => (f.y += over));
      const downs = floats[i].filter((f) => !f.up).sort((a, b) => a.y - b.y);
      for (let k = 1; k < downs.length; k++) downs[k].y = Math.max(downs[k].y, downs[k - 1].y + gap);
      const under = (downs.length ? downs[downs.length - 1].y : -Infinity) - (floatHi - gap / 2);
      if (under > 0) downs.forEach((f) => (f.y -= under));
      for (const f of [...ups, ...downs]) {
        const cx = slotStart(i) + slot / 2;
        const hit = rects.find((r) => r.role.startsWith("seg:") && cx >= r.x && cx <= r.x + r.w && f.y >= r.y && f.y <= r.y + r.h);
        text({ x: slotStart(i), y: f.y - lineH / 2, w: slot, h: lineH, text: f.text, bold: f.bold, color: hit ? contrastText(hit.fill) : style.textColor, align: "center", valign: "middle", role: f.role });
      }
    }
  }

  // ---- Series names ----
  if (legend === "right") {
    const desired = spec.series.map((_, j) => {
      if (isLine(j)) return lastPointY.get(j) ?? (pos(hi) + pos(lo)) / 2;
      const last = segs.filter((s) => s.j === j).sort((a, b) => b.i - a.i)[0];
      if (!last) return (pos(hi) + pos(lo)) / 2;
      return (pos(last.v0) + pos(last.v1)) / 2;
    });
    const ys = placeLabels(desired, lineH, frame.top + legendH, frame.top + frame.height - lineH);
    spec.series.forEach((s, j) =>
      text({ x: plotB + pad, y: ys[j] - lineH / 2, w: rightW, h: lineH, text: s.name, align: "left", valign: "middle", role: `ser:${j}` }),
    );
  } else if (legend === "top") {
    legendRows.forEach((row, r) => {
      let x = frame.left;
      const y = frame.top + r * lineH;
      for (const it of row) {
        const j = it.j;
        if (isLine(j)) {
          lines.push({ kind: "line", x1: x, y1: y + lineH / 2, x2: x + sw, y2: y + lineH / 2, color: lineColor(j), width: Math.max(1.25, fs * 0.15), role: `legsw:${j}` });
        } else {
          rects.push({ kind: "rect", x, y: y + (lineH - sw) / 2, w: sw, h: sw, fill: barColor(j), role: `legsw:${j}` });
        }
        const tx = x + sw + fs * 0.35;
        text({ x: tx, y, w: it.w - sw - fs * 0.35, h: lineH, text: spec.series[j].name, align: "left", valign: "middle", role: `leg:${j}` });
        x += it.w;
      }
    });
  }

  const prims: Prim[] = [...rects, ...lines, ...marks, ...textPrims];
  return { prims, warnings };
}
