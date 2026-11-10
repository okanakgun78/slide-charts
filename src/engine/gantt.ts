// Gantt / timeline layout.
// Rows: tasks (bars), milestones (diamonds) and group headers (bold label + summary bar over the
// rows below it, up to the next header). The timeline header stacks any combination of tiers
// (year, quarter, month, week, day); the finest tier sets the grid. "Filled" headers follow the
// user's sample: solid cells, white bold text, white separators, days as "02 / Mon".

import { fromDay, parseDate, shortDate, todayDay } from "./dates";
import { resolveTiers, timelineHeader, timelineRange } from "./timeline";
import type { Frame, GanttSpec, Layout, LinePrim, Measure, Prim, RectPrim, ShapePrim, Style, TextPrim } from "./types";

export { autoTiers, isoWeek, periodStart, resolveTiers, tierPeriods, TIER_ORDER, type Period } from "./timeline";

export const MILESTONE_COLOR = "#CD9700";
const TODAY_COLOR = "#C00000";
export function layoutGantt(spec: GanttSpec, frame: Frame, style: Style, measure: Measure, today = todayDay()): Layout {
  const warnings: string[] = [];
  const fs = style.fontSize;
  const small = Math.round(fs * 0.85 * 2) / 2;
  const pad = fs * 0.4;
  const order = spec.dateOrder;
  const filled = spec.headerStyle !== "plain";
  const headFill = spec.headerColor ?? style.palette[0];

  // ---- Parse rows ----
  const rows = spec.rows.map((r) => {
    if (r.kind === "header") return { r, s: null as number | null, e: null as number | null };
    const s = parseDate(r.start, order, "start");
    let e = r.kind === "milestone" ? s : parseDate(r.end || r.start, order, "end");
    if (s === null && (r.start.trim() || r.kind !== "task" || r.end.trim())) warnings.push(`"${r.label || "row"}": date not recognised`);
    if (s !== null && e === null) { warnings.push(`"${r.label || "row"}": end date not recognised`); e = s; }
    if (s !== null && e !== null && e < s) { warnings.push(`"${r.label}": end before start (swapped)`); e = s; }
    return { r, s, e: s === null ? null : e };
  });
  rows.forEach((row, k) => {
    if (row.r.kind !== "header") return;
    let s = Infinity, e = -Infinity;
    for (let q = k + 1; q < rows.length && rows[q].r.kind !== "header"; q++) {
      if (rows[q].s !== null) { s = Math.min(s, rows[q].s!); e = Math.max(e, rows[q].e!); }
    }
    if (Number.isFinite(s)) { row.s = s; row.e = e; }
  });
  const dated = rows.filter((x) => x.s !== null);
  if (!dated.length) return { prims: [], warnings: [...warnings, "Enter start and end dates"] };
  const lo = Math.min(...dated.map((x) => x.s!));
  const hi = Math.max(...dated.map((x) => x.e!));

  // ---- Timeline range and tiers ----
  const tiers = resolveTiers(spec, hi - lo + 1);
  const finest = tiers[tiers.length - 1];
  const { t0, t1 } = timelineRange(lo, hi, tiers);
  if (finest === "day" && t1 - t0 > 370) warnings.push("Day columns over more than a year are too narrow to label");

  // ---- Horizontal layout ----
  const grouped = rows.some((x) => x.r.kind === "header");
  const indent = grouped ? fs * 0.8 : 0;
  const labelW = Math.min(
    Math.max(...rows.map((x) => measure(x.r.label, fs, x.r.kind === "header") + (x.r.kind === "header" ? 0 : indent))) + pad,
    frame.width * 0.32,
  );
  const tlLeft = frame.left + labelW + pad;
  // Leave half a diamond of room so a milestone on the last day stays inside the frame.
  const msSize = fs * 1.15;
  const tlRight = frame.left + frame.width - (rows.some((r) => r.r.kind === "milestone") ? msSize / 2 : 0);
  const x = (d: number) => tlLeft + ((d - t0) / (t1 - t0)) * (tlRight - tlLeft);

  // ---- Header ----
  const hdr = timelineHeader({ tiers, t0, t1, x, top: frame.top, filled, fill: headFill, style, measure });
  const tierPs = hdr.periods;
  const headerH = hdr.height + (filled ? 0 : pad * 0.5);

  // ---- Vertical layout ----
  const showToday = spec.showToday && today >= t0 && today < t1;
  const footH = showToday ? small * 1.3 : 0;
  const avail = frame.height - headerH - footH;
  const n = rows.length;
  let rowH = avail / n;
  if (rowH < fs * 1.3) warnings.push("Too many rows for this height: make the chart taller or the font smaller");
  rowH = Math.min(Math.max(rowH, fs * 1.3), fs * 2.6);
  const bodyTop = frame.top + headerH;
  const bodyBottom = bodyTop + rowH * n;

  const shading: RectPrim[] = [];
  const grid: LinePrim[] = [];
  const head: RectPrim[] = [];
  const bars: (RectPrim | ShapePrim)[] = [];
  const texts: TextPrim[] = [];
  const text = (t: Omit<TextPrim, "kind" | "size" | "color"> & { size?: number; color?: string }) =>
    texts.push({ kind: "text", size: fs, color: style.textColor, ...t });

  head.push(...hdr.rects);
  grid.push(...hdr.lines);
  texts.push(...hdr.texts);
  if (!filled) grid.push({ kind: "line", x1: frame.left, y1: bodyTop, x2: tlRight, y2: bodyTop, color: style.axisColor, width: 0.75, role: "headline" });

  // ---- Body grid: finest-tier lines, stronger lines where the coarsest tier changes ----
  const coarse = new Set(tiers.length > 1 ? tierPs[0].slice(1).map((p) => p.start) : []);
  tierPs[tierPs.length - 1].forEach((p, k) => {
    if (k === 0) return;
    const a = x(p.start);
    const strong = coarse.has(p.start);
    grid.push({ kind: "line", x1: a, y1: bodyTop, x2: a, y2: bodyBottom, color: strong ? style.axisColor : style.gridColor, width: strong ? 0.75 : 0.5, role: `grid:${k}` });
  });
  grid.push({ kind: "line", x1: frame.left, y1: bodyBottom, x2: tlRight, y2: bodyBottom, color: style.axisColor, width: 0.75, role: "footline" });

  // ---- Weekend shading (day grid) ----
  if (finest === "day" && spec.shadeWeekends && style.textColor !== "#FFFFFF") {
    for (let d = t0; d < t1; d++) {
      if (fromDay(d).wd >= 5) shading.push({ kind: "rect", x: x(d), y: bodyTop, w: x(d + 1) - x(d), h: bodyBottom - bodyTop, fill: "#F2F2F2", role: `wkend:${d}` });
    }
  }

  // ---- Rows ----
  const barH = rowH * 0.5;
  rows.forEach((row, k) => {
    const y0 = bodyTop + k * rowH;
    const cy = y0 + rowH / 2;
    const { r } = row;
    const isHeader = r.kind === "header";
    if (k > 0) grid.push({ kind: "line", x1: frame.left, y1: y0, x2: tlRight, y2: y0, color: style.gridColor, width: 0.5, role: `rowline:${k}` });
    text({ x: frame.left + (isHeader ? 0 : indent), y: y0, w: labelW - (isHeader ? 0 : indent), h: rowH, text: r.label, bold: isHeader, wrap: measure(r.label, fs, isHeader) > labelW - (isHeader ? 0 : indent), align: "left", valign: "middle", role: `lab:${k}` });
    if (row.s === null || row.e === null) return;
    const xa = x(row.s), xb = x(row.e + 1);
    let endX = xb;
    if (isHeader) {
      const h = Math.max(2, rowH * 0.16);
      bars.push({ kind: "rect", x: xa, y: cy - h / 2, w: xb - xa, h, fill: r.color ?? style.textColor, role: `sum:${k}` });
    } else if (r.kind === "task") {
      bars.push({ kind: "rect", x: xa, y: cy - barH / 2, w: Math.max(xb - xa, 1), h: barH, fill: r.color ?? style.palette[0], role: `bar:${k}` });
    } else {
      const d = Math.min(rowH * 0.72, msSize);
      const xm = x(row.s + 0.5);
      bars.push({ kind: "shape", shape: "Diamond", x: xm - d / 2, y: cy - d / 2, w: d, h: d, fill: r.color ?? MILESTONE_COLOR, role: `ms:${k}` });
      endX = xm + d / 2;
    }
    // Milestone date and note to the right of the mark (or to the left if no room).
    const extra = [r.kind === "milestone" && spec.milestoneDates ? shortDate(row.s, order) : "", r.note ?? ""].filter(Boolean).join("  ");
    if (extra) {
      const w = measure(extra, small) + 2;
      const right = endX + pad + w <= tlRight;
      const startX = r.kind === "milestone" ? x(row.s + 0.5) - Math.min(rowH * 0.72, msSize) / 2 : xa;
      text({ x: right ? endX + pad : startX - pad - w, y: y0, w, h: rowH, text: extra, size: small, align: right ? "left" : "right", valign: "middle", role: `note:${k}` });
    }
  });

  // ---- Today ----
  const extras: Prim[] = [];
  if (showToday) {
    const xt = x(today + 0.5);
    extras.push({ kind: "line", x1: xt, y1: bodyTop, x2: xt, y2: bodyBottom, color: TODAY_COLOR, width: 1, dash: true, role: "today" });
    const w = measure("Today", small) + 2;
    text({ x: xt - w / 2, y: bodyBottom + 1, w, h: small * 1.25, text: "Today", size: small, color: TODAY_COLOR, align: "center", valign: "top", role: "todaylbl" });
  }

  return { prims: [...shading, ...grid, ...head, ...bars, ...extras, ...texts], warnings };
}
