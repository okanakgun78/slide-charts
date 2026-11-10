// Gantt / timeline layout.
// Rows: tasks (bars), milestones (diamonds) and group headers (bold label + summary bar over the
// rows below it, up to the next header). The timeline header stacks any combination of tiers
// (year, quarter, month, week, day); the finest tier sets the grid. "Filled" headers follow the
// user's sample: solid cells, white bold text, white separators, days as "02 / Mon".

import { fromDay, monthName, parseDate, shortDate, todayDay, toDay } from "./dates";
import type { Frame, GanttSpec, GanttTier, Layout, LinePrim, Measure, Prim, RectPrim, ShapePrim, Style, TextPrim } from "./types";

export const MILESTONE_COLOR = "#CD9700";
const TODAY_COLOR = "#C00000";
export const TIER_ORDER: GanttTier[] = ["year", "quarter", "month", "week", "day"];
const MONTH_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WEEKDAY = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

export interface Period {
  start: number;
  end: number;
  /** Single-line label candidates, longest first. */
  labels: string[];
  /** Two-line label candidates (day cells: "02" over "Mon"), longest first. */
  lines?: [string, string][];
}

/** Tiers chosen automatically from the plan's span (days). */
export function autoTiers(span: number): GanttTier[] {
  return span <= 31 ? ["month", "day"] : span <= 120 ? ["month", "week"] : span <= 450 ? ["year", "month"] : span <= 1830 ? ["year", "quarter"] : ["year"];
}

export function resolveTiers(spec: Pick<GanttSpec, "tiers">, span: number): GanttTier[] {
  const t = spec.tiers === "auto" || !spec.tiers?.length ? autoTiers(span) : spec.tiers;
  return TIER_ORDER.filter((k) => t.includes(k));
}

/** First day of the period of `tier` that contains day `d`. */
export function periodStart(d: number, tier: GanttTier): number {
  const { y, m, wd } = fromDay(d);
  switch (tier) {
    case "day": return d;
    case "week": return d - wd;
    case "month": return toDay(y, m, 1);
    case "quarter": return toDay(y, Math.floor((m - 1) / 3) * 3 + 1, 1);
    case "year": return toDay(y, 1, 1);
  }
}

/** Last day of the period of `tier` that starts at `start`. */
function periodEnd(start: number, tier: GanttTier): number {
  const { y, m } = fromDay(start);
  switch (tier) {
    case "day": return start;
    case "week": return start + 6;
    case "month": return toDay(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1;
    case "quarter": { const nm = m + 3; return toDay(nm > 12 ? y + 1 : y, nm > 12 ? nm - 12 : nm, 1) - 1; }
    case "year": return toDay(y + 1, 1, 1) - 1;
  }
}

/** ISO 8601 week number of a day. */
export function isoWeek(d: number): number {
  const thursday = d - fromDay(d).wd + 3;
  const jan1 = toDay(fromDay(thursday).y, 1, 1);
  return Math.floor((thursday - jan1) / 7) + 1;
}

/** Periods of one tier covering [t0, t1) — clipped to the range — with label candidates. */
export function tierPeriods(t0: number, t1: number, tier: GanttTier, ctx: { tiers: GanttTier[]; multiYear: boolean }): Period[] {
  const out: Period[] = [];
  const hasYear = ctx.tiers.includes("year");
  const hasMonth = ctx.tiers.includes("month");
  const hasDay = ctx.tiers.includes("day");
  for (let p = periodStart(t0, tier); p < t1; p = periodEnd(p, tier) + 1) {
    const { y, m, d, wd } = fromDay(p);
    const yy = `'${String(y).slice(2)}`;
    let labels: string[];
    let lines: [string, string][] | undefined;
    switch (tier) {
      case "year": labels = [String(y), yy]; break;
      case "quarter": { const q = `Q${Math.floor((m - 1) / 3) + 1}`; labels = hasYear ? [q] : [`${q} ${y}`, `${q} ${yy}`, q]; break; }
      case "month": {
        const full = MONTH_FULL[m - 1], short = monthName(m);
        labels = hasYear || !ctx.multiYear ? [full, short, short[0]] : [`${full} ${y}`, `${short} ${y}`, `${short} ${yy}`, short, short[0]];
        break;
      }
      // With a day row below, week numbers (W45) say more than repeating the Monday's date.
      case "week": labels = hasDay ? [`Week ${isoWeek(p)}`, `W${isoWeek(p)}`, String(isoWeek(p))] : hasMonth ? [String(d)] : [`${d} ${monthName(m)}`, String(d)]; break;
      case "day":
        lines = [[String(d).padStart(2, "0"), WEEKDAY[wd]], [String(d), WEEKDAY[wd][0]]];
        labels = [String(d).padStart(2, "0"), String(d), WEEKDAY[wd][0]];
        break;
    }
    const clipped = p < t0 || periodEnd(p, tier) > t1 - 1;
    // A period cut short by the plan's start or end gets no cryptic one-letter label.
    if (clipped && tier !== "day") labels = labels.filter((l) => l.length > 1);
    out.push({ start: Math.max(p, t0), end: Math.min(periodEnd(p, tier), t1 - 1), labels, lines });
  }
  return out;
}

export function layoutGantt(spec: GanttSpec, frame: Frame, style: Style, measure: Measure, today = todayDay()): Layout {
  const warnings: string[] = [];
  const fs = style.fontSize;
  const small = Math.round(fs * 0.85 * 2) / 2;
  const pad = fs * 0.4;
  const order = spec.dateOrder;
  const filled = spec.headerStyle !== "plain";
  const headFill = spec.headerColor ?? style.palette[0];
  const headInk = filled ? "#FFFFFF" : style.textColor;

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
  const t0 = finest === "day" ? lo : periodStart(lo, finest);
  const t1 = periodEnd(periodStart(hi, finest), finest) + 1;
  if (finest === "day" && t1 - t0 > 370) warnings.push("Day columns over more than a year are too narrow to label");
  const multiYear = fromDay(t0).y !== fromDay(t1 - 1).y;
  const tierPs = tiers.map((t) => tierPeriods(t0, t1, t, { tiers, multiYear }));

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

  // ---- Header tier heights: day cells use two lines ("02" over "Mon") when they fit ----
  const one = filled ? fs * 1.6 : fs * 1.3;
  const dayIdx = tiers.indexOf("day");
  const dayW = dayIdx >= 0 ? x(t0 + 1) - x(t0) : 0;
  const twoLine = dayIdx >= 0 && measure("Wed", fs, true) <= dayW - 2 && measure("00", fs, true) <= dayW - 2;
  const tierH = tiers.map((t) => (t === "day" && twoLine ? (filled ? fs * 2.75 : fs * 2.5) : one));
  const headerH = tierH.reduce((a, b) => a + b, 0) + (filled ? 0 : pad * 0.5);

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

  // ---- Header ----
  let ty = frame.top;
  tiers.forEach((tier, ti) => {
    const h = tierH[ti];
    const finestTier = ti === tiers.length - 1;
    tierPs[ti].forEach((p, k) => {
      const a = x(p.start), b = x(p.end + 1), w = b - a;
      if (filled) head.push({ kind: "rect", x: a, y: ty, w, h, fill: headFill, stroke: "#FFFFFF", strokeWidth: 0.75, role: `hd:${tier}:${k}` });
      else if (k > 0) grid.push({ kind: "line", x1: a, y1: ty, x2: a, y2: ty + h, color: style.gridColor, width: 0.5, role: `hdl:${tier}:${k}` });
      const size = filled || !finestTier || tiers.length === 1 ? fs : small;
      const bold = filled || !finestTier || tiers.length === 1;
      if (tier === "day" && twoLine && p.lines) {
        const pair = p.lines.find(([l1, l2]) => measure(l1, size, bold) <= w - 2 && measure(l2, size, bold) <= w - 2);
        if (pair) {
          text({ x: a, y: ty, w, h: h / 2, text: pair[0], size, bold, color: headInk, align: "center", valign: "bottom", role: `hdt:${tier}:${k}:1` });
          text({ x: a, y: ty + h / 2, w, h: h / 2, text: pair[1], size, bold, color: headInk, align: "center", valign: "top", role: `hdt:${tier}:${k}:2` });
          return;
        }
      }
      const label = p.labels.find((l) => measure(l, size, bold) <= w - 3);
      if (label) text({ x: a, y: ty, w, h, text: label, size, bold, color: headInk, align: "center", valign: "middle", role: `hdt:${tier}:${k}` });
    });
    ty += h;
  });
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
