// Calendar view: the timeline header (any tiers, sample style) over a grid of free-text cells —
// no bars. Rows have their own height, style, fill and alignment (as in the user's sample:
// a tall plan row, an italic row, a bold yellow totals row). Cells are keyed by column start date.

import { fromDay, isoDay, monthName, parseDate } from "./dates";
import { wrapLines } from "./geometry";
import { isoWeek, timelineHeader, timelineRange, TIER_ORDER, type Period } from "./timeline";
import type { CalendarSpec, Frame, GanttTier, Layout, LinePrim, Measure, Prim, RectPrim, Style, TextPrim } from "./types";

const WEEKDAY = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

/** Calendar columns need room for text, so days are chosen only for short ranges. */
export function autoCalendarTiers(span: number): GanttTier[] {
  return span <= 14 ? ["month", "day"] : span <= 70 ? ["month", "week"] : span <= 500 ? ["year", "month"] : span <= 1830 ? ["year", "quarter"] : ["year"];
}

export interface CalendarColumn {
  key: string; // ISO date of the column's first day
  label: string; // for the data sheet
  start: number;
  end: number;
}

/** Resolves the date range, tiers and columns of a calendar (null if the dates are unreadable). */
export function calendarColumns(spec: CalendarSpec): { tiers: GanttTier[]; t0: number; t1: number; columns: CalendarColumn[] } | null {
  let lo = parseDate(spec.start, spec.dateOrder, "start");
  let hi = parseDate(spec.end || spec.start, spec.dateOrder, "end");
  if (lo === null || hi === null) return null;
  if (hi < lo) [lo, hi] = [hi, lo];
  const span = hi - lo + 1;
  const tiers = spec.tiers === "auto" || !spec.tiers.length ? autoCalendarTiers(span) : TIER_ORDER.filter((t) => (spec.tiers as GanttTier[]).includes(t));
  const { t0, t1 } = timelineRange(lo, hi, tiers);
  const finest = tiers[tiers.length - 1];
  const cols: CalendarColumn[] = [];
  for (let d = t0; d < t1; ) {
    const { y, m, d: dd, wd } = fromDay(d);
    let end: number;
    let label: string;
    switch (finest) {
      case "day": end = d; label = `${WEEKDAY[wd]} ${String(dd).padStart(2, "0")} ${monthName(m)}`; break;
      case "week": end = Math.min(d - wd + 6, t1 - 1); label = `W${isoWeek(d)} · ${dd} ${monthName(m)}`; break;
      case "month": end = Math.min(isoEndOfMonth(y, m), t1 - 1); label = `${monthName(m)} ${y}`; break;
      case "quarter": end = Math.min(isoEndOfMonth(y, Math.floor((m - 1) / 3) * 3 + 3), t1 - 1); label = `Q${Math.floor((m - 1) / 3) + 1} ${y}`; break;
      default: end = Math.min(isoEndOfMonth(y, 12), t1 - 1); label = String(y);
    }
    cols.push({ key: isoDay(d), label, start: d, end });
    d = end + 1;
  }
  return { tiers, t0, t1, columns: cols };
}

function isoEndOfMonth(y: number, m: number): number {
  const next = m === 12 ? Date.UTC(y + 1, 0, 1) : Date.UTC(y, m, 1);
  return Math.round(next / 86400000) - 1;
}

export function layoutCalendar(spec: CalendarSpec, frame: Frame, style: Style, measure: Measure): Layout {
  const warnings: string[] = [];
  const cal = calendarColumns(spec);
  if (!cal) return { prims: [], warnings: ["Enter a start and end date for the calendar"] };
  const { tiers, t0, t1, columns } = cal;
  if (columns.length > 60) warnings.push(`${columns.length} columns are too many to read: choose a coarser timeline (e.g. weeks)`);
  const fs = style.fontSize;
  const pad = fs * 0.4;
  const lineH = fs * 1.2;
  const inset = Math.max(2, fs * 0.25);
  const filled = spec.headerStyle !== "plain";

  // ---- Horizontal ----
  const rows = spec.rows.length ? spec.rows : [{ label: "", height: 4, style: "normal" as const, align: "top-left" as const, cells: {} }];
  const showLabels = spec.showLabels && rows.some((r) => r.label.trim());
  const labelW = showLabels ? Math.min(Math.max(...rows.map((r) => measure(r.label, fs, true))) + 2 * pad, frame.width * 0.25) : 0;
  const tlLeft = frame.left + labelW;
  const tlRight = frame.left + frame.width;
  const x = (d: number) => tlLeft + ((d - t0) / (t1 - t0)) * (tlRight - tlLeft);

  // ---- Header ----
  const hdr = timelineHeader({ tiers, t0, t1, x, top: frame.top, filled, fill: spec.headerColor ?? style.palette[0], style, measure });
  const bodyTop = frame.top + hdr.height + (filled ? 0 : pad * 0.5);
  const bodyBottom = frame.top + frame.height;
  if (bodyBottom - bodyTop < lineH * rows.length) warnings.push("The calendar is too short for its rows: make it taller");
  const weights = rows.map((r) => Math.max(0.5, r.height || 1));
  const totalW = weights.reduce((a, b) => a + b, 0);
  const rowTops: number[] = [];
  let acc = bodyTop;
  weights.forEach((w) => { rowTops.push(acc); acc += ((bodyBottom - bodyTop) * w) / totalW; });
  const rowH = (k: number) => ((bodyBottom - bodyTop) * weights[k]) / totalW;

  const fills: RectPrim[] = [];
  const grid: LinePrim[] = [];
  const texts: TextPrim[] = [];
  const line = (x1: number, y1: number, x2: number, y2: number, strong: boolean, role: string) =>
    grid.push({ kind: "line", x1, y1, x2, y2, color: strong ? style.axisColor : style.gridColor, width: strong ? 0.75 : 0.5, role });

  // ---- Row fills and weekend shading ----
  const dayCols = tiers[tiers.length - 1] === "day";
  rows.forEach((r, k) => {
    const y = rowTops[k], h = rowH(k);
    if (r.fill) fills.push({ kind: "rect", x: frame.left, y, w: frame.width, h, fill: r.fill, role: `rowfill:${k}` });
    else if (dayCols && spec.shadeWeekends && style.textColor !== "#FFFFFF") {
      columns.forEach((c, ci) => {
        if (fromDay(c.start).wd >= 5) fills.push({ kind: "rect", x: x(c.start), y, w: x(c.end + 1) - x(c.start), h, fill: "#F2F2F2", role: `wkend:${k}:${ci}` });
      });
    }
  });

  // ---- Cells ----
  const overflow = new Set<number>();
  rows.forEach((r, k) => {
    const y = rowTops[k], h = rowH(k);
    const bold = r.style === "bold", italic = r.style === "italic";
    const centred = r.align === "center";
    if (showLabels && r.label.trim()) {
      texts.push({ kind: "text", x: frame.left, y, w: labelW, h, text: r.label, size: fs, color: style.textColor, bold: true, wrap: true, inset, align: "left", valign: centred ? "middle" : "top", role: `rlab:${k}` });
    }
    columns.forEach((c, ci) => {
      const raw = (r.cells[c.key] ?? "").replace(/\r\n?/g, "\n").replace(/\s+$/, "");
      if (!raw.trim()) return;
      const a = x(c.start), w = x(c.end + 1) - a;
      const lines = raw.split("\n").reduce((n, ln) => n + wrapLines(ln, w - 2 * inset, fs, measure, bold).length, 0);
      if (lines * lineH > h - 2 * inset + 0.5 && !overflow.has(k)) {
        overflow.add(k);
        warnings.push(`Row ${r.label ? `"${r.label}"` : k + 1}: text does not fit (e.g. ${c.label}) — make the row taller or the text shorter`);
      }
      const ink = r.fill ? "#1A1A1A" : style.textColor;
      texts.push({ kind: "text", x: a, y, w, h, text: raw, size: fs, color: ink, bold, italic, wrap: true, inset, align: centred ? "center" : "left", valign: centred ? "middle" : "top", role: `cell:${k}:${ci}` });
    });
  });

  // ---- Grid: column lines (stronger at the coarse tier), row lines, outer border ----
  const coarse = new Set(hdr.periods.length > 1 ? hdr.periods[0].slice(1).map((p: Period) => p.start) : []);
  columns.forEach((c, ci) => { if (ci > 0) line(x(c.start), bodyTop, x(c.start), bodyBottom, coarse.has(c.start), `vgrid:${ci}`); });
  rows.forEach((_, k) => { if (k > 0) line(frame.left, rowTops[k], tlRight, rowTops[k], false, `hgrid:${k}`); });
  if (showLabels) line(tlLeft, bodyTop, tlLeft, bodyBottom, true, "labelsep");
  line(frame.left, bodyBottom, tlRight, bodyBottom, true, "bottom");
  line(frame.left, bodyTop, frame.left, bodyBottom, true, "left");
  line(tlRight, bodyTop, tlRight, bodyBottom, true, "right");
  if (!filled || showLabels) line(frame.left, bodyTop, tlRight, bodyTop, true, "top");

  const prims: Prim[] = [...fills, ...grid, ...hdr.lines, ...hdr.rects, ...hdr.texts, ...texts];
  return { prims, warnings };
}
