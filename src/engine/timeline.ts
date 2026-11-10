// Shared time axis for Gantt and Calendar charts: tiers (year/quarter/month/week/day), periods
// with label candidates, range snapping, and the header (filled cells as in the user's sample, or plain).

import { fromDay, monthName, toDay } from "./dates";
import type { GanttSpec, GanttTier, LinePrim, Measure, RectPrim, Style, TextPrim } from "./types";

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
export function periodEnd(start: number, tier: GanttTier): number {
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

/** Range [t0, t1) covering days lo..hi, snapped to the finest tier (days start on the first day). */
export function timelineRange(lo: number, hi: number, tiers: GanttTier[]): { t0: number; t1: number } {
  const finest = tiers[tiers.length - 1];
  const t0 = finest === "day" ? lo : periodStart(lo, finest);
  const t1 = periodEnd(periodStart(hi, finest), finest) + 1;
  return { t0, t1 };
}

export interface HeaderOptions {
  tiers: GanttTier[];
  t0: number;
  t1: number;
  x: (d: number) => number;
  top: number;
  filled: boolean;
  fill: string;
  style: Style;
  measure: Measure;
}

/** Timeline header: one row per tier; day cells use two lines ("02" over "Mon") when they fit. */
export function timelineHeader(o: HeaderOptions): { rects: RectPrim[]; lines: LinePrim[]; texts: TextPrim[]; height: number; periods: Period[][] } {
  const { tiers, t0, t1, x, filled, style, measure } = o;
  const fs = style.fontSize;
  const small = Math.round(fs * 0.85 * 2) / 2;
  const ink = filled ? "#FFFFFF" : style.textColor;
  const multiYear = fromDay(t0).y !== fromDay(t1 - 1).y;
  const periods = tiers.map((t) => tierPeriods(t0, t1, t, { tiers, multiYear }));
  const one = filled ? fs * 1.6 : fs * 1.3;
  const dayW = tiers.includes("day") ? x(t0 + 1) - x(t0) : 0;
  const twoLine = tiers.includes("day") && measure("Wed", fs, true) <= dayW - 2 && measure("00", fs, true) <= dayW - 2;
  const tierH = tiers.map((t) => (t === "day" && twoLine ? (filled ? fs * 2.75 : fs * 2.5) : one));
  const rects: RectPrim[] = [];
  const lines: LinePrim[] = [];
  const texts: TextPrim[] = [];
  const text = (t: Omit<TextPrim, "kind">) => texts.push({ kind: "text", ...t });
  let ty = o.top;
  tiers.forEach((tier, ti) => {
    const h = tierH[ti];
    const finestTier = ti === tiers.length - 1;
    periods[ti].forEach((p, k) => {
      const a = x(p.start), b = x(p.end + 1), w = b - a;
      if (filled) rects.push({ kind: "rect", x: a, y: ty, w, h, fill: o.fill, stroke: "#FFFFFF", strokeWidth: 0.75, role: `hd:${tier}:${k}` });
      else if (k > 0) lines.push({ kind: "line", x1: a, y1: ty, x2: a, y2: ty + h, color: style.gridColor, width: 0.5, role: `hdl:${tier}:${k}` });
      const size = filled || !finestTier || tiers.length === 1 ? fs : small;
      const bold = filled || !finestTier || tiers.length === 1;
      if (tier === "day" && twoLine && p.lines) {
        const pair = p.lines.find(([l1, l2]) => measure(l1, size, bold) <= w - 2 && measure(l2, size, bold) <= w - 2);
        if (pair) {
          text({ x: a, y: ty, w, h: h / 2, text: pair[0], size, bold, color: ink, align: "center", valign: "bottom", role: `hdt:${tier}:${k}:1` });
          text({ x: a, y: ty + h / 2, w, h: h / 2, text: pair[1], size, bold, color: ink, align: "center", valign: "top", role: `hdt:${tier}:${k}:2` });
          return;
        }
      }
      const label = p.labels.find((l) => measure(l, size, bold) <= w - 3);
      if (label) text({ x: a, y: ty, w, h, text: label, size, bold, color: ink, align: "center", valign: "middle", role: `hdt:${tier}:${k}` });
    });
    ty += h;
  });
  return { rects, lines, texts, height: ty - o.top, periods };
}
