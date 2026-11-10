// Shared types for the chart engine.
// The engine is pure: spec + frame + style -> primitives (points, PowerPoint units).
// Renderers turn primitives into PowerPoint shapes (src/ppt) or SVG (src/engine/svg.ts).

export interface Frame {
  left: number;
  top: number;
  width: number;
  height: number;
}

export type Align = "left" | "center" | "right";
export type VAlign = "top" | "middle" | "bottom";

export interface RectPrim {
  kind: "rect";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  role: string;
}

export interface TextPrim {
  kind: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  size: number;
  color: string;
  bold?: boolean;
  italic?: boolean;
  /** Let PowerPoint wrap the text inside the box width. */
  wrap?: boolean;
  /** Inner margin (points) on all sides; 0 when omitted. */
  inset?: number;
  align: Align;
  valign: VAlign;
  role: string;
}

export interface LinePrim {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  width: number;
  dash?: boolean;
  role: string;
}

export interface ShapePrim {
  kind: "shape";
  shape: "Ellipse" | "Diamond";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  stroke?: string;
  role: string;
}

export type Prim = RectPrim | TextPrim | LinePrim | ShapePrim;

// ---------- Chart specifications (stored in the shape tag) ----------

export type Mode = "stacked" | "clustered" | "percent";

export interface Series {
  name: string;
  values: (number | null)[];
  /** Explicit colour (hex); palette colour by position when omitted. */
  color?: string;
  /** In combination charts: draw this series as a line. */
  asLine?: boolean;
}

export interface CategorySpec {
  v: 2;
  type: "column" | "bar" | "line" | "combo";
  /** Bar arrangement (ignored for pure line charts; combo supports stacked/clustered). */
  mode: Mode;
  categories: string[];
  series: Series[];
  labels: {
    /** Value labels on segments / points. */
    values: boolean;
    /** Totals above stacks (stacked/percent). */
    totals: boolean;
    /** Series names: right of the last category, legend on top, or hidden. */
    legend: "auto" | "right" | "top" | "none";
  };
  /** Fixed decimals for value labels; null = automatic. */
  decimals: number | null;
}

export type GanttRowKind = "task" | "milestone" | "header";

export interface GanttRow {
  label: string;
  kind: GanttRowKind;
  /** Dates as typed or pasted (parsed with the spec's date order). End is inclusive. */
  start: string;
  end: string;
  color?: string;
  note?: string;
}

export type GanttTier = "year" | "quarter" | "month" | "week" | "day";

export interface GanttSpec {
  v: 2;
  type: "gantt";
  rows: GanttRow[];
  /** Header tiers to show (finest one sets the grid), or "auto" from the plan's span. */
  tiers: GanttTier[] | "auto";
  /** "filled": solid header cells with white text (user's sample); "plain": text only. */
  headerStyle: "filled" | "plain";
  /** Header cell colour; palette navy when omitted. */
  headerColor?: string;
  /** Grey Saturday/Sunday columns when the grid is in days. */
  shadeWeekends: boolean;
  /** v0.4 setting, replaced by `tiers` (kept for reading old charts). */
  scale?: "auto" | "week" | "month" | "quarter" | "year";
  showToday: boolean;
  /** Show the date next to milestones. */
  milestoneDates: boolean;
  /** How to read ambiguous dates like 04/05/2026. */
  dateOrder: "DMY" | "MDY";
}

export interface CalendarRow {
  /** Shown in an optional label column on the left. */
  label: string;
  /** Relative height (1 = one line; the sample's plan row is about 6). */
  height: number;
  style: "normal" | "bold" | "italic";
  /** Cell background, e.g. yellow for a totals row. */
  fill?: string;
  align: "top-left" | "center";
  /** Cell text keyed by the ISO date of the column's first day ("2026-11-02"). Use \n for new lines. */
  cells: Record<string, string>;
}

export interface CalendarSpec {
  v: 2;
  type: "calendar";
  /** Date range as typed (inclusive). */
  start: string;
  end: string;
  tiers: GanttTier[] | "auto";
  headerStyle: "filled" | "plain";
  headerColor?: string;
  shadeWeekends: boolean;
  /** Show row labels in a column on the left. */
  showLabels: boolean;
  rows: CalendarRow[];
  dateOrder: "DMY" | "MDY";
}

export type ChartSpec = CategorySpec | GanttSpec | CalendarSpec;

export interface Style {
  fontSize: number;
  textColor: string;
  axisColor: string;
  gridColor: string;
  /** Fill colours for bars, by series position. */
  palette: string[];
  /** Line colours, by series position (more contrast than the bar ramp). */
  linePalette: string[];
  decimalSep: string;
  thousandsSep: string;
}

/** Returns the rendered width (points) of `text` at `size` points. */
export type Measure = (text: string, size: number, bold?: boolean) => number;

export interface Layout {
  prims: Prim[];
  warnings: string[];
}
