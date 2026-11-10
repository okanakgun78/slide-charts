// Shared types for the chart engine.
// The engine is pure: data + frame + style -> primitives (points, PowerPoint units).
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

export type Prim = RectPrim | TextPrim | LinePrim;

export interface Series {
  name: string;
  values: (number | null)[];
}

export interface StackedSpec {
  v: 1;
  type: "stacked";
  categories: string[];
  series: Series[];
  showTotals: boolean;
  /** Fixed decimals for labels; auto when omitted. */
  decimals?: number;
}

export type ChartSpec = StackedSpec;

export interface Style {
  fontSize: number;
  textColor: string;
  axisColor: string;
  palette: string[];
}

/** Returns the rendered width (points) of `text` at `size` points. */
export type Measure = (text: string, size: number, bold?: boolean) => number;

export interface Layout {
  prims: Prim[];
  warnings: string[];
}
