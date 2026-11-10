// PowerPoint renderer: primitives -> native shapes, grouped, with the chart spec stored in a tag.

import { layoutStacked } from "../engine/stacked";
import type { Align, ChartSpec, Frame, Layout, Measure, Prim, Style, VAlign } from "../engine/types";

export const TAG = "SLIDECHARTS";
export const APP_VERSION = "0.1.0-pilot";

export interface StoredChart {
  app: "slide-charts";
  ver: string;
  spec: ChartSpec;
  /** Frame the layout was computed for. */
  frame: Frame;
  /** Group bounds right after drawing; compared later to detect move/resize. */
  drawn: Frame;
}

export function layoutChart(spec: ChartSpec, frame: Frame, style: Style, measure: Measure): Layout {
  switch (spec.type) {
    case "stacked":
      return layoutStacked(spec, frame, style, measure);
  }
}

/** Maps the stored layout frame through the move/resize the user applied to the group. */
export function refitFrame(stored: StoredChart, current: Frame): Frame {
  const d = stored.drawn;
  const sx = d.width > 0 ? current.width / d.width : 1;
  const sy = d.height > 0 ? current.height / d.height : 1;
  return {
    left: current.left + (stored.frame.left - d.left) * sx,
    top: current.top + (stored.frame.top - d.top) * sy,
    width: stored.frame.width * sx,
    height: stored.frame.height * sy,
  };
}

const H_ALIGN: Record<Align, "Left" | "Center" | "Right"> = {
  left: "Left",
  center: "Center",
  right: "Right",
};
const V_ALIGN: Record<VAlign, "Top" | "Middle" | "Bottom"> = {
  top: "Top",
  middle: "Middle",
  bottom: "Bottom",
};

/** Adds one primitive to the slide (queued; caller syncs). */
export function addPrim(shapes: PowerPoint.ShapeCollection, p: Prim): PowerPoint.Shape {
  if (p.kind === "rect") {
    const s = shapes.addGeometricShape("Rectangle", {
      left: p.x, top: p.y, width: Math.max(p.w, 0.1), height: Math.max(p.h, 0.1),
    });
    s.fill.setSolidColor(p.fill);
    if (p.stroke) {
      s.lineFormat.color = p.stroke;
      s.lineFormat.weight = p.strokeWidth ?? 0.75;
    } else {
      s.lineFormat.visible = false;
    }
    s.name = p.role;
    return s;
  }
  if (p.kind === "line") {
    const dx = p.x2 - p.x1;
    const dy = p.y2 - p.y1;
    let s: PowerPoint.Shape;
    if (dx === 0 || dy === 0) {
      s = shapes.addLine("Straight", {
        left: Math.min(p.x1, p.x2), top: Math.min(p.y1, p.y2), width: Math.abs(dx), height: Math.abs(dy),
      });
    } else {
      // The API has no flip, so diagonal lines are drawn horizontally and rotated about their centre.
      const len = Math.hypot(dx, dy);
      const cx = (p.x1 + p.x2) / 2;
      const cy = (p.y1 + p.y2) / 2;
      s = shapes.addLine("Straight", { left: cx - len / 2, top: cy, width: len, height: 0 });
      s.rotation = (Math.atan2(dy, dx) * 180) / Math.PI;
    }
    s.lineFormat.color = p.color;
    s.lineFormat.weight = p.width;
    if (p.dash) s.lineFormat.dashStyle = "Dash";
    s.name = p.role;
    return s;
  }
  const s = shapes.addTextBox(p.text, { left: p.x, top: p.y, width: Math.max(p.w, 1), height: Math.max(p.h, 1) });
  const tf = s.textFrame;
  tf.autoSizeSetting = "AutoSizeNone";
  tf.wordWrap = false;
  tf.leftMargin = 0;
  tf.rightMargin = 0;
  tf.topMargin = 0;
  tf.bottomMargin = 0;
  tf.verticalAlignment = V_ALIGN[p.valign];
  const tr = tf.textRange;
  tr.font.size = p.size;
  tr.font.color = p.color;
  tr.font.bold = !!p.bold;
  tr.paragraphFormat.horizontalAlignment = H_ALIGN[p.align];
  s.name = p.role;
  return s;
}

export interface DrawResult {
  groupId: string;
  shapeCount: number;
  msShapes: number;
  msGroupTag: number;
  warnings: string[];
}

/** Draws a chart on `slide` and returns the new group's id. */
export async function drawChart(
  ctx: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  spec: ChartSpec,
  frame: Frame,
  style: Style,
  measure: Measure,
): Promise<DrawResult> {
  const t0 = performance.now();
  const { prims, warnings } = layoutChart(spec, frame, style, measure);
  if (prims.length === 0) throw new Error(warnings.join("; ") || "Nothing to draw");
  const created = prims.map((p) => addPrim(slide.shapes, p));
  created.forEach((s) => s.load("id"));
  await ctx.sync();
  const t1 = performance.now();

  const group = slide.shapes.addGroup(created.map((s) => s.id));
  group.name = `Slide Charts (${spec.type})`;
  group.load("id,left,top,width,height");
  await ctx.sync();
  const drawn: Frame = { left: group.left, top: group.top, width: group.width, height: group.height };
  const stored: StoredChart = { app: "slide-charts", ver: APP_VERSION, spec, frame, drawn };
  group.tags.add(TAG, JSON.stringify(stored));
  await ctx.sync();
  const t2 = performance.now();

  return { groupId: group.id, shapeCount: prims.length, msShapes: t1 - t0, msGroupTag: t2 - t1, warnings };
}

export interface SelectedChart {
  shape: PowerPoint.Shape;
  slide: PowerPoint.Slide;
  stored: StoredChart;
  current: Frame;
}

/** Returns the chart the user has selected (or a shape inside it), or null. */
export async function getSelectedChart(ctx: PowerPoint.RequestContext): Promise<SelectedChart | null> {
  const sel = ctx.presentation.getSelectedShapes();
  sel.load("items/id,items/level");
  await ctx.sync();
  if (sel.items.length === 0) return null;
  let shape = sel.items[0];
  // A click inside a group can select a child; walk up to the top-level group.
  for (let guard = 0; shape.level > 0 && guard < 5; guard++) {
    shape = shape.parentGroup;
    shape.load("id,level");
    await ctx.sync();
  }
  const tag = shape.tags.getItemOrNullObject(TAG);
  tag.load("value");
  shape.load("id,left,top,width,height");
  const slide = shape.getParentSlide();
  slide.load("id");
  await ctx.sync();
  if (tag.isNullObject) return null;
  const stored = JSON.parse(tag.value) as StoredChart;
  return {
    shape, slide, stored,
    current: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
  };
}

/** Reads the theme accent colours of a slide (PowerPointApi 1.10). */
export async function readThemeAccents(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide): Promise<string[]> {
  const scheme = slide.themeColorScheme;
  const keys: ("Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6")[] = [
    "Accent1", "Accent2", "Accent3",
    "Accent4", "Accent5", "Accent6",
  ];
  const results = keys.map((k) => scheme.getThemeColor(k));
  await ctx.sync();
  return results.map((r) => r.value);
}
