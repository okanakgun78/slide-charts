// PowerPoint renderer: primitives -> native shapes, grouped, with the chart spec stored in a tag.
//
// Mac findings from pilot v0.1 (PowerPoint 16.114):
// - A width/height of 0 in add options is replaced by a default (~72 pt), so flat lines came out
//   diagonal. Zero dimensions are now drawn as LINE_EPS.
// - Deleting a group failed (GeneralException) or was ignored, so removal tries several strategies.

import { layoutCategory } from "../engine/category";
import { luminance, normHex } from "../engine/format";
import { layoutGantt } from "../engine/gantt";
import { migrate } from "../engine/spec";
import type { Align, ChartSpec, Frame, Layout, Measure, Prim, Style, VAlign } from "../engine/types";

export const TAG = "SLIDECHARTS";
export const APP_VERSION = "0.5.0";
/** Thickness used instead of 0 for horizontal/vertical line boxes (points). */
export const LINE_EPS = 0.01;

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
  return spec.type === "gantt" ? layoutGantt(spec, frame, style, measure) : layoutCategory(spec, frame, style, measure);
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

/** Office.js error -> one line including where it failed. */
export function errText(e: unknown): string {
  const err = e as {
    code?: string;
    message?: string;
    debugInfo?: { errorLocation?: string; statement?: string; code?: string };
  };
  const loc = err.debugInfo?.errorLocation ? ` @ ${err.debugInfo.errorLocation}` : "";
  const stmt = err.debugInfo?.statement ? ` [${err.debugInfo.statement.slice(0, 120)}]` : "";
  return `${err.code ?? "Error"}: ${err.message ?? String(e)}${loc}${stmt}`;
}

const H_ALIGN: Record<Align, "Left" | "Center" | "Right"> = { left: "Left", center: "Center", right: "Right" };
const V_ALIGN: Record<VAlign, "Top" | "Middle" | "Bottom"> = { top: "Top", middle: "Middle", bottom: "Bottom" };

/** Geometry of a line primitive as PowerPoint add options + rotation. Pure; unit-tested. */
export function lineGeometry(x1: number, y1: number, x2: number, y2: number): { left: number; top: number; width: number; height: number; rotation: number } {
  const dx = x2 - x1;
  const dy = y2 - y1;
  if (Math.abs(dy) < 1e-9) return { left: Math.min(x1, x2), top: y1, width: Math.max(Math.abs(dx), LINE_EPS), height: LINE_EPS, rotation: 0 };
  if (Math.abs(dx) < 1e-9) return { left: x1, top: Math.min(y1, y2), width: LINE_EPS, height: Math.max(Math.abs(dy), LINE_EPS), rotation: 0 };
  // No flip in the API: draw a flat line of the right length centred on the midpoint, then rotate.
  const len = Math.hypot(dx, dy);
  const cx = (x1 + x2) / 2;
  const cy = (y1 + y2) / 2;
  return { left: cx - len / 2, top: cy - LINE_EPS / 2, width: len, height: LINE_EPS, rotation: (Math.atan2(dy, dx) * 180) / Math.PI };
}

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
  if (p.kind === "shape") {
    const s = shapes.addGeometricShape(p.shape, { left: p.x, top: p.y, width: Math.max(p.w, 0.1), height: Math.max(p.h, 0.1) });
    s.fill.setSolidColor(p.fill);
    if (p.stroke) s.lineFormat.color = p.stroke;
    else s.lineFormat.visible = false;
    s.name = p.role;
    return s;
  }
  if (p.kind === "line") {
    const g = lineGeometry(p.x1, p.y1, p.x2, p.y2);
    const s = shapes.addLine("Straight", { left: g.left, top: g.top, width: g.width, height: g.height });
    if (g.rotation !== 0) s.rotation = g.rotation;
    s.lineFormat.color = p.color;
    s.lineFormat.weight = p.width;
    if (p.dash) s.lineFormat.dashStyle = "Dash";
    s.name = p.role;
    return s;
  }
  const s = shapes.addTextBox(p.text, { left: p.x, top: p.y, width: Math.max(p.w, 1), height: Math.max(p.h, 1) });
  const tf = s.textFrame;
  tf.autoSizeSetting = "AutoSizeNone";
  tf.wordWrap = !!p.wrap;
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
  drawn: Frame;
  stored: StoredChart;
  warnings: string[];
}

/** Draws a chart on the slide with `slideId` and returns the new group's id. */
export async function drawChart(
  slideId: string,
  spec: ChartSpec,
  frame: Frame,
  style: Style,
  measure: Measure,
): Promise<DrawResult> {
  return PowerPoint.run(async (ctx) => {
    const t0 = performance.now();
    const slide = ctx.presentation.slides.getItem(slideId);
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
    return { groupId: group.id, shapeCount: prims.length, msShapes: t1 - t0, msGroupTag: t2 - t1, drawn, stored, warnings };
  });
}

export interface ChartRef {
  slideId: string;
  groupId: string;
  stored: StoredChart;
  current: Frame;
}

/** Reads a chart by ids. Returns null if the shape has no Slide Charts tag. */
export async function readChart(slideId: string, groupId: string): Promise<ChartRef | null> {
  return PowerPoint.run(async (ctx) => {
    const shape = ctx.presentation.slides.getItem(slideId).shapes.getItemOrNullObject(groupId);
    shape.load("id,left,top,width,height");
    await ctx.sync();
    if (shape.isNullObject) return null;
    const tag = shape.tags.getItemOrNullObject(TAG);
    tag.load("value");
    await ctx.sync();
    if (tag.isNullObject) return null;
    const stored = JSON.parse(tag.value) as StoredChart;
    stored.spec = migrate(stored.spec);
    return {
      slideId, groupId,
      stored,
      current: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
    };
  });
}

/** Finds the chart the user has selected (or a shape inside it). */
export async function findSelectedChart(): Promise<ChartRef | null> {
  const ids = await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    const sel = ctx.presentation.getSelectedShapes();
    sel.load("items/id,items/level");
    await ctx.sync();
    if (sel.items.length === 0 || slides.items.length === 0) return null;
    let shape = sel.items[0];
    // A click inside a group can select a child; walk up to the top-level group.
    for (let guard = 0; shape.level > 0 && guard < 5; guard++) {
      shape = shape.parentGroup;
      shape.load("id,level");
      await ctx.sync();
    }
    return { slideId: slides.items[0].id, groupId: shape.id };
  });
  return ids ? readChart(ids.slideId, ids.groupId) : null;
}

/** True if a shape with this id still exists on the slide. */
export async function shapeExists(slideId: string, shapeId: string): Promise<boolean> {
  return PowerPoint.run(async (ctx) => {
    const s = ctx.presentation.slides.getItem(slideId).shapes.getItemOrNullObject(shapeId);
    await ctx.sync();
    return !s.isNullObject;
  });
}

export type DeleteStrategy = "delete-by-id" | "delete-children" | "ungroup-then-delete";

/** Removes a group, trying strategies in order until the group is gone. */
export async function removeGroup(slideId: string, groupId: string): Promise<{ strategy: DeleteStrategy | null; log: string[] }> {
  const log: string[] = [];
  const strategies: [DeleteStrategy, (ctx: PowerPoint.RequestContext, shapes: PowerPoint.ShapeCollection) => Promise<void>][] = [
    ["delete-by-id", async (ctx, shapes) => {
      shapes.getItem(groupId).delete();
      await ctx.sync();
    }],
    ["delete-children", async (ctx, shapes) => {
      const kids = shapes.getItem(groupId).group.shapes;
      kids.load("items/id");
      await ctx.sync();
      kids.items.forEach((k) => k.delete());
      await ctx.sync();
    }],
    ["ungroup-then-delete", async (ctx, shapes) => {
      const g = shapes.getItem(groupId);
      const kids = g.group.shapes;
      kids.load("items/id");
      await ctx.sync();
      const ids = kids.items.map((k) => k.id);
      g.group.ungroup();
      await ctx.sync();
      ids.forEach((id) => shapes.getItemOrNullObject(id).delete());
      await ctx.sync();
    }],
  ];
  for (const [name, fn] of strategies) {
    try {
      await PowerPoint.run(async (ctx) => fn(ctx, ctx.presentation.slides.getItem(slideId).shapes));
    } catch (e) {
      log.push(`${name}: ${errText(e)}`);
    }
    if (!(await shapeExists(slideId, groupId))) {
      log.push(`${name}: ok`);
      return { strategy: name, log };
    }
    log.push(`${name}: group still present`);
  }
  return { strategy: null, log };
}

export interface RedrawResult {
  groupId: string;
  stored: StoredChart;
  warnings: string[];
  frame: Frame;
  moved: boolean;
  resized: boolean;
  deleteStrategy: DeleteStrategy | null;
  log: string[];
  ms: number;
}

/** Redraws a chart in place with `spec`, following any move/resize the user applied. */
export async function redrawChart(ref: ChartRef, spec: ChartSpec, style: Style, measure: Measure): Promise<RedrawResult> {
  const t0 = performance.now();
  const d = ref.stored.drawn;
  const c = ref.current;
  const moved = Math.abs(c.left - d.left) > 0.5 || Math.abs(c.top - d.top) > 0.5;
  const resized = Math.abs(c.width - d.width) > 0.5 || Math.abs(c.height - d.height) > 0.5;
  const frame = refitFrame(ref.stored, ref.current);
  const del = await removeGroup(ref.slideId, ref.groupId);
  if (!del.strategy) throw new Error(`could not remove old chart: ${del.log.join(" | ")}`);
  const r = await drawChart(ref.slideId, spec, frame, style, measure);
  try {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(ref.slideId).setSelectedShapes([r.groupId]);
      await ctx.sync();
    });
  } catch (e) {
    del.log.push(`reselect: ${errText(e)}`);
  }
  return { groupId: r.groupId, stored: r.stored, warnings: r.warnings, frame, moved, resized, deleteStrategy: del.strategy, log: del.log, ms: performance.now() - t0 };
}

/** Reads the theme accent colours of a slide (PowerPointApi 1.10). */
export async function readThemeAccents(slideId: string): Promise<string[]> {
  return PowerPoint.run(async (ctx) => {
    const scheme = ctx.presentation.slides.getItem(slideId).themeColorScheme;
    const keys: ("Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6")[] = [
      "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
    ];
    const results = keys.map((k) => scheme.getThemeColor(k));
    await ctx.sync();
    return results.map((r) => r.value);
  });
}

/**
 * Best-effort background colour of a slide (slide, then layout, then master; solid fills only).
 * Returns null when it cannot be determined (pictures, gradients, older APIs).
 */
export async function readSlideBackground(slideId: string): Promise<string | null> {
  try {
    return await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(slideId);
      const sources = [slide.background.fill, slide.layout.background.fill, slide.slideMaster.background.fill];
      const solids = sources.map((f) => f.getSolidFillOrNullObject());
      solids.forEach((sf) => sf.load("color"));
      slide.background.load("isMasterBackgroundFollowed");
      await ctx.sync();
      const order = slide.background.isMasterBackgroundFollowed ? solids.slice(1) : solids;
      for (const sf of order) if (!sf.isNullObject && normHex(sf.color)) return normHex(sf.color);
      return null;
    });
  } catch {
    return null;
  }
}

/** Text and axis colours that read on a background colour. */
export function inkFor(background: string | null): { textColor: string; axisColor: string; gridColor: string } {
  if (background && luminance(background) < 0.3) return { textColor: "#FFFFFF", axisColor: "#D9D9D9", gridColor: "#8C96B0" };
  return { textColor: "#1A1A1A", axisColor: "#7F7F7F", gridColor: "#D9D9D9" };
}
