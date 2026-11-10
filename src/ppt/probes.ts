// Pilot checks: each probe exercises one capability the full product depends on and records
// what PowerPoint on this machine actually does. Results are shown in the pane and copied
// into a report the user sends back.

import { DEFAULT_PALETTE, normHex } from "../engine/format";
import { canvasMeasure } from "../engine/measure";
import { defaultStyle, sampleCategory } from "../engine/spec";
import type { CategorySpec, Frame, Style } from "../engine/types";
import {
  addPrim, drawChart, errText, LINE_EPS, readChart, readThemeAccents, redrawChart, shapeExists, TAG,
} from "./render";

export interface ProbeResult {
  id: string;
  name: string;
  ok: boolean | null; // null = informational / needs a visual check
  detail: string;
  ms?: number;
}

export interface PilotContext {
  slideId?: string;
  slideW: number;
  slideH: number;
  palette: string[];
  fontName?: string;
  /** PowerPoint width / canvas width for the same text (text-measurement calibration). */
  measureRatio?: number;
  chartGroupId?: string;
}

export const SAMPLE_SPEC: CategorySpec = sampleCategory("column", "stacked");

export function chartStyle(palette?: string[]): Style {
  const st = defaultStyle();
  return palette ? { ...st, palette } : st;
}

export const FALLBACK_FONTS = '"Aptos", "Calibri", "Helvetica Neue", Arial, sans-serif';

export function measureFor(pc: Pick<PilotContext, "fontName" | "measureRatio"> | undefined) {
  return canvasMeasure(`"${pc?.fontName ?? "Aptos"}", ${FALLBACK_FONTS}`, pc?.measureRatio ?? 1);
}

/** New numbers: each value moves by up to ±25%, rounded like the input. */
export function perturb(spec: CategorySpec): CategorySpec {
  return {
    ...spec,
    series: spec.series.map((s) => ({
      ...s,
      values: s.values.map((v) => (v === null ? null : Math.max(1, Math.round(v * (0.75 + Math.random() * 0.5))))),
    })),
  };
}

const API_SETS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"];

type Probe = (pc: PilotContext) => Promise<Omit<ProbeResult, "id" | "name">>;

async function timed(id: string, name: string, pc: PilotContext, fn: Probe): Promise<ProbeResult> {
  const t0 = performance.now();
  try {
    const r = await fn(pc);
    return { id, name, ...r, ms: Math.round(performance.now() - t0) };
  } catch (e) {
    return { id, name, ok: false, detail: errText(e), ms: Math.round(performance.now() - t0) };
  }
}

const run = <T>(fn: (ctx: PowerPoint.RequestContext) => Promise<T>) => PowerPoint.run(fn);
const r1 = (v: number) => Math.round(v * 100) / 100;

/** Adds a small labelled caption next to a visual probe. */
function caption(slide: PowerPoint.Slide, text: string, left: number, top: number, width = 230) {
  const t = slide.shapes.addTextBox(text, { left, top, width, height: 16 });
  t.textFrame.textRange.font.size = 9;
  t.textFrame.textRange.font.color = "#404040";
  t.textFrame.wordWrap = true;
  return t;
}

/** Two small rectangles grouped together; returns the group id. */
async function makeTestGroup(slideId: string, left: number, top: number): Promise<string> {
  return run(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    const a = shapes.addGeometricShape("Rectangle", { left, top, width: 8, height: 8 });
    const b = shapes.addGeometricShape("Rectangle", { left: left + 10, top, width: 8, height: 8 });
    a.load("id"); b.load("id");
    await ctx.sync();
    const g = shapes.addGroup([a.id, b.id]);
    g.load("id");
    await ctx.sync();
    return g.id;
  });
}

const probes: [string, string, Probe][] = [
  ["env", "Environment and API versions", async () => {
    const d = Office.context.diagnostics;
    const sets = API_SETS.map((v) => `${v}:${Office.context.requirements.isSetSupported("PowerPointApi", v) ? "yes" : "NO"}`);
    const ok = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
    return { ok, detail: `${d.host} ${d.platform} ${d.version} | ${sets.join(" ")}` };
  }],

  ["slide", "Add a clean test slide and go to it", async (pc) => run(async (ctx) => {
    const master = ctx.presentation.slideMasters.getItemAt(0);
    master.layouts.load("items/id,items/name");
    await ctx.sync();
    const counts = master.layouts.items.map((l) => l.shapes.getCount());
    await ctx.sync();
    const ranked = master.layouts.items
      .map((l, k) => ({ l, n: counts[k].value, pref: /blank|leer|vide|empty/i.test(l.name) ? 0 : /title only/i.test(l.name) ? 1 : 2 }))
      .sort((a, b) => a.pref - b.pref || a.n - b.n);
    const layout = ranked[0]?.l;
    ctx.presentation.slides.add(layout ? { layoutId: layout.id } : undefined);
    ctx.presentation.slides.load("items/id");
    await ctx.sync();
    const slide = ctx.presentation.slides.items[ctx.presentation.slides.items.length - 1];
    pc.slideId = slide.id;
    // Remove the layout's placeholders so the test slide is empty.
    slide.shapes.load("items/id");
    await ctx.sync();
    const removed = slide.shapes.items.length;
    slide.shapes.items.forEach((s) => s.delete());
    ctx.presentation.setSelectedSlides([slide.id]);
    await ctx.sync();
    return { ok: true, detail: `slide ${slide.id}; layout "${layout?.name ?? "default"}" (${ranked[0]?.n ?? "?"} layout shapes); removed ${removed} placeholders` };
  })],

  ["page", "Read slide size", async (pc) => run(async (ctx) => {
    const ps = ctx.presentation.pageSetup;
    ps.load("slideWidth,slideHeight");
    await ctx.sync();
    pc.slideW = ps.slideWidth;
    pc.slideH = ps.slideHeight;
    return { ok: true, detail: `${ps.slideWidth} x ${ps.slideHeight} pt` };
  })],

  ["theme", "Read template theme colours", async (pc) => {
    const raw = await readThemeAccents(pc.slideId!);
    const hex = raw.map(normHex);
    if (hex.every((h) => h)) pc.palette = hex as string[];
    return { ok: hex.every((h) => h), detail: `accents raw=${JSON.stringify(raw)}` };
  }],

  ["autosize", "Measure text (PowerPoint vs task pane)", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const samples = ["Revenue 2024", "1,234.5", "North America", "WWW mmm iii"];
    const boxes = samples.map((t, k) => {
      const s = slide.shapes.addTextBox(t, { left: 10, top: 10 + k * 20, width: 400, height: 18 });
      s.textFrame.wordWrap = false;
      s.textFrame.leftMargin = 0;
      s.textFrame.rightMargin = 0;
      s.textFrame.textRange.font.size = 12;
      s.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";
      return s;
    });
    await ctx.sync();
    boxes.forEach((s) => { s.load("width,height"); s.textFrame.textRange.font.load("name"); });
    await ctx.sync();
    const font = boxes[0].textFrame.textRange.font.name ?? undefined;
    pc.fontName = font;
    const m = canvasMeasure(`"${font}", ${FALLBACK_FONTS}`);
    const fontVisible = document.fonts ? document.fonts.check(`12px "${font}"`) : false;
    const rows = samples.map((t, k) => ({ t, ppt: boxes[k].width, canvas: m(t, 12) }));
    const ratios = rows.map((r) => r.ppt / r.canvas);
    const autosized = rows.some((r) => Math.abs(r.ppt - 400) > 1);
    pc.measureRatio = autosized ? ratios.reduce((a, b) => a + b, 0) / ratios.length : undefined;
    boxes.forEach((s) => s.delete());
    await ctx.sync();
    return {
      ok: autosized,
      detail: `font="${font}" visibleToPane=${fontVisible} ratio=${pc.measureRatio?.toFixed(3)} ` +
        rows.map((r) => `[${r.t}: ppt=${r.ppt.toFixed(1)} canvas=${r.canvas.toFixed(1)}]`).join(" "),
    };
  })],

  ["lines", "Flat lines keep zero height", async (pc) => run(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(pc.slideId!).shapes;
    const y = pc.slideH * 0.9;
    const a = shapes.addLine("Straight", { left: 20, top: y, width: 100, height: 0 });
    const b = shapes.addLine("Straight", { left: 140, top: y, width: 100, height: LINE_EPS });
    const c = shapes.addLine("Straight", { left: 260, top: y, width: 100, height: 1 });
    c.height = 0;
    [a, b, c].forEach((s) => s.load("height,width"));
    await ctx.sync();
    const detail = `opts h=0 -> ${r1(a.height)}; opts h=${LINE_EPS} -> ${r1(b.height)}; set h=0 after -> ${r1(c.height)}`;
    [a, b, c].forEach((s) => s.delete());
    await ctx.sync();
    return { ok: b.height < 0.5, detail };
  })],

  ["chart", "Draw sample stacked chart (timed)", async (pc) => {
    const frame: Frame = { left: pc.slideW * 0.06, top: pc.slideH * 0.12, width: pc.slideW * 0.58, height: pc.slideH * 0.62 };
    const r = await drawChart(pc.slideId!, SAMPLE_SPEC, frame, chartStyle(), measureFor(pc));
    pc.chartGroupId = r.groupId;
    return {
      ok: true,
      detail: `${r.shapeCount} shapes; draw ${Math.round(r.msShapes)} ms, group+tag ${Math.round(r.msGroupTag)} ms; ` +
        `frame h=${r1(frame.height)} drawn h=${r1(r.drawn.height)}` + (r.warnings.length ? `; warnings: ${r.warnings.join(", ")}` : ""),
    };
  }],

  ["tag", "Read chart data back from the file", async (pc) => {
    const ref = await readChart(pc.slideId!, pc.chartGroupId!);
    if (!ref) return { ok: false, detail: "tag missing" };
    const ok = JSON.stringify(ref.stored.spec) === JSON.stringify(SAMPLE_SPEC);
    return { ok, detail: `spec round-trip ${ok ? "identical" : "DIFFERENT"}` };
  }],

  ["bigtag", "Store a large data payload (60 KB)", async (pc) => run(async (ctx) => {
    const g = ctx.presentation.slides.getItem(pc.slideId!).shapes.getItem(pc.chartGroupId!);
    const big = "x".repeat(60000);
    g.tags.add("SLIDECHARTS_PROBE", big);
    await ctx.sync();
    const t = g.tags.getItem("SLIDECHARTS_PROBE");
    t.load("value");
    await ctx.sync();
    const len = t.value.length;
    g.tags.delete("SLIDECHARTS_PROBE");
    await ctx.sync();
    return { ok: len === big.length, detail: `wrote ${big.length}, read ${len}` };
  })],

  ["delete", "Ways to delete a group", async (pc) => {
    const out: string[] = [];
    let anyOk = false;
    const variants: [string, (ctx: PowerPoint.RequestContext, shapes: PowerPoint.ShapeCollection, gid: string) => Promise<void>][] = [
      ["by-id", async (ctx, shapes, gid) => { shapes.getItem(gid).delete(); await ctx.sync(); }],
      ["children", async (ctx, shapes, gid) => {
        const kids = shapes.getItem(gid).group.shapes;
        kids.load("items/id");
        await ctx.sync();
        kids.items.forEach((k) => k.delete());
        await ctx.sync();
      }],
      ["ungroup", async (ctx, shapes, gid) => {
        const g = shapes.getItem(gid);
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
    for (let k = 0; k < variants.length; k++) {
      const [name, fn] = variants[k];
      let gid = "";
      try {
        gid = await makeTestGroup(pc.slideId!, 20 + k * 30, pc.slideH * 0.95);
        await run((ctx) => fn(ctx, ctx.presentation.slides.getItem(pc.slideId!).shapes, gid));
        const gone = !(await shapeExists(pc.slideId!, gid));
        anyOk ||= gone;
        out.push(`${name}: ${gone ? "removed" : "still present"}`);
      } catch (e) {
        const gone = gid ? !(await shapeExists(pc.slideId!, gid).catch(() => false)) : false;
        anyOk ||= gone;
        out.push(`${name}: ${errText(e)}${gone ? " (but removed)" : ""}`);
      }
    }
    return { ok: anyOk, detail: out.join(" | ") };
  }],

  ["roundtrip", "Resize chart, then redraw with new data", async (pc) => {
    // Simulate the user resizing the chart: shrink and move the group via the API.
    let resizeNote = "";
    try {
      await run(async (ctx) => {
        const g = ctx.presentation.slides.getItem(pc.slideId!).shapes.getItem(pc.chartGroupId!);
        g.load("left,width,height");
        await ctx.sync();
        g.width = g.width * 0.85;
        g.height = g.height * 0.9;
        g.left = g.left + 10;
        await ctx.sync();
      });
      resizeNote = "resized via API";
    } catch (e) {
      resizeNote = `resize failed: ${errText(e)}`;
    }
    const ref = await readChart(pc.slideId!, pc.chartGroupId!);
    if (!ref) return { ok: false, detail: `${resizeNote}; chart not found` };
    const r = await redrawChart(ref, perturb(SAMPLE_SPEC), chartStyle(), measureFor(pc));
    const oldGone = !(await shapeExists(pc.slideId!, pc.chartGroupId!));
    const back = await readChart(pc.slideId!, r.groupId);
    pc.chartGroupId = r.groupId;
    const ok = oldGone && !!back;
    return {
      ok,
      detail: `${resizeNote}; moved=${r.moved} resized=${r.resized}; delete=${r.deleteStrategy}; ` +
        `old removed=${oldGone}; new tagged=${!!back}; new frame w=${r1(r.frame.width)} h=${r1(r.frame.height)}; ${r.log.join(" | ")}`,
    };
  }],

  ["harvey", "Harvey ball via shape adjustments (75%)", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.12, d = 28;
    const ring = slide.shapes.addGeometricShape("Ellipse", { left: x, top: y, width: d, height: d });
    ring.fill.setSolidColor("#FFFFFF");
    ring.lineFormat.color = pc.palette[0];
    ring.lineFormat.weight = 1;
    const pie = slide.shapes.addGeometricShape("Pie", { left: x, top: y, width: d, height: d });
    pie.fill.setSolidColor(pc.palette[0]);
    pie.lineFormat.visible = false;
    // Pilot v0.1 showed: adjustments are degrees, clockwise from 3 o'clock, normalised to -180..180.
    pie.adjustments.set(0, -90); // start at 12 o'clock
    pie.adjustments.set(1, 180); // end at 9 o'clock -> 75% clockwise
    caption(slide, "Harvey ball: 3/4 filled from 12 o'clock", x + d + 8, y + 6);
    await ctx.sync();
    return { ok: null, detail: "check visually" };
  })],

  ["arrows", "Rising line, arrowhead and triangle orientation", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.26;
    // Rising line from bottom-left to top-right.
    const line = addPrim(slide.shapes, { kind: "line", x1: x, y1: y + 50, x2: x + 70, y2: y, color: "#C00000", width: 1.5, role: "probe-diag" });
    const theta = Math.atan2(-50, 70);
    const hw = 9, hh = 8;
    const cx = x + 70 - (hh / 2) * Math.cos(theta), cy = y - (hh / 2) * Math.sin(theta);
    const head = slide.shapes.addGeometricShape("Triangle", { left: cx - hw / 2, top: cy - hh / 2, width: hw, height: hh });
    head.fill.setSolidColor("#C00000");
    head.lineFormat.visible = false;
    head.rotation = 90 + (theta * 180) / Math.PI;
    caption(slide, "Red line rises left to right; arrow tip at its top-right end", x + 80, y + 14, 180);
    // Orientation strip: rotations 0, 90, 180, 270.
    const ty = pc.slideH * 0.44;
    const tri = [0, 90, 180, 270].map((rot, k) => {
      const t = slide.shapes.addGeometricShape("Triangle", { left: x + k * 26, top: ty, width: 14, height: 12 });
      t.fill.setSolidColor("#003760");
      t.lineFormat.visible = false;
      if (rot) t.rotation = rot;
      return t;
    });
    caption(slide, "Triangles should point: up, right, down, left", x + 110, ty - 2, 160);
    line.load("rotation,width,height");
    head.load("rotation");
    tri.forEach((t) => t.load("rotation"));
    await ctx.sync();
    return {
      ok: null,
      detail: `line rot=${r1(line.rotation)} w=${r1(line.width)} h=${r1(line.height)}; head rot=${r1(head.rotation)} (set ${r1(90 + (theta * 180) / Math.PI)}); ` +
        `strip rot=${tri.map((t) => r1(t.rotation)).join(",")} (check visually)`,
    };
  })],

  ["dash", "Dashed value line", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.58;
    const l = addPrim(slide.shapes, { kind: "line", x1: x, y1: y, x2: x + 120, y2: y, color: "#404040", width: 1, dash: true, role: "probe-dash" });
    caption(slide, "Dashed line should be flat", x, y + 6, 160);
    l.load("height");
    l.lineFormat.load("dashStyle");
    await ctx.sync();
    return { ok: l.lineFormat.dashStyle === "Dash" && l.height < 0.5, detail: `dashStyle=${l.lineFormat.dashStyle} height=${r1(l.height)}` };
  })],
];

const NEEDS_CHART = new Set(["tag", "bigtag", "roundtrip"]);

export async function runPilot(onProgress: (r: ProbeResult) => void): Promise<{ results: ProbeResult[]; pc: PilotContext }> {
  const pc: PilotContext = { slideW: 960, slideH: 540, palette: DEFAULT_PALETTE };
  const results: ProbeResult[] = [];
  for (const [id, name, fn] of probes) {
    let r: ProbeResult;
    if (id !== "env" && id !== "slide" && !pc.slideId) r = { id, name, ok: false, detail: "skipped: no test slide" };
    else if (NEEDS_CHART.has(id) && !pc.chartGroupId) r = { id, name, ok: false, detail: "skipped: chart not drawn" };
    else r = await timed(id, name, pc, fn);
    results.push(r);
    onProgress(r);
  }
  return { results, pc };
}

export { TAG };
