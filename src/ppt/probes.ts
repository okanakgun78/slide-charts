// Pilot checks: each probe exercises one capability the full product depends on and records
// what PowerPoint on this machine actually does. Results are shown in the pane and copied
// into a report the user sends back.

import { DEFAULT_PALETTE, normHex } from "../engine/format";
import { canvasMeasure } from "../engine/measure";
import type { Frame, StackedSpec, Style } from "../engine/types";
import { addPrim, drawChart, readThemeAccents, TAG } from "./render";

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

export const SAMPLE_SPEC: StackedSpec = {
  v: 1,
  type: "stacked",
  categories: ["2021", "2022", "2023", "2024", "2025E"],
  series: [
    { name: "Europe", values: [42, 46, 51, 55, 61] },
    { name: "North America", values: [31, 33, 38, 41, 47] },
    { name: "Asia", values: [12, 16, 21, 27, 34] },
    { name: "Rest of world", values: [5, 6, 6, 8, 9] },
  ],
  showTotals: true,
};

export function chartStyle(palette: string[]): Style {
  return { fontSize: 12, textColor: "#1A1A1A", axisColor: "#7F7F7F", palette };
}

export const FALLBACK_FONTS = '"Aptos", "Calibri", "Helvetica Neue", Arial, sans-serif';

const API_SETS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"];

type Probe = (pc: PilotContext) => Promise<Omit<ProbeResult, "id" | "name">>;

async function timed(id: string, name: string, pc: PilotContext, fn: Probe): Promise<ProbeResult> {
  const t0 = performance.now();
  try {
    const r = await fn(pc);
    return { id, name, ...r, ms: Math.round(performance.now() - t0) };
  } catch (e) {
    const err = e as { code?: string; message?: string; debugInfo?: { errorLocation?: string } };
    const where = err.debugInfo?.errorLocation ? ` @ ${err.debugInfo.errorLocation}` : "";
    return { id, name, ok: false, detail: `${err.code ?? "Error"}: ${err.message ?? String(e)}${where}`, ms: Math.round(performance.now() - t0) };
  }
}

const run = <T>(fn: (ctx: PowerPoint.RequestContext) => Promise<T>) => PowerPoint.run(fn);

const probes: [string, string, Probe][] = [
  ["env", "Environment and API versions", async () => {
    const d = Office.context.diagnostics;
    const sets = API_SETS.map((v) => `${v}:${Office.context.requirements.isSetSupported("PowerPointApi", v) ? "yes" : "NO"}`);
    const ok = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
    return { ok, detail: `${d.host} ${d.platform} ${d.version} | ${sets.join(" ")}` };
  }],

  ["slide", "Add a test slide and go to it", async (pc) => run(async (ctx) => {
    let layoutId: string | undefined;
    const master = ctx.presentation.slideMasters.getItemAt(0);
    master.layouts.load("items/id,items/name");
    await ctx.sync();
    const blank = master.layouts.items.find((l) => /blank|leer|vide/i.test(l.name));
    if (blank) layoutId = blank.id;
    ctx.presentation.slides.add(layoutId ? { layoutId } : undefined);
    ctx.presentation.slides.load("items/id");
    await ctx.sync();
    const slide = ctx.presentation.slides.items[ctx.presentation.slides.items.length - 1];
    pc.slideId = slide.id;
    ctx.presentation.setSelectedSlides([slide.id]);
    await ctx.sync();
    return { ok: true, detail: `slide ${slide.id}; layout: ${blank ? blank.name : "default (no Blank layout found)"}` };
  })],

  ["page", "Read slide size", async (pc) => run(async (ctx) => {
    const ps = ctx.presentation.pageSetup;
    ps.load("slideWidth,slideHeight");
    await ctx.sync();
    pc.slideW = ps.slideWidth;
    pc.slideH = ps.slideHeight;
    return { ok: true, detail: `${ps.slideWidth} x ${ps.slideHeight} pt` };
  })],

  ["theme", "Read template theme colours", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const raw = await readThemeAccents(ctx, slide);
    const hex = raw.map(normHex);
    if (hex.every((h) => h)) pc.palette = hex as string[];
    return { ok: hex.every((h) => h), detail: `accents raw=${JSON.stringify(raw)}` };
  })],

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
    const font = boxes[0].textFrame.textRange.font.name;
    pc.fontName = font ?? undefined;
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
      detail: `font="${font}" visibleToPane=${fontVisible} autosizeReadBack=${autosized} ` +
        rows.map((r) => `[${r.t}: ppt=${r.ppt.toFixed(1)} canvas=${r.canvas.toFixed(1)}]`).join(" "),
    };
  })],

  ["chart", "Draw sample stacked chart (timed)", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const frame: Frame = { left: pc.slideW * 0.06, top: pc.slideH * 0.12, width: pc.slideW * 0.58, height: pc.slideH * 0.62 };
    const m = canvasMeasure(`"${pc.fontName ?? "Aptos"}", ${FALLBACK_FONTS}`, pc.measureRatio ?? 1);
    const r = await drawChart(ctx, slide, SAMPLE_SPEC, frame, chartStyle(pc.palette), m);
    pc.chartGroupId = r.groupId;
    return {
      ok: true,
      detail: `${r.shapeCount} shapes; draw ${Math.round(r.msShapes)} ms, group+tag ${Math.round(r.msGroupTag)} ms` +
        (r.warnings.length ? `; warnings: ${r.warnings.join(", ")}` : ""),
    };
  })],

  ["tag", "Read chart data back from the file", async (pc) => run(async (ctx) => {
    const g = ctx.presentation.slides.getItem(pc.slideId!).shapes.getItem(pc.chartGroupId!);
    const tag = g.tags.getItemOrNullObject(TAG);
    tag.load("value");
    g.load("type");
    await ctx.sync();
    if (tag.isNullObject) return { ok: false, detail: "tag missing" };
    const back = JSON.parse(tag.value);
    const ok = JSON.stringify(back.spec) === JSON.stringify(SAMPLE_SPEC);
    return { ok, detail: `type=${g.type}; ${tag.value.length} chars; spec round-trip ${ok ? "identical" : "DIFFERENT"}` };
  })],

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

  ["harvey", "Harvey ball via shape adjustments (75%)", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.14, d = 28;
    const ring = slide.shapes.addGeometricShape("Ellipse", { left: x, top: y, width: d, height: d });
    ring.fill.setSolidColor("#FFFFFF");
    ring.lineFormat.color = pc.palette[0];
    ring.lineFormat.weight = 1;
    const pie = slide.shapes.addGeometricShape("Pie", { left: x, top: y, width: d, height: d });
    pie.fill.setSolidColor(pc.palette[0]);
    pie.lineFormat.visible = false;
    pie.adjustments.load("count");
    await ctx.sync();
    const n = pie.adjustments.count;
    const before = Array.from({ length: n }, (_, k) => pie.adjustments.get(k));
    await ctx.sync();
    const vals = before.map((r) => r.value);
    // Default pie is 270 degrees from 3 o'clock. Detect units (degrees vs 60000ths) from the defaults.
    const unit = vals.some((v) => Math.abs(v) > 1000) ? 60000 : 1;
    pie.adjustments.set(0, 270 * unit); // start at 12 o'clock
    pie.adjustments.set(1, 180 * unit); // end at 9 o'clock -> 75% clockwise
    const label = slide.shapes.addTextBox("< should be 3/4 filled, from 12 o'clock clockwise", { left: x + d + 6, top: y + 6, width: 260, height: 16 });
    label.textFrame.textRange.font.size = 10;
    await ctx.sync();
    const after = Array.from({ length: n }, (_, k) => pie.adjustments.get(k));
    await ctx.sync();
    return { ok: null, detail: `count=${n} defaults=${JSON.stringify(vals)} after=${JSON.stringify(after.map((r) => r.value))} (check visually)` };
  })],

  ["diag", "Diagonal line via rotation + arrowhead", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.30;
    const line = addPrim(slide.shapes, { kind: "line", x1: x, y1: y + 50, x2: x + 70, y2: y, color: "#C00000", width: 1.5, role: "probe-diag" });
    // Arrowhead: an upward triangle rotated to the line's direction, its tip on the line end.
    const theta = Math.atan2(-50, 70);
    const hw = 9, hh = 8;
    const cx = x + 70 - (hh / 2) * Math.cos(theta), cy = y - (hh / 2) * Math.sin(theta);
    const head = slide.shapes.addGeometricShape("Triangle", { left: cx - hw / 2, top: cy - hh / 2, width: hw, height: hh });
    head.fill.setSolidColor("#C00000");
    head.lineFormat.visible = false;
    head.rotation = 90 + (theta * 180) / Math.PI;
    const label = slide.shapes.addTextBox("< red line should rise left-to-right, arrow tip at top-right", { left: x + 80, top: y + 16, width: 230, height: 28 });
    label.textFrame.textRange.font.size = 10;
    line.load("rotation");
    await ctx.sync();
    return { ok: null, detail: `line rotation read back=${line.rotation.toFixed(1)} (expected -35.5; check visually)` };
  })],

  ["dash", "Dashed value line", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const x = pc.slideW * 0.72, y = pc.slideH * 0.50;
    const l = addPrim(slide.shapes, { kind: "line", x1: x, y1: y, x2: x + 120, y2: y, color: "#404040", width: 1, dash: true, role: "probe-dash" });
    l.lineFormat.load("dashStyle");
    await ctx.sync();
    return { ok: l.lineFormat.dashStyle === "Dash", detail: `dashStyle=${l.lineFormat.dashStyle}` };
  })],

  ["delete", "Delete a group removes its parts", async (pc) => run(async (ctx) => {
    const slide = ctx.presentation.slides.getItem(pc.slideId!);
    const a = slide.shapes.addGeometricShape("Rectangle", { left: 5, top: 5, width: 10, height: 10 });
    const b = slide.shapes.addGeometricShape("Rectangle", { left: 20, top: 5, width: 10, height: 10 });
    a.load("id"); b.load("id");
    const before = slide.shapes.getCount();
    await ctx.sync();
    const g = slide.shapes.addGroup([a.id, b.id]);
    await ctx.sync();
    g.delete();
    const after = slide.shapes.getCount();
    await ctx.sync();
    return { ok: after.value === before.value - 2, detail: `top-level shapes before=${before.value} after=${after.value}` };
  })],
];

export async function runPilot(onProgress: (r: ProbeResult) => void): Promise<{ results: ProbeResult[]; pc: PilotContext }> {
  const pc: PilotContext = { slideW: 960, slideH: 540, palette: DEFAULT_PALETTE };
  const results: ProbeResult[] = [];
  for (const [id, name, fn] of probes) {
    if (id !== "env" && id !== "slide" && !pc.slideId) {
      const r: ProbeResult = { id, name, ok: false, detail: "skipped: no test slide" };
      results.push(r); onProgress(r); continue;
    }
    if ((id === "tag" || id === "bigtag") && !pc.chartGroupId) {
      const r: ProbeResult = { id, name, ok: false, detail: "skipped: chart not drawn" };
      results.push(r); onProgress(r); continue;
    }
    const r = await timed(id, name, pc, fn);
    results.push(r);
    onProgress(r);
  }
  return { results, pc };
}
