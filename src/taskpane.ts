import { DEFAULT_PALETTE, normHex } from "./engine/format";
import { canvasMeasure } from "./engine/measure";
import { layoutStacked } from "./engine/stacked";
import { toSvg } from "./engine/svg";
import type { Frame } from "./engine/types";
import {
  chartStyle, FALLBACK_FONTS, measureFor, perturb, runPilot, SAMPLE_SPEC, type PilotContext, type ProbeResult,
} from "./ppt/probes";
import { APP_VERSION, drawChart, errText, findSelectedChart, readThemeAccents, redrawChart } from "./ppt/render";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const state: {
  pc?: PilotContext;
  results: ProbeResult[];
  actions: Record<string, unknown>[];
} = { results: [], actions: [] };

function setBusy(busy: boolean) {
  for (const id of ["run", "redraw", "insert"]) $<HTMLButtonElement>(id).disabled = busy;
}

function renderResult(r: ProbeResult) {
  const li = document.createElement("li");
  const cls = r.ok === true ? "ok" : r.ok === false ? "bad" : "info";
  const mark = r.ok === true ? "✓" : r.ok === false ? "✗" : "◐";
  li.innerHTML = `<span class="st ${cls}">${mark}</span>${r.name}${r.ms !== undefined ? ` <span class="muted">(${r.ms} ms)</span>` : ""}<span class="d"></span>`;
  (li.querySelector(".d") as HTMLElement).textContent = r.detail;
  $("results").appendChild(li);
}

function updateReport() {
  const d = typeof Office !== "undefined" && Office.context?.diagnostics;
  const report = {
    app: APP_VERSION,
    when: new Date().toISOString(),
    office: d ? `${d.host} ${d.platform} ${d.version}` : "n/a",
    calibration: state.pc ? { font: state.pc.fontName, measureRatio: state.pc.measureRatio, palette: state.pc.palette } : null,
    results: state.results.map((r) => ({ id: r.id, ok: r.ok, ms: r.ms, detail: r.detail })),
    actions: state.actions,
  };
  $<HTMLTextAreaElement>("report").value = JSON.stringify(report, null, 1);
}

function logAction(a: Record<string, unknown>) {
  state.actions.push({ at: new Date().toISOString().slice(11, 19), ...a });
  updateReport();
}

async function themePalette(slideId: string): Promise<string[]> {
  try {
    const hex = (await readThemeAccents(slideId)).map(normHex);
    if (hex.every((h) => h)) return hex as string[];
  } catch {
    /* fall back below */
  }
  return state.pc?.palette ?? DEFAULT_PALETTE;
}

async function onRun() {
  setBusy(true);
  $("results").innerHTML = "";
  state.results = [];
  try {
    const { results, pc } = await runPilot((r) => {
      state.results.push(r);
      renderResult(r);
      updateReport();
    });
    state.pc = pc;
    state.results = results;
  } catch (e) {
    logAction({ action: "run", error: errText(e) });
  } finally {
    updateReport();
    setBusy(false);
  }
}

async function onRedraw() {
  setBusy(true);
  const t0 = performance.now();
  let step = "find selected chart";
  try {
    const ref = await findSelectedChart();
    if (!ref) {
      $("selection").textContent = "Select a Slide Charts chart first.";
      logAction({ action: "redraw", result: "no chart selected" });
      return;
    }
    step = "read theme";
    const palette = await themePalette(ref.slideId);
    step = "redraw";
    const r = await redrawChart(ref, perturb(ref.stored.spec), chartStyle(palette), measureFor(state.pc));
    $("selection").textContent = "Redrawn ✓";
    logAction({
      action: "redraw", ok: true, moved: r.moved, resized: r.resized, deleteStrategy: r.deleteStrategy,
      before: ref.current, frameAfter: r.frame, log: r.log, ms: Math.round(performance.now() - t0),
    });
  } catch (e) {
    $("selection").textContent = `Redraw failed at "${step}" — details in the report.`;
    logAction({ action: "redraw", step, error: errText(e), ms: Math.round(performance.now() - t0) });
  } finally {
    setBusy(false);
  }
}

async function onInsert() {
  setBusy(true);
  const t0 = performance.now();
  try {
    const where = await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items/id");
      const ps = ctx.presentation.pageSetup;
      ps.load("slideWidth,slideHeight");
      await ctx.sync();
      return { slideId: slides.items[0].id, w: ps.slideWidth, h: ps.slideHeight };
    });
    const frame: Frame = { left: where.w * 0.2, top: where.h * 0.2, width: where.w * 0.6, height: where.h * 0.6 };
    const r = await drawChart(where.slideId, SAMPLE_SPEC, frame, chartStyle(await themePalette(where.slideId)), measureFor(state.pc));
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(where.slideId).setSelectedShapes([r.groupId]);
      await ctx.sync();
    });
    logAction({ action: "insert", shapes: r.shapeCount, ms: Math.round(performance.now() - t0) });
  } catch (e) {
    logAction({ action: "insert", error: errText(e), ms: Math.round(performance.now() - t0) });
  } finally {
    setBusy(false);
  }
}

let selTimer: number | undefined;
function onSelectionChanged() {
  window.clearTimeout(selTimer);
  selTimer = window.setTimeout(async () => {
    try {
      const ref = await findSelectedChart();
      $("selection").textContent = ref
        ? `Selection: chart (${ref.stored.spec.type}, ${ref.stored.spec.categories.length} categories × ${ref.stored.spec.series.length} series)`
        : "Selection: not a Slide Charts chart";
      if (!state.actions.some((a) => a.action === "selection-event")) logAction({ action: "selection-event", ok: true });
    } catch (e) {
      $("selection").textContent = `Selection: (${errText(e)})`;
    }
  }, 150);
}

async function onCopy() {
  const ta = $<HTMLTextAreaElement>("report");
  try {
    await navigator.clipboard.writeText(ta.value);
  } catch {
    ta.focus();
    ta.select();
    document.execCommand("copy");
  }
  $("copy").textContent = "Copied ✓";
  window.setTimeout(() => ($("copy").textContent = "Copy report"), 1500);
}

function showPreview() {
  const frame: Frame = { left: 0, top: 0, width: 560, height: 320 };
  const { prims } = layoutStacked(SAMPLE_SPEC, frame, chartStyle(DEFAULT_PALETTE), canvasMeasure(FALLBACK_FONTS));
  $("preview").innerHTML = toSvg(prims, frame);
}

function wire() {
  $("run").addEventListener("click", onRun);
  $("redraw").addEventListener("click", onRedraw);
  $("insert").addEventListener("click", onInsert);
  $("copy").addEventListener("click", onCopy);
  $("version").textContent = `v${APP_VERSION}`;
  updateReport();
}

wire();
if (typeof Office === "undefined") {
  $("not-office").hidden = false;
  setBusy(true);
  showPreview();
} else {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.PowerPoint) {
      $("not-office").hidden = false;
      setBusy(true);
      showPreview();
      return;
    }
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged, (res) => {
      logAction({ action: "selection-handler", registered: res.status === Office.AsyncResultStatus.Succeeded });
    });
  });
}
