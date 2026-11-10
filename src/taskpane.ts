import { DEFAULT_PALETTE } from "./engine/format";
import { approxMeasure, canvasMeasure } from "./engine/measure";
import { layoutStacked } from "./engine/stacked";
import { toSvg } from "./engine/svg";
import type { Frame, StackedSpec } from "./engine/types";
import { chartStyle, FALLBACK_FONTS, runPilot, SAMPLE_SPEC, type PilotContext, type ProbeResult } from "./ppt/probes";
import { APP_VERSION, drawChart, getSelectedChart, readThemeAccents, refitFrame } from "./ppt/render";

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
    ua: navigator.userAgent,
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

function measureFn() {
  const pc = state.pc;
  return canvasMeasure(`"${pc?.fontName ?? "Aptos"}", ${FALLBACK_FONTS}`, pc?.measureRatio ?? 1);
}

function errText(e: unknown): string {
  const err = e as { code?: string; message?: string };
  return `${err.code ?? "Error"}: ${err.message ?? String(e)}`;
}

/** New numbers: each value moves by up to ±25%, rounded like the input. */
function perturb(spec: StackedSpec): StackedSpec {
  return {
    ...spec,
    series: spec.series.map((s) => ({
      ...s,
      values: s.values.map((v) => (v === null ? null : Math.max(1, Math.round(v * (0.75 + Math.random() * 0.5))))),
    })),
  };
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
  try {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelectedChart(ctx);
      if (!sel) {
        $("selection").textContent = "Select a Slide Charts chart first.";
        logAction({ action: "redraw", result: "no chart selected" });
        return;
      }
      const moved =
        Math.abs(sel.current.left - sel.stored.drawn.left) > 0.5 || Math.abs(sel.current.top - sel.stored.drawn.top) > 0.5;
      const resized =
        Math.abs(sel.current.width - sel.stored.drawn.width) > 0.5 || Math.abs(sel.current.height - sel.stored.drawn.height) > 0.5;
      const frame: Frame = refitFrame(sel.stored, sel.current);
      let palette = state.pc?.palette ?? DEFAULT_PALETTE;
      try {
        palette = await readThemeAccents(ctx, sel.slide);
      } catch {
        /* keep fallback palette */
      }
      sel.shape.delete();
      await ctx.sync();
      const r = await drawChart(ctx, sel.slide, perturb(sel.stored.spec), frame, chartStyle(palette), measureFn());
      sel.slide.setSelectedShapes([r.groupId]);
      await ctx.sync();
      logAction({
        action: "redraw", moved, resized,
        before: sel.current, frameAfter: frame,
        ms: Math.round(performance.now() - t0),
      });
    });
  } catch (e) {
    logAction({ action: "redraw", error: errText(e), ms: Math.round(performance.now() - t0) });
  } finally {
    setBusy(false);
  }
}

async function onInsert() {
  setBusy(true);
  const t0 = performance.now();
  try {
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
      const ps = ctx.presentation.pageSetup;
      ps.load("slideWidth,slideHeight");
      await ctx.sync();
      let palette = DEFAULT_PALETTE;
      try {
        palette = await readThemeAccents(ctx, slide);
      } catch {
        /* keep fallback palette */
      }
      const frame: Frame = { left: ps.slideWidth * 0.2, top: ps.slideHeight * 0.2, width: ps.slideWidth * 0.6, height: ps.slideHeight * 0.6 };
      const r = await drawChart(ctx, slide, SAMPLE_SPEC, frame, chartStyle(palette), measureFn());
      slide.setSelectedShapes([r.groupId]);
      await ctx.sync();
      logAction({ action: "insert", shapes: r.shapeCount, ms: Math.round(performance.now() - t0) });
    });
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
      await PowerPoint.run(async (ctx) => {
        const sel = await getSelectedChart(ctx);
        $("selection").textContent = sel
          ? `Selection: chart (${sel.stored.spec.type}, ${sel.stored.spec.categories.length} categories × ${sel.stored.spec.series.length} series)`
          : "Selection: not a Slide Charts chart";
      });
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
  const m = typeof document !== "undefined" ? canvasMeasure(FALLBACK_FONTS) : approxMeasure;
  const { prims } = layoutStacked(SAMPLE_SPEC, frame, chartStyle(DEFAULT_PALETTE), m);
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
