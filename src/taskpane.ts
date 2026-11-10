// Side panel: insert charts, edit the selected chart's data and options, settings, diagnostics.

import { layoutCategory } from "./engine/category";
import { canvasMeasure } from "./engine/measure";
import { calendarColumns } from "./engine/calendar";
import { parseDate } from "./engine/dates";
import { defaultStyle, describe, HIGHLIGHTS, PALETTE_A, sampleCalendar, sampleCategory, sampleGantt } from "./engine/spec";
import { toSvg } from "./engine/svg";
import type { CalendarSpec, CategorySpec, ChartSpec, Frame, GanttSpec, GanttTier, Mode } from "./engine/types";
import { FALLBACK_FONTS, runPilot, type ProbeResult } from "./ppt/probes";
import {
  APP_VERSION, detectLineBreak, drawChart, errText, findSelectedChart, inkFor, readChart, readSlideBackground, redrawChart, setLineBreak, TAG,
  type ChartRef,
} from "./ppt/render";
import { loadSettings, saveSettings, separators, styleFrom, type Settings } from "./ui/settings";
import { CalendarSheet } from "./ui/calendarSheet";
import { GanttSheet } from "./ui/ganttSheet";
import { Sheet } from "./ui/sheet";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const settings: Settings = loadSettings();
const state: {
  ref: ChartRef | null;
  busy: boolean;
  pending: boolean;
  calib: { font?: string; ratio?: number; lineBreak?: string } | null;
  ink: Map<string, ReturnType<typeof inkFor>>;
  probe: { results: ProbeResult[]; actions: Record<string, unknown>[] };
} = { ref: null, busy: false, pending: false, calib: null, ink: new Map(), probe: { results: [], actions: [] } };

// ---------- Helpers ----------

let toastTimer: number | undefined;
function toast(msg: string, ms = 5000) {
  const t = $("toast");
  t.textContent = msg;
  t.hidden = false;
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => (t.hidden = true), ms);
}

function status(msg: string) {
  $("ed-status").textContent = msg;
}

function measure() {
  return canvasMeasure(`"${state.calib?.font ?? "Aptos"}", ${FALLBACK_FONTS}`, state.calib?.ratio ?? 1);
}

/** Measures the presentation's body font once per session (temporary off-slide text boxes). */
async function ensureCalibrated(slideId: string) {
  if (state.calib) return;
  state.calib = {};
  try {
    const samples = ["Revenue 2024", "1,234.5", "North America", "Q3 2025E"];
    const out = await PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.slides.getItem(slideId).shapes;
      const boxes = samples.map((t, k) => {
        const s = shapes.addTextBox(t, { left: -2000, top: k * 30, width: 400, height: 20 });
        s.textFrame.wordWrap = false;
        s.textFrame.leftMargin = 0;
        s.textFrame.rightMargin = 0;
        s.textFrame.textRange.font.size = 12;
        s.textFrame.autoSizeSetting = "AutoSizeShapeToFitText";
        return s;
      });
      await ctx.sync();
      boxes.forEach((b) => { b.load("width"); b.textFrame.textRange.font.load("name"); });
      await ctx.sync();
      const font = boxes[0].textFrame.textRange.font.name ?? undefined;
      const widths = boxes.map((b) => b.width);
      boxes.forEach((b) => b.delete());
      await ctx.sync();
      return { font, widths };
    });
    const m = canvasMeasure(`"${out.font}", ${FALLBACK_FONTS}`);
    const ratios = samples.map((t, k) => out.widths[k] / m(t, 12)).filter((r) => r > 0.5 && r < 2);
    state.calib = { font: out.font, ratio: ratios.length ? ratios.reduce((a, b) => a + b, 0) / ratios.length : 1 };
    const lb = await detectLineBreak(slideId);
    setLineBreak(lb.sep);
    state.calib.lineBreak = JSON.stringify(lb.sep);
  } catch {
    /* keep defaults */
  }
}

async function styleForSlide(slideId: string) {
  let ink = state.ink.get(slideId);
  if (!ink) {
    ink = inkFor(await readSlideBackground(slideId));
    state.ink.set(slideId, ink);
  }
  return styleFrom(settings, ink);
}

// ---------- Editor ----------

const sheet = new Sheet($("sheet"), {
  decimalSep: () => separators(settings).decimalSep,
  combo: () => $<HTMLSelectElement>("ed-type").value.startsWith("combo"),
  palette: () => PALETTE_A,
  onChange: () => scheduleApply(),
});

const gsheet = new GanttSheet($("gsheet"), {
  palette: () => PALETTE_A,
  dateOrder: () => settings.dateOrder,
  onChange: () => scheduleApply(),
});

function ganttFromEditor(): GanttSpec {
  return {
    v: 2,
    type: "gantt",
    rows: gsheet.data(),
    tiers: selectedTiers(),
    headerStyle: $<HTMLSelectElement>("g-hstyle").value as GanttSpec["headerStyle"],
    headerColor: $<HTMLSelectElement>("g-hcolor").value || undefined,
    shadeWeekends: $<HTMLInputElement>("g-weekends").checked,
    showToday: $<HTMLInputElement>("g-today").checked,
    milestoneDates: $<HTMLInputElement>("g-msdates").checked,
    dateOrder: settings.dateOrder,
  };
}

const tierBoxes = () => Array.from(document.querySelectorAll<HTMLInputElement>("#ed-gantt input.g-tier"));

function selectedTiers(): GanttSpec["tiers"] {
  if ($<HTMLInputElement>("g-auto").checked) return "auto";
  const t = tierBoxes().filter((b) => b.checked).map((b) => b.value as GanttTier);
  return t.length ? t : "auto";
}

function showTiers(t: GanttSpec["tiers"]) {
  $<HTMLInputElement>("g-auto").checked = t === "auto";
  tierBoxes().forEach((b) => (b.checked = t !== "auto" && t.includes(b.value as GanttTier)));
}

const calTierBoxes = () => Array.from(document.querySelectorAll<HTMLInputElement>("#ed-cal input.c-tier"));

function calFromEditor(): CalendarSpec {
  const auto = $<HTMLInputElement>("c-auto").checked;
  const t = calTierBoxes().filter((b) => b.checked).map((b) => b.value as GanttTier);
  return {
    v: 2,
    type: "calendar",
    start: $<HTMLInputElement>("c-start").value,
    end: $<HTMLInputElement>("c-end").value,
    tiers: auto || !t.length ? "auto" : t,
    headerStyle: $<HTMLSelectElement>("c-hstyle").value as CalendarSpec["headerStyle"],
    headerColor: $<HTMLSelectElement>("c-hcolor").value || undefined,
    shadeWeekends: $<HTMLInputElement>("c-weekends").checked,
    showLabels: $<HTMLInputElement>("c-labels").checked,
    rows: csheet.data(),
    dateOrder: settings.dateOrder,
  };
}

const csheet = new CalendarSheet($("crows"), $("ccells"), {
  columns: () => calendarColumns(calFromEditor())?.columns ?? [],
  dateOrder: () => settings.dateOrder,
  onChange: () => scheduleApply(),
  onRange: (start, end) => {
    $<HTMLInputElement>("c-start").value = start;
    $<HTMLInputElement>("c-end").value = end;
  },
});

function openType(): ChartSpec["type"] | null {
  return state.ref?.stored.spec.type ?? null;
}

function specFromEditor(): ChartSpec {
  const t = openType();
  return t === "gantt" ? ganttFromEditor() : t === "calendar" ? calFromEditor() : categoryFromEditor();
}

function categoryFromEditor(): CategorySpec {
  const [type, mode] = $<HTMLSelectElement>("ed-type").value.split(":") as [CategorySpec["type"], Mode];
  const d = sheet.data();
  const dec = $<HTMLSelectElement>("ed-dec").value;
  return {
    v: 2, type, mode,
    categories: d.categories,
    series: d.series,
    labels: {
      values: $<HTMLInputElement>("ed-values").checked,
      totals: $<HTMLInputElement>("ed-totals").checked,
      legend: $<HTMLSelectElement>("ed-legend").value as CategorySpec["labels"]["legend"],
    },
    decimals: dec === "" ? null : Number(dec),
  };
}

function syncOptionStates() {
  const [type, mode] = $<HTMLSelectElement>("ed-type").value.split(":");
  $<HTMLInputElement>("ed-totals").disabled = mode === "clustered" || type === "line";
}

function openEditor(ref: ChartRef) {
  state.ref = ref;
  const spec = ref.stored.spec;
  $("editor").hidden = false;
  $("ed-cat").hidden = spec.type === "gantt" || spec.type === "calendar";
  $("ed-gantt").hidden = spec.type !== "gantt";
  $("ed-cal").hidden = spec.type !== "calendar";
  $("ed-title").textContent = describe(spec);
  if (spec.type === "calendar") {
    $<HTMLInputElement>("c-start").value = spec.start;
    $<HTMLInputElement>("c-end").value = spec.end;
    $<HTMLInputElement>("c-auto").checked = spec.tiers === "auto";
    calTierBoxes().forEach((b) => (b.checked = spec.tiers !== "auto" && spec.tiers.includes(b.value as GanttTier)));
    $<HTMLSelectElement>("c-hstyle").value = spec.headerStyle;
    $<HTMLSelectElement>("c-hcolor").value = spec.headerColor ?? "";
    $<HTMLInputElement>("c-weekends").checked = spec.shadeWeekends;
    $<HTMLInputElement>("c-labels").checked = spec.showLabels;
    csheet.load(spec.rows);
    status("");
    return;
  }
  if (spec.type === "gantt") {
    showTiers(spec.tiers);
    $<HTMLSelectElement>("g-hstyle").value = spec.headerStyle;
    $<HTMLSelectElement>("g-hcolor").value = spec.headerColor ?? "";
    $<HTMLInputElement>("g-weekends").checked = spec.shadeWeekends;
    $<HTMLInputElement>("g-today").checked = spec.showToday;
    $<HTMLInputElement>("g-msdates").checked = spec.milestoneDates;
    gsheet.load(spec.rows);
    status("");
    return;
  }
  $<HTMLSelectElement>("ed-type").value = `${spec.type}:${spec.type === "combo" && spec.mode === "percent" ? "stacked" : spec.type === "line" ? "stacked" : spec.mode}`;
  $<HTMLInputElement>("ed-values").checked = spec.labels.values;
  $<HTMLInputElement>("ed-totals").checked = spec.labels.totals;
  $<HTMLSelectElement>("ed-legend").value = spec.labels.legend;
  $<HTMLSelectElement>("ed-dec").value = spec.decimals === null ? "" : String(spec.decimals);
  syncOptionStates();
  sheet.load({ categories: spec.categories, series: spec.series });
  status("");
}

function closeEditor() {
  state.ref = null;
  $("editor").hidden = true;
}

let applyTimer: number | undefined;
function scheduleApply(delay = 600) {
  if (!state.ref) return;
  status("…");
  window.clearTimeout(applyTimer);
  applyTimer = window.setTimeout(apply, delay);
}

async function apply() {
  if (!state.ref) return;
  if (state.busy) {
    state.pending = true;
    return;
  }
  state.busy = true;
  const t0 = performance.now();
  try {
    const spec = specFromEditor();
    const fresh = await readChart(state.ref.slideId, state.ref.groupId);
    if (!fresh) {
      closeEditor();
      toast("That chart no longer exists (deleted or undone). Select a chart to edit it.");
      return;
    }
    await ensureCalibrated(fresh.slideId);
    const r = await redrawChart(fresh, spec, await styleForSlide(fresh.slideId), measure());
    state.ref = { slideId: fresh.slideId, groupId: r.groupId, stored: r.stored, current: r.stored.drawn };
    $("ed-title").textContent = describe(spec);
    status(r.warnings.length ? `⚠ ${r.warnings.join("; ")}` : `Updated · ${(performance.now() - t0) / 1000 < 1 ? "<1" : ((performance.now() - t0) / 1000).toFixed(1)} s`);
  } catch (e) {
    status("Update failed");
    toast(`Update failed: ${errText(e)}`, 10000);
  } finally {
    state.busy = false;
    if (state.pending) {
      state.pending = false;
      void apply();
    }
  }
}

// ---------- Insert ----------

async function insertChart(type: ChartSpec["type"], mode: Mode) {
  if (state.busy) return;
  state.busy = true;
  try {
    const target = await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items/id");
      const sel = ctx.presentation.getSelectedShapes();
      sel.load("items/id,items/left,items/top,items/width,items/height,items/level");
      const ps = ctx.presentation.pageSetup;
      ps.load("slideWidth,slideHeight");
      await ctx.sync();
      if (!slides.items.length) throw new Error("Select a slide first.");
      let frame: Frame = { left: ps.slideWidth * 0.05, top: ps.slideHeight * 0.2, width: ps.slideWidth * 0.9, height: ps.slideHeight * 0.7 };
      const s = sel.items[0];
      if (sel.items.length === 1 && s.level === 0 && s.width > 60 && s.height > 40) {
        const tag = s.tags.getItemOrNullObject(TAG);
        await ctx.sync();
        if (tag.isNullObject) frame = { left: s.left, top: s.top, width: s.width, height: s.height };
      }
      return { slideId: slides.items[0].id, frame };
    });
    await ensureCalibrated(target.slideId);
    const spec: ChartSpec = type === "gantt" ? sampleGantt(settings.dateOrder) : type === "calendar" ? sampleCalendar(settings.dateOrder) : sampleCategory(type, mode);
    const r = await drawChart(target.slideId, spec, target.frame, await styleForSlide(target.slideId), measure());
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(target.slideId).setSelectedShapes([r.groupId]);
      await ctx.sync();
    });
    openEditor({ slideId: target.slideId, groupId: r.groupId, stored: r.stored, current: r.drawn });
    status("Inserted with sample data — paste or type yours below.");
    if (type === "gantt") gsheet.focusCell(1, "label");
    else if (type === "calendar") { /* data sheet already shown */ }
    else sheet.focusCell(1, 1);
  } catch (e) {
    toast(`Insert failed: ${errText(e)}`, 10000);
  } finally {
    state.busy = false;
  }
}

// ---------- Selection ----------

let selTimer: number | undefined;
function onSelectionChanged() {
  window.clearTimeout(selTimer);
  selTimer = window.setTimeout(async () => {
    if (state.busy) return;
    try {
      const ref = await findSelectedChart();
      if (!ref) {
        if (!$("editor").contains(document.activeElement)) closeEditor();
        return;
      }
      if (state.ref && ref.groupId === state.ref.groupId) return;
      openEditor(ref);
    } catch {
      /* selection of unsupported objects: ignore */
    }
  }, 200);
}

// ---------- Settings ----------

function initSettings() {
  $<HTMLSelectElement>("set-font").value = String(settings.fontSize);
  $<HTMLSelectElement>("set-num").value = settings.numberFormat;
  $<HTMLSelectElement>("set-dates").value = settings.dateOrder;
  $("set-palette").innerHTML = [...PALETTE_A, ...HIGHLIGHTS].map((c) => `<span class="chip" style="background:${c}" title="${c}"></span>`).join("");
  const onChange = () => {
    // Read the sheet with the old decimal separator before switching formats.
    const data = openType() && openType() !== "gantt" && openType() !== "calendar" ? sheet.data() : null;
    settings.fontSize = Number($<HTMLSelectElement>("set-font").value);
    settings.numberFormat = $<HTMLSelectElement>("set-num").value as Settings["numberFormat"];
    settings.dateOrder = $<HTMLSelectElement>("set-dates").value as Settings["dateOrder"];
    saveSettings(settings);
    if (openType() === "gantt" || openType() === "calendar") {
      if (openType() === "gantt") gsheet.render(); // re-validate dates with the new order
      else csheet.renderCells();
      scheduleApply(0);
      return;
    }
    if (data) {
      sheet.load(data); // re-render cells with the new decimal separator
      scheduleApply(0);
    }
  };
  $("set-font").addEventListener("change", onChange);
  $("set-num").addEventListener("change", onChange);
  $("set-dates").addEventListener("change", onChange);
}

// ---------- Diagnostics ----------

function updateReport() {
  const d = typeof Office !== "undefined" && Office.context?.diagnostics;
  $<HTMLTextAreaElement>("report").value = JSON.stringify({
    app: APP_VERSION,
    when: new Date().toISOString(),
    office: d ? `${d.host} ${d.platform} ${d.version}` : "n/a",
    calibration: state.calib,
    results: state.probe.results.map((r) => ({ id: r.id, ok: r.ok, ms: r.ms, detail: r.detail })),
    actions: state.probe.actions,
  }, null, 1);
}

async function onRunChecks() {
  $("results").innerHTML = "";
  state.probe.results = [];
  const btn = $<HTMLButtonElement>("run");
  btn.disabled = true;
  try {
    await runPilot((r) => {
      state.probe.results.push(r);
      const li = document.createElement("li");
      const cls = r.ok === true ? "ok" : r.ok === false ? "bad" : "info";
      li.innerHTML = `<span class="st ${cls}">${r.ok === true ? "✓" : r.ok === false ? "✗" : "◐"}</span>${r.name}<span class="d"></span>`;
      (li.querySelector(".d") as HTMLElement).textContent = r.detail;
      $("results").appendChild(li);
      updateReport();
    });
  } catch (e) {
    state.probe.actions.push({ action: "run", error: errText(e) });
  } finally {
    btn.disabled = false;
    updateReport();
  }
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

// ---------- Preview outside Office ----------

function showPreview() {
  const frame: Frame = { left: 0, top: 0, width: 520, height: 300 };
  const m = canvasMeasure(FALLBACK_FONTS);
  const kinds: [CategorySpec["type"], Mode, string][] = [
    ["column", "stacked", "Stacked"], ["column", "clustered", "Clustered"], ["column", "percent", "100%"],
    ["bar", "stacked", "Stacked bar"], ["line", "stacked", "Line"], ["combo", "stacked", "Combination"],
  ];
  $("preview").innerHTML = kinds
    .map(([t, mo, name]) => `<figure><figcaption>${name}</figcaption>${toSvg(layoutCategory(sampleCategory(t, mo), frame, defaultStyle(), m).prims, frame, FALLBACK_FONTS, m)}</figure>`)
    .join("");
}

// ---------- Wiring ----------

function wire() {
  $("version").textContent = `v${APP_VERSION}`;
  document.querySelectorAll<HTMLButtonElement>("#insert .types button").forEach((b) =>
    b.addEventListener("click", () => insertChart(b.dataset.type as ChartSpec["type"], b.dataset.mode as Mode)),
  );
  $("ed-type").addEventListener("change", () => {
    syncOptionStates();
    sheet.render();
    scheduleApply(0);
  });
  for (const id of ["ed-values", "ed-totals", "ed-legend", "ed-dec"]) $(id).addEventListener("change", () => scheduleApply(0));
  $("add-series").addEventListener("click", () => sheet.addSeries());
  $("add-cat").addEventListener("click", () => sheet.addCategory());
  $("del-series").addEventListener("click", () => { const m = sheet.deleteSeries(); if (m) toast(m); });
  $("del-cat").addEventListener("click", () => { const m = sheet.deleteCategory(); if (m) toast(m); });
  $("transpose").addEventListener("click", () => sheet.transpose());
  for (const id of ["g-today", "g-msdates", "g-hstyle", "g-hcolor", "g-weekends"]) $(id).addEventListener("change", () => scheduleApply(0));
  $("g-auto").addEventListener("change", () => {
    if ($<HTMLInputElement>("g-auto").checked) tierBoxes().forEach((b) => (b.checked = false));
    else if (!tierBoxes().some((b) => b.checked)) $<HTMLInputElement>("g-auto").checked = true;
    scheduleApply(0);
  });
  tierBoxes().forEach((b) => b.addEventListener("change", () => {
    $<HTMLInputElement>("g-auto").checked = !tierBoxes().some((x) => x.checked);
    scheduleApply(0);
  }));
  // Calendar controls: date range and tiers change the columns of the text grid.
  const calRefresh = () => {
    for (const id of ["c-start", "c-end"]) {
      const el = $<HTMLInputElement>(id);
      el.classList.toggle("bad", el.value.trim() !== "" && parseDate(el.value, settings.dateOrder) === null);
    }
    csheet.renderCells();
    scheduleApply();
  };
  $("c-start").addEventListener("input", calRefresh);
  $("c-end").addEventListener("input", calRefresh);
  for (const id of ["c-hstyle", "c-hcolor", "c-weekends", "c-labels"]) $(id).addEventListener("change", () => scheduleApply(0));
  $("c-auto").addEventListener("change", () => {
    if ($<HTMLInputElement>("c-auto").checked) calTierBoxes().forEach((b) => (b.checked = false));
    else if (!calTierBoxes().some((b) => b.checked)) $<HTMLInputElement>("c-auto").checked = true;
    csheet.renderCells();
    scheduleApply(0);
  });
  calTierBoxes().forEach((b) => b.addEventListener("change", () => {
    $<HTMLInputElement>("c-auto").checked = !calTierBoxes().some((x) => x.checked);
    csheet.renderCells();
    scheduleApply(0);
  }));
  $("c-addrow").addEventListener("click", () => csheet.addRow());
  $("g-add").addEventListener("click", () => gsheet.addRow());
  $("g-del").addEventListener("click", () => { const m = gsheet.deleteRow(); if (m) toast(m); });
  $("g-up").addEventListener("click", () => gsheet.move(-1));
  $("g-down").addEventListener("click", () => gsheet.move(1));
  $("run").addEventListener("click", onRunChecks);
  $("copy").addEventListener("click", onCopy);
  initSettings();
  updateReport();
}

wire();
if (typeof Office === "undefined") {
  $("not-office").hidden = false;
  showPreview();
} else {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.PowerPoint) {
      $("not-office").hidden = false;
      showPreview();
      return;
    }
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    onSelectionChanged(); // a chart may already be selected when the panel opens
  });
}
