// Chart spec defaults, samples, migration and the house style.

import { fromDay, isoDay, todayDay, toDay, type DateOrder } from "./dates";
import type { CalendarSpec, CategorySpec, ChartSpec, GanttSpec, Mode, Style } from "./types";

/** Option A: navy-to-light-blue ramp (from the template navy), greys after. */
export const PALETTE_A = ["#003760", "#2E5B85", "#5B82A8", "#8EAAC7", "#BFD0E2", "#7F7F7F", "#BFBFBF"];
/** Line colours: same family, ordered for contrast between neighbouring lines. */
export const LINE_PALETTE_A = ["#003760", "#5B82A8", "#CD9700", "#7F7F7F", "#C00000", "#8EAAC7"];
/** Highlight colours from the template, offered per series. */
export const HIGHLIGHTS = ["#CD9700", "#C00000"];

export function defaultStyle(): Style {
  return {
    fontSize: 12,
    textColor: "#1A1A1A",
    axisColor: "#7F7F7F",
    gridColor: "#D9D9D9",
    palette: PALETTE_A,
    linePalette: LINE_PALETTE_A,
    decimalSep: ".",
    thousandsSep: ",",
  };
}

export type CategoryKind = CategorySpec["type"];

export function sampleCategory(type: CategoryKind, mode: Mode): CategorySpec {
  const categories = ["2021", "2022", "2023", "2024", "2025E"];
  const base: CategorySpec = {
    v: 2,
    type,
    mode,
    categories,
    series: [
      { name: "Europe", values: [42, 46, 51, 55, 61] },
      { name: "North America", values: [31, 33, 38, 41, 47] },
      { name: "Asia", values: [12, 16, 21, 27, 34] },
    ],
    labels: { values: true, totals: mode !== "clustered", legend: "auto" },
    decimals: null,
  };
  if (type === "line") {
    base.series = [
      { name: "Revenue growth", values: [4.2, 5.1, 3.8, 6.0, 6.4] },
      { name: "Market growth", values: [3.1, 3.4, 2.9, 3.6, 3.9] },
    ];
    base.labels.totals = false;
  }
  if (type === "combo") {
    base.series = [
      { name: "Europe", values: [42, 46, 51, 55, 61] },
      { name: "North America", values: [31, 33, 38, 41, 47] },
      { name: "Target", values: [70, 80, 90, 100, 110], asLine: true },
    ];
  }
  return base;
}

/** A two-phase project starting this month, so the sample always looks current. */
export function sampleGantt(order: DateOrder, today = todayDay()): GanttSpec {
  const { y, m } = fromDay(today);
  const base = toDay(y, m, 1);
  const d = (k: number) => isoDay(base + k);
  return {
    v: 2,
    type: "gantt",
    rows: [
      { label: "Phase 1 – Diagnose", kind: "header", start: "", end: "" },
      { label: "Interviews and data request", kind: "task", start: d(0), end: d(20) },
      { label: "Baseline analysis", kind: "task", start: d(14), end: d(41) },
      { label: "Steering committee 1", kind: "milestone", start: d(42), end: "" },
      { label: "Phase 2 – Design", kind: "header", start: "", end: "" },
      { label: "Option development", kind: "task", start: d(42), end: d(76) },
      { label: "Business case", kind: "task", start: d(63), end: d(90) },
      { label: "Final steering committee", kind: "milestone", start: d(91), end: "" },
    ],
    tiers: "auto",
    headerStyle: "filled",
    shadeWeekends: true,
    showToday: true,
    milestoneDates: true,
    dateOrder: order,
  };
}

/** This week (Monday to Sunday) as a header-only calendar; rows are added by the user. */
export function sampleCalendar(order: DateOrder, today = todayDay()): CalendarSpec {
  const monday = today - fromDay(today).wd;
  return {
    v: 2,
    type: "calendar",
    start: isoDay(monday),
    end: isoDay(monday + 6),
    tiers: ["month", "day"],
    headerStyle: "filled",
    shadeWeekends: false,
    showLabels: false,
    dateOrder: order,
    rows: [],
  };
}

/** Upgrades older stored specs to the current version. */
export function migrate(raw: unknown): ChartSpec {
  const s = raw as Record<string, unknown>;
  if (s && s.v === 1 && s.type === "stacked") {
    return {
      v: 2,
      type: "column",
      mode: "stacked",
      categories: s.categories as string[],
      series: s.series as CategorySpec["series"],
      labels: { values: true, totals: s.showTotals !== false, legend: "auto" },
      decimals: typeof s.decimals === "number" ? s.decimals : null,
    };
  }
  if (s && s.type === "gantt" && !("tiers" in s)) {
    const g = s as unknown as GanttSpec;
    const map: Record<string, GanttSpec["tiers"]> = {
      auto: "auto", week: ["month", "week"], month: ["year", "month"], quarter: ["year", "quarter"], year: ["year"],
    };
    return { ...g, tiers: map[g.scale ?? "auto"] ?? "auto", headerStyle: "filled", shadeWeekends: true };
  }
  return s as unknown as ChartSpec;
}

export function isGantt(spec: ChartSpec): spec is GanttSpec {
  return spec.type === "gantt";
}

/** Human-readable chart name for the panel. */
export function describe(spec: ChartSpec): string {
  if (spec.type === "gantt") return `Gantt (${spec.rows.length} rows)`;
  if (spec.type === "calendar") return `Calendar (${spec.rows.length} row${spec.rows.length === 1 ? "" : "s"})`;
  const modes: Record<Mode, string> = { stacked: "stacked", clustered: "clustered", percent: "100%" };
  const t = spec.type === "line" ? "line" : spec.type === "combo" ? `combination (${modes[spec.mode]})` : `${modes[spec.mode]} ${spec.type}`;
  return `${t[0].toUpperCase()}${t.slice(1)} · ${spec.categories.length} categories × ${spec.series.length} series`;
}
