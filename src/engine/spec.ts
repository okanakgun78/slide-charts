// Chart spec defaults, samples, migration and the house style.

import type { CategorySpec, ChartSpec, GanttSpec, Mode, Style } from "./types";

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
  return s as unknown as ChartSpec;
}

export function isGantt(spec: ChartSpec): spec is GanttSpec {
  return spec.type === "gantt";
}

/** Human-readable chart name for the panel. */
export function describe(spec: ChartSpec): string {
  if (spec.type === "gantt") return `Gantt (${spec.rows.length} rows)`;
  const modes: Record<Mode, string> = { stacked: "stacked", clustered: "clustered", percent: "100%" };
  const t = spec.type === "line" ? "line" : spec.type === "combo" ? `combination (${modes[spec.mode]})` : `${modes[spec.mode]} ${spec.type}`;
  return `${t[0].toUpperCase()}${t.slice(1)} · ${spec.categories.length} categories × ${spec.series.length} series`;
}
