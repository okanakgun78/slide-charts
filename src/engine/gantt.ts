// Gantt/timeline layout (implemented in v0.4).
import type { Frame, GanttSpec, Layout, Measure, Style } from "./types";

export function layoutGantt(_spec: GanttSpec, _frame: Frame, _style: Style, _measure: Measure): Layout {
  return { prims: [], warnings: ["Gantt charts arrive in v0.4"] };
}
