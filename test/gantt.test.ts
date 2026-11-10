import { describe, expect, it } from "vitest";
import { toDay } from "../src/engine/dates";
import { layoutGantt, periods, pickScale } from "../src/engine/gantt";
import { bbox } from "../src/engine/geometry";
import { approxMeasure } from "../src/engine/measure";
import { defaultStyle, sampleGantt } from "../src/engine/spec";
import type { GanttSpec, Prim, RectPrim, ShapePrim, TextPrim } from "../src/engine/types";

const style = defaultStyle();
const frame = { left: 40, top: 60, width: 860, height: 400 };
const today = toDay(2026, 11, 10);
const run = (s: GanttSpec) => layoutGantt(s, frame, style, approxMeasure, today);
const byRole = (ps: Prim[], role: string) => ps.find((p) => p.role === role);

describe("scale and periods", () => {
  it("picks a scale from the span", () => {
    expect(pickScale(60)).toBe("week");
    expect(pickScale(300)).toBe("month");
    expect(pickScale(1000)).toBe("quarter");
    expect(pickScale(547)).toBe("quarter");
    expect(pickScale(3000)).toBe("year");
  });
  it("weeks start on Monday and are grouped by month", () => {
    const { bottom, top } = periods(toDay(2026, 11, 4), toDay(2026, 12, 20), "week");
    expect(bottom[0].start).toBe(toDay(2026, 11, 2)); // Monday
    expect(top.map((t) => t.label)).toEqual(["Nov 2026", "Dec 2026"]);
  });
  it("quarters snap to quarter starts", () => {
    const { bottom, top } = periods(toDay(2026, 2, 10), toDay(2027, 5, 1), "quarter");
    expect(bottom.map((b) => b.label)).toEqual(["Q1", "Q2", "Q3", "Q4", "Q1", "Q2"]);
    expect(top.map((t) => t.label)).toEqual(["2026", "2027"]);
  });
});

describe("gantt layout", () => {
  const spec = sampleGantt("DMY", today);
  const { prims, warnings } = run(spec);

  it("lays out the sample inside the frame without warnings", () => {
    expect(warnings).toEqual([]);
    const b = bbox(prims);
    expect(b.left).toBeGreaterThanOrEqual(frame.left - 0.5);
    expect(b.top).toBeGreaterThanOrEqual(frame.top - 0.5);
    expect(b.left + b.width).toBeLessThanOrEqual(frame.left + frame.width + 0.5);
    expect(b.top + b.height).toBeLessThanOrEqual(frame.top + frame.height + 0.5);
  });

  it("draws tasks as bars, milestones as diamonds, headers as summary bars", () => {
    expect(prims.filter((p) => p.role.startsWith("bar:")).length).toBe(4);
    expect(prims.filter((p) => p.role.startsWith("ms:")).every((p) => (p as ShapePrim).shape === "Diamond")).toBe(true);
    expect(prims.filter((p) => p.role.startsWith("sum:")).length).toBe(2);
  });

  it("sizes bars by duration (end date inclusive)", () => {
    const a = byRole(prims, "bar:1") as RectPrim; // 21 days
    const b = byRole(prims, "bar:5") as RectPrim; // 35 days
    expect(b.w / a.w).toBeCloseTo(35 / 21, 6);
  });

  it("summary bar spans its tasks", () => {
    const sum = byRole(prims, "sum:0") as RectPrim;
    const t1 = byRole(prims, "bar:1") as RectPrim;
    const ms = byRole(prims, "ms:3") as ShapePrim;
    expect(sum.x).toBeCloseTo(t1.x, 6);
    expect(sum.x + sum.w).toBeGreaterThanOrEqual(ms.x + ms.w / 2 - 0.01);
  });

  it("labels milestone dates and the today line", () => {
    expect((byRole(prims, "note:3") as TextPrim).text).toMatch(/^\d{1,2} [A-Z][a-z]{2}$/);
    expect(byRole(prims, "today")).toBeTruthy();
  });

  it("warns about unreadable dates and keeps going", () => {
    const bad: GanttSpec = { ...spec, rows: [...spec.rows, { label: "Oops", kind: "task", start: "someday", end: "" }] };
    const r = run(bad);
    expect(r.warnings.some((w) => w.includes("Oops"))).toBe(true);
    expect(r.prims.length).toBeGreaterThan(10);
  });

  it("handles quarter-level roadmaps", () => {
    const road: GanttSpec = { ...spec, rows: [
      { label: "Pilot", kind: "task", start: "Q1 2027", end: "Q2 2027" },
      { label: "Roll-out", kind: "task", start: "Q3 2027", end: "Q2 2028" },
      { label: "Go-live", kind: "milestone", start: "1 Jul 2028", end: "" },
    ] };
    const r = run(road);
    expect(r.warnings).toEqual([]);
    expect(r.prims.filter((p) => p.role.startsWith("per:")).map((p) => (p as TextPrim).text).slice(0, 4)).toEqual(["Q1", "Q2", "Q3", "Q4"]);
  });
});

describe("bad end dates", () => {
  it("warns and falls back to the start date instead of corrupting the range", () => {
    const spec = sampleGantt("DMY", today);
    const bad: GanttSpec = { ...spec, rows: spec.rows.map((r, k) => (k === 2 ? { ...r, end: "sometime" } : r)) };
    const r = run(bad);
    expect(r.warnings.some((w) => w.includes("end date not recognised"))).toBe(true);
    const per = r.prims.filter((p) => p.role.startsWith("per:")).length;
    expect(per).toBeGreaterThan(5); // range still spans the plan
  });
});
