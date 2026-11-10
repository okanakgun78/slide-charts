import { describe, expect, it } from "vitest";
import { toDay } from "../src/engine/dates";
import { autoTiers, layoutGantt, tierPeriods } from "../src/engine/gantt";
import { bbox } from "../src/engine/geometry";
import { approxMeasure } from "../src/engine/measure";
import { defaultStyle, sampleGantt } from "../src/engine/spec";
import type { GanttSpec, Prim, RectPrim, ShapePrim, TextPrim } from "../src/engine/types";

const style = defaultStyle();
const frame = { left: 40, top: 60, width: 860, height: 400 };
const today = toDay(2026, 11, 10);
const run = (s: GanttSpec) => layoutGantt(s, frame, style, approxMeasure, today);
const byRole = (ps: Prim[], role: string) => ps.find((p) => p.role === role);

describe("tiers and periods", () => {
  it("picks tiers from the span", () => {
    expect(autoTiers(14)).toEqual(["month", "day"]);
    expect(autoTiers(60)).toEqual(["month", "week"]);
    expect(autoTiers(300)).toEqual(["year", "month"]);
    expect(autoTiers(547)).toEqual(["year", "quarter"]);
    expect(autoTiers(3000)).toEqual(["year"]);
  });
  it("weeks start on Monday; month labels prefer full names", () => {
    const tiers = ["month", "week"] as const;
    const weeks = tierPeriods(toDay(2026, 11, 2), toDay(2026, 12, 21), "week", { tiers: [...tiers], multiYear: false });
    expect(weeks[0].start).toBe(toDay(2026, 11, 2));
    const months = tierPeriods(toDay(2026, 11, 2), toDay(2026, 12, 21), "month", { tiers: [...tiers], multiYear: false });
    expect(months.map((m) => m.labels[0])).toEqual(["November", "December"]);
    expect(months[0].start).toBe(toDay(2026, 11, 2)); // clipped to the range
  });
  it("days carry two-line labels like the sample (02 / Mon)", () => {
    const days = tierPeriods(toDay(2026, 11, 2), toDay(2026, 11, 9), "day", { tiers: ["month", "day"], multiYear: false });
    expect(days.map((d) => d.lines![0].join(" "))).toEqual(["02 Mon", "03 Tue", "04 Wed", "05 Thu", "06 Fri", "07 Sat", "08 Sun"]);
  });
  it("quarters snap to quarter starts", () => {
    const q = tierPeriods(toDay(2026, 1, 1), toDay(2027, 7, 1), "quarter", { tiers: ["year", "quarter"], multiYear: true });
    expect(q.map((b) => b.labels[0])).toEqual(["Q1", "Q2", "Q3", "Q4", "Q1", "Q2"]);
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
    expect(r.prims.filter((p) => p.role.startsWith("hdt:quarter:")).map((p) => (p as TextPrim).text).slice(0, 4)).toEqual(["Q1", "Q2", "Q3", "Q4"]);
  });
});

describe("bad end dates", () => {
  it("warns and falls back to the start date instead of corrupting the range", () => {
    const spec = sampleGantt("DMY", today);
    const bad: GanttSpec = { ...spec, rows: spec.rows.map((r, k) => (k === 2 ? { ...r, end: "sometime" } : r)) };
    const r = run(bad);
    expect(r.warnings.some((w) => w.includes("end date not recognised"))).toBe(true);
    const per = r.prims.filter((p) => p.role.startsWith("hdt:")).length;
    expect(per).toBeGreaterThan(5); // range still spans the plan
  });
});

describe("sample-style header (month + day)", () => {
  const week: GanttSpec = { ...sampleGantt("DMY", today), tiers: ["month", "day"], rows: [
    { label: "Zone 2 run", kind: "task", start: "3 Nov 2026", end: "3 Nov 2026" },
    { label: "Fartlek", kind: "task", start: "5 Nov 2026", end: "5 Nov 2026" },
    { label: "Long run", kind: "milestone", start: "8 Nov 2026", end: "" },
    { label: "Recovery", kind: "task", start: "2 Nov 2026", end: "4 Nov 2026" },
  ] };
  const { prims, warnings } = run(week);
  const t = (prefix: string) => prims.filter((p) => p.role.startsWith(prefix)).map((p) => (p as TextPrim).text);
  it("shows the month band and '02 / Mon' day cells", () => {
    expect(warnings).toEqual([]);
    expect(t("hdt:month:")).toEqual(["November"]);
    expect(t("hdt:day:0:")).toEqual(["02", "Mon"]);
    expect(t("hdt:day:6:")).toEqual(["08", "Sun"]);
  });
  it("fills header cells with white bold text and shades the weekend", () => {
    const cells = prims.filter((p) => p.role.startsWith("hd:")) as RectPrim[];
    expect(cells.length).toBe(8);
    expect(cells.every((c) => c.fill === "#003760" && c.stroke === "#FFFFFF")).toBe(true);
    const lbl = prims.find((p) => p.role === "hdt:day:0:1") as TextPrim;
    expect(lbl.color).toBe("#FFFFFF");
    expect(lbl.bold).toBe(true);
    expect(prims.filter((p) => p.role.startsWith("wkend:")).length).toBe(2);
  });
  it("one column per day: a one-day task is one column wide", () => {
    const bar = prims.find((p) => p.role === "bar:0") as RectPrim;
    const cell = prims.find((p) => p.role === "hd:day:1") as RectPrim;
    expect(bar.w).toBeCloseTo(cell.w, 6);
    expect(bar.x).toBeCloseTo(cell.x, 6);
  });
  it("plain style keeps text-only headers", () => {
    const r = run({ ...week, headerStyle: "plain" });
    expect(r.prims.some((p) => p.role.startsWith("hd:"))).toBe(false);
    expect(r.prims.some((p) => p.role === "hdt:month:0")).toBe(true);
  });
  it("narrow days fall back to single letters, then to an unlabelled grid", () => {
    const long: GanttSpec = { ...week, rows: [{ label: "x", kind: "task", start: "1 Nov 2026", end: "31 Dec 2026" }] };
    const r = layoutGantt(long, { ...frame, width: 500 }, style, approxMeasure, today);
    expect(r.prims.filter((p) => p.role.startsWith("hd:day:")).length).toBe(61);
    expect(r.prims.filter((p) => p.role.startsWith("hdt:day:")).every((p) => (p as TextPrim).text.length <= 2)).toBe(true);
  });
});

describe("migration of v0.4 Gantt charts", () => {
  it("maps the old scale to tiers", async () => {
    const { migrate } = await import("../src/engine/spec");
    const old = { ...sampleGantt("DMY", today), scale: "quarter" } as Record<string, unknown>;
    delete old.tiers; delete old.headerStyle; delete old.shadeWeekends;
    expect(migrate(old)).toMatchObject({ tiers: ["year", "quarter"], headerStyle: "filled" });
  });
});

describe("header refinements", () => {
  it("week numbers follow ISO 8601", async () => {
    const { isoWeek } = await import("../src/engine/gantt");
    expect(isoWeek(toDay(2026, 11, 2))).toBe(45);
    expect(isoWeek(toDay(2027, 1, 1))).toBe(53); // Friday 1 Jan 2027 belongs to week 53 of 2026
    expect(isoWeek(toDay(2027, 1, 4))).toBe(1);
  });
  it("labels weeks by number when days are shown, and drops one-letter labels on cut-off cells", () => {
    const ws = tierPeriods(toDay(2026, 10, 26), toDay(2026, 11, 16), "week", { tiers: ["month", "week", "day"], multiYear: false });
    expect(ws.map((w) => w.labels[0])).toEqual(["Week 44", "Week 45", "Week 46"]);
    const ms = tierPeriods(toDay(2026, 10, 26), toDay(2026, 11, 16), "month", { tiers: ["month", "week"], multiYear: false });
    expect(ms[0].labels).not.toContain("O");
  });
});
