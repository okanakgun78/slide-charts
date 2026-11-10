import { describe, expect, it } from "vitest";
import { calendarColumns, layoutCalendar, naturalHeaderHeight } from "../src/engine/calendar";
import { toDay } from "../src/engine/dates";
import { bbox } from "../src/engine/geometry";
import { approxMeasure } from "../src/engine/measure";
import { defaultStyle, sampleCalendar } from "../src/engine/spec";
import type { CalendarSpec, RectPrim, TextPrim } from "../src/engine/types";

const style = defaultStyle();
const frame = { left: 24, top: 96, width: 729, height: 440 };
const today = toDay(2026, 11, 4); // Wednesday
const run = (s: CalendarSpec) => layoutCalendar(s, frame, style, approxMeasure);

// The user's sample: 2–8 November, a tall plan row, an italic shoes row, a bold yellow totals row.
const sample: CalendarSpec = {
  ...sampleCalendar("DMY", today),
  rows: [
    { label: "Plan", height: 6, style: "normal", align: "top-left", cells: {
      "2026-11-02": "1. Rest\n2. Stretching",
      "2026-11-05": "1. WU: 2.1K @4:49\n2. Drills\n3. Fartlek: 20K @3:51\n• 5K @4:11\n• 5K @3:55\n4. Strength",
    } },
    { label: "Shoes", height: 1, style: "italic", align: "center", cells: { "2026-11-03": "Hyper Edge 2 – 16K" } },
    { label: "Total", height: 1, style: "bold", align: "center", fill: "#FFFF00", cells: { "2026-11-02": "0.0 km", "2026-11-08": "97.4 km" } },
  ],
};

describe("calendar columns", () => {
  it("one column per day for a week, keyed by ISO date", () => {
    const c = calendarColumns(sample)!;
    expect(c.tiers).toEqual(["month", "day"]);
    expect(c.columns.map((x) => x.key)).toEqual(["2026-11-02", "2026-11-03", "2026-11-04", "2026-11-05", "2026-11-06", "2026-11-07", "2026-11-08"]);
    expect(c.columns[0].label).toBe("Mon 02 Nov");
  });
  it("auto picks weeks for two months and months for a year", () => {
    expect(calendarColumns({ ...sample, start: "1 Nov 2026", end: "31 Dec 2026", tiers: "auto" })!.tiers).toEqual(["month", "week"]);
    expect(calendarColumns({ ...sample, start: "Jan 2027", end: "Dec 2027", tiers: "auto" })!.tiers).toEqual(["year", "month"]);
  });
  it("returns null for unreadable dates", () => {
    expect(calendarColumns({ ...sample, start: "soon" })).toBeNull();
  });
});

describe("calendar layout (sample slide)", () => {
  const { prims, warnings } = run(sample);
  const byRole = (r: string) => prims.find((p) => p.role === r);
  it("stays inside the frame without warnings", () => {
    expect(warnings).toEqual([]);
    const b = bbox(prims);
    expect(b.left).toBeGreaterThanOrEqual(frame.left - 0.5);
    expect(b.left + b.width).toBeLessThanOrEqual(frame.left + frame.width + 0.5);
    expect(b.top + b.height).toBeLessThanOrEqual(frame.top + frame.height + 0.5);
  });
  it("has the sample header: November band and '02 / Mon' cells", () => {
    expect((byRole("hdt:month:0") as TextPrim).text).toBe("November");
    expect((byRole("hdt:day:0:1") as TextPrim).text).toBe("02");
    expect((byRole("hdt:day:0:2") as TextPrim).text).toBe("Mon");
  });
  it("keeps multi-line text in one wrapped, top-left cell", () => {
    const c = byRole("cell:0:3") as TextPrim;
    expect(c.text.split("\n").length).toBe(6);
    expect(c.wrap).toBe(true);
    expect(c.align).toBe("left");
    expect(c.valign).toBe("top");
  });
  it("applies row styles: italic centred row, bold yellow totals row", () => {
    expect((byRole("cell:1:1") as TextPrim).italic).toBe(true);
    expect((byRole("cell:2:6") as TextPrim).bold).toBe(true);
    expect((byRole("rowfill:2") as RectPrim).fill).toBe("#FFFF00");
  });
  it("sizes rows by their relative heights", () => {
    const plan = byRole("cell:0:0") as TextPrim, total = byRole("cell:2:0") as TextPrim;
    expect(plan.h / total.h).toBeCloseTo(6, 6);
  });
  it("warns when text does not fit its row", () => {
    const tight: CalendarSpec = { ...sample, rows: [{ ...sample.rows[0], height: 1 }, ...sample.rows.slice(1).map((r) => ({ ...r, height: 6 }))] };
    expect(run(tight).warnings.some((w) => w.includes("does not fit"))).toBe(true);
  });
  it("shows an optional label column", () => {
    const r = run({ ...sample, showLabels: true });
    expect((r.prims.find((p) => p.role === "rlab:2") as TextPrim).text).toBe("Total");
  });
});

describe("header-only calendar (no rows)", () => {
  const empty: CalendarSpec = { ...sample, rows: [] };
  it("is the default for a new calendar", () => {
    expect(sampleCalendar("DMY", today).rows).toEqual([]);
  });
  it("draws only the header, at natural height when the frame matches", () => {
    const natural = naturalHeaderHeight(empty, frame.width, style, approxMeasure);
    const r = layoutCalendar(empty, { ...frame, height: natural }, style, approxMeasure);
    expect(r.warnings).toEqual([]);
    expect(r.prims.some((p) => p.role.startsWith("cell:") || p.role.startsWith("vgrid:"))).toBe(false);
    expect(bbox(r.prims).height).toBeCloseTo(natural, 6);
    expect((r.prims.find((p) => p.role === "hdt:day:0:2") as TextPrim).text).toBe("Mon");
  });
  it("stretches with the frame within limits (a tall left-over frame is capped)", () => {
    const natural = naturalHeaderHeight(empty, frame.width, style, approxMeasure);
    const taller = layoutCalendar(empty, { ...frame, height: natural * 1.3 }, style, approxMeasure);
    expect(bbox(taller.prims).height).toBeCloseTo(natural * 1.3, 6);
    const huge = layoutCalendar(empty, { ...frame, height: natural * 6 }, style, approxMeasure);
    expect(bbox(huge.prims).height).toBeCloseTo(natural * 1.6, 6);
  });
});
