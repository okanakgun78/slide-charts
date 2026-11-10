import { describe, expect, it } from "vitest";
import { autoDecimals, contrastText, formatNumber, normHex } from "../src/engine/format";
import { approxMeasure } from "../src/engine/measure";
import { bbox, layoutStacked, placeLabels } from "../src/engine/stacked";
import type { StackedSpec, Style } from "../src/engine/types";

const style: Style = { fontSize: 12, textColor: "#1A1A1A", axisColor: "#7F7F7F", palette: ["#1F3A5F", "#3C6E9F", "#7FA7CF", "#B9CFE4"] };
const spec: StackedSpec = {
  v: 1, type: "stacked", categories: ["2021", "2022", "2023"], showTotals: true,
  series: [{ name: "A", values: [10, 20, 30] }, { name: "Bee", values: [5, null, 15] }],
};
const frame = { left: 50, top: 40, width: 500, height: 300 };

describe("format", () => {
  it("formats with thousands separators and decimals", () => {
    expect(formatNumber(1234567.891, 1)).toBe("1,234,567.9");
    expect(formatNumber(-0.04, 1)).toBe("0.0");
    expect(formatNumber(-12.5, 0)).toBe("-13");
    expect(formatNumber(1234.5, 1, ".", ",")).toBe("1.234,5");
  });
  it("picks decimals and colours", () => {
    expect(autoDecimals([1, 2, 3])).toBe(0);
    expect(autoDecimals([1, 2.5])).toBe(1);
    expect(normHex("abc")).toBe("#AABBCC");
    expect(normHex("#1f3a5f")).toBe("#1F3A5F");
    expect(normHex("rgb(1,2,3)")).toBeNull();
    expect(contrastText("#1F3A5F")).toBe("#FFFFFF");
    expect(contrastText("#B9CFE4")).toBe("#1A1A1A");
  });
});

describe("stacked layout", () => {
  const { prims, warnings } = layoutStacked(spec, frame, style, approxMeasure);
  const rects = prims.filter((p) => p.kind === "rect");

  it("draws one segment per positive value", () => {
    expect(warnings).toEqual([]);
    expect(rects).toHaveLength(5); // 6 cells, one null
  });

  it("stays inside the frame", () => {
    const b = bbox(prims);
    expect(b.left).toBeGreaterThanOrEqual(frame.left - 0.01);
    expect(b.top).toBeGreaterThanOrEqual(frame.top - 0.01);
    expect(b.left + b.width).toBeLessThanOrEqual(frame.left + frame.width + 0.01);
    expect(b.top + b.height).toBeLessThanOrEqual(frame.top + frame.height + 0.01);
  });

  it("scales segment heights proportionally to values", () => {
    const h = (role: string) => (rects.find((r) => r.role === role) as { h: number }).h;
    expect(h("seg:0:2") / h("seg:0:0")).toBeCloseTo(3, 6);
    expect(h("seg:1:2") / h("seg:1:0")).toBeCloseTo(3, 6);
  });

  it("stacks segments on a common baseline", () => {
    const axis = prims.find((p) => p.role === "axis") as { y1: number };
    const bottoms = rects.filter((r) => r.role.startsWith("seg:0:")).map((r) => (r as { y: number; h: number }).y + (r as { h: number }).h);
    bottoms.forEach((b) => expect(b).toBeCloseTo(axis.y1, 6));
  });

  it("labels totals correctly", () => {
    const totals = prims.filter((p) => p.role.startsWith("tot:")).map((p) => (p as { text: string }).text);
    expect(totals).toEqual(["15", "20", "45"]);
  });

  it("warns on negative values instead of drawing them", () => {
    const neg = { ...spec, series: [{ name: "A", values: [10, -5, 3] }] };
    expect(layoutStacked(neg, frame, style, approxMeasure).warnings[0]).toMatch(/Negative/);
  });
});

describe("label placement", () => {
  it("removes overlaps and keeps order", () => {
    const ys = placeLabels([100, 102, 104], 15, 0, 400);
    expect(ys[1] - ys[0]).toBeGreaterThanOrEqual(15 - 1e-9);
    expect(ys[2] - ys[1]).toBeGreaterThanOrEqual(15 - 1e-9);
  });
  it("shifts the block up when it overflows the bottom", () => {
    const ys = placeLabels([390, 395], 20, 0, 400);
    expect(Math.max(...ys)).toBeLessThanOrEqual(390 + 1e-9);
  });
});
