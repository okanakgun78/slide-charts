import { describe, expect, it } from "vitest";
import { autoDecimals, contrastText, formatNumber, normHex } from "../src/engine/format";
import { approxMeasure } from "../src/engine/measure";
import { fitFontSize, placeLabels, wrapLines } from "../src/engine/geometry";
import { migrate } from "../src/engine/spec";

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
    expect(contrastText("#003760")).toBe("#FFFFFF");
    expect(contrastText("#BFD0E2")).toBe("#1A1A1A");
  });
});

describe("geometry helpers", () => {
  it("removes overlaps and keeps order", () => {
    const ys = placeLabels([100, 102, 104], 15, 0, 400);
    expect(ys[1] - ys[0]).toBeGreaterThanOrEqual(15 - 1e-9);
    expect(ys[2] - ys[1]).toBeGreaterThanOrEqual(15 - 1e-9);
  });
  it("shifts the block up when it overflows the bottom", () => {
    const ys = placeLabels([390, 395], 20, 0, 400);
    expect(Math.max(...ys)).toBeLessThanOrEqual(390 + 1e-9);
  });
  it("wraps words to the width", () => {
    const lines = wrapLines("Rest of the world markets", 60, 12, approxMeasure);
    expect(lines.length).toBeGreaterThan(1);
    lines.forEach((l) => expect(approxMeasure(l, 12) <= 60 || !l.includes(" ")).toBe(true));
  });
  it("shrinks font until texts fit", () => {
    const s = fitFontSize(["1,234,567"], 40, 12, approxMeasure);
    expect(approxMeasure("1,234,567", s)).toBeLessThanOrEqual(40);
  });
});

describe("migrate", () => {
  it("upgrades pilot v1 stacked specs", () => {
    const s = migrate({ v: 1, type: "stacked", categories: ["a"], series: [{ name: "x", values: [1] }], showTotals: true });
    expect(s).toMatchObject({ v: 2, type: "column", mode: "stacked", labels: { totals: true } });
  });
});
