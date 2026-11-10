import { describe, expect, it } from "vitest";
import { layoutCategory } from "../src/engine/category";
import { bbox } from "../src/engine/geometry";
import { approxMeasure } from "../src/engine/measure";
import { defaultStyle, sampleCategory } from "../src/engine/spec";
import type { CategorySpec, Prim, RectPrim, TextPrim } from "../src/engine/types";

const style = defaultStyle();
const frame = { left: 50, top: 40, width: 600, height: 340 };
const run = (s: CategorySpec) => layoutCategory(s, frame, style, approxMeasure);
const rect = (ps: Prim[], role: string) => ps.find((p) => p.role === role) as RectPrim;
const texts = (ps: Prim[], prefix: string) => ps.filter((p) => p.role.startsWith(prefix)) as TextPrim[];
const inside = (ps: Prim[]) => {
  const b = bbox(ps);
  expect(b.left).toBeGreaterThanOrEqual(frame.left - 0.5);
  expect(b.top).toBeGreaterThanOrEqual(frame.top - 0.5);
  expect(b.left + b.width).toBeLessThanOrEqual(frame.left + frame.width + 0.5);
  expect(b.top + b.height).toBeLessThanOrEqual(frame.top + frame.height + 0.5);
};

function spec(over: Partial<CategorySpec>): CategorySpec {
  return { ...sampleCategory("column", "stacked"), ...over };
}

describe.each([
  ["column", "stacked"], ["column", "clustered"], ["column", "percent"],
  ["bar", "stacked"], ["bar", "clustered"], ["bar", "percent"],
  ["line", "stacked"], ["combo", "stacked"], ["combo", "clustered"],
] as const)("%s %s", (type, mode) => {
  it("lays out inside the frame without warnings", () => {
    const { prims, warnings } = run(sampleCategory(type, mode));
    expect(warnings).toEqual([]);
    expect(prims.length).toBeGreaterThan(10);
    inside(prims);
  });
});

describe("stacked columns", () => {
  const s = spec({ series: [{ name: "Top", values: [10, 20] }, { name: "Bottom", values: [30, 30] }], categories: ["A", "B"] });
  const { prims } = run(s);
  it("puts the first data row on top", () => {
    expect(rect(prims, "seg:0:0").y).toBeLessThan(rect(prims, "seg:1:0").y);
  });
  it("scales heights proportionally", () => {
    expect(rect(prims, "seg:0:1").h / rect(prims, "seg:0:0").h).toBeCloseTo(2, 6);
    expect(rect(prims, "seg:1:0").h / rect(prims, "seg:0:0").h).toBeCloseTo(3, 6);
  });
  it("labels totals", () => {
    expect(texts(prims, "tot:").map((t) => t.text)).toEqual(["40", "50"]);
  });
  it("puts series names on the right, aligned with the last column", () => {
    const names = texts(prims, "ser:");
    expect(names.map((t) => t.text)).toEqual(["Top", "Bottom"]);
    expect(names[0].y).toBeLessThan(names[1].y);
  });
});

describe("negative values", () => {
  const s = spec({ categories: ["A", "B"], series: [{ name: "P", values: [20, 10] }, { name: "N", values: [-5, -15] }] });
  const { prims } = run(s);
  const axis = prims.find((p) => p.role === "axis") as { y1: number };
  it("stacks negatives below the axis and positives above", () => {
    const p = rect(prims, "seg:0:1"), n = rect(prims, "seg:1:1");
    expect(p.y + p.h).toBeCloseTo(axis.y1, 6);
    expect(n.y).toBeCloseTo(axis.y1, 6);
  });
  it("shows net totals", () => {
    expect(texts(prims, "tot:").map((t) => t.text)).toEqual(["15", "-5"]);
  });
});

describe("100% columns", () => {
  const s = spec({ mode: "percent", categories: ["A"], series: [{ name: "x", values: [1] }, { name: "y", values: [3] }] });
  const { prims } = run(s);
  it("labels shares in percent and keeps the absolute total", () => {
    expect(texts(prims, "lbl:").map((t) => t.text).sort()).toEqual(["25%", "75%"]);
    expect(texts(prims, "tot:")[0].text).toBe("4");
  });
});

describe("clustered", () => {
  const { prims } = run(sampleCategory("column", "clustered"));
  it("places bars side by side with labels above", () => {
    const a = rect(prims, "seg:0:0"), b = rect(prims, "seg:1:0");
    expect(b.x).toBeCloseTo(a.x + a.w, 6);
    const l = prims.find((p) => p.role === "lbl:0:0") as TextPrim;
    expect(l.y + l.h).toBeLessThanOrEqual(a.y + 0.01);
  });
  it("uses a legend on top", () => {
    expect(texts(prims, "leg:").length).toBe(3);
  });
});

describe("bars", () => {
  const { prims } = run(sampleCategory("bar", "stacked"));
  it("runs bars horizontally with categories on the left", () => {
    const a = rect(prims, "seg:0:0");
    expect(a.w).toBeGreaterThan(a.h * 0.2);
    const cat = prims.find((p) => p.role === "cat:0") as TextPrim;
    expect(cat.x + cat.w).toBeLessThanOrEqual(Math.min(...prims.filter((p) => p.kind === "rect" && p.role.startsWith("seg")).map((p) => (p as RectPrim).x)) + 0.01);
  });
});

describe("lines", () => {
  const s = sampleCategory("line", "stacked");
  const { prims } = run(s);
  it("draws n-1 segments and n markers per series", () => {
    expect(prims.filter((p) => p.role.startsWith("line:0:")).length).toBe(4);
    expect(prims.filter((p) => p.role.startsWith("mk:0:")).length).toBe(5);
  });
  it("skips gaps", () => {
    const g = { ...s, series: [{ name: "a", values: [1, null, 3, 4, 5] }] };
    expect(run(g).prims.filter((p) => p.role.startsWith("line:0:")).length).toBe(2);
  });
  it("labels points without overlap within a category", () => {
    const lbl = texts(prims, "plbl:").filter((t) => t.role.endsWith(":0"));
    expect(Math.abs(lbl[0].y - lbl[1].y)).toBeGreaterThanOrEqual(12 * 1.3 * 0.95 - 1e-6);
  });
});
