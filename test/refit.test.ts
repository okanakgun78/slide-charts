import { describe, expect, it } from "vitest";

// Importing the renderer must not touch the PowerPoint global (it only exists after Office loads).
import { refitFrame } from "../src/ppt/render";

const stored = {
  app: "slide-charts" as const, ver: "t",
  spec: { v: 1 as const, type: "stacked" as const, categories: [], series: [], showTotals: true },
  frame: { left: 100, top: 100, width: 400, height: 200 },
  drawn: { left: 100, top: 110, width: 400, height: 190 },
};

describe("refitFrame", () => {
  it("is the identity when the chart was not touched", () => {
    expect(refitFrame(stored, stored.drawn)).toEqual(stored.frame);
  });
  it("follows a move", () => {
    const f = refitFrame(stored, { left: 150, top: 60, width: 400, height: 190 });
    expect(f).toEqual({ left: 150, top: 50, width: 400, height: 200 });
  });
  it("follows a resize", () => {
    const f = refitFrame(stored, { left: 100, top: 110, width: 800, height: 380 });
    expect(f.width).toBeCloseTo(800);
    expect(f.height).toBeCloseTo(400);
    expect(f.top).toBeCloseTo(90);
  });
});
