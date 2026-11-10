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

import { LINE_EPS, lineGeometry } from "../src/ppt/render";

/** Endpoints of a PowerPoint line box after rotation about its centre (top-left -> bottom-right). */
function endpoints(g: ReturnType<typeof lineGeometry>) {
  const cx = g.left + g.width / 2, cy = g.top + g.height / 2;
  const t = (g.rotation * Math.PI) / 180;
  const hx = g.width / 2, hy = g.height / 2;
  const rot = (x: number, y: number) => [cx + x * Math.cos(t) - y * Math.sin(t), cy + x * Math.sin(t) + y * Math.cos(t)];
  return [rot(-hx, -hy), rot(hx, hy)];
}

describe("lineGeometry", () => {
  it("never passes a zero height or width (Mac substitutes ~72 pt)", () => {
    const h = lineGeometry(10, 50, 200, 50);
    expect(h.height).toBe(LINE_EPS);
    expect(h.rotation).toBe(0);
    const v = lineGeometry(30, 10, 30, 90);
    expect(v.width).toBe(LINE_EPS);
  });
  it.each([
    [0, 50, 70, 0],    // rising
    [0, 0, 70, 50],    // falling
    [70, 0, 0, 50],    // right-to-left
    [10, 10, 12, 200], // steep
  ])("hits both endpoints for (%d,%d)->(%d,%d)", (x1, y1, x2, y2) => {
    const [a, b] = endpoints(lineGeometry(x1, y1, x2, y2));
    const near = (p: number[], x: number, y: number) => Math.hypot(p[0] - x, p[1] - y) < 0.05;
    expect((near(a, x1, y1) && near(b, x2, y2)) || (near(a, x2, y2) && near(b, x1, y1))).toBe(true);
  });
});
