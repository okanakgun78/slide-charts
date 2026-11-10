// SVG renderer for primitives: used for local visual checks and in-pane previews.

import { wrapLines } from "./geometry";
import { approxMeasure } from "./measure";
import type { Frame, Measure, Prim } from "./types";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export function toSvg(prims: Prim[], frame: Frame, fontFamily = "Helvetica, Arial, sans-serif", measure: Measure = approxMeasure): string {
  const m = 8;
  const vb = `${frame.left - m} ${frame.top - m} ${frame.width + 2 * m} ${frame.height + 2 * m}`;
  const body = prims
    .map((p) => {
      if (p.kind === "rect") {
        const stroke = p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.strokeWidth ?? 0.75}"` : "";
        return `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="${p.fill}"${stroke}/>`;
      }
      if (p.kind === "line") {
        const dash = p.dash ? ` stroke-dasharray="4 3"` : "";
        return `<line x1="${p.x1}" y1="${p.y1}" x2="${p.x2}" y2="${p.y2}" stroke="${p.color}" stroke-width="${p.width}"${dash}/>`;
      }
      if (p.kind === "shape") {
        const stroke = p.stroke ? ` stroke="${p.stroke}"` : "";
        if (p.shape === "Ellipse") return `<ellipse cx="${p.x + p.w / 2}" cy="${p.y + p.h / 2}" rx="${p.w / 2}" ry="${p.h / 2}" fill="${p.fill}"${stroke}/>`;
        const cx = p.x + p.w / 2, cy = p.y + p.h / 2;
        return `<polygon points="${cx},${p.y} ${p.x + p.w},${cy} ${cx},${p.y + p.h} ${p.x},${cy}" fill="${p.fill}"${stroke}/>`;
      }
      const ins = p.inset ?? 0;
      const lines = p.text.split("\n").flatMap((l) => (p.wrap ? wrapLines(l, p.w - 2 * ins, p.size, measure, p.bold) : [l]));
      const lh = p.size * 1.2;
      const blockH = lines.length * lh;
      const tx = p.align === "left" ? p.x + ins : p.align === "right" ? p.x + p.w - ins : p.x + p.w / 2;
      const anchor = p.align === "left" ? "start" : p.align === "right" ? "end" : "middle";
      const y0 = p.valign === "top" ? p.y + ins : p.valign === "bottom" ? p.y + p.h - ins - blockH : p.y + (p.h - blockH) / 2;
      const weight = p.bold ? ` font-weight="bold"` : "";
      const italic = p.italic ? ` font-style="italic"` : "";
      return lines
        .map((t, k) => `<text x="${tx}" y="${y0 + k * lh + p.size * 0.95}" font-size="${p.size}" fill="${p.color}" text-anchor="${anchor}"${weight}${italic}>${esc(t)}</text>`)
        .join("");
    })
    .join("\n");
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" font-family="${fontFamily}">\n${body}\n</svg>`;
}
