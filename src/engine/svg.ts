// SVG renderer for primitives: used for local visual checks and in-pane previews.

import type { Frame, Prim } from "./types";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export function toSvg(prims: Prim[], frame: Frame, fontFamily = "Helvetica, Arial, sans-serif"): string {
  const m = 6;
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
      const tx = p.align === "left" ? p.x : p.align === "right" ? p.x + p.w : p.x + p.w / 2;
      const anchor = p.align === "left" ? "start" : p.align === "right" ? "end" : "middle";
      const ty = p.valign === "top" ? p.y + p.size : p.valign === "bottom" ? p.y + p.h - p.size * 0.25 : p.y + p.h / 2 + p.size * 0.35;
      const weight = p.bold ? ` font-weight="bold"` : "";
      return `<text x="${tx}" y="${ty}" font-size="${p.size}" fill="${p.color}" text-anchor="${anchor}"${weight}>${esc(p.text)}</text>`;
    })
    .join("\n");
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" font-family="${fontFamily}">\n${body}\n</svg>`;
}
