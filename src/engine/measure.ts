// Text width measurement.
// In the task pane we measure with a canvas; PowerPoint's own fonts may not be visible to the
// web view, so the pilot reports canvas vs PowerPoint autosize to calibrate this.

import type { Measure } from "./types";

/** Font-independent estimate (average glyph ~0.53 em, digits ~0.56 em). Used in tests. */
export const approxMeasure: Measure = (text, size, bold) => {
  let em = 0;
  for (const ch of text) em += /[0-9]/.test(ch) ? 0.56 : /[ .,:;'|il!]/.test(ch) ? 0.28 : /[MW@%]/.test(ch) ? 0.85 : 0.53;
  return em * size * (bold ? 1.06 : 1);
};

export function canvasMeasure(fontFamily: string, correction = 1): Measure {
  const ctx = document.createElement("canvas").getContext("2d");
  if (!ctx) return approxMeasure;
  return (text, size, bold) => {
    ctx.font = `${bold ? "bold " : ""}${size}px ${fontFamily}`;
    return ctx.measureText(text).width * correction;
  };
}
