// User settings, kept in this computer's browser storage for the add-in.

import { defaultStyle } from "../engine/spec";
import type { Style } from "../engine/types";

export interface Settings {
  fontSize: number;
  /** en: 1,234.5   eu: 1.234,5   space: 1 234,5 */
  numberFormat: "en" | "eu" | "space";
}

const KEY = "slidecharts.settings.v1";
const DEFAULTS: Settings = { fontSize: 12, numberFormat: "en" };

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return { ...DEFAULTS, ...(JSON.parse(raw) as Partial<Settings>) };
  } catch {
    /* storage unavailable: defaults */
  }
  return { ...DEFAULTS };
}

export function saveSettings(s: Settings) {
  try {
    localStorage.setItem(KEY, JSON.stringify(s));
  } catch {
    /* ignore */
  }
}

export function separators(s: Settings): { decimalSep: string; thousandsSep: string } {
  if (s.numberFormat === "eu") return { decimalSep: ",", thousandsSep: "." };
  if (s.numberFormat === "space") return { decimalSep: ",", thousandsSep: " " };
  return { decimalSep: ".", thousandsSep: "," };
}

export function styleFrom(s: Settings, ink: Pick<Style, "textColor" | "axisColor" | "gridColor">): Style {
  return { ...defaultStyle(), ...ink, ...separators(s), fontSize: s.fontSize };
}
