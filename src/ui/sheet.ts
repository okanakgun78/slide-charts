// Data sheet for category charts: series in rows, categories across (as in Excel and think-cell).
// Keeps raw cell strings so partially typed numbers ("1.") are not lost.

import { HIGHLIGHTS } from "../engine/spec";
import type { Series } from "../engine/types";
import { cellText, parseNumber, parseTSV } from "./parse";

export interface SheetData {
  categories: string[];
  series: Series[];
}

interface Meta {
  color?: string;
  asLine?: boolean;
}

export interface SheetOptions {
  decimalSep: () => string;
  combo: () => boolean;
  palette: () => string[];
  onChange: () => void;
}

export class Sheet {
  private grid: string[][] = [[""]];
  private meta: Meta[] = [];
  private focus = { r: 1, c: 1 };
  private popover: HTMLDivElement | null = null;

  constructor(private root: HTMLElement, private opts: SheetOptions) {
    root.addEventListener("input", (e) => this.onInput(e));
    root.addEventListener("focusin", (e) => this.onFocus(e));
    root.addEventListener("paste", (e) => this.onPaste(e as ClipboardEvent));
    root.addEventListener("keydown", (e) => this.onKey(e as KeyboardEvent));
    root.addEventListener("click", (e) => this.onClick(e));
    document.addEventListener("click", (e) => {
      if (this.popover && !this.popover.contains(e.target as Node) && !(e.target as HTMLElement).classList.contains("sw")) this.closePopover();
    });
  }

  load(d: SheetData) {
    const sep = this.opts.decimalSep();
    this.grid = [["", ...d.categories], ...d.series.map((s) => [s.name, ...d.categories.map((_, i) => cellText(s.values[i] ?? null, sep))])];
    this.meta = d.series.map((s) => ({ color: s.color, asLine: s.asLine }));
    this.render();
  }

  data(): SheetData {
    const sep = this.opts.decimalSep();
    const nCols = this.grid[0].length - 1;
    const categories = this.grid[0].slice(1).map((c, i) => c.trim() || `Category ${i + 1}`);
    const series = this.grid.slice(1).map((row, k) => {
      const s: Series = { name: row[0].trim() || `Series ${k + 1}`, values: Array.from({ length: nCols }, (_, i) => parseNumber(row[i + 1] ?? "", sep)) };
      if (this.meta[k]?.color) s.color = this.meta[k].color;
      if (this.meta[k]?.asLine) s.asLine = true;
      return s;
    });
    return { categories, series };
  }

  // ---------- Structure edits ----------

  addSeries() {
    this.grid.push(["", ...new Array(this.grid[0].length - 1).fill("")]);
    this.meta.push({});
    this.focus = { r: this.grid.length - 1, c: 0 };
    this.changed(true);
  }

  addCategory() {
    this.grid.forEach((row) => row.push(""));
    this.focus = { r: 0, c: this.grid[0].length - 1 };
    this.changed(true);
  }

  /** Deletes the series (row) of the last focused cell. Returns an error message or "". */
  deleteSeries(): string {
    const r = this.focus.r;
    if (r < 1) return "Click a cell in the series to delete first.";
    if (this.grid.length <= 2) return "A chart needs at least one series.";
    this.grid.splice(r, 1);
    this.meta.splice(r - 1, 1);
    this.focus = { r: Math.min(r, this.grid.length - 1), c: this.focus.c };
    this.changed(true);
    return "";
  }

  /** Deletes the category (column) of the last focused cell. Returns an error message or "". */
  deleteCategory(): string {
    const c = this.focus.c;
    if (c < 1) return "Click a cell in the category to delete first.";
    if (this.grid[0].length <= 2) return "A chart needs at least one category.";
    this.grid.forEach((row) => row.splice(c, 1));
    this.focus = { r: this.focus.r, c: Math.min(c, this.grid[0].length - 1) };
    this.changed(true);
    return "";
  }

  transpose() {
    const d = this.data();
    const cats = d.series.map((s) => s.name);
    const series: Series[] = d.categories.map((name, i) => ({ name, values: d.series.map((s) => s.values[i] ?? null) }));
    this.load({ categories: cats, series });
    this.opts.onChange();
  }

  // ---------- Rendering ----------

  render() {
    this.closePopover();
    const pal = this.opts.palette();
    const combo = this.opts.combo();
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    const head = this.grid[0]
      .map((v, c) => (c === 0 ? `<th class="corner"></th><th class="corner">Series</th>` : `<th><input data-r="0" data-c="${c}" value="${esc(v)}" placeholder="Cat ${c}"></th>`))
      .join("") + (combo ? `<th class="corner">Line</th>` : "");
    let barPos = 0;
    const body = this.grid.slice(1).map((row, k) => {
      const r = k + 1;
      const m = this.meta[k] ?? {};
      const auto = m.asLine && combo ? "#7F7F7F" : pal[barPos % pal.length];
      if (!(m.asLine && combo)) barPos++;
      const cells = row.map((v, c) =>
        c === 0
          ? `<td><button class="sw" data-k="${k}" title="Series colour" style="background:${m.color ?? auto}"></button></td><td><input class="name" data-r="${r}" data-c="0" value="${esc(v)}" placeholder="Series ${r}"></td>`
          : `<td><input class="num" data-r="${r}" data-c="${c}" value="${esc(v)}" inputmode="decimal"></td>`,
      ).join("");
      const line = combo ? `<td class="ctr"><input type="checkbox" class="asline" data-k="${k}" ${m.asLine ? "checked" : ""}></td>` : "";
      return `<tr>${cells}${line}</tr>`;
    }).join("");
    this.root.innerHTML = `<table class="sheet"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
    const f = this.root.querySelector<HTMLInputElement>(`input[data-r="${this.focus.r}"][data-c="${this.focus.c}"]`);
    if (f && this.root.contains(document.activeElement)) f.focus();
  }

  focusCell(r: number, c: number) {
    this.focus = { r, c };
    this.root.querySelector<HTMLInputElement>(`input[data-r="${r}"][data-c="${c}"]`)?.focus();
  }

  // ---------- Events ----------

  private changed(rerender: boolean) {
    if (rerender) {
      this.render();
      this.focusCell(this.focus.r, this.focus.c);
    }
    this.opts.onChange();
  }

  private cellOf(t: EventTarget | null): { r: number; c: number } | null {
    const el = t as HTMLElement;
    if (!el?.dataset || el.dataset.r === undefined) return null;
    return { r: Number(el.dataset.r), c: Number(el.dataset.c) };
  }

  private onFocus(e: Event) {
    const cell = this.cellOf(e.target);
    if (cell) this.focus = cell;
  }

  private onInput(e: Event) {
    const t = e.target as HTMLInputElement;
    if (t.classList.contains("asline")) {
      const k = Number(t.dataset.k);
      this.meta[k] = { ...this.meta[k], asLine: t.checked };
      this.changed(true);
      return;
    }
    const cell = this.cellOf(t);
    if (!cell) return;
    this.grid[cell.r][cell.c] = t.value;
    this.opts.onChange();
  }

  private onKey(e: KeyboardEvent) {
    const cell = this.cellOf(e.target);
    if (!cell) return;
    const move = (dr: number, dc: number) => {
      e.preventDefault();
      const r = Math.max(0, Math.min(this.grid.length - 1, cell.r + dr));
      const c = Math.max(0, Math.min(this.grid[0].length - 1, cell.c + dc));
      this.focusCell(r === 0 && c === 0 ? 1 : r, c);
    };
    if (e.key === "Enter" || e.key === "ArrowDown") move(1, 0);
    else if (e.key === "ArrowUp") move(-1, 0);
  }

  private onPaste(e: ClipboardEvent) {
    const cell = this.cellOf(e.target);
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!cell || !/[\t\n]/.test(text.replace(/\r?\n$/, ""))) return; // single value: normal paste
    e.preventDefault();
    const block = parseTSV(text);
    // A block with an empty top-left cell is a whole table (header row + names): replace everything.
    const whole = block.length > 1 && (block[0]?.length ?? 0) > 1 && block[0][0].trim() === "";
    const r0 = whole ? 0 : cell.r;
    const c0 = whole ? 0 : cell.c;
    if (whole) {
      this.grid = [[""]];
      this.meta = [];
    }
    const width = Math.max(...block.map((b) => b.length));
    const needRows = r0 + block.length;
    const needCols = c0 + width;
    while (this.grid.length < needRows) { this.grid.push([]); this.meta.push({}); }
    const cols = Math.max(needCols, ...this.grid.map((row) => row.length));
    this.grid.forEach((row) => { while (row.length < cols) row.push(""); });
    block.forEach((brow, dr) => brow.forEach((v, dc) => {
      if (r0 + dr === 0 && c0 + dc === 0) return;
      this.grid[r0 + dr][c0 + dc] = v.trim();
    }));
    if (this.grid.length < 2) { this.grid.push(new Array(cols).fill("")); this.meta.push({}); }
    this.focus = { r: Math.max(1, r0), c: Math.max(1, c0) };
    this.changed(true);
  }

  private onClick(e: Event) {
    const t = e.target as HTMLElement;
    if (!t.classList.contains("sw")) return;
    const k = Number(t.dataset.k);
    this.openPopover(t, k);
  }

  private openPopover(anchor: HTMLElement, k: number) {
    this.closePopover();
    const pop = document.createElement("div");
    pop.className = "popover";
    const colours = [...this.opts.palette(), ...HIGHLIGHTS];
    pop.innerHTML =
      colours.map((c) => `<button class="sw pick" data-c="${c}" style="background:${c}" title="${c}"></button>`).join("") +
      `<button class="auto" data-c="">Auto</button>`;
    pop.addEventListener("click", (ev) => {
      const b = ev.target as HTMLElement;
      if (b.dataset.c === undefined) return;
      this.meta[k] = { ...this.meta[k], color: b.dataset.c || undefined };
      this.changed(true);
    });
    const r = anchor.getBoundingClientRect();
    pop.style.left = `${r.left + window.scrollX}px`;
    pop.style.top = `${r.bottom + window.scrollY + 4}px`;
    document.body.appendChild(pop);
    this.popover = pop;
  }

  private closePopover() {
    this.popover?.remove();
    this.popover = null;
  }
}
