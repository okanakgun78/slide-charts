// Data sheet for calendar charts.
// Rows table: one line per calendar row (label, height, style, fill, alignment).
// Text grid: one line per column of the calendar (date), one text box per calendar row.
// Paste: an Excel calendar with dates across the top or down the left replaces the content and
// sets the date range; any other block fills cells from the focused one.

import type { CalendarColumn } from "../engine/calendar";
import { isoDay, parseDate, type DateOrder } from "../engine/dates";
import type { CalendarRow } from "../engine/types";
import { parseTSV } from "./parse";

export interface CalendarSheetOptions {
  columns: () => CalendarColumn[];
  dateOrder: () => DateOrder;
  onChange: () => void;
  /** Called when a pasted calendar sets a new date range (ISO dates). */
  onRange: (start: string, end: string) => void;
}

const FILLS: [string, string][] = [["", "None"], ["#FFFF00", "Yellow"], ["#F2F2F2", "Light grey"], ["#DCE6F2", "Light blue"], ["#FDE9D9", "Light red"]];

const newRow = (k: number): CalendarRow => ({
  label: "", height: k === 0 ? 4 : 1, style: "normal", align: k === 0 ? "top-left" : "center", cells: {},
});

export class CalendarSheet {
  private rows: CalendarRow[] = [];
  private focus = { ci: 0, k: 0 };

  constructor(private rowsRoot: HTMLElement, private cellsRoot: HTMLElement, private opts: CalendarSheetOptions) {
    rowsRoot.addEventListener("input", (e) => this.onRowInput(e));
    rowsRoot.addEventListener("change", (e) => this.onRowInput(e));
    rowsRoot.addEventListener("click", (e) => this.onRowClick(e));
    cellsRoot.addEventListener("input", (e) => this.onCellInput(e));
    cellsRoot.addEventListener("focusin", (e) => {
      const t = e.target as HTMLElement;
      if (t.dataset?.k !== undefined) this.focus = { ci: Number(t.dataset.ci), k: Number(t.dataset.k) };
    });
    cellsRoot.addEventListener("paste", (e) => this.onPaste(e as ClipboardEvent));
  }

  load(rows: CalendarRow[]) {
    this.rows = rows.map((r) => ({ ...r, cells: { ...r.cells } }));
    this.render();
  }

  data(): CalendarRow[] {
    return this.rows.map((r) => ({ ...r, cells: { ...r.cells } }));
  }

  addRow() {
    this.rows.push(newRow(this.rows.length));
    this.render();
    this.opts.onChange();
  }

  render() {
    this.renderRows();
    this.renderCells();
  }

  renderRows() {
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    const opt = (v: string, t: string, cur: string) => `<option value="${v}" ${v === cur ? "selected" : ""}>${t}</option>`;
    const body = this.rows.map((r, k) => `<tr>
      <td><input class="name" data-k="${k}" data-f="label" value="${esc(r.label)}" placeholder="Row ${k + 1}"></td>
      <td><select data-k="${k}" data-f="height">${[1, 2, 3, 4, 5, 6, 8, 10].map((h) => opt(String(h), String(h), String(r.height))).join("")}</select></td>
      <td><select data-k="${k}" data-f="style">${opt("normal", "Normal", r.style)}${opt("bold", "Bold", r.style)}${opt("italic", "Italic", r.style)}</select></td>
      <td><select data-k="${k}" data-f="fill">${FILLS.map(([v, t]) => opt(v, t, r.fill ?? "")).join("")}</select></td>
      <td><select data-k="${k}" data-f="align">${opt("top-left", "Top left", r.align)}${opt("center", "Centre", r.align)}</select></td>
      <td class="ctr"><button class="mini" data-k="${k}" data-act="up" title="Move up">↑</button><button class="mini" data-k="${k}" data-act="del" title="Delete row">✕</button></td>
    </tr>`).join("");
    this.rowsRoot.innerHTML = `<table class="sheet cal"><thead><tr><th class="corner">Row</th><th class="corner">Height</th><th class="corner">Style</th><th class="corner">Fill</th><th class="corner">Align</th><th class="corner"></th></tr></thead><tbody>${body}</tbody></table>`;
  }

  renderCells() {
    const cols = this.opts.columns();
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    if (!cols.length) {
      this.cellsRoot.innerHTML = `<p class="muted pad">Enter a start and end date above.</p>`;
      return;
    }
    const head = `<th class="corner">Date</th>` + this.rows.map((r, k) => `<th class="corner">${esc(r.label || `Row ${k + 1}`)}</th>`).join("");
    const body = cols.map((c, ci) => `<tr><th class="dlabel">${esc(c.label)}</th>` + this.rows.map((r, k) => {
      const v = r.cells[c.key] ?? "";
      const lines = Math.min(6, Math.max(1, v.split("\n").length));
      return `<td><textarea data-ci="${ci}" data-k="${k}" data-key="${c.key}" rows="${lines}">${esc(v)}</textarea></td>`;
    }).join("") + `</tr>`).join("");
    this.cellsRoot.innerHTML = `<table class="sheet calcells"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
  }

  // ---------- Events ----------

  private onRowInput(e: Event) {
    const t = e.target as HTMLInputElement | HTMLSelectElement;
    if (t.dataset?.k === undefined || !t.dataset.f) return;
    if (t.tagName === "INPUT" && e.type === "change") return;
    if (t.tagName === "SELECT" && e.type === "input") return;
    const r = this.rows[Number(t.dataset.k)];
    switch (t.dataset.f) {
      case "label": r.label = t.value; this.renderCells(); break;
      case "height": r.height = Number(t.value); break;
      case "style": r.style = t.value as CalendarRow["style"]; break;
      case "fill": r.fill = t.value || undefined; break;
      case "align": r.align = t.value as CalendarRow["align"]; break;
    }
    this.opts.onChange();
  }

  private onRowClick(e: Event) {
    const t = e.target as HTMLElement;
    if (!t.dataset?.act) return;
    const k = Number(t.dataset.k);
    if (t.dataset.act === "del") {
      if (this.rows.length <= 1) return;
      this.rows.splice(k, 1);
    } else if (t.dataset.act === "up" && k > 0) {
      [this.rows[k - 1], this.rows[k]] = [this.rows[k], this.rows[k - 1]];
    }
    this.render();
    this.opts.onChange();
  }

  private onCellInput(e: Event) {
    const t = e.target as HTMLTextAreaElement;
    if (t.dataset?.key === undefined) return;
    const r = this.rows[Number(t.dataset.k)];
    if (t.value.trim()) r.cells[t.dataset.key] = t.value;
    else delete r.cells[t.dataset.key];
    t.rows = Math.min(6, Math.max(1, t.value.split("\n").length));
    this.opts.onChange();
  }

  private onPaste(e: ClipboardEvent) {
    const t = e.target as HTMLElement;
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (t.dataset?.key === undefined || !/[\t\n]/.test(text.replace(/\r?\n$/, ""))) return; // plain paste into one cell
    const block = parseTSV(text);
    const order = this.opts.dateOrder();
    const date = (v: string | undefined) => (v && v.trim() ? parseDate(v, order) : null);
    const across = block[0] ? block[0].slice(1).filter((v) => date(v) !== null).length : 0;
    const down = block.slice(1).filter((row) => date(row[0]) !== null).length;
    e.preventDefault();
    if (across >= 2 || down >= 2) {
      // A whole calendar: dates across the top (rows = calendar rows) or down the left.
      const entries: { label: string; byDate: [number, string][] }[] = [];
      if (across >= down) {
        const dates = block[0].slice(1).map((v) => date(v));
        block.slice(1).forEach((row) => entries.push({ label: (row[0] ?? "").trim(), byDate: dates.map((d, i) => [d ?? NaN, row[i + 1] ?? ""] as [number, string]).filter(([d]) => !Number.isNaN(d)) }));
      } else {
        const header = date(block[0][0]) === null ? block[0] : null;
        const data = header ? block.slice(1) : block;
        const width = Math.max(...data.map((r) => r.length)) - 1;
        for (let j = 0; j < width; j++) {
          entries.push({ label: (header?.[j + 1] ?? "").trim(), byDate: data.map((row) => [date(row[0]) ?? NaN, row[j + 1] ?? ""] as [number, string]).filter(([d]) => !Number.isNaN(d)) });
        }
      }
      const all = entries.flatMap((x) => x.byDate.map(([d]) => d));
      this.opts.onRange(isoDay(Math.min(...all)), isoDay(Math.max(...all)));
      const cols = this.opts.columns();
      const keyOf = (d: number) => cols.find((c) => d >= c.start && d <= c.end)?.key;
      this.rows = entries.map((x, j) => {
        const base = this.rows[j] ?? newRow(j);
        const cells: Record<string, string> = {};
        for (const [d, v] of x.byDate) {
          const key = keyOf(d);
          if (!key || !v.trim()) continue;
          cells[key] = cells[key] ? `${cells[key]}\n${v.trim()}` : v.trim();
        }
        return { ...base, label: x.label || base.label, cells };
      });
    } else {
      // A block of text: fill from the focused cell (columns down, calendar rows across).
      const cols = this.opts.columns();
      block.forEach((brow, di) => brow.forEach((v, dj) => {
        const col = cols[this.focus.ci + di];
        const k = this.focus.k + dj;
        if (!col) return;
        while (this.rows.length <= k) this.rows.push(newRow(this.rows.length));
        if (v.trim()) this.rows[k].cells[col.key] = v.trim();
        else delete this.rows[k].cells[col.key];
      }));
    }
    this.render();
    this.opts.onChange();
  }
}
