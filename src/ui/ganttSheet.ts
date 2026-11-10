// Data sheet for Gantt charts: one row per task, milestone or group header.

import { parseDate, type DateOrder } from "../engine/dates";
import { MILESTONE_COLOR } from "../engine/gantt";
import { HIGHLIGHTS } from "../engine/spec";
import type { GanttRow, GanttRowKind } from "../engine/types";
import { parseTSV } from "./parse";

export interface GanttSheetOptions {
  palette: () => string[];
  dateOrder: () => DateOrder;
  onChange: () => void;
}

const COLS = ["label", "kind", "start", "end", "note"] as const;
type Col = (typeof COLS)[number];

export class GanttSheet {
  private rows: GanttRow[] = [];
  private focus = { r: 0, c: 0 };
  private popover: HTMLDivElement | null = null;

  constructor(private root: HTMLElement, private opts: GanttSheetOptions) {
    root.addEventListener("input", (e) => this.onInput(e));
    root.addEventListener("change", (e) => this.onInput(e));
    root.addEventListener("focusin", (e) => this.onFocus(e));
    root.addEventListener("paste", (e) => this.onPaste(e as ClipboardEvent));
    root.addEventListener("click", (e) => this.onClick(e));
    document.addEventListener("click", (e) => {
      if (this.popover && !this.popover.contains(e.target as Node) && !(e.target as HTMLElement).classList.contains("sw")) this.closePopover();
    });
  }

  load(rows: GanttRow[]) {
    this.rows = rows.map((r) => ({ ...r }));
    this.render();
  }

  data(): GanttRow[] {
    return this.rows.map((r) => ({ ...r, label: r.label.trim() }));
  }

  addRow() {
    const at = Math.min(this.focus.r + 1, this.rows.length);
    this.rows.splice(at, 0, { label: "", kind: "task", start: "", end: "" });
    this.focus = { r: at, c: 0 };
    this.changed();
  }

  deleteRow(): string {
    if (this.rows.length <= 1) return "A Gantt chart needs at least one row.";
    this.rows.splice(this.focus.r, 1);
    this.focus.r = Math.min(this.focus.r, this.rows.length - 1);
    this.changed();
    return "";
  }

  move(delta: -1 | 1) {
    const r = this.focus.r, to = r + delta;
    if (to < 0 || to >= this.rows.length) return;
    [this.rows[r], this.rows[to]] = [this.rows[to], this.rows[r]];
    this.focus.r = to;
    this.changed();
  }

  render() {
    this.closePopover();
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    const order = this.opts.dateOrder();
    const kinds: [GanttRowKind, string][] = [["task", "Task"], ["milestone", "Milestone"], ["header", "Group"]];
    const body = this.rows.map((row, r) => {
      const auto = row.kind === "milestone" ? MILESTONE_COLOR : row.kind === "header" ? "#1A1A1A" : this.opts.palette()[0];
      const bad = (v: string, edge: "start" | "end") => (v.trim() && parseDate(v, order, edge) === null ? " bad" : "");
      const noDates = row.kind === "header";
      return `<tr class="${row.kind}">
        <td><button class="sw" data-r="${r}" title="Colour" style="background:${row.color ?? auto}"></button></td>
        <td><input class="name" data-r="${r}" data-c="label" value="${esc(row.label)}" placeholder="Label"></td>
        <td><select data-r="${r}" data-c="kind">${kinds.map(([k, t]) => `<option value="${k}" ${k === row.kind ? "selected" : ""}>${t}</option>`).join("")}</select></td>
        <td><input class="date${bad(row.start, "start")}" data-r="${r}" data-c="start" value="${esc(row.start)}" placeholder="${noDates ? "–" : "Start"}" ${noDates ? "disabled" : ""}></td>
        <td><input class="date${bad(row.end, "end")}" data-r="${r}" data-c="end" value="${esc(row.end)}" placeholder="${row.kind === "task" ? "End" : "–"}" ${row.kind !== "task" ? "disabled" : ""}></td>
        <td><input class="note" data-r="${r}" data-c="note" value="${esc(row.note ?? "")}" placeholder="Note"></td>
      </tr>`;
    }).join("");
    this.root.innerHTML = `<table class="sheet gantt"><thead><tr><th class="corner"></th><th class="corner">Label</th><th class="corner">Type</th><th class="corner">Start</th><th class="corner">End</th><th class="corner">Note</th></tr></thead><tbody>${body}</tbody></table>`;
  }

  focusCell(r: number, c: Col) {
    this.focus = { r, c: COLS.indexOf(c) };
    this.root.querySelector<HTMLInputElement>(`[data-r="${r}"][data-c="${c}"]`)?.focus();
  }

  private changed() {
    this.render();
    this.focusCell(this.focus.r, COLS[this.focus.c] ?? "label");
    this.opts.onChange();
  }

  private onFocus(e: Event) {
    const t = e.target as HTMLElement;
    if (t.dataset?.r === undefined) return;
    this.focus = { r: Number(t.dataset.r), c: Math.max(0, COLS.indexOf(t.dataset.c as Col)) };
  }

  private onInput(e: Event) {
    const t = e.target as HTMLInputElement | HTMLSelectElement;
    if (t.dataset?.r === undefined || !t.dataset.c) return;
    const r = Number(t.dataset.r);
    const c = t.dataset.c as Col;
    if (c === "kind") {
      if (e.type !== "change") return;
      this.rows[r].kind = t.value as GanttRowKind;
      this.changed();
      return;
    }
    if (e.type === "change") return; // inputs are handled on "input"
    (this.rows[r] as unknown as Record<string, string>)[c] = t.value;
    if (c === "start" || c === "end") {
      const bad = t.value.trim() !== "" && parseDate(t.value, this.opts.dateOrder(), c) === null;
      t.classList.toggle("bad", bad);
    }
    this.opts.onChange();
  }

  /**
   * Paste rows from Excel: Label | Start | End | Note (End and Note optional).
   * Rows without dates become group headers; rows with only a start become milestones.
   * A first row of column titles ("Task, Start, End") is skipped.
   */
  private onPaste(e: ClipboardEvent) {
    const t = e.target as HTMLElement;
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (t.dataset?.r === undefined || !/[\t\n]/.test(text.replace(/\r?\n$/, ""))) return;
    e.preventDefault();
    const order = this.opts.dateOrder();
    let block = parseTSV(text).filter((row) => row.some((v) => v.trim()));
    const isDate = (v: string | undefined) => !!v && parseDate(v, order) !== null;
    // Column titles: text (not dates) in the date columns, while later rows do have dates.
    const b0 = block[0] ?? [];
    const titles = !isDate(b0[1]) && !isDate(b0[2]) && !!((b0[1] ?? "").trim() || (b0[2] ?? "").trim());
    if (block.length > 1 && titles && block.slice(1).some((row) => isDate(row[1]) || isDate(row[2]))) block = block.slice(1);
    const parsed: GanttRow[] = block.map((cells) => {
      const [label = "", start = "", end = "", note = ""] = cells.map((v) => v.trim());
      const kind: GanttRowKind = !start && !end ? "header" : !end ? "milestone" : "task";
      const row: GanttRow = { label, kind, start: kind === "header" ? "" : start, end: kind === "task" ? end : "" };
      if (note) row.note = note;
      return row;
    });
    const r0 = Number(t.dataset.r);
    const whole = r0 === 0 && t.dataset.c === "label";
    if (whole) this.rows = parsed;
    else this.rows.splice(r0, Math.min(parsed.length, this.rows.length - r0), ...parsed);
    this.focus = { r: whole ? 0 : r0, c: 0 };
    this.changed();
  }

  private onClick(e: Event) {
    const t = e.target as HTMLElement;
    if (!t.classList.contains("sw") || t.dataset.r === undefined) return;
    const r = Number(t.dataset.r);
    this.closePopover();
    const pop = document.createElement("div");
    pop.className = "popover";
    pop.innerHTML = [...this.opts.palette(), ...HIGHLIGHTS].map((c) => `<button class="sw pick" data-c="${c}" style="background:${c}"></button>`).join("") + `<button class="auto" data-c="">Auto</button>`;
    pop.addEventListener("click", (ev) => {
      const b = ev.target as HTMLElement;
      if (b.dataset.c === undefined) return;
      this.rows[r].color = b.dataset.c || undefined;
      this.changed();
    });
    const rect = t.getBoundingClientRect();
    pop.style.left = `${rect.left + window.scrollX}px`;
    pop.style.top = `${rect.bottom + window.scrollY + 4}px`;
    document.body.appendChild(pop);
    this.popover = pop;
  }

  private closePopover() {
    this.popover?.remove();
    this.popover = null;
  }
}
