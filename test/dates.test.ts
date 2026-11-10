import { describe, expect, it } from "vitest";
import { fromDay, isoDay, parseDate, shortDate, toDay } from "../src/engine/dates";

const iso = (raw: string, order: "DMY" | "MDY" = "DMY", edge: "start" | "end" = "start") => {
  const d = parseDate(raw, order, edge);
  return d === null ? null : isoDay(d);
};

describe("parseDate", () => {
  it.each([
    ["2026-03-15", "DMY", "start", "2026-03-15"],
    ["2026-03-15T10:00:00", "DMY", "start", "2026-03-15"],
    ["15/03/2026", "DMY", "start", "2026-03-15"],
    ["03/15/2026", "DMY", "start", "2026-03-15"], // auto-resolved: 15 > 12
    ["04/05/2026", "DMY", "start", "2026-05-04"],
    ["04/05/2026", "MDY", "start", "2026-04-05"],
    ["4.5.26", "DMY", "start", "2026-05-04"],
    ["15 Mar 2026", "DMY", "start", "2026-03-15"],
    ["15-Mar-26", "MDY", "start", "2026-03-15"],
    ["Mar 15, 2026", "DMY", "start", "2026-03-15"],
    ["March 1st 2026", "DMY", "start", "2026-03-01"],
    ["15 Mart 2026", "DMY", "start", "2026-03-15"],
    ["3 Şubat 2027", "DMY", "start", "2027-02-03"],
    ["46096", "DMY", "start", "2026-03-15"], // Excel serial
    ["Mar 2026", "DMY", "start", "2026-03-01"],
    ["Mar 2026", "DMY", "end", "2026-03-31"],
    ["2026-02", "DMY", "end", "2026-02-28"],
    ["Q1 2026", "DMY", "start", "2026-01-01"],
    ["Q1 2026", "DMY", "end", "2026-03-31"],
    ["2026 Q4", "DMY", "end", "2026-12-31"],
    ["H2 2027", "DMY", "start", "2027-07-01"],
    ["Q3-26", "DMY", "end", "2026-09-30"],
    ["2027", "DMY", "end", "2027-12-31"],
  ] as const)("%s (%s, %s) -> %s", (raw, order, edge, out) => {
    expect(iso(raw, order, edge)).toBe(out);
  });
  it.each(["", "31/02/2026", "13/13/2026", "hello", "Q5 2026", "H3 2026", "12345"])("rejects %s", (raw) => {
    expect(parseDate(raw)).toBeNull();
  });
});

describe("day numbers", () => {
  it("round-trip and weekday", () => {
    const n = toDay(2026, 11, 10);
    expect(fromDay(n)).toMatchObject({ y: 2026, m: 11, d: 10, wd: 1 }); // Tuesday
    expect(shortDate(n, "DMY")).toBe("10 Nov");
    expect(shortDate(n, "MDY")).toBe("Nov 10");
  });
});
