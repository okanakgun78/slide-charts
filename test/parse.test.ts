import { describe, expect, it } from "vitest";
import { cellText, parseNumber, parseTSV } from "../src/ui/parse";

describe("parseNumber", () => {
  it.each([
    ["1,234.5", ".", 1234.5], ["1.234,5", ",", 1234.5], ["(12)", ".", -12], ["−7", ".", -7],
    ["$1,000", ".", 1000], ["45%", ".", 45], ["  3 ", ".", 3], ["1 234,5", ",", 1234.5], ["1e3", ".", 1000],
    ["", ".", null], ["n/a", ".", null], ["-", ".", null], ["12abc", ".", null],
  ] as const)("%s (%s) -> %s", (raw, sep, out) => {
    expect(parseNumber(raw, sep)).toBe(out);
  });
});

describe("parseTSV", () => {
  it("splits Excel clipboard text", () => {
    expect(parseTSV("\t2021\t2022\r\nEurope\t42\t46\r\n")).toEqual([["", "2021", "2022"], ["Europe", "42", "46"]]);
  });
});

describe("parseTSV with quoted cells", () => {
  it("keeps line breaks and tabs inside quoted cells", () => {
    expect(parseTSV('Day\tPlan\r\n02/11\t"1. Rest\n2. Stretching"\r\n03/11\t"say ""hi"""\r\n')).toEqual([
      ["Day", "Plan"], ["02/11", "1. Rest\n2. Stretching"], ["03/11", 'say "hi"'],
    ]);
  });
  it("leaves quotes inside unquoted cells alone", () => {
    expect(parseTSV('5" screen\tok')).toEqual([['5" screen', "ok"]]);
  });
  it("keeps trailing empty cells", () => {
    expect(parseTSV("a\t\tb\t\n")).toEqual([["a", "", "b", ""]]);
  });
});

describe("cellText", () => {
  it("round-trips", () => {
    expect(cellText(0.1 + 0.2)).toBe("0.3");
    expect(cellText(1234.5, ",")).toBe("1234,5");
    expect(cellText(null)).toBe("");
  });
});
