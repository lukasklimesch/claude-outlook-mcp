import { describe, expect, test } from "bun:test";
import {
  addDays,
  normalizeLocalDateTime,
  startOfDay,
  toRestrictFormat,
} from "../src/util/dates.js";

describe("normalizeLocalDateTime", () => {
  test("preserves local wall-clock components", () => {
    expect(normalizeLocalDateTime("2026-06-05T14:00:00")).toBe(
      "2026-06-05T14:00:00",
    );
  });

  test("zero-pads single digit fields", () => {
    expect(normalizeLocalDateTime("2026-01-02T03:04:05")).toBe(
      "2026-01-02T03:04:05",
    );
  });

  test("rejects invalid input", () => {
    expect(() => normalizeLocalDateTime("not-a-date")).toThrow();
    expect(() => normalizeLocalDateTime("")).toThrow();
  });

  test("pins a bare date to local midnight (no timezone day-shift)", () => {
    // Regardless of the host timezone, a date-only value must keep its day.
    expect(normalizeLocalDateTime("2026-06-05")).toBe("2026-06-05T00:00:00");
  });
});

describe("toRestrictFormat", () => {
  test("formats US-style 12-hour with AM/PM", () => {
    expect(toRestrictFormat(new Date(2026, 5, 5, 14, 30, 0))).toBe(
      "06/05/2026 02:30 PM",
    );
  });

  test("midnight renders as 12 AM", () => {
    expect(toRestrictFormat(new Date(2026, 0, 1, 0, 0, 0))).toBe(
      "01/01/2026 12:00 AM",
    );
  });

  test("noon renders as 12 PM", () => {
    expect(toRestrictFormat(new Date(2026, 11, 31, 12, 5, 0))).toBe(
      "12/31/2026 12:05 PM",
    );
  });
});

describe("startOfDay / addDays", () => {
  test("startOfDay zeroes the time", () => {
    const s = startOfDay(new Date(2026, 5, 5, 13, 22, 9));
    expect(s.getHours()).toBe(0);
    expect(s.getMinutes()).toBe(0);
    expect(s.getDate()).toBe(5);
  });

  test("addDays advances and does not mutate input", () => {
    const base = new Date(2026, 5, 5, 9, 0, 0);
    const plus2 = addDays(base, 2);
    expect(plus2.getDate()).toBe(7);
    expect(base.getDate()).toBe(5);
  });
});
