import { describe, expect, it } from "vitest";
import { addDays, hasExplicitOffset, parseDateOnly, parseDateTimeInput, parseRangeInput } from "../src/owa/time.js";

describe("parseDateTimeInput", () => {
  it("converts values with an offset to UTC", () => {
    expect(parseDateTimeInput("2026-11-02T14:30:00+01:00", "start")).toEqual({ value: "2026-11-02T13:30:00Z", floating: false });
    expect(parseDateTimeInput("2026-11-02T14:30:00.000Z", "start")).toEqual({ value: "2026-11-02T14:30:00Z", floating: false });
  });

  it("keeps values without an offset as floating wall-clock time", () => {
    expect(parseDateTimeInput("2026-11-02T14:30", "start")).toEqual({ value: "2026-11-02T14:30:00", floating: true });
    expect(parseDateTimeInput("2026-11-02", "start")).toEqual({ value: "2026-11-02T00:00:00", floating: true });
  });

  it("converts wall-clock time in an IANA zone, including daylight saving time", () => {
    expect(parseDateTimeInput("2027-06-05T10:00:00", "start", "Europe/London").value).toBe("2027-06-05T09:00:00Z");
    expect(parseDateTimeInput("2026-12-05T10:00:00", "start", "Europe/London").value).toBe("2026-12-05T10:00:00Z");
    expect(parseDateTimeInput("2026-07-01T09:00:00", "start", "America/New_York").value).toBe("2026-07-01T13:00:00Z");
    expect(parseDateTimeInput("2026-01-15T09:00:00", "start", "Asia/Kolkata").value).toBe("2026-01-15T03:30:00Z");
  });

  it("rejects invalid input", () => {
    expect(() => parseDateTimeInput("tomorrow", "start")).toThrow(/not a valid ISO 8601/);
    expect(() => parseDateTimeInput("2026-02-30T10:00", "start")).toThrow();
    expect(() => parseDateTimeInput("2026-11-02T10:00", "start", "Mars/Olympus")).toThrow(/Unknown IANA time zone/);
  });
});

describe("helpers", () => {
  it("detects explicit offsets", () => {
    expect(hasExplicitOffset("2026-11-02T10:00:00Z")).toBe(true);
    expect(hasExplicitOffset("2026-11-02T10:00:00-0500")).toBe(true);
    expect(hasExplicitOffset("2026-11-02T10:00:00")).toBe(false);
    expect(hasExplicitOffset("2026-11-02")).toBe(false);
  });

  it("treats range values without an offset as UTC", () => {
    expect(parseRangeInput("2026-11-02", "start").toISOString()).toBe("2026-11-02T00:00:00.000Z");
  });

  it("handles all-day dates", () => {
    expect(parseDateOnly("2026-11-02", "start")).toBe("2026-11-02");
    expect(parseDateOnly("2026-11-02T00:00:00", "start")).toBe("2026-11-02");
    expect(() => parseDateOnly("2026-11-02T10:00", "start")).toThrow(/all-day/);
    expect(addDays("2026-12-31", 1)).toBe("2027-01-01");
  });
});
