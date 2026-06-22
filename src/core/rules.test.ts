import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS } from "../config/defaults";
import { computePeriod, minutesToHHMM, parseDuration, periodFromMonth } from "./dateTime";
import { calculateShortage, classifyLeaveType, hasRestrictedWindow, permittedWindowEnd } from "./rules";

describe("period utilities", () => {
  it("computes the containing HR period", () => {
    expect(computePeriod("2026-04-28")).toMatchObject({ start: "2026-04-21", end: "2026-05-20" });
    expect(computePeriod("2026-01-10")).toMatchObject({ start: "2025-12-21", end: "2026-01-20" });
    expect(periodFromMonth(5, 2026)).toMatchObject({ start: "2026-04-21", end: "2026-05-20" });
  });

  it("formats and parses durations like Python", () => {
    expect(parseDuration("1:30")).toBe(90);
    expect(parseDuration("1:29:30")).toBe(90);
    expect(minutesToHHMM(75)).toBe("1:15");
  });
});

describe("shortage windows", () => {
  it("uses schedule-based permitted windows", () => {
    expect(permittedWindowEnd("2026-04-22", 55, "flexy", DEFAULT_SETTINGS)).toBe("18:00");
    expect(permittedWindowEnd("2026-04-22", 55, "unknown", DEFAULT_SETTINGS)).toBe("16:00");
  });

  it("uses special-rule alternating employees", () => {
    expect(hasRestrictedWindow(1052, "2026-02-21", DEFAULT_SETTINGS)).toBe(true);
    expect(hasRestrictedWindow(100, "2026-03-21", DEFAULT_SETTINGS)).toBe(true);
  });

  it("uses Ramadan permitted window end from Python source", () => {
    expect(permittedWindowEnd("2026-02-25", 55, "flexy", DEFAULT_SETTINGS)).toBe("16:30");
  });

  it("calculates presence overlap shortage", () => {
    expect(calculateShortage("0:10", "08:00 AM", "04:00 PM", "2026-04-22", 55, "standard", DEFAULT_SETTINGS)).toBe("0:10");
    expect(calculateShortage("0:00", "10:00 AM", "04:00 PM", "2026-04-22", 55, "standard", DEFAULT_SETTINGS)).toBe("2:00");
    expect(calculateShortage("0:00", "08:00 AM", "02:30 PM", "2026-02-25", 55, "standard", DEFAULT_SETTINGS)).toBe("0:00");
  });
});

describe("leave classification", () => {
  it("classifies exact and generic half-day labels", () => {
    expect(classifyLeaveType("1st Half Day Annual", 8 * 60, 12 * 60)).toBe("1st_half_annual");
    expect(classifyLeaveType("Half Day Annual (Daylight Saving)", 13 * 60, 16 * 60)).toBe("2nd_half_annual");
  });

  it("classifies permission-style entries", () => {
    expect(classifyLeaveType("Permission")).toBe("permission");
    expect(classifyLeaveType("Educational Leave - Core of Business")).toBe("permission");
  });
});
