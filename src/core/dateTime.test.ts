import { describe, expect, it } from "vitest";
import { computePeriod, periodFromMonth } from "./dateTime";

describe("payroll period helpers", () => {
  it("detects the 21-to-20 period containing a date", () => {
    expect(computePeriod("2026-04-28")).toMatchObject({
      start: "2026-04-21",
      end: "2026-05-20",
    });
  });

  it("builds a period from the month that ends on the 20th", () => {
    expect(periodFromMonth(5, 2026)).toMatchObject({
      start: "2026-04-21",
      end: "2026-05-20",
    });
  });
});
