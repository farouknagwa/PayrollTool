import { describe, expect, it } from "vitest";
import { transformRawPermissionRows } from "./prepare";

describe("permission preparation", () => {
  const rows = [
    {
      "Employee Code": 101,
      "Employee Name": "Synthetic One",
      "Request Date": "20/05/2026",
      "Effective Date": "19/05/2026",
      "Start Time": "09:00 AM",
      "End Time": "10:30 AM",
      "Time ": "09:00 AM",
      "Transaction Type": "Permission",
      "Transaction Sub Type": "Permission",
      "WF Template": "WF",
      Status: "Approved",
      "Branch Code": "B",
      "Workflow ID": "W",
    },
    {
      "Employee Code": 102,
      "Employee Name": "Synthetic Two",
      "Request Date": "25/05/2026",
      "Effective Date": "19/05/2026",
      "Start Time": "09:00 AM",
      "End Time": "10:00 AM",
      "Time ": "09:00 AM",
      "Transaction Type": "Permission",
      "Transaction Sub Type": "Permission",
      "WF Template": "WF",
      Status: "Approved",
    },
    {
      "Employee Code": 103,
      "Employee Name": "Synthetic Three",
      "Request Date": "19/05/2026",
      "Effective Date": "18/05/2026",
      "Start Time": "09:00 AM",
      "End Time": "10:00 AM",
      "Time ": "09:00 AM",
      "Transaction Type": "Permission",
      "Transaction Sub Type": "Permission",
      "WF Template": "WF",
      Status: "Rejected",
    },
  ];

  it("filters approved in-period rows and applies cutoff", () => {
    const result = transformRawPermissionRows(rows, {
      month: 5,
      year: 2026,
      requestCutoffDays: 0,
      noRequestCutoff: false,
    });
    expect(result.loaded).toBe(3);
    expect(result.kept).toBe(1);
    expect(result.excludedByCutoff).toBe(1);
    expect(result.rows[0].totalPermissionPeriod).toBe("1:30:00");
  });

  it("can disable request cutoff", () => {
    const result = transformRawPermissionRows(rows, {
      month: 5,
      year: 2026,
      requestCutoffDays: 0,
      noRequestCutoff: true,
    });
    expect(result.kept).toBe(2);
  });
});
