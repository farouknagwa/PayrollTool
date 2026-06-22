import * as XLSX from "xlsx";
import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS } from "../config/defaults";
import type { PayrollInputFiles } from "../core/types";
import { runPayrollPipeline } from "./payrollPipeline";

function workbookFile(name: string, rows: unknown[][], sheetName = "Worksheet"): File {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  const array = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  return {
    name,
    lastModified: Date.now(),
    webkitRelativePath: "",
    arrayBuffer: async () => array,
  } as File;
}

function baseInputs(permissionFile: File, raw = false): PayrollInputFiles {
  return {
    attendance: workbookFile("Attendance Report.xls", [
      ["Report"],
      ["Headers"],
      [101, "Synthetic One", "", "", "19/05/2026", "", "08:00 AM", "0:00", "", "", "", "04:00 PM"],
    ]),
    absences: workbookFile("Absence Report.xls", [
      ["Report"],
      ["Headers"],
      [102, "Synthetic Two", "", "", "19/05/2026"],
    ]),
    vacations: workbookFile("Employee Transactions_vacations.xls", [
      ["Report"],
      ["Headers"],
      [103, "Synthetic Three", "19/05/2026", "19/05/2026", "Unpaid Leave"],
    ]),
    [raw ? "rawPermissions" : "preparedPermissions"]: permissionFile,
  };
}

describe("payroll pipeline integration", () => {
  it("runs with prepared permissions", async () => {
    const permissions = workbookFile("Nagwa_Permission_Request_permission_details.xls", [
      [
        "Employee Code",
        "Employee Name",
        "Request Date",
        "Effective Date",
        "Start Time",
        "End Time",
        "Total Permission Period",
        "Time ",
        "Transaction Type",
        "Transaction Sub Type",
        "WF Template",
        "Status",
      ],
      [101, "Synthetic One", "19/05/2026", "19/05/2026", "09:00 AM", "10:00 AM", "1:00:00", "09:00 AM", "Permission", "Permission", "WF", "Approved"],
    ]);
    const result = await runPayrollPipeline(baseInputs(permissions), DEFAULT_SETTINGS, {
      month: 5,
      year: 2026,
      requestCutoffDays: 0,
      noRequestCutoff: false,
    });
    expect(result.detailedWorkbook.byteLength).toBeGreaterThan(1000);
    expect(result.finalWorkbook.byteLength).toBeGreaterThan(1000);
    expect(result.metrics.some((metric) => metric.step === "complete_final")).toBe(true);
    expect(result.period).toMatchObject({ start: "2026-04-21", end: "2026-05-20" });
  });

  it("runs with raw permissions and returns a prepared workbook", async () => {
    const rawPermissions = workbookFile("Nagwa_Permission_Request_Report.xls", [
      [
        "Branch Code",
        "Employee Code",
        "Employee Name",
        "Request Date",
        "Effective Date",
        "Start Time",
        "End Time",
        "Time ",
        "Transaction Type",
        "Transaction Sub Type",
        "WF Template",
        "Status",
        "Workflow ID",
      ],
      ["B", 101, "Synthetic One", "19/05/2026", "19/05/2026", "09:00 AM", "10:00 AM", "09:00 AM", "Permission", "Permission", "WF", "Approved", "W"],
    ]);
    const result = await runPayrollPipeline(baseInputs(rawPermissions, true), DEFAULT_SETTINGS, {
      month: 5,
      year: 2026,
      requestCutoffDays: 0,
      noRequestCutoff: false,
    });
    expect(result.preparedPermissionsWorkbook?.byteLength).toBeGreaterThan(500);
  });
});
