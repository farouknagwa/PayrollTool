import * as XLSX from "xlsx";
import {
  addDays,
  formatDDMMYYYY,
  formatTimeValue,
  parseDateValue,
  periodFromMonth,
  permissionDuration,
  minutesToHMS,
} from "../core/dateTime";
import type { PermissionPrepOptions, PermissionPrepResult, PermissionRecord } from "../core/types";
import { parsePreparedPermissions, readWorkbook, worksheetToArrayBuffer } from "../io/excel";

export const PERMISSION_OUTPUT_COLUMNS = [
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
] as const;

export const COLUMN_WIDTHS = [4522, 10240, 4096, 4181, 3541, 3328, 6186, 2602, 5760, 8576, 6058, 2773];

const REQUIRED_RAW_COLUMNS = [
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
];

function normalizeHeader(value: string): string {
  return value === "Time " ? "Time " : value.trim();
}

function rowsFromWorkbook(workbook: XLSX.WorkBook): Record<string, unknown>[] {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true, defval: "" }).map((row) => {
    const normalized: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(row)) normalized[normalizeHeader(key)] = value;
    return normalized;
  });
}

function ensureRawColumns(rows: Record<string, unknown>[]): void {
  const keys = new Set(rows.length > 0 ? Object.keys(rows[0]) : REQUIRED_RAW_COLUMNS);
  const missing = REQUIRED_RAW_COLUMNS.filter((column) => !keys.has(column));
  if (missing.length > 0) throw new Error(`Missing required permission column(s): ${missing.join(", ")}`);
}

function asText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function asEmployeeCode(value: unknown): number {
  const parsed = Number(asText(value));
  if (!Number.isFinite(parsed)) throw new Error(`Invalid Employee Code '${asText(value)}'.`);
  return Math.trunc(parsed);
}

function rawRowToPermission(row: Record<string, unknown>): PermissionRecord {
  const durationMinutes = permissionDuration(row["Start Time"], row["End Time"]);
  const requestDate = parseDateValue(row["Request Date"]);
  const effectiveDate = parseDateValue(row["Effective Date"]);
  return {
    employeeCode: asEmployeeCode(row["Employee Code"]),
    employeeName: asText(row["Employee Name"]),
    requestDate,
    effectiveDate,
    startTime: formatTimeValue(row["Start Time"]),
    endTime: formatTimeValue(row["End Time"]),
    totalPermissionPeriod: minutesToHMS(durationMinutes),
    time: formatTimeValue(row["Time "]),
    transactionType: asText(row["Transaction Type"]),
    transactionSubType: asText(row["Transaction Sub Type"]),
    wfTemplate: asText(row["WF Template"]),
    status: asText(row["Status"]),
  };
}

export function transformRawPermissionRows(rows: Record<string, unknown>[], options: PermissionPrepOptions): PermissionPrepResult {
  ensureRawColumns(rows);
  const period = periodFromMonth(options.month, options.year);
  const warnings: string[] = [];
  let excludedByCutoff = 0;
  const cutoffIso = options.noRequestCutoff ? null : addDays(period.end, options.requestCutoffDays);

  const kept: PermissionRecord[] = [];
  for (const row of rows) {
    const effectiveDate = parseDateValue(row["Effective Date"]);
    if (effectiveDate === null) {
      warnings.push("A row with unparseable Effective Date was skipped.");
      continue;
    }
    if (effectiveDate < period.start || effectiveDate > period.end) continue;
    if (asText(row["Status"]) !== "Approved") continue;
    const requestDate = parseDateValue(row["Request Date"]);
    if (!options.noRequestCutoff) {
      if (requestDate === null) {
        warnings.push("An approved in-period row with unparseable Request Date was skipped.");
        excludedByCutoff += 1;
        continue;
      }
      if (cutoffIso !== null && requestDate > cutoffIso) {
        excludedByCutoff += 1;
        continue;
      }
    }
    kept.push(rawRowToPermission(row));
  }

  return {
    rows: kept,
    period,
    loaded: rows.length,
    kept: kept.length,
    excludedByCutoff,
    warnings,
  };
}

export async function preparePermissionFile(file: File, options: PermissionPrepOptions): Promise<PermissionPrepResult> {
  const workbook = await readWorkbook(file);
  const rows = rowsFromWorkbook(workbook);
  return transformRawPermissionRows(rows, options);
}

export async function readPreparedPermissionFile(file: File): Promise<PermissionRecord[]> {
  const workbook = await readWorkbook(file);
  return parsePreparedPermissions(workbook);
}

function serialTimeFromHms(hms: string): number {
  const [hours, minutes, seconds] = hms.split(":").map(Number);
  return ((hours || 0) * 3600 + (minutes || 0) * 60 + (seconds || 0)) / 86400;
}

export function createPreparedPermissionWorkbook(rows: PermissionRecord[]): XLSX.WorkBook {
  const aoa: unknown[][] = [PERMISSION_OUTPUT_COLUMNS.slice()];
  for (const row of rows) {
    aoa.push([
      row.employeeCode,
      row.employeeName,
      formatDDMMYYYY(row.requestDate),
      formatDDMMYYYY(row.effectiveDate),
      row.startTime,
      row.endTime,
      serialTimeFromHms(row.totalPermissionPeriod),
      row.time,
      row.transactionType,
      row.transactionSubType,
      row.wfTemplate,
      row.status,
    ]);
  }
  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  sheet["!cols"] = COLUMN_WIDTHS.map((width) => ({ wch: Math.round(width / 256) }));
  sheet["!freeze"] = { xSplit: 0, ySplit: 1 };
  for (let rowIndex = 1; rowIndex <= rows.length; rowIndex += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: 6 })];
    if (cell) cell.z = "h:mm";
  }
  return {
    SheetNames: ["Worksheet"],
    Sheets: {
      Worksheet: sheet,
    },
  };
}

export function writePreparedPermissionWorkbook(rows: PermissionRecord[], format: "xls" | "xlsx" = "xls"): ArrayBuffer {
  return worksheetToArrayBuffer(createPreparedPermissionWorkbook(rows), format === "xls" ? "biff8" : "xlsx");
}
