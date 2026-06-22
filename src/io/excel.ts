import * as XLSX from "xlsx";
import { formatTimeValue, parseDateValue } from "../core/dateTime";
import type {
  AbsenceRecord,
  AttendanceRecord,
  PermissionRecord,
  PublicHolidayRecord,
  ResignationRecord,
  RosterEmployee,
  VacationRecord,
} from "../core/types";

export async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  const buffer = typeof file.arrayBuffer === "function"
    ? await file.arrayBuffer()
    : await new Response(file).arrayBuffer();
  return XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    dense: false,
  });
}

function firstSheet(workbook: XLSX.WorkBook): XLSX.WorkSheet {
  const name = workbook.SheetNames[0];
  const sheet = workbook.Sheets[name];
  if (!sheet) throw new Error("Workbook has no worksheets.");
  return sheet;
}

export function sheetRows(workbook: XLSX.WorkBook, sheetName?: string): unknown[][] {
  const sheet = sheetName && workbook.Sheets[sheetName] ? workbook.Sheets[sheetName] : firstSheet(workbook);
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    raw: true,
    defval: "",
  });
}

function asText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function asCode(value: unknown): number | null {
  const text = asText(value);
  if (!text) return null;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
}

function valueAt(row: unknown[], index: number): unknown {
  return row[index] ?? "";
}

export function parseAttendance(workbook: XLSX.WorkBook): AttendanceRecord[] {
  const rows = sheetRows(workbook);
  const dataRows = rows.slice(2);
  const out: AttendanceRecord[] = [];
  for (const row of dataRows) {
    const code = asCode(valueAt(row, 0));
    const attendanceDay = parseDateValue(valueAt(row, 4));
    if (code === null || attendanceDay === null) continue;
    out.push({
      code,
      name: asText(valueAt(row, 1)),
      attendanceDay,
      entryTime: formatTimeValue(valueAt(row, 6)),
      late: asText(valueAt(row, 7)) || "0:00",
      exitTime: formatTimeValue(valueAt(row, 11)),
    });
  }
  return out;
}

export function detectAttendancePeriodDate(workbook: XLSX.WorkBook): string | null {
  const rows = sheetRows(workbook);
  const headerIndex = rows.findIndex((row) => row.some((cell) => asText(cell) === "Attendance Day"));
  if (headerIndex >= 0) {
    const colIndex = rows[headerIndex].findIndex((cell) => asText(cell) === "Attendance Day");
    for (const row of rows.slice(headerIndex + 1)) {
      const parsed = parseDateValue(valueAt(row, colIndex));
      if (parsed !== null) return parsed;
    }
  }
  return parseAttendance(workbook)[0]?.attendanceDay ?? null;
}

export function parseAbsences(workbook: XLSX.WorkBook): AbsenceRecord[] {
  return sheetRows(workbook).slice(2).flatMap((row) => {
    const code = asCode(valueAt(row, 0));
    const absenceDate = parseDateValue(valueAt(row, 4));
    return code !== null && absenceDate !== null ? [{ code, absenceDate }] : [];
  });
}

export function parseVacations(workbook: XLSX.WorkBook): VacationRecord[] {
  return sheetRows(workbook).slice(2).flatMap((row) => {
    const code = asCode(valueAt(row, 0));
    const startDate = parseDateValue(valueAt(row, 2));
    const endDate = parseDateValue(valueAt(row, 3));
    const vacationType = asText(valueAt(row, 4));
    return code !== null && startDate !== null && endDate !== null
      ? [{ code, startDate, endDate, vacationType }]
      : [];
  });
}

function objectRows(workbook: XLSX.WorkBook): Record<string, unknown>[] {
  const sheet = firstSheet(workbook);
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { raw: true, defval: "" }).map((row) => {
    const normalized: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(row)) {
      normalized[key === "Time " ? "Time " : key.trim()] = value;
    }
    return normalized;
  });
}

export function parseResignations(workbook: XLSX.WorkBook): ResignationRecord[] {
  return objectRows(workbook).flatMap((row) => {
    const code = asCode(row.ID);
    const resignationDate = parseDateValue(row["Resignation Date"]);
    return code !== null && resignationDate !== null ? [{ code, resignationDate }] : [];
  });
}

export function parsePublicHolidays(workbook: XLSX.WorkBook): PublicHolidayRecord[] {
  return objectRows(workbook).flatMap((row) => {
    const holidayDate = parseDateValue(row["Date of Public Holiday"]);
    return holidayDate !== null ? [{ holidayDate }] : [];
  });
}

export function parsePreparedPermissions(workbook: XLSX.WorkBook): PermissionRecord[] {
  return objectRows(workbook).flatMap((row) => {
    const employeeCode = asCode(row["Employee Code"]);
    if (employeeCode === null) return [];
    return [
      {
        employeeCode,
        employeeName: asText(row["Employee Name"]),
        requestDate: parseDateValue(row["Request Date"]),
        effectiveDate: parseDateValue(row["Effective Date"]),
        startTime: formatTimeValue(row["Start Time"]),
        endTime: formatTimeValue(row["End Time"]),
        totalPermissionPeriod: asText(row["Total Permission Period"]),
        time: formatTimeValue(row["Time "]),
        transactionType: asText(row["Transaction Type"]),
        transactionSubType: asText(row["Transaction Sub Type"]),
        wfTemplate: asText(row["WF Template"]),
        status: asText(row.Status),
      },
    ];
  });
}

export function detectRawPermissionWorkbook(workbook: XLSX.WorkBook): boolean {
  const rows = sheetRows(workbook);
  const header = rows[0]?.map(asText) ?? [];
  return header.includes("Workflow ID") || header.includes("Branch Code") || !header.includes("Total Permission Period");
}

export function parseRosterFromTemplate(workbook: XLSX.WorkBook): RosterEmployee[] {
  const rows = sheetRows(workbook, "Nagwa Technologies");
  return rows.slice(3).flatMap((row) => {
    const code = asCode(valueAt(row, 0));
    if (code === null) return [];
    const name = asText(valueAt(row, 1)) || `Employee ${code}`;
    const schedule = asText(valueAt(row, 4)).toLowerCase() || "standard";
    return [{ code, name, schedule }];
  });
}

export function buildRosterFromReports(attendance: AttendanceRecord[], templateRoster: RosterEmployee[]): RosterEmployee[] {
  const byCode = new Map<number, RosterEmployee>();
  for (const employee of templateRoster) byCode.set(employee.code, employee);
  for (const row of attendance) {
    if (!byCode.has(row.code)) {
      byCode.set(row.code, {
        code: row.code,
        name: row.name || `Employee ${row.code}`,
        schedule: "standard",
      });
    }
  }
  return [...byCode.values()].sort((a, b) => a.code - b.code);
}

export function worksheetToArrayBuffer(workbook: XLSX.WorkBook, bookType: XLSX.BookType): ArrayBuffer {
  const array = XLSX.write(workbook, { type: "array", bookType });
  return array instanceof ArrayBuffer ? array : array.buffer.slice(array.byteOffset, array.byteOffset + array.byteLength);
}
