import * as XLSX from "xlsx";
import { parseDateValue } from "../core/dateTime";

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

function sheetRows(workbook: XLSX.WorkBook): unknown[][] {
  return XLSX.utils.sheet_to_json<unknown[]>(firstSheet(workbook), {
    header: 1,
    raw: true,
    defval: "",
  });
}

function asText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function valueAt(row: unknown[], index: number): unknown {
  return row[index] ?? "";
}

// Lightweight, browser-side preview of the payroll period. The authoritative
// period detection happens inside the Python pipeline; this only drives the
// "Detected period" hint and the permission month/year defaults in the UI.
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
  return null;
}
