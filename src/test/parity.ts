import * as XLSX from "xlsx";

export interface WorkbookDiff {
  sheet: string;
  cell: string;
  expected: unknown;
  actual: unknown;
}

export function diffWorkbookArrayBuffers(expected: ArrayBuffer, actual: ArrayBuffer): WorkbookDiff[] {
  const expectedBook = XLSX.read(expected, { type: "array", cellDates: true });
  const actualBook = XLSX.read(actual, { type: "array", cellDates: true });
  const diffs: WorkbookDiff[] = [];
  for (const sheetName of expectedBook.SheetNames) {
    const expectedSheet = expectedBook.Sheets[sheetName];
    const actualSheet = actualBook.Sheets[sheetName];
    if (!actualSheet) {
      diffs.push({ sheet: sheetName, cell: "<sheet>", expected: "present", actual: "missing" });
      continue;
    }
    const expectedRange = XLSX.utils.decode_range(expectedSheet["!ref"] ?? "A1:A1");
    const actualRange = XLSX.utils.decode_range(actualSheet["!ref"] ?? "A1:A1");
    const maxRow = Math.max(expectedRange.e.r, actualRange.e.r);
    const maxCol = Math.max(expectedRange.e.c, actualRange.e.c);
    for (let row = 0; row <= maxRow; row += 1) {
      for (let col = 0; col <= maxCol; col += 1) {
        const cell = XLSX.utils.encode_cell({ r: row, c: col });
        const expectedValue = expectedSheet[cell]?.v ?? "";
        const actualValue = actualSheet[cell]?.v ?? "";
        if (String(expectedValue) !== String(actualValue)) {
          diffs.push({ sheet: sheetName, cell, expected: expectedValue, actual: actualValue });
        }
      }
    }
  }
  return diffs;
}
