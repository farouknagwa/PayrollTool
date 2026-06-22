import ExcelJS from "exceljs";
import {
  buildDateRange,
  computePeriod,
  minutesToHMS,
  parseDuration,
  parseISODate,
} from "../core/dateTime";
import {
  applyHourReduction,
  applyPermittedDelays,
  applyWfhAndWorkdayOverrides,
  applyWorkMissionLunchExemption,
  calculateShortage,
  fillMissingPunches,
  joinLeaveEntry,
  recalculateShortageFromLeave,
} from "../core/rules";
import type {
  DetailedWorkbookModel,
  ISODate,
  PayrollInputFiles,
  PayrollRunResult,
  PayrollSettings,
  ParsedReports,
  PermissionPrepOptions,
  RunLogEntry,
  StepMetrics,
} from "../core/types";
import {
  buildRosterFromReports,
  detectAttendancePeriodDate,
  parseAbsences,
  parseAttendance,
  parsePublicHolidays,
  parseResignations,
  parseRosterFromTemplate,
  parseVacations,
  readWorkbook,
} from "../io/excel";
import { missingRequiredInputs } from "../io/files";
import { preparePermissionFile, readPreparedPermissionFile, writePreparedPermissionWorkbook } from "../permissions/prepare";

type LogSink = (entry: RunLogEntry) => void;

function nowLog(step: string, message: string, level: RunLogEntry["level"] = "info"): RunLogEntry {
  return { at: new Date().toISOString(), step, level, message };
}

function metricMessage(metric: StepMetrics): string {
  const parts: string[] = [];
  if (typeof metric.filled === "number") parts.push(`${metric.filled} filled`);
  if (typeof metric.appended === "number") parts.push(`${metric.appended} appended`);
  if (typeof metric.adjusted === "number") parts.push(`${metric.adjusted} adjusted`);
  if (typeof metric.blank === "number") parts.push(`${metric.blank} blank`);
  if (typeof metric.skippedCode === "number") parts.push(`${metric.skippedCode} skipped by employee code`);
  if (typeof metric.skippedDate === "number") parts.push(`${metric.skippedDate} skipped by date`);
  if (metric.warnings && metric.warnings.length > 0) parts.push(`${metric.warnings.length} warnings`);
  return `${metric.step}: ${parts.length > 0 ? parts.join(", ") : "completed"}`;
}

function isWeekend(date: ISODate): boolean {
  const day = parseISODate(date).getDay();
  return day === 5 || day === 6;
}

function ensureModelCell(model: DetailedWorkbookModel, employeeCode: number, date: ISODate) {
  model.cells[employeeCode] ??= {};
  model.cells[employeeCode][date] ??= {};
  return model.cells[employeeCode][date];
}

function makeModel(periodStart: ISODate, reports: ParsedReports): DetailedWorkbookModel {
  const period = computePeriod(periodStart);
  const employees = reports.roster;
  const cells: DetailedWorkbookModel["cells"] = {};
  for (const employee of employees) {
    cells[employee.code] = {};
    for (const date of period.dates) cells[employee.code][date] = {};
  }
  return { period, employees, cells };
}

function formatDateHeader(date: ISODate): Date {
  return parseISODate(date);
}

function toArrayBuffer(buffer: unknown): ArrayBuffer {
  if (buffer instanceof ArrayBuffer) return buffer;
  const view = buffer as Uint8Array;
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  return copy.buffer;
}

async function parseReports(inputs: PayrollInputFiles, permissionOptions: PermissionPrepOptions, log: LogSink) {
  const missing = missingRequiredInputs(inputs);
  if (missing.length > 0) throw new Error(`Missing required input(s): ${missing.join(", ")}`);
  const attendanceWorkbook = await readWorkbook(inputs.attendance as File);
  const attendance = parseAttendance(attendanceWorkbook);
  if (attendance.length === 0) throw new Error("Attendance Report contains no parseable attendance rows.");

  const absences = parseAbsences(await readWorkbook(inputs.absences as File));
  const vacations = parseVacations(await readWorkbook(inputs.vacations as File));
  const resignations = inputs.resignations ? parseResignations(await readWorkbook(inputs.resignations)) : [];
  const publicHolidays = inputs.publicHoliday ? parsePublicHolidays(await readWorkbook(inputs.publicHoliday)) : [];
  const templateRoster = inputs.nagwaTemplate ? parseRosterFromTemplate(await readWorkbook(inputs.nagwaTemplate)) : [];
  const roster = buildRosterFromReports(attendance, templateRoster);
  let preparedPermissionsWorkbook: ArrayBuffer | undefined;
  let permissions;

  if (inputs.rawPermissions) {
    const prepared = await preparePermissionFile(inputs.rawPermissions, permissionOptions);
    permissions = prepared.rows;
    preparedPermissionsWorkbook = writePreparedPermissionWorkbook(prepared.rows, "xls");
    log(nowLog("prepare_permission_report", `${prepared.loaded} raw permission rows loaded; ${prepared.kept} approved in-period rows kept.`));
    for (const warning of prepared.warnings) log(nowLog("prepare_permission_report", warning, "warn"));
  } else {
    permissions = await readPreparedPermissionFile(inputs.preparedPermissions as File);
  }

  return {
    reports: {
      attendance,
      absences,
      vacations,
      permissions,
      resignations,
      publicHolidays,
      roster,
    } satisfies ParsedReports,
    attendanceWorkbook,
    preparedPermissionsWorkbook,
  };
}

function fillAttendance(model: DetailedWorkbookModel, reports: ParsedReports, settings: PayrollSettings): StepMetrics {
  let filled = 0;
  let skippedCode = 0;
  let skippedDate = 0;
  const employeeCodes = new Set(model.employees.map((employee) => employee.code));
  const employees = new Map(model.employees.map((employee) => [employee.code, employee]));
  for (const row of reports.attendance) {
    if (!employeeCodes.has(row.code)) {
      skippedCode += 1;
      continue;
    }
    if (!model.period.dates.includes(row.attendanceDay) || isWeekend(row.attendanceDay)) {
      skippedDate += 1;
      continue;
    }
    const employee = employees.get(row.code);
    if (!employee) continue;
    const cells = ensureModelCell(model, row.code, row.attendanceDay);
    cells.in = row.entryTime;
    cells.out = row.exitTime;
    cells.shortage = calculateShortage(row.late, row.entryTime, row.exitTime, row.attendanceDay, row.code, employee.schedule, settings);
    filled += 1;
  }
  return { step: "fill_attendance", filled, skippedCode, skippedDate };
}

function fillAbsences(model: DetailedWorkbookModel, reports: ParsedReports): StepMetrics {
  let filled = 0;
  let skippedCode = 0;
  let skippedDate = 0;
  const employeeCodes = new Set(model.employees.map((employee) => employee.code));
  for (const row of reports.absences) {
    if (!employeeCodes.has(row.code)) {
      skippedCode += 1;
      continue;
    }
    if (!model.period.dates.includes(row.absenceDate) || isWeekend(row.absenceDate)) {
      skippedDate += 1;
      continue;
    }
    const cells = ensureModelCell(model, row.code, row.absenceDate);
    cells.in = "absent";
    cells.out = "absent";
    filled += 1;
  }
  return { step: "fill_absences", filled, skippedCode, skippedDate };
}

function fillVacations(model: DetailedWorkbookModel, reports: ParsedReports): StepMetrics {
  let filled = 0;
  let skippedCode = 0;
  let skippedDate = 0;
  const employeeCodes = new Set(model.employees.map((employee) => employee.code));
  for (const row of reports.vacations) {
    if (!employeeCodes.has(row.code)) {
      skippedCode += 1;
      continue;
    }
    for (const date of buildDateRange(row.startDate, row.endDate)) {
      if (!model.period.dates.includes(date)) {
        skippedDate += 1;
        continue;
      }
      const cells = ensureModelCell(model, row.code, date);
      if (isWeekend(date)) {
        cells.single = row.vacationType;
      } else {
        cells.in = row.vacationType;
        cells.out = row.vacationType;
        cells.shortage = row.vacationType;
      }
      filled += 1;
    }
  }
  return { step: "fill_vacations", filled, skippedCode, skippedDate };
}

function fillResignations(model: DetailedWorkbookModel, reports: ParsedReports): StepMetrics {
  let filled = 0;
  let skippedCode = 0;
  const employeeCodes = new Set(model.employees.map((employee) => employee.code));
  for (const row of reports.resignations) {
    if (!employeeCodes.has(row.code)) {
      skippedCode += 1;
      continue;
    }
    for (const date of model.period.dates) {
      if (date < row.resignationDate || isWeekend(date)) continue;
      const cells = ensureModelCell(model, row.code, date);
      cells.in = "Resigned";
      cells.out = "Resigned";
      filled += 1;
    }
  }
  return { step: "fill_resignations", filled, skippedCode };
}

function fillPublicHolidays(model: DetailedWorkbookModel, reports: ParsedReports): StepMetrics {
  let filled = 0;
  let skippedDate = 0;
  const warnings: string[] = [];
  for (const row of reports.publicHolidays) {
    if (!model.period.dates.includes(row.holidayDate) || isWeekend(row.holidayDate)) {
      skippedDate += 1;
      warnings.push(`Holiday date ${row.holidayDate} not found in Nagwa workday columns.`);
      continue;
    }
    for (const employee of model.employees) {
      const cells = ensureModelCell(model, employee.code, row.holidayDate);
      cells.in = "Public Holiday";
      cells.out = "Public Holiday";
      cells.shortage = "Public Holiday";
      filled += 1;
    }
  }
  return { step: "fill_public_holidays", filled, skippedDate, warnings };
}

function fillPermissions(model: DetailedWorkbookModel, reports: ParsedReports): StepMetrics {
  let filled = 0;
  let appended = 0;
  let skippedCode = 0;
  let skippedDate = 0;
  const employeeCodes = new Set(model.employees.map((employee) => employee.code));
  for (const row of reports.permissions) {
    if (!employeeCodes.has(row.employeeCode)) {
      skippedCode += 1;
      continue;
    }
    if (row.effectiveDate === null || !model.period.dates.includes(row.effectiveDate) || isWeekend(row.effectiveDate)) {
      skippedDate += 1;
      continue;
    }
    const cells = ensureModelCell(model, row.employeeCode, row.effectiveDate);
    const leaveValue = `${row.transactionSubType}, ${row.startTime}, ${row.endTime}`;
    const result = joinLeaveEntry(cells.leave, leaveValue);
    cells.leave = result.value;
    if (result.appended) appended += 1;
    else filled += 1;
  }
  return { step: "fill_permissions", filled, appended, skippedCode, skippedDate };
}

function finalValue(model: DetailedWorkbookModel, employeeCode: number, date: ISODate): string | undefined {
  const cells = model.cells[employeeCode]?.[date];
  if (!cells) return undefined;
  return isWeekend(date) ? cells.single : cells.shortage;
}

function normalizeLabel(text: string): string {
  return text.toLowerCase().replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
}

function abbreviate(value: string | undefined, settings: PayrollSettings): string | undefined {
  if (value === undefined || value.trim() === "") return undefined;
  const lookup = new Map(Object.entries(settings.abbreviations).map(([key, code]) => [normalizeLabel(key), code]));
  return lookup.get(normalizeLabel(value)) ?? value;
}

function formatFinalDuration(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return /^\d{1,3}:[0-5]?\d$/.test(value.trim()) ? `${value.trim()}:00` : value;
}

async function writeDetailedWorkbook(model: DetailedWorkbookModel): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Nagwa Technologies");
  sheet.getCell(1, 1).value = "Employee Code";
  sheet.getCell(1, 2).value = "Employee Name";
  sheet.getCell(1, 5).value = "Schedule";
  for (const [index, employee] of model.employees.entries()) {
    const row = 4 + index;
    sheet.getCell(row, 1).value = employee.code;
    sheet.getCell(row, 2).value = employee.name;
    sheet.getCell(row, 5).value = employee.schedule;
  }
  let col = 10;
  let counter = 1;
  for (const date of model.period.dates) {
    if (isWeekend(date)) {
      sheet.getCell(1, col).value = counter;
      sheet.mergeCells(2, col, 3, col);
      sheet.getCell(2, col).value = formatDateHeader(date);
      for (const [index, employee] of model.employees.entries()) {
        sheet.getCell(4 + index, col).value = ensureModelCell(model, employee.code, date).single ?? "";
      }
      col += 1;
      counter += 1;
    } else {
      sheet.getCell(1, col).value = counter;
      sheet.getCell(1, col + 1).value = counter + 1;
      sheet.getCell(1, col + 2).value = counter + 2;
      sheet.getCell(1, col + 3).value = counter + 3;
      sheet.mergeCells(2, col, 2, col + 3);
      sheet.getCell(2, col).value = formatDateHeader(date);
      ["in", "out", "Leave", "Shortage"].forEach((label, offset) => {
        sheet.getCell(3, col + offset).value = label;
      });
      for (const [index, employee] of model.employees.entries()) {
        const cells = ensureModelCell(model, employee.code, date);
        const row = 4 + index;
        sheet.getCell(row, col).value = cells.in ?? "";
        sheet.getCell(row, col + 1).value = cells.out ?? "";
        sheet.getCell(row, col + 2).value = cells.leave ?? "";
        sheet.getCell(row, col + 3).value = cells.shortage ?? "";
      }
      col += 4;
      counter += 4;
    }
  }
  sheet.views = [{ state: "frozen", ySplit: 3, xSplit: 5 }];
  return toArrayBuffer(await workbook.xlsx.writeBuffer());
}

async function writeFinalWorkbook(model: DetailedWorkbookModel, settings: PayrollSettings): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Final Nagwa Technologies");
  sheet.getCell(3, 1).value = "Employee Code";
  sheet.getCell(3, 2).value = "Employee Name";
  for (const [dateIndex, date] of model.period.dates.entries()) {
    sheet.getCell(3, 9 + dateIndex).value = formatDateHeader(date);
    sheet.getCell(3, 9 + dateIndex).numFmt = "dd-mmm";
  }
  sheet.getCell(3, 40).value = "Total";
  for (const [employeeIndex, employee] of model.employees.entries()) {
    const row = 5 + employeeIndex;
    sheet.getCell(row, 1).value = employee.code;
    sheet.getCell(row, 2).value = employee.name;
    let total = 0;
    for (const [dateIndex, date] of model.period.dates.entries()) {
      const value = formatFinalDuration(abbreviate(finalValue(model, employee.code, date), settings));
      sheet.getCell(row, 9 + dateIndex).value = value ?? "";
      total += parseDuration(value);
    }
    sheet.getCell(row, 40).value = minutesToHMS(total);
  }
  sheet.views = [{ state: "frozen", ySplit: 3, xSplit: 2 }];
  return toArrayBuffer(await workbook.xlsx.writeBuffer());
}

export async function runPayrollPipeline(
  inputs: PayrollInputFiles,
  settings: PayrollSettings,
  permissionOptions: PermissionPrepOptions,
  onLog?: LogSink,
): Promise<PayrollRunResult> {
  const logs: RunLogEntry[] = [];
  const metrics: StepMetrics[] = [];
  const log: LogSink = (entry) => {
    logs.push(entry);
    onLog?.(entry);
  };

  log(nowLog("run", "Starting browser payroll pipeline."));
  log(nowLog("extend_nagwa_technologies", "Detecting first attendance date."));
  const { reports, attendanceWorkbook, preparedPermissionsWorkbook } = await parseReports(inputs, permissionOptions, log);
  const firstDate = detectAttendancePeriodDate(attendanceWorkbook) ?? reports.attendance[0].attendanceDay;
  const model = makeModel(firstDate as ISODate, reports);
  const workdayCount = model.period.dates.filter((date) => !isWeekend(date)).length;
  const weekendCount = model.period.dates.length - workdayCount;
  metrics.push({ step: "extend_nagwa_technologies", filled: model.period.dates.length, workdayCount, weekendCount });
  log(nowLog("extend_nagwa_technologies", `Replaced calendar with period ${model.period.start} -> ${model.period.end}.`));

  log(nowLog("fill_attendance", "Applying attendance and HR business rules."));
  const fillSteps = [
    fillAttendance(model, reports, settings),
    fillAbsences(model, reports),
    fillVacations(model, reports),
    fillResignations(model, reports),
    fillPublicHolidays(model, reports),
    fillPermissions(model, reports),
    recalculateShortageFromLeave(model, settings),
    applyPermittedDelays(model),
    applyWorkMissionLunchExemption(model, settings),
    applyHourReduction(model, settings),
    fillMissingPunches(model),
    applyWfhAndWorkdayOverrides(model),
  ];
  metrics.push(...fillSteps);
  for (const metric of fillSteps) {
    log(nowLog("fill_attendance", metricMessage(metric)));
    for (const warning of metric.warnings ?? []) log(nowLog("fill_attendance", warning, "warn"));
  }

  log(nowLog("extend_final_nagwa_technologies", `Wrote ${model.period.dates.length} final report dates.`));
  metrics.push({ step: "extend_final_nagwa_technologies", filled: model.period.dates.length });
  const detailedWorkbook = await writeDetailedWorkbook(model);
  const finalWorkbook = await writeFinalWorkbook(model, settings);
  metrics.push({ step: "complete_final", filled: model.employees.length * model.period.dates.length, totalsWritten: model.employees.length });
  log(nowLog("complete_final", "Copied shortage values, applied abbreviations, and wrote totals."));
  log(nowLog("run", "All scripts completed successfully.", "success"));

  const warnings = metrics.flatMap((metric) => metric.warnings ?? []);
  return {
    detailedWorkbook,
    finalWorkbook,
    preparedPermissionsWorkbook,
    metrics,
    logs,
    period: model.period,
    employeesProcessed: model.employees.length,
    warnings,
  };
}
