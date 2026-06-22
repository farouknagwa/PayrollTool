export type ISODate = `${number}-${number}-${number}`;

export type Minutes = number;

export interface PayrollPeriod {
  start: ISODate;
  end: ISODate;
  dates: ISODate[];
}

export interface TimeWindow {
  start: string;
  end: string;
}

export interface HourReductionWindow {
  employeeCode: number;
  startDate: ISODate | null;
  endDate: ISODate | null;
}

export interface SpecialRulePair {
  employeeA: number;
  employeeB: number;
  anchorDate: ISODate;
}

export interface PayrollSettings {
  ramadanStart: ISODate;
  ramadanEnd: ISODate;
  workdayStart: string;
  workdayEndNormal: string;
  workdayEndRamadan: string;
  fullDayNormalMinutes: Minutes;
  fullDayRamadanMinutes: Minutes;
  permittedWindowEndNormal: string;
  permittedWindowEndRamadan: string;
  permittedWindowEndRestricted: string;
  scheduleWindowEnd: Record<string, string>;
  defaultScheduleWindowEnd: string;
  specialRulePairs: SpecialRulePair[];
  lunchWindowSwitchDate: ISODate;
  lunchWindowBefore: TimeWindow;
  lunchWindowFrom: TimeWindow;
  hourReductionWindows: HourReductionWindow[];
  abbreviations: Record<string, string>;
  requestCutoffDaysDefault: number;
  debugMode: boolean;
}

export interface RosterEmployee {
  code: number;
  name: string;
  schedule: string;
}

export interface AttendanceRecord {
  code: number;
  name: string;
  attendanceDay: ISODate;
  entryTime: string;
  exitTime: string;
  late: string;
}

export interface AbsenceRecord {
  code: number;
  absenceDate: ISODate;
}

export interface VacationRecord {
  code: number;
  startDate: ISODate;
  endDate: ISODate;
  vacationType: string;
}

export interface ResignationRecord {
  code: number;
  resignationDate: ISODate;
}

export interface PublicHolidayRecord {
  holidayDate: ISODate;
}

export interface PermissionRecord {
  employeeCode: number;
  employeeName: string;
  requestDate: ISODate | null;
  effectiveDate: ISODate | null;
  startTime: string;
  endTime: string;
  totalPermissionPeriod: string;
  time: string;
  transactionType: string;
  transactionSubType: string;
  wfTemplate: string;
  status: string;
}

export interface RawPermissionRecord extends Omit<PermissionRecord, "totalPermissionPeriod"> {
  branchCode?: string;
  workflowId?: string;
}

export interface DayCells {
  in?: string;
  out?: string;
  leave?: string;
  shortage?: string;
  single?: string;
}

export interface DetailedWorkbookModel {
  period: PayrollPeriod;
  employees: RosterEmployee[];
  cells: Record<number, Record<ISODate, DayCells>>;
}

export type LogLevel = "info" | "warn" | "error" | "success";

export interface RunLogEntry {
  at: string;
  step: string;
  level: LogLevel;
  message: string;
}

export interface StepMetrics {
  step: string;
  filled?: number;
  appended?: number;
  adjusted?: number;
  blank?: number;
  skippedCode?: number;
  skippedDate?: number;
  warnings?: string[];
  [key: string]: unknown;
}

export interface PayrollInputFiles {
  attendance?: File;
  absences?: File;
  vacations?: File;
  preparedPermissions?: File;
  rawPermissions?: File;
  resignations?: File;
  publicHoliday?: File;
  alternatePermissions?: File;
  nagwaTemplate?: File;
  finalTemplate?: File;
}

export interface PermissionPrepOptions {
  month: number;
  year: number;
  requestCutoffDays: number;
  noRequestCutoff: boolean;
}

export interface PermissionPrepResult {
  rows: PermissionRecord[];
  period: PayrollPeriod;
  loaded: number;
  kept: number;
  excludedByCutoff: number;
  warnings: string[];
}

export interface PayrollRunResult {
  detailedWorkbook: ArrayBuffer;
  finalWorkbook: ArrayBuffer;
  preparedPermissionsWorkbook?: ArrayBuffer;
  metrics: StepMetrics[];
  logs: RunLogEntry[];
  period: PayrollPeriod;
  employeesProcessed: number;
  warnings: string[];
}

export interface ParsedReports {
  attendance: AttendanceRecord[];
  absences: AbsenceRecord[];
  vacations: VacationRecord[];
  permissions: PermissionRecord[];
  resignations: ResignationRecord[];
  publicHolidays: PublicHolidayRecord[];
  roster: RosterEmployee[];
}
