import {
  compareISO,
  isBetweenISO,
  mergeCoveredMinutes,
  minutesToHHMM,
  overlapMinutes,
  parseDuration,
  timeToMinutes,
} from "./dateTime";
import type {
  DayCells,
  DetailedWorkbookModel,
  ISODate,
  PayrollSettings,
  StepMetrics,
} from "./types";

const LEAVE_ENTRY_SEPARATOR = " | ";
const HALFDAY_CATEGORIES = [
  "1st_half_annual",
  "2nd_half_annual",
  "1st_half_ramadan",
  "2nd_half_ramadan",
] as const;

type HalfDayCategory = (typeof HALFDAY_CATEGORIES)[number];
type RuleCategory = HalfDayCategory | "permission";

interface ParsedLeaveEntry {
  leaveType: string;
  start: number;
  end: number;
  category: RuleCategory | null;
}

export function splitLeaveEntries(value: unknown): string[] {
  if (value === null || value === undefined) return [];
  const raw = String(value).trim();
  if (!raw || raw === "nan" || raw === "None") return [];
  return raw.split(LEAVE_ENTRY_SEPARATOR).map((part) => part.trim()).filter(Boolean);
}

export function joinLeaveEntry(existing: unknown, next: string): { value: string; appended: boolean } {
  const current = existing === null || existing === undefined ? "" : String(existing).trim();
  if (!current || current === "nan" || current === "None") return { value: next, appended: false };
  return { value: `${current}${LEAVE_ENTRY_SEPARATOR}${next}`, appended: true };
}

export function parseLeaveCell(value: string): ParsedLeaveEntry {
  const parts = value.split(",").map((part) => part.trim());
  if (parts.length < 3) {
    throw new Error(`Expected at least 3 comma-separated parts, got ${parts.length}: '${value}'`);
  }
  const endRaw = parts.at(-1) ?? "";
  const startRaw = parts.at(-2) ?? "";
  const leaveType = parts.slice(0, -2).join(", ");
  const start = timeToMinutes(startRaw);
  const end = timeToMinutes(endRaw);
  if (start === null) throw new Error(`Cannot parse leave start time: '${startRaw}'`);
  if (end === null) throw new Error(`Cannot parse leave end time: '${endRaw}'`);
  return {
    leaveType,
    start,
    end,
    category: classifyLeaveType(leaveType, start, end),
  };
}

function isHalfDay(category: RuleCategory | null): category is HalfDayCategory {
  return category !== null && HALFDAY_CATEGORIES.includes(category as HalfDayCategory);
}

export function classifyLeaveType(leaveType: string, leaveStart?: number, leaveEnd?: number): RuleCategory | null {
  const normalized = leaveType.trim().toLowerCase().replace(/\s+/g, " ");
  if (normalized.includes("1st half day") && normalized.includes("annual")) return "1st_half_annual";
  if (normalized.includes("2nd half day") && normalized.includes("annual")) return "2nd_half_annual";
  if (normalized.includes("1st half day") && normalized.includes("ramadan")) return "1st_half_ramadan";
  if (normalized.includes("2nd half day") && normalized.includes("ramadan")) return "2nd_half_ramadan";
  if (normalized.includes("half day") && normalized.includes("annual")) {
    if (leaveStart === undefined || leaveEnd === undefined) return null;
    return leaveStart < 12 * 60 ? "1st_half_annual" : "2nd_half_annual";
  }
  if (normalized.includes("half day") && normalized.includes("ramadan")) {
    if (leaveStart === undefined || leaveEnd === undefined) return null;
    return leaveStart < 12 * 60 ? "1st_half_ramadan" : "2nd_half_ramadan";
  }
  if (normalized === "permission" || normalized === "work mission") return "permission";
  if (normalized.includes("educational leave") && normalized.includes("core of business")) return "permission";
  return null;
}

export function specialRulePeriodIndex(date: ISODate, anchorDate: ISODate): number {
  if (compareISO(date, anchorDate) < 0) return -1;
  const parsed = new Date(`${date}T00:00:00`);
  const periodYear = parsed.getDate() >= 21 ? parsed.getFullYear() : parsed.getMonth() === 0 ? parsed.getFullYear() - 1 : parsed.getFullYear();
  const periodMonth = parsed.getDate() >= 21 ? parsed.getMonth() + 1 : parsed.getMonth() === 0 ? 12 : parsed.getMonth();
  const anchor = new Date(`${anchorDate}T00:00:00`);
  return (periodYear - anchor.getFullYear()) * 12 + (periodMonth - (anchor.getMonth() + 1));
}

export function hasRestrictedWindow(empCode: number, date: ISODate, settings: PayrollSettings): boolean {
  for (const pair of settings.specialRulePairs) {
    if (empCode !== pair.employeeA && empCode !== pair.employeeB) continue;
    const periodIndex = specialRulePeriodIndex(date, pair.anchorDate);
    if (periodIndex < 0) continue;
    const restrictedEmployee = periodIndex % 2 === 0 ? pair.employeeA : pair.employeeB;
    if (empCode === restrictedEmployee) return true;
  }
  return false;
}

export function permittedWindowEnd(date: ISODate, empCode: number, schedule: string, settings: PayrollSettings): string {
  if (isBetweenISO(date, settings.ramadanStart, settings.ramadanEnd)) return settings.permittedWindowEndRamadan;
  if (hasRestrictedWindow(empCode, date, settings)) return settings.permittedWindowEndRestricted;
  const isSpecial = settings.specialRulePairs.some((pair) => pair.employeeA === empCode || pair.employeeB === empCode);
  if (isSpecial) return settings.permittedWindowEndNormal;
  const key = (schedule || "").trim().toLowerCase();
  return settings.scheduleWindowEnd[key] ?? settings.defaultScheduleWindowEnd;
}

export function calculateShortage(
  lateValue: string,
  entryTimeValue: string,
  exitTimeValue: string,
  date: ISODate,
  empCode: number,
  schedule: string,
  settings: PayrollSettings,
): string {
  const lateMinutes = parseDuration(lateValue);
  const fullDay = isBetweenISO(date, settings.ramadanStart, settings.ramadanEnd)
    ? settings.fullDayRamadanMinutes
    : settings.fullDayNormalMinutes;
  const start = timeToMinutes(settings.workdayStart) ?? 8 * 60;
  const end = timeToMinutes(permittedWindowEnd(date, empCode, schedule, settings)) ?? 16 * 60;
  const entry = timeToMinutes(entryTimeValue);
  const exit = timeToMinutes(exitTimeValue);
  const effectiveMinutes = entry !== null && exit !== null ? Math.max(0, Math.min(exit, end) - Math.max(entry, start)) : 0;
  return effectiveMinutes >= fullDay ? minutesToHHMM(lateMinutes) : minutesToHHMM(fullDay - effectiveMinutes);
}

function getCells(model: DetailedWorkbookModel, empCode: number, date: ISODate): DayCells {
  const row = model.cells[empCode] ?? {};
  const cells = row[date] ?? {};
  model.cells[empCode] = row;
  row[date] = cells;
  return cells;
}

function parsePunch(value: unknown): number | null {
  return timeToMinutes(String(value ?? "").trim());
}

function applyHalfDayRule(
  cells: DayCells,
  parsed: ParsedLeaveEntry,
  date: ISODate,
  category: HalfDayCategory,
  otherLeaves: ParsedLeaveEntry[],
  settings: PayrollSettings,
): boolean {
  const inTime = parsePunch(cells.in);
  const outTime = parsePunch(cells.out);
  if (inTime === null || outTime === null) {
    if (inTime === null) cells.in = "Missing Punch";
    if (outTime === null) cells.out = "Missing Punch";
    cells.shortage = "Missing Punch";
    return true;
  }

  const workdayStart = timeToMinutes(settings.workdayStart) ?? 8 * 60;
  const normalEnd = timeToMinutes(settings.workdayEndNormal) ?? 16 * 60;
  const ramadanEnd = timeToMinutes(settings.workdayEndRamadan) ?? 14 * 60 + 30;
  let windowStart: number;
  let windowEnd: number;
  if (category === "1st_half_annual") {
    windowStart = parsed.end;
    windowEnd = normalEnd;
  } else if (category === "2nd_half_annual") {
    windowStart = workdayStart;
    windowEnd = parsed.start;
  } else if (category === "1st_half_ramadan") {
    windowStart = parsed.end;
    windowEnd = ramadanEnd;
  } else {
    windowStart = workdayStart;
    windowEnd = parsed.start;
  }
  if (windowEnd <= windowStart) return false;
  const intervals: Array<[number, number]> = [[inTime, outTime]];
  for (const leave of otherLeaves) intervals.push([leave.start, leave.end]);
  const covered = mergeCoveredMinutes(intervals.map(([start, end]) => [Math.max(start, windowStart), Math.min(end, windowEnd)]));
  cells.shortage = minutesToHHMM(Math.max(0, windowEnd - windowStart - covered));
  void date;
  return true;
}

function applyPermissionRule(
  cells: DayCells,
  parsed: ParsedLeaveEntry,
  date: ISODate,
  empCode: number,
  schedule: string,
  settings: PayrollSettings,
): boolean {
  const inTime = parsePunch(cells.in);
  const outTime = parsePunch(cells.out);
  const fullDay = isBetweenISO(date, settings.ramadanStart, settings.ramadanEnd)
    ? settings.fullDayRamadanMinutes
    : settings.fullDayNormalMinutes;
  const windowStart = timeToMinutes(settings.workdayStart) ?? 8 * 60;
  const windowEnd = timeToMinutes(permittedWindowEnd(date, empCode, schedule, settings)) ?? 16 * 60;

  if (inTime === null) {
    cells.in = "Missing Punch";
    if (outTime === null) cells.out = "Missing Punch";
    cells.shortage = "Missing Punch";
    return true;
  }
  if (parsed.start > inTime) {
    if (outTime === null) {
      cells.out = "Missing Punch";
      cells.shortage = "Missing Punch";
      return true;
    }
    if (parsed.end <= outTime) return false;
    const leaveDuration = parsed.end - parsed.start;
    const effectiveIn = Math.max(inTime, windowStart);
    const effectiveOut = Math.min(outTime, windowEnd);
    const overlap = overlapMinutes(effectiveIn, effectiveOut, parsed.start, parsed.end);
    const actualWork = Math.max(0, effectiveOut - effectiveIn) - overlap;
    const hoursShortage = Math.max(0, fullDay - leaveDuration - actualWork);
    const latePenalty = Math.max(0, inTime - 10 * 60);
    cells.shortage = minutesToHHMM(Math.max(latePenalty, hoursShortage));
    return true;
  }

  if (outTime === null) {
    cells.out = "Missing Punch";
    cells.shortage = "Missing Punch";
    return true;
  }
  const cutoff = Math.max(10 * 60, parsed.end);
  const latePenalty = Math.max(0, inTime - cutoff);
  const requiredPresence = fullDay - (parsed.end - parsed.start);
  const effectiveIn = Math.max(inTime, windowStart);
  const effectiveOut = Math.min(outTime, windowEnd);
  const overlap = overlapMinutes(effectiveIn, effectiveOut, parsed.start, parsed.end);
  const actualPresence = Math.max(0, effectiveOut - effectiveIn) - overlap;
  cells.shortage = minutesToHHMM(Math.max(latePenalty, Math.max(0, requiredPresence - actualPresence)));
  return true;
}

function subtractLeaveFromShortage(cells: DayCells, parsed: ParsedLeaveEntry): boolean {
  const current = String(cells.shortage ?? "").trim();
  if (!/^\d+:\d{2}$/.test(current)) return false;
  const duration = parsed.end - parsed.start;
  if (duration <= 0) return false;
  cells.shortage = minutesToHHMM(Math.max(0, parseDuration(current) - duration));
  return true;
}

export function recalculateShortageFromLeave(model: DetailedWorkbookModel, settings: PayrollSettings): StepMetrics {
  let overridden = 0;
  const warnings: string[] = [];
  for (const employee of model.employees) {
    for (const date of model.period.dates) {
      const cells = getCells(model, employee.code, date);
      const rawEntries = splitLeaveEntries(cells.leave);
      if (rawEntries.length === 0) continue;
      const parsed = rawEntries.map((entry) => {
        try {
          return parseLeaveCell(entry);
        } catch (error) {
          warnings.push(`Employee ${employee.code}, Date ${date}: ${(error as Error).message}`);
          return null;
        }
      });
      let ruleApplied = false;
      let anyChange = false;
      for (const [index, entry] of parsed.entries()) {
        if (entry === null || entry.category === null) continue;
        if (isHalfDay(entry.category)) {
          const otherLeaves = parsed.filter((other, otherIndex) => otherIndex !== index && other !== null && !isHalfDay(other.category)) as ParsedLeaveEntry[];
          if (applyHalfDayRule(cells, entry, date, entry.category, otherLeaves, settings)) {
            ruleApplied = true;
            anyChange = true;
          }
        } else if (!ruleApplied) {
          if (applyPermissionRule(cells, entry, date, employee.code, employee.schedule, settings)) {
            ruleApplied = true;
            anyChange = true;
          }
        } else if (subtractLeaveFromShortage(cells, entry)) {
          anyChange = true;
        }
      }
      if (anyChange) overridden += 1;
    }
  }
  return { step: "recalculate_shortage_from_leave", adjusted: overridden, warnings };
}

export function applyPermittedDelays(model: DetailedWorkbookModel): StepMetrics {
  let adjusted = 0;
  const warnings: string[] = [];
  for (const employee of model.employees) {
    for (const date of model.period.dates) {
      const cells = getCells(model, employee.code, date);
      const entries = splitLeaveEntries(cells.leave);
      if (entries.length === 0) continue;
      const parsed = entries.flatMap((entry) => {
        try {
          return [parseLeaveCell(entry)];
        } catch {
          return [];
        }
      });
      if (parsed.some((entry) => isHalfDay(entry.category))) continue;
      for (const entry of parsed) {
        if (!entry.leaveType.toLowerCase().includes("permitted delays")) continue;
        const duration = entry.end - entry.start;
        if (duration <= 0) {
          warnings.push(`Employee ${employee.code}, Date ${date}: Permitted-delay duration is zero or negative.`);
          continue;
        }
        const current = String(cells.shortage ?? "").trim();
        if (current === "Missing Punch") continue;
        cells.shortage = minutesToHHMM(Math.max(0, parseDuration(current) - duration));
        adjusted += 1;
      }
    }
  }
  return { step: "apply_permitted_delays", adjusted, warnings };
}

export function applyWorkMissionLunchExemption(model: DetailedWorkbookModel, settings: PayrollSettings): StepMetrics {
  let adjusted = 0;
  for (const employee of model.employees) {
    for (const date of model.period.dates) {
      const cells = getCells(model, employee.code, date);
      const entries = splitLeaveEntries(cells.leave).flatMap((entry) => {
        try {
          return [parseLeaveCell(entry)];
        } catch {
          return [];
        }
      });
      if (!entries.some((entry) => entry.leaveType.toLowerCase().includes("work mission"))) continue;
      const current = String(cells.shortage ?? "").trim();
      if (!/^\d+:\d{2}$/.test(current)) continue;
      const lunch = compareISO(date, settings.lunchWindowSwitchDate) < 0 ? settings.lunchWindowBefore : settings.lunchWindowFrom;
      const lunchStart = timeToMinutes(lunch.start) ?? 12 * 60;
      const lunchEnd = timeToMinutes(lunch.end) ?? 13 * 60;
      const intervals: Array<[number, number]> = entries.map((entry) => [entry.start, entry.end]);
      const inTime = parsePunch(cells.in);
      const outTime = parsePunch(cells.out);
      if (inTime !== null && outTime !== null) intervals.push([inTime, outTime]);
      const covered = mergeCoveredMinutes(intervals.map(([start, end]) => [Math.max(start, lunchStart), Math.min(end, lunchEnd)]));
      const uncovered = lunchEnd - lunchStart - covered;
      if (uncovered <= 0) continue;
      cells.shortage = minutesToHHMM(Math.max(0, parseDuration(current) - uncovered));
      adjusted += 1;
    }
  }
  return { step: "apply_work_mission_lunch_exemption", adjusted };
}

export function applyHourReduction(model: DetailedWorkbookModel, settings: PayrollSettings): StepMetrics {
  let adjusted = 0;
  for (const window of settings.hourReductionWindows) {
    for (const date of model.period.dates) {
      if (!isBetweenISO(date, window.startDate, window.endDate)) continue;
      const cells = getCells(model, window.employeeCode, date);
      const current = String(cells.shortage ?? "").trim();
      if (!/^\d+:\d{2}$/.test(current)) continue;
      cells.shortage = minutesToHHMM(Math.max(0, parseDuration(current) - 60));
      adjusted += 1;
    }
  }
  return { step: "apply_hour_reduction", adjusted };
}

export function fillMissingPunches(model: DetailedWorkbookModel): StepMetrics {
  let filled = 0;
  for (const employee of model.employees) {
    for (const date of model.period.dates) {
      const cells = getCells(model, employee.code, date);
      const inText = String(cells.in ?? "").trim();
      const outText = String(cells.out ?? "").trim();
      const inBlank = !inText || inText === "nan" || inText === "None";
      const outBlank = !outText || outText === "nan" || outText === "None";
      if (inBlank === outBlank) continue;
      if (inBlank) {
        if (parsePunch(outText) === null) continue;
        cells.in = "Missing Punch";
      } else {
        if (parsePunch(inText) === null) continue;
        cells.out = "Missing Punch";
      }
      cells.shortage = "Missing Punch";
      filled += 1;
    }
  }
  return { step: "fill_missing_punches", filled };
}

export function applyWfhAndWorkdayOverrides(model: DetailedWorkbookModel): StepMetrics {
  let filled = 0;
  for (const employee of model.employees) {
    for (const date of model.period.dates) {
      const cells = getCells(model, employee.code, date);
      const leave = String(cells.leave ?? "").toLowerCase();
      const label = leave.includes("workday") ? "Workday" : leave.includes("work from home") ? "WFH" : null;
      if (label === null) continue;
      cells.in = label;
      cells.out = label;
      cells.shortage = label;
      filled += 1;
    }
  }
  return { step: "apply_wfh_and_workday_overrides", filled };
}
