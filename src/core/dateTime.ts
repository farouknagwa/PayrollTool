import type { ISODate, PayrollPeriod } from "./types";

const DAY_MS = 24 * 60 * 60 * 1000;

export function toISODate(date: Date): ISODate {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}` as ISODate;
}

export function parseISODate(value: ISODate | string): Date {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

export function compareISO(a: ISODate, b: ISODate): number {
  return parseISODate(a).getTime() - parseISODate(b).getTime();
}

export function isBetweenISO(value: ISODate, start: ISODate | null, end: ISODate | null): boolean {
  if (start !== null && compareISO(value, start) < 0) return false;
  if (end !== null && compareISO(value, end) > 0) return false;
  return true;
}

export function addDays(date: ISODate, days: number): ISODate {
  const out = parseISODate(date);
  out.setDate(out.getDate() + days);
  return toISODate(out);
}

export function buildDateRange(start: ISODate, end: ISODate): ISODate[] {
  const out: ISODate[] = [];
  let cursor = parseISODate(start);
  const endTime = parseISODate(end).getTime();
  while (cursor.getTime() <= endTime) {
    out.push(toISODate(cursor));
    cursor = new Date(cursor.getTime() + DAY_MS);
  }
  return out;
}

export function computePeriod(refDate: ISODate): PayrollPeriod {
  const ref = parseISODate(refDate);
  let start: Date;
  if (ref.getDate() >= 21) {
    start = new Date(ref.getFullYear(), ref.getMonth(), 21);
  } else {
    start = new Date(ref.getFullYear(), ref.getMonth() - 1, 21);
  }
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 20);
  const startIso = toISODate(start);
  const endIso = toISODate(end);
  return {
    start: startIso,
    end: endIso,
    dates: buildDateRange(startIso, endIso),
  };
}

export function periodFromMonth(month: number, year: number): PayrollPeriod {
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error(`Month must be between 1 and 12, got ${month}.`);
  }
  const start = new Date(year, month - 2, 21);
  const end = new Date(year, month - 1, 20);
  const startIso = toISODate(start);
  const endIso = toISODate(end);
  return {
    start: startIso,
    end: endIso,
    dates: buildDateRange(startIso, endIso),
  };
}

export function isEgyptWeekend(date: ISODate): boolean {
  const jsDay = parseISODate(date).getDay();
  return jsDay === 5 || jsDay === 6;
}

export function parseDateValue(value: unknown): ISODate | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toISODate(value);
  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return toISODate(new Date(excelEpoch + value * DAY_MS));
  }
  const text = String(value).trim();
  if (!text || text.toLowerCase() === "nan") return null;

  const slash = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slash) {
    const [, day, month, year] = slash;
    return toISODate(new Date(Number(year), Number(month) - 1, Number(day)));
  }

  const dash = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (dash) {
    const [, year, month, day] = dash;
    return toISODate(new Date(Number(year), Number(month) - 1, Number(day)));
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : toISODate(parsed);
}

export function formatDDMMYYYY(date: ISODate | null): string {
  if (date === null) return "";
  const parsed = parseISODate(date);
  return `${`${parsed.getDate()}`.padStart(2, "0")}/${`${parsed.getMonth() + 1}`.padStart(2, "0")}/${parsed.getFullYear()}`;
}

export function timeToMinutes(value: string): number | null {
  const text = String(value ?? "").trim();
  if (!text || text === "nan" || text === "None") return null;
  const normalized = text.replace(/\s+/g, " ").toUpperCase();
  const twelve = normalized.match(/^(\d{1,2}):([0-5]\d)\s*([AP]M)$/);
  if (twelve) {
    let hour = Number(twelve[1]);
    const minute = Number(twelve[2]);
    const marker = twelve[3];
    if (marker === "PM" && hour !== 12) hour += 12;
    if (marker === "AM" && hour === 12) hour = 0;
    return hour * 60 + minute;
  }
  const hms = normalized.match(/^(\d{1,2}):([0-5]\d)(?::[0-5]\d)?$/);
  if (hms) return Number(hms[1]) * 60 + Number(hms[2]);
  return null;
}

export function minutesToTime(totalMinutes: number): string {
  const minutes = ((Math.trunc(totalMinutes) % 1440) + 1440) % 1440;
  const hour24 = Math.floor(minutes / 60);
  const minute = minutes % 60;
  const marker = hour24 >= 12 ? "PM" : "AM";
  const hour12 = hour24 % 12 || 12;
  return `${`${hour12}`.padStart(2, "0")}:${`${minute}`.padStart(2, "0")} ${marker}`;
}

export function parseDuration(value: unknown): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.round(value * 24 * 60);
  }
  const text = String(value ?? "").trim();
  if (!text || text === "nan" || text === "None") return 0;
  const parts = text.match(/^(\d{1,3}):([0-5]?\d)(?::([0-5]?\d))?$/);
  if (!parts) return 0;
  const seconds = parts[3] ? Number(parts[3]) : 0;
  return Number(parts[1]) * 60 + Number(parts[2]) + (seconds >= 30 ? 1 : 0);
}

export function minutesToHHMM(totalMinutes: number): string {
  const safe = Math.max(0, Math.trunc(totalMinutes));
  const hours = Math.floor(safe / 60);
  const minutes = safe % 60;
  return `${hours}:${`${minutes}`.padStart(2, "0")}`;
}

export function minutesToHMS(totalMinutes: number): string {
  return `${minutesToHHMM(totalMinutes)}:00`;
}

export function formatTimeValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return minutesToTime(value.getHours() * 60 + value.getMinutes());
  if (typeof value === "number" && Number.isFinite(value)) return minutesToTime(Math.round(value * 24 * 60));
  const minutes = timeToMinutes(String(value));
  return minutes === null ? String(value).trim() : minutesToTime(minutes);
}

export function permissionDuration(startValue: unknown, endValue: unknown): number {
  const start = timeToMinutes(formatTimeValue(startValue));
  const end = timeToMinutes(formatTimeValue(endValue));
  if (start === null || end === null) return 0;
  return end >= start ? end - start : end + 24 * 60 - start;
}

export function overlapMinutes(startA: number, endA: number, startB: number, endB: number): number {
  return Math.max(0, Math.min(endA, endB) - Math.max(startA, startB));
}

export function mergeCoveredMinutes(intervals: Array<[number, number]>): number {
  const clipped = intervals.filter(([, end]) => Number.isFinite(end)).filter(([start, end]) => end > start);
  clipped.sort((a, b) => a[0] - b[0]);
  let covered = 0;
  let current: [number, number] | null = null;
  for (const [start, end] of clipped) {
    if (current === null || start > current[1]) {
      if (current !== null) covered += current[1] - current[0];
      current = [start, end];
    } else {
      current[1] = Math.max(current[1], end);
    }
  }
  if (current !== null) covered += current[1] - current[0];
  return covered;
}
