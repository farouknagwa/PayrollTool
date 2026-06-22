import { DEFAULT_SETTINGS } from "./defaults";
import { compareISO } from "../core/dateTime";
import type { PayrollSettings } from "../core/types";

const STORAGE_KEY = "payrolltool.settings.v1";

function isTime(value: string): boolean {
  return /^\d{1,2}:[0-5]\d$/.test(value);
}

export function cloneDefaultSettings(): PayrollSettings {
  return JSON.parse(JSON.stringify(DEFAULT_SETTINGS)) as PayrollSettings;
}

export function normalizeSettings(input: Partial<PayrollSettings>): PayrollSettings {
  const defaults = cloneDefaultSettings();
  const merged = {
    ...defaults,
    ...input,
    scheduleWindowEnd: {
      ...defaults.scheduleWindowEnd,
      ...(input.scheduleWindowEnd ?? {}),
    },
    abbreviations: {
      ...defaults.abbreviations,
      ...(input.abbreviations ?? {}),
    },
    specialRulePairs: input.specialRulePairs ?? defaults.specialRulePairs,
    hourReductionWindows: input.hourReductionWindows ?? defaults.hourReductionWindows,
    lunchWindowBefore: {
      ...defaults.lunchWindowBefore,
      ...(input.lunchWindowBefore ?? {}),
    },
    lunchWindowFrom: {
      ...defaults.lunchWindowFrom,
      ...(input.lunchWindowFrom ?? {}),
    },
  };
  return merged;
}

export function validateSettings(settings: PayrollSettings): string[] {
  const errors: string[] = [];
  if (compareISO(settings.ramadanStart, settings.ramadanEnd) > 0) {
    errors.push("Ramadan start date must be on or before Ramadan end date.");
  }
  const timeFields = [
    ["Workday start", settings.workdayStart],
    ["Normal workday end", settings.workdayEndNormal],
    ["Ramadan workday end", settings.workdayEndRamadan],
    ["Normal permitted window end", settings.permittedWindowEndNormal],
    ["Ramadan permitted window end", settings.permittedWindowEndRamadan],
    ["Restricted permitted window end", settings.permittedWindowEndRestricted],
    ["Default schedule window end", settings.defaultScheduleWindowEnd],
    ["Lunch before start", settings.lunchWindowBefore.start],
    ["Lunch before end", settings.lunchWindowBefore.end],
    ["Lunch from start", settings.lunchWindowFrom.start],
    ["Lunch from end", settings.lunchWindowFrom.end],
  ];
  for (const [label, value] of timeFields) {
    if (!isTime(value)) errors.push(`${label} must be HH:MM.`);
  }
  for (const [name, value] of Object.entries(settings.scheduleWindowEnd)) {
    if (!name.trim()) errors.push("Schedule type names cannot be blank.");
    if (!isTime(value)) errors.push(`Schedule '${name}' must use HH:MM window end.`);
  }
  for (const pair of settings.specialRulePairs) {
    if (!Number.isInteger(pair.employeeA) || !Number.isInteger(pair.employeeB)) {
      errors.push("Special-rule employee codes must be integers.");
    }
    if (pair.employeeA === pair.employeeB) {
      errors.push("Special-rule pair employees must be different.");
    }
  }
  for (const window of settings.hourReductionWindows) {
    if (!Number.isInteger(window.employeeCode)) {
      errors.push("Hour-reduction employee code must be an integer.");
    }
    if (window.startDate && window.endDate && compareISO(window.startDate, window.endDate) > 0) {
      errors.push(`Hour-reduction range for ${window.employeeCode} has start after end.`);
    }
  }
  if (settings.requestCutoffDaysDefault < 0) {
    errors.push("Request cutoff days cannot be negative.");
  }
  return errors;
}

export function loadSettings(): PayrollSettings {
  if (typeof localStorage === "undefined") return cloneDefaultSettings();
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return cloneDefaultSettings();
  try {
    return normalizeSettings(JSON.parse(raw) as Partial<PayrollSettings>);
  } catch {
    return cloneDefaultSettings();
  }
}

export function saveSettings(settings: PayrollSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function resetSettings(): PayrollSettings {
  const defaults = cloneDefaultSettings();
  if (typeof localStorage !== "undefined") {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(defaults));
  }
  return defaults;
}
