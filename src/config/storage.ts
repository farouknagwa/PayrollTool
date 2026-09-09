import { DEFAULT_SETTINGS } from "./defaults";
import { compareISO } from "../core/dateTime";
import type { PayrollSettings } from "../core/types";

const LEGACY_STORAGE_KEY = "payrolltool.settings.v1";

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
    lunchWindowSwitchDateCancel: !("lunchWindowSwitchDateCancel" in input)
      ? defaults.lunchWindowSwitchDateCancel
      : input.lunchWindowSwitchDateCancel
        ? input.lunchWindowSwitchDateCancel
        : null,
  };
  return merged;
}

export function validateSettings(settings: PayrollSettings): string[] {
  const errors: string[] = [];
  if (compareISO(settings.ramadanStart, settings.ramadanEnd) > 0) {
    errors.push("Ramadan start date must be on or before Ramadan end date.");
  }
  if (
    settings.lunchWindowSwitchDateCancel &&
    compareISO(settings.lunchWindowSwitchDateCancel, settings.lunchWindowSwitchDate) <= 0
  ) {
    errors.push("Lunch switch cancel date must be after lunch switch date.");
  }
  const timeFields = [
    ["Workday start", settings.workdayStart],
    ["Normal workday end", settings.workdayEndNormal],
    ["Ramadan workday end", settings.workdayEndRamadan],
    ["Normal permitted window end", settings.permittedWindowEndNormal],
    ["Ramadan permitted window end", settings.permittedWindowEndRamadan],
    ["Restricted permitted window end", settings.permittedWindowEndRestricted],
    ["Default schedule window end", settings.defaultScheduleWindowEnd],
    ["Lunch start", settings.lunchWindowBefore.start],
    ["Lunch end", settings.lunchWindowBefore.end],
    ["Summer Lunch start", settings.lunchWindowFrom.start],
    ["Summer Lunch end", settings.lunchWindowFrom.end],
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

function clearLegacyBrowserSettings(): void {
  if (typeof localStorage === "undefined") return;
  localStorage.removeItem(LEGACY_STORAGE_KEY);
}

export function loadSettings(): PayrollSettings {
  clearLegacyBrowserSettings();
  return cloneDefaultSettings();
}

export function resetSettings(): PayrollSettings {
  clearLegacyBrowserSettings();
  return cloneDefaultSettings();
}
