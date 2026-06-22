import type { PayrollSettings } from "../core/types";

export const DEFAULT_SETTINGS: PayrollSettings = {
  ramadanStart: "2026-02-20",
  ramadanEnd: "2026-03-18",
  workdayStart: "08:00",
  workdayEndNormal: "16:00",
  workdayEndRamadan: "14:30",
  fullDayNormalMinutes: 480,
  fullDayRamadanMinutes: 390,
  permittedWindowEndNormal: "18:00",
  permittedWindowEndRamadan: "16:30",
  permittedWindowEndRestricted: "16:00",
  scheduleWindowEnd: {
    standard: "16:00",
    undefined: "16:00",
    flexy: "18:00",
    flexy9: "17:00",
  },
  defaultScheduleWindowEnd: "16:00",
  specialRulePairs: [
    {
      employeeA: 1052,
      employeeB: 100,
      anchorDate: "2026-02-21",
    },
  ],
  lunchWindowSwitchDate: "2026-04-26",
  lunchWindowBefore: {
    start: "12:00",
    end: "13:00",
  },
  lunchWindowFrom: {
    start: "12:30",
    end: "13:30",
  },
  hourReductionWindows: [
    { employeeCode: 514, startDate: "2025-03-09", endDate: "2026-04-07" },
    { employeeCode: 2148, startDate: "2024-09-18", endDate: "2026-06-20" },
    { employeeCode: 350, startDate: "2025-08-03", endDate: "2026-10-27" },
    { employeeCode: 1809, startDate: null, endDate: "2026-11-01" },
    { employeeCode: 822, startDate: "2025-03-10", endDate: "2026-12-10" },
    { employeeCode: 315, startDate: "2025-09-07", endDate: "2027-06-06" },
  ],
  abbreviations: {
    absent: "A",
    "unpaid leave": "UL",
    "permitted absence during probation": "PD",
    "permitted absence": "PA",
    "5 days sick leave": "5S",
    "severe illness sick leave": "SS",
    "severe sick 85%": "S",
    "unpaid sick leave": "US",
    "work from home": "WFH",
  },
  requestCutoffDaysDefault: 0,
  debugMode: false,
};

export const REQUIRED_BASENAMES = {
  attendance: "Attendance Report",
  absences: "Absence Report",
  vacations: "Employee Transactions_vacations",
  preparedPermissions: "Nagwa_Permission_Request_permission_details",
} as const;

export const OPTIONAL_BASENAMES = {
  resignations: "Resignations",
  publicHoliday: "Public Holiday",
  alternatePermissions: "permissions",
} as const;

export const RAW_PERMISSION_BASENAME = "Nagwa_Permission_Request_Report";

export const TEMPLATE_BASENAMES = {
  nagwaTemplate: "Nagwa Technologies",
  finalTemplate: "Final Nagwa Technologies",
} as const;

export const GOTCHA_MESSAGES = [
  "Pipeline order is critical: WFH/Workday override always runs last.",
  "Permitted delays are skipped when a half-day leave exists on the same day.",
  "Missing Punch sweep ignores known status labels such as absent, Public Holiday, Resigned, and vacation types.",
  "Multiple permissions on the same day are joined with \" | \" in the Leave column.",
  "Workday takes priority over Work from Home when both keywords are present.",
  "File names must match exact basenames; the resolver tries .xlsx first, then .xls.",
  "If detected period differs from selected month, the Attendance Report period controls the full payroll pipeline.",
  "Final Total = 0 is normal when all daily cells are text labels rather than durations.",
];
