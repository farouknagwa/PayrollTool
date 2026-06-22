import { OPTIONAL_BASENAMES, RAW_PERMISSION_BASENAME, REQUIRED_BASENAMES, TEMPLATE_BASENAMES } from "../config/defaults";
import type { PayrollInputFiles } from "../core/types";

type FileKey = keyof PayrollInputFiles;

const basenameToKey: Array<[FileKey, string]> = [
  ["attendance", REQUIRED_BASENAMES.attendance],
  ["absences", REQUIRED_BASENAMES.absences],
  ["vacations", REQUIRED_BASENAMES.vacations],
  ["preparedPermissions", REQUIRED_BASENAMES.preparedPermissions],
  ["rawPermissions", RAW_PERMISSION_BASENAME],
  ["resignations", OPTIONAL_BASENAMES.resignations],
  ["publicHoliday", OPTIONAL_BASENAMES.publicHoliday],
  ["alternatePermissions", OPTIONAL_BASENAMES.alternatePermissions],
  ["nagwaTemplate", TEMPLATE_BASENAMES.nagwaTemplate],
  ["finalTemplate", TEMPLATE_BASENAMES.finalTemplate],
];

function cleanName(name: string): string {
  const file = name.split("/").at(-1)?.split("\\").at(-1) ?? name;
  return file.replace(/\.(xlsx|xls)$/i, "");
}

function extensionRank(name: string): number {
  if (/\.xlsx$/i.test(name)) return 0;
  if (/\.xls$/i.test(name)) return 1;
  return 2;
}

export function mapInputFiles(files: File[]): PayrollInputFiles {
  const mapped: PayrollInputFiles = {};
  const sorted = [...files].sort((a, b) => extensionRank(a.name) - extensionRank(b.name));
  for (const file of sorted) {
    const base = cleanName(file.name);
    const match = basenameToKey.find(([, expected]) => base === expected);
    if (match && mapped[match[0]] === undefined) {
      mapped[match[0]] = file;
    }
  }
  if (mapped.preparedPermissions === undefined && mapped.alternatePermissions !== undefined) {
    mapped.preparedPermissions = mapped.alternatePermissions;
  }
  return mapped;
}

export function fileChecklist(inputs: PayrollInputFiles): Array<{ label: string; state: "found" | "missing" | "optional" | "warning"; detail: string }> {
  return [
    {
      label: "Attendance Report",
      state: inputs.attendance ? "found" : "missing",
      detail: inputs.attendance?.name ?? "Required: Attendance Report.xls or .xlsx",
    },
    {
      label: "Absence Report",
      state: inputs.absences ? "found" : "missing",
      detail: inputs.absences?.name ?? "Required: Absence Report.xls or .xlsx",
    },
    {
      label: "Employee Transactions_vacations",
      state: inputs.vacations ? "found" : "missing",
      detail: inputs.vacations?.name ?? "Required: Employee Transactions_vacations.xls or .xlsx",
    },
    {
      label: "Permissions",
      state: inputs.preparedPermissions || inputs.rawPermissions ? "found" : "missing",
      detail: inputs.preparedPermissions?.name ?? inputs.rawPermissions?.name ?? "Required: prepared details or raw request report",
    },
    {
      label: "Resignations",
      state: inputs.resignations ? "found" : "optional",
      detail: inputs.resignations?.name ?? "Optional; skipped if missing",
    },
    {
      label: "Public Holiday",
      state: inputs.publicHoliday ? "found" : "optional",
      detail: inputs.publicHoliday?.name ?? "Optional; skipped if missing",
    },
  ];
}

export function missingRequiredInputs(inputs: PayrollInputFiles): string[] {
  const missing: string[] = [];
  if (!inputs.attendance) missing.push("Attendance Report.xls[x]");
  if (!inputs.absences) missing.push("Absence Report.xls[x]");
  if (!inputs.vacations) missing.push("Employee Transactions_vacations.xls[x]");
  if (!inputs.preparedPermissions && !inputs.rawPermissions) {
    missing.push("Nagwa_Permission_Request_permission_details.xls[x] or Nagwa_Permission_Request_Report.xls[x]");
  }
  return missing;
}

export function permissionMode(inputs: PayrollInputFiles): "prepared" | "raw" | "missing" {
  if (inputs.rawPermissions) return "raw";
  if (inputs.preparedPermissions) return "prepared";
  return "missing";
}
