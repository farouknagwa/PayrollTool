import type { ISODate, PayrollInputFiles, PayrollRunResult, PayrollSettings, PermissionPrepOptions, RunLogEntry, StepMetrics } from "../core/types";
import extendNagwaSource from "../py/extend_nagwa_technologies.py?raw";
import fillAttendanceSource from "../py/fill_attendance.py?raw";
import extendFinalSource from "../py/extend_final_nagwa_technologies.py?raw";
import completeFinalSource from "../py/complete_final.py?raw";
import preparePermissionSource from "../py/prepare_permission_report.py?raw";

type PyodideLoadModule = {
  loadPyodide: (options: { indexURL: string; createPyodideModule?: unknown }) => Promise<Pyodide>;
};

type PyodideAsmModule = {
  default: unknown;
};

type Pyodide = {
  FS: {
    mkdirTree: (path: string) => void;
    writeFile: (path: string, data: Uint8Array | string) => void;
    readFile: (path: string, options?: { encoding?: "binary" }) => Uint8Array;
  };
  loadPackage: (packages: string | string[]) => Promise<void>;
  runPython: (code: string) => unknown;
  runPythonAsync: (code: string) => Promise<unknown>;
  setStdout: (options: { batched: (line: string) => void }) => void;
  globals: {
    set: (name: string, value: unknown) => void;
    delete: (name: string) => void;
  };
};

type RunRequest = {
  type: "run";
  inputs: PayrollInputFiles;
  settings: PayrollSettings;
  permissionOptions: PermissionPrepOptions;
};

type PrepareRequest = {
  type: "preparePermissions";
  file: File;
  permissionOptions: PermissionPrepOptions;
};

type WorkerRequest = RunRequest | PrepareRequest;

type WorkerResponse =
  | { type: "status"; message: string }
  | { type: "log"; entry: RunLogEntry }
  | { type: "done"; result: PayrollRunResult }
  | { type: "prepared"; workbook: ArrayBuffer; summary: string; logs: RunLogEntry[] }
  | { type: "error"; message: string };

const PYODIDE_MODULE_URLS = [
  "https://cdn.jsdelivr.net/pyodide/v314.0.0/full/pyodide.mjs",
  "https://cdn.jsdelivr.net/pyodide/v0.29.0/full/pyodide.mjs",
];

const PROJECT_ROOT = "/project";
const RAW_DATA_DIR = `${PROJECT_ROOT}/raw data`;
const TEMPLATES_DIR = `${PROJECT_ROOT}/templates`;
const OUTPUT_DIR = `${PROJECT_ROOT}/output`;
const SCRIPTS_DIR = `${PROJECT_ROOT}/scripts`;
const PREPARE_DIR = `${PROJECT_ROOT}/Prepare_Nagwa_Permission_Request_permission_details_Report`;

let pyodidePromise: Promise<Pyodide> | null = null;
let currentStep = "engine";
let logs: RunLogEntry[] = [];

function nowLog(step: string, message: string, level: RunLogEntry["level"] = "info"): RunLogEntry {
  return { at: new Date().toISOString(), step, level, message };
}

function post(message: WorkerResponse) {
  self.postMessage(message);
}

function postStatus(message: string) {
  post({ type: "status", message });
}

function levelFromLine(line: string): RunLogEntry["level"] {
  if (/error|traceback|failed/i.test(line)) return "error";
  if (/warning/i.test(line)) return "warn";
  if (/saved|done|completed successfully/i.test(line)) return "success";
  return "info";
}

function postLog(step: string, message: string, level = levelFromLine(message)) {
  const trimmed = message.trim();
  if (!trimmed) return;
  if (/holiday date .* not found in Nagwa workday columns/i.test(trimmed)) return;
  if (/holiday date\(s\) not found in Nagwa sheet/i.test(trimmed)) return;
  if (/\b(rows?|day-slots?|dates?) skipped\b/i.test(trimmed)) return;
  if (/No data warnings\./i.test(trimmed)) return;
  const entry = nowLog(step, trimmed, level);
  logs.push(entry);
  post({ type: "log", entry });
}

async function importPyodideLoader(): Promise<{ module: PyodideLoadModule; createPyodideModule: unknown; indexURL: string }> {
  let lastError: unknown;
  for (const url of PYODIDE_MODULE_URLS) {
    try {
      const module = await import(/* @vite-ignore */ url) as PyodideLoadModule;
      const asmModule = await import(/* @vite-ignore */ url.replace(/pyodide\.mjs$/, "pyodide.asm.mjs")) as PyodideAsmModule;
      return { module, createPyodideModule: asmModule.default, indexURL: url.replace(/pyodide\.mjs$/, "") };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError instanceof Error ? lastError : new Error(String(lastError));
}

async function getPyodide(): Promise<Pyodide> {
  if (pyodidePromise) return pyodidePromise;
  pyodidePromise = (async () => {
    postStatus("Loading Python engine. First load can take a few minutes; later runs use the browser cache.");
    const { module, createPyodideModule, indexURL } = await importPyodideLoader();
    const pyodide = await module.loadPyodide({ indexURL, createPyodideModule });
    pyodide.setStdout({
      batched: (line) => postLog(currentStep, line),
    });

    postStatus("Loading payroll Python packages.");
    await pyodide.loadPackage(["pandas", "xlrd", "micropip"]);
    try {
      await pyodide.loadPackage("python-calamine");
    } catch {
      await pyodide.runPythonAsync(`
import micropip
await micropip.install("python-calamine")
`);
    }
    await pyodide.runPythonAsync(`
import micropip
await micropip.install(["openpyxl", "xlwt"])
`);
    postStatus("Python engine ready.");
    return pyodide;
  })();
  return pyodidePromise;
}

function resetLogs() {
  logs = [];
  currentStep = "engine";
}

async function resetProjectFS(pyodide: Pyodide) {
  pyodide.runPython(`
import os
import shutil
shutil.rmtree("${PROJECT_ROOT}", ignore_errors=True)
os.makedirs("${RAW_DATA_DIR}", exist_ok=True)
os.makedirs("${TEMPLATES_DIR}", exist_ok=True)
os.makedirs("${OUTPUT_DIR}", exist_ok=True)
os.makedirs("${SCRIPTS_DIR}", exist_ok=True)
os.makedirs("${PREPARE_DIR}", exist_ok=True)
`);
  pyodide.FS.writeFile(`${SCRIPTS_DIR}/__init__.py`, "");
}

function writePythonSources(pyodide: Pyodide) {
  pyodide.FS.writeFile(`${SCRIPTS_DIR}/extend_nagwa_technologies.py`, extendNagwaSource);
  pyodide.FS.writeFile(`${SCRIPTS_DIR}/fill_attendance.py`, fillAttendanceSource);
  pyodide.FS.writeFile(`${SCRIPTS_DIR}/extend_final_nagwa_technologies.py`, extendFinalSource);
  pyodide.FS.writeFile(`${SCRIPTS_DIR}/complete_final.py`, completeFinalSource);
  pyodide.FS.writeFile(`${PREPARE_DIR}/prepare_permission_report.py`, preparePermissionSource);
}

function extensionOf(file: File, fallback: "xls" | "xlsx" = "xlsx"): "xls" | "xlsx" {
  return /\.xls$/i.test(file.name) ? "xls" : /\.xlsx$/i.test(file.name) ? "xlsx" : fallback;
}

function requireFile(file: File | undefined, label: string): File {
  if (!file) throw new Error(`Missing required file: ${label}`);
  return file;
}

async function writeUploadedFile(pyodide: Pyodide, file: File, path: string) {
  const buffer = await file.arrayBuffer();
  pyodide.FS.writeFile(path, new Uint8Array(buffer));
}

async function writePipelineInputs(pyodide: Pyodide, inputs: PayrollInputFiles) {
  const attendance = requireFile(inputs.attendance, "Attendance Report");
  const absences = requireFile(inputs.absences, "Absence Report");
  const vacations = requireFile(inputs.vacations, "Employee Transactions_vacations");
  const publicHoliday = requireFile(inputs.publicHoliday, "Public Holiday");
  const nagwaTemplate = requireFile(inputs.nagwaTemplate, "Nagwa Technologies template");
  const finalTemplate = requireFile(inputs.finalTemplate, "Final Nagwa Technologies template");

  await writeUploadedFile(pyodide, attendance, `${RAW_DATA_DIR}/Attendance Report.${extensionOf(attendance)}`);
  await writeUploadedFile(pyodide, absences, `${RAW_DATA_DIR}/Absence Report.${extensionOf(absences)}`);
  await writeUploadedFile(pyodide, vacations, `${RAW_DATA_DIR}/Employee Transactions_vacations.${extensionOf(vacations)}`);
  await writeUploadedFile(pyodide, publicHoliday, `${RAW_DATA_DIR}/Public Holiday.${extensionOf(publicHoliday)}`);
  await writeUploadedFile(pyodide, nagwaTemplate, `${TEMPLATES_DIR}/Nagwa Technologies.xlsx`);
  await writeUploadedFile(pyodide, finalTemplate, `${TEMPLATES_DIR}/Final Nagwa Technologies.xlsx`);

  if (inputs.resignations) {
    await writeUploadedFile(pyodide, inputs.resignations, `${RAW_DATA_DIR}/Resignations.${extensionOf(inputs.resignations)}`);
  }
}

function runPythonWithGlobals(pyodide: Pyodide, code: string, globals: Record<string, unknown> = {}) {
  for (const [key, value] of Object.entries(globals)) pyodide.globals.set(key, value);
  return pyodide.runPythonAsync(code).finally(() => {
    for (const key of Object.keys(globals)) pyodide.globals.delete(key);
  });
}

async function injectSettings(pyodide: Pyodide, settings: PayrollSettings) {
  currentStep = "settings";
  await runPythonWithGlobals(pyodide, `
import datetime as _dt
import json as _json
import sys as _sys

for _name in [
    "scripts.fill_attendance",
    "scripts.complete_final",
    "scripts.extend_nagwa_technologies",
    "scripts.extend_final_nagwa_technologies",
    "prepare_permission_report",
]:
    _sys.modules.pop(_name, None)

import scripts.fill_attendance as fill_attendance
import scripts.complete_final as complete_final
import prepare_permission_report

_settings = _json.loads(_settings_json)

def _date(value):
    return _dt.datetime.strptime(value, "%Y-%m-%d").date()

def _time(value):
    hour, minute = str(value).split(":")[:2]
    return _dt.time(int(hour), int(minute))

fill_attendance.RAMADAN_START = _date(_settings["ramadanStart"])
fill_attendance.RAMADAN_END = _date(_settings["ramadanEnd"])
fill_attendance.WORKDAY_START = _time(_settings["workdayStart"])
fill_attendance.WORKDAY_END_NORMAL = _time(_settings["workdayEndNormal"])
fill_attendance.WORKDAY_END_RAMADAN = _time(_settings["workdayEndRamadan"])
fill_attendance.FULL_DAY_NORMAL = int(_settings["fullDayNormalMinutes"])
fill_attendance.FULL_DAY_RAMADAN = int(_settings["fullDayRamadanMinutes"])
fill_attendance.PERMITTED_WINDOW_END_NORMAL = _time(_settings["permittedWindowEndNormal"])
fill_attendance.PERMITTED_WINDOW_END_RAMADAN = _time(_settings["permittedWindowEndRamadan"])
fill_attendance.PERMITTED_WINDOW_END_RESTRICTED = _time(_settings["permittedWindowEndRestricted"])
fill_attendance.SCHEDULE_WINDOW_END = {str(k).strip().lower(): _time(v) for k, v in _settings["scheduleWindowEnd"].items()}
fill_attendance.DEFAULT_SCHEDULE_WINDOW_END = _time(_settings["defaultScheduleWindowEnd"])

_pairs = [
    (int(pair["employeeA"]), int(pair["employeeB"]), _date(pair["anchorDate"]))
    for pair in _settings["specialRulePairs"]
]
if not _pairs:
    _pairs = [(1052, 100, _dt.date(2026, 2, 21))]
fill_attendance.SPECIAL_RULE_PAIRS = _pairs
fill_attendance.SPECIAL_RULE_EMPLOYEE_A = _pairs[0][0]
fill_attendance.SPECIAL_RULE_EMPLOYEE_B = _pairs[0][1]
fill_attendance.SPECIAL_RULE_ANCHOR = _pairs[0][2]
fill_attendance.SPECIAL_RULE_EMPLOYEES = {employee for pair in _pairs for employee in pair[:2]}

fill_attendance.LUNCH_WINDOW_SWITCH_DATE = _date(_settings["lunchWindowSwitchDate"])
fill_attendance.LUNCH_WINDOW_BEFORE = (_time(_settings["lunchWindowBefore"]["start"]), _time(_settings["lunchWindowBefore"]["end"]))
fill_attendance.LUNCH_WINDOW_FROM = (_time(_settings["lunchWindowFrom"]["start"]), _time(_settings["lunchWindowFrom"]["end"]))
fill_attendance.HOUR_REDUCTION_WINDOWS = {
    int(row["employeeCode"]): (
        None if row.get("startDate") is None else _date(row["startDate"]),
        None if row.get("endDate") is None else _date(row["endDate"]),
    )
    for row in _settings["hourReductionWindows"]
}

complete_final.ABBREVIATIONS = {str(k): str(v) for k, v in _settings["abbreviations"].items()}
complete_final._ABBREV_LOOKUP = {
    complete_final._normalise_label(k): v for k, v in complete_final.ABBREVIATIONS.items()
}
`, { _settings_json: JSON.stringify(settings) });
}

function prepareArgs(inputPath: string, outputPath: string, options: PermissionPrepOptions): string[] {
  const args = [
    "prepare_permission_report.py",
    "--input",
    inputPath,
    "--output",
    outputPath,
    "--month",
    String(options.month),
    "--year",
    String(options.year),
  ];
  if (options.noRequestCutoff) {
    args.push("--no-request-cutoff");
  } else {
    args.push("--request-cutoff-days", String(options.requestCutoffDays));
  }
  return args;
}

async function runPrepare(pyodide: Pyodide, inputPath: string, outputPath: string, options: PermissionPrepOptions) {
  currentStep = "prepare_permission_report";
  postLog(currentStep, "Preparing permission details report.");
  await runPythonWithGlobals(pyodide, `
import json as _json
import sys as _sys
import prepare_permission_report
_sys.argv = _json.loads(_argv_json)
prepare_permission_report.main()
`, { _argv_json: JSON.stringify(prepareArgs(inputPath, outputPath, options)) });
}

async function runScriptMain(pyodide: Pyodide, step: string, moduleName: string) {
  currentStep = step;
  postLog(step, `Starting ${step}.`);
  await runPythonWithGlobals(pyodide, `
import sys as _sys
import ${moduleName} as _module
_sys.argv = ["${step}.py"]
_module.main()
`);
}

function readOutput(pyodide: Pyodide, path: string): ArrayBuffer {
  const bytes = pyodide.FS.readFile(path, { encoding: "binary" });
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

async function getRunInfo(pyodide: Pyodide): Promise<{ period: PayrollRunResult["period"]; employeesProcessed: number }> {
  const infoJson = await runPythonWithGlobals(pyodide, `
import json
from openpyxl import load_workbook
from scripts.extend_nagwa_technologies import compute_period, daterange, detect_first_attendance_date

_first = detect_first_attendance_date(_attendance_path)
_start, _end = compute_period(_first)
_dates = [day.isoformat() for day in daterange(_start, _end)]
_wb = load_workbook("${OUTPUT_DIR}/Nagwa Technologies.xlsx", read_only=True, data_only=True)
_ws = _wb["Nagwa Technologies"]
_employees = 0
for _row in range(4, _ws.max_row + 1):
    if _ws.cell(_row, 1).value is not None:
        _employees += 1
_wb.close()
json.dumps({
    "period": {"start": _start.isoformat(), "end": _end.isoformat(), "dates": _dates},
    "employeesProcessed": _employees,
})
`, { _attendance_path: `${RAW_DATA_DIR}/Attendance Report.xlsx` }).catch(async () => {
    const fallbackJson = await runPythonWithGlobals(pyodide, `
import json
from openpyxl import load_workbook
from scripts.extend_nagwa_technologies import compute_period, daterange, detect_first_attendance_date

_first = detect_first_attendance_date("${RAW_DATA_DIR}/Attendance Report.xls")
_start, _end = compute_period(_first)
_dates = [day.isoformat() for day in daterange(_start, _end)]
_wb = load_workbook("${OUTPUT_DIR}/Nagwa Technologies.xlsx", read_only=True, data_only=True)
_ws = _wb["Nagwa Technologies"]
_employees = 0
for _row in range(4, _ws.max_row + 1):
    if _ws.cell(_row, 1).value is not None:
        _employees += 1
_wb.close()
json.dumps({
    "period": {"start": _start.isoformat(), "end": _end.isoformat(), "dates": _dates},
    "employeesProcessed": _employees,
})
`);
    return fallbackJson;
  });
  const parsed = JSON.parse(String(infoJson)) as { period: { start: ISODate; end: ISODate; dates: ISODate[] }; employeesProcessed: number };
  return parsed;
}

function standardMetrics(): StepMetrics[] {
  return [
    { step: "extend_nagwa_technologies" },
    { step: "fill_attendance" },
    { step: "extend_final_nagwa_technologies" },
    { step: "complete_final" },
  ];
}

async function preparePyodideRun(pyodide: Pyodide) {
  await resetProjectFS(pyodide);
  writePythonSources(pyodide);
  pyodide.runPython(`
import sys
sys.path.insert(0, "${PROJECT_ROOT}")
sys.path.insert(0, "${PREPARE_DIR}")
`);
}

async function runFullPipeline(request: RunRequest) {
  resetLogs();
  const pyodide = await getPyodide();
  await preparePyodideRun(pyodide);
  await writePipelineInputs(pyodide, request.inputs);
  await injectSettings(pyodide, request.settings);

  let preparedPermissionsWorkbook: ArrayBuffer | undefined;
  if (request.inputs.rawPermissions) {
    const rawExt = extensionOf(request.inputs.rawPermissions, "xls");
    const rawPath = `${RAW_DATA_DIR}/Nagwa_Permission_Request_Report.${rawExt}`;
    const preparedPath = `${RAW_DATA_DIR}/Nagwa_Permission_Request_permission_details.xls`;
    await writeUploadedFile(pyodide, request.inputs.rawPermissions, rawPath);
    await runPrepare(pyodide, rawPath, preparedPath, request.permissionOptions);
    preparedPermissionsWorkbook = readOutput(pyodide, preparedPath);
  } else {
    const prepared = requireFile(request.inputs.preparedPermissions, "Nagwa_Permission_Request_permission_details");
    await writeUploadedFile(pyodide, prepared, `${RAW_DATA_DIR}/Nagwa_Permission_Request_permission_details.${extensionOf(prepared, "xls")}`);
  }

  await runScriptMain(pyodide, "extend_nagwa_technologies", "scripts.extend_nagwa_technologies");
  await runScriptMain(pyodide, "fill_attendance", "scripts.fill_attendance");
  await runScriptMain(pyodide, "extend_final_nagwa_technologies", "scripts.extend_final_nagwa_technologies");
  await runScriptMain(pyodide, "complete_final", "scripts.complete_final");
  postLog("run", "All scripts completed successfully.", "success");

  const detailedWorkbook = readOutput(pyodide, `${OUTPUT_DIR}/Nagwa Technologies.xlsx`);
  const finalWorkbook = readOutput(pyodide, `${OUTPUT_DIR}/Final Nagwa Technologies.xlsx`);
  const runInfo = await getRunInfo(pyodide);
  const warnings = logs.filter((entry) => entry.level === "warn").map((entry) => entry.message);

  const result: PayrollRunResult = {
    detailedWorkbook,
    finalWorkbook,
    preparedPermissionsWorkbook,
    metrics: standardMetrics(),
    logs,
    period: runInfo.period,
    employeesProcessed: runInfo.employeesProcessed,
    warnings,
  };
  post({ type: "done", result });
}

async function runStandalonePrepare(request: PrepareRequest) {
  resetLogs();
  const pyodide = await getPyodide();
  await preparePyodideRun(pyodide);
  const inputPath = `${PREPARE_DIR}/Nagwa_Permission_Request_Report.${extensionOf(request.file, "xls")}`;
  const outputPath = `${PREPARE_DIR}/Nagwa_Permission_Request_permission_details.xls`;
  await writeUploadedFile(pyodide, request.file, inputPath);
  await runPrepare(pyodide, inputPath, outputPath, request.permissionOptions);
  const workbook = readOutput(pyodide, outputPath);
  post({
    type: "prepared",
    workbook,
    summary: "Prepared permissions report is ready. See the Run Log for row counts.",
    logs,
  });
}

self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  try {
    if (event.data.type === "run") {
      await runFullPipeline(event.data);
    } else {
      await runStandalonePrepare(event.data);
    }
  } catch (error) {
    post({ type: "error", message: error instanceof Error ? error.message : String(error) });
  }
};
