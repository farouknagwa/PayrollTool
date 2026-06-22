import { useMemo, useRef, useState } from "react";
import { saveAs } from "file-saver";
import "./App.css";
import { GOTCHA_MESSAGES } from "./config/defaults";
import { loadSettings, resetSettings, saveSettings, validateSettings } from "./config/storage";
import { computePeriod } from "./core/dateTime";
import type { ISODate, PayrollInputFiles, PayrollRunResult, PayrollSettings, PermissionPrepOptions, RunLogEntry, StepMetrics } from "./core/types";
import { mapInputFiles, missingRequiredInputs, permissionMode } from "./io/files";
import { detectAttendancePeriodDate, readWorkbook } from "./io/excel";

type WorkerResponse =
  | { type: "status"; message: string }
  | { type: "log"; entry: RunLogEntry }
  | { type: "done"; result: PayrollRunResult }
  | { type: "prepared"; workbook: ArrayBuffer; summary: string; logs: RunLogEntry[] }
  | { type: "error"; message: string };

const currentYear = new Date().getFullYear();
const currentMonth = new Date().getMonth() + 1;

function downloadBuffer(buffer: ArrayBuffer, filename: string, mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
  saveAs(new Blob([buffer], { type: mime }), filename);
}

function formatLog(entry: RunLogEntry): string {
  return `${new Date(entry.at).toLocaleTimeString()}  ${stepLabel(entry.step)}  ${entry.message}`;
}

function formatMetric(metric: StepMetrics): string {
  const parts: string[] = [];
  if (typeof metric.filled === "number") parts.push(`${metric.filled} filled`);
  if (typeof metric.appended === "number") parts.push(`${metric.appended} appended`);
  if (typeof metric.adjusted === "number") parts.push(`${metric.adjusted} adjusted`);
  if (typeof metric.blank === "number") parts.push(`${metric.blank} blank`);
  if (typeof metric.skippedCode === "number") parts.push(`${metric.skippedCode} skipped by employee code`);
  if (typeof metric.skippedDate === "number") parts.push(`${metric.skippedDate} skipped by date`);
  if (typeof metric.warnings?.length === "number" && metric.warnings.length > 0) parts.push(`${metric.warnings.length} warnings`);
  return parts.length > 0 ? parts.join(", ") : "completed";
}

function stepLabel(step: string): string {
  const labels: Record<string, string> = {
    engine: "Engine",
    settings: "Settings",
    prepare_permission_report: "Prepare permissions",
    extend_nagwa_technologies: "Build detailed calendar",
    fill_attendance: "Fill attendance",
    extend_final_nagwa_technologies: "Build final calendar",
    complete_final: "Complete final report",
    run: "Run",
  };
  return labels[step] ?? step;
}

type AbbreviationRow = {
  id: string;
  label: string;
  code: string;
};

let rowCounter = 0;

function nextRowId(prefix: string): string {
  rowCounter += 1;
  return `${prefix}-${Date.now()}-${rowCounter}`;
}

function abbreviationRowsFrom(abbreviations: PayrollSettings["abbreviations"]): AbbreviationRow[] {
  return Object.entries(abbreviations).map(([label, code], index) => ({
    id: `abbr-${index}-${label}`,
    label,
    code,
  }));
}

function abbreviationsFromRows(rows: AbbreviationRow[]): PayrollSettings["abbreviations"] {
  const abbreviations: PayrollSettings["abbreviations"] = {};
  for (const row of rows) {
    const label = row.label.trim();
    if (label) abbreviations[label] = row.code;
  }
  return abbreviations;
}

function TextInput(props: {
  label: string;
  value: string | number;
  type?: string;
  onChange: (value: string) => void;
}) {
  return (
    <label className="field">
      <span>{props.label}</span>
      <input type={props.type ?? "text"} value={props.value} onChange={(event) => props.onChange(event.target.value)} />
    </label>
  );
}

function UploadTile(props: {
  title: string;
  description: string;
  file?: File;
  optional?: boolean;
  accept?: string;
  multiple?: boolean;
  onFiles: (files: FileList | File[]) => void;
}) {
  const state = props.file ? "found" : props.optional ? "optional" : "missing";
  return (
    <div
      className={`upload-tile ${state}`}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        props.onFiles(event.dataTransfer.files);
      }}
    >
      <div className="upload-tile-header">
        <strong>{props.file ? "Found" : props.optional ? "Optional" : "Missing"}</strong>
        <span>{props.title}</span>
      </div>
      <small>{props.file?.name ?? props.description}</small>
      <label className="file-button">
        Choose File
        <input
          type="file"
          accept={props.accept ?? ".xls,.xlsx"}
          multiple={props.multiple}
          onChange={(event) => {
            if (event.target.files) props.onFiles(event.target.files);
            event.currentTarget.value = "";
          }}
        />
      </label>
      <span className="drop-hint">or drag and drop here</span>
    </div>
  );
}

function App() {
  const [inputs, setInputs] = useState<PayrollInputFiles>({});
  const [settings, setSettings] = useState<PayrollSettings>(() => loadSettings());
  const [abbreviationRows, setAbbreviationRows] = useState<AbbreviationRow[]>(() => abbreviationRowsFrom(loadSettings().abbreviations));
  const [permissionOptions, setPermissionOptions] = useState<PermissionPrepOptions>({
    month: currentMonth,
    year: currentYear,
    requestCutoffDays: settings.requestCutoffDaysDefault,
    noRequestCutoff: false,
  });
  const [standalonePermissionFile, setStandalonePermissionFile] = useState<File | null>(null);
  const [standaloneSummary, setStandaloneSummary] = useState<string>("");
  const [logs, setLogs] = useState<RunLogEntry[]>([]);
  const [metrics, setMetrics] = useState<StepMetrics[]>([]);
  const [result, setResult] = useState<PayrollRunResult | null>(null);
  const [detectedPeriod, setDetectedPeriod] = useState<string>("");
  const [running, setRunning] = useState(false);
  const [preparingPermissions, setPreparingPermissions] = useState(false);
  const [engineStatus, setEngineStatus] = useState("");
  const [error, setError] = useState<string>("");
  const workerRef = useRef<Worker | null>(null);

  const validationErrors = useMemo(() => validateSettings(settings), [settings]);
  const missing = useMemo(() => missingRequiredInputs(inputs), [inputs]);

  async function addFiles(fileList: FileList | File[]) {
    const files = Array.from(fileList);
    const mapped = mapInputFiles(files);
    setInputs((prev) => ({ ...prev, ...mapped }));
    if (mapped.attendance) {
      try {
        const workbook = await readWorkbook(mapped.attendance);
        const firstDate = detectAttendancePeriodDate(workbook);
        if (firstDate) {
          const period = computePeriod(firstDate as ISODate);
          setDetectedPeriod(`${period.start} -> ${period.end}`);
          setPermissionOptions((prev) => ({
            ...prev,
            month: Number(period.end.slice(5, 7)),
            year: Number(period.end.slice(0, 4)),
          }));
        }
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError));
      }
    }
  }

  async function addTemplateFiles(fileList: FileList | File[]) {
    const files = Array.from(fileList);
    const mapped = mapInputFiles(files);
    setInputs((prev) => ({ ...prev, ...mapped }));
  }

  function addStandalonePermissionFiles(fileList: FileList | File[]) {
    const [file] = Array.from(fileList);
    if (!file) return;
    setStandalonePermissionFile(file);
    setStandaloneSummary("");
    setError("");
  }

  function updateSettings(next: PayrollSettings) {
    setSettings(next);
    saveSettings(next);
  }

  function getWorker(): Worker {
    if (workerRef.current === null) {
      workerRef.current = new Worker(new URL("./workers/pyodideWorker.ts", import.meta.url), { type: "module" });
    }
    return workerRef.current;
  }

  function stopWorker() {
    workerRef.current?.terminate();
    workerRef.current = null;
    setRunning(false);
    setPreparingPermissions(false);
    setEngineStatus("Stopped.");
  }

  async function preparePermissionsOnly() {
    if (!standalonePermissionFile) {
      setError("Upload Nagwa_Permission_Request_Report.xls[x] first.");
      return;
    }
    setError("");
    setResult(null);
    setMetrics([]);
    setLogs([]);
    setPreparingPermissions(true);
    setStandaloneSummary("Preparing permissions...");
    const worker = getWorker();
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const message = event.data;
      if (message.type === "status") {
        setEngineStatus(message.message);
      } else if (message.type === "log") {
        setLogs((prev) => [...prev, message.entry]);
      } else if (message.type === "prepared") {
        downloadBuffer(message.workbook, "Nagwa_Permission_Request_permission_details.xls", "application/vnd.ms-excel");
        setLogs(message.logs);
        setStandaloneSummary(message.summary);
        setPreparingPermissions(false);
      } else if (message.type === "error") {
        setStandaloneSummary("");
        setError(message.message);
        setPreparingPermissions(false);
      }
    };
    worker.onerror = (event) => {
      setStandaloneSummary("");
      setError(event.message);
      setPreparingPermissions(false);
    };
    worker.postMessage({ type: "preparePermissions", file: standalonePermissionFile, permissionOptions });
  }

  async function runPayroll() {
    setError("");
    setResult(null);
    setMetrics([]);
    setLogs([]);
    setStandaloneSummary("");
    if (missing.length > 0) {
      setError(`Missing required files: ${missing.join(", ")}`);
      return;
    }
    if (validationErrors.length > 0) {
      setError(`Fix settings before running: ${validationErrors.join(" ")}`);
      return;
    }
    setRunning(true);
    const worker = getWorker();
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const message = event.data;
      if (message.type === "status") {
        setEngineStatus(message.message);
      } else if (message.type === "log") {
        setLogs((prev) => [...prev, message.entry]);
      } else if (message.type === "done") {
        setResult(message.result);
        setMetrics(message.result.metrics);
        setLogs(message.result.logs);
        setRunning(false);
      } else if (message.type === "error") {
        setError(message.message);
        setRunning(false);
      }
    };
    worker.onerror = (event) => {
      setError(event.message);
      setRunning(false);
    };
    worker.postMessage({
      type: "run",
      inputs,
      settings,
      permissionOptions,
    });
  }

  function resetAllSettings() {
    const defaults = resetSettings();
    setSettings(defaults);
    setAbbreviationRows(abbreviationRowsFrom(defaults.abbreviations));
    setPermissionOptions((prev) => ({ ...prev, requestCutoffDays: defaults.requestCutoffDaysDefault }));
  }

  function updateHourReduction(index: number, field: "employeeCode" | "startDate" | "endDate", value: string) {
    const next = [...settings.hourReductionWindows];
    next[index] = {
      ...next[index],
      [field]: field === "employeeCode" ? Number(value) : value === "" ? null : value,
    };
    updateSettings({ ...settings, hourReductionWindows: next });
  }

  function addHourReduction() {
    updateSettings({
      ...settings,
      hourReductionWindows: [
        ...settings.hourReductionWindows,
        { employeeCode: 0, startDate: null, endDate: null },
      ],
    });
  }

  function removeHourReduction(index: number) {
    updateSettings({
      ...settings,
      hourReductionWindows: settings.hourReductionWindows.filter((_, rowIndex) => rowIndex !== index),
    });
  }

  function commitAbbreviationRows(nextRows: AbbreviationRow[]) {
    setAbbreviationRows(nextRows);
    updateSettings({ ...settings, abbreviations: abbreviationsFromRows(nextRows) });
  }

  function updateAbbreviation(id: string, field: "label" | "code", value: string) {
    commitAbbreviationRows(abbreviationRows.map((row) => row.id === id ? { ...row, [field]: value } : row));
  }

  function addAbbreviation() {
    commitAbbreviationRows([...abbreviationRows, { id: nextRowId("abbr"), label: "new leave type", code: "" }]);
  }

  function removeAbbreviation(id: string) {
    commitAbbreviationRows(abbreviationRows.filter((row) => row.id !== id));
  }

  return (
    <main className="app-shell">
      <header className="hero-panel">
        <div>
          <p className="eyebrow">100% browser-side payroll processing</p>
          <h1>PayrollTool Web</h1>
          <p className="lead">
            Run the Nagwa payroll pipeline in the browser. Files stay on this device; no backend, API key, or upload server is used.
          </p>
        </div>
        <div className="hero-actions">
          <button className="primary" type="button" onClick={runPayroll} disabled={running || preparingPermissions || missing.length > 0 || validationErrors.length > 0}>
            {running ? "Running..." : "Run Payroll"}
          </button>
          <button type="button" onClick={stopWorker} disabled={!running && !preparingPermissions}>
            Stop
          </button>
        </div>
      </header>

      {error && <div className="toast error">{error}</div>}
      {engineStatus && <div className="toast info">{engineStatus}</div>}

      <section className="grid two">
        <div
          className="card"
          onDragOver={(event) => event.preventDefault()}
          onDrop={(event) => {
            event.preventDefault();
            void addFiles(event.dataTransfer.files);
          }}
        >
          <h2>Input Files</h2>
          <p>Choose multiple report files, choose the whole raw-data folder, or drag and drop the folder into this box.</p>
          <div className="bulk-upload">
            <label className="file-button">
              Choose Files
              <input
                type="file"
                multiple
                accept=".xls,.xlsx"
                onChange={(event) => {
                  if (event.target.files) void addFiles(event.target.files);
                  event.currentTarget.value = "";
                }}
              />
            </label>
            <label className="file-button">
              Choose Folder
              <input
                type="file"
                multiple
                {...{ webkitdirectory: "true", directory: "true" }}
                onChange={(event) => {
                  if (event.target.files) void addFiles(event.target.files);
                  event.currentTarget.value = "";
                }}
              />
            </label>
            <span className="drop-hint">or drag and drop the whole folder here</span>
          </div>
          <div className="upload-grid">
            <UploadTile
              title="Attendance Report"
              description="Required: Attendance Report.xls or .xlsx"
              file={inputs.attendance}
              onFiles={(files) => void addFiles(files)}
            />
            <UploadTile
              title="Absence Report"
              description="Required: Absence Report.xls or .xlsx"
              file={inputs.absences}
              onFiles={(files) => void addFiles(files)}
            />
            <UploadTile
              title="Employee Transactions_vacations"
              description="Required: Employee Transactions_vacations.xls or .xlsx"
              file={inputs.vacations}
              onFiles={(files) => void addFiles(files)}
            />
            <UploadTile
              title="Permissions"
              description="Required: prepared details or raw request report"
              file={inputs.preparedPermissions ?? inputs.rawPermissions}
              onFiles={(files) => void addFiles(files)}
            />
            <UploadTile
              title="Resignations"
              description="Optional: Resignations.xls or .xlsx"
              file={inputs.resignations}
              optional
              onFiles={(files) => void addFiles(files)}
            />
            <UploadTile
              title="Public Holiday"
              description="Optional: Public Holiday.xls or .xlsx"
              file={inputs.publicHoliday}
              optional
              onFiles={(files) => void addFiles(files)}
            />
          </div>
          <div className="meta-row">
            <span>Permission mode: <strong>{permissionMode(inputs)}</strong></span>
            <span>Detected period: <strong>{detectedPeriod || "Upload Attendance Report"}</strong></span>
          </div>
        </div>

        <div
          className="card"
          onDragOver={(event) => event.preventDefault()}
          onDrop={(event) => {
            event.preventDefault();
            addStandalonePermissionFiles(event.dataTransfer.files);
          }}
        >
          <h2>Prepare Permissions Only</h2>
          <p>Use this when HR only needs to convert `Nagwa_Permission_Request_Report` into the prepared permission-details report.</p>
          <UploadTile
            title="Raw Permission Request Report"
            description="Required: Nagwa_Permission_Request_Report.xls or .xlsx"
            file={standalonePermissionFile ?? undefined}
            onFiles={addStandalonePermissionFiles}
          />
          <div className="controls">
            <TextInput label="Payroll month" type="number" value={permissionOptions.month} onChange={(value) => setPermissionOptions((prev) => ({ ...prev, month: Number(value) }))} />
            <TextInput label="Year" type="number" value={permissionOptions.year} onChange={(value) => setPermissionOptions((prev) => ({ ...prev, year: Number(value) }))} />
            <TextInput label="Cutoff days" type="number" value={permissionOptions.requestCutoffDays} onChange={(value) => setPermissionOptions((prev) => ({ ...prev, requestCutoffDays: Number(value) }))} />
            <label className="toggle">
              <input
                type="checkbox"
                checked={permissionOptions.noRequestCutoff}
                onChange={(event) => setPermissionOptions((prev) => ({ ...prev, noRequestCutoff: event.target.checked }))}
              />
              No request cutoff
            </label>
          </div>
          <button type="button" onClick={preparePermissionsOnly} disabled={running || preparingPermissions}>
            {preparingPermissions ? "Preparing..." : "Prepare Permissions Only"}
          </button>
          {standaloneSummary && <p className="summary-line">{standaloneSummary}</p>}
        </div>
      </section>

      <section
        className="card"
        onDragOver={(event) => event.preventDefault()}
        onDrop={(event) => {
          event.preventDefault();
          void addTemplateFiles(event.dataTransfer.files);
        }}
      >
        <h2>Private Templates</h2>
        <p>Required for real runs: upload the two current styled templates here. They stay in the browser and are not bundled in the public repo.</p>
        <div className="bulk-upload">
          <label className="file-button">
            Choose Files
            <input
              type="file"
              multiple
              accept=".xlsx"
              onChange={(event) => {
                if (event.target.files) void addTemplateFiles(event.target.files);
                event.currentTarget.value = "";
              }}
            />
          </label>
          <span className="drop-hint">or drag and drop the two template files here</span>
        </div>
        <div className="upload-grid two-tiles">
          <UploadTile
            title="Nagwa Technologies template"
            description="Required: Nagwa Technologies.xlsx"
            file={inputs.nagwaTemplate}
            accept=".xlsx"
            onFiles={(files) => void addTemplateFiles(files)}
          />
          <UploadTile
            title="Final Nagwa Technologies template"
            description="Required: Final Nagwa Technologies.xlsx"
            file={inputs.finalTemplate}
            accept=".xlsx"
            onFiles={(files) => void addTemplateFiles(files)}
          />
        </div>
      </section>

      <section className="card">
        <details>
          <summary>Settings / Rules Panel</summary>
          <div className="settings-grid">
            <TextInput label="Ramadan start" type="date" value={settings.ramadanStart} onChange={(value) => updateSettings({ ...settings, ramadanStart: value as PayrollSettings["ramadanStart"] })} />
            <TextInput label="Ramadan end" type="date" value={settings.ramadanEnd} onChange={(value) => updateSettings({ ...settings, ramadanEnd: value as PayrollSettings["ramadanEnd"] })} />
            <TextInput label="Ramadan permitted end" value={settings.permittedWindowEndRamadan} onChange={(value) => updateSettings({ ...settings, permittedWindowEndRamadan: value })} />
            <TextInput label="Lunch switch date" type="date" value={settings.lunchWindowSwitchDate} onChange={(value) => updateSettings({ ...settings, lunchWindowSwitchDate: value as PayrollSettings["lunchWindowSwitchDate"] })} />
            <TextInput label="Lunch before start" value={settings.lunchWindowBefore.start} onChange={(value) => updateSettings({ ...settings, lunchWindowBefore: { ...settings.lunchWindowBefore, start: value } })} />
            <TextInput label="Lunch before end" value={settings.lunchWindowBefore.end} onChange={(value) => updateSettings({ ...settings, lunchWindowBefore: { ...settings.lunchWindowBefore, end: value } })} />
            <TextInput label="Lunch from start" value={settings.lunchWindowFrom.start} onChange={(value) => updateSettings({ ...settings, lunchWindowFrom: { ...settings.lunchWindowFrom, start: value } })} />
            <TextInput label="Lunch from end" value={settings.lunchWindowFrom.end} onChange={(value) => updateSettings({ ...settings, lunchWindowFrom: { ...settings.lunchWindowFrom, end: value } })} />
            <TextInput label="Default request cutoff" type="number" value={settings.requestCutoffDaysDefault} onChange={(value) => updateSettings({ ...settings, requestCutoffDaysDefault: Number(value) })} />
            <label className="toggle field">
              <input type="checkbox" checked={settings.debugMode} onChange={(event) => updateSettings({ ...settings, debugMode: event.target.checked })} />
              Debug mode
            </label>
          </div>

          <div className="rule-columns">
            <section>
              <h3>Schedule Types</h3>
              <p className="helper-text">These names come from the Schedule column in the employee template. The time is the latest time that counts toward attendance.</p>
              {Object.entries(settings.scheduleWindowEnd).map(([name, value]) => (
                <div className="inline-row" key={name}>
                  <input value={name} readOnly />
                  <input value={value} onChange={(event) => updateSettings({ ...settings, scheduleWindowEnd: { ...settings.scheduleWindowEnd, [name]: event.target.value } })} />
                </div>
              ))}
              <button type="button" onClick={() => updateSettings({ ...settings, scheduleWindowEnd: { ...settings.scheduleWindowEnd, newSchedule: "16:00" } })}>Add schedule</button>
            </section>

            <section>
              <h3>Special-Rule Pairs</h3>
              <p className="helper-text">
                Employee A is restricted in the first payroll cycle starting from the “First cycle starts” date. Employee B is restricted in the next cycle, then they alternate every cycle.
              </p>
              {settings.specialRulePairs.map((pair, index) => (
                <div className="editable-card" key={`special-rule-pair-${index}`}>
                  <TextInput label="Employee A" type="number" value={pair.employeeA} onChange={(value) => {
                    const next = [...settings.specialRulePairs];
                    next[index] = { ...pair, employeeA: Number(value) };
                    updateSettings({ ...settings, specialRulePairs: next });
                  }} />
                  <TextInput label="Employee B" type="number" value={pair.employeeB} onChange={(value) => {
                    const next = [...settings.specialRulePairs];
                    next[index] = { ...pair, employeeB: Number(value) };
                    updateSettings({ ...settings, specialRulePairs: next });
                  }} />
                  <TextInput label="First cycle starts" type="date" value={pair.anchorDate} onChange={(value) => {
                    const next = [...settings.specialRulePairs];
                    next[index] = { ...pair, anchorDate: value as PayrollSettings["specialRulePairs"][number]["anchorDate"] };
                    updateSettings({ ...settings, specialRulePairs: next });
                  }} />
                  <button type="button" className="secondary-action" onClick={() => {
                    const next = [...settings.specialRulePairs];
                    next[index] = { ...pair, employeeA: pair.employeeB, employeeB: pair.employeeA };
                    updateSettings({ ...settings, specialRulePairs: next });
                  }}>
                    ↔ Swap A/B
                  </button>
                </div>
              ))}
              <button type="button" onClick={() => updateSettings({ ...settings, specialRulePairs: [...settings.specialRulePairs, { employeeA: 0, employeeB: 0, anchorDate: "2026-02-21" }] })}>Add pair</button>
            </section>
          </div>

          <div className="friendly-tables">
            <section>
              <h3>Hour-Reduction Employees</h3>
              <p className="helper-text">Use this when an employee gets 1 hour deducted from daily shortage during a date range. Leave Start or End empty for an open range.</p>
              <div className="editable-table">
                <div className="table-head four-cols">
                  <span>Employee Code</span>
                  <span>Start Date</span>
                  <span>End Date</span>
                  <span>Action</span>
                </div>
                {settings.hourReductionWindows.map((window, index) => (
                  <div className="table-row four-cols" key={`hour-reduction-${index}`}>
                    <input type="number" value={window.employeeCode} onChange={(event) => updateHourReduction(index, "employeeCode", event.target.value)} />
                    <input type="date" value={window.startDate ?? ""} onChange={(event) => updateHourReduction(index, "startDate", event.target.value)} />
                    <input type="date" value={window.endDate ?? ""} onChange={(event) => updateHourReduction(index, "endDate", event.target.value)} />
                    <button type="button" className="secondary-action" onClick={() => removeHourReduction(index)}>Remove</button>
                  </div>
                ))}
              </div>
              <button type="button" onClick={addHourReduction}>Add employee</button>
            </section>

            <section>
              <h3>Final Report Abbreviations</h3>
              <p className="helper-text">When the final report sees a long leave label, it writes the short code here. Matching ignores uppercase/lowercase and extra spaces.</p>
              <div className="editable-table">
                <div className="table-head three-cols">
                  <span>Leave Label</span>
                  <span>Short Code</span>
                  <span>Action</span>
                </div>
                {abbreviationRows.map((row) => (
                  <div className="table-row three-cols" key={row.id}>
                    <input value={row.label} onChange={(event) => updateAbbreviation(row.id, "label", event.target.value)} />
                    <input value={row.code} onChange={(event) => updateAbbreviation(row.id, "code", event.target.value)} />
                    <button type="button" className="secondary-action" onClick={() => removeAbbreviation(row.id)}>Remove</button>
                  </div>
                ))}
              </div>
              <button type="button" onClick={addAbbreviation}>Add abbreviation</button>
            </section>
          </div>
          <div className="settings-reset-actions">
            <div className="reset-action-card">
              <button type="button" onClick={resetAllSettings}>Restore default settings</button>
              <p>Use this to return all payroll rules and the permission cutoff to the standard values shipped with the tool.</p>
            </div>
          </div>
          {validationErrors.length > 0 && (
            <ul className="validation">
              {validationErrors.map((item) => <li key={item}>{item}</li>)}
            </ul>
          )}
        </details>
      </section>

      <section className="grid two">
        <div className="card">
          <h2>Run Log</h2>
          <p className="helper-text">
            If a public-holiday date says it was not found, that date is outside the detected 21-to-20 payroll period or falls on a Friday/Saturday, so there is no workday column to mark.
          </p>
          <ol className="steps">
            {["extend_nagwa_technologies", "fill_attendance", "extend_final_nagwa_technologies", "complete_final"].map((step, index) => (
              <li key={step} className={logs.some((entry) => entry.step === step) ? "active" : ""}>{index + 1}/4 {stepLabel(step)}</li>
            ))}
          </ol>
          <div className="log-panel">
            {logs.length === 0 ? <p>No run yet.</p> : logs.map((entry, index) => <pre key={`${entry.at}-${index}`} className={entry.level}>{formatLog(entry)}</pre>)}
          </div>
        </div>

        <div className="card">
          <h2>Output</h2>
          {result ? (
            <div className="downloads">
              <p><strong>Success.</strong> {result.employeesProcessed} employees, {result.period.dates.length} dates, {result.warnings.length} warnings.</p>
              <button type="button" onClick={() => downloadBuffer(result.detailedWorkbook, "Nagwa Technologies.xlsx")}>Download Nagwa Technologies.xlsx</button>
              <button type="button" onClick={() => downloadBuffer(result.finalWorkbook, "Final Nagwa Technologies.xlsx")}>Download Final Nagwa Technologies.xlsx</button>
              {result.preparedPermissionsWorkbook && <button type="button" onClick={() => downloadBuffer(result.preparedPermissionsWorkbook as ArrayBuffer, "Nagwa_Permission_Request_permission_details.xls", "application/vnd.ms-excel")}>Download prepared permissions</button>}
            </div>
          ) : (
            <p>Outputs will appear after a successful run.</p>
          )}
          <div className="metrics">
            {metrics.map((metric) => <p key={metric.step}><strong>{stepLabel(metric.step)}</strong>: {formatMetric(metric)}</p>)}
          </div>
        </div>
      </section>

      <section className="card">
        <h2>Gotchas surfaced before every run</h2>
        <div className="gotchas">
          {GOTCHA_MESSAGES.map((message) => <span key={message}>{message}</span>)}
        </div>
      </section>
    </main>
  );
}

export default App;
