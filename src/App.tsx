import { useMemo, useRef, useState } from "react";
import { saveAs } from "file-saver";
import "./App.css";
import { loadSettings, resetSettings, saveSettings, validateSettings } from "./config/storage";
import { computePeriod } from "./core/dateTime";
import type { ISODate, PayrollInputFiles, PayrollRunResult, PayrollSettings, PermissionPrepOptions, RunLogEntry, StepMetrics } from "./core/types";
import { missingRequiredInputs, permissionMode } from "./io/files";
import { detectAttendancePeriodDate, readWorkbook } from "./io/excel";

type WorkerResponse =
  | { type: "status"; message: string }
  | { type: "log"; entry: RunLogEntry }
  | { type: "done"; result: PayrollRunResult }
  | { type: "error"; message: string };

type AppTab = "inputs" | "settings" | "log" | "outputs";

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
  onFile: (file: File) => void;
}) {
  const state = props.file ? "found" : props.optional ? "optional" : "missing";
  return (
    <div
      className={`upload-tile ${state}`}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const file = event.dataTransfer.files.item(0);
        if (file) props.onFile(file);
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
          onChange={(event) => {
            const file = event.target.files?.item(0);
            if (file) props.onFile(file);
            event.currentTarget.value = "";
          }}
        />
      </label>
      <span className="drop-hint">or drag and drop here</span>
    </div>
  );
}

function App() {
  const [activeTab, setActiveTab] = useState<AppTab>("inputs");
  const [inputs, setInputs] = useState<PayrollInputFiles>({});
  const [settings, setSettings] = useState<PayrollSettings>(() => loadSettings());
  const [permissionOptions, setPermissionOptions] = useState<PermissionPrepOptions>({
    month: currentMonth,
    year: currentYear,
    requestCutoffDays: settings.requestCutoffDaysDefault,
    noRequestCutoff: false,
  });
  const [logs, setLogs] = useState<RunLogEntry[]>([]);
  const [metrics, setMetrics] = useState<StepMetrics[]>([]);
  const [result, setResult] = useState<PayrollRunResult | null>(null);
  const [detectedPeriod, setDetectedPeriod] = useState<string>("");
  const [running, setRunning] = useState(false);
  const [engineStatus, setEngineStatus] = useState("");
  const [error, setError] = useState<string>("");
  const workerRef = useRef<Worker | null>(null);

  const validationErrors = useMemo(() => validateSettings(settings), [settings]);
  const missing = useMemo(() => missingRequiredInputs(inputs), [inputs]);

  async function assignInputFile(key: keyof PayrollInputFiles, file: File) {
    setInputs((prev) => {
      const next = { ...prev, [key]: file };
      if (key === "alternatePermissions" && next.preparedPermissions === undefined) {
        next.preparedPermissions = file;
      }
      return next;
    });
    setError("");
    if (key === "attendance") {
      try {
        const workbook = await readWorkbook(file);
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
    setEngineStatus("Stopped.");
  }

  async function runPayroll() {
    setError("");
    setResult(null);
    setMetrics([]);
    setLogs([]);
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

  const tabs: Array<{ id: AppTab; label: string }> = [
    { id: "inputs", label: "Inputs" },
    { id: "settings", label: "Settings" },
    { id: "log", label: "Run Log" },
    { id: "outputs", label: "Outputs" },
  ];

  return (
    <div className="app-layout">
      <header className="site-header">
        <div className="site-header-inner">
          <img className="site-logo" src={`${import.meta.env.BASE_URL}logo.svg`} alt="Nagwa" />
        </div>
      </header>

      <main className="app-shell">
        <header className="hero-panel">
          <div>
            <p className="eyebrow">100% browser-side payroll processing</p>
            <h1>PayrollTool</h1>
            <p className="lead">
              Run the Nagwa payroll pipeline in the browser. Files stay on this device; no backend, API key, or upload server is used.
            </p>
            <div className="meta-row hero-meta">
              <span>Permission mode: <strong>{permissionMode(inputs)}</strong></span>
              <span>Detected period: <strong>{detectedPeriod || "Upload Attendance Report"}</strong></span>
            </div>
          </div>
          <div className="hero-actions">
            <button className="primary" type="button" onClick={runPayroll} disabled={running || missing.length > 0 || validationErrors.length > 0}>
              {running ? "Running..." : "Run Payroll"}
            </button>
            <button type="button" onClick={stopWorker} disabled={!running}>
              Stop
            </button>
          </div>
        </header>

      {error && <div className="toast error">{error}</div>}
      {engineStatus && <div className="toast info">{engineStatus}</div>}

      <nav className="tabs" aria-label="Payroll dashboard sections">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            className={activeTab === tab.id ? "active" : ""}
            onClick={() => setActiveTab(tab.id)}
          >
            {tab.label}
          </button>
        ))}
      </nav>

      {activeTab === "inputs" && (
        <section className="grid two tab-panel">
          <div className="card">
            <h2>Input Files</h2>
            <p>Choose or drag each report into its own cell. The cell decides the file role, so the filename does not need to match exactly.</p>
            <div className="upload-grid">
              <UploadTile title="Attendance Report" description="Required attendance export" file={inputs.attendance} onFile={(file) => void assignInputFile("attendance", file)} />
              <UploadTile title="Absence Report" description="Required absence export" file={inputs.absences} onFile={(file) => void assignInputFile("absences", file)} />
              <UploadTile title="Employee Transactions_vacations" description="Required vacation/transaction export" file={inputs.vacations} onFile={(file) => void assignInputFile("vacations", file)} />
              <UploadTile title="Public Holiday" description="Required public holiday workbook" file={inputs.publicHoliday} onFile={(file) => void assignInputFile("publicHoliday", file)} />
              <UploadTile title="Resignations" description="Optional resignation workbook" file={inputs.resignations} optional onFile={(file) => void assignInputFile("resignations", file)} />
            </div>

            <div className="permission-input">
              <h3>Permission Request</h3>
              <p className="helper-text">
                Use either the raw request report or a prepared details report. If both are uploaded, the raw report is preferred and prepared again.
              </p>
              <div className="upload-grid two-tiles">
                <UploadTile title="Raw Permission Request Report" description="Nagwa_Permission_Request_Report export" file={inputs.rawPermissions} onFile={(file) => void assignInputFile("rawPermissions", file)} />
                <UploadTile title="Prepared Permission Details" description="Nagwa_Permission_Request_permission_details export" file={inputs.preparedPermissions} onFile={(file) => void assignInputFile("preparedPermissions", file)} />
                <UploadTile title="Permissions alternate" description="Optional permissions.xls fallback" file={inputs.alternatePermissions} optional onFile={(file) => void assignInputFile("alternatePermissions", file)} />
              </div>
              {permissionMode(inputs) === "raw" && (
                <div className="permission-options">
                  <p className="helper-text">
                    A raw permission report was uploaded, so the pipeline will prepare it first. These options control that preparation step.
                  </p>
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
                </div>
              )}
            </div>
          </div>

          <div className="card">
            <h2>Private Templates</h2>
            <p>Required for real runs. Drop each styled template into its own cell; templates stay in the browser.</p>
            <div className="upload-grid two-tiles">
              <UploadTile title="Nagwa Technologies template" description="Detailed workbook template" file={inputs.nagwaTemplate} accept=".xlsx" onFile={(file) => void assignInputFile("nagwaTemplate", file)} />
              <UploadTile title="Final Nagwa Technologies template" description="Final report template" file={inputs.finalTemplate} accept=".xlsx" onFile={(file) => void assignInputFile("finalTemplate", file)} />
            </div>
            {missing.length > 0 && (
              <div className="missing-list">
                <strong>Still needed:</strong>
                <ul>
                  {missing.map((item) => <li key={item}>{item}</li>)}
                </ul>
              </div>
            )}
          </div>
        </section>
      )}

      {activeTab === "settings" && (
        <section className="card tab-panel">
          <h2>Settings / Rules Panel</h2>
          <div className="settings-grid">
            <TextInput label="Ramadan start" type="date" value={settings.ramadanStart} onChange={(value) => updateSettings({ ...settings, ramadanStart: value as PayrollSettings["ramadanStart"] })} />
            <TextInput label="Ramadan end" type="date" value={settings.ramadanEnd} onChange={(value) => updateSettings({ ...settings, ramadanEnd: value as PayrollSettings["ramadanEnd"] })} />
            <TextInput label="Ramadan permitted end" value={settings.permittedWindowEndRamadan} onChange={(value) => updateSettings({ ...settings, permittedWindowEndRamadan: value })} />
            <TextInput label="Lunch switch date" type="date" value={settings.lunchWindowSwitchDate} onChange={(value) => updateSettings({ ...settings, lunchWindowSwitchDate: value as PayrollSettings["lunchWindowSwitchDate"] })} />
            <TextInput label="Lunch before start" value={settings.lunchWindowBefore.start} onChange={(value) => updateSettings({ ...settings, lunchWindowBefore: { ...settings.lunchWindowBefore, start: value } })} />
            <TextInput label="Lunch before end" value={settings.lunchWindowBefore.end} onChange={(value) => updateSettings({ ...settings, lunchWindowBefore: { ...settings.lunchWindowBefore, end: value } })} />
            <TextInput label="Lunch from start" value={settings.lunchWindowFrom.start} onChange={(value) => updateSettings({ ...settings, lunchWindowFrom: { ...settings.lunchWindowFrom, start: value } })} />
            <TextInput label="Lunch from end" value={settings.lunchWindowFrom.end} onChange={(value) => updateSettings({ ...settings, lunchWindowFrom: { ...settings.lunchWindowFrom, end: value } })} />
          </div>

          <div className="rule-columns">
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
                    Swap A/B
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
          </div>
          <div className="settings-reset-actions">
            <div className="reset-action-card">
              <button type="button" onClick={resetAllSettings}>Restore default settings</button>
              <p>Use this to return all payroll rules to the standard values shipped with the tool.</p>
            </div>
          </div>
          {validationErrors.length > 0 && (
            <ul className="validation">
              {validationErrors.map((item) => <li key={item}>{item}</li>)}
            </ul>
          )}
        </section>
      )}

      {activeTab === "log" && (
        <section className="card tab-panel">
          <h2>Run Log</h2>
          <ol className="steps">
            {["extend_nagwa_technologies", "fill_attendance", "extend_final_nagwa_technologies", "complete_final"].map((step, index) => (
              <li key={step} className={logs.some((entry) => entry.step === step) ? "active" : ""}>{index + 1}/4 {stepLabel(step)}</li>
            ))}
          </ol>
          <div className="log-panel">
            {logs.length === 0 ? <p>No run yet.</p> : logs.map((entry, index) => <pre key={`${entry.at}-${index}`} className={entry.level}>{formatLog(entry)}</pre>)}
          </div>
        </section>
      )}

      {activeTab === "outputs" && (
        <section className="card tab-panel">
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
        </section>
      )}
      </main>

      <footer className="site-footer">
        <p>Copyright © 2026 Nagwa</p>
        <p>All Rights Reserved</p>
      </footer>
    </div>
  );
}

export default App;
