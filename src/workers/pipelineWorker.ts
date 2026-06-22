import { runPayrollPipeline } from "../pipeline/payrollPipeline";
import type { PayrollInputFiles, PayrollSettings, PermissionPrepOptions, RunLogEntry } from "../core/types";

export interface PipelineWorkerRequest {
  type: "run";
  inputs: PayrollInputFiles;
  settings: PayrollSettings;
  permissionOptions: PermissionPrepOptions;
}

type WorkerResponse =
  | { type: "log"; entry: RunLogEntry }
  | { type: "done"; result: Awaited<ReturnType<typeof runPayrollPipeline>> }
  | { type: "error"; message: string };

const post = (message: WorkerResponse) => {
  self.postMessage(message);
};

self.onmessage = async (event: MessageEvent<PipelineWorkerRequest>) => {
  if (event.data.type !== "run") return;
  try {
    const result = await runPayrollPipeline(event.data.inputs, event.data.settings, event.data.permissionOptions, (entry) => {
      post({ type: "log", entry });
    });
    post({ type: "done", result });
  } catch (error) {
    post({ type: "error", message: error instanceof Error ? error.message : String(error) });
  }
};
