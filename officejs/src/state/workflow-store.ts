import { parseSteps, Workflow } from "./workflow";
import { validateSnapshot } from "../excel/recovery";
import { parseAddress } from "../tools/protocol";

const KEY = "AIForExcel.Workflow.v1";
export function decodeWorkflow(raw: string): Workflow {
  if (typeof raw !== "string" || raw.length > 1800000) throw new Error("任务进度过大或格式无效");
  const task = JSON.parse(raw) as Workflow;
  if (task?.version !== 1 || typeof task.id !== "string" || !/^[a-f0-9-]{36}$/i.test(task.id) || typeof task.task !== "string" || task.task.length > 12000 ||
      !["planned", "running", "paused", "completed", "recovery"].includes(task.status) || !Array.isArray(task.steps) || !task.anchor || typeof task.anchor.sheetId !== "string" || typeof task.anchor.sheet !== "string" || task.revision !== undefined && (!Number.isSafeInteger(task.revision) || task.revision < 0)) throw new Error("任务进度无效，未执行任何步骤");
  parseAddress(task.anchor.address);
  parseSteps({ summary: task.summary, steps: task.steps.map(({ title, instruction, sheet, address, expected }) => ({ title, instruction, sheet, address, expected })) });
  if (!Array.isArray(task.watched) || task.watched.length > 1000 || task.guard !== undefined && !/^[a-f0-9]{64}$/.test(task.guard)) throw new Error("任务检查记录无效");
  for (const ref of task.watched) { if (typeof ref.sheet !== "string" || /^_AI_backup_/i.test(ref.sheet)) throw new Error("任务检查目标无效"); parseAddress(ref.address); }
  let unfinished = false;
  for (const step of task.steps) {
    if (!["pending", "running", "done", "failed", "restoring"].includes(step.status) || [step.result, step.error, step.correction].some(value => value !== undefined && (typeof value !== "string" || value.length > 12000))) throw new Error("任务步骤状态无效");
    if (step.status !== "done") unfinished = true;
    else if (unfinished && task.status !== "recovery") throw new Error("任务完成顺序无效");
    if (step.undo) validateSnapshot(step.undo);
  }
  return task;
}
export async function saveWorkflow(task: Workflow): Promise<void> {
  await assertWorkflowRevision(task);
  const revision = (task.revision ?? 0) + 1;
  const raw = JSON.stringify({ ...task, revision });
  decodeWorkflow(raw);
  Office.context.document.settings.set(KEY, raw);
  await new Promise<void>((resolve, reject) => Office.context.document.settings.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else reject(new Error("无法保存任务进度，已停止后续操作；请保留窗格及备份，检查工作簿是否可编辑。"));
  }));
  task.revision = revision;
}
export async function assertWorkflowRevision(task: Workflow): Promise<void> {
  await new Promise<void>((resolve, reject) => Office.context.document.settings.refreshAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve(); else reject(new Error("无法核对工作簿中的最新任务进度，已停止操作"));
  }));
  const raw = Office.context.document.settings.get(KEY);
  const stored = raw === null || raw === undefined ? undefined : decodeWorkflow(raw);
  if (stored ? stored.id !== task.id || (stored.revision ?? 0) !== (task.revision ?? 0) : (task.revision ?? 0) !== 0) throw new Error("另一个窗格已更新任务进度，请重新打开此窗格后继续");
}
export function loadWorkflow(): Workflow | undefined {
  const raw = Office.context.document.settings.get(KEY);
  if (raw === undefined || raw === null) return;
  const task = decodeWorkflow(raw);
  if (task.status === "running" || task.steps.some(step => step.status === "running" || step.status === "restoring")) task.status = "recovery";
  return task;
}
export async function clearWorkflow(): Promise<void> {
  Office.context.document.settings.remove(KEY);
  await new Promise<void>((resolve, reject) => Office.context.document.settings.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve(); else reject(new Error("任务记录未能清除，备份仍保留"));
  }));
}
