import type { UndoSnapshot } from "../excel/recovery";
import { parseAddress } from "../tools/protocol";

export interface StepDefinition { title: string; instruction: string; sheet: string; address: string; expected: string; }
export interface WorkflowStep extends StepDefinition {
  status: "pending" | "running" | "done" | "failed" | "restoring";
  result?: string; error?: string; correction?: string; undo?: UndoSnapshot;
}
export interface RangeRef { sheet: string; address: string; }
export interface Workflow {
  version: 1; id: string; task: string; summary: string; revision?: number;
  anchor: { sheet: string; sheetId: string; address: string };
  status: "planned" | "running" | "paused" | "completed" | "recovery";
  steps: WorkflowStep[]; guard?: string; watched: RangeRef[];
}
export function parseSteps(value: unknown): { summary: string; steps: StepDefinition[] } {
  const object = value as { summary?: unknown; steps?: unknown };
  if (!object || typeof object.summary !== "string" || object.summary.length > 1500 || !Array.isArray(object.steps) || !object.steps.length || object.steps.length > 8 || Object.keys(object).some(key => !["summary", "steps"].includes(key))) throw new Error("任务计划应包含 1–8 个步骤及简短说明");
  const steps = object.steps.map((item: StepDefinition) => {
    if (!item || Object.keys(item).some(key => !["title", "instruction", "sheet", "address", "expected"].includes(key))) throw new Error("步骤包含不支持的参数");
    for (const [key, max] of [["title", 100], ["instruction", 3000], ["sheet", 31], ["address", 40], ["expected", 1000]] as const) {
      if (typeof item[key] !== "string" || !item[key].trim() || item[key].length > max) throw new Error("任务步骤缺少有效的标题、要求、工作表、区域或检查标准");
    }
    if (/^_AI_backup_/i.test(item.sheet)) throw new Error("不能以恢复备份作为任务目标");
    const bounds = parseAddress(item.address);
    if (bounds.rows * bounds.cols > 100000) throw new Error("单个步骤上下文最多 100000 格");
    return { title: item.title, instruction: item.instruction, sheet: item.sheet, address: item.address, expected: item.expected };
  });
  return { summary: object.summary, steps };
}
export interface WorkflowIO {
  save: (task: Workflow) => Promise<void>;
  guard: (task: Workflow) => Promise<string>;
  execute: (task: Workflow, index: number, signal: AbortSignal, checkpoint: (state: UndoSnapshot) => Promise<void>) => Promise<{ result: string; passed: boolean; watched: RangeRef[] }>;
  restore: (state: UndoSnapshot) => Promise<void>;
  dispose: (state: UndoSnapshot) => Promise<void>;
  changed?: () => void;
}
export class WorkflowRunner {
  private active = false;
  constructor(public task: Workflow, private io: WorkflowIO) {}
  private async save(): Promise<void> { await this.io.save(this.task); this.io.changed?.(); }
  private async capture(): Promise<void> { this.task.guard = await this.io.guard(this.task); }
  private async unchanged(): Promise<void> {
    if (this.task.guard && this.task.guard !== await this.io.guard(this.task)) throw new Error("任务涉及的工作表内容、格式或对象已改变；为保留手动修改，已停止继续或重做。请结束此任务后重新规划。");
  }
  async initialize(): Promise<void> { await this.capture(); await this.save(); }
  async run(signal: AbortSignal): Promise<void> {
    if (this.active) return;
    this.active = true;
    try {
      if (this.task.status !== "recovery") await this.unchanged();
      let index = this.task.steps.findIndex(step => step.status !== "done");
      if (index < 0) return;
      if (this.task.steps[index].undo) await this.rewindInternal(index);
      this.task.status = "running"; await this.save();
      for (; index < this.task.steps.length; index++) {
        if (signal.aborted) break;
        await this.unchanged();
        const step = this.task.steps[index];
        step.status = "running"; step.error = undefined; await this.save();
        try {
          const result = await this.io.execute(this.task, index, signal, async snapshot => { step.undo = snapshot; await this.save(); });
          step.result = result.result;
          for (const ref of result.watched) if (!this.task.watched.some(item => item.sheet === ref.sheet && item.address === ref.address)) this.task.watched.push(ref);
          step.status = result.passed ? "done" : "failed";
          if (!result.passed) step.error = "结果检查未通过，可补充要求后从此步重做。";
          await this.capture(); await this.save();
          if (!result.passed) break;
        } catch (error) {
          step.status = "failed"; step.error = error instanceof Error ? error.message : "步骤失败";
          this.task.status = error instanceof Error && error.name === "RecoveryError" ? "recovery" : "paused";
          // Keep prepared backups even if execution already rolled back. Reapplying
          // restoration is safe, and avoids a dangling journal after a window closes.
          if (!(error instanceof Error && error.name === "WorkbookChangedError")) {
            try { await this.capture(); } catch { this.task.status = "recovery"; }
          }
          await this.save(); break;
        }
      }
      if (this.task.status !== "recovery") this.task.status = this.task.steps.every(step => step.status === "done") ? "completed" : "paused";
      await this.save();
    } finally { this.active = false; this.io.changed?.(); }
  }
  private async rewindInternal(from: number): Promise<void> {
    // Conservatively invalidate all later steps: none can continue with stale inputs.
    this.task.status = "recovery";
    for (let index = this.task.steps.length - 1; index >= from; index--) {
      const step = this.task.steps[index], undo = step.undo;
      if (undo) {
        step.status = "restoring"; await this.save();
        await this.io.restore(undo);
        step.undo = undefined;
      }
      step.status = "pending"; step.result = step.error = undefined;
      // Save before removing backups; a failed save must leave them recoverable.
      await this.save();
      if (undo) await this.io.dispose(undo).catch(() => undefined);
    }
    // Watches may include outputs deleted by rollback. Rebuild from surviving steps.
    this.task.watched = [{ sheet: this.task.anchor.sheet, address: this.task.anchor.address },
      ...this.task.steps.slice(0, from).flatMap(step => step.undo?.backups.map(item => ({ sheet: item.sheetId ?? item.sheet, address: item.address })) ?? [])];
    await this.capture(); this.task.status = "paused"; await this.save();
  }
  async rewind(from: number, correction = ""): Promise<void> {
    if (this.active || !Number.isInteger(from) || from < 0 || from >= this.task.steps.length || correction.length > 3000) throw new Error("重做步骤无效或任务仍在执行");
    this.active = true;
    try {
      if (this.task.status !== "recovery") await this.unchanged();
      this.task.steps[from].correction = correction.trim() || this.task.steps[from].correction;
      await this.rewindInternal(from);
    } finally { this.active = false; this.io.changed?.(); }
  }
}
