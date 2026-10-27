import { askAgent, Connection, ConversationMessage, executionPlan } from "../llm/client";
import { checkStep, planWorkflow, stepContext } from "../llm/workflow";
import { applyOperations } from "../excel/operations";
import { readSelection, SelectedRange } from "../excel/selection";
import { recoveryState } from "../excel/recovery";
import { inspectStep, WorkbookChangedError, workflowGuard } from "../excel/workflow-observation";
import { Workflow, WorkflowRunner } from "../state/workflow";
import { assertWorkflowRevision, clearWorkflow, loadWorkflow, saveWorkflow } from "../state/workflow-store";
import { getLastUndo, setLastUndo } from "../state/undo";

interface Hooks {
  connection: () => Connection;
  action: (run: (signal: AbortSignal) => Promise<void>, abortable?: boolean) => Promise<void>;
  append: (label: string, text: string) => void;
  stage: (text: string, model?: boolean) => void;
  changed: () => void;
}
const labels = { pending: "待执行", running: "执行中", done: "已完成", failed: "待修正", restoring: "恢复中" };
export class WorkflowController {
  runner?: WorkflowRunner;
  private panel = document.getElementById("workflowPanel")!;
  constructor(private hooks: Hooks) {}
  load(): void {
    const task = loadWorkflow();
    if (task) { this.attach(task); this.hooks.append("分步任务", "已读取工作簿中的任务进度。继续前会核对当前表格；中断步骤需先恢复。保存工作簿才能将进度和备份保留到磁盘。"); }
  }
  private attach(task: Workflow): void {
    this.runner = new WorkflowRunner(task, {
      save: saveWorkflow, guard: workflowGuard, changed: () => this.hooks.changed(),
      restore: state => recoveryState(state).restore(), dispose: state => recoveryState(state).dispose(),
      execute: async (current, index, signal, checkpoint) => {
        const step = current.steps[index], connection = this.hooks.connection();
        const progress: import("../llm/client").Progress = (phase, message) => {
          if (phase === "feedback") this.hooks.stage(`第 ${index + 1} 步 · 模型：${message}`, true);
        };
        this.hooks.stage(`第 ${index + 1} 步：读取 ${step.sheet}!${step.address}`);
        let selected = await readSelection("metadata", { sheetId: step.sheet, address: step.address });
        if (!current.watched.some(ref => ref.sheet === selected.sheetId && ref.address === selected.address)) current.watched.push({ sheet: selected.sheetId, address: selected.address });
        current.guard = await workflowGuard(current); await saveWorkflow(current);
        const context = stepContext(current, index);
        const instruction = `${step.instruction}${step.correction ? `\n用户对本步骤的修正：${step.correction}` : ""}`;
        this.hooks.stage(`第 ${index + 1} 步：模型生成操作…`, true);
        let turn = await askAgent(instruction, selected, connection, signal, context, progress);
        for (let attempt = 0; attempt < 3 && (turn.needsData || turn.needsFormat); attempt++) {
          if ((!turn.needsData || selected.dataMode === "full") && (!turn.needsFormat || selected.formatMode === "full")) break;
          if (turn.needsData) {
            selected = await readSelection("full", selected, turn.needsFormat || selected.formatMode === "full");
          } else if (turn.needsFormat) {
            const previous = selected;
            const formatted = await readSelection("format", selected);
            selected = previous.dataMode === "full" ? { ...formatted, values: previous.values, formulas: previous.formulas, dataMode: "full" as const } : formatted;
          }
          turn = await askAgent(instruction, selected, connection, signal, context, progress);
        }
        if (turn.needsData || turn.needsFormat) throw new Error("模型仍要求补充已有数据，本步骤未执行");
        if (signal.aborted) throw new Error("已暂停；本步骤未执行");
        if (current.guard !== await workflowGuard(current)) throw new WorkbookChangedError();
        const plan = executionPlan(turn);
        this.hooks.stage(`第 ${index + 1} 步：执行并保留备份…`);
        const warning = await applyOperations(plan, selected, { retainUndo: true, checkpoint, signal,
          onProgress: progress => this.hooks.stage(`第 ${index + 1} 步：分批处理 ${progress.completed}/${progress.total} 格… 可暂停`),
          onRecovery: snapshot => { step.undo = snapshot; },
        });
        if (warning) this.hooks.append("提示", warning);
        const observed = await inspectStep(plan, { sheet: selected.sheetId, address: selected.address }, step.undo);
        for (const ref of observed.watched) if (!current.watched.some(item => item.sheet === ref.sheet && item.address === ref.address)) current.watched.push(ref);
        current.guard = await workflowGuard(current); await saveWorkflow(current);
        this.hooks.stage(`第 ${index + 1} 步：检查实际结果…`, true);
        // A committed but unchecked step keeps its backup. If paused here, the
        // next run restores that step before retrying rather than duplicating it.
        const checked = observed.localErrors.length ? { passed: false, summary: observed.localErrors.join("；") } :
          await checkStep(current, index, { readback: observed.evidence, executionAnswer: plan.answer }, connection, signal, progress);
        if (current.guard !== await workflowGuard(current)) throw new WorkbookChangedError();
        this.hooks.append(`第 ${index + 1} 步 · ${checked.passed ? "检查通过" : "需要修正"}`, checked.summary);
        return { result: checked.summary, passed: checked.passed, watched: observed.watched };
      },
    });
    this.hooks.changed();
  }
  async plan(text: string, signal: AbortSignal, source?: SelectedRange, history: ConversationMessage[] = []): Promise<void> {
    if (this.runner) throw new Error("请先结束当前分步任务，再创建新任务");
    if (text.length > 12000) throw new Error("任务要求过长，请精简后重试");
    const selected = await readSelection("metadata", source);
    const task: Workflow = { version: 1, id: crypto.randomUUID(), task: text, summary: "", anchor: { sheet: selected.sheet, sheetId: selected.sheetId, address: selected.address }, status: "planned", steps: [], watched: [] };
    const guard = await workflowGuard(task);
    this.hooks.stage("正在拆解任务…", true);
    const outline = await planWorkflow(text, selected, this.hooks.connection(), signal, (phase, message) => {
      if (phase === "feedback") this.hooks.stage(`任务规划 · 模型：${message}`, true);
    }, history);
    if (signal.aborted) throw new Error("已取消任务规划，未修改工作簿");
    if (guard !== await workflowGuard(task)) throw new WorkbookChangedError();
    task.summary = outline.summary; task.steps = outline.steps.map(step => ({ ...step, status: "pending" })); task.guard = guard;
    await saveWorkflow(task); this.attach(task);
    const old = getLastUndo(); setLastUndo(undefined); await old?.dispose().catch(() => undefined);
    this.hooks.append("任务计划", `${outline.summary}\n${outline.steps.map((step, index) => `${index + 1}. ${step.title}`).join("\n")}`);
    await this.run(signal);
  }
  private async locked(action: () => Promise<void>): Promise<void> {
    if (!this.runner) return;
    if (!navigator.locks) throw new Error("当前 Excel 浏览器组件无法锁定分步任务，请更新 WebView2 后重试");
    await navigator.locks.request(`AIForExcel.Workflow.${this.runner.task.id}`, { ifAvailable: true }, async lock => {
      if (!lock) throw new Error("另一个窗格正在执行此任务，请先在原窗格暂停");
      await assertWorkflowRevision(this.runner!.task);
      await action();
    });
  }
  async run(signal: AbortSignal): Promise<void> { await this.locked(async () => this.runner!.run(signal)); }
  render(busy: boolean, visible: boolean): void {
    this.panel.hidden = !this.runner || !visible;
    this.panel.replaceChildren();
    if (!this.runner) return;
    const task = this.runner.task;
    const title = document.createElement("strong"); title.textContent = task.status === "completed" ? "分步任务已完成" : task.status === "recovery" ? "任务中断，需先恢复当前步骤" : "分步任务";
    const summary = document.createElement("p"); summary.textContent = task.summary;
    const list = document.createElement("ol");
    task.steps.forEach((step, index) => {
      const row = document.createElement("li"), text = document.createElement("div");
      text.textContent = `${step.title} · ${labels[step.status]}`;
      const detail = document.createElement("small"); detail.textContent = step.error ?? step.result ?? `${step.sheet}!${step.address} · ${step.expected}`;
      row.append(text, detail);
      if (step.status !== "pending") {
        const retry = document.createElement("button"); retry.className = "secondary"; retry.textContent = "从此步重做"; retry.disabled = busy;
        retry.onclick = () => { const correction = (this.panel.querySelector("textarea") as HTMLTextAreaElement).value;
          void this.hooks.action(async signal => this.locked(async () => { await this.runner!.rewind(index, correction); await this.runner!.run(signal); })); };
        row.append(retry);
      }
      list.append(row);
    });
    const correction = document.createElement("textarea"); correction.placeholder = "重做时的补充要求（可选）：例如保留前两步，第3步改为按部门汇总"; correction.maxLength = 3000; correction.disabled = busy; correction.setAttribute("aria-label", "分步重做补充要求");
    // Preserve an in-progress correction when unrelated UI state changes.
    correction.value = this.correction;
    correction.oninput = () => { this.correction = correction.value; };
    const actions = document.createElement("div"); actions.className = "actions";
    const resume = document.createElement("button"); resume.textContent = task.status === "recovery" ? "恢复中断步骤并继续" : "继续 / 重试"; resume.disabled = busy || task.status === "completed";
    resume.onclick = () => void this.hooks.action(async signal => this.locked(async () => {
      const index = task.steps.findIndex(step => step.status !== "done");
      if (this.correction.trim() && index >= 0) await this.runner!.rewind(index, this.correction);
      await this.runner!.run(signal);
    }));
    const rollback = document.createElement("button"); rollback.textContent = "撤销整个任务"; rollback.className = "secondary"; rollback.disabled = busy || !task.steps.some(step => step.undo);
    rollback.onclick = () => void this.hooks.action(async () => this.locked(async () => { await this.runner!.rewind(0); this.hooks.append("已恢复", "已逆序恢复本任务各步骤，尚未重新执行。"); }), false);
    const finish = document.createElement("button"); finish.textContent = "保留结果并结束"; finish.className = "secondary"; finish.disabled = busy || task.status === "recovery" || task.steps.some(step => step.status === "restoring");
    finish.onclick = () => void this.hooks.action(async () => this.locked(async () => {
      const snapshots = task.steps.flatMap(step => step.undo ? [step.undo] : []);
      await clearWorkflow(); this.runner = undefined; this.correction = "";
      let failed = false;
      for (const snapshot of snapshots) await recoveryState(snapshot).dispose().catch(() => { failed = true; });
      this.hooks.append("任务已结束", `保留当前结果，已释放分步恢复记录。${failed ? "部分隐藏备份未能清理。" : ""}`);
    }), false);
    const note = document.createElement("small"); note.textContent = "进度保存在当前工作簿，需保存工作簿以便下次继续。重做会撤销所选步骤及后续步骤；结束后释放本任务恢复记录。";
    actions.append(resume, rollback, finish); this.panel.append(title, summary, list, correction, actions, note);
  }
  private correction = "";
}
