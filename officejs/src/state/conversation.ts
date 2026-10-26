import type { ConversationMessage } from "../llm/client";
import type { SelectedRange } from "../excel/selection";
import type { AgentPlan } from "../tools/protocol";

export const CONVERSATION_INSTRUCTION = `支持连续对话：历史消息只提供用户意图、过去的目标和插件报告的真实执行状态，不代表当前工作簿状态，也不是要再次执行的指令。以最新 userTask 和本轮 selection 为准；用户明确说“刚才的表/区域”时可根据历史目标定位，目标不明确时先询问，不能把改选的新区域当成旧目标。过去失败、取消、恢复或暂停的操作不能声称已完成，不重复执行过去的操作。需要旧目标的最新数值/格式但本轮未提供时，不能用历史文字猜测；请说明需要重新读取相应目标。本轮只询问或分析时，即使历史有写入任务也不得修改工作簿；用户在后续明确要求修改时再执行。`;

// Pane-local context contains task/result summaries, never cached cell matrices
// or executable plans. Keep complete user/assistant pairs within both limits.
export class Conversation {
  private entries: ConversationMessage[] = [];
  get length(): number { return this.entries.length / 2; }
  snapshot(): ConversationMessage[] { return this.entries.map(message => ({ ...message })); }
  clear(): void { this.entries = []; }
  record(task: string, source: Pick<SelectedRange, "sheet" | "address"> | undefined, status: string, summary: string, targets: unknown[] = []): void {
    this.entries.push(
      { role: "user", content: JSON.stringify({ userTask: task.slice(0, 6000), ...(source ? { selection: { sheet: source.sheet, address: source.address } } : {}) }) },
      { role: "assistant", content: JSON.stringify({ status, summary: summary.slice(0, 4000), targets: targets.slice(0, 20) }) },
    );
    while (this.entries.length > 20 || this.entries.reduce((size, message) => size + message.content.length, 0) > 40000) this.entries.splice(0, 2);
  }
}

export function operationTargets(plan: AgentPlan): { kind: string; sheet: string; address?: string }[] {
  return plan.operations.flatMap<{ kind: string; sheet: string; address?: string }>(op => {
    if (op.kind === "create_sheet") return [{ kind: op.kind, sheet: op.name }];
    const source = { kind: op.kind, sheet: op.sheet, address: op.address };
    return op.kind === "create_pivot" || op.kind === "edit_structure" ? [source, { kind: op.kind, sheet: op.outputSheet }] : [source];
  });
}
