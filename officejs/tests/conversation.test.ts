import { test, afterEach } from "node:test";
import assert from "node:assert/strict";
import { Conversation, operationTargets } from "../src/state/conversation";
import { askAgent, askAnalysis } from "../src/llm/client";
import { planWorkflow } from "../src/llm/workflow";
import { parseAddress, parsePlan } from "../src/tools/protocol";

const originalFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = originalFetch; });
const config = { baseUrl: "https://example.com/v1", model: "test-model", apiKey: "test-only" };
const selected = { workbook: "test", sheet: "数据", sheetId: "id", address: "B2", bounds: parseAddress("B2"), values: [[9]], formulas: [[9]], dataMode: "full" as const };
const response = (value: unknown) => new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(value) } }] }));

test("conversation stores target references without old cell matrices or executable plans", () => {
  const conversation = new Conversation();
  const plan = parsePlan(JSON.stringify({ answer: "已写入", operations: [
    { kind: "write_values", sheet: "数据", address: "A1", data: [["old-cell-content"]] },
    { kind: "create_sheet", name: "汇总", data: [["old-output-content"]] },
    { kind: "create_pivot", sheet: "数据", address: "A1:B3", outputSheet: "透视", rows: ["类别"], columns: [], values: [{ field: "金额", aggregation: "Sum" }] },
  ] }));
  conversation.record("写入数据", selected, "applied", plan.answer, operationTargets(plan));
  const history = conversation.snapshot();
  assert.deepEqual(JSON.parse(history[0].content).selection, { sheet: "数据", address: "B2" });
  assert.deepEqual(JSON.parse(history[1].content).targets.map((ref: any) => ref.sheet), ["数据", "汇总", "数据", "透视"]);
  const structure = parsePlan(JSON.stringify({ answer: "副本插行", operations: [{ kind: "edit_structure", sheet: "数据", address: "A2", outputSheet: "副本", action: "insert_rows" }] }));
  assert.deepEqual(operationTargets(structure).map(ref => ref.sheet), ["数据", "副本"]);
  assert.doesNotMatch(JSON.stringify(history), /old-cell-content|old-output-content|formulas|operations/);
  history[0].content = "mutated";
  history.pop();
  assert.equal(conversation.snapshot().length, 2);
  assert.notEqual(conversation.snapshot()[0].content, "mutated");
});

test("history limits retain whole pairs and clearing starts an empty conversation", () => {
  const conversation = new Conversation();
  for (let i = 0; i < 12; i++) conversation.record(`task${i}`, selected, "applied", "完成");
  assert.equal(conversation.length, 10);
  assert.equal(JSON.parse(conversation.snapshot()[0].content).userTask, "task2");
  for (let i = 0; i < 12; i++) conversation.record("x".repeat(7000), selected, "analysis_answered", "y".repeat(5000));
  const history = conversation.snapshot();
  assert.ok(history.reduce((size, item) => size + item.content.length, 0) <= 40000);
  assert.ok(history.length > 0);
  history.forEach((item, i) => assert.equal(item.role, i % 2 ? "assistant" : "user"));
  conversation.clear();
  assert.equal(conversation.length, 0);
  assert.deepEqual(conversation.snapshot(), []);
});

test("table follow-up sends previous outcome and fresh selection, then clear removes that context", async () => {
  const conversation = new Conversation();
  conversation.record("将 A1 加粗", { sheet: "数据", address: "A1" }, "applied", "已加粗", [{ kind: "format_range", sheet: "数据", address: "A1" }]);
  let calls = 0;
  globalThis.fetch = (async (_url, init) => {
    const messages = JSON.parse(String(init?.body)).messages;
    assert.match(messages[0].content, /不重复执行过去的操作/);
    if (++calls === 1) {
      assert.equal(messages.length, 4);
      assert.equal(JSON.parse(messages[1].content).userTask, "将 A1 加粗");
      assert.equal(JSON.parse(messages[2].content).status, "applied");
      assert.equal(JSON.parse(messages[2].content).targets[0].address, "A1");
    } else assert.equal(messages.length, 2);
    const current = JSON.parse(messages.at(-1).content);
    assert.equal(current.userTask, "刚才的区域再居中");
    assert.equal(current.selection.address, "B2");
    assert.deepEqual(current.selection.values, [[9]]);
    return response({ answer: "请确认目标", operations: [] });
  }) as typeof fetch;
  await askAgent("刚才的区域再居中", selected, config, undefined, undefined, undefined, conversation.snapshot());
  conversation.clear();
  await askAgent("刚才的区域再居中", selected, config, undefined, undefined, undefined, conversation.snapshot());
  assert.equal(calls, 2);
});

test("shared context preserves failure and restoration statuses while analysis still rejects writes", async () => {
  const conversation = new Conversation();
  for (const status of ["stopped_not_completed", "recovery_required", "restored", "restore_incomplete", "workflow_paused"])
    conversation.record("之前的任务", selected, status, status);
  globalThis.fetch = (async (_url, init) => {
    const messages = JSON.parse(String(init?.body)).messages;
    assert.deepEqual(messages.slice(1, -1).filter((m: any) => m.role === "assistant").map((m: any) => JSON.parse(m.content).status),
      ["stopped_not_completed", "recovery_required", "restored", "restore_incomplete", "workflow_paused"]);
    return response({ answer: "执行", operations: [{ kind: "clear_contents", sheet: "数据", address: "B2" }] });
  }) as typeof fetch;
  await assert.rejects(askAnalysis("继续", selected, config, conversation.snapshot()), /不会修改工作簿/);
});

test("automatic workflow planning receives shared conversation for a dependent follow-up", async () => {
  const conversation = new Conversation();
  conversation.record("按部门汇总", selected, "applied", "已创建汇总表", [{ sheet: "汇总" }]);
  globalThis.fetch = (async (_url, init) => {
    const messages = JSON.parse(String(init?.body)).messages;
    assert.deepEqual(messages.slice(1, -1), conversation.snapshot());
    assert.equal(JSON.parse(messages.at(-1).content).task, "把刚才的汇总再计算占比后绘图");
    return response({ summary: "计算并绘图", steps: [{ title: "计算占比", instruction: "计算占比", sheet: "汇总", address: "A1:B3", expected: "占比合计为1" }] });
  }) as typeof fetch;
  assert.equal((await planWorkflow("把刚才的汇总再计算占比后绘图", selected, config, undefined, undefined, conversation.snapshot())).steps[0].sheet, "汇总");
});
