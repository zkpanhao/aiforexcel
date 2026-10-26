import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { parseSteps, Workflow, WorkflowIO, WorkflowRunner } from "../src/state/workflow";
import { assertWorkflowRevision, decodeWorkflow, loadWorkflow, saveWorkflow } from "../src/state/workflow-store";
import { UndoSnapshot } from "../src/excel/recovery";
import { planWorkflow, checkStep } from "../src/llm/workflow";
import { askAgent } from "../src/llm/client";
const token = "11111111-1111-4111-8111-111111111111";
function task(): Workflow {
  return { version: 1, id: token, task: "计算后汇总并绘图", summary: "四步任务", anchor: { sheet: "数据", sheetId: "source", address: "A1" }, status: "planned", watched: [],
    steps: Array.from({ length: 4 }, (_, i) => ({ title: `步骤${i + 1}`, instruction: `完成第${i + 1}步`, sheet: "数据", address: "A1", expected: `结果为${i + 1}`, status: "pending" })) };
}
function harness() {
  const state = task(), values = [0, 0, 0, 0], calls: number[] = [], restores: number[] = [], disposed: number[] = [], saved: string[] = [];
  let failAt = -1, restoreFails = false, saveFails = false, pauseAt = -1;
  const controller = new AbortController();
  const io: WorkflowIO = {
    save: async value => { if (saveFails) throw new Error("save failed"); saved.push(JSON.stringify(value)); },
    guard: async () => createHash("sha256").update(JSON.stringify(values)).digest("hex"),
    execute: async (_task, index, _signal, checkpoint) => {
      calls.push(index);
      const snapshot: UndoSnapshot = { version: 1, token, sourceSheetId: String(index), backups: [], created: [] };
      await checkpoint(snapshot);
      assert.ok(JSON.parse(saved.at(-1)!).steps[index].undo, "journal must precede writes");
      values[index] = 1;
      if (pauseAt === index) controller.abort();
      return { passed: failAt !== index, result: `实际结果${index}`, watched: [] };
    },
    restore: async snapshot => { const index = Number(snapshot.sourceSheetId); restores.push(index); if (restoreFails) throw new Error("protected"); values[index] = 0; },
    dispose: async snapshot => { disposed.push(Number(snapshot.sourceSheetId)); },
  };
  const runner = new WorkflowRunner(state, io);
  return { state, values, calls, restores, disposed, saved, io, runner, controller, fail: (index: number) => { failAt = index; }, pause: (index: number) => { pauseAt = index; }, failRestore: (v: boolean) => { restoreFails = v; }, failSave: (v: boolean) => { saveFails = v; } };
}
test("workflow outline is bounded and rejects executable or malformed step payloads", () => {
  const outline = { summary: "计划", steps: task().steps.map(({ status, ...step }) => step) };
  assert.equal(parseSteps(outline).steps.length, 4);
  for (const value of [null, { ...outline, steps: [] }, { ...outline, steps: Array(9).fill(outline.steps[0]) }, { ...outline, operations: [] },
    { ...outline, steps: [{ ...outline.steps[0], script: "run()" }] }, { ...outline, steps: [{ ...outline.steps[0], sheet: "_AI_backup_x" }] }, { ...outline, steps: [{ ...outline.steps[0], address: "A:A" }] }]) assert.throws(() => parseSteps(value));
});
test("failure in step 3 preserves steps 1 and 2; continuation restores and retries only step 3 onward", async () => {
  const h = harness(); await h.runner.initialize(); h.fail(2);
  await h.runner.run(h.controller.signal);
  assert.deepEqual(h.calls, [0, 1, 2]); assert.equal(h.state.status, "paused");
  assert.deepEqual(h.state.steps.map(step => step.status), ["done", "done", "failed", "pending"]);
  h.fail(-1); await h.runner.run(new AbortController().signal);
  assert.deepEqual(h.calls, [0, 1, 2, 2, 3]); assert.deepEqual(h.restores, [2]);
  assert.deepEqual(h.values, [1, 1, 1, 1]); assert.equal(h.state.status, "completed");
});
test("redo reverses dependent steps in reverse order and keeps earlier results", async () => {
  const h = harness(); await h.runner.initialize(); await h.runner.run(h.controller.signal);
  await h.runner.rewind(2, "第三步改为按部门汇总");
  assert.deepEqual(h.restores, [3, 2]); assert.deepEqual(h.values, [1, 1, 0, 0]);
  assert.equal(h.state.steps[2].correction, "第三步改为按部门汇总");
  await h.runner.run(new AbortController().signal);
  assert.deepEqual(h.calls, [0, 1, 2, 3, 2, 3]);
  await h.runner.rewind(0); assert.deepEqual(h.values, [0, 0, 0, 0]);
});
test("pause retains a committed checked step and resumes at the next step", async () => {
  const h = harness(); h.pause(0); await h.runner.initialize(); await h.runner.run(h.controller.signal);
  assert.deepEqual(h.calls, [0]); assert.equal(h.state.steps[0].status, "done");
  await h.runner.run(new AbortController().signal); assert.deepEqual(h.calls, [0, 1, 2, 3]);
});
test("manual changes block continuation and rollback without touching successful steps", async () => {
  const h = harness(); h.pause(0); await h.runner.initialize(); await h.runner.run(h.controller.signal); h.values[0] = 9;
  await assert.rejects(h.runner.run(new AbortController().signal), /已改变/);
  await assert.rejects(h.runner.rewind(0), /已改变/); assert.deepEqual(h.restores, []); assert.equal(h.values[0], 9);
});
test("failed restoration keeps backups and blocks later rollback steps; retry recovers", async () => {
  const h = harness(); await h.runner.initialize(); await h.runner.run(h.controller.signal); h.failRestore(true);
  await assert.rejects(h.runner.rewind(2), /protected/);
  assert.equal(h.state.status, "recovery"); assert.ok(h.state.steps[3].undo); assert.deepEqual(h.disposed, []);
  h.failRestore(false); await h.runner.rewind(2);
  assert.deepEqual(h.restores, [3, 3, 2]); assert.deepEqual(h.values, [1, 1, 0, 0]);
});
test("checkpoint storage failure prevents workbook mutation", async () => {
  const h = harness(); await h.runner.initialize();
  const execute = h.io.execute;
  h.io.execute = async (...args) => { h.failSave(true); return execute(...args); };
  await assert.rejects(h.runner.run(h.controller.signal), /save failed/);
  assert.deepEqual(h.values, [0, 0, 0, 0]);
});
test("serialized task resumes after a window interruption without replaying completed steps", async () => {
  const h = harness(); h.pause(1); await h.runner.initialize(); await h.runner.run(h.controller.signal);
  const saved = decodeWorkflow(h.saved.at(-1)!);
  saved.steps[2].status = "running"; saved.steps[2].undo = { version: 1, token, sourceSheetId: "2", backups: [], created: [] }; saved.status = "running"; h.values[2] = 1;
  (globalThis as any).Office = { context: { document: { settings: { get: () => JSON.stringify(saved) } } } };
  const recovered = loadWorkflow()!; assert.equal(recovered.status, "recovery");
  const resumed = new WorkflowRunner(recovered, h.io); await resumed.run(new AbortController().signal);
  assert.deepEqual(h.calls, [0, 1, 2, 3]); assert.deepEqual(h.restores, [2]);
});
test("document storage uses only declarative task state and reports a failed save", async () => {
  let stored = "";
  (globalThis as any).Office = { AsyncResultStatus: { Succeeded: "succeeded" }, context: { document: { settings: { get: () => undefined, refreshAsync: (callback: (value: unknown) => void) => callback({ status: "succeeded" }), set: (_key: string, value: string) => { stored = value; }, saveAsync: (callback: (value: unknown) => void) => callback({ status: "failed" }) } } } };
  await assert.rejects(saveWorkflow(task()), /无法保存/);
  assert.equal(stored.includes("apiKey"), false); assert.equal(stored.includes("baseUrl"), false);
  const bad = task(); (bad.steps[0] as any).undo = { version: 1, token: "bad", backups: [], created: [] }; assert.throws(() => decodeWorkflow(JSON.stringify(bad)), /恢复记录/);
});
test("a stale pane cannot save over a newer task revision", async () => {
  let stored: string | undefined;
  (globalThis as any).Office = { AsyncResultStatus: { Succeeded: "ok" }, context: { document: { settings: {
    get: () => stored, set: (_key: string, value: string) => { stored = value; },
    refreshAsync: (callback: (value: unknown) => void) => callback({ status: "ok" }), saveAsync: (callback: (value: unknown) => void) => callback({ status: "ok" }),
  } } } };
  const current = task(); await saveWorkflow(current);
  const stale = decodeWorkflow(stored!); await saveWorkflow(current);
  assert.equal(current.revision, 2);
  await assert.rejects(assertWorkflowRevision(stale), /另一个窗格/);
  await assert.rejects(saveWorkflow(stale), /另一个窗格/);
  assert.equal(decodeWorkflow(stored!).revision, 2);
});
test("manual edits during model verification keep the old guard and prevent a retry from overwriting them", async () => {
  const h = harness(); await h.runner.initialize();
  h.io.execute = async () => { h.values[0] = 99; const error = new Error("changed while checking"); error.name = "WorkbookChangedError"; throw error; };
  await h.runner.run(h.controller.signal);
  await assert.rejects(h.runner.run(new AbortController().signal), /已改变/);
  assert.equal(h.values[0], 99); assert.deepEqual(h.restores, []);
});
const config = { baseUrl: "https://example.com/v1", model: "test-model", apiKey: "test-only" };
const selected = { workbook: "test", sheet: "数据", sheetId: "source", address: "A1", bounds: { row: 0, col: 0, rows: 1, cols: 1 }, values: [[1]], formulas: [[1]], dataMode: "full" as const };
test("planning, step execution context, and verification share the configured compatible model", async () => {
  const original = globalThis.fetch; const received: any[] = [];
  const outline = { summary: "先计算后汇总", steps: task().steps.map(({ status, ...step }) => step) };
  const replies = [outline, { answer: "生成本步", operations: [] }, { passed: false, summary: "汇总口径不符" }];
  globalThis.fetch = (async (_url, init) => { const body = JSON.parse(String(init?.body)); received.push(body); return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(replies.shift()) } }] })); }) as typeof fetch;
  try {
    assert.equal((await planWorkflow("原任务", selected, config)).steps.length, 4);
    await askAgent("第三步", selected, config, undefined, { completed: ["第一步已完成", "第二步已完成"], currentStep: 3 });
    const result = await checkStep(task(), 2, { actual: [[2]] }, config);
    assert.equal(result.passed, false);
    assert.equal(JSON.parse(received[1].messages[1].content).taskContext.currentStep, 3);
    assert.deepEqual(received.map(body => body.model), ["test-model", "test-model", "test-model"]);
    assert.equal(received[0].messages[1].content.includes("test-only"), false);
  } finally { globalThis.fetch = original; }
});
