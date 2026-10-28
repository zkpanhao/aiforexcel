import { test } from "node:test";
import assert from "node:assert/strict";
import { applyOperations } from "../src/excel/operations";
import { canUndo, undoLast } from "../src/state/undo";
import { parseAddress, parsePlan } from "../src/tools/protocol";
import { askAgent, executionPlan, parseTurn } from "../src/llm/client";
import { recoveryState, UndoSnapshot } from "../src/excel/recovery";

// A queued host simulation tests sequencing and recovery, not Excel's copy semantics.
function host(failWrite = false, includeTarget = false, includePivot = false) {
  const events: string[] = [], queue: (() => void)[] = [];
  let counter = 0, failed = false;
  const sheets = new Map<string, any>();
  const range = (owner: any) => ({
    data: [[1], [2]], formats: ["red-bold-border", "blue-italic"],
    get values() { return this.data; },
    set values(value) { queue.push(() => { events.push(`write:${owner.id}`); this.data = structuredClone(value); }); },
    get formulas() { return this.data; },
    set formulas(value) { this.values = value; },
    hasSpill: false, savedAsArray: false, load() {},
    getMergedAreasOrNullObject() { return { isNullObject: true, load() {} }; },
    getTables() { return { items: [], load() {} }; },
    getPivotTables() { return { items: includePivot ? [{}] : [], load() {} }; },
    getVisibleView() {
      const current = this;
      return {
        rowCount: current.data.length, columnCount: 1,
        cellAddresses: current.data.map((_: unknown, index: number) => [`A${index + 1}`]),
        values: current.data,
        load() {},
        set(properties: { values?: unknown[][] }) { if (properties.values) queue.push(() => { events.push(`visible:${owner.id}`); current.data = structuredClone(properties.values); }); },
      };
    },
    copyFrom(from: any, copyType?: string) { queue.push(() => { events.push(`copy:${owner.id}:${copyType ?? "All"}`); this.data = structuredClone(from.data); this.formats = structuredClone(from.formats); }); },
    clear() { queue.push(() => { events.push(`write:${owner.id}`); this.data = [[""], [""]]; this.formats = ["normal", "normal"]; }); },
  });
  function sheet(name: string) {
    const item: any = { name, id: `id-${++counter}`, isNullObject: false, protection: { protected: false, load() {} }, load() {}, visibility: "visible" };
    const props = new Map();
    item.customProperties = {
      add: (key: string, value: string) => queue.push(() => props.set(key, value)),
      getItemOrNullObject: (key: string) => ({ value: props.get(key), isNullObject: !props.has(key), load() {} }),
    };
    item.range = range(item);
    item.getRange = () => item.range;
    item.getRangeByIndexes = () => item.range;
    item.delete = () => queue.push(() => { events.push(`delete:${item.id}`); sheets.delete(item.id); });
    queue.push(() => sheets.set(item.id, item));
    return item;
  }
  const source = sheet("数据"); queue.splice(0).forEach(fn => fn());
  const target = includeTarget ? sheet("其他") : undefined; queue.splice(0).forEach(fn => fn());
  const collection = {
    get items() { return [...sheets.values()]; }, load() {}, add: sheet,
    getItem(id: string) { const found = [...sheets.values()].find(s => s.id === id || s.name === id); if (!found) throw new Error("sheet missing"); return found; },
    getItemOrNullObject(id: string) { return [...sheets.values()].find(s => s.id === id || s.name === id) ?? { isNullObject: true, load() {} }; },
  };
  const context = { workbook: { worksheets: collection, protection: { protected: false, load() {} } }, async sync() {
    const pending = queue.splice(0); pending.forEach(fn => fn()); events.push("sync");
    if (failWrite && !failed && events.includes(`write:${source.id}`)) { failed = true; throw new Error("simulated host failure"); }
  } };
  (globalThis as any).Excel = {
    run: async (callback: (context: any) => unknown) => callback(context),
    RangeCopyType: { all: "All", formats: "Formats" }, ClearApplyTo: { contents: "Contents", formats: "Formats" }, SheetVisibility: { visible: "Visible", veryHidden: "VeryHidden" },
  };
  const selected = { workbook: "test", sheet: source.name, sheetId: source.id, address: "A1:A2", bounds: parseAddress("A1:A2"), values: [[1], [2]], formulas: [[1], [2]] };
  return { source, target, events, sheets, selected };
}
const clear = () => parsePlan(JSON.stringify({ answer: "已清除", operations: [{ kind: "clear_contents", sheet: "数据", address: "A1:A2" }] }));
test("copy_formats uses native format-only copy and keeps cell contents", async () => {
  const h = host();
  const selected = { ...h.selected, dataMode: "metadata" as const, values: [], formulas: [] };
  const plan = parsePlan(JSON.stringify({ answer: "已复制格式", operations: [{ kind: "copy_formats", sheet: "数据", sourceAddress: "A1:A1", address: "A2:A2" }] }));
  await applyOperations(plan, selected);
  assert.ok(h.events.some(event => event.endsWith(":Formats")));
  assert.deepEqual(h.source.range.data, [[1], [2]]);
});
test("copy operations can use the selected sheet as source and target another sheet", async () => {
  const h = host(false, true);
  const selected = { ...h.selected, dataMode: "metadata" as const, values: [], formulas: [] };
  const formatPlan = parsePlan(JSON.stringify({ answer: "已复制格式", operations: [{ kind: "copy_formats", sourceSheet: "数据", sheet: "其他", sourceAddress: "A1:A1", address: "A2:A2" }] }));
  await applyOperations(formatPlan, selected);
  assert.ok(h.events.some(event => event.endsWith(":Formats")));
  const visiblePlan = parsePlan(JSON.stringify({ answer: "已复制可见值", operations: [{ kind: "copy_visible_values", sheet: "其他", sourceAddress: "A1:A2", address: "A1:A2" }] }));
  await applyOperations(visiblePlan, selected);
  assert.ok(h.events.some(event => event.startsWith("visible:")));
});
test("copy_visible_values writes static values through the visible view", async () => {
  const h = host();
  const selected = { ...h.selected, dataMode: "metadata" as const, values: [], formulas: [] };
  const plan = parsePlan(JSON.stringify({ answer: "已复制可见值", operations: [{ kind: "copy_visible_values", sheet: "数据", sourceAddress: "A1:A2", address: "A1:A2" }] }));
  await applyOperations(plan, selected);
  assert.ok(h.events.some(event => event.startsWith("visible:")));
  assert.deepEqual(h.source.range.data, [[1], [2]]);
});
test("snapshot sync precedes mutation; undo restores snapshot and cleans up", async () => {
  const h = host();
  await applyOperations(clear(), h.selected);
  assert.equal(h.sheets.size, 2); assert.equal(canUndo(), true);
  const copy = h.events.findIndex(e => e.startsWith("copy:")), write = h.events.indexOf(`write:${h.source.id}`);
  assert.ok(copy >= 0 && h.events.slice(copy + 1, write).includes("sync"));
  assert.deepEqual(h.source.range.data, [[""], [""]]);
  await undoLast();
  assert.deepEqual(h.source.range.data, [[1], [2]]);
  assert.deepEqual(h.source.range.formats, ["red-bold-border", "blue-italic"]);
  assert.equal(h.sheets.size, 1);
});
test("cross-sheet writes keep an independent backup and report occupied targets", async () => {
  const h = host(false, true);
  const plan = parsePlan(JSON.stringify({ answer: "已写入", operations: [{ kind: "write_values", sheet: "其他", address: "D3", data: [[9]] }] }));
  const warning = await applyOperations(plan, h.selected);
  assert.match(warning ?? "", /其他!D3 已有/);
  assert.deepEqual(h.target?.range.data, [[9]]);
  await undoLast();
  assert.deepEqual(h.target?.range.data, [[1], [2]]);
  assert.equal(h.sheets.size, 2);
});
test("partial host failure triggers rollback and removes backup", async () => {
  const h = host(true);
  await assert.rejects(applyOperations(clear(), h.selected), /已恢复修改前状态/);
  assert.deepEqual(h.source.range.data, [[1], [2]]);
  assert.equal(h.sheets.size, 1); assert.equal(canUndo(), false);
});
test("stale selection stops, while an explicit out-of-selection target executes", async () => {
  const h = host(); h.source.range.data = [[9], [2]];
  await assert.rejects(applyOperations(clear(), h.selected), /内容已改变/);
  assert.equal(h.sheets.size, 1); assert.ok(!h.events.some(e => e.startsWith("write:")));
  const outside = parsePlan('{"answer":"","operations":[{"kind":"clear_contents","sheet":"数据","address":"B1"}]}');
  await applyOperations(outside, { ...h.selected, values: [[9], [2]], formulas: [[9], [2]] });
});

test("model response flows through conversation state into execution and undo", async () => {
  const originalFetch = globalThis.fetch;
  const h = host();
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify({ answer: "清除", needs_data: false, data_request: { reason: "不需额外数据" }, operations: clear().operations }) } }] }))) as typeof fetch;
  try {
    const turn = await askAgent("清除内容", h.selected, { baseUrl: "https://example.com/v1", model: "test", apiKey: "test-only" });
    assert.equal(turn.needsData, false);
    assert.equal(turn.dataReason, "不需额外数据");
    await applyOperations(executionPlan(turn), h.selected);
    assert.deepEqual(h.source.range.data, [[""], [""]]);
    await undoLast();
    assert.deepEqual(h.source.range.data, [[1], [2]]);
    assert.equal(h.sheets.size, 1);
  } finally { globalThis.fetch = originalFetch; }
});

test("execution handoff keeps operation validation strict while allowing explicit targets", async () => {
  const h = host();
  const turn = parseTurn(JSON.stringify({ answer: "", operations: clear().operations }));
  (turn.operations[0] as any).script = "forbidden";
  await assert.rejects(applyOperations(executionPlan(turn), h.selected), /不支持的参数/);
  delete (turn.operations[0] as any).script;
  (turn.operations[0] as any).address = "B1";
  await applyOperations(executionPlan(turn), h.selected);
  assert.equal(h.sheets.size, 2);
});

test("pivot content operations are rejected before creating a backup", async () => {
  const h = host(false, false, true);
  await assert.rejects(applyOperations(clear(), h.selected), /透视表内部只允许修改格式/);
  assert.equal(h.sheets.size, 1);
  assert.ok(!h.events.some(event => event.startsWith("copy:")));
});

test("pivot format operations are allowed and restore formats only", async () => {
  const h = host(false, false, true);
  const plan = parsePlan(JSON.stringify({ answer: "已清除格式", operations: [{ kind: "clear_formats", sheet: "数据", address: "A1:A2" }] }));
  await applyOperations(plan, h.selected);
  assert.deepEqual(h.source.range.formats, ["normal", "normal"]);
  assert.ok(h.events.filter(event => event.startsWith("copy:")).every(event => event.endsWith(":Formats")));
  await undoLast();
  assert.deepEqual(h.source.range.formats, ["red-bold-border", "blue-italic"]);
  assert.deepEqual(h.source.range.data, [[1], [2]]);
});
test("retained step snapshots survive serialization and restore in reverse order without deleting earlier backups", async () => {
  const h = host(); const snapshots: UndoSnapshot[] = [];
  await applyOperations(clear(), h.selected, { retainUndo: true, checkpoint: async state => {
    assert.ok(!h.events.includes(`write:${h.source.id}`)); snapshots.push(state);
  } });
  await applyOperations(parsePlan(JSON.stringify({ answer: "写入第二步", operations: [{ kind: "write_values", sheet: "数据", address: "A1:A2", data: [[9], [8]] }] })),
    { ...h.selected, values: [[""], [""]], formulas: [[""], [""]] }, { retainUndo: true, checkpoint: async state => { snapshots.push(state); } });
  assert.equal(h.sheets.size, 3);
  const second = recoveryState(JSON.parse(JSON.stringify(snapshots[1])));
  await second.restore(); await second.dispose();
  assert.deepEqual(h.source.range.data, [[""], [""]]); assert.equal(h.sheets.size, 2);
  const first = recoveryState(JSON.parse(JSON.stringify(snapshots[0])));
  await first.restore(); await first.dispose();
  assert.deepEqual(h.source.range.data, [[1], [2]]); assert.equal(h.sheets.size, 1);
});
test("a checkpoint failure leaves the source unchanged and retains a recoverable backup", async () => {
  const h = host(); let snapshot: UndoSnapshot | undefined;
  await assert.rejects(applyOperations(clear(), h.selected, { retainUndo: true, checkpoint: async state => { snapshot = state; throw new Error("progress storage failed"); } }), /progress storage failed/);
  assert.deepEqual(h.source.range.data, [[1], [2]]); assert.ok(snapshot); assert.equal(h.sheets.size, 2);
  await recoveryState(snapshot).restore(); await recoveryState(snapshot).dispose(); assert.equal(h.sheets.size, 1);
});
