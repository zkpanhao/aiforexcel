import { test, afterEach } from "node:test";
import assert from "node:assert/strict";
import { cellBatches, transformText, writeMatrix } from "../src/excel/batches";
import { applyOperations } from "../src/excel/operations";
import { recoveryState, UndoSnapshot } from "../src/excel/recovery";
import { setLastUndo, undoLast } from "../src/state/undo";
import { addressFromBounds, CellMatrix, parseAddress, parsePlan } from "../src/tools/protocol";
import { askAgent, executionPlan } from "../src/llm/client";

afterEach(() => setLastUndo(undefined));
const cleanup = (address: string, extra: object = {}) => parsePlan(JSON.stringify({ answer: "清理文本", operations: [{ kind: "transform_text", sheet: "数据", address, whitespace: "remove", removeInvisible: true, ...extra }] }));

// Queued, address-aware host: batch writes must preserve all cells outside each tile.
function host(rows = 96015, failAt = 0) {
  const sheets = new Map<string, any>(), queue: (() => void)[] = [];
  const writes: { address: string; cells: number }[] = [];
  let nextId = 0, writeCount = 0, failed = false;
  function sheet(name: string) {
    const item: any = { name, id: String(++nextId), data: [] as any[][], styles: [] as string[], load() {}, protection: { protected: false, load() {} }, isNullObject: false };
    const props = new Map();
    item.customProperties = { add: (key: string, value: string) => queue.push(() => props.set(key, value)), getItemOrNullObject: (key: string) => ({ load() {}, isNullObject: !props.has(key), value: props.get(key) }) };
    item.delete = () => queue.push(() => sheets.delete(item.id));
    item.getRange = (address: string) => {
      const b = parseAddress(address);
      const read = () => Array.from({ length: b.rows }, (_, row) => Array.from({ length: b.cols }, (_, col) => item.data[b.row + row]?.[b.col + col] ?? ""));
      const assign = (matrix: any[][]) => matrix.forEach((row, r) => row.forEach((value, c) => {
        item.data[b.row + r] ??= [];
        item.data[b.row + r][b.col + c] = value;
      }));
      return {
        load() {}, hasSpill: false, savedAsArray: false,
        get values() { return read(); }, get formulas() { return read(); },
        set values(data: any[][]) { assert.equal(data.length, b.rows); assert.equal(data[0].length, b.cols); queue.push(() => {
          const value = data.map(row => row.map(cell => typeof cell === "string" && cell.startsWith("'") ? cell.slice(1)
            : typeof cell === "string" && /^\s*\d+\s*$/.test(cell) ? Number(cell) : cell));
          assign(value); writes.push({ address, cells: b.rows * b.cols }); writeCount++;
        }); },
        set formulas(data: any[][]) { this.values = data; },
        getTables: () => ({ items: [], load() {} }), getPivotTables: () => ({ items: [], load() {} }),
        getMergedAreasOrNullObject: () => ({ isNullObject: true, load() {} }),
        copyFrom(from: any) { queue.push(() => { assign(structuredClone(from.values)); item.styles = structuredClone(from.styles); }); },
        get styles() { return item.styles; },
      };
    };
    sheets.set(item.id, item);
    return item;
  }
  const source = sheet("数据");
  source.data = Array.from({ length: rows+2 }, () => [" A \u200bB\u00a0 "]);
  source.styles = ["font:blue", "numberFormat:@", "border:red"];
  const original = structuredClone(source.data);
  const collection = { get items() { return [...sheets.values()]; }, load() {}, add: sheet,
    getItem: (id: string) => { const result = [...sheets.values()].find(s => s.id === id || s.name === id); if (!result) throw new Error("Missing sheet"); return result; },
    getItemOrNullObject: (id: string) => [...sheets.values()].find(s => s.id === id || s.name === id) ?? { isNullObject: true, load() {} },
  };
  const context: any = { workbook: { worksheets: collection, protection: { protected: false, load() {} } }, async sync() {
    const pending = queue.splice(0); pending.forEach(action => action());
    if (!failed && failAt && writeCount >= failAt) { failed = true; throw new Error("Simulated failure after a committed batch"); }
  } };
  (globalThis as any).Excel = { run: async (run: (context: any) => unknown) => run(context), RangeCopyType: { all: "All", formats: "Formats" }, SheetVisibility: { visible: "Visible", veryHidden: "VeryHidden" } };
  const address = `A2:A${rows+1}`;
  const selected: any = { workbook: "synthetic", sheet: "数据", sheetId: source.id, address, bounds: parseAddress(address), dataMode: "metadata", values: [], formulas: [] };
  return { source, original, selected, sheets, writes, context };
}

test("tiles cover tall, wide and non-origin rectangles exactly within 2000 cells", () => {
  for (const address of ["K2:K96016", "B3:CX60", "C8:FFF8", "XFA1048574:XFD1048576"]) {
    const bounds = parseAddress(address), batches = [...cellBatches(bounds)];
    assert.equal(batches.reduce((sum, batch) => sum + batch.rows * batch.cols, 0), bounds.rows * bounds.cols);
    const seen = new Set<string>();
    for (const b of batches) {
      assert.ok(b.rows * b.cols <= 2000);
      for (let row = b.row; row < b.row + b.rows; row++) for (let col = b.col; col < b.col + b.cols; col++) {
        assert.ok(row >= bounds.row && row < bounds.row + bounds.rows && col >= bounds.col && col < bounds.col + bounds.cols);
        assert.ok(!seen.has(`${row},${col}`)); seen.add(`${row},${col}`);
      }
    }
  }
});

test("explicit text rules handle Unicode whitespace and placeholders without guessing", () => {
  assert.equal(transformText(" \tA\u00a0B\u3000\n\u200b\ufeff", { whitespace: "remove", removeInvisible: true }), "AB");
  assert.equal(transformText(" A  B ", { whitespace: "trim" }), "A  B");
  assert.equal(transformText(" A\t\nB ", { whitespace: "collapse" }), "A B");
  assert.equal(transformText("0-/N/A", { whitespace: "remove" }), "0-/N/A");
  assert.equal(transformText("--", { replacements: [{ find: "-", replace: "", wholeCell: true }] }), "--");
  assert.equal(transformText("待定待定", { replacements: [{ find: "待定", replace: "=1" }] }), "=1=1");
});

test("96015-cell model plan executes in 49 host batches and one undo restores all data and formats", async () => {
  const h = host();
  const progress: number[] = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (_url, init) => {
    const body = JSON.parse(String(init?.body)), payload = JSON.parse(body.messages.at(-1).content);
    assert.equal(payload.selection.cellCount, 96015); assert.equal(payload.selection.values, undefined);
    assert.match(body.messages[0].content, /2000格是内部批次大小/);
    return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify({ ...cleanup(h.selected.address), needs_data: false }) } }] }));
  }) as typeof fetch;
  try {
    const turn = await askAgent("清除空格及零宽空格", h.selected, { baseUrl: "https://example.com/v1", model: "test", apiKey: "test-only" });
    const report = await applyOperations(executionPlan(turn), h.selected, { onProgress: update => progress.push(update.completed) });
    assert.match(report!, /修改 96015 格/); assert.equal(progress.length, 49); assert.equal(progress.at(-1), 96015);
    assert.equal(h.writes.length, 49); assert.ok(h.writes.every(write => write.cells <= 2000));
    assert.equal(h.source.data[0][0], h.original[0][0]); assert.equal(h.source.data[96016][0], h.original[96016][0]);
    assert.ok(h.source.data.slice(1,96016).every((row: any[]) => row[0] === "AB"));
    await undoLast(); assert.deepEqual(h.source.data, h.original); assert.deepEqual(h.source.styles, ["font:blue", "numberFormat:@", "border:red"]); assert.equal(h.sheets.size, 1);
  } finally { globalThis.fetch = originalFetch; }
});

test("text cleanup keeps formulas, numbers, booleans and untouched string literals", async () => {
  const h = host(8);
  h.source.data = [["outside"], ["=SUM(1, 2)"], [0], [false], [" A B "], [" - "], [" +text "], [" 'tag "], [" N/A "], ["outside"]];
  const original = structuredClone(h.source.data);
  const report = await applyOperations(cleanup(h.selected.address), h.selected);
  assert.match(report!, /保留 1 格公式/);
  assert.deepEqual(h.source.data, [["outside"], ["=SUM(1, 2)"], [0], [false], ["AB"], ["-"], ["+text"], ["'tag"], ["N/A"], ["outside"]]);
  await undoLast(); assert.deepEqual(h.source.data, original);
});

for (const mode of ["failure", "cancel"] as const) test(`${mode} after two committed batches restores the entire transaction`, async () => {
  const h = host(6200, mode === "failure" ? 2 : 0), controller = new AbortController();
  await assert.rejects(applyOperations(cleanup(h.selected.address), h.selected, { signal: controller.signal, onProgress: progress => { if (mode === "cancel" && progress.completed >= 4000) controller.abort(); } }), /已恢复修改前状态/);
  assert.deepEqual(h.source.data, h.original); assert.equal(h.sheets.size, 1); assert.equal(h.writes.length, 2);
});

test("retained batched steps can serialize and undo in reverse order", async () => {
  const h = host(4200), snapshots: UndoSnapshot[] = [];
  const options = { retainUndo: true, checkpoint: async (snapshot: UndoSnapshot) => { snapshots.push(snapshot); } };
  await applyOperations(cleanup(h.selected.address), h.selected, options);
  await applyOperations(cleanup(h.selected.address, { replacements: [{ find: "AB", replace: "Z", wholeCell: true }] }), h.selected, options);
  assert.equal(h.source.data[1][0], "Z");
  const second = recoveryState(JSON.parse(JSON.stringify(snapshots[1]))); await second.restore(); await second.dispose();
  assert.equal(h.source.data[1][0], "AB");
  const first = recoveryState(JSON.parse(JSON.stringify(snapshots[0]))); await first.restore(); await first.dispose();
  assert.deepEqual(h.source.data, h.original); assert.equal(h.sheets.size, 1);
});

test("generic values, formulas and new-sheet data use host batches and preserve literal text", async () => {
  for (const kind of ["write_values", "write_formulas", "create_sheet"] as const) {
    const h = host(4200);
    const data = Array.from({ length: 4200 }, () => [kind === "write_formulas" ? "=1+1" : "=literal"]);
    if (kind !== "write_formulas") { data[10] = ["0012"]; data[11] = ["12345678901234567890"]; }
    const op = kind === "create_sheet" ? { kind, name: "输出", data } : { kind, sheet: "数据", address: h.selected.address, data };
    await applyOperations(parsePlan(JSON.stringify({ answer: "写入", operations: [op] })), h.selected);
    assert.deepEqual(h.writes.map(write => write.cells), [2000,2000,200]);
    const output = kind === "create_sheet" ? [...h.sheets.values()].find(sheet => sheet.name === "输出").data : h.source.data.slice(1,4201);
    assert.deepEqual(output, data);
    await undoLast(); assert.deepEqual(h.source.data, h.original); assert.equal(h.sheets.size, 1);
  }
});

test("large string payloads split below the cell-count batch size", async () => {
  const h = host(40);
  const data: CellMatrix = Array.from({ length: 40 }, () => ["x".repeat(30000)]);
  await writeMatrix(h.context, h.source, parseAddress("A2:A41"), data, false);
  assert.ok(h.writes.length > 1); assert.equal(h.writes.reduce((sum, write) => sum+write.cells, 0), 40);
  assert.deepEqual(h.source.data.slice(1,41), data);
});
