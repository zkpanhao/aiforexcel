import { test } from "node:test";
import assert from "node:assert/strict";
import { parseAddress, parsePlan, validateScope } from "../src/tools/protocol";
import { literalValues } from "../src/excel/operations";

const scope = { sheet: "数据", address: "B2:C3", bounds: parseAddress("B2:C3") };
const parse = (operations: unknown[]) => parsePlan(JSON.stringify({ answer: "完成", operations }));
test("accept fenced JSON and absolute cell coordinates", () => {
  const plan = parsePlan(' \n```json\n{"answer":"完成","operations":[{"kind":"write_values","sheet":"数据","address":"$B$2:$C$3","data":[[1,2],[3,null]]}]}\n```\n');
  assert.doesNotThrow(() => validateScope(plan, scope));
  assert.deepEqual(parseAddress("XFD1048576"), { row: 1048575, col: 16383, rows: 1, cols: 1 });
});
for (const address of ["A0", "XFE1", "A1048577", "C3:B2", "Sheet1!A1", "A1,C2", "A:A"]) {
  test(`reject invalid address ${address}`, () => assert.throws(() => parseAddress(address)));
}
test("allow cross-sheet and out-of-selection operations", () => {
  for (const [sheet, address] of [["其他", "B2"], ["数据", "A2"], ["数据", "C3:D3"], ["数据", "B1:C2"]]) {
    assert.doesNotThrow(() => validateScope(parse([{ kind: "clear_contents", sheet, address }]), scope));
  }
});
test("reject executable code, unknown keys and invalid dimensions", () => {
  const bad = [
    { kind: "execute_js", code: "anything" },
    { kind: "clear_contents", sheet: "数据", address: "B2", script: "anything" },
    { kind: "write_values", sheet: "数据", address: "B2:C3", data: [[1, 2], [3]] },
    { kind: "write_values", sheet: "数据", address: "B2", data: [[{ x: 1 }]] },
    { kind: "format_range", sheet: "数据", address: "B2", bold: "true" },
    { kind: "format_range", sheet: "数据", address: "B2", fill: "red" },
    { kind: "write_values", sheet: "数据", address: "B2:C3", data: [[1]] },
  ];
  for (const op of bad) assert.throws(() => parse([op]));
});
test("reject external formulas, accept ordinary formulas", () => {
  for (const formula of ['=WEBSERVICE("https://x")', '=HYPERLINK("https://x")', "=[other.xlsx]Sheet1!A1", "=RTD(1)", '=IMAGE("https://x")', "SUM(A1:A2)"]) {
    assert.throws(() => parse([{ kind: "write_formulas", sheet: "数据", address: "B2", data: [[formula]] }]));
  }
  assert.doesNotThrow(() => parse([{ kind: "write_formulas", sheet: "数据", address: "B2", data: [["=SUM(A1:A2)"]] }]));
});
test("literal data cannot become formulas and null clears a cell", () => {
  assert.deepEqual(literalValues([["=1+1", "+1", "-1", "@x", "'a", null, -1, false, "文本"]]), [["'=1+1", "'+1", "'-1", "'@x", "''a", "", -1, false, "文本"]]);
});
test("allow large broadcast formatting and batched value writes", () => {
  assert.throws(() => parse([{ kind: "create_sheet", name: "_ai_BACKUP_test" }]));
  const large = { sheet: "数据", address: "A1:Z20", bounds: parseAddress("A1:Z20") };
  const plan = parse(Array.from({ length: 4 }, () => ({ kind: "clear_formats", sheet: "数据", address: "A1:Z20" })));
  assert.doesNotThrow(() => validateScope(plan, large));
  assert.doesNotThrow(() => parse([{ kind: "write_values", sheet: "数据", address: "A1:Z100", data: Array.from({ length: 100 }, () => Array(26).fill(1)) }]));
});

test("large text cleanup uses bounded declarative rules and aggregate limits", () => {
  const operation = { kind: "transform_text", sheet: "数据", address: "K2:K96016", whitespace: "remove", removeInvisible: true };
  assert.doesNotThrow(() => validateScope(parse([operation]), scope));
  for (const extra of [{ whitespace: "regex" }, { script: "anything" }, { replacements: [{ find: "-", replace: "", regex: true }] }, { replacements: [{ find: "", replace: "" }] }, { removeInvisible: "yes" }]) {
    assert.throws(() => parse([{ ...operation, ...extra }]));
  }
  assert.throws(() => parse([{ kind: "transform_text", sheet: "数据", address: "K2:K96016" }]));
  assert.throws(() => parse([{ ...operation, address: "K1:K100001" }]));
  assert.throws(() => validateScope(parse([operation, operation, operation]), scope), /200000/);
  assert.throws(() => parse([{ kind: "create_sheet", name: "宽表", data: [Array(16385).fill(0)] }]), /行列范围/);
});
test("new-sheet cells count toward aggregate budget", () => {
  const plan = parse([{ kind: "create_sheet", name: "汇总", data: [Array(2000).fill(0)] }, { kind: "clear_contents", sheet: "数据", address: "B2" }]);
  assert.doesNotThrow(() => validateScope(plan, scope));
  const broad = { sheet: "数据", address: "A1:CV1000", bounds: parseAddress("A1:CV1000") };
  assert.throws(() => validateScope(parse([
    { kind: "clear_contents", sheet: "数据", address: "A1:CV1000" },
    { kind: "clear_contents", sheet: "数据", address: "A1:CV1000" },
    { kind: "clear_contents", sheet: "数据", address: "A1:CV1000" },
  ]), broad));
  assert.throws(() => parse(Array(21).fill({ kind: "clear_contents", sheet: "数据", address: "B2" })));
});
