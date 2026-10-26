import { test } from "node:test";
import assert from "node:assert/strict";
import { applyPivotConditional, supportsPivotConditional } from "../src/excel/pivot-format";

test("pivot fallback accepts simple custom and cell-value rules", () => {
  assert.equal(supportsPivotConditional({ type: "Custom", formula1: "=N4<0", fontColor: "#FF0000" }), true);
  assert.equal(supportsPivotConditional({ type: "CellValue", operator: "LessThan", formula1: "0", fontColor: "#FF0000" }), true);
  assert.equal(supportsPivotConditional({ type: "ColorScale", lowColor: "#F8696B", highColor: "#63BE7B" }), false);
});

test("pivot fallback applies direct font formatting to matching rows", async () => {
  const formatted: { color?: string }[] = [{}, {}];
  const target: any = {
    values: [[1], [-2]],
    load() {},
    getCell(row: number) { return { format: { font: formatted[row], fill: {}, } }; },
  };
  const references: Record<string, number> = { N4: 1, N5: -2 };
  const sheet: any = { getRange(address: string) { return { values: [[references[address]]], load() {} }; } };
  const context: any = { async sync() {} };
  await applyPivotConditional(context, sheet, target, "F4:F5", { type: "Custom", formula1: "=N4<0", fontColor: "#FF0000" });
  assert.deepEqual(formatted, [{}, { color: "#FF0000" }]);
});
