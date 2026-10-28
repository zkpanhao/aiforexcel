import { test } from "node:test";
import assert from "node:assert/strict";
import { parsePlan, parseAddress, validateScope, addressFromBounds } from "../src/tools/protocol";
import { applyRangeStyle, outputNames, editsCells, restoreFilterCriteria, resolveWorksheet } from "../src/excel/extended";
const base = { sheet: "数据", address: "A1:C4" };
const parse = (operations: unknown[]) => parsePlan(JSON.stringify({ answer: "", operations }));
test("selection addresses preserve merged rectangle dimensions across column boundaries", () => {
  for (const address of ["A6:C6", "Z9:AB11", "XFD1048576", "A1"]) assert.equal(addressFromBounds(parseAddress(address)), address);
});
test("lookup replacement resolves sheet names locally and reports missing sheets clearly", () => {
  const context = { workbook: { worksheets: { items: [{ name: "111表" }] } } } as any;
  assert.equal(resolveWorksheet(context, "111表").name, "111表");
  assert.equal(resolveWorksheet(context, "111表 ").name, "111表");
  assert.throws(() => resolveWorksheet(context, "不存在"), /找不到参考工作表/);
});
test("extended operations accept the advertised bounded parameters", () => {
  const operations = [
    { kind: "format_range", ...base, fontName: "微软雅黑", fontSize: 14, italic: true, underline: "Double", horizontalAlign: "Center", verticalAlign: "Top", borders: [{ edges: "All", style: "Continuous", weight: "Thin", color: "#336699" }] },
    { kind: "copy_formats", ...base, sourceAddress: "A1:C1" },
    { kind: "copy_visible_values", ...base, sourceAddress: "A1:C4" },
    { kind: "lookup_replace", ...base, address: "A1:A4", lookupSheet: "111表", lookupAddress: "E15:F24" },
    { kind: "resize_range", ...base, rowHeight: 25, autofitColumns: true },
    { kind: "merge_cells", ...base, across: true }, { kind: "unmerge_cells", ...base },
    { kind: "sort_range", ...base, fields: [{ column: 2, ascending: false }], hasHeaders: true },
    { kind: "filter_range", ...base, rules: [{ column: 1, criterion1: ">10", criterion2: "<20", operator: "And" }] },
    { kind: "filter_range", ...base, rules: [{ column: 0, values: ["甲"] }] }, { kind: "clear_filter", ...base },
    { kind: "create_chart", ...base, chartType: "Line", title: "趋势" },
    { kind: "create_pivot", ...base, outputSheet: "汇总", rows: ["类别"], values: [{ field: "金额", aggregation: "Sum" }] },
  ];
  for (const op of operations) assert.doesNotThrow(() => validateScope(parse([op]), { ...base, bounds: parseAddress(base.address) }));
  for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) assert.doesNotThrow(() => parse([{ kind: "edit_structure", ...base, action, outputSheet: "副本" }]));
});
test("conditional formatting allows common rule types and rejects external formulas", () => {
  for (const rule of [
    { type: "CellValue", operator: "Between", formula1: "10", formula2: "20", fill: "#FFFF00" },
    { type: "ContainsText", text: "待办", bold: true }, { type: "Custom", formula1: "=$B2>10", fontColor: "#FF0000" },
    { type: "ColorScale", lowColor: "#FF0000", highColor: "#00FF00" }, { type: "DataBar" },
  ]) assert.doesNotThrow(() => parse([{ kind: "conditional_format", ...base, rule }]));
  for (const rule of [
    { type: "Custom", formula1: '=WEBSERVICE("https://x")' }, { type: "Custom", formula1: "=[x.xlsx]A1" },
    { type: "CellValue", operator: "Between", formula1: "1" }, { type: "ContainsText", text: "x", script: "run()" },
    { type: "DataBar", formula1: "=1" },
  ]) assert.throws(() => parse([{ kind: "conditional_format", ...base, rule }]));
});
test("extended protocol rejects malformed nested parameters, excess scope and destructive mixing", () => {
  for (const op of [
    { kind: "format_range", ...base, fontSize: 0 }, { kind: "format_range", ...base, fontSize: 410 },
    { kind: "format_range", ...base, borders: [{ edges: "All", script: "x" }] },
    { kind: "format_range", ...base, horizontalAlign: "execute" },
    { kind: "copy_formats", ...base, sourceAddress: "not-a-range" },
    { kind: "copy_visible_values", ...base, sourceAddress: "A1:A2", address: "B1:B3" },
    { kind: "lookup_replace", ...base, address: "A1:B4", lookupSheet: "111表", lookupAddress: "E15:F24" },
    { kind: "resize_range", ...base, rowHeight: 20, autofitRows: true },
    { kind: "sort_range", ...base, fields: [{ column: 3, ascending: true }], hasHeaders: true },
    { kind: "filter_range", ...base, rules: [{ column: 0, values: ["x"], criterion1: "=x" }] },
    { kind: "filter_range", ...base, rules: [{ column: 0, values: ["x"] }, { column: 0, values: ["y"] }] },
    { kind: "create_chart", ...base, chartType: "javascript", title: "x" },
    { kind: "create_pivot", ...base, outputSheet: "_AI_backup_x", rows: [], values: [{ field: "x", aggregation: "Sum" }] },
    { kind: "edit_structure", ...base, outputSheet: "副本", action: "run_macro" },
  ]) assert.throws(() => parse([op]));
  assert.doesNotThrow(() => parse([{ kind: "copy_formats", ...base, sourceAddress: "A1:D1", address: "A2:D100" }]));
  assert.throws(() => parse([{ kind: "edit_structure", ...base, outputSheet: "副本", action: "delete_rows" }, { kind: "clear_contents", ...base }]));
  assert.doesNotThrow(() => validateScope(parse([{ kind: "create_chart", ...base, address: "A1:D4", chartType: "Line", title: "x" }]), { ...base, bounds: parseAddress(base.address) }));
});
test("borders map to native edges without applying inside borders to one cell", () => {
  const borders = new Map<string, any>();
  const target: any = { format: { font: {}, fill: {}, borders: { getItem: (edge: string) => { const value = {}; borders.set(edge, value); return value; } } } };
  applyRangeStyle(target, { fontSize: 16, italic: true, horizontalAlign: "Center", borders: [{ edges: "All", color: "#336699", weight: "Medium" }] }, "A1");
  assert.equal(target.format.font.size, 16); assert.equal(target.format.font.italic, true); assert.equal(target.format.horizontalAlignment, "Center");
  assert.deepEqual([...borders.keys()].sort(), ["EdgeBottom", "EdgeLeft", "EdgeRight", "EdgeTop"]);
  for (const border of borders.values()) assert.deepEqual(border, { style: "Continuous", color: "#336699", weight: "Medium" });
});
test("outputs and undo classification include pivots and structural copies", () => {
  const plan = parse([{ kind: "create_pivot", ...base, outputSheet: "透视", rows: [], values: [{ field: "金额", aggregation: "Sum" }] }, { kind: "create_sheet", name: "新表" }]);
  assert.deepEqual(outputNames(plan), ["透视", "新表"]);
  assert.equal(editsCells(plan.operations[0]), false);
  assert.equal(editsCells(parse([{ kind: "filter_range", ...base, rules: [] }]).operations[0]), true);
});
test("macro functions remain closed while aggregate writes use host batching", () => {
  for (const formula of ['=CALL("x","y")', '=EXEC("x")', '=RUN("x")', '=EVALUATE("x")']) assert.throws(() => parse([{ kind: "write_formulas", sheet: "数据", address: "A1", data: [[formula]] }]));
  const plan = parse([{ kind: "write_values", sheet: "数据", address: "A1:A1500", data: Array.from({ length: 1500 }, () => [1]) }, { kind: "write_values", sheet: "数据", address: "B1:B1500", data: Array.from({ length: 1500 }, () => [2]) }]);
  assert.doesNotThrow(() => validateScope(plan, { sheet: "数据", address: "A1:B1500", bounds: parseAddress("A1:B1500") }));
});
test("filter restoration strips inactive/default fields returned by Excel", () => {
  assert.equal(restoreFilterCriteria(null), undefined);
  assert.equal(restoreFilterCriteria({ filterOn: "Values", values: [] }), undefined);
  assert.equal(restoreFilterCriteria({ filterOn: "BottomItems", criterion1: null } as any), undefined);
  assert.deepEqual(restoreFilterCriteria({ filterOn: "Custom", criterion1: "=甲", operator: "Or", criterion2: null } as any), { filterOn: "Custom", criterion1: "=甲" });
  assert.deepEqual(restoreFilterCriteria({ filterOn: "Values", values: ["甲"], operator: "invalid", criterion1: null } as any), { filterOn: "Values", values: ["甲"] });
  assert.deepEqual(restoreFilterCriteria({ filterOn: "Custom", criterion1: ">10", criterion2: "<20", operator: "And", color: "" }), { filterOn: "Custom", criterion1: ">10", criterion2: "<20", operator: "And" });
});
