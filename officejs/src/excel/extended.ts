import { AgentPlan, ExcelOperation, RangeStyle, ConditionalRule, parseAddress } from "../tools/protocol";
import { SelectedRange } from "./selection";

const localAddress = (address: string) => address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, "");
export function restoreFilterCriteria(criteria: Excel.FilterCriteria | null): Excel.FilterCriteria | undefined {
  if (!criteria?.filterOn) return;
  const fields: Record<string, string[]> = { Values: ["values"], Custom: ["criterion1", "criterion2", "operator"], CellColor: ["color"], FontColor: ["color"], Dynamic: ["dynamicCriteria"], Icon: ["icon"], TopItems: ["criterion1"], TopPercent: ["criterion1"], BottomItems: ["criterion1"], BottomPercent: ["criterion1"] };
  const keys = fields[criteria.filterOn];
  if (!keys) throw new Error("原筛选规则无法可靠恢复，请先手动清除筛选");
  if (criteria.filterOn === "Values" && !criteria.values?.length || criteria.filterOn === "Custom" && !criteria.criterion1) return;
  // Excel reports inactive columns as BottomItems with null criterion1.
  if (["TopItems", "TopPercent", "BottomItems", "BottomPercent"].includes(criteria.filterOn) && !criteria.criterion1) return;
  if (["CellColor", "FontColor"].includes(criteria.filterOn) && !criteria.color || criteria.filterOn === "Icon" && !criteria.icon || criteria.filterOn === "Dynamic" && (!criteria.dynamicCriteria || criteria.dynamicCriteria === "Unknown")) return;
  const result: Record<string, unknown> = { filterOn: criteria.filterOn };
  for (const key of [...keys, "subField"]) {
    const value = (criteria as unknown as Record<string, unknown>)[key];
    if (value !== undefined && value !== null && value !== "" && (key !== "operator" || Boolean(criteria.criterion2) && (value === "And" || value === "Or"))) result[key] = structuredClone(value);
  }
  return result as unknown as Excel.FilterCriteria;
}
export function outputNames(plan: AgentPlan): string[] {
  return plan.operations.flatMap(op => op.kind === "create_sheet" ? [op.name] : op.kind === "create_pivot" || op.kind === "edit_structure" ? [op.outputSheet] : []);
}
export function editsCells(op: ExcelOperation): boolean {
  return !["create_sheet", "create_pivot", "edit_structure", "create_chart"].includes(op.kind);
}
export function applyRangeStyle(target: Excel.Range, op: RangeStyle, address: string): void {
  const font = target.format.font;
  if (op.bold !== undefined) font.bold = op.bold;
  if (op.italic !== undefined) font.italic = op.italic;
  if (op.underline !== undefined) font.underline = op.underline;
  if (op.fontName !== undefined) font.name = op.fontName;
  if (op.fontSize !== undefined) font.size = op.fontSize;
  if (op.fontColor !== undefined) font.color = op.fontColor;
  if (op.fill !== undefined) target.format.fill.color = op.fill;
  if (op.wrap !== undefined) target.format.wrapText = op.wrap;
  if (op.horizontalAlign !== undefined) target.format.horizontalAlignment = op.horizontalAlign;
  if (op.verticalAlign !== undefined) target.format.verticalAlignment = op.verticalAlign;
  const bounds = parseAddress(address);
  if (op.numberFormat !== undefined) target.numberFormat = Array.from({ length: bounds.rows }, () => Array(bounds.cols).fill(op.numberFormat!));
  const outside = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
  const inside = [...(bounds.rows > 1 ? ["InsideHorizontal"] : []), ...(bounds.cols > 1 ? ["InsideVertical"] : [])];
  for (const spec of op.borders ?? []) {
    const edges = spec.edges === "All" ? [...outside, ...inside] : spec.edges === "Outside" ? outside : spec.edges === "Inside" ? inside : [`Edge${spec.edges}`];
    for (const edge of edges) {
      const border = target.format.borders.getItem(edge as Excel.BorderIndex);
      border.style = spec.style ?? "Continuous";
      if (spec.color !== undefined) border.color = spec.color;
      if (spec.weight !== undefined) border.weight = spec.weight;
    }
  }
}
function conditional(target: Excel.Range, rule: ConditionalRule): void {
  const type = rule.type === "ContainsText" ? "ContainsText" : rule.type;
  const cf = target.conditionalFormats.add(type as Excel.ConditionalFormatType);
  let format: Excel.ConditionalRangeFormat | undefined;
  switch (rule.type) {
    case "CellValue":
      cf.cellValue.rule = { operator: rule.operator === "NotEqual" ? "NotEqualTo" : rule.operator!, formula1: rule.formula1!, ...(rule.formula2 ? { formula2: rule.formula2 } : {}) };
      format = cf.cellValue.format; break;
    case "ContainsText": cf.textComparison.rule = { operator: "Contains", text: rule.text! }; format = cf.textComparison.format; break;
    case "Custom": cf.custom.rule.formula = rule.formula1!; format = cf.custom.format; break;
    case "ColorScale": cf.colorScale.criteria = { minimum: { type: "LowestValue", color: rule.lowColor ?? "#F8696B" }, maximum: { type: "HighestValue", color: rule.highColor ?? "#63BE7B" } }; break;
    case "DataBar": cf.dataBar.positiveFormat.fillColor = rule.highColor ?? "#638EC6"; cf.dataBar.negativeFormat.fillColor = rule.lowColor ?? "#FF0000"; break;
  }
  if (format) {
    if (rule.fill !== undefined) format.fill.color = rule.fill;
    if (rule.fontColor !== undefined) format.font.color = rule.fontColor;
    if (rule.bold !== undefined) format.font.bold = rule.bold;
  }
}

// Captures only metadata affected by extended operations; cell data stays in the
// native range backup managed by operations.ts. No model-generated code is run.
export interface ExtendedSnapshot {
  dimensions: { axis: "row" | "column"; index: number; value: number }[];
  conditions: string[]; merges: string[];
  charts: { name: string; id?: string }[];
  filter?: { enabled: boolean; criteria: (Excel.FilterCriteria | null | undefined)[]; address?: string; column: number; rows: { index: number; hidden: boolean }[] };
}
export async function beforeRestoreExtended(context: Excel.RequestContext, sheet: Excel.Worksheet, state: ExtendedSnapshot): Promise<void> {
  state.conditions.forEach(address => sheet.getRange(address).conditionalFormats.clearAll());
  state.merges.forEach(address => sheet.getRange(address).unmerge());
  const charts = state.charts.map(record => sheet.charts.getItemOrNullObject(record.id ?? record.name));
  charts.forEach(chart => chart.load("id"));
  await context.sync();
  charts.filter(chart => !chart.isNullObject).forEach(chart => chart.delete());
}
export function afterRestoreExtended(sheet: Excel.Worksheet, state: ExtendedSnapshot): void {
  for (const item of state.dimensions) {
    const range = sheet.getRangeByIndexes(item.axis === "row" ? item.index : 0, item.axis === "column" ? item.index : 0, 1, 1);
    if (item.axis === "row") range.format.rowHeight = item.value; else range.format.columnWidth = item.value;
  }
  if (state.filter) {
    const filter = state.filter;
    sheet.autoFilter.remove();
    if (filter.enabled && filter.address) {
      sheet.autoFilter.apply(sheet.getRange(filter.address));
      filter.criteria.forEach((criteria, column) => { if (criteria?.filterOn) sheet.autoFilter.apply(sheet.getRange(filter.address!), column, criteria); });
    }
    for (const item of filter.rows) sheet.getRangeByIndexes(item.index, filter.column, 1, 1).rowHidden = item.hidden;
  }
}
export async function prepareExtended(context: Excel.RequestContext, source: Excel.Worksheet, plan: AgentPlan, selected: SelectedRange, token: string) {
  const dimensions: { axis: "row" | "column"; index: number; value: number }[] = [];
  const dimensionReads: { axis: "row" | "column"; index: number; range: Excel.Range }[] = [];
  const dimensionKeys = new Set<string>();
  const mergeOps = plan.operations.filter(op => op.kind === "merge_cells" || op.kind === "unmerge_cells");
  const mergeReads = plan.operations.filter(op => op.kind === "merge_cells").map(op => ({ op, range: source.getRange(op.address) }));
  const conditionOps = plan.operations.filter(op => op.kind === "conditional_format");
  const conditionRanges = conditionOps.map(op => ({ op, range: source.getRange(op.address), formats: source.getRange(op.address).conditionalFormats }));
  conditionRanges.forEach(item => item.formats.load("items/id"));
  const filterOps = plan.operations.filter(op => op.kind === "filter_range" || op.kind === "clear_filter");
  const filter = filterOps.length ? source.autoFilter : undefined;
  const oldFilterRange = filter?.getRangeOrNullObject();
  filter?.load(["enabled", "criteria"]); oldFilterRange?.load("address");
  const rowReads: { index: number; range: Excel.Range }[] = [];
  for (const op of plan.operations) {
    if (op.kind !== "resize_range") continue;
    const b = parseAddress(op.address);
    for (const axis of ["row", "column"] as const) {
      const applies = axis === "row" ? op.rowHeight !== undefined || op.autofitRows : op.columnWidth !== undefined || op.autofitColumns;
      if (!applies) continue;
      const start = axis === "row" ? b.row : b.col, count = axis === "row" ? b.rows : b.cols;
      for (let index = start; index < start + count; index++) {
        const key = `${axis}:${index}`; if (dimensionKeys.has(key)) continue; dimensionKeys.add(key);
        const range = source.getRangeByIndexes(axis === "row" ? index : 0, axis === "column" ? index : 0, 1, 1);
        range.format.load(axis === "row" ? "rowHeight" : "columnWidth"); dimensionReads.push({ axis, index, range });
      }
    }
  }
  const filterBounds = filterOps.length ? parseAddress(filterOps[0].address) : undefined;
  if (filter && filterBounds) {
    if (filterBounds.rows > 2000) throw new Error("可恢复的筛选目前最多 2000 行，请缩小目标区域");
    for (let index = filterBounds.row; index < filterBounds.row + filterBounds.rows; index++) {
      const range = source.getRangeByIndexes(index, filterBounds.col, 1, 1); range.load("rowHidden"); rowReads.push({ index, range });
    }
  }
  mergeReads.forEach(item => item.range.load("formulas"));
  const pivots = plan.operations.filter(op => op.kind === "create_pivot").map(op => ({ op, header: source.getRange(op.address).getRow(0) }));
  pivots.forEach(item => item.header.load("values"));
  const structure = plan.operations.find(op => op.kind === "edit_structure");
  const used = structure ? source.getUsedRangeOrNullObject() : undefined;
  used?.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();
  if (conditionRanges.some(item => item.formats.items.length)) throw new Error("目标区域已有条件格式；为避免覆盖原规则，请指定没有条件格式的区域");
  for (const { op, range } of mergeReads) {
    const rows = op.across ? range.formulas : [range.formulas.flat()];
    if (rows.some(row => row.slice(1).some(value => value !== "" && value !== null))) throw new Error("合并会丢弃非左上角单元格内容，请先整理数据；本轮未执行");
  }
  for (const { op, header } of pivots) {
    const names = header.values[0];
    if (names.some(name => typeof name !== "string" || !name.trim()) || new Set(names).size !== names.length) throw new Error("透视表要求第一行是非空且不重复的文字字段名");
    const axes = [...op.rows, ...(op.columns ?? [])];
    if (new Set(axes).size !== axes.length || [...axes, ...op.values.map(value => value.field)].some(name => !names.includes(name))) throw new Error("透视字段不存在或行列字段重复，请核对表头");
  }
  if (structure && used && !used.isNullObject) {
    if (used.rowCount * used.columnCount > 100000) throw new Error("行列插删需复制整张工作表；当前已用区域超过 100000 格，请拆分后处理");
    const b = parseAddress(structure.address);
    if (structure.action === "insert_rows" && used.rowIndex + used.rowCount + b.rows > 1048576 || structure.action === "insert_columns" && used.columnIndex + used.columnCount + b.cols > 16384) throw new Error("插入将超过工作表边界");
  }
  const filterState = filter ? { enabled: filter.enabled, criteria: (filter.criteria ?? []).map(restoreFilterCriteria), address: oldFilterRange && !oldFilterRange.isNullObject ? localAddress(oldFilterRange.address) : undefined } : undefined;
  if (filterState?.address && filterOps.some(op => localAddress(op.address) !== filterState.address)) throw new Error("工作表已有其它区域的筛选，请选中原筛选区域后再操作");
  const hiddenRows = rowReads.map(item => ({ index: item.index, hidden: item.range.rowHidden }));
  for (const item of dimensionReads) dimensions.push({ axis: item.axis, index: item.index, value: item.axis === "row" ? item.range.format.rowHeight : item.range.format.columnWidth });
  const chartRecords: { name: string; proxy: Excel.Chart; id?: string }[] = [];
  const snapshot: ExtendedSnapshot = {
    dimensions, conditions: conditionOps.map(op => op.address), merges: mergeOps.map(op => op.address),
    charts: plan.operations.filter(op => op.kind === "create_chart").map((_, i) => ({ name: `AI_${token.replace(/-/g, "").slice(0, 20)}_${i}` })),
    ...(filterState ? { filter: { ...filterState, column: filterBounds?.col ?? selected.bounds.col, rows: hiddenRows } } : {}),
  };
  let chartIndex = 0;
  return {
    snapshot,
    captureCreated: () => { for (const [index, record] of chartRecords.entries()) { record.id = record.proxy.id; snapshot.charts[index].id = record.id; } },
    execute: async (op: ExcelOperation, target: Excel.Range | undefined, created: Excel.Worksheet[], ownerKey: string) => {
      switch (op.kind) {
        case "format_range": applyRangeStyle(target!, op, op.address); break;
        case "resize_range":
          if (op.rowHeight !== undefined) target!.format.rowHeight = op.rowHeight;
          if (op.columnWidth !== undefined) target!.format.columnWidth = op.columnWidth;
          if (op.autofitRows) target!.format.autofitRows();
          if (op.autofitColumns) target!.format.autofitColumns();
          break;
        case "merge_cells": target!.merge(op.across ?? false); break;
        case "unmerge_cells": target!.unmerge(); break;
        case "sort_range": target!.sort.apply(op.fields.map(field => ({ key: field.column, ascending: field.ascending })), op.matchCase ?? false, op.hasHeaders); break;
        case "filter_range":
          source.autoFilter.remove(); source.autoFilter.apply(target!);
          for (const rule of op.rules) source.autoFilter.apply(target!, rule.column, rule.values ? { filterOn: "Values", values: rule.values } : { filterOn: "Custom", criterion1: rule.criterion1, ...(rule.criterion2 ? { criterion2: rule.criterion2 } : {}), ...(rule.operator ? { operator: rule.operator } : {}) });
          break;
        case "clear_filter": source.autoFilter.remove(); break;
        case "conditional_format": conditional(target!, op.rule); break;
        case "create_chart": {
          const chart = source.charts.add(op.chartType, target!, op.seriesBy ?? "Columns");
          const name = `AI_${token.replace(/-/g, "").slice(0, 20)}_${chartIndex++}`;
          chartRecords.push({ name, proxy: chart }); chart.load("id"); chart.name = name; chart.title.text = op.title; chart.title.visible = true;
          chart.setPosition(target!.getCell(0, 0)); chart.width = 480; chart.height = 288; break;
        }
        case "create_pivot": {
          const output = context.workbook.worksheets.add(op.outputSheet);
          output.customProperties.add(ownerKey, token); output.load("id"); created.push(output);
          const pivot = output.pivotTables.add(`AI_Pivot_${token.replace(/-/g, "").slice(0, 12)}`, target!, output.getRange("A1"));
          op.rows.forEach(name => pivot.rowHierarchies.add(pivot.hierarchies.getItem(name)));
          op.columns?.forEach(name => pivot.columnHierarchies.add(pivot.hierarchies.getItem(name)));
          for (const value of op.values) { const field = pivot.dataHierarchies.add(pivot.hierarchies.getItem(value.field)); field.summarizeBy = value.aggregation; }
          break;
        }
        case "edit_structure": {
          const output = source.copy("After", source); output.name = op.outputSheet;
          output.customProperties.add(ownerKey, token); output.load("id"); created.push(output);
          const range = output.getRange(op.address);
          if (op.action === "insert_rows") range.getEntireRow().insert("Down");
          if (op.action === "delete_rows") range.getEntireRow().delete("Up");
          if (op.action === "insert_columns") range.getEntireColumn().insert("Right");
          if (op.action === "delete_columns") range.getEntireColumn().delete("Left");
          break;
        }
        default: throw new Error("扩展执行器不支持此操作");
      }
    },
  };
}
