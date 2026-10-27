export type CellValue = string | number | boolean | null;
export type CellMatrix = CellValue[][];
export interface TextTransform { whitespace?: "remove" | "trim" | "collapse"; removeInvisible?: boolean; replacements?: { find: string; replace: string; wholeCell?: boolean }[]; }
export interface BorderSpec { edges: "All" | "Outside" | "Inside" | "Top" | "Bottom" | "Left" | "Right"; style?: "Continuous" | "Dash" | "Dot" | "Double" | "None"; color?: string; weight?: "Hairline" | "Thin" | "Medium" | "Thick"; }
export interface RangeStyle { bold?: boolean; italic?: boolean; underline?: "None" | "Single" | "Double"; fontName?: string; fontSize?: number; fill?: string; fontColor?: string; numberFormat?: string; wrap?: boolean; horizontalAlign?: "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed"; verticalAlign?: "Top" | "Center" | "Bottom" | "Justify" | "Distributed"; borders?: BorderSpec[]; }
export interface FilterRule { column: number; values?: string[]; criterion1?: string; criterion2?: string; operator?: "And" | "Or"; }
export interface ConditionalRule { type: "CellValue" | "ContainsText" | "Custom" | "ColorScale" | "DataBar"; operator?: "EqualTo" | "NotEqual" | "GreaterThan" | "GreaterThanOrEqual" | "LessThan" | "LessThanOrEqual" | "Between" | "NotBetween"; formula1?: string; formula2?: string; text?: string; fill?: string; fontColor?: string; bold?: boolean; lowColor?: string; highColor?: string; }
export type ExcelOperation =
  | { kind: "write_values"; sheet: string; address: string; data: CellMatrix }
  | { kind: "write_formulas"; sheet: string; address: string; data: string[][] }
  | ({ kind: "transform_text"; sheet: string; address: string } & TextTransform)
  | { kind: "clear_contents" | "clear_formats"; sheet: string; address: string }
  | ({ kind: "format_range"; sheet: string; address: string } & RangeStyle)
  | { kind: "copy_formats"; sheet: string; sourceAddress: string; address: string }
  | { kind: "copy_visible_values"; sheet: string; sourceAddress: string; address: string }
  | { kind: "lookup_replace"; sheet: string; address: string; lookupSheet: string; lookupAddress: string }
  | { kind: "resize_range"; sheet: string; address: string; rowHeight?: number; columnWidth?: number; autofitRows?: boolean; autofitColumns?: boolean }
  | { kind: "merge_cells"; sheet: string; address: string; across?: boolean }
  | { kind: "unmerge_cells"; sheet: string; address: string }
  | { kind: "sort_range"; sheet: string; address: string; fields: { column: number; ascending: boolean }[]; hasHeaders: boolean; matchCase?: boolean }
  | { kind: "filter_range"; sheet: string; address: string; rules: FilterRule[] }
  | { kind: "clear_filter"; sheet: string; address: string }
  | { kind: "conditional_format"; sheet: string; address: string; rule: ConditionalRule }
  | { kind: "create_chart"; sheet: string; address: string; chartType: "ColumnClustered" | "BarClustered" | "Line" | "Pie" | "Area" | "XYScatter"; title: string; seriesBy?: "Columns" | "Rows" }
  | { kind: "create_pivot"; sheet: string; address: string; outputSheet: string; rows: string[]; columns?: string[]; values: { field: string; aggregation: "Sum" | "Count" | "Average" | "Min" | "Max" }[] }
  | { kind: "edit_structure"; sheet: string; address: string; outputSheet: string; action: "insert_rows" | "delete_rows" | "insert_columns" | "delete_columns" }
  | { kind: "create_sheet"; name: string; data?: CellMatrix };
export interface AgentPlan { answer: string; operations: ExcelOperation[]; }
export interface Bounds { row: number; col: number; rows: number; cols: number; }
export function addressFromBounds(bounds: Bounds): string {
  const column = (index: number) => { let text = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text; return text; };
  const first = `${column(bounds.col)}${bounds.row + 1}`, last = `${column(bounds.col + bounds.cols - 1)}${bounds.row + bounds.rows}`;
  return first === last ? first : `${first}:${last}`;
}
export interface SelectionScope { sheet: string; address: string; bounds: Bounds; }
export const WRITE_BATCH_CELLS = 2000;
const WRITE_CELL_LIMIT = 100000;
const RANGE_CELL_LIMIT = 100000;
const TOTAL_CELL_LIMIT = 200000;
const record = (v: unknown): v is Record<string, unknown> => Boolean(v) && typeof v === "object" && !Array.isArray(v);
function only(value: Record<string, unknown>, keys: string[]): void { if (Object.keys(value).some(key => !keys.includes(key))) throw new Error("操作包含不支持的参数"); }
function oneOf(value: unknown, choices: readonly unknown[]): void { if (!choices.includes(value)) throw new Error("操作参数不在支持的选项中"); }
function text(value: unknown, limit = 255): void { if (typeof value !== "string" || !value.trim() || value.length > limit) throw new Error("文字参数无效或过长"); }
function flag(value: unknown): void { if (value !== undefined && typeof value !== "boolean") throw new Error("格式布尔参数无效"); }
function size(value: unknown, min: number, max: number): void { if (value !== undefined && (typeof value !== "number" || !Number.isFinite(value) || value < min || value > max)) throw new Error("尺寸参数超出支持范围"); }
function color(value: unknown): void { if (value !== undefined && (typeof value !== "string" || !/^#[0-9a-f]{6}$/i.test(value))) throw new Error("颜色必须是 #RRGGBB 格式"); }
function sheetName(value: unknown): void { text(value, 31); if (/[\\/*?:\[\]']/.test(value as string) || /^_AI_backup_/i.test(value as string)) throw new Error("新工作表名称无效"); }
function safeFormula(value: unknown): void { text(value, 1000); if (/\[|\]|\||https?:|WEBSERVICE|HYPERLINK|IMAGE\s*\(|RTD\s*\(|(?:CALL|EXEC|REGISTER(?:\.ID)?|RUN|EVALUATE)\s*\(/i.test(value as string)) throw new Error("公式不支持外部数据、链接或宏函数"); }
export const STYLE_KEYS = ["bold", "italic", "underline", "fontName", "fontSize", "fill", "fontColor", "numberFormat", "wrap", "horizontalAlign", "verticalAlign", "borders"];
function validateStyle(op: Record<string, unknown>): void {
  for (const key of ["bold", "italic", "wrap"]) flag(op[key]);
  for (const key of ["fill", "fontColor"]) color(op[key]);
  if (op.fontName !== undefined) text(op.fontName, 100);
  size(op.fontSize, 1, 409);
  if (op.underline !== undefined) oneOf(op.underline, ["None", "Single", "Double"]);
  if (op.horizontalAlign !== undefined) oneOf(op.horizontalAlign, ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"]);
  if (op.verticalAlign !== undefined) oneOf(op.verticalAlign, ["Top", "Center", "Bottom", "Justify", "Distributed"]);
  if (op.numberFormat !== undefined) text(op.numberFormat);
  if (op.borders !== undefined) {
    if (!Array.isArray(op.borders) || !op.borders.length || op.borders.length > 8) throw new Error("边框设置无效");
    for (const border of op.borders) {
      if (!record(border)) throw new Error("边框设置无效");
      only(border, ["edges", "style", "color", "weight"]); oneOf(border.edges, ["All", "Outside", "Inside", "Top", "Bottom", "Left", "Right"]);
      if (border.style !== undefined) oneOf(border.style, ["Continuous", "Dash", "Dot", "Double", "None"]);
      if (border.weight !== undefined) oneOf(border.weight, ["Hairline", "Thin", "Medium", "Thick"]);
      color(border.color);
    }
  }
}

export function parseAddress(address: string): Bounds {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?::\$?([A-Z]{1,3})\$?([1-9]\d{0,6}))?$/i.exec(address);
  if (!match) throw new Error("只支持连续区域，例如 A1:C20");
  const column = (s: string) => [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const row = Number(match[2]) - 1, col = column(match[1]) - 1;
  const endRow = Number(match[4] ?? match[2]), endCol = column(match[3] ?? match[1]);
  if (endRow <= row || endCol <= col || endRow > 1048576 || endCol > 16384) throw new Error("区域地址超出 Excel 范围");
  return { row, col, rows: endRow - row, cols: endCol - col };
}
function matrix(value: unknown, formulas = false): CellMatrix {
  if (!Array.isArray(value) || !value.length || !Array.isArray(value[0]) || !value[0].length) throw new Error("数据必须是非空二维数组");
  const cols = value[0].length;
  if (value.length * cols > WRITE_CELL_LIMIT) throw new Error("单个写入区域最多 100000 格，插件会自动分批写入");
  for (const row of value) {
    if (!Array.isArray(row) || row.length !== cols) throw new Error("数据每行列数必须一致");
    for (const cell of row) {
      if (cell !== null && !["string", "number", "boolean"].includes(typeof cell)) throw new Error("单元格值类型无效");
      if (typeof cell === "number" && !Number.isFinite(cell)) throw new Error("数值必须有限");
      if (typeof cell === "string" && cell.length > 32767) throw new Error("单元格文字过长");
      if (formulas) {
        if (typeof cell !== "string" || !cell.startsWith("=")) throw new Error("公式必须以等号开头，且不支持外部链接");
        safeFormula(cell);
      }
    }
  }
  return value as CellMatrix;
}
export function parsePlan(raw: string): AgentPlan {
  if (raw.length > 500000) throw new Error("操作计划过大，请缩小任务范围");
  const parsed: unknown = JSON.parse(raw.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim());
  if (!record(parsed) || typeof parsed.answer !== "string" || !Array.isArray(parsed.operations)) throw new Error("模型返回的操作计划格式无效");
  if (Object.keys(parsed).some(k => !["answer", "operations"].includes(k))) throw new Error("操作计划包含不支持的参数");
  if (parsed.operations.length > 20) throw new Error("单次最多执行 20 项操作");
  for (const op of parsed.operations) {
    if (!record(op)) throw new Error("操作格式无效");
    if (op.kind === "create_sheet") {
      sheetName(op.name);
      if (op.data !== undefined) { const data = matrix(op.data); if (data.length > 1048576 || data[0].length > 16384) throw new Error("新表数据超出 Excel 行列范围"); }
      if (Object.keys(op).some(k => !["kind", "name", "data"].includes(k))) throw new Error("新建工作表包含不支持的参数");
      continue;
    }
    if (!["write_values", "write_formulas", "transform_text", "clear_contents", "clear_formats", "format_range", "copy_formats", "copy_visible_values", "lookup_replace", "resize_range", "merge_cells", "unmerge_cells", "sort_range", "filter_range", "clear_filter", "conditional_format", "create_chart", "create_pivot", "edit_structure"].includes(String(op.kind))) throw new Error(`不支持的操作：${String(op.kind)}`);
    if (typeof op.sheet !== "string" || typeof op.address !== "string") throw new Error("操作缺少工作表或区域");
    const bounds = parseAddress(op.address);
    if (bounds.rows * bounds.cols > RANGE_CELL_LIMIT) throw new Error("目标区域最多 100000 格");
    const allowed = ["kind", "sheet", "address"];
    if (op.kind === "write_values" || op.kind === "write_formulas") {
      allowed.push("data");
      const data = matrix(op.data, op.kind === "write_formulas");
      if (bounds.rows * bounds.cols > WRITE_CELL_LIMIT) throw new Error("单个写入区域最多 100000 格");
      if (data.length !== bounds.rows || data[0].length !== bounds.cols) throw new Error("写入数据与目标区域的行列数不一致");
    }
    if (op.kind === "transform_text") {
      allowed.push("whitespace", "removeInvisible", "replacements");
      if (op.whitespace !== undefined) oneOf(op.whitespace, ["remove", "trim", "collapse"]);
      flag(op.removeInvisible);
      if (op.replacements !== undefined) {
        if (!Array.isArray(op.replacements) || !op.replacements.length || op.replacements.length > 20) throw new Error("文本替换规则需为 1–20 项");
        for (const rule of op.replacements) {
          if (!record(rule)) throw new Error("文本替换规则无效");
          only(rule, ["find", "replace", "wholeCell"]); text(rule.find, 1000); flag(rule.wholeCell);
          if (typeof rule.replace !== "string" || rule.replace.length > 1000) throw new Error("替换文字无效或过长");
        }
      }
      if (op.whitespace === undefined && op.removeInvisible !== true && op.replacements === undefined) throw new Error("请指定文本清理或替换规则");
    }
    if (op.kind === "format_range") {
      allowed.push(...STYLE_KEYS); validateStyle(op);
    }
    if (op.kind === "copy_formats" || op.kind === "copy_visible_values") {
      allowed.push("sourceAddress");
      if (typeof op.sourceAddress !== "string") throw new Error("复制操作缺少源区域");
      const sourceBounds = parseAddress(op.sourceAddress);
      if (sourceBounds.rows * sourceBounds.cols > RANGE_CELL_LIMIT) throw new Error("复制源区域最多 100000 格");
      if (op.kind === "copy_visible_values" && (sourceBounds.rows !== bounds.rows || sourceBounds.cols !== bounds.cols)) throw new Error("可见单元格复制要求源区域与目标区域尺寸一致");
    }
    if (op.kind === "lookup_replace") {
      allowed.push("lookupSheet", "lookupAddress");
      if (bounds.cols !== 1) throw new Error("映射替换目标必须是单列区域");
      if (typeof op.lookupSheet !== "string" || !op.lookupSheet.trim() || /^_AI_backup_/i.test(op.lookupSheet)) throw new Error("映射表工作表名称无效");
      if (typeof op.lookupAddress !== "string") throw new Error("映射替换缺少参考区域");
      const lookupBounds = parseAddress(op.lookupAddress);
      if (lookupBounds.cols !== 2) throw new Error("映射参考区域必须是两列：旧名称和新名称");
      if (lookupBounds.rows * lookupBounds.cols > RANGE_CELL_LIMIT) throw new Error("映射参考区域最多 100000 格");
    }
    if (op.kind === "resize_range") {
      allowed.push("rowHeight", "columnWidth", "autofitRows", "autofitColumns");
      size(op.rowHeight, 1, 409); size(op.columnWidth, 1, 1500); flag(op.autofitRows); flag(op.autofitColumns);
      if (!allowed.slice(3).some(key => op[key] !== undefined)) throw new Error("请指定行高、列宽或自动调整");
      if (op.autofitRows && op.rowHeight !== undefined || op.autofitColumns && op.columnWidth !== undefined) throw new Error("同一维度不能同时指定尺寸和自动调整");
      if (bounds.rows > 2000 || bounds.cols > 500) throw new Error("尺寸调整最多 2000 行、500 列，请分批处理");
    }
    if (op.kind === "merge_cells") { allowed.push("across"); flag(op.across); }
    if (op.kind === "sort_range") {
      allowed.push("fields", "hasHeaders", "matchCase"); flag(op.hasHeaders); flag(op.matchCase);
      if (op.hasHeaders === undefined || !Array.isArray(op.fields) || !op.fields.length || op.fields.length > 8) throw new Error("排序需指定表头及排序列");
      for (const field of op.fields) {
        if (!record(field)) throw new Error("排序列无效"); only(field, ["column", "ascending"]);
        if (!Number.isInteger(field.column) || (field.column as number) < 0 || (field.column as number) >= bounds.cols || typeof field.ascending !== "boolean") throw new Error("排序列超出区域或方向无效");
      }
    }
    if (op.kind === "filter_range") {
      allowed.push("rules");
      if (!Array.isArray(op.rules) || op.rules.length > Math.min(bounds.cols, 20)) throw new Error("筛选规则过多或无效");
      const seen = new Set();
      for (const rule of op.rules) {
        if (!record(rule)) throw new Error("筛选规则无效"); only(rule, ["column", "values", "criterion1", "criterion2", "operator"]);
        if (!Number.isInteger(rule.column) || (rule.column as number) < 0 || (rule.column as number) >= bounds.cols || seen.has(rule.column)) throw new Error("筛选列超出范围或重复");
        seen.add(rule.column);
        if (rule.values !== undefined) {
          if (rule.criterion1 !== undefined || rule.criterion2 !== undefined || rule.operator !== undefined || !Array.isArray(rule.values) || !rule.values.length || rule.values.length > 100) throw new Error("值筛选不能混用条件筛选");
          rule.values.forEach(value => text(value));
        } else { text(rule.criterion1); if (rule.criterion2 !== undefined) text(rule.criterion2); if (rule.operator !== undefined) oneOf(rule.operator, ["And", "Or"]); }
      }
    }
    if (op.kind === "conditional_format") {
      allowed.push("rule"); const rule = op.rule;
      if (!record(rule)) throw new Error("条件格式规则无效");
      oneOf(rule.type, ["CellValue", "ContainsText", "Custom", "ColorScale", "DataBar"]);
      const ruleKeys = rule.type === "CellValue" ? ["operator", "formula1", "formula2", "fill", "fontColor", "bold"] : rule.type === "ContainsText" ? ["text", "fill", "fontColor", "bold"] : rule.type === "Custom" ? ["formula1", "fill", "fontColor", "bold"] : ["lowColor", "highColor"];
      only(rule, ["type", ...ruleKeys]);
      for (const key of ["fill", "fontColor", "lowColor", "highColor"]) color(rule[key]); flag(rule.bold);
      if (rule.type === "CellValue") {
        oneOf(rule.operator, ["EqualTo", "NotEqual", "GreaterThan", "GreaterThanOrEqual", "LessThan", "LessThanOrEqual", "Between", "NotBetween"]);
        safeFormula(rule.formula1); if (["Between", "NotBetween"].includes(String(rule.operator))) safeFormula(rule.formula2);
        else if (rule.formula2 !== undefined) throw new Error("此比较条件不支持第二个公式");
      }
      if (rule.type === "ContainsText") text(rule.text);
      if (rule.type === "Custom") { safeFormula(rule.formula1); if (!(rule.formula1 as string).startsWith("=")) throw new Error("自定义条件格式公式必须以 = 开头"); }
    }
    if (op.kind === "create_chart") {
      allowed.push("chartType", "title", "seriesBy");
      oneOf(op.chartType, ["ColumnClustered", "BarClustered", "Line", "Pie", "Area", "XYScatter"]); text(op.title, 200);
      if (op.seriesBy !== undefined) oneOf(op.seriesBy, ["Columns", "Rows"]);
    }
    if (op.kind === "create_pivot") {
      allowed.push("outputSheet", "rows", "columns", "values"); sheetName(op.outputSheet);
      for (const key of ["rows", "columns"]) {
        if (op[key] === undefined && key === "columns") continue;
        if (!Array.isArray(op[key]) || (op[key] as unknown[]).length > 8) throw new Error("透视字段列表无效");
        (op[key] as unknown[]).forEach(value => text(value));
      }
      if (!Array.isArray(op.values) || !op.values.length || op.values.length > 8 || bounds.rows < 2) throw new Error("透视表需要表头、数据和汇总字段");
      for (const field of op.values) { if (!record(field)) throw new Error("汇总字段无效"); only(field, ["field", "aggregation"]); text(field.field); oneOf(field.aggregation, ["Sum", "Count", "Average", "Min", "Max"]); }
    }
    if (op.kind === "edit_structure") {
      allowed.push("outputSheet", "action"); sheetName(op.outputSheet); oneOf(op.action, ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]);
      if (parsed.operations.length !== 1) throw new Error("行列插删在工作表副本中单独执行，请将其它操作另行发送");
      if (String(op.action).endsWith("rows") ? bounds.rows > 100 : bounds.cols > 50) throw new Error("一次最多插删 100 行或 50 列");
    }
    if (Object.keys(op).some(k => !allowed.includes(k))) throw new Error("操作包含不支持的参数");
  }
  return parsed as unknown as AgentPlan;
}
export function validateScope(plan: AgentPlan, scope: SelectionScope): void {
  let total = 0;
  let writes = 0;
  for (const op of plan.operations) {
    if (op.kind === "create_sheet") { const count = op.data ? op.data.length * op.data[0].length : 0; total += count; writes += count; }
    else {
      // The selected range is the model's input context, not a write sandbox.
      // Targets are still constrained by parseAddress and the aggregate limits
      // below, while the executor checks the target worksheet and protection.
      const b = parseAddress(op.address);
      total += b.rows * b.cols;
      if (op.kind === "write_values" || op.kind === "write_formulas") writes += b.rows * b.cols;
    }
  }
  if (total > TOTAL_CELL_LIMIT) throw new Error("单次操作累计最多处理 200000 格");
  if (writes > TOTAL_CELL_LIMIT) throw new Error("单次写入累计最多 200000 格，插件会自动分批写入");
}
