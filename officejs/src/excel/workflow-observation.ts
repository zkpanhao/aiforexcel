import { AgentPlan, addressFromBounds, parseAddress } from "../tools/protocol";
import { RangeRef, Workflow } from "../state/workflow";
import type { UndoSnapshot } from "./recovery";
import { transformText } from "./batches";

async function hash(value: unknown): Promise<string> {
  const data = new TextEncoder().encode(JSON.stringify(value));
  return Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", data)), byte => byte.toString(16).padStart(2, "0")).join("");
}
const properties: Excel.CellPropertiesLoadOptions = { format: {
  fill: { color: true }, font: { bold: true, italic: true, color: true, name: true, size: true, underline: true },
  borders: { color: true, style: true, weight: true }, horizontalAlignment: true, verticalAlignment: true, wrapText: true,
} };
export class WorkbookChangedError extends Error { constructor() { super("等待期间任务涉及的内容、格式或对象已改变，任务已暂停。请结束此任务后重新规划，以保留手动修改。"); this.name = "WorkbookChangedError"; } }

// Hash complete watched ranges in bounded batches. Only hashes enter saved task
// state; cell values and rich per-cell formats stay in the workbook/runtime.
export async function workflowGuard(task: Workflow): Promise<string> {
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets; sheets.load("items/name,items/id");
    await context.sync();
    const visible = sheets.items.filter(sheet => !/^_AI_backup_/i.test(sheet.name));
    const refs: RangeRef[] = [{ sheet: task.anchor.sheetId, address: task.anchor.address }, ...task.watched];
    for (const step of task.steps) for (const item of step.undo?.backups ?? []) refs.push({ sheet: item.sheetId ?? item.sheet, address: item.address });
    for (const step of task.steps) for (const item of step.undo?.created ?? []) {
      const sheet = sheets.getItemOrNullObject(item.id ?? item.name); sheet.load("id"); await context.sync();
      if (sheet.isNullObject) continue;
      const used = sheet.getUsedRangeOrNullObject(); used.load(["address", "rowCount", "columnCount"]); await context.sync();
      if (!used.isNullObject) {
        if (used.rowCount * used.columnCount > 100000) throw new Error("任务输出工作表过大，无法检查；备份已保留");
        refs.push({ sheet: sheet.id, address: used.address.slice(used.address.lastIndexOf("!") + 1) });
      }
    }
    let digest = await hash(visible.map(sheet => ({ id: sheet.id, name: sheet.name })));
    const touched = new Set<string>();
    for (const ref of [...new Map(refs.map(ref => [`${ref.sheet}!${ref.address}`, ref])).values()]) {
      const sheet = sheets.getItem(ref.sheet); sheet.load("id"); await context.sync(); touched.add(sheet.id);
      const bounds = parseAddress(ref.address);
      if (bounds.rows * bounds.cols > 100000) throw new Error("任务检查范围超过 100000 格");
      for (let row = 0; row < bounds.rows; row += Math.max(1, Math.floor(1000 / bounds.cols))) {
        const rows = Math.min(bounds.rows - row, Math.max(1, Math.floor(1000 / bounds.cols)));
        const range = sheet.getRangeByIndexes(bounds.row + row, bounds.col, rows, bounds.cols);
        range.load(["values", "formulas", "numberFormat"]);
        const styles = range.getCellProperties(properties);
        const rowState = range.getRowProperties({ rowHidden: true, format: { rowHeight: true } });
        const columnState = row === 0 ? range.getColumnProperties({ columnHidden: true, format: { columnWidth: true } }) : undefined;
        const merged = range.getMergedAreasOrNullObject(); merged.load("address");
        await context.sync();
        digest = await hash([digest, sheet.id, bounds, row, range.values, range.formulas, range.numberFormat, styles.value, rowState.value, columnState?.value, merged.isNullObject ? "" : merged.address]);
      }
    }
    for (const id of [...touched].sort()) {
      const sheet = sheets.getItem(id);
      const charts = sheet.charts; charts.load("items/id,items/name,items/chartType,items/title/text");
      const filter = sheet.autoFilter; filter.load(["enabled", "criteria"]);
      const pivots = sheet.pivotTables; pivots.load("items/id,items/name");
      await context.sync();
      digest = await hash([digest, id, charts.items.map(chart => ({ id: chart.id, name: chart.name, type: chart.chartType, title: chart.title.text })), filter.enabled, filter.criteria, pivots.items.map(pivot => ({ id: pivot.id, name: pivot.name }))]);
    }
    return digest;
  });
}
export function operationRefs(plan: AgentPlan): RangeRef[] {
  return plan.operations.flatMap(op => op.kind === "create_sheet" ? [{ sheet: op.name, address: `A1:${columnName(op.data?.[0]?.length ?? 1)}${op.data?.length ?? 1}` }] :
    op.kind === "create_pivot" || op.kind === "edit_structure" ? [{ sheet: op.sheet, address: op.address }, { sheet: op.outputSheet, address: "A1" }] : [{ sheet: op.sheet, address: op.address }]);
}
function columnName(count: number): string { let result = ""; for (let n = count; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result; return result; }

export async function inspectStep(plan: AgentPlan, input: RangeRef, snapshot?: UndoSnapshot): Promise<{ evidence: unknown; watched: RangeRef[]; localErrors: string[] }> {
  return Excel.run(async context => {
    const refs = [...operationRefs(plan), input];
    const watched: RangeRef[] = [], evidence: unknown[] = [], localErrors: string[] = [];
    let remaining = 2000;
    for (const ref of [...new Map(refs.map(ref => [`${ref.sheet}!${ref.address}`, ref])).values()]) {
      const sheet = context.workbook.worksheets.getItem(ref.sheet); sheet.load(["id", "name"]);
      const isOutput = plan.operations.some(op => (op.kind === "create_pivot" || op.kind === "edit_structure") && op.outputSheet === ref.sheet);
      const range = isOutput ? sheet.getUsedRange() : sheet.getRange(ref.address);
      range.load(["values", "formulas", "address", "rowCount", "columnCount"]);
      const charts = sheet.charts; charts.load("items/name,items/chartType,items/title/text");
      await context.sync();
      if (range.rowCount * range.columnCount > 100000) throw new Error("步骤输出超出检查上限，已保留备份");
      const address = range.address.slice(range.address.lastIndexOf("!") + 1).replace(/\$/g, "");
      const textOps = plan.operations.filter((op): op is Extract<typeof plan.operations[number], { kind: "transform_text" }> => op.kind === "transform_text" && op.sheet === sheet.name && addressFromBounds(parseAddress(op.address)) === address);
      let textVerification: { complete: boolean; checkedCells: number; passed: boolean } | undefined;
      if (textOps.length && snapshot) {
        const backup = snapshot.backups.find(item => item.sheet === sheet.name && addressFromBounds(parseAddress(item.address)) === address);
        const b = parseAddress(address);
        const conflicts = plan.operations.some(op => {
          if (op.kind === "create_sheet" || op.sheet !== sheet.name || !["write_values", "write_formulas", "clear_contents", "sort_range", "merge_cells", "unmerge_cells", "transform_text"].includes(op.kind)) return false;
          const other = parseAddress(op.address);
          const overlaps = b.row < other.row + other.rows && other.row < b.row + b.rows && b.col < other.col + other.cols && other.col < b.col + b.cols;
          return overlaps && !(op.kind === "transform_text" && addressFromBounds(other) === address);
        });
        if (backup && !conflicts) {
          const before = context.workbook.worksheets.getItem(backup.id ?? backup.name).getRange(backup.address);
          before.load(["values", "formulas"]); await context.sync();
          let passed = true;
          for (let row = 0; row < b.rows; row++) for (let col = 0; col < b.cols; col++) {
            const value = before.values[row][col], formula = before.formulas[row][col];
            const isFormula = typeof formula === "string" && formula.startsWith("=");
            const expected = isFormula || typeof value !== "string" ? value : textOps.reduce((text, op) => transformText(text, op), value);
            if (range.values[row][col] !== expected || isFormula && range.formulas[row][col] !== formula) passed = false;
          }
          textVerification = { complete: true, checkedCells: b.rows * b.cols, passed };
          if (!passed) localErrors.push(`${sheet.name}!${address} 分批文本清理与原始备份逐格核对不一致`);
        }
      }
      watched.push({ sheet: sheet.id, address });
      const cells = range.rowCount * range.columnCount;
      const rows = Math.min(range.rowCount, Math.floor(remaining / range.columnCount));
      remaining -= rows * range.columnCount;
      // Explicitly mark partial observations. Verification must not treat a sample
      // as a complete read of a large result.
      const styles = rows ? range.getRow(0).getCellProperties(properties) : undefined;
      await context.sync();
      evidence.push({ sheet: sheet.name, address, cellCount: cells, complete: rows === range.rowCount, observedRows: rows, values: range.values.slice(0, rows), formulas: range.formulas.slice(0, rows), firstRowFormats: styles?.value, ...(textVerification ? { textVerification } : {}),
        charts: charts.items.map(chart => ({ name: chart.name, type: chart.chartType, title: chart.title.text })) });
      if (isOutput || plan.operations.some(op => op.kind === "write_formulas" && op.sheet === sheet.name && op.address.replace(/\$/g, "") === address)) {
        for (const row of range.values) for (const value of row) if (typeof value === "string" && /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|NUM!|N\/A$|SPILL!|CALC!)/.test(value)) localErrors.push(`${sheet.name}!${address} 含 Excel 错误 ${value}`);
      }
      for (const op of plan.operations) {
        if (op.kind === "create_sheet" || op.sheet !== sheet.name || op.address.replace(/\$/g, "") !== address) continue;
        if (op.kind === "write_values" && JSON.stringify(range.values) !== JSON.stringify(op.data.map(row => row.map(value => value ?? "")))) localErrors.push(`${sheet.name}!${address} 写入值与计划不一致`);
        if (op.kind === "write_formulas" && JSON.stringify(range.formulas).toUpperCase() !== JSON.stringify(op.data).toUpperCase()) localErrors.push(`${sheet.name}!${address} 公式与计划不一致`);
        if (op.kind === "clear_contents" && range.values.some(row => row.some(value => value !== "" && value !== null))) localErrors.push(`${sheet.name}!${address} 未完全清空`);
      }
    }
    return { evidence, watched, localErrors: [...new Set(localErrors)].slice(0, 20) };
  });
}
