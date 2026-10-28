import { AgentPlan, parseAddress, parsePlan, validateScope } from "../tools/protocol";
import { SelectedRange } from "./selection";
import { getLastUndo, setLastUndo } from "../state/undo";
import { editsCells, outputNames, prepareExtended } from "./extended";
import { applyPivotConditional, supportsPivotConditional } from "./pivot-format";
import { BatchOptions, checkCancellation, cleanText, writeMatrix } from "./batches";
export { literalValues } from "./batches";

import { OWNER, recoveryState, UndoSnapshot } from "./recovery";
export interface ExecutionOptions extends BatchOptions { retainUndo?: boolean; checkpoint?: (snapshot: UndoSnapshot) => Promise<void>; onRecovery?: (snapshot: UndoSnapshot) => void; }
export class RecoveryError extends Error { constructor(message: string) { super(message); this.name = "RecoveryError"; } }

// Native range copies preserve mixed cell formats, borders, and formulas.
export async function applyOperations(plan: AgentPlan, selected: SelectedRange, options: ExecutionOptions = {}): Promise<string | undefined> {
  plan = parsePlan(JSON.stringify(plan));
  validateScope(plan, selected);
  if (!plan.operations.length) return;
  checkCancellation(options.signal);
  const token = crypto.randomUUID();
  const backupName = `_AI_backup_${token.replace(/-/g, "").slice(0, 16)}`;
  const hasRangeEdits = plan.operations.some(editsCells);
  const newNames = outputNames(plan);
  if (new Set(newNames.map(name => name.toLowerCase())).size !== newNames.length) throw new Error("新建工作表名称重复");
  type BackupRecord = { name: string; id?: string; sheet: string; sheetId?: string; address: string; copyType: Excel.RangeCopyType };
  const backups: BackupRecord[] = [];
  let overwriteWarning: string | undefined;
  const textReports: string[] = [];
  let createdIds: string[] = [];
  let mutationStarted = false;
  let activeOperation: string | undefined;
  const previous = getLastUndo();
  let extended: Awaited<ReturnType<typeof prepareExtended>> | undefined;
  const writesOnly = (op: typeof plan.operations[number]) => op.kind === "write_values" || op.kind === "write_formulas";
  const crossSheetExtended = plan.operations.some(op => {
    if (op.kind === "create_sheet" || op.kind === "transform_text" || op.kind === "copy_formats" || op.kind === "copy_visible_values" || writesOnly(op)) return false;
    return op.sheet !== selected.sheet;
  });
  if (crossSheetExtended) throw new Error("跨工作表目前支持写入值、公式、格式复制和可见值复制；请将其它格式、筛选、排序等操作单独发送到目标工作表");

  const snapshot: UndoSnapshot = { version: 1, token, sourceSheetId: selected.sheetId, backups, created: newNames.map(name => ({ name })) };
  const state = recoveryState(snapshot);

  try {
    await Excel.run(async context => {
      const source = context.workbook.worksheets.getItem(selected.sheetId);
      const range = source.getRange(selected.address);
      source.load("name");
      source.protection.load("protected");
      context.workbook.protection.load("protected");
      // Legacy callers without dataMode captured full values; only explicit metadata
      // mode skips the stale-content check.
      if (selected.dataMode !== "metadata") range.load(["values", "formulas"]);
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      const targets = plan.operations.flatMap(op => op.kind === "create_sheet" ? [] : (() => {
        const targetSheet = sheets.getItem(op.sheet), targetRange = targetSheet.getRange(op.address);
        const pivotTables = typeof (targetRange as any).getPivotTables === "function" ? (targetRange as any).getPivotTables(false) as { items: unknown[]; load: (propertyNames: string) => void } : undefined;
        return [{ op, sheet: targetSheet, range: targetRange, tables: targetRange.getTables(false), merged: targetRange.getMergedAreasOrNullObject(), pivotTables }];
      })());
      targets.forEach(item => {
        item.sheet.load(["name", "id"]);
        item.sheet.protection.load("protected");
        item.tables.load("items/name");
        item.merged.load("address");
        item.pivotTables?.load("items/name");
        item.range.load(["hasSpill", "savedAsArray"]);
        if (writesOnly(item.op)) item.range.load(["values", "formulas"]);
      });
      await context.sync();
      if (source.name !== selected.sheet) throw new Error("工作表名称已改变，请重新发送");
      if (source.protection.protected || context.workbook.protection.protected) throw new Error("请先解除工作表和工作簿保护");
      if (selected.dataMode !== "metadata" && (JSON.stringify(range.values) !== JSON.stringify(selected.values) || JSON.stringify(range.formulas) !== JSON.stringify(selected.formulas))) throw new Error("等待期间选区内容已改变，本轮未执行，请重新发送");
      if (targets.some(item => item.sheet.protection.protected)) throw new Error("目标工作表受保护，无法执行本轮操作");
      if (targets.some(item => item.tables.items.length && !["format_range", "copy_formats", "resize_range", "create_chart", "create_pivot", "edit_structure"].includes(item.op.kind))) throw new Error("Excel 表格目前仅支持格式调整、图表或透视分析；请将表格转为普通区域后进行其它操作");
      const pivotContentKinds = ["write_values", "write_formulas", "copy_visible_values", "lookup_replace", "transform_text", "clear_contents", "sort_range", "filter_range", "clear_filter", "merge_cells", "unmerge_cells"];
      if (targets.some(item => item.pivotTables?.items.length && pivotContentKinds.includes(item.op.kind))) throw new Error("透视表内部只允许修改格式，不允许写入、清除内容、排序、筛选或合并单元格");
      const occupied = targets.filter(item => writesOnly(item.op)).map(item => {
        const values = item.range.values as unknown[][], formulas = item.range.formulas as unknown[][];
        let count = 0;
        for (let row = 0; row < values.length; row++) for (let col = 0; col < (values[row]?.length ?? 0); col++) {
          const value = values[row]?.[col], formula = formulas[row]?.[col];
          if ((value !== null && value !== undefined && value !== "") || (formula !== null && formula !== undefined && formula !== "")) count++;
        }
        return count ? { sheet: item.op.sheet, address: item.op.address, count } : undefined;
      }).filter((item): item is { sheet: string; address: string; count: number } => Boolean(item));
      if (occupied.length) overwriteWarning = `提示：${occupied.map(item => `${item.sheet}!${item.address} 已有 ${item.count} 个非空单元格`).join("；")}，本次写入已覆盖这些内容。`;
      if (targets.some(item => hasRangeEdits && (item.range.hasSpill !== false || item.range.savedAsArray !== false))) throw new Error("暂不修改数组或溢出公式区域，请选择普通区域");
      for (const item of targets) {
        if (!item.merged.isNullObject && hasRangeEdits && ["sort_range", "filter_range", "clear_filter", "conditional_format", "write_values", "write_formulas", "copy_visible_values", "lookup_replace", "transform_text"].includes(item.op.kind)) throw new Error("请先拆分目标区域中的合并单元格，再进行数据写入、排序、筛选或条件格式操作");
      }
      if (newNames.some(name => sheets.items.some(sheet => sheet.name.toLowerCase() === name.toLowerCase()))) throw new Error("新工作表名称已存在，请换一个名称");
      const pivotConditional = new Set(targets.filter(item => item.pivotTables?.items.length && item.op.kind === "conditional_format").map(item => item.op));
      pivotConditional.forEach(op => { if (op.kind === "conditional_format" && !supportsPivotConditional(op.rule)) throw new Error("透视表区域不支持此类条件格式，请改用直接字体或填充格式"); });
      const extendedPlan = pivotConditional.size ? { ...plan, operations: plan.operations.filter(op => !(op.kind === "conditional_format" && pivotConditional.has(op))) } : plan;
      extended = await prepareExtended(context, source, extendedPlan, selected, token);
      snapshot.extended = extended.snapshot;
      const pivotConditionalQueue: { op: Extract<typeof plan.operations[number], { kind: "conditional_format" }>; target: Excel.Range }[] = [];
      if (hasRangeEdits) {
        const uniqueTargets = [...new Map(targets.filter(item => editsCells(item.op)).map(item => [`${item.op.sheet}!${item.op.address}`, item])).values()];
        uniqueTargets.forEach((item, index) => {
          const name = `${backupName}_${index + 1}`;
          const backup = sheets.add(name);
          backups.push({ name, sheet: item.op.sheet, sheetId: item.sheet.id, address: item.op.address, copyType: item.pivotTables?.items.length ? Excel.RangeCopyType.formats : Excel.RangeCopyType.all });
          backup.customProperties.add(OWNER, token);
          backup.load("id");
          // Keep the original coordinates so relative formulas are restored exactly.
          backup.getRange(item.op.address).copyFrom(item.range, backups[backups.length - 1].copyType);
          backup.visibility = Excel.SheetVisibility.veryHidden;
        });
        await context.sync();
      }
      await options.checkpoint?.(snapshot);
      checkCancellation(options.signal);
      mutationStarted = true;
      const created: Excel.Worksheet[] = [];
      for (const op of plan.operations) {
        checkCancellation(options.signal);
        if (op.kind === "create_sheet") {
          activeOperation = `create_sheet ${op.name}`;
          const sheet = sheets.add(op.name);
          sheet.customProperties.add(OWNER, token);
          sheet.load("id");
          created.push(sheet);
          if (op.data) await writeMatrix(context, sheet, { row: 0, col: 0, rows: op.data.length, cols: op.data[0].length }, op.data, false, options);
          continue;
        }
        activeOperation = `${op.kind} ${op.sheet}!${op.address}`;
        const target = sheets.getItem(op.sheet).getRange(op.address);
        if (op.kind === "conditional_format" && pivotConditional.has(op)) {
          pivotConditionalQueue.push({ op, target });
          continue;
        }
        switch (op.kind) {
          case "write_values": await writeMatrix(context, sheets.getItem(op.sheet), parseAddress(op.address), op.data, false, options); break;
          case "write_formulas": await writeMatrix(context, sheets.getItem(op.sheet), parseAddress(op.address), op.data, true, options); break;
          case "transform_text": {
            const result = await cleanText(context, sheets.getItem(op.sheet), parseAddress(op.address), op, options);
            textReports.push(`${op.sheet}!${op.address}：已检查 ${parseAddress(op.address).rows * parseAddress(op.address).cols} 格，修改 ${result.changed} 格${result.skippedFormulas ? `，保留 ${result.skippedFormulas} 格公式` : ""}。`);
            break;
          }
          case "clear_contents": target.clear(Excel.ClearApplyTo.contents); break;
          case "clear_formats": target.clear(Excel.ClearApplyTo.formats); break;
          default: await extended.execute(op, target, created, OWNER); break;
        }
      }
      await context.sync();
      checkCancellation(options.signal);
      for (const item of pivotConditionalQueue) await applyPivotConditional(context, sheets.getItem(item.op.sheet), item.target, item.op.address, item.op.rule);
      if (pivotConditionalQueue.length) await context.sync();
      extended.captureCreated();
      createdIds = created.map(sheet => sheet.id);
      snapshot.created.forEach((item, i) => { item.id = createdIds[i]; });
    });
  } catch (error) {
    if (mutationStarted) {
      try { await state.restore(); }
      catch {
        if (options.retainUndo) options.onRecovery?.(snapshot); else setLastUndo(state);
        const detail = error instanceof Error ? ` 原始失败：${error.message}` : "";
        throw new RecoveryError(`执行中断且自动恢复未完成。失败操作：${activeOperation ?? "批量提交"}。${options.retainUndo ? "请使用分步任务的恢复或撤销按钮。" : "请点击“恢复上一步”。"}${backups.length ? `备份 ${backupName} 已保留，请勿删除。` : "恢复记录已保留。"}${detail}`);
      }
      if (!options.retainUndo) await state.dispose().catch(() => undefined);
      throw new Error(`执行失败，已恢复修改前状态。${error instanceof Error ? error.message : "请重试"}`);
    }
    if (options.retainUndo && backups.length) options.onRecovery?.(snapshot);
    else await state.dispose().catch(() => undefined);
    throw error;
  }
  overwriteWarning = [overwriteWarning, ...textReports].filter(Boolean).join(" ") || undefined;
  if (options.retainUndo) return overwriteWarning;
  setLastUndo(state);
  try { await previous?.dispose(); }
  catch { return [overwriteWarning, "操作已完成；上一份隐藏备份未能清理，可继续使用恢复按钮。"].filter(Boolean).join(" "); }
  return overwriteWarning;
}
