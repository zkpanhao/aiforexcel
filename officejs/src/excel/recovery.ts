import { UndoState } from "../state/undo";
import { afterRestoreExtended, beforeRestoreExtended, ExtendedSnapshot } from "./extended";
import { parseAddress } from "../tools/protocol";

export const OWNER = "AIForExcelTransaction";
export interface UndoSnapshot {
  version: 1; token: string; sourceSheetId: string;
  backups: { name: string; id?: string; sheet: string; sheetId?: string; address: string; copyType: Excel.RangeCopyType }[];
  created: { name: string; id?: string }[];
  extended?: ExtendedSnapshot;
}
// Only accept bounded, declarative recovery metadata; never execute saved code.
export function validateSnapshot(state: UndoSnapshot): void {
  if (state?.version !== 1 || !/^[a-f0-9-]{36}$/i.test(state.token) || typeof state.sourceSheetId !== "string" || !state.sourceSheetId ||
      !Array.isArray(state.backups) || state.backups.length > 100 || !Array.isArray(state.created) || state.created.length > 100) throw new Error("恢复记录无效，已停止操作");
  const expected = `_AI_backup_${state.token.replace(/-/g, "").slice(0, 16)}_`;
  for (const item of state.backups) {
    if (!item.name.startsWith(expected) || typeof item.sheet !== "string" || !["All", "Formats"].includes(item.copyType)) throw new Error("恢复目标无效");
    const bounds = parseAddress(item.address);
    if (bounds.rows * bounds.cols > 100000) throw new Error("恢复范围过大");
  }
  for (const item of state.created) if (typeof item.name !== "string" || item.name.length > 31 || item.name.startsWith("_AI_backup_")) throw new Error("恢复工作表无效");
  if (state.extended) {
    const extra = state.extended;
    if (!Array.isArray(extra.dimensions) || extra.dimensions.length > 10000 || !Array.isArray(extra.conditions) || !Array.isArray(extra.merges) || !Array.isArray(extra.charts)) throw new Error("恢复元数据无效");
    [...extra.conditions, ...extra.merges].forEach(parseAddress);
    for (const item of extra.dimensions) if (!["row", "column"].includes(item.axis) || !Number.isInteger(item.index) || item.index < 0 || item.index >= (item.axis === "row" ? 1048576 : 16384) || !Number.isFinite(item.value) || item.value < 0) throw new Error("恢复尺寸无效");
    for (const item of extra.charts) if (!item.name.startsWith(`AI_${state.token.replace(/-/g, "").slice(0, 20)}_`)) throw new Error("恢复图表无效");
    if (extra.filter) {
      if (extra.filter.address) parseAddress(extra.filter.address);
      if (!Array.isArray(extra.filter.rows) || extra.filter.rows.length > 2000 || !Array.isArray(extra.filter.criteria)) throw new Error("恢复筛选无效");
    }
  }
}
export function recoveryState(snapshot: UndoSnapshot): UndoState {
  validateSnapshot(snapshot);
  return {
    snapshot,
    restore: async () => Excel.run(async context => {
      validateSnapshot(snapshot);
      const sheets = context.workbook.worksheets;
      const source = sheets.getItem(snapshot.sourceSheetId);
      source.protection.load("protected"); context.workbook.protection.load("protected");
      const backups = snapshot.backups.map(item => sheets.getItem(item.id ?? item.name));
      const targets = snapshot.backups.map(item => sheets.getItem(item.sheetId ?? item.sheet));
      const owners = backups.map(sheet => sheet.customProperties.getItemOrNullObject(OWNER));
      owners.forEach(owner => owner.load("value"));
      targets.forEach(sheet => sheet.protection.load("protected"));
      const created = snapshot.created.map(item => sheets.getItemOrNullObject(item.id ?? item.name));
      created.forEach(sheet => sheet.load("id"));
      await context.sync();
      if (source.protection.protected || targets.some(sheet => sheet.protection.protected) || context.workbook.protection.protected) throw new Error("工作表或工作簿受保护，无法恢复；备份已保留");
      if (owners.some(owner => owner.isNullObject || owner.value !== snapshot.token)) throw new Error("备份标识不匹配，已停止恢复");
      const createdOwners = created.filter(sheet => !sheet.isNullObject).map(sheet => ({ sheet, owner: sheet.customProperties.getItemOrNullObject(OWNER) }));
      createdOwners.forEach(item => item.owner.load("value")); await context.sync();
      if (snapshot.extended) await beforeRestoreExtended(context, source, snapshot.extended);
      snapshot.backups.forEach((item, index) => targets[index].getRange(item.address).copyFrom(backups[index].getRange(item.address), item.copyType));
      if (snapshot.extended) afterRestoreExtended(source, snapshot.extended);
      for (const { sheet, owner } of createdOwners) if (!owner.isNullObject && owner.value === snapshot.token) { sheet.visibility = Excel.SheetVisibility.visible; sheet.delete(); }
      await context.sync();
    }),
    dispose: async () => {
      if (!snapshot.backups.length) return;
      await Excel.run(async context => {
        const sheets = snapshot.backups.map(item => context.workbook.worksheets.getItemOrNullObject(item.id ?? item.name));
        sheets.forEach(sheet => sheet.load("id")); await context.sync();
        const owned = sheets.filter(sheet => !sheet.isNullObject).map(sheet => ({ sheet, owner: sheet.customProperties.getItemOrNullObject(OWNER) }));
        owned.forEach(item => item.owner.load("value")); await context.sync();
        if (owned.some(item => item.owner.isNullObject || item.owner.value !== snapshot.token)) throw new Error("备份标识不匹配，未清理备份");
        owned.forEach(({ sheet }) => { sheet.visibility = Excel.SheetVisibility.visible; sheet.delete(); }); await context.sync();
      });
    },
  };
}
