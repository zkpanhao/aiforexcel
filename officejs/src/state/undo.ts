import type { UndoSnapshot } from "../excel/recovery";
export interface UndoState { restore: () => Promise<void>; dispose: () => Promise<void>; snapshot?: UndoSnapshot; }
let lastUndo: UndoState | undefined;
export function setLastUndo(state: UndoState | undefined): void { lastUndo = state; }
export function getLastUndo(): UndoState | undefined { return lastUndo; }
export function canUndo(): boolean { return Boolean(lastUndo); }
export async function undoLast(): Promise<string | undefined> {
  if (!lastUndo) throw new Error("没有可恢复的操作");
  const state = lastUndo;
  await state.restore();
  if (lastUndo === state) lastUndo = undefined;
  try { await state.dispose(); }
  catch { return "已恢复，隐藏备份未能清理。"; }
}
