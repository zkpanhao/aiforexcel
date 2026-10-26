import { applyOperations } from "./operations";
import { writeMatrix } from "./batches";
import { readSelection } from "./selection";
import { parsePlan, parseAddress } from "../tools/protocol";
import { recoveryState, UndoSnapshot } from "./recovery";
import { undoLast } from "../state/undo";
import { inspectStep } from "./workflow-observation";

// Called only by the verifier after its generated test-workbook identity check.
export async function verifyNativeBatches(sheetId: string, sheetName: string): Promise<string[]> {
  const completed: string[] = [], address = "K2:K96016";
  const ensure = (value: unknown, message: string) => { if (!value) throw new Error(message); };
  const data: (string | number | boolean)[][] = Array.from({ length: 96015 }, () => [" A \u200bB\u00a0 "]);
  data[1] = [0]; data[2] = [false]; data[3] = [" - "]; data[4] = [" N/A "];
  data[5] = [" 0012 "]; data[6] = [" 12345678901234567890 "]; data[7] = [" =1+1 "];
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    await writeMatrix(context, sheet, parseAddress(address), data, false);
    sheet.getRange("K2").formulas = [['=" A B "']];
    sheet.getRange("K3").format.font.color = "#336699";
    sheet.getRange("K3").numberFormat = [["0.00"]];
    await context.sync();
  });
  const selected = await readSelection("metadata", { sheetId, address });
  const original = await readSelection("full", selected);
  const plan = (range = address, replacements?: { find: string; replace: string; wholeCell: boolean }[]) => parsePlan(JSON.stringify({ answer: "原生分批文本清理", operations: [{ kind: "transform_text", sheet: sheetName, address: range, whitespace: "remove", removeInvisible: true, ...(replacements ? { replacements } : {}) }] }));
  const restored = async () => {
    const actual = await readSelection("full", selected);
    ensure(JSON.stringify(actual.values) === JSON.stringify(original.values) && JSON.stringify(actual.formulas) === JSON.stringify(original.formulas), "Batched recovery lost original data or formulas");
    await Excel.run(async context => { const cell = context.workbook.worksheets.getItem(sheetId).getRange("K3"); cell.load("numberFormat"); cell.format.font.load("color"); await context.sync(); ensure(cell.numberFormat[0][0] === "0.00" && cell.format.font.color === "#336699", "Batched recovery lost formatting"); });
  };
  let batches = 0;
  await applyOperations(plan(), selected, { onProgress: () => { batches++; } });
  ensure(batches === 49, "96015 cells were not submitted in 49 batches");
  const actual = await readSelection("full", selected);
  ensure(actual.formulas[0][0] === '=" A B "' && actual.values[0][0] === " A B ", "Cleanup changed an existing formula");
  ensure(actual.values[1][0] === 0 && actual.values[2][0] === false && actual.values[3][0] === "-" && actual.values[4][0] === "N/A", "Cleanup guessed placeholders or changed non-text values");
  ensure(actual.values[5][0] === "0012" && actual.values[6][0] === "12345678901234567890" && actual.values[7][0] === "=1+1", `Cleanup coerced identifier text or a literal into a formula: ${JSON.stringify({ values: actual.values.slice(5,8), formulas: actual.formulas.slice(5,8) })}`);
  ensure(actual.values.slice(8).every(row => row[0] === "AB"), "Cleanup did not finish all 96015 cells");
  await undoLast(); await restored(); completed.push("分批：96015 格清理/49 批/文本类型与公式格式保留/整轮恢复");
  for (const cancel of [false, true]) {
    const controller = new AbortController(); let rolledBack = false;
    try {
      await applyOperations(plan(), selected, { signal: controller.signal, onProgress: progress => {
        if (progress.completed >= 4000) { if (cancel) controller.abort(); else throw new Error("Injected failure after two native commits"); }
      } });
    } catch (error) { rolledBack = error instanceof Error && error.message.includes("已恢复修改前状态"); }
    ensure(rolledBack, "Partial batches were not restored"); await restored();
    completed.push(cancel ? "分批：提交两批后取消自动恢复全部数据" : "分批：提交两批后失败自动恢复全部数据");
  }
  const small = await readSelection("metadata", { sheetId, address: "K100:K4299" });
  const snapshots: UndoSnapshot[] = [];
  const options = { retainUndo: true, checkpoint: async (snapshot: UndoSnapshot) => { snapshots.push(snapshot); } };
  await applyOperations(plan(small.address), small, options);
  const observed = await inspectStep(plan(small.address), { sheet: sheetId, address: small.address }, snapshots[0]);
  ensure(!observed.localErrors.length && (observed.evidence as any[]).some(item => item.textVerification?.checkedCells === 4200 && item.textVerification.passed), "Full deterministic text verification was not captured");
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).getRange("K4299").values = [["wrong tail"]]; await context.sync(); });
  const corrupted = await inspectStep(plan(small.address), { sheet: sheetId, address: small.address }, snapshots[0]);
  ensure(corrupted.localErrors.length > 0, "Text verification failed to detect corruption outside the model sample");
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).getRange("K4299").values = [["AB"]]; await context.sync(); });
  await applyOperations(plan(small.address, [{ find: "AB", replace: "ZZ", wholeCell: true }]), small, options);
  const second = recoveryState(JSON.parse(JSON.stringify(snapshots[1]))); await second.restore(); await second.dispose();
  const between = await readSelection("full", small); ensure(between.values.every(row => row[0] === "AB"), "Undoing second batch step removed first-step results");
  const first = recoveryState(JSON.parse(JSON.stringify(snapshots[0]))); await first.restore(); await first.dispose();
  await restored(); completed.push("分批：全量结果核对/样本外错误检测/两步独立备份逆序恢复");
  return completed;
}
