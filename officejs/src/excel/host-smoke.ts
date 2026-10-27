import { applyOperations } from "./operations";
import { SelectedRange } from "./selection";
import { parseAddress, parsePlan } from "../tools/protocol";
import { canUndo, undoLast } from "../state/undo";

// Opt-in development check. Never run against a user's ordinary workbook.
export async function runOfficeSmokeTest(): Promise<string> {
  let selected: SelectedRange;
  let testId = "";
  const testName = `AI_Test_${crypto.randomUUID().slice(0, 8)}`;
  const expectedWorkbook = "Excel add-in e6e1a25b-a2dc-4c80-9177-6c1e9e0e9d61";
  const ensure = (condition: boolean, message: string) => { if (!condition) throw new Error(message); };
  await Excel.run(async context => {
    context.workbook.load("name");
    await context.sync();
    ensure(new RegExp(`^${expectedWorkbook}(?:\\.\\d+)?(?:\\.xlsx)?$`).test(context.workbook.name), "Skipped: workbook is not the generated development test workbook");
    ensure(!canUndo(), "Skipped: a user recovery record is already present");
    const sheet = context.workbook.worksheets.add(testName);
    sheet.load("id");
    const range = sheet.getRange("A1:B3");
    range.values = [["保留文字", 2], [12.345, 0], [false, "前缀"]];
    sheet.getRange("B2").formulas = [["=SUM(A2,1)"]];
    const first = sheet.getRange("A1"), last = sheet.getRange("B3");
    first.format.font.name = "Arial"; first.format.font.size = 9; first.format.font.bold = true;
    first.format.font.color = "#C00000"; first.format.fill.color = "#FFE699"; first.format.wrapText = true;
    const border = first.format.borders.getItem(Excel.BorderIndex.edgeBottom);
    border.style = Excel.BorderLineStyle.double; border.color = "#0070C0";
    last.format.font.name = "Consolas"; last.format.font.size = 16; last.format.font.italic = true;
    sheet.getRange("A2").numberFormat = [["0.000"]];
    sheet.getRange("B2").numberFormat = [["0.00"]];
    await context.sync();
    testId = sheet.id;
    range.load(["values", "formulas"]);
    await context.sync();
    selected = { workbook: context.workbook.name, sheet: testName, sheetId: sheet.id, address: "A1:B3", bounds: parseAddress("A1:B3"), values: range.values, formulas: range.formulas, dataMode: "full" };
  });

  const inspect = async () => Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(testId), range = sheet.getRange("A1:B3");
    const first = sheet.getRange("A1"), last = sheet.getRange("B3");
    const border = first.format.borders.getItem(Excel.BorderIndex.edgeBottom);
    range.load(["values", "formulas", "numberFormat"]);
    first.load(["format/font/name", "format/font/size", "format/font/bold", "format/font/color", "format/fill/color", "format/wrapText"]);
    last.load(["format/font/name", "format/font/size", "format/font/italic", "format/font/bold"]);
    border.load(["style", "color"]);
    await context.sync();
    return { values: range.values, formulas: range.formulas, numbers: range.numberFormat,
      first: { name: first.format.font.name, size: first.format.font.size, bold: first.format.font.bold, color: first.format.font.color, fill: first.format.fill.color, wrap: first.format.wrapText },
      last: { name: last.format.font.name, size: last.format.font.size, italic: last.format.font.italic, bold: last.format.font.bold }, border: { style: border.style, color: border.color } };
  });
  const before = await inspect();
  const copyPlan = parsePlan(JSON.stringify({ answer: "测试复制格式", operations: [
    { kind: "copy_formats", sheet: testName, sourceAddress: "A1", address: "A2:B3" },
  ] }));
  await applyOperations(copyPlan, selected!);
  const copied = await inspect();
  ensure(JSON.stringify(copied.values) === JSON.stringify(before.values) && JSON.stringify(copied.formulas) === JSON.stringify(before.formulas), "Format-only copy changed cell contents");
  ensure(copied.last.name === copied.first.name && copied.last.size === copied.first.size && copied.last.bold === copied.first.bold, "Native format-only copy did not repeat source format");
  await undoLast();
  ensure(JSON.stringify(await inspect()) === JSON.stringify(before), "Format-only copy undo mismatch");

  const plan = parsePlan(JSON.stringify({ answer: "测试清除", operations: [
    { kind: "clear_contents", sheet: testName, address: "A1:B3" },
    { kind: "clear_formats", sheet: testName, address: "A1:B3" },
  ] }));
  await applyOperations(plan, selected!);
  const cleared = await inspect();
  ensure(cleared.values.every(row => row.every(value => value === "")), "Native clear did not empty all cells");
  ensure(cleared.first.bold === false && cleared.border.style === Excel.BorderLineStyle.none, "Native clear did not reset font/border");
  await undoLast();
  ensure(JSON.stringify(await inspect()) === JSON.stringify(before), "Native undo mismatch: values/formulas/fonts/borders/number formats");

  const literalPlan = parsePlan(JSON.stringify({ answer: "测试写入", operations: [
    { kind: "write_values", sheet: testName, address: "A1:B3", data: [["=1+1", null], ["+test", "@text"], [123, true]] },
    { kind: "format_range", sheet: testName, address: "A1:B3", numberFormat: "0.00" },
  ] }));
  await applyOperations(literalPlan, selected!);
  const literal = await inspect();
  ensure(literal.values[0][0] === "=1+1" && literal.values[0][1] === "" && literal.values[1][0] === "+test" && literal.values[1][1] === "@text", "Literal text or null clear mismatch");
  ensure(literal.numbers.every(row => row.every(format => format === "0.00")), "Number format matrix mismatch");
  await undoLast();
  ensure(JSON.stringify(await inspect()) === JSON.stringify(before), "Second undo mismatch");
  await Excel.run(async context => { context.workbook.worksheets.getItem(testId).delete(); await context.sync(); });
  return "Passed in native Excel: format-only copy and undo; clear contents/formats; mixed fonts, sizes, colors, fill, wrap, borders, number formats and formula restore; literal text/null writes; format matrix; three undo cycles. Synthetic test sheet removed.";
}
