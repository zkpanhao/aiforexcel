import { applyOperations } from "./operations";
import { readSelection } from "./selection";
import { parsePlan, ExcelOperation } from "../tools/protocol";
import { undoLast, canUndo } from "../state/undo";
import { executionPlan, parseTurn } from "../llm/client";
import { Workflow, WorkflowRunner } from "../state/workflow";
import { clearWorkflow, decodeWorkflow, loadWorkflow, saveWorkflow } from "../state/workflow-store";
import { recoveryState } from "./recovery";
import { inspectStep, workflowGuard } from "./workflow-observation";
import { WorkflowController } from "../taskpane/workflow-controller";
import { verifyNativeBatches } from "./native-batch-check";
import { readFileCommands, saveFileCommand, updateFileCommand, deleteFileCommand } from "../state/common-commands";
import { CommonCommandList } from '../taskpane/common-command-list';
import { readFileConnection, saveFileConnection, forgetFileKey } from "../state/connection-file";
const GUID = "8264d257-5475-42b8-96bf-15fc2785c524";
const completed: string[] = [];
let stage = "初始化", filterDebug = "";
function ensure(value: unknown, message: string): asserts value { if (!value) throw new Error(message); }
async function verify(): Promise<void> {
  const name = `AI_Check_${crypto.randomUUID().slice(0, 7)}`;
  let sheetId = "";
  const data = [["类别", "数量", "金额"], ["甲", 3, 30], ["乙", 1, 10], ["甲", 2, 20]];
  await Excel.run(async context => {
    context.workbook.load("name"); await context.sync();
    ensure(new RegExp(`^Excel add-in ${GUID}(?:\\.\\d+)?(?:\\.xlsx)?$`, "i").test(context.workbook.name), "Verification is restricted to its generated empty workbook");
    ensure(!canUndo(), "Existing undo record: verification refused");
    const sheet = context.workbook.worksheets.add(name); sheet.load("id"); sheet.getRange("A1:C4").values = data; sheet.getRange("A6").values = [["标题"]];
    await context.sync(); sheetId = sheet.id;
  });
  const run = async (label: string, operations: any[], inspect: (sheet: Excel.Worksheet, context: Excel.RequestContext) => Promise<void>, address = "A1:C4") => {
    stage = label;
    const selected = await readSelection("full", { sheetId, address });
    // Use the same conversation-to-execution handoff as the production taskpane.
    const plan = executionPlan(parseTurn(JSON.stringify({ answer: label, needs_data: false, data_request: { reason: "合成测试已提供数据" }, operations: operations.map(op => ({ ...op, sheet: name, address: op.address ?? address })) })));
    try { await applyOperations(plan, selected); }
    catch (error) {
      const actual = await readSelection("full", { sheetId, address: selected.address });
      throw new Error(`${label}: ${error instanceof Error ? error.message : error}; expected=${JSON.stringify({ address: selected.address, values: selected.values, formulas: selected.formulas })}; actual=${JSON.stringify({ address: actual.address, values: actual.values, formulas: actual.formulas })}`);
    }
    await Excel.run(async context => inspect(context.workbook.worksheets.getItem(sheetId), context));
    stage = `恢复 ${label}`; await undoLast();
    await Excel.run(async context => { const range = context.workbook.worksheets.getItem(sheetId).getRange("A1:C4"); range.load("values"); await context.sync(); ensure(JSON.stringify(range.values) === JSON.stringify(data), `${label}: undo changed original values`); });
    completed.push(label);
  };
  const initial = await readSelection("full", { sheetId, address: "A1:C4" });
  stage = "建立可见单元格复制筛选";
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.autoFilter.apply(sheet.getRange("A1:C4")); sheet.autoFilter.apply(sheet.getRange("A1:C4"), 0, { filterOn: "Values", values: ["甲"] }); await context.sync(); });
  await run("可见单元格数值化", [{ kind: "copy_visible_values", sourceAddress: "C2:C4", address: "B2:B4" }], async (sheet, context) => {
    const values = sheet.getRange("B2:B4"), hidden = sheet.getRange("B3"); values.load(["values", "formulas"]); hidden.load("rowHidden"); await context.sync();
    ensure(JSON.stringify(values.values) === JSON.stringify([[30], [1], [20]]) && values.formulas.every(row => row.every(value => value !== "=C2")), "Visible value copy changed hidden rows or kept formulas");
    ensure(hidden.rowHidden, "Visible value copy did not preserve the filtered hidden row");
  });
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).autoFilter.remove(); await context.sync(); });
  await Excel.run(async context => { const lookup = context.workbook.worksheets.add("AI_Lookup"); lookup.getRange("E15:F16").values = [["甲", "甲店"], ["乙", "乙店"]]; await context.sync(); });
  await run("跨表映射替换", [{ kind: "lookup_replace", lookupSheet: "AI_Lookup", lookupAddress: "E15:F16", address: "A2:A4" }], async (sheet, context) => {
    const values = sheet.getRange("A2:A4"); values.load(["values", "formulas"]); await context.sync();
    ensure(JSON.stringify(values.values) === JSON.stringify([["甲店"], ["乙店"], ["甲店"]]) && values.formulas.every(row => row.every(value => value !== "=A2")), "Cross-sheet mapping replacement mismatch");
  });
  await Excel.run(async context => { context.workbook.worksheets.getItem("AI_Lookup").delete(); await context.sync(); });

  let rejected = false;
  try { await applyOperations(parsePlan(JSON.stringify({ answer: "合并保护", operations: [{ kind: "merge_cells", sheet: name, address: "A1:C4" }] })), initial); }
  catch (error) { rejected = error instanceof Error && error.message.includes("丢弃"); }
  ensure(rejected && !canUndo(), "Non-anchor contents were not protected from merge"); completed.push("合并数据保护");
  let rolledBack = false;
  try { await applyOperations(parsePlan(JSON.stringify({ answer: "失败回退", operations: [{ kind: "format_range", sheet: name, address: "A1:C4", bold: true }, { kind: "write_formulas", sheet: name, address: "A2", data: [["=SUM("]] }] })), initial); }
  catch (error) { rolledBack = error instanceof Error && error.message.includes("已恢复修改前状态"); }
  ensure(rolledBack, "Partial failure did not restore the transaction");
  await Excel.run(async context => { const cell = context.workbook.worksheets.getItem(sheetId).getRange("A1"); cell.format.font.load("bold"); const range = context.workbook.worksheets.getItem(sheetId).getRange("A1:C4"); range.load("values"); await context.sync(); ensure(!cell.format.font.bold && JSON.stringify(range.values) === JSON.stringify(data), "Failure rollback lost values or formats"); });
  completed.push("部分失败自动回退");
  await run("字体/字号/对齐/边框", [{ kind: "format_range", fontSize: 15, italic: true, underline: "Single", horizontalAlign: "Center", borders: [{ edges: "All", style: "Dash", color: "#336699", weight: "Thin" }] }], async (sheet, context) => {
    const cell = sheet.getRange("A1"), border = cell.format.borders.getItem("EdgeBottom"); cell.format.font.load(["size", "italic", "underline"]); cell.format.load("horizontalAlignment"); border.load(["style", "color"]); await context.sync();
    ensure(cell.format.font.size === 15 && cell.format.font.italic && cell.format.font.underline === "Single" && cell.format.horizontalAlignment === "Center" && border.style === "Dash", "Extended formatting mismatch");
  });
  const dimensions = await Excel.run(async context => { const row = context.workbook.worksheets.getItem(sheetId).getRange("A1").format; row.load(["rowHeight", "columnWidth"]); await context.sync(); return { row: row.rowHeight, column: row.columnWidth }; });
  await run("行高列宽", [{ kind: "resize_range", rowHeight: 30, columnWidth: 100 }], async (sheet, context) => { const f = sheet.getRange("A1").format; f.load(["rowHeight", "columnWidth"]); await context.sync(); ensure(f.rowHeight === 30 && f.columnWidth > 95, "Dimension mismatch"); });
  await Excel.run(async context => { const f = context.workbook.worksheets.getItem(sheetId).getRange("A1").format; f.load(["rowHeight", "columnWidth"]); await context.sync(); ensure(f.rowHeight === dimensions.row && f.columnWidth === dimensions.column, "Dimension undo mismatch"); });
  await run("自动行高列宽", [{ kind: "resize_range", autofitRows: true, autofitColumns: true }], async (_sheet, context) => { await context.sync(); });
  await run("合并", [{ kind: "merge_cells" }], async (sheet, context) => { const merged = sheet.getRange("A6:C6").getMergedAreasOrNullObject(); merged.load("address"); await context.sync(); ensure(!merged.isNullObject, "Merge failed"); }, "A6:C6");
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).getRange("A6:C6").merge(); await context.sync(); });
  await run("拆分及合并恢复", [{ kind: "unmerge_cells" }], async (sheet, context) => { const merged = sheet.getRange("A6:C6").getMergedAreasOrNullObject(); merged.load("address"); await context.sync(); ensure(merged.isNullObject, "Unmerge failed"); }, "A6:C6");
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId), merged = sheet.getRange("A6:C6").getMergedAreasOrNullObject(); merged.load("address"); await context.sync(); ensure(!merged.isNullObject, "Merged state not restored"); sheet.getRange("A6:C6").unmerge(); await context.sync(); });
  await run("排序", [{ kind: "sort_range", fields: [{ column: 2, ascending: true }], hasHeaders: true }], async (sheet, context) => { const row = sheet.getRange("A2:C2"); row.load("values"); await context.sync(); ensure(row.values[0][0] === "乙" && row.values[0][2] === 10, "Sort broke records"); });
  await run("筛选", [{ kind: "filter_range", rules: [{ column: 0, values: ["甲"] }] }], async (sheet, context) => { const row = sheet.getRange("A3"); row.load("rowHidden"); await context.sync(); ensure(row.rowHidden, "Filter did not hide unmatched row"); });
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.autoFilter.load("enabled"); const row = sheet.getRange("A3"); row.load("rowHidden"); await context.sync(); ensure(!sheet.autoFilter.enabled && !row.rowHidden, "Filter undo mismatch"); });
  stage = "建立原筛选";
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.autoFilter.apply(sheet.getRange("A1:C4")); sheet.autoFilter.apply(sheet.getRange("A1:C4"), 0, { filterOn: "Values", values: ["甲"] }); sheet.autoFilter.load("criteria"); await context.sync(); filterDebug = JSON.stringify(sheet.autoFilter.criteria); });
  await run("清除筛选及原筛选恢复", [{ kind: "clear_filter" }], async (sheet, context) => { sheet.autoFilter.load("enabled"); await context.sync(); ensure(!sheet.autoFilter.enabled, "Clear filter failed"); });
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.autoFilter.load("enabled"); const row = sheet.getRange("A3"); row.load("rowHidden"); await context.sync(); ensure(sheet.autoFilter.enabled && row.rowHidden, "Original filter was not restored"); sheet.autoFilter.remove(); await context.sync(); });
  for (const rule of [{ type: "CellValue", operator: "GreaterThan", formula1: "15", fill: "#FFFF00" }, { type: "ContainsText", text: "甲", bold: true }, { type: "Custom", formula1: "=$C2>15", fill: "#FFFF00" }, { type: "ColorScale" }, { type: "DataBar" }]) {
    await run(`条件格式 ${rule.type}`, [{ kind: "conditional_format", rule }], async (sheet, context) => { const rules = sheet.getRange("A1:C4").conditionalFormats; rules.load("items/id"); await context.sync(); ensure(rules.items.length === 1, "Conditional format missing"); });
    await Excel.run(async context => { const rules = context.workbook.worksheets.getItem(sheetId).getRange("A1:C4").conditionalFormats; rules.load("items/id"); await context.sync(); ensure(rules.items.length === 0, "Conditional format undo failed"); });
  }
  for (const chartType of ["ColumnClustered", "BarClustered", "Line", "Pie", "Area", "XYScatter"]) {
    await run(`图表 ${chartType}`, [{ kind: "create_chart", chartType, title: "测试图" }], async (sheet, context) => { sheet.charts.load("items/name"); await context.sync(); ensure(sheet.charts.items.length === 1, "Chart missing"); if (chartType === "ColumnClustered") { sheet.charts.items[0].name = "用户改名后的测试图"; await context.sync(); } });
    await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.charts.load("items/id"); await context.sync(); ensure(!sheet.charts.items.length, "Chart undo failed"); });
  }
  await run("透视表", [{ kind: "create_pivot", outputSheet: "AI_TestPivot", rows: ["类别"], values: [{ field: "金额", aggregation: "Sum" }] }], async (_sheet, context) => { const output = context.workbook.worksheets.getItem("AI_TestPivot"); output.pivotTables.load("items/name"); const range = output.getUsedRange(); range.load("values"); await context.sync(); ensure(output.pivotTables.items.length === 1 && range.values.flat().includes(50), "Pivot aggregation mismatch"); });
  for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns"]) {
    await run(`副本 ${action}`, [{ kind: "edit_structure", action, outputSheet: "AI_TestStructure", address: action.endsWith("rows") ? "A3:C3" : "B1:B4" }], async (_sheet, context) => { const copy = context.workbook.worksheets.getItem("AI_TestStructure"); const range = copy.getRange("A1:D5"); range.load("values"); await context.sync(); ensure(action === "insert_rows" ? range.values[2][0] === "" && range.values[3][0] === "乙" : action === "delete_rows" ? range.values[2][0] === "甲" : action === "insert_columns" ? range.values[1][1] === "" && range.values[1][2] === 3 : range.values[1][1] === 30, "Structural copy mismatch"); });
  }
  stage = "分步任务检查";
  const workflow: Workflow = { version: 1, id: crypto.randomUUID(), task: "原生分步恢复验证", summary: "计算、汇总及格式", anchor: { sheet: name, sheetId, address: "A1:C4" }, status: "planned", watched: [],
    steps: [
      { title: "修改数量", instruction: "写入9", sheet: name, address: "B2", expected: "9", status: "pending" },
      { title: "计算金额", instruction: "数量乘10", sheet: name, address: "C2", expected: "90", status: "pending" },
      { title: "创建汇总", instruction: "创建测试表", sheet: name, address: "A1:C4", expected: "汇总表存在", status: "pending" },
      { title: "美化汇总", instruction: "设置粗体", sheet: "AI_Workflow_Check", address: "A1:B2", expected: "粗体", status: "pending" },
    ] };
  let failThird = true;
  const calls: number[] = [];
  const io = {
    save: saveWorkflow, guard: workflowGuard,
    restore: async (state: Parameters<typeof recoveryState>[0]) => recoveryState(state).restore(),
    dispose: async (state: Parameters<typeof recoveryState>[0]) => recoveryState(state).dispose(),
    execute: async (current: Workflow, index: number, _signal: AbortSignal, checkpoint: (state: Parameters<typeof recoveryState>[0]) => Promise<void>) => {
      calls.push(index);
      const step = current.steps[index];
      const selection = await readSelection("full", { sheetId: step.sheet, address: step.address });
      const operations: ExcelOperation[][] = [
        [{ kind: "write_values", sheet: name, address: "B2", data: [[9]] }],
        [{ kind: "write_formulas", sheet: name, address: "C2", data: [["=B2*10"]] }],
        [{ kind: "create_sheet", name: "AI_Workflow_Check", data: [["结果", "金额"], ["汇总", 90]] }],
        [{ kind: "format_range", sheet: "AI_Workflow_Check", address: "A1:B2", bold: true }],
      ];
      const plan = parsePlan(JSON.stringify({ answer: step.title, operations: operations[index] }));
      await applyOperations(plan, selection, { retainUndo: true, checkpoint });
      const observed = await inspectStep(plan, { sheet: selection.sheetId, address: selection.address });
      ensure(!observed.localErrors.length, observed.localErrors.join(";"));
      return { passed: index !== 2 || !failThird, result: "原生步骤结果已读取", watched: observed.watched };
    },
  };
  const first = new WorkflowRunner(workflow, io); await first.initialize(); await first.run(new AbortController().signal);
  ensure(workflow.steps[0].status === "done" && workflow.steps[1].status === "done" && workflow.steps[2].status === "failed", "Failure did not preserve first two steps");
  completed.push("分步：第三步失败保留前两步");
  const stored = loadWorkflow(); ensure(stored?.id === workflow.id && stored.steps[2].undo, "Workbook settings did not retain step recovery");
  const resumed = new WorkflowRunner(decodeWorkflow(JSON.stringify(stored)), io);
  failThird = false; await resumed.run(new AbortController().signal);
  ensure(resumed.task.status === "completed" && JSON.stringify(calls) === "[0,1,2,2,3]", "Continuation replayed completed steps");
  completed.push("分步：持久化重载后仅重做第三步及后续");
  await resumed.rewind(2, "重做汇总");
  await Excel.run(async context => { const values = context.workbook.worksheets.getItem(sheetId).getRange("B2:C2"); values.load("values"); const output = context.workbook.worksheets.getItemOrNullObject("AI_Workflow_Check"); output.load("id"); await context.sync(); ensure(output.isNullObject && values.values[0][0] === 9 && values.values[0][1] === 90, "Dependent undo did not preserve earlier data"); });
  completed.push("分步：逆序撤销后续新表及格式，保留早期数据");
  await resumed.run(new AbortController().signal);
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).getRange("B2").values = [[12]]; await context.sync(); });
  let changeBlocked = false;
  try { await resumed.rewind(2); } catch { changeBlocked = true; }
  ensure(changeBlocked, "Manual edits did not stop rollback");
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).getRange("B2").values = [[9]]; await context.sync(); });
  completed.push("分步：手动改表后拒绝重做");
  await resumed.rewind(0); await clearWorkflow();
  await Excel.run(async context => { const values = context.workbook.worksheets.getItem(sheetId).getRange("A1:C4"); values.load("values"); await context.sync(); ensure(JSON.stringify(values.values) === JSON.stringify(data), "Whole task undo mismatch"); });
  completed.push("分步：整任务恢复原值和公式");
  stage = "分步界面与模拟模型全链路";
  await Excel.run(async context => { const sheet = context.workbook.worksheets.getItem(sheetId); sheet.activate(); sheet.getRange("A1:C4").select(); await context.sync(); });
  const originalFetch = window.fetch;
  let requests = 0;
  window.fetch = (async (url, init) => {
    if (String(url) !== "https://workflow-test.invalid/v1/chat/completions") return originalFetch(url, init);
    requests++;
    const payload = JSON.parse(JSON.parse(String(init?.body)).messages[1].content);
    const reply = payload.taskContext ? { answer: "合成操作", operations: payload.taskContext.currentStep === 1
      ? [{ kind: "write_values", sheet: name, address: "B2", data: [[8]] }]
      : [{ kind: "write_formulas", sheet: name, address: "C2", data: [["=B2*10"]] }] } : payload.evidence ? { passed: true, summary: "根据合成结果验证本步骤通过" } :
      { summary: "两步合成任务", steps: [
        { title: "更新数量", instruction: "B2改为8", sheet: name, address: "B2", expected: "B2为8" },
        { title: "更新金额", instruction: "C2改为=B2*10", sheet: name, address: "C2", expected: "C2为80" },
      ] };
    return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(reply) } }] }), { headers: { "Content-Type": "application/json" } });
  }) as typeof fetch;
  try {
    const controller = new WorkflowController({ connection: () => ({ baseUrl: "https://workflow-test.invalid/v1", model: "synthetic", apiKey: "test-only-not-a-real-key" }),
      action: async run => run(new AbortController().signal), append: () => undefined, stage: () => undefined, changed: () => undefined });
    await controller.plan("先改数量再计算金额", new AbortController().signal);
    ensure(controller.runner?.task.status === "completed" && requests === 5, "Controller did not finish plan/execution/check loop");
    controller.render(false, true);
    ensure(document.querySelectorAll("#workflowPanel li").length === 2 && document.getElementById("workflowPanel")?.textContent?.includes("从此步重做"), "Task panel did not show completed steps and retry actions");
    await controller.runner.rewind(0); await clearWorkflow();
    completed.push("分步：模拟模型规划/生成/检查与真实 Excel 及界面完整链路");
  } finally { window.fetch = originalFetch; }
  stage = "大区域分批写入与整轮恢复";
  completed.push(...await verifyNativeBatches(sheetId, name));
  stage = "常用命令文件迁移与独立持久保存";
  let legacy: string | null = JSON.stringify(["旧缓存迁移测试"]);
  const storage = { getItem: () => legacy, removeItem: (key: string) => {
    ensure(key === "ai-for-excel.common-commands.v1", "Migration attempted to touch another browser setting"); legacy = null;
  } };
  ensure(JSON.stringify(await readFileCommands(storage)) === JSON.stringify(["旧缓存迁移测试"]) && legacy === null, "Legacy command migration failed");
  completed.push("常用命令：Excel 内迁移旧缓存到独立 JSON 文件");
  await saveFileCommand("原生文件保存测试", storage);
  ensure(JSON.stringify(await readFileCommands({ getItem: () => null, removeItem: () => {} })) === JSON.stringify(["原生文件保存测试", "旧缓存迁移测试"]), "File commands lost after simulated empty cache");
  completed.push("常用命令：Excel 内保存后空缓存重读，服务端核验文件落盘");
  stage = "常用命令右键修改和删除";
  const draft = document.getElementById('prompt') as HTMLTextAreaElement;
  draft.value = '未发送的原生测试草稿';
  let commandMutation: Promise<void> = Promise.resolve();
  const view = new CommonCommandList(document, window, {
    select: () => { throw new Error('Right click unexpectedly inserted a command'); },
    update: (original, replacement) => commandMutation = (async () => { view.render(await updateFileCommand(original, replacement)); })(),
    remove: original => commandMutation = (async () => { view.render(await deleteFileCommand(original)); })(),
  });
  view.render(await readFileCommands()); view.open();
  const openFirstMenu = () => {
    const first = document.querySelector<HTMLButtonElement>('#commonCommandList .common-command-item');
    ensure(first, 'Command row missing');
    first.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: window.innerWidth - 1, clientY: window.innerHeight - 1 }));
    ensure(!document.getElementById('commandMenu')!.hidden, 'Context menu did not open');
  };
  openFirstMenu(); (document.getElementById('commandEdit') as HTMLButtonElement).click();
  const commandText = document.getElementById('commandEditText') as HTMLTextAreaElement;
  ensure(commandText.value === '原生文件保存测试', 'Editor targeted the wrong command');
  commandText.value = '原生修改后命令\n多行';
  (document.getElementById('commandEditSave') as HTMLButtonElement).click(); await commandMutation; await Promise.resolve();
  ensure((await readFileCommands())[0] === '原生修改后命令\n多行' && draft.value === '未发送的原生测试草稿', 'Command edit did not persist or changed prompt draft');
  completed.push('常用命令：真实右键菜单修改多行命令，保存文件并保留聊天草稿');
  openFirstMenu(); (document.getElementById('commandDelete') as HTMLButtonElement).click(); await commandMutation; await Promise.resolve();
  ensure(JSON.stringify(await readFileCommands()) === JSON.stringify(['旧缓存迁移测试']), 'Delete targeted the wrong command');
  openFirstMenu(); (document.getElementById('commandDelete') as HTMLButtonElement).click(); await commandMutation; await Promise.resolve();
  const oldCommands = { getItem: () => JSON.stringify(['原生文件保存测试', '旧缓存迁移测试']), removeItem: () => {} };
  ensure((await readFileCommands(oldCommands)).length === 0 && draft.value === '未发送的原生测试草稿', 'Deleted commands resurrected from legacy cache');
  completed.push('常用命令：右键逐条删除及空列表持久保存，旧缓存不会恢复已删命令');
  view.close();
  stage = "API 连接文件迁移/保存/清除";
  const testConnection = { provider: "custom" as const, baseUrl: "https://connection-test.invalid/v1", model: "synthetic", apiKey: `test-only-${crypto.randomUUID()}`, rememberKey: true };
  let oldConnection: string | null = JSON.stringify({ version: 2, ...testConnection, keyBaseUrl: testConnection.baseUrl });
  const connectionStorage = { getItem: () => oldConnection, removeItem: (key: string) => {
    ensure(key === "ai-for-excel.connection.v1", "Connection migration attempted to touch unrelated storage"); oldConnection = null;
  } };
  ensure((await readFileConnection(connectionStorage))?.apiKey === testConnection.apiKey && oldConnection === null, "Connection migration failed");
  completed.push("API 配置：Excel 内迁移旧缓存到独立文件");
  await saveFileConnection({ ...testConnection, model: "file-model" });
  const emptyCache = { getItem: () => null, removeItem: () => {} };
  const reopened = await readFileConnection(emptyCache);
  ensure(reopened?.apiKey === testConnection.apiKey && reopened.model === "file-model", "Connection lost after empty cache reload");
  completed.push("API 配置：保存后空缓存重读地址、模型及记住的密钥");
  await forgetFileKey();
  ensure(!(await readFileConnection(emptyCache))?.apiKey, "Forgotten file credential was restored");
  completed.push("API 配置：清除密钥并核验文件及备份不含密钥");
  await Excel.run(async context => { context.workbook.worksheets.getItem(sheetId).delete(); await context.sync(); });
}
Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) return;
  let result: { passed: boolean; completed: string[]; error?: string };
  try { await verify(); result = { passed: true, completed }; }
  catch (error) { result = { passed: false, completed, error: error instanceof Error ? `${stage}: ${error.message}; location=${(error as { debugInfo?: { errorLocation?: string } }).debugInfo?.errorLocation ?? "unknown"}; filter=${filterDebug}` : "Unknown host failure" }; }
  document.body.textContent = JSON.stringify(result, null, 2);
  await fetch("/result", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(result) });
});
