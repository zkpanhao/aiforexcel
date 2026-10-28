import "./taskpane.css";
import { readSelection } from "../excel/selection";
import { applyOperations, RecoveryError } from "../excel/operations";
import { askAgent, executionPlan, inferProvider, listModels, providerPreset, PROVIDER_PRESETS, ProviderId, Progress } from "../llm/client";
import { canUndo, undoLast } from "../state/undo";
import { canonicalBaseUrl } from "../state/settings";
import { readFileConnection, saveFileConnection, forgetFileKey } from "../state/connection-file";
import { readFileCommands, saveFileCommand, updateFileCommand, deleteFileCommand } from "../state/common-commands";
import { CommonCommandList } from "./common-command-list";
import { Conversation, operationTargets } from "../state/conversation";
import type { Workflow } from "../state/workflow";
import type { SelectedRange } from "../excel/selection";
import type { ExcelOperation } from "../tools/protocol";
import { reportClientError } from "../state/client-log";
import { WorkflowController } from "./workflow-controller";
declare const __HOST_SMOKE_TEST__: boolean;

const element = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const messages = element("messages"), status = element("status"), stopwatch = element("stopwatch"), selectionLabel = element("selection");
const prompt = element<HTMLTextAreaElement>("prompt");
const clearConversation = element<HTMLButtonElement>("clearConversation");
const sendButton = element<HTMLButtonElement>("send"), undoButton = element<HTMLButtonElement>("undo");
const commonCommandsButton = element<HTMLButtonElement>("commonCommands"), saveCommonCommandButton = element<HTMLButtonElement>("saveCommonCommand");
const cancelButton = element<HTMLButtonElement>("cancel"), modelsButton = element<HTMLButtonElement>("modelsRefresh");
const provider = element<HTMLSelectElement>("provider"), baseUrl = element<HTMLInputElement>("baseUrl"), model = element<HTMLInputElement>("model"), apiKey = element<HTMLInputElement>("apiKey");
const rememberKey = element<HTMLInputElement>("rememberKey"), saveSettings = element<HTMLButtonElement>("saveSettings"), forgetKey = element<HTMLButtonElement>("forgetKey"), settingsStatus = element("settingsStatus");
const modelPicker = element<HTMLSelectElement>("modelPicker"), modelPickerLabel = element("modelPickerLabel"), modelListStatus = element("modelListStatus");
for (const item of PROVIDER_PRESETS) { const option = document.createElement("option"); option.value = item.id; option.textContent = item.label; provider.append(option); }
let busy = false, ready = false, controller: AbortController | undefined;
let workflowRunning = false;
const conversation = new Conversation();
let commonCommands: string[] = [];
let commandsBusy = false;
let connectionLoading = true, settingsPending = 0, settingsRevision = 0;
let modelStopwatchTimer: ReturnType<typeof setInterval> | undefined;
let modelStopwatchStartedAt: number | undefined;
let activeProvider: ProviderId = inferProvider(baseUrl.value);
provider.value = activeProvider;
let keyEndpoint = canonicalBaseUrl(baseUrl.value);
const config = () => ({ provider: provider.value as ProviderId, baseUrl: baseUrl.value.trim(), model: model.value.trim(), apiKey: apiKey.value.trim() });
const workflow = new WorkflowController({ connection: config, append, changed: updateButtons,
  stage: (text, modelStage = false) => { status.textContent = text; if (modelStage) { if (modelStopwatchStartedAt === undefined) startModelStopwatch(); } else stopModelStopwatch(); },
  action: workflowAction,
});
function rememberWorkflow(task: Workflow): void {
  conversation.record(task.task, task.anchor, `workflow_${task.status}`,
    JSON.stringify({ summary: task.summary, steps: task.steps.map(step => ({ title: step.title, status: step.status, result: step.result, error: step.error })) }),
    task.steps.map(step => ({ sheet: step.sheet, address: step.address })));
}
async function workflowAction(action: (signal: AbortSignal) => Promise<void>, abortable = true): Promise<void> {
  if (busy || !ready || settingsPending) return;
  const previousTask = workflow.runner?.task;
  busy = true; workflowRunning = true; const cancellation = new AbortController(); controller = abortable ? cancellation : undefined; updateButtons();
  try { await action(cancellation.signal); status.textContent = workflow.runner?.task.status === "completed" ? "分步任务已完成" : "分步任务已暂停或结束"; }
  catch (error) { void reportClientError({ stage: "分步任务", error: error instanceof Error ? error.message : "操作失败，已停止后续步骤" }); append("分步任务", error instanceof Error ? error.message : "操作失败，已停止后续步骤"); status.textContent = "分步任务已停止"; }
  finally { const task = workflow.runner?.task ?? previousTask; if (task) rememberWorkflow(task); stopModelStopwatch(); controller = undefined; workflowRunning = false; busy = false; updateButtons(); }
}
function queueSettings(action: () => Promise<void>, message: string): void {
  const revision = ++settingsRevision;
  settingsPending++; updateButtons(); settingsStatus.textContent = "正在保存本机连接配置…";
  void action().then(() => { if (revision === settingsRevision) settingsStatus.textContent = message; })
    .catch(error => { if (revision === settingsRevision) settingsStatus.textContent = error instanceof Error ? error.message : "连接配置文件操作失败，未确认保存或清除。"; })
    .finally(() => { settingsPending--; updateButtons(); });
}
function persistSettings(message?: string): void {
  if (connectionLoading || busy) return;
  const snapshot = { ...config(), rememberKey: rememberKey.checked };
  queueSettings(() => saveFileConnection(snapshot, commandMigrationStorage()), message ?? (snapshot.rememberKey && snapshot.apiKey
    ? "已保存到本机文件，重新打开窗格会自动填入密钥。" : "地址和模型已保存到本机文件；密钥仅本次使用。"));
}
function endpointChanged(): void {
  // Clear a credential before it could be sent to a different provider.
  let next: string;
  try { next = canonicalBaseUrl(baseUrl.value); }
  catch { apiKey.value = ""; rememberKey.checked = false; settingsStatus.textContent = "请输入有效的接口地址。"; return; }
  if (next !== keyEndpoint) {
    modelPicker.replaceChildren(); modelPickerLabel.hidden = true; modelListStatus.textContent = "";
    apiKey.value = ""; rememberKey.checked = false; keyEndpoint = next;
    persistSettings("新接口地址已保存到本机文件，请重新填写密钥。");
  } else persistSettings();
}
function providerChanged(): void {
  const next = provider.value as ProviderId;
  const preset = providerPreset(next);
  if (next !== activeProvider) {
    apiKey.value = "";
    rememberKey.checked = false;
    activeProvider = next;
  }
  if (preset.id !== "custom") { baseUrl.value = preset.baseUrl; model.value = preset.defaultModel; }
  endpointChanged();
}
async function restoreConnection(): Promise<void> {
  try {
    const saved = await readFileConnection(commandMigrationStorage());
    if (saved) {
      baseUrl.value = saved.baseUrl; model.value = saved.model; apiKey.value = saved.apiKey; rememberKey.checked = saved.rememberKey;
      activeProvider = saved.provider ?? inferProvider(saved.baseUrl); provider.value = activeProvider;
      keyEndpoint = canonicalBaseUrl(saved.baseUrl);
      settingsStatus.textContent = saved.apiKey ? "已从本机文件载入连接设置和密钥。" : "已从本机文件载入连接设置，请填写密钥。";
    }
  } catch (error) { settingsStatus.textContent = error instanceof Error ? error.message : "读取本机连接配置文件失败。"; }
  finally { connectionLoading = false; updateButtons(); }
}
baseUrl.addEventListener("change", () => { provider.value = inferProvider(baseUrl.value); activeProvider = provider.value as ProviderId; endpointChanged(); });
provider.addEventListener("change", providerChanged);
model.addEventListener("change", () => persistSettings());
modelPicker.addEventListener("change", () => {
  if (!modelPicker.value) return;
  model.value = modelPicker.value;
  persistSettings();
});
apiKey.addEventListener("input", () => { if (rememberKey.checked) persistSettings(); });
rememberKey.addEventListener("change", () => persistSettings());
saveSettings.onclick = () => { endpointChanged(); };
forgetKey.onclick = () => {
  if (busy || connectionLoading) return;
  apiKey.value = ""; rememberKey.checked = false;
  queueSettings(() => forgetFileKey(commandMigrationStorage()), "已清除本机文件及备份中的密钥。");
};
const commandList = new CommonCommandList(document, window, {
  select: command => {
    const current = prompt.value;
    prompt.value = current ? `${current}${/\s$/.test(current) ? "" : "\n"}${command}` : command;
    prompt.focus(); prompt.selectionStart = prompt.selectionEnd = prompt.value.length;
  },
  update: (command, replacement) => mutateCommonCommands(() => updateFileCommand(command, replacement), "常用命令已修改并保存到本机文件"),
  remove: command => mutateCommonCommands(() => deleteFileCommand(command), "常用命令已删除"),
});
function renderCommonCommands(): void { commandList.render(commonCommands); }
async function mutateCommonCommands(action: () => Promise<string[]>, message: string): Promise<void> {
  if (busy || commandsBusy || !ready) throw new Error("请等待当前操作完成后重试");
  commandsBusy = true; updateButtons();
  try { commonCommands = await action(); renderCommonCommands(); status.textContent = message; }
  catch (error) { status.textContent = error instanceof Error ? error.message : "常用命令文件操作失败"; throw error; }
  finally { commandsBusy = false; updateButtons(); }
}
renderCommonCommands();
function commandMigrationStorage(): Storage | undefined { try { return window.localStorage; } catch { return undefined; } }
async function refreshCommonCommands(open = false): Promise<void> {
  if (busy || commandsBusy) return;
  commandsBusy = true; updateButtons();
  try {
    commonCommands = await readFileCommands(commandMigrationStorage());
    renderCommonCommands();
    if (open) {
      commandList.open();
      if (!commonCommands.length) status.textContent = "还没有保存的常用命令";
    }
  } catch (error) {
    commandList.close();
    append("常用命令", error instanceof Error ? error.message : "读取常用命令文件失败");
  } finally { commandsBusy = false; updateButtons(); }
}
commonCommandsButton.onclick = () => {
  if (busy || commandsBusy) return;
  if (!commandList.hidden) { commandList.close(); return; }
  void refreshCommonCommands(true);
};
saveCommonCommandButton.onclick = () => {
  if (busy || commandsBusy) return;
  const text = prompt.value;
  commandsBusy = true; updateButtons();
  void saveFileCommand(text, commandMigrationStorage()).then(commands => {
    commonCommands = commands; renderCommonCommands(); status.textContent = "常用命令已保存到本机文件";
  }).catch(error => { status.textContent = error instanceof Error ? error.message : "常用命令文件保存失败"; })
    .finally(() => { commandsBusy = false; updateButtons(); });
};
function reportDevelopmentHost(supported: boolean, selectionReady: boolean): void {
  // Local diagnostics contain no addresses, workbook data or credentials.
  const endpoint = process.env.NODE_ENV === "development" ? "/__dev/office-ready" : "/__local/ready";
  void fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ host: "Excel", supported, selectionReady }) }).catch(() => undefined);
}
function append(label: string, text: string): void {
  const item = document.createElement("div"), title = document.createElement("strong"), content = document.createElement("p");
  title.textContent = label; content.textContent = text; item.append(title, content); messages.append(item);
  while (messages.childElementCount > 60) messages.firstElementChild?.remove();
  messages.scrollTop = messages.scrollHeight;
}
function updateModelStopwatch(): void {
  if (modelStopwatchStartedAt === undefined) return;
  stopwatch.textContent = `· 模型处理 ${((performance.now() - modelStopwatchStartedAt) / 1000).toFixed(1)} 秒`;
  stopwatch.hidden = false;
}
function startModelStopwatch(): void {
  if (modelStopwatchTimer !== undefined) clearInterval(modelStopwatchTimer);
  modelStopwatchStartedAt = performance.now();
  updateModelStopwatch();
  modelStopwatchTimer = setInterval(updateModelStopwatch, 100);
}
function stopModelStopwatch(): void {
  if (modelStopwatchTimer !== undefined) clearInterval(modelStopwatchTimer);
  modelStopwatchTimer = undefined;
  modelStopwatchStartedAt = undefined;
  stopwatch.textContent = "";
  stopwatch.hidden = true;
}
function updateButtons(): void {
  sendButton.disabled = !ready || busy || commandsBusy || Boolean(settingsPending); undoButton.disabled = !ready || busy || commandsBusy || Boolean(workflow.runner) || !canUndo();
  commonCommandsButton.disabled = !ready || busy || commandsBusy; saveCommonCommandButton.disabled = commonCommandsButton.disabled; commandList.setDisabled(commonCommandsButton.disabled);
  modelsButton.disabled = busy || connectionLoading || Boolean(settingsPending); cancelButton.hidden = !controller;
  clearConversation.disabled = busy || !ready;
  provider.disabled = baseUrl.disabled = model.disabled = apiKey.disabled = busy || connectionLoading;
  rememberKey.disabled = saveSettings.disabled = forgetKey.disabled = busy || connectionLoading;
  modelPicker.disabled = busy || connectionLoading;
  cancelButton.textContent = workflowRunning ? "暂停" : "取消";
  sendButton.textContent = "发送";
  prompt.placeholder = workflow.runner ? "可询问当前数据；修改请在任务面板继续、重做，或结束任务后提出（Ctrl+Enter 发送）" : "输入分析问题或修改要求，可连续追问（Ctrl+Enter 发送）";
  workflow.render(busy, true);
}
async function refreshSelection(): Promise<void> {
  // Only read location metadata on clicks; cell data is captured once on send.
  await Excel.run(async context => {
    const range = context.workbook.getSelectedRange(); range.load("address"); await context.sync();
    selectionLabel.textContent = range.address;
  });
}
async function send(): Promise<void> {
  if (busy || commandsBusy || settingsPending || !ready || !prompt.value.trim()) return;
  const text = prompt.value.trim(), connection = config();
  busy = true; controller = new AbortController(); updateButtons();
  const started = performance.now();
  let executed = false;
  let source: SelectedRange | undefined;
  let targets: ReturnType<typeof operationTargets> = [];
  let attemptedOperations: readonly ExcelOperation[] = [];
  let failureStage = "读取选区";
  const history = conversation.snapshot();
  let readOnly = Boolean(workflow.runner);
  try {
    status.textContent = "读取选区…";
    // Start with metadata so formatting requests never pull the selected cell
    // contents merely to decide whether a format read is needed.
    let selected = await readSelection("metadata");
    source = selected;
    const captured = performance.now();
    if (controller.signal.aborted) throw new Error("已取消；未修改工作簿");
    append(`你 · ${selected.sheet}!${selected.address}`, text);
    status.textContent = `已捕获 ${selected.bounds.rows * selected.bounds.cols} 格；首轮只发送选区信息…`;
    let requestRound = 0;
    const progress: Progress = (phase, message) => {
      if (phase === "feedback") { status.textContent = `模型：${message}`; return; }
      const round = `第 ${requestRound} 轮 · `;
      status.textContent = round + (phase === "connecting" ? "正在提交请求… 可取消"
        : phase === "waiting" ? "模型处理中… 可取消" : phase === "received" ? "已收到模型响应，正在解析…"
        : phase === "validating" ? "正在校验模型结果…" : "响应格式异常，正在重新请求… 可取消");
    };
    status.textContent = selected.dataMode === "metadata" || selected.formatMode === "metadata" ? "让模型判断是否需要读取数据或格式… 可取消" : "等待模型… 可取消";
    const ask = async (current: typeof selected) => {
      failureStage = "请求模型";
      requestRound++;
      const turn = await askAgent(text, current, connection, controller!.signal, undefined, progress, history, readOnly);
      readOnly ||= turn.readOnly === true;
      if (readOnly && (turn.needsWorkflow || turn.operations.length)) throw new Error("本轮只读请求不能修改表格或创建分步任务");
      return turn;
    };
    let dataRead = selected.dataMode === "full", formatRead = selected.formatMode === "full";
    startModelStopwatch();
    let plan = await ask(selected);
    for (let attempt = 0; attempt < 3 && (plan.needsData || plan.needsFormat); attempt++) {
      const needData = plan.needsData && !dataRead;
      const needFormat = plan.needsFormat && !formatRead;
      if (!needData && !needFormat) break;
      status.textContent = `模型需要补充${needData ? "单元格数据" : ""}${needData && needFormat ? "和" : ""}${needFormat ? "颜色/格式" : ""}${plan.dataReason || plan.formatReason ? `（${plan.dataReason ?? plan.formatReason}）` : ""}…`;
      if (needData) {
        selected = await readSelection("full", selected, needFormat || formatRead);
      } else if (needFormat) {
        const previous = selected;
        const formatted = await readSelection("format", selected);
        selected = previous.dataMode === "full" ? { ...formatted, values: previous.values, formulas: previous.formulas, dataMode: "full" as const } : formatted;
      }
      dataRead = selected.dataMode === "full"; formatRead = selected.formatMode === "full";
      if (controller.signal.aborted) throw new Error("已取消；未修改工作簿");
      plan = await ask(selected);
    }
    if (plan.needsData || plan.needsFormat) throw new Error("模型仍无法完成数据或格式读取请求；未修改工作簿");
    if (plan.needsWorkflow) {
      append("助手", `将自动分步处理：${plan.workflowReason}`);
      workflowRunning = true; updateButtons();
      await workflow.plan(text, controller!.signal, selected, history);
      if (workflow.runner) rememberWorkflow(workflow.runner.task);
      prompt.value = "";
      status.textContent = workflow.runner?.task.status === "completed" ? "任务已完成" : "任务已暂停，可在步骤列表继续";
      return;
    }
    const answered = performance.now();
    stopModelStopwatch();
    if (plan.operations.length) {
      failureStage = "备份并执行操作";
      status.textContent = `正在备份并执行 ${plan.operations.length} 项操作…`;
      const currentPlan = executionPlan(plan);
      attemptedOperations = currentPlan.operations;
      targets = operationTargets(currentPlan);
      const warning = await applyOperations(currentPlan, selected, { signal: controller!.signal,
        onProgress: progress => { status.textContent = `正在分批处理 ${progress.completed}/${progress.total} 格… 可取消，取消后恢复本轮修改`; },
      });
      executed = true;
      append("助手 · 已执行", plan.answer);
      if (warning) append("提示", warning);
    } else append("助手", plan.answer);
    conversation.record(text, selected, executed ? "applied" : "answered_without_changes", plan.answer, targets);
    prompt.value = "";
    const seconds = (ms: number) => `${(ms / 1000).toFixed(1)}秒`;
    status.textContent = `${executed ? "执行完成" : "回答完成"} · 选区 ${seconds(captured - started)} · 模型 ${seconds(answered - captured)} · 执行 ${seconds(performance.now() - answered)}`;
  } catch (error) {
    void reportClientError({ stage: failureStage, error: error instanceof Error ? error.message : "操作失败", selected: source ? { sheet: source.sheet, address: source.address } : undefined, operations: attemptedOperations });
    if (workflow.runner) rememberWorkflow(workflow.runner.task);
    conversation.record(text, source, error instanceof RecoveryError ? "recovery_required" : executed ? "applied" : "stopped_not_completed", error instanceof Error ? error.message : "操作失败", targets);
    append(error instanceof RecoveryError ? "需要恢复" : "提示", error instanceof Error ? error.message : "操作失败，请重试");
    status.textContent = executed ? "修改已完成" : "本轮已停止";
  } finally { stopModelStopwatch(); controller = undefined; workflowRunning = false; busy = false; updateButtons(); }
}

sendButton.onclick = () => void send();
clearConversation.onclick = () => {
  if (busy) return;
  conversation.clear(); messages.replaceChildren();
  append("提示", "对话已清除，下一条消息从新对话开始。表格内容、恢复记录和分步任务进度保持不变。");
  status.textContent = "新对话已就绪"; updateButtons();
};
cancelButton.onclick = () => { controller?.abort(); status.textContent = workflowRunning ? "正在暂停；未完成检查的步骤保留备份，继续时先恢复再重试" : "正在取消…"; };
undoButton.onclick = () => {
  if (busy || !canUndo()) return;
  busy = true; updateButtons(); status.textContent = "正在恢复…";
  void undoLast().then(warning => { conversation.record("恢复上一步", undefined, "restored", "最近一轮表格修改已恢复，历史执行结果已不代表当前内容"); append("已恢复", warning ?? "已还原上次操作前的选区和本轮新建工作表。"); status.textContent = "就绪"; })
    .catch((error: Error) => { void reportClientError({ stage: "恢复上一步", error: error.message }); conversation.record("恢复上一步", undefined, "restore_incomplete", error.message); append("恢复未完成", error.message); status.textContent = "恢复记录已保留"; })
    .finally(() => { busy = false; updateButtons(); });
};
modelsButton.onclick = () => {
  if (busy || connectionLoading || settingsPending) return;
  busy = true; updateButtons(); status.textContent = "获取模型列表…";
  modelListStatus.textContent = "正在请求接口模型列表…";
  void listModels(config()).then(ids => {
    modelPicker.replaceChildren();
    const placeholder = document.createElement("option"); placeholder.value = ""; placeholder.textContent = `选择模型（接口返回 ${ids.length} 个）`; modelPicker.append(placeholder);
    for (const id of ids) { const option = document.createElement("option"); option.value = id; option.textContent = id; modelPicker.append(option); }
    modelPicker.value = ids.includes(model.value.trim()) ? model.value.trim() : "";
    modelPickerLabel.hidden = ids.length === 0;
    const description = ids.length === 0 ? "接口未返回模型列表，仍可手动输入模型名称。" : ids.length === 1 ? "接口仅返回 1 个模型；其他模型是否可用以中转实际支持为准，也可手动输入。" : `接口返回 ${ids.length} 个模型，下拉列表显示全部，也可手动输入。`;
    modelListStatus.textContent = description;
    status.textContent = `模型列表已更新（${ids.length} 个）`;
  }).catch((error: Error) => { append("提示", error.message); modelListStatus.textContent = "获取失败；可重试或手动输入模型名称。"; status.textContent = "就绪"; })
    .finally(() => { busy = false; updateButtons(); });
};
prompt.addEventListener("keydown", event => { if (event.ctrlKey && event.key === "Enter") { event.preventDefault(); void send(); } });
updateButtons();
Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) { status.textContent = "请在 Excel 中打开此加载项"; return; }
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) { reportDevelopmentHost(false, false); status.textContent = "需要支持 ExcelApi 1.13 的 Microsoft 365，请先更新 Excel"; return; }
  reportDevelopmentHost(true, false);
  if (__HOST_SMOKE_TEST__) {
    status.textContent = "正在验证加载项…";
    let result: { passed: boolean; detail: string };
    try { result = { passed: true, detail: await (await import("../excel/host-smoke")).runOfficeSmokeTest() }; }
    catch (error) { result = { passed: false, detail: error instanceof Error ? error.message : "宿主检查失败" }; }
    await fetch("/__dev/office-test-result", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(result) }).catch(() => undefined);
    status.textContent = result.passed ? "就绪" : "加载项验证未通过";
  }
  await restoreConnection();
  ready = true; updateButtons();
  await refreshCommonCommands();
  append("助手", "直接输入分析问题或表格修改要求，可连续追问。只询问或分析时直接回答；要求修改时执行操作，必要时自动分步处理。点击“清除对话”开始新对话。直接处理的恢复记录仅在当前窗格有效；分步进度随工作簿保存。");
  try { workflow.load(); if (workflow.runner) rememberWorkflow(workflow.runner.task); updateButtons(); }
  catch (error) { append("任务进度", error instanceof Error ? error.message : "无法读取任务进度"); }
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    await Excel.run(async context => {
      context.workbook.onSelectionChanged.add(async () => {
        clearTimeout(timer); timer = setTimeout(() => { void refreshSelection().catch(() => { selectionLabel.textContent = "请选择连续区域"; }); }, 200);
      });
      await context.sync();
    });
  } catch { append("提示", "选区自动跟踪暂不可用，发送时仍会读取当前选区。"); }
  try { await refreshSelection(); reportDevelopmentHost(true, true); }
  catch { selectionLabel.textContent = "请选择连续区域"; }
});
