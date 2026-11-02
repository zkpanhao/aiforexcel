import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import ts from "typescript";
import * as client from "../src/llm/client";
import * as settings from "../src/state/settings";
import * as conversation from "../src/state/conversation";
import { parseAddress } from "../src/tools/protocol";
import { CommonCommandList } from '../src/taskpane/common-command-list';
import { userFacingError } from "../src/state/user-error";

// Exercise the actual pane handlers with an isolated DOM and workbook adapter.
// No browser package, model requests, user storage or real workbook is involved.
class Element {
  value = ""; textContent = ""; hidden = false; disabled = false; checked = false;
  children: Element[] = []; scrollTop = 0; scrollHeight = 0;
  onclick?: () => void;
  onfocus?: () => void;
  title = ''; type = ''; className = ''; style = { left: '', top: '' };
  attributes: Record<string, string> = {};
  listeners: Record<string, (event?: any) => void> = {};
  addEventListener(name: string, callback: (event?: any) => void) { this.listeners[name] = callback; }
  append(...children: Element[]) { this.children.push(...children); }
  replaceChildren() { this.children = []; }
  focus() { this.onfocus?.(); }
  click() { if (!this.disabled) this.onclick?.(); }
  setAttribute(key: string, value: string) { this.attributes[key] = value; }
  contains(target: Element): boolean { return this === target || this.children.some(child => child.contains(target)); }
  getBoundingClientRect() { return { left: 10, top: 10, bottom: 40, width: 120, height: 60 }; }
  get childElementCount() { return this.children.length; }
}
type CommandFiles = { readFileCommands: () => Promise<string[]>; saveFileCommand: (text: string) => Promise<string[]>;
  updateFileCommand?: (command: string, replacement: string) => Promise<string[]>; deleteFileCommand?: (command: string) => Promise<string[]> };
async function pane(commandFiles: CommandFiles = { readFileCommands: async () => [] as string[], saveFileCommand: async (text: string) => [text] },
  connectionFile = { readFileConnection: async (): Promise<settings.ConnectionSettings | undefined> => undefined, saveFileConnection: async (_value: settings.ConnectionSettings) => {}, forgetFileKey: async () => {} }) {
  const elements = new Map<string, Element>();
  let active: Element | undefined;
  const createElement = () => { const e = new Element(); e.onfocus = () => { active = e; }; return e; };
  const element = (id: string) => { if (!elements.has(id)) elements.set(id, createElement()); return elements.get(id)!; };
  const documentEvents: Record<string, (event?: any) => void> = {};
  element("baseUrl").value = "https://example.com/v1";
  element("model").value = "test-model"; element("apiKey").value = "test-only";
  let ready!: (info: unknown) => Promise<void>;
  const calls: { readOnly: boolean; task: string; selected: any; history: client.ConversationMessage[] }[] = [];
  const state = {
    selected: { workbook: "test", sheet: "数据", sheetId: "id", address: "A1", bounds: parseAddress("A1"), values: [[1]], formulas: [[1]], dataMode: "full" as const },
    failApply: false, failUndo: false, undo: false, applied: 0, restored: 0, analysisReply: false, turns: [] as any[],
  };
  let releaseRequest: (() => void) | undefined;
  let hold = false;
  async function answer(readOnly: boolean, task: string, selected: any, history: client.ConversationMessage[]) {
    calls.push({ readOnly, task, selected, history });
    if (hold) await new Promise<void>(resolve => { releaseRequest = resolve; });
    if (state.turns.length) return state.turns.shift();
    const analysis = readOnly || state.analysisReply;
    return { answer: analysis ? "分析结论" : "已加粗", readOnly: analysis, needsData: false, needsFormat: false,
      operations: analysis ? [] : [{ kind: "format_range", sheet: selected.sheet, address: selected.address, bold: true }] };
  }
  class RecoveryError extends Error {}
  class WorkflowController {
    runner?: any;
    constructor() { controller = this; }
    render() {} load() {}
  }
  let controller!: WorkflowController;
  const modules: Record<string, unknown> = {
    "./taskpane.css": {},
    "../excel/selection": { readSelection: async () => state.selected, readSupplementalRanges: async (ranges: { sheet: string; address: string }[]) => ranges.map(range => ({ ...state.selected, sheet: range.sheet, sheetId: range.sheet, address: range.address, bounds: parseAddress(range.address), values: [[`${range.sheet}!${range.address}`]], formulas: [[""]] })) },
    "../excel/operations": { RecoveryError, applyOperations: async () => {
      if (state.failApply) throw new Error("写入失败，已回退");
      state.undo = true; state.applied++;
    } },
    "../llm/client": { ...client,
      askAgent: (task: string, selected: any, _config: unknown, _signal: unknown, _context: unknown, _progress: unknown, history: client.ConversationMessage[], readOnly: boolean) => answer(readOnly, task, selected, history),
    },
    "../state/undo": { canUndo: () => state.undo, undoLast: async () => {
      if (state.failUndo) throw new Error("恢复失败");
      state.undo = false; state.restored++;
    } },
    "../state/settings": settings, "../state/conversation": conversation,
    "../state/client-log": { reportClientError: async () => {} },
    "../state/user-error": { userFacingError },
    "../state/common-commands": { updateFileCommand: async (_command: string, replacement: string) => [replacement], deleteFileCommand: async () => [], ...commandFiles },
    './common-command-list': { CommonCommandList },
    "../state/connection-file": connectionFile,
    "./workflow-controller": { WorkflowController },
  };
  const source = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.ts"), "utf8");
  const code = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
  vm.runInNewContext(code, {
    exports: {}, require: (name: string) => { assert.ok(name in modules, name); return modules[name]; },
    document: { getElementById: element, createElement, get activeElement() { return active; }, addEventListener: (type: string, action: (event?: any) => void) => { documentEvents[type] = action; } },
    window: { localStorage: { getItem: () => null }, innerWidth: 320, innerHeight: 600, addEventListener: () => {} },
    Office: { onReady: (callback: typeof ready) => { ready = callback; }, HostType: { Excel: "Excel" }, context: { requirements: { isSetSupported: () => true } } },
    Excel: { run: async (action: (context: unknown) => Promise<void>) => action({ sync: async () => {}, workbook: {
      onSelectionChanged: { add: () => {} }, getSelectedRange: () => ({ load: () => {}, address: "数据!A1" }),
    } }) },
    __HOST_SMOKE_TEST__: false, process: { env: { NODE_ENV: "production" } },
    fetch: async () => ({}), Error, AbortController, performance, setInterval, clearInterval, setTimeout, clearTimeout,
  });
  await ready({ host: "Excel" });
  const idle = async () => {
    for (let i = 0; i < 50 && element("provider").disabled; i++) await new Promise(resolve => setImmediate(resolve));
    assert.equal(element("provider").disabled, false, "pane should finish the action");
  };
  const send = async (text: string) => { element("prompt").value = text; element("send").onclick!(); await idle(); };
  return { state, calls, element, send, idle, workflow: controller, documentEvents,
    hold: () => { hold = true; }, release: () => { hold = false; releaseRequest!(); } };
}

test("pane confirms file save only after success and reloads commands without clearing prompt or undo", async () => {
  let commands = ["原命令"], fail = false;
  const h = await pane({ readFileCommands: async () => commands, saveFileCommand: async text => {
    if (fail) throw new Error("文件无法写入，未确认保存");
    commands = [text, ...commands]; return commands;
  } });
  const settle = async () => { for (let i = 0; i < 10; i++) await new Promise(resolve => setImmediate(resolve)); };
  await h.send("加粗");
  h.element("prompt").value = "待保存命令";
  h.element("saveCommonCommand").onclick!();
  assert.equal(h.element("send").disabled, true);
  await settle();
  assert.match(h.element("status").textContent, /已保存到本机文件/);
  assert.equal(h.element("prompt").value, "待保存命令");
  assert.equal(h.state.undo, true);
  commands = ["文件中新加的命令"];
  h.element("commonCommandList").hidden = true;
  h.element("commonCommands").onclick!(); await settle();
  assert.equal(h.element("commonCommandList").children[0].textContent, "文件中新加的命令");
  fail = true;
  h.element("saveCommonCommand").onclick!(); await settle();
  assert.match(h.element("status").textContent, /未确认保存/);
  assert.equal(h.element("prompt").value, "待保存命令");
  assert.equal(h.state.undo, true);
});

test("pane welcome message explains common conversation examples", async () => {
  const h = await pane();
  const welcome = h.element("messages").children[0].children[1].textContent;
  assert.match(welcome, /写公式：在 O 列计算 M 列占 N 列的比例/);
  assert.match(welcome, /把 A 列城市所属的省份填入 B 列/);
  assert.match(welcome, /销售额大于 10 万的整行标为浅黄色/);
  assert.match(welcome, /按日期从早到晚排序当前表格/);
});

test('right-click edits/deletes the clicked command, preserves draft/undo and persists exact target', async () => {
  let commands = ['第一条', '第二条\n多行'];
  const updates: string[][] = [], deletes: string[] = [];
  const h = await pane({ readFileCommands: async () => [...commands], saveFileCommand: async text => [text],
    updateFileCommand: async (old, replacement) => { updates.push([old, replacement]); commands = commands.map(text => text === old ? replacement : text); return [...commands]; },
    deleteFileCommand: async command => { deletes.push(command); commands = commands.filter(text => text !== command); return [...commands]; } });
  const settle = async () => { for (let i = 0; i < 10; i++) await new Promise(resolve => setImmediate(resolve)); };
  await h.send('加粗'); h.element('prompt').value = '尚未发送的草稿';
  h.element('commonCommands').click(); await settle();
  const second = h.element('commonCommandList').children[1];
  let prevented = false;
  second.listeners.contextmenu({ clientX: 319, clientY: 599, preventDefault: () => { prevented = true; } });
  assert.equal(prevented, true); assert.equal(h.element('commandMenu').hidden, false);
  assert.ok(parseInt(h.element('commandMenu').style.left) <= 196);
  assert.ok(parseInt(h.element('commandMenu').style.top) <= 536);
  assert.equal(h.element('prompt').value, '尚未发送的草稿');
  h.element('commandEdit').click();
  assert.equal(h.element('commandEditText').value, '第二条\n多行');
  h.element('commandEditText').value = '修改后的\n多行'; h.element('commandEditSave').click();
  await settle();
  assert.deepEqual(updates, [['第二条\n多行', '修改后的\n多行']]);
  assert.equal(h.element('commandEditor').hidden, true);
  assert.deepEqual(commands, ['第一条', '修改后的\n多行']);
  const first = h.element('commonCommandList').children[0];
  first.listeners.keydown({ key: 'F10', shiftKey: true, preventDefault: () => {} });
  h.element('commandDelete').click(); await settle();
  assert.deepEqual(deletes, ['第一条']); assert.deepEqual(commands, ['修改后的\n多行']);
  assert.equal(h.element('prompt').value, '尚未发送的草稿'); assert.equal(h.state.undo, true);
  h.element('commonCommandList').children[0].click();
  assert.equal(h.element('prompt').value, '尚未发送的草稿\n修改后的\n多行');
});

test('command menu dismiss/cancel/validation never saves and failed edits/deletes keep list and edit draft', async () => {
  let changes = 0;
  const h = await pane({ readFileCommands: async () => ['保留命令'], saveFileCommand: async text => [text],
    updateFileCommand: async () => { changes++; throw new Error('文件保存失败'); }, deleteFileCommand: async () => { changes++; throw new Error('文件删除失败'); } });
  const settle = async () => { for (let i = 0; i < 10; i++) await new Promise(resolve => setImmediate(resolve)); };
  h.element('commonCommands').click(); await settle();
  const row = h.element('commonCommandList').children[0];
  const open = () => row.listeners.contextmenu({ clientX: 20, clientY: 30, preventDefault: () => {} });
  open(); h.documentEvents.pointerdown({ target: h.element('prompt') }); assert.equal(h.element('commandMenu').hidden, true);
  open(); h.element('commandMenu').listeners.keydown({ key: 'Escape', preventDefault: () => {} }); assert.equal(h.element('commandMenu').hidden, true);
  open(); h.element('commandEdit').click(); h.element('commandEditCancel').click(); assert.equal(changes, 0);
  open(); h.element('commandEdit').click(); h.element('commandEditText').value = ' '; h.element('commandEditSave').click();
  assert.match(h.element('commandError').textContent, /不能为空/); assert.equal(changes, 0);
  h.element('commandEditText').value = '不能丢掉的修改草稿'; h.element('commandEditSave').click(); await settle();
  assert.equal(h.element('commandEditor').hidden, false); assert.equal(h.element('commandEditText').value, '不能丢掉的修改草稿');
  assert.match(h.element('commandError').textContent, /保存失败/);
  open(); h.element('commandDelete').click(); await settle();
  assert.match(h.element('commandError').textContent, /删除失败/); assert.equal(h.element('commonCommandList').children[0].title, '保留命令');
  assert.equal(changes, 2);
});

test("pane loads file connection before use, clears key on provider change and reports failed forget honestly", async () => {
  const loaded: settings.ConnectionSettings = { provider: 'custom', baseUrl: 'https://example.com/v1', model: 'file-model', apiKey: 'test-only-pane-key', rememberKey: true };
  const saved: settings.ConnectionSettings[] = [];
  let failForget = true;
  const h = await pane(undefined, { readFileConnection: async () => loaded,
    saveFileConnection: async value => { saved.push(value); }, forgetFileKey: async () => { if (failForget) throw new Error('文件未清除，未确认清除'); } });
  const settle = async () => { for (let i = 0; i < 10; i++) await new Promise(resolve => setImmediate(resolve)); };
  assert.equal(h.element('model').value, 'file-model'); assert.equal(h.element('apiKey').value, loaded.apiKey);
  h.element('model').value = 'new-model'; h.element('model').listeners.change(); await settle();
  assert.equal(saved[0].model, 'new-model'); assert.match(h.element('settingsStatus').textContent, /已保存到本机文件/);
  h.element('provider').value = 'deepseek'; h.element('provider').listeners.change();
  assert.equal(h.element('apiKey').value, ''); assert.equal(h.element('rememberKey').checked, false);
  await settle();
  assert.equal(saved.at(-1)?.baseUrl, 'https://api.deepseek.com'); assert.equal(saved.at(-1)?.apiKey, '');
  h.element('forgetKey').onclick!(); await settle();
  assert.match(h.element('settingsStatus').textContent, /未确认清除/);
  failForget = false; h.element('forgetKey').onclick!(); await settle();
  assert.match(h.element('settingsStatus').textContent, /已清除本机文件及备份/);
});

test("one pane alternates modifications and analysis with continuous context and fresh selections", async () => {
  const h = await pane();
  await h.send("加粗");
  h.state.selected = { ...h.state.selected, address: "B2", bounds: parseAddress("B2"), values: [[5]], formulas: [[5]] };
  await h.send("再处理这里");
  assert.equal(h.calls[1].selected.address, "B2");
  assert.equal(JSON.parse(h.calls[1].history[1].content).status, "applied");
  assert.equal(JSON.parse(h.calls[1].history[1].content).targets[0].address, "A1");
  h.state.analysisReply = true; await h.send("解释刚才的操作");
  assert.equal(h.state.applied, 2);
  assert.equal(h.element("undo").disabled, false);
  assert.equal(h.calls[2].history.length, 4);
  h.state.analysisReply = false; await h.send("继续修改");
  assert.equal(JSON.parse(h.calls[3].history[5].content).status, "answered_without_changes");
});

test("clear button resets messages and context, preserves input/undo/workflow, and is disabled while busy", async () => {
  const h = await pane();
  await h.send("加粗");
  h.element("prompt").value = "尚未发送";
  const task = { task: { status: "paused" } };
  h.workflow.runner = task;
  h.element("clearConversation").onclick!();
  assert.equal(h.element("messages").childElementCount, 1);
  assert.equal(h.element("prompt").value, "尚未发送");
  assert.equal(h.state.undo, true); assert.equal(h.state.applied, 1); assert.equal(h.state.restored, 0);
  assert.equal(h.workflow.runner, task);
  h.workflow.runner = undefined;
  await h.send("新任务"); assert.equal(h.calls[1].history.length, 0);
  h.hold();
  const pending = h.send("等待中的任务");
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(h.element("clearConversation").disabled, true);
  const count = h.element("messages").childElementCount;
  h.element("clearConversation").onclick!();
  assert.equal(h.element("messages").childElementCount, count);
  h.release(); await pending;
  await h.send("继续"); assert.equal(h.calls[3].history.length, 4);
});

test("pane records actual failed and restored outcomes instead of the model's unexecuted success answer", async () => {
  const h = await pane();
  h.state.failApply = true; await h.send("加粗");
  h.state.analysisReply = true; await h.send("是否完成");
  let outcome = JSON.parse(h.calls[1].history[1].content);
  assert.equal(outcome.status, "stopped_not_completed"); assert.doesNotMatch(outcome.summary, /已加粗/);
  h.state.failApply = false; h.state.analysisReply = false; await h.send("重试");
  h.state.failUndo = true; h.element("undo").onclick!(); await h.idle();
  h.state.failUndo = false; h.element("undo").onclick!(); await h.idle();
  h.state.analysisReply = true; await h.send("现在呢");
  const statuses = h.calls[3].history.filter(item => item.role === "assistant").map(item => JSON.parse(item.content).status);
  assert.deepEqual(statuses, ["stopped_not_completed", "answered_without_changes", "applied", "restore_incomplete", "restored"]);
});

test("existing workflow keeps its panel and permits questions while blocking new writes or planning", async () => {
  const h = await pane();
  const task = { task: { task: "原任务", anchor: { sheet: "数据", address: "A1" }, summary: "原任务", status: "paused", steps: [] } };
  h.workflow.runner = task;
  await h.send("解释当前结果");
  assert.equal(h.calls[0].readOnly, true);
  assert.equal(h.element("send").disabled, false);
  assert.equal(h.workflow.runner, task); assert.equal(h.state.applied, 0);
  h.state.turns.push({ answer: "错误写入", operations: [{ kind: "clear_contents", sheet: "数据", address: "A1" }] });
  await h.send("试图写入");
  assert.equal(h.state.applied, 0); assert.equal(h.workflow.runner, task);
  h.state.turns.push({ answer: "错误规划", needsWorkflow: true, operations: [] });
  await h.send("新任务");
  assert.equal(h.state.applied, 0); assert.equal(h.workflow.runner, task);
});

test("read-only decision stays locked through data supplementation and resets for the next user request", async () => {
  const h = await pane();
  h.state.selected = { ...h.state.selected, dataMode: "metadata" as any };
  h.state.turns.push(
    { answer: "需要数据", readOnly: true, needsData: true, operations: [] },
    { answer: "不能在补充后写入", operations: [{ kind: "clear_contents", sheet: "数据", address: "A1" }] },
  );
  await h.send("分析原因");
  assert.equal(h.calls.length, 2); assert.equal(h.calls[0].readOnly, false); assert.equal(h.calls[1].readOnly, true);
  assert.equal(h.state.applied, 0);
  await h.send("现在修改");
  assert.equal(h.calls[2].readOnly, false); assert.equal(h.state.applied, 1);
});
test("model can request and receive supplemental ranges outside the current selection", async () => {
  const h = await pane();
  h.state.selected = { ...h.state.selected, values: [], formulas: [], dataMode: "metadata" as any };
  h.state.turns.push(
    { answer: "需要指定区域", needsData: true, dataRanges: [{ sheet: "数据", address: "L273" }], operations: [] },
    { answer: "已检查 L273", needsData: false, operations: [] },
  );
  await h.send("检查 L273");
  assert.equal(h.calls.length, 2);
  assert.equal(h.calls[1].selected.supplementalRanges[0].address, "L273");
  assert.deepEqual(h.calls[1].selected.supplementalRanges[0].values, [["数据!L273"]]);
});

test("unified HTML contains the clear button and no mode selector", () => {
  const html = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.html"), "utf8");
  assert.doesNotMatch(html, /id="mode"|工作模式/);
  assert.doesNotMatch(html, /<strong>AI for Excel<\/strong>/);
  assert.doesNotMatch(html, /<header/);
  assert.match(html, /class="conversation-actions"><span id="selection"[^>]*>.*<button id="clearConversation"/s);
  assert.match(html, /id="clearConversation"[^>]*>清除对话/);
  assert.match(html, /id="undo"[^>]*>恢复<\/button>/);
  assert.match(html, /id="commonCommands"[^>]*>常用<\/button>/);
  assert.match(html, /id="saveCommonCommand"[^>]*>保存<\/button>/);
  assert.doesNotMatch(html, /id="undo"[^>]*>恢复上一步<\/button>/);
  assert.doesNotMatch(html, /id="commonCommands"[^>]*>常用命令<\/button>/);
  assert.doesNotMatch(html, /id="saveCommonCommand"[^>]*>保存常用命令<\/button>/);
});
