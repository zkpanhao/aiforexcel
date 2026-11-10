import { test, afterEach } from "node:test";
import assert from "node:assert/strict";
import { askAgent, askAnalysis, endpoint, executionPlan, listModels, parseTurn, PROVIDER_PRESETS, selectionPayload } from "../src/llm/client";
import { parseAddress } from "../src/tools/protocol";
const originalFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = originalFetch; });
const config = { baseUrl: "https://example.com/v1", model: "test-model", apiKey: "test-only-not-a-real-key" };
const selected = { workbook: "test", sheet: "Sheet1", sheetId: "id", address: "A1", bounds: parseAddress("A1"), values: [[1]], formulas: [[1]], dataMode: "full" as const };
test("built-in providers use their documented compatible API base URLs", () => {
  assert.deepEqual(PROVIDER_PRESETS.filter(item => item.id !== "custom").map(item => [item.id, item.baseUrl]), [
    ["openai", "https://api.openai.com/v1"],
    ["deepseek", "https://api.deepseek.com"],
    ["qwen", "https://dashscope.aliyuncs.com/compatible-mode/v1"],
    ["glm", "https://open.bigmodel.cn/api/paas/v4"],
    ["minimax", "https://api.minimax.io/v1"],
  ]);
});
test("endpoints use configured base and reject credentials or insecure remote URL", () => {
  assert.equal(endpoint(config.baseUrl, "models"), "https://example.com/v1/models");
  for (const url of ["http://example.com", "https://user:key@example.com", "https://example.com?key=x"]) assert.throws(() => endpoint(url, "models"));
});
test("send selected data, use exact configured model, avoid unsupported temperature", async () => {
  globalThis.fetch = (async (url, init) => {
    assert.equal(url, "https://example.com/v1/chat/completions");
    const body = JSON.parse(String(init?.body));
    assert.equal(body.model, config.model); assert.equal(body.temperature, undefined);
    const payload = JSON.parse(body.messages[1].content);
    assert.equal(payload.selection.sheet, selected.sheet); assert.equal(payload.selection.address, selected.address);
    assert.equal(payload.selection.dataMode, "full"); assert.deepEqual(payload.selection.values, selected.values); assert.equal(payload.userTask, "说明数据");
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"1","operations":[]}' } }] }));
  }) as typeof fetch;
  assert.equal((await askAgent("说明数据", selected, config)).answer, "1");
});
test("provider preset routes requests through that provider base URL", async () => {
  const deepseek = { provider: "deepseek" as const, baseUrl: "https://api.deepseek.com", model: "deepseek-chat", apiKey: "test-only-not-a-real-key" };
  globalThis.fetch = (async (url, init) => {
    assert.equal(url, "https://api.deepseek.com/chat/completions");
    assert.equal(JSON.parse(String(init?.body)).model, "deepseek-chat");
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"ok","operations":[]}' } }] }));
  }) as typeof fetch;
  assert.equal((await askAgent("说明数据", selected, deepseek)).answer, "ok");
});
test("analysis mode keeps conversation context and rejects write operations", async () => {
  globalThis.fetch = (async (_url, init) => {
    const body = JSON.parse(String(init?.body));
    assert.equal(body.messages[0].role, "system");
    assert.equal(body.messages[1].content, "上一轮分析结论");
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"继续分析后结论"}' } }] }));
  }) as typeof fetch;
  const turn = await askAnalysis("继续分析", selected, config, [{ role: "assistant", content: "上一轮分析结论" }]);
  assert.equal(turn.answer, "继续分析后结论");
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"不应修改","operations":[{"kind":"clear_contents","sheet":"Sheet1","address":"A1"}]}' } }] }))) as typeof fetch;
  await assert.rejects(askAnalysis("清除", selected, config), /不会修改工作簿/);
});
test("model list is live, not hard-coded", async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({ data: [{ id: "future-model" }, { id: "gpt-6.1-sol" }] }))) as typeof fetch;
  assert.deepEqual(await listModels(config), ["future-model", "gpt-6.1-sol"]);
});
test("large selections send metadata only on the first turn", () => {
  const values = Array.from({ length: 100 }, (_, row) => Array.from({ length: 20 }, (_, col) => `${row}:${col}`));
  const large = { ...selected, address: "A1:T100", bounds: { row: 0, col: 0, rows: 100, cols: 20 }, values, formulas: values, dataMode: "metadata" as const };
  const payload = selectionPayload(large) as any;
  assert.equal(payload.dataMode, "metadata");
  assert.equal(payload.cellCount, 2000);
  assert.equal("values" in payload, false);
  assert.equal("formulas" in payload, false);
});
test("selection payload includes per-cell formats only after format capture", () => {
  const metadata = { ...selected, formatMode: "metadata" as const };
  assert.equal((selectionPayload(metadata) as any).formatMode, "metadata");
  assert.equal("formats" in (selectionPayload(metadata) as any), false);
  const full = { ...selected, formatMode: "full" as const, formats: [[{ fill: "#00FF00", fontColor: "#000000", bold: false, italic: false, underline: "None", numberFormat: "0.00" }]] };
  assert.deepEqual((selectionPayload(full) as any).formats[0][0], full.formats[0][0]);
  const formatOnly = { ...selected, values: [], formulas: [], dataMode: "metadata" as const, formatMode: "full" as const, formats: full.formats };
  const formatPayload = selectionPayload(formatOnly) as any;
  assert.equal(formatPayload.dataMode, "metadata");
  assert.equal("values" in formatPayload, false);
  assert.deepEqual(formatPayload.formats[0][0], full.formats[0][0]);
  const column = { ...metadata, selectionKind: "column" as const, selectionLabel: "D:D", address: "D1:D266" };
  const columnPayload = selectionPayload(column) as any;
  assert.equal(columnPayload.kind, "column");
  assert.equal(columnPayload.label, "D:D");
  assert.equal(columnPayload.address, "D1:D266");
});
test("selection payload includes supplemental range values and formulas", () => {
  const payload = selectionPayload({ ...selected, dataMode: "metadata", supplementalRanges: [{ ...selected, address: "L273", bounds: parseAddress("L273"), values: [["2026-11-01"]], formulas: [[""]] }] }) as any;
  assert.equal(payload.dataMode, "metadata");
  assert.equal(payload.supplemental_ranges[0].sheet, selected.sheet);
  assert.equal(payload.supplemental_ranges[0].address, "L273");
  assert.deepEqual(payload.supplemental_ranges[0].values, [["2026-11-01"]]);
});
test("parseTurn accepts a data request and rejects mixed request plus operations", () => {
  const turn = parseTurn('{"answer":"需要读取内容","needs_data":true,"data_request":{"reason":"需要查找具体值"},"operations":[]}');
  assert.equal(turn.needsData, true);
  assert.equal(turn.dataReason, "需要查找具体值");
  assert.throws(() => parseTurn('{"answer":"读取","needs_data":true,"data_request":{"reason":"查找"},"operations":[{"kind":"clear_contents","sheet":"Sheet1","address":"A1"}]}'), /不能同时执行操作/);
});
test("parseTurn accepts explicit supplemental data ranges", () => {
  const turn = parseTurn('{"answer":"需要读取指定区域","needs_data":true,"data_request":{"reason":"检查 L273 和日期行","ranges":[{"sheet":"预估收入目标","address":"L273"},{"sheet":"预估收入目标","address":"A2:BN2"}]},"operations":[]}');
  assert.deepEqual(turn.dataRanges, [
    { sheet: "预估收入目标", address: "L273" },
    { sheet: "预估收入目标", address: "A2:BN2" },
  ]);
});
test("parseTurn rejects invalid or oversized supplemental data ranges", () => {
  assert.throws(() => parseTurn('{"answer":"读取","needs_data":true,"data_request":{"reason":"查找","ranges":[{"sheet":"Sheet1","address":"A0"}]},"operations":[]}'), /连续区域/);
  assert.throws(() => parseTurn('{"answer":"读取","needs_data":true,"data_request":{"reason":"查找","ranges":[{"sheet":"Sheet1","address":"A1:CV100000"}]},"operations":[]}'), /100000/);
});
test("parseTurn validates data_request reason", () => {
  assert.throws(() => parseTurn('{"answer":"读取","needs_data":true,"data_request":{"reason":1},"operations":[]}'), /data_request/);
});

test("parseTurn tolerates omitted empty fields for data requests and read-only answers", () => {
  const request = parseTurn('{"needs_data":true,"data_request":{"reason":"查找浙江门店"}}');
  assert.equal(request.answer, "");
  assert.equal(request.operations.length, 0);
  const answer = parseTurn('{"answer":"浙江门店A下降最多"}');
  assert.equal(answer.answer, "浙江门店A下降最多");
  assert.equal(answer.operations.length, 0);
  const format = parseTurn('{"needs_format":true,"format_request":{"reason":"识别绿色填充"}}');
  assert.equal(format.needsFormat, true);
  assert.equal(format.formatReason, "识别绿色填充");
  assert.equal(format.operations.length, 0);
});
test("provider errors never echo response secrets", async () => {
  globalThis.fetch = (async () => new Response("SECRET-test", { status: 401 })) as typeof fetch;
  await assert.rejects(listModels(config), error => error instanceof Error && error.message.includes("401") && !error.message.includes("SECRET"));
});
test("cancellation rejects without returning an operation plan", async () => {
  const controller = new AbortController(); controller.abort();
  globalThis.fetch = (async (_url, init) => { assert.equal(init?.signal?.aborted, true); throw new DOMException("Aborted", "AbortError"); }) as typeof fetch;
  await assert.rejects(askAgent("清空", selected, config, controller.signal), /已取消/);
});

test("accepts one JSON object inside prose, fences or a closed reasoning block", () => {
  const json = JSON.stringify({ answer: '颜色 {标签} 与 "边框"', operations: [{ kind: "format_range", sheet: "Sheet1", address: "A1", fill: "#D9E2F3", borders: [{ edges: "All", weight: "Thin" }] }] });
  for (const raw of [json, `说明：\n${json}\n完成。`, `以下是结果：\n\`\`\`json\n${json}\n\`\`\``, `<think>分析 {不是操作}</think>\n${json}`]) {
    assert.equal(parseTurn(raw).operations[0].kind, "format_range");
  }
});

test("does not repair malformed syntax, extract nested partial plans or select among multiple plans", () => {
  const json = '{"answer":"ok","operations":[]}';
  for (const raw of [`${json}\n${json}`, '{"answer":"ok","operations":[],}', `{"partial":${json}`, '<think>unfinished', "{'answer':'ok','operations':[]}"]) {
    assert.throws(() => parseTurn(raw));
  }
  assert.throws(() => parseTurn(`说明 {"answer":"ok","operations":[{"kind":"run_script","code":"anything"}]}`));
});

test("syntax failure regenerates once with original task, never echoes bad output", async () => {
  let calls = 0;
  globalThis.fetch = (async (_url, init) => {
    const body = JSON.parse(String(init?.body));
    assert.equal(JSON.parse(body.messages[1].content).userTask, "按标题、标签、数据增加边框和颜色");
    assert.equal(JSON.stringify(body).includes("PRIVATE-invalid-response"), false);
    calls++;
    if (calls === 2) assert.match(body.messages[0].content, /上次响应无法解析/);
    const content = calls === 1 ? "PRIVATE-invalid-response" : '{"answer":"已设置","operations":[{"kind":"format_range","sheet":"Sheet1","address":"A1","fill":"#D9E2F3","borders":[{"edges":"All","weight":"Thin"}]}]}';
    return new Response(JSON.stringify({ choices: [{ finish_reason: "stop", message: { content } }] }));
  }) as typeof fetch;
  assert.equal((await askAgent("按标题、标签、数据增加边框和颜色", selected, config)).operations.length, 1);
  assert.equal(calls, 2);
});

test("invalid second response stops after exactly two requests", async () => {
  let calls = 0;
  globalThis.fetch = (async () => { calls++; return new Response(JSON.stringify({ choices: [{ message: { content: "invalid" } }] })); }) as typeof fetch;
  await assert.rejects(askAgent("美化", selected, config), /JSON 格式无效/);
  assert.equal(calls, 2);
});

test("truncation, refusal and invalid operations stop without syntax retry", async () => {
  for (const choice of [
    { finish_reason: "length", message: { content: '{"answer":' } },
    { finish_reason: "content_filter", message: { content: "blocked" } },
    { message: { refusal: "blocked", content: "blocked" } },
    { message: { content: '{"answer":"ok","operations":[{"kind":"run_script"}]}' } },
  ]) {
    let calls = 0;
    globalThis.fetch = (async () => { calls++; return new Response(JSON.stringify({ choices: [choice] })); }) as typeof fetch;
    await assert.rejects(askAgent("美化", selected, config));
    assert.equal(calls, 1);
  }
});

test("cancellation between malformed response and retry prevents another request", async () => {
  let calls = 0;
  const controller = new AbortController();
  globalThis.fetch = (async () => { calls++; controller.abort(); return new Response(JSON.stringify({ choices: [{ message: { content: "invalid" } }] })); }) as typeof fetch;
  await assert.rejects(askAgent("美化", selected, config, controller.signal), /已取消/);
  assert.equal(calls, 1);
});

test("non JSON HTTP envelope has a sanitized distinct error", async () => {
  globalThis.fetch = (async () => new Response("SECRET-provider-page")) as typeof fetch;
  await assert.rejects(askAgent("美化", selected, config), error => error instanceof Error && error.message.includes("接口返回") && !error.message.includes("SECRET"));
});

 test("table progress reports waiting before headers arrive and validates the response", async () => {
  const phases: string[] = [];
  let respond!: (response: Response) => void;
  globalThis.fetch = (() => new Promise<Response>(resolve => { respond = resolve; })) as typeof fetch;
  const pending = askAgent("说明数据", selected, config, undefined, undefined, phase => phases.push(phase));
  assert.deepEqual(phases, ["connecting", "waiting"]);
  respond(new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"完成","operations":[]}' } }] })));
  assert.equal((await pending).answer, "完成");
  assert.deepEqual(phases, ["connecting", "waiting", "received", "validating"]);
});

test("automatic workflow routing is a non-executable decision and never mixes with data requests or writes", () => {
  const turn = parseTurn('{"needs_workflow":true,"workflow_reason":"汇总依赖计算结果"}');
  assert.equal(turn.needsWorkflow, true);
  assert.equal(turn.workflowReason, "汇总依赖计算结果");
  assert.deepEqual(turn.operations, []);
  assert.throws(() => executionPlan(turn), /先规划/);
  for (const value of [
    { needs_workflow: "true" },
    { needs_workflow: true, operations: [{ kind: "clear_contents", sheet: "Sheet1", address: "A1" }] },
    { needs_workflow: true, needs_data: true },
    { needs_workflow: true, needs_format: true },
    { needs_workflow: true, workflow_reason: {} },
  ]) assert.throws(() => parseTurn(JSON.stringify(value)));
});

test("simple processing keeps one request; a workflow step cannot recursively create another workflow", async () => {
  let calls = 0;
  globalThis.fetch = (async (_url, init) => {
    calls++;
    const messages = JSON.parse(String(init?.body)).messages;
    if (calls === 1) {
      assert.match(messages[0].content, /用户无需判断复杂程度/);
      return new Response(JSON.stringify({ choices: [{ message: { content: '{"answer":"设置格式","operations":[{"kind":"format_range","sheet":"Sheet1","address":"A1","bold":true}]}' } }] }));
    }
    assert.match(messages[0].content, /本轮只能执行当前步骤/);
    return new Response(JSON.stringify({ choices: [{ message: { content: '{"needs_workflow":true,"operations":[]}' } }] }));
  }) as typeof fetch;
  const direct = await askAgent("加粗", selected, config);
  assert.equal(direct.needsWorkflow, undefined);
  assert.equal(executionPlan(direct).operations.length, 1);
  assert.equal(calls, 1);
  await assert.rejects(askAgent("当前步", selected, config, undefined, { currentStep: 1 }), /不能再次拆分/);
  assert.equal(calls, 2);
});

test("analysis cannot escalate into a modifying workflow", async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: { content: '{"needs_workflow":true,"operations":[]}' } }] }))) as typeof fetch;
  await assert.rejects(askAnalysis("分析", selected, config), /只读请求不会修改/);
});

test("unified requests return analysis or modifications without a separate classification call", async () => {
  let calls = 0;
  globalThis.fetch = (async (_url, init) => {
    calls++;
    const messages = JSON.parse(String(init?.body)).messages;
    assert.match(messages[0].content, /无需用户选择模式/);
    assert.match(messages[0].content, /历史修改请求不构成本轮修改授权/);
    const content = calls === 1 ? { answer: "销售下降主要来自门店A", read_only: true } :
      { answer: "写入结论", read_only: false, operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", data: [["销售下降主要来自门店A"]] }] };
    return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(content) } }] }));
  }) as typeof fetch;
  const analysis = await askAgent("只分析销售下降原因，不修改", selected, config);
  assert.equal(analysis.readOnly, true); assert.deepEqual(analysis.operations, []);
  assert.equal(calls, 1);
  assert.equal((await askAgent("把结论写入A1", selected, config)).operations.length, 1);
  assert.equal(calls, 2);
});

test("read-only validation rejects writes and workflows, including an omitted read_only flag under a forced restriction", async () => {
  const write = { kind: "clear_contents", sheet: "Sheet1", address: "A1" };
  assert.throws(() => parseTurn(JSON.stringify({ answer: "分析", read_only: "true" })), /read_only/);
  for (const value of [
    { answer: "分析", read_only: true, operations: [write] },
    { answer: "分析", read_only: true, needs_workflow: true },
  ]) assert.throws(() => parseTurn(JSON.stringify(value)), /只读请求/);
  assert.equal(parseTurn('{"read_only":true,"needs_data":true}').needsData, true);
  assert.equal(parseTurn('{"read_only":true,"needs_format":true}').needsFormat, true);
  assert.throws(() => executionPlan({ answer: "分析", needsData: false, needsFormat: false, readOnly: true, operations: [write] } as any), /只读请求/);
  for (const result of [{ answer: "错误修改", operations: [write] }, { needs_workflow: true }]) {
    globalThis.fetch = (async (_url, init) => {
      assert.match(JSON.parse(String(init?.body)).messages[0].content, /本轮由插件限定只读/);
      return new Response(JSON.stringify({ choices: [{ message: { content: JSON.stringify(result) } }] }));
    }) as typeof fetch;
    await assert.rejects(askAgent("分析", selected, config, undefined, undefined, undefined, [], true), /只读请求不会修改/);
  }
});
