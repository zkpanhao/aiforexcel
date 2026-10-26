import { test, afterEach } from "node:test";
import assert from "node:assert/strict";
import { askAgent, askAnalysis, askStructured } from "../src/llm/client";
import { parseAddress } from "../src/tools/protocol";
import { readChatResponse } from "../src/llm/stream";

const originalFetch = globalThis.fetch;
afterEach(() => { globalThis.fetch = originalFetch; });
const connection = { baseUrl: "https://example.com/v1", model: "test", apiKey: "test-only" };
const selected = { workbook: "test", sheet: "Sheet1", sheetId: "id", address: "A1", bounds: parseAddress("A1"), values: [[1]], formulas: [[1]], dataMode: "full" as const };
const encoder = new TextEncoder();
function frame(content: string): string { return `data: ${JSON.stringify({ choices: [{ index: 0, delta: { content } }] })}\r\n\r\n`; }
const end = 'data: {"choices":[{"index":0,"delta":{},"finish_reason":"stop"}]}\n\ndata: [DONE]\n\n';
const tick = () => new Promise(resolve => setTimeout(resolve, 0));

test("model feedback reaches the caller before completion across fragmented SSE and UTF-8", async () => {
  let source!: ReadableStreamDefaultController<Uint8Array>;
  globalThis.fetch = (async (_url, init) => {
    assert.equal(JSON.parse(String(init?.body)).stream, true);
    return new Response(new ReadableStream<Uint8Array>({ start(controller) { source = controller; } }), { headers: { "Content-Type": "text/event-stream; charset=utf-8" } });
  }) as typeof fetch;
  const feedback: string[] = [];
  let completed = false;
  const pending = askAgent("汇总", selected, connection, undefined, undefined, (phase, message) => { if (phase === "feedback") feedback.push(message!); }).then(value => { completed = true; return value; });
  await tick();
  // Provider reasoning is never surfaced as a progress message.
  source.enqueue(encoder.encode('data: {"choices":[{"index":0,"delta":{"reasoning_content":"PRIVATE"}}]}\n\n'));
  const bytes = encoder.encode(frame('{"progress":"正在核对字段"}\n'));
  for (let i = 0; i < bytes.length; i += 7) source.enqueue(bytes.slice(i, i + 7));
  await tick();
  assert.deepEqual(feedback, ["正在核对字段"]);
  assert.equal(completed, false);
  source.enqueue(encoder.encode(frame('{"progress":"正在生成汇总方案"}\n')));
  await tick();
  assert.deepEqual(feedback, ["正在核对字段", "正在生成汇总方案"]);
  assert.equal(completed, false);
  source.enqueue(encoder.encode(frame('{"answer":"方案就绪","operations":[]}') + end));
  source.close();
  assert.equal((await pending).answer, "方案就绪");
});

test("analysis and workflow requests receive feedback and retain strict final schemas", async () => {
  for (const kind of ["analysis", "workflow"]) {
    const feedback: string[] = [];
    const result = kind === "analysis" ? '{"answer":"结论","operations":[]}' : '{"passed":true,"summary":"核对完成"}';
    globalThis.fetch = (async () => new Response(frame('{"progress":"正在检查结果"}\n') + frame(result) + end, { headers: { "content-type": "text/event-stream" } })) as typeof fetch;
    const progress = (phase: string, message?: string) => { if (phase === "feedback") feedback.push(message!); };
    if (kind === "analysis") assert.equal((await askAnalysis("分析", selected, connection, [], undefined, progress)).answer, "结论");
    else assert.deepEqual(await askStructured("检查", {}, connection, undefined, progress), { passed: true, summary: "核对完成" });
    assert.deepEqual(feedback, ["正在检查结果"]);
  }
});

test("interrupted, refused, malformed, or invalid streamed results cannot yield an operation plan", async () => {
  for (const text of [
    frame('{"answer":"incomplete"'),
    'data: PRIVATE-invalid-event\n\n',
    frame('{"answer":"ok","operations":[{"kind":"run_script"}]}') + end,
    frame('{"answer":"ok","operations":[]}') + 'data: {"choices":[{"delta":{"refusal":"refused"},"finish_reason":"stop"}]}\n\n',
  ]) {
    globalThis.fetch = (async () => new Response(text, { headers: { "content-type": "text/event-stream" } })) as typeof fetch;
    await assert.rejects(askAgent("处理", selected, connection), error => error instanceof Error && !error.message.includes("PRIVATE"));
  }
});

test("cancelling a pending stream stops reading without another request", async () => {
  const controller = new AbortController();
  let cancelled = false, calls = 0;
  globalThis.fetch = (async () => {
    calls++;
    return new Response(new ReadableStream({ cancel() { cancelled = true; } }), { headers: { "content-type": "text/event-stream" } });
  }) as typeof fetch;
  const pending = askAgent("处理", selected, connection, controller.signal);
  await tick(); controller.abort();
  await assert.rejects(pending, /已取消/);
  assert.equal(cancelled, true); assert.equal(calls, 1);
});

test("a provider returning ordinary JSON remains compatible with progress prefixes", async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: { content: '{"progress":"正在核对"}\n{"answer":"ok","operations":[]}' } }] }))) as typeof fetch;
  assert.equal((await askAgent("处理", selected, connection)).answer, "ok");
});

test("SSE reasoning and framing beyond the former 4 MB cap still return a short valid answer", async () => {
  const reasoning = `data: ${JSON.stringify({ choices: [{ index: 0, delta: { reasoning_content: "PRIVATE".repeat(2000) } }] })}\n\n`;
  let index = 0;
  globalThis.fetch = (async () => new Response(new ReadableStream<Uint8Array>({
    pull(controller) {
      if (index++ < 400) controller.enqueue(encoder.encode(reasoning));
      else { controller.enqueue(encoder.encode(frame('{"answer":"ok","operations":[]}') + end)); controller.close(); }
    },
  }), { headers: { "content-type": "text/event-stream" } })) as typeof fetch;
  assert.equal((await askAgent("处理", selected, connection)).answer, "ok");
});

test("ordinary JSON limits actual content independently of ignored reasoning and JSON escaping", async () => {
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: {
    reasoning_content: "PRIVATE".repeat(100000), content: '{"answer":"ok","operations":[]}',
  } }] }))) as typeof fetch;
  assert.equal((await askAgent("处理", selected, connection)).answer, "ok");
  const content = JSON.stringify({ answer: '"'.repeat(180000), operations: [] });
  const wrapped = JSON.stringify({ choices: [{ message: { content } }] });
  assert.ok(content.length < 500000 && wrapped.length > 600000);
  globalThis.fetch = (async () => new Response(wrapped)) as typeof fetch;
  assert.equal((await askAgent("处理", selected, connection)).answer.length, 180000);
});

test("oversized final content is rejected immediately and cancels an unfinished stream", async () => {
  let cancelled = false;
  const response = new Response(new ReadableStream<Uint8Array>({
    start(controller) { controller.enqueue(encoder.encode(frame("x".repeat(600001)))); },
    cancel() { cancelled = true; },
  }), { headers: { "content-type": "text/event-stream" } });
  await assert.rejects(readChatResponse(response, new AbortController().signal), /最终输出超过 600000/);
  assert.equal(cancelled, true);
  globalThis.fetch = (async () => new Response(JSON.stringify({ choices: [{ message: { content: "x".repeat(600001) } }] }))) as typeof fetch;
  await assert.rejects(askAgent("处理", selected, connection), /最终输出超过 600000/);
});

test("unbounded or unterminated SSE events are stopped before the transport cap", async () => {
  for (const fragment of ["data: " + "x".repeat(100000), "data: " + "x".repeat(100000) + "\n"]) {
    let cancelled = false;
    const response = new Response(new ReadableStream<Uint8Array>({
      pull(controller) { controller.enqueue(encoder.encode(fragment)); },
      cancel() { cancelled = true; },
    }), { headers: { "content-type": "text/event-stream" } });
    await assert.rejects(readChatResponse(response, new AbortController().signal), /单条流式消息过大/);
    assert.equal(cancelled, true);
  }
});

test("transport limits remain bounded for endless SSE and ordinary JSON responses", async () => {
  for (const streaming of [true, false]) {
    let cancelled = false;
    const block = encoder.encode(streaming ? `: ${"x".repeat(100000)}\n\n` : " ".repeat(100000));
    const response = new Response(new ReadableStream<Uint8Array>({
      pull(controller) { controller.enqueue(block); },
      cancel() { cancelled = true; },
    }), { headers: { "content-type": streaming ? "text/event-stream" : "application/json" } });
    await assert.rejects(readChatResponse(response, new AbortController().signal), /传输量超过 64 MiB/);
    assert.equal(cancelled, true);
  }
});

test("ordinary JSON response reading can be cancelled while awaiting the body", async () => {
  const controller = new AbortController();
  let cancelled = false;
  const response = new Response(new ReadableStream<Uint8Array>({ cancel() { cancelled = true; } }));
  const pending = readChatResponse(response, controller.signal);
  await tick(); controller.abort();
  await assert.rejects(pending, /已取消/);
  assert.equal(cancelled, true);
});
