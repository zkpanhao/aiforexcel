// Only explicit, user-facing progress summaries are displayed. Reasoning deltas
// are ignored; partial output never becomes an executable plan.
export const FEEDBACK_INSTRUCTION = `输出过程允许先发送简短的任务进展说明，每条独占一行，格式为 {"progress":"正在做什么"}，例如 {"progress":"正在核对汇总口径"}。在任务阶段变化时发送，使用面向用户的简短说明，不输出内部推理、思考过程、单元格原文或密钥，不声称已经修改工作簿。进展行之后输出原协议要求的一个完整最终 JSON 对象；最终 JSON 开始后不得再插入进展行。简单任务可直接输出最终 JSON。`;

// SSE framing and ignored reasoning can greatly exceed the actual answer.
// Bound transport, individual events and retained output independently.
const MAX_RESPONSE_BYTES = 64 * 1024 * 1024;
const MAX_EVENT_CHARS = 4 * 1024 * 1024;
const MAX_CONTENT_CHARS = 600000;
const transportLimit = () => new Error("接口响应传输量超过 64 MiB，已停止读取；请检查接口是否重复推送或输出了过多推理内容，本次响应未执行");
const contentLimit = () => new Error("模型最终输出超过 600000 字符，请分批生成结果，避免返回整表明细；本次响应未执行");

class Output {
  private prefix = "";
  private final = false;
  content = "";
  constructor(private feedback?: (message: string) => void) {}
  add(text: string): void {
    if (this.content.length + this.prefix.length + text.length > MAX_CONTENT_CHARS) throw contentLimit();
    if (this.final) { this.content += text; return; }
    this.prefix += text;
    let newline: number;
    while ((newline = this.prefix.indexOf("\n")) !== -1) {
      const line = this.prefix.slice(0, newline).trim();
      if (!line) { this.prefix = this.prefix.slice(newline + 1); continue; }
      let value: unknown;
      try { value = JSON.parse(line); } catch { /* final JSON may span lines */ }
      if (value && typeof value === "object" && Object.keys(value).length === 1 && "progress" in value && typeof value.progress === "string") {
        const message = value.progress.replace(/[\u0000-\u001f\u007f]/g, " ").trim().slice(0, 160);
        if (message) this.feedback?.(message);
        this.prefix = this.prefix.slice(newline + 1);
      } else {
        this.final = true; this.content += this.prefix; this.prefix = ""; break;
      }
    }
  }
  finish(): string { return this.content + this.prefix; }
}

async function readJsonText(response: Response, signal: AbortSignal): Promise<string> {
  if (!response.body) throw new Error("接口未返回可读取的响应；本次响应未执行");
  const reader = response.body.getReader(), decoder = new TextDecoder();
  let text = "", size = 0;
  const abort = () => { void reader.cancel().catch(() => undefined); };
  signal.addEventListener("abort", abort, { once: true });
  try {
    while (true) {
      if (signal.aborted) throw new Error("请求已取消；本次响应未执行");
      const part = await reader.read();
      if (signal.aborted) throw new Error("请求已取消；本次响应未执行");
      if (part.done) return text + decoder.decode();
      size += part.value.byteLength;
      if (size > MAX_RESPONSE_BYTES) throw transportLimit();
      text += decoder.decode(part.value, { stream: true });
    }
  } finally { signal.removeEventListener("abort", abort); await reader.cancel().catch(() => undefined); reader.releaseLock(); }
}

export async function readChatResponse(response: Response, signal: AbortSignal, feedback?: (message: string) => void): Promise<unknown> {
  const output = new Output(feedback);
  if (!response.headers.get("content-type")?.includes("text/event-stream")) {
    const text = await readJsonText(response, signal);
    let body;
    try { body = JSON.parse(text); } catch { throw new Error("接口返回的响应不是有效 JSON，请检查接口地址和服务状态；未修改工作簿"); }
    const message = body?.choices?.[0]?.message;
    if (typeof message?.content === "string") { output.add(message.content); message.content = output.finish(); }
    return body;
  }
  if (!response.body) throw new Error("接口未返回可读取的响应流；未修改工作簿");
  const reader = response.body.getReader(), decoder = new TextDecoder();
  let buffer = "", event = "", size = 0, finish: string | undefined, refusal = "", done = false;
  const abort = () => { void reader.cancel().catch(() => undefined); };
  signal.addEventListener("abort", abort, { once: true });
  const dispatch = () => {
    const data = event.trim(); event = "";
    if (!data) return;
    if (data === "[DONE]") { done = true; return; }
    let chunk;
    try { chunk = JSON.parse(data); } catch { throw new Error("接口流式响应格式无效；未修改工作簿"); }
    if (chunk.error) throw new Error("模型流式请求失败；未修改工作簿");
    const choice = chunk.choices?.find((item: { index?: number }) => item.index === 0 || item.index === undefined);
    if (!choice) return;
    if (typeof choice.delta?.content === "string") output.add(choice.delta.content);
    if (typeof choice.delta?.refusal === "string") refusal += choice.delta.refusal;
    if (typeof choice.finish_reason === "string") finish = choice.finish_reason;
  };
  const line = (value: string) => {
    if (value.length + event.length > MAX_EVENT_CHARS) throw new Error("接口单条流式消息过大，请检查接口响应格式；本次响应未执行");
    if (!value) dispatch();
    else if (value.startsWith("data:")) event += `${event ? "\n" : ""}${value.slice(5).trimStart()}`;
  };
  try {
    while (!done) {
      if (signal.aborted) throw new Error("请求已取消；未修改工作簿");
      const part = await reader.read();
      if (signal.aborted) throw new Error("请求已取消；未修改工作簿");
      if (part.done) { buffer += decoder.decode(); break; }
      size += part.value.byteLength;
      if (size > MAX_RESPONSE_BYTES) throw transportLimit();
      buffer += decoder.decode(part.value, { stream: true });
      let newline: number;
      while (!done && (newline = buffer.indexOf("\n")) !== -1) {
        line(buffer.slice(0, newline).replace(/\r$/, "")); buffer = buffer.slice(newline + 1);
      }
      if (!done && buffer.length + event.length > MAX_EVENT_CHARS) throw new Error("接口单条流式消息过大，请检查接口响应格式；本次响应未执行");
    }
    if (!done) { if (buffer) line(buffer.replace(/\r$/, "")); dispatch(); }
    if (!finish) throw new Error("模型响应流中断或缺少结束标记；未修改工作簿");
    const content = output.finish();
    return { choices: [{ finish_reason: finish, message: { content, ...(refusal ? { refusal } : {}) } }] };
  } finally { signal.removeEventListener("abort", abort); await reader.cancel().catch(() => undefined); reader.releaseLock(); }
}
