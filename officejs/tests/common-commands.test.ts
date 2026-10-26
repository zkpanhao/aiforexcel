import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import https from "node:https";
import { readFileCommands, saveFileCommand, updateFileCommand, deleteFileCommand } from "../src/state/common-commands";
const { filePath, readCommands, addCommand, migrateCommands, changeCommand, handleCommonCommands } = require("../scripts/common-commands.cjs");
const mkcert = require("mkcert");
const legacyKey = "ai-for-excel.common-commands.v1";

function temporary() { return fs.mkdtempSync(path.join(os.tmpdir(), "ai-excel-commands-")); }

test("command file initializes legacy data once, preserves saved order and backs up changes", () => {
  const root = temporary();
  try {
    assert.equal(readCommands(root).exists, false);
    migrateCommands(root, ["旧命令", " 重复 ", "重复"]);
    addCommand(root, "新命令");
    assert.deepEqual(migrateCommands(root, ["重复", "补回命令"]).commands, ["新命令", "旧命令", "重复"]);
    const original = fs.readFileSync(filePath(root), "utf8");
    migrateCommands(root, ["旧命令"]);
    assert.equal(fs.readFileSync(filePath(root), "utf8"), original);
    assert.deepEqual(JSON.parse(fs.readFileSync(filePath(root) + ".bak", "utf8")).commands, ["旧命令", "重复"]);
    assert.deepEqual(addCommand(root, " 重复 ").commands, ["重复", "新命令", "旧命令"]);
    for (let i = 0; i < 55; i++) addCommand(root, `命令${i}`);
    assert.equal(readCommands(root).commands.length, 50);
    assert.equal(readCommands(root).commands[0], "命令54");
    assert.throws(() => addCommand(root, " "));
    assert.throws(() => addCommand(root, "x".repeat(2001)));
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test("edit/delete target exact text, preserve other commands and reject stale or duplicate edits", () => {
  const root = temporary();
  try {
    migrateCommands(root, ['第一条', '第二条\n多行', '第三条']);
    assert.deepEqual(changeCommand(root, '第二条\n多行', ' 修改后的\n多行 ').commands, ['第一条', '修改后的\n多行', '第三条']);
    const saved = fs.readFileSync(filePath(root), 'utf8');
    assert.throws(() => changeCommand(root, '第一条', '第三条'), /COMMAND_EXISTS/);
    assert.throws(() => changeCommand(root, '第二条\n多行', '过期编辑'), /COMMAND_CHANGED/);
    assert.equal(fs.readFileSync(filePath(root), 'utf8'), saved);
    assert.deepEqual(changeCommand(root, '第一条').commands, ['修改后的\n多行', '第三条']);
    assert.throws(() => changeCommand(root, '第一条'), /COMMAND_CHANGED/);
    assert.deepEqual(migrateCommands(root, ['第一条', '第二条\n多行']).commands, ['修改后的\n多行', '第三条']);
    changeCommand(root, '修改后的\n多行'); changeCommand(root, '第三条');
    assert.equal(readCommands(root).exists, true); assert.deepEqual(readCommands(root).commands, []);
    assert.deepEqual(migrateCommands(root, ['第一条']).commands, []);
    assert.deepEqual(JSON.parse(fs.readFileSync(filePath(root) + '.bak', 'utf8')).commands, ['第三条']);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test("malformed command file is never replaced; failed atomic replacement preserves original", () => {
  const root = temporary(), rename = fs.renameSync;
  try {
    addCommand(root, "保留");
    const original = fs.readFileSync(filePath(root), "utf8");
    fs.renameSync = ((from, to) => {
      if (to === filePath(root)) throw new Error("simulated disk failure");
      rename(from, to);
    }) as typeof fs.renameSync;
    assert.throws(() => addCommand(root, "不能丢掉原文件"));
    assert.throws(() => changeCommand(root, "保留", "不能丢掉原文件"));
    assert.throws(() => changeCommand(root, "保留"));
    fs.renameSync = rename;
    assert.equal(fs.readFileSync(filePath(root), "utf8"), original);
    assert.equal(fs.readdirSync(path.dirname(filePath(root))).some(name => name.endsWith(".tmp")), false);
    fs.writeFileSync(filePath(root), "{broken");
    assert.throws(() => addCommand(root, "不能覆盖"));
    assert.throws(() => migrateCommands(root, []));
    assert.equal(fs.readFileSync(filePath(root), "utf8"), "{broken");
    assert.equal(fs.readFileSync(filePath(root) + ".bak", "utf8"), original);
  } finally { fs.renameSync = rename; fs.rmSync(root, { recursive: true, force: true }); }
});

test("command endpoint persists concurrent adds, enforces local origin and rejects invalid requests", async () => {
  const root = temporary();
  const ca = await mkcert.createCA({ organization: "Commands test", countryCode: "CN", locality: "Test", state: "Test", validity: 1 });
  const cert = await mkcert.createCert({ ca, domains: ["localhost", "127.0.0.1"], validity: 1 });
  const server = https.createServer({ cert: cert.cert, key: cert.key }, (req, res) => {
    if (!handleCommonCommands(root, req, res)) { res.writeHead(404); res.end(); }
  });
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  const port = (server.address() as { port: number }).port;
  function request(method = "GET", body?: unknown, headers: Record<string, string> = {}) {
    return new Promise<{ status: number; data: any }>((resolve, reject) => {
      const req = https.request({ hostname: "127.0.0.1", servername: "localhost", port, path: "/__local/common-commands", method, ca: ca.cert,
        headers: { Origin: `https://localhost:${port}`, "Content-Type": "application/json", ...headers } }, res => {
        let text = ""; res.on("data", chunk => text += chunk); res.on("end", () => resolve({ status: res.statusCode!, data: JSON.parse(text) }));
      });
      req.on("error", reject); req.end(body === undefined ? undefined : JSON.stringify(body));
    });
  }
  try {
    assert.equal((await request()).data.exists, false);
    const added = await Promise.all(Array.from({ length: 12 }, (_, i) => request("POST", { action: "add", command: `命令${i}` })));
    assert.ok(added.every(result => result.status === 200));
    assert.equal(new Set((await request()).data.commands).size, 12);
    assert.equal(readCommands(root).commands.length, 12);
    assert.equal((await request('POST', { action: 'update', command: '命令1', replacement: '修改后的命令' })).status, 200);
    assert.equal((await request('POST', { action: 'delete', command: '命令0' })).status, 200);
    assert.equal((await request('POST', { action: 'delete', command: '命令0' })).status, 409);
    assert.equal((await request('POST', { action: 'update', command: '命令1', replacement: '不能覆盖' })).status, 409);
    assert.equal((await request('POST', { action: 'update', command: '命令2', replacement: '命令3' })).status, 409);
    assert.equal(readCommands(root).commands.length, 11);
    assert.ok(readCommands(root).commands.includes('修改后的命令'));
    assert.equal((await request("GET", undefined, { Origin: "https://external.invalid" })).status, 403);
    assert.equal((await request("GET", undefined, { Host: "external.invalid" })).status, 403);
    assert.equal((await request("POST", { action: "add", command: "x" }, { Origin: "" })).status, 403);
    assert.equal((await request("POST", { action: "add", command: "x" }, { "Content-Type": "text/plain" })).status, 403);
    for (const body of [{ action: "delete" }, { action: 'delete', command: '命令2', replacement: 'x' }, { action: 'update', command: '命令2', replacement: ' ' }, { action: 'update', command: '命令2', replacement: 'x'.repeat(2001) }, { action: "add", command: "" }, { action: "add", command: "x", path: "outside" }, { action: "migrate", commands: [1] }]) {
      assert.equal((await request("POST", body)).status, 400);
    }
    assert.equal((await request("DELETE")).status, 405);
    assert.equal((await request("POST", { action: "add", command: "x".repeat(600000) })).status, 413);
    fs.writeFileSync(filePath(root), "bad file");
    assert.equal((await request()).status, 503);
    assert.equal((await request("POST", { action: "add", command: "x" })).status, 503);
  } finally {
    await new Promise<void>(resolve => server.close(() => resolve()));
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("client removes only legacy commands after durable migration, and file survives empty browser storage", async () => {
  const root = temporary(), originalFetch = globalThis.fetch;
  const values = new Map([[legacyKey, JSON.stringify(["旧命令"])], ["ai-for-excel.connection.v1", "untouched"]]);
  const storage = { getItem: (key: string) => values.get(key) ?? null, removeItem: (key: string) => { values.delete(key); } };
  globalThis.fetch = (async (_url, init) => {
    const body = init?.body ? JSON.parse(String(init.body)) : undefined;
    const data = !body ? readCommands(root) : body.action === "migrate" ? migrateCommands(root, body.commands) : body.action === 'add' ? addCommand(root, body.command) : changeCommand(root, body.command, body.action === 'update' ? body.replacement : undefined);
    return new Response(JSON.stringify(data));
  }) as typeof fetch;
  try {
    assert.deepEqual(await readFileCommands(storage), ["旧命令"]);
    assert.equal(values.has(legacyKey), false);
    assert.equal(values.get("ai-for-excel.connection.v1"), "untouched");
    assert.deepEqual(await saveFileCommand(" 新命令 ", storage), ["新命令", "旧命令"]);
    const blocked = { getItem: () => { throw new Error("browser cache unavailable"); }, removeItem: () => { throw new Error("blocked"); } };
    assert.deepEqual(await readFileCommands(blocked), ["新命令", "旧命令"]);
    assert.deepEqual(await readFileCommands({ getItem: () => null, removeItem: () => {} }), ["新命令", "旧命令"]);
    assert.deepEqual(await updateFileCommand('旧命令', '修改后的命令'), ['新命令', '修改后的命令']);
    assert.deepEqual(await deleteFileCommand('新命令'), ['修改后的命令']);
    values.set(legacyKey, JSON.stringify(['新命令', '旧命令']));
    assert.deepEqual(await readFileCommands(storage), ['修改后的命令']);
    assert.equal(values.has(legacyKey), false);
  } finally { globalThis.fetch = originalFetch; fs.rmSync(root, { recursive: true, force: true }); }
});

test('client reports edit conflicts and delete failures without claiming success', async () => {
  const originalFetch = globalThis.fetch;
  try {
    globalThis.fetch = (async () => new Response(JSON.stringify({ error: 'COMMAND_EXISTS' }), { status: 409 })) as typeof fetch;
    await assert.rejects(updateFileCommand('旧命令', '重复命令'), /已存在/);
    globalThis.fetch = (async () => new Response(JSON.stringify({ error: 'COMMAND_CHANGED' }), { status: 409 })) as typeof fetch;
    await assert.rejects(deleteFileCommand('过期命令'), /已被修改或删除/);
    globalThis.fetch = (async () => new Response('{}', { status: 503 })) as typeof fetch;
    await assert.rejects(deleteFileCommand('保留命令'), /未确认保存或删除/);
    await assert.rejects(updateFileCommand('旧命令', ' '), /不能为空/);
    await assert.rejects(updateFileCommand('旧命令', 'x'.repeat(2001)), /2000/);
  } finally { globalThis.fetch = originalFetch; }
});

test("failed migration or malformed response retains legacy data and never confirms saving", async () => {
  const originalFetch = globalThis.fetch;
  let removed = false;
  const storage = { getItem: () => JSON.stringify(["旧命令"]), removeItem: () => { removed = true; } };
  try {
    globalThis.fetch = (async (_url, init) => init?.method === "GET"
      ? new Response(JSON.stringify({ version: 1, exists: false, commands: [] }))
      : new Response("{}", { status: 503 })) as typeof fetch;
    await assert.rejects(saveFileCommand("新命令", storage), /未确认保存/);
    assert.equal(removed, false);
    globalThis.fetch = (async () => new Response(JSON.stringify({ version: 1, exists: true, commands: [123] }))) as typeof fetch;
    await assert.rejects(readFileCommands(storage), /未确认保存/);
    assert.equal(removed, false);
    await assert.rejects(saveFileCommand(" "), /输入命令/);
    await assert.rejects(saveFileCommand("x".repeat(2001)), /2000/);
  } finally { globalThis.fetch = originalFetch; }
});
