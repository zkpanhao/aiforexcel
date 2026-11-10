import { test } from "node:test";
import assert from "node:assert/strict";
import { canonicalBaseUrl, forgetSavedKey, loadCommonCommands, loadConnection, saveCommonCommand, saveConnection } from "../src/state/settings";

function memoryStorage() {
  const values = new Map<string, string>();
  return { values, getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value); }, removeItem: (key: string) => { values.delete(key); } };
}
const settings = { provider: "custom" as const, baseUrl: "https://example.com/v1", model: "test-model", apiKey: "test-only-credential", rememberKey: true };
test("reopening settings restores remembered key, endpoint and model", () => {
  const storage = memoryStorage(); saveConnection(storage, settings);
  assert.deepEqual(loadConnection(storage), settings);
});
test("key is never persisted without opt-in", () => {
  const storage = memoryStorage(); saveConnection(storage, { ...settings, rememberKey: false });
  assert.equal(loadConnection(storage)?.apiKey, "");
  assert.ok([...storage.values.values()].every(value => !value.includes(settings.apiKey)));
});
test("unchecking remember removes an existing key immediately", () => {
  const storage = memoryStorage(); saveConnection(storage, settings);
  saveConnection(storage, { ...settings, rememberKey: false });
  assert.equal(loadConnection(storage)?.rememberKey, false);
  assert.ok([...storage.values.values()].every(value => !value.includes(settings.apiKey)));
});
test("forgetting keeps address and model but removes key and consent", () => {
  const storage = memoryStorage(); saveConnection(storage, settings); forgetSavedKey(storage);
  assert.deepEqual(loadConnection(storage), { ...settings, apiKey: "", rememberKey: false });
  assert.ok([...storage.values.values()].every(value => !value.includes(settings.apiKey)));
});
test("a key stored for a different endpoint is not restored", () => {
  const storage = memoryStorage(); saveConnection(storage, settings);
  const key = [...storage.values.keys()][0];
  const stored = JSON.parse(storage.values.get(key)!); stored.baseUrl = "https://other.example/v1";
  storage.setItem(key, JSON.stringify(stored));
  assert.equal(loadConnection(storage)?.apiKey, "");
});
test("canonical trailing slash preserves endpoint scope", () => {
  const storage = memoryStorage(); saveConnection(storage, { ...settings, baseUrl: `${settings.baseUrl}/` });
  assert.deepEqual(loadConnection(storage), settings);
  assert.equal(canonicalBaseUrl("https://EXAMPLE.com/v1/"), "https://example.com/v1");
});
test("corrupt settings are ignored and can be cleared", () => {
  const storage = memoryStorage(); saveConnection(storage, settings);
  const key = [...storage.values.keys()][0]; storage.setItem(key, "{");
  assert.equal(loadConnection(storage), undefined); forgetSavedKey(storage);
  assert.equal(storage.values.size, 0);
});
test("storage failures propagate so UI cannot claim success", () => {
  const storage = { ...memoryStorage(), setItem: () => { throw new Error("blocked"); } };
  assert.throws(() => saveConnection(storage, settings), /blocked/);
});
test("common commands are saved newest first and duplicate commands move to the front", () => {
  const storage = memoryStorage();
  assert.deepEqual(saveCommonCommand(storage, "按金额降序"), ["按金额降序"]);
  assert.deepEqual(saveCommonCommand(storage, "标记异常"), ["标记异常", "按金额降序"]);
  assert.deepEqual(saveCommonCommand(storage, "按金额降序"), ["按金额降序", "标记异常"]);
  assert.deepEqual(loadCommonCommands(storage), ["按金额降序", "标记异常"]);
});
test("blank common commands are rejected and corrupt command storage is ignored", () => {
  const storage = memoryStorage();
  assert.throws(() => saveCommonCommand(storage, "  \n  "), /输入命令/);
  storage.setItem("ai-for-excel.common-commands.v1", "{");
  assert.deepEqual(loadCommonCommands(storage), []);
});
