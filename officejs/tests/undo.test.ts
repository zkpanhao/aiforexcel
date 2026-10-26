import { test } from "node:test";
import assert from "node:assert/strict";
import { canUndo, setLastUndo, undoLast } from "../src/state/undo";
test("failed undo retains recovery; successful undo clears it and disposes backup", async () => {
  let fail = true, disposed = 0;
  setLastUndo({ restore: async () => { if (fail) throw new Error("protected"); }, dispose: async () => { disposed++; } });
  await assert.rejects(undoLast(), /protected/);
  assert.equal(canUndo(), true); assert.equal(disposed, 0);
  fail = false; await undoLast();
  assert.equal(canUndo(), false); assert.equal(disposed, 1);
});
test("cleanup failure does not report a successful restore as failed", async () => {
  setLastUndo({ restore: async () => undefined, dispose: async () => { throw new Error("protected"); } });
  assert.match((await undoLast())!, /已恢复/);
  assert.equal(canUndo(), false);
});
