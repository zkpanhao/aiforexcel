import { test } from "node:test";
import assert from "node:assert/strict";
import { userFacingError } from "../src/state/user-error";

test("incomplete cross-range writes explain how to recapture target and reference areas", () => {
  const message = userFacingError(new Error("写入数据与目标区域的行列数不一致"), { sheet: "即时零售门店对照", address: "D2:E969" });
  assert.match(message, /即时零售门店对照!D2:E969/);
  assert.match(message, /参考工作表和连续地址/);
  assert.match(message, /发送后切换工作表不会改变/);
});

test("unrelated errors keep their original message", () => {
  assert.equal(userFacingError(new Error("找不到工作表")), "找不到工作表");
});
