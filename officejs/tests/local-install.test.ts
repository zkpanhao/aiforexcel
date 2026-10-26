import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import https from "node:https";
const { createServer, activate, assetPath, releasePath } = require("../scripts/local-server.cjs");
const mkcert = require("mkcert");

test("installed host switches complete releases, supports rollback and isolates local files", async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "ai-excel-host-test-"));
  const ca = await mkcert.createCA({ organization: "AI Excel test only", countryCode: "CN", locality: "Test", state: "Test", validity: 1 });
  const cert = await mkcert.createCert({ ca, domains: ["localhost", "127.0.0.1"], validity: 1 });
  fs.mkdirSync(path.join(root, "logs"));
  fs.writeFileSync(path.join(root, "cert.pem"), cert.cert);
  fs.writeFileSync(path.join(root, "key.pem"), cert.key);
  fs.writeFileSync(path.join(root, "install.json"), JSON.stringify({ installId: "test-install", cert: path.join(root, "cert.pem"), key: path.join(root, "key.pem") }));
  const first = "1111111111111111", second = "2222222222222222";
  for (const id of [first, second]) {
    fs.mkdirSync(releasePath(root, id), { recursive: true });
    for (const name of ["taskpane.html", "taskpane.js", "commands.html", "manifest.xml"]) fs.writeFileSync(path.join(releasePath(root, id), name), id);
  }
  assert.throws(() => activate(root, "3333333333333333"));
  assert.throws(() => releasePath(root, "../../outside"));
  for (const url of ["/../install.json", "/%2e%2e/key.pem", "/assets/../../key.pem", "/%ZZ", "/install.json", "/user-data/common-commands.json", "/user-data/common-commands.json.bak", "/user-data/connection.json", "/user-data/connection.json.bak"]) assert.equal(assetPath(root, url), null);
  activate(root, first);
  const server = createServer(root);
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  function request(url: string, method = "GET", body?: unknown, origin?: string): Promise<{ status: number; body: string }> {
    return new Promise((resolve, reject) => {
      const req = https.request({ hostname: "127.0.0.1", port, path: url, method, ca: ca.cert, headers: origin ? { Origin: origin } : {} }, response => {
        let text = ""; response.on("data", data => text += data); response.on("end", () => resolve({ status: response.statusCode!, body: text }));
      });
      req.on("error", reject); req.end(body ? JSON.stringify(body) : undefined);
    });
  }
  try {
    assert.equal((await request("/taskpane.js")).body, first);
    activate(root, second);
    assert.equal((await request("/taskpane.js")).body, second);
    assert.equal(JSON.parse((await request("/__local/health")).body).release, second);
    const previous = JSON.parse(fs.readFileSync(path.join(root, "current.json"), "utf8")).previous;
    activate(root, previous);
    assert.equal((await request("/taskpane.js")).body, first);
    assert.equal((await request("/install.json")).status, 404);
    assert.equal((await request("/taskpane.html", "POST")).status, 405);
    assert.equal((await request("/__local/ready", "POST", { host: "Excel" }, "https://example.com")).status, 403);
    assert.equal((await request("/__local/ready", "POST", { host: "Excel", supported: true, selectionReady: true, secret: "must-not-be-recorded" }, "https://localhost:3000")).status, 204);
    const ready = JSON.parse(fs.readFileSync(path.join(root, "logs", "office-ready.json"), "utf8"));
    assert.equal(ready.selectionReady, true); assert.equal(ready.release, first); assert.equal(ready.secret, undefined);
  } finally {
    await new Promise<void>(resolve => server.close(() => resolve()));
    fs.rmSync(root, { recursive: true, force: true });
  }
});
