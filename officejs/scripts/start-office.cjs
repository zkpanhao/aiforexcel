const path = require("node:path");
const fs = require("node:fs");
const https = require("node:https");
const { spawn } = require("node:child_process");
const { getHttpsServerOptions } = require("office-addin-dev-certs");
const { startDebugging, AppType } = require("office-addin-debugging");
const { OfficeApp } = require("office-addin-manifest");
const { saveDevServerProcessId } = require("office-addin-debugging/lib/debugInfo");
const root = path.resolve(__dirname, "..");
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

function probe(ca) {
  return new Promise((resolve, reject) => {
    const request = https.get("https://localhost:3000/taskpane.html", { ca }, response => {
      let body = "";
      response.on("data", chunk => { body += chunk; if (body.length > 20000) request.destroy(new Error("Unexpected response on port 3000")); });
      response.on("end", () => response.statusCode === 200 && /<title>AI for Excel<\/title>/.test(body) ? resolve(true) : reject(new Error("Port 3000 is serving another page; close that server first.")));
      response.on("error", reject);
    });
    request.setTimeout(3000, () => request.destroy(new Error("Local HTTPS server did not respond")));
    request.on("error", error => error.code === "ECONNREFUSED" || error.code === "ECONNRESET" ? resolve(false) : reject(error));
  });
}

async function main() {
  const tls = await getHttpsServerOptions();
  if (!(await probe(tls.ca))) {
    fs.mkdirSync(path.join(root, "work"), { recursive: true });
    const output = fs.openSync(path.join(root, "work", "dev-server.log"), "a");
    const child = spawn(process.execPath, [require.resolve("webpack-cli/bin/cli.js"), "serve", "--mode", "development"], {
      cwd: root, detached: true, windowsHide: true, stdio: ["ignore", output, output], env: process.env,
    });
    fs.closeSync(output);
    let startError;
    child.on("error", error => { startError = error; });
    child.unref();
    let online = false;
    for (let n = 0; n < 80; n++) {
      if (startError) throw startError;
      await delay(250);
      // Webpack may accept TLS while its first compilation is still running.
      if (await probe(tls.ca).catch(() => false)) { online = true; break; }
    }
    if (!online) { child.kill(); throw new Error("Server start timed out; see work/dev-server.log"); }
    await saveDevServerProcessId(child.pid);
    console.log("Local HTTPS server is ready.");
  }
  await startDebugging(path.join(root, "manifest.xml"), {
    appType: AppType.Desktop, app: OfficeApp.Excel, enableDebugging: false, enableLiveReload: false,
    enableSideload: !process.argv.includes("--no-sideload"),
  });
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
