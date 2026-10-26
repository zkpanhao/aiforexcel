const https = require("node:https");
const fs = require("node:fs");
const { caCertificatePath } = require("office-addin-dev-certs/lib/defaults");
const ca = fs.readFileSync(caCertificatePath);
const paths = ["taskpane.html", "taskpane.js", "commands.html", "assets/icon-16.png", "assets/icon-32.png", "assets/icon-64.png", "assets/icon-80.png", "manifest.xml"];
async function main() {
  for (const path of paths) {
    await new Promise((resolve, reject) => {
      const request = https.get(`https://localhost:3000/${path}`, { ca }, response => {
        let bytes = 0;
        response.on("data", chunk => { bytes += chunk.length; });
        response.on("end", () => {
          if (response.statusCode !== 200 || !bytes) return reject(new Error(`${path}: HTTP ${response.statusCode}, ${bytes} bytes`));
          console.log(`${path}: 200, ${bytes} bytes (TLS verified)`); resolve();
        });
        response.on("error", reject);
      });
      request.setTimeout(10000, () => request.destroy(new Error("Local server timeout")));
      request.on("error", reject);
    });
  }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
