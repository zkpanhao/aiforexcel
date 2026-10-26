const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');
const os = require('node:os');
const { handleCommonCommands, readCommands } = require('./common-commands.cjs');
const { handleConnection, readConnection, filePath: connectionPath } = require('./connection-settings.cjs');
const { execFileSync, spawn } = require('node:child_process');
const { assetPath } = require('./local-server.cjs');
const source = path.resolve(__dirname, '..');
const GUID = '8264d257-5475-42b8-96bf-15fc2785c524';
async function main() {
  execFileSync(process.execPath, [require.resolve('webpack-cli/bin/cli.js'), '--config', 'webpack.native-check.cjs', '--mode', 'production'], { cwd: source, stdio: 'inherit', windowsHide: true });
  const directory = path.join(source, 'build', 'native-check');
  const manifest = path.join(directory, 'manifest.xml');
  fs.writeFileSync(manifest, fs.readFileSync(path.join(source, 'manifest.xml'), 'utf8').replace(/e6e1a25b-a2dc-4c80-9177-6c1e9e0e9d61/g, GUID).replace(/localhost:3000/g, 'localhost:3001').replace('DefaultValue="AI for Excel"', 'DefaultValue="AI for Excel 验证"'));
  const defaults = require('office-addin-dev-certs/lib/defaults');
  let finish;
  const done = new Promise(resolve => { finish = resolve; });
  const commandRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'ai-excel-native-commands-'));
  const server = https.createServer({ cert: fs.readFileSync(defaults.localhostCertificatePath), key: fs.readFileSync(defaults.localhostKeyPath) }, (req, res) => {
    if (handleCommonCommands(commandRoot, req, res)) return;
    if (handleConnection(commandRoot, req, res)) return;
    if (req.url === '/result' && req.method === 'POST') {
      if (req.headers.origin !== 'https://localhost:3001') { res.writeHead(403); return res.end(); }
      let body = '';
      req.on('data', chunk => { body += chunk; if (body.length > 10000) req.destroy(); });
      req.on('end', () => {
        try {
          const data = JSON.parse(body);
          const result = { at: new Date().toISOString(), passed: data.passed === true, completed: data.completed, error: data.error };
          if (result.passed && (!readCommands(commandRoot).exists || readCommands(commandRoot).commands.length !== 0)) {
            result.passed = false; result.error = 'Native commands did not reach the isolated JSON file';
          }
          if (result.passed && (readConnection(commandRoot).settings?.rememberKey !== false || readConnection(commandRoot).settings?.apiKey ||
            fs.readFileSync(connectionPath(commandRoot), 'utf8').includes('apiKey') || fs.readFileSync(connectionPath(commandRoot) + '.bak', 'utf8').includes('apiKey'))) {
            result.passed = false; result.error = 'Native credential forget did not clear the isolated JSON files';
          }
          fs.mkdirSync(path.join(source, 'work'), { recursive: true });
          fs.writeFileSync(path.join(source, 'work', 'native-check.json'), JSON.stringify(result, null, 2));
          res.writeHead(204); res.end(); finish(result);
        } catch { res.writeHead(400); res.end(); }
      });
      return;
    }
    const file = assetPath(directory, req.url);
    if (req.method !== 'GET' || !file || !fs.existsSync(file)) { res.writeHead(404); return res.end(); }
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Content-Type', file.endsWith('.js') ? 'text/javascript' : file.endsWith('.html') ? 'text/html; charset=utf-8' : file.endsWith('.png') ? 'image/png' : 'application/xml');
    res.end(fs.readFileSync(file));
  });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(3001, 'localhost', resolve); });
  const { registerAddIn, unregisterAddIn } = require('office-addin-dev-settings');
  let timer;
  try {
    await registerAddIn(manifest);
    const info = await require('office-addin-manifest').OfficeAddinManifest.readManifestFile(manifest);
    const workbook = await require('office-addin-dev-settings/lib/sideload').generateSideloadFile('excel', info);
    const excelExe = execFileSync('powershell.exe', ['-NoProfile', '-Command', "(Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe').'(default)'"], { encoding: 'utf8', windowsHide: true }).trim();
    const excel = spawn(excelExe, ['/x', workbook], { detached: true, windowsHide: false, stdio: 'ignore' });
    excel.unref();
    const timeout = new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('Native verification timed out; no result was received.')), 180000); });
    const result = await Promise.race([done, timeout]);
    console.log(JSON.stringify(result, null, 2));
    if (!result.passed) throw new Error('Native verification failed; see work/native-check.json');
  } finally {
    clearTimeout(timer);
    await unregisterAddIn(manifest);
    await new Promise(resolve => { server.close(resolve); server.closeAllConnections(); });
    fs.rmSync(commandRoot, { recursive: true, force: true });
  }
}
// Office tooling telemetry may retain retry timers after the test server closes.
main().then(() => process.exit(0), error => { console.error(error.message); process.exit(1); });
