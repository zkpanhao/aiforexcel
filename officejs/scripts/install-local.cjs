const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');
const crypto = require('node:crypto');
const { execFileSync, spawn } = require('node:child_process');
const { activate } = require('./local-server.cjs');
const source = path.resolve(__dirname, '..');
const root = path.join(process.env.USERPROFILE || '', 'AIForExcel', 'officejs');
const command = process.argv[2] || 'install';
function json(file) { return JSON.parse(fs.readFileSync(file, 'utf8')); }
function write(file, data) { fs.writeFileSync(file, JSON.stringify(data, null, 2)); }
function health(config) {
  return new Promise((resolve, reject) => {
    const request = https.get('https://localhost:3000/__local/health', { ca: fs.readFileSync(config.ca) }, response => {
      let body = '';
      response.on('data', chunk => { body += chunk; if (body.length > 4000) request.destroy(new Error('Unexpected host response')); });
      response.on('end', () => { try { const data = JSON.parse(body); if (response.statusCode !== 200 || data.app !== 'AIForExcel' || data.installId !== config.installId || data.protocol !== 1) throw new Error('Port 3000 belongs to another host'); resolve(data); } catch (error) { reject(error); } });
    });
    request.on('error', reject); request.setTimeout(2000, () => request.destroy(new Error('Local host timeout')));
  });
}
async function start(config) {
  try { return await health(config); } catch (error) { if (error.code !== 'ECONNREFUSED') throw error; }
  const log = fs.openSync(path.join(root, 'logs', 'server.log'), 'a');
  const child = spawn(path.join(root, 'runtime', 'node.exe'), [path.join(root, 'local-server.cjs'), root], { cwd: root, windowsHide: true, detached: true, stdio: ['ignore', log, log] });
  fs.closeSync(log);
  let startError; child.on('error', error => { startError = error; }); child.unref();
  for (let i = 0; i < 30; i++) {
    if (startError) throw startError;
    await new Promise(resolve => setTimeout(resolve, 200));
    try { return await health(config); } catch { /* wait for startup */ }
  }
  throw new Error('Local host failed to start. See local logs/server.log.');
}
function powershell(args) {
  execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(source, 'scripts', 'local-windows.ps1'), '-InstallRoot', root, ...args], { stdio: 'inherit', windowsHide: true });
}
async function main() {
  if (process.platform !== 'win32' || !process.env.LOCALAPPDATA) throw new Error('Windows is required.');
  if (Number(process.versions.node.split('.')[0]) !== 24) throw new Error('Use Node.js 24 to install or update this project.');
  if (command === 'status') { console.log(JSON.stringify({ root, current: json(path.join(root, 'current.json')), host: await health(json(path.join(root, 'install.json'))) }, null, 2)); return; }
  if (command === 'start') { console.log(await start(json(path.join(root, 'install.json')))); return; }
  if (command === 'catalog') { powershell(['-Action', 'Catalog']); return; }
  if (command === 'rollback') {
    const old = json(path.join(root, 'current.json'));
    if (!old.previous) throw new Error('No previous release is available.');
    activate(root, old.previous);
    fs.copyFileSync(path.join(root, 'releases', old.previous, 'manifest.xml'), path.join(root, 'catalog', 'manifest.xml'));
    try { await start(json(path.join(root, 'install.json'))); } catch (error) { write(path.join(root, 'current.json'), old); fs.copyFileSync(path.join(root, 'releases', old.release, 'manifest.xml'), path.join(root, 'catalog', 'manifest.xml')); throw error; }
    console.log('Rolled back. Reopen the task pane.'); return;
  }
  if (command !== 'install') throw new Error('Supported commands: install, start, status, rollback, catalog');
  // Always verify source before publishing a release. Never use a stale dist folder.
  function sourceHash() {
    const hash = crypto.createHash('sha256');
    for (const name of ['src', 'assets', 'scripts', 'tests', '.npmrc', 'manifest.xml', 'package.json', 'package-lock.json', 'tsconfig.json', 'webpack.config.js']) {
      function walk(file) { if (fs.statSync(file).isDirectory()) for (const entry of fs.readdirSync(file).sort()) walk(path.join(file, entry)); else hash.update(path.relative(source, file)).update(fs.readFileSync(file)); }
      walk(path.join(source, name));
    }
    return hash.digest('hex');
  }
  const before = sourceHash();
  for (const [entry, args] of [ ['typescript/bin/tsc', ['--noEmit']], ['tsx/cli', ['--test', ...fs.readdirSync(path.join(source, 'tests')).filter(name => name.endsWith('.test.ts')).map(name => `tests/${name}`)]], ['webpack-cli/bin/cli.js', ['--mode', 'production']], ['office-addin-manifest/cli.js', ['validate', 'manifest.xml']] ]) {
    execFileSync(process.execPath, [require.resolve(entry), ...args], { cwd: source, stdio: 'inherit', windowsHide: true });
  }
  if (before !== sourceHash()) throw new Error('Source changed during checks; wait for OneDrive sync to finish and retry.');
  const tls = await require('office-addin-dev-certs').getHttpsServerOptions(365);
  void tls;
  const certs = require('office-addin-dev-certs/lib/defaults');
  for (const dir of ['runtime', 'releases', 'catalog', 'logs']) fs.mkdirSync(path.join(root, dir), { recursive: true });
  const configFile = path.join(root, 'install.json');
  const config = fs.existsSync(configFile) ? json(configFile) : { installId: crypto.randomUUID() };
  Object.assign(config, { ca: certs.caCertificatePath, cert: certs.localhostCertificatePath, key: certs.localhostKeyPath });
  const node = path.join(root, 'runtime', 'node.exe');
  const licenseFile = path.join(root, 'runtime', 'LICENSE.txt');
  if (!fs.existsSync(licenseFile)) {
    const adjacentLicense = path.join(path.dirname(process.execPath), 'LICENSE');
    if (fs.existsSync(adjacentLicense)) fs.copyFileSync(adjacentLicense, licenseFile);
    else {
      const response = await fetch(`https://raw.githubusercontent.com/nodejs/node/${process.version}/LICENSE`, { signal: AbortSignal.timeout(30000) });
      if (!response.ok) throw new Error('Could not obtain the Node.js license for the bundled runtime.');
      fs.writeFileSync(licenseFile, await response.text());
    }
  }
  if (!fs.existsSync(node)) {
    fs.copyFileSync(process.execPath, node);
  }
  const version = execFileSync(node, ['--version'], { encoding: 'utf8', windowsHide: true }).trim();
  if (!version.startsWith('v24.')) throw new Error('Installed runtime must be Node.js 24.');
  const release = crypto.createHash('sha256').update(before).digest('hex').slice(0, 16);
  const releaseDir = path.join(root, 'releases', release);
  if (!fs.existsSync(releaseDir)) {
    const staging = fs.mkdtempSync(path.join(root, 'releases', 'staging-'));
    fs.cpSync(path.join(source, 'dist'), staging, { recursive: true });
    write(path.join(staging, 'release.json'), { release, sourceHash: before, createdAt: new Date().toISOString() });
    fs.renameSync(staging, releaseDir);
  }
  const old = fs.existsSync(path.join(root, 'current.json')) ? json(path.join(root, 'current.json')) : null;
  // Verify ownership of an old developer host before switching port 3000.
  powershell(['-Action', 'Migrate', '-SourceRoot', source]);
  write(configFile, config);
  fs.copyFileSync(path.join(__dirname, 'local-server.cjs'), path.join(root, 'local-server.cjs'));
  fs.copyFileSync(path.join(__dirname, 'common-commands.cjs'), path.join(root, 'common-commands.cjs'));
  fs.copyFileSync(path.join(__dirname, 'connection-settings.cjs'), path.join(root, 'connection-settings.cjs'));
  if (!old || old.release !== release) activate(root, release);
  fs.copyFileSync(path.join(releaseDir, 'manifest.xml'), path.join(root, 'catalog', 'manifest.xml'));
  try {
    const status = await start(config);
    if (status.release !== release) throw new Error('Installed release does not match host.');
    powershell(['-Action', 'Register']);
    console.log(JSON.stringify({ installed: root, release, node: version, host: status, next: 'Configure catalog, then add from Excel Home > Add-ins > Advanced > Shared Folder.' }, null, 2));
  } catch (error) {
    if (old) { write(path.join(root, 'current.json'), old); fs.copyFileSync(path.join(root, 'releases', old.release, 'manifest.xml'), path.join(root, 'catalog', 'manifest.xml')); }
    throw error;
  }
}
if (require.main === module) main().catch(error => { console.error(error.message); process.exitCode = 1; });
module.exports = { health };
