// Installed host serves releases and same-origin per-device settings APIs.
const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');
const { X509Certificate } = require('node:crypto');
const { handleCommonCommands } = require('./common-commands.cjs');
const { handleConnection } = require('./connection-settings.cjs');

function releasePath(root, id) {
  if (!/^[a-f0-9]{16}$/.test(id)) throw new Error('Invalid release id');
  return path.join(root, 'releases', id);
}
function assetPath(directory, rawUrl) {
  let name;
  try { name = decodeURIComponent(rawUrl.split('?')[0]); } catch { return null; }
  if (name === '/') name = '/taskpane.html';
  if (!/^\/(?:taskpane\.(?:html|js)|commands\.html|manifest\.xml|assets\/[\w.-]+\.png)$/.test(name)) return null;
  return path.join(directory, name.slice(1));
}
function clipped(value, limit) {
  return typeof value === 'string' && value.trim() ? value.replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, limit) : undefined;
}
function clientLogRecord(current, data) {
  const result = { at: new Date().toISOString(), release: current.release, event: 'client-error' };
  const stage = clipped(data?.stage, 80), error = clipped(data?.error, 1000);
  if (stage) result.stage = stage;
  if (error) result.error = error;
  if (data?.selected && typeof data.selected === 'object') {
    const selected = { sheet: clipped(data.selected.sheet, 120), address: clipped(data.selected.address, 80) };
    if (selected.sheet || selected.address) result.selected = selected;
  }
  if (Array.isArray(data?.operations)) {
    result.operations = data.operations.slice(0, 20).flatMap(item => {
      if (!item || typeof item !== 'object') return [];
      const operation = { kind: clipped(item.kind, 60) };
      for (const key of ['sheet', 'address', 'lookupSheet', 'lookupAddress']) {
        const value = clipped(item[key], 120);
        if (value) operation[key] = value;
      }
      return operation.kind ? [operation] : [];
    });
  }
  return result;
}
function appendClientLog(root, record) {
  const file = path.join(root, 'logs', 'client.log');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  try {
    if (fs.statSync(file).size > 2 * 1024 * 1024) {
      const previous = `${file}.1`;
      try { fs.rmSync(previous, { force: true }); } catch { /* keep the current log if rotation fails */ }
      try { fs.renameSync(file, previous); } catch { /* keep appending to the current log */ }
    }
  } catch { /* the log is created below */ }
  fs.appendFileSync(file, `${JSON.stringify(record)}\n`, 'utf8');
}
function activate(root, release) {
  releasePath(root, release);
  for (const name of ['taskpane.html', 'taskpane.js', 'commands.html', 'manifest.xml']) {
    if (!fs.statSync(path.join(releasePath(root, release), name)).isFile()) throw new Error(`Missing ${name}`);
  }
  const previous = fs.existsSync(path.join(root, 'current.json')) ? JSON.parse(fs.readFileSync(path.join(root, 'current.json'), 'utf8')).release : null;
  const next = { release, previous: previous === release ? null : previous, updatedAt: new Date().toISOString() };
  const temporary = path.join(root, 'current.json.tmp');
  fs.writeFileSync(temporary, JSON.stringify(next, null, 2));
  fs.renameSync(temporary, path.join(root, 'current.json'));
  return next;
}
function createServer(root) {
  const config = JSON.parse(fs.readFileSync(path.join(root, 'install.json'), 'utf8'));
  const cert = fs.readFileSync(config.cert);
  if (Date.parse(new X509Certificate(cert).validTo) <= Date.now()) throw new Error('HTTPS certificate expired; run the installer to renew it.');
  const server = https.createServer({ cert, key: fs.readFileSync(config.key) }, (request, response) => {
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('X-Content-Type-Options', 'nosniff');
    if (handleCommonCommands(root, request, response)) return;
    if (handleConnection(root, request, response)) return;
    try {
      const current = JSON.parse(fs.readFileSync(path.join(root, 'current.json'), 'utf8'));
      if (request.url === '/__local/health' && request.method === 'GET') {
        response.setHeader('Content-Type', 'application/json');
        return response.end(JSON.stringify({ app: 'AIForExcel', protocol: 1, installId: config.installId, release: current.release, pid: process.pid }));
      }
      if (request.url === '/__local/ready' && request.method === 'POST') {
        if (request.headers.origin !== 'https://localhost:3000') { response.writeHead(403); return response.end(); }
        let body = '';
        request.on('data', chunk => { body += chunk; if (body.length > 1000) request.destroy(); });
        request.on('end', () => {
          try {
            const data = JSON.parse(body);
            const result = { at: new Date().toISOString(), release: current.release, host: data.host === 'Excel' ? 'Excel' : 'Other', supported: data.supported === true, selectionReady: data.selectionReady === true };
            fs.writeFileSync(path.join(root, 'logs', 'office-ready.json'), JSON.stringify(result, null, 2));
            response.writeHead(204); response.end();
          } catch { response.writeHead(400); response.end(); }
        });
        return;
      }
      if (request.url === '/__local/client-log' && request.method === 'POST') {
        if (request.headers.origin !== 'https://localhost:3000') { response.writeHead(403); return response.end(); }
        let body = '', tooLarge = false;
        request.on('data', chunk => { if (body.length < 10000) body += chunk; else tooLarge = true; });
        request.on('end', () => {
          try {
            if (tooLarge) throw new Error('log too large');
            appendClientLog(root, clientLogRecord(current, JSON.parse(body)));
            response.writeHead(204); response.end();
          } catch { response.writeHead(400); response.end(); }
        });
        return;
      }
      if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405); return response.end(); }
      const file = assetPath(releasePath(root, current.release), request.url);
      if (!file || !fs.existsSync(file)) { response.writeHead(404); return response.end(); }
      const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.xml': 'application/xml', '.png': 'image/png' };
      response.setHeader('Content-Type', types[path.extname(file)] || 'application/octet-stream');
      response.end(request.method === 'HEAD' ? undefined : fs.readFileSync(file));
    } catch { response.writeHead(503); response.end('Local installation needs repair.'); }
  });
  return server;
}
if (require.main === module) {
  const root = path.resolve(process.argv[2] || path.join(__dirname, '..'));
  const server = createServer(root);
  server.on('error', error => { console.error(error.message); process.exitCode = 1; });
  // localhost resolves to a loopback address; never bind to the LAN interfaces.
  server.listen(3000, 'localhost', () => console.log(`AI for Excel local host started (${process.pid}).`));
}
module.exports = { releasePath, assetPath, activate, createServer, clientLogRecord };
