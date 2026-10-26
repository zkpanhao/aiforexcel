// This file API is available only to the local taskpane, never static serving.
const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const LIMIT = 128 * 1024;
function filePath(root) { return path.join(root, 'user-data', 'connection.json'); }
function normalize(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).some(key => !['provider', 'baseUrl', 'model', 'apiKey', 'rememberKey'].includes(key))) throw new Error('INVALID_SETTINGS');
  if (typeof value.baseUrl !== 'string' || value.baseUrl.length > 4096 || typeof value.model !== 'string' || value.model.length > 2000 || typeof value.apiKey !== 'string' || value.apiKey.length > 32768 || typeof value.rememberKey !== 'boolean') throw new Error('INVALID_SETTINGS');
  const url = new URL(value.baseUrl.trim());
  if (url.username || url.password || url.search || url.hash || (url.protocol !== 'https:' && !(url.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(url.hostname)))) throw new Error('INVALID_SETTINGS');
  const provider = value.provider ?? 'custom';
  if (!['custom', 'openai', 'deepseek', 'qwen', 'glm', 'minimax'].includes(provider)) throw new Error('INVALID_SETTINGS');
  return { provider, baseUrl: url.href.replace(/\/$/, ''), model: value.model.trim(), rememberKey: value.rememberKey, apiKey: value.rememberKey ? value.apiKey.trim() : '' };
}
function stored(settings) {
  if (!settings) return null;
  const { apiKey, ...fields } = settings;
  return { ...fields, ...(fields.rememberKey && apiKey ? { apiKey, keyBaseUrl: fields.baseUrl } : {}) };
}
function readConnection(root) {
  const file = filePath(root);
  if (!fs.existsSync(file)) return { version: 1, exists: false, settings: null };
  try {
    if (fs.statSync(file).size > LIMIT) throw new Error('TOO_LARGE');
    const data = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
    if (!data || data.version !== 1 || Object.keys(data).some(key => !['version', 'updatedAt', 'settings'].includes(key))) throw new Error('INVALID_FILE');
    let settings = null;
    if (data.settings !== null) {
      const { keyBaseUrl, ...fields } = data.settings;
      settings = normalize({ ...fields, apiKey: fields.apiKey ?? '' });
      // A manually edited URL must never restore a credential for another host.
      if (keyBaseUrl !== settings.baseUrl) settings.apiKey = '';
    }
    return { version: 1, exists: true, settings };
  } catch { throw new Error('CONNECTION_FILE_INVALID'); }
}
function atomicWrite(file, content) {
  const temp = `${file}.${randomUUID()}.tmp`;
  let fd;
  try {
    fd = fs.openSync(temp, 'wx', 0o600); fs.writeFileSync(fd, content, 'utf8');
    fs.fsyncSync(fd); fs.closeSync(fd); fd = undefined; fs.renameSync(temp, file);
  } finally {
    if (fd !== undefined) fs.closeSync(fd);
    if (fs.existsSync(temp)) fs.unlinkSync(temp);
  }
}
function writeConnection(root, settings) {
  const file = filePath(root), previous = readConnection(root);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const serialize = value => JSON.stringify({ version: 1, updatedAt: new Date().toISOString(), settings: stored(value) }, null, 2) + '\n';
  // Backups preserve endpoint/model only. Forget/opt-out must remove the key
  // from every file we manage, rather than retaining it in an old backup.
  if (previous.exists) atomicWrite(`${file}.bak`, serialize(previous.settings && { ...previous.settings, rememberKey: false, apiKey: '' }));
  atomicWrite(file, serialize(settings));
  return readConnection(root);
}
function saveConnection(root, value) { return writeConnection(root, normalize(value)); }
function migrateConnection(root, value) {
  const settings = value === null ? null : normalize(value), current = readConnection(root);
  return current.exists ? current : writeConnection(root, settings);
}
function forgetKey(root) {
  const current = readConnection(root);
  return writeConnection(root, current.settings && { ...current.settings, apiKey: '', rememberKey: false });
}
function handleConnection(root, req, res) {
  if (req.url !== '/__local/connection') return false;
  const port = req.socket.localPort, origin = `https://localhost:${port}`;
  const send = (code, data) => { res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' }); res.end(JSON.stringify(data)); };
  // A custom header also prevents cross-site simple requests and navigation.
  if (![`localhost:${port}`, `127.0.0.1:${port}`].includes(req.headers.host) || req.headers['x-aiforexcel-settings'] !== '1' || (req.headers.origin && req.headers.origin !== origin) || req.headers['sec-fetch-site'] === 'cross-site') {
    send(403, { error: 'FORBIDDEN' }); return true;
  }
  if (req.method === 'GET') {
    try { send(200, readConnection(root)); } catch { send(503, { error: 'CONNECTION_FILE_UNAVAILABLE' }); }
    return true;
  }
  if (req.method !== 'POST') { send(405, { error: 'METHOD_NOT_ALLOWED' }); return true; }
  if (req.headers.origin !== origin || req.headers['content-type']?.split(';')[0].trim().toLowerCase() !== 'application/json') { send(403, { error: 'FORBIDDEN' }); return true; }
  let bytes = 0, chunks = [], oversized = false;
  req.on('error', () => { chunks = []; });
  req.on('data', chunk => {
    bytes += chunk.length;
    if (bytes > LIMIT) { if (!oversized) send(413, { error: 'BODY_TOO_LARGE' }); oversized = true; chunks = []; return; }
    if (!oversized) chunks.push(chunk);
  });
  req.on('end', () => {
    if (oversized) return;
    let body;
    try {
      body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      if (!body || Array.isArray(body) || typeof body !== 'object') throw new Error('INVALID');
      const keys = body.action === 'forget' ? ['action'] : ['save', 'migrate'].includes(body.action) ? ['action', 'settings'] : [];
      if (!keys.length || Object.keys(body).some(key => !keys.includes(key))) throw new Error('INVALID');
      if (body.action !== 'forget' && !(body.action === 'migrate' && body.settings === null)) normalize(body.settings);
    } catch { send(400, { error: 'INVALID_REQUEST' }); return; }
    try { send(200, body.action === 'save' ? saveConnection(root, body.settings) : body.action === 'migrate' ? migrateConnection(root, body.settings) : forgetKey(root)); }
    catch { send(503, { error: 'CONNECTION_FILE_UNAVAILABLE' }); }
  });
  return true;
}
module.exports = { filePath, normalize, readConnection, saveConnection, migrateConnection, forgetKey, handleConnection };
