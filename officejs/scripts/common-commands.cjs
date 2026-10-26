// Per-device command text lives outside immutable releases and browser profiles.
const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const MAX_COMMANDS = 50;
const MAX_LENGTH = 2000;
const MAX_BYTES = 512 * 1024;
const ROUTE = '/__local/common-commands';

function command(value) {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > MAX_LENGTH) throw new Error('INVALID_COMMAND');
  return value.trim();
}
function commands(value) {
  if (!Array.isArray(value) || value.length > MAX_COMMANDS) throw new Error('INVALID_COMMANDS');
  return [...new Set(value.map(command))];
}
function filePath(root) { return path.join(root, 'user-data', 'common-commands.json'); }
function readCommands(root) {
  const file = filePath(root);
  if (!fs.existsSync(file)) return { version: 1, commands: [], exists: false };
  try {
    if (fs.statSync(file).size > MAX_BYTES) throw new Error('TOO_LARGE');
    const data = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
    if (!data || data.version !== 1 || Object.keys(data).some(key => !['version', 'commands', 'updatedAt'].includes(key))) throw new Error('INVALID_FILE');
    return { version: 1, commands: commands(data.commands), exists: true };
  } catch { throw new Error('COMMAND_FILE_INVALID'); }
}
function atomicWrite(file, content) {
  const temporary = `${file}.${randomUUID()}.tmp`;
  let fd;
  try {
    fd = fs.openSync(temporary, 'wx', 0o600);
    fs.writeFileSync(fd, content, 'utf8');
    fs.fsyncSync(fd); fs.closeSync(fd); fd = undefined;
    fs.renameSync(temporary, file);
  } finally {
    if (fd !== undefined) fs.closeSync(fd);
    if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
  }
}
function writeCommands(root, list) {
  const file = filePath(root), validated = commands(list);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  // Never overwrite malformed user data or the last good backup with it.
  const previous = readCommands(root);
  if (previous.exists) atomicWrite(`${file}.bak`, fs.readFileSync(file));
  atomicWrite(file, JSON.stringify({ version: 1, updatedAt: new Date().toISOString(), commands: validated }, null, 2) + '\n');
  return { version: 1, commands: validated, exists: true };
}
function addCommand(root, value) {
  const normalized = command(value), previous = readCommands(root);
  return writeCommands(root, [normalized, ...previous.commands.filter(item => item !== normalized)].slice(0, MAX_COMMANDS));
}
function migrateCommands(root, values) {
  const legacy = commands(values), previous = readCommands(root);
  // The file is authoritative once created. An old pane/cache must not
  // resurrect entries that a newer pane has deleted or edited.
  return previous.exists ? previous : writeCommands(root, legacy);
}
function changeCommand(root, value, replacement) {
  const original = command(value), next = replacement === undefined ? undefined : command(replacement);
  const previous = readCommands(root), index = previous.commands.indexOf(original);
  if (index === -1) throw new Error('COMMAND_CHANGED');
  if (next === original) return previous;
  if (next !== undefined && previous.commands.includes(next)) throw new Error('COMMAND_EXISTS');
  const list = [...previous.commands];
  if (next === undefined) list.splice(index, 1); else list[index] = next;
  return writeCommands(root, list);
}
function handleCommonCommands(root, request, response) {
  if (request.url !== ROUTE) return false;
  const port = request.socket.localPort;
  const origin = `https://localhost:${port}`;
  const send = (status, value) => {
    response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
    response.end(JSON.stringify(value));
  };
  // Reject DNS rebinding hosts and all cross-origin access, including native GETs
  // carrying an Origin. No permissive CORS headers or generic filesystem API.
  if (request.headers.host !== `localhost:${port}` && request.headers.host !== `127.0.0.1:${port}`) { send(403, { error: 'FORBIDDEN' }); return true; }
  if (request.headers.origin && request.headers.origin !== origin) { send(403, { error: 'FORBIDDEN' }); return true; }
  if (request.method === 'GET') {
    try { send(200, readCommands(root)); } catch { send(503, { error: 'COMMAND_FILE_UNAVAILABLE' }); }
    return true;
  }
  if (request.method !== 'POST') { send(405, { error: 'METHOD_NOT_ALLOWED' }); return true; }
  if (request.headers.origin !== origin || request.headers['content-type']?.split(';')[0].trim().toLowerCase() !== 'application/json') { send(403, { error: 'FORBIDDEN' }); return true; }
  let bytes = 0, chunks = [], oversized = false;
  request.on('error', () => { chunks = []; });
  request.on('data', chunk => {
    bytes += chunk.length;
    if (bytes > MAX_BYTES) { if (!oversized) send(413, { error: 'BODY_TOO_LARGE' }); oversized = true; chunks = []; return; }
    if (!oversized) chunks.push(chunk);
  });
  request.on('end', () => {
    if (oversized) return;
    let body;
    try {
      body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      if (!body || typeof body !== 'object' || Array.isArray(body)) throw new Error('INVALID_REQUEST');
      const keys = ['add', 'delete'].includes(body.action) ? ['action', 'command'] : body.action === 'update' ? ['action', 'command', 'replacement'] : body.action === 'migrate' ? ['action', 'commands'] : [];
      if (!keys.length || Object.keys(body).some(key => !keys.includes(key))) throw new Error('INVALID_REQUEST');
      if (body.action === 'migrate') commands(body.commands);
      else { command(body.command); if (body.action === 'update') command(body.replacement); }
    } catch { send(400, { error: 'INVALID_REQUEST' }); return; }
    try { send(200, body.action === 'add' ? addCommand(root, body.command) : body.action === 'migrate' ? migrateCommands(root, body.commands) : changeCommand(root, body.command, body.action === 'update' ? body.replacement : undefined)); }
    catch (error) {
      if (['COMMAND_CHANGED', 'COMMAND_EXISTS'].includes(error.message)) send(409, { error: error.message });
      else send(503, { error: 'COMMAND_FILE_UNAVAILABLE' });
    }
  });
  return true;
}
module.exports = { filePath, readCommands, addCommand, migrateCommands, changeCommand, handleCommonCommands };
