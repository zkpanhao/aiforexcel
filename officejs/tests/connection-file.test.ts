import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import https from 'node:https';
import { randomUUID } from 'node:crypto';
import { readFileConnection, saveFileConnection, forgetFileKey } from '../src/state/connection-file';
const { filePath, readConnection, saveConnection, migrateConnection, forgetKey, handleConnection } = require('../scripts/connection-settings.cjs');
const mkcert = require('mkcert');
const fixture = () => ({ provider: 'custom' as const, baseUrl: 'https://example.com/v1', model: 'test-model', apiKey: `test-only-${randomUUID()}`, rememberKey: true });
const temporary = () => fs.mkdtempSync(path.join(os.tmpdir(), 'ai-excel-connection-'));

test('connection file persists opted-in key, binds it to endpoint and never keeps a backup credential', () => {
  const root = temporary(), value = fixture();
  try {
    assert.equal(readConnection(root).exists, false);
    assert.deepEqual(migrateConnection(root, value).settings, value);
    assert.deepEqual(migrateConnection(root, { ...value, apiKey: 'stale-cache' }).settings, value);
    const changed = { ...value, model: 'other-model' };
    saveConnection(root, changed);
    const backup = fs.readFileSync(filePath(root) + '.bak', 'utf8');
    assert.equal(backup.includes(value.apiKey), false); assert.equal(backup.includes('keyBaseUrl'), false);
    const raw = JSON.parse(fs.readFileSync(filePath(root), 'utf8'));
    raw.settings.baseUrl = 'https://other.example/v1';
    fs.writeFileSync(filePath(root), JSON.stringify(raw));
    assert.equal(readConnection(root).settings.apiKey, '');
    saveConnection(root, changed);
    forgetKey(root);
    assert.deepEqual(readConnection(root).settings, { ...changed, apiKey: '', rememberKey: false });
    for (const name of fs.readdirSync(path.dirname(filePath(root)))) assert.equal(fs.readFileSync(path.join(path.dirname(filePath(root)), name), 'utf8').includes(value.apiKey), false);
    assert.equal(migrateConnection(root, value).settings.apiKey, '', 'stale legacy cache cannot resurrect a forgotten key');
    saveConnection(root, { ...value, rememberKey: false });
    assert.equal(readConnection(root).settings.apiKey, '');
    assert.equal(fs.readFileSync(filePath(root), 'utf8').includes(value.apiKey), false);
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test('corrupt configuration and interrupted atomic write retain original file and expose no secret errors', () => {
  const root = temporary(), value = fixture(), rename = fs.renameSync;
  try {
    saveConnection(root, value);
    const original = fs.readFileSync(filePath(root), 'utf8');
    fs.renameSync = ((from, to) => { if (to === filePath(root)) throw new Error('simulated disk failure'); rename(from, to); }) as typeof fs.renameSync;
    assert.throws(() => forgetKey(root));
    fs.renameSync = rename;
    assert.equal(fs.readFileSync(filePath(root), 'utf8'), original);
    assert.equal(fs.readdirSync(path.dirname(filePath(root))).some(name => name.endsWith('.tmp')), false);
    fs.writeFileSync(filePath(root), '{broken');
    for (const action of [() => readConnection(root), () => saveConnection(root, value), () => migrateConnection(root, value), () => forgetKey(root)]) assert.throws(action, /CONNECTION_FILE_INVALID/);
    assert.equal(fs.readFileSync(filePath(root), 'utf8'), '{broken');
  } finally { fs.renameSync = rename; fs.rmSync(root, { recursive: true, force: true }); }
});

test('connection API requires local host and custom header, rejects cross-site/invalid/oversized requests', async () => {
  const root = temporary(), value = fixture();
  const ca = await mkcert.createCA({ organization: 'Connection test', countryCode: 'CN', locality: 'Test', state: 'Test', validity: 1 });
  const cert = await mkcert.createCert({ ca, domains: ['localhost', '127.0.0.1'], validity: 1 });
  const server = https.createServer({ cert: cert.cert, key: cert.key }, (req, res) => { if (!handleConnection(root, req, res)) { res.writeHead(404); res.end(); } });
  await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve));
  const port = (server.address() as { port: number }).port;
  function request(method = 'GET', body?: unknown, headers: Record<string, string> = {}) {
    return new Promise<{ status: number; data: any }>((resolve, reject) => {
      const req = https.request({ hostname: '127.0.0.1', servername: 'localhost', port, path: '/__local/connection', method, ca: ca.cert,
        headers: { Origin: `https://localhost:${port}`, 'X-AIForExcel-Settings': '1', 'Content-Type': 'application/json', ...headers } }, res => {
        let data = ''; res.on('data', chunk => data += chunk); res.on('end', () => resolve({ status: res.statusCode!, data: JSON.parse(data) }));
      });
      req.on('error', reject); req.end(body === undefined ? undefined : JSON.stringify(body));
    });
  }
  try {
    assert.equal((await request()).data.exists, false);
    assert.equal((await request('POST', { action: 'save', settings: value })).status, 200);
    assert.deepEqual((await request()).data.settings, value);
    for (const headers of [{ 'X-AIForExcel-Settings': '' }, { Origin: 'https://external.invalid' }, { Host: 'external.invalid' }, { 'Sec-Fetch-Site': 'cross-site' }]) {
      const denied = await request('GET', undefined, headers);
      assert.equal(denied.status, 403); assert.equal(JSON.stringify(denied.data).includes(value.apiKey), false);
    }
    assert.equal((await request('POST', { action: 'forget' }, { Origin: '' })).status, 403);
    assert.equal((await request('POST', { action: 'forget' }, { 'Content-Type': 'text/plain' })).status, 403);
    for (const body of [{ action: 'save', settings: { ...value, baseUrl: 'https://name:password@example.com' } }, { action: 'save', settings: { ...value, rememberKey: 'yes' } }, { action: 'save', settings: { ...value, apiKey: 'x'.repeat(32769) } }, { action: 'forget', path: '../outside' }, { action: 'unknown' }]) assert.equal((await request('POST', body)).status, 400);
    assert.equal((await request('POST', { action: 'save', settings: { ...value, apiKey: 'x'.repeat(150000) } })).status, 413);
    assert.equal((await request('DELETE')).status, 405);
    assert.equal((await request('POST', { action: 'forget' })).data.settings.apiKey, '');
    fs.writeFileSync(filePath(root), 'broken');
    const failed = await request(); assert.equal(failed.status, 503); assert.equal(JSON.stringify(failed.data).includes(value.apiKey), false);
  } finally { await new Promise<void>(resolve => server.close(() => resolve())); fs.rmSync(root, { recursive: true, force: true }); }
});

test('file client migrates old cache once, works with blocked cache, and serializes rapid saves/forget', async () => {
  const root = temporary(), originalFetch = globalThis.fetch, value = fixture();
  const legacy = new Map([['ai-for-excel.connection.v1', JSON.stringify({ version: 2, ...value, keyBaseUrl: value.baseUrl })], ['ai-for-excel.common-commands.v1', 'untouched']]);
  const storage = { getItem: (key: string) => legacy.get(key) ?? null, removeItem: (key: string) => { legacy.delete(key); } };
  globalThis.fetch = (async (_url, init) => {
    assert.equal((init?.headers as Record<string, string>)['X-AIForExcel-Settings'], '1');
    const body = init?.body ? JSON.parse(String(init.body)) : null;
    await new Promise(resolve => setTimeout(resolve, body?.settings?.apiKey === value.apiKey ? 15 : 1));
    const data = !body ? readConnection(root) : body.action === 'migrate' ? migrateConnection(root, body.settings) : body.action === 'save' ? saveConnection(root, body.settings) : forgetKey(root);
    return new Response(JSON.stringify(data));
  }) as typeof fetch;
  try {
    assert.deepEqual(await readFileConnection(storage), value);
    assert.equal(legacy.has('ai-for-excel.connection.v1'), false); assert.equal(legacy.get('ai-for-excel.common-commands.v1'), 'untouched');
    const blocked = { getItem: () => { throw new Error('blocked'); }, removeItem: () => { throw new Error('blocked'); } };
    assert.deepEqual(await readFileConnection(blocked), value);
    const snapshot = { ...value };
    const first = saveFileConnection(snapshot); snapshot.baseUrl = 'https://must-not-save.invalid';
    await Promise.all([first, saveFileConnection({ ...value, model: 'latest' }), forgetFileKey()]);
    assert.deepEqual(await readFileConnection(blocked), { ...value, model: 'latest', apiKey: '', rememberKey: false });
    legacy.set('ai-for-excel.connection.v1', JSON.stringify({ version: 2, ...value, keyBaseUrl: value.baseUrl }));
    assert.equal((await readFileConnection(storage))?.apiKey, '');
  } finally { globalThis.fetch = originalFetch; fs.rmSync(root, { recursive: true, force: true }); }
});

test('failed file operation retains cache and rejects without echoing credentials, later retries still work', async () => {
  const originalFetch = globalThis.fetch, value = fixture(); let removed = false;
  const storage = { getItem: () => JSON.stringify({ version: 2, ...value, keyBaseUrl: value.baseUrl }), removeItem: () => { removed = true; } };
  try {
    globalThis.fetch = (async (_url, init) => !init?.body ? new Response(JSON.stringify({ version: 1, exists: false, settings: null })) : new Response(value.apiKey, { status: 503 })) as typeof fetch;
    await assert.rejects(readFileConnection(storage), error => error instanceof Error && /未确认/.test(error.message) && !error.message.includes(value.apiKey));
    assert.equal(removed, false);
    await assert.rejects(saveFileConnection(value, storage), /未确认/);
    await assert.rejects(forgetFileKey(storage), /未确认/);
    assert.equal(removed, false);
    globalThis.fetch = (async () => new Response(JSON.stringify({ version: 1, exists: true, settings: { ...value, apiKey: 123 } }))) as typeof fetch;
    await assert.rejects(readFileConnection(storage), /未确认/);
    globalThis.fetch = (async () => new Response(JSON.stringify({ version: 1, exists: true, settings: { ...value, apiKey: '', rememberKey: false } }))) as typeof fetch;
    await forgetFileKey(storage); assert.equal(removed, true);
  } finally { globalThis.fetch = originalFetch; }
});
