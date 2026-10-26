import { canonicalBaseUrl, ConnectionSettings, loadConnection } from './settings';
import { inferProvider } from '../llm/client';
type LegacyStorage = Pick<Storage, 'getItem' | 'removeItem'>;
const LEGACY_KEY = 'ai-for-excel.connection.v1';
interface Saved { version: 1; exists: boolean; settings: ConnectionSettings | null; }
type Action = { action: 'save' | 'migrate'; settings: ConnectionSettings | null } | { action: 'forget' };
function normalized(settings: ConnectionSettings): ConnectionSettings {
  if (!settings || typeof settings.baseUrl !== 'string' || settings.baseUrl.length > 4096 || typeof settings.model !== 'string' || settings.model.length > 2000 || typeof settings.apiKey !== 'string' || settings.apiKey.length > 32768 || typeof settings.rememberKey !== 'boolean') throw new Error('INVALID_SETTINGS');
  const baseUrl = canonicalBaseUrl(settings.baseUrl), provider = settings.provider ?? inferProvider(baseUrl);
  if (!['custom', 'openai', 'deepseek', 'qwen', 'glm', 'minimax'].includes(provider)) throw new Error('INVALID_SETTINGS');
  return { provider, baseUrl, model: settings.model.trim(), rememberKey: settings.rememberKey, apiKey: settings.rememberKey ? settings.apiKey.trim() : '' };
}
async function request(body?: Action): Promise<Saved> {
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 10000);
  try {
    const response = await fetch('/__local/connection', { method: body ? 'POST' : 'GET', cache: 'no-store', credentials: 'omit', redirect: 'error', signal: controller.signal,
      headers: { 'X-AIForExcel-Settings': '1', ...(body ? { 'Content-Type': 'application/json' } : {}) }, ...(body ? { body: JSON.stringify(body) } : {}) });
    if (!response.ok) throw new Error('UNAVAILABLE');
    const data = await response.json() as Saved;
    if (!data || data.version !== 1 || typeof data.exists !== 'boolean' || (body && !data.exists) || (!data.exists && data.settings !== null)) throw new Error('INVALID_RESPONSE');
    return { ...data, settings: data.settings === null ? null : normalized(data.settings) };
  } catch { throw new Error('无法读写本机连接配置文件，请检查本地服务和文件；本次未确认保存或清除。'); }
  finally { clearTimeout(timer); }
}
let pending: Promise<unknown> = Promise.resolve();
function serial<T>(action: () => Promise<T>): Promise<T> {
  const next = pending.then(action); pending = next.catch(() => undefined); return next;
}
function clearLegacy(storage?: LegacyStorage): void { try { storage?.removeItem(LEGACY_KEY); } catch { /* Retrying reads the authoritative file, never reimports a forgotten key. */ } }
export function readFileConnection(storage?: LegacyStorage): Promise<ConnectionSettings | undefined> {
  return serial(async () => {
    let result = await request();
    if (!result.exists) {
      let legacy: ConnectionSettings | undefined;
      try { if (storage) legacy = loadConnection(storage); } catch { /* File access is independent of browser storage. */ }
      result = await request({ action: 'migrate', settings: legacy ? normalized(legacy) : null });
    }
    clearLegacy(storage);
    return result.settings ?? undefined;
  });
}
export function saveFileConnection(settings: ConnectionSettings, storage?: LegacyStorage): Promise<void> {
  // Snapshot before queuing: quick typing/provider changes must keep event order.
  let snapshot: ConnectionSettings;
  try { snapshot = normalized(settings); } catch { return Promise.reject(new Error('未能保存：请检查接口地址、模型和密钥长度。')); }
  return serial(async () => { await request({ action: 'save', settings: snapshot }); clearLegacy(storage); });
}
export function forgetFileKey(storage?: LegacyStorage): Promise<void> {
  return serial(async () => { await request({ action: 'forget' }); clearLegacy(storage); });
}
