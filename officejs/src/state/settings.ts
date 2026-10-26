import { Connection, endpoint, inferProvider, ProviderId } from "../llm/client";

const STORAGE_KEY = "ai-for-excel.connection.v1";
const COMMON_COMMANDS_KEY = "ai-for-excel.common-commands.v1";
const MAX_COMMON_COMMANDS = 50;
const MAX_COMMON_COMMAND_LENGTH = 2000;
type SettingsStorage = Pick<Storage, "getItem" | "setItem" | "removeItem">;
export interface ConnectionSettings extends Connection { rememberKey: boolean; }
export const canonicalBaseUrl = (value: string): string => endpoint(value.trim(), "").replace(/\/$/, "");

export function saveConnection(storage: SettingsStorage, settings: ConnectionSettings): void {
  const baseUrl = canonicalBaseUrl(settings.baseUrl);
  storage.setItem(STORAGE_KEY, JSON.stringify({
    version: 2, ...(settings.provider ? { provider: settings.provider } : {}), baseUrl, model: settings.model.trim(), rememberKey: settings.rememberKey,
    ...(settings.rememberKey && settings.apiKey.trim() ? { apiKey: settings.apiKey.trim(), keyBaseUrl: baseUrl } : {}),
  }));
}

export function loadConnection(storage: Pick<Storage, 'getItem'>): ConnectionSettings | undefined {
  const raw = storage.getItem(STORAGE_KEY);
  if (!raw) return;
  try {
    const data = JSON.parse(raw);
    if (!data || (data.version !== 1 && data.version !== 2) || typeof data.baseUrl !== "string" || typeof data.model !== "string" || typeof data.rememberKey !== "boolean") return;
    const baseUrl = canonicalBaseUrl(data.baseUrl);
    // A remembered credential belongs to one endpoint, never a newly edited URL.
    const apiKey = data.rememberKey && data.keyBaseUrl === baseUrl && typeof data.apiKey === "string" ? data.apiKey : "";
    const provider = typeof data.provider === "string" && ["custom", "openai", "deepseek", "qwen", "glm", "minimax"].includes(data.provider) ? data.provider as ProviderId : inferProvider(baseUrl);
    return { provider, baseUrl, model: data.model, rememberKey: data.rememberKey, apiKey };
  } catch { return; }
}

export function forgetSavedKey(storage: SettingsStorage): void {
  const current = loadConnection(storage);
  if (current) saveConnection(storage, { ...current, apiKey: "", rememberKey: false });
  else storage.removeItem(STORAGE_KEY);
}

export function loadCommonCommands(storage: Pick<Storage, "getItem">): string[] {
  const raw = storage.getItem(COMMON_COMMANDS_KEY);
  if (!raw) return [];
  try {
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item): item is string => typeof item === "string" && item.trim().length > 0)
      .map(item => item.trim().slice(0, MAX_COMMON_COMMAND_LENGTH)).slice(0, MAX_COMMON_COMMANDS);
  } catch { return []; }
}

export function saveCommonCommand(storage: SettingsStorage, command: string): string[] {
  const normalized = command.trim().slice(0, MAX_COMMON_COMMAND_LENGTH);
  if (!normalized) throw new Error("请先在对话框中输入命令");
  const next = [normalized, ...loadCommonCommands(storage).filter(item => item !== normalized)].slice(0, MAX_COMMON_COMMANDS);
  storage.setItem(COMMON_COMMANDS_KEY, JSON.stringify(next));
  return next;
}
