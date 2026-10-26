import { loadCommonCommands } from "./settings";

const ROUTE = "/__local/common-commands";
const LEGACY_KEY = "ai-for-excel.common-commands.v1";
type LegacyStorage = Pick<Storage, "getItem" | "removeItem">;
interface SavedCommands { version: 1; commands: string[]; exists: boolean; }
class CommandConflict extends Error {}
async function request(body?: { action: "add" | "delete"; command: string } | { action: "update"; command: string; replacement: string } | { action: "migrate"; commands: string[] }): Promise<SavedCommands> {
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 10000);
  try {
    const response = await fetch(ROUTE, { method: body ? "POST" : "GET", cache: "no-store", credentials: "omit", redirect: "error", signal: controller.signal,
      ...(body ? { headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) } : {}),
    });
    if (response.status === 409) {
      const conflict = await response.json();
      throw new CommandConflict(conflict?.error === 'COMMAND_EXISTS' ? '已存在相同的常用命令，请修改内容后再保存。' : '这条命令已被修改或删除，请重新打开常用命令列表后再操作。');
    }
    if (!response.ok) throw new Error("FILE_UNAVAILABLE");
    const data: unknown = await response.json();
    const saved = data as SavedCommands;
    if (!saved || saved.version !== 1 || typeof saved.exists !== "boolean" || (body && !saved.exists) || !Array.isArray(saved.commands) || saved.commands.length > 50 || saved.commands.some(item => typeof item !== "string" || !item.trim() || item.length > 2000)) throw new Error("INVALID_FILE_RESPONSE");
    return saved;
  } catch (error) { if (error instanceof CommandConflict) throw error; throw new Error("无法读写本机常用命令文件，请检查本地服务和文件；原文件已保留，本次未确认保存或删除。"); }
  finally { clearTimeout(timer); }
}
export async function readFileCommands(storage?: LegacyStorage): Promise<string[]> {
  let legacy: string[] = [];
  try { if (storage) legacy = loadCommonCommands(storage); } catch { /* File storage works even if browser storage is blocked. */ }
  const current = await request();
  if (!current.exists) {
    const migrated = await request({ action: "migrate", commands: legacy });
    // Remove only the old command key, after durable storage succeeded. Never
    // touch the separately stored API connection settings.
    if (legacy.length) { try { storage?.removeItem(LEGACY_KEY); } catch { /* Idempotent migration can safely retry. */ } }
    return migrated.commands;
  }
  if (legacy.length) { try { storage?.removeItem(LEGACY_KEY); } catch { /* Existing files never reimport stale cached commands. */ } }
  return current.commands;
}
export async function saveFileCommand(command: string, storage?: LegacyStorage): Promise<string[]> {
  const normalized = command.trim();
  if (!normalized) throw new Error("请先在对话框中输入命令");
  if (normalized.length > 2000) throw new Error("单条常用命令最多 2000 字符，请精简后保存");
  await readFileCommands(storage);
  return (await request({ action: "add", command: normalized })).commands;
}
export async function updateFileCommand(command: string, replacement: string): Promise<string[]> {
  const normalized = replacement.trim();
  if (!normalized) throw new Error("命令内容不能为空");
  if (normalized.length > 2000) throw new Error("单条常用命令最多 2000 字符，请精简后保存");
  return (await request({ action: "update", command, replacement: normalized })).commands;
}
export async function deleteFileCommand(command: string): Promise<string[]> {
  return (await request({ action: "delete", command })).commands;
}
